import itertools
import json
from dataclasses import replace
from unittest.mock import patch

import httpx
import pytest
from azure.core.credentials import AccessToken

from provision import API_VERSION, DeploymentType, Settings, build_resources, main, send_requests


@pytest.fixture
def settings(monkeypatch: pytest.MonkeyPatch) -> Settings:
    for name in (
        "SEARCH_ADMIN_KEY",
        "SEARCH_DATASOURCE_NAME",
        "SEARCH_SKILLSET_NAME",
        "SEARCH_INDEXER_NAME",
        "SHAREPOINT_APPLICATION_SECRET",
        "SHAREPOINT_LIBRARY_URL",
        "AZURE_OPENAI_KEY",
        "AZURE_AI_SERVICES_KEY",
        "INDEXER_SCHEDULE",
    ):
        monkeypatch.delenv(name, raising=False)
    for name, value in {
        "SEARCH_ENDPOINT": "https://example.search.windows.net",
        "DEPLOYMENT_TYPE": "all_embeddings",
        "SEARCH_INDEX_NAME": "test-rag",
        "SHAREPOINT_SITE_URL": "https://example.sharepoint.com/sites/Hub",
        "SHAREPOINT_APPLICATION_ID": "11111111-1111-1111-1111-111111111111",
        "SHAREPOINT_TENANT_ID": "22222222-2222-2222-2222-222222222222",
        "SEARCH_IDENTITY_CLIENT_ID": "33333333-3333-3333-3333-333333333333",
        "AZURE_OPENAI_ENDPOINT": "https://example.openai.azure.com",
        "AZURE_OPENAI_DEPLOYMENT": "my-embedding-deployment",
        "AZURE_AI_SERVICES_ENDPOINT": "https://example.cognitiveservices.azure.com",
    }.items():
        monkeypatch.setenv(name, value)
    return Settings.from_env()


@pytest.mark.parametrize("deployment_type", list(DeploymentType))
def test_no_acl_or_wizard_identity(settings: Settings, deployment_type: DeploymentType) -> None:
    settings = Settings.from_env(deployment_type)
    resources = build_resources(settings)
    expected = ["datasources", "indexes", "indexers"]
    if deployment_type != DeploymentType.NO_EMBEDDING:
        expected.insert(2, "skillsets")
    assert list(resources) == expected
    source = resources["datasources"]
    assert source["type"] == "sharepoint"
    assert source["container"] == {"name": "defaultSiteLibrary"}
    assert "identity" not in source and "dataSource" not in source
    assert "indexerPermissionOptions" not in source
    document = json.dumps(resources)
    for removed in (
        "permissionFilter",
        "UserIds",
        "GroupIds",
        "metadata_user_ids",
        "metadata_group_ids",
        "sharePointConnectorAppRegistration",
        "@odata.etag",
        "@odata.context",
    ):
        assert removed not in document
    connection = source["credentials"]["connectionString"]
    assert f"ApplicationId={settings.application_id};" in connection
    assert f"FederatedCredentialApplicationId={settings.identity_client_id}" in connection
    assert "ApplicationSecret" not in connection


def test_preserves_reference_pipeline(settings: Settings) -> None:
    resources = build_resources(settings)
    index, skillset, indexer = (resources[name] for name in ("indexes", "skillsets", "indexers"))
    fields = {item["name"]: item for item in index["fields"]}
    assert fields["chunk_id"]["key"] and fields["chunk_id"]["analyzer"] == "keyword"
    assert fields["text_parent_id"]["filterable"] and fields["image_parent_id"]["filterable"]
    assert fields["text_vector"]["dimensions"] == 3072
    assert fields["image_vector"]["dimensions"] == 1024
    assert index["semantic"]["configurations"][0]["prioritizedFields"]["titleField"] == {
        "fieldName": "title"
    }
    ocr, merge, split, text, image = skillset["skills"]
    assert ocr["@odata.type"] == "#Microsoft.Skills.Vision.OcrSkill"
    assert merge["inputs"][1]["source"] == "/document/normalized_images/*/text"
    assert split["maximumPageLength"] == 2000 and split["pageOverlapLength"] == 500
    assert split["unit"] == "characters"
    assert text["inputs"][0]["source"] == "/document/pages/*"
    assert text["dimensions"] == fields["text_vector"]["dimensions"]
    assert image["outputs"][0]["targetName"] == "image_vector"
    vectorizers = index["vectorSearch"]["vectorizers"]
    assert vectorizers[0]["azureOpenAIParameters"]["deploymentId"] == settings.openai_deployment
    assert image["modelVersion"] == vectorizers[1]["aiServicesVisionParameters"]["modelVersion"]
    projections = skillset["indexProjections"]
    assert projections["parameters"]["projectionMode"] == "skipIndexingParentDocuments"
    for selector in projections["selectors"]:
        assert selector["targetIndexName"] == index["name"]
        assert selector["parentKeyFieldName"] in fields
        mappings = {mapping["name"]: mapping["source"] for mapping in selector["mappings"]}
        assert mappings["title"] == "/document/metadata_spo_item_name"
        assert "chunk_id" not in mappings and selector["parentKeyFieldName"] not in mappings
        assert set(mappings) <= fields.keys()
    assert indexer["dataSourceName"] == resources["datasources"]["name"]
    assert indexer["targetIndexName"] == index["name"]
    assert indexer["skillsetName"] == skillset["name"]
    assert indexer["parameters"]["configuration"]["imageAction"] == "generateNormalizedImages"
    assert indexer["schedule"] == {"interval": "PT1H"}


@pytest.mark.parametrize("keys", list(itertools.product([False, True], repeat=4)))
@pytest.mark.parametrize("deployment_type", list(DeploymentType))
def test_independent_auth_modes(
    settings: Settings,
    keys: tuple[bool, bool, bool, bool],
    deployment_type: DeploymentType,
) -> None:
    search_key, openai_key, vision_key, app_secret = keys
    config = replace(
        Settings.from_env(deployment_type),
        search_key="search-test-secret" if search_key else "",
        openai_key="openai-test-secret" if openai_key else "",
        ai_services_key="vision-test-secret" if vision_key else "",
        application_secret="sharepoint-test-secret" if app_secret else "",
    )
    resources = build_resources(config)
    source, index = (resources[name] for name in ("datasources", "indexes"))
    connection = source["credentials"]["connectionString"]
    assert ("ApplicationSecret=" in connection) is app_secret
    assert ("FederatedCredentialApplicationId=" in connection) is not app_secret
    if deployment_type != DeploymentType.NO_EMBEDDING:
        skillset = resources["skillsets"]
        vectorizers = {item["kind"]: item for item in index["vectorSearch"]["vectorizers"]}
        skills = {item["name"]: item for item in skillset["skills"]}
        if config.use_text:
            for item in (skills["embed-text"], vectorizers["azureOpenAI"]["azureOpenAIParameters"]):
                assert ("apiKey" in item) is openai_key
                assert "authIdentity" not in item
                assert item["resourceUri"] == config.openai_endpoint
        if config.use_images:
            image_vectorizer = vectorizers["aiServicesVision"]["aiServicesVisionParameters"]
            assert ("apiKey" in image_vectorizer) is vision_key
            assert image_vectorizer["resourceUri"] == config.ai_services_endpoint
            billing = skillset["cognitiveServices"]
            assert billing["@odata.type"].endswith("ByKey" if vision_key else "ByIdentity")
            assert ("key" in billing) is vision_key
            assert billing["subdomainUrl"] == config.ai_services_endpoint
            assert "identity" not in billing

    calls = []

    def handler(request: httpx.Request) -> httpx.Response:
        calls.append(request)
        assert request.method == "PUT"
        assert request.url.params["api-version"] == API_VERSION
        if search_key:
            assert request.headers["api-key"] == config.search_key
            assert "Authorization" not in request.headers
        else:
            assert request.headers["Authorization"] == "Bearer managed-test-token"
            assert "api-key" not in request.headers
        return httpx.Response(201, json={})

    with patch("provision.ManagedIdentityCredential") as identity:
        credential = identity.return_value.__enter__.return_value
        credential.get_token.return_value = AccessToken("managed-test-token", 4_000_000_000)
        send_requests(config, resources, transport=httpx.MockTransport(handler))
        if search_key:
            identity.assert_not_called()
        else:
            identity.assert_called_once_with()
            credential.get_token.assert_called_with("https://search.azure.com/.default")
    assert len(calls) == len(resources)
    for request, (collection, body) in zip(calls, resources.items(), strict=True):
        assert request.url.path == f"/{collection}/{body['name']}"
        assert json.loads(request.content) == body


@pytest.mark.parametrize(
    ("name", "value"),
    [
        ("SEARCH_ENDPOINT", "http://example.search.windows.net"),
        ("SEARCH_ENDPOINT", "https://example.search.windows.net.attacker.com"),
        ("AZURE_OPENAI_ENDPOINT", "https://example.openai.azure.com/openai/v1"),
        ("AZURE_AI_SERVICES_ENDPOINT", "https://example.com"),
        ("SHAREPOINT_APPLICATION_ID", "bad-guid"),
        ("SEARCH_IDENTITY_CLIENT_ID", ""),
        ("SHAREPOINT_SITE_URL", "https://example.sharepoint.com;ApplicationSecret=bad"),
        ("SHAREPOINT_LIBRARY_URL", "https://another.sharepoint.com/sites/Hub/Library"),
        ("SHAREPOINT_LIBRARY_URL", "https://example.sharepoint.com/sites/Hub/../Private"),
        ("SHAREPOINT_APPLICATION_SECRET", "bad;secret"),
        ("SEARCH_INDEX_NAME", "../index"),
        ("SEARCH_INDEX_NAME", "double--dash"),
        ("INDEXER_SCHEDULE", "PT4M"),
        ("INDEXER_SCHEDULE", "PT25H"),
        ("INDEXER_SCHEDULE", "invalid"),
        ("AZURE_OPENAI_DEPLOYMENT", ""),
        ("DEPLOYMENT_TYPE", "invalid-mode"),
    ],
)
def test_invalid_environment(
    settings: Settings, monkeypatch: pytest.MonkeyPatch, name: str, value: str
) -> None:
    monkeypatch.setenv(name, value)
    with pytest.raises(ValueError):
        Settings.from_env()


def test_secret_auth_does_not_require_identity_id(
    settings: Settings, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setenv("SHAREPOINT_APPLICATION_SECRET", "secret-for-test")
    monkeypatch.delenv("SEARCH_IDENTITY_CLIENT_ID")
    config = Settings.from_env()
    assert config.application_secret == "secret-for-test" and config.identity_client_id == ""


def test_custom_library_names_and_schedule(
    settings: Settings, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setenv("SHAREPOINT_LIBRARY_URL", settings.site_url + "/Shared Documents")
    monkeypatch.setenv("SEARCH_DATASOURCE_NAME", "custom-source")
    monkeypatch.setenv("INDEXER_SCHEDULE", "")
    resources = build_resources(Settings.from_env())
    assert resources["datasources"]["name"] == "custom-source-all-embeddings"
    assert resources["datasources"]["container"] == {
        "name": "useQuery",
        "query": f"includeLibrary={settings.site_url}/Shared Documents",
    }
    assert resources["indexers"]["schedule"] is None


def test_dry_run_is_offline_and_redacted(
    settings: Settings, monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str]
) -> None:
    secrets = {
        "SEARCH_ADMIN_KEY": "search-key-sentinel",
        "AZURE_OPENAI_KEY": 'openai-key-"sentinel"',
        "AZURE_AI_SERVICES_KEY": "vision-key-sentinel",
        "SHAREPOINT_APPLICATION_SECRET": "sharepoint-secret-sentinel",
    }
    for name, value in secrets.items():
        monkeypatch.setenv(name, value)
    with patch("provision.ManagedIdentityCredential") as identity, patch("provision.httpx.Client"):
        assert main(["--dry-run"]) == 0
    identity.assert_not_called()
    output = capsys.readouterr().out
    assert "<redacted>" in output
    for secret in secrets.values():
        assert secret not in output
        assert json.dumps(secret)[1:-1] not in output
    assert set(json.loads(output)) == {"datasources", "indexes", "skillsets", "indexers"}


@pytest.mark.parametrize("code", [400, 401, 403, 409, 429, 500])
def test_failure_stops_provisioning_and_redacts_response(settings: Settings, code: int) -> None:
    config = replace(settings, search_key="api-key-sentinel")
    calls = []

    def handler(request: httpx.Request) -> httpx.Response:
        calls.append(request)
        if len(calls) == 1:
            return httpx.Response(201)
        return httpx.Response(code, text="Rejected api-key-sentinel")

    with pytest.raises(RuntimeError, match=f"HTTP {code}: Rejected <redacted>"):
        send_requests(config, build_resources(config), transport=httpx.MockTransport(handler))
    assert len(calls) == 2


def test_status_only(settings: Settings, capsys: pytest.CaptureFixture[str]) -> None:
    config = replace(settings, search_key="search-key-sentinel")
    calls = []

    def handler(request: httpx.Request) -> httpx.Response:
        calls.append(request)
        assert request.method == "GET"
        assert request.url.path == f"/indexers/{config.names['indexers']}/status"
        return httpx.Response(200, json={"lastResult": {"status": "transientFailure"}})

    send_requests(
        config, build_resources(config), status_only=True, transport=httpx.MockTransport(handler)
    )
    assert len(calls) == 1
    assert json.loads(capsys.readouterr().out)["lastResult"]["status"] == "transientFailure"


def test_main_reports_error(settings: Settings, capsys: pytest.CaptureFixture[str]) -> None:
    with patch("provision.send_requests", side_effect=RuntimeError("HTTP 403: denied")):
        assert main([]) == 1
    assert "HTTP 403: denied" in capsys.readouterr().err


@pytest.mark.parametrize("deployment_type", list(DeploymentType))
def test_exact_resources_for_each_type(settings: Settings, deployment_type: DeploymentType) -> None:
    config = Settings.from_env(deployment_type)
    resources = build_resources(config)
    index = resources["indexes"]
    fields = {item["name"]: item for item in index["fields"]}
    indexer = resources["indexers"]
    assert ("text_vector" in fields) is config.use_text
    assert ("image_vector" in fields) is config.use_images
    assert ("image_parent_id" in fields) is config.use_images
    assert indexer["parameters"]["configuration"]["imageAction"] == (
        "generateNormalizedImages" if config.use_images else "none"
    )
    if deployment_type == DeploymentType.NO_EMBEDDING:
        assert set(fields) == {"chunk_id", "chunk", "title", "metadata_spo_item_path"}
        assert "skillsets" not in resources
        assert "skillsetName" not in indexer
        assert "vectorSearch" not in index
        mappings = {item["targetFieldName"]: item for item in indexer["fieldMappings"]}
        assert mappings["chunk"]["sourceFieldName"] == "content"
        assert mappings["chunk_id"] == {
            "sourceFieldName": "metadata_spo_site_library_item_id",
            "targetFieldName": "chunk_id",
            "mappingFunction": {"name": "base64Encode"},
        }
        assert "resourceUri" not in json.dumps(resources)
        return

    skillset = resources["skillsets"]
    skills = {item["name"]: item for item in skillset["skills"]}
    assert set(skills) == (
        {"split"}
        | ({"embed-text"} if config.use_text else set())
        | ({"ocr", "merge", "embed-image"} if config.use_images else set())
    )
    assert ("cognitiveServices" in skillset) is config.use_images
    assert skills["split"]["inputs"][0]["source"] == (
        "/document/mergedText" if config.use_images else "/document/content"
    )
    vectorizers = {item["kind"]: item for item in index["vectorSearch"]["vectorizers"]}
    assert set(vectorizers) == (
        ({"azureOpenAI"} if config.use_text else set())
        | ({"aiServicesVision"} if config.use_images else set())
    )
    selectors = skillset["indexProjections"]["selectors"]
    assert len(selectors) == (2 if config.use_images else 1)
    text_mappings = {item["name"] for item in selectors[0]["mappings"]}
    assert ("text_vector" in text_mappings) is config.use_text
    assert "chunk" in text_mappings
    for selector in selectors:
        assert selector["parentKeyFieldName"] in fields
        assert {item["name"] for item in selector["mappings"]} <= fields.keys()


@pytest.mark.parametrize("deployment_type", list(DeploymentType))
def test_only_selected_service_settings_are_required(
    settings: Settings, monkeypatch: pytest.MonkeyPatch, deployment_type: DeploymentType
) -> None:
    for name in ("AZURE_OPENAI_ENDPOINT", "AZURE_OPENAI_DEPLOYMENT", "AZURE_AI_SERVICES_ENDPOINT"):
        monkeypatch.delenv(name)
    use_text = deployment_type in (DeploymentType.TEXT_EMBEDDINGS, DeploymentType.ALL_EMBEDDINGS)
    use_images = deployment_type in (DeploymentType.IMAGE_EMBEDDINGS, DeploymentType.ALL_EMBEDDINGS)
    if use_text:
        with pytest.raises(ValueError, match="AZURE_OPENAI_ENDPOINT"):
            Settings.from_env(deployment_type)
        monkeypatch.setenv("AZURE_OPENAI_ENDPOINT", settings.openai_endpoint)
        monkeypatch.setenv("AZURE_OPENAI_DEPLOYMENT", settings.openai_deployment)
    if use_images:
        with pytest.raises(ValueError, match="AZURE_AI_SERVICES_ENDPOINT"):
            Settings.from_env(deployment_type)
        monkeypatch.setenv("AZURE_AI_SERVICES_ENDPOINT", settings.ai_services_endpoint)
    config = Settings.from_env(deployment_type)
    assert bool(config.openai_endpoint) is use_text
    assert bool(config.ai_services_endpoint) is use_images


def test_default_and_cli_override(
    settings: Settings, monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str]
) -> None:
    monkeypatch.delenv("DEPLOYMENT_TYPE")
    assert Settings.from_env().deployment_type == DeploymentType.NO_EMBEDDING
    monkeypatch.setenv("DEPLOYMENT_TYPE", "all_embeddings")
    for name in ("AZURE_OPENAI_ENDPOINT", "AZURE_OPENAI_DEPLOYMENT", "AZURE_AI_SERVICES_ENDPOINT"):
        monkeypatch.setenv(name, "unused-invalid-value")
    assert main(["--type", "no_embedding", "--dry-run"]) == 0
    resources = json.loads(capsys.readouterr().out)
    assert set(resources) == {"datasources", "indexes", "indexers"}
    assert resources["indexes"]["name"] == "test-rag-no-embedding"


def test_type_names_are_stable_and_isolated(
    settings: Settings, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setenv("SEARCH_DATASOURCE_NAME", "my-source")
    monkeypatch.setenv("SEARCH_SKILLSET_NAME", "my-skills")
    monkeypatch.setenv("SEARCH_INDEXER_NAME", "my-indexer")
    configurations = [Settings.from_env(item) for item in DeploymentType]
    for collection in ("indexes", "datasources", "skillsets", "indexers"):
        assert len({config.names[collection] for config in configurations}) == 4
    for config in configurations:
        assert config.names == Settings.from_env(config.deployment_type).names
        suffix = config.deployment_type.replace("_", "-")
        assert config.names["indexes"] == f"test-rag-{suffix}"
        assert config.names["datasources"] == f"my-source-{suffix}"
        assert config.names["skillsets"] == f"my-skills-{suffix}"
        assert config.names["indexers"] == f"my-indexer-{suffix}"


def test_repeated_deployments_reuse_resources_and_request_incremental_run(
    settings: Settings,
) -> None:
    objects: dict[str, bytes] = {}
    calls: list[tuple[str, str]] = []

    def handler(request: httpx.Request) -> httpx.Response:
        calls.append((request.method, request.url.path))
        if request.method == "POST":
            assert request.url.path.endswith("/run")
            return httpx.Response(202)
        assert request.method == "PUT"
        existed = request.url.path in objects
        objects[request.url.path] = request.content
        return httpx.Response(200 if existed else 201)

    expected_count = 0
    for deployment_type in DeploymentType:
        config = replace(Settings.from_env(deployment_type), search_key="test-key")
        resources = build_resources(config)
        transport = httpx.MockTransport(handler)
        start = len(calls)
        send_requests(config, resources, transport=transport)
        assert all(method == "PUT" for method, _ in calls[start:])
        start = len(calls)
        send_requests(config, resources, transport=transport)
        assert len(calls[start:]) == len(resources) + 1
        assert calls[-1] == ("POST", f"/indexers/{config.names['indexers']}/run")
        expected_count += len(resources)
        assert len(objects) == expected_count
    assert not any("reset" in path for _, path in calls)


@pytest.mark.parametrize("code", [409, 429, 500])
def test_rerun_errors_are_not_silently_ignored(settings: Settings, code: int) -> None:
    config = replace(Settings.from_env("no_embedding"), search_key="test-key")

    def handler(request: httpx.Request) -> httpx.Response:
        return httpx.Response(code if request.method == "POST" else 200)

    with pytest.raises(RuntimeError, match=f"HTTP {code}"):
        send_requests(config, build_resources(config), transport=httpx.MockTransport(handler))


@pytest.mark.parametrize("suffix", ["", "/", "///"])
@pytest.mark.parametrize("deployment_type", ["text_embeddings", "all_embeddings"])
def test_openai_endpoint_has_one_trailing_slash(
    settings: Settings,
    monkeypatch: pytest.MonkeyPatch,
    suffix: str,
    deployment_type: str,
) -> None:
    monkeypatch.setenv("AZURE_OPENAI_ENDPOINT", "https://example.openai.azure.com" + suffix)
    config = Settings.from_env(deployment_type)
    expected = "https://example.openai.azure.com/"
    assert config.openai_endpoint == expected
    assert config.search_endpoint == settings.search_endpoint
    resources = build_resources(config)
    embedding = next(
        skill for skill in resources["skillsets"]["skills"] if skill["name"] == "embed-text"
    )
    vectorizer = next(
        item
        for item in resources["indexes"]["vectorSearch"]["vectorizers"]
        if item["kind"] == "azureOpenAI"
    )
    assert embedding["resourceUri"] == expected
    assert vectorizer["azureOpenAIParameters"]["resourceUri"] == expected
