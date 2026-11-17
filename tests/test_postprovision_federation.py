import json
import os
import subprocess
from collections.abc import Iterator
from unittest.mock import MagicMock, patch

import httpx
import pytest

from postprovision_federation import (
    configure_sharepoint_federation,
    verify_sharepoint_federation,
)

TENANT = "22222222-2222-2222-2222-222222222222"
APPLICATION = "11111111-1111-1111-1111-111111111111"
PRINCIPAL = "33333333-3333-3333-3333-333333333333"
CLIENT_ID = "44444444-4444-4444-4444-444444444444"
ISSUER = f"https://login.microsoftonline.com/{TENANT}/v2.0"
IDENTITY_PATH = f"/v1.0/servicePrincipals/{PRINCIPAL}"
CREDENTIALS_PATH = f"/v1.0/applications(appId='{APPLICATION}')/federatedIdentityCredentials"


@pytest.fixture(autouse=True)
def isolate_hook_environment(monkeypatch: pytest.MonkeyPatch) -> None:
    for name in ("CONFIGURE_SHAREPOINT_FEDERATION", "SEARCH_IDENTITY_PRINCIPAL_ID"):
        monkeypatch.delenv(name, raising=False)


@pytest.fixture
def federation_environment(monkeypatch: pytest.MonkeyPatch) -> Iterator[MagicMock]:
    for name in (
        "SHAREPOINT_APPLICATION_SECRET",
        "SEARCH_DATASOURCE_NAME",
        "SEARCH_SKILLSET_NAME",
        "SEARCH_INDEXER_NAME",
        "SHAREPOINT_LIBRARY_URL",
        "INDEXER_SCHEDULE",
    ):
        monkeypatch.delenv(name, raising=False)
    for name, value in {
        "AZURE_ENV_NAME": "qa",
        "AZURE_TENANT_ID": TENANT,
        "SHAREPOINT_TENANT_ID": TENANT,
        "SHAREPOINT_APPLICATION_ID": APPLICATION,
        "SEARCH_IDENTITY_PRINCIPAL_ID": PRINCIPAL,
        "SEARCH_IDENTITY_CLIENT_ID": "55555555-5555-5555-5555-555555555555",
        "AZURE_SEARCH_SERVICE_NAME": "test-search-service",
        "DEPLOYMENT_TYPE": "no_embedding",
        "SEARCH_ENDPOINT": "https://example.search.windows.net",
        "SHAREPOINT_SITE_URL": "https://example.sharepoint.com/sites/Hub",
        "SEARCH_INDEX_NAME": "test-rag",
    }.items():
        monkeypatch.setenv(name, value)
    with (
        patch("postprovision_federation.AzureDeveloperCliCredential") as factory,
        patch("postprovision_federation.subprocess.run") as save,
    ):
        factory.return_value.__enter__.return_value.get_token.return_value.token = "test-token"
        yield save


def graph_transport(credentials: list[dict], requests: list[httpx.Request]) -> httpx.MockTransport:
    def handle(request: httpx.Request) -> httpx.Response:
        requests.append(request)
        assert request.headers["Authorization"] == "Bearer test-token"
        if request.method == "GET" and request.url.path == IDENTITY_PATH:
            return httpx.Response(
                200,
                json={
                    "id": PRINCIPAL,
                    "appId": CLIENT_ID,
                    "servicePrincipalType": "ManagedIdentity",
                },
            )
        if request.url.path == CREDENTIALS_PATH:
            if request.method == "GET":
                return httpx.Response(200, json={"value": credentials})
            if request.method == "POST":
                return httpx.Response(201, json={"id": "new-credential"})
        if request.method == "PATCH":
            for credential in credentials:
                if request.url.path == f"{CREDENTIALS_PATH}/{credential.get('id')}":
                    credential.update(json.loads(request.content))
                    return httpx.Response(204)
        pytest.fail(f"Unexpected request: {request.method} {request.url}")

    return httpx.MockTransport(handle)


@pytest.fixture
def stale_credential() -> dict:
    return {
        "id": "66666666-6666-6666-6666-666666666666",
        "name": "test-search-service",
        "issuer": ISSUER,
        "subject": "55555555-5555-5555-5555-555555555555",
        "audiences": ["api://AzureADTokenExchange"],
        "description": "Search identity for azd environment qa.",
    }


@pytest.mark.parametrize("at_credential_limit", [False, True])
def test_updates_recreated_search_trust_in_place_and_reuses_it_on_retry(
    federation_environment: MagicMock, stale_credential: dict, at_credential_limit: bool
) -> None:
    credentials = [stale_credential]
    if at_credential_limit:
        credentials.extend({"name": f"unrelated-{i}"} for i in range(19))
    before = [dict(item) for item in credentials]
    requests: list[httpx.Request] = []
    transport = graph_transport(credentials, requests)

    configure_sharepoint_federation(transport=transport)

    assert [(request.method, request.url.path) for request in requests] == [
        ("GET", IDENTITY_PATH),
        ("GET", CREDENTIALS_PATH),
        ("PATCH", f"{CREDENTIALS_PATH}/{stale_credential['id']}"),
    ]
    assert json.loads(requests[-1].content) == {"subject": PRINCIPAL}
    assert credentials == [{**before[0], "subject": PRINCIPAL}, *before[1:]]
    assert os.environ["SEARCH_IDENTITY_CLIENT_ID"] == CLIENT_ID
    federation_environment.assert_called_once()

    requests.clear()
    federation_environment.reset_mock()
    configure_sharepoint_federation(transport=transport)
    assert [request.method for request in requests] == ["GET", "GET"]
    federation_environment.assert_called_once()


@pytest.mark.parametrize(
    ("overrides", "message"),
    [
        ({"issuer": "https://another-issuer.example"}, "name conflicts"),
        ({"audiences": ["unexpected-audience"]}, "name conflicts"),
        ({"description": "Search identity for azd environment prd."}, "name conflicts"),
        ({"description": None}, "name conflicts"),
        ({"id": None}, "no valid ID"),
        ({"id": "not-a-guid"}, "no valid ID"),
    ],
)
def test_stale_trust_update_rejects_unrelated_or_malformed_credentials(
    federation_environment: MagicMock, stale_credential: dict, overrides: dict, message: str
) -> None:
    requests: list[httpx.Request] = []
    with pytest.raises(ValueError, match=message):
        configure_sharepoint_federation(
            transport=graph_transport([{**stale_credential, **overrides}], requests)
        )
    assert [request.method for request in requests] == ["GET", "GET"]
    federation_environment.assert_not_called()


@pytest.mark.parametrize("status", [403, 404, 409])
def test_failed_trust_update_does_not_delete_create_or_save(
    federation_environment: MagicMock,
    stale_credential: dict,
    capsys: pytest.CaptureFixture[str],
    status: int,
) -> None:
    requests: list[httpx.Request] = []
    transport = graph_transport([stale_credential], requests)

    def handle(request: httpx.Request) -> httpx.Response:
        if request.method == "PATCH":
            requests.append(request)
            return httpx.Response(status, json={"error": {"message": "Update rejected"}})
        return transport.handle_request(request)

    with patch("postprovision_federation.time.sleep") as sleep:
        with pytest.raises(RuntimeError, match=f"HTTP {status}.*Update rejected"):
            configure_sharepoint_federation(transport=httpx.MockTransport(handle))
    assert [request.method for request in requests] == ["GET", "GET", "PATCH"]
    sleep.assert_not_called()
    federation_environment.assert_not_called()
    assert "Updated SharePoint" not in capsys.readouterr().out


def test_verification_does_not_update_stale_managed_trust(
    federation_environment: MagicMock, stale_credential: dict
) -> None:
    requests: list[httpx.Request] = []
    with pytest.raises(ValueError, match="No matching"):
        verify_sharepoint_federation(transport=graph_transport([stale_credential], requests))
    assert [request.method for request in requests] == ["GET", "GET"]
    federation_environment.assert_not_called()


@pytest.mark.parametrize("search_name", ["test-search-service", "another-search-service"])
def test_adds_new_trust_without_replacing_old_credential(
    monkeypatch: pytest.MonkeyPatch, federation_environment: MagicMock, search_name: str
) -> None:
    monkeypatch.setenv("AZURE_SEARCH_SERVICE_NAME", search_name)
    old_credential = {"name": "aisearch", "subject": "old-principal", "issuer": ISSUER}
    requests: list[httpx.Request] = []
    configure_sharepoint_federation(transport=graph_transport([old_credential], requests))
    assert [(request.method, request.url.path) for request in requests] == [
        ("GET", IDENTITY_PATH),
        ("GET", CREDENTIALS_PATH),
        ("POST", CREDENTIALS_PATH),
    ]
    body = json.loads(requests[-1].content)
    assert body["subject"] == PRINCIPAL
    assert body["issuer"] == ISSUER
    assert body["audiences"] == ["api://AzureADTokenExchange"]
    assert body["name"] == search_name
    assert os.environ["SEARCH_IDENTITY_CLIENT_ID"] == CLIENT_ID
    federation_environment.assert_called_once_with(
        [
            "azd",
            "env",
            "set",
            "SEARCH_IDENTITY_CLIENT_ID",
            CLIENT_ID,
            "--environment",
            "qa",
        ],
        check=True,
        capture_output=True,
        text=True,
        timeout=60,
    )


def test_reuses_matching_trust_with_any_name(federation_environment: MagicMock) -> None:
    credentials = [
        {
            "name": "manually-created",
            "issuer": ISSUER,
            "subject": PRINCIPAL,
            "audiences": ["api://AzureADTokenExchange"],
        }
    ]
    requests: list[httpx.Request] = []
    configure_sharepoint_federation(transport=graph_transport(credentials, requests))
    assert all(request.method == "GET" for request in requests)
    federation_environment.assert_called_once()


def test_federation_rejects_secret_auth_before_graph_access(
    monkeypatch: pytest.MonkeyPatch, federation_environment: MagicMock
) -> None:
    monkeypatch.setenv("SHAREPOINT_APPLICATION_SECRET", "test-secret")
    requests: list[httpx.Request] = []
    with pytest.raises(ValueError, match="requires SHAREPOINT_APPLICATION_SECRET to be unset"):
        configure_sharepoint_federation(transport=graph_transport([], requests))
    assert not requests
    federation_environment.assert_not_called()


def test_federation_does_not_require_pipeline_settings(
    monkeypatch: pytest.MonkeyPatch, federation_environment: MagicMock
) -> None:
    for name in ("SHAREPOINT_SITE_URL", "SEARCH_ENDPOINT", "AZURE_OPENAI_ENDPOINT"):
        monkeypatch.delenv(name, raising=False)
    monkeypatch.setenv("DEPLOYMENT_TYPE", "not-a-pipeline-mode")
    requests: list[httpx.Request] = []
    configure_sharepoint_federation(transport=graph_transport([], requests))
    assert [request.method for request in requests] == ["GET", "GET", "POST"]
    federation_environment.assert_called_once()


def test_credential_limit_does_not_remove_existing_trusts(
    federation_environment: MagicMock,
) -> None:
    requests: list[httpx.Request] = []
    credentials = [{"name": f"existing-{i}"} for i in range(20)]
    with pytest.raises(ValueError, match="20 federated credential limit"):
        configure_sharepoint_federation(transport=graph_transport(credentials, requests))
    assert all(request.method == "GET" for request in requests)
    federation_environment.assert_not_called()


def test_client_id_save_failure_is_not_reported_as_success(
    federation_environment: MagicMock,
) -> None:
    federation_environment.side_effect = subprocess.CalledProcessError(
        1, "azd", stderr="environment save failed"
    )
    with pytest.raises(RuntimeError, match="environment save failed"):
        configure_sharepoint_federation(transport=graph_transport([], []))


@pytest.mark.parametrize(
    ("variable", "value", "message"),
    [
        ("SEARCH_IDENTITY_PRINCIPAL_ID", "not-a-guid", "must be a GUID"),
        ("SHAREPOINT_TENANT_ID", CLIENT_ID, "same tenant"),
        ("AZURE_ENV_NAME", "", "Set AZURE_ENV_NAME"),
        ("AZURE_SEARCH_SERVICE_NAME", "", "Set AZURE_SEARCH_SERVICE_NAME"),
    ],
)
def test_invalid_federation_inputs_prevent_graph_access(
    monkeypatch: pytest.MonkeyPatch,
    federation_environment: MagicMock,
    variable: str,
    value: str,
    message: str,
) -> None:
    monkeypatch.setenv(variable, value)
    requests: list[httpx.Request] = []
    with pytest.raises(ValueError, match=message):
        configure_sharepoint_federation(transport=graph_transport([], requests))
    assert not requests
    federation_environment.assert_not_called()


@pytest.mark.parametrize(
    ("overrides", "message"),
    [
        ({"servicePrincipalType": "Application"}, "expected managed identity"),
        ({"id": APPLICATION}, "expected managed identity"),
        ({"appId": None}, "no valid application/client ID"),
    ],
)
def test_unexpected_identity_prevents_app_mutation(
    federation_environment: MagicMock, overrides: dict, message: str
) -> None:
    requests: list[httpx.Request] = []

    def handle(request: httpx.Request) -> httpx.Response:
        requests.append(request)
        return httpx.Response(
            200,
            json={
                "id": PRINCIPAL,
                "appId": CLIENT_ID,
                "servicePrincipalType": "ManagedIdentity",
                **overrides,
            },
        )

    with pytest.raises(ValueError, match=message):
        configure_sharepoint_federation(transport=httpx.MockTransport(handle))
    assert len(requests) == 1
    federation_environment.assert_not_called()


@pytest.mark.parametrize("operation", ["identity", "list", "create"])
def test_graph_permission_errors_stop_setup_without_retry(
    federation_environment: MagicMock, operation: str
) -> None:
    requests: list[httpx.Request] = []
    transport = graph_transport([], requests)

    def handle(request: httpx.Request) -> httpx.Response:
        current = (
            "identity"
            if request.url.path == IDENTITY_PATH
            else "list"
            if request.method == "GET"
            else "create"
        )
        if current == operation:
            requests.append(request)
            return httpx.Response(403, json={"error": {"message": "Permission denied"}})
        return transport.handle_request(request)

    with patch("postprovision_federation.time.sleep") as sleep:
        with pytest.raises(RuntimeError, match="HTTP 403.*Permission denied"):
            configure_sharepoint_federation(transport=httpx.MockTransport(handle))
    sleep.assert_not_called()
    federation_environment.assert_not_called()


@pytest.mark.parametrize("eventually_visible", [True, False])
def test_new_identity_directory_replication_retry_is_bounded(
    federation_environment: MagicMock, eventually_visible: bool
) -> None:
    attempts = 0
    transport = graph_transport([], [])

    def handle(request: httpx.Request) -> httpx.Response:
        nonlocal attempts
        if request.url.path == IDENTITY_PATH:
            attempts += 1
            if not eventually_visible or attempts == 1:
                return httpx.Response(404, json={"error": {"message": "Not yet visible"}})
        return transport.handle_request(request)

    with patch("postprovision_federation.time.sleep") as sleep:
        if eventually_visible:
            configure_sharepoint_federation(transport=httpx.MockTransport(handle))
            assert attempts == 2
            sleep.assert_called_once_with(30)
            federation_environment.assert_called_once()
        else:
            with pytest.raises(RuntimeError, match="HTTP 404"):
                configure_sharepoint_federation(transport=httpx.MockTransport(handle))
            assert attempts == 11
            assert sleep.call_count == 10
            federation_environment.assert_not_called()


@pytest.mark.parametrize("conflict", ["audience", "name"])
def test_existing_conflicting_trust_is_not_replaced(
    federation_environment: MagicMock, conflict: str
) -> None:
    credential = (
        {"issuer": ISSUER, "subject": PRINCIPAL, "audiences": ["unexpected-audience"]}
        if conflict == "audience"
        else {"name": "test-search-service", "issuer": ISSUER, "subject": "another-identity"}
    )
    requests: list[httpx.Request] = []
    with pytest.raises(ValueError, match="incompatible audience|name conflicts"):
        configure_sharepoint_federation(transport=graph_transport([credential], requests))
    assert all(request.method == "GET" for request in requests)
    federation_environment.assert_not_called()


@pytest.mark.parametrize(
    "next_link",
    [
        "https://graph.microsoft.com/v1.0/next-page",
        "https://untrusted.example/v1.0/next-page",
        f"https://graph.microsoft.com{CREDENTIALS_PATH}",
    ],
)
def test_pagination_reuses_matching_trust_and_rejects_unsafe_links(
    federation_environment: MagicMock, next_link: str
) -> None:
    transport = graph_transport([], [])
    requests: list[httpx.Request] = []

    def handle(request: httpx.Request) -> httpx.Response:
        requests.append(request)
        if request.url.path == IDENTITY_PATH:
            return transport.handle_request(request)
        if request.url.path == CREDENTIALS_PATH:
            return httpx.Response(200, json={"value": [], "@odata.nextLink": next_link})
        assert str(request.url) == "https://graph.microsoft.com/v1.0/next-page"
        return httpx.Response(
            200,
            json={
                "value": [
                    {
                        "issuer": ISSUER,
                        "subject": PRINCIPAL,
                        "audiences": ["api://AzureADTokenExchange"],
                    }
                ]
            },
        )

    if next_link.endswith("/next-page") and "untrusted" not in next_link:
        configure_sharepoint_federation(transport=httpx.MockTransport(handle))
        assert len(requests) == 3
        federation_environment.assert_called_once()
    else:
        with pytest.raises(ValueError, match="invalid federation pagination URL"):
            configure_sharepoint_federation(transport=httpx.MockTransport(handle))
        assert len(requests) == 2
        federation_environment.assert_not_called()


@pytest.mark.parametrize("response", [{}, {"value": None}, {"value": [None]}, []])
def test_malformed_graph_list_prevents_app_mutation(
    federation_environment: MagicMock, response: object
) -> None:
    transport = graph_transport([], [])
    methods: list[str] = []

    def handle(request: httpx.Request) -> httpx.Response:
        methods.append(request.method)
        if request.url.path == IDENTITY_PATH:
            return transport.handle_request(request)
        return httpx.Response(200, json=response)

    with pytest.raises(ValueError, match="invalid"):
        configure_sharepoint_federation(transport=httpx.MockTransport(handle))
    assert methods == ["GET", "GET"]
    federation_environment.assert_not_called()


def test_verification_discovers_client_id_without_granting_or_saving(
    monkeypatch: pytest.MonkeyPatch, federation_environment: MagicMock
) -> None:
    for name in ("AZURE_ENV_NAME", "AZURE_SEARCH_SERVICE_NAME"):
        monkeypatch.delenv(name, raising=False)
    stale_client_id = "55555555-5555-5555-5555-555555555555"
    monkeypatch.setenv("SEARCH_IDENTITY_CLIENT_ID", stale_client_id)
    requests: list[httpx.Request] = []
    credentials = [
        {"issuer": ISSUER, "subject": PRINCIPAL, "audiences": ["api://AzureADTokenExchange"]}
    ]
    assert (
        verify_sharepoint_federation(transport=graph_transport(credentials, requests)) == CLIENT_ID
    )
    assert os.environ["SEARCH_IDENTITY_CLIENT_ID"] == stale_client_id
    assert [request.method for request in requests] == ["GET", "GET"]
    federation_environment.assert_not_called()


@pytest.mark.parametrize(
    "credentials",
    [
        [],
        [{"issuer": ISSUER, "subject": "previous-identity"}],
        [{"issuer": "https://another-issuer.example", "subject": PRINCIPAL}],
        [{"issuer": ISSUER, "subject": PRINCIPAL, "audiences": ["wrong-audience"]}],
    ],
)
def test_verification_rejects_missing_or_incorrect_trust_without_writes(
    federation_environment: MagicMock, credentials: list[dict]
) -> None:
    requests: list[httpx.Request] = []
    with pytest.raises(ValueError, match="No matching|incompatible audience"):
        verify_sharepoint_federation(transport=graph_transport(credentials, requests))
    assert [request.method for request in requests] == ["GET", "GET"]
    federation_environment.assert_not_called()
