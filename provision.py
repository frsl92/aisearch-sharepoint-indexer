"""Create a SharePoint indexing pipeline with optional text/image embeddings."""

import argparse
import json
import os
import re
import sys
from contextlib import ExitStack
from dataclasses import dataclass, field
from enum import StrEnum
from typing import Any
from urllib.parse import unquote, urlsplit
from uuid import UUID

import httpx
from azure.core.exceptions import AzureError
from azure.identity import ManagedIdentityCredential

API_VERSION = "2026-08-01-preview"
TEXT_MODEL = "text-embedding-3-large"
VISION_MODEL = "2023-04-15"


class DeploymentType(StrEnum):
    NO_EMBEDDING = "no_embedding"
    IMAGE_EMBEDDINGS = "image_embeddings"
    TEXT_EMBEDDINGS = "text_embeddings"
    ALL_EMBEDDINGS = "all_embeddings"


def env(name: str, default: str = "") -> str:
    return os.environ.get(name, default).strip()


def required(name: str) -> str:
    value = env(name)
    if not value:
        raise ValueError(f"Set {name}.")
    return value


def endpoint(name: str, suffixes: tuple[str, ...], *, site: bool = False) -> str:
    value = required(name).rstrip("/")
    parts = urlsplit(value)
    if (
        parts.scheme != "https"
        or not parts.hostname
        or not any(parts.hostname.endswith("." + suffix) for suffix in suffixes)
        or parts.username is not None
        or parts.password is not None
        or parts.port is not None
        or parts.query
        or parts.fragment
        or ";" in value
        or any(char.isspace() and char != " " for char in value)
        or (not site and parts.path)
    ):
        raise ValueError(f"{name} must be an HTTPS resource URL under {', '.join(suffixes)}.")
    return value


@dataclass(frozen=True)
class Settings:
    search_endpoint: str
    names: dict[str, str]
    deployment_type: DeploymentType
    site_url: str
    application_id: str
    tenant_id: str
    identity_client_id: str
    library_url: str
    openai_endpoint: str
    openai_deployment: str
    ai_services_endpoint: str
    schedule: str
    search_key: str = field(repr=False)
    application_secret: str = field(repr=False)
    openai_key: str = field(repr=False)
    ai_services_key: str = field(repr=False)

    @classmethod
    def from_env(cls, deployment_type: str | None = None) -> "Settings":
        try:
            selected_type = DeploymentType(
                deployment_type or env("DEPLOYMENT_TYPE", "no_embedding")
            )
        except ValueError as exc:
            raise ValueError(
                "DEPLOYMENT_TYPE must be one of: " + ", ".join(DeploymentType)
            ) from exc
        use_text = selected_type in (DeploymentType.TEXT_EMBEDDINGS, DeploymentType.ALL_EMBEDDINGS)
        use_images = selected_type in (
            DeploymentType.IMAGE_EMBEDDINGS,
            DeploymentType.ALL_EMBEDDINGS,
        )
        suffix = selected_type.replace("_", "-")
        index = f"{env('SEARCH_INDEX_NAME', 'sharepoint-rag')}-{suffix}"
        names = {
            "indexes": index,
            **{
                collection: (f"{env(variable)}-{suffix}" if env(variable) else f"{index}-{label}")
                for collection, variable, label in (
                    ("datasources", "SEARCH_DATASOURCE_NAME", "datasource"),
                    ("skillsets", "SEARCH_SKILLSET_NAME", "skillset"),
                    ("indexers", "SEARCH_INDEXER_NAME", "indexer"),
                )
            },
        }
        for collection, name in names.items():
            if (
                not re.fullmatch(r"[a-z0-9][a-z0-9-]*[a-z0-9]", name)
                or "--" in name
                or len(name) > 128
            ):
                raise ValueError(
                    f"Invalid {collection} name: use 2-128 lowercase letters/digits/dashes."
                )
        ids = {}
        for variable in ("SHAREPOINT_APPLICATION_ID", "SHAREPOINT_TENANT_ID"):
            try:
                ids[variable] = str(UUID(required(variable)))
            except ValueError as exc:
                raise ValueError(f"{variable} must be a GUID.") from exc
        secret = env("SHAREPOINT_APPLICATION_SECRET")
        identity_id = ""
        if not secret:
            try:
                identity_id = str(UUID(required("SEARCH_IDENTITY_CLIENT_ID")))
            except ValueError as exc:
                raise ValueError(
                    "SEARCH_IDENTITY_CLIENT_ID must be the Search identity client GUID."
                ) from exc
        if any(char in secret for char in ";\r\n"):
            raise ValueError("SHAREPOINT_APPLICATION_SECRET cannot contain semicolons or newlines.")
        site_url = endpoint("SHAREPOINT_SITE_URL", ("sharepoint.com",), site=True)
        library_url = env("SHAREPOINT_LIBRARY_URL")
        if library_url:
            library_url = endpoint("SHAREPOINT_LIBRARY_URL", ("sharepoint.com",), site=True)
            site_parts, library_parts = urlsplit(site_url), urlsplit(library_url)
            path = unquote(library_parts.path)
            if (
                site_parts.hostname != library_parts.hostname
                or not path.startswith(unquote(site_parts.path) + "/")
                or any(segment in (".", "..") for segment in path.split("/"))
            ):
                raise ValueError("SHAREPOINT_LIBRARY_URL must be inside SHAREPOINT_SITE_URL.")
        schedule = env("INDEXER_SCHEDULE", "PT1H")
        if schedule:
            match = re.fullmatch(r"PT(?:(\d+)H)?(?:(\d+)M)?", schedule)
            minutes = int(match[1] or 0) * 60 + int(match[2] or 0) if match else 0
            if not 5 <= minutes <= 1440:
                raise ValueError("INDEXER_SCHEDULE must be 5 minutes to 24 hours, e.g. PT1H.")
        return cls(
            search_endpoint=endpoint("SEARCH_ENDPOINT", ("search.windows.net",)),
            names=names,
            deployment_type=selected_type,
            site_url=site_url,
            application_id=ids["SHAREPOINT_APPLICATION_ID"],
            tenant_id=ids["SHAREPOINT_TENANT_ID"],
            identity_client_id=identity_id,
            library_url=library_url,
            openai_endpoint=endpoint(
                "AZURE_OPENAI_ENDPOINT",
                ("openai.azure.com", "services.ai.azure.com", "cognitiveservices.azure.com"),
            )
            + "/"
            if use_text
            else "",
            openai_deployment=required("AZURE_OPENAI_DEPLOYMENT") if use_text else "",
            ai_services_endpoint=endpoint(
                "AZURE_AI_SERVICES_ENDPOINT",
                ("cognitiveservices.azure.com", "services.ai.azure.com"),
            )
            if use_images
            else "",
            schedule=schedule,
            search_key=env("SEARCH_ADMIN_KEY"),
            application_secret=secret,
            openai_key=env("AZURE_OPENAI_KEY"),
            ai_services_key=env("AZURE_AI_SERVICES_KEY"),
        )

    @property
    def use_text(self) -> bool:
        return self.deployment_type in (
            DeploymentType.TEXT_EMBEDDINGS,
            DeploymentType.ALL_EMBEDDINGS,
        )

    @property
    def use_images(self) -> bool:
        return self.deployment_type in (
            DeploymentType.IMAGE_EMBEDDINGS,
            DeploymentType.ALL_EMBEDDINGS,
        )

    def hide_secrets(self, text: str) -> str:
        for secret in (
            self.search_key,
            self.application_secret,
            self.openai_key,
            self.ai_services_key,
        ):
            if secret:
                text = text.replace(json.dumps(secret)[1:-1], "<redacted>")
                text = text.replace(secret, "<redacted>")
        return text


def build_resources(settings: Settings) -> dict[str, dict[str, Any]]:
    names = settings.names
    connection = (
        f"SharePointOnlineEndpoint={settings.site_url};"
        f"ApplicationId={settings.application_id};TenantId={settings.tenant_id};"
    )
    connection += (
        f"ApplicationSecret={settings.application_secret}"
        if settings.application_secret
        else f"FederatedCredentialApplicationId={settings.identity_client_id}"
    )
    source = {
        "name": names["datasources"],
        "type": "sharepoint",
        "credentials": {"connectionString": connection},
        "container": (
            {"name": "useQuery", "query": f"includeLibrary={settings.library_url}"}
            if settings.library_url
            else {"name": "defaultSiteLibrary"}
        ),
    }
    openai = {
        "resourceUri": settings.openai_endpoint,
        "deploymentId": settings.openai_deployment,
        "modelName": TEXT_MODEL,
        **({"apiKey": settings.openai_key} if settings.openai_key else {}),
    }
    vision = {
        "resourceUri": settings.ai_services_endpoint,
        "modelVersion": VISION_MODEL,
        **({"apiKey": settings.ai_services_key} if settings.ai_services_key else {}),
    }
    index: dict[str, Any] = {
        "name": names["indexes"],
        "fields": [
            {
                "name": "chunk_id",
                "type": "Edm.String",
                "key": True,
                "searchable": True,
                "sortable": True,
                "analyzer": "keyword",
            },
            {"name": "text_parent_id", "type": "Edm.String", "filterable": True},
            {"name": "chunk", "type": "Edm.String", "searchable": True},
            {"name": "title", "type": "Edm.String", "searchable": True},
            {"name": "image_parent_id", "type": "Edm.String", "filterable": True},
            {
                "name": "text_vector",
                "type": "Collection(Edm.Single)",
                "searchable": True,
                "dimensions": 3072,
                "vectorSearchProfile": "text-profile",
            },
            {
                "name": "image_vector",
                "type": "Collection(Edm.Single)",
                "searchable": True,
                "dimensions": 1024,
                "vectorSearchProfile": "image-profile",
            },
            {"name": "metadata_spo_item_path", "type": "Edm.String", "retrievable": False},
        ],
        "semantic": {
            "defaultConfiguration": "semantic-config",
            "configurations": [
                {
                    "name": "semantic-config",
                    "rankingOrder": "BoostedRerankerScore",
                    "prioritizedFields": {
                        "titleField": {"fieldName": "title"},
                        "prioritizedContentFields": [{"fieldName": "chunk"}],
                        "prioritizedKeywordsFields": [],
                    },
                }
            ],
        },
        "vectorSearch": {
            "algorithms": [
                {
                    "name": "hnsw",
                    "kind": "hnsw",
                    "hnswParameters": {
                        "metric": "cosine",
                        "m": 4,
                        "efConstruction": 400,
                        "efSearch": 500,
                    },
                }
            ],
            "profiles": [
                {"name": "text-profile", "algorithm": "hnsw", "vectorizer": "text-vectorizer"},
                {"name": "image-profile", "algorithm": "hnsw", "vectorizer": "image-vectorizer"},
            ],
            "vectorizers": [
                {
                    "name": "text-vectorizer",
                    "kind": "azureOpenAI",
                    "azureOpenAIParameters": openai.copy(),
                },
                {
                    "name": "image-vectorizer",
                    "kind": "aiServicesVision",
                    "aiServicesVisionParameters": vision,
                },
            ],
        },
    }
    image_path = "/document/normalized_images/*"
    page_path = "/document/pages/*"
    shared_mappings = [
        {"name": "title", "source": "/document/metadata_spo_item_name"},
        {"name": "metadata_spo_item_path", "source": "/document/metadata_spo_item_path"},
    ]
    skillset: dict[str, Any] = {
        "name": names["skillsets"],
        "description": "OCR, text chunks, OpenAI text embeddings and Vision image embeddings",
        "skills": [
            {
                "@odata.type": "#Microsoft.Skills.Vision.OcrSkill",
                "name": "ocr",
                "context": image_path,
                "lineEnding": "Space",
                "defaultLanguageCode": "en",
                "detectOrientation": True,
                "inputs": [{"name": "image", "source": image_path}],
                "outputs": [{"name": "text", "targetName": "text"}],
            },
            {
                "@odata.type": "#Microsoft.Skills.Text.MergeSkill",
                "name": "merge",
                "context": "/document",
                "insertPreTag": " ",
                "insertPostTag": " ",
                "inputs": [
                    {"name": "text", "source": "/document/content"},
                    {"name": "itemsToInsert", "source": f"{image_path}/text"},
                    {"name": "offsets", "source": f"{image_path}/contentOffset"},
                ],
                "outputs": [{"name": "mergedText", "targetName": "mergedText"}],
            },
            {
                "@odata.type": "#Microsoft.Skills.Text.SplitSkill",
                "name": "split",
                "context": "/document",
                "defaultLanguageCode": "en",
                "textSplitMode": "pages",
                "maximumPageLength": 2000,
                "pageOverlapLength": 500,
                "maximumPagesToTake": 0,
                "unit": "characters",
                "inputs": [{"name": "text", "source": "/document/mergedText"}],
                "outputs": [{"name": "textItems", "targetName": "pages"}],
            },
            {
                "@odata.type": "#Microsoft.Skills.Text.AzureOpenAIEmbeddingSkill",
                "name": "embed-text",
                "context": page_path,
                **openai,
                "dimensions": 3072,
                "inputs": [{"name": "text", "source": page_path}],
                "outputs": [{"name": "embedding", "targetName": "text_vector"}],
            },
            {
                "@odata.type": "#Microsoft.Skills.Vision.VectorizeSkill",
                "name": "embed-image",
                "context": image_path,
                "modelVersion": VISION_MODEL,
                "inputs": [{"name": "image", "source": image_path}],
                "outputs": [{"name": "vector", "targetName": "image_vector"}],
            },
        ],
        "cognitiveServices": {
            "@odata.type": (
                "#Microsoft.Azure.Search.AIServicesByKey"
                if settings.ai_services_key
                else "#Microsoft.Azure.Search.AIServicesByIdentity"
            ),
            "subdomainUrl": settings.ai_services_endpoint,
            **({"key": settings.ai_services_key} if settings.ai_services_key else {}),
        },
        "indexProjections": {
            "selectors": [
                {
                    "targetIndexName": names["indexes"],
                    "parentKeyFieldName": "text_parent_id",
                    "sourceContext": page_path,
                    "mappings": [
                        {"name": "text_vector", "source": f"{page_path}/text_vector"},
                        {"name": "chunk", "source": page_path},
                        *shared_mappings,
                    ],
                },
                {
                    "targetIndexName": names["indexes"],
                    "parentKeyFieldName": "image_parent_id",
                    "sourceContext": image_path,
                    "mappings": [
                        {"name": "image_vector", "source": f"{image_path}/image_vector"},
                        *shared_mappings,
                    ],
                },
            ],
            "parameters": {"projectionMode": "skipIndexingParentDocuments"},
        },
    }
    indexer: dict[str, Any] = {
        "name": names["indexers"],
        "dataSourceName": names["datasources"],
        "skillsetName": names["skillsets"],
        "targetIndexName": names["indexes"],
        "parameters": {
            "maxFailedItems": 0,
            "maxFailedItemsPerBatch": 0,
            "configuration": {
                "imageAction": "generateNormalizedImages",
                "parsingMode": "default",
                "dataToExtract": "contentAndMetadata",
            },
        },
        "fieldMappings": [
            {"sourceFieldName": "metadata_spo_item_name", "targetFieldName": "title"},
        ],
        "schedule": {"interval": settings.schedule} if settings.schedule else None,
    }
    if settings.deployment_type == DeploymentType.NO_EMBEDDING:
        index["fields"] = [
            item
            for item in index["fields"]
            if item["name"] in ("chunk_id", "chunk", "title", "metadata_spo_item_path")
        ]
        del index["vectorSearch"]
        del indexer["skillsetName"]
        indexer["parameters"]["configuration"]["imageAction"] = "none"
        indexer["fieldMappings"].extend(
            [
                {
                    "sourceFieldName": "metadata_spo_site_library_item_id",
                    "targetFieldName": "chunk_id",
                    "mappingFunction": {"name": "base64Encode"},
                },
                {"sourceFieldName": "content", "targetFieldName": "chunk"},
            ]
        )
        return {"datasources": source, "indexes": index, "indexers": indexer}

    if not settings.use_text:
        index["fields"] = [item for item in index["fields"] if item["name"] != "text_vector"]
        index["vectorSearch"]["profiles"] = [
            item for item in index["vectorSearch"]["profiles"] if item["name"] != "text-profile"
        ]
        index["vectorSearch"]["vectorizers"] = [
            item
            for item in index["vectorSearch"]["vectorizers"]
            if item["name"] != "text-vectorizer"
        ]
        skillset["skills"] = [item for item in skillset["skills"] if item["name"] != "embed-text"]
        text_selector = skillset["indexProjections"]["selectors"][0]
        text_selector["mappings"] = [
            item for item in text_selector["mappings"] if item["name"] != "text_vector"
        ]
    if not settings.use_images:
        index["fields"] = [
            item
            for item in index["fields"]
            if item["name"] not in ("image_parent_id", "image_vector")
        ]
        index["vectorSearch"]["profiles"] = [
            item for item in index["vectorSearch"]["profiles"] if item["name"] != "image-profile"
        ]
        index["vectorSearch"]["vectorizers"] = [
            item
            for item in index["vectorSearch"]["vectorizers"]
            if item["name"] != "image-vectorizer"
        ]
        skillset["skills"] = [
            item
            for item in skillset["skills"]
            if item["name"] not in ("ocr", "merge", "embed-image")
        ]
        skillset["skills"][0]["inputs"][0]["source"] = "/document/content"
        del skillset["cognitiveServices"]
        skillset["indexProjections"]["selectors"] = skillset["indexProjections"]["selectors"][:1]
        indexer["parameters"]["configuration"]["imageAction"] = "none"
    skillset["description"] = f"SharePoint {settings.deployment_type} pipeline"
    return {"datasources": source, "indexes": index, "skillsets": skillset, "indexers": indexer}


def send_requests(
    settings: Settings,
    resources: dict[str, dict[str, Any]],
    *,
    status_only: bool = False,
    transport: httpx.BaseTransport | None = None,
) -> None:
    with ExitStack() as stack:
        client = stack.enter_context(httpx.Client(timeout=60, transport=transport))
        credential = (
            stack.enter_context(ManagedIdentityCredential()) if not settings.search_key else None
        )

        def request(method: str, path: str, body: dict[str, Any] | None = None) -> httpx.Response:
            if credential is not None:
                headers = {
                    "Authorization": "Bearer "
                    + credential.get_token("https://search.azure.com/.default").token
                }
            else:
                headers = {"api-key": settings.search_key}
            response = client.request(
                method,
                settings.search_endpoint + path,
                params={"api-version": API_VERSION},
                headers=headers,
                json=body,
            )
            if not response.is_success:
                raise RuntimeError(
                    f"{method} {path}: HTTP {response.status_code}: "
                    + settings.hide_secrets(response.text)
                )
            return response

        if status_only:
            response = request("GET", f"/indexers/{settings.names['indexers']}/status")
            print(settings.hide_secrets(json.dumps(response.json(), indent=2)))
            return
        for collection, body in resources.items():
            path = f"/{collection}/{body['name']}"
            response = request("PUT", path, body)
            print(f"Created/updated {path}")
            if collection == "indexers" and response.status_code == 200:
                request("POST", f"{path}/run")
                print(f"Requested incremental run for {path}")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--type",
        choices=list(DeploymentType),
        help="Override DEPLOYMENT_TYPE (default: no_embedding).",
    )
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument(
        "--dry-run", action="store_true", help="Print redacted JSON without Azure calls."
    )
    mode.add_argument("--status", action="store_true", help="Print indexer execution status.")
    args = parser.parse_args(argv)
    settings = None
    try:
        settings = Settings.from_env(args.type)
        resources = build_resources(settings)
        if args.dry_run:
            print(settings.hide_secrets(json.dumps(resources, indent=2)))
        else:
            send_requests(settings, resources, status_only=args.status)
            if not args.status:
                print("New indexers start automatically. Use --status to check indexing results.")
        return 0
    except (ValueError, RuntimeError, AzureError, httpx.HTTPError) as exc:
        message = settings.hide_secrets(str(exc)) if settings else str(exc)
        print(f"Error: {message}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
