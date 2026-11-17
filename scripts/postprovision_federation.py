"""Configure or verify the SharePoint app's trust of the Search identity."""

import os
import subprocess
import sys
import time
from typing import Any
from uuid import UUID

import httpx
from azure.core.exceptions import AzureError
from azure.identity import AzureDeveloperCliCredential

from provision import env, env_bool, required

GRAPH_ENDPOINT = "https://graph.microsoft.com/v1.0"
TOKEN_EXCHANGE_AUDIENCE = "api://AzureADTokenExchange"


def _ensure_sharepoint_federation(
    *, create: bool, transport: httpx.BaseTransport | None = None
) -> str:
    ids = {}
    for variable in (
        "AZURE_TENANT_ID",
        "SHAREPOINT_TENANT_ID",
        "SHAREPOINT_APPLICATION_ID",
        "SEARCH_IDENTITY_PRINCIPAL_ID",
    ):
        try:
            ids[variable] = str(UUID(required(variable)))
        except ValueError as exc:
            raise ValueError(f"{variable} must be a GUID.") from exc
    tenant_id = ids["AZURE_TENANT_ID"]
    if ids["SHAREPOINT_TENANT_ID"] != tenant_id:
        raise ValueError("Federation requires the app and Search in the same tenant.")
    environment = required("AZURE_ENV_NAME") if create else ""
    principal_id = ids["SEARCH_IDENTITY_PRINCIPAL_ID"]
    issuer = f"https://login.microsoftonline.com/{tenant_id}/v2.0"
    credential_name = required("AZURE_SEARCH_SERVICE_NAME") if create else ""
    app_id = ids["SHAREPOINT_APPLICATION_ID"]
    credentials_url = (
        f"{GRAPH_ENDPOINT}/applications(appId='{app_id}')/federatedIdentityCredentials"
    )

    with (
        AzureDeveloperCliCredential(tenant_id=tenant_id) as credential,
        httpx.Client(timeout=30, transport=transport) as client,
    ):

        def request(
            method: str,
            url: str,
            body: dict[str, Any] | None = None,
            *,
            retry_not_found: bool = False,
        ) -> httpx.Response:
            for attempt in range(11):
                token = credential.get_token("https://graph.microsoft.com/.default").token
                response = client.request(
                    method, url, headers={"Authorization": f"Bearer {token}"}, json=body
                )
                if response.status_code != 404 or not retry_not_found or attempt == 10:
                    break
                print("Search identity is not yet visible in Entra; retrying in 30 seconds.")
                time.sleep(30)
            if not response.is_success:
                raise RuntimeError(
                    f"Microsoft Graph {method} failed with HTTP {response.status_code}: "
                    f"{response.text}"
                )
            return response

        def payload(response: httpx.Response) -> dict[str, Any]:
            value = response.json()
            if not isinstance(value, dict):
                raise ValueError("Microsoft Graph returned an invalid response object.")
            return value

        identity = payload(
            request(
                "GET",
                f"{GRAPH_ENDPOINT}/servicePrincipals/{principal_id}"
                "?$select=id,appId,servicePrincipalType",
                retry_not_found=True,
            )
        )
        if (
            identity.get("id") != principal_id
            or identity.get("servicePrincipalType") != "ManagedIdentity"
        ):
            raise ValueError("SEARCH_IDENTITY_PRINCIPAL_ID is not the expected managed identity.")
        try:
            client_id = str(UUID(identity.get("appId", "")))
        except (ValueError, TypeError, AttributeError) as exc:
            raise ValueError("Search managed identity has no valid application/client ID.") from exc
        credentials: list[dict[str, Any]] = []
        next_url = credentials_url
        visited: set[str] = set()
        while next_url:
            if not next_url.startswith(GRAPH_ENDPOINT + "/") or next_url in visited:
                raise ValueError("Microsoft Graph returned an invalid federation pagination URL.")
            visited.add(next_url)
            page = payload(request("GET", next_url))
            entries = page.get("value")
            if not isinstance(entries, list) or any(not isinstance(item, dict) for item in entries):
                raise ValueError("Microsoft Graph returned an invalid federated credentials list.")
            credentials.extend(entries)
            link = page.get("@odata.nextLink")
            if link is not None and not isinstance(link, str):
                raise ValueError("Microsoft Graph returned an invalid federation pagination URL.")
            next_url = link or ""

        matching = [
            item
            for item in credentials
            if item.get("issuer") == issuer and item.get("subject") == principal_id
        ]
        if matching:
            if any(item.get("audiences") != [TOKEN_EXCHANGE_AUDIENCE] for item in matching):
                raise ValueError(
                    "Existing federation for this identity has an incompatible audience."
                )
            print("The SharePoint app already trusts this Search identity.")
        else:
            if not create:
                raise ValueError(
                    "No matching SharePoint federated credential exists for this Search identity. "
                    "Have an app owner/administrator configure it, or enable "
                    "CONFIGURE_SHAREPOINT_FEDERATION with the required permissions."
                )
            description = f"Search identity for azd environment {environment}."
            existing = next(
                (item for item in credentials if item.get("name") == credential_name), None
            )
            if existing is not None:
                if (
                    existing.get("issuer") != issuer
                    or existing.get("audiences") != [TOKEN_EXCHANGE_AUDIENCE]
                    or existing.get("description") != description
                ):
                    raise ValueError(
                        "Federated credential name conflicts with another trust; not updated. "
                        "Automatic updates require the same issuer, audience and "
                        f"description {description!r}. Have an app owner/administrator "
                        "review this credential."
                    )
                try:
                    credential_id = str(UUID(existing.get("id", "")))
                except (ValueError, TypeError, AttributeError) as exc:
                    raise ValueError(
                        "Existing federated credential has no valid ID; not updated."
                    ) from exc
                request(
                    "PATCH",
                    f"{credentials_url}/{credential_id}",
                    {"subject": principal_id},
                )
                print(
                    f"Updated SharePoint app federated credential {credential_name} "
                    "to trust the current Search identity."
                )
            else:
                if len(credentials) >= 20:
                    raise ValueError(
                        "The app has reached its 20 federated credential limit; "
                        "review obsolete credentials before retrying."
                    )
                request(
                    "POST",
                    credentials_url,
                    {
                        "name": credential_name,
                        "issuer": issuer,
                        "subject": principal_id,
                        "audiences": [TOKEN_EXCHANGE_AUDIENCE],
                        "description": description,
                    },
                )
                print(f"Added SharePoint app federated credential {credential_name}.")
            print("Entra trust propagation can take time; check indexer status after setup.")
    return client_id


def verify_sharepoint_federation(*, transport: httpx.BaseTransport | None = None) -> str:
    return _ensure_sharepoint_federation(create=False, transport=transport)


def save_search_identity_client_id(client_id: str) -> None:
    environment = required("AZURE_ENV_NAME")

    try:
        subprocess.run(
            [
                "azd",
                "env",
                "set",
                "SEARCH_IDENTITY_CLIENT_ID",
                client_id,
                "--environment",
                environment,
            ],
            check=True,
            capture_output=True,
            text=True,
            timeout=60,
        )
    except subprocess.CalledProcessError as exc:
        raise RuntimeError(
            f"Could not save SEARCH_IDENTITY_CLIENT_ID in the azd environment: {exc.stderr}"
        ) from exc
    except (OSError, subprocess.TimeoutExpired) as exc:
        raise RuntimeError(f"Could not save SEARCH_IDENTITY_CLIENT_ID: {exc}") from exc
    os.environ["SEARCH_IDENTITY_CLIENT_ID"] = client_id
    print("Saved the Search identity client ID in the selected azd environment.")


def configure_sharepoint_federation(*, transport: httpx.BaseTransport | None = None) -> None:
    if env("SHAREPOINT_APPLICATION_SECRET"):
        raise ValueError(
            "CONFIGURE_SHAREPOINT_FEDERATION=true requires "
            "SHAREPOINT_APPLICATION_SECRET to be unset."
        )
    client_id = _ensure_sharepoint_federation(create=True, transport=transport)
    save_search_identity_client_id(client_id)


def postprovision_federation() -> int:
    try:
        if not env_bool("CONFIGURE_SHAREPOINT_FEDERATION"):
            print("Federation setup skipped (CONFIGURE_SHAREPOINT_FEDERATION=false).")
            return 0
        configure_sharepoint_federation()
    except (AzureError, httpx.HTTPError, RuntimeError, ValueError) as exc:
        print(f"Error configuring SharePoint federation: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(postprovision_federation())
