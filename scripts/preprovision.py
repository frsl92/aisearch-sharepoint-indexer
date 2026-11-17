"""azd hook: enable the Search serverless preview in the selected subscription."""

import os
import sys
import time
from collections.abc import Callable
from typing import Any
from uuid import UUID

import httpx
from azure.core.exceptions import AzureError
from azure.identity import AzureDeveloperCliCredential

from provision import env, required

ARM_ENDPOINT = "https://management.azure.com"
FEATURE_API_VERSION = "2021-07-01"
PROVIDER_API_VERSION = "2021-04-01"
REGISTRATION_TIMEOUT_SECONDS = 600
POLL_INTERVAL_SECONDS = 10


def registration_state(response: httpx.Response, *keys: str) -> str:
    try:
        value: Any = response.json()
    except ValueError as exc:
        raise RuntimeError("Azure returned invalid registration JSON.") from exc
    for key in keys:
        if not isinstance(value, dict) or key not in value:
            raise RuntimeError(f"Azure registration response is missing {'.'.join(keys)}.")
        value = value[key]
    if not isinstance(value, str) or not value.strip():
        raise RuntimeError(f"Azure registration response has an invalid {'.'.join(keys)}.")
    return value.strip()


def check_registration_state(name: str, state: str) -> None:
    if state.casefold() == "pending":
        raise RuntimeError(
            f"{name} registration is pending approval. Contact Azure support before retrying."
        )
    if state.casefold() not in {"registered", "registering", "notregistered", "unregistered"}:
        raise RuntimeError(f"{name} registration is in unexpected state {state!r}.")


def wait_until_registered(name: str, read_state: Callable[[], str]) -> None:
    deadline = time.monotonic() + REGISTRATION_TIMEOUT_SECONDS
    previous_state = ""
    while True:
        state = read_state()
        check_registration_state(name, state)
        if state.casefold() == "registered":
            return
        if state != previous_state:
            print(f"{name} registration state: {state}.", flush=True)
            previous_state = state
        remaining = deadline - time.monotonic()
        if remaining <= 0:
            raise RuntimeError(
                f"{name} registration did not reach Registered within "
                f"{REGISTRATION_TIMEOUT_SECONDS} seconds; last state: {state}. "
                "Retry provisioning after registration completes."
            )
        time.sleep(min(POLL_INTERVAL_SECONDS, remaining))


def register_serverless(
    subscription_id: str, *, transport: httpx.BaseTransport | None = None
) -> None:
    subscription = f"/subscriptions/{subscription_id}"
    feature = (
        f"{subscription}/providers/Microsoft.Features"
        "/providers/Microsoft.Search/features/Serverless"
    )
    provider = f"{subscription}/providers/Microsoft.Search"

    with (
        AzureDeveloperCliCredential(tenant_id=env("AZURE_TENANT_ID")) as credential,
        httpx.Client(base_url=ARM_ENDPOINT, timeout=30, transport=transport) as client,
    ):

        def request(method: str, path: str, api_version: str) -> httpx.Response:
            token = credential.get_token(f"{ARM_ENDPOINT}/.default").token
            response = client.request(
                method,
                path,
                params={"api-version": api_version},
                headers={"Authorization": f"Bearer {token}"},
            )
            if not response.is_success:
                raise RuntimeError(
                    f"Azure registration request {method} {path} failed "
                    f"with HTTP {response.status_code}: {response.text}"
                )
            return response

        def feature_state() -> str:
            return registration_state(
                request("GET", feature, FEATURE_API_VERSION), "properties", "state"
            )

        def provider_state() -> str:
            return registration_state(
                request("GET", provider, PROVIDER_API_VERSION), "registrationState"
            )

        state = feature_state()
        check_registration_state("Microsoft.Search/Serverless", state)
        if state.casefold() != "registered":
            if state.casefold() in {"notregistered", "unregistered"}:
                print(f"Registering Search serverless preview for subscription {subscription_id}.")
                request("POST", f"{feature}/register", FEATURE_API_VERSION)
            wait_until_registered("Microsoft.Search/Serverless", feature_state)
        else:
            print("Microsoft.Search/Serverless is already registered.")

        print("Refreshing the Microsoft.Search provider registration.", flush=True)
        request("POST", f"{provider}/register", PROVIDER_API_VERSION)
        wait_until_registered("Microsoft.Search provider", provider_state)
        print("Search serverless preview registration is ready.")


def preprovision() -> int:
    if os.environ.get("SEARCH_SKU", "basic").lower() != "serverless":
        return 0

    try:
        raw_subscription_id = required("AZURE_SUBSCRIPTION_ID")
        try:
            subscription_id = str(UUID(raw_subscription_id))
        except ValueError as exc:
            raise ValueError("AZURE_SUBSCRIPTION_ID must be a subscription UUID.") from exc
        register_serverless(subscription_id)
    except (AzureError, httpx.HTTPError, RuntimeError, ValueError) as exc:
        print(f"Error preparing Search serverless preview: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(preprovision())
