import subprocess
import sys
from collections.abc import Iterator
from pathlib import Path
from unittest.mock import MagicMock, patch

import httpx
import pytest

import preprovision

SUBSCRIPTION = "11111111-1111-1111-1111-111111111111"
FEATURE = (
    f"/subscriptions/{SUBSCRIPTION}/providers/Microsoft.Features"
    "/providers/Microsoft.Search/features/Serverless"
)
PROVIDER = f"/subscriptions/{SUBSCRIPTION}/providers/Microsoft.Search"


@pytest.fixture
def azd_credential() -> Iterator[MagicMock]:
    with patch("preprovision.AzureDeveloperCliCredential") as factory:
        credential = factory.return_value.__enter__.return_value
        credential.get_token.return_value.token = "test-token"
        yield factory


def arm_transport(steps: list[tuple[str, str, int, dict]]) -> httpx.MockTransport:
    def handle(request: httpx.Request) -> httpx.Response:
        assert steps, f"Unexpected request: {request.method} {request.url.path}"
        method, path, status, body = steps.pop(0)
        assert (request.method, request.url.path) == (method, path)
        assert request.headers["Authorization"] == "Bearer test-token"
        expected_version = "2021-07-01" if "Microsoft.Features" in path else "2021-04-01"
        assert request.url.params["api-version"] == expected_version
        return httpx.Response(status, json=body)

    return httpx.MockTransport(handle)


def feature_state(state: str) -> dict:
    return {"properties": {"state": state}}


def provider_state(state: str) -> dict:
    return {"registrationState": state}


def test_registers_feature_then_refreshes_provider(
    monkeypatch: pytest.MonkeyPatch, azd_credential: MagicMock
) -> None:
    monkeypatch.setenv("AZURE_TENANT_ID", "tenant-id")
    monkeypatch.setattr(preprovision.time, "sleep", lambda _: None)
    steps = [
        ("GET", FEATURE, 200, feature_state("NotRegistered")),
        ("POST", f"{FEATURE}/register", 200, {}),
        ("GET", FEATURE, 200, feature_state("Registering")),
        ("GET", FEATURE, 200, feature_state("Registered")),
        ("POST", f"{PROVIDER}/register", 200, {}),
        ("GET", PROVIDER, 200, provider_state("Registering")),
        ("GET", PROVIDER, 200, provider_state("Registered")),
    ]

    preprovision.register_serverless(SUBSCRIPTION, transport=arm_transport(steps))

    assert not steps
    azd_credential.assert_called_once_with(tenant_id="tenant-id")
    credential = azd_credential.return_value.__enter__.return_value
    credential.get_token.assert_called_with("https://management.azure.com/.default")


def test_already_registered_feature_still_refreshes_provider(
    azd_credential: MagicMock,
) -> None:
    steps = [
        ("GET", FEATURE, 200, feature_state("Registered")),
        ("POST", f"{PROVIDER}/register", 200, {}),
        ("GET", PROVIDER, 200, provider_state("Registered")),
    ]
    preprovision.register_serverless(SUBSCRIPTION, transport=arm_transport(steps))
    assert not steps


def test_registration_in_progress_is_not_restarted(
    monkeypatch: pytest.MonkeyPatch, azd_credential: MagicMock
) -> None:
    monkeypatch.setattr(preprovision.time, "sleep", lambda _: None)
    steps = [
        ("GET", FEATURE, 200, feature_state("Registering")),
        ("GET", FEATURE, 200, feature_state("Registered")),
        ("POST", f"{PROVIDER}/register", 200, {}),
        ("GET", PROVIDER, 200, provider_state("Registered")),
    ]
    preprovision.register_serverless(SUBSCRIPTION, transport=arm_transport(steps))
    assert not steps


def test_pending_approval_stops_before_provider_refresh(azd_credential: MagicMock) -> None:
    steps = [
        ("GET", FEATURE, 200, feature_state("NotRegistered")),
        ("POST", f"{FEATURE}/register", 200, {}),
        ("GET", FEATURE, 200, feature_state("Pending")),
    ]
    with pytest.raises(RuntimeError, match="pending approval"):
        preprovision.register_serverless(SUBSCRIPTION, transport=arm_transport(steps))
    assert not steps


def test_timeout_stops_before_provider_refresh(
    monkeypatch: pytest.MonkeyPatch, azd_credential: MagicMock
) -> None:
    class Clock:
        elapsed = 0.0

        def monotonic(self) -> float:
            return self.elapsed

        def sleep(self, seconds: float) -> None:
            self.elapsed += seconds

    clock = Clock()
    monkeypatch.setattr(preprovision, "time", clock)

    def handle(request: httpx.Request) -> httpx.Response:
        assert (request.method, request.url.path) == ("GET", FEATURE)
        return httpx.Response(200, json=feature_state("Registering"))

    with pytest.raises(RuntimeError, match="did not reach Registered"):
        preprovision.register_serverless(SUBSCRIPTION, transport=httpx.MockTransport(handle))
    assert clock.elapsed == 600


def test_provider_permission_error_is_reported(azd_credential: MagicMock) -> None:
    steps = [
        ("GET", FEATURE, 200, feature_state("Registered")),
        ("POST", f"{PROVIDER}/register", 403, {"error": {"message": "Access denied"}}),
    ]
    with pytest.raises(RuntimeError, match="HTTP 403.*Access denied"):
        preprovision.register_serverless(SUBSCRIPTION, transport=arm_transport(steps))
    assert not steps


def test_invalid_registration_response_does_not_continue(azd_credential: MagicMock) -> None:
    steps = [("GET", FEATURE, 200, {"properties": {}})]
    with pytest.raises(RuntimeError, match="missing properties.state"):
        preprovision.register_serverless(SUBSCRIPTION, transport=arm_transport(steps))
    assert not steps


@pytest.mark.parametrize("sku", [None, "basic", "standard"])
def test_dedicated_skus_do_not_access_azure(
    monkeypatch: pytest.MonkeyPatch, sku: str | None
) -> None:
    if sku is None:
        monkeypatch.delenv("SEARCH_SKU", raising=False)
    else:
        monkeypatch.setenv("SEARCH_SKU", sku)
    monkeypatch.delenv("AZURE_SUBSCRIPTION_ID", raising=False)
    with patch("preprovision.register_serverless") as register:
        assert preprovision.preprovision() == 0
    register.assert_not_called()


def test_serverless_uses_selected_subscription(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("SEARCH_SKU", "SERVERLESS")
    monkeypatch.setenv("AZURE_SUBSCRIPTION_ID", SUBSCRIPTION)
    with patch("preprovision.register_serverless") as register:
        assert preprovision.preprovision() == 0
    register.assert_called_once_with(SUBSCRIPTION)


@pytest.mark.parametrize("subscription", ["", "not-a-subscription"])
def test_missing_or_invalid_subscription_fails_without_registration(
    monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str], subscription: str
) -> None:
    monkeypatch.setenv("SEARCH_SKU", "serverless")
    monkeypatch.setenv("AZURE_SUBSCRIPTION_ID", subscription)
    with patch("preprovision.register_serverless") as register:
        assert preprovision.preprovision() == 1
    register.assert_not_called()
    assert "AZURE_SUBSCRIPTION_ID" in capsys.readouterr().err


def test_registration_failure_stops_the_hook(
    monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str]
) -> None:
    monkeypatch.setenv("SEARCH_SKU", "serverless")
    monkeypatch.setenv("AZURE_SUBSCRIPTION_ID", SUBSCRIPTION)
    with patch("preprovision.register_serverless", side_effect=RuntimeError("HTTP 403: denied")):
        assert preprovision.preprovision() == 1
    assert "HTTP 403: denied" in capsys.readouterr().err


def test_hook_entry_point_fails_without_subscription(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("SEARCH_SKU", "serverless")
    monkeypatch.delenv("AZURE_SUBSCRIPTION_ID", raising=False)
    root = Path(__file__).resolve().parents[1]
    result = subprocess.run(
        [sys.executable, str(root / "scripts" / "preprovision.py")],
        cwd=root,
        capture_output=True,
        text=True,
        check=False,
    )
    assert result.returncode == 1
    assert "AZURE_SUBSCRIPTION_ID" in result.stderr
