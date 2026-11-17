import os
from unittest.mock import patch

import pytest

from postprovision_search import postprovision_search


@pytest.fixture(autouse=True)
def pipeline_environment(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.delenv("CONFIGURE_SEARCH_PIPELINE", raising=False)
    monkeypatch.delenv("SHAREPOINT_APPLICATION_SECRET", raising=False)
    monkeypatch.setenv("SEARCH_IDENTITY_CLIENT_ID", "previous-client-id")


@pytest.mark.parametrize("result", [0, 1])
def test_pipeline_saves_verified_client_id_before_provisioning(result: int) -> None:
    events: list[str] = []

    def verified_client_id() -> str:
        events.append("verify")
        return "verified-client-id"

    def save_client_id(client_id: str) -> None:
        events.append("save")
        assert client_id == "verified-client-id"
        os.environ["SEARCH_IDENTITY_CLIENT_ID"] = client_id

    def provision(arguments: list[str]) -> int:
        assert events == ["verify", "save"]
        assert arguments == ["--search-auth", "azd"]
        assert os.environ["SEARCH_IDENTITY_CLIENT_ID"] == "verified-client-id"
        return result

    with (
        patch("postprovision_search.verify_sharepoint_federation", verified_client_id),
        patch("postprovision_search.save_search_identity_client_id", save_client_id),
        patch("postprovision_search.main", provision),
    ):
        assert postprovision_search() == result


def test_pipeline_does_not_run_if_saving_verified_client_id_fails(
    capsys: pytest.CaptureFixture[str],
) -> None:
    with (
        patch("postprovision_search.verify_sharepoint_federation", return_value="verified-id"),
        patch(
            "postprovision_search.save_search_identity_client_id",
            side_effect=RuntimeError("azd env set failed"),
        ),
        patch("postprovision_search.main") as provision,
    ):
        assert postprovision_search() == 1
    provision.assert_not_called()
    assert "azd env set failed" in capsys.readouterr().err


@pytest.mark.parametrize("value", ["false", "invalid"])
def test_skipped_or_invalid_stage_does_not_check_or_mutate_anything(
    monkeypatch: pytest.MonkeyPatch, value: str
) -> None:
    monkeypatch.setenv("CONFIGURE_SEARCH_PIPELINE", value)
    with (
        patch("postprovision_search.verify_sharepoint_federation") as verify,
        patch("postprovision_search.save_search_identity_client_id") as save,
        patch("postprovision_search.main") as provision,
    ):
        assert postprovision_search() == (0 if value == "false" else 1)
    verify.assert_not_called()
    save.assert_not_called()
    provision.assert_not_called()


@pytest.mark.parametrize("error", ["No matching trust exists", "Microsoft Graph HTTP 403"])
def test_verification_failure_blocks_pipeline(
    capsys: pytest.CaptureFixture[str], error: str
) -> None:
    with (
        patch(
            "postprovision_search.verify_sharepoint_federation",
            side_effect=RuntimeError(error),
        ),
        patch("postprovision_search.main") as provision,
    ):
        assert postprovision_search() == 1
    provision.assert_not_called()
    assert error in capsys.readouterr().err


def test_secret_auth_does_not_require_federation(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("SHAREPOINT_APPLICATION_SECRET", "test-secret")
    with (
        patch("postprovision_search.verify_sharepoint_federation") as verify,
        patch("postprovision_search.save_search_identity_client_id") as save,
        patch("postprovision_search.main", return_value=0) as provision,
    ):
        assert postprovision_search() == 0
    verify.assert_not_called()
    save.assert_not_called()
    provision.assert_called_once_with(["--search-auth", "azd"])
