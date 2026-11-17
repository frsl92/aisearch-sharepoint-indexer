import os
import subprocess
import sys
from pathlib import Path
from unittest.mock import patch

import pytest

from postprovision import postprovision


@pytest.fixture(autouse=True)
def isolate_search_flags(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.delenv("CONFIGURE_SEARCH_PIPELINE", raising=False)


@pytest.mark.parametrize(
    ("configure", "search", "expected"),
    [
        ("true", "true", ["federation", "verify", "save", "pipeline"]),
        ("true", "false", ["federation"]),
        ("false", "true", ["verify", "save", "pipeline"]),
        ("false", "false", []),
    ],
)
def test_stage_controls_are_independent_and_ordered(
    monkeypatch: pytest.MonkeyPatch, configure: str, search: str, expected: list[str]
) -> None:
    monkeypatch.setenv("CONFIGURE_SHAREPOINT_FEDERATION", configure)
    monkeypatch.setenv("CONFIGURE_SEARCH_PIPELINE", search)
    monkeypatch.setenv("AZURE_ENV_NAME", "qa")
    monkeypatch.delenv("SHAREPOINT_APPLICATION_SECRET", raising=False)
    monkeypatch.delenv("SEARCH_IDENTITY_CLIENT_ID", raising=False)
    events: list[str] = []
    client_id = "44444444-4444-4444-4444-444444444444"

    def configure_trust() -> None:
        events.append("federation")
        monkeypatch.setenv("SEARCH_IDENTITY_CLIENT_ID", client_id)

    def verify_trust() -> str:
        events.append("verify")
        if configure == "true":
            assert os.environ["SEARCH_IDENTITY_CLIENT_ID"] == client_id
        return client_id

    def configure_pipeline(arguments: list[str]) -> int:
        assert arguments == ["--search-auth", "azd"]
        assert os.environ["SEARCH_IDENTITY_CLIENT_ID"] == client_id
        events.append("pipeline")
        return 0

    def save_client_id(value: str) -> None:
        assert value == client_id
        events.append("save")
        monkeypatch.setenv("SEARCH_IDENTITY_CLIENT_ID", value)

    with (
        patch("postprovision_federation.configure_sharepoint_federation", configure_trust),
        patch("postprovision_search.verify_sharepoint_federation", verify_trust),
        patch("postprovision_search.save_search_identity_client_id", save_client_id),
        patch("postprovision_search.main", configure_pipeline),
    ):
        assert postprovision() == 0
    assert events == expected


@pytest.mark.parametrize("flag", ["CONFIGURE_SHAREPOINT_FEDERATION", "CONFIGURE_SEARCH_PIPELINE"])
def test_all_flags_are_validated_before_any_stage(
    monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str], flag: str
) -> None:
    monkeypatch.setenv("CONFIGURE_SHAREPOINT_FEDERATION", "true")
    monkeypatch.setenv(flag, "invalid")
    with (
        patch("postprovision.postprovision_federation") as federation,
        patch("postprovision.postprovision_search") as search,
    ):
        assert postprovision() == 1
    federation.assert_not_called()
    search.assert_not_called()
    assert f"{flag} must be true or false" in capsys.readouterr().err


def test_federation_failure_prevents_search_stage(
    monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str]
) -> None:
    monkeypatch.setenv("CONFIGURE_SHAREPOINT_FEDERATION", "true")
    monkeypatch.setenv("CONFIGURE_SEARCH_PIPELINE", "true")
    with (
        patch(
            "postprovision_federation.configure_sharepoint_federation",
            side_effect=RuntimeError("HTTP 403: denied"),
        ),
        patch("postprovision.postprovision_search") as search,
    ):
        assert postprovision() == 1
    search.assert_not_called()
    assert "HTTP 403: denied" in capsys.readouterr().err


@pytest.mark.parametrize(
    ("script", "flag", "value", "expected"),
    [
        ("postprovision.py", "CONFIGURE_SEARCH_PIPELINE", "false", 0),
        ("postprovision.py", "CONFIGURE_SEARCH_PIPELINE", "invalid", 1),
        ("postprovision_federation.py", "CONFIGURE_SHAREPOINT_FEDERATION", "false", 0),
        ("postprovision_federation.py", "CONFIGURE_SHAREPOINT_FEDERATION", "invalid", 1),
        ("postprovision_search.py", "CONFIGURE_SEARCH_PIPELINE", "false", 0),
        ("postprovision_search.py", "CONFIGURE_SEARCH_PIPELINE", "invalid", 1),
    ],
)
def test_stage_entry_points(
    monkeypatch: pytest.MonkeyPatch, script: str, flag: str, value: str, expected: int
) -> None:
    monkeypatch.setenv("CONFIGURE_SHAREPOINT_FEDERATION", "false")
    monkeypatch.setenv(flag, value)
    root = Path(__file__).resolve().parents[1]
    result = subprocess.run(
        [sys.executable, str(root / "scripts" / script)],
        cwd=root,
        capture_output=True,
        text=True,
        check=False,
    )
    assert result.returncode == expected
    assert flag in result.stdout + result.stderr
