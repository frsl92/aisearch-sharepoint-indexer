import json
import os
import shutil
import subprocess
from pathlib import Path
from typing import Any

import pytest


@pytest.fixture(scope="module")
def compiled_template() -> dict[str, Any]:
    compiler = shutil.which("bicep")
    if not compiler:
        executable = "bicep.exe" if os.name == "nt" else "bicep"
        installed = Path.home() / ".azure" / "bin" / executable
        if not installed.is_file():
            pytest.skip("Bicep CLI is required for infrastructure contract tests.")
        compiler = str(installed)
    root = Path(__file__).resolve().parents[1]
    result = subprocess.run(
        [compiler, "build", str(root / "infra" / "main.bicep"), "--stdout"],
        capture_output=True,
        text=True,
        check=True,
        timeout=120,
    )
    template: dict[str, Any] = json.loads(result.stdout)
    return template


def test_multi_service_account_and_rbac_are_available_for_image_modes(
    compiled_template: dict[str, Any],
) -> None:
    assert {"image_embeddings", "all_embeddings"} <= set(
        compiled_template["parameters"]["deploymentType"]["allowedValues"]
    )
    module = next(
        resource
        for resource in compiled_template["resources"]
        if resource["type"] == "Microsoft.Resources/deployments"
    )
    template = module["properties"]["template"]
    resources = template["resources"]
    account = next(r for r in resources if r["type"] == "Microsoft.CognitiveServices/accounts")
    search = next(r for r in resources if r["type"] == "Microsoft.Search/searchServices")
    # These are unconditional, so both image modes receive the same account and access.
    for resource in (module, account, search):
        assert resource.get("condition", True) is True
    assert account["kind"] == "AIServices"
    assert account["sku"]["name"] == "S0"
    assert account["properties"]["customSubDomainName"] == account["name"]
    assert account["properties"]["disableLocalAuth"] is True
    assert search["identity"]["type"] == "SystemAssigned"

    identity = template["outputs"]["searchPrincipalId"]["value"]
    account_roles = [
        r
        for r in resources
        if r["type"] == "Microsoft.Authorization/roleAssignments"
        and r["properties"]["principalId"] == identity
    ]
    roles = {
        template["variables"][r["properties"]["roleDefinitionId"].split("'")[1]]
        for r in account_roles
    }
    assert any("a97b65f3-24c7-4388-baec-2e87135dc908" in role for role in roles)
    assert any("5e0bd9bd-7b93-4f28-af87-19fc36ad61bd" in role for role in roles)
    for role in account_roles:
        assert role.get("condition", True) is True
        assert role["properties"]["principalType"] == "ServicePrincipal"
        assert "Microsoft.CognitiveServices/accounts" in role["scope"]
        assert account["name"][1:-1] in role["scope"]

    assert (
        "Microsoft.CognitiveServices/accounts" in template["outputs"]["aiServicesEndpoint"]["value"]
    )
    assert (
        ".outputs.aiServicesEndpoint.value"
        in (compiled_template["outputs"]["AZURE_AI_SERVICES_ENDPOINT"]["value"])
    )
