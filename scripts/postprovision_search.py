"""Configure Search objects after verifying the selected SharePoint authentication."""

import sys

import httpx
from azure.core.exceptions import AzureError

from postprovision_federation import (
    save_search_identity_client_id,
    verify_sharepoint_federation,
)
from provision import env, env_bool, main


def postprovision_search(enabled: bool | None = None) -> int:
    try:
        if enabled is None:
            enabled = env_bool("CONFIGURE_SEARCH_PIPELINE", default=True)
        if not enabled:
            print("Search pipeline setup skipped (CONFIGURE_SEARCH_PIPELINE=false).")
            return 0
        if not env("SHAREPOINT_APPLICATION_SECRET"):
            client_id = verify_sharepoint_federation()
            save_search_identity_client_id(client_id)
    except (AzureError, httpx.HTTPError, RuntimeError, ValueError) as exc:
        print(f"Error verifying SharePoint federation: {exc}", file=sys.stderr)
        return 1
    return main(["--search-auth", "azd"])


if __name__ == "__main__":
    sys.exit(postprovision_search())
