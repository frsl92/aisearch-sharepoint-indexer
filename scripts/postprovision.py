"""azd hook: run independently controlled federation and Search setup in order."""

import sys

from postprovision_federation import postprovision_federation
from postprovision_search import postprovision_search
from provision import env_bool


def postprovision() -> int:
    try:
        env_bool("CONFIGURE_SHAREPOINT_FEDERATION")
        search_enabled = env_bool("CONFIGURE_SEARCH_PIPELINE", default=True)
    except ValueError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1
    result = postprovision_federation()
    if result:
        return result
    return postprovision_search(enabled=search_enabled)


if __name__ == "__main__":
    sys.exit(postprovision())
