"""Review-first synchronization example with anonymous data.

The default performs reads and prints a dry-run plan. Writes and pruning have
separate explicit environment confirmations and require a dedicated list name.
"""

from __future__ import annotations

import json
import os
from typing import TYPE_CHECKING

from azure.identity import ClientSecretCredential

from graphbridge import GraphBridgeClient
from graphbridge.models import SyncResult

if TYPE_CHECKING:
    from graphbridge.resources import SharePointListResource

WRITE_CONFIRMATION = "APPLY_SYNC_TO_DEDICATED_LIST"
PRUNE_CONFIRMATION = "PRUNE_DEDICATED_LIST"
DEDICATED_PREFIX = "GraphBridge Integration - "

SOURCE_ROWS = [
    {"ExternalId": "DEMO-001", "Title": "Anonymous task", "Status": "Open"},
    {"ExternalId": "DEMO-002", "Title": "Second task", "Status": "Done"},
]


def synchronize(
    tasks: SharePointListResource,
    *,
    apply_changes: bool = False,
    prune: bool = False,
) -> SyncResult:
    """Plan, print, and optionally apply the anonymous example dataset."""

    if apply_changes:
        if os.environ.get("GRAPHBRIDGE_ALLOW_WRITES") != WRITE_CONFIRMATION:
            raise RuntimeError("Refusing synchronization writes without explicit confirmation")
        if not (tasks.display_name or "").startswith(DEDICATED_PREFIX):
            raise RuntimeError("Synchronization writes require a dedicated test list")
    if prune and apply_changes and os.environ.get("GRAPHBRIDGE_ALLOW_PRUNE") != PRUNE_CONFIRMATION:
        raise RuntimeError("Refusing prune without its separate explicit confirmation")

    plan = tasks.sync.plan(
        rows=SOURCE_ROWS,
        key_field="ExternalId",
        prune=prune,
        dry_run=not apply_changes,
    )
    print(json.dumps(plan.to_dict(), indent=2, default=str))
    return tasks.sync.apply(plan)


def main() -> None:
    credential = ClientSecretCredential(
        tenant_id=os.environ["AZURE_TENANT_ID"],
        client_id=os.environ["AZURE_CLIENT_ID"],
        client_secret=os.environ["AZURE_CLIENT_SECRET"],
    )
    with GraphBridgeClient(credential=credential) as client:
        site = client.sites.get(os.environ["SHAREPOINT_SITE_ID"])
        tasks = site.lists.get_by_id(os.environ["SHAREPOINT_LIST_ID"])
        synchronize(tasks)  # dry-run only


if __name__ == "__main__":
    main()
