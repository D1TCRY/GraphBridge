"""Explicitly gated create/update example for a dedicated test list.

The example never deletes data and leaves its created item in the dedicated list
for manual inspection.
"""

from __future__ import annotations

import os
from uuid import uuid4

from azure.identity import ClientSecretCredential

from graphbridge import GraphBridgeClient

WRITE_CONFIRMATION = "CREATE_UPDATE_IN_DEDICATED_LIST"
DEDICATED_PREFIX = "GraphBridge Integration - "


def main() -> None:
    if os.environ.get("GRAPHBRIDGE_ALLOW_WRITES") != WRITE_CONFIRMATION:
        raise RuntimeError(
            f"Refusing writes: set GRAPHBRIDGE_ALLOW_WRITES={WRITE_CONFIRMATION} only for a dedicated test list"
        )
    expected_name = os.environ["SHAREPOINT_LIST_NAME"]
    if not expected_name.startswith(DEDICATED_PREFIX):
        raise RuntimeError(f"Dedicated list name must start with {DEDICATED_PREFIX!r}")

    credential = ClientSecretCredential(
        tenant_id=os.environ["AZURE_TENANT_ID"],
        client_id=os.environ["AZURE_CLIENT_ID"],
        client_secret=os.environ["AZURE_CLIENT_SECRET"],
    )
    with GraphBridgeClient(credential=credential) as client:
        site = client.sites.get(os.environ["SHAREPOINT_SITE_ID"])
        sharepoint_list = site.lists.get_by_id(os.environ["SHAREPOINT_LIST_ID"])
        if sharepoint_list.display_name != expected_name:
            raise RuntimeError("Configured list ID does not match the dedicated list name")

        marker = f"GraphBridge example {uuid4()}"
        created = sharepoint_list.items.create({"Title": marker})
        current = sharepoint_list.items.get(created.id, fields=("Title",))
        if current.etag is None:
            raise RuntimeError("Refusing an update without a current eTag")
        updated = sharepoint_list.items.update(
            created.id,
            {"Title": f"{marker} updated"},
            etag=current.etag,
        )
        print("Created and updated dedicated test item:", updated.id)
        print("No cleanup was attempted; review the marked item manually.")


if __name__ == "__main__":
    main()
