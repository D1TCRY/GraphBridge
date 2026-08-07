"""Read items from an explicitly configured SharePoint site and list."""

from __future__ import annotations

import os

from azure.identity import ClientSecretCredential

from graphbridge import GraphBridgeClient


def main() -> None:
    credential = ClientSecretCredential(
        tenant_id=os.environ["AZURE_TENANT_ID"],
        client_id=os.environ["AZURE_CLIENT_ID"],
        client_secret=os.environ["AZURE_CLIENT_SECRET"],
    )
    with GraphBridgeClient(credential=credential) as client:
        site = client.sites.get(os.environ["SHAREPOINT_SITE_ID"])
        sharepoint_list = site.lists.get_by_id(os.environ["SHAREPOINT_LIST_ID"])
        for item in sharepoint_list.items.iter_all(
            fields=("Title",),
            top=100,
        ):
            print(item.id, item.fields.get("Title"))


if __name__ == "__main__":
    main()
