"""Read-only GraphBridge example using placeholder environment variables."""

import os

from graphbridge import GbAuth, GbList, GbSite

if __name__ == "__main__":
    auth = GbAuth(
        tenant_id=os.environ["AZURE_TENANT_ID"],
        client_id=os.environ["AZURE_CLIENT_ID"],
        client_secret=os.environ["AZURE_CLIENT_SECRET"],
    )
    site = GbSite(
        hostname=os.environ["SHAREPOINT_HOSTNAME"],
        site_path=os.environ["SHAREPOINT_SITE_PATH"],
        gb_auth=auth,
    )
    sharepoint_list = GbList(
        list_name=os.environ["SHAREPOINT_LIST_NAME"],
        gb_site=site,
    )

    # This example intentionally performs a read-only operation.
    print(sharepoint_list.list_rows[:5])
