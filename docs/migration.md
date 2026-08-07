# Migration guide

The composed API is the target for new code. `GbAuth`, `GbSite`, and `GbList`
remain temporarily available so applications can migrate one boundary at a time.
`GbList.upload()` already emits `DeprecationWarning`; no removal version is
committed until a stable 1.0 contract and a completed migration window exist.

## Authentication and composition

Before:

```python
import os

from graphbridge import GbAuth, GbList, GbSite

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
tasks = GbList(list_name=os.environ["SHAREPOINT_LIST_NAME"], gb_site=site)
rows = tasks.list_rows
```

After:

```python
import os

from azure.identity import ClientSecretCredential
from graphbridge import GraphBridgeClient

credential = ClientSecretCredential(
    tenant_id=os.environ["AZURE_TENANT_ID"],
    client_id=os.environ["AZURE_CLIENT_ID"],
    client_secret=os.environ["AZURE_CLIENT_SECRET"],
)
with GraphBridgeClient(credential=credential) as client:
    site = client.sites.get_by_path(
        hostname=os.environ["SHAREPOINT_HOSTNAME"],
        path=os.environ["SHAREPOINT_SITE_PATH"],
    )
    tasks = site.lists.get_by_id(os.environ["SHAREPOINT_LIST_ID"])
    rows = list(tasks.items.iter_all(fields=("Title", "Status")))
```

The credential implements the structural `TokenCredential` protocol. The same
client, transport, session, and credential are reused by every composed
resource. Azure Identity, rather than GraphBridge, owns token cache/renewal.

## CRUD result shapes

Legacy methods return dictionaries containing `successes` and `failures`. Modern
single-item methods return `ListItem` and raise typed exceptions; batch methods
return `BatchResult` with ordered per-input outcomes.

Before:

```python
result = tasks.update("42", {"Status": "Done"})
if result["failures"]:
    handle(result["failures"])
```

After:

```python
current = tasks.items.get("42", fields=("Status",))
updated = tasks.items.update(
    current.id,
    {"Status": "Done"},
    etag=current.etag,
)
```

Catch `GraphPreconditionFailedError` for an eTag conflict and re-read/re-plan;
do not silently overwrite.

## Replace `upload()` with plan/review/apply

Before:

```python
# Legacy: delete=True can prune remote-only rows.
result = tasks.upload(ids, rows, force=False, delete=True)
```

After:

```python
plan = tasks.sync.plan(
    rows=rows,
    key_field="ExternalId",
    prune=False,
    dry_run=True,
)
review(plan.to_dict())

# Build a non-dry plan only after review.
approved = tasks.sync.plan(
    rows=rows,
    key_field="ExternalId",
    prune=False,
)
result = tasks.sync.apply(approved)
```

Use a stable unique business key stored in the list, not the SharePoint item ID
unless that is explicitly your legacy contract. `prune=False` retains remote-only
items. Turning prune on must be a separate reviewed decision. Apply order is
create, PATCH update, delete; create failures defer deletes. Use
`tasks.sync.retry(result)` only for failed/deferred work.

`force=True` in the adapter no longer deletes and recreates an item: it PATCHes
the existing item and preserves the old `replaced` result section. Code that
depended on a new item ID must be redesigned explicitly.

## Replace heuristic field encoding with schema mapping

Before:

```python
encoded = tasks.encode_row({"Project Name": "Apollo"})
decoded = tasks.decode_row(encoded)
```

The legacy codec guesses SharePoint's encoded internal name from characters. It
cannot prove the actual column identity and remains only for compatibility.

After:

```python
name_map = tasks.columns.display_name_map()
payload = tasks.columns.to_internal_fields({"Project Name": "Apollo"})
created = tasks.items.create(
    {"Project Name": "Apollo"},
    field_names="display",
)
```

The composed API reads the list schema and maps display names to authoritative
internal names. Duplicate display names and unknown names fail explicitly.

## Warning and compatibility policy

- `GbAuth`, `GbSite`, and `GbList` are exported and characterized by unit tests.
- `GbList.upload()` emits a caller-facing `DeprecationWarning`.
- Legacy cached dictionary shapes, common exceptions, and batch result shapes
  remain supported during the migration window.
- Legacy cached tokens and heuristic encoding are not guarantees for the modern
  API and should not appear in new code.
- Adding names to `graphbridge.__all__` exposes the modern API without removing
  legacy names; code should nevertheless use explicit imports.

## Stable v1.0 limitations

GraphBridge does not fall back to Microsoft Graph beta. The current public API
does not provide:

- automatic full-state recovery after an expired delta cursor (HTTP 410);
- reliable create/update delta classification without caller-owned known IDs;
- durable delta cursor storage;
- document-set version creation/restoration;
- list-item `$orderby`, which the v1.0 list-items operation does not advertise;
- automatic business conflict resolution after HTTP 412;
- simulated endpoints for capabilities absent from v1.0.

## Release roadmap

- **0.1.0** — modern internal transport/composition while the legacy API remains
  functional.
- **0.2.0** — public composed API and formal legacy deprecation.
- **0.3.0** — schema mapping, delta, versions, and advanced safe synchronization.
- **1.0.0** — stable contract; possible legacy removal only after the migration
  guide and deprecation window have been completed.

The repository reports `0.1.0`. This release establishes the modern transport
and resource composition while retaining the legacy compatibility layer. The
later version numbers above are roadmap milestones, not the current package
version.
