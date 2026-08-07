# GraphBridge

GraphBridge is a small app-only Python client for stable Microsoft Graph v1.0
SharePoint site, list, column, and list-item operations. It supports Python 3.10+
and keeps the original `GbAuth`/`GbSite`/`GbList` API as a compatibility layer.

No Graph beta endpoint is used.

The package version is `0.1.0`. This release introduces the modern internal
transport and resource composition while retaining the original API as a
compatibility layer. See the
[architecture](docs/architecture.md), [permissions](docs/permissions.md),
[migration guide](docs/migration.md), and
[release matrix](docs/release-readiness.md).

## Installation

```bash
pip install graphbridge
```

For development:

```bash
pip install -e ".[dev]"
```

An Entra ID application normally needs `Sites.Read.All` for reads,
`Sites.ReadWrite.All` for list-item writes, and `Sites.Manage.All` for list and
column schema mutations. For a single controlled site/list, prefer
`Sites.Selected` or `Lists.SelectedOperations.Selected` plus an explicit object
assignment. Grant only the permissions required by your workload; Entra consent
alone grants no object access for Selected permissions.

## Composed API

Create one client and reuse its credential, transport, and HTTP session:

```python
import os

from azure.identity import ClientSecretCredential
from graphbridge import GraphBridgeClient

credential = ClientSecretCredential(
    tenant_id=os.environ["AZURE_TENANT_ID"],
    client_id=os.environ["AZURE_CLIENT_ID"],
    client_secret=os.environ["AZURE_CLIENT_SECRET"],
)
client = GraphBridgeClient(credential=credential)
site = client.sites.get_by_path(
    hostname="contoso.sharepoint.com",
    path="/sites/Marketing",
)
tasks = site.lists.get_by_name("Tasks")
```

Use `get_by_id()` when an ID is available. `get_by_name()` performs one complete,
paginated list enumeration so duplicate titles are detected and reported with
`GraphAmbiguousMatchError`. The compatibility `get()` method still accepts the
Graph ID-or-title path segment.

## Filtered reads and lazy pagination

A mapping builds an escaped equality filter and sends it to the server by default:

```python
page = tasks.items.list(
    fields=("Title", "Status", "Owner"),
    filter={"Status": "Open"},
    top=50,
)

for item in page.items:
    print(item.id, item.fields)
```

Controlled expressions support `eq`, `ne`, `lt`, `gt`, `le`, `ge`, and
`startswith`. String literals are escaped by doubling embedded single quotes:

```python
from graphbridge.query import compare, startswith

query = compare("fields/Priority", "ge", 3) & startswith("fields/Title", "Q1")
for item in tasks.items.iter_all(filter=query, fields=("Title", "Priority")):
    print(item.id)
```

`iter_all()` is lazy. It requests the first page only when iteration starts and
requests subsequent pages only as needed. Graph's `@odata.nextLink` is forwarded
verbatim and the first-page parameters are not reapplied.

Local filtering is an explicit fallback, never an automatic behavior:

```python
open_items = tasks.items.iter_all(
    filter={"Status": "Open"},
    filter_mode="local",
)
```

The fallback evaluates each downloaded page once. Raw string filters are accepted
for server compatibility but cannot be used with `filter_mode="local"`.

## Item CRUD and eTags

All item operations return `ListItem` models except delete, which returns `None`.
`ListItem.raw` preserves the complete Graph payload and `ListItem.etag` preserves
the concurrency token when Graph returns it. An empty success is represented by
`response_empty=True`, which is distinct from an empty JSON object.

```python
# Illustrative write flow: use a disposable test list and a previously read item.
current = tasks.items.get("42", fields=("Title", "Status"))

created = tasks.items.create({"Title": "Review", "Status": "Open"})
updated = tasks.items.update(
    current.id,
    {"Status": "Done"},
    etag=current.etag,
)
```

When `If-Match` does not match, Graph's HTTP 412 is exposed as
`GraphPreconditionFailedError`; GraphBridge does not silently overwrite the item.
Deletion is intentionally omitted from runnable examples. The explicit
`tasks.items.delete(item_id, etag=...)` API is documented in the API reference.

## Lists and relationships

Lists can be enumerated lazily or created with initial columns on the stable
endpoint:

```python
for info in site.lists.iter_all(select=("id", "displayName", "list")):
    print(info.id, info.display_name, info.template)

# Illustrative schema creation; keep it in a disposable test site.
books = site.lists.create(
    "Books - GraphBridge Example",
    template="genericList",
    columns=(
        {"name": "Author", "text": {}},
        {"name": "PageCount", "number": {}},
    ),
)
```

`get_by_id(..., select=..., expand=...)` can retrieve v1.0 metadata and stable
relationships such as `columns`, `contentTypes`, `drive`, `items`, `operations`,
and `subscriptions`. Expanded values remain in `ListInfo.raw` and can be read with
`SharePointListResource.relationship()`.

## Columns and internal field names

Column definitions retain their known type facet (`text`, `choice`, `number`,
`dateTime`, `lookup`, and the other v1.0 facets) plus the entire raw payload, so
unknown future properties are not discarded.

```python
for column in tasks.columns.iter_all():
    print(column.id, column.display_name, column.name, column.column_type)

name_map = tasks.columns.display_name_map()
# {'Project Name': 'Project_x0020_Name', ...}

payload = tasks.columns.to_internal_fields({"Project Name": "Apollo"})

# Illustrative CRUD on a disposable list.
priority = tasks.columns.create(
    {"name": "Priority", "choice": {"choices": ["Low", "High"]}}
)
priority = tasks.columns.update(priority.id, {"required": True})
```

Column deletion is intentionally omitted from examples. Column creation,
updates, and deletion require schema-management authorization and should be
performed only in a reviewed administrative workflow.

Item reads and writes can opt into the schema map directly:

```python
item = tasks.items.create(
    {"Project Name": "Apollo"},
    field_names="display",
)
```

The map is cached and is invalidated after column mutations. Legacy
`GbList.encode_row()`/`decode_row()` remain available only as a temporary fallback;
the composed API does not use the character codec when schema is available.

## Generic JSON batch

The generic executor accepts any relative v1.0 `BatchRequest`. Item helpers expose
create, update, and delete batches:

```python
created = tasks.items.create_many([{"Title": f"Row {index}"} for index in range(45)])

updated = tasks.items.update_many(
    [
        ("10", {"Status": "Done"}, 'W/"etag-10"'),
        {"id": "11", "fields": {"Status": "Blocked"}, "etag": 'W/"etag-11"'},
    ],
    max_attempts=3,
    backoff_factor=0.5,
)

```

`delete_many()` is available for explicitly reviewed cleanup and supports
per-item eTags, but destructive batch code is intentionally omitted here.

Collections are split into batches of at most 20. Subrequest IDs are deterministic
global input indexes. Responses are correlated by ID even if Graph returns them
out of order. `BatchResult.results` retains one ordered `BatchItemResult` per input,
while `successes` and `failures` provide convenient partial-result views.

Only individual 408, 429, and 5xx subrequests are retried. Successful requests
are never replayed. Per-response `Retry-After` is honored (the longest delay is
used for a retry wave); otherwise exponential backoff applies. `max_attempts`
includes the initial call and prevents infinite loops.

## Safe synchronization: plan, review, apply

Choose a stable, unique business key stored in a SharePoint column. Do not use a
display title that users can freely rename. `plan()` performs reads only, validates
the key on both sides, and records field-level differences and eTags:

```python
source_rows = [
    {"ExternalId": "TASK-001", "Title": "Review", "Status": "Open"},
    {"ExternalId": "TASK-002", "Title": "Approve", "Status": "Done"},
]

plan = tasks.sync.plan(
    rows=source_rows,
    key_field="ExternalId",
    prune=False,
    dry_run=True,
)

print(plan.to_dict())  # review creates, PATCH updates, differences and reasons
preview = tasks.sync.apply(plan)  # no writes because the plan is dry-run
```

After review, build a non-dry plan (or override explicitly) and apply it:

```python
plan = tasks.sync.plan(rows=source_rows, key_field="ExternalId", prune=False)
result = tasks.sync.apply(plan, max_attempts=3, backoff_factor=0.5)

if result.failures:
    retry = tasks.sync.retry(result)  # contains only failed/deferred operations
```

`prune=False` is the default: remote-only rows are retained and appear as
`unchanged` with a reason. `prune=True` plans deletes, but apply always runs
create → PATCH update → delete. Any create failure defers every planned delete,
so a failed replacement cannot be preceded by data removal. Updates and deletes
carry the eTag read during planning and surface HTTP 412 conflicts for review.

Batch is used by default and inputs larger than 20 are chunked. `SyncResult`
retains successful creates, updates and deletes, unchanged operations, individual
errors, and correlated `results`. `result.retry_plan()` and `tasks.sync.retry()`
never include already successful operations. Re-plan rather than blindly retrying
an eTag conflict if remote state may have changed.

See [`examples/read_items.py`](examples/read_items.py),
[`examples/write_items.py`](examples/write_items.py), and
[`examples/synchronize.py`](examples/synchronize.py). Reads are the default;
writes and prune require separate exact environment confirmations and a list
whose verified name starts with `GraphBridge Integration - `.

## Delta changes

`delta()` follows every `@odata.nextLink` and returns the final opaque
`@odata.deltaLink` unchanged. Persist and pass the complete link, not a parsed
token:

```python
delta = tasks.items.delta(known_ids=local_item_ids, fields=("ExternalId", "Status"))
apply_to_local_cache(delta.created, delta.modified, delta.deleted)
save_cursor(delta.delta_link)

later = tasks.items.delta(link=load_cursor(), known_ids=local_item_ids)
```

Graph's delta feed does not label a non-deleted entry as create versus update.
`known_ids` lets GraphBridge classify against caller-owned state; without it,
entries are returned in `DeltaResult.unclassified` instead of being guessed.
The last occurrence of an item in a delta round wins, as required by Graph.

An expired cursor can return HTTP 410 with a resync strategy and a `Location`
header. GraphBridge raises `DeltaResetRequiredError` with `strategy` and the
unchanged `restart_link`; it never hides a full rebuild behind the incremental
call. The application must decide how to reconcile local and remote state.

## List item versions

```python
for version in tasks.items.versions("42"):
    print(version.id, version.last_modified_date_time, version.fields)

```

The stable restore action creates a new current version while retaining existing
history. Version history may be disabled or retained only for a finite period by
SharePoint configuration and tenant policy. Listing requires read permission;
restoring requires write permission. The mutating
`tasks.items.restore_version(item_id, version_id)` call is intentionally not in a
runnable example.

## Stable endpoints used

| Capability | Microsoft Graph v1.0 path |
|---|---|
| Resolve site | `GET /sites/{hostname}:/{path}` or `GET /sites/{site-id}` |
| Enumerate/create lists | `GET/POST /sites/{site-id}/lists` |
| List metadata by ID/title | `GET /sites/{site-id}/lists/{list-id-or-title}` |
| Item list/get/create | `GET/POST /sites/{site-id}/lists/{list-id}/items[...]` |
| Item field update | `PATCH /sites/{site-id}/lists/{list-id}/items/{item-id}/fields` |
| Item delete | `DELETE /sites/{site-id}/lists/{list-id}/items/{item-id}` |
| Item delta | `GET /sites/{site-id}/lists/{list-id}/items/delta` |
| List item versions | `GET /sites/{site-id}/lists/{list-id}/items/{item-id}/versions` |
| Restore item version | `POST /sites/{site-id}/lists/{list-id}/items/{item-id}/versions/{version-id}/restoreVersion` |
| Column CRUD | `GET/POST/PATCH/DELETE /sites/{site-id}/lists/{list-id}/columns[...]` |
| JSON batch | `POST /$batch` |

## Explicitly unsupported in this phase

- Microsoft Graph beta endpoints.
- Silent local-filter fallback.
- Automatic full-state recovery after delta HTTP 410.
- Reliable create/update classification from a delta feed without caller-provided
  known IDs; such entries remain explicitly unclassified.
- Document-set version creation/restoration (distinct from list item versions).
- List-item `$orderby`: the v1.0 list-items operation documents `$filter` and
  `$expand`, not `$orderby`; the generic `ODataQuery` supports `$orderby` only for
  endpoints that advertise it.
- Simulated endpoints for capabilities absent from Graph v1.0.

## Legacy compatibility

`GbAuth`, `GbSite`, and `GbList` retain their characterized constructors, cached
dictionary shapes, exceptions, and legacy result sections. Their shared transport
uses the composed client. In particular:

- `GbList.get_items_by_features()` is a local legacy filter;
- `upload()` emits `DeprecationWarning` with a caller-facing stack level and adapts
  to the new plan/apply service;
- `upload(force=True)` now PATCHes existing rows and reports them in the legacy
  `replaced` section; it intentionally never delete-and-recreates;
- `upload(delete=True)` maps to explicit pruning after creates and updates;
- `encode_row()`/`decode_row()` use a character codec rather than list schema;
- legacy CRUD and batch methods return dictionaries instead of typed models.

See [API reference](docs/api-reference.md) and
[public API characterization](docs/api-test-matrix.md) for the complete contracts.

## Security

Never commit client secrets or access tokens. Use environment variables or a
secret vault, grant least privilege, and test mutations only against disposable
lists. The included tests use simulated HTTP responses and block real network
connections.

Every HTTP call has a finite timeout. Retries have finite budgets; non-idempotent
outer POST/PATCH requests are not replayed automatically, and batch retry replays
only transient failed subrequests. Access tokens reflected by Graph in JSON or
response headers are redacted before values or exceptions reach the caller.

The integration suite is excluded by default. It requires `--run-integration`,
dedicated site/list IDs and verified names, and an exact environment confirmation.
Write integration has a second confirmation and removes only the item it created.
Review [the integration risk guide](tests/integration/README.md); do not run those
tests against production.

`publish.bat` defaults to build-only verification. TestPyPI/PyPI upload requires
an explicit mode plus `GRAPHBRIDGE_PUBLISH_CONFIRM`; never store Twine credentials
in the repository.

## Release roadmap

- 0.1.0: modern internal infrastructure with the legacy API still functional.
- 0.2.0: public composed API and formal legacy deprecation.
- 0.3.0: schema mapping, delta, versions, and advanced synchronization.
- 1.0.0: stable contract and possible legacy removal only after the migration
  guide and a completed deprecation window.
