# GraphBridge composed API reference

This reference covers the stable Microsoft Graph v1.0 composed API. The legacy
surface is characterized separately in `api-test-matrix.md`.

## Client and transport

`GraphBridgeClient` accepts a `TokenCredential` or an injected `GraphTransport`.
The default timeout is `(3.05, 30.0)`, the default retry budget is three, and the
default maximum single delay is 120 seconds. `max_retries` is limited to ten;
timeouts, backoff, and delay caps must be finite. Configured base URLs must use
HTTPS and terminate at a v1.0 root. POST/PATCH are not retried by default.

The package root exports the client, transport, credential protocol, common
models/errors and the legacy classes. Query and resource helpers are imported
from `graphbridge.query` and `graphbridge.resources`.

## Query helpers

`graphbridge.query` exports:

- `escape_odata_string(value)` and `odata_literal(value)`;
- `compare(field_name, operator, value)` for `eq`, `ne`, `lt`, `gt`, `le`, `ge`;
- `eq(...)`, `ne(...)`, and `startswith(...)` shortcuts;
- `filter_from_mapping(values, field_prefix=None)`;
- `fields_expand(fields=None)`;
- `FilterExpression`, including `&` and `|` composition;
- `ODataQuery(select=(), expand=(), filter=None, top=None, orderby=())`.

Identifiers are validated, scalar values are serialized, and embedded single
quotes are doubled. A raw string filter remains accepted for compatibility but
does not have a local predicate.

## Models

- `SiteInfo`, `ListInfo`, `ListItem`, and `ColumnInfo` expose common typed
  properties and preserve the entire Graph object in `raw`.
- `Page[T]` stores `items`, `next_link`, `delta_link`, and `raw`.
- `BatchItemResult[T]` stores the input index, request ID, status, value/error,
  attempts, response headers, and response body.
- `BatchResult[T]` stores ordered `results` plus partial `successes` and
  `failures` views.
- `SyncFieldDifference`, `SyncOperation`, `SyncPlan`, `SyncOperationResult`, and
  `SyncResult` retain review reasons, source/remote correlation, eTags, partial
  outcomes, and retry-plan construction. Plans and results expose `to_dict()`.
- `DeletedListItem` and `DeltaResult` retain delta tombstones and the exact final
  cursor. `ListItemVersion` retains fields, publisher metadata and raw properties.

## ListsResource

```python
list() -> Page[ListInfo]
iter_pages() -> Iterator[Page[ListInfo]]
iter_all() -> Iterator[ListInfo]
get(identifier) -> SharePointListResource       # compatibility ID-or-title
get_by_id(list_id, *, select=None, expand=None) -> SharePointListResource
get_by_name(name) -> SharePointListResource
create(display_name, *, template="genericList", columns=None, description=None)
bind(metadata) -> SharePointListResource
```

`get_by_name()` enumerates once and raises `GraphAmbiguousMatchError` if more than
one exact `displayName`/`name` match exists. `create()` uses the stable list-create
endpoint, which supports initial `columnDefinition` objects.

`SharePointListResource.metadata` exposes raw metadata. `relationship(name)` reads
a stable relationship already included by `$expand`; it never invents a separate
endpoint. `refresh(select=..., expand=...)` refreshes the bound metadata by ID.

## ListItemsResource

```python
list(*, fields=None, filter=None, top=None, select=None,
     filter_mode="server", field_names="internal") -> Page[ListItem]
iter_pages(...) -> Iterator[Page[ListItem]]
iter_all(...) -> Iterator[ListItem]
get(item_id, *, fields=None, field_names="internal") -> ListItem
create(fields, *, field_names="internal") -> ListItem
update(item_id, fields, *, etag=None, field_names="internal") -> ListItem
delete(item_id, *, etag=None) -> None
create_many(records, *, max_attempts=None, backoff_factor=None, sleep=None,
            field_names="internal") -> BatchResult[ListItem]
update_many(updates, fields=None, *, etag=None, etags=None, max_attempts=None,
            backoff_factor=None, sleep=None,
            field_names="internal") -> BatchResult[ListItem]
delete_many(item_ids, *, etag=None, etags=None, max_attempts=None,
            backoff_factor=None, sleep=None) -> BatchResult[str]
delta(*, link=None, token=None, fields=None, select=None, top=None,
      known_ids=None, field_names="internal") -> DeltaResult
versions(item_id) -> list[ListItemVersion]
restore_version(item_id, version_id) -> None
```

Mapping filters are equality comparisons on `fields/<name>` and run server-side.
`filter_mode="local"` is the explicit one-pass fallback. `field_names="display"`
uses the cached column schema map for selections, mapping filters, and writes.

`update_many()` accepts `(id, fields)`, `(id, fields, etag)`, mappings containing
`id`/`item_id`, `fields`, and optional `etag`, or parallel ID/field sequences.
`delete_many()` accepts IDs, `(id, etag)` pairs, or equivalent mappings. A scalar
`etag`, per-input sequence, or ID-to-eTag mapping can override record eTags.

`delta()` follows all pages and returns the exact `@odata.deltaLink`. A supplied
opaque link cannot be combined with initial query options. With `known_ids`,
non-deleted changes are classified as created or modified; without caller-owned
state they remain `unclassified`. Recognized delta-reset HTTP 410 responses raise
`DeltaResetRequiredError` rather than automatically starting a full enumeration.

## SyncService

Every `SharePointListResource` exposes `sync`:

```python
plan(*, rows, key_field, prune=False, dry_run=False,
     field_names="internal") -> SyncPlan
apply(plan, *, dry_run=None, use_batch=True, max_attempts=None,
      backoff_factor=None, sleep=None) -> SyncResult
retry(result, **apply_options) -> SyncResult
```

`plan()` performs no mutation. It materializes the source once, validates missing,
empty, non-scalar and duplicate keys before reading remotely, then reads the list
once and validates the same invariants remotely. Only source fields are compared;
extra server fields do not produce updates. Display-name mode uses the cached
authoritative column map.

Creates contain complete source fields. Updates contain only changed fields and
use PATCH with the planned eTag. Deletes are absent unless `prune=True`;
remote-only records otherwise appear under `unchanged` with a retention reason.

`apply()` runs create, update, then delete phases. It uses the generic batch layer
by default and retains every individual outcome. If any create fails, all deletes
are recorded as `deleteDeferred` and are not sent. `retry()` builds a new plan from
only failed/deferred outcomes, so successes are not repeated. An eTag 412 remains
visible and normally calls for a fresh plan.

## VersionsResource

```python
list(item_id) -> Page[ListItemVersion]
iter_pages(item_id) -> Iterator[Page[ListItemVersion]]
iter_all(item_id) -> Iterator[ListItemVersion]
versions(item_id) -> list[ListItemVersion]
get(item_id, version_id) -> ListItemVersion
restore_version(item_id, version_id) -> None
```

The endpoints are stable v1.0. SharePoint policy controls whether history exists
and how long it is retained. Restore uses `restoreVersion`, creates a new current
version, and preserves earlier versions.

## ColumnsResource

```python
list(*, select=None, top=None, orderby=None) -> Page[ColumnInfo]
iter_pages(...) -> Iterator[Page[ColumnInfo]]
iter_all(...) -> Iterator[ColumnInfo]
get(column_id) -> ColumnInfo                     # compatibility alias
get_by_id(column_id) -> ColumnInfo
get_by_name(name, *, include_display_name=True) -> ColumnInfo
create(definition) -> ColumnInfo
update(column_id, changes) -> ColumnInfo
delete(column_id) -> None
display_name_map(*, refresh=False) -> dict[str, str]
to_internal_fields(fields, *, strict=True) -> dict[str, object]
to_display_fields(fields) -> dict[str, object]
invalidate_schema() -> None
```

Internal `name` matches take priority over display-name matches. Duplicate display
names are rejected rather than overwritten in the map. Column mutations invalidate
the cached schema. `ColumnInfo.raw` is authoritative for properties GraphBridge
does not yet model explicitly.

## Generic batch API

`BatchRequest(id, method, url, headers={}, body=None, input_index=None)` requires a
relative v1.0 URL. `batch_payload()` enforces Graph's 20-request maximum.

`execute_batch(transport, requests, *, max_attempts=3, backoff_factor=0.5,
max_retry_delay=120.0, sleep=time.sleep)` automatically chunks larger inputs, uses case-insensitive IDs
for safe correlation, returns responses in request order, and retries only inner
408/429/5xx failures. Attempts are capped at eleven including the initial call.
Subrequests cannot carry an `Authorization` header; the outer request supplies
authentication. Batch response IDs must be present, unique, and known.

## Errors

HTTP failures retain the existing typed hierarchy, including
`GraphPreconditionFailedError` (412) and `GraphThrottlingError` (429).
`GraphAmbiguousMatchError` reports the duplicate candidate IDs.
`GraphInvalidResponseError` reports malformed or uncorrelatable Graph responses.
`GraphGoneError` represents a generic 410. `DeltaResetRequiredError` is reserved
for the two documented delta resync strategies and carries the exact Location
link. `SyncMissingKeyError` and `SyncDuplicateKeyError` reject unsafe plans.

## Stable endpoint notes

- Delta: `GET /sites/{site-id}/lists/{list-id}/items/delta`.
- Versions: `GET .../items/{item-id}/versions`.
- Restore: `POST .../items/{item-id}/versions/{version-id}/restoreVersion`.
- No beta endpoint, document-set version mutation, or automatic delta rebuild is
  used or simulated.
