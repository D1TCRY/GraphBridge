# Phase 2 architecture: authentication, HTTP transport and composition

> Historical phase note. The current architecture and release boundaries are
> documented in [`architecture.md`](architecture.md).

## Implemented boundary

`GraphBridgeClient` is the owner of one `GraphTransport`. Every resource created
from the client keeps a reference to that same client and transport:

```text
GraphBridgeClient
  -> SitesResource
       -> SiteResource
            -> ListsResource
                 -> SharePointListResource
                      -> ListItemsResource
                      -> ColumnsResource
                      -> SyncService
                      -> VersionsResource
```

No resource inherits from authentication. The original inheritance hierarchy is
retained only by `GbAuth`, `GbSite` and `GbList`, whose HTTP operations now adapt
to one shared composed client.

## Request flow

1. A resource builds a path relative to the configurable Graph v1.0 base URL.
2. `GraphTransport` asks the injected credential for the Graph `.default` scope.
3. The transport creates per-request authorization headers without caching the
   raw access token.
4. One reusable `requests.Session` sends the request with the configured timeout,
   User-Agent and uniform JSON settings.
5. A successful response becomes `None`, a JSON value, or a typed resource model.
6. A Graph failure becomes a typed exception with a `GraphError`. Tokens reflected
   by an error response are redacted.
7. Network failures, 429 and 5xx responses are retried only within the configured
   budget. Automatic retries default to idempotent methods; POST and PATCH are not
   replayed unless a caller explicitly opts in.
8. Pagination sends `@odata.nextLink` back to the transport verbatim and does not
   reapply the first page's query parameters. Absolute URLs outside the configured
   v1.0 root are rejected before acquiring or sending a token.

## Legacy migration status

Migrated to the shared transport:

- site and list metadata reads;
- item reads and pagination;
- item create, update and delete;
- JSON batch create and delete;
- credential reuse between injected `GbAuth`, `GbSite` and `GbList` instances.

Intentionally still legacy:

- inheritance and all public constructor signatures/properties;
- cached `GbAuth.token`, required by the characterized legacy contract;
- dictionary return shapes and legacy `RuntimeError`/partial-result behavior;
- field-name encoding and decoding;
- metadata cache invalidation quirks;
- different CRUD and batch outcome formats.

## Deferred work

- removal of the legacy inheritance hierarchy and cached token behavior.

## Stable CRUD/query/batch extension

The composed layer now also owns the phase-3 stable v1.0 surface:

- item CRUD plus typed create/update/delete batch helpers;
- controlled server-side OData filters and explicit one-pass local fallback;
- explicit list ID/name resolution, duplicate-title detection, list creation, and
  expanded relationship metadata;
- full list-column CRUD and cached `displayName -> name` schema mapping;
- a generic 20-request JSON batch executor with deterministic IDs, ordered
  correlation, partial outcomes, per-subrequest throttling retry, `Retry-After`,
  exponential backoff, and a finite attempt budget.

The schema map is authoritative for the composed API when
`field_names="display"`. The legacy character codec remains isolated in `GbList`.

## Safe synchronization, delta, and version extension

`SyncService.plan()` reads the item collection once and produces reviewable,
serializable operations keyed by a caller-selected business column. It rejects
missing and duplicate keys on either side. Updates contain changed fields only,
retain the read eTag, and are always PATCH operations.

`SyncService.apply()` uses three ordered phases: create, update, prune. Batch
subrequest correlation and retry come from the generic executor. Partial outcomes
remain visible, and `SyncResult.retry_plan()` includes only failed/deferred work.
A create failure installs a conservative barrier that defers all deletes.

`GbList.upload()` now builds this same typed plan from its legacy item-ID/row
pairs, applies it without delete-and-recreate, maps results back to the old
dictionary sections, and emits `DeprecationWarning`.

Delta traversal follows Graph links verbatim through a complete round and retains
the final cursor. Recognized HTTP 410 resync responses raise a dedicated exception
with the server strategy and Location link; rebuilding is an explicit caller
decision. Caller-owned known IDs are required to label non-deleted changes as
create versus update, because the feed itself reports latest state rather than an
operation label.

`VersionsResource` models and paginates retained list-item versions. Its stable
`restoreVersion` action creates a new current version without erasing history.

Still deferred:

- durable cursor/state storage and application-specific full-state rebuilding;
- conflict-resolution policy beyond exposing eTag 412;
- dependency ordering between business records in a batch;
- removal of the legacy inheritance hierarchy and cached token behavior.
