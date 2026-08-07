# Architecture

GraphBridge is an app-only, synchronous Python client for the stable Microsoft
Graph `v1.0` SharePoint site, list, column, and list-item endpoints. The modern
API composes small resource objects around one client and one HTTP transport;
the inheritance-based API remains an adapter for a time-bounded migration.

## Layering

```text
TokenCredential
    -> GraphAuthenticator
        -> GraphTransport
            -> GraphBridgeClient
                -> SitesResource -> SiteResource
                    -> ListsResource -> SharePointListResource
                        -> items / columns / versions / sync
```

- `auth.py` defines the structural `TokenCredential` contract and acquires a
  token for every HTTP attempt. Azure Identity owns token caching and renewal.
- `transport.py` owns the session, authentication header, finite timeout,
  retry policy, response parsing, token redaction, URL confinement, and typed
  HTTP errors.
- `client.py` owns the transport lifecycle and exposes `client.sites`.
- `resources/` implements v1.0 operations and resource composition.
- `query.py`, `pagination.py`, and `batch.py` provide controlled OData,
  continuation traversal, and the 20-subrequest JSON batch protocol.
- `models.py` contains dependency-free result models; `exceptions.py` contains
  the public error hierarchy.
- `graph_bridge.py` and `legacy.py` preserve the characterized legacy surface.

Dependencies point inward: resources depend on transport, models, and helpers;
the transport does not depend on resource objects. Circular runtime imports are
avoided with local composition imports and `TYPE_CHECKING` imports.

## Authentication and transport invariants

`GraphBridgeClient` accepts either a `TokenCredential` or an already configured
`GraphTransport`, never both. No credential or token is rendered by client,
transport, authenticator, legacy object, network error, or Graph error
representations. If Graph reflects the active access token in a successful JSON
body, error body, nested error, or response header, the transport replaces it
before returning or storing the value.

The default base is `https://graph.microsoft.com/v1.0`. Configured bases must be
absolute HTTPS URLs, contain no user information/query/fragment, and end at a
`v1.0` root. Absolute pagination and delta links must remain on the configured
origin and below that root. This prevents a service-controlled continuation link
from forwarding a bearer token to another origin or to `/beta`.

Every request has a finite positive timeout. GET, HEAD, OPTIONS, PUT, and DELETE
may retry network failures, HTTP 429, and HTTP 5xx up to the configured finite
budget. POST and PATCH are not replayed by the outer transport by default.
`Retry-After` is honored within a finite delay cap; otherwise capped exponential
backoff is used. Batch POSTs are
never replayed as a whole: only transient 408, 429, or 5xx subrequests are retried.

## Resource composition and data flow

A normal flow resolves a site, then a list, then uses its subordinate resources:

```python
site = client.sites.get_by_path(hostname=hostname, path=site_path)
tasks = site.lists.get_by_id(list_id)
items = tasks.items.iter_all(fields=("Title", "Status"))
```

IDs are preferred for mutation. Name lookups enumerate all pages and reject an
ambiguous exact name rather than selecting an arbitrary list. All collection
iterators are lazy and forward `@odata.nextLink` exactly; first-page query
parameters are not reapplied to continuation links.

Typed models preserve common fields and the original payload. Unknown Graph
properties are not discarded. Invalid collection, item, batch, delta, and
version shapes raise `GraphInvalidResponseError` rather than being guessed.

## Query, schema, and batch boundaries

Controlled OData builders validate property paths and escape literals. Mapping
filters run server-side by default. Local filtering is an explicit fallback and
never triggers a second download. The list-items v1.0 surface intentionally does
not expose `$orderby`.

The authoritative field-name translation is the cached list column schema
(`displayName -> name`). Column mutations invalidate that cache. The legacy
character codec is isolated in `GbList` and is not used by the composed API.

JSON batch input is split at 20 operations. IDs are unique case-insensitively,
responses are correlated even when returned out of order, and every result keeps
its original input index, attempts, status, headers, and body.

## Synchronization safety

`SyncService.plan()` performs reads only. It materializes the source once,
requires a non-empty unique scalar business key on both sides, compares only
source-owned fields, and records field differences and remote eTags.

`prune=False` is the default. With `prune=True`, deletes are merely planned until
`apply()` is called. Apply order is create, PATCH update, then delete. A create
failure defers all deletes. Updates and deletes use planned eTags when available,
and HTTP 412 remains visible to the caller. Dry-run never mutates. Retry plans
contain only failed or safely deferred operations.

Delta processing keeps the final opaque cursor, lets the last event for an item
win, and never invents create/update classification without caller-owned known
IDs. A recognized HTTP 410 raises `DeltaResetRequiredError`; full reconciliation
is an application decision.

## Legacy boundary

`GbAuth`, `GbSite`, and `GbList` remain available and share the modern transport
when composed. Characterized constructors, properties, dictionary result shapes,
and legacy exceptions remain. `GbList.upload()` emits `DeprecationWarning` and
adapts to plan/apply; force mode PATCHes instead of delete-and-recreate.

The legacy cached token and heuristic field-name codec are compatibility
liabilities and must not be copied into new code. See `migration.md`.

## Verification and release boundary

Unit tests block real socket connections. Integration tests are marked and
skipped unless `--run-integration` plus dedicated-environment confirmations are
provided. They are not a release-suite default. Packaging uses the `src` layout
and includes only `graphbridge*` packages. `publish.bat` builds by default and
requires an explicit mode and environment confirmation before any upload.

GraphBridge does not implement beta endpoints, automatic permission assignment,
durable delta cursor storage, automatic full rebuild after HTTP 410, or a generic
business conflict-resolution policy.
