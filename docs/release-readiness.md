# Release-readiness conformity matrix

This matrix is based on the current files and local test results, not on prior
completion claims. Status values are `complete`, `partial`, `missing`, or
`blocked`.

| Area | Status | Files and implemented APIs | Existing tests | Remaining gap | Required intervention |
|---|---|---|---|---|---|
| Characterization tests | complete | `graph_bridge.py`, `legacy.py`; `GbAuth`, `GbSite`, `GbList`, legacy CRUD/batch/upload/codec | `test_auth.py`, `test_site.py`, `test_list_read.py`, `test_list_write.py`, `test_list_batch.py`, `test_list_upload.py`, `test_legacy_composition.py` | Legacy ambiguities remain intentionally documented | Keep tests for the migration window; remove legacy only after an announced deprecation window |
| Authentication and HTTP transport | complete | `auth.py`, `transport.py`, `_version.py`; `TokenCredential`, `GraphAuthenticator`, `GraphTransport`, finite timeout/retry, v1.0 URL confinement, redaction | `test_transport.py`, `test_legacy_composition.py` | No asynchronous transport | None for the current synchronous contract |
| Class composition | complete | `client.py`, `resources/sites.py`, `resources/lists.py`; `GraphBridgeClient -> sites -> lists -> items/columns/versions/sync` | `test_composed_api.py`, `test_legacy_composition.py` | Legacy inheritance still exists | Continue adapter isolation until eventual removal |
| Item CRUD and query | complete | `resources/items.py`, `query.py`, `pagination.py`, `batch.py`; typed CRUD, OData, lazy pages, 20-item batches, eTags | `test_query_items_v1.py`, `test_batch_v1.py`, `test_composed_api.py`, `test_transport.py` | List-item `$orderby` is not exposed because v1.0 does not advertise it | Keep unsupported rather than use beta |
| Lists and schema | complete | `resources/lists.py`, `resources/columns.py`; ID/name lookup, list create, relationships, column CRUD, schema name map | `test_lists_columns_v1.py`, `test_composed_api.py` | Schema mutation requires stronger authorization | Deploy with documented least privilege and dedicated administration identity |
| Safe synchronization | complete | `resources/sync.py`, sync models; plan/apply/retry, dry-run, explicit prune, eTag, delete barrier | `test_sync_v1.py`, `test_list_upload.py` | Business-specific conflict policy and durable state are intentionally external | Re-plan on conflicts; keep persistence in the application |
| Models and errors | complete | `models.py`, `exceptions.py`; typed payloads/results and HTTP/domain exceptions | `test_primitives.py`, transport/batch/delta/sync tests | Models intentionally retain arbitrary raw Graph data | Callers must treat raw business fields as sensitive application data |
| Security and permissions | partial | `.gitignore`, `.env.example`, `transport.py`, guarded examples, `docs/permissions.md`, hardened `publish.bat` | `test_repository_hygiene.py`, `test_transport.py`, representation tests | A local `.env` exists and Git metadata is absent, so whether it was ever tracked cannot be verified | Restore/attach `.git`, run `git ls-files`/history secret scan, rotate any credential if tracking is discovered |
| Tests and quality | complete | `pyproject.toml`, `tests/unit`, `tests/integration`; strict markers, network blocker, opt-in safety gates | Unit suite plus skipped integration smoke/write tests | Integration tests are deliberately not run in this audit | Run read-only then write canaries only in the dedicated environment before rollout |
| Migration and release | partial | `README.md`, `docs/architecture.md`, `docs/migration.md`, examples, `pyproject.toml`, `publish.bat` | hygiene/version tests; local build and artifact audit | Git diff/provenance is blocked by the missing `.git`; no real integration canary was authorized | Restore Git metadata, review the real diff, run opt-in canaries, then make a dedicated version/release commit |

## Release recommendation

The working and recommended initial release version is `0.1.0`. It identifies
the first migration release with modern internals and a functioning legacy API;
later roadmap milestones do not change the current package version. No push,
tag, package-index upload, or publication is part of this preparation.
