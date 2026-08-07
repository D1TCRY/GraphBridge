# Public API characterization matrix

This matrix records the characterized public surface of GraphBridge 0.1.0.
Tests use synthetic Microsoft Graph responses and block real network
connections.

| Public API | Characterized behavior | Primary tests |
|---|---|---|
| `graphbridge.__all__` | Modern client/transport/models/errors plus legacy classes | `test_auth.py::test_public_package_exports` |
| `deduplicate_dicts()` | Stable first-occurrence de-duplication | `test_auth.py::test_deduplicate_dicts_preserves_first_occurrence` |
| `GbAuth(...)` | Validation and lazy initialization | `test_auth.py::test_auth_*` |
| `GbAuth.tenant_id` | Get, validate, invalidate auth caches | `test_auth.py::test_auth_*`, `test_changing_auth_input_*` |
| `GbAuth.client_id` | Get, validate, invalidate auth caches | `test_auth.py::test_auth_*`, `test_changing_auth_input_*` |
| `GbAuth.client_secret` | Get, validate, invalidate auth caches | `test_auth.py::test_auth_*`, `test_changing_auth_input_*` |
| `GbAuth.credential` | Lazy cached credential | `test_auth.py::test_credential_and_token_are_lazy_and_cached` |
| `GbAuth.token` | Lazy cached token and wrapped acquisition error | `test_auth.py::test_credential_*`, `test_token_acquisition_error_is_wrapped` |
| `GbAuth.headers` | Bearer authorization header | `test_auth.py::test_credential_and_token_are_lazy_and_cached` |
| `str(GbAuth)`, `repr(GbAuth)` | Secret/token redaction | `test_auth.py::test_auth_string_representations_redact_secret_and_token` |
| `GbSite(...)` | Direct auth or injected `GbAuth`; type validation | `test_site.py::test_site_can_be_built_*`, `test_site_rejects_*` |
| `GbSite.hostname` | Get and validation | `test_site.py::test_site_validates_location` |
| `GbSite.site_path` | Get and validation | `test_site.py::test_site_validates_location` |
| `GbSite.site_url` | Graph URL construction | `test_site.py::test_site_url_shape` |
| `GbSite.site_data` | GET shape, error, and cache | `test_site.py::test_site_data_*` |
| `GbSite.site_id` | ID extraction and missing-ID warning | `test_site.py::test_site_data_*`, `test_missing_site_id_*` |
| `str(GbSite)`, `repr(GbSite)` | Metadata rendering and inherited secret redaction | `test_site.py::test_site_url_shape`, `test_site_data_is_requested_once_and_cached` |
| `GbList(...)` | Direct site/auth construction and injected-site validation | `test_list_read.py::test_list_can_be_built_*`, `test_list_rejects_*` |
| `GbList.list_name` | Get and validation | `test_list_read.py::test_list_name_validation` |
| `GbList.encode_map`, `decode_map` | Legacy character mapping | `test_list_read.py::test_field_name_encoding_and_decoding` |
| `GbList.encode_row()`, `decode_row()` | Field-key round trip, including digits and punctuation | `test_list_read.py::test_field_name_encoding_and_decoding` |
| `GbList.list_url` | Site/list URL and title quoting | `test_list_read.py::test_list_url_encodes_spaces` |
| `GbList.list_data` | GET shape, error, and cache | `test_list_read.py::test_list_data_*` |
| `GbList.list_id` | ID extraction and missing-ID warning | `test_list_read.py::test_list_data_*`, `test_missing_list_id_*` |
| `GbList.list_items_all` | Expanded fields, `$top=200`, pagination and errors | `test_list_read.py::test_list_items_all_*` |
| `GbList.list_items` | Single-page response and errors | `test_list_read.py::test_list_items_*` |
| `GbList.list_rows` | Fields-only shape | `test_list_read.py::test_rows_and_ids_preserve_legacy_shapes` |
| `GbList.list_ids` | Item-ID shape | `test_list_read.py::test_rows_and_ids_preserve_legacy_shapes` |
| `GbList.list_fields` | First-row keys with exactly one list read | `test_list_read.py::test_list_fields_*` |
| `GbList.get_items_by_features()` | OR/AND/nested match, de-duplication and empty input | `test_list_read.py::test_get_items_by_features_*` |
| `GbList.update()` | Normalization, PATCH request/result/error shapes | `test_list_write.py::test_update_*` |
| `GbList.create()` | Normalization, POST request/result/error shapes | `test_list_write.py::test_create_*` |
| `GbList.delete()` | Normalization, DELETE request/result/error shapes | `test_list_write.py::test_delete_*` |
| `GbList.upload()` | Deprecated adapter, legacy result sections, PATCH force mode, safe prune order | `test_list_upload.py::test_upload_*` |
| `GbList.create_many()` | 20-request chunks and inner/outer failures | `test_list_batch.py::test_create_many_*` |
| `GbList.delete_many()` | ID mapping, `If-Match`, inner/outer failures | `test_list_batch.py::test_delete_many_*` |
| `str(GbList)`, `repr(GbList)` | List metadata and secret/token redaction | `test_list_read.py::test_list_repr_and_str_redact_secret_and_token` |

## Composed API introduced after characterization

The package root exports the preferred client, transport, common models/errors,
and the three legacy classes. Supporting query/resource helpers remain available
from their defining modules.

| Public API | Contract | Primary tests |
|---|---|---|
| `AccessToken`, `TokenCredential` | Structural credential protocols compatible with Azure Identity | `test_transport.py::test_token_is_requested_for_every_attempt_and_never_cached` |
| `GraphAuthenticator.scope`, `get_access_token()` | Graph `.default` scope, acquisition per attempt, safe failures | `test_transport.py::test_token_is_requested_for_every_attempt_and_never_cached`, `test_invalid_credential_and_acquisition_failure_are_safe` |
| `GraphTransport(...)`, `request()` and verb helpers | Shared session, configurable URL/timeout/User-Agent/JSON, empty responses | `test_transport.py::test_timeout_user_agent_json_and_absolute_url_are_forwarded_unchanged`, `test_successful_empty_and_invalid_json_responses_are_distinguished` |
| `GraphTransport.close()` and context manager | Close the owned or injected session | `test_transport.py::test_context_manager_closes_injected_session` |
| `GraphTransport` retry policy | Network, 429, 5xx, `Retry-After`; no default POST/PATCH replay | `test_transport.py::test_network_errors_retry_safe_requests_with_injected_sleep`, `test_retry_after_seconds_is_respected_and_token_is_renewed`, `test_server_errors_retry_safe_requests`, `test_non_idempotent_request_is_not_replayed_by_default` |
| Typed Graph exceptions | Status mapping for 401/403/404/409/412/429/5xx and safe messages | `test_transport.py::test_http_statuses_map_to_typed_errors_and_redact_tokens` |
| `GraphError`, `SiteInfo`, `ListInfo`, `ListItem`, `ColumnInfo`, `Page`, `BatchResult`, `SyncPlan`, `SyncResult` | Lightweight typed data and retained raw payloads | `test_primitives.py::test_lightweight_models_preserve_typed_fields_and_raw_payloads` |
| `ODataQuery.to_params()` | Stable v1.0 query serialization and `$top` validation | `test_primitives.py::test_odata_query_serializes_supported_v1_parameters` |
| `BatchRequest.to_payload()`, `chunks()`, `batch_payload()` | Shared JSON batch shape and 20-request limit | `test_primitives.py::test_batch_primitives_chunk_and_validate_graph_limit`, `test_composed_api.py::test_new_batch_primitives_report_partial_results` |
| `iter_pages()`, `iter_items()` | Typed pages/items and verbatim `@odata.nextLink` traversal | `test_composed_api.py::test_item_pagination_uses_exact_next_link_and_models`, `test_primitives.py::test_iter_items_uses_page_parser_and_exact_next_link` |
| `GraphBridgeClient(...)` | Credential or injected transport, one shared transport/session, safe representation | `test_composed_api.py::test_documented_composition_uses_one_client_transport_and_session`, `test_client_validation_representation_and_injected_transport` |
| `SitesResource.get_by_path()`, `get()`, `bind()` | Resolve or bind a `SiteResource` | `test_composed_api.py::test_documented_composition_uses_one_client_transport_and_session`, `test_binding_and_resource_types_require_no_network`, `test_invalid_resource_shapes_raise_graph_response_error` |
| `SiteResource.id`, `.lists` | Site model plus composed list navigation | `test_composed_api.py::test_documented_composition_uses_one_client_transport_and_session` |
| `ListsResource.get()`, `list()`, `iter_pages()`, `iter_all()`, `bind()` | Resolve, enumerate, paginate, or bind lists | `test_composed_api.py::test_documented_composition_uses_one_client_transport_and_session`, `test_list_and_column_discovery_support_pages_and_direct_lookup` |
| `SharePointListResource.id`, `.items`, `.columns`, `.sync`, `.versions` | List model plus composed subordinate resources | `test_composed_api.py::test_binding_and_resource_types_require_no_network`, `test_sync_v1.py` |
| `ListItemsResource.list()`, `iter_pages()`, `iter_all()`, `get()` | Typed item reads and exact pagination | `test_composed_api.py::test_item_pagination_uses_exact_next_link_and_models`, `test_columns_and_item_crud_return_models_and_preserve_request_shapes`, `test_invalid_resource_shapes_raise_graph_response_error` |
| `ListItemsResource.create()`, `update()`, `delete()` | Typed CRUD, including optional `If-Match` | `test_composed_api.py::test_columns_and_item_crud_return_models_and_preserve_request_shapes` |
| `ListItemsResource.create_many()`, `delete_many()` | Batch partial successes and typed failures | `test_composed_api.py::test_new_batch_primitives_report_partial_results` |
| `ColumnsResource.list()`, `iter_pages()`, `iter_all()`, `get()` | Typed column discovery and pagination | `test_composed_api.py::test_columns_and_item_crud_return_models_and_preserve_request_shapes`, `test_list_and_column_discovery_support_pages_and_direct_lookup` |
| `VersionsResource` | Paginated models, single-version read, stable restore action | `test_delta_versions_v1.py::test_versions_*`, `test_restore_version_*` |
| Legacy adapters over composed transport | Shared session/token renewal and equivalent raw data | `test_legacy_composition.py` |

## Intentionally characterized legacy ambiguities

- Tokens are cached as strings indefinitely after the first acquisition.
- Changing `hostname` or `site_path` after resolving a site does not invalidate
  cached site metadata.
- Changing `list_name` after resolving a list does not invalidate cached list
  metadata.
- `get_items_by_features()` performs one complete remote read per feature dict.
- `create()` wraps a non-dict member of a list in `{"fields": value}` instead of
  raising the documented `TypeError`.
- `upload(force=True)` is intentionally changed: it PATCHes existing items and
  preserves the legacy `replaced` result section without changing the item ID.
- `upload()` emits a caller-facing `DeprecationWarning` and applies prune only
  after successful creates and PATCH updates.
- `upload()` treats supplied identifiers as SharePoint item IDs; a newly created
  item can receive a different Graph-assigned ID.
- Normal CRUD and batch methods expose different success/failure record shapes.
- `graphbridge.__all__` now adds the modern public surface without removing any
  legacy name; explicit imports remain recommended.

## Stable CRUD, query, schema, and generic batch extension

All tests below use simulated responses; real network access is blocked.

| Public API | Contract | Primary tests |
|---|---|---|
| Controlled filter helpers and escaping | Valid identifiers/literals, quote doubling, comparisons, `startswith`, AND/OR | `test_query_items_v1.py::test_controlled_odata_query_*` |
| Item server filtering and local fallback | Server by default; explicit single-pass local mode | `test_query_items_v1.py::test_item_mapping_filter_*`, `test_local_filter_*` |
| Lazy item pagination | No request before iteration; exact nextLink | `test_query_items_v1.py::test_iter_all_is_lazy_*` |
| Empty/JSON item outcomes and eTags | `response_empty`, `If-Match`, typed 412 | `test_query_items_v1.py::test_item_crud_*`, `test_item_update_surfaces_412_*` |
| Explicit list ID/name lookup | Enumeration, exact match, duplicate detection | `test_lists_columns_v1.py::test_list_enumeration_*`, `test_duplicate_list_names_*` |
| List creation and relationships | Template, initial columns, metadata and expanded relations | `test_lists_columns_v1.py::test_create_list_*`, `test_list_metadata_*` |
| Column CRUD and type retention | GET/POST/PATCH/DELETE, known facet and unknown raw properties | `test_lists_columns_v1.py::test_column_crud_*` |
| Schema display-name mapping | Cached `displayName -> name`, item write translation | `test_lists_columns_v1.py::test_display_name_mapping_*` |
| Batch sizes | 1, 20, 21, and 45 operations with automatic chunking | `test_batch_v1.py::test_batch_create_chunks_*` |
| Batch partial/order correlation | Out-of-order subresponses mapped to ordered input outcomes | `test_batch_v1.py::test_partial_batch_response_*` |
| Batch retry | Only failed subrequests, inner `Retry-After`, finite budget | `test_batch_v1.py::test_batch_retries_*`, `test_batch_retry_budget_*` |
| Batch update/delete | Per-item eTags, typed successes and 412 failure | `test_batch_v1.py::test_update_many_and_delete_many_*` |

## Safe synchronization, delta, and versions extension

All tests use simulated responses and the global real-network blocker.

| Public API | Contract | Primary tests |
|---|---|---|
| `SyncService.plan()` | One remote read, key validation, create/update/delete/unchanged reasons, prune and dry-run | `test_sync_v1.py::test_sync_plan_*`, `test_empty_source_*` |
| `SyncService.apply()` | Create → PATCH → prune, eTags, partial outcomes and no mutation in dry-run | `test_sync_v1.py::test_sync_apply_*`, `test_sync_dry_run_*` |
| `SyncResult.retry_plan()`, `SyncService.retry()` | Retry only failed/deferred operations | `test_sync_v1.py::test_sync_retry_*` |
| Delete safety barrier | Failed create defers all deletes; no delete-before-create | `test_sync_v1.py::test_failed_create_defers_*` |
| Large synchronization | 45 creates split into 20/20/5 subrequests | `test_sync_v1.py::test_large_dataset_*` |
| `ListItemsResource.delta()` | One/many pages, exact cursor, tombstones, last occurrence and classification | `test_delta_versions_v1.py::test_delta_*` |
| Delta HTTP 410 | Explicit reset exception with strategy/Location; generic 410 remains generic | `test_delta_versions_v1.py::test_delta_410_*`, `test_generic_410_*` |
| Item versions | Paginated `ListItemVersion` models and stable `restoreVersion` POST | `test_delta_versions_v1.py::test_versions_*`, `test_restore_version_*` |
