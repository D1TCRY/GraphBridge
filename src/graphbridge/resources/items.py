"""SharePoint list item CRUD, controlled queries and generic batch operations."""

from __future__ import annotations

import builtins
import time
from collections.abc import Callable, Collection, Iterator, Mapping, Sequence
from typing import TYPE_CHECKING, Any, Literal, cast
from urllib.parse import quote

from ..batch import BatchRequest, BatchResponse, execute_batch
from ..exceptions import DeltaResetRequiredError, GraphGoneError, GraphInvalidResponseError
from ..models import (
    BatchItemResult,
    BatchResult,
    DeletedListItem,
    DeltaResult,
    GraphError,
    ListItem,
    ListItemVersion,
    Page,
)
from ..pagination import iter_pages
from ..query import FilterExpression, ODataQuery, fields_expand, filter_from_mapping

if TYPE_CHECKING:
    from ..client import GraphBridgeClient
    from .lists import SharePointListResource

FilterValue = str | FilterExpression | Mapping[str, object] | None
FieldNameMode = Literal["internal", "display"]
FilterMode = Literal["server", "local"]


class ListItemsResource:
    """Uniform item operations for the stable Graph v1.0 listItem endpoints."""

    def __init__(self, client: GraphBridgeClient, sharepoint_list: SharePointListResource) -> None:
        self.client = client
        self.transport = client.transport
        self.sharepoint_list = sharepoint_list

    def list(
        self,
        *,
        fields: Sequence[str] | None = None,
        filter: FilterValue = None,
        top: int | None = None,
        select: Sequence[str] | None = None,
        filter_mode: FilterMode = "server",
        field_names: FieldNameMode = "internal",
    ) -> Page[ListItem]:
        """Return the first page; filtering is server-side unless explicitly local."""

        return next(
            self.iter_pages(
                fields=fields,
                filter=filter,
                top=top,
                select=select,
                filter_mode=filter_mode,
                field_names=field_names,
            )
        )

    def iter_pages(
        self,
        *,
        fields: Sequence[str] | None = None,
        filter: FilterValue = None,
        top: int | None = None,
        select: Sequence[str] | None = None,
        filter_mode: FilterMode = "server",
        field_names: FieldNameMode = "internal",
    ) -> Iterator[Page[ListItem]]:
        """Lazily yield pages while forwarding ``@odata.nextLink`` verbatim.

        ``filter_mode='local'`` is an explicit fallback. It omits ``$filter``
        from the request and evaluates a controlled filter once as each page is
        downloaded; it never performs a second list download.
        """

        if filter_mode not in {"server", "local"}:
            raise ValueError("filter_mode must be 'server' or 'local'")
        selected_fields = self._field_names(fields, field_names=field_names)
        expression = self._filter_expression(filter, field_names=field_names)
        if filter_mode == "local" and isinstance(expression, str):
            raise ValueError("raw string filters cannot be evaluated locally")
        query = ODataQuery(
            select=tuple(select or ()),
            expand=(fields_expand(selected_fields),),
            filter=expression if filter_mode == "server" else None,
            top=top,
        )
        pages = iter_pages(
            self.transport,
            self._base_path,
            params=query.to_params(),
            parser=ListItem.from_payload,
        )
        if filter_mode == "server" or expression is None:
            return pages
        assert isinstance(expression, FilterExpression)
        return self._locally_filtered_pages(pages, expression)

    def iter_all(
        self,
        *,
        fields: Sequence[str] | None = None,
        filter: FilterValue = None,
        top: int | None = None,
        select: Sequence[str] | None = None,
        filter_mode: FilterMode = "server",
        field_names: FieldNameMode = "internal",
    ) -> Iterator[ListItem]:
        """Lazily yield items without eagerly materializing all pages."""

        for page in self.iter_pages(
            fields=fields,
            filter=filter,
            top=top,
            select=select,
            filter_mode=filter_mode,
            field_names=field_names,
        ):
            yield from page.items

    def get(
        self,
        item_id: str,
        *,
        fields: Sequence[str] | None = None,
        field_names: FieldNameMode = "internal",
    ) -> ListItem:
        if not item_id:
            raise ValueError("item_id cannot be empty")
        selected_fields = self._field_names(fields, field_names=field_names)
        payload = self.transport.get(
            f"{self._base_path}/{quote(str(item_id), safe='')}",
            params={"$expand": fields_expand(selected_fields)},
        )
        return self._item_from_payload(payload, fallback_id=str(item_id))

    def create(
        self,
        fields: Mapping[str, Any],
        *,
        field_names: FieldNameMode = "internal",
    ) -> ListItem:
        values = self._write_fields(fields, field_names=field_names)
        payload = self.transport.post(self._base_path, json={"fields": values})
        return self._created_item(payload, values)

    def update(
        self,
        item_id: str,
        fields: Mapping[str, Any],
        *,
        etag: str | None = None,
        field_names: FieldNameMode = "internal",
    ) -> ListItem:
        if not item_id:
            raise ValueError("item_id cannot be empty")
        values = self._write_fields(fields, field_names=field_names)
        headers = {"If-Match": etag} if etag is not None else None
        payload = self.transport.patch(
            f"{self._base_path}/{quote(str(item_id), safe='')}/fields",
            json=values,
            headers=headers,
        )
        return self._updated_item(str(item_id), values, payload)

    def delete(self, item_id: str, *, etag: str | None = None) -> None:
        if not item_id:
            raise ValueError("item_id cannot be empty")
        headers = {"If-Match": etag} if etag is not None else None
        payload = self.transport.delete(
            f"{self._base_path}/{quote(str(item_id), safe='')}",
            headers=headers,
        )
        if payload is not None:
            raise GraphInvalidResponseError("Microsoft Graph item delete response must be empty")

    def delta(
        self,
        *,
        link: str | None = None,
        token: str | None = None,
        fields: Sequence[str] | None = None,
        select: Sequence[str] | None = None,
        top: int | None = None,
        known_ids: Collection[str] | None = None,
        field_names: FieldNameMode = "internal",
    ) -> DeltaResult:
        """Traverse one delta round and return its opaque continuation cursor.

        When ``known_ids`` is supplied, non-deleted entries are classified as
        ``created`` or ``modified`` against that caller-owned state. Without it,
        Graph cannot reliably distinguish the two, so they are returned in
        ``unclassified``. A reset-indicating HTTP 410 raises
        :class:`DeltaResetRequiredError`; no full enumeration is started
        automatically.
        """

        if link is not None and any(
            value is not None for value in (token, fields, select, top)
        ):
            raise ValueError("an opaque delta link cannot be combined with query options")
        if link is not None and (not isinstance(link, str) or not link):
            raise ValueError("delta link cannot be empty")
        selected_fields = self._field_names(fields, field_names=field_names)
        query = ODataQuery(
            select=tuple(select or ()),
            expand=(fields_expand(selected_fields),),
            top=top,
        )
        params: dict[str, str | int] = query.to_params() or {}
        if token is not None:
            if not isinstance(token, str) or not token:
                raise ValueError("delta token cannot be empty")
            params["token"] = token
        current_url: str | None = link or f"{self._base_path}/delta"
        current_params = None if link is not None else (params or None)
        known = {str(item_id) for item_id in known_ids} if known_ids is not None else None
        changes: dict[str, tuple[str, ListItem | DeletedListItem]] = {}
        delta_link: str | None = None
        page_count = 0

        while current_url is not None:
            try:
                payload = self.transport.get(current_url, params=current_params)
            except GraphGoneError as error:
                if error.error.code not in {
                    "resyncChangesApplyDifferences",
                    "resyncChangesUploadDifferences",
                }:
                    raise
                raise DeltaResetRequiredError(
                    error.error,
                    restart_link=self._header(error.response_headers, "Location"),
                    strategy=error.error.code,
                ) from error
            if not isinstance(payload, Mapping):
                raise GraphInvalidResponseError(
                    "Microsoft Graph delta response must be a JSON object"
                )
            raw_items = payload.get("value", [])
            if not isinstance(raw_items, list):
                raise GraphInvalidResponseError(
                    "Microsoft Graph delta response 'value' must be a list"
                )
            page_count += 1
            for raw_item in raw_items:
                if not isinstance(raw_item, Mapping):
                    raise GraphInvalidResponseError(
                        "Microsoft Graph delta entries must be JSON objects"
                    )
                item_id = str(raw_item.get("id", ""))
                if not item_id:
                    raise GraphInvalidResponseError(
                        "Microsoft Graph delta entry does not contain an id"
                    )
                if item_id in changes:
                    del changes[item_id]
                if "deleted" in raw_item:
                    changes[item_id] = (
                        "deleted",
                        DeletedListItem.from_payload(raw_item),
                    )
                else:
                    item = ListItem.from_payload(raw_item)
                    kind = (
                        "unclassified"
                        if known is None
                        else ("modified" if item_id in known else "created")
                    )
                    changes[item_id] = (kind, item)

            next_link = self._delta_link(payload.get("@odata.nextLink"))
            page_delta_link = self._delta_link(payload.get("@odata.deltaLink"))
            if next_link is not None and page_delta_link is not None:
                raise GraphInvalidResponseError(
                    "Microsoft Graph delta page cannot contain both nextLink and deltaLink"
                )
            if next_link is not None:
                current_url = next_link
                current_params = None
                continue
            if page_delta_link is None:
                raise GraphInvalidResponseError(
                    "the final Microsoft Graph delta page must contain a deltaLink"
                )
            delta_link = page_delta_link
            current_url = None

        created: builtins.list[ListItem] = []
        modified: builtins.list[ListItem] = []
        deleted: builtins.list[DeletedListItem] = []
        unclassified: builtins.list[ListItem] = []
        for kind, value in changes.values():
            if kind == "deleted":
                assert isinstance(value, DeletedListItem)
                deleted.append(value)
            else:
                assert isinstance(value, ListItem)
                if kind == "created":
                    created.append(value)
                elif kind == "modified":
                    modified.append(value)
                else:
                    unclassified.append(value)
        return DeltaResult(
            created=created,
            modified=modified,
            deleted=deleted,
            unclassified=unclassified,
            delta_link=delta_link,
            pages=page_count,
        )

    def versions(self, item_id: str) -> builtins.list[ListItemVersion]:
        """Return every retained stable v1.0 version for one item."""

        return self.sharepoint_list.versions.versions(item_id)

    def restore_version(self, item_id: str, version_id: str) -> None:
        """Restore a retained version as a new current version."""

        return self.sharepoint_list.versions.restore_version(item_id, version_id)

    def create_many(
        self,
        records: Sequence[Mapping[str, Any]],
        *,
        max_attempts: int | None = None,
        backoff_factor: float | None = None,
        sleep: Callable[[float], None] | None = None,
        field_names: FieldNameMode = "internal",
    ) -> BatchResult[ListItem]:
        values = [self._write_fields(record, field_names=field_names) for record in records]
        requests = [
            BatchRequest(
                id=str(index),
                method="POST",
                url=self._base_path,
                body={"fields": record},
                input_index=index,
            )
            for index, record in enumerate(values)
        ]
        responses = self._execute_batch(
            requests,
            max_attempts=max_attempts,
            backoff_factor=backoff_factor,
            sleep=sleep,
        )
        return self._create_batch_result(responses, values)

    def update_many(
        self,
        updates: Sequence[Any],
        fields: Sequence[Mapping[str, Any]] | None = None,
        *,
        etag: str | None = None,
        etags: Sequence[str | None] | Mapping[str, str] | str | None = None,
        max_attempts: int | None = None,
        backoff_factor: float | None = None,
        sleep: Callable[[float], None] | None = None,
        field_names: FieldNameMode = "internal",
    ) -> BatchResult[ListItem]:
        """Batch updates from records or parallel ``ids``/``fields`` sequences.

        A record can be ``(id, fields)``, ``(id, fields, etag)`` or a mapping
        containing ``id``/``item_id``, ``fields`` and optional ``etag``.
        """

        if etag is not None and etags is not None:
            raise ValueError("etag and etags cannot both be supplied")
        normalized = self._normalize_updates(updates, fields)
        configured_etags = etags if etags is not None else etag
        prepared: builtins.list[tuple[str, dict[str, Any], str | None]] = []
        for index, (item_id, values, record_etag) in enumerate(normalized):
            override = self._etag_at(configured_etags, index, item_id)
            prepared.append(
                (
                    item_id,
                    self._write_fields(values, field_names=field_names),
                    override if override is not None else record_etag,
                )
            )
        requests: builtins.list[BatchRequest] = []
        for index, (item_id, values, item_etag) in enumerate(prepared):
            headers = {"If-Match": item_etag} if item_etag is not None else {}
            requests.append(
                BatchRequest(
                    id=str(index),
                    method="PATCH",
                    url=f"{self._base_path}/{quote(item_id, safe='')}/fields",
                    headers=headers,
                    body=values,
                    input_index=index,
                )
            )
        responses = self._execute_batch(
            requests,
            max_attempts=max_attempts,
            backoff_factor=backoff_factor,
            sleep=sleep,
        )
        return self._update_batch_result(responses, prepared)

    def delete_many(
        self,
        item_ids: Sequence[Any] | str | int,
        *,
        etag: str | None = None,
        etags: Sequence[str | None] | Mapping[str, str] | str | None = None,
        max_attempts: int | None = None,
        backoff_factor: float | None = None,
        sleep: Callable[[float], None] | None = None,
    ) -> BatchResult[str]:
        """Delete items in correlated chunks of at most 20 subrequests."""

        if etag is not None and etags is not None:
            raise ValueError("etag and etags cannot both be supplied")
        configured_etags = etags if etags is not None else etag
        normalized = self._normalize_deletes(item_ids)
        prepared: builtins.list[tuple[str, str | None]] = []
        for index, (item_id, record_etag) in enumerate(normalized):
            override = self._etag_at(configured_etags, index, item_id)
            prepared.append((item_id, override if override is not None else record_etag))
        requests: builtins.list[BatchRequest] = []
        for index, (item_id, item_etag) in enumerate(prepared):
            headers = {"If-Match": item_etag} if item_etag is not None else {}
            requests.append(
                BatchRequest(
                    id=str(index),
                    method="DELETE",
                    url=f"{self._base_path}/{quote(item_id, safe='')}",
                    headers=headers,
                    input_index=index,
                )
            )
        responses = self._execute_batch(
            requests,
            max_attempts=max_attempts,
            backoff_factor=backoff_factor,
            sleep=sleep,
        )
        return self._delete_batch_result(responses, prepared)

    @property
    def _base_path(self) -> str:
        site_id = quote(self.sharepoint_list.site.id, safe=",")
        list_id = quote(self.sharepoint_list.id, safe="")
        return f"/sites/{site_id}/lists/{list_id}/items"

    def _field_names(
        self, fields: Sequence[str] | None, *, field_names: FieldNameMode
    ) -> Sequence[str] | None:
        if fields is None or field_names == "internal":
            return fields
        if field_names != "display":
            raise ValueError("field_names must be 'internal' or 'display'")
        mapping = self.sharepoint_list.columns.display_name_map()
        internal = set(mapping.values())
        selected: builtins.list[str] = []
        for name in fields:
            if name in internal:
                selected.append(name)
            elif name in mapping:
                selected.append(mapping[name])
            else:
                raise KeyError(f"unknown SharePoint column displayName: {name!r}")
        return selected

    def _write_fields(
        self, fields: Mapping[str, Any], *, field_names: FieldNameMode
    ) -> dict[str, Any]:
        if not isinstance(fields, Mapping):
            raise TypeError("fields must be a mapping")
        if field_names == "internal":
            return dict(fields)
        if field_names == "display":
            return self.sharepoint_list.columns.to_internal_fields(fields)
        raise ValueError("field_names must be 'internal' or 'display'")

    def _filter_expression(
        self, value: FilterValue, *, field_names: FieldNameMode
    ) -> str | FilterExpression | None:
        if not isinstance(value, Mapping):
            return value
        fields: Mapping[str, object] = value
        if field_names == "display":
            fields = self.sharepoint_list.columns.to_internal_fields(value)
        elif field_names != "internal":
            raise ValueError("field_names must be 'internal' or 'display'")
        return filter_from_mapping(fields, field_prefix="fields")

    @staticmethod
    def _locally_filtered_pages(
        pages: Iterator[Page[ListItem]], expression: FilterExpression
    ) -> Iterator[Page[ListItem]]:
        for page in pages:
            yield Page(
                items=[item for item in page.items if expression.matches(item.raw)],
                next_link=page.next_link,
                delta_link=page.delta_link,
                raw=page.raw,
            )

    def _execute_batch(
        self,
        requests: Sequence[BatchRequest],
        *,
        max_attempts: int | None,
        backoff_factor: float | None,
        sleep: Callable[[float], None] | None,
    ) -> builtins.list[BatchResponse]:
        attempts = (
            max_attempts
            if max_attempts is not None
            else int(getattr(self.transport, "max_retries", 2)) + 1
        )
        factor = (
            backoff_factor
            if backoff_factor is not None
            else float(getattr(self.transport, "backoff_factor", 0.5))
        )
        sleeper = (
            sleep
            if sleep is not None
            else cast(Callable[[float], None], getattr(self.transport, "_sleep", time.sleep))
        )
        return execute_batch(
            self.transport,
            requests,
            max_attempts=attempts,
            backoff_factor=factor,
            max_retry_delay=float(getattr(self.transport, "max_retry_delay", 120.0)),
            sleep=sleeper,
        )

    @staticmethod
    def _item_from_payload(payload: Any, *, fallback_id: str = "") -> ListItem:
        if not isinstance(payload, Mapping):
            raise GraphInvalidResponseError("Microsoft Graph item response must be a JSON object")
        return ListItem.from_payload(payload, fallback_id=fallback_id)

    @classmethod
    def _created_item(cls, payload: Any, fields: Mapping[str, Any]) -> ListItem:
        if payload is None:
            return ListItem(id="", fields=dict(fields), response_empty=True)
        item = cls._item_from_payload(payload)
        return cls._with_fallback_fields(item, fields)

    @classmethod
    def _updated_item(
        cls, item_id: str, fields: Mapping[str, Any], payload: Any
    ) -> ListItem:
        if payload is None:
            return ListItem(
                id=item_id,
                fields=dict(fields),
                response_empty=True,
            )
        if not isinstance(payload, Mapping):
            raise GraphInvalidResponseError("Microsoft Graph item update response must be a JSON object")
        if "fields" in payload:
            return ListItem.from_payload(payload, fallback_id=item_id)
        field_values = {
            str(name): value
            for name, value in payload.items()
            if not str(name).startswith("@odata.") and name != "eTag"
        }
        return ListItem(
            id=item_id,
            fields=field_values,
            etag=cls._optional_string(payload.get("eTag") or payload.get("@odata.etag")),
            raw=dict(payload),
        )

    @staticmethod
    def _with_fallback_fields(item: ListItem, fields: Mapping[str, Any]) -> ListItem:
        if item.fields or not fields:
            return item
        return ListItem(
            id=item.id,
            fields=dict(fields),
            etag=item.etag,
            created_date_time=item.created_date_time,
            last_modified_date_time=item.last_modified_date_time,
            web_url=item.web_url,
            response_empty=item.response_empty,
            raw=item.raw,
        )

    def _create_batch_result(
        self, responses: Sequence[BatchResponse], records: Sequence[Mapping[str, Any]]
    ) -> BatchResult[ListItem]:
        successes: builtins.list[ListItem] = []
        failures: builtins.list[GraphError] = []
        results: builtins.list[BatchItemResult[ListItem]] = []
        for response, record in zip(responses, records, strict=True):
            if response.succeeded:
                try:
                    value = self._created_item(response.body, record)
                    value = self._with_response_etag(value, response.headers)
                except GraphInvalidResponseError as error:
                    graph_error = self._invalid_batch_error(response, str(error))
                    failures.append(graph_error)
                    results.append(self._failure_result(response, graph_error))
                    continue
                successes.append(value)
                results.append(self._success_result(response, value))
            else:
                graph_error = response.to_error()
                failures.append(graph_error)
                results.append(self._failure_result(response, graph_error))
        return BatchResult(successes=successes, failures=failures, results=results)

    def _update_batch_result(
        self,
        responses: Sequence[BatchResponse],
        updates: Sequence[tuple[str, Mapping[str, Any], str | None]],
    ) -> BatchResult[ListItem]:
        successes: builtins.list[ListItem] = []
        failures: builtins.list[GraphError] = []
        results: builtins.list[BatchItemResult[ListItem]] = []
        for response, (item_id, fields, _etag) in zip(responses, updates, strict=True):
            if response.succeeded:
                try:
                    value = self._updated_item(item_id, fields, response.body)
                    value = self._with_response_etag(value, response.headers)
                except GraphInvalidResponseError as error:
                    graph_error = self._invalid_batch_error(response, str(error))
                    failures.append(graph_error)
                    results.append(self._failure_result(response, graph_error))
                    continue
                successes.append(value)
                results.append(self._success_result(response, value))
            else:
                graph_error = response.to_error()
                failures.append(graph_error)
                results.append(self._failure_result(response, graph_error))
        return BatchResult(successes=successes, failures=failures, results=results)

    def _delete_batch_result(
        self,
        responses: Sequence[BatchResponse],
        deletes: Sequence[tuple[str, str | None]],
    ) -> BatchResult[str]:
        successes: builtins.list[str] = []
        failures: builtins.list[GraphError] = []
        results: builtins.list[BatchItemResult[str]] = []
        for response, (item_id, _etag) in zip(responses, deletes, strict=True):
            if response.succeeded:
                successes.append(item_id)
                results.append(self._success_result(response, item_id))
            else:
                graph_error = response.to_error()
                failures.append(graph_error)
                results.append(self._failure_result(response, graph_error))
        return BatchResult(successes=successes, failures=failures, results=results)

    @staticmethod
    def _success_result(response: BatchResponse, value: Any) -> BatchItemResult[Any]:
        return BatchItemResult(
            input_index=response.request.input_index or 0,
            request_id=response.request.id,
            status_code=response.status_code,
            value=value,
            attempts=response.attempts,
            response_headers=response.headers,
            response_body=response.body,
        )

    @staticmethod
    def _failure_result(
        response: BatchResponse, error: GraphError
    ) -> BatchItemResult[Any]:
        return BatchItemResult(
            input_index=response.request.input_index or 0,
            request_id=response.request.id,
            status_code=response.status_code,
            error=error,
            attempts=response.attempts,
            response_headers=response.headers,
            response_body=response.body,
        )

    @staticmethod
    def _invalid_batch_error(response: BatchResponse, message: str) -> GraphError:
        return GraphError(
            code="invalidResponse",
            message=message,
            status_code=response.status_code,
            inner_error={
                "batch_request_id": response.request.id,
                "input_index": response.request.input_index,
            },
        )

    @classmethod
    def _with_response_etag(
        cls, item: ListItem, headers: Mapping[str, Any]
    ) -> ListItem:
        if item.etag is not None:
            return item
        etag = cls._header(headers, "ETag")
        if etag is None:
            return item
        return ListItem(
            id=item.id,
            fields=item.fields,
            etag=etag,
            created_date_time=item.created_date_time,
            last_modified_date_time=item.last_modified_date_time,
            web_url=item.web_url,
            response_empty=item.response_empty,
            raw=item.raw,
        )

    @staticmethod
    def _normalize_updates(
        updates: Sequence[Any], fields: Sequence[Mapping[str, Any]] | None
    ) -> builtins.list[tuple[str, Mapping[str, Any], str | None]]:
        if fields is not None:
            if len(updates) != len(fields):
                raise ValueError("update ids and fields must have the same length")
            return [(str(item_id), values, None) for item_id, values in zip(updates, fields)]
        normalized: builtins.list[tuple[str, Mapping[str, Any], str | None]] = []
        for update in updates:
            if isinstance(update, Mapping):
                item_id = update.get("id", update.get("item_id"))
                values = update.get("fields")
                record_etag = update.get("etag")
            elif isinstance(update, (tuple, builtins.list)) and len(update) in {2, 3}:
                item_id, values = update[0], update[1]
                record_etag = update[2] if len(update) == 3 else None
            else:
                raise TypeError("each batch update must contain id, fields and optional etag")
            if item_id is None or not str(item_id):
                raise ValueError("batch update item ids cannot be empty")
            if not isinstance(values, Mapping):
                raise TypeError("batch update fields must be mappings")
            normalized.append(
                (str(item_id), values, str(record_etag) if record_etag is not None else None)
            )
        return normalized

    @staticmethod
    def _normalize_deletes(
        values: Sequence[Any] | str | int,
    ) -> builtins.list[tuple[str, str | None]]:
        source: Sequence[Any] = [values] if isinstance(values, (str, int)) else values
        normalized: builtins.list[tuple[str, str | None]] = []
        for value in source:
            record_etag: Any = None
            if isinstance(value, Mapping):
                item_id = value.get("id", value.get("item_id"))
                record_etag = value.get("etag")
            elif isinstance(value, (tuple, builtins.list)) and len(value) == 2:
                item_id, record_etag = value
            else:
                item_id = value
            if item_id is None or not str(item_id):
                raise ValueError("batch delete item ids cannot be empty")
            normalized.append(
                (str(item_id), str(record_etag) if record_etag is not None else None)
            )
        return normalized

    @staticmethod
    def _etag_at(
        etags: Sequence[str | None] | Mapping[str, str] | str | None,
        index: int,
        item_id: str,
    ) -> str | None:
        if etags is None:
            return None
        if isinstance(etags, str):
            return etags
        if isinstance(etags, Mapping):
            return etags.get(item_id)
        if len(etags) <= index:
            raise ValueError("not enough etags were supplied")
        return etags[index]

    @staticmethod
    def _header(headers: Mapping[str, Any], name: str) -> str | None:
        expected = name.casefold()
        for key, value in headers.items():
            if str(key).casefold() == expected:
                return str(value)
        return None

    @staticmethod
    def _optional_string(value: object) -> str | None:
        return str(value) if value is not None else None

    @staticmethod
    def _delta_link(value: object) -> str | None:
        if value is None:
            return None
        if not isinstance(value, str) or not value:
            raise GraphInvalidResponseError(
                "Microsoft Graph returned an invalid delta pagination link"
            )
        return value
