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
    """Read and mutate SharePoint list items.

    The resource combines typed CRUD, controlled queries, lazy pagination, delta
    traversal, version shortcuts, field-name translation, and batch helpers. It
    is available as ``sharepoint_list.items`` and is the primary content API.

    Args:
        client: Shared GraphBridge client.
        sharepoint_list: Parent SharePoint list.
    """

    def __init__(self, client: GraphBridgeClient, sharepoint_list: SharePointListResource) -> None:
        """Initialize the list-items resource.

        Construction stores the parent list scope and performs no network request.

        Args:
            client: Shared GraphBridge client.
            sharepoint_list: Parent SharePoint list.
        """
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
        """Return the first page of list items.

        Server-side filtering is the default. Local filtering must be requested
        explicitly and is evaluated once against each downloaded page. Use
        :meth:`iter_all` for a complete potentially multi-page result.

        Args:
            fields: Optional SharePoint fields to expand.
            filter: Optional raw, controlled, or mapping filter.
            top: Optional maximum page size.
            select: Optional item properties to return.
            filter_mode: Whether filtering runs on the server or locally.
            field_names: Whether supplied field names are internal or display names.
        """

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
        """Lazily iterate through list-item pages.

        The first request is delayed until iteration, and subsequent Graph links
        are forwarded exactly without rebuilding the original query. Stopping
        iteration early prevents later requests.

        Args:
            fields: Optional SharePoint fields to expand.
            filter: Optional raw, controlled, or mapping filter.
            top: Optional maximum page size.
            select: Optional item properties to return.
            filter_mode: Whether filtering runs on the server or locally.
            field_names: Whether supplied field names are internal or display names.

        Raises:
            ValueError: If the filter mode or local filter is invalid.
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
        """Lazily iterate through all matching list items.

        This flattens page results while retaining lazy retrieval, making it
        suitable for collections that should not be loaded into memory at once.
        The first request is still deferred until iteration begins.

        Args:
            fields: Optional SharePoint fields to expand.
            filter: Optional raw, controlled, or mapping filter.
            top: Optional maximum page size.
            select: Optional item properties to return.
            filter_mode: Whether filtering runs on the server or locally.
            field_names: Whether supplied field names are internal or display names.
        """

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
        """Retrieve one SharePoint list item.

        Requested fields are expanded in the same call and the response is parsed
        into a model that preserves both convenient fields and raw metadata. Its
        eTag can protect a later conditional update or delete.

        Args:
            item_id: Graph list-item identifier.
            fields: Optional SharePoint fields to expand.
            field_names: Whether supplied field names are internal or display names.

        Raises:
            ValueError: If the item identifier is empty.
        """
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
        """Create one SharePoint list item.

        Display-name mode resolves the list schema before writing. Empty Graph
        successes are represented explicitly rather than confused with empty JSON.
        The returned model therefore represents both full and minimal Graph
        success shapes consistently.

        Args:
            fields: Field values for the new item.
            field_names: Whether supplied field names are internal or display names.

        Raises:
            TypeError: If fields is not a mapping.
        """
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
        """Update fields on one SharePoint list item.

        The stable endpoint patches the item's ``fields`` facet. Supplying an eTag
        enables optimistic concurrency and exposes stale writes as HTTP 412. The
        library never silently retries a business-level concurrency conflict.

        Args:
            item_id: Graph list-item identifier.
            fields: Field values to update.
            etag: Optional concurrency token sent with ``If-Match``.
            field_names: Whether supplied field names are internal or display names.

        Raises:
            ValueError: If the item identifier is empty.
            GraphPreconditionFailedError: If the supplied eTag is stale.
        """
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
        """Delete one SharePoint list item.

        An optional eTag can protect against deleting a remotely changed item, and
        a successful response is required to have no JSON body. Because the
        operation is destructive, callers should resolve item IDs explicitly.

        Args:
            item_id: Graph list-item identifier.
            etag: Optional concurrency token sent with ``If-Match``.

        Raises:
            ValueError: If the item identifier is empty.
            GraphPreconditionFailedError: If the supplied eTag is stale.
            GraphInvalidResponseError: If Graph returns a non-empty success body.
        """
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
        """Traverse one complete list-item delta round.

        Every continuation page is consumed and the final opaque delta link is
        returned unchanged. The last occurrence of an item in the round wins.
        Without ``known_ids``, non-deleted changes remain explicitly unclassified.

        Args:
            link: Optional opaque continuation or delta link.
            token: Optional initial delta token.
            fields: Optional SharePoint fields to expand.
            select: Optional item properties to return.
            top: Optional maximum page size.
            known_ids: Optional caller-owned IDs used to classify changes.
            field_names: Whether supplied field names are internal or display names.

        Returns:
            Classified changes and the final opaque delta link.

        Raises:
            ValueError: If a link is combined with incompatible query options.
            DeltaResetRequiredError: If Graph requires explicit resynchronization.
            GraphInvalidResponseError: If the delta response shape is invalid.
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
        """Return every retained version of one item.

        This convenience method delegates to the list's version resource and
        materializes all paginated version entries.

        Args:
            item_id: Graph list-item identifier.
        """

        return self.sharepoint_list.versions.versions(item_id)

    def restore_version(self, item_id: str, version_id: str) -> None:
        """Restore a retained version as the current state.

        SharePoint creates a new current version while retaining the existing
        history; GraphBridge does not replace or remove historical entries.

        Args:
            item_id: Graph list-item identifier.
            version_id: Version identifier to restore.
        """

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
        """Create multiple list items with Graph batches.

        Records are assigned deterministic input IDs, divided into groups of at
        most twenty, and correlated back into ordered per-input outcomes.

        Args:
            records: Ordered field mappings to create.
            max_attempts: Maximum attempts per transient subrequest.
            backoff_factor: Base factor used for retry delays.
            sleep: Optional function used to wait between retries.
            field_names: Whether supplied field names are internal or display names.

        Returns:
            Aggregate and per-input create outcomes.
        """
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
        """Update multiple list items with Graph batches.

        Several convenient input shapes are normalized to item ID, fields, and
        optional eTag triples before batching. Each subrequest reports its own
        success, failure, and attempt count.

        Args:
            updates: Update records or a parallel sequence of item IDs.
            fields: Optional field mappings paired with item IDs.
            etag: Optional eTag applied to every update.
            etags: Optional per-item eTags.
            max_attempts: Maximum attempts per transient subrequest.
            backoff_factor: Base factor used for retry delays.
            sleep: Optional function used to wait between retries.
            field_names: Whether supplied field names are internal or display names.

        Raises:
            ValueError: If eTag options conflict or input lengths do not match.
            TypeError: If an update record has an invalid shape.
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
        """Delete multiple items with Graph batches.

        Scalar or per-item eTags can be attached to conditional deletes, and
        partial failures remain correlated with the original identifiers.

        Args:
            item_ids: Item IDs or records containing IDs and optional eTags.
            etag: Optional eTag applied to every delete.
            etags: Optional per-item eTags.
            max_attempts: Maximum attempts per transient subrequest.
            backoff_factor: Base factor used for retry delays.
            sleep: Optional function used to wait between retries.

        Raises:
            ValueError: If eTag options conflict or an ID is empty.
        """

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
        """Return the stable Graph path for items in the bound list.

        Site and list IDs are quoted independently before child paths are added.
        """
        site_id = quote(self.sharepoint_list.site.id, safe=",")
        list_id = quote(self.sharepoint_list.id, safe="")
        return f"/sites/{site_id}/lists/{list_id}/items"

    def _field_names(
        self, fields: Sequence[str] | None, *, field_names: FieldNameMode
    ) -> Sequence[str] | None:
        """Translate selected field names when display mode is used.

        The cached column schema is authoritative; unknown display names fail
        rather than falling back to the legacy character encoder.

        Args:
            fields: Optional field names to translate.
            field_names: Field-name convention used by the caller.

        Raises:
            ValueError: If the field-name mode is invalid.
            KeyError: If a display name is unknown.
        """
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
        """Normalize fields for a write operation.

        Args:
            fields: Field values to normalize.
            field_names: Field-name convention used by the caller.

        Raises:
            TypeError: If fields is not a mapping.
            ValueError: If the field-name mode is invalid.
        """
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
        """Normalize a caller-provided filter.

        Args:
            value: Raw, controlled, mapping, or empty filter.
            field_names: Field-name convention used by the caller.
        """
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
        """Filter already downloaded pages locally.

        Args:
            pages: Parsed pages to filter.
            expression: Controlled filter with a local predicate.
        """
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
        """Execute item subrequests with configured retry defaults.

        Explicit options override transport settings, while omitted values inherit
        its attempt budget, backoff factor, delay cap, and sleep function.

        Args:
            requests: Ordered batch subrequests.
            max_attempts: Optional attempt-budget override.
            backoff_factor: Optional retry-factor override.
            sleep: Optional wait-function override.
        """
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
        """Validate and parse an item payload.

        Args:
            payload: Decoded Graph response.
            fallback_id: ID used when the payload omits one.

        Raises:
            GraphInvalidResponseError: If the payload is not a mapping.
        """
        if not isinstance(payload, Mapping):
            raise GraphInvalidResponseError("Microsoft Graph item response must be a JSON object")
        return ListItem.from_payload(payload, fallback_id=fallback_id)

    @classmethod
    def _created_item(cls, payload: Any, fields: Mapping[str, Any]) -> ListItem:
        """Parse a create response with field fallbacks.

        Args:
            payload: Decoded Graph response or ``None``.
            fields: Submitted field values.
        """
        if payload is None:
            return ListItem(id="", fields=dict(fields), response_empty=True)
        item = cls._item_from_payload(payload)
        return cls._with_fallback_fields(item, fields)

    @classmethod
    def _updated_item(
        cls, item_id: str, fields: Mapping[str, Any], payload: Any
    ) -> ListItem:
        """Parse an update response with item and field fallbacks.

        Args:
            item_id: Updated item identifier.
            fields: Submitted field values.
            payload: Decoded Graph response or ``None``.
        """
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
        """Add submitted fields when Graph omits them.

        Args:
            item: Parsed list item.
            fields: Submitted field values.
        """
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
        """Convert create subresponses into a batch result.

        Args:
            responses: Correlated raw batch responses.
            records: Submitted create records.
        """
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
        """Convert update subresponses into a batch result.

        Args:
            responses: Correlated raw batch responses.
            updates: Normalized update records.
        """
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
        """Convert delete subresponses into a batch result.

        Args:
            responses: Correlated raw batch responses.
            deletes: Normalized delete records.
        """
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
        """Build one successful batch item result.

        Args:
            response: Correlated raw batch response.
            value: Parsed successful value.
        """
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
        """Build one failed batch item result.

        Args:
            response: Correlated raw batch response.
            error: Structured failure.
        """
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
        """Build an error for an invalid successful subresponse.

        Args:
            response: Correlated raw batch response.
            message: Validation failure message.
        """
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
        """Copy a response-header eTag onto an item when needed.

        Args:
            item: Parsed list item.
            headers: Batch subresponse headers.
        """
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
        """Normalize accepted batch-update input shapes.

        Parallel ID/field sequences, tuples, and mappings are converted to one
        consistent representation before request URLs and headers are created.

        Args:
            updates: Update records or item IDs.
            fields: Optional parallel field mappings.

        Raises:
            ValueError: If lengths differ or an item ID is empty.
            TypeError: If a record or field mapping is invalid.
        """
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
        """Normalize accepted batch-delete input shapes.

        Args:
            values: Item IDs or records containing IDs and eTags.

        Raises:
            ValueError: If an item ID is empty.
        """
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
        """Resolve an eTag override for one input.

        Args:
            etags: Scalar, sequence, mapping, or empty eTag configuration.
            index: Position in the input sequence.
            item_id: Graph list-item identifier.

        Raises:
            ValueError: If an eTag sequence is too short.
        """
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
        """Read a header case-insensitively.

        Args:
            headers: Response header mapping.
            name: Header name to retrieve.
        """
        expected = name.casefold()
        for key, value in headers.items():
            if str(key).casefold() == expected:
                return str(value)
        return None

    @staticmethod
    def _optional_string(value: object) -> str | None:
        """Convert a value to text while preserving ``None``.

        Args:
            value: Value to convert.
        """
        return str(value) if value is not None else None

    @staticmethod
    def _delta_link(value: object) -> str | None:
        """Validate an optional delta continuation link.

        Args:
            value: Link value returned by Graph.

        Raises:
            GraphInvalidResponseError: If a present link is invalid.
        """
        if value is None:
            return None
        if not isinstance(value, str) or not value:
            raise GraphInvalidResponseError(
                "Microsoft Graph returned an invalid delta pagination link"
            )
        return value
