"""Plan-first, loss-aware SharePoint list synchronization."""

from __future__ import annotations

from collections.abc import Callable, Iterable, Mapping, Sequence
from typing import TYPE_CHECKING, Any, Literal

from ..exceptions import (
    GraphBridgeError,
    GraphRequestError,
    SyncDuplicateKeyError,
    SyncMissingKeyError,
    SyncValidationError,
)
from ..models import (
    BatchItemResult,
    GraphError,
    ListItem,
    SyncFieldDifference,
    SyncOperation,
    SyncOperationResult,
    SyncPlan,
    SyncResult,
)

if TYPE_CHECKING:
    from ..client import GraphBridgeClient
    from .lists import SharePointListResource

FieldNameMode = Literal["internal", "display"]


class SyncService:
    """Build, review, apply and retry deterministic list synchronization plans."""

    def __init__(
        self, client: GraphBridgeClient, sharepoint_list: SharePointListResource
    ) -> None:
        self.client = client
        self.transport = client.transport
        self.sharepoint_list = sharepoint_list
        self.items = sharepoint_list.items

    def plan(
        self,
        *,
        rows: Iterable[Mapping[str, Any]],
        key_field: str,
        prune: bool = False,
        dry_run: bool = False,
        field_names: FieldNameMode = "internal",
    ) -> SyncPlan:
        """Read remote state once and return a mutation-free synchronization plan.

        Only source fields are compared. Extra SharePoint/system fields therefore
        do not cause updates. ``prune=False`` retains remote-only items as
        ``unchanged`` operations with an explicit reason.
        """

        source = self._normalize_rows(rows)
        if not isinstance(key_field, str) or not key_field:
            raise ValueError("key_field cannot be empty")
        if not isinstance(prune, bool):
            raise TypeError("prune must be a boolean")
        if not isinstance(dry_run, bool):
            raise TypeError("dry_run must be a boolean")
        if field_names == "display":
            mapping = self.sharepoint_list.columns.display_name_map()
            internal_names = set(mapping.values())
            if key_field in internal_names:
                internal_key = key_field
            elif key_field in mapping:
                internal_key = mapping[key_field]
            else:
                raise KeyError(f"unknown SharePoint column displayName: {key_field!r}")
            source = [self.sharepoint_list.columns.to_internal_fields(row) for row in source]
        elif field_names == "internal":
            internal_key = key_field
        else:
            raise ValueError("field_names must be 'internal' or 'display'")

        source_pairs = self._validate_source_keys(source, internal_key)
        source_duplicates = self._duplicates(
            [(key, f"source[{index}]") for key, index in source_pairs]
        )
        if source_duplicates:
            raise SyncDuplicateKeyError(internal_key, source_duplicates)

        selected_fields = list(
            dict.fromkeys(
                [internal_key]
                + [str(name) for row in source for name in row]
            )
        )
        remote = list(self.items.iter_all(fields=selected_fields))
        return self._build_key_plan(
            source,
            remote,
            key_field=internal_key,
            prune=prune,
            dry_run=dry_run,
            field_names=field_names,
        )

    def apply(
        self,
        plan: SyncPlan,
        *,
        dry_run: bool | None = None,
        use_batch: bool = True,
        max_attempts: int | None = None,
        backoff_factor: float | None = None,
        sleep: Callable[[float], None] | None = None,
    ) -> SyncResult:
        """Apply a reviewed plan in create, PATCH-update, then prune order.

        A failed create defers every planned delete. This conservative barrier
        prevents a prune from removing an item while a desired source item was
        not successfully created. Batch subrequest retries are delegated to the
        generic batch executor and never replay successful subrequests.
        """

        if not isinstance(plan, SyncPlan):
            raise TypeError("plan must be a SyncPlan")
        effective_dry_run = plan.dry_run if dry_run is None else dry_run
        if not isinstance(effective_dry_run, bool):
            raise TypeError("dry_run must be a boolean")
        creates = self._typed_operations(plan.creates, "create")
        updates = self._typed_operations(plan.updates, "update")
        deletes = self._typed_operations(plan.deletes, "delete")
        if effective_dry_run:
            return SyncResult(
                unchanged=list(plan.unchanged),
                plan=plan,
                applied=False,
                dry_run=True,
            )

        created: list[ListItem] = []
        updated: list[ListItem] = []
        deleted: list[str] = []
        failures: list[GraphError] = []
        outcomes: list[SyncOperationResult] = []

        if use_batch:
            self._apply_create_batches(
                creates,
                created,
                failures,
                outcomes,
                max_attempts=max_attempts,
                backoff_factor=backoff_factor,
                sleep=sleep,
            )
            self._apply_update_batches(
                updates,
                updated,
                failures,
                outcomes,
                max_attempts=max_attempts,
                backoff_factor=backoff_factor,
                sleep=sleep,
            )
        else:
            self._apply_direct_writes(
                creates, updates, created, updated, failures, outcomes
            )

        create_failed = any(
            result.operation.operation == "create" and not result.succeeded
            for result in outcomes
        )
        if create_failed:
            for operation in deletes:
                error = GraphError(
                    code="deleteDeferred",
                    message="delete was deferred because at least one create failed",
                    inner_error={"item_id": operation.item_id, "key": operation.key},
                )
                failures.append(error)
                outcomes.append(
                    SyncOperationResult(
                        operation=operation,
                        error=error,
                        attempts=0,
                        deferred=True,
                    )
                )
        elif use_batch:
            self._apply_delete_batches(
                deletes,
                deleted,
                failures,
                outcomes,
                max_attempts=max_attempts,
                backoff_factor=backoff_factor,
                sleep=sleep,
            )
        else:
            self._apply_direct_deletes(deletes, deleted, failures, outcomes)

        outcomes.extend(
            SyncOperationResult(
                operation=operation,
                value=operation.item_id,
                attempts=0,
            )
            for operation in plan.unchanged
        )
        return SyncResult(
            created=created,
            updated=updated,
            deleted=deleted,
            unchanged=list(plan.unchanged),
            failures=failures,
            results=outcomes,
            plan=plan,
            applied=True,
        )

    def retry(self, result: SyncResult, **apply_options: Any) -> SyncResult:
        """Retry only failed or safely deferred operations from ``result``."""

        if not isinstance(result, SyncResult):
            raise TypeError("result must be a SyncResult")
        return self.apply(result.retry_plan(), **apply_options)

    def _plan_by_item_ids(
        self,
        *,
        rows: Sequence[Mapping[str, Any]],
        item_ids: Sequence[str],
        remote_items: Sequence[ListItem],
        prune: bool,
    ) -> SyncPlan:
        """Build the legacy upload adapter plan without another remote read."""

        if len(rows) != len(item_ids):
            raise ValueError("item_ids and rows must have the same length")
        duplicate_ids = self._duplicates(
            [(item_id, f"source[{index}]") for index, item_id in enumerate(item_ids)]
        )
        if duplicate_ids:
            raise SyncDuplicateKeyError("@item_id", duplicate_ids)
        remote_by_id = {item.id: (index, item) for index, item in enumerate(remote_items)}
        creates: list[SyncOperation] = []
        updates: list[SyncOperation] = []
        deletes: list[SyncOperation] = []
        unchanged: list[SyncOperation] = []
        source_ids = set(item_ids)
        for source_index, (item_id, fields) in enumerate(zip(item_ids, rows, strict=True)):
            remote_entry = remote_by_id.get(item_id)
            if remote_entry is None:
                creates.append(
                    SyncOperation(
                        operation="create",
                        key=item_id,
                        fields=dict(fields),
                        reason="legacy item id is not present remotely",
                        source_index=source_index,
                    )
                )
                continue
            remote_index, remote = remote_entry
            differences = self._differences(fields, remote.fields)
            if differences:
                updates.append(
                    SyncOperation(
                        operation="update",
                        key=item_id,
                        fields={value.field: value.local_value for value in differences},
                        item_id=remote.id,
                        etag=remote.etag,
                        differences=differences,
                        reason="legacy row differs from the existing item",
                        source_index=source_index,
                        remote_index=remote_index,
                    )
                )
            else:
                unchanged.append(
                    SyncOperation(
                        operation="unchanged",
                        key=item_id,
                        fields=dict(fields),
                        item_id=remote.id,
                        etag=remote.etag,
                        reason="legacy row already matches the existing item",
                        source_index=source_index,
                        remote_index=remote_index,
                    )
                )
        for remote_index, remote in enumerate(remote_items):
            if remote.id in source_ids:
                continue
            operation = SyncOperation(
                operation="delete" if prune else "unchanged",
                key=remote.id,
                fields=dict(remote.fields),
                item_id=remote.id,
                etag=remote.etag,
                reason=(
                    "remote item is absent from the legacy source and prune is enabled"
                    if prune
                    else "remote item is retained because prune is disabled"
                ),
                remote_index=remote_index,
            )
            if prune:
                deletes.append(operation)
            else:
                unchanged.append(operation)
        return SyncPlan(
            creates=creates,
            updates=updates,
            deletes=deletes,
            unchanged=unchanged,
            key_field="@item_id",
            prune=prune,
            local_count=len(rows),
            remote_count=len(remote_items),
        )

    def _build_key_plan(
        self,
        source: Sequence[Mapping[str, Any]],
        remote: Sequence[ListItem],
        *,
        key_field: str,
        prune: bool,
        dry_run: bool,
        field_names: FieldNameMode,
    ) -> SyncPlan:
        source_pairs = self._validate_source_keys(source, key_field)
        remote_pairs = self._validate_remote_keys(remote, key_field)
        source_duplicates = self._duplicates(
            [(key, f"source[{index}]") for key, index in source_pairs]
        )
        remote_duplicates = self._duplicates(
            [(key, f"remote[{index}]") for key, index in remote_pairs]
        )
        duplicates = dict(source_duplicates)
        for key, locations in remote_duplicates.items():
            duplicates.setdefault(key, []).extend(locations)
        if duplicates:
            raise SyncDuplicateKeyError(key_field, duplicates)

        remote_by_key = {
            key: (index, remote[index]) for key, index in remote_pairs
        }
        source_keys = {key for key, _index in source_pairs}
        creates: list[SyncOperation] = []
        updates: list[SyncOperation] = []
        deletes: list[SyncOperation] = []
        unchanged: list[SyncOperation] = []
        for key, source_index in source_pairs:
            fields = source[source_index]
            remote_entry = remote_by_key.get(key)
            if remote_entry is None:
                creates.append(
                    SyncOperation(
                        operation="create",
                        key=key,
                        fields=dict(fields),
                        reason="source key does not exist remotely",
                        source_index=source_index,
                    )
                )
                continue
            remote_index, item = remote_entry
            differences = self._differences(fields, item.fields)
            if differences:
                updates.append(
                    SyncOperation(
                        operation="update",
                        key=key,
                        fields={value.field: value.local_value for value in differences},
                        item_id=item.id,
                        etag=item.etag,
                        differences=differences,
                        reason="one or more source fields differ from SharePoint",
                        source_index=source_index,
                        remote_index=remote_index,
                    )
                )
            else:
                unchanged.append(
                    SyncOperation(
                        operation="unchanged",
                        key=key,
                        fields=dict(fields),
                        item_id=item.id,
                        etag=item.etag,
                        reason="source fields already match SharePoint",
                        source_index=source_index,
                        remote_index=remote_index,
                    )
                )
        for key, remote_index in remote_pairs:
            if key in source_keys:
                continue
            item = remote[remote_index]
            operation = SyncOperation(
                operation="delete" if prune else "unchanged",
                key=key,
                fields=dict(item.fields),
                item_id=item.id,
                etag=item.etag,
                reason=(
                    "remote key is absent from the source and prune is enabled"
                    if prune
                    else "remote key is retained because prune is disabled"
                ),
                remote_index=remote_index,
            )
            if prune:
                deletes.append(operation)
            else:
                unchanged.append(operation)
        return SyncPlan(
            creates=creates,
            updates=updates,
            deletes=deletes,
            unchanged=unchanged,
            key_field=key_field,
            prune=prune,
            dry_run=dry_run,
            field_names=field_names,
            local_count=len(source),
            remote_count=len(remote),
        )

    def _apply_create_batches(
        self,
        operations: Sequence[SyncOperation],
        created: list[ListItem],
        failures: list[GraphError],
        outcomes: list[SyncOperationResult],
        **batch_options: Any,
    ) -> None:
        for operation_chunk in self._chunks(operations):
            try:
                batch = self.items.create_many(
                    [operation.fields for operation in operation_chunk],
                    **batch_options,
                )
            except GraphBridgeError as error:
                self._record_phase_error(operation_chunk, error, failures, outcomes)
                continue
            self._record_batch_results(
                operation_chunk, batch.results, created, failures, outcomes
            )

    def _apply_update_batches(
        self,
        operations: Sequence[SyncOperation],
        updated: list[ListItem],
        failures: list[GraphError],
        outcomes: list[SyncOperationResult],
        **batch_options: Any,
    ) -> None:
        for operation_chunk in self._chunks(operations):
            try:
                batch = self.items.update_many(
                    [
                        (operation.item_id, operation.fields, operation.etag)
                        for operation in operation_chunk
                    ],
                    **batch_options,
                )
            except GraphBridgeError as error:
                self._record_phase_error(operation_chunk, error, failures, outcomes)
                continue
            self._record_batch_results(
                operation_chunk, batch.results, updated, failures, outcomes
            )

    def _apply_delete_batches(
        self,
        operations: Sequence[SyncOperation],
        deleted: list[str],
        failures: list[GraphError],
        outcomes: list[SyncOperationResult],
        **batch_options: Any,
    ) -> None:
        for operation_chunk in self._chunks(operations):
            try:
                batch = self.items.delete_many(
                    [(operation.item_id, operation.etag) for operation in operation_chunk],
                    **batch_options,
                )
            except GraphBridgeError as error:
                self._record_phase_error(operation_chunk, error, failures, outcomes)
                continue
            self._record_batch_results(
                operation_chunk, batch.results, deleted, failures, outcomes
            )

    def _apply_direct_writes(
        self,
        creates: Sequence[SyncOperation],
        updates: Sequence[SyncOperation],
        created: list[ListItem],
        updated: list[ListItem],
        failures: list[GraphError],
        outcomes: list[SyncOperationResult],
    ) -> None:
        for operation in creates:
            try:
                value = self.items.create(operation.fields)
            except GraphBridgeError as error:
                self._record_phase_error([operation], error, failures, outcomes)
            else:
                created.append(value)
                outcomes.append(
                    SyncOperationResult(operation=operation, value=value, status_code=201)
                )
        for operation in updates:
            assert operation.item_id is not None
            try:
                value = self.items.update(
                    operation.item_id, operation.fields, etag=operation.etag
                )
            except GraphBridgeError as error:
                self._record_phase_error([operation], error, failures, outcomes)
            else:
                updated.append(value)
                outcomes.append(
                    SyncOperationResult(operation=operation, value=value, status_code=200)
                )

    def _apply_direct_deletes(
        self,
        operations: Sequence[SyncOperation],
        deleted: list[str],
        failures: list[GraphError],
        outcomes: list[SyncOperationResult],
    ) -> None:
        for operation in operations:
            assert operation.item_id is not None
            try:
                self.items.delete(operation.item_id, etag=operation.etag)
            except GraphBridgeError as error:
                self._record_phase_error([operation], error, failures, outcomes)
            else:
                deleted.append(operation.item_id)
                outcomes.append(
                    SyncOperationResult(
                        operation=operation,
                        value=operation.item_id,
                        status_code=204,
                    )
                )

    @staticmethod
    def _record_batch_results(
        operations: Sequence[SyncOperation],
        batch_results: Sequence[BatchItemResult[Any]],
        successes: list[Any],
        failures: list[GraphError],
        outcomes: list[SyncOperationResult],
    ) -> None:
        for result in batch_results:
            operation = operations[result.input_index]
            if result.error is not None:
                failures.append(result.error)
                outcomes.append(
                    SyncOperationResult(
                        operation=operation,
                        error=result.error,
                        status_code=result.status_code,
                        attempts=result.attempts,
                    )
                )
            else:
                value = result.value
                successes.append(value)
                outcomes.append(
                    SyncOperationResult(
                        operation=operation,
                        value=value,
                        status_code=result.status_code,
                        attempts=result.attempts,
                    )
                )

    @classmethod
    def _record_phase_error(
        cls,
        operations: Sequence[SyncOperation],
        caught: GraphBridgeError,
        failures: list[GraphError],
        outcomes: list[SyncOperationResult],
    ) -> None:
        error = cls._graph_error(caught)
        for operation in operations:
            contextual = GraphError(
                code=error.code,
                message=error.message,
                status_code=error.status_code,
                request_id=error.request_id,
                date=error.date,
                inner_error={
                    **dict(error.inner_error),
                    "item_id": operation.item_id,
                    "key": operation.key,
                },
            )
            failures.append(contextual)
            outcomes.append(
                SyncOperationResult(
                    operation=operation,
                    error=contextual,
                    status_code=contextual.status_code,
                )
            )

    @staticmethod
    def _graph_error(caught: GraphBridgeError) -> GraphError:
        if isinstance(caught, GraphRequestError):
            return caught.error
        candidate = getattr(caught, "error", None)
        if isinstance(candidate, GraphError):
            return candidate
        return GraphError(code=type(caught).__name__, message=str(caught))

    @staticmethod
    def _typed_operations(
        values: Sequence[object], expected: str
    ) -> list[SyncOperation]:
        operations: list[SyncOperation] = []
        for value in values:
            if not isinstance(value, SyncOperation) or value.operation != expected:
                raise TypeError(f"plan {expected} entries must be SyncOperation values")
            operations.append(value)
        return operations

    @staticmethod
    def _normalize_rows(
        rows: Iterable[Mapping[str, Any]],
    ) -> list[dict[str, Any]]:
        if isinstance(rows, Mapping):
            raise TypeError("rows must be an iterable of mappings, not one mapping")
        try:
            source = list(rows)
        except TypeError:
            raise TypeError("rows must be an iterable of mappings") from None
        if not all(isinstance(row, Mapping) for row in source):
            raise TypeError("every synchronization row must be a mapping")
        return [dict(row) for row in source]

    @classmethod
    def _validate_source_keys(
        cls, rows: Sequence[Mapping[str, Any]], key_field: str
    ) -> list[tuple[Any, int]]:
        missing: list[str] = []
        pairs: list[tuple[Any, int]] = []
        for index, row in enumerate(rows):
            if key_field not in row or cls._empty_key(row.get(key_field)):
                missing.append(f"source[{index}]")
                continue
            key = row[key_field]
            cls._ensure_hashable_key(key_field, key, f"source[{index}]")
            pairs.append((key, index))
        if missing:
            raise SyncMissingKeyError(key_field, missing)
        return pairs

    @classmethod
    def _validate_remote_keys(
        cls, items: Sequence[ListItem], key_field: str
    ) -> list[tuple[Any, int]]:
        missing: list[str] = []
        pairs: list[tuple[Any, int]] = []
        for index, item in enumerate(items):
            if key_field not in item.fields or cls._empty_key(item.fields.get(key_field)):
                missing.append(f"remote[{index}] item {item.id!r}")
                continue
            key = item.fields[key_field]
            cls._ensure_hashable_key(key_field, key, f"remote[{index}]")
            pairs.append((key, index))
        if missing:
            raise SyncMissingKeyError(key_field, missing)
        return pairs

    @staticmethod
    def _empty_key(value: Any) -> bool:
        return value is None or value == ""

    @staticmethod
    def _ensure_hashable_key(key_field: str, value: Any, location: str) -> None:
        try:
            hash(value)
        except TypeError:
            raise SyncValidationError(
                f"sync key {key_field!r} at {location} must be a scalar hashable value"
            ) from None

    @staticmethod
    def _duplicates(values: Sequence[tuple[Any, str]]) -> dict[Any, list[str]]:
        locations: dict[Any, list[str]] = {}
        for key, location in values:
            locations.setdefault(key, []).append(location)
        return {key: value for key, value in locations.items() if len(value) > 1}

    @staticmethod
    def _differences(
        local: Mapping[str, Any], remote: Mapping[str, Any]
    ) -> list[SyncFieldDifference]:
        return [
            SyncFieldDifference(
                field=str(name),
                local_value=value,
                remote_value=remote.get(name),
            )
            for name, value in local.items()
            if name not in remote or remote[name] != value
        ]

    @staticmethod
    def _chunks(values: Sequence[SyncOperation]) -> Iterable[Sequence[SyncOperation]]:
        for index in range(0, len(values), 20):
            yield values[index : index + 20]
