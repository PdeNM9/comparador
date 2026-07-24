"""Core comparison rules for judicial process spreadsheets."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any

import pandas as pd

from excel_utils import (
    ensure_columns_exist,
    first_non_empty_value,
    join_unique_values,
    normalize_cnj_series,
)


INTERNAL_CNJ_KEY = object()


class ComparisonError(ValueError):
    """Raised when the selected configuration cannot be compared."""


@dataclass(frozen=True)
class SheetConfig:
    """Selected columns for one uploaded worksheet."""

    cnj_column: Any
    responsible_column: Any
    annotation_columns: tuple[Any, ...] = ()


@dataclass(frozen=True)
class DuplicateReport:
    """Summary of duplicated normalized CNJs in one worksheet."""

    sheet_label: str
    duplicated_processes: int
    duplicated_rows: int
    sample: pd.DataFrame


@dataclass(frozen=True)
class ComparisonResult:
    """All data produced by a spreadsheet comparison."""

    old_total_rows: int
    new_total_rows: int
    old_valid_rows: int
    new_valid_rows: int
    old_unique_processes: int
    new_unique_processes: int
    excluded_processes: pd.DataFrame
    new_processes: pd.DataFrame
    maintained_processes: pd.DataFrame
    final_dataframe: pd.DataFrame
    duplicate_reports: tuple[DuplicateReport, ...]
    old_empty_cnj_rows: int
    new_empty_cnj_rows: int
    annotation_column_map: dict[Any, Any]

    @property
    def status_counts(self) -> dict[str, int]:
        """Counts used by cards and charts."""
        return {
            "Excluídos": len(self.excluded_processes),
            "Novos": len(self.new_processes),
            "Mantidos": len(self.maintained_processes),
        }


def compare_sheets(
    old_dataframe: pd.DataFrame,
    new_dataframe: pd.DataFrame,
    old_config: SheetConfig,
    new_config: SheetConfig,
) -> ComparisonResult:
    """Compare two worksheets and build the final annotation-enriched file."""
    _validate_inputs(old_dataframe, new_dataframe, old_config, new_config)

    old_prepared = _with_normalized_key(old_dataframe, old_config.cnj_column)
    new_prepared = _with_normalized_key(new_dataframe, new_config.cnj_column)

    old_valid = old_prepared[old_prepared[INTERNAL_CNJ_KEY].ne("")]
    new_valid = new_prepared[new_prepared[INTERNAL_CNJ_KEY].ne("")]

    if old_valid.empty:
        raise ComparisonError(
            "A Planilha 1 não possui nenhum CNJ válido na coluna selecionada."
        )
    if new_valid.empty:
        raise ComparisonError(
            "A Planilha 2 não possui nenhum CNJ válido na coluna selecionada."
        )

    old_unique_keys = old_valid[INTERNAL_CNJ_KEY].drop_duplicates()
    new_unique_keys = new_valid[INTERNAL_CNJ_KEY].drop_duplicates()

    excluded_keys = old_unique_keys[~old_unique_keys.isin(new_unique_keys)]
    new_keys = new_unique_keys[~new_unique_keys.isin(old_unique_keys)]
    maintained_keys = new_unique_keys[new_unique_keys.isin(old_unique_keys)]

    excluded_processes = _build_process_table(
        old_valid,
        old_config.cnj_column,
        old_config.responsible_column,
        excluded_keys,
    )
    new_processes = _build_process_table(
        new_valid,
        new_config.cnj_column,
        new_config.responsible_column,
        new_keys,
    )
    maintained_processes = _build_process_table(
        new_valid,
        new_config.cnj_column,
        new_config.responsible_column,
        maintained_keys,
    )

    annotation_values = _aggregate_annotations(
        old_valid,
        old_config.annotation_columns,
    )
    final_dataframe, annotation_column_map = _build_final_dataframe(
        new_dataframe,
        new_prepared[INTERNAL_CNJ_KEY],
        annotation_values,
        old_config.annotation_columns,
    )

    duplicate_reports = tuple(
        report
        for report in (
            _build_duplicate_report(
                old_valid,
                old_config.cnj_column,
                old_config.responsible_column,
                "Planilha 1",
            ),
            _build_duplicate_report(
                new_valid,
                new_config.cnj_column,
                new_config.responsible_column,
                "Planilha 2",
            ),
        )
        if report is not None
    )

    return ComparisonResult(
        old_total_rows=len(old_dataframe),
        new_total_rows=len(new_dataframe),
        old_valid_rows=len(old_valid),
        new_valid_rows=len(new_valid),
        old_unique_processes=int(old_unique_keys.size),
        new_unique_processes=int(new_unique_keys.size),
        excluded_processes=excluded_processes,
        new_processes=new_processes,
        maintained_processes=maintained_processes,
        final_dataframe=final_dataframe,
        duplicate_reports=duplicate_reports,
        old_empty_cnj_rows=int(old_prepared[INTERNAL_CNJ_KEY].eq("").sum()),
        new_empty_cnj_rows=int(new_prepared[INTERNAL_CNJ_KEY].eq("").sum()),
        annotation_column_map=annotation_column_map,
    )


def _validate_inputs(
    old_dataframe: pd.DataFrame,
    new_dataframe: pd.DataFrame,
    old_config: SheetConfig,
    new_config: SheetConfig,
) -> None:
    if old_dataframe.empty:
        raise ComparisonError("A Planilha 1 está vazia.")
    if new_dataframe.empty:
        raise ComparisonError("A Planilha 2 está vazia.")

    ensure_columns_exist(
        old_dataframe,
        [
            old_config.cnj_column,
            old_config.responsible_column,
            *old_config.annotation_columns,
        ],
        "Planilha 1",
    )
    ensure_columns_exist(
        new_dataframe,
        [new_config.cnj_column, new_config.responsible_column],
        "Planilha 2",
    )


def _with_normalized_key(dataframe: pd.DataFrame, cnj_column: Any) -> pd.DataFrame:
    prepared = dataframe.copy()
    prepared[INTERNAL_CNJ_KEY] = normalize_cnj_series(prepared[cnj_column])
    return prepared


def _build_process_table(
    prepared: pd.DataFrame,
    cnj_column: Any,
    responsible_column: Any,
    keys: pd.Series,
) -> pd.DataFrame:
    columns = ["Número CNJ", "Responsável"]
    if keys.empty:
        return pd.DataFrame(columns=columns)

    filtered = prepared[prepared[INTERNAL_CNJ_KEY].isin(keys)]
    if filtered.empty:
        return pd.DataFrame(columns=columns)

    grouped = filtered.groupby(INTERNAL_CNJ_KEY, sort=False)
    table = grouped.agg(
        **{
            "Número CNJ": (cnj_column, first_non_empty_value),
            "Responsável": (responsible_column, join_unique_values),
        }
    ).reset_index(drop=True)

    return table[columns]


def _aggregate_annotations(
    old_valid: pd.DataFrame,
    annotation_columns: tuple[Any, ...],
) -> pd.DataFrame:
    if not annotation_columns:
        return pd.DataFrame(index=old_valid[INTERNAL_CNJ_KEY].drop_duplicates())

    aggregations = {
        annotation_column: join_unique_values
        for annotation_column in annotation_columns
    }
    return old_valid.groupby(INTERNAL_CNJ_KEY, sort=False).agg(aggregations)


def _build_final_dataframe(
    new_dataframe: pd.DataFrame,
    new_keys: pd.Series,
    annotation_values: pd.DataFrame,
    annotation_columns: tuple[Any, ...],
) -> tuple[pd.DataFrame, dict[Any, Any]]:
    final_dataframe = new_dataframe.copy()
    annotation_column_map = _build_annotation_column_map(
        final_dataframe.columns,
        annotation_columns,
    )

    for source_column, target_column in annotation_column_map.items():
        if source_column not in annotation_values.columns:
            final_dataframe[target_column] = ""
            continue

        copied_values = new_keys.map(annotation_values[source_column])
        final_dataframe[target_column] = copied_values.fillna("")

    return final_dataframe, annotation_column_map


def _build_annotation_column_map(
    existing_columns: pd.Index,
    annotation_columns: tuple[Any, ...],
) -> dict[Any, Any]:
    used_columns = set(existing_columns)
    mapping: dict[Any, Any] = {}

    for source_column in annotation_columns:
        target_column = source_column
        if target_column in used_columns:
            base_name = f"{source_column} - Planilha 1"
            target_column = base_name
            counter = 2
            while target_column in used_columns:
                target_column = f"{base_name} ({counter})"
                counter += 1

        mapping[source_column] = target_column
        used_columns.add(target_column)

    return mapping


def _build_duplicate_report(
    valid_dataframe: pd.DataFrame,
    cnj_column: Any,
    responsible_column: Any,
    sheet_label: str,
) -> DuplicateReport | None:
    duplicated_mask = valid_dataframe.duplicated(
        INTERNAL_CNJ_KEY,
        keep=False,
    )
    if not duplicated_mask.any():
        return None

    duplicated = valid_dataframe.loc[duplicated_mask]
    counts = duplicated[INTERNAL_CNJ_KEY].value_counts()
    sample_keys = counts.head(10).index
    sample_records = []

    for normalized_key in sample_keys:
        rows = duplicated[duplicated[INTERNAL_CNJ_KEY].eq(normalized_key)]
        sample_records.append(
            {
                "CNJ normalizado": normalized_key,
                "Número CNJ": first_non_empty_value(rows[cnj_column]),
                "Responsável": join_unique_values(rows[responsible_column]),
                "Ocorrências": int(counts.loc[normalized_key]),
            }
        )

    return DuplicateReport(
        sheet_label=sheet_label,
        duplicated_processes=int(counts.size),
        duplicated_rows=int(duplicated.shape[0]),
        sample=pd.DataFrame(sample_records),
    )
