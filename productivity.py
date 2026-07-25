"""Productivity comparison rules for server-based judicial spreadsheets."""

from __future__ import annotations

from dataclasses import dataclass
from io import BytesIO
from typing import Any
from zipfile import BadZipFile
import unicodedata

import numpy as np
import pandas as pd
from openpyxl.utils.exceptions import InvalidFileException

from excel_utils import (
    ExcelProcessingError,
    first_non_empty_value,
    join_unique_values,
    normalize_cnj_series,
    validate_xlsx_filename,
)


PROCESS_KEY = object()

OLD_PROCESS_COLUMN = "Número do Processo"
OLD_STATUS_COLUMN = "Situação"
OLD_NOTE_COLUMN = "Observação"

CURRENT_REQUIRED_COLUMNS = (
    "Descrição Classe",
    "Número Processo",
    "Quantidade de Dias na Situação Atual Processo",
    "Situação Atual",
    "Última Tarefa PJE",
)
CURRENT_PROCESS_COLUMN = "Número Processo"

SERVER_COLUMN = "Servidor"
PREVIOUS_STATUS_COLUMN = "Situação anterior"
PREVIOUS_NOTE_COLUMN = "Observação anterior"
COMPARISON_STATUS_COLUMN = "Status comparativo"

STATUS_REMAINED = "Permaneceu"
STATUS_PRODUCTIVE = "Saiu"
STATUS_NEW = "Novo"
STATUS_EMPTY_CNJ = "CNJ vazio"


class ProductivityError(ValueError):
    """Raised when the productivity report cannot be produced."""


@dataclass(frozen=True)
class DuplicateReport:
    """Summary of duplicated normalized CNJs."""

    source_label: str
    duplicated_processes: int
    duplicated_rows: int
    sample: pd.DataFrame


@dataclass(frozen=True)
class ProductivityResult:
    """All tables and counters produced by the productivity comparison."""

    server_count: int
    old_total_rows: int
    old_valid_rows: int
    old_unique_processes: int
    current_total_rows: int
    current_valid_rows: int
    current_unique_processes: int
    productive_unique_processes: int
    remained_unique_processes: int
    new_unique_processes: int
    old_empty_cnj_rows: int
    current_empty_cnj_rows: int
    current_sheet_name: str
    skipped_empty_sheets: tuple[str, ...]
    duplicate_reports: tuple[DuplicateReport, ...]
    summary_by_server: pd.DataFrame
    productive_processes: pd.DataFrame
    new_processes: pd.DataFrame
    current_enriched: pd.DataFrame
    old_comparison: pd.DataFrame

    @property
    def status_counts(self) -> dict[str, int]:
        """Counts used by status charts."""
        return {
            "Produtivos": self.productive_unique_processes,
            "Permaneceram": self.remained_unique_processes,
            "Novos": self.new_unique_processes,
        }

    @property
    def export_sheets(self) -> dict[str, pd.DataFrame]:
        """Workbook sheets used by the full report download."""
        return {
            "Resumo por servidor": self.summary_by_server,
            "Atual enriquecida": self.current_enriched,
            "Produtivos - saíram": self.productive_processes,
            "Novos": self.new_processes,
            "Lista antiga comparada": self.old_comparison,
        }


def read_workbook_sheets(file_bytes: bytes, file_name: str) -> dict[str, pd.DataFrame]:
    """Read every worksheet from an uploaded .xlsx workbook."""
    validate_xlsx_filename(file_name)
    if not file_bytes:
        raise ExcelProcessingError("O arquivo enviado está vazio.")

    try:
        workbook = pd.ExcelFile(BytesIO(file_bytes), engine="openpyxl")
        try:
            sheets = {
                sheet_name: workbook.parse(
                    sheet_name=sheet_name,
                    keep_default_na=False,
                )
                for sheet_name in workbook.sheet_names
            }
        finally:
            workbook.close()
    except (BadZipFile, InvalidFileException, ValueError, OSError) as exc:
        raise ExcelProcessingError(
            "Não foi possível abrir o arquivo Excel. Verifique se ele não "
            "está corrompido e se realmente está no formato .xlsx."
        ) from exc

    if not sheets:
        raise ExcelProcessingError("A planilha não possui abas para leitura.")
    return sheets


def build_productivity_report(
    server_sheets: dict[str, pd.DataFrame],
    current_dataframe: pd.DataFrame,
    current_sheet_name: str,
) -> ProductivityResult:
    """Build a productivity report using server sheets and the current list."""
    old_records, skipped_empty_sheets = _prepare_server_records(server_sheets)
    current_prepared = _prepare_current_records(current_dataframe)

    old_valid = old_records[old_records[PROCESS_KEY].ne("")].copy()
    current_valid = current_prepared[current_prepared[PROCESS_KEY].ne("")]

    if old_valid.empty:
        raise ProductivityError(
            "A planilha dos servidores não possui nenhum CNJ válido."
        )
    if current_valid.empty:
        raise ProductivityError("A planilha atual não possui nenhum CNJ válido.")

    old_keys = set(old_valid[PROCESS_KEY])
    current_keys = set(current_valid[PROCESS_KEY])

    old_by_server_process = _aggregate_old_by_server_process(old_valid)
    old_by_process = _aggregate_old_by_process(old_valid)

    old_by_server_process[COMPARISON_STATUS_COLUMN] = np.where(
        old_by_server_process[PROCESS_KEY].isin(current_keys),
        STATUS_REMAINED,
        STATUS_PRODUCTIVE,
    )

    old_comparison = _finalize_old_comparison(old_by_server_process)
    productive_processes = old_comparison[
        old_comparison[COMPARISON_STATUS_COLUMN].eq(STATUS_PRODUCTIVE)
    ].reset_index(drop=True)

    current_enriched = _build_current_enriched(
        current_dataframe=current_dataframe,
        current_keys=current_prepared[PROCESS_KEY],
        old_by_process=old_by_process,
        old_keys=old_keys,
    )
    new_processes = current_enriched[
        current_enriched[COMPARISON_STATUS_COLUMN].eq(STATUS_NEW)
    ].reset_index(drop=True)

    summary_by_server = _build_server_summary(old_by_server_process)
    duplicate_reports = tuple(
        report
        for report in (
            _build_duplicate_report(
                old_valid,
                "Lista dos servidores",
                server_column=SERVER_COLUMN,
            ),
            _build_duplicate_report(current_valid, "Planilha atual"),
        )
        if report is not None
    )

    productive_unique = old_valid.loc[
        ~old_valid[PROCESS_KEY].isin(current_keys),
        PROCESS_KEY,
    ].nunique()
    remained_unique = old_valid.loc[
        old_valid[PROCESS_KEY].isin(current_keys),
        PROCESS_KEY,
    ].nunique()
    new_unique = current_valid.loc[
        ~current_valid[PROCESS_KEY].isin(old_keys),
        PROCESS_KEY,
    ].nunique()

    return ProductivityResult(
        server_count=len(server_sheets) - len(skipped_empty_sheets),
        old_total_rows=len(old_records),
        old_valid_rows=len(old_valid),
        old_unique_processes=old_valid[PROCESS_KEY].nunique(),
        current_total_rows=len(current_dataframe),
        current_valid_rows=len(current_valid),
        current_unique_processes=current_valid[PROCESS_KEY].nunique(),
        productive_unique_processes=int(productive_unique),
        remained_unique_processes=int(remained_unique),
        new_unique_processes=int(new_unique),
        old_empty_cnj_rows=int(old_records[PROCESS_KEY].eq("").sum()),
        current_empty_cnj_rows=int(current_prepared[PROCESS_KEY].eq("").sum()),
        current_sheet_name=current_sheet_name,
        skipped_empty_sheets=tuple(skipped_empty_sheets),
        duplicate_reports=duplicate_reports,
        summary_by_server=summary_by_server,
        productive_processes=productive_processes,
        new_processes=new_processes,
        current_enriched=current_enriched,
        old_comparison=old_comparison,
    )


def _prepare_server_records(
    server_sheets: dict[str, pd.DataFrame],
) -> tuple[pd.DataFrame, list[str]]:
    prepared_frames: list[pd.DataFrame] = []
    skipped_empty_sheets: list[str] = []

    for sheet_name, dataframe in server_sheets.items():
        if dataframe.empty or dataframe.shape[1] == 0:
            skipped_empty_sheets.append(sheet_name)
            continue

        process_column = _find_column(
            dataframe,
            (OLD_PROCESS_COLUMN, "Numero do Processo", "Número Processo"),
            f"aba {sheet_name}",
        )
        status_column = _find_column(
            dataframe,
            (OLD_STATUS_COLUMN, "Situacao"),
            f"aba {sheet_name}",
        )
        note_column = _find_column(
            dataframe,
            (OLD_NOTE_COLUMN, "Observacao"),
            f"aba {sheet_name}",
        )

        base_columns = {process_column, status_column, note_column}
        extra_columns = [
            column for column in dataframe.columns if column not in base_columns
        ]

        prepared = pd.DataFrame(
            {
                SERVER_COLUMN: sheet_name,
                OLD_PROCESS_COLUMN: dataframe[process_column],
                PREVIOUS_STATUS_COLUMN: dataframe[status_column],
                PREVIOUS_NOTE_COLUMN: dataframe[note_column],
            }
        )

        for extra_column in extra_columns:
            prepared[str(extra_column)] = dataframe[extra_column]

        prepared[PROCESS_KEY] = normalize_cnj_series(prepared[OLD_PROCESS_COLUMN])
        prepared_frames.append(prepared)

    if not prepared_frames:
        raise ProductivityError(
            "Nenhuma aba válida foi encontrada na planilha dos servidores."
        )

    return pd.concat(prepared_frames, ignore_index=True), skipped_empty_sheets


def _prepare_current_records(current_dataframe: pd.DataFrame) -> pd.DataFrame:
    if current_dataframe.empty or current_dataframe.shape[1] == 0:
        raise ProductivityError("A planilha atual está vazia.")

    for required_column in CURRENT_REQUIRED_COLUMNS:
        _find_column(current_dataframe, (required_column,), "planilha atual")

    process_column = _find_column(
        current_dataframe,
        (CURRENT_PROCESS_COLUMN, "Numero Processo", "Número do Processo"),
        "planilha atual",
    )

    prepared = current_dataframe.copy()
    prepared[PROCESS_KEY] = normalize_cnj_series(prepared[process_column])
    return prepared


def _aggregate_old_by_server_process(old_valid: pd.DataFrame) -> pd.DataFrame:
    aggregation_columns = _old_export_columns(old_valid, include_server=False)
    aggregations = _old_aggregations(aggregation_columns)

    return (
        old_valid.groupby([SERVER_COLUMN, PROCESS_KEY], sort=False)
        .agg(aggregations)
        .reset_index()
    )


def _aggregate_old_by_process(old_valid: pd.DataFrame) -> pd.DataFrame:
    aggregation_columns = _old_export_columns(old_valid, include_server=True)
    aggregations = _old_aggregations(aggregation_columns)

    return old_valid.groupby(PROCESS_KEY, sort=False).agg(aggregations)


def _old_export_columns(
    dataframe: pd.DataFrame,
    include_server: bool,
) -> list[str]:
    ignored = {PROCESS_KEY}
    if not include_server:
        ignored.add(SERVER_COLUMN)
    return [str(column) for column in dataframe.columns if column not in ignored]


def _old_aggregations(columns: list[str]) -> dict[str, Any]:
    aggregations: dict[str, Any] = {}
    for column in columns:
        if column == OLD_PROCESS_COLUMN:
            aggregations[column] = first_non_empty_value
        else:
            aggregations[column] = join_unique_values
    return aggregations


def _finalize_old_comparison(old_by_server_process: pd.DataFrame) -> pd.DataFrame:
    ordered_columns = [
        SERVER_COLUMN,
        OLD_PROCESS_COLUMN,
        PREVIOUS_STATUS_COLUMN,
        PREVIOUS_NOTE_COLUMN,
        COMPARISON_STATUS_COLUMN,
    ]
    extra_columns = [
        column
        for column in old_by_server_process.columns
        if column not in {*ordered_columns, PROCESS_KEY}
    ]
    ordered_columns.extend(extra_columns)
    return old_by_server_process[ordered_columns].sort_values(
        [SERVER_COLUMN, COMPARISON_STATUS_COLUMN, OLD_PROCESS_COLUMN],
        ascending=[True, True, True],
        kind="stable",
    ).reset_index(drop=True)


def _build_current_enriched(
    current_dataframe: pd.DataFrame,
    current_keys: pd.Series,
    old_by_process: pd.DataFrame,
    old_keys: set[str],
) -> pd.DataFrame:
    enriched = current_dataframe.copy()
    valid_current_key = current_keys.ne("")

    enriched[SERVER_COLUMN] = current_keys.map(
        old_by_process[SERVER_COLUMN]
    ).fillna("")
    enriched.loc[valid_current_key & ~current_keys.isin(old_keys), SERVER_COLUMN] = (
        STATUS_NEW.upper()
    )

    for source_column in old_by_process.columns:
        if source_column in {OLD_PROCESS_COLUMN, SERVER_COLUMN}:
            continue
        target_column = _previous_column_name(source_column)
        enriched[target_column] = current_keys.map(
            old_by_process[source_column]
        ).fillna("")

    enriched[COMPARISON_STATUS_COLUMN] = np.select(
        [
            current_keys.eq(""),
            current_keys.isin(old_keys),
            ~current_keys.isin(old_keys),
        ],
        [STATUS_EMPTY_CNJ, STATUS_REMAINED, STATUS_NEW],
        default=STATUS_NEW,
    )
    return enriched


def _previous_column_name(source_column: str) -> str:
    if source_column in {PREVIOUS_STATUS_COLUMN, PREVIOUS_NOTE_COLUMN}:
        return source_column
    return f"{source_column} anterior"


def _build_server_summary(old_by_server_process: pd.DataFrame) -> pd.DataFrame:
    grouped = old_by_server_process.groupby(SERVER_COLUMN, sort=False)
    summary = grouped.agg(
        **{
            "Processos na lista 01.26": (PROCESS_KEY, "nunique"),
            "Produtivos (saíram)": (
                COMPARISON_STATUS_COLUMN,
                lambda values: int(values.eq(STATUS_PRODUCTIVE).sum()),
            ),
            "Ainda no arquivo 120 dias": (
                COMPARISON_STATUS_COLUMN,
                lambda values: int(values.eq(STATUS_REMAINED).sum()),
            ),
        }
    ).reset_index()

    total_productive = summary["Produtivos (saíram)"].sum()
    summary["% produtividade"] = (
        summary["Produtivos (saíram)"]
        .div(summary["Processos na lista 01.26"])
        .mul(100)
        .round(1)
    )
    if total_productive:
        summary["Participação na produtividade"] = (
            summary["Produtivos (saíram)"].div(total_productive).mul(100).round(1)
        )
    else:
        summary["Participação na produtividade"] = 0.0

    return summary.sort_values(
        ["Produtivos (saíram)", "% produtividade", SERVER_COLUMN],
        ascending=[False, False, True],
        kind="stable",
    ).reset_index(drop=True)


def _build_duplicate_report(
    valid_dataframe: pd.DataFrame,
    source_label: str,
    server_column: str | None = None,
) -> DuplicateReport | None:
    duplicated = valid_dataframe[valid_dataframe.duplicated(PROCESS_KEY, keep=False)]
    if duplicated.empty:
        return None

    counts = duplicated[PROCESS_KEY].value_counts()
    sample_records = []
    process_column = _process_column_for_duplicate_report(duplicated)
    for process_key in counts.head(10).index:
        rows = duplicated[duplicated[PROCESS_KEY].eq(process_key)]
        record = {
            "CNJ normalizado": process_key,
            "Número do Processo": first_non_empty_value(rows[process_column]),
            "Ocorrências": int(counts.loc[process_key]),
        }
        if server_column and server_column in rows.columns:
            record[SERVER_COLUMN] = join_unique_values(rows[server_column])
        sample_records.append(record)

    return DuplicateReport(
        source_label=source_label,
        duplicated_processes=int(counts.size),
        duplicated_rows=int(duplicated.shape[0]),
        sample=pd.DataFrame(sample_records),
    )


def _process_column_for_duplicate_report(dataframe: pd.DataFrame) -> Any:
    if OLD_PROCESS_COLUMN in dataframe.columns:
        return OLD_PROCESS_COLUMN
    if CURRENT_PROCESS_COLUMN in dataframe.columns:
        return CURRENT_PROCESS_COLUMN
    return dataframe.columns[0]


def _find_column(
    dataframe: pd.DataFrame,
    candidates: tuple[str, ...],
    sheet_label: str,
) -> Any:
    normalized_columns = {_normalize_label(column): column for column in dataframe.columns}
    for candidate in candidates:
        column = normalized_columns.get(_normalize_label(candidate))
        if column is not None:
            return column

    readable_candidates = ", ".join(candidates)
    raise ProductivityError(
        f"{sheet_label}: coluna obrigatória não encontrada "
        f"({readable_candidates})."
    )


def _normalize_label(value: Any) -> str:
    text = str(value).strip()
    ascii_text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore")
    return " ".join(ascii_text.decode("ascii").casefold().split())
