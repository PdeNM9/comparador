"""Utilities for reading, validating and normalizing Excel data."""

from __future__ import annotations

from io import BytesIO
from pathlib import Path
from typing import Any, Iterable
from zipfile import BadZipFile

import pandas as pd
from openpyxl.utils.exceptions import InvalidFileException


class ExcelProcessingError(ValueError):
    """Raised when an uploaded workbook cannot be read safely."""


def validate_xlsx_filename(file_name: str) -> None:
    """Validate that the uploaded file has the expected .xlsx extension."""
    suffix = Path(file_name or "").suffix.lower()
    if suffix != ".xlsx":
        raise ExcelProcessingError(
            "Formato inválido. Envie um arquivo Excel no formato .xlsx."
        )


def read_sheet_names(file_bytes: bytes, file_name: str) -> list[str]:
    """Return all worksheet names from an .xlsx workbook."""
    validate_xlsx_filename(file_name)
    if not file_bytes:
        raise ExcelProcessingError("O arquivo enviado está vazio.")

    try:
        workbook = pd.ExcelFile(BytesIO(file_bytes), engine="openpyxl")
        try:
            sheet_names = list(workbook.sheet_names)
        finally:
            workbook.close()
    except (BadZipFile, InvalidFileException, ValueError, OSError) as exc:
        raise ExcelProcessingError(
            "Não foi possível abrir o arquivo Excel. Verifique se o arquivo "
            "não está corrompido e se realmente está no formato .xlsx."
        ) from exc

    if not sheet_names:
        raise ExcelProcessingError("A planilha não possui abas para leitura.")
    return sheet_names


def read_excel_sheet(
    file_bytes: bytes,
    file_name: str,
    sheet_name: str,
) -> pd.DataFrame:
    """Read one worksheet preserving the values needed for a later export."""
    validate_xlsx_filename(file_name)
    if not sheet_name:
        raise ExcelProcessingError("Selecione uma aba válida para leitura.")

    try:
        dataframe = pd.read_excel(
            BytesIO(file_bytes),
            sheet_name=sheet_name,
            engine="openpyxl",
            keep_default_na=False,
        )
    except (BadZipFile, InvalidFileException, ValueError, OSError) as exc:
        raise ExcelProcessingError(
            f"Não foi possível ler a aba '{sheet_name}'. Confirme se a aba "
            "existe e se o arquivo é um .xlsx válido."
        ) from exc

    if dataframe.empty or dataframe.shape[1] == 0:
        raise ExcelProcessingError(
            f"A aba '{sheet_name}' está vazia ou não possui colunas."
        )
    return dataframe


def ensure_columns_exist(
    dataframe: pd.DataFrame,
    columns: Iterable[Any],
    sheet_label: str,
) -> None:
    """Raise a friendly error if any selected column is absent."""
    missing = [column for column in columns if column not in dataframe.columns]
    if missing:
        readable = ", ".join(str(column) for column in missing)
        raise ExcelProcessingError(
            f"{sheet_label}: as seguintes colunas não foram encontradas: "
            f"{readable}."
        )


def normalize_cnj(value: Any) -> str:
    """Normalize a CNJ number by keeping digits only.

    This makes values such as ``0000000-00.0000.0.00.0000`` and
    ``00000000000000000000`` comparable.
    """
    if value is None:
        return ""

    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass

    if isinstance(value, float) and value.is_integer():
        text = str(int(value))
    else:
        text = str(value).strip()

    if text.endswith(".0") and text[:-2].isdigit():
        text = text[:-2]

    return "".join(character for character in text if character.isdigit())


def normalize_cnj_series(series: pd.Series) -> pd.Series:
    """Normalize an entire column of CNJ values."""
    return series.map(normalize_cnj)


def value_to_display(value: Any) -> str:
    """Convert scalar spreadsheet values to a clean display string."""
    if value is None:
        return ""

    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass

    if isinstance(value, float) and value.is_integer():
        return str(int(value))

    text = str(value).strip()
    if text.casefold() in {"nan", "nat", "none"}:
        return ""
    return text


def first_non_empty_value(values: Iterable[Any]) -> str:
    """Return the first non-empty value in source order."""
    for value in values:
        text = value_to_display(value)
        if text:
            return text
    return ""


def join_unique_values(values: Iterable[Any], separator: str = " | ") -> str:
    """Concatenate unique non-empty values preserving their original order."""
    seen: set[str] = set()
    output: list[str] = []

    for value in values:
        text = value_to_display(value)
        if not text or text in seen:
            continue
        seen.add(text)
        output.append(text)

    return separator.join(output)
