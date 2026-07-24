"""Streamlit dashboard components."""

from __future__ import annotations

from typing import Iterable
import unicodedata

import pandas as pd
import streamlit as st

from comparison import ComparisonResult


def configure_page() -> None:
    """Configure Streamlit before any page element is rendered."""
    st.set_page_config(
        page_title="Comparador de Planilhas Judiciais",
        layout="wide",
        initial_sidebar_state="expanded",
    )


def inject_styles() -> None:
    """Apply a restrained visual layer over Streamlit defaults."""
    st.markdown(
        """
        <style>
        :root {
            --surface: #ffffff;
            --muted: #64748b;
            --line: #e2e8f0;
            --ink: #111827;
            --blue: #3a6ea5;
            --green: #2a9d8f;
            --red: #d95f59;
        }
        .block-container {
            padding-top: 2rem;
            padding-bottom: 3rem;
        }
        .app-header {
            border-bottom: 1px solid var(--line);
            margin-bottom: 1.2rem;
            padding-bottom: 1rem;
        }
        .app-header h1 {
            color: var(--ink);
            font-size: 2rem;
            letter-spacing: 0;
            margin: 0;
        }
        .app-header p {
            color: var(--muted);
            margin: 0 0 .25rem 0;
            text-transform: uppercase;
            font-size: .78rem;
            font-weight: 700;
        }
        .metric-card {
            background: var(--surface);
            border: 1px solid var(--line);
            border-left: 5px solid var(--blue);
            border-radius: 8px;
            min-height: 118px;
            padding: .95rem 1rem;
            box-shadow: 0 10px 28px rgba(15, 23, 42, .05);
        }
        .metric-card.green { border-left-color: var(--green); }
        .metric-card.red { border-left-color: var(--red); }
        .metric-label {
            color: var(--muted);
            display: block;
            font-size: .82rem;
            font-weight: 700;
            margin-bottom: .4rem;
            text-transform: uppercase;
        }
        .metric-value {
            color: var(--ink);
            display: block;
            font-size: 1.85rem;
            font-weight: 800;
            line-height: 1.2;
        }
        .metric-note {
            color: var(--muted);
            display: block;
            font-size: .82rem;
            margin-top: .35rem;
        }
        [data-testid="stSidebar"] {
            border-right: 1px solid var(--line);
        }
        [data-testid="stFileUploader"] section {
            border-radius: 8px;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_header() -> None:
    st.markdown(
        """
        <div class="app-header">
            <p>Comparação de processos judiciais</p>
            <h1>Comparador de planilhas Excel</h1>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_metric_cards(result: ComparisonResult) -> None:
    """Render dashboard cards with totals and comparison counts."""
    metrics = [
        (
            "Total Planilha 1",
            result.old_total_rows,
            f"{format_int(result.old_unique_processes)} CNJs únicos",
            "",
        ),
        (
            "Total Planilha 2",
            result.new_total_rows,
            f"{format_int(result.new_unique_processes)} CNJs únicos",
            "",
        ),
        (
            "Excluídos",
            len(result.excluded_processes),
            "existiam apenas na Planilha 1",
            "red",
        ),
        (
            "Novos",
            len(result.new_processes),
            "existem apenas na Planilha 2",
            "green",
        ),
        (
            "Mantidos",
            len(result.maintained_processes),
            "presentes nas duas planilhas",
            "",
        ),
    ]

    columns = st.columns(len(metrics))
    for column, (label, value, note, tone) in zip(columns, metrics):
        with column:
            st.markdown(
                f"""
                <div class="metric-card {tone}">
                    <span class="metric-label">{label}</span>
                    <span class="metric-value">{format_int(value)}</span>
                    <span class="metric-note">{note}</span>
                </div>
                """,
                unsafe_allow_html=True,
            )


def render_quality_messages(result: ComparisonResult) -> None:
    """Show duplicate and empty-CNJ notices without blocking the workflow."""
    if result.old_empty_cnj_rows or result.new_empty_cnj_rows:
        st.info(
            "Linhas com CNJ vazio foram preservadas no arquivo final, mas "
            "não entraram no cálculo de novos, excluídos ou mantidos. "
            f"Planilha 1: {format_int(result.old_empty_cnj_rows)}; "
            f"Planilha 2: {format_int(result.new_empty_cnj_rows)}."
        )

    for report in result.duplicate_reports:
        st.warning(
            f"{report.sheet_label}: {format_int(report.duplicated_processes)} "
            "CNJs duplicados encontrados em "
            f"{format_int(report.duplicated_rows)} linhas. A comparação "
            "continuou e valores de responsável/anotação foram concatenados "
            "por CNJ."
        )
        with st.expander(f"Amostra de duplicados - {report.sheet_label}"):
            st.dataframe(
                report.sample,
                hide_index=True,
                use_container_width=True,
            )


def render_filterable_table(
    title: str,
    dataframe: pd.DataFrame,
    key: str,
    height: int = 380,
) -> pd.DataFrame:
    """Render a table with a simple text filter over all visible columns."""
    st.subheader(title)
    query = st.text_input(
        f"Filtrar {title.lower()}",
        key=f"filter_{key}",
        placeholder="Digite parte do CNJ ou do responsável",
    )
    filtered = filter_dataframe(dataframe, query)
    st.caption(
        f"{format_int(len(filtered))} de {format_int(len(dataframe))} "
        "registros exibidos"
    )
    st.dataframe(
        filtered,
        hide_index=True,
        use_container_width=True,
        height=height,
    )
    return filtered


def filter_dataframe(dataframe: pd.DataFrame, query: str) -> pd.DataFrame:
    """Filter a dataframe by searching every column as text."""
    query = (query or "").strip()
    if not query or dataframe.empty:
        return dataframe

    normalized_query = query.casefold()
    mask = pd.Series(False, index=dataframe.index)
    for column in dataframe.columns:
        mask = mask | dataframe[column].astype(str).str.casefold().str.contains(
            normalized_query,
            regex=False,
            na=False,
        )
    return dataframe[mask]


def guess_column(columns: Iterable[object], preferred_terms: Iterable[str]):
    """Return the first column whose label contains one preferred term."""
    columns = list(columns)
    normalized_terms = [_normalize_text(term) for term in preferred_terms]

    for term in normalized_terms:
        for column in columns:
            if term in _normalize_text(column):
                return column

    return columns[0] if columns else None


def guess_annotation_columns(columns: Iterable[object]) -> list[object]:
    """Guess annotation-like columns while still requiring user confirmation."""
    terms = (
        "anot",
        "observ",
        "coment",
        "nota",
        "andamento",
        "providencia",
        "pendencia",
    )
    guessed = []
    for column in columns:
        normalized = _normalize_text(column)
        if any(term in normalized for term in terms):
            guessed.append(column)
    return guessed


def format_int(value: int) -> str:
    """Format integers using the Brazilian thousands separator."""
    return f"{int(value):,}".replace(",", ".")


def _normalize_text(value: object) -> str:
    text = str(value)
    ascii_text = unicodedata.normalize("NFKD", text).encode(
        "ascii",
        "ignore",
    )
    return ascii_text.decode("ascii").casefold()
