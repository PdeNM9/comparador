"""Streamlit dashboard components."""

from __future__ import annotations

import pandas as pd
import streamlit as st


def configure_page() -> None:
    """Configure Streamlit before any page element is rendered."""
    st.set_page_config(
        page_title="Produtividade por Servidor",
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
            --yellow: #d9a441;
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
        .metric-card.yellow { border-left-color: var(--yellow); }
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
    """Render the app header."""
    st.markdown(
        """
        <div class="app-header">
            <p>Produtividade judicial</p>
            <h1>Comparativo por servidor</h1>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_productivity_metric_cards(result) -> None:
    """Render dashboard cards with productivity counters."""
    metrics = [
        ("Servidores", result.server_count, "abas processadas", ""),
        ("Lista 01.26", result.old_unique_processes, "processos únicos", ""),
        ("Arquivo 120 dias", result.current_unique_processes, "processos atuais", ""),
        (
            "Produtivos",
            result.productive_unique_processes,
            "saíram do arquivo 120 dias",
            "green",
        ),
        ("Novos", result.new_unique_processes, "não estavam na lista 01.26", "yellow"),
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


def render_productivity_quality_messages(result) -> None:
    """Show duplicate and empty-CNJ notices without blocking the workflow."""
    if result.old_empty_cnj_rows or result.current_empty_cnj_rows:
        st.info(
            "Linhas com CNJ vazio não entram na produtividade. "
            f"Lista 01.26: {format_int(result.old_empty_cnj_rows)}; "
            f"arquivo atual: {format_int(result.current_empty_cnj_rows)}."
        )

    if result.skipped_empty_sheets:
        sheets = ", ".join(result.skipped_empty_sheets)
        st.info(f"Abas vazias ignoradas: {sheets}.")

    for report in result.duplicate_reports:
        st.warning(
            f"{report.source_label}: {format_int(report.duplicated_processes)} "
            "CNJs duplicados encontrados em "
            f"{format_int(report.duplicated_rows)} linhas. A comparação "
            "continuou com consolidação por CNJ."
        )
        with st.expander(f"Amostra de duplicados - {report.source_label}"):
            st.dataframe(
                report.sample,
                hide_index=True,
                width="stretch",
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
        placeholder="Digite parte do CNJ, servidor ou situação",
    )
    filtered = filter_dataframe(dataframe, query)
    st.caption(
        f"{format_int(len(filtered))} de {format_int(len(dataframe))} "
        "registros exibidos"
    )
    st.dataframe(
        filtered,
        hide_index=True,
        width="stretch",
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


def format_int(value: int) -> str:
    """Format integers using the Brazilian thousands separator."""
    return f"{int(value):,}".replace(",", ".")


def format_percent(value: float) -> str:
    """Format a percentage using the Brazilian decimal separator."""
    return f"{float(value):.1f}%".replace(".", ",")
