"""Streamlit app for productivity comparison by judicial server."""

from __future__ import annotations

from dataclasses import dataclass
import hashlib
from pathlib import Path

import pandas as pd
import streamlit as st

from dashboard import (
    configure_page,
    format_int,
    inject_styles,
    render_filterable_table,
    render_header,
    render_productivity_metric_cards,
    render_productivity_quality_messages,
)
from excel_utils import ExcelProcessingError, read_excel_sheet, read_sheet_names
from export import (
    EXCEL_MIME_TYPE,
    dataframe_to_excel_bytes,
    timestamped_filename,
    workbook_to_excel_bytes,
)
from productivity import (
    ProductivityError,
    ProductivityResult,
    build_productivity_report,
    read_workbook_sheets,
)


SERVER_FILE_KEY = "server_file"
CURRENT_FILE_KEY = "current_file"
RESULT_KEY = "productivity_result"
SIGNATURE_KEY = "productivity_signature"
STATUS_COLORS = {
    "Produtivos": "#2A9D8F",
    "Permaneceram": "#3A6EA5",
    "Novos": "#D9A441",
}


@dataclass(frozen=True)
class WorkbookInput:
    """Workbook bytes selected by upload or local project file."""

    name: str
    bytes_data: bytes
    source_label: str


@st.cache_data(show_spinner=False)
def cached_sheet_names(file_bytes: bytes, file_name: str) -> list[str]:
    """Cache worksheet discovery for each uploaded file."""
    return read_sheet_names(file_bytes, file_name)


@st.cache_data(show_spinner=False)
def cached_read_sheet(
    file_bytes: bytes,
    file_name: str,
    sheet_name: str,
) -> pd.DataFrame:
    """Cache worksheet loading for the selected current sheet."""
    return read_excel_sheet(file_bytes, file_name, sheet_name)


@st.cache_data(show_spinner=False)
def cached_read_workbook(
    file_bytes: bytes,
    file_name: str,
) -> dict[str, pd.DataFrame]:
    """Cache loading of every server worksheet."""
    return read_workbook_sheets(file_bytes, file_name)


@st.cache_data(show_spinner=False)
def cached_local_file(path: str, mtime_ns: int) -> bytes:
    """Cache local workbook bytes from the project folder."""
    return Path(path).read_bytes()


@st.cache_data(show_spinner=False)
def cached_excel_bytes(dataframe: pd.DataFrame, sheet_name: str) -> bytes:
    """Cache one-table Excel serialization."""
    return dataframe_to_excel_bytes(dataframe, sheet_name)


@st.cache_data(show_spinner=False)
def cached_report_bytes(
    summary_by_server: pd.DataFrame,
    current_enriched: pd.DataFrame,
    productive_processes: pd.DataFrame,
    new_processes: pd.DataFrame,
    old_comparison: pd.DataFrame,
) -> bytes:
    """Cache the full multi-sheet report."""
    return workbook_to_excel_bytes(
        {
            "Resumo por servidor": summary_by_server,
            "Atual enriquecida": current_enriched,
            "Produtivos - saíram": productive_processes,
            "Novos": new_processes,
            "Lista antiga comparada": old_comparison,
        }
    )


def main() -> None:
    """Run the Streamlit application."""
    configure_page()
    inject_styles()
    render_header()

    workbook_inputs = _render_input_area()
    if workbook_inputs is None:
        st.info("Envie os dois arquivos .xlsx ou use os arquivos detectados na pasta.")
        return

    server_workbook, current_workbook = workbook_inputs

    try:
        server_sheet_names = cached_sheet_names(
            server_workbook.bytes_data,
            server_workbook.name,
        )
        current_sheet_names = cached_sheet_names(
            current_workbook.bytes_data,
            current_workbook.name,
        )
    except ExcelProcessingError as exc:
        st.error(str(exc))
        return

    current_sheet_name = _render_sidebar_configuration(
        server_workbook,
        current_workbook,
        server_sheet_names,
        current_sheet_names,
    )
    current_signature = _build_signature(
        server_workbook.bytes_data,
        current_workbook.bytes_data,
        current_sheet_name,
    )

    if st.sidebar.button(
        "Processar produtividade",
        type="primary",
        width="stretch",
    ):
        _run_productivity_report(
            server_workbook,
            current_workbook,
            current_sheet_name,
            current_signature,
        )

    result = st.session_state.get(RESULT_KEY)
    if result is None:
        st.info("Clique em Processar produtividade para montar o painel.")
        return

    if st.session_state.get(SIGNATURE_KEY) != current_signature:
        st.warning(
            "Arquivos ou aba atual mudaram depois do último processamento. "
            "Clique em Processar produtividade para atualizar o painel."
        )

    _render_results(result)


def _render_input_area() -> tuple[WorkbookInput, WorkbookInput] | None:
    st.subheader("Arquivos")
    left, right = st.columns(2)
    with left:
        server_upload = st.file_uploader(
            "Lista 01.26 - abas por servidor",
            type=["xlsx"],
            key=SERVER_FILE_KEY,
        )
    with right:
        current_upload = st.file_uploader(
            "Arquivo atual - 120 dias",
            type=["xlsx"],
            key=CURRENT_FILE_KEY,
        )

    local_files = _local_xlsx_files()
    use_local = False
    if local_files and not (server_upload and current_upload):
        use_local = st.checkbox(
            "Usar arquivos .xlsx detectados na pasta do projeto",
            value=server_upload is None and current_upload is None,
        )

    if use_local:
        local_left, local_right = st.columns(2)
        with local_left:
            server_path = st.selectbox(
                "Arquivo dos servidores",
                options=local_files,
                index=_guess_file_index(local_files, ("lista", "01.26")),
                format_func=lambda path: path.name,
            )
        with local_right:
            current_path = st.selectbox(
                "Arquivo atual",
                options=local_files,
                index=_guess_file_index(local_files, ("120", "dias")),
                format_func=lambda path: path.name,
            )

        if server_path == current_path:
            st.error("Selecione dois arquivos diferentes.")
            return None

        try:
            return (
                _local_workbook_input(server_path),
                _local_workbook_input(current_path),
            )
        except OSError as exc:
            st.error(f"Não foi possível ler os arquivos locais: {exc}")
            return None

    if not server_upload or not current_upload:
        return None

    return (
        WorkbookInput(
            name=server_upload.name,
            bytes_data=server_upload.getvalue(),
            source_label="upload",
        ),
        WorkbookInput(
            name=current_upload.name,
            bytes_data=current_upload.getvalue(),
            source_label="upload",
        ),
    )


def _render_sidebar_configuration(
    server_workbook: WorkbookInput,
    current_workbook: WorkbookInput,
    server_sheet_names: list[str],
    current_sheet_names: list[str],
) -> str:
    with st.sidebar:
        st.header("Configuração")
        st.caption(f"Servidores: {server_workbook.name}")
        st.caption(f"Abas detectadas: {format_int(len(server_sheet_names))}")
        st.caption(f"Atual: {current_workbook.name}")

        if len(current_sheet_names) == 1:
            current_sheet_name = current_sheet_names[0]
            st.caption(f"Aba atual: {current_sheet_name}")
        else:
            current_sheet_name = st.selectbox(
                "Aba do arquivo atual",
                options=current_sheet_names,
                index=0,
            )

        with st.expander("Servidores detectados"):
            st.write(", ".join(server_sheet_names))

    return current_sheet_name


def _run_productivity_report(
    server_workbook: WorkbookInput,
    current_workbook: WorkbookInput,
    current_sheet_name: str,
    current_signature: str,
) -> None:
    progress = st.progress(0, text="Lendo abas dos servidores")
    try:
        server_sheets = cached_read_workbook(
            server_workbook.bytes_data,
            server_workbook.name,
        )
        progress.progress(30, text="Lendo arquivo atual")
        current_dataframe = cached_read_sheet(
            current_workbook.bytes_data,
            current_workbook.name,
            current_sheet_name,
        )
        progress.progress(55, text="Normalizando CNJs")
        result = build_productivity_report(
            server_sheets=server_sheets,
            current_dataframe=current_dataframe,
            current_sheet_name=current_sheet_name,
        )
        progress.progress(85, text="Preparando tabelas e exportações")
        st.session_state[RESULT_KEY] = result
        st.session_state[SIGNATURE_KEY] = current_signature
        progress.progress(100, text="Produtividade processada")
        st.success("Produtividade processada com sucesso.")
    except (ExcelProcessingError, ProductivityError) as exc:
        st.error(str(exc))
    finally:
        progress.empty()


def _render_results(result: ProductivityResult) -> None:
    render_productivity_metric_cards(result)
    render_productivity_quality_messages(result)

    st.subheader("Produtividade por servidor")
    chart_left, chart_right = st.columns(2)
    with chart_left:
        _render_plotly_chart(
            build_productivity_by_server_chart(result.summary_by_server),
            width="stretch",
            config={"displayModeBar": False},
        )
    with chart_right:
        _render_plotly_chart(
            build_productivity_rate_chart(result.summary_by_server),
            width="stretch",
            config={"displayModeBar": False},
        )

    st.subheader("Distribuição")
    _render_plotly_chart(
        build_pie_chart(result.status_counts),
        width="stretch",
        config={"displayModeBar": False},
    )

    _render_exports(result)
    _render_tables(result)


def _render_exports(result: ProductivityResult) -> None:
    st.subheader("Exportações")
    first, second, third, fourth = st.columns(4)
    with first:
        st.download_button(
            "Relatório completo",
            data=cached_report_bytes(
                result.summary_by_server,
                result.current_enriched,
                result.productive_processes,
                result.new_processes,
                result.old_comparison,
            ),
            file_name=timestamped_filename("relatorio_produtividade"),
            mime=EXCEL_MIME_TYPE,
            width="stretch",
        )
    with second:
        st.download_button(
            "Produtivos",
            data=cached_excel_bytes(result.productive_processes, "Produtivos"),
            file_name=timestamped_filename("processos_produtivos"),
            mime=EXCEL_MIME_TYPE,
            width="stretch",
            disabled=result.productive_processes.empty,
        )
    with third:
        st.download_button(
            "Novos",
            data=cached_excel_bytes(result.new_processes, "Novos"),
            file_name=timestamped_filename("processos_novos"),
            mime=EXCEL_MIME_TYPE,
            width="stretch",
            disabled=result.new_processes.empty,
        )
    with fourth:
        st.download_button(
            "Atual enriquecida",
            data=cached_excel_bytes(result.current_enriched, "Atual enriquecida"),
            file_name=timestamped_filename("planilha_atual_enriquecida"),
            mime=EXCEL_MIME_TYPE,
            width="stretch",
        )


def _render_tables(result: ProductivityResult) -> None:
    st.subheader("Tabelas")
    summary_tab, productive_tab, new_tab, current_tab, old_tab = st.tabs(
        [
            "Resumo",
            "Produtivos",
            "Novos",
            "Atual enriquecida",
            "Lista antiga",
        ]
    )
    with summary_tab:
        st.dataframe(
            result.summary_by_server,
            hide_index=True,
            width="stretch",
        )
    with productive_tab:
        render_filterable_table(
            "Processos produtivos",
            result.productive_processes,
            "productive",
        )
    with new_tab:
        render_filterable_table(
            "Processos novos",
            result.new_processes,
            "new",
        )
    with current_tab:
        render_filterable_table(
            "Planilha atual enriquecida",
            result.current_enriched,
            "current",
            height=460,
        )
    with old_tab:
        render_filterable_table(
            "Lista 01.26 comparada",
            result.old_comparison,
            "old",
            height=460,
        )


def _render_plotly_chart(figure, **kwargs) -> None:
    if figure is None:
        st.warning(
            "Os gráficos não puderam ser carregados. Confirme se o pacote "
            "`plotly` está instalado no ambiente do Streamlit."
        )
        return
    st.plotly_chart(figure, **kwargs)


def build_productivity_by_server_chart(summary):
    """Create a stacked bar chart with productive and remaining processes."""
    go = _load_plotly()
    if go is None:
        return None
    if summary.empty:
        return _empty_figure(go, "Sem dados para exibir")

    ordered = summary.sort_values("Produtivos (saíram)", ascending=True)
    figure = go.Figure()
    figure.add_bar(
        y=ordered["Servidor"],
        x=ordered["Produtivos (saíram)"],
        name="Produtivos",
        orientation="h",
        marker_color="#2A9D8F",
        hovertemplate="%{y}: %{x} produtivos<extra></extra>",
    )
    figure.add_bar(
        y=ordered["Servidor"],
        x=ordered["Ainda no arquivo 120 dias"],
        name="Ainda no 120 dias",
        orientation="h",
        marker_color="#3A6EA5",
        hovertemplate="%{y}: %{x} ainda no arquivo<extra></extra>",
    )
    figure.update_layout(
        barmode="stack",
        height=420,
        margin=dict(l=24, r=24, t=30, b=24),
        template="plotly_white",
        xaxis_title="Processos",
        yaxis_title=None,
        legend=dict(orientation="h", yanchor="bottom", y=1.02),
    )
    return figure


def build_productivity_rate_chart(summary):
    """Create a ranking chart with productivity percentage by server."""
    go = _load_plotly()
    if go is None:
        return None
    if summary.empty:
        return _empty_figure(go, "Sem dados para exibir")

    ordered = summary.sort_values("% produtividade", ascending=False)
    figure = go.Figure(
        data=[
            go.Bar(
                x=ordered["Servidor"],
                y=ordered["% produtividade"],
                marker_color="#2A9D8F",
                text=ordered["% produtividade"].map(lambda value: f"{value:.1f}%"),
                textposition="outside",
                hovertemplate="%{x}: %{y:.1f}%<extra></extra>",
            )
        ]
    )
    figure.update_layout(
        height=360,
        margin=dict(l=24, r=24, t=30, b=24),
        template="plotly_white",
        xaxis_title=None,
        yaxis_title="% produtividade",
        uniformtext_minsize=10,
        uniformtext_mode="show",
    )
    figure.update_yaxes(rangemode="tozero", ticksuffix="%")
    return figure


def build_pie_chart(counts: dict[str, int]):
    """Create a donut chart showing the global comparison distribution."""
    go = _load_plotly()
    if go is None:
        return None

    labels = list(counts.keys())
    values = [counts[label] for label in labels]
    colors = [STATUS_COLORS.get(label, "#6B7280") for label in labels]

    if not any(values):
        return _empty_figure(go, "Sem dados para exibir")

    figure = go.Figure(
        data=[
            go.Pie(
                labels=labels,
                values=values,
                hole=0.48,
                marker=dict(colors=colors),
                textinfo="label+percent",
                hovertemplate="%{label}: %{value}<extra></extra>",
            )
        ]
    )
    figure.update_layout(
        height=360,
        margin=dict(l=24, r=24, t=30, b=24),
        template="plotly_white",
        showlegend=False,
    )
    return figure


def _empty_figure(go, message: str):
    figure = go.Figure()
    figure.add_annotation(
        text=message,
        x=0.5,
        y=0.5,
        xref="paper",
        yref="paper",
        showarrow=False,
        font=dict(size=16, color="#6B7280"),
    )
    figure.update_layout(
        height=360,
        margin=dict(l=24, r=24, t=30, b=24),
        template="plotly_white",
        xaxis=dict(visible=False),
        yaxis=dict(visible=False),
    )
    return figure


def _load_plotly():
    try:
        import plotly.graph_objects as go
    except ImportError:
        return None
    return go


def _local_xlsx_files() -> list[Path]:
    return sorted(
        path
        for path in Path.cwd().glob("*.xlsx")
        if path.is_file() and not path.name.startswith("~$")
    )


def _guess_file_index(paths: list[Path], terms: tuple[str, ...]) -> int:
    lowered_terms = tuple(term.casefold() for term in terms)
    for index, path in enumerate(paths):
        name = path.name.casefold()
        if all(term in name for term in lowered_terms):
            return index
    return 0


def _local_workbook_input(path: Path) -> WorkbookInput:
    return WorkbookInput(
        name=path.name,
        bytes_data=cached_local_file(str(path), path.stat().st_mtime_ns),
        source_label="pasta do projeto",
    )


def _build_signature(
    server_bytes: bytes,
    current_bytes: bytes,
    current_sheet_name: str,
) -> str:
    payload = "|".join(
        [
            _hash_bytes(server_bytes),
            _hash_bytes(current_bytes),
            current_sheet_name,
        ]
    )
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


def _hash_bytes(file_bytes: bytes) -> str:
    return hashlib.sha256(file_bytes).hexdigest()


if __name__ == "__main__":
    main()
