"""Streamlit app for comparing judicial process Excel spreadsheets."""

from __future__ import annotations

import hashlib
from typing import Any

import pandas as pd
import streamlit as st

from charts import build_bar_chart, build_pie_chart
from comparison import ComparisonError, ComparisonResult, SheetConfig, compare_sheets
from dashboard import (
    configure_page,
    guess_annotation_columns,
    guess_column,
    inject_styles,
    render_filterable_table,
    render_header,
    render_metric_cards,
    render_quality_messages,
)
from excel_utils import ExcelProcessingError, read_excel_sheet, read_sheet_names
from export import EXCEL_MIME_TYPE, dataframe_to_excel_bytes, timestamped_filename


OLD_FILE_KEY = "old_file"
NEW_FILE_KEY = "new_file"
RESULT_KEY = "comparison_result"
SIGNATURE_KEY = "comparison_signature"


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
    """Cache worksheet loading for the selected sheet."""
    return read_excel_sheet(file_bytes, file_name, sheet_name)


@st.cache_data(show_spinner=False)
def cached_excel_bytes(dataframe: pd.DataFrame, sheet_name: str) -> bytes:
    """Cache Excel serialization to keep download buttons responsive."""
    return dataframe_to_excel_bytes(dataframe, sheet_name)


def main() -> None:
    """Run the Streamlit application."""
    configure_page()
    inject_styles()
    render_header()

    old_upload, new_upload = _render_upload_area()
    if not old_upload or not new_upload:
        st.info("Envie as duas planilhas .xlsx para liberar a configuração.")
        return

    try:
        old_bytes = old_upload.getvalue()
        new_bytes = new_upload.getvalue()
        old_sheet_names = cached_sheet_names(old_bytes, old_upload.name)
        new_sheet_names = cached_sheet_names(new_bytes, new_upload.name)
    except ExcelProcessingError as exc:
        st.error(str(exc))
        return

    try:
        selection = _render_sidebar_configuration(
            old_bytes,
            old_upload.name,
            old_sheet_names,
            new_bytes,
            new_upload.name,
            new_sheet_names,
        )
    except ExcelProcessingError as exc:
        st.error(str(exc))
        return

    if selection is None:
        return

    (
        old_dataframe,
        new_dataframe,
        old_sheet_name,
        new_sheet_name,
        old_config,
        new_config,
    ) = selection

    current_signature = _build_signature(
        old_bytes,
        new_bytes,
        old_sheet_name,
        new_sheet_name,
        old_config,
        new_config,
    )

    if st.sidebar.button(
        "Comparar planilhas",
        type="primary",
        use_container_width=True,
    ):
        _run_comparison(
            old_dataframe,
            new_dataframe,
            old_config,
            new_config,
            current_signature,
        )

    result = st.session_state.get(RESULT_KEY)
    if result is None:
        st.info("Configure as colunas e clique em Comparar planilhas.")
        return

    if st.session_state.get(SIGNATURE_KEY) != current_signature:
        st.warning(
            "A configuração ou algum arquivo foi alterado depois da última "
            "comparação. Clique em Comparar planilhas para atualizar o painel."
        )

    _render_results(result)


def _render_upload_area():
    st.subheader("Upload")
    left, right = st.columns(2)
    with left:
        old_upload = st.file_uploader(
            "Planilha 1 - situação anterior",
            type=["xlsx"],
            key=OLD_FILE_KEY,
            help="Arquivo Excel .xlsx com a situação anterior.",
        )
    with right:
        new_upload = st.file_uploader(
            "Planilha 2 - situação atual",
            type=["xlsx"],
            key=NEW_FILE_KEY,
            help="Arquivo Excel .xlsx com a situação atual.",
        )
    return old_upload, new_upload


def _render_sidebar_configuration(
    old_bytes: bytes,
    old_file_name: str,
    old_sheet_names: list[str],
    new_bytes: bytes,
    new_file_name: str,
    new_sheet_names: list[str],
):
    with st.sidebar:
        st.header("Configuração")
        old_sheet_name = _safe_selectbox(
            "Aba da Planilha 1",
            old_sheet_names,
            "old_sheet_name",
            old_sheet_names[0],
        )
        new_sheet_name = _safe_selectbox(
            "Aba da Planilha 2",
            new_sheet_names,
            "new_sheet_name",
            new_sheet_names[0],
        )

    old_dataframe = cached_read_sheet(old_bytes, old_file_name, old_sheet_name)
    new_dataframe = cached_read_sheet(new_bytes, new_file_name, new_sheet_name)

    old_columns = list(old_dataframe.columns)
    new_columns = list(new_dataframe.columns)
    if not old_columns:
        st.error("A Planilha 1 não possui colunas para seleção.")
        return None
    if not new_columns:
        st.error("A Planilha 2 não possui colunas para seleção.")
        return None

    with st.sidebar:
        st.caption(
            f"Planilha 1: {len(old_dataframe):,} linhas".replace(",", ".")
        )
        st.caption(
            f"Planilha 2: {len(new_dataframe):,} linhas".replace(",", ".")
        )
        st.divider()
        st.subheader("Planilha 1")
        old_cnj_column = _safe_selectbox(
            "Coluna CNJ",
            old_columns,
            "old_cnj_column",
            guess_column(old_columns, ("cnj", "processo", "numero", "número")),
        )
        old_responsible_column = _safe_selectbox(
            "Coluna responsável",
            old_columns,
            "old_responsible_column",
            guess_column(
                old_columns,
                ("responsavel", "responsável", "advogado", "servidor"),
            ),
        )

        annotation_options = [
            column
            for column in old_columns
            if column not in {old_cnj_column, old_responsible_column}
        ]
        default_annotations = [
            column
            for column in guess_annotation_columns(annotation_options)
            if column in annotation_options
        ]
        old_annotation_columns = _safe_multiselect(
            "Colunas de anotações",
            annotation_options,
            "old_annotation_columns",
            default_annotations,
        )

        if not old_annotation_columns:
            st.warning(
                "Nenhuma coluna de anotação foi selecionada. O arquivo final "
                "será gerado, mas sem transferência de anotações."
            )

        st.divider()
        st.subheader("Planilha 2")
        new_cnj_column = _safe_selectbox(
            "Coluna CNJ",
            new_columns,
            "new_cnj_column",
            guess_column(new_columns, ("cnj", "processo", "numero", "número")),
        )
        new_responsible_column = _safe_selectbox(
            "Coluna responsável",
            new_columns,
            "new_responsible_column",
            guess_column(
                new_columns,
                ("responsavel", "responsável", "advogado", "servidor"),
            ),
        )

    old_config = SheetConfig(
        cnj_column=old_cnj_column,
        responsible_column=old_responsible_column,
        annotation_columns=tuple(old_annotation_columns),
    )
    new_config = SheetConfig(
        cnj_column=new_cnj_column,
        responsible_column=new_responsible_column,
    )
    return (
        old_dataframe,
        new_dataframe,
        old_sheet_name,
        new_sheet_name,
        old_config,
        new_config,
    )


def _safe_selectbox(
    label: str,
    options: list[Any],
    key: str,
    default: Any,
) -> Any:
    if key in st.session_state and st.session_state[key] not in options:
        del st.session_state[key]

    index = options.index(default) if default in options else 0
    return st.selectbox(
        label,
        options=options,
        index=index,
        key=key,
        format_func=str,
    )


def _safe_multiselect(
    label: str,
    options: list[Any],
    key: str,
    default: list[Any],
) -> list[Any]:
    if key in st.session_state:
        st.session_state[key] = [
            value for value in st.session_state[key] if value in options
        ]
        return st.multiselect(
            label,
            options=options,
            key=key,
            format_func=str,
        )

    return st.multiselect(
        label,
        options=options,
        default=default,
        key=key,
        format_func=str,
    )


def _run_comparison(
    old_dataframe: pd.DataFrame,
    new_dataframe: pd.DataFrame,
    old_config: SheetConfig,
    new_config: SheetConfig,
    current_signature: str,
) -> None:
    progress = st.progress(0, text="Validando seleção das colunas")
    try:
        progress.progress(25, text="Normalizando números CNJ")
        progress.progress(50, text="Comparando processos")
        result = compare_sheets(
            old_dataframe,
            new_dataframe,
            old_config,
            new_config,
        )
        progress.progress(85, text="Preparando painel e exportações")
        st.session_state[RESULT_KEY] = result
        st.session_state[SIGNATURE_KEY] = current_signature
        progress.progress(100, text="Comparação concluída")
        st.success("Comparação concluída com sucesso.")
    except (ComparisonError, ExcelProcessingError) as exc:
        st.error(str(exc))
    finally:
        progress.empty()


def _render_results(result: ComparisonResult) -> None:
    render_metric_cards(result)
    render_quality_messages(result)

    st.subheader("Gráficos")
    chart_left, chart_right = st.columns(2)
    with chart_left:
        st.plotly_chart(
            build_bar_chart(result.status_counts),
            use_container_width=True,
            config={"displayModeBar": False},
        )
    with chart_right:
        st.plotly_chart(
            build_pie_chart(result.status_counts),
            use_container_width=True,
            config={"displayModeBar": False},
        )

    st.subheader("Exportações")
    export_left, export_middle, export_right = st.columns(3)
    with export_left:
        st.download_button(
            "Baixar excluídos",
            data=cached_excel_bytes(result.excluded_processes, "Excluídos"),
            file_name=timestamped_filename("processos_excluidos"),
            mime=EXCEL_MIME_TYPE,
            use_container_width=True,
            disabled=result.excluded_processes.empty,
        )
    with export_middle:
        st.download_button(
            "Baixar novos",
            data=cached_excel_bytes(result.new_processes, "Novos"),
            file_name=timestamped_filename("processos_novos"),
            mime=EXCEL_MIME_TYPE,
            use_container_width=True,
            disabled=result.new_processes.empty,
        )
    with export_right:
        st.download_button(
            "Baixar completo com anotações",
            data=cached_excel_bytes(result.final_dataframe, "Planilha final"),
            file_name=timestamped_filename("planilha_com_anotacoes"),
            mime=EXCEL_MIME_TYPE,
            use_container_width=True,
        )

    st.subheader("Tabelas")
    excluded_tab, new_tab, maintained_tab = st.tabs(
        ["Excluídos", "Novos", "Mantidos"]
    )
    with excluded_tab:
        render_filterable_table(
            "Processos excluídos",
            result.excluded_processes,
            "excluded",
        )
    with new_tab:
        render_filterable_table(
            "Processos novos",
            result.new_processes,
            "new",
        )
    with maintained_tab:
        render_filterable_table(
            "Processos mantidos",
            result.maintained_processes,
            "maintained",
        )


def _build_signature(
    old_bytes: bytes,
    new_bytes: bytes,
    old_sheet_name: str,
    new_sheet_name: str,
    old_config: SheetConfig,
    new_config: SheetConfig,
) -> str:
    payload = "|".join(
        [
            _hash_bytes(old_bytes),
            _hash_bytes(new_bytes),
            old_sheet_name,
            new_sheet_name,
            repr(old_config),
            repr(new_config),
        ]
    )
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


def _hash_bytes(file_bytes: bytes) -> str:
    return hashlib.sha256(file_bytes).hexdigest()


if __name__ == "__main__":
    main()
