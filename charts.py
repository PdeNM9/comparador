"""Plotly charts used by the productivity dashboard."""

from __future__ import annotations

import plotly.graph_objects as go


STATUS_COLORS = {
    "Produtivos": "#2A9D8F",
    "Permaneceram": "#3A6EA5",
    "Novos": "#D9A441",
}


def build_productivity_by_server_chart(summary) -> go.Figure:
    """Create a stacked bar chart with productive and remaining processes."""
    if summary.empty:
        return _empty_figure("Sem dados para exibir")

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


def build_productivity_rate_chart(summary) -> go.Figure:
    """Create a ranking chart with productivity percentage by server."""
    if summary.empty:
        return _empty_figure("Sem dados para exibir")

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


def build_pie_chart(counts: dict[str, int]) -> go.Figure:
    """Create a donut chart showing the global comparison distribution."""
    labels = list(counts.keys())
    values = [counts[label] for label in labels]
    colors = [STATUS_COLORS.get(label, "#6B7280") for label in labels]

    if not any(values):
        return _empty_figure("Sem dados para exibir")

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


def _empty_figure(message: str) -> go.Figure:
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
