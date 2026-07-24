"""Plotly charts used by the dashboard."""

from __future__ import annotations

import plotly.graph_objects as go


STATUS_COLORS = {
    "Excluídos": "#D95F59",
    "Novos": "#2A9D8F",
    "Mantidos": "#3A6EA5",
}


def build_bar_chart(counts: dict[str, int]) -> go.Figure:
    """Create a bar chart comparing excluded, new and maintained processes."""
    labels = list(counts.keys())
    values = [counts[label] for label in labels]
    colors = [STATUS_COLORS.get(label, "#6B7280") for label in labels]

    figure = go.Figure(
        data=[
            go.Bar(
                x=labels,
                y=values,
                marker_color=colors,
                text=values,
                textposition="outside",
                hovertemplate="%{x}: %{y}<extra></extra>",
            )
        ]
    )
    figure.update_layout(
        height=360,
        margin=dict(l=24, r=24, t=30, b=24),
        template="plotly_white",
        xaxis_title=None,
        yaxis_title="Quantidade",
        uniformtext_minsize=10,
        uniformtext_mode="show",
    )
    figure.update_yaxes(rangemode="tozero")
    return figure


def build_pie_chart(counts: dict[str, int]) -> go.Figure:
    """Create a donut chart showing process distribution."""
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
