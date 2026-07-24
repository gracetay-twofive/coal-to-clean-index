from __future__ import annotations

import base64
import io
import os
from datetime import datetime
from pathlib import Path
from typing import Any

import matplotlib.pyplot as plt
import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen.canvas import Canvas
from reportlab.platypus import Paragraph, Table, TableStyle

PILLAR_COLUMNS = [
    "PILLAR 1 score\nEnergy system conditions",
    "PILLAR 2 score\nPolicy and transition commitment",
    "PILLAR 3 score\nGovernance and institutional capacity",
    "PILLAR 4 score\nCarbon market maturity",
    "PILLAR 5 score\nMacro-financial conditions",
    "PILLAR 6 score\nJust transition and social credibility",
]

PILLAR_HEADERS = [
    "Energy system conditions",
    "Policy and transition commitment",
    "Governance and institutional capacity",
    "Carbon market maturity",
    "Macro-financial conditions",
    "Just transition and social credibility",
]

SURVEY_LABELS = {
    "reliable_power": "Reliable power during the transition",
    "government_follow_through": "Government follow-through",
    "project_delivery": "Ability to deliver the project",
    "carbon_market_track_record": "Carbon-market track record",
    "stable_finance": "Stable finance and payments",
    "workers_communities": "Protection for workers and communities",
    "time_to_delivery": "Time to delivery",
    "credit_price": "Credit price",
}

ACCENT = "#2b5688"
ADJUSTED = "#8bb0da"
PALE = "#ecf2f9"
HEADER = "#e8f0f8"
INK = "#1a1a1a"
MUTED = "#5f6770"


def _register_fonts() -> tuple[str, str, str]:
    candidates = [
        (Path.home() / "Library/Fonts/Montserrat-Regular.ttf", Path.home() / "Library/Fonts/Montserrat-SemiBold.ttf"),
        (Path("/Library/Fonts/Montserrat-Regular.ttf"), Path("/Library/Fonts/Montserrat-SemiBold.ttf")),
        (Path("/usr/share/fonts/truetype/montserrat/Montserrat-Regular.ttf"), Path("/usr/share/fonts/truetype/montserrat/Montserrat-SemiBold.ttf")),
    ]
    for regular, bold in candidates:
        if regular.exists() and bold.exists():
            try:
                pdfmetrics.registerFont(TTFont("Montserrat", str(regular)))
                pdfmetrics.registerFont(TTFont("Montserrat-SemiBold", str(bold)))
                return "Montserrat", "Montserrat-SemiBold", "Montserrat"
            except Exception:
                pass
    return "Helvetica", "Helvetica-Bold", "DejaVu Sans"


FONT_REGULAR, FONT_BOLD, MPL_FONT = _register_fonts()


def _ranking_chart_png(results: pd.DataFrame, colour: str, overlay: bool, wide: bool = False) -> bytes:
    ordered = results.sort_values("Rank").sort_values("Rank", ascending=False)
    count = len(ordered)
    fig_height = max(9.4, count * 0.24)
    fig_width = 11.0 if wide else 7.35
    fig, ax = plt.subplots(figsize=(fig_width, fig_height))
    fig.patch.set_alpha(0)
    ax.set_facecolor("white")

    bars = ax.barh(ordered["Country"], ordered["Overall score"], color=colour, height=0.68)
    maximum = max(100.0, float(ordered["Overall score"].max()) + 6)
    ax.set_xlim(0, maximum)
    ax.set_xlabel("OVERALL SCORE", fontsize=8.5, fontfamily=MPL_FONT, labelpad=9)
    ax.grid(axis="x", alpha=0.14, linewidth=0.6)
    ax.set_axisbelow(True)
    ax.tick_params(axis="x", labelsize=7, length=0)
    ax.tick_params(axis="y", labelsize=7, length=0, pad=8)
    for label in ax.get_yticklabels() + ax.get_xticklabels():
        label.set_fontfamily(MPL_FONT)

    for bar, value in zip(bars, ordered["Overall score"]):
        if overlay:
            ax.text(
                1.4,
                bar.get_y() + bar.get_height() / 2,
                f"{float(value):.2f}",
                va="center",
                ha="left",
                fontsize=6.5,
                color=ACCENT,
                fontfamily=MPL_FONT,
            )
        else:
            ax.text(
                float(value) + 0.45,
                bar.get_y() + bar.get_height() / 2,
                f"{float(value):.2f}",
                va="center",
                ha="left",
                fontsize=6.5,
                color=INK,
                fontfamily=MPL_FONT,
            )

    if overlay and "Base score" in ordered.columns:
        ax.scatter(
            ordered["Base score"],
            range(count),
            marker="D",
            s=22,
            color=ACCENT,
            edgecolors="white",
            linewidths=0.6,
            zorder=4,
            label="Base index",
        )
        from matplotlib.patches import Patch
        handles = [
            Patch(facecolor=colour, edgecolor="none", label="Adjusted score"),
            plt.Line2D([0], [0], marker="D", color="none", markerfacecolor=ACCENT, markeredgecolor="white", markersize=5.5, label="Base index"),
        ]
        legend = ax.legend(
            handles=handles,
            loc="upper right",
            frameon=False,
            fontsize=7,
            ncol=1,
            bbox_to_anchor=(1.0, 1.015),
        )
        for item in legend.get_texts():
            item.set_fontfamily(MPL_FONT)

    for spine in ["top", "right", "left"]:
        ax.spines[spine].set_visible(False)
    ax.spines["bottom"].set_color("#cbd7e5")
    ax.spines["bottom"].set_linewidth(0.6)
    fig.tight_layout(pad=0.8)
    buffer = io.BytesIO()
    fig.savefig(buffer, format="png", dpi=190, bbox_inches="tight", transparent=True)
    plt.close(fig)
    return buffer.getvalue()


def _draw_page_background(canvas: Canvas, width: float, height: float) -> None:
    canvas.setFillColor(colors.HexColor(PALE))
    canvas.rect(0, 0, width, height, stroke=0, fill=1)


def _draw_card(canvas: Canvas, x: float, y: float, width: float, height: float) -> None:
    canvas.setFillColor(colors.white)
    canvas.setStrokeColor(colors.HexColor("#d7e1ec"))
    canvas.setLineWidth(0.6)
    canvas.roundRect(x, y, width, height, 9, stroke=1, fill=1)


def _draw_footer(canvas: Canvas, width: float, page_number: int) -> None:
    footer_font_size = 4.6
    note_style = ParagraphStyle(
        "FooterNote",
        fontName=FONT_REGULAR,
        fontSize=footer_font_size,
        leading=6.0,
        textColor=colors.HexColor(MUTED),
    )
    note = Paragraph(
        "<b>IMPORTANT NOTE:</b> THIS REPORT IS A JURISDICTION-LEVEL SCREENING OUTPUT. IT DOES NOT ASSESS INDIVIDUAL POWER PLANTS, PROJECTS, TRANSACTIONS OR INVESTMENTS, AND IT IS NOT A SUBSTITUTE FOR TRANSACTION-SPECIFIC DUE DILIGENCE.",
        note_style,
    )
    note_width = width - 24 * mm
    _, note_height = note.wrap(note_width, 12 * mm)
    note.drawOn(canvas, 12 * mm, 15.0 * mm)

    canvas.setStrokeColor(colors.HexColor("#d7e1ec"))
    canvas.setLineWidth(0.45)
    canvas.line(12 * mm, 11.8 * mm, width - 12 * mm, 11.8 * mm)
    canvas.setFillColor(colors.HexColor(MUTED))
    canvas.setFont(FONT_REGULAR, footer_font_size)
    generated = datetime.now().strftime("%d %B %Y").upper()
    footer_text = (
        "COAL-TO-CLEAN JURISDICTIONAL READINESS INDEX 2026  |  "
        "DEVELOPED BY GRACE TAY, EMSC SUSTAINABILITY MANAGEMENT  |  "
        "CONTACT: K2521144H@E.NTU.EDU.SG  |  "
        f"GENERATED {generated}"
    )
    canvas.drawString(12 * mm, 7.6 * mm, footer_text)
    canvas.drawRightString(width - 12 * mm, 4.8 * mm, f"PAGE {page_number}")


def _draw_title(canvas: Canvas, title: str, subtitle: str | None, x: float, y: float, max_width: float) -> float:
    canvas.setFillColor(colors.HexColor(INK))
    canvas.setFont(FONT_BOLD, 18)
    canvas.drawString(x, y, title)
    current_y = y - 7 * mm
    if subtitle:
        style = ParagraphStyle(
            "Subtitle",
            fontName=FONT_REGULAR,
            fontSize=8.5,
            leading=11,
            textColor=colors.HexColor(MUTED),
        )
        paragraph = Paragraph(subtitle, style)
        _, height = paragraph.wrap(max_width, 20 * mm)
        paragraph.drawOn(canvas, x, current_y - height + 2 * mm)
        current_y -= height + 2 * mm
    return current_y


def _draw_main_index_page(canvas: Canvas, results: pd.DataFrame, page_number: int) -> None:
    width, height = A4
    canvas.setPageSize(A4)
    _draw_page_background(canvas, width, height)
    margin = 12 * mm
    title_y = height - 18 * mm
    current_y = _draw_title(
        canvas,
        "Coal-to-Clean Jurisdictional Readiness Index 2026",
        "The base index primarily reflects 2024 data and provides the reference point for the customised views that follow.",
        margin,
        title_y,
        width - 2 * margin,
    )
    card_y = 31 * mm
    card_h = current_y - card_y - 4 * mm
    _draw_card(canvas, margin, card_y, width - 2 * margin, card_h)
    chart = _ranking_chart_png(results, ACCENT, overlay=False)
    canvas.drawImage(
        ImageReader(io.BytesIO(chart)),
        margin + 6 * mm,
        card_y + 5 * mm,
        width=width - 2 * margin - 12 * mm,
        height=card_h - 10 * mm,
        preserveAspectRatio=True,
        anchor="c",
        mask="auto",
    )
    _draw_footer(canvas, width, page_number)


def _table_header_paragraph(text: str) -> Paragraph:
    return Paragraph(
        text,
        ParagraphStyle(
            "TableHeader",
            fontName=FONT_BOLD,
            fontSize=6.3,
            leading=7.2,
            textColor=colors.HexColor(INK),
            alignment=0,
        ),
    )


def _draw_full_table_page(canvas: Canvas, results: pd.DataFrame, page_number: int) -> None:
    width, height = landscape(A4)
    canvas.setPageSize((width, height))
    _draw_page_background(canvas, width, height)
    margin = 10 * mm
    canvas.setFillColor(colors.HexColor(INK))
    canvas.setFont(FONT_BOLD, 16)
    canvas.drawString(margin, height - 15 * mm, "Jurisdiction Rankings and Pillar Scores")

    frame = results.sort_values("Rank").copy()
    headers = ["Rank", "Country", "Overall score", *PILLAR_HEADERS]
    data: list[list[Any]] = [[_table_header_paragraph(value) for value in headers]]
    for _, row in frame.iterrows():
        data.append(
            [
                int(row["Rank"]),
                str(row["Country"]),
                f"{float(row['Overall score']):.2f}",
                *[f"{float(row[column]):.2f}" for column in PILLAR_COLUMNS],
            ]
        )

    available_w = width - 2 * margin
    col_widths = [12 * mm, 40 * mm, 22 * mm] + [available_w - 74 * mm] * 6
    pillar_width = (available_w - 74 * mm) / 6
    col_widths = [12 * mm, 40 * mm, 22 * mm] + [pillar_width] * 6
    row_heights = [18 * mm] + [4.15 * mm] * len(frame)
    table = Table(data, colWidths=col_widths, rowHeights=row_heights, repeatRows=1)
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(HEADER)),
                ("TEXTCOLOR", (0, 0), (-1, -1), colors.HexColor(INK)),
                ("FONTNAME", (0, 1), (-1, -1), FONT_REGULAR),
                ("FONTSIZE", (0, 1), (-1, -1), 5.8),
                ("ALIGN", (0, 1), (0, -1), "CENTER"),
                ("ALIGN", (2, 1), (-1, -1), "RIGHT"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#d9e3ef")),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f8fbfe")]),
                ("LEFTPADDING", (0, 0), (-1, -1), 3),
                ("RIGHTPADDING", (0, 0), (-1, -1), 3),
                ("TOPPADDING", (0, 0), (-1, -1), 1.5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 1.5),
            ]
        )
    )
    table_width, table_height = table.wrap(available_w, height)
    card_y = 31 * mm
    card_h = height - 51 * mm
    _draw_card(canvas, margin, card_y, available_w, card_h)
    table.drawOn(canvas, margin, card_y + card_h - table_height)
    _draw_footer(canvas, width, page_number)


def _survey_table(version: dict[str, Any], available_width: float) -> Table:
    header_style = ParagraphStyle(
        "SurveyHeader",
        fontName=FONT_BOLD,
        fontSize=7.4,
        leading=9,
        textColor=colors.HexColor(INK),
    )
    cell_style = ParagraphStyle(
        "SurveyCell",
        fontName=FONT_REGULAR,
        fontSize=7.2,
        leading=9,
        textColor=colors.HexColor(INK),
    )
    data: list[list[Any]] = [[
        Paragraph("Factors affecting coal-to-clean opportunities", header_style),
        Paragraph("Importance", header_style),
    ]]
    for key, value in version.get("responses", {}).items():
        data.append([
            Paragraph(SURVEY_LABELS.get(key, key.replace("_", " ").title()), cell_style),
            Paragraph(f"{int(value)} / 5", cell_style),
        ])
    table = Table(data, colWidths=[available_width - 34 * mm, 34 * mm], repeatRows=1)
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(HEADER)),
                ("TEXTCOLOR", (0, 0), (-1, -1), colors.HexColor(INK)),
                ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#d9e3ef")),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f8fbfe")]),
                ("ALIGN", (1, 1), (1, -1), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
                ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ]
        )
    )
    return table


def _draw_custom_page(canvas: Canvas, version: dict[str, Any], page_number: int) -> None:
    width, height = A4
    canvas.setPageSize(A4)
    _draw_page_background(canvas, width, height)
    margin = 12 * mm
    title_y = height - 18 * mm
    current_y = _draw_title(
        canvas,
        "Your Priority-Adjusted Country Ranking",
        "The adjusted ranking reflects your industry and survey priorities.",
        margin,
        title_y,
        width - 2 * margin,
    )
    canvas.setFillColor(colors.HexColor(MUTED))
    canvas.setFont(FONT_BOLD, 7.5)
    canvas.drawString(margin, current_y - 1 * mm, f"CUSTOMISED VIEW {version['number']}")
    current_y -= 8 * mm

    card_x = margin
    card_w = width - 2 * margin
    chart_h = 156 * mm
    chart_y = current_y - chart_h
    _draw_card(canvas, card_x, chart_y, card_w, chart_h)
    chart = _ranking_chart_png(version["results"], ADJUSTED, overlay=True, wide=True)
    canvas.drawImage(
        ImageReader(io.BytesIO(chart)),
        card_x + 5 * mm,
        chart_y + 4 * mm,
        width=card_w - 10 * mm,
        height=chart_h - 8 * mm,
        preserveAspectRatio=True,
        anchor="c",
        mask="auto",
    )

    table_top = chart_y - 5 * mm
    table = _survey_table(version, card_w)
    table_width, table_height = table.wrap(card_w, 80 * mm)
    table_y = table_top - table_height
    table.drawOn(canvas, card_x, table_y)
    canvas.setFillColor(colors.HexColor(MUTED))
    canvas.setFont(FONT_REGULAR, 6.7)
    canvas.drawString(card_x, table_y - 5 * mm, "1 = Less important     5 = More important")
    _draw_footer(canvas, width, page_number)


def build_version_csv(version: dict[str, Any]) -> bytes:
    frame = version["results"].copy()
    return frame.to_csv(index=False).encode("utf-8")


def build_versions_pdf(base_results: pd.DataFrame, versions: list[dict[str, Any]]) -> bytes:
    buffer = io.BytesIO()
    canvas = Canvas(buffer, pagesize=A4)
    page_number = 1
    _draw_main_index_page(canvas, base_results, page_number)
    canvas.showPage()

    page_number += 1
    _draw_full_table_page(canvas, base_results, page_number)
    canvas.showPage()

    for version in versions:
        page_number += 1
        _draw_custom_page(canvas, version, page_number)
        canvas.showPage()

    canvas.save()
    return buffer.getvalue()


def encode_attachment(name: str, mime_type: str, content: bytes) -> dict[str, str]:
    return {
        "name": name,
        "mime_type": mime_type,
        "content_base64": base64.b64encode(content).decode("ascii"),
    }
