from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from skills.clean_data import ALL_METRICS, prepare_analysis_data
from skills.compute_metrics import compute_metrics

try:
    from jinja2 import Environment, FileSystemLoader
except ImportError:  # pragma: no cover - fallback renderer handles this case
    Environment = None
    FileSystemLoader = None


DEFAULT_TEMPLATE_PATH = Path("templates/report_template.html")
DEFAULT_OUTPUT_PATH = Path("output/monthly_review.html")
PERCENT_METRICS = {"ctr", "cvr"}


def format_metric_value(metric_name: str, value: Any) -> str:
    if value is None or pd.isna(value):
        return "N/A"
    if metric_name in PERCENT_METRICS:
        return f"{float(value):.2%}"
    if metric_name == "roi":
        return f"{float(value):.2f}"
    if metric_name in {"spend", "cpc", "cpm", "gmv", "aov"}:
        return f"{float(value):,.2f}"
    if float(value).is_integer():
        return f"{int(value):,}"
    return f"{float(value):,.2f}"


def _normalize_table_input(table: Any) -> pd.DataFrame:
    if isinstance(table, pd.DataFrame):
        return table.copy()
    if isinstance(table, list):
        return pd.DataFrame(table)
    if isinstance(table, dict):
        if "records" in table:
            return pd.DataFrame(table["records"])
        return pd.DataFrame([table])
    return pd.DataFrame()


def format_table_for_display(table: Any) -> dict[str, Any]:
    df = _normalize_table_input(table)
    if df.empty:
        return {"columns": [], "rows": []}

    display = df.copy()
    for column in display.columns:
        column_key = str(column).lower()
        if column_key in ALL_METRICS:
            display[column] = display[column].apply(
                lambda value, metric=column_key: format_metric_value(metric, value)
            )
        else:
            display[column] = display[column].fillna("N/A").astype(str)

    return {
        "columns": [str(column) for column in display.columns],
        "rows": display.to_dict(orient="records"),
    }


def build_kpi_summary_section(data: pd.DataFrame | None = None) -> dict[str, Any]:
    if data is None or data.empty:
        empty_row = {metric: "N/A" for metric in ALL_METRICS}
        return {"title": "KPI 概览", "columns": list(empty_row.keys()), "rows": [empty_row]}

    prepared = prepare_analysis_data(data)
    summary_df = compute_metrics(prepared)
    formatted = format_table_for_display(summary_df)
    formatted["title"] = "KPI 概览"
    return formatted


def normalize_chart_items(chart_paths: list[Any] | None) -> list[dict[str, str]]:
    charts: list[dict[str, str]] = []
    for index, item in enumerate(chart_paths or [], start=1):
        if isinstance(item, str):
            charts.append(
                {
                    "title": f"图表 {index}",
                    "path": item,
                    "caption": "",
                }
            )
            continue

        if isinstance(item, dict):
            path = str(item.get("path", "")).strip()
            if not path:
                continue
            charts.append(
                {
                    "title": str(item.get("title", f"图表 {index}")),
                    "path": path,
                    "caption": str(item.get("caption", "")),
                }
            )
    return charts


def normalize_diagnosis_sections(
    diagnosis_tables: list[dict[str, Any]] | None,
) -> list[dict[str, Any]]:
    sections: list[dict[str, Any]] = []
    for index, item in enumerate(diagnosis_tables or [], start=1):
        title = str(item.get("title", f"诊断表 {index}"))
        summary = str(item.get("summary", "")).strip()
        next_step = str(item.get("next_step", "")).strip()
        note = str(item.get("note", "")).strip()
        formatted_table = format_table_for_display(item.get("table"))

        sections.append(
            {
                "title": title,
                "summary": summary,
                "next_step": next_step,
                "note": note,
                "columns": formatted_table["columns"],
                "rows": formatted_table["rows"],
            }
        )
    return sections


def build_report_context(report_payload: dict[str, Any]) -> dict[str, Any]:
    input_data = report_payload.get("kpi_source_data")
    source_df = _normalize_table_input(input_data) if input_data is not None else None

    kpi_summary = report_payload.get("kpi_summary")
    if kpi_summary is None:
        kpi_summary_section = build_kpi_summary_section(source_df)
    else:
        kpi_summary_section = format_table_for_display(kpi_summary)
        kpi_summary_section["title"] = "KPI 概览"

    context = {
        "title": str(report_payload.get("title", "Skintific Indonesia Meta x Shopee 月报")),
        "subtitle": str(report_payload.get("subtitle", "业务复盘报告")),
        "executive_summary": list(report_payload.get("executive_summary", [])),
        "kpi_summary": kpi_summary_section,
        "charts": normalize_chart_items(report_payload.get("chart_paths")),
        "diagnosis_sections": normalize_diagnosis_sections(
            report_payload.get("diagnosis_tables")
        ),
        "highlights": list(report_payload.get("highlights", [])),
        "weaknesses": list(report_payload.get("weaknesses", [])),
        "actions": list(report_payload.get("actions", [])),
        "data_notes": list(report_payload.get("data_notes", [])),
    }
    return context


def _render_table_html(columns: list[str], rows: list[dict[str, Any]]) -> str:
    if not columns or not rows:
        return "<p class='empty-note'>当前没有可展示的表格数据。</p>"

    header_html = "".join(f"<th>{column}</th>" for column in columns)
    row_html = []
    for row in rows:
        cells = "".join(f"<td>{row.get(column, 'N/A')}</td>" for column in columns)
        row_html.append(f"<tr>{cells}</tr>")

    return (
        "<div class='table-wrap'><table><thead><tr>"
        f"{header_html}"
        "</tr></thead><tbody>"
        f"{''.join(row_html)}"
        "</tbody></table></div>"
    )


def _render_without_jinja(context: dict[str, Any]) -> str:
    chart_html = []
    for chart in context["charts"]:
        caption = f"<p class='chart-caption'>{chart['caption']}</p>" if chart["caption"] else ""
        chart_html.append(
            "<section class='card'>"
            f"<div class='section-tag'>趋势图</div><h2>{chart['title']}</h2>"
            f"<img src='{chart['path']}' alt='{chart['title']}'>"
            f"{caption}</section>"
        )

    diagnosis_html = []
    for section in context["diagnosis_sections"]:
        note_html = f"<p>{section['note']}</p>" if section["note"] else ""
        summary_html = f"<p>{section['summary']}</p>" if section["summary"] else ""
        next_step_html = (
            f"<p><strong>下一步：</strong>{section['next_step']}</p>"
            if section["next_step"]
            else ""
        )
        diagnosis_html.append(
            "<section class='card'>"
            f"<div class='section-tag'>诊断表</div><h2>{section['title']}</h2>"
            f"{note_html}{summary_html}"
            f"{_render_table_html(section['columns'], section['rows'])}"
            f"{next_step_html}</section>"
        )

    executive_summary_html = "".join(f"<li>{item}</li>" for item in context["executive_summary"])
    highlights_html = "".join(f"<li>{item}</li>" for item in context["highlights"])
    weaknesses_html = "".join(f"<li>{item}</li>" for item in context["weaknesses"])
    actions_html = "".join(f"<li>{item}</li>" for item in context["actions"])
    data_notes_html = "".join(f"<li>{item}</li>" for item in context["data_notes"])

    return f"""<!DOCTYPE html>
<html lang="zh">
<head>
  <meta charset="utf-8">
  <title>{context['title']}</title>
  <style>
    body {{ margin: 0; font-family: "Microsoft YaHei", "PingFang SC", sans-serif; background: linear-gradient(180deg, #f3eee1 0%, #f8f7f4 100%); color: #182026; }}
    .page {{ max-width: 1240px; margin: 0 auto; padding: 28px 18px 40px; }}
    .hero, .card {{ background: #fffdfa; border: 1px solid #ddd4c6; border-radius: 18px; box-shadow: 0 14px 32px rgba(24,32,38,0.06); }}
    .hero {{ padding: 28px; margin-bottom: 20px; background: radial-gradient(circle at top right, rgba(15,118,110,0.12), transparent 28%), #fffdfa; }}
    .card {{ padding: 22px; margin-bottom: 18px; }}
    .grid {{ display: grid; grid-template-columns: 1fr; gap: 18px; }}
    .two-col {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(280px, 1fr)); gap: 18px; }}
    h1, h2 {{ margin: 0 0 12px; }}
    p, li {{ color: #52606d; line-height: 1.7; }}
    ul {{ margin: 0; padding-left: 18px; }}
    .section-tag {{ display: inline-block; margin-bottom: 12px; padding: 4px 10px; border-radius: 999px; background: #dff3f1; color: #0f766e; font-size: 13px; }}
    .table-wrap {{ overflow-x: auto; margin-top: 12px; }}
    table {{ width: 100%; min-width: 840px; border-collapse: collapse; font-size: 13px; }}
    th, td {{ padding: 10px 8px; text-align: left; border-bottom: 1px solid #ddd4c6; vertical-align: top; }}
    th {{ background: #f1f5f4; color: #182026; }}
    img {{ max-width: 100%; border-radius: 14px; border: 1px solid #ddd4c6; }}
    .chart-caption {{ margin-top: 10px; }}
    .empty-note {{ color: #8a5a00; }}
  </style>
</head>
<body>
  <div class="page">
    <section class="hero">
      <div class="section-tag">Executive Summary</div>
      <h1>{context['title']}</h1>
      <p>{context['subtitle']}</p>
      <ul>{executive_summary_html}</ul>
    </section>

    <div class="grid">
      <section class="card">
        <div class="section-tag">KPI Summary</div>
        <h2>{context['kpi_summary']['title']}</h2>
        {_render_table_html(context['kpi_summary']['columns'], context['kpi_summary']['rows'])}
      </section>

      {''.join(chart_html)}
      {''.join(diagnosis_html)}

      <section class="two-col">
        <div class="card">
          <div class="section-tag">Highlights</div>
          <h2>亮点</h2>
          <ul>{highlights_html}</ul>
        </div>
        <div class="card">
          <div class="section-tag">Weaknesses</div>
          <h2>弱点</h2>
          <ul>{weaknesses_html}</ul>
        </div>
      </section>

      <section class="card">
        <div class="section-tag">Next Steps</div>
        <h2>下一步动作</h2>
        <ul>{actions_html}</ul>
      </section>

      <section class="card">
        <div class="section-tag">Data Notes</div>
        <h2>数据说明</h2>
        <ul>{data_notes_html}</ul>
      </section>
    </div>
  </div>
</body>
</html>"""


def render_html_report(
    context: dict[str, Any],
    template_path: str | Path = DEFAULT_TEMPLATE_PATH,
) -> str:
    template_path = Path(template_path)
    if Environment is None or FileSystemLoader is None or not template_path.exists():
        return _render_without_jinja(context)

    environment = Environment(loader=FileSystemLoader(str(template_path.parent)))
    template = environment.get_template(template_path.name)
    return template.render(**context)


def save_html_report(
    html: str,
    output_path: str | Path = DEFAULT_OUTPUT_PATH,
) -> Path:
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(html, encoding="utf-8")
    return output_path


def build_html_report(
    report_payload: dict[str, Any],
    template_path: str | Path = DEFAULT_TEMPLATE_PATH,
    output_path: str | Path = DEFAULT_OUTPUT_PATH,
) -> Path:
    context = build_report_context(report_payload)
    html = render_html_report(context=context, template_path=template_path)
    return save_html_report(html=html, output_path=output_path)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Render a Chinese business HTML report.")
    parser.add_argument(
        "--report-json",
        required=True,
        help="JSON payload path for the report content.",
    )
    parser.add_argument(
        "--template",
        default=str(DEFAULT_TEMPLATE_PATH),
        help="HTML template path.",
    )
    parser.add_argument(
        "--output-html",
        default=str(DEFAULT_OUTPUT_PATH),
        help="Final HTML output path under /output.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    payload = json.loads(Path(args.report_json).read_text(encoding="utf-8"))
    build_html_report(
        report_payload=payload,
        template_path=args.template,
        output_path=args.output_html,
    )


if __name__ == "__main__":
    main()
