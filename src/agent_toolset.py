import asyncio
import json
import logging
import os
import random
import re
import uuid
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

import requests
from openai import AsyncOpenAI
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.series import DataPoint
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.properties import PageSetupProperties
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)

# ----------------------------------------------------------------------------
# Excel styling constants
# ----------------------------------------------------------------------------
FONT_NAME = "Arial"
NUMERIC_TYPES = {"integer", "number", "currency", "percent"}

_BASE_PALETTE = ["F28E2B", "59A14F", "E15759", "76B7B2", "EDC948", "B07AA1", "9C755F", "BAB0AC"]
THEMES = {
    "blue":   {"dark": "1F3A5F", "accent": "2E75B6", "band": "EAF1FB", "palette": ["2E75B6"] + _BASE_PALETTE},
    "green":  {"dark": "1E4D3A", "accent": "2E8B57", "band": "E8F5EE", "palette": ["2E8B57"] + _BASE_PALETTE},
    "purple": {"dark": "3B2A5E", "accent": "7B5EA7", "band": "F1ECF8", "palette": ["7B5EA7"] + _BASE_PALETTE},
    "orange": {"dark": "5A2E12", "accent": "E07B39", "band": "FDF0E6", "palette": ["E07B39"] + _BASE_PALETTE},
}

THIN = Side(style="thin", color="D0D7DE")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


class DocumentCreatorToolset:
    def __init__(self):
        self.openai_key = os.getenv("OPENAI_API_KEY")
        self.pexels_key = os.getenv("PEXELS_API_KEY")
        self.client = AsyncOpenAI(api_key=self.openai_key)
        self.model_name = "gpt-4o"

        self.supported_formats = ["pptx_slide_deck", "xlsx_spreadsheet"]
        self.last_brief: Optional[str] = None
        self.last_plan: Optional[Dict[str, Any]] = None

        self.image_cache_dir = "image_cache"
        os.makedirs(self.image_cache_dir, exist_ok=True)

    # ------------------------------------------------------------------
    # TOOLS (exposed to the LLM agent)
    # ------------------------------------------------------------------
    async def generate_document(self, brief: str) -> str:
        """Generate a PowerPoint deck or Excel spreadsheet from the user's brief, choosing the best format automatically."""
        logger.info(f"📄 GENERATE_DOCUMENT CALLED: {brief}")
        try:
            self.last_brief = brief
            fmt = await self._detect_format(brief)

            if fmt == "xlsx_spreadsheet":
                logger.info("📊 Generating Spreadsheet...")
                plan = await self._generate_full_xlsx_plan_from_brief(brief)
                plan["detected_format"] = fmt
                doc_path = await asyncio.to_thread(self._render_xlsx_from_plan, plan)
            else:
                fmt = "pptx_slide_deck"
                logger.info("🖼️ Generating Presentation...")
                plan = await self._generate_full_ppt_plan_from_brief(brief)
                plan["detected_format"] = fmt
                doc_path = await asyncio.to_thread(self._render_ppt_from_plan, plan)

            self.last_plan = plan
            return json.dumps({"status": "success", "file_path": doc_path,
                               "message": f"Successfully generated {fmt}"})
        except Exception as e:
            logger.exception("Error in generate_document")
            return json.dumps({"status": "error", "message": str(e)})

    async def revise_document(self, instruction: str) -> str:
        """Revise the most recently generated document according to the user's change request."""
        logger.info(f"✏️ REVISE_DOCUMENT CALLED: {instruction}")
        if not self.last_plan:
            return json.dumps({"status": "error",
                               "message": "No document has been generated yet. Ask for a new document first."})
        try:
            fmt = self.last_plan.get("detected_format", "pptx_slide_deck")
            prompt = (
                "You are editing a document plan stored as JSON.\n"
                f"CURRENT PLAN:\n{json.dumps(self.last_plan)}\n\n"
                f"USER CHANGE REQUEST: {instruction}\n\n"
                "Apply the change and return the COMPLETE updated plan as JSON, "
                "keeping exactly the same schema and keys. Do not drop unrelated content."
            )
            response = await self.client.chat.completions.create(
                model=self.model_name,
                messages=[{"role": "user", "content": prompt}],
                response_format={"type": "json_object"},
                temperature=0.3,
                max_tokens=6000,
            )
            plan = json.loads(response.choices[0].message.content)
            plan["detected_format"] = fmt

            if fmt == "xlsx_spreadsheet":
                doc_path = await asyncio.to_thread(self._render_xlsx_from_plan, plan)
            else:
                doc_path = await asyncio.to_thread(self._render_ppt_from_plan, plan)

            self.last_plan = plan
            return json.dumps({"status": "success", "file_path": doc_path,
                               "message": f"Successfully revised {fmt}"})
        except Exception as e:
            logger.exception("Error in revise_document")
            return json.dumps({"status": "error", "message": str(e)})

    # ------------------------------------------------------------------
    # FORMAT DETECTION
    # ------------------------------------------------------------------
    async def _detect_format(self, brief: str) -> str:
        """Explicit keywords win; otherwise let the LLM judge what suits the request best."""
        text = brief.lower()
        xlsx_kw = r"\b(excel|spreadsheet|xlsx|worksheet|workbook|tracker|budget|ledger|csv)\b"
        ppt_kw = r"\b(ppt|pptx|powerpoint|slides?|slide deck|deck|presentation|pitch)\b"
        has_x, has_p = re.search(xlsx_kw, text), re.search(ppt_kw, text)
        if has_x and not has_p:
            logger.info("Format chosen by KEYWORD rule: xlsx")
            return "xlsx_spreadsheet"
        if has_p and not has_x:
            logger.info("Format chosen by KEYWORD rule: pptx")
            return "pptx_slide_deck"

        logger.info("No clear keywords, using LLM fallback to choose format")
        prompt = f'''Decide the best output format for this request: "{brief}"

Choose "xlsx_spreadsheet" when the request is mainly about numbers or structured data:
budgets, expenses, sales/financial data, forecasts, KPIs, trackers, schedules, inventories,
comparisons of metrics, calculations, or anything to be sorted/filtered/charted.

Choose "pptx_slide_deck" when the request is mainly narrative or persuasive:
pitches, explanations, overviews, training material, project summaries, reports to present, proposals.

Return JSON only: {{"detected_format": "pptx_slide_deck"}} or {{"detected_format": "xlsx_spreadsheet"}}'''
        response = await self.client.chat.completions.create(
            model=self.model_name,
            messages=[{"role": "user", "content": prompt}],
            response_format={"type": "json_object"},
            temperature=0.0,
        )
        fmt = json.loads(response.choices[0].message.content).get("detected_format")
        logger.info(f"LLM fallback chose format: {fmt}")
        return fmt if fmt in self.supported_formats else "pptx_slide_deck"

    # ------------------------------------------------------------------
    # PLAN GENERATION (LLM)
    # ------------------------------------------------------------------
    async def _generate_full_ppt_plan_from_brief(self, brief: str) -> Dict[str, Any]:
        prompt = f'''
        Create a detailed 10-slide presentation plan for: "{brief}".

        STRICT CONTENT RULES:
        1. SECTION 0 (TITLE SLIDE): Generate a professional, catchy, and creative title (DO NOT repeat the prompt) and a sophisticated subtitle.
        2. CONTENT SECTIONS (1-9): For EVERY slide, you MUST provide 4 to 5 DETAILED bullet points.
        3. TEXT QUALITY: Each bullet point must be a full, informative sentence (12-18 words). DO NOT use short phrases or simple headings.
        4. IMAGE QUERIES: Provide an aesthetic Pexels query for every slide.

        Return ONLY a JSON object:
        {{
          "detected_format": "pptx_slide_deck",
          "content_plan": {{
            "sections": [
              {{ "title": "Creative Title", "description": "Subtitle", "image_search_query": "term", "elements": [] }},
              {{ "title": "Slide Title", "description": "Context", "image_search_query": "term", "elements": [ {{"description": "Detailed Sentence 1..."}} ] }}
            ]
          }}
        }}
        '''
        response = await self.client.chat.completions.create(
            model=self.model_name, messages=[{"role": "user", "content": prompt}],
            response_format={"type": "json_object"}, temperature=0.7, max_tokens=4000,
        )
        return json.loads(response.choices[0].message.content)

    async def _generate_full_xlsx_plan_from_brief(self, brief: str) -> Dict[str, Any]:
        prompt = f'''
Design a professional, data-rich Excel workbook for this request: "{brief}"

RULES:
- 3 to 4 sheets. Each sheet: 8 to 15 rows of realistic, internally consistent sample data.
- Every sheet has at least one text "label" column and at least one numeric column.
- Numbers must be RAW numbers (no "$", "," or "%" characters). Percentages are fractions (0.15 = 15%).
- Do NOT include totals rows or computed columns that need formulas; the app adds totals itself.
- Column "type" must be one of: text, integer, number, currency, percent, date (dates as YYYY-MM-DD).
- Column "total" is "sum", "average" or "none" (use "none" for text/date columns and for ratios like percent).
- Each sheet has 1 or 2 charts. "type" is "bar", "pie" or "line".
    * pie: parts of a whole (max 8 rows, ONE value column)
    * bar: comparing categories
    * line: trends over time
  "label_column" and "value_columns" are 0-based column indexes; value columns must be numeric.
- theme is one of: blue, green, purple, orange (pick what suits the topic).
- currency_symbol: the symbol that suits the topic/region (default "$").
- note: one short line saying the figures are illustrative sample data.

Return ONLY JSON in exactly this shape:
{{
  "detected_format": "xlsx_spreadsheet",
  "workbook_title": "Catchy workbook title",
  "subtitle": "One-line description",
  "theme": "blue",
  "currency_symbol": "$",
  "note": "Figures are illustrative sample data.",
  "sheets": [
    {{
      "name": "Sales by Region",
      "title": "Sales by Region - FY2026",
      "description": "What this sheet shows",
      "columns": [
        {{"header": "Region", "type": "text", "total": "none"}},
        {{"header": "Revenue", "type": "currency", "total": "sum"}},
        {{"header": "Growth", "type": "percent", "total": "none"}}
      ],
      "rows": [["North", 120000, 0.12], ["South", 98000, 0.08]],
      "charts": [
        {{"type": "pie", "title": "Revenue Share", "label_column": 0, "value_columns": [1]}},
        {{"type": "bar", "title": "Revenue by Region", "label_column": 0, "value_columns": [1]}}
      ]
    }}
  ]
}}
'''
        response = await self.client.chat.completions.create(
            model=self.model_name, messages=[{"role": "user", "content": prompt}],
            response_format={"type": "json_object"}, temperature=0.3, max_tokens=6000,
        )
        plan = json.loads(response.choices[0].message.content)
        if not plan.get("sheets"):
            raise ValueError("The model returned a spreadsheet plan with no sheets.")
        return plan

    # ------------------------------------------------------------------
    # EXCEL RENDERING
    # ------------------------------------------------------------------
    @staticmethod
    def _safe_sheet_name(name: Any, used: set) -> str:
        base = re.sub(r"[\\/*?:\[\]]", "-", str(name or "Sheet")).strip()[:28] or "Sheet"
        candidate, n = base, 2
        while candidate.lower() in {u.lower() for u in used}:
            candidate = f"{base[:26]}_{n}"
            n += 1
        used.add(candidate)
        return candidate

    @staticmethod
    def _coerce(value: Any, ctype: str):
        """Convert LLM output into a clean Python value for the given column type."""
        if value is None or value == "":
            return None
        if ctype in NUMERIC_TYPES:
            had_pct = False
            if isinstance(value, bool):
                return str(value)
            if isinstance(value, (int, float)):
                num = float(value)
            else:
                raw = str(value).strip()
                had_pct = raw.endswith("%")
                cleaned = re.sub(r"[^\d.\-]", "", raw.replace(",", ""))
                try:
                    num = float(cleaned)
                except ValueError:
                    return raw
            if had_pct:
                num /= 100
            elif ctype == "percent" and abs(num) > 1.5:  # model wrote 15 instead of 0.15
                num /= 100
            return int(round(num)) if ctype == "integer" else num
        if ctype == "date":
            for f in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%b %Y", "%B %Y"):
                try:
                    return datetime.strptime(str(value).strip(), f)
                except ValueError:
                    continue
            return str(value)
        return str(value)

    @staticmethod
    def _number_format(ctype: str, symbol: str) -> str:
        return {
            "integer": "#,##0",
            "number": "#,##0.00",
            "currency": f'"{symbol}"#,##0.00',
            "percent": "0.0%",
            "date": "yyyy-mm-dd",
        }.get(ctype, "General")

    def _normalize_columns(self, sheet: Dict[str, Any]) -> List[Dict[str, str]]:
        cols = []
        for c in sheet.get("columns") or []:
            if isinstance(c, str):
                c = {"header": c}
            ctype = str(c.get("type", "text")).lower()
            if ctype not in NUMERIC_TYPES | {"text", "date"}:
                ctype = "text"
            total = str(c.get("total", "none")).lower()
            if total not in ("sum", "average") or ctype not in NUMERIC_TYPES:
                total = "none"
            cols.append({"header": str(c.get("header", f"Column {len(cols) + 1}")), "type": ctype, "total": total})
        return cols

    @staticmethod
    def _print_setup(ws):
        """Landscape, fit to one page wide, so printing/PDF export looks tidy."""
        ws.page_setup.orientation = "landscape"
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)

    def _build_data_sheet(self, ws, sheet: Dict[str, Any], theme: Dict[str, Any], symbol: str, note: str):
        cols = self._normalize_columns(sheet)
        if not cols:
            return None
        ncols = len(cols)
        banner_width = max(ncols, 6)
        HDR, FIRST = 4, 5

        # --- banner + description ---
        for c in range(1, banner_width + 1):
            ws.cell(row=1, column=c).fill = PatternFill("solid", fgColor=theme["dark"])
        t = ws.cell(row=1, column=1, value=str(sheet.get("title") or sheet.get("name") or "Sheet"))
        t.font = Font(name=FONT_NAME, size=18, bold=True, color="FFFFFF")
        t.alignment = Alignment(vertical="center", indent=1)
        ws.row_dimensions[1].height = 36
        d = ws.cell(row=2, column=1, value=str(sheet.get("description") or ""))
        d.font = Font(name=FONT_NAME, size=10, italic=True, color="666666")
        d.alignment = Alignment(indent=1)

        # --- header row ---
        for i, col in enumerate(cols, start=1):
            h = ws.cell(row=HDR, column=i, value=col["header"])
            h.font = Font(name=FONT_NAME, size=11, bold=True, color="FFFFFF")
            h.fill = PatternFill("solid", fgColor=theme["accent"])
            h.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            h.border = BORDER
        ws.row_dimensions[HDR].height = 26

        # --- data rows ---
        rows = []
        for r in sheet.get("rows") or []:
            if isinstance(r, dict):
                r = [r.get(c["header"]) for c in cols]
            if isinstance(r, (list, tuple)):
                rows.append(list(r)[:ncols] + [None] * (ncols - len(r)))
        widths = [len(c["header"]) + 4 for c in cols]
        band_fill = PatternFill("solid", fgColor=theme["band"])

        for ri, row in enumerate(rows):
            excel_row = FIRST + ri
            for ci, (col, raw) in enumerate(zip(cols, row), start=1):
                val = self._coerce(raw, col["type"])
                cell = ws.cell(row=excel_row, column=ci, value=val)
                cell.font = Font(name=FONT_NAME, size=10)
                cell.border = BORDER
                if ri % 2 == 1:
                    cell.fill = band_fill
                if col["type"] in NUMERIC_TYPES:
                    cell.number_format = self._number_format(col["type"], symbol)
                    cell.alignment = Alignment(horizontal="right")
                elif col["type"] == "date":
                    cell.number_format = "yyyy-mm-dd"
                    cell.alignment = Alignment(horizontal="center")
                else:
                    cell.alignment = Alignment(horizontal="left", wrap_text=True)
                if isinstance(val, (int, float)):
                    shown = len(f"{val:,.2f}") + (len(symbol) if col["type"] == "currency" else 0)
                else:
                    shown = len(str(raw if raw is not None else ""))
                widths[ci - 1] = max(widths[ci - 1], shown + 4)

        last = FIRST + len(rows) - 1
        next_row = last + 1

        # --- totals row (real formulas) ---
        if rows and any(c["total"] != "none" for c in cols):
            for ci, col in enumerate(cols, start=1):
                cell = ws.cell(row=next_row, column=ci)
                cell.font = Font(name=FONT_NAME, size=10, bold=True, color=theme["dark"])
                cell.fill = PatternFill("solid", fgColor=theme["band"])
                cell.border = Border(top=Side(style="medium", color=theme["dark"]),
                                     bottom=Side(style="medium", color=theme["dark"]),
                                     left=THIN, right=THIN)
                letter = get_column_letter(ci)
                if col["total"] != "none":
                    fn = "SUM" if col["total"] == "sum" else "AVERAGE"
                    cell.value = f"={fn}({letter}{FIRST}:{letter}{last})"
                    cell.number_format = self._number_format(col["type"], symbol)
                    cell.alignment = Alignment(horizontal="right")
                elif ci == 1:
                    cell.value = "Total"
            next_row += 1

        # --- footnote ---
        if note:
            n = ws.cell(row=next_row + 1, column=1, value=str(note))
            n.font = Font(name=FONT_NAME, size=9, italic=True, color="888888")

        # --- sheet polish ---
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(i)].width = min(max(w, 10), 42)
        ws.freeze_panes = ws.cell(row=FIRST, column=1)
        if rows:
            ws.auto_filter.ref = f"A{HDR}:{get_column_letter(ncols)}{last}"
        ws.sheet_view.showGridLines = False
        ws.sheet_properties.tabColor = theme["accent"]
        self._print_setup(ws)

        info = {"ws": ws, "columns": cols, "hdr": HDR, "first": FIRST, "last": last, "ncols": ncols}
        return info

    def _make_chart(self, info: Dict[str, Any], spec: Dict[str, Any], theme: Dict[str, Any]):
        ws, cols = info["ws"], info["columns"]
        hdr, first, last = info["hdr"], info["first"], info["last"]
        if last < first:
            return None
        numeric_idx = [i for i, c in enumerate(cols) if c["type"] in NUMERIC_TYPES]
        if not numeric_idx:
            return None

        ctype = str(spec.get("type", "bar")).lower()
        if ctype not in ("bar", "pie", "line"):
            ctype = "bar"
        try:
            label_idx = int(spec.get("label_column", 0))
        except (TypeError, ValueError):
            label_idx = 0
        if not 0 <= label_idx < len(cols):
            label_idx = 0

        vals: List[int] = []
        for v in spec.get("value_columns") or []:
            try:
                v = int(v)
            except (TypeError, ValueError):
                continue
            if v in numeric_idx and v != label_idx and v not in vals:
                vals.append(v)
        if not vals:
            vals = [i for i in numeric_idx if i != label_idx][:1]
        if not vals:
            return None
        vals = vals[:1] if ctype == "pie" else vals[:4]

        palette = theme["palette"]
        if ctype == "pie":
            chart = PieChart()
        elif ctype == "line":
            chart = LineChart()
        else:
            chart = BarChart()
            chart.type = "col"
            chart.grouping = "clustered"

        chart.title = spec.get("title") or f"{cols[vals[0]]['header']} by {cols[label_idx]['header']}"
        chart.height, chart.width = 8.5, 16

        for v in vals:
            chart.add_data(Reference(ws, min_col=v + 1, min_row=hdr, max_row=last), titles_from_data=True)
        chart.set_categories(Reference(ws, min_col=label_idx + 1, min_row=first, max_row=last))

        if ctype == "pie":
            series = chart.series[0]
            for i in range(last - first + 1):
                pt = DataPoint(idx=i)
                pt.graphicalProperties.solidFill = palette[i % len(palette)]
                series.dPt.append(pt)
            chart.dataLabels = DataLabelList()
            chart.dataLabels.showPercent = True
            chart.dataLabels.showVal = False
            chart.dataLabels.showCatName = False
            chart.dataLabels.showSerName = False
            chart.dataLabels.showLegendKey = False
            chart.legend.position = "r"
        else:
            for i, s in enumerate(chart.series):
                color = palette[i % len(palette)]
                s.graphicalProperties.solidFill = color
                s.graphicalProperties.line.solidFill = color
                if ctype == "line":
                    s.smooth = False
                    s.graphicalProperties.line.width = 28575
                    s.marker.symbol = "circle"
                    s.marker.size = 7
                    s.marker.graphicalProperties.solidFill = color
                    s.marker.graphicalProperties.line.solidFill = color
            chart.x_axis.delete = False   # required so axes show up in newer Excel
            chart.y_axis.delete = False
            chart.y_axis.title = cols[vals[0]]["header"] if len(vals) == 1 else None
            if len(vals) == 1:
                chart.legend = None
            else:
                chart.legend.position = "b"
        return chart

    def _build_overview(self, ws, plan: Dict[str, Any], built: list, theme: Dict[str, Any]):
        for c in range(1, 21):
            ws.cell(row=1, column=c).fill = PatternFill("solid", fgColor=theme["dark"])
            ws.cell(row=2, column=c).fill = PatternFill("solid", fgColor=theme["accent"])
        t = ws.cell(row=1, column=2, value=str(plan.get("workbook_title") or "Workbook"))
        t.font = Font(name=FONT_NAME, size=24, bold=True, color="FFFFFF")
        t.alignment = Alignment(vertical="center")
        ws.row_dimensions[1].height = 48
        s = ws.cell(row=2, column=2, value=str(plan.get("subtitle") or ""))
        s.font = Font(name=FONT_NAME, size=11, italic=True, color="FFFFFF")
        ws.row_dimensions[2].height = 22

        ws.column_dimensions["A"].width = 3
        ws.column_dimensions["B"].width = 26
        ws.cell(row=4, column=2, value="CONTENTS").font = Font(name=FONT_NAME, size=12, bold=True, color=theme["dark"])
        row = 5
        for data_ws, sheet, _info in built:
            link = ws.cell(row=row, column=2, value=data_ws.title)
            link.hyperlink = f"#'{data_ws.title}'!A1"
            link.font = Font(name=FONT_NAME, size=11, underline="single", color=theme["accent"])
            desc = ws.cell(row=row, column=3, value=str(sheet.get("description") or ""))
            desc.font = Font(name=FONT_NAME, size=10, color="555555")
            row += 1

        # dashboard: first chart of each sheet, two per row
        row += 1
        ws.cell(row=row, column=2, value="DASHBOARD").font = Font(name=FONT_NAME, size=12, bold=True, color=theme["dark"])
        row += 1
        col_anchors = ["B", "L"]
        placed = 0
        for _ws, _sheet, info in built:
            chart = self._first_chart(info, theme)
            if not chart:
                continue
            ws.add_chart(chart, f"{col_anchors[placed % 2]}{row + (placed // 2) * 18}")
            placed += 1
        ws.sheet_view.showGridLines = False
        ws.sheet_properties.tabColor = theme["dark"]
        self._print_setup(ws)

    def _first_chart(self, info: Dict[str, Any], theme: Dict[str, Any]):
        for spec in info.get("chart_specs", []):
            chart = self._make_chart(info, spec, theme)
            if chart:
                return chart
        return None

    def _render_xlsx_from_plan(self, plan: Dict[str, Any]) -> str:
        theme = THEMES.get(str(plan.get("theme", "blue")).lower(), THEMES["blue"])
        symbol = str(plan.get("currency_symbol") or "$")
        if len(symbol) > 3 or '"' in symbol:
            symbol = "$"
        note = plan.get("note") or "Figures are illustrative sample data."

        wb = Workbook()
        overview = wb.active
        overview.title = "Overview"
        used = {"Overview"}
        built = []

        for sheet in (plan.get("sheets") or [])[:6]:
            ws = wb.create_sheet(self._safe_sheet_name(sheet.get("name"), used))
            info = self._build_data_sheet(ws, sheet, theme, symbol, note)
            if info is None:
                wb.remove(ws)
                continue

            specs = [s for s in (sheet.get("charts") or []) if isinstance(s, dict)][:2]
            if not specs:  # fallback so every sheet has a visual
                specs = [{"type": "bar"}]
            info["chart_specs"] = specs

            for i, spec in enumerate(specs):
                chart = self._make_chart(info, spec, theme)
                if chart:
                    ws.add_chart(chart, f"{get_column_letter(max(info['ncols'], 6) + 2)}{4 + i * 18}")
            built.append((ws, sheet, info))

        if not built:
            raise ValueError("Spreadsheet plan contained no usable sheets.")
        self._build_overview(overview, plan, built, theme)

        out_dir = "outputs"
        os.makedirs(out_dir, exist_ok=True)
        ts = datetime.utcnow().strftime("%Y%m%dT%H%M%SZ") + "_" + uuid.uuid4().hex[:4]
        path = os.path.join(out_dir, f"spreadsheet_{ts}.xlsx")
        wb.save(path)
        return path

    # ------------------------------------------------------------------
    # POWERPOINT RENDERING
    # ------------------------------------------------------------------
    def _fetch_image_for_query(self, query: str) -> Optional[str]:
        if not self.pexels_key:
            return None
        safe_name = "".join([c if c.isalnum() or c in "-_" else "_" for c in query.lower()[:30]])
        cached_path = Path(self.image_cache_dir) / f"{safe_name}_{random.randint(1, 100)}.jpg"
        headers = {"Authorization": self.pexels_key}
        params = {"query": query, "per_page": 5}
        try:
            resp = requests.get("https://api.pexels.com/v1/search", headers=headers, params=params, timeout=10)
            if resp.status_code == 200 and resp.json().get("photos"):
                chosen_photo = resp.json()["photos"][0]
                img_resp = requests.get(chosen_photo["src"]["large"], timeout=10)
                with open(cached_path, "wb") as f:
                    f.write(img_resp.content)
                return str(cached_path)
        except Exception:
            logger.warning("Image fetch failed for %s", query, exc_info=True)
        return None

    def _render_ppt_from_plan(self, plan: Dict[str, Any]) -> str:
        prs = Presentation()
        out_dir = "outputs"
        os.makedirs(out_dir, exist_ok=True)
        sections = plan.get("content_plan", {}).get("sections", [])

        BG_COLOR = RGBColor(230, 235, 242)
        TITLE_COLOR = RGBColor(10, 50, 100)
        TEXT_COLOR = RGBColor(50, 60, 70)
        ACCENT_COLOR = RGBColor(220, 80, 50)

        for i, section in enumerate(sections):
            slide = prs.slides.add_slide(prs.slide_layouts[6])  # 6 = truly blank (5 has an empty title placeholder)
            slide.background.fill.solid()
            slide.background.fill.fore_color.rgb = BG_COLOR

            if i == 0:
                accent = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(10), Inches(0.4))
                accent.fill.solid()
                accent.fill.fore_color.rgb = ACCENT_COLOR
                accent.line.fill.background()

                title_box = slide.shapes.add_textbox(Inches(1), Inches(2.2), Inches(8), Inches(2.0))
                tf = title_box.text_frame
                tf.word_wrap = True
                p = tf.paragraphs[0]
                p.text = section.get("title", "Presentation Title")
                p.font.size = Pt(44)
                p.font.bold = True
                p.font.color.rgb = TITLE_COLOR
                p.alignment = PP_ALIGN.CENTER

                desc = section.get("description", "").strip()
                if desc:
                    desc_box = slide.shapes.add_textbox(Inches(1), Inches(4.2), Inches(8), Inches(1.5))
                    desc_box.text_frame.word_wrap = True
                    p_d = desc_box.text_frame.paragraphs[0]
                    p_d.text = desc
                    p_d.font.size = Pt(22)
                    p_d.font.italic = True
                    p_d.font.color.rgb = TEXT_COLOR
                    p_d.alignment = PP_ALIGN.CENTER
                continue

            accent = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(0.4), Inches(0.1), Inches(0.6))
            accent.fill.solid()
            accent.fill.fore_color.rgb = ACCENT_COLOR
            accent.line.fill.background()

            title_box = slide.shapes.add_textbox(Inches(0.7), Inches(0.3), Inches(9.0), Inches(0.8))
            p_t = title_box.text_frame.paragraphs[0]
            p_t.text = section.get("title", "Section Title")
            p_t.font.size = Pt(32)
            p_t.font.bold = True
            p_t.font.color.rgb = TITLE_COLOR

            img_path = self._fetch_image_for_query(section.get("image_search_query", "modern"))
            text_width = Inches(5.0) if img_path else Inches(9.0)
            if img_path:
                try:
                    slide.shapes.add_picture(img_path, Inches(5.8), Inches(1.5), width=Inches(3.8))
                except Exception:
                    logger.warning("Could not place image", exc_info=True)

            y_pos = 1.3
            desc = section.get("description", "").strip()
            if desc:
                d_box = slide.shapes.add_textbox(Inches(0.5), Inches(y_pos), text_width, Inches(0.8))
                d_box.text_frame.word_wrap = True
                p_d = d_box.text_frame.paragraphs[0]
                p_d.text = desc
                p_d.font.size = Pt(16)
                p_d.font.italic = True
                p_d.font.color.rgb = TEXT_COLOR
                y_pos += 0.9

            elements = section.get("elements", [])
            if elements:
                box = slide.shapes.add_textbox(Inches(0.5), Inches(y_pos), text_width, Inches(7.0 - y_pos))
                tf = box.text_frame
                tf.word_wrap = True
                for idx, el in enumerate(elements):
                    p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
                    val = el if isinstance(el, str) else el.get("description", el.get("text", ""))
                    p.text = f"• {val}"
                    p.font.size = Pt(12)
                    p.font.color.rgb = TEXT_COLOR
                    p.space_after = Pt(6)

        ts = datetime.utcnow().strftime("%Y%m%dT%H%M%SZ") + "_" + uuid.uuid4().hex[:4]
        path = os.path.join(out_dir, f"presentation_{ts}.pptx")
        prs.save(path)
        return path

    def get_tools(self) -> dict[str, Any]:
        return {
            "generate_document": self.generate_document,
            "revise_document": self.revise_document,
        }
    
# import asyncio
# import json
# import logging
# import os
# import random
# import re
# import uuid
# from datetime import datetime
# from pathlib import Path
# from typing import Any, Dict, List, Optional

# import requests
# from openai import AsyncOpenAI
# from openpyxl import Workbook
# from openpyxl.chart import BarChart, LineChart, PieChart, Reference
# from openpyxl.chart.label import DataLabelList
# from openpyxl.chart.series import DataPoint
# from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
# from openpyxl.utils import get_column_letter
# from openpyxl.worksheet.properties import PageSetupProperties
# from pptx import Presentation
# from pptx.dml.color import RGBColor
# from pptx.enum.shapes import MSO_SHAPE
# from pptx.enum.text import PP_ALIGN
# from pptx.util import Inches, Pt

# logger = logging.getLogger(__name__)
# logger.setLevel(logging.DEBUG)

# # ----------------------------------------------------------------------------
# # Excel styling constants
# # ----------------------------------------------------------------------------
# FONT_NAME = "Arial"
# NUMERIC_TYPES = {"integer", "number", "currency", "percent"}

# _BASE_PALETTE = ["F28E2B", "59A14F", "E15759", "76B7B2", "EDC948", "B07AA1", "9C755F", "BAB0AC"]
# THEMES = {
#     "blue":   {"dark": "1F3A5F", "accent": "2E75B6", "band": "EAF1FB", "palette": ["2E75B6"] + _BASE_PALETTE},
#     "green":  {"dark": "1E4D3A", "accent": "2E8B57", "band": "E8F5EE", "palette": ["2E8B57"] + _BASE_PALETTE},
#     "purple": {"dark": "3B2A5E", "accent": "7B5EA7", "band": "F1ECF8", "palette": ["7B5EA7"] + _BASE_PALETTE},
#     "orange": {"dark": "5A2E12", "accent": "E07B39", "band": "FDF0E6", "palette": ["E07B39"] + _BASE_PALETTE},
# }

# THIN = Side(style="thin", color="D0D7DE")
# BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


# class DocumentCreatorToolset:
#     def __init__(self):
#         self.openai_key = os.getenv("OPENAI_API_KEY")
#         self.pexels_key = os.getenv("PEXELS_API_KEY")
#         self.client = AsyncOpenAI(api_key=self.openai_key)
#         self.model_name = "gpt-4o"

#         self.supported_formats = ["pptx_slide_deck", "xlsx_spreadsheet"]
#         self.last_brief: Optional[str] = None
#         self.last_plan: Optional[Dict[str, Any]] = None

#         self.image_cache_dir = "image_cache"
#         os.makedirs(self.image_cache_dir, exist_ok=True)

#     # ------------------------------------------------------------------
#     # TOOLS (exposed to the LLM agent)
#     # ------------------------------------------------------------------
#     async def generate_document(self, brief: str) -> str:
#         """Generate a PowerPoint deck or Excel spreadsheet from the user's brief, choosing the best format automatically."""
#         logger.info(f"📄 GENERATE_DOCUMENT CALLED: {brief}")
#         try:
#             self.last_brief = brief
#             fmt = await self._detect_format(brief)

#             if fmt == "xlsx_spreadsheet":
#                 logger.info("📊 Generating Spreadsheet...")
#                 plan = await self._generate_full_xlsx_plan_from_brief(brief)
#                 plan["detected_format"] = fmt
#                 doc_path = await asyncio.to_thread(self._render_xlsx_from_plan, plan)
#             else:
#                 fmt = "pptx_slide_deck"
#                 logger.info("🖼️ Generating Presentation...")
#                 plan = await self._generate_full_ppt_plan_from_brief(brief)
#                 plan["detected_format"] = fmt
#                 doc_path = await asyncio.to_thread(self._render_ppt_from_plan, plan)

#             self.last_plan = plan
#             return json.dumps({"status": "success", "file_path": doc_path,
#                                "message": f"Successfully generated {fmt}"})
#         except Exception as e:
#             logger.exception("Error in generate_document")
#             return json.dumps({"status": "error", "message": str(e)})

#     async def revise_document(self, instruction: str) -> str:
#         """Revise the most recently generated document according to the user's change request."""
#         logger.info(f"✏️ REVISE_DOCUMENT CALLED: {instruction}")
#         if not self.last_plan:
#             return json.dumps({"status": "error",
#                                "message": "No document has been generated yet. Ask for a new document first."})
#         try:
#             fmt = self.last_plan.get("detected_format", "pptx_slide_deck")
#             prompt = (
#                 "You are editing a document plan stored as JSON.\n"
#                 f"CURRENT PLAN:\n{json.dumps(self.last_plan)}\n\n"
#                 f"USER CHANGE REQUEST: {instruction}\n\n"
#                 "Apply the change and return the COMPLETE updated plan as JSON, "
#                 "keeping exactly the same schema and keys. Do not drop unrelated content."
#             )
#             response = await self.client.chat.completions.create(
#                 model=self.model_name,
#                 messages=[{"role": "user", "content": prompt}],
#                 response_format={"type": "json_object"},
#                 temperature=0.3,
#                 max_tokens=6000,
#             )
#             plan = json.loads(response.choices[0].message.content)
#             plan["detected_format"] = fmt

#             if fmt == "xlsx_spreadsheet":
#                 doc_path = await asyncio.to_thread(self._render_xlsx_from_plan, plan)
#             else:
#                 doc_path = await asyncio.to_thread(self._render_ppt_from_plan, plan)

#             self.last_plan = plan
#             return json.dumps({"status": "success", "file_path": doc_path,
#                                "message": f"Successfully revised {fmt}"})
#         except Exception as e:
#             logger.exception("Error in revise_document")
#             return json.dumps({"status": "error", "message": str(e)})

#     # ------------------------------------------------------------------
#     # FORMAT DETECTION
#     # ------------------------------------------------------------------
#     async def _detect_format(self, brief: str) -> str:
#         """Explicit keywords win; otherwise let the LLM judge what suits the request best."""
#         text = brief.lower()
#         xlsx_kw = r"\b(excel|spreadsheet|xlsx|worksheet|workbook|tracker|budget|ledger|csv)\b"
#         ppt_kw = r"\b(ppt|pptx|powerpoint|slides?|slide deck|deck|presentation|pitch)\b"
#         has_x, has_p = re.search(xlsx_kw, text), re.search(ppt_kw, text)
#         if has_x and not has_p:
#             return "xlsx_spreadsheet"
#         if has_p and not has_x:
#             return "pptx_slide_deck"

#         prompt = f'''Decide the best output format for this request: "{brief}"

# Choose "xlsx_spreadsheet" when the request is mainly about numbers or structured data:
# budgets, expenses, sales/financial data, forecasts, KPIs, trackers, schedules, inventories,
# comparisons of metrics, calculations, or anything to be sorted/filtered/charted.

# Choose "pptx_slide_deck" when the request is mainly narrative or persuasive:
# pitches, explanations, overviews, training material, project summaries, reports to present, proposals.

# Return JSON only: {{"detected_format": "pptx_slide_deck"}} or {{"detected_format": "xlsx_spreadsheet"}}'''
#         response = await self.client.chat.completions.create(
#             model=self.model_name,
#             messages=[{"role": "user", "content": prompt}],
#             response_format={"type": "json_object"},
#             temperature=0.0,
#         )
#         fmt = json.loads(response.choices[0].message.content).get("detected_format")
#         return fmt if fmt in self.supported_formats else "pptx_slide_deck"

#     # ------------------------------------------------------------------
#     # PLAN GENERATION (LLM)
#     # ------------------------------------------------------------------
#     async def _generate_full_ppt_plan_from_brief(self, brief: str) -> Dict[str, Any]:
#         prompt = f'''
#         Create a detailed 10-slide presentation plan for: "{brief}".

#         STRICT CONTENT RULES:
#         1. SECTION 0 (TITLE SLIDE): Generate a professional, catchy, and creative title (DO NOT repeat the prompt) and a sophisticated subtitle.
#         2. CONTENT SECTIONS (1-9): For EVERY slide, you MUST provide 4 to 5 DETAILED bullet points.
#         3. TEXT QUALITY: Each bullet point must be a full, informative sentence (12-18 words). DO NOT use short phrases or simple headings.
#         4. IMAGE QUERIES: Provide an aesthetic Pexels query for every slide.

#         Return ONLY a JSON object:
#         {{
#           "detected_format": "pptx_slide_deck",
#           "content_plan": {{
#             "sections": [
#               {{ "title": "Creative Title", "description": "Subtitle", "image_search_query": "term", "elements": [] }},
#               {{ "title": "Slide Title", "description": "Context", "image_search_query": "term", "elements": [ {{"description": "Detailed Sentence 1..."}} ] }}
#             ]
#           }}
#         }}
#         '''
#         response = await self.client.chat.completions.create(
#             model=self.model_name, messages=[{"role": "user", "content": prompt}],
#             response_format={"type": "json_object"}, temperature=0.7, max_tokens=4000,
#         )
#         return json.loads(response.choices[0].message.content)

#     async def _generate_full_xlsx_plan_from_brief(self, brief: str) -> Dict[str, Any]:
#         prompt = f'''
# Design a professional, data-rich Excel workbook for this request: "{brief}"

# RULES:
# - 3 to 4 sheets. Each sheet: 8 to 15 rows of realistic, internally consistent sample data.
# - Every sheet has at least one text "label" column and at least one numeric column.
# - Numbers must be RAW numbers (no "$", "," or "%" characters). Percentages are fractions (0.15 = 15%).
# - Do NOT include totals rows or computed columns that need formulas; the app adds totals itself.
# - Column "type" must be one of: text, integer, number, currency, percent, date (dates as YYYY-MM-DD).
# - Column "total" is "sum", "average" or "none" (use "none" for text/date columns and for ratios like percent).
# - Each sheet has 1 or 2 charts. "type" is "bar", "pie" or "line".
#     * pie: parts of a whole (max 8 rows, ONE value column)
#     * bar: comparing categories
#     * line: trends over time
#   "label_column" and "value_columns" are 0-based column indexes; value columns must be numeric.
# - theme is one of: blue, green, purple, orange (pick what suits the topic).
# - currency_symbol: the symbol that suits the topic/region (default "$").
# - note: one short line saying the figures are illustrative sample data.

# Return ONLY JSON in exactly this shape:
# {{
#   "detected_format": "xlsx_spreadsheet",
#   "workbook_title": "Catchy workbook title",
#   "subtitle": "One-line description",
#   "theme": "blue",
#   "currency_symbol": "$",
#   "note": "Figures are illustrative sample data.",
#   "sheets": [
#     {{
#       "name": "Sales by Region",
#       "title": "Sales by Region - FY2026",
#       "description": "What this sheet shows",
#       "columns": [
#         {{"header": "Region", "type": "text", "total": "none"}},
#         {{"header": "Revenue", "type": "currency", "total": "sum"}},
#         {{"header": "Growth", "type": "percent", "total": "none"}}
#       ],
#       "rows": [["North", 120000, 0.12], ["South", 98000, 0.08]],
#       "charts": [
#         {{"type": "pie", "title": "Revenue Share", "label_column": 0, "value_columns": [1]}},
#         {{"type": "bar", "title": "Revenue by Region", "label_column": 0, "value_columns": [1]}}
#       ]
#     }}
#   ]
# }}
# '''
#         response = await self.client.chat.completions.create(
#             model=self.model_name, messages=[{"role": "user", "content": prompt}],
#             response_format={"type": "json_object"}, temperature=0.3, max_tokens=6000,
#         )
#         plan = json.loads(response.choices[0].message.content)
#         if not plan.get("sheets"):
#             raise ValueError("The model returned a spreadsheet plan with no sheets.")
#         return plan

#     # ------------------------------------------------------------------
#     # EXCEL RENDERING
#     # ------------------------------------------------------------------
#     @staticmethod
#     def _safe_sheet_name(name: Any, used: set) -> str:
#         base = re.sub(r"[\\/*?:\[\]]", "-", str(name or "Sheet")).strip()[:28] or "Sheet"
#         candidate, n = base, 2
#         while candidate.lower() in {u.lower() for u in used}:
#             candidate = f"{base[:26]}_{n}"
#             n += 1
#         used.add(candidate)
#         return candidate

#     @staticmethod
#     def _coerce(value: Any, ctype: str):
#         """Convert LLM output into a clean Python value for the given column type."""
#         if value is None or value == "":
#             return None
#         if ctype in NUMERIC_TYPES:
#             had_pct = False
#             if isinstance(value, bool):
#                 return str(value)
#             if isinstance(value, (int, float)):
#                 num = float(value)
#             else:
#                 raw = str(value).strip()
#                 had_pct = raw.endswith("%")
#                 cleaned = re.sub(r"[^\d.\-]", "", raw.replace(",", ""))
#                 try:
#                     num = float(cleaned)
#                 except ValueError:
#                     return raw
#             if had_pct:
#                 num /= 100
#             elif ctype == "percent" and abs(num) > 1.5:  # model wrote 15 instead of 0.15
#                 num /= 100
#             return int(round(num)) if ctype == "integer" else num
#         if ctype == "date":
#             for f in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%b %Y", "%B %Y"):
#                 try:
#                     return datetime.strptime(str(value).strip(), f)
#                 except ValueError:
#                     continue
#             return str(value)
#         return str(value)

#     @staticmethod
#     def _number_format(ctype: str, symbol: str) -> str:
#         return {
#             "integer": "#,##0",
#             "number": "#,##0.00",
#             "currency": f'"{symbol}"#,##0.00',
#             "percent": "0.0%",
#             "date": "yyyy-mm-dd",
#         }.get(ctype, "General")

#     def _normalize_columns(self, sheet: Dict[str, Any]) -> List[Dict[str, str]]:
#         cols = []
#         for c in sheet.get("columns") or []:
#             if isinstance(c, str):
#                 c = {"header": c}
#             ctype = str(c.get("type", "text")).lower()
#             if ctype not in NUMERIC_TYPES | {"text", "date"}:
#                 ctype = "text"
#             total = str(c.get("total", "none")).lower()
#             if total not in ("sum", "average") or ctype not in NUMERIC_TYPES:
#                 total = "none"
#             cols.append({"header": str(c.get("header", f"Column {len(cols) + 1}")), "type": ctype, "total": total})
#         return cols

#     @staticmethod
#     def _print_setup(ws):
#         """Landscape, fit to one page wide, so printing/PDF export looks tidy."""
#         ws.page_setup.orientation = "landscape"
#         ws.page_setup.fitToWidth = 1
#         ws.page_setup.fitToHeight = 0
#         ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)

#     def _build_data_sheet(self, ws, sheet: Dict[str, Any], theme: Dict[str, Any], symbol: str, note: str):
#         cols = self._normalize_columns(sheet)
#         if not cols:
#             return None
#         ncols = len(cols)
#         banner_width = max(ncols, 6)
#         HDR, FIRST = 4, 5

#         # --- banner + description ---
#         for c in range(1, banner_width + 1):
#             ws.cell(row=1, column=c).fill = PatternFill("solid", fgColor=theme["dark"])
#         t = ws.cell(row=1, column=1, value=str(sheet.get("title") or sheet.get("name") or "Sheet"))
#         t.font = Font(name=FONT_NAME, size=18, bold=True, color="FFFFFF")
#         t.alignment = Alignment(vertical="center", indent=1)
#         ws.row_dimensions[1].height = 36
#         d = ws.cell(row=2, column=1, value=str(sheet.get("description") or ""))
#         d.font = Font(name=FONT_NAME, size=10, italic=True, color="666666")
#         d.alignment = Alignment(indent=1)

#         # --- header row ---
#         for i, col in enumerate(cols, start=1):
#             h = ws.cell(row=HDR, column=i, value=col["header"])
#             h.font = Font(name=FONT_NAME, size=11, bold=True, color="FFFFFF")
#             h.fill = PatternFill("solid", fgColor=theme["accent"])
#             h.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
#             h.border = BORDER
#         ws.row_dimensions[HDR].height = 26

#         # --- data rows ---
#         rows = []
#         for r in sheet.get("rows") or []:
#             if isinstance(r, dict):
#                 r = [r.get(c["header"]) for c in cols]
#             if isinstance(r, (list, tuple)):
#                 rows.append(list(r)[:ncols] + [None] * (ncols - len(r)))
#         widths = [len(c["header"]) + 4 for c in cols]
#         band_fill = PatternFill("solid", fgColor=theme["band"])

#         for ri, row in enumerate(rows):
#             excel_row = FIRST + ri
#             for ci, (col, raw) in enumerate(zip(cols, row), start=1):
#                 val = self._coerce(raw, col["type"])
#                 cell = ws.cell(row=excel_row, column=ci, value=val)
#                 cell.font = Font(name=FONT_NAME, size=10)
#                 cell.border = BORDER
#                 if ri % 2 == 1:
#                     cell.fill = band_fill
#                 if col["type"] in NUMERIC_TYPES:
#                     cell.number_format = self._number_format(col["type"], symbol)
#                     cell.alignment = Alignment(horizontal="right")
#                 elif col["type"] == "date":
#                     cell.number_format = "yyyy-mm-dd"
#                     cell.alignment = Alignment(horizontal="center")
#                 else:
#                     cell.alignment = Alignment(horizontal="left", wrap_text=True)
#                 if isinstance(val, (int, float)):
#                     shown = len(f"{val:,.2f}") + (len(symbol) if col["type"] == "currency" else 0)
#                 else:
#                     shown = len(str(raw if raw is not None else ""))
#                 widths[ci - 1] = max(widths[ci - 1], shown + 4)

#         last = FIRST + len(rows) - 1
#         next_row = last + 1

#         # --- totals row (real formulas) ---
#         if rows and any(c["total"] != "none" for c in cols):
#             for ci, col in enumerate(cols, start=1):
#                 cell = ws.cell(row=next_row, column=ci)
#                 cell.font = Font(name=FONT_NAME, size=10, bold=True, color=theme["dark"])
#                 cell.fill = PatternFill("solid", fgColor=theme["band"])
#                 cell.border = Border(top=Side(style="medium", color=theme["dark"]),
#                                      bottom=Side(style="medium", color=theme["dark"]),
#                                      left=THIN, right=THIN)
#                 letter = get_column_letter(ci)
#                 if col["total"] != "none":
#                     fn = "SUM" if col["total"] == "sum" else "AVERAGE"
#                     cell.value = f"={fn}({letter}{FIRST}:{letter}{last})"
#                     cell.number_format = self._number_format(col["type"], symbol)
#                     cell.alignment = Alignment(horizontal="right")
#                 elif ci == 1:
#                     cell.value = "Total"
#             next_row += 1

#         # --- footnote ---
#         if note:
#             n = ws.cell(row=next_row + 1, column=1, value=str(note))
#             n.font = Font(name=FONT_NAME, size=9, italic=True, color="888888")

#         # --- sheet polish ---
#         for i, w in enumerate(widths, start=1):
#             ws.column_dimensions[get_column_letter(i)].width = min(max(w, 10), 42)
#         ws.freeze_panes = ws.cell(row=FIRST, column=1)
#         if rows:
#             ws.auto_filter.ref = f"A{HDR}:{get_column_letter(ncols)}{last}"
#         ws.sheet_view.showGridLines = False
#         ws.sheet_properties.tabColor = theme["accent"]
#         self._print_setup(ws)

#         info = {"ws": ws, "columns": cols, "hdr": HDR, "first": FIRST, "last": last, "ncols": ncols}
#         return info

#     def _make_chart(self, info: Dict[str, Any], spec: Dict[str, Any], theme: Dict[str, Any]):
#         ws, cols = info["ws"], info["columns"]
#         hdr, first, last = info["hdr"], info["first"], info["last"]
#         if last < first:
#             return None
#         numeric_idx = [i for i, c in enumerate(cols) if c["type"] in NUMERIC_TYPES]
#         if not numeric_idx:
#             return None

#         ctype = str(spec.get("type", "bar")).lower()
#         if ctype not in ("bar", "pie", "line"):
#             ctype = "bar"
#         try:
#             label_idx = int(spec.get("label_column", 0))
#         except (TypeError, ValueError):
#             label_idx = 0
#         if not 0 <= label_idx < len(cols):
#             label_idx = 0

#         vals: List[int] = []
#         for v in spec.get("value_columns") or []:
#             try:
#                 v = int(v)
#             except (TypeError, ValueError):
#                 continue
#             if v in numeric_idx and v != label_idx and v not in vals:
#                 vals.append(v)
#         if not vals:
#             vals = [i for i in numeric_idx if i != label_idx][:1]
#         if not vals:
#             return None
#         vals = vals[:1] if ctype == "pie" else vals[:4]

#         palette = theme["palette"]
#         if ctype == "pie":
#             chart = PieChart()
#         elif ctype == "line":
#             chart = LineChart()
#         else:
#             chart = BarChart()
#             chart.type = "col"
#             chart.grouping = "clustered"

#         chart.title = spec.get("title") or f"{cols[vals[0]]['header']} by {cols[label_idx]['header']}"
#         chart.height, chart.width = 8.5, 16

#         for v in vals:
#             chart.add_data(Reference(ws, min_col=v + 1, min_row=hdr, max_row=last), titles_from_data=True)
#         chart.set_categories(Reference(ws, min_col=label_idx + 1, min_row=first, max_row=last))

#         if ctype == "pie":
#             series = chart.series[0]
#             for i in range(last - first + 1):
#                 pt = DataPoint(idx=i)
#                 pt.graphicalProperties.solidFill = palette[i % len(palette)]
#                 series.dPt.append(pt)
#             chart.dataLabels = DataLabelList()
#             chart.dataLabels.showPercent = True
#             chart.dataLabels.showVal = False
#             chart.dataLabels.showCatName = False
#             chart.dataLabels.showSerName = False
#             chart.dataLabels.showLegendKey = False
#             chart.legend.position = "r"
#         else:
#             for i, s in enumerate(chart.series):
#                 color = palette[i % len(palette)]
#                 s.graphicalProperties.solidFill = color
#                 s.graphicalProperties.line.solidFill = color
#                 if ctype == "line":
#                     s.smooth = False
#                     s.graphicalProperties.line.width = 28575
#                     s.marker.symbol = "circle"
#                     s.marker.size = 7
#                     s.marker.graphicalProperties.solidFill = color
#                     s.marker.graphicalProperties.line.solidFill = color
#             chart.x_axis.delete = False   # required so axes show up in newer Excel
#             chart.y_axis.delete = False
#             chart.y_axis.title = cols[vals[0]]["header"] if len(vals) == 1 else None
#             if len(vals) == 1:
#                 chart.legend = None
#             else:
#                 chart.legend.position = "b"
#         return chart

#     def _build_overview(self, ws, plan: Dict[str, Any], built: list, theme: Dict[str, Any]):
#         for c in range(1, 21):
#             ws.cell(row=1, column=c).fill = PatternFill("solid", fgColor=theme["dark"])
#             ws.cell(row=2, column=c).fill = PatternFill("solid", fgColor=theme["accent"])
#         t = ws.cell(row=1, column=2, value=str(plan.get("workbook_title") or "Workbook"))
#         t.font = Font(name=FONT_NAME, size=24, bold=True, color="FFFFFF")
#         t.alignment = Alignment(vertical="center")
#         ws.row_dimensions[1].height = 48
#         s = ws.cell(row=2, column=2, value=str(plan.get("subtitle") or ""))
#         s.font = Font(name=FONT_NAME, size=11, italic=True, color="FFFFFF")
#         ws.row_dimensions[2].height = 22

#         ws.column_dimensions["A"].width = 3
#         ws.column_dimensions["B"].width = 26
#         ws.cell(row=4, column=2, value="CONTENTS").font = Font(name=FONT_NAME, size=12, bold=True, color=theme["dark"])
#         row = 5
#         for data_ws, sheet, _info in built:
#             link = ws.cell(row=row, column=2, value=data_ws.title)
#             link.hyperlink = f"#'{data_ws.title}'!A1"
#             link.font = Font(name=FONT_NAME, size=11, underline="single", color=theme["accent"])
#             desc = ws.cell(row=row, column=3, value=str(sheet.get("description") or ""))
#             desc.font = Font(name=FONT_NAME, size=10, color="555555")
#             row += 1

#         # dashboard: first chart of each sheet, two per row
#         row += 1
#         ws.cell(row=row, column=2, value="DASHBOARD").font = Font(name=FONT_NAME, size=12, bold=True, color=theme["dark"])
#         row += 1
#         col_anchors = ["B", "L"]
#         placed = 0
#         for _ws, _sheet, info in built:
#             chart = self._first_chart(info, theme)
#             if not chart:
#                 continue
#             ws.add_chart(chart, f"{col_anchors[placed % 2]}{row + (placed // 2) * 18}")
#             placed += 1
#         ws.sheet_view.showGridLines = False
#         ws.sheet_properties.tabColor = theme["dark"]
#         self._print_setup(ws)

#     def _first_chart(self, info: Dict[str, Any], theme: Dict[str, Any]):
#         for spec in info.get("chart_specs", []):
#             chart = self._make_chart(info, spec, theme)
#             if chart:
#                 return chart
#         return None

#     def _render_xlsx_from_plan(self, plan: Dict[str, Any]) -> str:
#         theme = THEMES.get(str(plan.get("theme", "blue")).lower(), THEMES["blue"])
#         symbol = str(plan.get("currency_symbol") or "$")
#         if len(symbol) > 3 or '"' in symbol:
#             symbol = "$"
#         note = plan.get("note") or "Figures are illustrative sample data."

#         wb = Workbook()
#         overview = wb.active
#         overview.title = "Overview"
#         used = {"Overview"}
#         built = []

#         for sheet in (plan.get("sheets") or [])[:6]:
#             ws = wb.create_sheet(self._safe_sheet_name(sheet.get("name"), used))
#             info = self._build_data_sheet(ws, sheet, theme, symbol, note)
#             if info is None:
#                 wb.remove(ws)
#                 continue

#             specs = [s for s in (sheet.get("charts") or []) if isinstance(s, dict)][:2]
#             if not specs:  # fallback so every sheet has a visual
#                 specs = [{"type": "bar"}]
#             info["chart_specs"] = specs

#             for i, spec in enumerate(specs):
#                 chart = self._make_chart(info, spec, theme)
#                 if chart:
#                     ws.add_chart(chart, f"{get_column_letter(max(info['ncols'], 6) + 2)}{4 + i * 18}")
#             built.append((ws, sheet, info))

#         if not built:
#             raise ValueError("Spreadsheet plan contained no usable sheets.")
#         self._build_overview(overview, plan, built, theme)

#         out_dir = "outputs"
#         os.makedirs(out_dir, exist_ok=True)
#         ts = datetime.utcnow().strftime("%Y%m%dT%H%M%SZ") + "_" + uuid.uuid4().hex[:4]
#         path = os.path.join(out_dir, f"spreadsheet_{ts}.xlsx")
#         wb.save(path)
#         return path

#     # ------------------------------------------------------------------
#     # POWERPOINT RENDERING
#     # ------------------------------------------------------------------
#     def _fetch_image_for_query(self, query: str) -> Optional[str]:
#         if not self.pexels_key:
#             return None
#         safe_name = "".join([c if c.isalnum() or c in "-_" else "_" for c in query.lower()[:30]])
#         cached_path = Path(self.image_cache_dir) / f"{safe_name}_{random.randint(1, 100)}.jpg"
#         headers = {"Authorization": self.pexels_key}
#         params = {"query": query, "per_page": 5}
#         try:
#             resp = requests.get("https://api.pexels.com/v1/search", headers=headers, params=params, timeout=10)
#             if resp.status_code == 200 and resp.json().get("photos"):
#                 chosen_photo = resp.json()["photos"][0]
#                 img_resp = requests.get(chosen_photo["src"]["large"], timeout=10)
#                 with open(cached_path, "wb") as f:
#                     f.write(img_resp.content)
#                 return str(cached_path)
#         except Exception:
#             logger.warning("Image fetch failed for %s", query, exc_info=True)
#         return None

#     def _render_ppt_from_plan(self, plan: Dict[str, Any]) -> str:
#         prs = Presentation()
#         out_dir = "outputs"
#         os.makedirs(out_dir, exist_ok=True)
#         sections = plan.get("content_plan", {}).get("sections", [])

#         BG_COLOR = RGBColor(230, 235, 242)
#         TITLE_COLOR = RGBColor(10, 50, 100)
#         TEXT_COLOR = RGBColor(50, 60, 70)
#         ACCENT_COLOR = RGBColor(220, 80, 50)

#         for i, section in enumerate(sections):
#             slide = prs.slides.add_slide(prs.slide_layouts[6])  # 6 = truly blank (5 has an empty title placeholder)
#             slide.background.fill.solid()
#             slide.background.fill.fore_color.rgb = BG_COLOR

#             if i == 0:
#                 accent = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(10), Inches(0.4))
#                 accent.fill.solid()
#                 accent.fill.fore_color.rgb = ACCENT_COLOR
#                 accent.line.fill.background()

#                 title_box = slide.shapes.add_textbox(Inches(1), Inches(2.2), Inches(8), Inches(2.0))
#                 tf = title_box.text_frame
#                 tf.word_wrap = True
#                 p = tf.paragraphs[0]
#                 p.text = section.get("title", "Presentation Title")
#                 p.font.size = Pt(44)
#                 p.font.bold = True
#                 p.font.color.rgb = TITLE_COLOR
#                 p.alignment = PP_ALIGN.CENTER

#                 desc = section.get("description", "").strip()
#                 if desc:
#                     desc_box = slide.shapes.add_textbox(Inches(1), Inches(4.2), Inches(8), Inches(1.5))
#                     desc_box.text_frame.word_wrap = True
#                     p_d = desc_box.text_frame.paragraphs[0]
#                     p_d.text = desc
#                     p_d.font.size = Pt(22)
#                     p_d.font.italic = True
#                     p_d.font.color.rgb = TEXT_COLOR
#                     p_d.alignment = PP_ALIGN.CENTER
#                 continue

#             accent = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(0.4), Inches(0.1), Inches(0.6))
#             accent.fill.solid()
#             accent.fill.fore_color.rgb = ACCENT_COLOR
#             accent.line.fill.background()

#             title_box = slide.shapes.add_textbox(Inches(0.7), Inches(0.3), Inches(9.0), Inches(0.8))
#             p_t = title_box.text_frame.paragraphs[0]
#             p_t.text = section.get("title", "Section Title")
#             p_t.font.size = Pt(32)
#             p_t.font.bold = True
#             p_t.font.color.rgb = TITLE_COLOR

#             img_path = self._fetch_image_for_query(section.get("image_search_query", "modern"))
#             text_width = Inches(5.0) if img_path else Inches(9.0)
#             if img_path:
#                 try:
#                     slide.shapes.add_picture(img_path, Inches(5.8), Inches(1.5), width=Inches(3.8))
#                 except Exception:
#                     logger.warning("Could not place image", exc_info=True)

#             y_pos = 1.3
#             desc = section.get("description", "").strip()
#             if desc:
#                 d_box = slide.shapes.add_textbox(Inches(0.5), Inches(y_pos), text_width, Inches(0.8))
#                 d_box.text_frame.word_wrap = True
#                 p_d = d_box.text_frame.paragraphs[0]
#                 p_d.text = desc
#                 p_d.font.size = Pt(16)
#                 p_d.font.italic = True
#                 p_d.font.color.rgb = TEXT_COLOR
#                 y_pos += 0.9

#             elements = section.get("elements", [])
#             if elements:
#                 box = slide.shapes.add_textbox(Inches(0.5), Inches(y_pos), text_width, Inches(7.0 - y_pos))
#                 tf = box.text_frame
#                 tf.word_wrap = True
#                 for idx, el in enumerate(elements):
#                     p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
#                     val = el if isinstance(el, str) else el.get("description", el.get("text", ""))
#                     p.text = f"• {val}"
#                     p.font.size = Pt(12)
#                     p.font.color.rgb = TEXT_COLOR
#                     p.space_after = Pt(6)

#         ts = datetime.utcnow().strftime("%Y%m%dT%H%M%SZ") + "_" + uuid.uuid4().hex[:4]
#         path = os.path.join(out_dir, f"presentation_{ts}.pptx")
#         prs.save(path)
#         return path

#     def get_tools(self) -> dict[str, Any]:
#         return {
#             "generate_document": self.generate_document,
#             "revise_document": self.revise_document,
#         }

# # import json
# # import os
# # import logging
# # import requests
# # import random
# # from pathlib import Path
# # from datetime import datetime
# # from typing import Dict, Any

# # from openai import AsyncOpenAI
# # from pptx import Presentation
# # from pptx.util import Pt, Inches
# # from pptx.enum.shapes import MSO_SHAPE
# # from pptx.dml.color import RGBColor
# # from openpyxl import Workbook
# # from openpyxl.chart import PieChart, BarChart, Reference
# # from openpyxl.styles import PatternFill, Font

# # logger = logging.getLogger(__name__)
# # logger.setLevel(logging.DEBUG)

# # class DocumentCreatorToolset:
# #     def __init__(self):
# #         self.openai_key = os.getenv("OPENAI_API_KEY")
# #         self.pexels_key = os.getenv("PEXELS_API_KEY")
# #         self.client = AsyncOpenAI(api_key=self.openai_key)
# #         self.model_name = "gpt-4o"
        
# #         self.supported_formats = ["pptx_slide_deck", "xlsx_spreadsheet"]
# #         self.last_brief = None
# #         self.last_plan = None
        
# #         self.image_cache_dir = "image_cache"
# #         os.makedirs(self.image_cache_dir, exist_ok=True)

# #     async def generate_document(self, brief: str) -> str:
# #         logger.info(f"📄 GENERATE_DOCUMENT CALLED: {brief}")
# #         try:
# #             self.last_brief = brief
# #             plan = await self._detect_format_and_plan(brief)
# #             fmt = plan.get("detected_format", "pptx_slide_deck")
            
# #             if fmt == "xlsx_spreadsheet":
# #                 logger.info("📊 Generating Spreadsheet...")
# #                 plan = await self._generate_full_xlsx_plan_from_brief(brief)
# #                 doc_path = self._render_xlsx_from_plan(plan)
# #             else:
# #                 logger.info("🖼️ Generating Presentation...")
# #                 plan = await self._generate_full_ppt_plan_from_brief(brief)
# #                 doc_path = self._render_ppt_from_plan(plan)

# #             self.last_plan = plan
# #             return json.dumps({"status": "success", "file_path": doc_path, "message": f"Successfully generated {fmt}"})
# #         except Exception as e:
# #             logger.exception("Error in generate_document")
# #             return json.dumps({"status": "error", "message": str(e)})

# #     async def _detect_format_and_plan(self, brief: str) -> Dict[str, Any]:
# #         prompt = f'Analyze "{brief}". Return JSON: {{ "detected_format": "pptx_slide_deck" }} or "xlsx_spreadsheet".'
# #         response = await self.client.chat.completions.create(
# #             model=self.model_name, messages=[{"role": "user", "content": prompt}],
# #             response_format={"type": "json_object"}, temperature=0.1
# #         )
# #         return json.loads(response.choices[0].message.content)

# #     async def _generate_full_ppt_plan_from_brief(self, brief: str) -> Dict[str, Any]:
# #         prompt = f'''
# #         Create a detailed 10-slide presentation plan for: "{brief}".
        
# #         STRICT CONTENT RULES:
# #         1. SECTION 0 (TITLE SLIDE): Generate a professional, catchy, and creative title (DO NOT repeat the prompt) and a sophisticated subtitle.
# #         2. CONTENT SECTIONS (1-9): For EVERY slide, you MUST provide 5 to 6 DETAILED bullet points.
# #         3. TEXT QUALITY: Each bullet point must be a full, informative sentence (15-20 words). DO NOT use short phrases or simple headings.
# #         4. IMAGE QUERIES: Provide an aesthetic Pexels query for every slide.
        
# #         Return ONLY a JSON object:
# #         {{
# #           "detected_format": "pptx_slide_deck",
# #           "content_plan": {{
# #             "sections": [
# #               {{ "title": "Creative Title", "description": "Subtitle", "image_search_query": "term", "elements": [] }},
# #               {{ "title": "Slide Title", "description": "Context", "image_search_query": "term", "elements": [ {{"description": "Detailed Sentence 1..."}}, ... ] }}
# #             ]
# #           }}
# #         }}
# #         '''
# #         response = await self.client.chat.completions.create(
# #             model=self.model_name, messages=[{"role": "user", "content": prompt}],
# #             response_format={"type": "json_object"}, temperature=0.7, max_tokens=4000
# #         )
# #         return json.loads(response.choices[0].message.content)

# #     async def _generate_full_xlsx_plan_from_brief(self, brief: str) -> Dict[str, Any]:
# #         prompt = f'Design a real-data spreadsheet for: {brief}. 4 sheets. RAW numbers. JSON ONLY.'
# #         response = await self.client.chat.completions.create(
# #             model=self.model_name, messages=[{"role": "user", "content": prompt}],
# #             response_format={"type": "json_object"}, temperature=0.2 
# #         )
# #         return json.loads(response.choices[0].message.content)

# #     def _fetch_image_for_query(self, query: str) -> str:
# #         # CLEAN SYNTAX: Fixed the parenthesis/bracket closing from previous error
# #         safe_name = "".join([c if c.isalnum() or c in "-_" else "_" for c in query.lower()[:30]])
# #         cached_path = Path(self.image_cache_dir) / f"{safe_name}_{random.randint(1,100)}.jpg"
# #         headers = {"Authorization": self.pexels_key}
# #         params = {"query": query, "per_page": 5} 
# #         try:
# #             resp = requests.get("https://api.pexels.com/v1/search", headers=headers, params=params, timeout=10)
# #             if resp.status_code == 200 and resp.json().get("photos"):
# #                 chosen_photo = resp.json()["photos"][0]
# #                 img_resp = requests.get(chosen_photo["src"]["large"], timeout=10)
# #                 with open(cached_path, "wb") as f:
# #                     f.write(img_resp.content)
# #                 return str(cached_path)
# #         except Exception:
# #             return None

# #     def _render_ppt_from_plan(self, plan: Dict[str, Any]) -> str:
# #         prs = Presentation()
# #         out_dir = "outputs"
# #         os.makedirs(out_dir, exist_ok=True)
# #         sections = plan.get("content_plan", {}).get("sections", [])

# #         BG_COLOR = RGBColor(230, 235, 242) 
# #         TITLE_COLOR = RGBColor(10, 50, 100) 
# #         TEXT_COLOR = RGBColor(50, 60, 70)  
# #         ACCENT_COLOR = RGBColor(220, 80, 50) # Orange

# #         for i, section in enumerate(sections):
# #             slide = prs.slides.add_slide(prs.slide_layouts[5]) # Blank
# #             slide.background.fill.solid()
# #             slide.background.fill.fore_color.rgb = BG_COLOR

# #             if i == 0:
# #                 # --- TITLE SLIDE (ORANGE BAR ON TOP) ---
# #                 accent = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(10), Inches(0.4))
# #                 accent.fill.solid()
# #                 accent.fill.fore_color.rgb = ACCENT_COLOR
# #                 accent.line.fill.background()

# #                 title_box = slide.shapes.add_textbox(Inches(1), Inches(2.2), Inches(8), Inches(2.0))
# #                 tf = title_box.text_frame
# #                 tf.word_wrap = True
# #                 p = tf.paragraphs[0]
# #                 p.text = section.get("title", "Presentation Title")
# #                 p.font.size = Pt(44)
# #                 p.font.bold = True
# #                 p.font.color.rgb = TITLE_COLOR
# #                 p.alignment = 1 # Center

# #                 desc = section.get("description", "").strip()
# #                 if desc:
# #                     desc_box = slide.shapes.add_textbox(Inches(1), Inches(4.2), Inches(8), Inches(1.5))
# #                     p_d = desc_box.text_frame.paragraphs[0]
# #                     p_d.text = desc
# #                     p_d.font.size = Pt(22)
# #                     p_d.font.italic = True
# #                     p_d.font.color.rgb = TEXT_COLOR
# #                     p_d.alignment = 1
# #                 continue

# #             # --- CONTENT SLIDES (ORANGE BAR ON LEFT) ---
# #             accent = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(0.4), Inches(0.1), Inches(0.6))
# #             accent.fill.solid()
# #             accent.fill.fore_color.rgb = ACCENT_COLOR
# #             accent.line.fill.background()

# #             title_box = slide.shapes.add_textbox(Inches(0.7), Inches(0.3), Inches(9.0), Inches(0.8))
# #             p_t = title_box.text_frame.paragraphs[0]
# #             p_t.text = section.get("title", "Section Title")
# #             p_t.font.size = Pt(32)
# #             p_t.font.bold = True
# #             p_t.font.color.rgb = TITLE_COLOR 

# #             img_path = self._fetch_image_for_query(section.get("image_search_query", "modern"))
# #             text_width = Inches(5.0) if img_path else Inches(9.0)
# #             if img_path:
# #                 try:
# #                     slide.shapes.add_picture(img_path, Inches(5.8), Inches(1.5), width=Inches(3.8))
# #                 except Exception: pass

# #             y_pos = 1.3
# #             desc = section.get("description", "").strip()
# #             if desc:
# #                 d_box = slide.shapes.add_textbox(Inches(0.5), Inches(y_pos), text_width, Inches(1.2))
# #                 d_box.text_frame.word_wrap = True
# #                 p_d = d_box.text_frame.paragraphs[0]
# #                 p_d.text = desc
# #                 p_d.font.size = Pt(16)
# #                 p_d.font.color.rgb = TEXT_COLOR 
# #                 y_pos += 1.0

# #             elements = section.get("elements", [])
# #             if elements:
# #                 box = slide.shapes.add_textbox(Inches(0.5), Inches(y_pos), text_width, Inches(5.0))
# #                 tf = box.text_frame
# #                 tf.word_wrap = True
# #                 for idx, el in enumerate(elements):
# #                     p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
# #                     val = el if isinstance(el, str) else el.get("description", el.get("text", ""))
# #                     p.text = f"• {val}"
# #                     p.font.size = Pt(14)
# #                     p.font.color.rgb = TEXT_COLOR
# #                     p.space_after = Pt(8)

# #         ts = datetime.utcnow().strftime("%Y%m%dT%H%M%SZ")
# #         path = os.path.join(out_dir, f"presentation_{ts}.pptx")
# #         prs.save(path)
# #         return path
        
# #     def _render_xlsx_from_plan(self, plan: Dict[str, Any]) -> str:
# #         wb = Workbook()
# #         out_dir = "outputs"
# #         os.makedirs(out_dir, exist_ok=True)
# #         ts = datetime.utcnow().strftime("%Y%m%dT%H%M%SZ")
# #         path = os.path.join(out_dir, f"spreadsheet_{ts}.xlsx")
# #         wb.save(path)
# #         return path

# #     def get_tools(self) -> dict[str, Any]:
# #         return {'generate_document': self.generate_document}

