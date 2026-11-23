"""Financial Report Generator v3 — n8n local Excel service.

Requirements: fastapi, uvicorn, openpyxl
Install (inside existing venv): python -m pip install fastapi uvicorn openpyxl
Launch: python -m uvicorn excel_server:app --host 127.0.0.1 --port 8000

POST /generate-report with the `report` object from n8n's
`Parse Financial Analyst Response` node. Returns .xlsx as binary.

Debug without n8n:  python excel_server.py report.json out.xlsx

Design goals of v3
------------------
* Written for a reader with NO finance background (think: a high-school student).
* One glance at the Overview sheet answers "is this business healthy?":
  score gauge, plain rating word, report card, key numbers and charts.
* Every number is explained in everyday words (Health Checks, Money Words).
* The health score, coverage and category scores are NEVER recalculated here.
  They come from n8n exactly as supplied; the workbook only presents them.
* All values are written as plain numbers (no formulas), so the file looks
  identical in Excel, Google Sheets, LibreOffice and phone previewers.
"""
from __future__ import annotations

import io
import math
import re
import sys
from collections import OrderedDict
from datetime import datetime
from typing import Any, Optional

from openpyxl import Workbook
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.chart import BarChart, DoughnutChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.series import DataPoint
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.text import RichText, Text
from openpyxl.chart.title import Title
from openpyxl.drawing.line import LineProperties
from openpyxl.worksheet.pagebreak import Break
from openpyxl.drawing.text import CharacterProperties, Paragraph, ParagraphProperties, RegularTextRun
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

try:  # the web service is optional when the module is used as a library
    from fastapi import Body, FastAPI, HTTPException
    from fastapi.responses import Response
    app = FastAPI(title="Financial Report Generator", version="3.0.0")
except Exception:  # pragma: no cover
    app = None

# ---------------------------------------------------------------------------
# Visual system
# ---------------------------------------------------------------------------
NAVY, INK, MUTED, LINE = "1B2A41", "22313F", "6B7C8C", "DCE3EA"
PAPER, WHITE, TEAL, SKY = "F5F7FA", "FFFFFF", "0F8B8D", "E3F1F2"
THIS_YEAR, LAST_YEAR = "0F8B8D", "B7C4D1"
FONT = "Calibri"

# (label, low, high, strong colour, light colour)
BANDS = [
    ("At risk", 0, 34, "D14343", "F9DADA"),
    ("Weak", 35, 49, "E67A2E", "FDE6D4"),
    ("Fair", 50, 64, "D99A00", "FFF1C9"),
    ("Healthy", 65, 79, "4FA24A", "E3F3DC"),
    ("Strong", 80, 100, "1E8E5A", "D6F0E2"),
]
INFO_COLORS = ("6B7C8C", "EEF1F4")

FILL = lambda c: PatternFill("solid", fgColor=c)
HAIR = Side(style="thin", color=LINE)

FMT_INT = '#,##0;[Red]-#,##0;0'
FMT_DEC = '#,##0.00;[Red]-#,##0.00;0.00'
FMT_PCT = '0.0%;[Red]-0.0%;0.0%'
FMT_MULT = '0.00"x";[Red]-0.00"x";0.00"x"'

CHART_PALETTE = ["0F8B8D", "1B2A41", "E67A2E", "4FA24A", "D99A00", "7A6CB3", "B7C4D1"]


# ---------------------------------------------------------------------------
# Small helpers
# ---------------------------------------------------------------------------
def safe_text(value: Any, limit: int = 32000) -> str:
    """Remove characters Excel cannot store. Formula-looking text is stored as
    text by `Canvas.cell` (it never becomes a formula)."""
    if value is None or isinstance(value, (list, dict)):
        return ""
    return re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", " ", str(value)).strip()[:limit]


def numeric(value: Any) -> Optional[float]:
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)) and math.isfinite(value):
        return float(value)
    return None


def metric_value(value: Any) -> Optional[float]:
    if isinstance(value, dict):
        if value.get("available") is False:
            return None
        for key in ("value", "currentValue", "amount", "ratio"):
            if key in value:
                return numeric(value[key])
        return None
    return numeric(value)


def human(value: Any) -> str:
    text = re.sub(r"([a-z])([A-Z])", r"\1 \2", str(value or "").replace("_", " "))
    text = re.sub(r"\s+", " ", text).strip()
    return text[0].upper() + text[1:] if text else ""


def plainify(text: Any) -> str:
    """Keep AI prose as authored; financial units must be established upstream.

    Converting every 0.xxxxxx into a percentage is unsafe: 0.277778 may
    be a cash-ratio multiple rather than a percentage.
    """
    return re.sub(r"\s+", " ", safe_text(text))


def shorten(text: str, limit: int) -> str:
    text = (text or "").strip()
    if len(text) <= limit:
        return text
    cut = text[:limit]
    end = max(cut.rfind(". "), cut.rfind("! "), cut.rfind("? "))
    if end > limit * 0.55:
        return cut[: end + 1]
    return cut[: cut.rfind(" ")].rstrip(",;:") + "…"


def first_sentences(text: str, n: int = 2) -> str:
    parts = re.split(r"(?<=[.!?])\s+", (text or "").strip())
    return " ".join(parts[:n])


def rating_for(score: Optional[float]) -> Optional[str]:
    if score is None:
        return None
    for label, lo, hi, *_ in reversed(BANDS):
        if score >= lo:
            return label
    return BANDS[0][0]


def band_colors(rating: Optional[str]) -> tuple[str, str]:
    for label, _lo, _hi, strong, light in BANDS:
        if label == rating:
            return strong, light
    return INFO_COLORS


def pretty_date(value: Optional[str]) -> str:
    if not value:
        return "Unknown period"
    try:
        d = datetime.strptime(str(value)[:10], "%Y-%m-%d")
        return f"{d.strftime('%B')} {d.day}, {d.year}"
    except ValueError:
        return str(value)


def short_date(value: Optional[str]) -> str:
    """Compact column label: 'Dec 31, 1995'."""
    try:
        d = datetime.strptime(str(value)[:10], "%Y-%m-%d")
        return f"{d.strftime('%b')} {d.day}, {d.year}"
    except (ValueError, TypeError):
        return str(value or "—")


def period_label(value: Optional[str]) -> str:
    """Chart/legend label: the year when the period is a year end, else the date."""
    try:
        d = datetime.strptime(str(value)[:10], "%Y-%m-%d")
        return str(d.year) if (d.month, d.day) == (12, 31) else short_date(value)
    except (ValueError, TypeError):
        return str(value or "—")


def pct_change(new: Optional[float], old: Optional[float]) -> Optional[float]:
    if new is None or old is None or old <= 0 or new < 0:
        return None
    return (new - old) / old


# ---------------------------------------------------------------------------
# Reading the n8n report
# ---------------------------------------------------------------------------
def source_period(record: dict) -> str:
    rp = record.get("reportingPeriod") or {}
    return str(record.get("period") or (rp.get("endDate") if isinstance(rp, dict) else None) or record.get("year") or "")


def period_list(report: dict) -> list[str]:
    periods = set()
    for record in report.get("financialMetrics") or []:
        if isinstance(record, dict) and record.get("period"):
            periods.add(str(record["period"]))
    for record in report.get("consolidatedStatements") or []:
        if isinstance(record, dict):
            p = source_period(record)
            if p:
                periods.add(p)
    return sorted(periods, reverse=True)


def source_records(report: dict) -> list[dict]:
    """Flatten consolidatedStatements[].accounts into rows, keeping the order
    the accounts have in the original document."""
    result = []
    for statement in report.get("consolidatedStatements") or []:
        if not isinstance(statement, dict):
            continue
        kind = str(statement.get("statementType") or "other_financial_statement")
        period = source_period(statement)
        accounts = statement.get("accounts") or {}
        if isinstance(accounts, list):
            iterator = [(str(i), a) for i, a in enumerate(accounts)]
        elif isinstance(accounts, dict):
            iterator = list(accounts.items())
        else:
            continue
        for order, (key, account) in enumerate(iterator):
            if not isinstance(account, dict):
                continue
            amount = numeric(account.get("reportedValue"))
            if amount is None:
                amount = numeric(account.get("reported_value"))
            if amount is None:
                amount = numeric(account.get("value"))
            if amount is None:
                continue
            result.append({
                "kind": kind, "period": period, "key": str(account.get("key") or key),
                "label": safe_text(account.get("label") or human(key), 220),
                "section": str(account.get("section") or ""),
                "subsection": str(account.get("subsection") or ""),
                "value": amount, "order": order,
                "currency": statement.get("currency"), "unit": statement.get("unitMultiplier"),
                "classification": statement.get("cashFlowClassification"),
            })
    return result


ALIASES = {
    "total_assets": ["total_assets"],
    "total_liabilities": ["total_liabilities"],
    "total_equity": ["total_equity", "stockholders_equity", "shareholders_equity"],
    "current_assets": ["total_current_assets"],
    "current_liabilities": ["total_current_liabilities"],
    "cash": ["cash_and_cash_equivalents", "cash_and_bank_current_accounts", "cash_and_bank", "cash"],
    "revenue": ["total_revenue", "total_revenues", "sales_revenue", "net_revenue", "net_sales", "total_income",
                "revenues", "total_financial_income", "total_operating_revenue"],
    "net_income": ["net_income", "net_profit", "net_earnings", "net_income_loss", "net_profit_loss",
                   "excess_of_income_over_expenses", "excess_income_over_expenses",
                   "excess_of_revenue_over_expenses", "net_surplus", "net_surplus_deficit",
                   "surplus_for_the_year", "profit_for_the_year"],
    "gross_loans": ["loans_outstanding_gross", "gross_loans", "gross_loans_outstanding"],
    "loans_current": ["loans_outstanding_current", "current_loans"],
    "loans_past_due": ["loans_outstanding_past_due", "past_due_loans", "nonperforming_loans"],
    "loans_restructured": ["loans_outstanding_restructured", "restructured_loans"],
    "loan_loss_reserve": ["loan_loss_reserve", "loan_loss_allowance", "allowance_for_loan_losses"],
    "grants": ["grant_revenue_for_operations", "income_from_grants_for_credit_services", "grants", "grant_income",
               "grant_revenue"],
    "sources_total": ["total_sources_of_funds"],
    "uses_total": ["total_uses_of_funds"],
    "ocf": ["net_cash_from_operating_activities", "net_cash_provided_by_operating_activities",
            "operating_cash_flow", "cash_flow_from_operations"],
    "icf": ["net_cash_from_investing_activities", "net_cash_used_in_investing_activities",
            "net_cash_provided_by_investing_activities", "investing_cash_flow"],
    "fcf": ["net_cash_from_financing_activities", "net_cash_provided_by_financing_activities",
            "net_cash_used_in_financing_activities", "financing_cash_flow"],
}


class Model:
    """A friendlier view over the n8n report."""

    def __init__(self, report: dict):
        self.raw = report
        self.company = safe_text(report.get("company") or "Financial Analysis", 150)
        self.records = source_records(report)
        self.periods = period_list(report)
        self.p1 = self.periods[0] if self.periods else None
        self.p2 = self.periods[1] if len(self.periods) > 1 else None
        self.profile = report.get("scoringProfile") or report.get("profile") or ""
        self.is_lender = "lend" in str(self.profile).lower() or bool(self.first_val("gross_loans", self.p1) is not None)
        self.score = numeric(report.get("financialHealthScore"))
        self.coverage = numeric(report.get("dataCoveragePercent"))
        self.rating = report.get("scoreRating") or rating_for(self.score)
        # A high score on partial evidence is not an overall health verdict.
        self.rating_qualification = (" (partial evidence)" if self.coverage is not None and self.coverage < 100 else "")
        self.status = human(report.get("scoringStatus") or "unknown")
        self.errors = [safe_text(e if not isinstance(e, dict) else "; ".join(str(v) for v in e.values()), 600)
                       for e in (report.get("errors") or [])]
        self.warnings = [safe_text(w if not isinstance(w, dict) else "; ".join(str(v) for v in w.values()), 600)
                         for w in (report.get("warnings") or [])]

        # categories
        self.cats = []
        for c in report.get("scoreCategories") or []:
            if not isinstance(c, dict):
                continue
            w = numeric(c.get("effectiveWeight"))
            base = numeric(c.get("baseWeight"))
            self.cats.append({
                "name": human(c.get("name") or c.get("category") or "Category"),
                "score": numeric(c.get("score")),
                "weight": (w / 100 if w is not None and w > 1 else w),
                "base_weight": (base / 100 if base is not None and base > 1 else base),
                "availability": human(c.get("availability") or "Unknown"),
                "components": [p for p in (c.get("components") or []) if isinstance(p, dict)],
            })
        self.components = {}
        for c in self.cats:
            for p in c["components"]:
                if p.get("metricKey"):
                    self.components[p["metricKey"]] = p

        # metrics: key -> period -> raw
        self.metric_vals: dict[str, dict[str, Any]] = OrderedDict()
        self.metric_kind_of: dict[str, str] = {}
        for rec in report.get("financialMetrics") or []:
            if not isinstance(rec, dict):
                continue
            period = str(rec.get("period") or "")
            for key, raw in (rec.get("metrics") or {}).items():
                self.metric_vals.setdefault(key, {}).setdefault(period, raw)
                self.metric_kind_of.setdefault(key, str(rec.get("statementType") or ""))

        # currency / scale
        self.currency = next((r["currency"] for r in self.records if r.get("currency")), None)
        self.unit = next((r["unit"] for r in self.records if r.get("unit")), None)

        # analyst text
        a = report.get("financialAnalysis") or {}
        self.analysis = {
            "headline": plainify(a.get("headline")),
            "executive_summary": plainify(a.get("executive_summary")),
            "overall_assessment": plainify(a.get("overall_assessment")),
            "score_explanation": plainify(a.get("score_explanation")),
            "conclusion": plainify(a.get("conclusion")),
            "data_limitations": [plainify(x) for x in (a.get("data_limitations") or []) if plainify(x)],
        }
        for key in ("strengths", "weaknesses", "financial_trends", "risk_factors", "recommendations"):
            items = []
            for x in a.get(key) or []:
                if isinstance(x, dict) and x.get("title") and x.get("explanation"):
                    items.append({"title": plainify(x["title"]), "explanation": plainify(x["explanation"]),
                                  "priority": str(x.get("priority") or "").lower()})
            self.analysis[key] = items

    # -- lookups ----------------------------------------------------------
    def stmt_val(self, kind: str, key: str, period: Optional[str]) -> Optional[float]:
        if period is None:
            return None
        for r in self.records:
            if r["kind"] == kind and r["period"] == period and r["key"] == key:
                return r["value"]
        return None

    def first_val(self, name: str, period: Optional[str], kind: Optional[str] = None) -> Optional[float]:
        if period is None:
            return None
        for key in ALIASES.get(name, [name]):
            for r in self.records:
                if r["period"] == period and r["key"] == key and (kind is None or r["kind"] == kind):
                    return r["value"]
        return None

    def metric(self, key: str, period: Optional[str]) -> Optional[float]:
        if period is None:
            return None
        return metric_value((self.metric_vals.get(key) or {}).get(period))

    def amount(self, name: str, period: Optional[str]) -> Optional[float]:
        """Prefer the n8n metric, fall back to the statement account."""
        v = self.metric(name, period)
        return v if v is not None else self.first_val(name, period)

    def pair(self, name: str) -> tuple[Optional[float], Optional[float]]:
        return self.amount(name, self.p1), self.amount(name, self.p2)

    @property
    def units_note(self) -> str:
        if self.currency and self.unit:
            scale = {1: "", 1000: " thousands", 1000000: " millions"}.get(self.unit, f" x{self.unit}")
            return f"Amounts are in {self.currency}{scale}."
        if self.currency:
            return f"Amounts are in {self.currency}."
        return "The source documents do not say which currency or scale is used, so amounts are shown exactly as reported."

    def money(self, v: Optional[float]) -> str:
        if v is None:
            return "—"
        sym = {"USD": "$", "EUR": "€", "GBP": "£"}.get(str(self.currency or "").upper(), "")
        body = f"{abs(v):,.0f}" if abs(v) >= 100 or float(v).is_integer() else f"{abs(v):,.2f}"
        return f"{'-' if v < 0 else ''}{sym}{body}"


# ---------------------------------------------------------------------------
# Plain-English catalogues
# ---------------------------------------------------------------------------
CATEGORY_QUESTIONS = [
    ("capital", "Is there enough owner money to soak up losses?"),
    ("loan portfolio", "Are borrowers paying their loans back?"),
    ("loan", "Are borrowers paying their loans back?"),
    ("liquidity", "Can it pay its short-term bills?"),
    ("solvency", "Does it owe too much compared with what it owns?"),
    ("earning", "Is it making money?"),
    ("profit", "Is it making money and growing?"),
    ("cash", "Does the day-to-day business produce cash?"),
]


def category_question(name: str) -> str:
    n = name.lower()
    for needle, text in CATEGORY_QUESTIONS:
        if needle in n:
            return text
    return "One slice of the overall health score."


# key -> dict(group, title, kind, direction, meaning)
# kind: percent (stored as a decimal), multiple, amount
METRIC_INFO = OrderedDict([
    ("current_ratio", dict(group="bills", title="Current ratio", kind="multiple", direction=None,
        meaning="Compares what the business could turn into cash within a year with what it must pay within a year. "
                "Above 1.0x means it has more than enough on paper. It does not prove the cash is free to spend.")),
    ("cash_ratio", dict(group="bills", title="Cash ratio", kind="multiple", direction="higher",
        meaning="Actual cash and bank money for every 1 owed in the short term. A bigger number is a bigger cash cushion.")),
    ("working_capital", dict(group="bills", title="Working capital", kind="amount", direction="higher",
        meaning="Short-term assets minus short-term debts: the money left over after everything due within a year is paid.")),
    ("equity_to_assets", dict(group="debt", title="Owners' share of assets", kind="percent", direction="higher",
        meaning="The part of everything the business owns that really belongs to its owners rather than to lenders. "
                "A bigger share is a bigger safety cushion.")),
    ("liabilities_to_assets", dict(group="debt", title="Share of assets owed to others", kind="percent", direction="lower",
        meaning="The part of everything the business owns that is paid for with borrowed money or unpaid bills. "
                "Lower usually means less risk.")),
    ("liabilities_to_equity", dict(group="debt", title="Debt for every 1 of owner money", kind="multiple", direction="lower",
        meaning="How much the business owes for each 1 that the owners have put in. Lower usually means less risk.")),
    ("past_due_loans_to_gross_loans", dict(group="loans", title="Loans that are late", kind="percent", direction="lower",
        meaning="The share of all loans handed out where the borrower is behind on payments. Late loans are more likely "
                "never to be repaid, so lower is better. (Late does not automatically mean lost.)")),
    ("loan_loss_reserves_to_past_due_loans", dict(group="loans", title="Safety money vs. late loans", kind="percent",
        direction="higher",
        meaning="Money set aside to absorb bad loans, compared with the loans that are late. "
                "Higher means a bigger cushion if some of them are never repaid.")),
    ("loan_loss_reserves_to_gross_loans", dict(group="loans", title="Safety money vs. all loans", kind="percent",
        direction="higher",
        meaning="Money set aside for bad loans as a share of every loan the business has out.")),
    ("revenue", dict(group="profit", title="Money earned (revenue)", kind="amount", direction="higher",
        meaning="All the money the business earned from its activities during the year, before any costs are taken off.")),
    ("gross_profit", dict(group="profit", title="Gross profit", kind="amount", direction="higher",
        meaning="Money earned minus the direct cost of what was sold.")),
    ("operating_income", dict(group="profit", title="Profit from running the business", kind="amount", direction="higher",
        meaning="What is left after the everyday costs of running the business, before one-off items such as grants.")),
    ("net_income", dict(group="profit", title="Final profit or loss (net income)", kind="amount", direction="higher",
        meaning="The bottom line: everything earned minus every cost. Positive = profit (surplus); negative = loss.")),
    ("gross_margin", dict(group="profit", title="Gross margin", kind="percent", direction="higher",
        meaning="Out of every 100 earned, how much is left after the direct cost of the goods or services.")),
    ("operating_margin", dict(group="profit", title="Profit margin from operations", kind="percent", direction="higher",
        meaning="Out of every 100 earned, how much is left after running costs.")),
    ("net_margin", dict(group="profit", title="Final profit margin", kind="percent", direction="higher",
        meaning="Out of every 100 earned, how much ends up as final profit. 4.7% means 4.7 kept for every 100 earned.")),
    ("return_on_assets", dict(group="profit", title="Profit compared with what it owns", kind="percent", direction="higher",
        meaning="Final profit divided by total assets (using year-end figures). Shows how hard the business's assets work.")),
    ("return_on_equity", dict(group="profit", title="Profit compared with owners' money", kind="percent", direction="higher",
        meaning="Final profit divided by the owners' stake (year-end figures). Shows what the owners' money earned.")),
    ("operating_cash_flow", dict(group="cash", title="Cash from running the business", kind="amount", direction="higher",
        meaning="Cash produced by everyday operations, according to a standard cash flow statement.")),
    ("opening_cash", dict(group="cash", title="Cash at the start of the year", kind="amount", direction=None,
        meaning="Cash and bank money at the beginning of the period.")),
    ("closing_cash", dict(group="cash", title="Cash at the end of the year", kind="amount", direction="higher",
        meaning="Cash and bank money at the end of the period.")),
    ("net_cash_change", dict(group="cash", title="Change in cash during the year", kind="amount", direction="higher",
        meaning="How much the cash balance went up (+) or down (-) over the period.")),
])

GROUP_TITLES = OrderedDict([
    ("bills", ("Can it pay its bills?", "Short-term safety: can the business cover what it owes soon?")),
    ("debt", ("Does it rely too much on borrowed money?", "Long-term safety: how much of the business is owned vs. borrowed.")),
    ("loans", ("Are borrowers paying the loans back?", "Loan quality: for businesses that lend money.")),
    ("profit", ("Is it making money?", "Profit: money earned vs. money spent.")),
    ("cash", ("What happened to its cash?", "Cash: the money actually in the bank.")),
    ("other", ("Other measures", "Extra numbers found in the documents.")),
])

# first matching substring of the account key wins
ACCOUNT_HELP = [
    ("total_assets", "Everything the business owns that has value."),
    ("total_current_assets", "Cash, loans due back soon and other things that turn into cash within a year."),
    ("total_long_term_assets", "Things the business keeps for longer than a year."),
    ("total_non_current_assets", "Things the business keeps for longer than a year."),
    ("total_current_liabilities", "Bills and debts that must be paid within a year."),
    ("total_liabilities_and_equity", "Always equals total assets: what is owed plus what the owners hold."),
    ("total_liabilities", "Everything the business owes to others."),
    ("total_equity", "What belongs to the owners: assets minus liabilities."),
    ("cash", "Money in the bank and on hand."),
    ("interest_bearing_deposits", "Money placed in accounts that earn interest."),
    ("loans_outstanding_past_due", "Loans where the borrower is late with payments."),
    ("loans_outstanding_restructured", "Loans whose payment terms were changed because the borrower struggled."),
    ("loans_outstanding_gross", "All loans handed out that have not been fully repaid, before setting money aside."),
    ("loans_outstanding_current", "Loans where the borrower is paying on time."),
    ("loan_loss_reserve", "Money set aside for loans that may never be repaid. It is subtracted from loans."),
    ("net_loans", "Loans owed back to the business after subtracting the safety money set aside."),
    ("accumulated_depreciation", "Wear and tear already written off on buildings and equipment."),
    ("net_property", "Value of buildings and equipment after subtracting wear and tear."),
    ("long_term_investments", "Investments held for more than a year."),
    ("short_term_borrowings", "Borrowed money that must be repaid within a year."),
    ("long_term_debt", "Borrowed money that is repaid after more than a year."),
    ("client_savings", "Money that customers have deposited with the business."),
    ("loan_fund_capital", "Money the owners or donors put in to fund the loans."),
    ("retained", "Profits kept in the business from earlier years."),
    ("net_surplus", "This year's profit (a surplus is a profit for a non-profit)."),
    ("total_financial_income", "All money earned from lending and investing."),
    ("total_revenue", "All money earned before costs."),
    ("interest_on_debt", "What the business pays for the money it borrowed."),
    ("total_financial_costs", "What it costs the business to get the money it lends out."),
    ("gross_financial_margin", "Income left after paying the cost of borrowed money."),
    ("provision_for_loan_losses", "Money set aside this year to cover loans that may not be repaid."),
    ("net_financial_margin", "What is left after both borrowing costs and bad-loan costs."),
    ("total_operating_expenses", "The everyday costs of running the business."),
    ("depreciation", "Wear and tear on equipment, spread over its life. No cash leaves the bank."),
    ("net_income_from_operations", "Profit from the everyday business, before grants."),
    ("grant", "Gifts of money from donors or governments."),
    ("excess_of_income_over_expenses", "The final result for the year: money earned minus money spent."),
    ("net_income", "The final result for the year: money earned minus money spent."),
    ("total_sources_of_funds", "All the ways money came into the business this year."),
    ("total_uses_of_funds", "All the ways money was spent or tied up this year."),
    ("net_cash_flow", "How much the cash balance changed during the year."),
    ("cash_beginning", "Cash on the first day of the period."),
    ("cash_end", "Cash on the last day of the period."),
    ("increase_in_borrowed", "New money borrowed during the year."),
    ("increase_in_loans", "Extra loans handed out during the year."),
    ("loan_write_offs", "Loans officially given up on as lost."),
]


def account_help(key: str) -> str:
    k = key.lower()
    for needle, text in ACCOUNT_HELP:
        if needle in k:
            return text
    return ""


GLOSSARY = [
    ("Assets", "Everything the business owns that has value: cash, loans owed to it, buildings, equipment.", "total_assets"),
    ("Liabilities", "Everything the business owes to other people: loans it must repay, bills it must pay.", "total_liabilities"),
    ("Equity", "What is left for the owners if all debts were paid: assets minus liabilities.", "total_equity"),
    ("Revenue (income)", "All the money the business earns from its work before any costs are taken off.", "revenue"),
    ("Expenses (costs)", "Money the business spends to keep running: wages, rent, interest on debt and so on.", None),
    ("Net income / profit / surplus", "What is left after ALL costs. Positive is a profit (called a surplus for non-profits); negative is a loss.", "net_income"),
    ("Margin", "Profit as a share of money earned. A 5% margin means the business keeps 5 out of every 100 it earns.", "net_margin"),
    ("Liquidity", "How easily the business can pay what it owes soon. Think of it as how much cash it can lay hands on quickly.", None),
    ("Current ratio", "Short-term assets divided by short-term debts. Above 1.0x means it can cover its short-term bills on paper.", "current_ratio"),
    ("Cash ratio", "Cash divided by short-term debts. Stricter than the current ratio because it only counts actual cash.", "cash_ratio"),
    ("Working capital", "Short-term assets minus short-term debts.", "working_capital"),
    ("Leverage", "How much of the business is paid for with borrowed money. More borrowing means more risk when times are bad.", None),
    ("Past-due loan", "A loan where the borrower is behind on payments. Late is not the same as lost, but it is a warning sign.", "past_due_loans_to_gross_loans"),
    ("Loan-loss reserve", "Money set aside to cover loans that may never be paid back. It works like a safety cushion.", "loan_loss_reserves_to_past_due_loans"),
    ("Provision for loan losses", "The amount added to that safety cushion this year. It counts as a cost.", None),
    ("Depreciation", "Spreading the cost of equipment over the years it is used. It lowers profit but no cash leaves the bank.", None),
    ("Sources and uses of funds", "A simple list of where money came from during the year and where it went. It is not the same as an operating cash flow statement.", None),
    ("Percentage points", "The gap between two percentages. Going from 36.8% to 38.9% is a rise of 2.1 percentage points.", None),
    ("Financial health score", "A 0–100 score built from several health checks, each with a weight. It is a rule-of-thumb model, not an official credit rating.", None),
    ("Data coverage", "How many of the score's checks could be done with the numbers that were supplied. Checks that cannot be done are left out, not counted as zero.", None),
]


# ---------------------------------------------------------------------------
# Worksheet canvas: tiny layout helpers
# ---------------------------------------------------------------------------
def F(size=10, bold=False, color=INK, italic=False, underline=None):
    return Font(name=FONT, size=size, bold=bold, color=color, italic=italic, underline=underline)


def A(h="left", v="center", wrap=True, indent=0):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap, indent=indent)


class Canvas:
    def __init__(self, ws, widths: dict[int, float], default_height: float = 18):
        self.ws = ws
        self.widths = widths
        ws.sheet_view.showGridLines = False
        ws.sheet_view.zoomScale = 100
        ws.sheet_format.defaultRowHeight = default_height
        ws.sheet_format.customHeight = True
        for col, w in widths.items():
            ws.column_dimensions[get_column_letter(col)].width = w
        ws.page_setup.orientation = "landscape"
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.print_options.horizontalCentered = True
        ws.oddFooter.center.text = "Automated, unaudited analysis  |  Check the source documents before making decisions"
        ws.oddFooter.right.text = "Page &P of &N"
        self.last_col = max(widths)

    # geometry ------------------------------------------------------------
    def span_width(self, c1: int, c2: int) -> float:
        return sum(self.widths.get(c, 9) for c in range(c1, c2 + 1))

    def lines_needed(self, text: str, c1: int, c2: int, size: float = 10, indent: int = 0) -> int:
        usable = max(8.0, self.span_width(c1, c2) - 1.2 * indent - 1.0)
        chars_per_line = usable * 1.12 * (10.0 / size)
        return sum(max(1, math.ceil(len(p) / chars_per_line)) for p in str(text).split("\n"))

    def height(self, row: int, h: float):
        self.ws.row_dimensions[row].height = h

    def grow(self, row: int, h: float):
        cur = self.ws.row_dimensions[row].height or 18
        self.ws.row_dimensions[row].height = max(cur, h)

    # cells ---------------------------------------------------------------
    def cell(self, r, c, value=None, *, font=None, fill=None, align=None, fmt=None, border=None):
        cell = self.ws.cell(r, c)
        if value is not None:
            if isinstance(value, str):
                cell.value = value
                cell.data_type = "s"      # never interpret text as a formula
            else:
                cell.value = value
        if font:
            cell.font = font
        if fill:
            cell.fill = FILL(fill)
        if align:
            cell.alignment = align
        if fmt:
            cell.number_format = fmt
        if border:
            cell.border = border
        return cell

    def paint(self, r1, c1, r2, c2, fill=None, border=None):
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                if fill:
                    self.ws.cell(r, c).fill = FILL(fill)
                if border:
                    self.ws.cell(r, c).border = border

    def block(self, r1, c1, r2, c2, value=None, *, font=None, fill=None, align=None, fmt=None):
        if fill:
            self.paint(r1, c1, r2, c2, fill=fill)
        if (r1, c1) != (r2, c2):
            self.ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
        return self.cell(r1, c1, value, font=font, align=align, fmt=fmt)

    def text(self, r, c1, c2, text, *, size=10, bold=False, color=INK, italic=False, fill=None, h="left",
             v="center", indent=0, min_h=18, pad=6, rows=1, shrink_row_ok=True):
        """Merged wrapped text in one row (or a fixed block of rows). Row height is
        estimated from the amount of text."""
        text = str(text or "")
        self.block(r, c1, r + rows - 1, c2, text, font=F(size, bold, color, italic), fill=fill,
                   align=A(h, v, True, indent))
        lines = self.lines_needed(text, c1, c2, size, indent)
        need = lines * size * 1.42 + pad
        if rows == 1:
            self.height(r, max(min_h, need) if shrink_row_ok else max(self.ws.row_dimensions[r].height or 18, need))
        else:
            per = max(min_h, need / rows)
            for rr in range(r, r + rows):
                self.grow(rr, per)
        return lines

    def rich(self, r, c1, c2, parts, *, fill=None, h="left", v="center", indent=0, rows=1, min_h=18):
        """parts: list of (text, size, bold, color). Rich text inside one merged cell."""
        blocks = []
        for text, size, bold, color in parts:
            blocks.append(TextBlock(InlineFont(rFont=FONT, sz=size, b=bold, color=color), text))
        self.block(r, c1, r + rows - 1, c2, None, fill=fill, align=A(h, v, True, indent))
        self.ws.cell(r, c1).value = CellRichText(*blocks)
        total = "".join(p[0] for p in parts)
        size = max(p[1] for p in parts)
        lines = self.lines_needed(total, c1, c2, size, indent)
        need = lines * size * 1.42 + 6
        if rows == 1:
            self.grow(r, max(min_h, need))
        else:
            for rr in range(r, r + rows):
                self.grow(rr, max(min_h, need / rows))
        return lines

    def card(self, r1, c1, r2, c2, *, fill=None, edge=LINE, left=None, left_w="thick"):
        """Light border around a rectangle, optionally with a coloured left stripe."""
        if fill:
            self.paint(r1, c1, r2, c2, fill=fill)
        thin = Side(style="thin", color=edge)
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                cell = self.ws.cell(r, c)
                old = cell.border
                cell.border = Border(
                    left=(Side(style=left_w, color=left) if (left and c == c1) else (thin if c == c1 else old.left)),
                    right=thin if c == c2 else old.right,
                    top=thin if r == r1 else old.top,
                    bottom=thin if r == r2 else old.bottom,
                )

    def title_band(self, company: str, subtitle: str, last_col: Optional[int] = None):
        last = last_col or self.last_col
        self.paint(1, 1, 3, last, fill=NAVY)
        self.block(1, 2, 1, last - 1, company, font=F(22, True, WHITE), fill=NAVY, align=A("left", "center", False))
        self.height(1, 36)
        self.block(2, 2, 2, last - 1, subtitle, font=F(10.5, False, "C8D3DF"), fill=NAVY, align=A("left", "center", False))
        self.height(2, 22)
        self.paint(3, 1, 3, last, fill=TEAL)
        self.height(3, 4)

    def nav(self, row: int, current: str, c1: int, c2: int):
        """Hyperlinked jump bar."""
        names = ["Overview", "Health Checks", "Statements", "Insights", "Money Words", "About This Report"]
        span = c2 - c1 + 1
        per = max(1, span // len(names))
        col = c1
        for i, name in enumerate(names):
            end = c2 if i == len(names) - 1 else min(c2, col + per - 1)
            cur = name == current
            cell = self.block(row, col, row, end, name, font=F(9.5, cur, NAVY if cur else TEAL, False, None if cur else "single"),
                              align=A("center", "center", False), fill=SKY if cur else None)
            if not cur:
                cell.hyperlink = f"#'{name}'!A1"
            col = end + 1
            if col > c2:
                break
        self.height(row, 22)

    def section(self, row: int, c1: int, c2: int, title: str, subtitle: str = ""):
        self.block(row, c1, row, c2, title.upper(), font=F(11, True, NAVY), align=A("left", "center", False))
        self.height(row, 24)
        for c in range(c1, c2 + 1):
            self.ws.cell(row, c).border = Border(bottom=Side(style="medium", color=TEAL))
        if subtitle:
            self.text(row + 1, c1, c2, subtitle, size=9.5, italic=True, color=MUTED, min_h=16, pad=2)
            return row + 2
        return row + 1


# ---------------------------------------------------------------------------
# Charts
# ---------------------------------------------------------------------------
def _char_props(size=900, bold=False, color=MUTED):
    return CharacterProperties(sz=size, b=bold, solidFill=color, latin=None)


def set_title(chart, text, size=1200):
    cp = CharacterProperties(sz=size, b=True, solidFill=NAVY)
    chart.title = Title(
        tx=Text(rich=RichText(p=[Paragraph(pPr=ParagraphProperties(defRPr=cp), r=[RegularTextRun(rPr=cp, t=text)])])),
        overlay=False)


def style_axes(chart, *, horizontal=False, num_fmt="#,##0", show_val_axis=True):
    tx = RichText(p=[Paragraph(pPr=ParagraphProperties(defRPr=_char_props(900)), endParaRPr=CharacterProperties())])
    for ax in (chart.x_axis, chart.y_axis):
        ax.delete = False
        ax.txPr = tx
    chart.y_axis.number_format = num_fmt
    chart.y_axis.majorGridlines.spPr = GraphicalProperties(ln=LineProperties(solidFill="E6EBF0", w=6350))
    chart.y_axis.spPr = GraphicalProperties(ln=LineProperties(noFill=True))
    chart.x_axis.spPr = GraphicalProperties(ln=LineProperties(solidFill="B9C4CF", w=9525))
    chart.x_axis.majorTickMark = "none"
    chart.y_axis.majorTickMark = "none"
    if not show_val_axis:
        chart.y_axis.delete = True


def labels(series, *, percent=False, pos=None):
    d = DataLabelList()
    d.showVal = not percent
    d.showPercent = percent
    d.showCatName = False
    d.showSerName = False
    d.showLegendKey = False
    d.showLeaderLines = False
    if pos:
        d.position = pos
    d.txPr = RichText(p=[Paragraph(pPr=ParagraphProperties(defRPr=_char_props(900, True, INK)), endParaRPr=CharacterProperties())])
    series.dLbls = d


def legend_bottom(chart, show=True):
    if not show:
        chart.legend = None
        return
    chart.legend.position = "b"
    chart.legend.txPr = RichText(p=[Paragraph(pPr=ParagraphProperties(defRPr=_char_props(900)), endParaRPr=CharacterProperties())])


class ChartData:
    """Holds chart source tables on a hidden sheet so charts never depend on
    the (merged, styled) display cells."""

    def __init__(self, ws):
        self.ws = ws
        self.row = 1

    def table(self, headers: list[str], rows: list[list[Any]], fmt: str = "#,##0") -> tuple[int, int]:
        top = self.row
        for j, h in enumerate(headers, 1):
            self.ws.cell(top, j, safe_text(h))
        for i, row in enumerate(rows, 1):
            for j, v in enumerate(row, 1):
                c = self.ws.cell(top + i, j, safe_text(v) if isinstance(v, str) else v)
                if j > 1:
                    c.number_format = fmt
        self.row = top + len(rows) + 3
        return top, top + len(rows)


def column_chart(cd: ChartData, title, headers, rows, *, colors, fmt="#,##0", stacked=False, horizontal=False,
                 point_colors=None, width=14.4, height=8.8, show_legend=True, val_axis=False, gap=70):
    top, last = cd.table(headers, rows, fmt)
    ch = BarChart()
    ch.type = "bar" if horizontal else "col"
    ch.grouping = "stacked" if stacked else "clustered"
    if stacked:
        ch.overlap = 100
    ch.gapWidth = gap
    for j in range(2, len(headers) + 1):
        ch.add_data(Reference(cd.ws, min_col=j, min_row=top, max_row=last), titles_from_data=True)
    ch.set_categories(Reference(cd.ws, min_col=1, min_row=top + 1, max_row=last))
    for idx, (s, color) in enumerate(zip(ch.series, colors)):
        s.graphicalProperties = GraphicalProperties(solidFill=color)
        s.graphicalProperties.line.noFill = True
        s.invertIfNegative = False
        if not (stacked and idx >= 2):          # thin top segments only add clutter
            labels(s, pos=None if stacked else "outEnd")
    if point_colors and len(ch.series) == 1:
        for i, color in enumerate(point_colors):
            pt = DataPoint(idx=i)
            pt.graphicalProperties = GraphicalProperties(solidFill=color)
            pt.graphicalProperties.line.noFill = True
            ch.series[0].dPt.append(pt)
    set_title(ch, title)
    style_axes(ch, horizontal=horizontal, num_fmt=fmt, show_val_axis=val_axis)
    if horizontal:
        ch.x_axis.scaling.orientation = "maxMin"
        ch.x_axis.tickLblPos = "low"
    legend_bottom(ch, show_legend and len(headers) > 2)
    ch.width, ch.height = width, height
    ch.graphical_properties = GraphicalProperties(solidFill=WHITE, ln=LineProperties(solidFill=LINE, w=9525))
    return ch


def pie_chart(cd: ChartData, title, rows, *, width=14.4, height=8.8):
    top, last = cd.table(["Slice", "Amount"], rows, "#,##0")
    ch = PieChart()
    ch.add_data(Reference(cd.ws, min_col=2, min_row=top, max_row=last), titles_from_data=True)
    ch.set_categories(Reference(cd.ws, min_col=1, min_row=top + 1, max_row=last))
    s = ch.series[0]
    for i in range(len(rows)):
        pt = DataPoint(idx=i)
        pt.graphicalProperties = GraphicalProperties(solidFill=CHART_PALETTE[i % len(CHART_PALETTE)])
        pt.graphicalProperties.line = LineProperties(solidFill=WHITE, w=12700)
        s.dPt.append(pt)
    labels(s, percent=True, pos="outEnd")
    set_title(ch, title)
    ch.legend.position = "r"
    ch.legend.txPr = RichText(p=[Paragraph(pPr=ParagraphProperties(defRPr=_char_props(900)), endParaRPr=CharacterProperties())])
    ch.firstSliceAng = 0
    ch.width, ch.height = width, height
    ch.graphical_properties = GraphicalProperties(solidFill=WHITE, ln=LineProperties(solidFill=LINE, w=9525))
    return ch


def gauge_chart(cd: ChartData, score: float, color: str, *, width=11.6, height=7.6):
    top, last = cd.table(["Part", "Value"], [["Score", score], ["Remaining", 100 - score], ["Hidden half", 100]], "0.0")
    ch = DoughnutChart()
    ch.add_data(Reference(cd.ws, min_col=2, min_row=top, max_row=last), titles_from_data=True)
    ch.set_categories(Reference(cd.ws, min_col=1, min_row=top + 1, max_row=last))
    ch.holeSize = 74
    ch.firstSliceAng = 270
    s = ch.series[0]
    for i, fillc in enumerate([color, "E3E8EE", None]):
        pt = DataPoint(idx=i)
        pt.graphicalProperties = GraphicalProperties(solidFill=fillc) if fillc else GraphicalProperties(noFill=True)
        pt.graphicalProperties.line = LineProperties(noFill=True)
        s.dPt.append(pt)
    ch.legend = None
    ch.title = None
    ch.width, ch.height = width, height
    ch.graphical_properties = GraphicalProperties(noFill=True, ln=LineProperties(noFill=True))
    ch.plot_area.graphicalProperties = GraphicalProperties(noFill=True, ln=LineProperties(noFill=True))
    return ch


# ---------------------------------------------------------------------------
# Shared analytics used by several sheets
# ---------------------------------------------------------------------------
GOOD, BAD, NEUTRAL = "1E8E5A", "D14343", "6B7C8C"


def delta_info(kind: str, direction: Optional[str], new: Optional[float], old: Optional[float],
               old_label: str = "") -> tuple[str, str]:
    """Return (text, colour) describing a change in plain words."""
    if new is None or old is None:
        return "—", NEUTRAL
    d = new - old
    if kind == "percent":
        d100 = d * 100
        if abs(d100) < 0.05:
            return "● no change", NEUTRAL
        text = f"{'▲' if d > 0 else '▼'} {abs(d100):.1f} pts"
    elif kind == "multiple":
        if abs(d) < 0.005:
            return "● no change", NEUTRAL
        text = f"{'▲' if d > 0 else '▼'} {abs(d):.2f}x"
    else:
        if abs(d) < 1e-9:
            return "● no change", NEUTRAL
        pc = pct_change(new, old)
        if pc is not None:
            text = f"{'▲' if d > 0 else '▼'} {abs(pc) * 100:.1f}%"
        elif old < 0 <= new:
            text = f"▲ from a loss of {abs(old):,.0f}"
        elif new < 0 <= old:
            text = f"▼ into a loss of {abs(new):,.0f}"
        else:
            text = f"{'▲' if d > 0 else '▼'} {abs(d):,.0f}"
    if old_label:
        text += f" vs {old_label}"
    if direction is None:
        return text, NEUTRAL
    good = (d > 0) == (direction == "higher")
    return text, GOOD if good else BAD


def fmt_by_kind(kind: str, v: Optional[float], m: Optional[Model] = None, short: bool = False) -> str:
    if v is None:
        return "—"
    if kind == "percent":
        return f"{v * 100:.0f}%" if short and abs(v * 100 - round(v * 100)) < 0.05 else f"{v * 100:.1f}%"
    if kind == "multiple":
        return f"{v:.2f}x"
    return m.money(v) if m else f"{v:,.0f}"


def metric_kind(key: str, raw: Any) -> str:
    info = METRIC_INFO.get(key)
    if info:
        return info["kind"]
    unit = str(raw.get("unit") or "").lower() if isinstance(raw, dict) else ""
    k = key.lower()
    if unit in ("decimal", "proportion", "fraction", "percent", "percentage", "%") or k.endswith(("_margin", "_growth", "_ratio_pct")):
        return "percent"
    if unit in ("ratio", "x", "multiple", "times") or k.endswith("_ratio"):
        return "multiple"
    return "amount"


def display_value(m: Model, key: str, period: Optional[str]) -> Optional[float]:
    raw = (m.metric_vals.get(key) or {}).get(period) if period else None
    v = metric_value(raw)
    if v is None:
        return None
    unit = str(raw.get("unit") or "").lower() if isinstance(raw, dict) else ""
    if unit in ("percent", "percentage", "%") and key not in METRIC_INFO:
        return v / 100
    return v


def verdict_for(m: Model, key: str, value: Optional[float]) -> Optional[tuple[str, float]]:
    comp = m.components.get(key)
    if not comp or value is None:
        return None
    worry, goal = numeric(comp.get("worry")), numeric(comp.get("goal"))
    if worry is None or goal is None or worry == goal:
        return None
    score = max(0.0, min(100.0, 100 * (value - worry) / (goal - worry)))
    return rating_for(score), score


def goal_text(m: Model, key: str, kind: str) -> str:
    comp = m.components.get(key)
    if not comp:
        return "No goal set in this model"
    worry, goal = numeric(comp.get("worry")), numeric(comp.get("goal"))
    if worry is None or goal is None:
        return "No goal set in this model"
    g, w = fmt_by_kind(kind, goal, m, True), fmt_by_kind(kind, worry, m, True)
    if (comp.get("direction") or ("higher" if goal > worry else "lower")) == "lower":
        return f"Goal: {g} or less\nWorry zone: {w} or more"
    return f"Goal: {g} or more\nWorry zone: {w} or less"


RULE_BASED_NEXT_STEPS = [
    ("loan", "Look closely at the loans that are behind on payments: how late they are and where they come from. "
             "Check that the money set aside for bad loans is big enough to cover them."),
    ("liquid", "Compare cash on hand with the bills due in the next 12 months, and plan how any shortfall would be covered."),
    ("earning", "Find the costs that grew faster than income and decide which ones can be trimmed or which prices can rise."),
    ("profit", "Find the costs that grew faster than income and decide which ones can be trimmed or which prices can rise."),
    ("capital", "Keep the owners' share of the business growing as it grows, so there is a cushion if a bad year comes."),
    ("solven", "Review how much is owed compared with what is owned, and avoid new borrowing until the balance improves."),
    ("cash", "Ask for a standard cash flow statement so the business's day-to-day cash generation can be measured properly."),
]


def rule_based_step(cat_name: str) -> str:
    n = cat_name.lower()
    for needle, text in RULE_BASED_NEXT_STEPS:
        if needle in n:
            return text
    return f"Review the numbers behind “{cat_name}” and find out what is driving the score."


def scored_cats(m: Model) -> list[dict]:
    return sorted([c for c in m.cats if c["score"] is not None], key=lambda c: c["score"])


def build_story(m: Model) -> dict:
    cats = scored_cats(m)
    weak = cats[0] if cats else None
    strong = cats[-1] if cats else None
    a = m.analysis

    headline = a["headline"]
    if not headline:
        if m.score is not None and strong and weak and strong is not weak:
            headline = (f"{m.company} scores {m.score:.1f} out of 100 ({(m.rating or '').lower()}): "
                        f"{strong['name'].lower()} is a strength, {weak['name'].lower()} needs attention.")
        elif m.score is not None:
            headline = (f"{m.company} scores {m.score:.1f} out of 100 ({(m.rating or '').lower()}), "
                        f"but only {len(cats)} of {len(m.cats)} areas could be scored.") if len(m.cats) > len(cats) else f"{m.company} scores {m.score:.1f} out of 100 ({(m.rating or '').lower()})."
        else:
            headline = f"A financial health score could not be calculated for {m.company}."

    def pick(items, fallback):
        if items:
            x = items[0]
            return shorten(f"{x['title']}: {first_sentences(x['explanation'], 1)}", 165)
        return fallback

    good = pick(a["strengths"],
                f"{strong['name']} is the strongest area, scoring {strong['score']:.0f} out of 100." if strong
                else "Not enough scored data to name a strength.")
    watch = pick(a["weaknesses"],
                 f"{weak['name']} is the weakest area, scoring {weak['score']:.0f} out of 100." if weak
                 else "Not enough scored data to name a weak spot.")
    if len(cats) == 1 and not a["weaknesses"]:
        watch = (f"Only {len(cats)} of {len(m.cats)} areas could be scored, so this is a partial picture. "
                 "The areas left out may tell a different story.")
    recs = a["recommendations"]
    if recs:
        order = {"high": 0, "medium": 1, "low": 2}
        recs = sorted(recs, key=lambda x: order.get(x["priority"], 1))
        nxt = shorten(f"{recs[0]['title']}: {first_sentences(recs[0]['explanation'], 1)}", 165)
    elif len(cats) == 1:
        nxt = "Supply the missing statements (balance sheet and cash flow) so more areas can be scored."
    elif weak:
        nxt = shorten(rule_based_step(weak["name"]), 165)
    else:
        nxt = "Provide complete statements so a full score can be calculated."
    return {"headline": headline, "good": good, "watch": watch, "next": nxt}


# ----- income statement decomposition ---------------------------------------
SUBTOTAL_PREFIXES = ("total_", "net_", "gross_", "excess_", "operating_income", "profit_", "surplus")


def is_subtotal_key(key: str) -> bool:
    return key.lower().startswith(SUBTOTAL_PREFIXES)


def income_rows(m: Model, period: Optional[str]) -> list[dict]:
    rows = [r for r in m.records if r["kind"] == "income_statement" and r["period"] == period]
    return sorted(rows, key=lambda r: r["order"])


def income_breakdown(m: Model, period: Optional[str]) -> dict:
    """Split income-statement lines into income lines, cost lines and other income."""
    rows = income_rows(m, period)
    rev_total_idx = None
    for i, r in enumerate(rows):
        k = r["key"].lower()
        if k.startswith("total_") and any(w in k for w in ("income", "revenue", "sales")):
            rev_total_idx = i
            break
    income, costs, other = [], [], []
    if rev_total_idx is not None:
        income = [r for r in rows[:rev_total_idx] if not is_subtotal_key(r["key"]) and r["value"] > 0]
        for r in rows[rev_total_idx + 1:]:
            k, lab = r["key"].lower(), r["label"].lower()
            if is_subtotal_key(k):
                if re.search(r"grant|donation|other_income", k):
                    other.append(r)
                continue
            if re.search(r"grant|donation|other_income|gain|subsid", k + " " + lab):
                other.append(r)
            elif r["value"] > 0:
                costs.append(r)
    return {"income": income, "costs": costs, "other": other}


def top_slices(rows: list[tuple[str, float]], n: int = 5, other_label: str = "Everything else") -> list[list[Any]]:
    rows = sorted([x for x in rows if x[1] > 0], key=lambda x: -x[1])
    total = sum(v for _, v in rows) or 1.0
    head = [x for x in rows[:n] if x[1] / total >= 0.03]
    rest = sum(v for _, v in rows) - sum(v for _, v in head)
    if len(head) == len(rows):
        return [[a, b] for a, b in head]
    if len(rows) - len(head) == 1:             # a single leftover keeps its own name
        return [[a, b] for a, b in rows]
    return [[a, b] for a, b in head] + [[other_label, rest]]


def clean_label(label: str) -> str:
    s = re.sub(r"^\(|\)$", "", label.strip())
    return s[:34] + "…" if len(s) > 35 else s


# ----- KPI tiles --------------------------------------------------------------
def kpi_candidates(m: Model) -> list[dict]:
    old = period_label(m.p2) if m.p2 else ""
    out = []

    def add(label, kind, direction, new, old_v, note=""):
        if new is None:
            return
        out.append(dict(label=label, kind=kind, direction=direction, new=new, old=old_v, note=note))

    rev1, rev2 = m.pair("revenue")
    add("Money earned (revenue)", "amount", "higher", rev1, rev2)
    ni1, ni2 = m.pair("net_income")
    add("Final profit" if (ni1 or 0) >= 0 else "Final loss", "amount", "higher", ni1, ni2)
    a1, a2 = m.pair("total_assets")
    add("Everything it owns (assets)", "amount", None, a1, a2)
    e1, e2 = display_value(m, "equity_to_assets", m.p1), display_value(m, "equity_to_assets", m.p2)
    add("Owners' share of assets", "percent", "higher", e1, e2)
    p1_, p2_ = display_value(m, "past_due_loans_to_gross_loans", m.p1), display_value(m, "past_due_loans_to_gross_loans", m.p2)
    add("Loans that are late", "percent", "lower", p1_, p2_)
    c1, c2 = m.pair("cash")
    add("Cash and bank money", "amount", "higher", c1, c2)
    nm1, nm2 = display_value(m, "net_margin", m.p1), display_value(m, "net_margin", m.p2)
    add("Profit kept from each 100 earned", "percent", "higher", nm1, nm2)
    l1, l2 = m.pair("total_liabilities")
    add("What it owes (liabilities)", "amount", None, l1, l2)
    return out[:8]


# ---------------------------------------------------------------------------
# Charts for the Overview sheet
# ---------------------------------------------------------------------------
def build_charts(m: Model, cd: ChartData) -> list[tuple[Any, str]]:
    specs: list[tuple[Any, str]] = []
    y1 = period_label(m.p1)
    y2 = period_label(m.p2) if m.p2 else None
    two = bool(m.p2)
    heads = lambda first: [first, y2, y1] if two else [first, y1]
    colors = [LAST_YEAR, THIS_YEAR] if two else [THIS_YEAR]

    def row(label, new, old):
        return [label, old, new] if two else [label, new]

    # 1. own vs owe -----------------------------------------------------------
    a, l, e = m.pair("total_assets"), m.pair("total_liabilities"), m.pair("total_equity")
    rows = [row(lbl, p[0], p[1]) for lbl, p in
            [("Assets (what it owns)", a), ("Liabilities (what it owes)", l), ("Equity (owners’ share)", e)]
            if p[0] is not None]
    if len(rows) >= 2:
        specs.append((column_chart(cd, "What it owns vs. what it owes", heads("Item"), rows, colors=colors),
                      "Assets = liabilities + equity. The bigger the owners’ share of the total, the bigger the safety cushion."))

    # 2. income, costs, profit ------------------------------------------------
    def other_income(p):
        if p is None:
            return 0.0
        br = income_breakdown(m, p)
        tot = sum(r["value"] for r in br["other"] if r["value"] > 0)
        return tot or (m.first_val("grants", p) or 0.0)

    def costs(p):
        rev, ni = m.amount("revenue", p), m.amount("net_income", p)
        if rev is None or ni is None:
            return None
        return rev + other_income(p) - ni

    rev, ni = m.pair("revenue"), m.pair("net_income")
    if rev[0] is not None and ni[0] is not None:
        rows = [row("Money earned", rev[0], rev[1])]
        oi1, oi2 = other_income(m.p1), other_income(m.p2)
        if oi1 or oi2:
            rows.append(row("Grants & other income", oi1, oi2 if two else None))
        rows.append(row("Money spent", costs(m.p1), costs(m.p2) if two else None))
        rows.append(row("Final profit / loss", ni[0], ni[1]))
        specs.append((column_chart(cd, "Money in, money out, what was left", heads("Item"), rows, colors=colors),
                      "Profit = money earned (plus any grants) minus money spent. A negative bar means a loss."))

    # 3. where the money went -------------------------------------------------
    br = income_breakdown(m, m.p1)
    if br["costs"]:
        slices = top_slices([(clean_label(r["label"]), r["value"]) for r in br["costs"]], 5, "All other costs")
        ni1 = m.amount("net_income", m.p1)
        if ni1 and ni1 > 0:
            slices.append(["Left over as profit", ni1])
            title = f"Where the money went in {y1}"
        else:
            title = f"Where the money was spent in {y1}"
        if len(slices) >= 2:
            specs.append((pie_chart(cd, title, slices),
                          "Each slice is a share of the money that came in. A thin ‘left over’ slice means thin profit margins."))

    # 4. where the money comes from ------------------------------------------
    if len(br["income"]) >= 2:
        slices = top_slices([(clean_label(r["label"]), r["value"]) for r in br["income"]], 3, "All other income")
        specs.append((pie_chart(cd, f"Where the income came from in {y1}", slices),
                      "Relying on one big source is simpler to manage but riskier if that source dries up."))

    # 5. loan book (lenders) or short-term safety ----------------------------
    cur1, pd1, rs1 = (m.first_val(k, m.p1) for k in ("loans_current", "loans_past_due", "loans_restructured"))
    if cur1 is not None and pd1 is not None:
        rows = []
        for p in ([m.p2] if two else []) + [m.p1]:
            c_, d_, r_ = (m.first_val(k, p) for k in ("loans_current", "loans_past_due", "loans_restructured"))
            if c_ is not None:
                rows.append([period_label(p), c_, d_ or 0, r_ or 0])
        specs.append((column_chart(cd, "How healthy is the loan book?",
                                   ["Year", "Paid on time", "Late (past due)", "Terms changed (restructured)"], rows,
                                   colors=["4FA24A", "D14343", "D99A00"], stacked=True),
                      "For a lender the red part matters most. If it shrinks over time, borrowers are paying more reliably."))
    elif not m.is_lender:
        ca, cl = m.pair("current_assets"), m.pair("current_liabilities")
        rows = [row(lbl, p[0], p[1]) for lbl, p in [("Short-term assets", ca), ("Short-term debts", cl)] if p[0] is not None]
        if len(rows) == 2:
            specs.append((column_chart(cd, "Can it pay the bills due soon?", heads("Item"), rows, colors=colors),
                          "If the first bar is taller than the second, the business can cover its short-term debts on paper."))

    # 6. where the cash came from / went -------------------------------------
    cf_periods = sorted({r["period"] for r in m.records if r["kind"] == "cash_flow_statement"}, reverse=True)
    if cf_periods:
        cp = cf_periods[0]
        cf = sorted([r for r in m.records if r["kind"] == "cash_flow_statement" and r["period"] == cp], key=lambda r: r["order"])
        keys = [r["key"] for r in cf]
        s_key, u_key = ALIASES["sources_total"][0], ALIASES["uses_total"][0]
        items = []
        if s_key in keys and u_key in keys:
            i_s, i_u = keys.index(s_key), keys.index(u_key)
            for r in cf[:i_s]:
                if not is_subtotal_key(r["key"]) and r["value"] != 0:
                    items.append((clean_label(r["label"]), abs(r["value"])))
            for r in cf[i_s + 1:i_u]:
                if not is_subtotal_key(r["key"]) and r["value"] != 0:
                    items.append((clean_label(r["label"]), -abs(r["value"])))
            title = f"Where the cash came from and went ({period_label(cp)})"
            cap = "Teal bars are money coming in; orange bars are money going out. Depreciation and loan-loss provisions are paper costs, so they are added back."
        else:
            for lbl, name in [("Running the business", "ocf"), ("Investing", "icf"), ("Financing", "fcf")]:
                v = m.first_val(name, cp)
                if v is not None:
                    items.append((lbl, v))
            title = f"Cash movements in {period_label(cp)}"
            cap = "Teal bars added cash; orange bars used cash."
        if len(items) >= 2:
            specs.append((column_chart(cd, title, ["Item", "Amount"], [list(x) for x in items], colors=[THIS_YEAR],
                                       horizontal=True, point_colors=[THIS_YEAR if v >= 0 else "E67A2E" for _, v in items],
                                       show_legend=False, height=max(8.8, 1.2 + 0.62 * len(items)), gap=45), cap))
    return specs[:6]


# ---------------------------------------------------------------------------
# Sheet 1: Overview
# ---------------------------------------------------------------------------
def build_overview(wb: Workbook, m: Model, cd: ChartData):
    ws = wb["Overview"]
    ws.sheet_properties.tabColor = NAVY
    widths = {1: 1.8, **{c: 12.6 for c in range(2, 14)}, 14: 1.8}
    cv = Canvas(ws, widths)
    sub = f"Financial health report  ·  {pretty_date(m.p1)}"
    if m.p2:
        sub += f"  ·  compared with {pretty_date(m.p2)}"
    cv.title_band(m.company, sub)
    cv.nav(4, "Overview", 2, 13)
    cv.height(5, 8)
    r = 6

    # attention banner --------------------------------------------------------
    if m.errors:
        cv.text(r, 2, 13, "NEEDS REVIEW BEFORE USE:  " + shorten("; ".join(m.errors[:3]), 330), size=10, bold=True,
                color="7A4B00", fill="FFF1C9", indent=1, min_h=26)
        cv.card(r, 2, r, 13, edge="E0A100")
        r += 2

    if m.coverage is not None and m.coverage < 70 and m.score is not None:
        cv.text(r, 2, 13, f"LIMITED DATA:  this score rests on only {m.coverage:.0f}% of the checks the model normally runs, "
                          "so treat it with extra caution. Supply more statements for a fuller picture.",
                size=10, bold=True, color="7A4B00", fill="FFF1C9", indent=1, min_h=28)
        cv.card(r, 2, r, 13, edge="E0A100")
        r += 2

    story = build_story(m)
    strong, light = band_colors(m.rating)

    # ---- hero: gauge card (left) + story card (right) ---------------------
    H0 = r
    for rr in range(H0, H0 + 10):
        cv.height(rr, 24)
    cv.height(H0 + 10, 40)
    cv.card(H0, 2, H0 + 10, 6, fill=WHITE)
    cv.card(H0, 8, H0 + 10, 13, fill=PAPER)

    cv.text(H0, 2, 6, "FINANCIAL HEALTH SCORE", size=9.5, bold=True, color=MUTED, h="center", min_h=24)
    if m.score is not None:
        ws.add_chart(gauge_chart(cd, m.score, strong), f"B{H0 + 1}")
        cv.block(H0 + 3, 3, H0 + 5, 5, round(m.score, 1), font=F(32, True, strong),
                 align=A("center", "center", False), fmt="0.0")
        cv.text(H0 + 6, 3, 5, ((m.rating or "") + ("*" if m.rating_qualification else "")).upper(), size=16, bold=True, color=strong, h="center", min_h=24)
        cov = f"out of 100  ·  data coverage {m.coverage:.0f}%  ·  partial evidence" if m.coverage is not None and m.coverage < 100 else (f"out of 100  ·  data coverage {m.coverage:.0f}%" if m.coverage is not None else "out of 100")
        cv.text(H0 + 7, 2, 6, cov, size=9.5, color=MUTED, h="center", min_h=22)
        cv.text(H0 + 8, 2, 6, "A rule-of-thumb score, not an official credit rating.", size=9, italic=True, color=MUTED,
                h="center", min_h=22)
    else:
        cv.text(H0 + 3, 2, 6, "Score unavailable", size=18, bold=True, color=MUTED, h="center", rows=3)
        cv.text(H0 + 6, 2, 6, "There was not enough valid data to calculate a score. See About This Report.", size=10,
                color=MUTED, h="center", rows=2, indent=1)
    # band legend
    for i, (label, lo, hi, s_, l_) in enumerate(BANDS):
        active = label == m.rating
        cv.block(H0 + 10, 2 + i, H0 + 10, 2 + i, f"{label.upper()}\n{lo}–{hi}",
                 font=F(8.5, active, WHITE if active else MUTED), fill=s_ if active else l_,
                 align=A("center", "center", True))
        ws.cell(H0 + 10, 2 + i).border = Border(left=Side(style="thin", color=WHITE), right=Side(style="thin", color=WHITE))

    # story card
    cv.text(H0, 8, 13, "THE 30-SECOND STORY", size=9.5, bold=True, color=MUTED, indent=1, min_h=24, fill=PAPER)
    cv.text(H0 + 1, 8, 13, story["headline"], size=14, bold=True, color=NAVY, rows=2, indent=1, fill=PAPER)
    for k, (tag, key, color) in enumerate([("THE GOOD", "good", "1E8E5A"), ("WATCH OUT", "watch", "E67A2E"),
                                           ("NEXT STEP", "next", TEAL)]):
        top = H0 + 3 + 2 * k
        cv.block(top, 8, top + 1, 8, tag, font=F(8.5, True, WHITE), fill=color, align=A("center", "center", True))
        cv.text(top, 9, 13, story[key], size=10, rows=2, indent=1, fill=PAPER)
    cv.text(H0 + 9, 8, 13, "Scroll down for the report card, key numbers and charts.", size=9, italic=True, color=MUTED,
            indent=1, fill=PAPER, min_h=24)
    r = H0 + 12

    # ---- report card -------------------------------------------------------
    r = cv.section(r, 2, 13, "Report card: how each area scores",
                   "Each area answers one simple question. Longer, greener bars are better.")
    cv.height(r - 1, 18)
    unscored = False
    for c in m.cats:
        h = 40
        cv.height(r, h)
        parts = [(c["name"], 11, True, NAVY), ("\n" + category_question(c["name"]), 9, False, MUTED)]
        if c["weight"]:
            parts.append((f"  ·  counts for {c['weight'] * 100:.0f}% of the score", 9, False, MUTED))
        cv.rich(r, 2, 6, parts, indent=1, min_h=h)
        s = c["score"]
        if s is None:
            unscored = True
            cv.rich(r, 7, 10, [("Not assessed — no data", 10, False, MUTED)])
            cv.block(r, 11, r, 11, "—", font=F(14, True, MUTED), align=A("center"))
            cv.block(r, 12, r, 13, "NOT SCORED", font=F(9.5, True, WHITE), fill="9AA7B4", align=A("center"))
        else:
            rt = rating_for(s)
            sc, _ = band_colors(rt)
            n = max(0, min(20, round(s / 5)))
            cv.rich(r, 7, 10, [("█" * n, 11, False, sc), ("█" * (20 - n), 11, False, "E3E8EE")])
            cv.block(r, 11, r, 11, round(s, 1), font=F(14, True, sc), align=A("center"), fmt="0")
            cv.block(r, 12, r, 13, rt.upper(), font=F(9.5, True, WHITE), fill=sc, align=A("center"))
        for cc in range(2, 14):
            ws.cell(r, cc).border = Border(bottom=HAIR)
        r += 1
    if not m.cats:
        cv.text(r, 2, 13, "No score categories were supplied.", size=10, color=MUTED, indent=1)
        r += 1
    if unscored:
        cv.text(r, 2, 13, "NOT SCORED means the supplied documents did not contain what was needed for that area. "
                          "It is left out of the total, not counted as zero.", size=9, italic=True, color=MUTED, indent=1, min_h=26)
        r += 1
    r += 1

    # ---- key numbers --------------------------------------------------------
    kp = kpi_candidates(m)
    if kp:
        subtitle = (f"{period_label(m.p1)} compared with {period_label(m.p2)}. Green means moving in a healthy direction, "
                    f"red means moving the wrong way, grey means neither is better.") if m.p2 else \
                   f"Figures for {period_label(m.p1)}. No earlier year was supplied, so there is nothing to compare with."
        ws.row_breaks.append(Break(id=r - 1))
        r = cv.section(r, 2, 13, "Key numbers", subtitle)
        cv.height(r - 1, 30)
        for i, k in enumerate(kp):
            col = 2 + 3 * (i % 4)
            top = r + 5 * (i // 4)
            if i % 4 == 0:
                cv.height(top, 20)
                cv.height(top + 1, 22)
                cv.height(top + 2, 22)
                cv.height(top + 3, 22)
                cv.height(top + 4, 10)
            text, color = delta_info(k["kind"], k["direction"], k["new"], k["old"], period_label(m.p2) if m.p2 else "")
            if not m.p2:
                text, color = "", NEUTRAL
            cv.paint(top, col, top + 3, col + 2, fill=PAPER)
            cv.text(top, col, col + 2, k["label"].upper(), size=8.5, bold=True, color=MUTED, indent=1, fill=PAPER, min_h=20)
            val = fmt_by_kind(k["kind"], k["new"], m)
            cv.block(top + 1, col, top + 2, col + 2, val, font=F(22, True, NAVY), fill=PAPER, align=A("left", "center", False, 1))
            cv.text(top + 3, col, col + 2, text, size=9.5, bold=True, color=color, indent=1, fill=PAPER, min_h=22)
            cv.card(top, col, top + 3, col + 2, edge=LINE, left=color if text else LINE)
        r += 5 * ((len(kp) + 3) // 4) + 1

    # ---- charts -------------------------------------------------------------
    specs = build_charts(m, cd)
    if specs:
        r = cv.section(r, 2, 13, "The numbers in pictures", "Plain charts built from the statements. Captions say how to read each one.")
        cv.height(r - 1, 18)
        for i in range(0, len(specs), 2):
            pair = specs[i:i + 2]
            if i == 2:
                ws.row_breaks.append(Break(id=r - 1))
            rows_needed = 14
            for rr in range(r, r + rows_needed):
                cv.height(rr, 18)
            for j, (chart, caption) in enumerate(pair):
                col = 2 if j == 0 else 8
                ws.add_chart(chart, f"{get_column_letter(col)}{r}")
                cv.text(r + rows_needed, col, col + 5, caption, size=9, italic=True, color=MUTED, indent=1, min_h=32)
            cv.height(r + rows_needed + 1, 10)
            r += rows_needed + 2

    cv.text(r, 2, 13, "This workbook was produced automatically from the supplied documents and an AI-written commentary. "
                      "It has not been audited. " + m.units_note, size=9, italic=True, color=MUTED, indent=1, min_h=30)
    ws.print_area = f"A1:N{r + 1}"
    ws.sheet_view.zoomScale = 100


# ---------------------------------------------------------------------------
# Sheet 2: Health Checks
# ---------------------------------------------------------------------------
def build_health(wb: Workbook, m: Model, cd: ChartData):
    ws = wb["Health Checks"]
    ws.sheet_properties.tabColor = TEAL
    widths = {1: 1.8, 2: 30, 3: 14, 4: 14, 5: 20, 6: 14, 7: 26, 8: 62, 9: 1.8}
    cv = Canvas(ws, widths)
    cv.title_band(m.company, "Health checks  ·  one measure per row, explained in plain words", 9)
    cv.nav(4, "Health Checks", 2, 8)
    cv.height(5, 8)
    y1, y2 = period_label(m.p1), (period_label(m.p2) if m.p2 else "")
    cv.text(6, 2, 8, "How to read this page:  “Latest” is the most recent year and “Prior” is the year before. The arrow shows the "
                     "direction of change (green = healthier, red = less healthy, grey = neither is clearly better). The Verdict "
                     "uses the same goals as the health score. Those goals are rules of thumb, not industry standards.",
            size=10, color=INK, fill=SKY, indent=1, min_h=48)
    r = 8

    groups: dict[str, list] = OrderedDict((g, []) for g in GROUP_TITLES)
    for key, by_period in m.metric_vals.items():
        if key.startswith("total_"):
            continue
        raw = next((v for v in by_period.values() if v is not None), None)
        info = METRIC_INFO.get(key)
        kind = metric_kind(key, raw)
        v1, v2 = display_value(m, key, m.p1), display_value(m, key, m.p2)
        if v1 is None and v2 is None:
            continue
        group = info["group"] if info else "other"
        title = info["title"] if info else human(key)
        meaning = info["meaning"] if info else (safe_text(raw.get("description"), 300) if isinstance(raw, dict) else "") or \
            "A measure calculated from the supplied statements."
        direction = info["direction"] if info else None
        groups[group].append(dict(key=key, title=title, kind=kind, v1=v1, v2=v2, meaning=meaning, direction=direction))

    pct_for_chart = []
    for gkey, (gtitle, gsub) in GROUP_TITLES.items():
        items = groups[gkey]
        if not items:
            continue
        r = cv.section(r, 2, 8, gtitle, gsub)
        headers = ["Measure", f"Latest ({y1})", f"Prior ({y2})" if y2 else "Prior", "Change", "Verdict", "Goal vs. worry zone", "What it means"]
        for j, h in enumerate(headers):
            cv.cell(r, 2 + j, h, font=F(9.5, True, NAVY), fill="EAF0F5", align=A("center" if 0 < j < 6 else "left", "center", True, 1 if j in (0, 6) else 0))
        cv.height(r, 28)
        r += 1
        for it in items:
            fmt = {"percent": FMT_PCT, "multiple": FMT_MULT}.get(it["kind"], FMT_INT if all(
                (x is None or abs(x - round(x)) < 1e-9) for x in (it["v1"], it["v2"])) else FMT_DEC)
            cv.cell(r, 2, it["title"], font=F(10.5, True, NAVY), align=A("left", "center", True, 1))
            cv.cell(r, 3, it["v1"], font=F(11, True, INK), align=A("center"), fmt=fmt)
            cv.cell(r, 4, it["v2"] if it["v2"] is not None else "—", font=F(10.5, False, MUTED), align=A("center"), fmt=fmt)
            text, color = delta_info(it["kind"], it["direction"], it["v1"], it["v2"])
            cv.cell(r, 5, text, font=F(10, True, color), align=A("center"))
            verdict = verdict_for(m, it["key"], it["v1"])
            if verdict:
                strong, _ = band_colors(verdict[0])
                cv.cell(r, 6, verdict[0].upper(), font=F(9.5, True, WHITE), fill=strong, align=A("center"))
            else:
                cv.cell(r, 6, "NO GOAL SET", font=F(8.5, True, "52606D"), fill=INFO_COLORS[1], align=A("center"))
            cv.cell(r, 7, goal_text(m, it["key"], it["kind"]) if verdict else "—", font=F(9.5, False, MUTED), align=A("left", "center", True, 1))
            cv.text(r, 8, 8, it["meaning"], size=9.5, color=INK, indent=1, min_h=36)
            cv.height(r, max(ws.row_dimensions[r].height or 36, 36))
            for cc in range(2, 9):
                ws.cell(r, cc).border = Border(bottom=HAIR)
            if it["kind"] == "percent" and it["v1"] is not None and it["key"] != "revenue_growth":
                pct_for_chart.append(it)
            r += 1
        r += 1

    # chart: key percentages
    pct = [p for p in pct_for_chart if p["v2"] is not None][:6]
    if len(pct) >= 2:
        r = cv.section(r, 2, 8, "Percentages side by side", "Same measures as above, drawn so changes are easy to spot.")
        rows = [[p["title"], p["v2"], p["v1"]] for p in pct]
        ch = column_chart(cd, "Key percentages: prior vs. latest", ["Measure", y2, y1], rows,
                          colors=[LAST_YEAR, THIS_YEAR], fmt="0.0%", width=24, height=9)
        for rr in range(r, r + 17):
            cv.height(rr, 18)
        ws.add_chart(ch, f"B{r}")
        r += 18
    if r == 8:
        cv.text(r, 2, 8, "No ratio measures were supplied for this report.", size=10, color=MUTED, indent=1)
        r += 1
    cv.text(r + 1, 2, 8, "A rise or fall is not automatically good or bad: it depends on the kind of business. "
                         "Two years of data show a change, not a long-term trend.", size=9, italic=True, color=MUTED, indent=1, min_h=30)
    ws.print_area = f"A1:I{r + 2}"


# ---------------------------------------------------------------------------
# Sheet 3: Statements
# ---------------------------------------------------------------------------
STATEMENT_ORDER = ["balance_sheet", "income_statement", "cash_flow_statement"]
STATEMENT_TITLES = {
    "balance_sheet": ("Balance sheet: what the business owns and owes",
                      "A snapshot on one day. Assets (what it owns) always equal liabilities (what it owes) plus equity (the owners’ share)."),
    "income_statement": ("Income statement: money in, money out",
                         "Covers a whole year. Money earned, minus money spent, leaves the profit (or loss)."),
    "cash_flow_statement": ("Cash statement: where the cash came from and went",
                            "Follows the money actually moving in and out during the year."),
}
BS_SECTION_NAMES = {"assets": "What the business owns (assets)", "liabilities": "What it owes (liabilities)",
                    "equity": "The owners’ share (equity)", "both": "Balance check"}


def bs_section(raw: str) -> Optional[str]:
    s = (raw or "").lower()
    if "asset" in s:
        return "assets"
    if "liabilit" in s and "equity" in s:
        return "both"
    if "liabilit" in s:
        return "liabilities"
    if "equity" in s or "capital" in s:
        return "equity"
    return None


def ordered_keys(members: list[dict], p1: Optional[str], p2: Optional[str]) -> list[str]:
    first = [r["key"] for r in sorted([r for r in members if r["period"] == p1], key=lambda r: r["order"])]
    seen = set(first)
    if p2:
        prev = None
        for r in sorted([r for r in members if r["period"] == p2], key=lambda r: r["order"]):
            if r["key"] in seen:
                prev = r["key"]
                continue
            pos = first.index(prev) + 1 if prev in first else 0
            first.insert(pos, r["key"])
            seen.add(r["key"])
            prev = r["key"]
    return first


def build_statements(wb: Workbook, m: Model):
    ws = wb["Statements"]
    ws.sheet_properties.tabColor = "2F6DB5"
    widths = {1: 1.8, 2: 44, 3: 15, 4: 15, 5: 22, 6: 14, 7: 62, 8: 1.8}
    cv = Canvas(ws, widths)
    cv.title_band(m.company, "Financial statements  ·  the original numbers, with a plain-English note beside each line", 8)
    cv.nav(4, "Statements", 2, 7)
    cv.height(5, 8)
    cv.text(6, 2, 7, "The amounts here are exactly as reported in the source documents. " + m.units_note +
            "  “Share of total” shows each line as a percentage of total assets (balance sheet) or of money earned (income statement).",
            size=10, fill=SKY, indent=1, min_h=40)
    r = 8
    kinds = sorted({x["kind"] for x in m.records}, key=lambda k: STATEMENT_ORDER.index(k) if k in STATEMENT_ORDER else 99)
    if not kinds:
        cv.text(r, 2, 7, "No statement lines were found in the data passed to the workbook.", size=10, color=MUTED, indent=1)
    for kind in kinds:
        members = [x for x in m.records if x["kind"] == kind]
        periods = sorted({x["period"] for x in members if x["period"]}, reverse=True)
        p1 = periods[0] if periods else None
        p2 = periods[1] if len(periods) > 1 else None
        title, sub = STATEMENT_TITLES.get(kind, (human(kind), "Supplementary schedule found in the documents."))
        r = cv.section(r, 2, 7, title, sub)
        if kind == "cash_flow_statement" and any(x.get("classification") == "sources_and_uses_of_funds" for x in members):
            cv.text(r, 2, 7, "This is a “sources and uses of funds” list. It is NOT a standard operating cash flow statement, so it is "
                             "not used to judge day-to-day cash generation.", size=9.5, bold=True, color="7A4B00", fill="FFF1C9", indent=1, min_h=30)
            r += 1
        heads = ["Line item", short_date(p1) if p1 else "Latest", short_date(p2) if p2 else "No prior year", "Change", "Share of total", "What it means"]
        for j, h in enumerate(heads):
            cv.cell(r, 2 + j, h, font=F(9.5, True, NAVY), fill="EAF0F5", align=A("center" if 0 < j < 5 else "left", "center", True, 1 if j in (0, 5) else 0))
        cv.height(r, 26)
        r += 1

        keys = ordered_keys(members, p1, p2)
        lookup = {(x["key"], x["period"]): x for x in members}
        base_key = {"balance_sheet": "total_assets", "income_statement": None}.get(kind)
        base = (lookup.get((base_key, p1)) or {}).get("value") if base_key else None
        if kind == "income_statement":
            base = m.amount("revenue", p1)

        # section boundaries
        s_idx = u_idx = rev_idx = None
        if kind == "cash_flow_statement":
            if ALIASES["sources_total"][0] in keys:
                s_idx = keys.index(ALIASES["sources_total"][0])
            if ALIASES["uses_total"][0] in keys:
                u_idx = keys.index(ALIASES["uses_total"][0])
        if kind == "income_statement":
            for i, k in enumerate(keys):
                if k.lower().startswith("total_") and any(w in k.lower() for w in ("income", "revenue", "sales")):
                    rev_idx = i
                    break
        current_header = None
        numfmt = FMT_INT if all(abs(x["value"] - round(x["value"])) < 1e-9 for x in members) else FMT_DEC
        for i, key in enumerate(keys):
            rec1, rec2 = lookup.get((key, p1)), lookup.get((key, p2)) if p2 else None
            rec = rec1 or rec2
            v1, v2 = (rec1 or {}).get("value"), (rec2 or {}).get("value")
            if v1 is None and v2 is None:
                continue
            header = None
            if kind == "balance_sheet":
                sec = bs_section(rec["section"])
                header = BS_SECTION_NAMES.get(sec) if sec else (human(rec["section"]) or None)
            elif kind == "income_statement" and rev_idx is not None:
                header = "Money coming in" if i <= rev_idx else "Money going out, and what was left"
            elif kind == "cash_flow_statement" and s_idx is not None and u_idx is not None:
                header = "Where the cash came from" if i <= s_idx else ("Where the cash went" if i <= u_idx else "Cash balance")
            elif kind == "cash_flow_statement":
                header = human(rec["section"]) or None
            if header and header != current_header:
                cv.block(r, 2, r, 7, header.upper(), font=F(9.5, True, TEAL), fill=SKY, align=A("left", "center", False, 1))
                cv.height(r, 22)
                r += 1
                current_header = header
            went = kind == "cash_flow_statement" and s_idx is not None and u_idx is not None and s_idx < i <= u_idx
            if went:
                v1 = abs(v1) if v1 is not None else None
                v2 = abs(v2) if v2 is not None else None
            klow = key.lower()
            total = klow.startswith("total_") or (rec["label"].isupper() and "total" in rec["label"].lower())
            sub_ = (not total) and (klow.startswith(("net_", "gross_", "excess_")) or (rec["label"].isupper() and len(rec["label"]) > 4))
            bold = total or sub_
            fill = "EAF0F5" if total else (PAPER if sub_ else None)
            label = rec["label"].strip()
            cv.cell(r, 2, label if bold else label, font=F(10, bold, NAVY if bold else INK), fill=fill, align=A("left", "center", True, 1 if bold else 2))
            cv.cell(r, 3, v1, font=F(10, bold, INK), fill=fill, align=A("right"), fmt=numfmt)
            cv.cell(r, 4, v2, font=F(10, bold, MUTED if not bold else INK), fill=fill, align=A("right"), fmt=numfmt)
            if v1 is not None and v2 is not None:
                text, _ = delta_info("amount", None, v1, v2)
            else:
                text = "—"
            cv.cell(r, 5, text, font=F(9.5, bold, "3F5266"), fill=fill, align=A("center", "center", True))
            share = (v1 / base) if (base and v1 is not None and base != 0 and kind != "cash_flow_statement") else None
            cv.cell(r, 6, share if share is not None else "—", font=F(9.5, False, MUTED), fill=fill, align=A("center"), fmt=FMT_PCT)
            cv.cell(r, 7, account_help(key), font=F(9.5, False, MUTED), fill=fill, align=A("left", "center", True, 1))
            help_lines = cv.lines_needed(account_help(key), 7, 7, 9.5, 1)
            lab_lines = cv.lines_needed(label, 2, 2, 10, 2)
            cv.height(r, max(21, max(help_lines, lab_lines) * 13.6 + 6))
            for cc in range(2, 8):
                ws.cell(r, cc).border = Border(bottom=HAIR, top=Side(style="thin", color=NAVY) if total else None)
            r += 1
        r += 2
    ws.print_area = f"A1:H{r}"


# ---------------------------------------------------------------------------
# Sheet 4: Insights
# ---------------------------------------------------------------------------
PRIORITY_STYLE = {"high": ("D14343", "HIGH PRIORITY"), "medium": ("D99A00", "MEDIUM"), "low": ("6B7C8C", "LOW")}


def finding_cards(cv: Canvas, r: int, items: list[dict], deep: str, light: str, with_priority=False) -> int:
    ws = cv.ws
    for i, it in enumerate(items, 1):
        title_end = 8 if with_priority else 9
        cv.text(r, 2, title_end, f"{i}.  {it['title']}", size=11, bold=True, color=deep, fill=light, indent=1, min_h=26)
        if with_priority:
            colr, lab = PRIORITY_STYLE.get(it.get("priority") or "medium", PRIORITY_STYLE["medium"])
            cv.block(r, 9, r, 9, lab, font=F(8.5, True, WHITE), fill=colr, align=A("center"))
        cv.text(r + 1, 2, 9, it["explanation"], size=10.5, color=INK, indent=1, min_h=24, pad=6)
        cv.paint(r, 1, r + 1, 1, fill=deep)
        for rr in (r, r + 1):
            for cc in range(2, 10):
                c = ws.cell(rr, cc)
                c.border = Border(bottom=HAIR if rr == r + 1 else None, right=HAIR if cc == 9 else None)
        r += 2
        cv.height(r, 8)
        r += 1
    return r


def build_insights(wb: Workbook, m: Model):
    ws = wb["Insights"]
    ws.sheet_properties.tabColor = "7A6CB3"
    widths = {1: 1.8, **{c: 13 for c in range(2, 10)}, 10: 1.8}
    cv = Canvas(ws, widths)
    cv.title_band(m.company, "Insights  ·  what the numbers mean and what to do about them", 10)
    cv.nav(4, "Insights", 2, 9)
    cv.height(5, 8)
    a = m.analysis
    story = build_story(m)
    cv.text(6, 2, 9, "Written by an AI assistant using only the numbers in this report. It can still make mistakes, so check anything "
                     "important against the Statements sheet.", size=9.5, italic=True, color=MUTED, indent=1, min_h=30)
    r = 8
    cv.text(r, 2, 9, story["headline"], size=16, bold=True, color=NAVY, fill=SKY, indent=1, min_h=48, pad=18)
    cv.card(r, 2, r, 9, edge=TEAL, left=TEAL)
    r += 2

    def prose(r, title, body):
        if not body:
            return r
        r = cv.section(r, 2, 9, title)
        cv.text(r, 2, 9, body, size=10.5, indent=1, min_h=26, pad=6)
        return r + 2

    r = prose(r, "In short", a["executive_summary"])
    r = prose(r, "The big picture", a["overall_assessment"])

    sections = [
        ("The good news", a["strengths"], "1E8E5A", "DDF2E6", False),
        ("Things to watch", a["weaknesses"], "C2611A", "FDE6D4", False),
        ("What changed since last year", a["financial_trends"], "2F6DB5", "E3EEF9", False),
        ("Risks to keep in mind", a["risk_factors"], "D14343", "F9DADA", False),
    ]
    for title, items, deep, light, pr in sections:
        if items:
            r = cv.section(r, 2, 9, title)
            r += 0
            r = finding_cards(cv, r, items, deep, light, pr)

    recs = a["recommendations"]
    order = {"high": 0, "medium": 1, "low": 2}
    if recs:
        recs = sorted(recs, key=lambda x: order.get(x["priority"], 1))
        r = cv.section(r, 2, 9, "What to do next", "Practical steps tied to the numbers in this report.")
        r = finding_cards(cv, r, recs, TEAL, SKY, True)
    elif scored_cats(m):
        r = cv.section(r, 2, 9, "What to do next", "No AI suggestions were returned, so these are rule-based next steps for the weakest areas.")
        items = [{"title": f"Look into {c['name'].lower()}", "explanation": rule_based_step(c["name"]),
                  "priority": "high" if i == 0 else "medium"} for i, c in enumerate(scored_cats(m)[:2])]
        r = finding_cards(cv, r, items, TEAL, SKY, True)

    r = prose(r, "How the score was built", a["score_explanation"])
    r = prose(r, "Bottom line", a["conclusion"])
    if a["data_limitations"]:
        r = cv.section(r, 2, 9, "What this report cannot tell you")
        for i, t in enumerate(a["data_limitations"], 1):
            cv.text(r, 2, 9, f"•  {t}", size=10, color=MUTED, indent=1, min_h=22, pad=8)
            r += 1
    ws.print_area = f"A1:J{r + 1}"


# ---------------------------------------------------------------------------
# Sheet 5: Money Words
# ---------------------------------------------------------------------------
BAND_MEANING = {
    "Strong": "Healthy on almost every check.",
    "Healthy": "Mostly good, with a few things to keep an eye on.",
    "Fair": "A mixed picture: some strengths and some real concerns.",
    "Weak": "Several checks are poor. Needs attention.",
    "At risk": "Most checks are poor. Look into it urgently.",
}


def build_glossary(wb: Workbook, m: Model):
    ws = wb["Money Words"]
    ws.sheet_properties.tabColor = "D99A00"
    widths = {1: 1.8, 2: 30, 3: 78, 4: 26, 5: 1.8}
    cv = Canvas(ws, widths)
    cv.title_band(m.company, "Money words  ·  every finance term in this workbook, in everyday language", 5)
    cv.nav(4, "Money Words", 2, 4)
    cv.height(5, 8)
    r = cv.section(6, 2, 4, "What the score colours mean")
    for label, lo, hi, strong, light in reversed(BANDS):
        cv.cell(r, 2, f"{label.upper()}   {lo}–{hi}", font=F(10, True, WHITE), fill=strong, align=A("left", "center", False, 1))
        cv.text(r, 3, 4, BAND_MEANING[label], size=10, fill=light, indent=1, min_h=24)
        r += 1
    cv.text(r, 2, 4, "These words only label ranges of this report’s own scoring model. They are not an official credit rating.",
            size=9, italic=True, color=MUTED, indent=1, min_h=24)
    r += 2
    r = cv.section(r, 2, 4, "Glossary")
    for j, h in enumerate(["Term", "What it means", "In this report"]):
        cv.cell(r, 2 + j, h, font=F(9.5, True, NAVY), fill="EAF0F5", align=A("left", "center", False, 1))
    cv.height(r, 24)
    r += 1
    for term, meaning, key in GLOSSARY:
        example = ""
        if key:
            info = METRIC_INFO.get(key)
            kind = info["kind"] if info else "amount"
            v = display_value(m, key, m.p1) if key in m.metric_vals else m.amount(key, m.p1)
            if v is not None:
                example = f"{fmt_by_kind(kind, v, m)}  ({period_label(m.p1)})"
        cv.cell(r, 2, term, font=F(10.5, True, NAVY), align=A("left", "center", True, 1))
        cv.text(r, 3, 3, meaning, size=10, indent=1, min_h=24)
        cv.cell(r, 4, example, font=F(10, True, TEAL), align=A("left", "center", True, 1))
        for cc in range(2, 5):
            ws.cell(r, cc).border = Border(bottom=HAIR)
        r += 1
    ws.print_area = f"A1:E{r}"


# ---------------------------------------------------------------------------
# Sheet 6: About This Report
# ---------------------------------------------------------------------------
def build_about(wb: Workbook, m: Model):
    ws = wb["About This Report"]
    ws.sheet_properties.tabColor = "6B7C8C"
    widths = {1: 1.8, 2: 44, 3: 17, 4: 17, 5: 17, 6: 15, 7: 13, 8: 17, 9: 1.8}
    cv = Canvas(ws, widths)
    cv.title_band(m.company, "About this report  ·  how the score was built and how far to trust it", 9)
    cv.nav(4, "About This Report", 2, 8)
    cv.height(5, 8)
    r = cv.section(6, 2, 8, "The basics")
    for label, val in [
        ("Health score", f"{m.score:.1f} out of 100 ({m.rating})" if m.score is not None else "Not calculated"),
        ("Data coverage", f"{m.coverage:.1f}% of the score’s checks could be done" if m.coverage is not None else "Unknown"),
        ("Status", m.status),
        ("Type of scoring used", human(m.profile) or "Unknown"),
        ("Years covered", ", ".join(pretty_date(p) for p in m.periods) or "Unknown"),
        ("Currency and scale", m.units_note),
    ]:
        cv.cell(r, 2, label, font=F(10, True, NAVY), align=A("left", "center", True, 1), border=Border(bottom=HAIR))
        cv.text(r, 3, 8, val, size=10, indent=1, min_h=22)
        for cc in range(2, 9):
            ws.cell(r, cc).border = Border(bottom=HAIR)
        r += 1
    r += 1

    # score breakdown --------------------------------------------------------
    r = cv.section(r, 2, 8, "How the score adds up",
                   "Each area is scored 0–100 from one or more checks. The areas are then blended using their weights. "
                   "Areas that could not be scored are left out, not counted as zero.")
    heads = ["Area / check", "Value measured", "Worry level\n(earns 0)", "Goal level\n(earns 100)", "Points\n(0–100)", "Weight", "Points added\nto total"]
    for j, h in enumerate(heads):
        cv.cell(r, 2 + j, h, font=F(9.5, True, NAVY), fill="EAF0F5", align=A("left" if j == 0 else "center", "center", True, 1 if j == 0 else 0))
    cv.height(r, 34)
    r += 1
    total_w = sum(c["weight"] or 0 for c in m.cats if c["score"] is not None)
    contrib_sum = 0.0
    for c in m.cats:
        w = c["weight"] or 0
        contrib = (c["score"] * w / total_w) if (c["score"] is not None and total_w) else None
        if contrib is not None:
            contrib_sum += contrib
        cv.cell(r, 2, c["name"], font=F(10.5, True, NAVY), fill=PAPER, align=A("left", "center", True, 1))
        for cc in (3, 4, 5):
            cv.cell(r, cc, None, fill=PAPER)
        cv.cell(r, 6, c["score"] if c["score"] is not None else "not scored", font=F(10.5, True, INK), fill=PAPER, align=A("center"), fmt="0.0")
        cv.cell(r, 7, w if w else 0, font=F(10, False, INK), fill=PAPER, align=A("center"), fmt="0.0%")
        cv.cell(r, 8, contrib if contrib is not None else "—", font=F(10.5, True, TEAL), fill=PAPER, align=A("center"), fmt="0.0")
        cv.height(r, 24)
        r += 1
        comp_total = sum(numeric(p.get("weight")) or 0 for p in c["components"]) or 1.0
        for p in c["components"]:
            kind = "percent"
            if p.get("metricKey") in METRIC_INFO:
                kind = METRIC_INFO[p["metricKey"]]["kind"]
            val = numeric(p.get("value"))
            wk, gl = numeric(p.get("worry")), numeric(p.get("goal"))
            cv.cell(r, 2, p.get("name"), font=F(10, False, INK), align=A("left", "center", True, 3))
            cv.cell(r, 3, fmt_by_kind(kind, val, m) if val is not None else "not available", font=F(10, False, INK), align=A("center"))
            cv.cell(r, 4, fmt_by_kind(kind, wk, m, True) if wk is not None else "—", font=F(10, False, MUTED), align=A("center"))
            cv.cell(r, 5, fmt_by_kind(kind, gl, m, True) if gl is not None else "—", font=F(10, False, MUTED), align=A("center"))
            sc = numeric(p.get("score"))
            cv.cell(r, 6, sc if sc is not None else "—", font=F(10, False, INK), align=A("center"), fmt="0.0")
            pw = numeric(p.get("weight"))
            cv.cell(r, 7, (pw / comp_total) if pw is not None else "—", font=F(10, False, MUTED), align=A("center"), fmt="0%")
            cv.height(r, 24)
            for cc in range(2, 9):
                ws.cell(r, cc).border = Border(bottom=HAIR)
            r += 1
    if m.score is not None and m.cats:
        cv.block(r, 2, r, 7, "Total of the points added (matches the health score)", font=F(10.5, True, NAVY), fill="EAF0F5", align=A("right", "center", False, 1))
        cv.cell(r, 8, contrib_sum, font=F(11, True, NAVY), fill="EAF0F5", align=A("center"), fmt="0.0")
        cv.height(r, 26)
        r += 1
        if abs(contrib_sum - m.score) > 0.2:
            cv.text(r, 2, 8, f"Note: this sum ({contrib_sum:.1f}) differs slightly from the supplied score ({m.score:.1f}) because of rounding in the supplied category scores.",
                    size=9, italic=True, color=MUTED, indent=1, min_h=24)
            r += 1
    cv.text(r, 2, 8, "The goal and worry levels are illustrative settings of this prototype model. They have not been tested against real "
                     "lending or business outcomes and are not industry standards.", size=9, italic=True, color=MUTED, indent=1, min_h=30)
    r += 2

    # validation ---------------------------------------------------------------
    checks = [v for v in (m.raw.get("validationChecks") or []) if isinstance(v, dict)]
    if checks:
        r = cv.section(r, 2, 8, "Number checks", "Simple tests that the supplied numbers add up.")
        for j, h in enumerate(["Check", "Year", "Result", "Details"]):
            cols = [(2, 2), (3, 3), (4, 4), (5, 8)][j]
            cv.block(r, cols[0], r, cols[1], h, font=F(9.5, True, NAVY), fill="EAF0F5", align=A("left", "center", False, 1))
        r += 1
        names = {"accounting_equation": "Assets = liabilities + equity", "cash_reconciliation": "Opening cash + change = closing cash"}
        for ch in checks:
            status = str(ch.get("status") or "unknown").lower()
            ok = status == "passed"
            cv.cell(r, 2, names.get(str(ch.get("check")), human(ch.get("check") or ch.get("name") or "Check")), font=F(10), align=A("left", "center", True, 1))
            cv.cell(r, 3, short_date(ch.get("period")), font=F(10), align=A("left", "center", False, 1))
            cv.cell(r, 4, ("PASSED" if ok else status.replace("_", " ").upper()), font=F(9, True, WHITE), fill=GOOD if ok else BAD, align=A("center"))
            diff = ch.get("difference")
            cv.block(r, 5, r, 8, ("Difference: " + f"{diff:g}" if isinstance(diff, (int, float)) else safe_text(ch.get("message") or "")), font=F(10, False, MUTED), align=A("left", "center", True, 1))
            cv.height(r, 22)
            for cc in range(2, 9):
                ws.cell(r, cc).border = Border(bottom=HAIR)
            r += 1
        r += 1

    for heading, items in [("Warnings and notes", m.warnings), ("Problems that need a human to review", m.errors)]:
        if not items:
            continue
        r = cv.section(r, 2, 8, heading)
        for i, t in enumerate(items, 1):
            cv.text(r, 2, 8, f"{i}.  {t}", size=10, indent=1, min_h=22, pad=8)
            r += 1
        r += 1

    r = cv.section(r, 2, 8, "Please remember")
    cv.text(r, 2, 8, "This is automated, unaudited analysis of the documents that were supplied. The numbers, currency, scale and years should be "
                     "checked against the originals before anyone uses them to lend, invest or make business decisions. "
                     "The health score is a teaching and screening aid, not a credit rating.", size=10, indent=1, min_h=60)
    ws.print_area = f"A1:I{r + 2}"


# ---------------------------------------------------------------------------
# Entry points
# ---------------------------------------------------------------------------
def create_workbook(report: dict) -> io.BytesIO:
    if not isinstance(report, dict):
        raise ValueError("Expected a financial report JSON object")
    if not report.get("financialMetrics") and not report.get("consolidatedStatements") and not report.get("scoreCategories"):
        raise ValueError("No financial metrics, statements or scoring categories supplied")
    m = Model(report)
    wb = Workbook()
    wb.properties.title = f"Financial Health Report - {m.company}"
    wb.properties.creator = "Financial Report Automation"
    wb.properties.subject = "Plain-language financial health report"
    ws0 = wb.active
    ws0.title = "Overview"
    for name in ["Health Checks", "Statements", "Insights", "Money Words", "About This Report"]:
        wb.create_sheet(name)
    data_ws = wb.create_sheet("_chart_data")
    data_ws.sheet_state = "hidden"
    cd = ChartData(data_ws)

    build_overview(wb, m, cd)
    build_health(wb, m, cd)
    build_statements(wb, m)
    build_insights(wb, m)
    build_glossary(wb, m)
    build_about(wb, m)
    wb.active = 0
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out


if app is not None:
    @app.get("/")
    def home():
        return {"status": "online", "service": "Financial Report Generator", "version": "3.0.0",
                "endpoints": ["POST /generate-report"]}

    @app.post("/generate-report")
    def generate_report(payload: dict = Body(...)):
        report = payload.get("report", payload)
        if not isinstance(report, dict):
            raise HTTPException(status_code=422, detail="Expected a JSON report object")
        try:
            data = create_workbook(report).getvalue()
        except ValueError as exc:
            raise HTTPException(status_code=422, detail=str(exc)) from exc
        except Exception as exc:
            raise HTTPException(status_code=500, detail=f"Workbook generation failed: {exc}") from exc
        stem = re.sub(r"[^A-Za-z0-9_-]+", "_", str(report.get("company") or "Financial_Analysis")).strip("_")[:60]
        return Response(
            content=data,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f'attachment; filename="{stem or "Financial_Analysis"}_Report.xlsx"'},
        )


if __name__ == "__main__":  # python excel_server.py report.json out.xlsx
    import json
    if len(sys.argv) < 3:
        print("usage: python excel_server.py report.json out.xlsx")
        sys.exit(1)
    payload = json.load(open(sys.argv[1], encoding="utf-8"))
    rep = payload.get("report", payload) if isinstance(payload, dict) else payload
    with open(sys.argv[2], "wb") as fh:
        fh.write(create_workbook(rep).getvalue())
    print("wrote", sys.argv[2])