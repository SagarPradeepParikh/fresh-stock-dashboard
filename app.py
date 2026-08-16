"""Evidence-first India/US equity research dashboard for Streamlit Cloud."""
from __future__ import annotations

import datetime as dt
import base64
import hashlib
import io
import json
import math
import re
import time
from dataclasses import dataclass
from typing import Any
from urllib.parse import urljoin, urlparse
from zoneinfo import ZoneInfo

import pandas as pd
import numpy as np
import plotly.graph_objects as go
import requests
import streamlit as st
import yfinance as yf
from bs4 import BeautifulSoup
from docx import Document
from pypdf import PdfReader
from streamlit_local_storage import LocalStorage

st.set_page_config(page_title="Evidence-First Equity Research", page_icon="📊", layout="wide")

BACKGROUND_SHARE_URL = "https://share.google/KZHe3f7dRE3dJAZ27"


@st.cache_data(ttl=86_400, show_spinner=False)
def public_background_data_url(url: str) -> str | None:
    """Use the supplied link only when it resolves to a publicly readable image."""
    try:
        response = requests.get(url, timeout=15, allow_redirects=True)
        content_type = response.headers.get("Content-Type", "").split(";", 1)[0].lower()
        if response.ok and content_type.startswith("image/"):
            return f"data:{content_type};base64,{base64.b64encode(response.content).decode('ascii')}"
    except requests.RequestException:
        pass
    return None


background_image = public_background_data_url(BACKGROUND_SHARE_URL)
chart_art = "linear-gradient(115deg, transparent 0 18%, rgba(0,234,255,.18) 18.2% 18.5%, transparent 18.7% 33%, rgba(255,51,197,.16) 33.2% 33.6%, transparent 33.8% 51%, rgba(255,197,61,.16) 51.2% 51.5%, transparent 51.7% 100%), radial-gradient(circle at 12% 20%, rgba(0,234,255,.3) 0 2px, transparent 3px), radial-gradient(circle at 72% 70%, rgba(255,61,180,.24) 0 2px, transparent 3px)"
background_css = (
    f"background-image: {chart_art}, linear-gradient(rgba(6,10,31,.82), rgba(6,10,31,.92)), url('{background_image}');"
    if background_image else f"background-image: {chart_art}; background-color: #060a1f;"
)

st.markdown("""<style>
.stApp { """ + background_css + """ background-attachment: fixed; background-size: 100% 100%, 54px 54px, 72px 72px, 100% 100%, cover; background-position: center; color: #f3f4f6; overflow-x: hidden; }
div[data-testid="stMetric"] { background: rgba(12, 20, 47, .72); border: 1px solid rgba(87, 220, 255, .26); border-radius: 12px; padding: 10px; backdrop-filter: blur(4px); }
section[data-testid="stSidebar"] { background: #111827; }
.amber-upload { border: 2px dashed #f59e0b; border-radius: 10px; padding: 12px; background: #3a2b11; }
div.stButton > button[kind="primary"] { background: #16a34a; border-color: #22c55e; color: white; font-weight: 700; }
</style>""", unsafe_allow_html=True)

US_UNIVERSE = {
    "Large cap": ["AAPL", "MSFT", "NVDA", "AMZN", "GOOGL", "META", "BRK-B", "AVGO", "TSLA", "JPM"],
    "Mid cap": ["PLTR", "COIN", "TOST", "HOOD", "AFRM", "RBLX", "DKNG", "CAVA", "RDDT", "DUOL"],
    "Small cap": ["RIG", "SENS", "COMP", "GPRO", "KOD", "BBIG", "FUBO", "SPCE", "BBAI", "MVIS"],
}
INDIA_UNIVERSE = {
    "Large cap": ["RELIANCE", "TCS", "HDFCBANK", "ICICIBANK", "INFY", "BHARTIARTL", "ITC", "LT", "SBIN", "HINDUNILVR"],
    "Mid cap": ["TATAPOWER", "BEL", "VOLTAS", "DIXON", "CUMMINSIND", "POLYCAB", "TRENT", "LUPIN", "PERSISTENT", "COFORGE"],
    "Small cap": ["SUZLON", "IRFC", "CAMS", "CERA", "KAYNES", "CLEAN", "BLS", "CDSL", "JYOTHYLAB", "EASEMYTRIP"],
}
DUAL_LISTINGS = {
    "Infosys": {"US": "INFY", "India": "INFY"},
    "Wipro": {"US": "WIT", "India": "WIPRO"},
    "HDFC Bank": {"US": "HDB", "India": "HDFCBANK"},
    "ICICI Bank": {"US": "IBN", "India": "ICICIBANK"},
    "Dr. Reddy's Laboratories": {"US": "RDY", "India": "DRREDDY"},
}

for key, value in {"watchlist": [], "active_ticker": "AAPL", "market": "US", "open_ir": False, "dual_listing": "Custom ticker", "last_good_prices": {}}.items():
    if key not in st.session_state:
        st.session_state[key] = value

# This component persists only public, provider-returned fundamentals in the
# current browser. API keys, uploaded reports and user details are never stored.
BROWSER_STORAGE = LocalStorage()


def secret(name: str) -> str | None:
    """Read an optional secret without crashing where no secrets.toml exists."""
    try:
        value = st.secrets.get(name)
        return str(value).strip() if value else None
    except Exception:
        return None


def symbol_for_market(ticker: str, market: str) -> str:
    ticker = ticker.strip().upper().replace(" ", "")
    if market == "India" and ticker and not ticker.endswith(".NS"):
        return f"{ticker}.NS"
    return ticker.replace(".NS", "") if market == "US" else ticker


def display_symbol(symbol: str) -> str:
    return symbol.removesuffix(".NS")


@st.cache_data(ttl=900, show_spinner=False)
def company_name_matches(query: str, market: str) -> list[dict[str, str]]:
    """Return real Yahoo Finance search matches for a company name or ticker."""
    if len(query.strip()) < 2:
        return []
    try:
        quotes = yf.Search(query.strip(), max_results=15, news_count=0).quotes or []
    except Exception:
        return []
    matches = []
    seen = set()
    for quote in quotes:
        symbol = str(quote.get("symbol", "")).upper()
        name = str(quote.get("longname") or quote.get("shortname") or "")
        exchange = str(quote.get("exchDisp") or quote.get("exchange") or "")
        country = str(quote.get("country") or "")
        if not symbol or symbol in seen or quote.get("quoteType") not in {"EQUITY", "MUTUALFUND", None}:
            continue
        india_match = symbol.endswith(".NS") or country.lower() == "india" or "nse" in exchange.lower()
        if (market == "India" and not india_match) or (market == "US" and india_match):
            continue
        seen.add(symbol)
        matches.append({"symbol": symbol, "name": name or symbol, "exchange": exchange or "Unavailable"})
    return matches


def num(value: Any) -> float | None:
    try:
        answer = float(value)
        return answer if pd.notna(answer) else None
    except (TypeError, ValueError):
        return None


def money(value: Any, currency: str) -> str:
    value = num(value)
    if value is None:
        return "Unavailable"
    signs = {"USD": "$", "INR": "₹", "EUR": "€", "GBP": "£"}
    return f"{signs.get(currency, '')}{value:,.2f}" if currency in signs else f"{value:,.2f} {currency}"


def ratio(value: Any, pct: bool = False) -> str:
    value = num(value)
    if value is None:
        return "Unavailable"
    return f"{value * 100:,.2f}%" if pct else f"{value:,.2f}"


def compact_money(value: Any, currency: str) -> str:
    value = num(value)
    if value is None:
        return "Unavailable"
    for divisor, label in [(1e12, "T"), (1e9, "B"), (1e6, "M")]:
        if abs(value) >= divisor:
            return f"{money(value / divisor, currency)}{label}"
    return money(value, currency)


@st.cache_data(ttl=300, show_spinner=False)
def usd_conversion_factor(currency: str) -> tuple[float | None, str]:
    """Return live Yahoo FX conversion from a reported currency to USD.

    The financial source's native amounts are converted for presentation only;
    ratios are deliberately not converted.  If Yahoo does not return a quote,
    no converted number is shown.
    """
    if currency == "USD":
        return 1.0, "USD native"
    try:
        if currency == "INR":
            quote = yf.Ticker("INR=X").history(period="5d", interval="1d")["Close"].dropna()
            rate = float(quote.iloc[-1])  # INR per USD
            return (1 / rate), f"1 USD = {rate:,.4f} INR"
        quote = yf.Ticker(f"{currency}USD=X").history(period="5d", interval="1d")["Close"].dropna()
        rate = float(quote.iloc[-1])  # USD per one unit of local currency
        return rate, f"1 {currency} = {rate:,.6f} USD"
    except Exception:
        return None, "Live FX quote unavailable"


def usd_millions(value: Any, conversion_factor: float | None) -> str:
    value = num(value)
    if value is None or conversion_factor is None:
        return "Unavailable"
    return f"${(value * conversion_factor / 1_000_000):,.2f}M"


def display_line_items(items: list[tuple[str, Any]]) -> None:
    """Render readable labels and values one per line, including on mobile."""
    for label, value in items:
        st.markdown(f"**{label}:** {value}")


def display_evidence(section: str, evidence: dict[str, list[dict[str, str]]], empty_message: str) -> None:
    """Display deterministic source excerpts with clickable source URLs."""
    rows = evidence.get(section, [])
    if rows:
        st.dataframe(pd.DataFrame(rows), width="stretch", hide_index=True, column_config={"Source": st.column_config.LinkColumn("Source")})
    else:
        st.info(empty_message)


BOTTLENECK_DEFINITIONS = {
    "Bottleneck": "Demand is high relative to available supply, so the relevant raw material, product, service, or by-product may become constrained and prices may rise.",
    "At par": "Demand and supply are broadly balanced; price movement is limited or mainly reflects general inflation or deflation.",
    "Eased out": "Supply or availability is high relative to demand, which can contribute to lower prices.",
}


def _document_text(url: str, headers: dict[str, str]) -> str:
    try:
        response = requests.get(url, headers=headers, timeout=15)
        response.raise_for_status()
        content_type = response.headers.get("Content-Type", "").lower()
        if "pdf" in content_type or url.lower().endswith(".pdf"):
            reader = PdfReader(io.BytesIO(response.content))
            return " ".join(reader.pages[index].extract_text() or "" for index in range(min(30, len(reader.pages))))
        return BeautifulSoup(response.text, "html.parser").get_text(" ", strip=True)
    except (requests.RequestException, ValueError, OSError):
        return ""


@st.cache_data(ttl=900, show_spinner=False)
def supply_demand_assessments(source_urls: tuple[str, ...], market: str, sec_user_agent: str | None) -> dict[str, dict[str, Any]]:
    """Create a disclosed-language heuristic, without making an investment forecast."""
    headers = {"User-Agent": sec_user_agent} if market == "US" and sec_user_agent else {"User-Agent": "Mozilla/5.0"}
    source_text = " ".join(_document_text(url, headers) for url in source_urls[:3])
    time.sleep(0.12 if market == "US" else 0.25)
    sentences = re.split(r"(?<=[.!?])\s+", re.sub(r"\s+", " ", source_text))
    periods = {
        "Past": ("previous", "prior", "historical", "formerly", "year ended"),
        "Present": ("current", "currently", "today", "now", "ongoing"),
        "Future": ("expect", "expected", "outlook", "forecast", "will", "anticipate", "future"),
    }
    positive = ("shortage", "supply constraint", "constrained supply", "capacity constraint", "supply disruption", "high demand", "demand exceeded", "backlog", "price increase")
    negative = ("oversupply", "excess capacity", "weak demand", "demand decline", "inventory build", "price decline", "supply glut", "overcapacity")
    neutral = ("balanced supply", "stable demand", "stable pricing", "normalized", "normalised", "equilibrium")
    result = {}
    for period, markers in periods.items():
        relevant = [sentence for sentence in sentences if any(marker in sentence.lower() for marker in markers)]
        text = " ".join(relevant) if relevant else source_text
        high = sum(text.lower().count(term) for term in positive)
        low = sum(text.lower().count(term) for term in negative)
        balanced = sum(text.lower().count(term) for term in neutral)
        raw_score = high - low
        # Signed logarithmic normalization: -100 (eased-out) to +100 (bottleneck).
        score = 0.0 if raw_score == 0 else (math.copysign(math.log1p(abs(raw_score)) / math.log(6) * 100, raw_score))
        score = max(-100.0, min(100.0, score))
        if not source_text:
            classification = "Insufficient public source text"
        elif score >= 25:
            classification = "Bottleneck evidence"
        elif score <= -25:
            classification = "Eased-out evidence"
        elif balanced or raw_score == 0:
            classification = "At-par / inconclusive evidence"
        else:
            classification = "At-par / inconclusive evidence"
        result[period] = {"score": score, "classification": classification, "high": high, "low": low, "balanced": balanced}
    return result


def supply_demand_gauge(period: str, result: dict[str, Any]) -> None:
    """Show the red-left, amber-centre, green-right logarithmic evidence scale."""
    figure = go.Figure(go.Indicator(
        mode="gauge+number",
        value=result["score"],
        number={"suffix": "", "font": {"size": 28, "color": "white"}},
        title={"text": f"{period}<br><span style='font-size:0.75em'>{result['classification']}</span>", "font": {"color": "white", "size": 16}},
        gauge={
            "axis": {"range": [-100, 100], "tickvals": [-100, 0, 100], "ticktext": ["Eased out", "At par", "Bottleneck"], "tickcolor": "white"},
            "bar": {"color": "white"},
            "bgcolor": "rgba(0,0,0,0)",
            "steps": [
                {"range": [-100, -25], "color": "#dc2626"},
                {"range": [-25, 25], "color": "#f59e0b"},
                {"range": [25, 100], "color": "#16a34a"},
            ],
        },
    ))
    figure.update_layout(height=260, margin={"l": 15, "r": 15, "t": 55, "b": 5}, paper_bgcolor="rgba(0,0,0,0)", font={"color": "white"})
    st.plotly_chart(figure, width="stretch", config={"displayModeBar": False})
    st.caption(f"Disclosed-language counts — bottleneck: {result['high']}; eased-out: {result['low']}; at-par: {result['balanced']}.")


def deterministic_supply_brief(assessment: dict[str, dict[str, Any]], evidence: dict[str, list[dict[str, str]]]) -> str:
    """Produce a sub-400-word explanation without an AI model."""
    readings = "; ".join(f"{period}: {result['classification']} (score {result['score']:.0f})" for period, result in assessment.items())
    excerpts = evidence.get("11. Supply-demand evidence", [])[:2]
    source_note = " ".join(f"Source excerpt: {item['Evidence excerpt'][:300]}" for item in excerpts)
    text = (
        f"The gauge is based on a logarithmically normalised count of supply-demand language in the selected public filings or Investor Relations pages. "
        f"Readings are {readings}. Terms such as shortage, constrained supply, high demand and backlog move the indicator right toward bottleneck; "
        f"oversupply, excess capacity, weak demand and inventory build move it left toward eased out; balanced supply, stable demand and normalised pricing support the centre. "
        f"The result is evidence-based, not a price forecast. Geopolitical effects are included only where the selected official source explicitly mentions them. {source_note}"
    )
    return " ".join(text.split()[:400])


def ai_supply_brief(company: str, assessment: dict[str, dict[str, Any]], evidence: dict[str, list[dict[str, str]]], api_key: str | None, model: str | None, news_context: list[dict[str, Any]]) -> tuple[str | None, str | None]:
    """Generate a concise source-bound explanation using an optional Gemini key."""
    if not api_key:
        return None, "Add GEMINI_API_KEY to Streamlit secrets to generate the optional AI explanation."
    source_rows = evidence.get("11. Supply-demand evidence", []) + evidence.get("5. Market and industry position", [])
    source_text = "\n".join(f"SOURCE: {item['Source']}\nEXCERPT: {item['Evidence excerpt']}" for item in source_rows[:5])
    news_text = "\n".join(f"NEWS: {item.get('headline', '')} — {item.get('summary', '')}" for item in news_context[:5])
    prompt = f"""Write no more than 350 words for Section 11 of an equity-research dashboard about {company}.
Explain why the disclosed-language gauge reads Past={assessment['Past']['classification']}, Present={assessment['Present']['classification']}, Future={assessment['Future']['classification']}.
Use only the supplied official filing excerpts and provider news context. Discuss geopolitical conditions only if directly supported by those sources. Do not give investment advice, do not invent facts, and explicitly state when source evidence is insufficient.

OFFICIAL SOURCES:
{source_text or 'No usable official source excerpt.'}

PROVIDER NEWS CONTEXT:
{news_text or 'No provider news supplied.'}
"""
    try:
        from google import genai
        response = genai.Client(api_key=api_key).models.generate_content(model=model or "gemini-1.5-flash", contents=prompt)
        return " ".join((response.text or "No AI explanation was returned.").split()[:400]), None
    except Exception as exc:
        return None, f"AI explanation was unavailable: {exc}"


def market_clock() -> dict[str, dict[str, str]]:
    now = dt.datetime.now(dt.timezone.utc)
    clocks = {"US": ("America/New_York", dt.time(9, 30), dt.time(16)), "India": ("Asia/Kolkata", dt.time(9, 15), dt.time(15, 30))}
    out = {}
    for name, (zone, start, end) in clocks.items():
        local = now.astimezone(ZoneInfo(zone))
        open_now = local.weekday() < 5 and start <= local.time() <= end
        # Regular-hours status only. Exchange holiday calendars need a licensed calendar source.
        out[name] = {"status": "🟢 OPEN" if open_now else "🔴 CLOSED", "time": local.strftime("%d %b, %I:%M %p %Z")}
    return out


def next_quarter_end(today: dt.date) -> dt.date:
    candidates = [dt.date(year, month, day) for year in (today.year, today.year + 1) for month, day in ((3, 31), (6, 30), (9, 30), (12, 31))]
    return min(day for day in candidates if day >= today)


def statutory_calendar(market: str) -> pd.DataFrame:
    """Show general upcoming compliance windows, not issuer-specific legal advice."""
    today = dt.date.today()
    period_end = next_quarter_end(today)
    if market == "India":
        financial_due_days = 60 if period_end.month == 3 else 45
        rows = [
            ("Integrated Filing (Governance)", period_end + dt.timedelta(days=30), "SEBI LODR general timeline; eligibility and amendments may vary."),
            ("Financial results / Integrated Filing (Financial)", period_end + dt.timedelta(days=financial_due_days), "45 days after ordinary quarter-end; 60 days after March year-end."),
            ("Shareholding pattern", period_end + dt.timedelta(days=21), "General quarterly SEBI LODR timeline."),
        ]
    else:
        rows = [
            ("Form 10-Q", period_end + dt.timedelta(days=40), "Accelerated / large accelerated filer general deadline; non-accelerated filers generally have 45 days."),
            ("Form 10-Q", period_end + dt.timedelta(days=45), "Non-accelerated filer general deadline; issuer fiscal year and eligibility control."),
            ("Form 8-K", "Event based", "Generally four business days after a reportable event."),
        ]
    calendar = pd.DataFrame(rows, columns=["Key statutory filing", "Indicative next due date", "Scope / qualification"])
    # Arrow requires one consistent type per column; statutory dates may include event-based text.
    calendar["Indicative next due date"] = calendar["Indicative next due date"].map(lambda value: value.isoformat() if isinstance(value, dt.date) else str(value))
    return calendar


def latest_statement_period(frame: pd.DataFrame) -> str:
    """Return the latest financial period date without representing it as a filing date."""
    if frame.empty:
        return "Unavailable"
    dates = []
    for column in frame.columns:
        try:
            dates.append(pd.Timestamp(column).date())
        except (TypeError, ValueError):
            pass
    return max(dates).isoformat() if dates else "Unavailable"


@st.cache_data(ttl=600, show_spinner=False)
def yahoo_price(symbol: str) -> dict[str, Any]:
    """Fresh price/history cache: five minutes, separate from the four-hour ratios cache."""
    try:
        history = yf.Ticker(symbol).history(period="1y", interval="1d", auto_adjust=False, actions=False)
        if history is None or history.empty:
            return {"history": pd.DataFrame(), "error": "Yahoo Finance returned no price history."}
        history.index = pd.to_datetime(history.index)
        close = history["Close"].dropna()
        latest = float(close.iloc[-1]) if not close.empty else None
        previous = float(close.iloc[-2]) if len(close) > 1 else None
        return {"history": history, "price": latest, "previous": previous, "as_of": history.index[-1], "error": None}
    except Exception as exc:
        return {"history": pd.DataFrame(), "error": f"Yahoo price request failed: {exc}"}


@st.cache_data(ttl=60, show_spinner=False)
def finnhub_quote(symbol: str, token: str | None) -> dict[str, Any]:
    """Use Finnhub for a fresh US quote when configured, avoiding Yahoo throttling."""
    if not token:
        return {"price": None, "previous": None, "as_of": None, "error": "Finnhub is not configured."}
    try:
        response = requests.get("https://finnhub.io/api/v1/quote", params={"symbol": symbol, "token": token}, timeout=12)
        response.raise_for_status()
        quote = response.json()
        current = num(quote.get("c"))
        previous = num(quote.get("pc"))
        if current is None or current == 0:
            return {"price": None, "previous": None, "as_of": None, "error": "Finnhub returned no current quote for this symbol."}
        timestamp = dt.datetime.fromtimestamp(int(quote.get("t", 0)), dt.timezone.utc) if quote.get("t") else None
        return {"price": current, "previous": previous, "as_of": timestamp, "error": None}
    except (requests.RequestException, ValueError, TypeError) as exc:
        return {"price": None, "previous": None, "as_of": None, "error": f"Finnhub quote request failed: {exc}"}


@st.cache_data(ttl=900, show_spinner=False)
def price_history(symbol: str, market: str, token: str | None) -> pd.DataFrame:
    """Return a one-year chart independently of the live-quote provider.

    Finnhub is tried first for US symbols when configured. Yahoo is retained as
    a read-only fallback, so a Finnhub live quote never removes the chart.
    """
    if market == "US" and token:
        try:
            end = int(dt.datetime.now(dt.timezone.utc).timestamp())
            start = end - 370 * 24 * 60 * 60
            response = requests.get(
                "https://finnhub.io/api/v1/stock/candle",
                params={"symbol": symbol, "resolution": "D", "from": start, "to": end, "token": token},
                timeout=15,
            )
            response.raise_for_status()
            payload = response.json()
            if payload.get("s") == "ok" and payload.get("t") and payload.get("c"):
                return pd.DataFrame({"Close": payload["c"]}, index=pd.to_datetime(payload["t"], unit="s", utc=True))
        except (requests.RequestException, ValueError, TypeError):
            pass
    try:
        history = yf.Ticker(symbol).history(period="1y", interval="1d", auto_adjust=False, actions=False)
        if history is None or history.empty or "Close" not in history:
            return pd.DataFrame()
        return history[["Close"]].dropna()
    except Exception:
        return pd.DataFrame()


@st.cache_data(ttl=14400, show_spinner=False)
def yahoo_fundamentals(symbol: str) -> dict[str, Any]:
    """Server-side four-hour cache for ratios and statements; never creates substitute data."""
    try:
        ticker = yf.Ticker(symbol)
        try:
            info = ticker.get_info() or {}
        except Exception:
            info = {}
        def frame(getter: str) -> pd.DataFrame:
            try:
                result = getattr(ticker, getter)
                return result if isinstance(result, pd.DataFrame) else pd.DataFrame()
            except Exception:
                return pd.DataFrame()
        return {"info": info, "income": frame("income_stmt"), "quarterly_income": frame("quarterly_income_stmt"), "balance": frame("balance_sheet"), "quarterly_balance": frame("quarterly_balance_sheet"), "cashflow": frame("cashflow"), "quarterly_cashflow": frame("quarterly_cashflow"), "error": None}
    except Exception as exc:
        return {"info": {}, "income": pd.DataFrame(), "quarterly_income": pd.DataFrame(), "balance": pd.DataFrame(), "quarterly_balance": pd.DataFrame(), "cashflow": pd.DataFrame(), "quarterly_cashflow": pd.DataFrame(), "error": f"Yahoo fundamental request failed: {exc}"}


def _frame_to_cache(frame: pd.DataFrame) -> str:
    """Make a DataFrame safe for the browser's JSON-only localStorage."""
    if frame.empty:
        return ""
    return frame.to_json(orient="split", date_format="iso", default_handler=str)


def _frame_from_cache(payload: str) -> pd.DataFrame:
    if not payload:
        return pd.DataFrame()
    try:
        return pd.read_json(io.StringIO(payload), orient="split")
    except (ValueError, TypeError):
        return pd.DataFrame()


def _json_safe_info(info: dict[str, Any]) -> dict[str, Any]:
    """Keep actual scalar fundamentals and provider text; omit non-JSON objects."""
    clean = {}
    for key, value in info.items():
        if isinstance(value, (str, int, float, bool)) or value is None:
            clean[key] = value
    return clean


def browser_cached_fundamentals(symbol: str) -> tuple[dict[str, Any], bool]:
    """Use a four-hour browser cache before Yahoo ratios/financial statements.

    `streamlit-local-storage` is asynchronous on the initial paint. A first
    visit may therefore call Yahoo once; later refreshes/revisits from the same
    browser use this localStorage payload and make no Yahoo fundamentals call.
    """
    cache_key = f"equity-dashboard:fundamentals:v1:{symbol}"
    raw = BROWSER_STORAGE.getItem(cache_key)
    try:
        saved = json.loads(raw) if isinstance(raw, str) else None
    except json.JSONDecodeError:
        saved = None

    now = dt.datetime.now(dt.timezone.utc).timestamp()
    if isinstance(saved, dict) and now - float(saved.get("saved_at", 0)) < 14_400:
        return {
            "info": saved.get("info", {}),
            "income": _frame_from_cache(saved.get("income", "")),
            "quarterly_income": _frame_from_cache(saved.get("quarterly_income", "")),
            "balance": _frame_from_cache(saved.get("balance", "")),
            "quarterly_balance": _frame_from_cache(saved.get("quarterly_balance", "")),
            "cashflow": _frame_from_cache(saved.get("cashflow", "")),
            "quarterly_cashflow": _frame_from_cache(saved.get("quarterly_cashflow", "")),
            "error": saved.get("error"),
        }, True

    live = yahoo_fundamentals(symbol)
    if not live.get("error"):
        browser_payload = {
            "saved_at": now,
            "info": _json_safe_info(live["info"]),
            "income": _frame_to_cache(live["income"]),
            "quarterly_income": _frame_to_cache(live["quarterly_income"]),
            "balance": _frame_to_cache(live["balance"]),
            "quarterly_balance": _frame_to_cache(live["quarterly_balance"]),
            "cashflow": _frame_to_cache(live["cashflow"]),
            "quarterly_cashflow": _frame_to_cache(live["quarterly_cashflow"]),
            "error": None,
        }
        BROWSER_STORAGE.setItem(cache_key, json.dumps(browser_payload, default=str))
    return live, False


@st.cache_data(ttl=900, show_spinner=False)
def finnhub_profile_and_news(symbol: str, token: str | None) -> dict[str, Any]:
    """Finnhub is used only for US-listed symbols and only when its user key exists."""
    if not token:
        return {"profile": {}, "news": [], "error": "Add FINNHUB_API_KEY to Streamlit secrets to enable Finnhub data."}
    base = "https://finnhub.io/api/v1"
    try:
        profile_response = requests.get(f"{base}/stock/profile2", params={"symbol": symbol, "token": token}, timeout=12)
        news_response = requests.get(f"{base}/company-news", params={"symbol": symbol, "from": (dt.date.today()-dt.timedelta(days=30)).isoformat(), "to": dt.date.today().isoformat(), "token": token}, timeout=12)
        profile = profile_response.json() if profile_response.ok else {}
        news = news_response.json() if news_response.ok else []
        return {"profile": profile if isinstance(profile, dict) else {}, "news": news if isinstance(news, list) else [], "error": None}
    except (requests.RequestException, ValueError) as exc:
        return {"profile": {}, "news": [], "error": f"Finnhub request failed: {exc}"}


SEC_FACTS = {
    "Revenue": ("RevenueFromContractWithCustomerExcludingAssessedTax", "Revenues", "SalesRevenueNet", "Revenue"),
    "Gross profit": ("GrossProfit",),
    "Operating income": ("OperatingIncomeLoss",),
    "Net income": ("NetIncomeLoss", "ProfitLoss"),
    "Cash and cash equivalents": ("CashAndCashEquivalentsAtCarryingValue",),
    "Total assets": ("Assets",),
    "Current assets": ("AssetsCurrent",),
    "Total liabilities": ("Liabilities",),
    "Current liabilities": ("LiabilitiesCurrent",),
    "Shareholders' equity": ("StockholdersEquity", "StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest"),
    "Long-term debt": ("LongTermDebtNoncurrent", "LongTermDebt"),
    "Operating cash flow": ("NetCashProvidedByUsedInOperatingActivities",),
    "Capital expenditure": ("PaymentsToAcquirePropertyPlantAndEquipment",),
}
SEC_STATEMENT_CONCEPTS = {
    "Income statement": {
        "Revenue": ("RevenueFromContractWithCustomerExcludingAssessedTax", "Revenues", "SalesRevenueNet", "Revenue"),
        "Gross profit": ("GrossProfit",),
        "Operating income": ("OperatingIncomeLoss",),
        "Net income": ("NetIncomeLoss", "ProfitLoss"),
    },
    "Balance sheet": {
        "Cash and cash equivalents": ("CashAndCashEquivalentsAtCarryingValue",),
        "Current assets": ("AssetsCurrent",),
        "Total assets": ("Assets",),
        "Current liabilities": ("LiabilitiesCurrent",),
        "Total liabilities": ("Liabilities",),
        "Shareholders' equity": ("StockholdersEquity", "StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest"),
        "Long-term debt": ("LongTermDebtNoncurrent", "LongTermDebt"),
    },
    "Cash-flow statement": {
        "Operating cash flow": ("NetCashProvidedByUsedInOperatingActivities",),
        "Capital expenditure": ("PaymentsToAcquirePropertyPlantAndEquipment",),
        "Cash dividends paid": ("PaymentsOfDividends",),
    },
}


def _latest_sec_fact(facts: dict[str, Any], concepts: tuple[str, ...]) -> float | None:
    """Get a latest annual USD fact, preserving the SEC filing's reported value."""
    for concept in concepts:
        for taxonomy in ("us-gaap", "ifrs-full"):
            fact = facts.get(taxonomy, {}).get(concept, {})
            units = fact.get("units", {}).get("USD", [])
            annual = [item for item in units if item.get("form") in {"10-K", "20-F", "40-F"} and item.get("fy")]
            candidates = annual or units
            if candidates:
                item = max(candidates, key=lambda row: (row.get("end", ""), row.get("filed", "")))
                return num(item.get("val"))
    return None


def _annual_sec_series(facts: dict[str, Any], concepts: tuple[str, ...]) -> dict[str, str]:
    """Return up to three annual SEC XBRL values, in original USD amounts shown as USD M."""
    for concept in concepts:
        for taxonomy in ("us-gaap", "ifrs-full"):
            units = facts.get(taxonomy, {}).get(concept, {}).get("units", {}).get("USD", [])
            annual = [item for item in units if item.get("form") in {"10-K", "20-F", "40-F"} and item.get("fy")]
            if annual:
                by_year = {}
                for item in annual:
                    year = str(item.get("fy"))
                    if year not in by_year or item.get("filed", "") > by_year[year].get("filed", ""):
                        by_year[year] = item
                return {year: usd_millions(item.get("val"), 1.0) for year, item in sorted(by_year.items(), reverse=True)[:3]}
    return {}


def sec_statement_frames(facts: dict[str, Any]) -> dict[str, pd.DataFrame]:
    """Build standardised annual SEC statement tables without commercial intermediaries."""
    tables = {}
    for statement_name, rows in SEC_STATEMENT_CONCEPTS.items():
        table_rows = []
        years = set()
        for label, concepts in rows.items():
            values = _annual_sec_series(facts, concepts)
            years.update(values)
            table_rows.append((label, values))
        columns = ["SEC XBRL line item", *sorted(years, reverse=True)]
        tables[statement_name] = pd.DataFrame([[label, *[values.get(year, "Unavailable") for year in columns[1:]]] for label, values in table_rows], columns=columns)
    return tables


@st.cache_data(ttl=900, show_spinner=False)
def sec_company_data(symbol: str, user_agent: str | None) -> dict[str, Any]:
    """Read public EDGAR filings and GAAP/IFRS-derived XBRL facts directly.

    SEC data is free and needs no API key.  The SEC requires a descriptive
    User-Agent with contact details, so the request is not attempted without it.
    """
    if not user_agent or "@" not in user_agent:
        return {"error": "Set SEC_USER_AGENT in Streamlit secrets, for example: Your App Name contact@example.com.", "facts": pd.DataFrame(), "filings": pd.DataFrame(), "statements": {}, "browse_url": None, "last_filed": None}
    headers = {"User-Agent": user_agent, "Accept-Encoding": "gzip, deflate", "Host": "data.sec.gov"}
    try:
        tickers_response = requests.get("https://www.sec.gov/files/company_tickers.json", headers={"User-Agent": user_agent}, timeout=15)
        tickers_response.raise_for_status()
        matches = [record for record in tickers_response.json().values() if record.get("ticker", "").upper() == symbol.upper()]
        if not matches:
            return {"error": "This ticker was not found in the SEC's company ticker file.", "facts": pd.DataFrame(), "filings": pd.DataFrame(), "statements": {}, "browse_url": None, "last_filed": None}
        record = matches[0]
        cik = str(record["cik_str"]).zfill(10)
        facts_response = requests.get(f"https://data.sec.gov/api/xbrl/companyfacts/CIK{cik}.json", headers=headers, timeout=20)
        submissions_response = requests.get(f"https://data.sec.gov/submissions/CIK{cik}.json", headers=headers, timeout=20)
        facts_response.raise_for_status()
        submissions_response.raise_for_status()
        facts_json = facts_response.json()
        rows = [{"SEC GAAP/IFRS metric": label, "Latest annual reported value (USD M)": usd_millions(_latest_sec_fact(facts_json.get("facts", {}), concepts), 1.0)} for label, concepts in SEC_FACTS.items()]
        recent = submissions_response.json().get("filings", {}).get("recent", {})
        filing_rows = []
        for form, filed, accession, document in zip(recent.get("form", []), recent.get("filingDate", []), recent.get("accessionNumber", []), recent.get("primaryDocument", [])):
            if form in {"10-K", "10-Q", "8-K", "20-F", "40-F", "DEF 14A"}:
                accession_path = accession.replace("-", "")
                filing_rows.append({"Form": form, "Filed": filed, "Official SEC filing": f"https://www.sec.gov/Archives/edgar/data/{int(cik)}/{accession_path}/{document}"})
            if len(filing_rows) == 12:
                break
        return {"error": None, "facts": pd.DataFrame(rows), "filings": pd.DataFrame(filing_rows), "statements": sec_statement_frames(facts_json.get("facts", {})), "browse_url": f"https://www.sec.gov/edgar/browse/?CIK={cik}", "company_name": facts_json.get("entityName"), "last_filed": filing_rows[0]["Filed"] if filing_rows else None}
    except (requests.RequestException, ValueError, KeyError) as exc:
        return {"error": f"SEC public-data request failed: {exc}", "facts": pd.DataFrame(), "filings": pd.DataFrame(), "statements": {}, "browse_url": None, "last_filed": None}


def official_exchange_links(symbol: str, market: str) -> dict[str, str]:
    """Official links only; this app does not bypass exchange access controls."""
    if market == "US":
        return {
            "SEC EDGAR company filings": "https://www.sec.gov/edgar/search/",
            "Nasdaq market activity": f"https://www.nasdaq.com/market-activity/stocks/{symbol.lower()}",
            "Dow Jones index information": "https://www.spglobal.com/spdji/en/indices/equity/dow-jones-industrial-average/",
        }
    base = display_symbol(symbol)
    return {
        "NSE company quote and disclosures": f"https://www.nseindia.com/get-quotes/equity?symbol={base}",
        "BSE corporate announcements": "https://www.bseindia.com/corporates/ann.html",
        "BSE board meetings": "https://www.bseindia.com/corporates/board_meeting.aspx",
    }


DISCLOSURE_KEYWORDS = {
    "4. Group structure": ("subsidiary", "parent company", "affiliate", "related party", "promoter"),
    "5. Market and industry position": ("market share", "largest", "competition", "competitor", "industry"),
    "8. Litigation, prospects and M&A": ("legal proceeding", "litigation", "contingent", "acquisition", "merger", "restructuring"),
    "9. Partnerships": ("partnership", "joint venture", "collaboration", "strategic alliance", "agreement"),
    "10. Key personnel": ("chief executive", "chief financial officer", "appointed", "resigned", "retired", "director"),
    "11. Supply-demand evidence": ("supply", "demand", "shortage", "capacity", "inventory", "geopolitical"),
}


def _keyword_snippets(text: str, keywords: tuple[str, ...], limit: int = 3) -> list[str]:
    normalized = re.sub(r"\s+", " ", text)
    snippets = []
    for keyword in keywords:
        match = re.search(re.escape(keyword), normalized, flags=re.IGNORECASE)
        if match:
            start = max(0, match.start() - 150)
            end = min(len(normalized), match.end() + 270)
            excerpt = normalized[start:end].strip()
            if excerpt not in snippets:
                snippets.append(excerpt)
        if len(snippets) >= limit:
            break
    return snippets


@st.cache_data(ttl=900, show_spinner=False)
def deterministic_disclosure_evidence(source_urls: tuple[str, ...], market: str, sec_user_agent: str | None) -> dict[str, list[dict[str, str]]]:
    """Extract disclosed keyword evidence from a few public official pages only.

    This is deterministic text matching, not AI analysis. It returns evidence
    excerpts, never inferred facts, forecasts, or a business classification.
    """
    results = {section: [] for section in DISCLOSURE_KEYWORDS}
    headers = {"User-Agent": sec_user_agent} if market == "US" and sec_user_agent else {"User-Agent": "Mozilla/5.0"}
    for url in source_urls[:3]:
        try:
            response = requests.get(url, headers=headers, timeout=15)
            response.raise_for_status()
            content_type = response.headers.get("Content-Type", "").lower()
            if "pdf" in content_type or url.lower().endswith(".pdf"):
                reader = PdfReader(io.BytesIO(response.content))
                text = " ".join(reader.pages[index].extract_text() or "" for index in range(min(30, len(reader.pages))))
            else:
                text = BeautifulSoup(response.text, "html.parser").get_text(" ", strip=True)
            for section, keywords in DISCLOSURE_KEYWORDS.items():
                for snippet in _keyword_snippets(text, keywords):
                    results[section].append({"Source": url, "Evidence excerpt": snippet})
            time.sleep(0.12 if market == "US" else 0.25)
        except (requests.RequestException, ValueError, OSError):
            continue
    return results


@st.cache_data(ttl=3600, show_spinner=False)
def resolve_investor_relations(website: str) -> dict[str, str | None]:
    """Find IR from the official domain, including common issuer-owned routes."""
    if not website or not website.startswith(("http://", "https://")):
        return {"website": website, "ir_url": None, "error": "No official website was supplied by the data provider."}
    try:
        response = requests.get(website, headers={"User-Agent": "Mozilla/5.0"}, timeout=12)
        response.raise_for_status()
        soup = BeautifulSoup(response.text, "html.parser")
        keywords = ("investor", "investor relations", "shareholder", "financial reports", "annual report")
        for anchor in soup.find_all("a", href=True):
            text = anchor.get_text(" ", strip=True).lower()
            href = anchor["href"]
            if any(word in text.lower() or word.replace(" ", "-") in href.lower() for word in keywords):
                candidate = urljoin(website, href)
                if urlparse(candidate).netloc == urlparse(website).netloc:
                    return {"website": website, "ir_url": candidate, "error": None}
        base = f"{urlparse(website).scheme}://{urlparse(website).netloc}"
        for path in ("/investors", "/investor-relations", "/investor", "/financial-information", "/annual-reports"):
            candidate = f"{base}{path}"
            probe = requests.get(candidate, headers={"User-Agent": "Mozilla/5.0"}, timeout=10, allow_redirects=True)
            if probe.ok and urlparse(probe.url).netloc == urlparse(website).netloc and len(probe.text) > 800:
                return {"website": website, "ir_url": probe.url, "error": None}
        return {"website": website, "ir_url": None, "error": "No investor-relations route was discoverable on the official domain. Use the SEC/NSE/BSE sources below or upload the issuer annual report."}
    except requests.RequestException as exc:
        return {"website": website, "ir_url": None, "error": f"Could not reach the official website: {exc}"}


def extract_uploaded_document(file) -> tuple[str, str | None]:
    """Read user-supplied official PDF, DOCX, XLSX, XLS, or CSV files locally."""
    raw = file.getvalue()
    suffix = file.name.rsplit(".", 1)[-1].lower()
    try:
        if suffix == "pdf":
            text = "\n".join(page.extract_text() or "" for page in PdfReader(io.BytesIO(raw)).pages)
        elif suffix == "docx":
            text = "\n".join(paragraph.text for paragraph in Document(io.BytesIO(raw)).paragraphs)
        elif suffix in {"xlsx", "xls", "csv"}:
            table = pd.read_csv(io.BytesIO(raw)) if suffix == "csv" else pd.read_excel(io.BytesIO(raw))
            text = table.to_csv(index=False)
        else:
            return "", "This file type is not supported. Upload PDF, DOCX, XLSX, XLS, or CSV."
        return text[:180000], None
    except Exception as exc:
        return "", f"Could not read the uploaded document: {exc}"


def gemini_analysis(text: str, company: str, api_key: str | None, model: str | None) -> dict[str, str]:
    """Sequential, source-grounded AI analysis. Does not produce output without a key and source text."""
    if not api_key:
        return {"error": "Add GEMINI_API_KEY to Streamlit secrets before using document analysis."}
    if not text.strip():
        return {"error": "Upload a readable official document before requesting analysis."}
    try:
        from google import genai
        client = genai.Client(api_key=api_key)
        prompts = [
            ("corporate_web", "Extract only disclosed group structure: parent, subsidiaries, associates, promoter entities, related parties, their business, investment plans, and cited page/section references. If absent, state not disclosed."),
            ("disclosures", "Return valid JSON only, with these string keys: section_6, section_8, section_9, section_10. section_6 covers AGM, investor/business meets and official communications. section_8 covers investor-report facts, prospects, litigation, contingent liabilities and M&A. section_9 covers current, past and announced partnerships. section_10 covers key personnel, joins and departures. Use only disclosed facts with page/section references. State not disclosed for missing information."),
            ("bottleneck", "Assess past, present and forward supply-demand conditions only where the document gives evidence. Label each as bottleneck, eased-out, par, or insufficient disclosed evidence. Explain in concise bullets and cite page/section references."),
        ]
        result = {}
        for key, instruction in prompts:  # Deliberately sequential: one request at a time.
            prompt = f"You are analysing an official document for {company}. {instruction}\n\nSOURCE DOCUMENT:\n{text}"
            response = client.models.generate_content(model=model or "gemini-1.5-flash", contents=prompt)
            result[key] = response.text or "No analysis text returned."
        try:
            parsed = json.loads(result["disclosures"].removeprefix("```json").removesuffix("```").strip())
            for key in ("section_6", "section_8", "section_9", "section_10"):
                result[key] = str(parsed.get(key, "Not disclosed in the uploaded source."))
        except (json.JSONDecodeError, AttributeError):
            result["section_6"] = result["disclosures"]
            result["section_8"] = "The structured disclosure response could not be parsed; review the Section 6 source output."
            result["section_9"] = "The structured disclosure response could not be parsed; review the Section 6 source output."
            result["section_10"] = "The structured disclosure response could not be parsed; review the Section 6 source output."
        return result
    except Exception as exc:
        return {"error": f"Gemini analysis failed: {exc}"}


def statement_view(frame: pd.DataFrame, title: str, conversion_factor: float | None) -> None:
    st.subheader(title)
    if frame.empty:
        st.info("The provider did not return this statement for the selected company.")
        return
    visible = frame.copy()
    visible.columns = [str(c.date()) if hasattr(c, "date") else str(c) for c in visible.columns]
    # Statements are source-reported absolute amounts. Convert only numeric values.
    for column in visible.columns:
        visible[column] = visible[column].map(lambda value: usd_millions(value, conversion_factor))
    st.dataframe(visible, width="stretch")


def price_chart(history: pd.DataFrame, symbol: str, currency: str) -> None:
    if history.empty:
        st.info("One-year price-chart data is temporarily unavailable from the configured providers.")
        return
    figure = go.Figure(go.Scatter(x=history.index, y=history["Close"], mode="lines", line={"color": "#38bdf8"}, name="Close"))
    figure.update_layout(title=f"{symbol}: one-year closing-price history", paper_bgcolor="#0b0f19", plot_bgcolor="#111827", font={"color": "#f3f4f6"}, height=380, margin={"l": 10, "r": 10, "t": 45, "b": 10})
    figure.update_yaxes(title=f"Price ({currency})", gridcolor="#374151")
    st.plotly_chart(figure, width="stretch")


def ranking_source_links(symbol: str, market: str, company: str) -> list[tuple[str, str]]:
    """Links to the requested third-party ranking pages; their figures remain provider-owned."""
    clean_symbol = display_symbol(symbol)
    company_query = company.replace(" ", "+")
    links = [("CSIMarket live US company/industry rankings", "https://csimarket.com/screening/most_valuable.php")]
    if market == "India":
        links.append(("Screener.in live company page", f"https://www.screener.in/company/{clean_symbol}/consolidated/"))
        links.append(("Screener.in company search", f"https://www.screener.in/search/?q={company_query}"))
    else:
        links.append(("CSIMarket live sector and industry screens", "https://csimarket.com/screening/index.php"))
    return links


def estimated_agm_schedule(market: str, sec_data: dict[str, Any] | None, income: pd.DataFrame) -> pd.DataFrame:
    """Show an estimate only when there is a clearly stated source basis."""
    rows: list[dict[str, str]] = []
    if market == "US" and sec_data is not None and not sec_data.get("filings", pd.DataFrame()).empty:
        proxies = sec_data["filings"].loc[sec_data["filings"]["Form"] == "DEF 14A"]
        if not proxies.empty:
            filed = pd.to_datetime(proxies.iloc[0]["Filed"], errors="coerce")
            if pd.notna(filed):
                rows.append({"Meeting": "Annual meeting / proxy season", "Past official proxy filing": filed.date().isoformat(), "Indicative next window": (filed.date() + dt.timedelta(days=365)).isoformat(), "Method": "One year after latest DEF 14A filing; confirm in the current proxy statement."})
    if market == "India":
        period = latest_statement_period(income)
        if period != "Unavailable":
            end = pd.Timestamp(period).date()
            rows.append({"Meeting": "AGM", "Past annual statement period": end.isoformat(), "Indicative next window": (end + dt.timedelta(days=180)).isoformat(), "Method": "Statutory outer-window estimate from the latest provider statement period; issuer notice controls."})
    return pd.DataFrame(rows)


US_RS_PEERS = ["AAPL", "MSFT", "NVDA", "AMZN", "GOOGL", "META", "AVGO", "JPM", "XOM", "LLY", "COST"]
INDIA_RS_PEERS = ["RELIANCE.NS", "TCS.NS", "HDFCBANK.NS", "ICICIBANK.NS", "INFY.NS", "BHARTIARTL.NS", "ITC.NS", "LT.NS", "SBIN.NS", "HINDUNILVR.NS", "SUNPHARMA.NS"]


def _close_from_download(downloaded: pd.DataFrame, ticker: str) -> pd.Series:
    """Read adjusted Close from yfinance's single or multi-index download output."""
    if downloaded.empty:
        return pd.Series(dtype=float)
    try:
        if isinstance(downloaded.columns, pd.MultiIndex):
            if ticker in downloaded.columns.get_level_values(0):
                series = downloaded[ticker]["Close"]
            else:
                series = downloaded["Close"][ticker]
        else:
            series = downloaded["Close"]
        return pd.to_numeric(series, errors="coerce").dropna()
    except (KeyError, TypeError):
        return pd.Series(dtype=float)


@st.cache_data(ttl=900, show_spinner=False)
def calculate_stock_rs_and_beta_engine(target_ticker: str) -> dict[str, Any]:
    """Calculate 40/20/20/20 proxy-peer RS percentile and 60-day covariance beta.

    The percentile is relative to the disclosed internal high-cap proxy pool,
    not a claim about every listed company. Errors are returned as data so the
    Streamlit page stays usable when a market provider throttles requests.
    """
    ticker = target_ticker.strip().upper()
    india = ticker.endswith((".NS", ".BO"))
    benchmark = "^NSEI" if india else "^GSPC"
    pipeline = "Bharat Market Pipeline" if india else "Wall Street Pipeline"
    peers = INDIA_RS_PEERS if india else US_RS_PEERS
    pool = list(dict.fromkeys([ticker, *peers]))
    print(f"[RS/Beta] {pipeline} engaged for {ticker}; benchmark={benchmark}; proxy peers={len(pool) - 1}")
    try:
        data = yf.download(pool + [benchmark], period="15mo", interval="1d", auto_adjust=True, progress=False, threads=True, group_by="ticker")
        scores: dict[str, float] = {}
        histories: dict[str, pd.Series] = {}
        for name in pool:
            close = _close_from_download(data, name)
            if len(close) < 240:
                continue
            histories[name] = close
            # Four non-overlapping trailing quarters, with the most recent weighted 40%.
            values = [float(close.iloc[-1] / close.iloc[-64] - 1), float(close.iloc[-64] / close.iloc[-127] - 1), float(close.iloc[-127] / close.iloc[-190] - 1), float(close.iloc[-190] / close.iloc[-253] - 1)]
            scores[name] = 0.40 * values[0] + 0.20 * values[1] + 0.20 * values[2] + 0.20 * values[3]
        if ticker not in scores:
            return {"ok": False, "Pipeline": pipeline, "Benchmark": benchmark, "error": "Insufficient price history: at least 240 active trading rows are required for the selected ticker."}
        if len(scores) < 2:
            return {"ok": False, "Pipeline": pipeline, "Benchmark": benchmark, "error": "Insufficient active proxy-peer price histories to calculate a percentile."}
        ordered = sorted(scores.items(), key=lambda item: item[1])
        rank = next(index + 1 for index, item in enumerate(ordered) if item[0] == ticker)
        rs = int(round(1 + ((rank - 1) / (len(ordered) - 1)) * 98))
        asset_returns = histories[ticker].pct_change().dropna().tail(60)
        benchmark_close = _close_from_download(data, benchmark)
        benchmark_returns = benchmark_close.pct_change().dropna().tail(60)
        paired = pd.concat([asset_returns.rename("asset"), benchmark_returns.rename("benchmark")], axis=1).dropna().tail(60)
        if len(paired) < 40 or float(np.var(paired["benchmark"], ddof=1)) == 0:
            return {"ok": False, "Pipeline": pipeline, "Benchmark": benchmark, "error": "Insufficient aligned 60-day benchmark history to calculate beta."}
        beta = float(np.cov(paired["asset"], paired["benchmark"], ddof=1)[0, 1] / np.var(paired["benchmark"], ddof=1))
        if rs >= 85 and beta < 1.10:
            signal = "🔥 EXPLOSIVE HOLY GRAIL DECOUPLING"
        elif rs >= 80:
            signal = "🚀 Outperforming Momentum Trend"
        elif rs <= 30:
            signal = "⚠️ Severe Capital Underperformance Track"
        else:
            signal = "Normal Structural Performance Market Track"
        return {"ok": True, "Pipeline": pipeline, "Benchmark": benchmark, "RS Percentile Rating": rs, "Recent 60-Day Beta": beta, "System Diagnostic Signal": signal, "Proxy peer observations": len(scores), "Weighted raw score": scores[ticker]}
    except Exception as exc:
        return {"ok": False, "Pipeline": pipeline, "Benchmark": benchmark, "error": f"RS/Beta provider request failed: {exc}"}


def _net_income_row(income: pd.DataFrame) -> pd.Series:
    for label in ("Net Income", "Net Income Common Stockholders", "Net Income Including Noncontrolling Interests", "Normalized Income"):
        if label in income.index:
            return pd.to_numeric(income.loc[label], errors="coerce")
    return pd.Series(dtype=float)


@st.cache_data(ttl=14_400, show_spinner=False)
def pat_and_market_price_table(symbol: str, income_json: str, market: str) -> pd.DataFrame:
    """Return five fiscal-year PAT growth and end-year share-price proxy.

    Product selling prices are not universally disclosed in structured feeds;
    the price field is explicitly a market-share-price proxy and never labelled
    as a product price.
    """
    income = _frame_from_cache(income_json)
    net_income = _net_income_row(income)
    if net_income.empty:
        return pd.DataFrame()
    try:
        history = yf.Ticker(symbol).history(period="6y", interval="1d", auto_adjust=True, actions=False)["Close"].dropna()
    except Exception:
        history = pd.Series(dtype=float)
    records = []
    annual_values = []
    for column, raw_value in net_income.items():
        try:
            end = pd.Timestamp(column)
        except (TypeError, ValueError):
            continue
        value = num(raw_value)
        if value is not None:
            annual_values.append((end, value))
    for end, value in sorted(annual_values, key=lambda item: item[0])[-5:]:
        if value is None:
            continue
        before = history.loc[:pd.Timestamp(end).tz_localize(history.index.tz) if getattr(history.index, "tz", None) else pd.Timestamp(end)] if not history.empty else pd.Series(dtype=float)
        close = float(before.iloc[-1]) if not before.empty else None
        records.append({"FY": str(end.year), "PAT (reported currency M)": value / 1_000_000, "Market share-price proxy": close})
    if not records:
        return pd.DataFrame()
    table = pd.DataFrame(records)
    table["PAT % increase"] = table["PAT (reported currency M)"].pct_change() * 100
    table["Market share-price % rise"] = table["Market share-price proxy"].pct_change() * 100
    table["Multiplier"] = table.apply(lambda row: row["PAT % increase"] / row["Market share-price % rise"] if pd.notna(row["PAT % increase"]) and pd.notna(row["Market share-price % rise"]) and row["Market share-price % rise"] != 0 else np.nan, axis=1)
    threshold = 2.0 if market == "India" else 1.5
    table["Multiplier status"] = table["Multiplier"].map(lambda value: "Green" if pd.notna(value) and value >= threshold else ("Amber" if pd.notna(value) else "Unavailable"))
    return table


def render() -> None:
    clocks = market_clock()
    a, b = st.columns(2)
    a.markdown(f"### United States: {clocks['US']['status']}")
    a.caption(clocks["US"]["time"])
    b.markdown(f"### India: {clocks['India']['status']}")
    b.caption(clocks["India"]["time"])
    st.header("Upcoming key statutory filing windows")
    us_deadlines, india_deadlines = st.columns(2)
    with us_deadlines:
        st.subheader("United States")
        st.dataframe(statutory_calendar("US"), width="stretch", hide_index=True)
    with india_deadlines:
        st.subheader("India")
        st.dataframe(statutory_calendar("India"), width="stretch", hide_index=True)

    with st.sidebar:
        st.header("Research controls")
        page = st.radio("Page", ["Research dashboard", "Watchlist"], label_visibility="collapsed")
        previous_market = st.session_state.market
        market = st.radio("Market", ["US", "India"], index=0 if st.session_state.market == "US" else 1)
        st.session_state.market = market
        dual_options = ["Custom ticker", *DUAL_LISTINGS.keys()]
        dual_choice = st.selectbox("Dual-listed company switch", dual_options, index=dual_options.index(st.session_state.dual_listing))
        if dual_choice != st.session_state.dual_listing:
            st.session_state.dual_listing = dual_choice
            if dual_choice != "Custom ticker":
                st.session_state.active_ticker = DUAL_LISTINGS[dual_choice][market]
        if market != previous_market and st.session_state.dual_listing != "Custom ticker":
            st.session_state.active_ticker = DUAL_LISTINGS[st.session_state.dual_listing][market]
        st.caption("Select a dual-listed company, then switch market to load its US ticker/ADR or Indian NSE ticker.")
        st.subheader("Company-name search")
        name_query = st.text_input("Search company name", placeholder="Example: Apple, Infosys, Reliance")
        possible_matches = company_name_matches(name_query, market)
        if name_query.strip():
            if possible_matches:
                labels = [f"{match['name']} — {match['symbol']} ({match['exchange']})" for match in possible_matches]
                selected_label = st.selectbox("Possible matches", labels)
                chosen_match = possible_matches[labels.index(selected_label)]
                if st.button("Load company match", width="stretch"):
                    st.session_state.active_ticker = display_symbol(chosen_match["symbol"])
                    st.session_state.dual_listing = "Custom ticker"
                    st.rerun()
            else:
                st.info("No matching listed instruments were returned by Yahoo Finance for this market.")
        universe = US_UNIVERSE if market == "US" else INDIA_UNIVERSE
        tier = st.selectbox("Company size", list(universe))
        selected = st.selectbox("Top-ten selection", universe[tier])
        entry = st.text_input("Ticker", value=display_symbol(st.session_state.active_ticker))
        if st.button("Load ticker", width="stretch"):
            st.session_state.active_ticker = entry.strip().upper() or selected
            st.session_state.dual_listing = "Custom ticker"
            st.rerun()
        if st.button("Load selection", width="stretch"):
            st.session_state.active_ticker = selected
            st.session_state.dual_listing = "Custom ticker"
            st.rerun()

    if page == "Watchlist":
        st.title("Watchlist")
        if not st.session_state.watchlist:
            st.info("No saved companies yet.")
            return
        saved = pd.DataFrame(st.session_state.watchlist)
        st.dataframe(saved, width="stretch", hide_index=True)
        target = st.selectbox("Open saved company", saved["Symbol"].unique())
        if st.button("Open deep dive"):
            row = saved.loc[saved.Symbol == target].iloc[0]
            st.session_state.active_ticker = display_symbol(row.Symbol)
            st.session_state.market = row.Market
            st.rerun()
        return

    symbol = symbol_for_market(st.session_state.active_ticker, market)
    if not symbol:
        st.error("Enter a ticker symbol.")
        return
    finnhub_token = secret("FINNHUB_API_KEY")
    if market == "US" and finnhub_token:
        quote = finnhub_quote(symbol, finnhub_token)
        price = {
            "history": pd.DataFrame(),
            "price": quote.get("price"),
            "previous": quote.get("previous"),
            "as_of": quote.get("as_of"),
            "error": quote.get("error"),
            "source": "Finnhub",
        }
    else:
        price = yahoo_price(symbol)
        price["source"] = "Yahoo Finance"
    # Quote and chart requests are deliberately separate: a successful Finnhub
    # quote must not leave the earlier Yahoo-only chart blank.
    chart_history = price_history(symbol, market, finnhub_token)
    if not chart_history.empty:
        price["history"] = chart_history
    fundamentals, browser_cache_hit = browser_cached_fundamentals(symbol)
    info = fundamentals["info"]
    sec_data = sec_company_data(symbol, secret("SEC_USER_AGENT")) if market == "US" else None
    currency = info.get("currency") or ("INR" if market == "India" else "USD")
    financial_currency = info.get("financialCurrency") or currency
    fx_factor, fx_note = usd_conversion_factor(financial_currency)
    company = info.get("longName") or info.get("shortName") or symbol
    st.title(company)
    cache_label = "browser cache (under four hours old)" if browser_cache_hit else "provider refresh; browser cache updated"
    st.caption(f"Symbol: {symbol} | Market: {market} | Price source: {price['source']} | Fundamentals: {cache_label} | Data shown only when returned by a provider.")
    price_cache_key = f"{market}:{symbol}"
    price_warning = None
    if price.get("error"):
        saved_price = st.session_state.last_good_prices.get(price_cache_key)
        if saved_price:
            price = {**price, **saved_price, "error": None, "source": f"cached {saved_price.get('source', 'quote')}"}
            price_warning = "Live quote is temporarily unavailable; displaying the last valid quote from this browser session."
        else:
            price_warning = price["error"]
    elif price.get("price") is not None:
        st.session_state.last_good_prices[price_cache_key] = {
            "price": price.get("price"),
            "previous": price.get("previous"),
            "as_of": price.get("as_of"),
            "history": price.get("history", pd.DataFrame()),
            "source": price.get("source"),
        }
    if price_warning:
        st.warning(price_warning)
    if st.button("➕ Add to watchlist", type="primary", key=f"watch-top-{market}-{symbol}"):
        if symbol not in [item["Symbol"] for item in st.session_state.watchlist]:
            st.session_state.watchlist.append({"Symbol": symbol, "Market": market, "Company": company, "Saved at UTC": dt.datetime.now(dt.timezone.utc).strftime("%Y-%m-%d %H:%M")})
            st.success("Added to watchlist.")
        else:
            st.info("Already in watchlist.")
    latest, previous = price.get("price"), price.get("previous")
    change = latest - previous if latest is not None and previous is not None else None
    last_filing = sec_data.get("last_filed") if sec_data else None
    st.header("1. Company ticker, exchange and latest available stock price")
    display_line_items([
        ("Ticker", symbol),
        ("Stock exchange", info.get("fullExchangeName") or info.get("exchange") or "Unavailable"),
        ("Latest available close", money(latest, currency)),
        ("Daily price change", f"{change:+,.2f}" if change is not None else "Unavailable"),
        ("Market capitalisation (absolute USD millions)", usd_millions(info.get("marketCap"), fx_factor)),
        ("Last official filing date", last_filing or "Unavailable"),
    ])
    if market == "India" and not last_filing:
        st.caption(f"Latest provider financial statement period: {latest_statement_period(fundamentals['income'])}. This is not represented as an NSE/BSE filing date; use the official exchange links in Section 4 for issuer filings.")
    price_chart(price["history"], symbol, currency)
    st.caption(f"All absolute financial values below are presented in USD millions. FX basis: {fx_note}. Ratios are not converted.")
    st.header("2. Company full name")
    st.write(company)
    st.header("3. Company business and location")
    display_line_items([
        ("Country", info.get("country") or "Unavailable"),
        ("City", info.get("city") or "Unavailable"),
        ("Sector", info.get("sector") or "Unavailable"),
        ("Industry", info.get("industry") or "Unavailable"),
    ])
    if info.get("longBusinessSummary"):
        st.write(info["longBusinessSummary"])

    st.header("4. Corporate group, subsidiaries, associates and promoters")
    website = info.get("website")
    resolved = resolve_investor_relations(website) if website else {"website": None, "ir_url": None, "error": "Provider did not return an official company website."}
    if market == "US" and sec_data is not None and not sec_data.get("filings", pd.DataFrame()).empty:
        evidence_urls = tuple(sec_data["filings"]["Official SEC filing"].head(3).tolist())
    else:
        evidence_urls = tuple(url for url in (resolved.get("ir_url"), resolved.get("website")) if url)
    disclosure_evidence = deterministic_disclosure_evidence(evidence_urls, market, secret("SEC_USER_AGENT")) if evidence_urls else {section: [] for section in DISCLOSURE_KEYWORDS}
    supply_assessment = supply_demand_assessments(evidence_urls, market, secret("SEC_USER_AGENT")) if evidence_urls else {period: {"score": 0.0, "classification": "Insufficient public source text", "high": 0, "low": 0, "balanced": 0} for period in ("Past", "Present", "Future")}
    if resolved.get("website"):
        st.link_button("Open company website", resolved["website"])
    if resolved.get("ir_url"):
        st.link_button("Auto-Fetch: open verified Investor Relations page", resolved["ir_url"])
    else:
        st.warning(f"Investor Relations link unavailable. {resolved.get('error', '')}")
    st.markdown("<div class='amber-upload'>To populate corporate group, subsidiaries, associates, promoters and related-party information, upload the latest official annual report / 10-K / 20-F / annual return or a company investor presentation. The notes titled <em>Subsidiaries</em>, <em>Related Parties</em>, <em>Promoters / Shareholding</em>, or <em>Business Combinations</em> are most useful. Proxy block encountered or report link unavailable? Download the official report using the sources below, then drop it here for source-grounded analysis.</div>", unsafe_allow_html=True)
    st.subheader("Official public filing sources")
    if market == "US":
        if sec_data and sec_data.get("browse_url"):
            st.link_button("Open SEC EDGAR company filings", sec_data["browse_url"])
        elif sec_data:
            st.info(sec_data["error"])
    else:
        st.info("NSE corporate-data API access is a paid product. These official exchange pages are provided without scraping or bypassing access controls.")
    for label, url in official_exchange_links(symbol, market).items():
        st.link_button(label, url)
    uploaded = st.file_uploader("Official report upload", type=["pdf", "docx", "xlsx", "xls", "csv"], help="Only upload documents downloaded from the company or stock exchange website.")
    analysis = None
    if uploaded:
        document_text, document_error = extract_uploaded_document(uploaded)
        if document_error:
            st.error(document_error)
        else:
            st.success(f"Read {len(document_text):,} characters from {uploaded.name}.")
            if st.button("Analyze official report sequentially"):
                progress = st.progress(0, text="Analyzing Corporate Web (1/3)...")
                analysis = gemini_analysis(document_text, company, secret("GEMINI_API_KEY"), secret("GEMINI_MODEL"))
                if analysis.get("error"):
                    st.error(analysis["error"])
                else:
                    progress.progress(34, text="Analyzing Litigation Ledger (2/3)...")
                    progress.progress(67, text="Calculating Bottleneck State (3/3)...")
                    progress.progress(100, text="Analysis complete.")

    display_evidence("4. Group structure", disclosure_evidence, "No matching group-structure terms were found in the limited public source pages checked. Open the official filings above for complete disclosure.")
    if analysis and not analysis.get("error"):
        with st.expander("Optional uploaded-report analysis"):
            st.write(analysis["corporate_web"])

    # These placeholders preserve the visual chronology: 1 through 11.
    section_5_slot = st.container()
    section_6_slot = st.container()
    section_7_slot = st.container()

    st.header("8. Investor-report facts, future prospects, litigation and M&A")
    st.caption("Official filing/report text is used first. Gemini runs only against an official report you upload; it is never asked to invent historic, current, or future M&A.")
    display_evidence("8. Litigation, prospects and M&A", disclosure_evidence, "No matching litigation, prospect, or M&A terms were found in the limited public source pages checked.")
    if analysis and not analysis.get("error"):
        with st.expander("Optional uploaded-report analysis"):
            st.write(analysis["section_8"])
    else:
        st.info("Upload an annual report, 10-K/20-F, merger circular, or investor presentation in Section 4 and run the source-grounded analysis for disclosed prospects, litigation, contingent liabilities and M&A.")
    for label, url in ranking_source_links(symbol, market, company):
        st.link_button(f"Research source: {label}", url, key=f"ma-{label}")

    st.header("9. Current, past and announced partnerships")
    st.caption("Only partnerships, collaborations, joint ventures, or alliances explicitly disclosed in the reviewed source are shown. ‘Future’ means announced, not predicted.")
    display_evidence("9. Partnerships", disclosure_evidence, "No matching partnership terms were found in the limited public source pages checked.")
    if analysis and not analysis.get("error"):
        with st.expander("Optional uploaded-report analysis"):
            st.write(analysis["section_9"])

    st.header("10. Key personnel, expected joins and departures")
    st.caption("Personnel information is limited to appointments, resignations, retirements and planned changes stated in filings, official announcements or your uploaded official source.")
    display_evidence("10. Key personnel", disclosure_evidence, "No matching executive appointment or departure terms were found in the limited public source pages checked.")
    if analysis and not analysis.get("error"):
        with st.expander("Optional uploaded-report analysis"):
            st.write(analysis["section_10"])

    st.header("11. Past, present and forward bottleneck assessment")
    st.caption("This is a transparent, logarithmically normalised disclosed-language heuristic, not an investment recommendation or verified economic forecast.")
    st.subheader("Five-year PAT and market-price proxy multiplier")
    st.caption("PAT uses provider-reported net income. Product selling-price data is not a universal structured-data field, so the table uses a clearly labelled end-of-year market share-price proxy. Upload an official report if you need a disclosed product-price series. Green threshold: ≥1.5x for US; ≥2.0x for India.")
    current_pat_table = pat_and_market_price_table(symbol, _frame_to_cache(fundamentals["income"]), market)
    if current_pat_table.empty:
        st.info("The provider did not return five-year annual net-income data for a PAT multiplier table.")
    else:
        def multiplier_style(value: Any) -> str:
            return "background-color: #166534; color: white" if value == "Green" else ("background-color: #b45309; color: white" if value == "Amber" else "")
        st.dataframe(current_pat_table.style.map(multiplier_style, subset=["Multiplier status"]).format({"PAT (reported currency M)": "{:,.2f}", "Market share-price proxy": "{:,.2f}", "PAT % increase": "{:.2f}%", "Market share-price % rise": "{:.2f}%", "Multiplier": "{:.2f}x"}, na_rep="Unavailable"), width="stretch", hide_index=True)
    if st.session_state.dual_listing != "Custom ticker":
        other_market = "India" if market == "US" else "US"
        other_symbol = symbol_for_market(DUAL_LISTINGS[st.session_state.dual_listing][other_market], other_market)
        other_fundamentals = yahoo_fundamentals(other_symbol)
        other_table = pat_and_market_price_table(other_symbol, _frame_to_cache(other_fundamentals["income"]), other_market)
        st.subheader(f"Dual-listed reference: {other_market} ({other_symbol})")
        if other_table.empty:
            st.info("No usable five-year annual net-income history was returned for the other listing.")
        else:
            st.dataframe(other_table.style.map(multiplier_style, subset=["Multiplier status"]).format({"PAT (reported currency M)": "{:,.2f}", "Market share-price proxy": "{:,.2f}", "PAT % increase": "{:.2f}%", "Market share-price % rise": "{:.2f}%", "Multiplier": "{:.2f}x"}, na_rep="Unavailable"), width="stretch", hide_index=True)
    st.subheader("Relative Strength and 60-day covariance beta")
    rs_beta = calculate_stock_rs_and_beta_engine(symbol)
    if rs_beta.get("ok"):
        rs_col, beta_col, signal_col = st.columns(3)
        rs_col.metric("RS percentile rating", f"{rs_beta['RS Percentile Rating']} / 99")
        beta_col.metric("Recent 60-day beta", f"{rs_beta['Recent 60-Day Beta']:.2f}")
        signal_col.metric("System diagnostic signal", rs_beta["System Diagnostic Signal"])
        st.caption(f"{rs_beta['Pipeline']} • Benchmark: {rs_beta['Benchmark']} • Proxy peer observations: {rs_beta['Proxy peer observations']} • 40/20/20/20 quarterly weighted return score: {rs_beta['Weighted raw score']:.2%}. The percentile is against this transparent high-cap proxy pool, not the complete market.")
    else:
        st.info(f"RS/Beta engine unavailable: {rs_beta.get('error', 'Unknown provider error')} ({rs_beta.get('Pipeline', 'routing unavailable')}).")
    with st.expander("Definitions used by this scale"):
        display_line_items([(label, definition) for label, definition in BOTTLENECK_DEFINITIONS.items()])
    past_column, present_column, future_column = st.columns(3)
    with past_column:
        supply_demand_gauge("Past", supply_assessment["Past"])
    with present_column:
        supply_demand_gauge("Present", supply_assessment["Present"])
    with future_column:
        supply_demand_gauge("Future", supply_assessment["Future"])
    st.subheader("Why this reading appears")
    brief_state_key = f"supply-brief:{symbol}"
    if brief_state_key not in st.session_state:
        st.session_state[brief_state_key] = deterministic_supply_brief(supply_assessment, disclosure_evidence)
    if st.button("Refresh this explanation with Gemini", key=f"gemini-supply-{symbol}"):
        news_context = []
        if market == "US":
            finnhub_context = finnhub_profile_and_news(symbol, secret("FINNHUB_API_KEY"))
            news_context = finnhub_context.get("news", []) if not finnhub_context.get("error") else []
        ai_brief, ai_error = ai_supply_brief(company, supply_assessment, disclosure_evidence, secret("GEMINI_API_KEY"), secret("GEMINI_MODEL"), news_context)
        if ai_brief:
            st.session_state[brief_state_key] = ai_brief
        elif ai_error:
            st.info(ai_error)
    st.write(st.session_state[brief_state_key])
    display_evidence("11. Supply-demand evidence", disclosure_evidence, "No matching supply-demand, capacity, inventory, or geopolitical terms were found in the limited public source pages checked.")
    if analysis and not analysis.get("error"):
        with st.expander("Optional uploaded-report analysis"):
            st.write(analysis["bottleneck"])

    with section_5_slot:
        st.header("5. Market and industry position")
        st.info("Live rankings are provider-owned and can use different reporting periods. The requested CSIMarket/Screener pages are linked below. The dashboard will not invent a numeric rank when a provider does not return a verified company match.")
        display_evidence("5. Market and industry position", disclosure_evidence, "No market-position terms were found in the limited public source pages checked. A numeric industry rank is not inferred.")
        source_rows = pd.DataFrame([{"Provider": label, "Live ranking / company page": url, "Status": "Open provider page for latest published ranking"} for label, url in ranking_source_links(symbol, market, company)])
        st.dataframe(source_rows, width="stretch", hide_index=True, column_config={"Live ranking / company page": st.column_config.LinkColumn("Live ranking / company page")})
        for label, url in official_exchange_links(symbol, market).items():
            st.link_button(f"Open source: {label}", url, key=f"comparison-{label}")
    section_7_slot.header("7. Company financial ratios, P&L, balance sheet and cash flow")
    metrics = pd.DataFrame({"Company metric": ["Trailing P/E", "Forward P/E", "Price / book", "Return on equity", "Debt / equity", "Current ratio", "Quick ratio"], "Latest reported value": [ratio(info.get("trailingPE")), ratio(info.get("forwardPE")), ratio(info.get("priceToBook")), ratio(info.get("returnOnEquity"), True), ratio(info.get("debtToEquity")), ratio(info.get("currentRatio")), ratio(info.get("quickRatio"))]})
    section_7_slot.dataframe(metrics, width="stretch", hide_index=True)
    if market == "US":
        section_7_slot.subheader("SEC Company Facts: latest annual GAAP/IFRS XBRL facts (USD millions)")
        if sec_data and sec_data.get("error"):
            section_7_slot.info(sec_data["error"])
        elif sec_data is not None:
            section_7_slot.dataframe(sec_data["facts"], width="stretch", hide_index=True)
            for statement_name, statement_frame in sec_data.get("statements", {}).items():
                section_7_slot.subheader(f"SEC XBRL annual {statement_name.lower()} (USD millions)")
                section_7_slot.dataframe(statement_frame, width="stretch", hide_index=True)
            section_7_slot.subheader("Recent official SEC filings")
            section_7_slot.dataframe(sec_data["filings"], width="stretch", hide_index=True, column_config={"Official SEC filing": st.column_config.LinkColumn("Official SEC filing")})
    with section_7_slot:
        statement_view(fundamentals["income"], "Annual income statement / P&L (USD millions)", fx_factor)
        statement_view(fundamentals["balance"], "Annual balance sheet (USD millions)", fx_factor)
        statement_view(fundamentals["cashflow"], "Annual cash-flow statement (USD millions)", fx_factor)
    with section_7_slot.expander("Quarterly statements"):
        statement_view(fundamentals["quarterly_income"], "Quarterly income statement (USD millions)", fx_factor)
        statement_view(fundamentals["quarterly_balance"], "Quarterly balance sheet (USD millions)", fx_factor)
        statement_view(fundamentals["quarterly_cashflow"], "Quarterly cash flow (USD millions)", fx_factor)

    section_6_slot.header("6. AGM, business meets, announcements, news and social sources")
    if analysis and not analysis.get("error"):
        section_6_slot.subheader("Official-report disclosures")
        section_6_slot.write(analysis["section_6"])
    else:
        section_6_slot.info("AGM and business-meet disclosures from reports require an official upload and Gemini analysis.")
    if market == "US":
        finnhub = finnhub_profile_and_news(symbol, secret("FINNHUB_API_KEY"))
        if finnhub["error"]:
            section_6_slot.info(finnhub["error"])
        for story in finnhub["news"][:10]:
            section_6_slot.markdown(f"- [{story.get('headline', 'Untitled')}]({story.get('url', '')}) — {dt.datetime.fromtimestamp(story.get('datetime', 0), dt.timezone.utc).strftime('%d %b %Y')}")
    else:
        section_6_slot.info("For India, use the official company Investor Relations link above and exchange disclosures. Finnhub's company-news endpoint is configured here for US symbols only.")
    if website:
        section_6_slot.link_button("Search company posts on X", f"https://x.com/search?q={company.replace(' ', '%20')}&src=typed_query")
    section_6_slot.caption("X/Twitter posts are not ingested or analysed without an authorised X API plan. Links are provided instead of unauthenticated scraping.")


if __name__ == "__main__":
    render()
