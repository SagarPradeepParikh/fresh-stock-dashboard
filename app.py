"""Evidence-first India/US equity research dashboard for Streamlit Cloud."""
from __future__ import annotations

import datetime as dt
import base64
import hashlib
import io
import json
import re
from dataclasses import dataclass
from typing import Any
from urllib.parse import urljoin, urlparse
from zoneinfo import ZoneInfo

import pandas as pd
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
background_css = (
    f"background-image: linear-gradient(rgba(11,15,25,.88), rgba(11,15,25,.88)), url('{background_image}');"
    if background_image else "background-color: #0b0f19;"
)

st.markdown("""<style>
.stApp { """ + background_css + """ background-attachment: fixed; background-size: cover; background-position: center; color: #f3f4f6; }
section[data-testid="stSidebar"] { background: #111827; }
.amber-upload { border: 2px dashed #f59e0b; border-radius: 10px; padding: 12px; background: #3a2b11; }
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

for key, value in {"watchlist": [], "active_ticker": "AAPL", "market": "US", "open_ir": False, "dual_listing": "Custom ticker"}.items():
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
    return pd.DataFrame(rows, columns=["Key statutory filing", "Indicative next due date", "Scope / qualification"])


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


@st.cache_data(ttl=300, show_spinner=False)
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
    "Revenue": ("RevenueFromContractWithCustomerExcludingAssessedTax", "Revenues", "SalesRevenueNet"),
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


def _latest_sec_fact(facts: dict[str, Any], concepts: tuple[str, ...]) -> float | None:
    """Get a latest annual USD fact, preserving the SEC filing's reported value."""
    for concept in concepts:
        fact = facts.get("us-gaap", {}).get(concept, {})
        units = fact.get("units", {}).get("USD", [])
        annual = [item for item in units if item.get("form") in {"10-K", "20-F", "40-F"} and item.get("fy")]
        candidates = annual or units
        if candidates:
            item = max(candidates, key=lambda row: (row.get("end", ""), row.get("filed", "")))
            return num(item.get("val"))
    return None


@st.cache_data(ttl=900, show_spinner=False)
def sec_company_data(symbol: str, user_agent: str | None) -> dict[str, Any]:
    """Read public EDGAR filings and GAAP/IFRS-derived XBRL facts directly.

    SEC data is free and needs no API key.  The SEC requires a descriptive
    User-Agent with contact details, so the request is not attempted without it.
    """
    if not user_agent or "@" not in user_agent:
        return {"error": "Set SEC_USER_AGENT in Streamlit secrets, for example: Your App Name contact@example.com.", "facts": pd.DataFrame(), "filings": pd.DataFrame(), "browse_url": None, "last_filed": None}
    headers = {"User-Agent": user_agent, "Accept-Encoding": "gzip, deflate", "Host": "data.sec.gov"}
    try:
        tickers_response = requests.get("https://www.sec.gov/files/company_tickers.json", headers={"User-Agent": user_agent}, timeout=15)
        tickers_response.raise_for_status()
        matches = [record for record in tickers_response.json().values() if record.get("ticker", "").upper() == symbol.upper()]
        if not matches:
            return {"error": "This ticker was not found in the SEC's company ticker file.", "facts": pd.DataFrame(), "filings": pd.DataFrame(), "browse_url": None, "last_filed": None}
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
        return {"error": None, "facts": pd.DataFrame(rows), "filings": pd.DataFrame(filing_rows), "browse_url": f"https://www.sec.gov/edgar/browse/?CIK={cik}", "company_name": facts_json.get("entityName"), "last_filed": filing_rows[0]["Filed"] if filing_rows else None}
    except (requests.RequestException, ValueError, KeyError) as exc:
        return {"error": f"SEC public-data request failed: {exc}", "facts": pd.DataFrame(), "filings": pd.DataFrame(), "browse_url": None, "last_filed": None}


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


@st.cache_data(ttl=3600, show_spinner=False)
def resolve_investor_relations(website: str) -> dict[str, str | None]:
    """Find an investor-relations link only by crawling the company's reported official site."""
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
        return {"website": website, "ir_url": None, "error": "No investor-relations link was found on the official homepage."}
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
    st.dataframe(visible, use_container_width=True)


def price_chart(history: pd.DataFrame, symbol: str, currency: str) -> None:
    if history.empty:
        return
    figure = go.Figure(go.Scatter(x=history.index, y=history["Close"], mode="lines", line={"color": "#38bdf8"}, name="Close"))
    figure.update_layout(title=f"{symbol}: one-year closing-price history", paper_bgcolor="#0b0f19", plot_bgcolor="#111827", font={"color": "#f3f4f6"}, height=380, margin={"l": 10, "r": 10, "t": 45, "b": 10})
    figure.update_yaxes(title=f"Price ({currency})", gridcolor="#374151")
    st.plotly_chart(figure, use_container_width=True)


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
        st.dataframe(statutory_calendar("US"), use_container_width=True, hide_index=True)
    with india_deadlines:
        st.subheader("India")
        st.dataframe(statutory_calendar("India"), use_container_width=True, hide_index=True)

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
        universe = US_UNIVERSE if market == "US" else INDIA_UNIVERSE
        tier = st.selectbox("Company size", list(universe))
        selected = st.selectbox("Top-ten selection", universe[tier])
        entry = st.text_input("Ticker", value=display_symbol(st.session_state.active_ticker))
        if st.button("Load ticker", use_container_width=True):
            st.session_state.active_ticker = entry.strip().upper() or selected
            st.session_state.dual_listing = "Custom ticker"
            st.rerun()
        if st.button("Load selection", use_container_width=True):
            st.session_state.active_ticker = selected
            st.session_state.dual_listing = "Custom ticker"
            st.rerun()

    if page == "Watchlist":
        st.title("Watchlist")
        if not st.session_state.watchlist:
            st.info("No saved companies yet.")
            return
        saved = pd.DataFrame(st.session_state.watchlist)
        st.dataframe(saved, use_container_width=True, hide_index=True)
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
    price = yahoo_price(symbol)
    fundamentals, browser_cache_hit = browser_cached_fundamentals(symbol)
    info = fundamentals["info"]
    sec_data = sec_company_data(symbol, secret("SEC_USER_AGENT")) if market == "US" else None
    currency = info.get("currency") or ("INR" if market == "India" else "USD")
    financial_currency = info.get("financialCurrency") or currency
    fx_factor, fx_note = usd_conversion_factor(financial_currency)
    company = info.get("longName") or info.get("shortName") or symbol
    st.title(company)
    cache_label = "browser cache (under four hours old)" if browser_cache_hit else "provider refresh; browser cache updated"
    st.caption(f"Symbol: {symbol} | Market: {market} | Source: Yahoo Finance | Fundamentals: {cache_label} | Data shown only when returned by a provider.")
    if price["error"]:
        st.error(price["error"])
        return
    latest, previous = price.get("price"), price.get("previous")
    change = latest - previous if latest is not None and previous is not None else None
    last_filing = sec_data.get("last_filed") if sec_data else None
    st.header("1. Company ticker, exchange and latest available stock price")
    m1, m2, m3, m4, m5, m6 = st.columns(6)
    m1.metric("Latest available close", money(latest, currency), f"{change:+,.2f}" if change is not None else None)
    m2.metric("Trailing P/E", ratio(info.get("trailingPE")))
    m3.metric("Forward P/E", ratio(info.get("forwardPE")))
    m4.metric("Price / book", ratio(info.get("priceToBook")))
    m5.metric("Market cap (USD M)", usd_millions(info.get("marketCap"), fx_factor))
    m6.metric("Last official filing date", last_filing or "Unavailable")
    if market == "India" and not last_filing:
        st.caption(f"Latest provider financial statement period: {latest_statement_period(fundamentals['income'])}. This is not represented as an NSE/BSE filing date; use the official exchange links in Section 4 for issuer filings.")
    price_chart(price["history"], symbol, currency)
    st.caption(f"All absolute financial values below are presented in USD millions. FX basis: {fx_note}. Ratios are not converted.")
    if st.button("Add to watchlist"):
        if symbol not in [item["Symbol"] for item in st.session_state.watchlist]:
            st.session_state.watchlist.append({"Symbol": symbol, "Market": market, "Company": company, "Saved at UTC": dt.datetime.now(dt.timezone.utc).strftime("%Y-%m-%d %H:%M")})
            st.success("Added to watchlist.")
        else:
            st.info("Already in watchlist.")

    st.header("2. Company full name")
    st.write(company)
    st.header("3. Company business and location")
    st.write({"Exchange": info.get("fullExchangeName") or info.get("exchange") or "Unavailable", "Country": info.get("country") or "Unavailable", "City": info.get("city") or "Unavailable", "Sector": info.get("sector") or "Unavailable", "Industry": info.get("industry") or "Unavailable"})
    if info.get("longBusinessSummary"):
        st.write(info["longBusinessSummary"])

    st.header("4. Corporate group, subsidiaries, associates and promoters")
    website = info.get("website")
    resolved = resolve_investor_relations(website) if website else {"website": None, "ir_url": None, "error": "Provider did not return an official company website."}
    if resolved.get("website"):
        st.link_button("Open company website", resolved["website"])
    if resolved.get("ir_url"):
        st.link_button("Auto-Fetch: open verified Investor Relations page", resolved["ir_url"])
    else:
        st.warning(f"Investor Relations link unavailable. {resolved.get('error', '')}")
    st.markdown("<div class='amber-upload'>Proxy block encountered or report link unavailable? Download the official report using the link above, then drop it here for source-grounded analysis.</div>", unsafe_allow_html=True)
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

    if analysis and not analysis.get("error"):
        st.write(analysis["corporate_web"])
    else:
        st.info("Upload an official company or exchange report and run source-grounded analysis to populate this section.")

    st.header("8. Investor-report facts, future prospects, litigation and M&A")
    if analysis and not analysis.get("error"):
        st.write(analysis["section_8"])
    else:
        st.info("This section requires a readable official report upload and configured Gemini analysis.")

    st.header("9. Current, past and announced partnerships")
    if analysis and not analysis.get("error"):
        st.write(analysis["section_9"])
    else:
        st.info("This section requires a readable official report upload and configured Gemini analysis.")

    st.header("10. Key personnel, expected joins and departures")
    if analysis and not analysis.get("error"):
        st.write(analysis["section_10"])
    else:
        st.info("This section requires a readable official report upload and configured Gemini analysis.")

    st.header("11. Past, present and forward bottleneck assessment")
    if analysis and not analysis.get("error"):
        st.write(analysis["bottleneck"])
    else:
        st.info("This section requires a readable official report upload and configured Gemini analysis.")

    st.header("5. Market and industry position")
    st.info("Public exchange sites provide listings and index context, but they do not expose a complete free, machine-readable peer-ranking universe. This dashboard therefore links the authoritative source pages and does not fabricate rank or peer averages.")
    for label, url in official_exchange_links(symbol, market).items():
        st.link_button(f"Open source: {label}", url, key=f"comparison-{label}")
    st.header("7. Company financial ratios, P&L, balance sheet and cash flow")
    metrics = pd.DataFrame({"Company metric": ["Trailing P/E", "Forward P/E", "Price / book", "Return on equity", "Debt / equity", "Current ratio", "Quick ratio"], "Latest reported value": [ratio(info.get("trailingPE")), ratio(info.get("forwardPE")), ratio(info.get("priceToBook")), ratio(info.get("returnOnEquity"), True), ratio(info.get("debtToEquity")), ratio(info.get("currentRatio")), ratio(info.get("quickRatio"))]})
    st.dataframe(metrics, use_container_width=True, hide_index=True)
    if market == "US":
        st.subheader("SEC Company Facts: latest annual GAAP/IFRS XBRL facts (USD millions)")
        if sec_data and sec_data.get("error"):
            st.info(sec_data["error"])
        elif sec_data is not None:
            st.dataframe(sec_data["facts"], use_container_width=True, hide_index=True)
            st.subheader("Recent official SEC filings")
            st.dataframe(sec_data["filings"], use_container_width=True, hide_index=True, column_config={"Official SEC filing": st.column_config.LinkColumn("Official SEC filing")})
    statement_view(fundamentals["income"], "Annual income statement / P&L (USD millions)", fx_factor)
    statement_view(fundamentals["balance"], "Annual balance sheet (USD millions)", fx_factor)
    statement_view(fundamentals["cashflow"], "Annual cash-flow statement (USD millions)", fx_factor)
    with st.expander("Quarterly statements"):
        statement_view(fundamentals["quarterly_income"], "Quarterly income statement (USD millions)", fx_factor)
        statement_view(fundamentals["quarterly_balance"], "Quarterly balance sheet (USD millions)", fx_factor)
        statement_view(fundamentals["quarterly_cashflow"], "Quarterly cash flow (USD millions)", fx_factor)

    st.header("6. AGM, business meets, announcements, news and social sources")
    if analysis and not analysis.get("error"):
        st.subheader("Official-report disclosures")
        st.write(analysis["section_6"])
    else:
        st.info("AGM and business-meet disclosures from reports require an official upload and Gemini analysis.")
    if market == "US":
        finnhub = finnhub_profile_and_news(symbol, secret("FINNHUB_API_KEY"))
        if finnhub["error"]:
            st.info(finnhub["error"])
        for story in finnhub["news"][:10]:
            st.markdown(f"- [{story.get('headline', 'Untitled')}]({story.get('url', '')}) — {dt.datetime.fromtimestamp(story.get('datetime', 0), dt.timezone.utc).strftime('%d %b %Y')}")
    else:
        st.info("For India, use the official company Investor Relations link above and exchange disclosures. Finnhub's company-news endpoint is configured here for US symbols only.")
    if website:
        st.link_button("Search company posts on X", f"https://x.com/search?q={company.replace(' ', '%20')}&src=typed_query")
    st.caption("X/Twitter posts are not ingested or analysed without an authorised X API plan. Links are provided instead of unauthenticated scraping.")


if __name__ == "__main__":
    render()
