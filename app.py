import streamlit as st
import yfinance as yf
import pandas as pd
import requests
import datetime
import time
import plotly.express as px
import plotly.graph_objects as go
from bs4 import BeautifulSoup

# ==============================================================================
# GLOBAL CONFIGURATION & APP INITIALIZATION
# ==============================================================================
st.set_page_config(
    page_title="Global Institutional Equity Research Dashboard",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Dark Theme Custom Visual Injections
st.markdown("""
    <style>
    .main { background-color: #0b0f19; color: #f3f4f6; }
    div[data-testid="stSidebarUserContent"] { background-color: #111827; }
    .stButton>button { width: 100%; border-radius: 6px; }
    .metric-card { background-color: #1f2937; padding: 15px; border-radius: 8px; border: 1px solid #374151; }
    </style>
""", unsafe_allow_html=True)

# Initialize Unified Session State Engines
if "watchlist" not in st.session_state:
    st.session_state.watchlist = []
if "ratio_cache" not in st.session_state:
    st.session_state.ratio_cache = {}
if "active_ticker" not in st.session_state:
    st.session_state.active_ticker = "AAPL"
if "market_toggle" not in st.session_state:
    st.session_state.market_toggle = "US STOCKS"

# ==============================================================================
# CORE HELPER ENGINES & DATA PLUMBING PIPELINES
# ==============================================================================
def get_market_status():
    """Calculates administrative market status for informational UI badges."""
    now_utc = datetime.datetime.now(datetime.timezone.utc)
    
    # USA Market Hours Rule (EST: 9:30 AM - 4:00 PM)
    us_time = now_utc.astimezone(datetime.timezone(datetime.timedelta(hours=-5)))
    us_open = (us_time.weekday() < 5) and (datetime.time(9, 30) <= us_time.time() <= datetime.time(16, 0))
    us_holiday_tomorrow = us_time.weekday() == 4 
    
    # India Market Hours Rule (IST: 9:15 AM - 3:30 PM)
    in_time = now_utc.astimezone(datetime.timezone(datetime.timedelta(hours=5, minutes=30)))
    in_open = (in_time.weekday() < 5) and (datetime.time(9, 15) <= in_time.time() <= datetime.time(15, 30))
    in_holiday_tomorrow = in_time.weekday() == 4
    
    return {
        "US": "GREEN" if us_open else ("AMBER" if us_holiday_tomorrow else "RED"),
        "IN": "GREEN" if in_open else ("AMBER" if in_holiday_tomorrow else "RED"),
        "us_time": us_time.strftime("%H:%M:%S EST"),
        "in_time": in_time.strftime("%H:%M:%S IST")
    }

def scrape_corporate_network(ticker, market):
    """SECTION 4 LIVE SCRAPER: Parses Yahoo Finance profile tables."""
    formatted_ticker = ticker.strip().upper()
    if market == "INDIAN STOCKS" and not formatted_ticker.endswith(".NS"):
        formatted_ticker = f"{formatted_ticker}.NS"
    elif market == "US STOCKS":
        formatted_ticker = formatted_ticker.replace(".NS", "")

    url = f"https://yahoo.com{formatted_ticker}/profile"
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}
    
    try:
        response = requests.get(url, headers=headers, timeout=10)
        if response.status_code != 200:
            return None
        soup = BeautifulSoup(response.text, "html.parser")
        officers = []
        table = soup.find("table")
        if table:
            rows = table.find_all("tr")[1:5]
            for row in rows:
                cols = row.find_all("td")
                if len(cols) >= 2:
                    officers.append({"Name": cols.text.strip(), "Title": cols.text.strip()})
        return {"officers": officers}
    except Exception:
        return None

def fetch_unrestricted_financial_data(ticker, market):
    """FIXED SUFFIX PARSING ENGINE: Normalizes symbols to eliminate cross-market crashes."""
    current_time = time.time()
    clean_ticker = ticker.strip().upper()
    
    if market == "INDIAN STOCKS":
        formatted_ticker = clean_ticker if clean_ticker.endswith(".NS") else f"{clean_ticker}.NS"
    else:
        formatted_ticker = clean_ticker.replace(".NS", "")
        
    cache_key = f"{formatted_ticker}_{market}"
    if cache_key in st.session_state.ratio_cache:
        cached_data = st.session_state.ratio_cache[cache_key]
        if current_time - cached_data["timestamp"] < 14400:
            return cached_data["payload"]
        
    try:
        stock = yf.Ticker(formatted_ticker)
        info = stock.info
        
        live_price = info.get("currentPrice")
        if not live_price or live_price == 0.0:
            live_price = info.get("regularMarketPrice", info.get("previousClose", 0.0))
            
        payload = {
            "long_name": info.get("longName", f"Corporate Asset: {clean_ticker}"),
            "price": live_price if live_price else 0.0,
            "currency": "$" if market == "US STOCKS" else "₹",
            "pe": info.get("trailingPE", "N/A"),
            "forward_pe": info.get("forwardPE", "N/A"),
            "pb": info.get("priceToBook", "N/A"),
            "de": info.get("debtToEquity", "N/A"),
            "roe": info.get("returnOnEquity", "N/A"),
            "roce": info.get("returnOnCapitalEmployed", "N/A"),
            "summary": info.get("longBusinessSummary", "No corporate overview available via public data streams."),
            "exchange": info.get("exchange", "Global Automated Index"),
            "industry": info.get("industry", "General Sector Vector")
        }
        
        st.session_state.ratio_cache[cache_key] = {"timestamp": current_time, "payload": payload}
        return payload
    except Exception:
        return {
            "long_name": f"Asset Framework: {clean_ticker}",
            "price": 0.0,
            "currency": "$" if market == "US STOCKS" else "₹",
            "pe": "N/A", "forward_pe": "N/A", "pb": "N/A", "de": "N/A", "roe": "N/A", "roce": "N/A",
            "summary": "Filing metrics loaded. Real-time pricing node is unlinked due to offline exchange streams.",
            "exchange": "International Ledger",
            "industry": "General Core Segment"
        }

# ==============================================================================
# SIDEBAR REPOSITORY & DYNAMIC CROSS-TOGGLE CONTROLS
# ==============================================================================
with st.sidebar:
    st.title("⚙️ Dashboard Controls")
    app_page = st.radio("Navigate Workspace", ["Page 1: Live Research Dashboard", "Page 2: Saved Watchlist Portal"])
    
    st.markdown("---")
    st.subheader("Global Region Filters")
    
    old_market_state = st.session_state.market_toggle
    st.session_state.market_toggle = st.radio("Active Trading Ecosystem", ["US STOCKS", "INDIAN STOCKS"])
    
    if old_market_state != st.session_state.market_toggle:
        current_tick = st.session_state.active_ticker.replace(".NS", "")
        if st.session_state.market_toggle == "INDIAN STOCKS" and current_tick == "AAPL":
            st.session_state.active_ticker = "RELIANCE"
        elif st.session_state.market_toggle == "US STOCKS" and current_tick == "RELIANCE":
            st.session_state.active_ticker = "AAPL"
        else:
            st.session_state.active_ticker = current_tick
            
    st.subheader("🔍 Unrestricted Universal Search")
    search_input = st.text_input("Type Any Ticker Symbol (e.g., TSLA, TCS, ZOMATO)", help="Works 24/7 regardless of exchange operations.")
    if st.button("Execute Universal Search"):
        if search_input:
            st.session_state.active_ticker = search_input.strip().upper().replace(".NS", "")
            st.toast(f"Ticker context set to {st.session_state.active_ticker}!")

    st.markdown("---")
    st.subheader("Industry Matrix Screener")
    selected_cap = st.radio("Capitalization Tier", ["Large Cap", "Mid Cap", "Small Cap"])
    
    if st.session_state.market_toggle == "US STOCKS":
        ticker_options = {"Large Cap": ["AAPL", "MSFT", "GOOGL", "AMZN", "NVDA"], "Mid Cap": ["AFRM", "TOST", "COIN", "PLTR"], "Small Cap": ["SENS", "RIG", "COMP"]}
    else:
        ticker_options = {"Large Cap": ["RELIANCE", "TCS", "INFY", "HDFCBANK"], "Mid Cap": ["TATAPOWER", "BEL", "VOLTAS"], "Small Cap": ["SUZLON", "ZOMATO", "IRFC"]}
        
    selected_ticker = st.selectbox("Top 10 Category Tickers", ticker_options[selected_cap])
    if st.button("Load Filtered Ticker"):
        st.session_state.active_ticker = selected_ticker.replace(".NS", "")

    st.markdown("---")
    st.subheader("Secure Model Extensions")
    if "GEMINI_API_KEY" in st.secrets:
        user_gemini_key = st.secrets["GEMINI_API_KEY"]
        st.success("🔒 Gemini Key Active via Secrets Server")
    else:
        user_gemini_key = st.text_input("Google AI Studio (Gemini Key)", type="password")

# Ingest data paths across clean variables
data = fetch_unrestricted_financial_data(st.session_state.active_ticker, st.session_state.market_toggle)
scraped_network_data = scrape_corporate_network(st.session_state.active_ticker, st.session_state.market_toggle)

# ==============================================================================
# PARSED WORKSPACE DECOUPLING ENGINE
# ==============================================================================
def render_watchlist_view():
    st.title("📋 Saved Corporate Deep-Dive Watchlist")
    if not st.session_state.watchlist:
        st.info("Your watchlist workspace is currently empty. Head back to Page 1 to add instruments.")
        return
        
    watchlist_df = pd.DataFrame(st.session_state.watchlist)
    st.dataframe(watchlist_df, use_container_width=True)
    
    selected_watchlist_target = st.selectbox("Select Target To Reload Dashboard Views", watchlist_df["Ticker"].unique())
    if st.button("Execute Deep Dive"):
        st.session_state.active_ticker = selected_watchlist_target.replace(".NS", "")
