import yfinance as yf
import pandas as pd
from datetime import date


# ---------- Scoring functions ----------

def momentum_score(hist):
    if hist.empty or len(hist) < 50:
        return 1

    start = hist["Close"].iloc[0]
    end = hist["Close"].iloc[-1]
    ret = ((end - start) / start) * 100

    dma200 = hist["Close"].rolling(200).mean().iloc[-1]

    if ret > 20 and end > dma200:
        return 10
    elif ret > 10:
        return 7
    elif ret > 0:
        return 4
    else:
        return 1


def quality_score(roe):
    if roe is None:
        return 4

    roe = roe * 100
    if roe > 20:
        return 10
    elif roe > 15:
        return 7
    elif roe > 10:
        return 4
    else:
        return 1


def valuation_score(pe):
    if pe is None:
        return 4

    if pe < 15:
        return 10
    elif pe < 25:
        return 7
    elif pe < 40:
        return 4
    else:
        return 1


def stability_score(de):
    if de is None:
        return 4

    if de < 0.3:
        return 10
    elif de < 0.7:
        return 7
    elif de < 1.0:
        return 4
    else:
        return 1

def sanity_checks(info, hist):
    flags = []

    # 1. Trend sanity (200 DMA)
    if not hist.empty and len(hist) >= 200:
        price = hist["Close"].iloc[-1]
        dma200 = hist["Close"].rolling(200).mean().iloc[-1]
        if price < dma200:
            flags.append("Below 200 DMA")

    # 2. Debt sanity
    debt = info.get("debtToEquity")
    if debt is not None and debt > 1:
        flags.append("High Debt")

    # 3. Valuation sanity
    pe = info.get("trailingPE")
    if pe is not None and pe > 50:
        flags.append("Very Expensive")

    return flags


# ---------- Main ranking logic ----------
def rank_stock(symbol):
    stock = yf.Ticker(symbol)
    info = stock.info
    hist = stock.history(period="6mo")

    rank = (
        momentum_score(hist) * 0.20 +
        quality_score(info.get("returnOnEquity")) * 0.20 +
        valuation_score(info.get("trailingPE")) * 0.20 +
        stability_score(info.get("debtToEquity")) * 0.15
    )

    flags = sanity_checks(info, hist)

    # Final decision logic
    if rank >= 8 and len(flags) == 0:
        action = "BUY"
    elif rank >= 6 and len(flags) <= 1:
        action = "WATCH"
    else:
        action = "IGNORE"

    return round(rank, 2), flags, action

# ---------- Read Excel ----------
df = pd.read_excel("stocks.xlsx")

results = []

for stock_symbol in df["Stock"]:
    try:
        rank, flags, action = rank_stock(stock_symbol)

        results.append({
            "Stock": stock_symbol,
            "Rank": rank,
            "Flags": ", ".join(flags) if flags else "None",
            "Action": action
        })

    except Exception as e:
        results.append({
            "Stock": stock_symbol,
            "Rank": None,
            "Flags": f"Error: {e}",
            "Action": "ERROR"
        })

# ---------- Convert to DataFrame ----------
result_df = pd.DataFrame(results)

# ---------- Save back to Excel ----------
result_df.to_excel("stocks.xlsx", index=False)

print("\nRanking Updated ✅")
print("Saved file: stocks.xlsx")