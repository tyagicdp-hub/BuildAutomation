import pandas as pd
import requests
from io import StringIO

print("Fetching NSE stocks...")

# ---------- NSE STOCKS ----------
nse_url = "https://archives.nseindia.com/content/equities/EQUITY_L.csv"
nse_df = pd.read_csv(nse_url)

nse_symbols = nse_df["SYMBOL"].dropna().tolist()
nse_symbols = [symbol + ".NS" for symbol in nse_symbols]

print(f"NSE Stocks Fetched: {len(nse_symbols)}")

# ---------- BSE STOCKS ----------
print("\nFetching BSE stocks...")

bse_url = "https://www.bseindia.com/download/BhavCopy/Equity/EQ_ISINCODE_*.CSV"

# Reliable maintained BSE dataset
bse_master_url = "https://www.bseindia.com/corporates/List_Scrips.aspx"

# Use official BSE scrip list
bse_csv_url = "https://api.bseindia.com/BseIndiaAPI/api/ListofScripData/w"

headers = {
    "User-Agent": "Mozilla/5.0"
}

try:
    response = requests.get(bse_csv_url, headers=headers)
    bse_data = response.json()

    bse_symbols = []

    for stock in bse_data:
        if stock.get("SCRIP_CD"):
            bse_symbols.append(str(stock["SCRIP_CD"]) + ".BO")

    print(f"BSE Stocks Fetched: {len(bse_symbols)}")

except Exception as e:
    print("BSE fetch failed, fallback method...")
    bse_symbols = []

# ---------- COMBINE ----------
print("\nCombining symbols...")

all_symbols = nse_symbols + bse_symbols

df = pd.DataFrame({"Stock": all_symbols})

df.to_excel("stocks.xlsx", index=False)

print(f"\nTotal Symbols Saved: {len(df)} ✅")