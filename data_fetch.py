import pandas as pd
import yfinance as yf
import os
from datetime import date, timedelta
from dotenv import load_dotenv
from logger import Logger

load_dotenv()

logger = Logger("data_fetch")
DATA_DIRECTORY = os.getenv("DATA_DIRECTORY", "data/")

def get_last_trading_day() -> date:
    today = date.today()
    weekday = today.weekday()
    if weekday == 6:  # Sunday
        return today - timedelta(days=2)
    elif weekday == 5:  # Saturday
        return today - timedelta(days=1)
    else:
        return today


def is_crypto(ticker: str) -> bool:
    return ticker.endswith("-USD")


def should_update(ticker: str) -> bool:
    parquet_path = DATA_DIRECTORY + ticker + ".parquet"
    if not os.path.exists(parquet_path): return True
    file_mod_date = date.fromtimestamp(os.path.getmtime(parquet_path))

    if is_crypto(ticker):
        return date.today() > file_mod_date
    else:
        last_trading_day = get_last_trading_day()
        return file_mod_date < last_trading_day

def pull_ticker_data(ticker: str) -> pd.DataFrame:
    parquet_path = DATA_DIRECTORY + ticker + ".parquet"
    logger.debug("Downloading fresh {ticker} data...")
    data = yf.download(ticker, period="max")
    logger.debug("Saving {ticker} data to {parquet_path}")
    data.to_parquet(parquet_path)
    return data

def get_ticker_data(ticker: str) -> pd.DataFrame:
    parquet_path = DATA_DIRECTORY + ticker + ".parquet"
    if should_update(ticker): return pull_ticker_data(ticker)
    return pd.read_parquet(parquet_path)
