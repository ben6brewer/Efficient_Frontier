from data_fetch import get_ticker_data
from logger import Logger
from gui import run_gui


def main():
    logger = Logger("main")
    tickers = ["AAPL", "GOOGL", "MSFT", "BTC-USD", "ETH-USD"]

    for ticker in tickers:
        data = get_ticker_data(ticker)
        logger.debug(f"{ticker}:\n{data.head()}")
    
    run_gui()


if __name__ == "__main__":
    main()