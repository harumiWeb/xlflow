# stock-price

`stock-price` is an Excel VBA sample created and updated by an AI agent using xlflow.
It retrieves stock and cryptocurrency information from the Twelve Data API and displays it in a dashboard format.

## Features

- Retrieves prices for AAPL / MSFT / BTC/USD / ETH/USD
- Displays current price, daily change, weekly percentage change, and range information
- Shows key indicators in a card format
- Automatically generates lists and charts on the sheet
- Caches API responses for a short time to reduce load during consecutive executions

## Requirements

- Windows
- Microsoft Excel
- xlflow
- `TWELVE_DATA_API_KEY` environment variable
- Internet connection

## How to Run

1. Set `TWELVE_DATA_API_KEY` as an environment variable.
2. Open `build/Book.xlsm`.
3. Execute `Main.Run`.
   If you want to launch it from a button, you can use `Ui.RunFromButton`.

When running from xlflow, use `xlflow run --diagnostic` in this folder.

## Customizable Areas

- `src/modules/App.bas`
  - Contains the API key reading logic and the entry point for running the dashboard.
- `src/modules/MarketDashboard.bas`
  - Defines the monitored tickers, display layout, and chart configuration.
- `src/modules/TwelveDataClient.bas`
  - Handles access to the Twelve Data API and caching.
- `src/modules/Ui.bas`
  - A thin wrapper for buttons.

## Notes

- `build/Book.xlsm` is the distributed file.
- `.xlflow/` contains the execution state and backups.
- The `ApiCache` sheet is for internal caching and is typically hidden.
