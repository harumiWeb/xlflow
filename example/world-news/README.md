# world-news

`world-news` is an Excel VBA sample managed with xlflow.  
It fetches the latest world news from NewsAPI and lists the results on the `News` sheet.

## What it does

- Calls NewsAPI's `everything` endpoint
- Fetches articles with the default `q=world` and `language=en` settings
- Shows published time, source, image, title, description, and link
- Refreshes automatically when the workbook opens
- Displays a visible failure state on the sheet if the request fails

## Requirements

- `NEWSAPI_KEY` environment variable
- Microsoft Excel
- xlflow

## How to run

1. Set the `NEWSAPI_KEY` environment variable.
2. Open `build/Book.xlsm`.
3. The workbook refreshes automatically on open.  
   To refresh manually, run `Main.Run` or `Ui.RefreshNews`.

If you run it through xlflow, use `xlflow run --diagnostic` from this folder.

## Easy places to change

- `src/modules/NewsApi.bas`
  - Adjust the query, language, sort order, and page size.
- `src/modules/NewsSheet.bas`
  - Adjust the sheet layout, image handling, and link rendering.

## Notes

- Images are cached in a temporary folder at runtime.
- The sheet shows up to 15 articles.
