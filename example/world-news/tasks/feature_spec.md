# Feature Spec

## Goal

Open the workbook and automatically fetch recent world-news articles from NewsAPI, then display them in a worksheet.

## Contract

- The workbook reads the API key from the `NEWSAPI_KEY` environment variable.
- The workbook fetches data from NewsAPI's `everything` endpoint with a default `q=world` query.
- The workbook does not prompt the user for input.
- `ThisWorkbook.Workbook_Open` triggers the refresh flow by calling `Main.Run`.
- The rendered sheet shows a title area, refresh metadata, and one row per article with published time, source, image, title, description, and URL.
- If the API key is missing or the request fails, the workbook renders a visible failure state on the sheet.

## Module Responsibilities

- `Main`: stable macro entrypoint.
- `App`: orchestration and error-to-sheet handling.
- `NewsApi`: URL construction, environment lookup, HTTP request, and JSON decode entrypoint.
- `JsonParser`: parse the NewsAPI JSON payload into dictionaries and collections.
- `NewsSheet`: create/reset the sheet, download article images when `urlToImage` exists, and render success or failure output.
- `ThisWorkbook`: call `Main.Run` on workbook open.
