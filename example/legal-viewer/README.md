# e-Gov Statute Search Viewer

A Windows application that extracts only the required statutes from the e-Gov Statute API Version 2, enabling search, viewing, comparison, and citation directly within Excel.
You can verify the full text of each statute by publication date as available on e-GOV, then locate and cite specific articles, clauses, items, or supplementary tables for inclusion in your documents.

## Key Features

- Search by statute name, statutory number, statutory ID, or registered alias names
- Select from historical versions, current version, or unimplemented provisions based on implementation date
- View articles, clauses, and items; supplementary rules; and supplementary tables
- Perform full-text searches across a single selected statute/implementation date combination
- Supports AND / OR logical operations
- Enables exact match searching using double quotation marks
- Generate citations for statutory text or supplementary tables
- Basic citation format
- Citation with source attribution
- Markdown-style citation
- Markdown table format
- Ascii table formatting with divider lines
- Tab-delimited export suitable for Excel pasting
- Edit and copy selected citations to clipboard
- Review, rerun, or delete search history
- Register tags and notes for bookmarks
- Compare different implementation dates or select arbitrary articles/clauses
- View API communication logs and error logs
- Customize settings including network conditions, searches, citation formats, and management sheets

## System Requirements

- Windows version of Microsoft Excel 365
- Internet connection required
- Environment capable of executing VBA macros
  Note: This application is not compatible with Mac versions of Excel.

## Launch Instructions

1. Save `law-viewer.xlsm` to your local folder.
2. For files downloaded from the internet, open the file properties and enable "Allow" if prompted.
3. Open the file in Excel and select either "Enable Content" or activate macros.
4. Click the "Open Statute Viewer" button in the main sheet.

## Basic Usage Guide

### 1. Adding a Statute

1. Press "Add Statute" on the main screen.
2. Enter the statute name, statutory number, ID, or alias name to search for it.
3. From the search results, select the desired statute.
4. Choose an implementation date from the available options.
5. Click "Add."
   The application will retrieve only the full text of the selected "Statute ID × Implementation Date" combination from e-Gov. If retrieval and parsing are successful, the added statute will appear in the main screen's list of managed statutes.

### 2. Reading the Full Text

After adding a statute, navigate to the full text section and select articles, supplementary rules, or supplementary tables from the navigation menu.

- Upper preview: Full text of a provision or appended table
- Lower preview: Selected units such as specific clauses, items, or rows in an appended table
  Appended tables can be displayed in either Markdown format or ASCII grid format. If table structure cannot be retrieved from the original data, the display will default to plain text.

### 3. Search the Text Content

Searching targets a single legal act and implementation date currently selected in the main screen:

- Enter multiple search terms separated by spaces
- Select AND / OR logical operators
- Use double quotes for exact phrase matching (e.g., "break time")
  Upon selecting a search result, you can preview both the full relevant provision and its smallest constituent unit.

### 4. Generate Quotations

1. Choose the content to quote from either "Text Content" or "Search Results."
2. Select the desired quotation format.
3. Click "Generate."
4. Edit the quotation field as needed if necessary.
5. Click "Copy."
   When regenerating or reselecting an edited quotation, you'll be prompted to confirm overwriting.

## Additional Features

### History

View the history of both legal act searches and text content searches. You can restore previously used criteria to reopen them, or delete entries individually, by type, or all at once.

### Bookmarks

Allows registration of specific clauses, items, appended tables, rows, and sections. Saves only references, tags, and note information—not the actual text content. When opening, if the source text hasn't been retrieved yet, the system will automatically fetch the required legal act from e-Gov.

### Comparison

Select both a legal act and a text unit on either side, then compare different implementation dates or specific provisions in a Git-style diff format for comparison.

### Alias Master

Register alternate names for frequently used legal acts. Can add, update, or delete aliases, along with their official legal titles and notes.

### Settings

Modify various options including quotation formats, content selection criteria, search modes, API communication log limits, management sheet display settings, retry counts, timeout duration, and API call intervals.

### Logs

View both API communication logs and standard error logs. For inquiries or issue verification, please check the occurrence date, processing name, HTTP status code, and error message when needed.

## Data Storage

The following data is saved within the application's book:

- Search history records
- Bookmarks, tags, and note information
- User preferences settings
- Alias master database
- API communication logs and error logs

The following data is treated as temporary cache and will be deleted when the book is closed, opened, or manually removed:

- Full text of legal statutes
- Legal statute metadata listings
- Potential enforcement dates
- Search results
- Comparison reference data
  When closing the book, the system will delete temporary data and overwrite the current file. If termination processing or saving fails, the operation will be canceled for data protection purposes.

## What This Application Does Not Do

- Massive batch retrieval of all legal statute texts
- Permanent storage of all legal statute texts
- Cross-searching across full text keywords for all e-Gov legal statutes
- Cross-text searching spanning multiple added legal statutes

## Usage Notes

- Legal statute texts and enforcement date information are obtained from the e-Gov Statute API. Retrieval may fail depending on network conditions or the API's operational status.
- Always verify the target statute, enforcement date, article number, and full text content before citing any information.
- This application serves as an auxiliary tool for verifying legal information and does not provide legal advice.
- Bookmark notes, history, settings, and logs are stored within the book. Please exercise caution when entering or storing sensitive or personal information.
- ASCII border tables may display inconsistent line positioning depending on the font used or destination of pasted content.

## Troubleshooting

### Screen Does Not Open

- Check that macros are enabled.
- Click the "Open Legal Statute View" button in the startup sheet.
- If "Allow" is displayed in the file properties, enable it and reopen the application.

### Unable to Retrieve Legal Statutes

- Confirm your internet connection is working properly.
- Review the "API Log" and "Error Log."
- Check the timeout settings, retry counts, and API call intervals in the settings panel.

### Search Results Not Found

- Try using the official statute name, legal statute number, or statutory ID.
- Verify that any alternate names registered in the alternative naming master are correct.
- For text searches, please confirm that you are selecting the correct statute and enforcement date as the search target.

### Can't Close Application

Please verify write permissions to the destination, whether the file is read-only, and whether a file with the same name is already in use. If the save operation fails, the app will cancel closing.
