# gen-qrcode

gen-qrcode is an Excel VBA sample managed by xlflow.
It generates a QR code from a string entered into a cell and draws it directly on the worksheet.

## Features

- Generates a QR code from a string entered in cell B2.
- Displays the generated result as a two-dimensional matrix using cell fill colors.
- Automatically redraws when the input changes.
- Handles UTF-8 strings.
- Works without external APIs or add-ins.

## How to run

1. Create `build/Book.xlsm` and use `xlflow push` to transfer the code.
2. Enter the string you want to convert into a QR code in cell B2 of the sheet.
3. The QR code will be generated automatically based on the input.

Main.Run is also executed immediately after opening the workbook. To re-run from the CLI, use `xlflow run --diagnostic` in this folder.

## Areas to customize

- src/modules/App.bas
  - You can adjust the sheet layout, input cell, output position, and drawing process.
- src/modules/QrCode.bas
  - Contains the core logic for QR code encoding, error correction, and mask selection.
- src/workbook/Sheet1.bas
  - Contains the event handling for regeneration when B2 is changed.
- src/modules/Ui.bas
  - A thin wrapper for calling Main.Run from a button.

## Notes

- Entering an empty string will clear the output.
- This sample is intended for demonstration and handles QR Code Versions 1 through 4.
- Strings with UTF-8 payloads exceeding 78 bytes are not supported.
