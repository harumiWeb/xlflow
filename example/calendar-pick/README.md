# calendar-pick

`calendar-pick` is an Excel VBA sample managed with xlflow.
It provides a UserForm-based calendar picker and writes the selected date into the currently selected worksheet cell.

## Features

- Shows a reusable calendar picker UserForm
- Lets the user move between months and choose a year
- Highlights weekends, dates outside the current month, and the current date
- Writes the confirmed date to the active single cell using `yyyy/mm/dd` formatting
- Includes a helper macro that installs a worksheet launch button

## Requirements

- Windows
- Microsoft Excel
- xlflow

## How to run

1. Create `build/Book.xlsm` and use `xlflow push` to transfer the code.
2. Open `build/Book.xlsm`.
3. Select a single cell.
4. Run `Main.Run` to open the calendar picker.
5. Choose a date. The selected date is written to the selected cell.

To install a launch button on the first worksheet, run `Ui.InstallCalendarPickerButton`.

To run from the CLI, use `xlflow run --diagnostic` in this folder. `xlflow.toml` already configures `entry = "Main.Run"`, so no extra macro argument is needed.

## Areas to customize

- `src/forms/CalendarPicker.frm`
  - Contains the UserForm behavior, month navigation, year selection, and calendar rendering.
- `src/classes/CalendarDayButton.cls`
  - Connects dynamically created day buttons to the form click handler.
- `src/modules/App.bas`
  - Resolves the target cell, opens the picker, writes the selected date, and installs the launch button.
- `src/modules/Ui.bas`
  - Thin wrappers for worksheet buttons.

## Notes

- The picker expects exactly one cell to be selected before it runs.
- Closing or canceling the form leaves the worksheet unchanged.
- The form supports years from 1900 through 2100.
