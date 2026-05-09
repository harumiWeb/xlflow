# tetris

`tetris` is an Excel VBA sample managed with xlflow.
It renders a Tetris play screen on a worksheet, allowing you to play using only keyboard operations.

## Features

- Play Tetris on a 20x10 grid
- Randomly generate 7 types of tetrominoes
- Supports lateral movement, rotation, soft drop, hard drop, and restart
- Displays score and number of cleared lines on the sheet

## How to run

1. Create `build/Book.xlsm` and use `xlflow push` to transfer the code.
2. The game will start automatically when you open the workbook.
3. Operate with the following keys:

| Key   | Action          |
| ----- | --------------- |
| ← / → | Move left/right |
| ↑     | Rotate          |
| ↓     | Soft drop       |
| Space | Hard drop       |
| R     | Restart         |

## Areas for customization

- `src/modules/App.bas`
  - Contains board size, drop speed, score calculation, key operations, and rendering logic.
- `src/modules/Main.bas`
  - The entry point.
- `src/workbook/ThisWorkbook.bas`
  - Contains the logic to start the game when the workbook is opened.

## Notes

- No external APIs or additional configurations are required.
- It uses Excel's `Application.OnKey` and `Application.OnTime` to handle input and automatic falling.
- The game loop stops when Excel is closed.
