# space-invader

`space-invader` is an Excel VBA sample managed with xlflow.
It runs a Space Invader-style game on a modeless UserForm.

## Features

- Displays the game in a VBA UserForm
- Moves the player with keyboard input
- Spawns enemy rows, bullets, lives, score, and game-over state
- Uses pixel-like labels to render the player, enemies, bullets, and UFO
- Starts automatically when the workbook opens

## Requirements

- Windows
- Microsoft Excel
- xlflow

## How to run

1. Create `build/Book.xlsm` and use `xlflow push` to transfer the code.
2. Open `build/Book.xlsm`.
3. The game starts automatically on workbook open.
4. To start it manually, run `Main.Run` or `Ui.RunFromButton`.

To run from the CLI, use `xlflow run --diagnostic` in this folder. `xlflow.toml` already configures `entry = "Main.Run"`, so no extra macro argument is needed.

## Controls

| Key   | Action         |
| ----- | -------------- |
| Left  | Move left      |
| Right | Move right     |
| E     | Fire           |
| R     | Restart        |
| Esc   | Close the game |

## Areas to customize

- `src/forms/specs/frmInvader.yaml`
  - Contains the source-controlled UserForm Designer spec.
- `src/forms/code/frmInvader.bas`
  - Contains the game loop, rendering, collision handling, and keyboard input.
- `src/modules/GameConstants.bas`
  - Defines board size, speeds, enemy layout, scoring, and sprite patterns.
- `src/modules/GameMain.bas`
  - Opens and stops the modeless game form.
- `src/modules/Main.bas`
  - The xlflow entry point.

## Notes

- This sample intentionally uses interactive UserForm behavior, so `forbid_interactive_input` is disabled in `xlflow.toml`.
- Keyboard polling uses the Windows `GetAsyncKeyState` API.
- The workbook `Workbook_Open` event starts the game automatically.
