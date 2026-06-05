# Maze Chase

Maze Chase is an xlflow example project that builds a Pac-Man inspired arcade game entirely in Excel VBA.

The game runs inside a UserForm and uses Label and Frame controls for the maze, pellets, sprite rendering, score UI, and game loop updates. It is designed as a practical example of how to use xlflow for a larger interactive VBA project instead of a sheet-automation macro.

## Features

- UserForm-based gameplay with no sheet rendering
- Maze walls drawn from control geometry
- Pac-Man style player and ghost sprite rendering
- Dot collection, cherry power-up items, scoring, and lives
- Directional input and animated player mouth states
- Simple ghost chase AI with frightened mode
- Focused xlflow tests for core game rules

## Project Layout

- [src/modules/Main.bas](src/modules/Main.bas): default entry point wired from xlflow
- [src/modules/App.bas](src/modules/App.bas): interactive game launcher
- [src/modules/MazeChaseGame.bas](src/modules/MazeChaseGame.bas): game state, movement, collisions, scoring, frightened mode
- [src/modules/MazeChaseRenderer.bas](src/modules/MazeChaseRenderer.bas): UserForm rendering and sprite updates
- [src/modules/MazeChaseLoop.bas](src/modules/MazeChaseLoop.bas): Win32 timer-driven game loop
- [src/modules/MazeChaseInput.bas](src/modules/MazeChaseInput.bas): keyboard input handling
- [src/forms/code/MazeChaseForm.bas](src/forms/code/MazeChaseForm.bas): UserForm lifecycle hooks
- [src/modules/Tests/TestMazeChase.bas](src/modules/Tests/TestMazeChase.bas): gameplay tests

## Running The Example

Start an xlflow session, push the current source, and run the default macro:

```powershell
xlflow session start --json
xlflow push --fast --session --no-save --json
xlflow run Main.Run --direct --interactive --session --json
```

When you are done, save and close the session:

```powershell
xlflow save --json
xlflow session stop --json
```

## Testing

Run the focused gameplay suite while iterating:

```powershell
xlflow test --module TestMazeChase --session --json
```

Run the full suite before publishing or sharing changes:

```powershell
xlflow test --session --json
```

The current example project is expected to pass cleanly end-to-end. The
`SampleTests` module contains only a minimal smoke test, while the Maze Chase
behavior is covered by `TestMazeChase`.

Run lint before finalizing changes:

```powershell
xlflow lint --json
```

## Why This Example Exists

This example demonstrates that xlflow can support:

- non-trivial VBA program structure
- interactive UserForm applications
- repeatable source-controlled VBA development
- testable game logic separated from UI code

It is intended to be a reference project for anyone building richer Excel VBA experiences with xlflow.
