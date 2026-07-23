# Test-driven VBA development

Testing means writing a small VBA procedure that proves a result, then letting xlflow run it like any other macro. You do not need a testing framework before starting: xlflow can create the basic test module for you.

**Example goal:** `CalculateTotal` writes a value to `Result!B2`. The test first expects that value; it fails until the implementation is correct.

```powershell
xlflow generate test CalculatorTests
xlflow test --json
```

The first command adds a test file under `src/`; the second runs discovered `Test*` procedures. A failure is useful evidence, not a broken setup—read its procedure name and assertion message, then implement the smallest change in `src/`.

```powershell
xlflow fmt --check
xlflow lint --json
xlflow analyze --json
xlflow session start --json
xlflow push --fast --session --no-save --json
xlflow test --session --json
xlflow inspect range --sheet Result --address B2 --session --json
xlflow save --session --json
xlflow session stop --json
```

The source checks happen before Excel opens. The session keeps Excel open while you repeat the test. `inspect` is a second proof: it checks the workbook value independently of the test assertion. `--no-save` leaves experiments in the live workbook until `save --session`; omit the save and discard only when you intentionally want to abandon them.

Use a fresh non-session workbook for isolation modes and retries. A session-backed test intentionally observes the same live workbook and supports only the documented session isolation mode.
