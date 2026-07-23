# CodeLens and testing

CodeLens appears above no-argument runnable `Sub` procedures and test procedures named `Test*`, `test*`, or `*_Test`. Use **Run** for a procedure and **Run Test** for a test. Argument-bearing procedures and functions are intentionally excluded.

The Testing view discovers source-defined tests when `xlflow.testing.autoDiscover` is enabled. Execution remains `xlflow test`; use the output channel and JSON result for detailed failures. Save modified buffers before CodeLens execution when `xlflow.run.saveBeforeRun` is enabled.
