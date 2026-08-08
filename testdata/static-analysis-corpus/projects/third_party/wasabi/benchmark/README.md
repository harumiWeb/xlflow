# Wasabi Benchmark Suite

> [!NOTE]
> These benchmarks can run and display results in any ![](../resources/svg/ms-office.svg) **Microsoft Office** program.

> [!NOTE]
> <img src="../resources/logo.png" width="20" /> **Wasabi Version targeted:** [v2.3.7-beta](https://github.com/uesleibros/wasabi/releases/tag/v2.3.7-beta)

This directory contains a self contained, high resolution benchmark harness for the core primitives of Wasabi. It measures raw CPU throughput and memory allocation overhead, completely independent of network latency. The framework utilizes the native Windows `QueryPerformanceCounter` API for sub millisecond precision.

## Evaluated Operations

| Operation Category | Targeted Internal Functions | Description |
|:---|:---|:---|
| **Memory Boundaries** | `WasabiMemFind` | Evaluates the speed of internal byte array scanning for HTTP boundaries and payload separators. |
| **Crypto and Encoding** | `DecodeBase64`, `Base64Encode` | Measures the execution time of the `CryptStringToBinaryW` and `CryptBinaryToStringW` APIs, specifically validating the safe NTLM token decoding throughput. |
| **String Conversions** | `StringToUtf8`, `Utf8ToString` | Measures the overhead of wide character to UTF-8 byte array transformations. |
| **WebSocket Framing** | `BuildWSFrame` | Tests allocation overhead, native masking speed, and MTU fragmentation limits when constructing heavy 64KB payload frames. |

## Prerequisites

1. The target version of `Wasabi.bas` must be present in the active VBA project.
2. To allow the benchmark harness to measure internal performance, the targeted internal functions mentioned above must be temporarily exposed by changing their scope from `Private` to `Public` within `Wasabi.bas`.

## Execution Procedure

1. Import all `.bas` files from this benchmark folder into your VBA project.
2. Open the VBA Immediate Window (Ctrl+G).
3. Execute the suite by typing `Benchmark_Runner.RunAllBenchmarks` and pressing Enter.

## Interpreting the Telemetry

* **Avg Latency (µs):** The average microseconds consumed per individual operation.
* **Throughput (MB/s):** The volume of megabytes the operation successfully processes per second.
* **Ops/s (k):** The total thousands of operations completed per second.

The suite dynamically adapts iteration counts based on the payload size to ensure statistically significant timing samples.
