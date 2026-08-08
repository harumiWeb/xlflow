# Wasabi Unit Testing Framework

> [!NOTE]
> These unit tests can run and display results in any ![](../resources/svg/ms-office.svg) **Microsoft Office** program.

> [!NOTE]
> <img src="../resources/logo.png" width="20" /> **Wasabi Version targeted:** [v2.3.7-beta](https://github.com/uesleibros/wasabi/releases/tag/v2.3.7-beta)

This directory contains the comprehensive unit and integration testing suite for Wasabi. The architecture is written entirely in pure VBA and executes directly within the native VBA IDE, requiring no external dependencies or third party testing frameworks. 

## Test Suite Architecture

The framework is orchestrated by `Test_Runner.bas`, which tracks assertions and aggregates pass/fail metrics. The suite is divided into specific domains:
* `Test_Utils.bas`: Validates data transformations, specifically the cryptography API implementations for Base64 and UTF-8 handling.
* `Test_Memory.bas`: Ensures internal memory boundary logic operates safely without buffer overflows.
* `Test_WebSockets.bas`: Verifies framing, handshakes, and payload integrity.
* `Test_TCP.bas`: Validates raw socket connections and error state handling.
* `Test_MQTT.bas`: Tests the protocol negotiation and MQTT broker handshake implementation.

## Prerequisites

To perform deep internal validation, certain core functions within `Wasabi.bas` must be temporarily exposed. You will need to change the scope of the following functions from `Private` to `Public` prior to running the suite:

| Internal Function | Validation Target |
|:---|:---|
| `DecodeBase64` | CryptStringToBinaryW implementation and NTLM integrity |
| `Base64Encode` | Standard Base64 encoding logic |
| `WasabiMemFind` | Internal byte boundary detection and memory scanning |
| `SHA1` | Cryptographic hashing for WebSocket handshakes |
| `GenerateWSKey` | Sec-WebSocket-Key generation |
| `ComputeWebSocketAccept` | Sec-WebSocket-Accept validation |
| `ParseURL` | URI scheme and port extraction logic |
| `StringToUtf8` | Wide string to UTF-8 byte array conversion |
| `Utf8ToString` | UTF-8 byte array to wide string conversion |
| `BuildWSFrame` | RFC 6455 compliant frame construction |

*Disclaimer: These functions are not part of the stable public API. Their signatures and internal memory management routines may change between releases without notice.*

## Execution Procedure

1. Import your target version of `Wasabi.bas` into a new VBA project (Microsoft Excel is the recommended host environment).
2. Import all `.bas` testing modules from this repository directory into the same project.
3. Open the VBA Immediate Window (`Ctrl+G`).
4. Execute the suite by typing `Test_Runner.RunAllTests` and pressing Enter.

The engine will output the execution flow in real time to the Immediate Window, culminating in a summarized telemetry report detailing the total passed and failed assertions.
