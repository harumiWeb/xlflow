<!-- 実行計画 -->

# xlflow C# bridge 化実行計画

## 1. 方針

xlflow は Go CLI を維持し、Windows/Excel 固有の実行部分を C# bridge へ段階移行する。

```txt
Go core:
  - CLI command
  - project/source tree management
  - config loading
  - lint / format / static analysis
  - JSON envelope mapping
  - release packaging
  - bridge provider selection

C# bridge:
  - Excel COM automation
  - VBIDE automation
  - workbook/session handling
  - macro execution
  - UserForm import/export
  - Win32 dialog watcher
  - runtime/compile error capture
  - export-image fallback
  - process/window/clipboard control

PowerShell bridge:
  - legacy fallback
  - unsupported .NET command fallback in auto mode only
```

狙いは、PowerShell 実行ポリシーに依存しない企業環境対応と、COM/VBIDE/Win32 周辺の制御品質向上である。

## 2. 重要な設計決定

### 2.1 C# bridge は別 exe

`xlflow.exe` に .NET runtime を埋め込まない。Windows release archive では `xlflow.exe` と `xlflow-excel-bridge.exe` を同梱し、Go 側が bridge provider として起動する。

理由:

- Go CLI の軽量性とクロスプラットフォーム性を保てる。
- Excel COM/VBIDE 操作を STA 前提の C# process に閉じ込められる。
- C# bridge が壊れても PowerShell fallback を残せる。
- bridge protocol を固定すれば段階移行できる。

### 2.2 stdin/stdout JSON protocol

初期通信方式は stdin/stdout JSON とする。gRPC、named pipe daemon、常駐 server は初期 scope から外す。

```txt
Go CLI
  -> starts xlflow-excel-bridge.exe
  -> writes one request JSON to stdin
  -> reads one response JSON from stdout
```

制約:

- stdout は bridge response JSON 専用にする。
- ログは response の `logs` または stderr に出す。
- C# bridge の fatal error も JSON response に変換する。
- request/response には `protocol_version` と `request_id` を必ず入れる。

### 2.3 STA/async 境界

Excel COM 操作は dedicated STA thread に閉じ込める。`async/await` 後に COM 操作が ThreadPool/MTA へ戻る設計は禁止する。

初期実装では `Program.Main` と command handler を同期実行に寄せる。将来 async I/O が必要になった場合も、COM 操作は STA dispatcher に投げる。

### 2.4 fallback rule

bridge mode は `auto | dotnet | powershell`。

```txt
--bridge dotnet:
  - dotnet bridge が無い、protocol 不一致、unsupported command の場合は失敗
  - PowerShell へ暗黙 fallback しない

--bridge powershell:
  - 常に PowerShell bridge を使う

--bridge auto:
  - 現段階では PowerShell first
  - .NET 対応 command が安定した段階で dotnet first に切り替える
  - unsupported command は PowerShell fallback してよい
```

明示指定時に fallback しない理由は、CI や企業環境検証で「PowerShell を使っていない」ことを証明できるようにするため。

### 2.5 企業配布リスク

C# exe は PowerShell 実行ポリシーを避けられるが、AppLocker、WDAC、Defender、EDR、署名ポリシーには影響される可能性がある。default bridge を .NET に切り替える前に、release packaging と code signing 方針を ADR か docs に残す。

## 3. 目標ディレクトリ構成

```txt
xlflow/
├─ cmd/xlflow/
├─ internal/
│  ├─ config/
│  ├─ output/
│  ├─ project/
│  ├─ lint/
│  └─ excel/
│     ├─ runner.go
│     ├─ options.go
│     ├─ result.go
│     ├─ payload.go
│     ├─ bridge/
│     │  ├─ provider.go
│     │  ├─ resolver.go
│     │  ├─ contract.go
│     │  ├─ protocol.go
│     │  ├─ dotnet.go
│     │  ├─ powershell.go
│     │  ├─ errors.go
│     │  └─ version.go
│     ├─ scripts/
│     └─ forms/
├─ bridge/dotnet/
│  ├─ Xlflow.ExcelBridge.sln
│  ├─ src/Xlflow.ExcelBridge/
│  │  ├─ Xlflow.ExcelBridge.csproj
│  │  ├─ Program.cs
│  │  ├─ Contract/
│  │  ├─ Commands/
│  │  ├─ Excel/
│  │  ├─ Runtime/
│  │  ├─ Windows/
│  │  ├─ Serialization/
│  │  └─ Diagnostics/
│  └─ tests/Xlflow.ExcelBridge.Tests/
├─ docs/adr/
│  ├─ ADR-0008-dotnet-excel-bridge.md
│  └─ ADR-0009-bridge-provider-contract.md
├─ docs/bridge/
│  ├─ bridge-protocol.md
│  ├─ dotnet-bridge.md
│  └─ powershell-bridge-legacy.md
├─ scripts/
│  ├─ build-dotnet-bridge.ps1
│  └─ test-dotnet-bridge.ps1
├─ global.json
└─ .goreleaser.yaml
```

ADR 番号は既存の `ADR-0006` / `ADR-0007` と衝突しないように `ADR-0008` 以降を使う。

## 4. Go 側設計

`internal/excel` は Excel を直接操作しない。xlflow command options/config を bridge request payload に変換し、bridge provider を呼ぶ層にする。

```go
type Runner struct {
    RootDir string
    Bridge  bridge.Provider
}
```

provider interface:

```go
type Provider interface {
    Name() string
    Supports(command string) bool
    Info(ctx context.Context) (Info, error)
    Execute(ctx context.Context, req Request) (Response, error)
}
```

設定優先順位:

```txt
CLI option --bridge
  > XLFLOW_EXCEL_BRIDGE
  > [excel].bridge in xlflow.toml
  > default auto
```

Phase 1 では `internal/excel/bridge.go` の既存挙動を変えずに PowerShellProvider へ隔離する。CLI 出力・exit code・script args は互換を維持する。

## 5. Bridge protocol

request:

```json
{
  "protocol_version": 1,
  "request_id": "01J...",
  "command": "inspect",
  "timeout_ms": 60000,
  "payload": {
    "workbook_path": "C:\\dev\\Book.xlsm",
    "visible": false,
    "sheet": "Sheet1",
    "address": "A1:D20"
  }
}
```

response:

```json
{
  "protocol_version": 1,
  "request_id": "01J...",
  "status": "ok",
  "command": "inspect",
  "logs": [],
  "error": null,
  "bridge": {
    "name": "dotnet",
    "version": "0.1.0",
    "protocol_version": 1
  },
  "inspect": {}
}
```

error:

```json
{
  "protocol_version": 1,
  "request_id": "01J...",
  "status": "failed",
  "command": "run",
  "logs": [],
  "error": {
    "code": "EXCEL_COM_ERROR",
    "message": "Failed to run macro.",
    "phase": "macro.run",
    "source": "Microsoft Excel",
    "number": -2146827284,
    "hresult": "0x800A03EC",
    "details": {}
  },
  "suggestions": [
    "Run `xlflow doctor --json` to verify Excel and VBIDE access.",
    "Check that the macro name is fully qualified, for example `Module1.Main`.",
    "Run `xlflow macros --json` to list runnable macros."
  ]
}
```

## 6. C# bridge 設計

`Program.cs` は薄く保つ。

責務:

- `[STAThread]` entrypoint
- `--version-json` / `--capabilities-json`
- stdin JSON の読み取り
- command dispatch
- stdout JSON response
- fatal error の JSON 化

初期 handler interface:

```csharp
public interface ICommandHandler
{
    string CommandName { get; }
    bool Supports(BridgeRequest request);
    BridgeResponse Handle(BridgeRequest request, CancellationToken cancellationToken);
}
```

大きな `ExcelBridge.cs` は作らない。PowerShell `common.ps1` の巨大化を C# で再現しないため、責務ごとに分ける。

```txt
Commands/      command dispatch and command-specific orchestration
Contract/      request/response/error/capability models
Serialization/ JSON options and stdin/stdout transport
Diagnostics/   environment, Office, VBIDE probes
Excel/         Excel.Application, Workbook, VBProject wrappers
Runtime/       runtime injection, debug, UI/debug stream
Windows/       Win32, UI Automation, process, clipboard, dialogs
```

## 7. 移行順序

`run` は重要だが、runtime injection、dialog watcher、VBA error capture、timeout、Excel process cleanup が絡むため後回しにする。

```txt
1. ADR / bridge protocol docs
2. Go bridge provider abstraction
3. PowerShellProvider 化
4. --bridge auto|powershell|dotnet
5. .NET bridge skeleton
6. --version-json / --capabilities-json
7. doctor
8. inspect / process
9. pull / push
10. macros
11. run
12. dialog watcher
13. test / debug / runtime injection
14. form / export-image
15. .NET bridge packaging / release
16. remaining PowerShell-only commands
17. .NET bridge default 化
18. PowerShell bridge legacy 化
```

## 8. 実装フェーズ

### Phase 0: ADR と protocol docs

作るもの:

- `docs/adr/ADR-0008-dotnet-excel-bridge.md`
- `docs/adr/ADR-0009-bridge-provider-contract.md`
- `docs/bridge/bridge-protocol.md`
- `docs/bridge/dotnet-bridge.md`
- `docs/bridge/powershell-bridge-legacy.md`

完了条件:

- Go/C#/PowerShell の責務分割が docs に残っている。
- fallback rule と protocol versioning が明文化されている。

### Phase 1: Go bridge provider abstraction

scope:

- `internal/excel/bridge` package を追加する。
- `Provider` / `Request` / `Response` / `Info` を定義する。
- 既存 PowerShell 実行を `PowerShellProvider` に隔離する。
- `--bridge` option、`XLFLOW_EXCEL_BRIDGE`、`[excel].bridge` を追加する。
- default は `auto` だが、初期は PowerShell 挙動を維持する。

non-goals:

- C# bridge は追加しない。
- JSON envelope と exit code を変えない。

### Phase 2: .NET bridge skeleton

scope:

- `bridge/dotnet/Xlflow.ExcelBridge.sln`
- `src/Xlflow.ExcelBridge`
- `tests/Xlflow.ExcelBridge.Tests`
- `global.json`
- `scripts/build-dotnet-bridge.ps1`
- `scripts/test-dotnet-bridge.ps1`
- `--version-json`
- `--capabilities-json`
- stdin/stdout JSON transport
- minimal `doctor`

完了条件:

- `dotnet build bridge/dotnet/Xlflow.ExcelBridge.sln` が通る。
- `dotnet test bridge/dotnet/Xlflow.ExcelBridge.sln` が通る。
- bridge exe が stdout に valid JSON を返す。

### Phase 3: .NET doctor

診断項目:

- OS / architecture / .NET runtime
- Excel COM activation
- Excel version/build
- VBIDE access
- AutomationSecurity
- Trust access to VBA project object model

完了条件:

- `xlflow doctor --bridge dotnet --json` が PowerShell なしで動く。
- Excel COM errors は structured error で返る。

### Phase 4: .NET inspect / process

scope:

- workbook/sheet/range inspection
- process list
- process cleanup
- session-owned process handling

完了条件:

- `xlflow inspect --bridge dotnet --json`
- `xlflow process list --bridge dotnet --json`
- `xlflow process cleanup --bridge dotnet --json`

### Phase 5: .NET pull / push

scope:

- standard module / class module / document module export/import
- UserForm `.frm` / `.frx`
- sidecar code source
- Rubberduck folder annotation
- changed-only push
- backup mode

完了条件:

- PowerShell bridge と同等の round-trip E2E が通る。
- output envelope が互換である。

### Phase 6: .NET macros / run

scope:

- macro listing
- `Application.Run`
- fully qualified macro name
- args
- timeout
- save/no-save
- compile/runtime error capture

完了条件:

- success は `status: ok`
- runtime error は structured error
- timeout は Excel/dialog 状態つき error

### Phase 7: .NET dialog watcher

scope:

- runtime error dialog
- compile error dialog
- MsgBox/InputBox/FileDialog support
- dialog text/buttons capture
- close/respond behavior

完了条件:

- modal dialog が agent 実行をブロックしない。
- dialog snapshot が JSON diagnostics に残る。

### Phase 8: .NET test / debug / runtime injection

scope:

- test runner
- debug pipe injection
- defined name runtime mode
- MsgBox/InputBox/FileDialog response injection
- debug stream
- UI stream

完了条件:

- `xlflow test --bridge dotnet --json`
- `xlflow run Module1.Main --bridge dotnet --json`
- stream output が envelope に統合される。

### Phase 9: .NET form / export-image

scope:

- form inspect/build
- form image export
- sheet/range/shape image export
- clipboard retry
- window activation / minimized Excel 対策

完了条件:

- `xlflow form inspect --bridge dotnet --json`
- `xlflow export-image --bridge dotnet --json`

### Phase 10: packaging / release

scope:

- Windows archive に `xlflow-excel-bridge.exe` を同梱する。
- Linux/macOS archive には含めない。
- CGO 依存を含むため、release build は OS ごとの native runner で行う。
- Windows は MSYS2 UCRT64 GCC、Linux は Ubuntu native GCC を使い、Linux asset は Windows release 作成後に同じ GitHub Release へ追加する。
- checksum は初期対応として Windows `checksums.txt` と Linux `checksums-linux.txt` に分ける。
- self-contained single-file publish を使う。
- trimming は使わない。
- code signing 方針を検討する。

publish example:

```powershell
dotnet publish bridge/dotnet/src/Xlflow.ExcelBridge/Xlflow.ExcelBridge.csproj `
  -c Release `
  -r win-x64 `
  -p:PublishSingleFile=true `
  -p:SelfContained=true `
  -p:PublishTrimmed=false
```

### Phase 11: remaining PowerShell-only commands

scope:

- `new`
- `session`
- `runner`
- `attach`
- `list`
- `ui`
- `edit`

方針:

- Go 側の既存 command handler と request payload key は原則維持し、C# bridge 側に同じ command key の handler を追加する。
- `CommandRegistry` と `--capabilities-json` に未移行 command を追加し、`--bridge dotnet` で明示実行できるようにする。
- `new` は macro-enabled workbook 作成と初期 VBA bootstrap を .NET 側で実装し、PowerShell path と同じ workbook/source envelope を返す。
- `session` / `attach` は live Excel workbook と `.xlflow/session.json` の対応を .NET 側で管理し、dirty / save_required / active workbook mismatch の意味を既存 JSON contract と揃える。
- `runner` は runner helper module の install / status / remove を .NET 側に移し、`run` / `test` が使う helper contract と統合する。
- `list` は UserForm listing を .NET 側で実装し、既存の `list forms` output shape を維持する。
- `ui` は button add / list / remove を session-aware に実装し、macro verification と workbook save state を既存 behavior と揃える。
- `edit` は cell / range / rows / columns の mutation と検証を .NET 側で実装し、どの operation が変更されたかを JSON に明確に残す。
- 明示 `--bridge dotnet` は fallback しない。未対応 action は PowerShell に逃がさず structured error として返す。

完了条件:

- `xlflow <command> --bridge dotnet --json` が上記 command と sub-action で動く。
- `xlflow <command> --bridge powershell --json` は互換用に維持される。
- Go unit test が全 bridge command key の .NET routing / strict dotnet / explicit powershell を確認する。
- C# unit test が argument mapping、unsupported action、envelope shape を確認する。
- Windows + Excel COM E2E で `new/init`、session start/status/save/stop、attach、list forms、ui button add/list/remove、edit cell/range/rows/columns、既存 pull/push/run/test workflow が通る。

### Phase 12: .NET default 化

条件:

- `doctor`
- `inspect`
- `pull`
- `push`
- `run`
- `macros`
- `test`
- removed trace command surface
- `inspect-form`
- `form-write`
- `form-export-image`
- `export-image`
- `process`
- `new`
- `session`
- `runner`
- `attach`
- `list`
- `ui`
- `edit`

上記が C# bridge で安定し、Windows + Excel COM E2E が通っていること。

変更:

```txt
before:
  auto -> PowerShell first

after:
  auto -> .NET first on Windows, PowerShell fallback for unsupported commands
```

### Phase 13: PowerShell legacy 化

scope:

- PowerShell bridge docs を legacy fallback として更新する。
- doctor に selected bridge / fallback status を出す。
- 新機能は原則 C# bridge first にする。

## 9. CI / test 戦略

常時 CI:

- `go test ./...`
- `dotnet build bridge/dotnet/Xlflow.ExcelBridge.sln`
- `dotnet test bridge/dotnet/Xlflow.ExcelBridge.sln`
- bridge protocol unit tests

Windows + Excel self-hosted:

- `xlflow doctor --bridge dotnet --json`
- blank workbook
- standard module round-trip
- class module round-trip
- UserForm + `.frx` round-trip
- `session start -> push --fast --session --no-save -> run/test --session -> save --session -> session stop`

## 10. Issue 分割

この計画は以下の GitHub issues に分割して進める。

1. [#72 Document the .NET Excel bridge architecture and protocol](https://github.com/harumiWeb/xlflow/issues/72)
2. [#73 Introduce Go bridge provider abstraction](https://github.com/harumiWeb/xlflow/issues/73)
3. [#74 Add .NET Excel bridge skeleton and development tooling](https://github.com/harumiWeb/xlflow/issues/74)
4. [#75 Implement .NET doctor command](https://github.com/harumiWeb/xlflow/issues/75)
5. [#76 Implement .NET inspect and process commands](https://github.com/harumiWeb/xlflow/issues/76)
6. [#77 Implement .NET pull and push commands](https://github.com/harumiWeb/xlflow/issues/77)
7. [#78 Implement .NET macro listing and run command](https://github.com/harumiWeb/xlflow/issues/78)
8. [#79 Implement .NET dialog watcher and modal error handling](https://github.com/harumiWeb/xlflow/issues/79)
9. [#80 Migrate test, debug, and runtime injection to .NET bridge](https://github.com/harumiWeb/xlflow/issues/80)
10. [#81 Migrate form and export-image commands to .NET bridge](https://github.com/harumiWeb/xlflow/issues/81)
11. [#82 Package .NET bridge in Windows releases](https://github.com/harumiWeb/xlflow/issues/82)
12. [#97 Migrate remaining PowerShell-only commands to .NET bridge](https://github.com/harumiWeb/xlflow/issues/97)
13. [#83 Make .NET bridge the default Windows bridge](https://github.com/harumiWeb/xlflow/issues/83)
14. [#84 Mark PowerShell bridge as legacy fallback](https://github.com/harumiWeb/xlflow/issues/84)
