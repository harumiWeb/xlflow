<!-- 設計メモ -->

# xlflow C# bridge 化デザイン案

## 1. 目標

xlflow の責務を以下のように分割する。

```txt
Go core:
  - CLI command
  - project/source tree management
  - config loading
  - lint
  - formatter
  - static analysis
  - JSON envelope
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
```

狙いは、**Go の軽量 CLI と C# の Windows/COM 適性を両取りする**ことです。

---

# 2. 最終ディレクトリ構成

おすすめ構成はこれです。

```txt
xlflow/
├─ cmd/
│  └─ xlflow/
│     └─ main.go
│
├─ internal/
│  ├─ config/
│  ├─ output/
│  ├─ project/
│  ├─ lint/
│  ├─ format/
│  │
│  └─ excel/
│     ├─ runner.go
│     ├─ options.go
│     ├─ result.go
│     ├─ payload.go
│     │
│     ├─ bridge/
│     │  ├─ provider.go
│     │  ├─ resolver.go
│     │  ├─ contract.go
│     │  ├─ protocol.go
│     │  ├─ dotnet.go
│     │  ├─ powershell.go
│     │  ├─ errors.go
│     │  └─ version.go
│     │
│     ├─ scripts/
│     │  ├─ common.ps1
│     │  ├─ doctor.ps1
│     │  ├─ pull.ps1
│     │  ├─ push.ps1
│     │  ├─ run.ps1
│     │  ├─ inspect.ps1
│     │  └─ ...
│     │
│     └─ forms/
│
├─ bridge/
│  └─ dotnet/
│     ├─ Xlflow.ExcelBridge.sln
│     │
│     ├─ src/
│     │  └─ Xlflow.ExcelBridge/
│     │     ├─ Xlflow.ExcelBridge.csproj
│     │     ├─ Program.cs
│     │     │
│     │     ├─ Contract/
│     │     │  ├─ BridgeRequest.cs
│     │     │  ├─ BridgeResponse.cs
│     │     │  ├─ BridgeError.cs
│     │     │  ├─ BridgeInfo.cs
│     │     │  ├─ BridgeEnvelope.cs
│     │     │  └─ ProtocolVersion.cs
│     │     │
│     │     ├─ Commands/
│     │     │  ├─ ICommandHandler.cs
│     │     │  ├─ CommandRegistry.cs
│     │     │  ├─ DoctorCommand.cs
│     │     │  ├─ InspectCommand.cs
│     │     │  ├─ ProcessCommand.cs
│     │     │  ├─ PullCommand.cs
│     │     │  ├─ PushCommand.cs
│     │     │  ├─ RunCommand.cs
│     │     │  ├─ MacrosCommand.cs
│     │     │  ├─ SessionCommand.cs
│     │     │  ├─ TraceCommand.cs
│     │     │  ├─ FormCommand.cs
│     │     │  └─ ExportImageCommand.cs
│     │     │
│     │     ├─ Excel/
│     │     │  ├─ ExcelApplication.cs
│     │     │  ├─ ExcelApplicationFactory.cs
│     │     │  ├─ WorkbookHandle.cs
│     │     │  ├─ WorkbookResolver.cs
│     │     │  ├─ WorkbookSession.cs
│     │     │  ├─ WorkbookState.cs
│     │     │  ├─ VbeProject.cs
│     │     │  ├─ ComponentImporter.cs
│     │     │  ├─ ComponentExporter.cs
│     │     │  ├─ MacroRunner.cs
│     │     │  ├─ SheetInspector.cs
│     │     │  └─ ImageExporter.cs
│     │     │
│     │     ├─ Runtime/
│     │     │  ├─ RuntimeInjection.cs
│     │     │  ├─ DefinedNameStore.cs
│     │     │  ├─ UiResponseInjection.cs
│     │     │  ├─ DebugPipeServer.cs
│     │     │  ├─ UiStreamPipeServer.cs
│     │     │  └─ TraceInjector.cs
│     │     │
│     │     ├─ Windows/
│     │     │  ├─ NativeMethods.cs
│     │     │  ├─ DialogWatcher.cs
│     │     │  ├─ DialogResult.cs
│     │     │  ├─ WindowFinder.cs
│     │     │  ├─ ProcessGuard.cs
│     │     │  ├─ ClipboardGuard.cs
│     │     │  └─ UiAutomationProbe.cs
│     │     │
│     │     ├─ Serialization/
│     │     │  ├─ JsonOptions.cs
│     │     │  └─ StdinStdoutTransport.cs
│     │     │
│     │     └─ Diagnostics/
│     │        ├─ EnvironmentProbe.cs
│     │        ├─ OfficeProbe.cs
│     │        ├─ VbideProbe.cs
│     │        └─ BridgeSelfCheck.cs
│     │
│     └─ tests/
│        ├─ Xlflow.ExcelBridge.Tests/
│        └─ Xlflow.ExcelBridge.IntegrationTests/
│
├─ docs/
│  ├─ adr/
│  │  ├─ ADR-0006-dotnet-excel-bridge.md
│  │  └─ ADR-0007-bridge-provider-contract.md
│  │
│  └─ bridge/
│     ├─ bridge-protocol.md
│     ├─ dotnet-bridge.md
│     └─ powershell-bridge-legacy.md
│
├─ scripts/
│  ├─ build.ps1
│  ├─ build-dotnet-bridge.ps1
│  ├─ test-dotnet-bridge.ps1
│  └─ release.ps1
│
├─ testdata/
│  └─ workbooks/
│
├─ go.mod
├─ go.sum
├─ README.md
└─ .goreleaser.yaml
```

---

# 3. Go 側の設計

## 3.1 `internal/excel` の責務

Go 側の `internal/excel` は、Excel を直接操作しない。
あくまで **xlflow command を bridge request に変換し、bridge provider を呼び出す層**にする。

```txt
internal/excel/
  runner.go   -> Runner facade
  options.go  -> command options
  payload.go  -> options/config から bridge payload を作る
  result.go   -> bridge response を output.Envelope に変換
```

`Runner` はこういう形を目指す。

```go
type Runner struct {
    RootDir string
    Bridge  bridge.Provider
}
```

各 command は、PowerShell script を直接呼ぶのではなく bridge provider 経由にする。

```go
func (r Runner) Run(ctx context.Context, cfg config.Config, opts RunOptions) (output.Envelope, int, error) {
    req := bridge.Request{
        ProtocolVersion: bridge.CurrentProtocolVersion,
        Command: "run",
        Payload: buildRunPayload(r.RootDir, cfg, opts),
    }

    res, err := r.Bridge.Execute(ctx, req)
    return mapBridgeResponseToEnvelope(res, err)
}
```

---

## 3.2 bridge provider interface

`internal/excel/bridge/provider.go`

```go
package bridge

import "context"

type Provider interface {
    Name() string
    Supports(command string) bool
    Info(ctx context.Context) (Info, error)
    Execute(ctx context.Context, req Request) (Response, error)
}
```

`Supports(command string)` を持たせるのが重要です。

最初から全 command を C# 対応する必要はないため、

```txt
doctor  -> dotnet
inspect -> dotnet
run     -> powershell
push    -> powershell
```

のような段階移行ができます。

---

## 3.3 bridge resolver

`internal/excel/bridge/resolver.go`

```go
type Mode string

const (
    ModeAuto       Mode = "auto"
    ModeDotnet     Mode = "dotnet"
    ModePowerShell Mode = "powershell"
)

type Resolver struct {
    RootDir string
    Mode    Mode
}

func (r Resolver) Resolve(ctx context.Context, command string) (Provider, error) {
    switch r.Mode {
    case ModeDotnet:
        return r.resolveDotnet(ctx, command)
    case ModePowerShell:
        return r.resolvePowerShell(ctx, command)
    default:
        return r.resolveAuto(ctx, command)
    }
}
```

`auto` の選択ロジックはこうする。

```txt
1. xlflow-excel-bridge.exe が xlflow.exe の隣にある
2. bridge --version-json が成功する
3. protocol_version が compatible
4. 対象 command を supports している
5. すべて満たすなら dotnet
6. そうでなければ powershell fallback
```

---

## 3.4 bridge 選択オプション

設定・環境変数・CLI option の順に上書きできる形がよいです。

```toml
[excel]
bridge = "auto" # auto | dotnet | powershell
```

```bash
XLFLOW_EXCEL_BRIDGE=dotnet
```

```bash
xlflow run Main --bridge dotnet --json
xlflow inspect sheet --bridge powershell --json
```

優先順位はこれ。

```txt
CLI option
  > environment variable
  > xlflow config
  > default auto
```

---

# 4. Bridge protocol

## 4.1 通信方式

最初は **stdin/stdout JSON** がよいです。

gRPC や常駐 daemon は後回しで十分です。

```txt
Go CLI
  -> starts xlflow-excel-bridge.exe
  -> sends request JSON to stdin
  -> receives response JSON from stdout
```

理由は、

```txt
- 実装が単純
- デバッグしやすい
- Docker credential helper 的な設計に近い
- command 単位の実行モデルに合う
- PowerShell provider と contract を合わせやすい
```

です。

---

## 4.2 Request schema

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

## 4.3 Response schema

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

## 4.4 Error schema

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

---

# 5. C# bridge の設計

## 5.1 entrypoint

`Program.cs` は薄くする。

責務はこれだけ。

```txt
- STA thread で起動
- --version-json / --capabilities-json を処理
- stdin JSON を読む
- command handler に dispatch
- stdout JSON を返す
- fatal error を JSON に変換する
```

イメージ。

```csharp
[STAThread]
public static int Main(string[] args)
{
    return BridgeHost.Run(args);
}
```

`BridgeHost` 内で、

```csharp
var request = await StdinStdoutTransport.ReadRequestAsync();
var handler = registry.Resolve(request.Command);
var response = await handler.HandleAsync(request, cancellationToken);
await StdinStdoutTransport.WriteResponseAsync(response);
```

とする。

---

## 5.2 command handler

各 command は `ICommandHandler` を実装する。

```csharp
public interface ICommandHandler
{
    string CommandName { get; }
    bool Supports(BridgeRequest request);
    Task<BridgeResponse> HandleAsync(BridgeRequest request, CancellationToken cancellationToken);
}
```

最初から大きな `ExcelBridge.cs` を作らないのが重要です。
PowerShell の `common.ps1` 巨大化問題を C# で再現しないようにします。

---

## 5.3 Excel COM layer

Excel COM 操作は `Excel/` に隔離する。

```txt
ExcelApplication:
  - Excel.Application の生成/接続
  - DisplayAlerts / Visible / AutomationSecurity 設定
  - COM release

WorkbookHandle:
  - Workbook open/save/close
  - session workbook attach
  - dirty state

VbeProject:
  - VBProject/VBComponents 操作
  - component import/export
  - compile checks

MacroRunner:
  - Application.Run
  - runtime error capture
  - dialog watcher 連携
```

C# bridge は必ず STA 前提にする。

```csharp
[STAThread]
public static int Main(string[] args)
{
    ...
}
```

---

## 5.4 Windows layer

Win32 / UI Automation は `Windows/` に分ける。

```txt
DialogWatcher:
  - Excel runtime error dialog
  - compile error dialog
  - MsgBox/InputBox/FileDialog 補助
  - timeout 時の window snapshot

ProcessGuard:
  - Excel process tree
  - orphan Excel cleanup
  - session-owned process 判定

ClipboardGuard:
  - export-image / CopyPicture fallback
  - clipboard contention handling
```

この層は C# 化の大きな価値になります。

---

# 6. コマンド移行方針

最初から全部移植しない方がいいです。

## 移行優先度

```txt
Priority 1:
  doctor
  process list
  process cleanup
  inspect

Priority 2:
  pull
  push
  macros list

Priority 3:
  run
  compile check
  runtime error capture
  dialog watcher

Priority 4:
  test
  trace
  runtime injection
  debug/ui stream

Priority 5:
  form
  export-image
  advanced UI automation
```

理由は、`run` が一番重要に見える一方で、runtime injection、dialog watcher、VBA error capture、timeout、Excel process cleanup が絡んで一番難しいからです。

まずは `doctor` / `inspect` で bridge contract と COM lifecycle を固めるのが安全です。

---

# 7. 実装フェーズ

## Phase 0: ADR と設計固定

### 目的

設計判断を先に固定する。

### 作るもの

```txt
docs/adr/ADR-0006-dotnet-excel-bridge.md
docs/adr/ADR-0007-bridge-provider-contract.md
docs/bridge/bridge-protocol.md
```

### 内容

```txt
- Go core と C# bridge の責務分割
- C# bridge を別 exe にする理由
- PowerShell bridge を legacy fallback として残す理由
- stdin/stdout JSON protocol を採用する理由
- protocol_version を導入する理由
```

### 完了条件

```txt
- 設計方針が docs に残っている
- 以降の PR がこの設計に沿って切れる
```

---

## Phase 1: Go 側 bridge abstraction 導入

### 目的

C# bridge をまだ作らず、Go 側を差し替え可能にする。

### 実装内容

```txt
internal/excel/bridge/
  provider.go
  resolver.go
  contract.go
  powershell.go
  errors.go
```

既存 PowerShell 実行処理を `PowerShellProvider` に寄せる。

### 重要ポイント

この Phase では **挙動変更なし** にする。

```txt
before:
  Runner -> PowerShell script

after:
  Runner -> BridgeResolver -> PowerShellProvider -> PowerShell script
```

### 完了条件

```txt
- 既存テストが通る
- 既存コマンドの出力が変わらない
- --bridge powershell が使える
- --bridge auto の default が powershell fallback になる
```

---

## Phase 2: C# bridge skeleton 追加

### 目的

C# bridge の最小実行体を追加する。

### 実装内容

```txt
bridge/dotnet/
  Xlflow.ExcelBridge.sln
  src/Xlflow.ExcelBridge/
    Program.cs
    Contract/
    Serialization/
    Commands/
```

最初に対応するのは以下のみ。

```txt
--version-json
--capabilities-json
doctor minimal
```

### `--version-json` 出力例

```json
{
  "name": "xlflow-excel-bridge",
  "version": "0.1.0",
  "protocol_version": 1,
  "commit": "dev",
  "runtime": ".NET 8.0",
  "architecture": "x64"
}
```

### `--capabilities-json` 出力例

```json
{
  "commands": ["doctor"]
}
```

### 完了条件

```txt
- dotnet build が通る
- dotnet test が通る
- bridge exe が --version-json を返す
- Go 側から bridge info を取得できる
- xlflow doctor --bridge dotnet --json が最低限動く
```

---

## Phase 3: `doctor` を C# bridge 対応

### 目的

C# bridge で Excel 環境診断をできるようにする。

### 実装内容

```txt
Diagnostics/
  EnvironmentProbe.cs
  OfficeProbe.cs
  VbideProbe.cs

Commands/
  DoctorCommand.cs
```

診断項目。

```txt
- OS
- architecture
- .NET runtime
- Excel COM activation
- Excel version
- VBIDE access
- AutomationSecurity setting
- Trust access to VBA project object model
```

### 完了条件

```bash
xlflow doctor --bridge dotnet --json
```

が PowerShell なしで成功する。

---

## Phase 4: `inspect` / `process` を C# bridge 対応

### 目的

低リスクな読み取り系 command を C# 化する。

### 対象

```txt
inspect workbook
inspect sheet
inspect range
process list
process cleanup
```

### 実装内容

```txt
Excel/
  SheetInspector.cs

Windows/
  ProcessGuard.cs
```

### 完了条件

```bash
xlflow inspect --bridge dotnet --json
xlflow process list --bridge dotnet --json
xlflow process cleanup --bridge dotnet --json
```

が動く。

---

## Phase 5: `pull` / `push` を C# bridge 対応

### 目的

VBIDE 操作を C# bridge に移す。

### 実装内容

```txt
Excel/
  VbeProject.cs
  ComponentImporter.cs
  ComponentExporter.cs
```

### 注意点

```txt
- .bas / .cls / .frm / .frx の import/export
- ThisWorkbook / Sheet module の扱い
- UserForm frx 同期
- Rubberduck compatible folder annotation
- changed-only push
- backup mode
```

### 完了条件

```bash
xlflow pull --bridge dotnet --json
xlflow push --bridge dotnet --json
```

が既存 PowerShell bridge と同等に動く。

この Phase では PowerShell との出力差分テストを入れたいです。

---

## Phase 6: `macros` / `run` を C# bridge 対応

### 目的

xlflow の中核である macro 実行を C# 化する。

### 実装内容

```txt
Excel/
  MacroRunner.cs

Runtime/
  RuntimeInjection.cs
  DefinedNameStore.cs
  UiResponseInjection.cs

Windows/
  DialogWatcher.cs
```

### 対応するもの

```txt
- Application.Run
- fully qualified macro name
- args
- timeout
- save/no-save
- runtime mode
- MsgBox/InputBox/FileDialog response injection
- compile/runtime error capture
- modal dialog watcher
```

### 完了条件

```bash
xlflow run Module1.Main --bridge dotnet --json
```

で、

```txt
- 成功時に ok JSON
- 実行時エラーで structured error
- compile error で structured error
- timeout で Excel/dialog 状態つき error
```

が返る。

---

## Phase 7: `test` / `trace` / stream 系を C# bridge 対応

### 目的

agent-native runtime の高度機能を C# bridge に移す。

### 実装内容

```txt
Runtime/
  DebugPipeServer.cs
  UiStreamPipeServer.cs
  TraceInjector.cs
```

### 対象

```txt
- xlflow test
- xlflow trace
- debug stream
- UI stream
- runtime mode injection
```

### 完了条件

```bash
xlflow test --bridge dotnet --json
xlflow trace enable --bridge dotnet --json
```

が動く。

---

## Phase 8: `form` / `export-image` を C# bridge 対応

### 目的

Excel/VBE/UI/clipboard の泥臭い部分を C# に寄せる。

### 対象

```txt
- form build
- inspect form
- export form image
- export sheet/range image
- export shape image
```

### 実装ポイント

```txt
- CopyPicture fallback
- clipboard retry
- window activation
- offscreen/minimized Excel 対策
- GDI/Win32 fallback
```

### 完了条件

```bash
xlflow export-image --bridge dotnet --json
xlflow form inspect --bridge dotnet --json
```

が PowerShell なしで動く。

---

## Phase 9: default bridge を dotnet に変更

### 条件

以下が満たされるまで default を dotnet にしない。

```txt
- doctor
- inspect
- pull
- push
- run
- macros
- process
```

が C# bridge で十分安定していること。

### 変更前

```txt
auto:
  powershell first
  dotnet opt-in
```

### 変更後

```txt
auto:
  dotnet first
  powershell fallback
```

### 完了条件

```bash
xlflow run Main --json
```

が Windows では C# bridge を自動選択する。

---

## Phase 10: PowerShell bridge legacy 化

### 目的

PowerShell を消すのではなく、legacy fallback にする。

### 対応

```txt
- docs/bridge/powershell-bridge-legacy.md を追加
- doctor に legacy status を表示
- warning を出すか検討
- 新機能は原則 C# bridge のみに追加
```

### 完了条件

```txt
- PowerShell bridge は fallback として残る
- 新規 command は C# bridge first
- PowerShell script の肥大化を止める
```

---

# 8. Release / packaging 設計

## Windows archive

```txt
xlflow_Windows_x86_64.zip
  xlflow.exe
  xlflow-excel-bridge.exe
  README.md
  LICENSE
```

## Linux/macOS archive

```txt
xlflow_Linux_x86_64.tar.gz
  xlflow
  README.md
  LICENSE

xlflow_Darwin_arm64.tar.gz
  xlflow
  README.md
  LICENSE
```

C# bridge は Windows release にだけ含める。

---

## C# bridge publish

初期は self-contained single-file。

```powershell
dotnet publish bridge/dotnet/src/Xlflow.ExcelBridge/Xlflow.ExcelBridge.csproj `
  -c Release `
  -r win-x64 `
  -p:PublishSingleFile=true `
  -p:SelfContained=true
```

最初は trimming なしでよいです。
COM / dynamic / reflection / UI Automation と trimming は事故りやすいので、サイズ最適化は後回しが安全です。

---

# 9. CI 設計

## jobs

```txt
go-test:
  - go test ./...

dotnet-test:
  - dotnet test bridge/dotnet/Xlflow.ExcelBridge.sln

windows-integration:
  - Windows runner
  - dotnet publish
  - go test integration subset
  - bridge --version-json
  - xlflow doctor --bridge dotnet --json
```

Excel が必要な integration test は GitHub-hosted runner では難しい可能性があるため、最初は以下で分ける。

```txt
unit:
  GitHub Actions で常時実行

integration:
  self-hosted Windows + Excel 環境で実行
```

---

# 10. テスト戦略

## Go 側

```txt
- bridge resolver test
- bridge protocol compatibility test
- powershell fallback test
- envelope mapping test
- command payload builder test
```

## C# 側 unit test

```txt
- JSON request/response serialization
- command registry
- error mapping
- protocol version check
- path normalization
- response envelope generation
```

## C# 側 integration test

```txt
- Excel activation
- workbook open/save
- inspect range
- pull/push component
- run macro success
- run macro runtime error
- compile error capture
- dialog watcher
```

---

# 11. 失敗時の設計

bridge まわりは失敗パターンをかなり明示した方がよいです。

## bridge not found

```json
{
  "status": "failed",
  "error": {
    "code": "BRIDGE_NOT_FOUND",
    "message": "xlflow-excel-bridge.exe was not found.",
    "phase": "bridge.resolve"
  },
  "suggestions": [
    "Reinstall xlflow from the official Windows release archive.",
    "Ensure xlflow.exe and xlflow-excel-bridge.exe are located in the same directory.",
    "Use `--bridge powershell` to fallback to the legacy PowerShell bridge."
  ]
}
```

## version mismatch

```json
{
  "status": "failed",
  "error": {
    "code": "BRIDGE_VERSION_MISMATCH",
    "message": "xlflow.exe and xlflow-excel-bridge.exe versions do not match.",
    "phase": "bridge.version"
  },
  "suggestions": [
    "Reinstall xlflow from the same release archive.",
    "Check `xlflow doctor --json` for bridge diagnostics."
  ]
}
```

## unsupported command

```json
{
  "status": "failed",
  "error": {
    "code": "BRIDGE_COMMAND_UNSUPPORTED",
    "message": "The selected bridge does not support this command.",
    "phase": "bridge.capability"
  },
  "suggestions": [
    "Use `--bridge auto` to allow fallback.",
    "Use `--bridge powershell` for legacy command support."
  ]
}
```

---

# 12. 開発順序まとめ

実装順としてはこれが一番安全です。

```txt
1. ADR / bridge protocol docs
2. Go bridge provider abstraction
3. PowerShellProvider 化
4. --bridge auto|powershell|dotnet の導入
5. C# bridge skeleton
6. --version-json / --capabilities-json
7. doctor
8. inspect / process
9. pull / push
10. macros
11. run
12. dialog watcher
13. test / trace / runtime injection
14. form / export-image
15. dotnet bridge を default 化
16. PowerShell bridge を legacy fallback 化
```

---

# 13. Issue 分割案

GitHub issue にするなら、以下の粒度がよいです。

```md
## Issue 1: Introduce bridge provider abstraction in Go

### Goal

Refactor the Go Excel runner so that Excel automation is executed through a bridge provider abstraction instead of directly invoking bundled PowerShell scripts.

### Scope

- Add `internal/excel/bridge` package.
- Define `Provider` interface.
- Add `Request` / `Response` contract types.
- Move existing PowerShell invocation behind `PowerShellProvider`.
- Add bridge mode resolution:
  - `auto`
  - `powershell`
  - `dotnet`
- Keep current behavior unchanged by default.

### Non-goals

- Do not add the C# bridge yet.
- Do not change existing JSON output.
- Do not migrate commands to .NET yet.

### Acceptance criteria

- Existing commands still work.
- Existing tests pass.
- `--bridge powershell` works.
- `--bridge auto` falls back to the existing PowerShell behavior.
- PowerShell invocation is isolated behind `PowerShellProvider`.
```

```md
## Issue 2: Add .NET Excel bridge skeleton

### Goal

Add a C#/.NET bridge executable that can be invoked by the Go CLI without using PowerShell.

### Scope

- Add `bridge/dotnet/Xlflow.ExcelBridge.sln`.
- Add `Xlflow.ExcelBridge` console project.
- Implement:
  - `--version-json`
  - `--capabilities-json`
  - stdin/stdout JSON transport
  - protocol version field
  - basic command registry
- Add Go-side `DotnetProvider`.

### Initial capabilities

- `doctor`

### Acceptance criteria

- `xlflow-excel-bridge.exe --version-json` returns structured JSON.
- `xlflow-excel-bridge.exe --capabilities-json` returns supported commands.
- Go can locate and invoke the bridge.
- `xlflow doctor --bridge dotnet --json` can execute the bridge.
```

```md
## Issue 3: Implement .NET doctor command

### Goal

Implement Excel environment diagnostics in the .NET bridge.

### Scope

- Detect OS/runtime/architecture.
- Detect Excel COM availability.
- Start Excel through COM.
- Read Excel version/build.
- Check VBIDE access.
- Return diagnostics through the existing xlflow JSON envelope.

### Acceptance criteria

- `xlflow doctor --bridge dotnet --json` works without launching PowerShell.
- Diagnostic output includes selected bridge and protocol version.
- Excel COM errors are returned as structured errors.
- PowerShell bridge remains available as fallback.
```

```md
## Issue 4: Implement .NET inspect and process commands

### Goal

Migrate low-risk read-only Excel automation commands to the .NET bridge.

### Scope

- Implement workbook/sheet/range inspection.
- Implement Excel process list.
- Implement session-owned process cleanup where applicable.
- Add tests for JSON output compatibility.

### Acceptance criteria

- `xlflow inspect --bridge dotnet --json` works.
- `xlflow process list --bridge dotnet --json` works.
- `xlflow process cleanup --bridge dotnet --json` works.
- `--bridge auto` uses dotnet when supported and falls back otherwise.
```

```md
## Issue 5: Implement .NET pull/push commands

### Goal

Move VBA component import/export from PowerShell to the .NET bridge.

### Scope

- Export standard modules, class modules, document modules, and forms.
- Import standard modules, class modules, document modules, and forms.
- Preserve `.frm` / `.frx` handling.
- Preserve folder annotations and source layout behavior.
- Preserve backup and changed-only behavior.

### Acceptance criteria

- `xlflow pull --bridge dotnet --json` works.
- `xlflow push --bridge dotnet --json` works.
- Output is compatible with the current command envelope.
- Existing PowerShell behavior remains available as fallback.
```

```md
## Issue 6: Implement .NET macro run command

### Goal

Move macro execution to the .NET bridge and improve structured error handling.

### Scope

- Invoke macros through Excel COM.
- Support fully qualified macro names.
- Support arguments.
- Support save/no-save.
- Support timeout.
- Return runtime errors as structured JSON.
- Prepare integration point for dialog watcher.

### Acceptance criteria

- `xlflow run Module1.Main --bridge dotnet --json` works.
- Successful runs return `status: ok`.
- Runtime errors return structured error fields.
- Timeouts return actionable suggestions.
```

```md
## Issue 7: Implement .NET dialog watcher and modal error suppression

### Goal

Use C# Win32/UI Automation support to improve modal dialog detection and suppression during macro execution.

### Scope

- Detect Excel runtime error dialogs.
- Detect compile error dialogs.
- Capture dialog text/buttons.
- Close or respond to supported dialogs.
- Include dialog snapshot in JSON error output.

### Acceptance criteria

- Runtime error dialogs do not block agent execution.
- Compile error dialogs are detected and reported.
- Dialog text is included in structured diagnostics.
- `xlflow run --bridge dotnet --json` returns reliably after dialog failures.
```

```md
## Issue 8: Migrate test/trace/runtime injection to .NET bridge

### Goal

Move xlflow runtime injection and test/trace support to the .NET bridge.

### Scope

- Defined name based runtime mode injection.
- MsgBox/InputBox/FileDialog response injection.
- Debug stream pipe.
- UI stream pipe.
- Trace module injection.

### Acceptance criteria

- `xlflow test --bridge dotnet --json` works.
- `xlflow trace enable --bridge dotnet --json` works.
- Existing runtime mode behavior is preserved.
- UI/debug stream output is merged into the normal xlflow envelope.
```

```md
## Issue 9: Migrate form and export-image commands to .NET bridge

### Goal

Move UserForm and image export operations to the .NET bridge for better COM/Win32/clipboard control.

### Scope

- Form inspect.
- Form build/import.
- Form image export.
- Sheet/range/shape image export.
- Clipboard retry and fallback logic.

### Acceptance criteria

- `xlflow form inspect --bridge dotnet --json` works.
- `xlflow export-image --bridge dotnet --json` works.
- Clipboard failures produce actionable JSON errors.
- Existing PowerShell bridge remains available as fallback.
```

```md
## Issue 10: Make .NET bridge the default Windows bridge

### Goal

Switch Windows bridge resolution so that .NET is preferred by default once major commands are supported.

### Scope

- Update `auto` bridge resolution.
- Prefer dotnet bridge on Windows when available.
- Fall back to PowerShell for unsupported commands.
- Update docs and doctor output.
- Mark PowerShell bridge as legacy.

### Acceptance criteria

- `xlflow run --json` uses .NET bridge by default on Windows.
- `xlflow doctor --json` reports bridge selection clearly.
- `--bridge powershell` still works.
- Unsupported .NET commands fall back cleanly in `auto` mode.
```
