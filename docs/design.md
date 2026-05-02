現状、`scriptPath()` は次の順で `new.ps1` などを探しています。

1. 実行カレント配下の `scripts/new.ps1`
2. 開発時のソースツリー配下の `scripts/new.ps1`
3. `xlflow.exe` と同じディレクトリ配下の `scripts/new.ps1`

つまり、Releaseから `xlflow.exe` だけを取り出して実行すると失敗する構造です。実際に `scripts/new.ps1` や `scripts/run.ps1` などはリポジトリには存在していますが、Release ZIPには含まれていません。GoReleaser設定の `archives.files` も今は `LICENSE` と `README.md` だけです。

## まずやるべき修正

`.goreleaser.yaml` の `archives.files` に `scripts/*.ps1` を追加してください。

```yaml
archives:
  - id: xlflow
    ids:
      - xlflow
    formats:
      - zip
    name_template: >-
      {{ .ProjectName }}_
      {{- .Os }}_
      {{- if eq .Arch "amd64" }}x86_64{{ else }}{{ .Arch }}{{ end }}
    files:
      - LICENSE
      - README.md
      - scripts/*.ps1
```

この場合、Release ZIPはこうなります。

```txt
xlflow.exe
LICENSE
README.md
scripts/
  new.ps1
  doctor.ps1
  pull.ps1
  push.ps1
  run.ps1
  test.ps1
  macros.ps1
  trace.ps1
  session.ps1
  runner.ps1
  attach.ps1
  ui.ps1
```

`scriptPath()` は `xlflow.exe` と同じディレクトリの `scripts/<command>.ps1` を見に行く実装になっているので、この構成なら動きます。

## ただし、これだけだとまだ弱いです

ユーザーがGitHub ReleasesからZIPを落としても、ありがちな操作はこれです。

```txt
ZIPを解凍
↓
xlflow.exe だけを PATH の通った場所へコピー
↓
scripts がないので失敗
```

なので、**Release ZIP同梱だけでは「半分解決」**です。Scoopやwinget経由なら比較的きれいに配置できますが、直DLユーザーには事故りやすいです。

## おすすめの最終形

### 方針A: 短期リリース向け

まずはこれで十分です。

1. Release ZIPに `scripts/*.ps1` を含める
2. READMEに「ZIPを展開したディレクトリごとPATHに通す」と書く
3. `xlflow doctor` で `scripts` の有無を診断する
4. エラーメッセージを改善する

例えばエラーはこうした方が親切です。

```txt
Error: script new.ps1 was not found

xlflow requires bundled PowerShell scripts.
Expected location:
  C:\path\to\xlflow\scripts\new.ps1

If you installed from GitHub Releases, extract the whole ZIP file and keep the scripts directory next to xlflow.exe.
```

現状の `script_not_found` は正しいですが、ユーザーには「exe単体では動かない」が伝わりません。

### 方針B: 本命

**PowerShellスクリプトをGoバイナリに `embed` する**のが一番よいです。

理由は明確です。

```txt
xlflow.exe 単体で動く
go install でも壊れない
GitHub Releases直DLでも壊れない
Scoop/wingetでも構成が単純
ユーザーが scripts を消しても壊れない
```

`go install` を配布経路に入れるなら、特に `embed` はほぼ必須です。なぜなら `go install github.com/harumiWeb/xlflow/cmd/xlflow@latest` では、基本的にユーザーの手元に `scripts/*.ps1` を隣接配置できません。

実装イメージはこうです。

```go
// internal/excel/scripts_embed.go
package excel

import "embed"

//go:embed ../../scripts/*.ps1
var embeddedScripts embed.FS
```

ただし `go:embed` はパッケージディレクトリ外の `../../scripts` を直接埋め込めない制約があるため、実際には次のようにした方がきれいです。

```txt
internal/excel/scripts/
  new.ps1
  doctor.ps1
  pull.ps1
  push.ps1
  run.ps1
  ...
  embed.go
```

```go
package scripts

import "embed"

//go:embed *.ps1
var FS embed.FS
```

そして実行時は、一時ディレクトリに書き出して `powershell -File` で実行します。

```go
func materializeScript(commandName string) (string, func(), error) {
    name := commandName + ".ps1"

    data, err := scripts.FS.ReadFile(name)
    if err != nil {
        return "", nil, fmt.Errorf("embedded script %s was not found", name)
    }

    dir, err := os.MkdirTemp("", "xlflow-scripts-*")
    if err != nil {
        return "", nil, err
    }

    path := filepath.Join(dir, name)
    if err := os.WriteFile(path, data, 0600); err != nil {
        os.RemoveAll(dir)
        return "", nil, err
    }

    cleanup := func() {
        _ = os.RemoveAll(dir)
    }

    return path, cleanup, nil
}
```

`runWithOptions()` 側では、外部ファイルが見つかればそれを使い、見つからなければ埋め込み版にフォールバック、という形が移行しやすいです。

```go
script, err := scriptPath(r.RootDir, commandName)
var cleanup func()

if err != nil {
    script, cleanup, err = materializeScript(commandName)
    if err != nil {
        env = output.Failure(commandName, output.Error{
            Code:    "script_not_found",
            Message: err.Error(),
            Source:  "xlflow",
        })
        return env, output.ExitEnvironment, nil
    }
}

if cleanup != nil {
    defer cleanup()
}
```

## 私ならこう進めます

### v0.1.1 hotfix

まず即修正。

```txt
- GoReleaser ZIPに scripts/*.ps1 を同梱
- READMEのRelease直DL手順を修正
- script_not_found のエラー文を改善
```

これは今日中に直せるタイプの不具合です。

### v0.2.0

次に本命対応。

```txt
- scripts を internal/excel/scripts に移動
- go:embed でバイナリに同梱
- 外部 scripts があれば優先、なければ embed を使用
- go install での動作を保証
- Release ZIPから scripts 同梱を削除、またはデバッグ用として残す
```

外部 `scripts` を優先する設計にしておくと、開発中はPS1だけ差し替えて検証しやすいです。つまり、**開発体験と配布安定性を両立**できます。

## 配布経路ごとの結論

| 配布経路                |      `scripts` 同梱方式 | `go:embed` 方式 |
| ------------------- | ------------------: | ------------: |
| GitHub Releases ZIP | 動くが、exeだけコピーされると壊れる |            強い |
| Scoop               |                  動く |            強い |
| winget              |      動くがインストーラ設計に注意 |            強い |
| go install          |               ほぼ厳しい |           必須級 |
| AIエージェント利用          |       scripts欠落で事故る |            強い |

なので、xlflowのようなCLI-first/AI-agent-firstツールでは、最終的には **単一exeで完結**させるべきです。