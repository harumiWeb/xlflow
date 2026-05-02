# デザインドキュメント

## 全体方針

まず `xlflow` は Go 製 CLI なので、基本はこの構成がよいです。

```bash
go install github.com/harumiWeb/xlflow/cmd/xlflow@latest
```

GitHub Releases では以下を配布します。

```text
xlflow_Windows_x86_64.zip
checksums.txt
```

Windows は `.zip`、macOS/Linux は `.tar.gz` が無難です。

Scoop は GitHub Release の Windows zip を参照する manifest を生成します。

---

# 1. GoReleaser の基本構成

`.goreleaser.yaml` はまずこのくらいからで十分です。

```yaml
version: 2

project_name: xlflow

before:
  hooks:
    - go mod tidy

builds:
  - id: xlflow
    main: ./cmd/xlflow
    binary: xlflow
    env:
      - CGO_ENABLED=0
    flags:
      - -trimpath
    ldflags:
      - -s -w
      - -X main.version={{ .Version }}
      - -X main.commit={{ .Commit }}
      - -X main.date={{ .Date }}
    goos:
      - windows
      - darwin
      - linux
    goarch:
      - amd64
      - arm64
    ignore:
      - goos: windows
        goarch: arm64

archives:
  - id: default
    formats: [tar.gz]
    name_template: >-
      {{ .ProjectName }}_
      {{- .Os }}_
      {{- if eq .Arch "amd64" }}x86_64
      {{- else }}{{ .Arch }}{{ end }}
    format_overrides:
      - goos: windows
        formats: [zip]
    files:
      - LICENSE
      - README.md

checksum:
  name_template: checksums.txt

snapshot:
  version_template: "{{ incpatch .Version }}-next"

changelog:
  sort: asc
  filters:
    exclude:
      - "^docs:"
      - "^test:"
      - "^chore:"
```

`main.version` などは、実装側でこう受ける形にします。

```go
package main

var (
	version = "dev"
	commit  = "none"
	date    = "unknown"
)
```

CLI 側に `xlflow version` があると、配布後の検証がかなり楽です。

```bash
xlflow version
```

---

# 2. GitHub Actions

`.github/workflows/release.yml` は以下でよいです。

```yaml
name: Release

on:
  push:
    tags:
      - "v*"

permissions:
  contents: write

concurrency:
  group: release-${{ github.ref }}
  cancel-in-progress: false

jobs:
  release:
    runs-on: ubuntu-latest

    steps:
      - name: Check out repository
        uses: actions/checkout@v4
        with:
          fetch-depth: 0

      - name: Set up Go
        uses: actions/setup-go@v5
        with:
          go-version-file: go.mod
          cache: true

      - name: Run tests
        run: go test ./...

      - name: Build
        run: go build ./...

      - name: Run GoReleaser
        uses: goreleaser/goreleaser-action@v6
        with:
          distribution: goreleaser
          version: "~> v2"
          args: release --clean
        env:
          GITHUB_TOKEN: ${{ secrets.GITHUB_TOKEN }}
```

これで以下を実行するとリリースされます。

```bash
git tag v0.1.0
git push origin v0.1.0
```

---

# 3. Scoop 配布の考え方

Scoop には大きく 2 パターンあります。

## パターン A: 自分の bucket を作る

最初はこれが一番おすすめです。

```text
harumiWeb/scoop-bucket
```

ユーザーにはこう案内します。

```bash
scoop bucket add harumiweb https://github.com/harumiWeb/scoop-bucket
scoop install xlflow
```

メリットは、審査待ちがなく、すぐ配布できることです。

## パターン B: Scoop 本家 bucket に PR する

将来的にはありですが、最初から狙わなくてよいです。

本家 bucket に入れるにはある程度の実績、安定性、命名、manifest 品質が見られます。`xlflow` は開発者向け CLI なので、まずは自前 bucket で十分です。

---

# 4. Scoop bucket を作る

新しい GitHub リポジトリを作ります。

```text
harumiWeb/scoop-bucket
```

構成はシンプルです。

```text
scoop-bucket/
├─ bucket/
│  └─ xlflow.json
└─ README.md
```

`bucket/xlflow.json` はこのような内容になります。

```json
{
  "version": "0.1.0",
  "description": "AI-Agent-ready CLI framework for editing, testing, running, tracing, and diffing Excel VBA projects.",
  "homepage": "https://github.com/harumiWeb/xlflow",
  "license": "MIT",
  "architecture": {
    "64bit": {
      "url": "https://github.com/harumiWeb/xlflow/releases/download/v0.1.0/xlflow_windows_x86_64.zip",
      "hash": "sha256:REPLACE_ME"
    }
  },
  "bin": "xlflow.exe",
  "checkver": {
    "github": "https://github.com/harumiWeb/xlflow"
  },
  "autoupdate": {
    "architecture": {
      "64bit": {
        "url": "https://github.com/harumiWeb/xlflow/releases/download/v$version/xlflow_windows_x86_64.zip"
      }
    }
  }
}
```

ただし、重要なのは **GoReleaser の archive 名と Scoop manifest の URL を必ず一致させること**です。

上の GoReleaser 設定だと Windows artifact 名はおそらく以下になります。

```text
xlflow_windows_x86_64.zip
```

Scoop の URL もそれに合わせます。

---

# 5. GoReleaser で Scoop manifest を自動更新する

GoReleaser には Scoop manifest を別リポジトリに更新する機能があります。

`.goreleaser.yaml` に `scoops` を追加します。

```yaml
scoops:
  - name: xlflow
    repository:
      owner: harumiWeb
      name: scoop-bucket
      branch: main
      token: "{{ .Env.SCOOP_BUCKET_GITHUB_TOKEN }}"
    directory: bucket
    homepage: https://github.com/harumiWeb/xlflow
    description: AI-Agent-ready CLI framework for editing, testing, running, tracing, and diffing Excel VBA projects.
    license: MIT
```

GitHub Actions 側に token を追加します。

```yaml
env:
  GITHUB_TOKEN: ${{ secrets.GITHUB_TOKEN }}
  SCOOP_BUCKET_GITHUB_TOKEN: ${{ secrets.SCOOP_BUCKET_GITHUB_TOKEN }}
```

全体ではこのようになります。

```yaml
- name: Run GoReleaser
  uses: goreleaser/goreleaser-action@v6
  with:
    distribution: goreleaser
    version: "~> v2"
    args: release --clean
  env:
    GITHUB_TOKEN: ${{ secrets.GITHUB_TOKEN }}
    SCOOP_BUCKET_GITHUB_TOKEN: ${{ secrets.SCOOP_BUCKET_GITHUB_TOKEN }}
```

`SCOOP_BUCKET_GITHUB_TOKEN` は Personal Access Token で、`scoop-bucket` リポジトリへ push できる権限が必要です。

fine-grained PAT なら対象リポジトリを `scoop-bucket` に限定し、Contents read/write を付ければ十分です。

---

# 6. Scoop の README

`scoop-bucket` の README は最低限これでよいです。

````md
# harumiWeb Scoop Bucket

```powershell
scoop bucket add harumiweb https://github.com/harumiWeb/scoop-bucket
scoop install xlflow
```
````

## Packages

- `xlflow` - AI-Agent-ready CLI framework for Excel VBA development

````

`xlflow` 本体 README にはこう書くとよいです。

```md
## Installation

### Go install

```bash
go install github.com/harumiWeb/xlflow/cmd/xlflow@latest
````

### Scoop

```powershell
scoop bucket add harumiweb https://github.com/harumiWeb/scoop-bucket
scoop install xlflow
```

### GitHub Releases

Download prebuilt binaries from:

[https://github.com/harumiWeb/xlflow/releases](https://github.com/harumiWeb/xlflow/releases)

````

---

# 7. Scoop のローカル検証

Scoop manifest はローカルで検証できます。

```powershell
scoop bucket add harumiweb https://github.com/harumiWeb/scoop-bucket
scoop install xlflow
xlflow version
````

更新テストはこうです。

```powershell
scoop update
scoop update xlflow
```

manifest の構文チェックには以下も使えます。

```powershell
scoop checkup
```

また、自前 bucket 開発中は直接 manifest を試せます。

```powershell
scoop install .\bucket\xlflow.json
```

---

# 8. winget は Scoop の後でよい

winget は Windows 一般ユーザー向けには強いですが、初回 PR は少し面倒です。

流れはこうです。

1. GitHub Releases で Windows zip または installer を出す
2. `wingetcreate` を使う
3. `microsoft/winget-pkgs` に PR
4. CLA 同意
5. Validation 待ち
6. マージ待ち

ただし CLI ツールの場合、winget より Scoop のほうが開発者には自然です。

`xlflow` のような AI agent / Excel VBA 開発者向け CLI なら、最初の告知導線は以下で十分です。

```powershell
scoop bucket add harumiweb https://github.com/harumiWeb/scoop-bucket
scoop install xlflow
```

その後、落ち着いてから winget 対応でよいです。

---

# 9. 推奨リリース順

現実的にはこの順番が一番安全です。

## Step 1: GitHub Releases だけ完成させる

まず GoReleaser で release artifact が正しく出ることを確認します。

```bash
goreleaser release --snapshot --clean
```

ローカルで成功したらタグを切ります。

```bash
git tag v0.1.0
git push origin v0.1.0
```

## Step 2: Scoop bucket を手動で作る

最初の `xlflow.json` は手で作ってもよいです。

```text
harumiWeb/scoop-bucket
```

## Step 3: Scoop でインストール確認

```powershell
scoop bucket add harumiweb https://github.com/harumiWeb/scoop-bucket
scoop install xlflow
xlflow version
xlflow doctor
```

## Step 4: GoReleaser から Scoop bucket 自動更新

`v0.1.1` 以降で自動更新を試すのが安全です。

## Step 5: winget

Scoop と GitHub Releases が安定してからで十分です。

---

# 10. xlflow で特に注意すべき点

`xlflow` は Excel/VBA を扱う CLI なので、配布前に Windows で次を確認した方がよいです。

```powershell
xlflow version
xlflow doctor
xlflow init
xlflow pull
xlflow lint
xlflow push
xlflow run
```

特に `doctor` は重要です。

Scoop 経由のユーザーは、Excel COM、VBA プロジェクトへのアクセス許可、Trust Center 設定、Office のビット数などで詰まる可能性があります。

なので README に以下のような注意書きを入れるとよいです。

```md
> [!IMPORTANT]
> xlflow requires Microsoft Excel on Windows for commands that interact with workbooks through Excel COM automation.
>
> Some commands may require enabling "Trust access to the VBA project object model" in Excel Trust Center settings.
```

GitHub README では Alerts を使うと見やすいです。

---

## 結論

`xlflow` の配布はこの構成がかなり良いです。

```text
GitHub Releases
  └─ GoReleaser で自動生成

go install
  └─ Go ユーザー向け

Scoop
  └─ 自前 bucket から開始
  └─ GoReleaser で manifest 自動更新

winget
  └─ Scoop 安定後に追加
```

特に Scoop は、最初から本家 bucket を狙わずに、

```text
harumiWeb/scoop-bucket
```

を作るのがよいです。

そして README にはこの一行を大きく出すのが一番わかりやすいです。

```powershell
scoop bucket add harumiweb https://github.com/harumiWeb/scoop-bucket
scoop install xlflow
```
