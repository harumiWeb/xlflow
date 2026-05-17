<!-- 設計メモ -->

## 推奨ディレクトリ設計

xlflow なら、まずはこの構成がよいです。

```txt
xlflow/
├─ cmd/
├─ internal/
├─ scripts/
├─ examples/
├─ vitepress/
│  ├─ .vitepress/
│  │  ├─ config.ts
│  │  └─ theme/
│  │     ├─ index.ts
│  │     └─ custom.css
│  │
│  ├─ public/
│  │  ├─ logo.svg
│  │  ├─ favicon.ico
│  │  └─ images/
│  │
│  ├─ index.md
│  ├─ getting-started.md
│  ├─ installation.md
│  ├─ quickstart.md
│  ├─ concepts/
│  │  ├─ project-model.md
│  │  ├─ workbook-session-source.md
│  │  ├─ source-of-truth.md
│  │  ├─ backup-and-rollback.md
│  │  └─ ai-agent-workflow.md
│  │
│  ├─ commands/
│  │  ├─ index.md
│  │  ├─ new.md
│  │  ├─ init.md
│  │  ├─ doctor.md
│  │  ├─ attach.md
│  │  ├─ pull.md
│  │  ├─ push.md
│  │  ├─ session.md
│  │  ├─ save.md
│  │  ├─ run.md
│  │  ├─ lint.md
│  │  ├─ inspect.md
│  │  ├─ export-image.md
│  │  ├─ ui.md
│  │  └─ forms.md
│  │
│  ├─ guides/
│  │  ├─ ai-agent-first-project.md
│  │  ├─ build-weather-app.md
│  │  ├─ build-news-app.md
│  │  ├─ build-qr-generator.md
│  │  ├─ build-userform-game.md
│  │  ├─ userform-development.md
│  │  ├─ error-handling.md
│  │  └─ ci-local-workflow.md
│  │
│  ├─ reference/
│  │  ├─ json-output.md
│  │  ├─ project-structure.md
│  │  ├─ config-file.md
│  │  ├─ exit-codes.md
│  │  ├─ error-codes.md
│  │  ├─ environment-variables.md
│  │  └─ troubleshooting.md
│  │
│  ├─ ai-agents/
│  │  ├─ overview.md
│  │  ├─ recommended-prompts.md
│  │  ├─ codex.md
│  │  ├─ claude-code.md
│  │  ├─ github-copilot.md
│  │  └─ skills.md
│  │
│  ├─ design/
│  │  ├─ architecture.md
│  │  ├─ command-design.md
│  │  ├─ vba-import-export.md
│  │  ├─ session-runner.md
│  │  └─ linter.md
│  │
│  └─ changelog.md
│
├─ package.json
├─ .github/
│  └─ workflows/
│     └─ vitepress.yml
└─ README.md
```

ポイントは、**commands / guides / reference / ai-agents / demos** を分けることです。

README ではなくドキュメントサイトで一番価値が出るのは、たぶんこの3つです。

1. コマンドの仕様を探せること
2. AI エージェントが迷わず開発ループを回せること
3. 初見ユーザーが「何ができるツールか」をデモで直感的に理解できること

---

## vitepress 配下に置く理由

VitePress は `vitepress/.vitepress/config.ts` のように、プロジェクトルートとは別に vitepress ルートを持たせる構成が自然です。公式でも、VitePress の config は `<root>/.vitepress/config.[ext]` から解決され、TypeScript config も標準対応しています。([VitePress][2])

xlflow のような Go CLI プロジェクトでは、ルート直下に VitePress 関連ファイルを散らすより、`vitepress/` に閉じ込めた方が保守しやすいです。

---

## 最小の `vitepress/.vitepress/config.ts`

まずはこれくらいで十分です。

```ts
import { defineConfig } from 'vitepress'

export default defineConfig({
  lang: 'en-US',
  title: 'xlflow',
  description: 'AI-Agent-ready CLI framework for Excel VBA development',

  // GitHub Pages で https://harumiweb.github.io/xlflow/ に置くなら必要
  base: '/xlflow/',

  cleanUrls: true,
  lastUpdated: true,

  head: [
    ['link', { rel: 'icon', href: '/xlflow/favicon.ico' }],
    ['meta', { name: 'theme-color', content: '#2E7D32' }]
  ],

  themeConfig: {
    logo: '/logo.svg',

    nav: [
      { text: 'Guide', link: '/getting-started' },
      { text: 'Commands', link: '/commands/' },
      { text: 'AI Agents', link: '/ai-agents/' },
      { text: 'Demos', link: '/demos/' },
      { text: 'Reference', link: '/reference/json-output' }
    ],

    sidebar: {
      '/commands/': [
        {
          text: 'Commands',
          items: [
            { text: 'Overview', link: '/commands/' },
            { text: 'new', link: '/commands/new' },
            { text: 'init', link: '/commands/init' },
            { text: 'doctor', link: '/commands/doctor' },
            { text: 'attach', link: '/commands/attach' },
            { text: 'pull', link: '/commands/pull' },
            { text: 'push', link: '/commands/push' },
            { text: 'session', link: '/commands/session' },
            { text: 'save', link: '/commands/save' },
            { text: 'run', link: '/commands/run' },
            { text: 'lint', link: '/commands/lint' },
            { text: 'inspect', link: '/commands/inspect' },
            { text: 'export-image', link: '/commands/export-image' },
            { text: 'ui', link: '/commands/ui' },
            { text: 'forms', link: '/commands/forms' }
          ]
        }
      ],

      '/guides/': [
        {
          text: 'Guides',
          items: [
            { text: 'AI Agent First Project', link: '/guides/ai-agent-first-project' },
            { text: 'Weather App', link: '/guides/build-weather-app' },
            { text: 'News App', link: '/guides/build-news-app' },
            { text: 'QR Generator', link: '/guides/build-qr-generator' },
            { text: 'UserForm Game', link: '/guides/build-userform-game' },
            { text: 'Troubleshooting VBA Errors', link: '/guides/error-handling' }
          ]
        }
      ],

      '/reference/': [
        {
          text: 'Reference',
          items: [
            { text: 'JSON Output', link: '/reference/json-output' },
            { text: 'Project Structure', link: '/reference/project-structure' },
            { text: 'Config File', link: '/reference/config-file' },
            { text: 'Exit Codes', link: '/reference/exit-codes' },
            { text: 'Error Codes', link: '/reference/error-codes' },
            { text: 'Environment Variables', link: '/reference/environment-variables' },
            { text: 'Troubleshooting', link: '/reference/troubleshooting' }
          ]
        }
      ],

      '/ai-agents/': [
        {
          text: 'AI Agents',
          items: [
            { text: 'Overview', link: '/ai-agents/' },
            { text: 'Recommended Prompts', link: '/ai-agents/recommended-prompts' },
            { text: 'Codex', link: '/ai-agents/codex' },
            { text: 'Claude Code', link: '/ai-agents/claude-code' },
            { text: 'GitHub Copilot', link: '/ai-agents/github-copilot' },
            { text: 'Skills', link: '/ai-agents/skills' }
          ]
        }
      ],

      '/demos/': [
        {
          text: 'Demos',
          items: [
            { text: 'Overview', link: '/demos/' },
            { text: 'Weather App', link: '/demos/weather' },
            { text: 'News API App', link: '/demos/newsapi' },
            { text: 'Stock Dashboard', link: '/demos/stock-dashboard' },
            { text: 'QR Code Generator', link: '/demos/qr-code' },
            { text: 'Tetris', link: '/demos/tetris' },
            { text: 'Invader Game', link: '/demos/invader-game' }
          ]
        }
      ]
    },

    socialLinks: [
      { icon: 'github', link: 'https://github.com/harumiWeb/xlflow' }
    ],

    search: {
      provider: 'local'
    },

    editLink: {
      pattern: 'https://github.com/harumiWeb/xlflow/edit/main/vitepress/:path',
      text: 'Edit this page on GitHub'
    },

    footer: {
      message: 'Released under the BSD-3-Clause License.',
      copyright: 'Copyright © 2026 harumiWeb'
    }
  }
})
```

GitHub Pages で `https://harumiweb.github.io/xlflow/` 配下に公開するなら `base: '/xlflow/'` が重要です。独自ドメインでルート公開するなら `base` は不要、または `'/'` でよいです。

---

## トップページ設計

`vitepress/index.md` は、文章中心よりも「何ができるか」を一瞬で伝える構成がよいです。

```md
---
layout: home

hero:
  name: xlflow
  text: AI-Agent-ready Excel VBA development CLI
  tagline: Edit, lint, push, run, inspect, and automate Excel VBA projects from the command line.
  image:
    src: /logo.svg
    alt: xlflow
  actions:
    - theme: brand
      text: Get Started
      link: /getting-started
    - theme: alt
      text: Command Reference
      link: /commands/
    - theme: alt
      text: View on GitHub
      link: https://github.com/harumiWeb/xlflow

features:
  - title: Source-first VBA development
    details: Export and import VBA components as text files, making Excel VBA projects easier to version, review, and automate.
  - title: Built for AI agents
    details: JSON outputs, deterministic workflows, session-based execution, and diagnostics help coding agents develop Excel VBA autonomously.
  - title: Safer Excel automation
    details: Backup, lint, doctor, inspect, and run commands reduce the risk of breaking workbooks during iterative development.
---
```

トップページでは **コマンド名を全部並べない** 方がいいです。README と同じ問題が再発します。

---

## コマンドページのテンプレート

`vitepress/commands/push.md` などは、全コマンドで同じフォーマットにするとかなり見やすくなります。

````md
# xlflow push

Import VBA source files from `src/` into the target Excel workbook.

## Usage

```bash
xlflow push [options]
````

## When to use

Use this command after editing VBA source files on disk and before running or saving the workbook.

## Examples

```bash
xlflow push
xlflow push --json
xlflow push --session --json
```

## Options

| Option      | Description                    |
| ----------- | ------------------------------ |
| `--json`    | Output machine-readable JSON.  |
| `--session` | Use the active xlflow session. |

## Output

```json
{
  "ok": true,
  "command": "push",
  "backup": {
    "created": true,
    "path": ".xlflow/backups/Book_20260517_120000.xlsm"
  }
}
```

## Common failures

| Error                 | Cause                                                     | Fix                                              |
| --------------------- | --------------------------------------------------------- | ------------------------------------------------ |
| `workbook_not_found`  | The target workbook cannot be found.                      | Check the project config or run `xlflow attach`. |
| `vbide_access_denied` | Trust access to the VBA project object model is disabled. | Enable VBIDE access and run `xlflow doctor`.     |

## Related commands

* [`xlflow pull`](./pull.md)
* [`xlflow run`](./run.md)
* [`xlflow session`](./session.md)

````

この形式を全コマンドに適用すると、AI エージェントにも人間にも読みやすいです。

---

## コマンド一覧は将来的に自動生成する

最初は手書きでよいですが、いずれは以下のどちらかを入れると強いです。

### 案A: `xlflow help --json` から生成

```bash
xlflow help --json > vitepress/generated/commands.json
go run ./tools/gen-vitepress
````

生成先：

```txt
vitepress/commands/generated/
├─ new.md
├─ init.md
├─ push.md
├─ run.md
└─ ...
```

ただし、完全自動生成にすると文章が無機質になるので、個人的には次がよいです。

### 案B: YAML metadata + 手書き解説

```txt
vitepress/_data/commands/
├─ new.yaml
├─ init.yaml
├─ push.yaml
├─ run.yaml
└─ lint.yaml
```

例：

```yaml
name: push
summary: Import VBA source files into the workbook.
usage: xlflow push [options]
category: workbook
options:
  - name: --json
    description: Output machine-readable JSON.
  - name: --session
    description: Use the active xlflow session.
related:
  - pull
  - run
  - session
```

この YAML から `Options` や `Usage` だけ生成し、本文は Markdown で手書きにするのがベストです。

---

## xlflow のドキュメントで特に作るべきページ

優先順位はこの順番でよいです。

### 1. `getting-started.md`

初見向け。

````md
# Getting Started

## Install xlflow

## Create a new project

```bash
xlflow new Book.xlsm
````

## Start an Excel session

```bash
xlflow session start
```

## Edit VBA source

## Push and run

```bash
xlflow push --json
xlflow run Main.Main --json
```

## Save the workbook

```bash
xlflow save --session --json
```

````

---

### 2. `concepts/workbook-session-source.md`

これは xlflow 特有なので重要です。

説明すべきこと：

```txt
source files
   ↓ push
Excel workbook
   ↓ session/run/save
Excel instance
   ↓ pull
source files
````

このページがあると、AI エージェントも人間も「今どこが正か」を理解しやすくなります。

---

### 3. `ai-agents/overview.md`

xlflow の差別化ポイントなので、専用カテゴリにした方がいいです。

書くべき内容：

````md
# AI Agent Workflow

xlflow is designed to let coding agents develop Excel VBA projects without manual Excel operations.

## Recommended loop

```bash
xlflow doctor --json
xlflow session start --json
xlflow push --json
xlflow lint --json
xlflow run Main.Main --json
xlflow inspect --json
xlflow save --session --json
````

## Rules for agents

* Prefer `--json`.
* Run `doctor` before debugging environment issues.
* Use `session` for iterative development.
* Use `inspect` to verify workbook state.
* Avoid file picker dialogs and interactive MsgBox flows.

````

ここはかなり強い訴求ポイントになります。

---

### 4. `reference/json-output.md`

AI エージェント向けには必須です。

```md
# JSON Output

Most xlflow commands support `--json`.

## Success response

```json
{
  "ok": true,
  "command": "run",
  "durationMs": 1234
}
````

## Error response

```json
{
  "ok": false,
  "error": {
    "code": "vba_runtime_error",
    "message": "Division by zero",
    "source": "Main.Main",
    "line": 20
  }
}
```

## Error handling policy

Agents should inspect `error.code` first, then follow the suggested action if available.

````

これは xlflow の「AI-Agent-ready」感を強く出せます。

---

### 5. `demos/index.md`

ここは画像多めでよいです。

```md
# Demos

xlflow can be used by AI coding agents to build practical Excel VBA applications.

## Weather API App

![Weather app](/images/demo-weather.png)

## News API App

![News app](/images/demo-news.png)

## QR Code Generator

![QR code generator](/images/demo-qr.png)

## UserForm Invader Game

![Invader game](/images/demo-invader.png)
````

OSS 初見ユーザーには、コマンド仕様よりデモの方が刺さることが多いです。

---

## GitHub Pages デプロイ

GitHub Pages に出すなら `.github/workflows/vitepress.yml` を作ります。

```yaml
name: Deploy vitepress

on:
  push:
    branches:
      - main
    paths:
      - 'vitepress/**'
      - 'package.json'
      - 'package-lock.json'
      - '.github/workflows/vitepress.yml'
  workflow_dispatch:

permissions:
  contents: read
  pages: write
  id-token: write

concurrency:
  group: vitepress
  cancel-in-progress: true

jobs:
  build:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v5

      - uses: actions/setup-node@v6
        with:
          node-version: 22
          cache: npm

      - run: npm ci
      - run: npm run vitepress:build

      - uses: actions/upload-pages-artifact@v4
        with:
          path: vitepress/.vitepress/dist

  deploy:
    environment:
      name: github-pages
      url: ${{ steps.deployment.outputs.page_url }}
    needs: build
    runs-on: ubuntu-latest
    steps:
      - id: deployment
        uses: actions/deploy-pages@v4
```

VitePress はデフォルトでビルド出力を `.vitepress/dist` に置くため、`vitepress/.vitepress/dist` を Pages artifact に渡します。公式にも、dev server cache と production build output は `.vitepress/cache` と `.vitepress/dist` に置かれるため、Git 管理から除外するのが推奨されています。([VitePress][1])

`.gitignore` には追加します。

```gitignore
vitepress/.vitepress/cache
vitepress/.vitepress/dist
```

---

## README との役割分担

README は短くして、こういう構成にした方がいいです。

````md
# xlflow

AI-Agent-ready CLI framework for Excel VBA development.

## Quick demo

画像 or GIF

## Install

go install / GitHub Releases / winget / scoop

## Quick start

```bash
xlflow new Book.xlsm
xlflow session start
xlflow push --json
xlflow run Main.Main --json
````

## Documentation

Full documentation is available at:
[https://harumiweb.github.io/xlflow/](https://harumiweb.github.io/xlflow/)

## Examples

* Weather API App
* News API App
* QR Code Generator
* UserForm Invader Game

````

README は **広告・入口・最短導線** に寄せる。  
VitePress は **体系的な仕様・チュートリアル・リファレンス** に寄せる。  
この分担がよいです。

---

## 最初の実装順

いきなり全部作るより、この順番が安全です。

```txt
Phase 1: サイトの骨組み
- docs/index.md
- docs/getting-started.md
- docs/installation.md
- docs/commands/index.md
- docs/.vitepress/config.ts
- GitHub Pages workflow

Phase 2: コマンドリファレンス
- new / init / doctor / pull / push / run / lint / session / save
- 各ページに usage / examples / json output / common errors を入れる

Phase 3: AI エージェント向け導線
- ai-agents/overview.md
- ai-agents/recommended-prompts.md
- ai-agents/skills.md
- reference/json-output.md
- reference/error-codes.md

Phase 4: 自動生成
- xlflow help --json
- tools/gen-docs
- commands metadata
````

---

[1]: https://vitepress.dev/guide/getting-started "Getting Started | VitePress"
[2]: https://vitepress.dev/reference/site-config "Site Config | VitePress"
