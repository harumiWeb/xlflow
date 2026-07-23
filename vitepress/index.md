---
layout: home

hero:
  name: xlflow
  text: Modern Excel VBA development for VS Code, Git, and AI agents.
  tagline: 'Move VBA into a source-controlled workflow with completion, diagnostics, testing, safe Excel sessions, and reproducible automation.<br><span class="hero-install-label">Install on Windows</span><code class="hero-install-command">winget install HarumiWeb.Xlflow</code>'
  image:
    src: /images/icon.png
    alt: xlflow
  actions:
    - theme: brand
      text: Get Started
      link: /getting-started
    - theme: alt
      text: Choose your workflow
      link: /choose-workflow
    - theme: alt
      text: View on GitHub
      link: https://github.com/harumiWeb/xlflow
---

## Start with your goal

<div class="workflow-grid workflow-grid-primary">
  <a href="./tutorials/vscode-development"><strong>Develop in VS Code</strong><span>Use completion, diagnostics, navigation, CodeLens, and tests while editing source files.</span><em>VS Code workflow →</em></a>
  <a href="./tutorials/ai-agent"><strong>Develop with an AI agent</strong><span>Give Codex, Claude Code, or Copilot a structured terminal loop for Excel.</span><em>Agent workflow →</em></a>
  <a href="./tutorials/existing-workbook"><strong>Bring an existing workbook to Git</strong><span>Import an <code>.xlsm</code>, edit source, push, run, inspect, and review the diff.</span><em>Existing workbook tutorial →</em></a>
</div>

<div class="workflow-more">
  <span>Also:</span>
  <a href="./guides/ci">Automate from the CLI</a>
  <a href="./tutorials/wsl">Work from WSL</a>
  <a href="./help/troubleshooting">Troubleshoot a setup or Excel failure</a>
</div>

## Built for confident iteration

<div class="confidence-grid">
  <div><strong>Catch problems early</strong><span>Run formatter, lint, static analysis, and live editor diagnostics before Excel imports source.</span></div>
  <div><strong>Keep Excel responsive</strong><span>Use managed sessions for edit-run-inspect loops and structured diagnostic output instead of waiting behind dialogs.</span></div>
  <div><strong>Prove the result</strong><span>Run tests, inspect workbook ranges, export images, compare artifacts, and keep backups for recovery.</span></div>
</div>

## Explore examples

<div class="demo-grid">
  <a href="./demos/world-news"><img src="/images/world-news.png" alt="World news workbook"><span>World News</span></a>
  <a href="./demos/stock-dashboard"><img src="/images/stock-price.png" alt="Stock Dashboard"><span>Stock Dashboard</span></a>
  <a href="./demos/qr-code"><img src="/images/gen-qrcode.png" alt="QR generator workbook"><span>QR Generator</span></a>
  <a href="./demos/invader-game"><img src="/images/space-invader.gif" alt="UserForm invader game"><span>UserForm Game</span></a>
</div>

<style>
.workflow-grid {
  display: grid;
  grid-template-columns: repeat(3, minmax(0, 1fr));
  gap: 16px;
  margin: 20px 0 12px;
}
.VPHomeHero .tagline .hero-install-label {
  display: block;
  margin-top: 20px;
  font-size: 0.78em;
  font-weight: 600;
  color: var(--vp-c-text-1);
}
.VPHomeHero .tagline .hero-install-command {
  display: inline-block;
  margin-top: 6px;
  padding: 8px 12px;
  overflow-x: auto;
  color: var(--vp-c-text-1);
  border: 1px solid var(--vp-c-divider);
  border-radius: 8px;
  background: var(--vp-c-bg-soft);
  font-size: 0.72em;
  line-height: 1.35;
  white-space: nowrap;
}
.workflow-grid a,
.confidence-grid > div {
  display: flex;
  min-height: 182px;
  flex-direction: column;
  gap: 10px;
  padding: 22px;
  color: inherit;
  border: 1px solid var(--vp-c-divider);
  border-radius: 12px;
  background: var(--vp-c-bg-soft);
  text-decoration: none;
}
.workflow-grid a:hover { border-color: var(--vp-c-brand-1); }
.workflow-grid strong,
.confidence-grid strong { font-size: 1.08rem; }
.workflow-grid span,
.confidence-grid span { color: var(--vp-c-text-2); line-height: 1.55; }
.workflow-grid em { margin-top: auto; color: var(--vp-c-brand-1); font-style: normal; font-weight: 600; }
.workflow-more { display: flex; flex-wrap: wrap; gap: 8px 16px; margin-bottom: 42px; color: var(--vp-c-text-2); }
.workflow-more a { color: var(--vp-c-brand-1); font-weight: 600; }
.confidence-grid { display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 16px; margin-top: 20px; }
.confidence-grid > div { min-height: 144px; }
.demo-grid { display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap: 16px; margin-top: 20px; }
.demo-grid a { color: inherit; font-weight: 600; text-decoration: none; }
.demo-grid img { width: 100%; aspect-ratio: 16 / 10; object-fit: cover; border: 1px solid var(--vp-c-divider); border-radius: 10px; }
.demo-grid span { display: block; margin-top: 8px; }
@media (max-width: 960px) {
  .workflow-grid, .confidence-grid { grid-template-columns: 1fr; }
  .workflow-grid a, .confidence-grid > div { min-height: 0; }
  .demo-grid { grid-template-columns: repeat(2, minmax(0, 1fr)); }
}
@media (max-width: 640px) { .demo-grid { grid-template-columns: 1fr; } }
</style>
