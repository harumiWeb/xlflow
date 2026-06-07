---
layout: home

hero:
  name: xlflow
  text: AI-agent-ready Excel VBA development CLI
  tagline: Export, edit, lint, push, run, inspect, test, and review Excel VBA projects from the command line.
  image:
    src: /images/icon.png
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
    details: Export workbook VBA into normal source files, review changes, and import edits back into Excel from a repeatable CLI workflow.
  - title: Built for AI agents
    details: Stable JSON envelopes, explicit exit codes, session-aware commands, diagnostics, and bundled agent skills make automated VBA changes practical.
  - title: Safer Excel automation
    details: Doctor, lint, analyze, inspect, backup, diff, and structured debug workflows make workbook state and runtime failures easier to reason about.
---

## Demos

These samples show the kind of workbook applications xlflow can help build and verify.

<div class="demo-grid">
  <a href="/demos/world-news"><img src="/images/world-news.png" alt="World news workbook"><span>World News</span></a>
  <a href="/demos/stock-dashboard"><img src="/images/stock-price.png" alt="Stock dashboard workbook"><span>Stock Dashboard</span></a>
  <a href="/demos/qr-code"><img src="/images/gen-qrcode.png" alt="QR generator workbook"><span>QR Generator</span></a>
  <a href="/demos/invader-game"><img src="/images/space-invader.gif" alt="UserForm invader game"><span>UserForm Game</span></a>
</div>

<style>
.demo-grid {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
  gap: 16px;
  margin-top: 24px;
}
.demo-grid a {
  color: inherit;
  font-weight: 600;
  text-decoration: none;
}
.demo-grid img {
  width: 100%;
  aspect-ratio: 16 / 10;
  object-fit: cover;
  border: 1px solid var(--vp-c-divider);
  border-radius: 8px;
}
.demo-grid span {
  display: block;
  margin-top: 8px;
}
</style>
