import { defineConfig } from "vitepress";

export default defineConfig({
  lang: "en-US",
  title: "xlflow",
  description: "AI-agent-ready Excel VBA development CLI",
  base: "/xlflow/",
  cleanUrls: true,
  lastUpdated: true,
  head: [
    ["link", { rel: "icon", href: "/xlflow/images/icon.png" }],
    ["meta", { name: "theme-color", content: "#2E7D32" }],
  ],
  themeConfig: {
    logo: "/images/icon.png",
    nav: [
      { text: "Guide", link: "/getting-started" },
      { text: "Commands", link: "/commands/" },
      { text: "AI Agents", link: "/ai-agents/" },
      { text: "Demos", link: "/demos/" },
      { text: "Reference", link: "/reference/json-output" },
      { text: "Design", link: "/design/architecture" },
    ],
    sidebar: {
      "/commands/": commandSidebar(),
      "/concepts/": conceptSidebar(),
      "/reference/": referenceSidebar(),
      "/ai-agents/": agentSidebar(),
      "/demos/": demoSidebar(),
      "/design/": designSidebar(),
    },
    search: { provider: "local" },
    socialLinks: [{ icon: "github", link: "https://github.com/harumiWeb/xlflow" }],
    editLink: {
      pattern: "https://github.com/harumiWeb/xlflow/edit/main/vitepress/:path",
      text: "Edit this page on GitHub",
    },
    footer: {
      message: "Released under the MIT License.",
      copyright: "Copyright © 2026 harumiWeb",
    },
  },
});

function commandSidebar() {
  return [
    {
      text: "Commands",
      items: [
        { text: "Overview", link: "/commands/" },
        { text: "new", link: "/commands/new" },
        { text: "init", link: "/commands/init" },
        { text: "doctor", link: "/commands/doctor" },
        { text: "attach", link: "/commands/attach" },
        { text: "backup", link: "/commands/backup" },
        { text: "list", link: "/commands/list" },
        { text: "form", link: "/commands/form" },
        { text: "pull", link: "/commands/pull" },
        { text: "push", link: "/commands/push" },
        { text: "rollback", link: "/commands/rollback" },
        { text: "session", link: "/commands/session" },
        { text: "save", link: "/commands/save" },
        { text: "runner", link: "/commands/runner" },
        { text: "run", link: "/commands/run" },
        { text: "export-image", link: "/commands/export-image" },
        { text: "edit", link: "/commands/edit" },
        { text: "macros", link: "/commands/macros" },
        { text: "ui", link: "/commands/ui" },
        { text: "test", link: "/commands/test" },
        { text: "diff", link: "/commands/diff" },
        { text: "inspect", link: "/commands/inspect" },
        { text: "inspect-gui", link: "/commands/inspect-gui" },
        { text: "lint", link: "/commands/lint" },
        { text: "analyze", link: "/commands/analyze" },
        { text: "check", link: "/commands/check" },
        { text: "completion", link: "/commands/completion" },
        { text: "process", link: "/commands/process" },
        { text: "skill", link: "/commands/skill" },
        { text: "version", link: "/commands/version" },
        { text: "update", link: "/commands/update" },
      ],
    },
  ];
}

function conceptSidebar() {
  return [
    {
      text: "Concepts",
      items: [
        { text: "Project Model", link: "/concepts/project-model" },
        { text: "Workbook, Session, Source", link: "/concepts/workbook-session-source" },
        { text: "Source of Truth", link: "/concepts/source-of-truth" },
        { text: "Backup and Rollback", link: "/concepts/backup-and-rollback" },
        { text: "AI Agent Workflow", link: "/concepts/ai-agent-workflow" },
      ],
    },
  ];
}

function referenceSidebar() {
  return [
    {
      text: "Reference",
      items: [
        { text: "JSON Output", link: "/reference/json-output" },
        { text: "Project Structure", link: "/reference/project-structure" },
        { text: "Config File", link: "/reference/config-file" },
        { text: "XlflowUI", link: "/reference/xlflow-ui" },
        { text: "Exit Codes", link: "/reference/exit-codes" },
        { text: "Error Codes", link: "/reference/error-codes" },
        { text: "Environment Variables", link: "/reference/environment-variables" },
        { text: "Troubleshooting", link: "/reference/troubleshooting" },
      ],
    },
  ];
}

function agentSidebar() {
  return [
    {
      text: "AI Agents",
      items: [
        { text: "Overview", link: "/ai-agents/" },
        { text: "Recommended Prompts", link: "/ai-agents/recommended-prompts" },
        { text: "Codex", link: "/ai-agents/codex" },
        { text: "Claude Code", link: "/ai-agents/claude-code" },
        { text: "GitHub Copilot", link: "/ai-agents/github-copilot" },
        { text: "Skills", link: "/ai-agents/skills" },
      ],
    },
  ];
}

function demoSidebar() {
  return [
    {
      text: "Demos",
      items: [
        { text: "Overview", link: "/demos/" },
        { text: "World News", link: "/demos/world-news" },
        { text: "Stock Dashboard", link: "/demos/stock-dashboard" },
        { text: "QR Generator", link: "/demos/qr-code" },
        { text: "Tetris", link: "/demos/tetris" },
        { text: "Invader Game", link: "/demos/invader-game" },
        { text: "Calendar Picker", link: "/demos/calendar-picker" },
      ],
    },
  ];
}

function designSidebar() {
  return [
    {
      text: "Design",
      items: [
        { text: "Architecture", link: "/design/architecture" },
        { text: "Command Design", link: "/design/command-design" },
        { text: "VBA Import/Export", link: "/design/vba-import-export" },
        { text: "Session Runner", link: "/design/session-runner" },
        { text: "Linter", link: "/design/linter" },
      ],
    },
  ];
}
