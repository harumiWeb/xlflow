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
      { text: "Get Started", link: "/getting-started" },
      {
        text: "Explore",
        items: [
          { text: "Tutorials", link: "/tutorials/" },
          { text: "How-to Guides", link: "/guides/" },
          { text: "VS Code", link: "/vscode/" },
          { text: "AI Agents", link: "/ai-agents/" },
          { text: "Demos", link: "/demos/" },
        ],
      },
      {
        text: "Reference",
        items: [
          { text: "Commands", link: "/commands/" },
          { text: "Technical Reference", link: "/reference/json-output" },
          { text: "Help & Troubleshooting", link: "/help/" },
        ],
      },
      { text: "日本語", link: "/ja/" },
    ],
    sidebar: {
      "/commands/": commandSidebar(),
      "/concepts/": conceptSidebar(),
      "/tutorials/": tutorialSidebar(),
      "/guides/": guideSidebar(),
      "/vscode/": vscodeSidebar(),
      "/help/": helpSidebar(),
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
        { text: "capabilities", link: "/commands/capabilities" },
        { text: "new", link: "/commands/new" },
        { text: "init", link: "/commands/init" },
        { text: "pack", link: "/commands/pack" },
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
        { text: "recovery", link: "/commands/recovery" },
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
        { text: "generate", link: "/commands/generate" },
        { text: "module", link: "/commands/module" },
        { text: "completion", link: "/commands/completion" },
        { text: "process", link: "/commands/process" },
        { text: "skill", link: "/commands/skill" },
        { text: "type-db", link: "/commands/type-db" },
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

function tutorialSidebar() {
  return [
    {
      text: "Tutorials",
      items: [
        { text: "Overview", link: "/tutorials/" },
        { text: "First xlflow project", link: "/tutorials/first-project" },
        { text: "Import an existing workbook", link: "/tutorials/existing-workbook" },
        { text: "Develop VBA in VS Code", link: "/tutorials/vscode-development" },
        { text: "Develop with an AI agent", link: "/tutorials/ai-agent" },
        { text: "Test-driven VBA development", link: "/tutorials/tdd" },
        { text: "Work from WSL", link: "/tutorials/wsl" },
      ],
    },
  ];
}

function guideSidebar() {
  return [
    {
      text: "How-to Guides",
      items: [
        { text: "Overview", link: "/guides/" },
        { text: "Source control", link: "/guides/source-control" },
        { text: "Sessions and recovery", link: "/guides/sessions" },
        { text: "Run and debug VBA", link: "/guides/debugging" },
        { text: "Testing and static analysis", link: "/guides/testing" },
        { text: "Configure VS Code", link: "/guides/vscode" },
        { text: "Use xlflow with AI agents", link: "/guides/ai-agents" },
        { text: "Manage UserForms", link: "/guides/userforms" },
        { text: "Inspect workbooks and formulas", link: "/guides/workbook-inspection" },
        { text: "CI and automation", link: "/guides/ci" },
      ],
    },
  ];
}

function vscodeSidebar() {
  return [
    {
      text: "VS Code extension",
      items: [
        { text: "Overview", link: "/vscode/" },
        { text: "Installation", link: "/vscode/installation" },
        { text: "First project", link: "/vscode/first-project" },
        { text: "Project sidebar", link: "/vscode/sidebar" },
        { text: "Completion and type inference", link: "/vscode/completion" },
        { text: "Diagnostics and navigation", link: "/vscode/diagnostics" },
        { text: "CodeLens and testing", link: "/vscode/codelens-testing" },
        { text: "Commands and settings", link: "/vscode/settings" },
        { text: "Troubleshooting", link: "/vscode/troubleshooting" },
      ],
    },
  ];
}

function helpSidebar() {
  return [
    {
      text: "Help",
      items: [
        { text: "Overview", link: "/help/" },
        { text: "Troubleshooting", link: "/help/troubleshooting" },
        { text: "FAQ", link: "/help/faq" },
        { text: "Known limitations", link: "/help/known-limitations" },
        { text: "Reporting bugs", link: "/help/reporting-bugs" },
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
        { text: "CLI Inventory", link: "/reference/cli-command-inventory" },
        { text: "UserForm Specification", link: "/reference/userform-spec" },
        { text: "Project Structure", link: "/reference/project-structure" },
        { text: "Config File", link: "/reference/config-file" },
        { text: "XlflowUI", link: "/reference/xlflow-ui" },
        { text: "Exit Codes", link: "/reference/exit-codes" },
        { text: "Error Codes", link: "/reference/error-codes" },
        { text: "Diagnostic Inventory", link: "/reference/diagnostics" },
        { text: "Error-code Inventory", link: "/reference/error-code-inventory" },
        { text: "Documentation Maintenance", link: "/reference/documentation-maintenance" },
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
