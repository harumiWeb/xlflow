package agentskill

import (
	"embed"
	"errors"
	"fmt"
	"io/fs"
	"os"
	"path/filepath"
	"strings"
)

//go:embed all:templates/xlflow
var bundled embed.FS

const Name = "xlflow"

type Provider struct {
	Name        string
	Dir         string
	Description string
}

type InstallOptions struct {
	RootDir string
	Agent   string
	Target  string
	Force   bool
}

type InstallResult struct {
	Path    string `json:"path"`
	Agent   string `json:"agent,omitempty"`
	Created bool   `json:"created"`
}

var providers = []Provider{
	{Name: "agents", Dir: ".agents/skills", Description: "Repository-local shared agent skills"},
	{Name: "codex", Dir: ".codex/skills", Description: "Codex project skills"},
	{Name: "claude", Dir: ".claude/skills", Description: "Claude project skills"},
	{Name: "cursor", Dir: ".cursor/skills", Description: "Cursor project skills"},
	{Name: "gemini", Dir: ".gemini/skills", Description: "Gemini project skills"},
}

func Providers() []Provider {
	out := make([]Provider, len(providers))
	copy(out, providers)
	return out
}

func ProviderByName(name string) (Provider, bool) {
	name = strings.ToLower(strings.TrimSpace(name))
	for _, provider := range providers {
		if provider.Name == name {
			return provider, true
		}
	}
	return Provider{}, false
}

func Install(opts InstallOptions) (InstallResult, error) {
	root := opts.RootDir
	if root == "" {
		root = "."
	}
	if opts.Agent != "" && opts.Target != "" {
		return InstallResult{}, errors.New("--agent and --target cannot be combined")
	}

	agent := strings.ToLower(strings.TrimSpace(opts.Agent))
	target := strings.TrimSpace(opts.Target)
	if target == "" {
		provider, ok := ProviderByName(agent)
		if !ok {
			return InstallResult{}, fmt.Errorf("unsupported skill agent %q", opts.Agent)
		}
		target = provider.Dir
		agent = provider.Name
	}

	dest := target
	if !filepath.IsAbs(dest) {
		dest = filepath.Join(root, dest)
	}
	dest = filepath.Join(dest, Name)

	if _, err := os.Stat(dest); err == nil {
		if !opts.Force {
			return InstallResult{}, fmt.Errorf("refusing to overwrite existing skill: %s", dest)
		}
		if err := os.RemoveAll(dest); err != nil {
			return InstallResult{}, err
		}
	} else if !errors.Is(err, os.ErrNotExist) {
		return InstallResult{}, err
	}

	if err := copyBundled(dest); err != nil {
		return InstallResult{}, err
	}
	return InstallResult{
		Path:    filepath.ToSlash(rel(root, dest)),
		Agent:   agent,
		Created: true,
	}, nil
}

func copyBundled(dest string) error {
	return fs.WalkDir(bundled, "templates/xlflow", func(path string, d fs.DirEntry, err error) error {
		if err != nil {
			return err
		}
		relPath, err := filepath.Rel("templates/xlflow", path)
		if err != nil {
			return err
		}
		if relPath == "." {
			return nil
		}
		outPath := filepath.Join(dest, relPath)
		if d.IsDir() {
			return os.MkdirAll(outPath, 0o755)
		}
		if err := os.MkdirAll(filepath.Dir(outPath), 0o755); err != nil {
			return err
		}
		body, err := bundled.ReadFile(path)
		if err != nil {
			return err
		}
		return os.WriteFile(outPath, body, 0o644)
	})
}

func rel(base, path string) string {
	r, err := filepath.Rel(base, path)
	if err != nil {
		return path
	}
	return r
}
