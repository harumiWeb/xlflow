package corpus

import (
	"context"
	"crypto/sha256"
	"encoding/hex"
	"errors"
	"fmt"
	"io"
	"os"
	"os/exec"
	"path/filepath"
	"strings"
)

// SyncOptions controls the developer-only corpus materialisation operation.
// UpstreamCheckout is intended for offline, locally reviewed updates. When it
// is empty, Sync fetches the exact commit recorded by the manifest.
type SyncOptions struct {
	ManifestPath     string
	CorpusRoot       string
	UpstreamCheckout string
}

// Sync validates the manifest, materialises every listed project in a private
// staging tree, and atomically replaces projects/third_party. It never edits
// the manifest and does not perform network access when UpstreamCheckout is
// supplied.
func Sync(ctx context.Context, opts SyncOptions) error {
	if strings.TrimSpace(opts.ManifestPath) == "" {
		return errors.New("corpus sync requires a manifest path")
	}
	manifest, manifestRoot, err := LoadManifest(opts.ManifestPath)
	if err != nil {
		return err
	}
	corpusRoot := manifestRoot
	if strings.TrimSpace(opts.CorpusRoot) != "" {
		corpusRoot, err = filepath.Abs(opts.CorpusRoot)
		if err != nil {
			return fmt.Errorf("resolve corpus root: %w", err)
		}
	}
	if err := ensureDirectory(corpusRoot); err != nil {
		return fmt.Errorf("corpus root: %w", err)
	}
	if err := ensureManagedTreeClean(corpusRoot); err != nil {
		return err
	}

	checkout, ownedCheckout, err := resolveCheckout(ctx, manifest, opts.UpstreamCheckout)
	if err != nil {
		return err
	}
	if ownedCheckout {
		defer func() { _ = os.RemoveAll(checkout) }()
	}
	if err := verifyCheckoutHead(ctx, checkout, manifest.Upstream.Commit); err != nil {
		return err
	}

	stageRoot, err := os.MkdirTemp(corpusRoot, ".static-analysis-corpus-stage-")
	if err != nil {
		return fmt.Errorf("create corpus staging directory: %w", err)
	}
	defer func() { _ = os.RemoveAll(stageRoot) }()
	stageTarget := filepath.Join(stageRoot, "projects", "third_party")
	if err := os.MkdirAll(stageTarget, 0o755); err != nil {
		return fmt.Errorf("create corpus staging tree: %w", err)
	}
	for _, project := range manifest.Projects {
		source, err := confinedJoin(checkout, project.Source.Path)
		if err != nil {
			return fmt.Errorf("project %q source: %w", project.ID, err)
		}
		destination, err := confinedJoin(stageRoot, project.Path)
		if err != nil {
			return fmt.Errorf("project %q destination: %w", project.ID, err)
		}
		if err := copyProject(source, destination, checkout); err != nil {
			return fmt.Errorf("copy project %q: %w", project.ID, err)
		}
		if err := verifyProjectMetadata(destination, project); err != nil {
			return fmt.Errorf("project %q metadata: %w", project.ID, err)
		}
	}
	if err := publishManagedTree(stageTarget, filepath.Join(corpusRoot, "projects", "third_party"), corpusRoot); err != nil {
		return err
	}
	return nil
}

func resolveCheckout(ctx context.Context, manifest Manifest, requested string) (string, bool, error) {
	if strings.TrimSpace(requested) != "" {
		checkout, err := filepath.Abs(requested)
		if err != nil {
			return "", false, fmt.Errorf("resolve upstream checkout: %w", err)
		}
		if err := ensureDirectory(checkout); err != nil {
			return "", false, fmt.Errorf("upstream checkout: %w", err)
		}
		return checkout, false, nil
	}
	tempRoot, err := os.MkdirTemp("", "xlflow-static-analysis-corpus-")
	if err != nil {
		return "", false, fmt.Errorf("create upstream checkout: %w", err)
	}
	if err := os.Remove(tempRoot); err != nil {
		return "", false, fmt.Errorf("prepare upstream checkout: %w", err)
	}
	if err := os.MkdirAll(tempRoot, 0o755); err != nil {
		return "", false, fmt.Errorf("prepare upstream checkout: %w", err)
	}
	repoURL := "https://github.com/" + manifest.Upstream.Repository + ".git"
	if err := runGit(ctx, tempRoot, "init", "--quiet"); err != nil {
		_ = os.RemoveAll(tempRoot)
		return "", false, fmt.Errorf("initialize upstream checkout: %w", err)
	}
	if err := runGit(ctx, tempRoot, "remote", "add", "origin", repoURL); err != nil {
		_ = os.RemoveAll(tempRoot)
		return "", false, fmt.Errorf("configure upstream checkout: %w", err)
	}
	if err := runGit(ctx, tempRoot, "fetch", "--quiet", "--depth", "1", "origin", manifest.Upstream.Commit); err != nil {
		_ = os.RemoveAll(tempRoot)
		return "", false, fmt.Errorf("fetch pinned upstream commit: %w", err)
	}
	if err := runGit(ctx, tempRoot, "checkout", "--quiet", "--detach", "FETCH_HEAD"); err != nil {
		_ = os.RemoveAll(tempRoot)
		return "", false, fmt.Errorf("checkout pinned upstream commit: %w", err)
	}
	return tempRoot, true, nil
}

func verifyCheckoutHead(ctx context.Context, checkout, want string) error {
	out, err := runGitOutput(ctx, checkout, "rev-parse", "HEAD")
	if err != nil {
		return fmt.Errorf("read upstream checkout HEAD: %w", err)
	}
	got := strings.TrimSpace(out)
	if got != want {
		return fmt.Errorf("upstream checkout HEAD %q does not match pinned commit %q", got, want)
	}
	return nil
}

func runGit(ctx context.Context, dir string, args ...string) error {
	_, err := runGitOutput(ctx, dir, args...)
	return err
}

func runGitOutput(ctx context.Context, dir string, args ...string) (string, error) {
	cmd := exec.CommandContext(ctx, "git", args...)
	cmd.Dir = dir
	var output strings.Builder
	cmd.Stdout = &output
	cmd.Stderr = &output
	if err := cmd.Run(); err != nil {
		return output.String(), fmt.Errorf("git %s: %w (%s)", strings.Join(args, " "), err, strings.TrimSpace(output.String()))
	}
	return output.String(), nil
}

func ensureDirectory(path string) error {
	info, err := os.Stat(path)
	if err != nil {
		return err
	}
	if !info.IsDir() {
		return fmt.Errorf("%q is not a directory", path)
	}
	return nil
}

func ensureManagedTreeClean(corpusRoot string) error {
	repoRoot, err := findGitRoot(corpusRoot)
	if err != nil {
		return nil
	}
	target := filepath.Join(corpusRoot, "projects", "third_party")
	rel, err := filepath.Rel(repoRoot, target)
	if err != nil || filepath.IsAbs(rel) || rel == ".." || strings.HasPrefix(rel, ".."+string(filepath.Separator)) {
		return nil
	}
	out, err := runGitOutput(context.Background(), repoRoot, "status", "--porcelain", "--", filepath.ToSlash(rel))
	if err != nil {
		return fmt.Errorf("check managed corpus tree: %w", err)
	}
	if strings.TrimSpace(out) != "" {
		return fmt.Errorf("managed corpus tree has uncommitted changes: %s", strings.TrimSpace(out))
	}
	return nil
}

func findGitRoot(start string) (string, error) {
	current, err := filepath.Abs(start)
	if err != nil {
		return "", err
	}
	for {
		candidate := filepath.Join(current, ".git")
		if _, err := os.Stat(candidate); err == nil {
			return current, nil
		}
		parent := filepath.Dir(current)
		if parent == current {
			return "", os.ErrNotExist
		}
		current = parent
	}
}

func confinedJoin(root, relative string) (string, error) {
	rootAbs, err := filepath.Abs(root)
	if err != nil {
		return "", err
	}
	joined := filepath.Join(rootAbs, filepath.FromSlash(relative))
	joinedAbs, err := filepath.Abs(joined)
	if err != nil {
		return "", err
	}
	rel, err := filepath.Rel(rootAbs, joinedAbs)
	if err != nil || rel == ".." || strings.HasPrefix(rel, ".."+string(filepath.Separator)) || filepath.IsAbs(rel) {
		return "", fmt.Errorf("path %q escapes root", relative)
	}
	return joinedAbs, nil
}

func copyProject(source, destination, checkout string) error {
	info, err := os.Lstat(source)
	if err != nil {
		return err
	}
	if !info.IsDir() {
		return errors.New("source project is not a directory")
	}
	if err := ensureResolvedWithin(source, checkout); err != nil {
		return err
	}
	if err := os.MkdirAll(destination, 0o755); err != nil {
		return err
	}
	return filepath.WalkDir(source, func(current string, entry os.DirEntry, walkErr error) error {
		if walkErr != nil {
			return walkErr
		}
		if entry.Name() == ".git" {
			return fmt.Errorf(".git content is not allowed: %s", current)
		}
		if entry.Type()&os.ModeSymlink != 0 {
			return fmt.Errorf("symlink/reparse content is not allowed: %s", current)
		}
		if err := ensureResolvedWithin(current, checkout); err != nil {
			return err
		}
		rel, err := filepath.Rel(source, current)
		if err != nil {
			return err
		}
		if rel == "." {
			return nil
		}
		target := filepath.Join(destination, rel)
		if entry.IsDir() {
			return os.MkdirAll(target, 0o755)
		}
		fileInfo, err := entry.Info()
		if err != nil {
			return err
		}
		if !fileInfo.Mode().IsRegular() {
			return fmt.Errorf("irregular file is not allowed: %s", current)
		}
		return copyFile(current, target)
	})
}

func ensureResolvedWithin(candidate, root string) error {
	resolved, err := filepath.EvalSymlinks(candidate)
	if err != nil {
		return err
	}
	return ensurePathWithin(resolved, root)
}

func ensurePathWithin(candidate, root string) error {
	rootAbs, err := filepath.Abs(root)
	if err != nil {
		return err
	}
	candidateAbs, err := filepath.Abs(candidate)
	if err != nil {
		return err
	}
	rel, err := filepath.Rel(rootAbs, candidateAbs)
	if err != nil || rel == ".." || strings.HasPrefix(rel, ".."+string(filepath.Separator)) || filepath.IsAbs(rel) {
		return fmt.Errorf("resolved path %q escapes root %q", candidate, root)
	}
	return nil
}

func copyFile(source, destination string) error {
	in, err := os.Open(source)
	if err != nil {
		return err
	}
	defer func() { _ = in.Close() }()
	out, err := os.OpenFile(destination, os.O_CREATE|os.O_EXCL|os.O_WRONLY, 0o644)
	if err != nil {
		return err
	}
	if _, err := io.Copy(out, in); err != nil {
		_ = out.Close()
		return err
	}
	return out.Close()
}

func verifyProjectMetadata(destination string, project Project) error {
	for label, relative := range map[string]string{"license_file": project.Provenance.LicenseFile, "source_file": project.Provenance.SourceFile} {
		path, err := confinedJoin(destination, relative)
		if err != nil {
			return fmt.Errorf("%s: %w", label, err)
		}
		info, err := os.Lstat(path)
		if err != nil {
			return fmt.Errorf("%s %q: %w", label, relative, err)
		}
		if !info.Mode().IsRegular() {
			return fmt.Errorf("%s %q is not a regular file", label, relative)
		}
	}
	return nil
}

func publishManagedTree(staged, target, corpusRoot string) error {
	return publishManagedTreeWith(staged, target, corpusRoot, os.Rename)
}

func publishManagedTreeWith(staged, target, corpusRoot string, rename func(string, string) error) error {
	if err := os.MkdirAll(filepath.Dir(target), 0o755); err != nil {
		return fmt.Errorf("create corpus destination: %w", err)
	}
	if err := ensureExistingAncestorWithin(filepath.Dir(target), corpusRoot); err != nil {
		return fmt.Errorf("corpus destination boundary: %w", err)
	}
	if info, err := os.Lstat(target); err == nil && info.Mode()&os.ModeSymlink != 0 {
		return fmt.Errorf("managed corpus destination is a symlink: %s", target)
	} else if err != nil && !errors.Is(err, os.ErrNotExist) {
		return fmt.Errorf("inspect managed corpus destination: %w", err)
	}
	backup, err := os.MkdirTemp(filepath.Dir(target), ".static-analysis-corpus-backup-")
	if err != nil {
		return fmt.Errorf("create corpus backup path: %w", err)
	}
	if err := os.Remove(backup); err != nil {
		return fmt.Errorf("prepare corpus backup path: %w", err)
	}
	hadTarget := false
	if _, err := os.Lstat(target); err == nil {
		hadTarget = true
		if err := rename(target, backup); err != nil {
			return fmt.Errorf("stage existing corpus tree: %w", err)
		}
	}
	if err := rename(staged, target); err != nil {
		if hadTarget {
			if restoreErr := rename(backup, target); restoreErr != nil {
				return fmt.Errorf("publish corpus tree: %w; restore old tree: %v", err, restoreErr)
			}
		}
		return fmt.Errorf("publish corpus tree: %w", err)
	}
	if hadTarget {
		if err := os.RemoveAll(backup); err != nil {
			return fmt.Errorf("remove corpus backup: %w", err)
		}
	}
	return nil
}

func ensureExistingAncestorWithin(candidate, root string) error {
	rootResolved, err := filepath.EvalSymlinks(root)
	if err != nil {
		return err
	}
	current := candidate
	for {
		if _, err := os.Lstat(current); err == nil {
			resolved, err := filepath.EvalSymlinks(current)
			if err != nil {
				return err
			}
			return ensurePathWithin(resolved, rootResolved)
		} else if !errors.Is(err, os.ErrNotExist) {
			return err
		}
		parent := filepath.Dir(current)
		if parent == current {
			return os.ErrNotExist
		}
		current = parent
	}
}

// TreeDigest returns a deterministic digest of regular files below root. It
// is useful for tests and developer verification, not for manifest identity.
func TreeDigest(root string) (string, error) {
	type item struct{ name, digest string }
	items := make([]item, 0)
	err := filepath.WalkDir(root, func(path string, entry os.DirEntry, err error) error {
		if err != nil {
			return err
		}
		if entry.IsDir() {
			return nil
		}
		info, err := entry.Info()
		if err != nil {
			return err
		}
		if !info.Mode().IsRegular() {
			return fmt.Errorf("digest cannot include irregular file %s", path)
		}
		body, err := os.ReadFile(path)
		if err != nil {
			return err
		}
		sum := sha256.Sum256(body)
		rel, err := filepath.Rel(root, path)
		if err != nil {
			return err
		}
		items = append(items, item{filepath.ToSlash(rel), hex.EncodeToString(sum[:])})
		return nil
	})
	if err != nil {
		return "", err
	}
	var builder strings.Builder
	for _, value := range items {
		builder.WriteString(value.name)
		builder.WriteByte('\x00')
		builder.WriteString(value.digest)
		builder.WriteByte('\n')
	}
	sum := sha256.Sum256([]byte(builder.String()))
	return hex.EncodeToString(sum[:]), nil
}
