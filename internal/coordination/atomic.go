package coordination

import (
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"
)

func writeJSONAtomic(dir, target, tempPattern string, value any) error {
	tmp, err := os.CreateTemp(dir, tempPattern+"*.tmp")
	if err != nil {
		return fmt.Errorf("create metadata: %w", err)
	}
	tmpName := tmp.Name()
	defer func() { _ = os.Remove(tmpName) }()
	if err := tmp.Chmod(0o600); err != nil {
		_ = tmp.Close()
		return fmt.Errorf("secure metadata: %w", err)
	}
	encoder := json.NewEncoder(tmp)
	encoder.SetEscapeHTML(false)
	if err := encoder.Encode(value); err != nil {
		_ = tmp.Close()
		return fmt.Errorf("encode metadata: %w", err)
	}
	if err := tmp.Sync(); err != nil {
		_ = tmp.Close()
		return fmt.Errorf("flush metadata: %w", err)
	}
	if err := tmp.Close(); err != nil {
		return fmt.Errorf("close metadata: %w", err)
	}
	if err := platformAtomicReplace(tmpName, target); err != nil {
		return fmt.Errorf("publish metadata %q: %w", filepath.Base(target), err)
	}
	return nil
}
