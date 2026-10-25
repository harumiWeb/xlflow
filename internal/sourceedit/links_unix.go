//go:build unix

package sourceedit

func platformIsReparsePoint(string) (bool, error) {
	return false, nil
}
