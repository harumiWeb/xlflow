//go:build !windows

package workbookuse

func Detect(string) (State, error) {
	return State{}, ErrUnsupportedPlatform
}
