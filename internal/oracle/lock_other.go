//go:build !windows

package oracle

import "context"

type noopBatchLock struct{}

func newBatchLock() (BatchLock, error) {
	return noopBatchLock{}, nil
}

func (noopBatchLock) Acquire(ctx context.Context) (func(), error) {
	if ctx != nil {
		if err := ctx.Err(); err != nil {
			return nil, err
		}
	}
	return func() {}, nil
}
