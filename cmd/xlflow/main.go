package main

import (
	"os"

	"github.com/harumiWeb/xlflow/internal/cli"
	"github.com/harumiWeb/xlflow/internal/output"
)

func main() {
	if err := cli.Execute(); err != nil {
		os.Exit(output.ExitCode(err))
	}
}
