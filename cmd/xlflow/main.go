package main

import (
	"os"

	"github.com/harumiWeb/xlflow/internal/cli"
	"github.com/harumiWeb/xlflow/internal/output"
)

var (
	version = "dev"
	commit  = "none"
	date    = "unknown"
)

func main() {
	if err := cli.ExecuteWithBuildInfo(cli.BuildInfo{
		Version: version,
		Commit:  commit,
		Date:    date,
	}); err != nil {
		os.Exit(output.ExitCode(err))
	}
}
