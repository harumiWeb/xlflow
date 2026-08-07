package main

import (
	"context"
	"flag"
	"fmt"
	"os"
	"os/signal"

	"github.com/harumiWeb/xlflow/internal/staticanalysis/corpus"
)

func main() {
	if len(os.Args) < 2 || os.Args[1] != "sync" {
		fmt.Fprintln(os.Stderr, "usage: xlflow-static-analysis-corpus sync [--manifest path] [--corpus-root path] [--upstream-checkout path]")
		os.Exit(2)
	}
	flags := flag.NewFlagSet("sync", flag.ContinueOnError)
	flags.SetOutput(os.Stderr)
	manifest := flags.String("manifest", "testdata/static-analysis-corpus/manifest.json", "corpus manifest")
	corpusRoot := flags.String("corpus-root", "testdata/static-analysis-corpus", "corpus root")
	checkout := flags.String("upstream-checkout", "", "local checkout at the pinned upstream commit")
	if err := flags.Parse(os.Args[2:]); err != nil {
		os.Exit(2)
	}
	if flags.NArg() != 0 {
		fmt.Fprintln(os.Stderr, "sync does not accept positional arguments")
		os.Exit(2)
	}
	ctx, stop := signal.NotifyContext(context.Background(), os.Interrupt)
	defer stop()
	if err := corpus.Sync(ctx, corpus.SyncOptions{ManifestPath: *manifest, CorpusRoot: *corpusRoot, UpstreamCheckout: *checkout}); err != nil {
		fmt.Fprintln(os.Stderr, err)
		os.Exit(1)
	}
}
