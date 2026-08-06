package main

import (
	"context"
	"encoding/json"
	"errors"
	"flag"
	"fmt"
	"io"
	"os"
	"strings"
	"time"

	"github.com/harumiWeb/xlflow/internal/oracle"
)

type caseFlags []string

func (f *caseFlags) String() string { return strings.Join(*f, ",") }
func (f *caseFlags) Set(value string) error {
	value = strings.TrimSpace(value)
	if value == "" {
		return errors.New("--case cannot be empty")
	}
	*f = append(*f, value)
	return nil
}

func main() {
	var cases caseFlags
	var manifest string
	var strict, promote bool
	var timeout time.Duration
	meanings := map[string]string{}
	flags := flag.NewFlagSet("xlflow-vbe-oracle", flag.ContinueOnError)
	flags.SetOutput(io.Discard)
	flags.StringVar(&manifest, "manifest", "testdata/vbe-oracle/manifest.json", "oracle fixture manifest")
	flags.Var(&cases, "case", "case ID to execute (repeatable)")
	flags.BoolVar(&strict, "strict", false, "fail when observed behavior differs from asserted expectation")
	flags.BoolVar(&promote, "promote-observed", false, "promote selected observe fixtures")
	flags.DurationVar(&timeout, "timeout", oracle.DefaultTimeout, "per-case timeout")
	var meaningFlags caseFlags
	flags.Var(&meaningFlags, "diagnostic-meaning", "accepted promotion meaning as case-id=specification|policy|maintainability")
	if err := flags.Parse(os.Args[1:]); err != nil {
		writeFailure(2, err.Error())
		return
	}
	for _, value := range meaningFlags {
		id, meaning, ok := strings.Cut(value, "=")
		if !ok || strings.TrimSpace(id) == "" || strings.TrimSpace(meaning) == "" {
			writeFailure(2, "--diagnostic-meaning must be case-id=meaning")
			return
		}
		meanings[strings.TrimSpace(id)] = strings.TrimSpace(meaning)
	}
	report, err := oracle.Run(context.Background(), oracle.Options{
		ManifestPath:      manifest,
		CaseIDs:           cases,
		Strict:            strict,
		PromoteObserved:   promote,
		DiagnosticMeaning: meanings,
		Timeout:           timeout,
	})
	if err != nil {
		code := 3
		var exitErr *oracle.ExitError
		if errors.As(err, &exitErr) {
			code = exitErr.Code
		}
		writeReport(report)
		os.Exit(code)
	}
	writeReport(report)
}

func writeReport(report oracle.Report) {
	body, err := json.MarshalIndent(report, "", "  ")
	if err != nil {
		writeFailure(3, err.Error())
		return
	}
	fmt.Println(string(body))
}

func writeFailure(code int, message string) {
	report := oracle.Report{SchemaVersion: oracle.SchemaVersion, Status: "failed", Outcome: oracle.OutcomeInfrastructureFailure, Error: &oracle.ReportError{Code: "cli_error", Message: message}}
	writeReport(report)
	if code != 0 {
		// Keep this helper testable while ensuring the process exits for CLI use.
		os.Exit(code)
	}
}
