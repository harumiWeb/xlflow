package cli

import (
	"errors"
	"fmt"

	"github.com/harumiWeb/xlflow/internal/lint"
	"github.com/harumiWeb/xlflow/internal/output"
	packpkg "github.com/harumiWeb/xlflow/internal/pack"
)

func (a *app) writePackDiagnostics(err error, blank bool, template string) (bool, error) {
	_, encoding := errors.AsType[*packpkg.SourceEncodingError](err)
	_, aggregate := errors.AsType[*packpkg.DiagnosticsError](err)
	if !encoding && !aggregate {
		return false, nil
	}
	var issues []lint.Issue
	var details []map[string]any
	var collect func(error)
	collect = func(err error) {
		if group, ok := errors.AsType[*packpkg.DiagnosticsError](err); ok {
			for _, item := range group.Errors {
				collect(item)
			}
			return
		}
		if group, ok := errors.AsType[*packpkg.SourceEncodingErrors](err); ok {
			for _, item := range group.Errors {
				collect(item)
			}
			return
		}
		if item, ok := errors.AsType[*packpkg.SourceEncodingError](err); ok {
			suggestion := "The template's VBA codepage is authoritative. Use characters it supports or a template with the intended codepage; UTF-8 source encoding does not change the VBA project encoding."
			origin := "template"
			if blank {
				origin = "pack.blank.code_page"
				suggestion = "Set [pack.blank].code_page in xlflow.toml to the intended VBA project codepage (932 for Japanese). Keep source files as UTF-8."
			}
			issues = append(issues, lint.Issue{Severity: "error", Code: "pack_source_encoding_failed", File: item.Path, Line: item.Line, Column: item.Column, Message: fmt.Sprintf("Character %q (U+%04X) cannot be represented by VBA codepage %d.", item.Rune, item.Rune, item.CodePage), Suggestion: suggestion})
			payload := item.Details()
			payload["path"] = item.Path
			payload["codepage_source"] = origin
			if !blank {
				payload["template"] = displayPath(a.cwd, template)
			}
			details = append(details, payload)
			return
		}
		code := "pack_validation_failed"
		if userFormCode, ok := userFormEngineErrorCode(err); ok {
			code = "pack_" + userFormCode
		}
		issues = append(issues, lint.Issue{Severity: "error", Code: code, Message: err.Error()})
	}
	collect(err)
	code := "pack_validation_failed"
	if !aggregate {
		code = "pack_source_encoding_failed"
	}
	env := output.Failure("pack", output.Error{Code: code, Message: fmt.Sprintf("%d pack input issue(s) must be fixed before generation", len(issues)), Source: "xlflow", Phase: "preflight", Details: details})
	env.Issues = issues
	return true, a.write(env, output.ExitValidation)
}
