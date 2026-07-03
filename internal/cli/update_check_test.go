package cli

import (
	"bytes"
	"context"
	"encoding/json"
	"errors"
	"net/http"
	"net/http/httptest"
	"os"
	"testing"

	"github.com/harumiWeb/xlflow/internal/output"
)

func TestCheckForUpdate(t *testing.T) {
	tests := []struct {
		name           string
		currentVersion string
		latestVersion  string
		want           string
	}{
		{
			name:           "newer release available",
			currentVersion: "1.2.3",
			latestVersion:  "v1.2.4",
			want:           "v1.2.4",
		},
		{
			name:           "same release",
			currentVersion: "v1.2.3",
			latestVersion:  "1.2.3",
			want:           "",
		},
		{
			name:           "dev build skips comparison",
			currentVersion: "dev",
			latestVersion:  "v1.2.4",
			want:           "",
		},
		{
			name:           "stable release is newer than prerelease",
			currentVersion: "1.2.3-beta.1",
			latestVersion:  "v1.2.3",
			want:           "v1.2.3",
		},
		{
			name:           "invalid latest release is ignored",
			currentVersion: "1.2.3",
			latestVersion:  "latest",
			want:           "",
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got, err := checkForUpdate(context.Background(), stubReleaseChecker{
				release: latestRelease{Version: tt.latestVersion, ReleaseURL: "https://example.com/release"},
			}, tt.currentVersion)
			if err != nil {
				t.Fatalf("checkForUpdate() error = %v", err)
			}
			if got.LatestVersion != tt.want {
				t.Fatalf("checkForUpdate() latest = %q, want %q", got.LatestVersion, tt.want)
			}
			if tt.want != "" && got.ReleaseURL != "https://example.com/release" {
				t.Fatalf("checkForUpdate() release URL = %q, want https://example.com/release", got.ReleaseURL)
			}
		})
	}
}

func TestCheckForUpdateBestEffortWarnings(t *testing.T) {
	tests := []struct {
		name           string
		currentVersion string
		checker        stubReleaseChecker
		wantLatest     string
		wantWarning    string
	}{
		{
			name:           "newer release available",
			currentVersion: "1.2.3",
			checker: stubReleaseChecker{
				release: latestRelease{Version: "v1.2.4", ReleaseURL: "https://example.com/v1.2.4"},
			},
			wantLatest: "v1.2.4",
		},
		{
			name:           "current version is not semantic",
			currentVersion: "dev",
			checker:        stubReleaseChecker{release: latestRelease{Version: "v1.2.4"}},
			wantWarning:    `update check skipped: current version "dev" is not a release version`,
		},
		{
			name:           "release lookup fails",
			currentVersion: "1.2.3",
			checker:        stubReleaseChecker{err: errors.New("network blocked")},
			wantWarning:    "update check failed: network blocked",
		},
		{
			name:           "latest version is not semantic",
			currentVersion: "1.2.3",
			checker:        stubReleaseChecker{release: latestRelease{Version: "latest"}},
			wantWarning:    `update check skipped: latest release tag "latest" is not semantic version`,
		},
		{
			name:           "up to date",
			currentVersion: "1.2.3",
			checker:        stubReleaseChecker{release: latestRelease{Version: "v1.2.3"}},
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got := checkForUpdateBestEffort(context.Background(), tt.checker, tt.currentVersion)
			if got.LatestVersion != tt.wantLatest {
				t.Fatalf("latest = %q, want %q", got.LatestVersion, tt.wantLatest)
			}
			if got.Warning != tt.wantWarning {
				t.Fatalf("warning = %q, want %q", got.Warning, tt.wantWarning)
			}
		})
	}
}

func TestScaffoldWelcomeModelRecordsUpdateWarnings(t *testing.T) {
	tests := []struct {
		name        string
		app         app
		skip        bool
		envValue    *string
		wantLatest  string
		wantWarning string
	}{
		{
			name: "newer release",
			app: app{
				buildInfo:     BuildInfo{Version: "1.2.3"},
				updateChecker: stubReleaseChecker{release: latestRelease{Version: "v1.2.4"}},
			},
			wantLatest: "v1.2.4",
		},
		{
			name: "checker error warning",
			app: app{
				buildInfo:     BuildInfo{Version: "1.2.3"},
				updateChecker: stubReleaseChecker{err: errors.New("network blocked")},
			},
			wantWarning: "update check failed: network blocked",
		},
		{
			name: "non semantic current warning",
			app: app{
				buildInfo:     BuildInfo{Version: "dev"},
				updateChecker: stubReleaseChecker{release: latestRelease{Version: "v1.2.4"}},
			},
			wantWarning: `update check skipped: current version "dev" is not a release version`,
		},
		{
			name: "explicit flag skip has no warning",
			app: app{
				buildInfo:     BuildInfo{Version: "dev"},
				updateChecker: stubReleaseChecker{release: latestRelease{Version: "v1.2.4"}},
			},
			skip: true,
		},
		{
			name: "env skip has no warning",
			app: app{
				buildInfo:     BuildInfo{Version: "dev"},
				updateChecker: stubReleaseChecker{release: latestRelease{Version: "v1.2.4"}},
			},
			envValue: ptrString("1"),
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			if tt.envValue != nil {
				t.Setenv(noUpdateCheckEnvVar, *tt.envValue)
			} else {
				t.Setenv(noUpdateCheckEnvVar, "0")
			}
			got := tt.app.scaffoldWelcomeModel(tt.skip)
			if got.UpdateVersion != tt.wantLatest {
				t.Fatalf("UpdateVersion = %q, want %q", got.UpdateVersion, tt.wantLatest)
			}
			if got.UpdateWarning != tt.wantWarning {
				t.Fatalf("UpdateWarning = %q, want %q", got.UpdateWarning, tt.wantWarning)
			}
		})
	}
}

func TestGitHubReleaseCheckerLatestRelease(t *testing.T) {
	var gotAccept string
	var gotUserAgent string
	server := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		gotAccept = r.Header.Get("Accept")
		gotUserAgent = r.Header.Get("User-Agent")
		w.Header().Set("Content-Type", "application/json")
		_, _ = w.Write([]byte(`{"tag_name":"v1.2.4","html_url":"https://github.com/harumiWeb/xlflow/releases/tag/v1.2.4"}`))
	}))
	defer server.Close()

	checker := newGitHubReleaseChecker(server.Client())
	checker.url = server.URL

	got, err := checker.LatestRelease(context.Background())
	if err != nil {
		t.Fatalf("LatestRelease() error = %v", err)
	}
	if got.Version != "v1.2.4" {
		t.Fatalf("LatestRelease() version = %q, want v1.2.4", got.Version)
	}
	if got.ReleaseURL != "https://github.com/harumiWeb/xlflow/releases/tag/v1.2.4" {
		t.Fatalf("LatestRelease() release URL = %q", got.ReleaseURL)
	}
	if gotAccept != "application/vnd.github+json" {
		t.Fatalf("Accept header = %q, want application/vnd.github+json", gotAccept)
	}
	if gotUserAgent == "" {
		t.Fatal("User-Agent header should not be empty")
	}
}

func TestUpdateCheckCommandWritesJSON(t *testing.T) {
	tests := []struct {
		name           string
		currentVersion string
		release        latestRelease
		wantAvailable  bool
		wantLatest     string
		wantReleaseURL string
		wantExit       int
	}{
		{
			name:           "newer release available",
			currentVersion: "1.2.3",
			release: latestRelease{
				Version:    "v1.2.4",
				ReleaseURL: "https://example.com/v1.2.4",
			},
			wantAvailable:  true,
			wantLatest:     "v1.2.4",
			wantReleaseURL: "https://example.com/v1.2.4",
			wantExit:       0,
		},
		{
			name:           "up to date",
			currentVersion: "1.2.3",
			release:        latestRelease{Version: "v1.2.3"},
			wantAvailable:  false,
			wantLatest:     "",
			wantExit:       0,
		},
		{
			name:           "dev build skips comparison",
			currentVersion: "dev",
			release:        latestRelease{Version: "v1.2.4"},
			wantAvailable:  false,
			wantLatest:     "",
			wantExit:       0,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			var stdout bytes.Buffer
			a := &app{
				stdout:        &stdout,
				stderr:        &bytes.Buffer{},
				buildInfo:     BuildInfo{Version: tt.currentVersion},
				updateChecker: stubReleaseChecker{release: tt.release},
			}
			root := a.rootCommand()
			root.SetArgs([]string{"--json", "update", "check"})

			err := root.Execute()
			if gotExit := output.ExitCode(err); gotExit != tt.wantExit {
				t.Fatalf("exit = %d, want %d, err = %v", gotExit, tt.wantExit, err)
			}

			var got struct {
				Status  string             `json:"status"`
				Command string             `json:"command"`
				Update  updateCheckPayload `json:"update"`
			}
			if err := json.Unmarshal(stdout.Bytes(), &got); err != nil {
				t.Fatal(err)
			}
			if got.Status != output.StatusOK {
				t.Fatalf("status = %q, want %q", got.Status, output.StatusOK)
			}
			if got.Command != "update check" {
				t.Fatalf("command = %q, want update check", got.Command)
			}
			if got.Update.CurrentVersion != tt.currentVersion {
				t.Fatalf("current = %q, want %q", got.Update.CurrentVersion, tt.currentVersion)
			}
			if got.Update.UpdateAvailable != tt.wantAvailable {
				t.Fatalf("available = %t, want %t", got.Update.UpdateAvailable, tt.wantAvailable)
			}
			if got.Update.LatestVersion != tt.wantLatest {
				t.Fatalf("latest = %q, want %q", got.Update.LatestVersion, tt.wantLatest)
			}
			if got.Update.ReleaseURL != tt.wantReleaseURL {
				t.Fatalf("release URL = %q, want %q", got.Update.ReleaseURL, tt.wantReleaseURL)
			}
		})
	}
}

func TestUpdateCheckCommandFailsWhenReleaseLookupFails(t *testing.T) {
	var stdout bytes.Buffer
	a := &app{
		stdout:        &stdout,
		stderr:        &bytes.Buffer{},
		buildInfo:     BuildInfo{Version: "1.2.3"},
		updateChecker: stubReleaseChecker{err: errors.New("network down")},
	}
	root := a.rootCommand()
	root.SetArgs([]string{"--json", "update", "check"})

	err := root.Execute()
	if gotExit := output.ExitCode(err); gotExit != output.ExitEnvironment {
		t.Fatalf("exit = %d, want %d, err = %v", gotExit, output.ExitEnvironment, err)
	}

	var got struct {
		Status string             `json:"status"`
		Error  *output.Error      `json:"error"`
		Update updateCheckPayload `json:"update"`
	}
	if err := json.Unmarshal(stdout.Bytes(), &got); err != nil {
		t.Fatal(err)
	}
	if got.Status != output.StatusFailed {
		t.Fatalf("status = %q, want %q", got.Status, output.StatusFailed)
	}
	if got.Error == nil || got.Error.Code != "update_check_failed" {
		t.Fatalf("unexpected error: %#v", got.Error)
	}
	if got.Update.CurrentVersion != "1.2.3" {
		t.Fatalf("current = %q, want 1.2.3", got.Update.CurrentVersion)
	}
}

func TestGitHubReleaseCheckerRejectsUnexpectedStatus(t *testing.T) {
	server := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		http.Error(w, "boom", http.StatusBadGateway)
	}))
	defer server.Close()

	checker := newGitHubReleaseChecker(server.Client())
	checker.url = server.URL

	if _, err := checker.LatestRelease(context.Background()); err == nil {
		t.Fatal("LatestRelease() error = nil, want non-nil")
	}
}

func TestShouldSkipScaffoldUpdateCheck(t *testing.T) {
	t.Run("unset", func(t *testing.T) {
		prev, had := os.LookupEnv(noUpdateCheckEnvVar)
		if err := os.Unsetenv(noUpdateCheckEnvVar); err != nil {
			t.Fatalf("os.Unsetenv(%q) error = %v", noUpdateCheckEnvVar, err)
		}
		t.Cleanup(func() {
			var err error
			if had {
				err = os.Setenv(noUpdateCheckEnvVar, prev)
			} else {
				err = os.Unsetenv(noUpdateCheckEnvVar)
			}
			if err != nil {
				t.Fatalf("restore %q error = %v", noUpdateCheckEnvVar, err)
			}
		})
		if shouldSkipScaffoldUpdateCheck() {
			t.Fatal("shouldSkipScaffoldUpdateCheck() = true, want false")
		}
	})

	t.Run("empty", func(t *testing.T) {
		t.Setenv(noUpdateCheckEnvVar, "")
		if shouldSkipScaffoldUpdateCheck() {
			t.Fatal("shouldSkipScaffoldUpdateCheck() = true, want false")
		}
	})

	for _, raw := range []string{"1", "true", "TRUE", "yes", "on"} {
		t.Run(raw, func(t *testing.T) {
			t.Setenv(noUpdateCheckEnvVar, raw)
			if !shouldSkipScaffoldUpdateCheck() {
				t.Fatalf("shouldSkipScaffoldUpdateCheck() = false for %q, want true", raw)
			}
		})
	}

	for _, raw := range []string{"0", "false", "FALSE", "no", "off", "   "} {
		t.Run(raw, func(t *testing.T) {
			t.Setenv(noUpdateCheckEnvVar, raw)
			if shouldSkipScaffoldUpdateCheck() {
				t.Fatalf("shouldSkipScaffoldUpdateCheck() = true for %q, want false", raw)
			}
		})
	}
}

func ptrString(value string) *string {
	return &value
}
