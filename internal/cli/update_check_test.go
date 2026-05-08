package cli

import (
	"context"
	"net/http"
	"net/http/httptest"
	"os"
	"testing"
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
				release: latestRelease{Version: tt.latestVersion},
			}, tt.currentVersion)
			if err != nil {
				t.Fatalf("checkForUpdate() error = %v", err)
			}
			if got.LatestVersion != tt.want {
				t.Fatalf("checkForUpdate() latest = %q, want %q", got.LatestVersion, tt.want)
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
		_, _ = w.Write([]byte(`{"tag_name":"v1.2.4"}`))
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
	if gotAccept != "application/vnd.github+json" {
		t.Fatalf("Accept header = %q, want application/vnd.github+json", gotAccept)
	}
	if gotUserAgent == "" {
		t.Fatal("User-Agent header should not be empty")
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
