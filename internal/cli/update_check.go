package cli

import (
	"context"
	"encoding/json"
	"fmt"
	"net/http"
	"os"
	"strconv"
	"strings"
	"time"
)

const (
	gitHubLatestReleaseURL = "https://api.github.com/repos/harumiWeb/xlflow/releases/latest"
	updateCheckTimeout     = 3 * time.Second
	noUpdateCheckEnvVar    = "XLFLOW_NO_UPDATE_CHECK"
)

type releaseChecker interface {
	LatestRelease(ctx context.Context) (latestRelease, error)
}

type latestRelease struct {
	Version    string
	ReleaseURL string
}

type gitHubReleaseChecker struct {
	client *http.Client
	url    string
}

type scaffoldUpdateInfo struct {
	LatestVersion string
	ReleaseURL    string
}

type semanticVersion struct {
	major      int
	minor      int
	patch      int
	prerelease string
}

func newGitHubReleaseChecker(client *http.Client) gitHubReleaseChecker {
	if client == nil {
		client = &http.Client{Timeout: updateCheckTimeout}
	}
	return gitHubReleaseChecker{
		client: client,
		url:    gitHubLatestReleaseURL,
	}
}

func (a *app) scaffoldWelcomeModel(skipUpdateCheck bool) scaffoldWelcomeModel {
	model := scaffoldWelcomeModel{
		Version: a.buildInfo.Version,
	}
	if skipUpdateCheck || shouldSkipScaffoldUpdateCheck() || a.updateChecker == nil {
		return model
	}
	ctx, cancel := context.WithTimeout(context.Background(), updateCheckTimeout)
	defer cancel()
	update, err := checkForUpdate(ctx, a.updateChecker, model.Version)
	if err != nil {
		return model
	}
	model.UpdateVersion = update.LatestVersion
	return model
}

func shouldSkipScaffoldUpdateCheck() bool {
	value, ok := os.LookupEnv(noUpdateCheckEnvVar)
	if !ok {
		return false
	}
	value = strings.TrimSpace(strings.ToLower(value))
	switch value {
	case "", "0", "false", "no", "off":
		return false
	default:
		return true
	}
}

func checkForUpdate(ctx context.Context, checker releaseChecker, currentVersion string) (scaffoldUpdateInfo, error) {
	current, ok := parseSemanticVersion(currentVersion)
	if !ok {
		return scaffoldUpdateInfo{}, nil
	}
	release, err := checker.LatestRelease(ctx)
	if err != nil {
		return scaffoldUpdateInfo{}, err
	}
	latest, ok := parseSemanticVersion(release.Version)
	if !ok {
		return scaffoldUpdateInfo{}, nil
	}
	if latest.compare(current) <= 0 {
		return scaffoldUpdateInfo{}, nil
	}
	return scaffoldUpdateInfo{
		LatestVersion: strings.TrimSpace(release.Version),
		ReleaseURL:    strings.TrimSpace(release.ReleaseURL),
	}, nil
}

func (c gitHubReleaseChecker) LatestRelease(ctx context.Context) (latestRelease, error) {
	req, err := http.NewRequestWithContext(ctx, http.MethodGet, c.url, nil)
	if err != nil {
		return latestRelease{}, err
	}
	req.Header.Set("Accept", "application/vnd.github+json")
	req.Header.Set("User-Agent", "xlflow update-check")

	resp, err := c.client.Do(req)
	if err != nil {
		return latestRelease{}, err
	}
	defer func() {
		_ = resp.Body.Close()
	}()

	if resp.StatusCode != http.StatusOK {
		return latestRelease{}, fmt.Errorf("github releases request failed: %s", resp.Status)
	}

	var payload struct {
		TagName string `json:"tag_name"`
		HTMLURL string `json:"html_url"`
	}
	if err := json.NewDecoder(resp.Body).Decode(&payload); err != nil {
		return latestRelease{}, err
	}
	if strings.TrimSpace(payload.TagName) == "" {
		return latestRelease{}, fmt.Errorf("github release tag_name is empty")
	}
	return latestRelease{
		Version:    strings.TrimSpace(payload.TagName),
		ReleaseURL: strings.TrimSpace(payload.HTMLURL),
	}, nil
}

func parseSemanticVersion(raw string) (semanticVersion, bool) {
	value := strings.TrimSpace(raw)
	if value == "" {
		return semanticVersion{}, false
	}
	value = strings.TrimPrefix(value, "v")
	value = strings.TrimPrefix(value, "V")
	if idx := strings.Index(value, "+"); idx >= 0 {
		value = value[:idx]
	}
	version := semanticVersion{}
	if idx := strings.Index(value, "-"); idx >= 0 {
		version.prerelease = value[idx+1:]
		value = value[:idx]
	}
	parts := strings.Split(value, ".")
	if len(parts) != 3 {
		return semanticVersion{}, false
	}
	numbers := []*int{&version.major, &version.minor, &version.patch}
	for i, part := range parts {
		n, err := strconv.Atoi(part)
		if err != nil || n < 0 {
			return semanticVersion{}, false
		}
		*numbers[i] = n
	}
	return version, true
}

func (v semanticVersion) compare(other semanticVersion) int {
	left := []int{v.major, v.minor, v.patch}
	right := []int{other.major, other.minor, other.patch}
	for i := range left {
		switch {
		case left[i] < right[i]:
			return -1
		case left[i] > right[i]:
			return 1
		}
	}
	switch {
	case v.prerelease == other.prerelease:
		return 0
	case v.prerelease == "":
		return 1
	case other.prerelease == "":
		return -1
	default:
		return strings.Compare(v.prerelease, other.prerelease)
	}
}
