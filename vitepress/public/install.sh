#!/bin/sh
set -eu

action="install"
install_dir="${XLFLOW_INSTALL_DIR:-$HOME/.local/bin}"
owner="${XLFLOW_OWNER:-harumiWeb}"
repo="${XLFLOW_REPO:-xlflow}"
modify_path="1"

info() {
  printf '[xlflow] %s\n' "$1"
}

fail() {
  printf '[xlflow] ERROR: %s\n' "$1" >&2
  exit 1
}

usage() {
  cat <<'EOF'
Usage: install.sh [options]

Install the xlflow Linux frontend binary for WSL development.

Options:
  --install-dir DIR   Install directory (default: $HOME/.local/bin)
  --uninstall         Remove xlflow from the install directory
  --no-modify-path    Do not add the install directory to a shell profile
  --owner OWNER       GitHub owner (default: harumiWeb)
  --repo REPO         GitHub repo (default: xlflow)
  -h, --help          Show this help

Environment:
  XLFLOW_INSTALL_DIR  Install directory override
  XLFLOW_OWNER        GitHub owner override
  XLFLOW_REPO         GitHub repo override
EOF
}

while [ "$#" -gt 0 ]; do
  case "$1" in
    --install-dir)
      [ "$#" -ge 2 ] || fail "--install-dir requires a value"
      install_dir="$2"
      shift 2
      ;;
    --uninstall)
      action="uninstall"
      shift
      ;;
    --no-modify-path)
      modify_path="0"
      shift
      ;;
    --owner)
      [ "$#" -ge 2 ] || fail "--owner requires a value"
      owner="$2"
      shift 2
      ;;
    --repo)
      [ "$#" -ge 2 ] || fail "--repo requires a value"
      repo="$2"
      shift 2
      ;;
    -h | --help)
      usage
      exit 0
      ;;
    *)
      fail "unknown option: $1"
      ;;
  esac
done

need_cmd() {
  command -v "$1" >/dev/null 2>&1 || fail "required command not found: $1"
}

download_to_stdout() {
  url="$1"
  if command -v curl >/dev/null 2>&1; then
    curl -fsSL -H "Accept: application/vnd.github+json" -H "User-Agent: xlflow-install-script" "$url"
    return
  fi
  if command -v wget >/dev/null 2>&1; then
    wget -qO- --header="Accept: application/vnd.github+json" --header="User-Agent: xlflow-install-script" "$url"
    return
  fi
  fail "curl or wget is required"
}

download_to_file() {
  url="$1"
  path="$2"
  if command -v curl >/dev/null 2>&1; then
    curl -fL -H "User-Agent: xlflow-install-script" -o "$path" "$url"
    return
  fi
  if command -v wget >/dev/null 2>&1; then
    wget -qO "$path" --header="User-Agent: xlflow-install-script" "$url"
    return
  fi
  fail "curl or wget is required"
}

select_linux_asset_url() {
  release_json="$1"
  preferred=$(printf '%s\n' "$release_json" |
    sed -n 's/.*"browser_download_url":[[:space:]]*"\([^"]*xlflow_linux_x86_64\.tar\.gz\)".*/\1/p' |
    head -n 1)
  if [ -n "$preferred" ]; then
    printf '%s\n' "$preferred"
    return
  fi

  fallback=$(printf '%s\n' "$release_json" |
    sed -n 's/.*"browser_download_url":[[:space:]]*"\([^"]*\)".*/\1/p' |
    grep -Ei 'linux.*(x86_64|amd64).*\.tar\.gz$' |
    head -n 1 || true)
  if [ -n "$fallback" ]; then
    printf '%s\n' "$fallback"
    return
  fi

  fail "could not find a Linux x64 tar.gz asset in the latest release"
}

select_checksums_url() {
  release_json="$1"
  url=$(printf '%s\n' "$release_json" |
    sed -n 's/.*"browser_download_url":[[:space:]]*"\([^"]*checksums\.txt\)".*/\1/p' |
    head -n 1)
  if [ -n "$url" ]; then
    printf '%s\n' "$url"
    return
  fi

  fail "could not find checksums.txt in the latest release"
}

verify_checksum() {
  archive_path="$1"
  archive_name="$2"
  checksums_url="$3"

  checksums_path=$(mktemp)
  download_to_file "$checksums_url" "$checksums_path"

  expected=$(awk -v name="$archive_name" '$2 == name || $2 == "*" name { print $1; exit }' "$checksums_path")
  rm -f "$checksums_path"
  [ -n "$expected" ] || fail "could not find checksum for $archive_name"

  if command -v sha256sum >/dev/null 2>&1; then
    actual=$(sha256sum "$archive_path" | awk '{print $1}')
  elif command -v shasum >/dev/null 2>&1; then
    actual=$(shasum -a 256 "$archive_path" | awk '{print $1}')
  else
    fail "sha256sum or shasum is required for checksum verification"
  fi

  [ "$actual" = "$expected" ] || fail "checksum verification failed for $archive_name"
  info "Verified checksum for $archive_name"
}

profile_file() {
  shell_name=$(basename "${SHELL:-sh}")
  case "$shell_name" in
    zsh)
      printf '%s\n' "$HOME/.zshrc"
      ;;
    bash)
      printf '%s\n' "$HOME/.bashrc"
      ;;
    *)
      printf '%s\n' "$HOME/.profile"
      ;;
  esac
}

path_entry_expr() {
  case "$install_dir" in
    "$HOME"/*)
      suffix=${install_dir#"$HOME"/}
      printf '$HOME/%s\n' "$suffix"
      ;;
    *)
      printf '%s\n' "$install_dir"
      ;;
  esac
}

ensure_path_profile() {
  [ "$modify_path" = "1" ] || return 0

  case ":$PATH:" in
    *":$install_dir:"*) return 0 ;;
  esac

  profile=$(profile_file)
  entry=$(path_entry_expr)

  touch "$profile"
  if grep -q '>>> xlflow installer >>>' "$profile"; then
    info "PATH block already exists in $profile"
    return 0
  fi

  {
    printf '\n# >>> xlflow installer >>>\n'
    printf 'export PATH="%s:$PATH"\n' "$entry"
    printf '# <<< xlflow installer <<<\n'
  } >>"$profile"

  info "Added $install_dir to PATH in $profile"
  info "Open a new shell or run: export PATH=\"$install_dir:\$PATH\""
}

remove_path_profile() {
  profile=$(profile_file)
  [ -f "$profile" ] || return 0
  tmp_file=$(mktemp)
  sed '/# >>> xlflow installer >>>/,/# <<< xlflow installer <<</d' "$profile" >"$tmp_file"
  if ! cmp -s "$profile" "$tmp_file"; then
    mv "$tmp_file" "$profile"
    info "Removed xlflow PATH block from $profile"
  else
    rm -f "$tmp_file"
  fi
}

install_xlflow() {
  need_cmd tar
  need_cmd sed
  need_cmd grep
  need_cmd awk
  need_cmd find
  need_cmd mktemp

  api_url="https://api.github.com/repos/$owner/$repo/releases/latest"
  release_url="https://github.com/$owner/$repo/releases/latest"
  tmp_root=$(mktemp -d)

  cleanup() {
    rm -rf "$tmp_root"
  }
  trap cleanup EXIT INT TERM

  info "Fetching latest release metadata from $api_url"
  release_json=$(download_to_stdout "$api_url")
  asset_url=$(select_linux_asset_url "$release_json")
  checksums_url=$(select_checksums_url "$release_json")
  archive_path="$tmp_root/xlflow-linux.tar.gz"
  extract_dir="$tmp_root/extract"
  archive_name=$(basename "$asset_url")

  info "Downloading $archive_name"
  download_to_file "$asset_url" "$archive_path"
  verify_checksum "$archive_path" "$archive_name" "$checksums_url"

  info "Extracting archive"
  mkdir -p "$extract_dir"
  tar -xzf "$archive_path" -C "$extract_dir"

  xlflow_bin=$(find "$extract_dir" -type f -name xlflow -perm -u+x | head -n 1)
  if [ -z "$xlflow_bin" ]; then
    xlflow_bin=$(find "$extract_dir" -type f -name xlflow | head -n 1)
  fi
  [ -n "$xlflow_bin" ] || fail "downloaded archive did not contain xlflow"

  info "Installing to $install_dir"
  mkdir -p "$install_dir"
  cp "$xlflow_bin" "$install_dir/xlflow"
  chmod 0755 "$install_dir/xlflow"

  ensure_path_profile

  info "Verifying installation"
  "$install_dir/xlflow" version

  info "Install complete."
  info "Next step: xlflow doctor --json"
  info "Release source: $release_url"
}

uninstall_xlflow() {
  if [ -f "$install_dir/xlflow" ]; then
    rm -f "$install_dir/xlflow"
    info "Removed $install_dir/xlflow"
  else
    info "No xlflow binary found at $install_dir/xlflow"
  fi
  remove_path_profile
  info "Uninstall complete."
}

case "$action" in
  install)
    install_xlflow
    ;;
  uninstall)
    uninstall_xlflow
    ;;
  *)
    fail "unsupported action: $action"
    ;;
esac
