#!/usr/bin/env bash
# Cross-platform dependency installer for /єдрср skill
# Works on macOS, Linux (Debian/Ubuntu/Fedora/Arch), and Windows (Git Bash/WSL)
set -e

RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m'

ok()   { echo -e "${GREEN}✓${NC} $1"; }
warn() { echo -e "${YELLOW}⚠${NC} $1"; }
fail() { echo -e "${RED}✗${NC} $1"; }

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"

# --- Detect OS ---
detect_os() {
  case "$(uname -s 2>/dev/null || echo Windows)" in
    Darwin*)  echo "macos" ;;
    Linux*)
      if grep -qi microsoft /proc/version 2>/dev/null; then
        echo "wsl"
      else
        echo "linux"
      fi
      ;;
    MINGW*|MSYS*|CYGWIN*) echo "windows" ;;
    *)        echo "unknown" ;;
  esac
}

detect_linux_pkg() {
  if command -v apt-get &>/dev/null; then echo "apt"
  elif command -v dnf &>/dev/null; then echo "dnf"
  elif command -v yum &>/dev/null; then echo "yum"
  elif command -v pacman &>/dev/null; then echo "pacman"
  elif command -v apk &>/dev/null; then echo "apk"
  else echo "none"
  fi
}

OS=$(detect_os)
echo "OS: $OS"

# --- 1. Python 3 ---
if command -v python3 &>/dev/null; then
  ok "Python3: $(python3 --version 2>&1)"
elif command -v python &>/dev/null && python --version 2>&1 | grep -q "3\."; then
  ok "Python: $(python --version 2>&1)"
  alias python3=python
else
  warn "Python3 not found — installing..."
  case $OS in
    macos)
      if command -v brew &>/dev/null; then
        brew install python3
      else
        warn "brew not found — installing Homebrew first..."
        /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
        brew install python3
      fi
      ;;
    linux|wsl)
      PKG=$(detect_linux_pkg)
      case $PKG in
        apt) sudo apt-get update -qq && sudo apt-get install -y -qq python3 python3-pip ;;
        dnf) sudo dnf install -y python3 python3-pip ;;
        yum) sudo yum install -y python3 python3-pip ;;
        pacman) sudo pacman -Sy --noconfirm python python-pip ;;
        apk) sudo apk add python3 py3-pip ;;
        *) fail "No package manager found. Install Python3 manually: https://python.org" ;;
      esac
      ;;
    windows)
      if command -v winget &>/dev/null; then
        winget install Python.Python.3 --accept-package-agreements --accept-source-agreements
      elif command -v choco &>/dev/null; then
        choco install python3 -y
      elif command -v scoop &>/dev/null; then
        scoop install python
      else
        fail "Install Python3 manually: https://python.org/downloads/"
      fi
      ;;
  esac
fi

# --- 2. pip ---
if command -v pip3 &>/dev/null; then
  PIP=pip3
elif command -v pip &>/dev/null; then
  PIP=pip
else
  warn "pip not found — installing..."
  python3 -m ensurepip --upgrade 2>/dev/null || {
    curl -sS https://bootstrap.pypa.io/get-pip.py | python3
  }
  PIP=pip3
fi

# --- 3. Node.js + npm ---
if command -v node &>/dev/null; then
  ok "Node.js: $(node -v)"
else
  warn "Node.js not found — installing..."
  case $OS in
    macos)
      command -v brew &>/dev/null && brew install node || fail "Install Node.js: https://nodejs.org"
      ;;
    linux|wsl)
      PKG=$(detect_linux_pkg)
      case $PKG in
        apt) sudo apt-get update -qq && sudo apt-get install -y -qq nodejs npm ;;
        dnf) sudo dnf install -y nodejs npm ;;
        yum) sudo yum install -y nodejs npm ;;
        pacman) sudo pacman -Sy --noconfirm nodejs npm ;;
        apk) sudo apk add nodejs npm ;;
        *) fail "Install Node.js manually: https://nodejs.org" ;;
      esac
      ;;
    windows)
      if command -v winget &>/dev/null; then
        winget install OpenJS.NodeJS.LTS --accept-package-agreements --accept-source-agreements
      elif command -v choco &>/dev/null; then
        choco install nodejs-lts -y
      elif command -v scoop &>/dev/null; then
        scoop install nodejs-lts
      else
        fail "Install Node.js manually: https://nodejs.org"
      fi
      ;;
  esac
fi

# --- 4. pdftotext (poppler) ---
if command -v pdftotext &>/dev/null; then
  ok "pdftotext: installed"
else
  warn "pdftotext not found — installing poppler..."
  case $OS in
    macos)
      command -v brew &>/dev/null && brew install poppler || warn "Install poppler manually: brew install poppler"
      ;;
    linux|wsl)
      PKG=$(detect_linux_pkg)
      case $PKG in
        apt) sudo apt-get install -y -qq poppler-utils ;;
        dnf) sudo dnf install -y poppler-utils ;;
        yum) sudo yum install -y poppler-utils ;;
        pacman) sudo pacman -Sy --noconfirm poppler ;;
        apk) sudo apk add poppler-utils ;;
        *) warn "Install poppler-utils manually" ;;
      esac
      ;;
    windows)
      if command -v choco &>/dev/null; then
        choco install poppler -y
      elif command -v scoop &>/dev/null; then
        scoop install poppler
      else
        warn "Install poppler for Windows: https://github.com/oschwartz10612/poppler-windows/releases"
      fi
      ;;
  esac
fi

# --- 5. Python packages ---
echo ""
echo "Python packages:"

python3 -c "import docx" 2>/dev/null && ok "python-docx" || {
  warn "Installing python-docx..."
  $PIP install python-docx -q
  ok "python-docx installed"
}

python3 -c "import playwright" 2>/dev/null && ok "playwright" || {
  warn "Installing Playwright..."
  $PIP install playwright -q
  python3 -m playwright install chromium 2>/dev/null && ok "playwright + chromium installed" || warn "Playwright installed but chromium failed — run: python3 -m playwright install chromium"
}

# --- 6. npm docx package ---
echo ""
echo "npm packages:"
node -e "require('docx')" 2>/dev/null && ok "docx (npm)" || {
  warn "Installing docx npm package..."
  cd "$SCRIPT_DIR" && npm install -q
  ok "docx installed"
}

echo ""
echo -e "${GREEN}All dependencies checked.${NC}"
