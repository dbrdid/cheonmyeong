# 천명 프로그램 (WorshipSlide Pro)

Electron-based desktop app for worship slide presentations.  
예배 슬라이드 프레젠테이션 앱 — Electron 기반 데스크톱 애플리케이션.

> 일반 사용자 설치·사용법은 **[사용자_매뉴얼.md](사용자_매뉴얼.md)** 를 참고하세요.

---

## Supported Platforms

| OS | Status | PPTX conversion priority |
|----|--------|--------------------------|
| macOS (Apple Silicon / Intel) | ✅ | PowerPoint (AppleScript) → LibreOffice → Python+CoreGraphics |
| Windows 10 / 11 (x64) | ✅ | PowerPoint (PowerShell COM) → LibreOffice → Windows.Data.Pdf |
| Windows (ARM64) | ⚠️ experimental | Same as above |

---

## Development

**Prerequisites:** Node.js 20+, npm

```bash
git clone https://github.com/dbrdid/cheonmyeong.git
cd cheonmyeong
npm install
npm start
```

Core stack: Electron 31 · electron-builder 24 · electron-updater 6 · @google/genai · adm-zip

---

## Build

### Windows

```bash
# x64 portable single executable
npm run build-win

# x64 NSIS installer wizard (Korean UI)
npm run build-win-installer

# ARM64 portable (ARM Windows)
npm run build-win-arm
```

Output: `dist/천명프로그램-{version}-portable.exe` or `dist/천명프로그램-{version}-Setup.exe`

> **Cross-compile note:** When building on Apple Silicon Mac, `--x64` is already specified in the npm scripts. Omitting it would produce an ARM64 binary that won't run on standard Windows PCs.

### macOS

```bash
npm run build-mac
```

Output: `dist/천명프로그램-{version}.dmg` (universal: arm64 + x64)

---

## Release (GitHub Actions)

Pushing a `v*` tag triggers automatic Mac + Windows builds and uploads artifacts to GitHub Releases.

```bash
git tag v1.0.1
git push origin v1.0.1
```

Manual trigger: GitHub → Actions tab → **Release** → **Run workflow**.

### First-time CI setup checklist

- [ ] `package.json` → `build.publish.owner` set to `dbrdid`
- [ ] `package.json` → `build.publish.repo` set to `cheonmyeong`
- [ ] GitHub repo → Settings → Actions → General → Workflow permissions → **Read and write permissions** ✅
- [ ] Verify `GH_TOKEN` is available (automatically provided by `GITHUB_TOKEN` secret)

---

## Auto-update

`electron-updater` is integrated. In packaged builds the app checks GitHub Releases on startup and downloads updates in the background. A restart prompt appears in the UI when an update is ready.

- `autoDownload: true` — downloads silently
- `autoInstallOnAppQuit: true` — installs on next quit
- Disabled automatically in dev mode (`app.isPackaged === false`)

---

## Project Structure

```
.
├── main.js              # Electron main process — OS branching, IPC handlers, updater
├── preload.js           # Renderer ↔ main IPC bridge (contextBridge)
├── index.html           # Editor UI (renderer)
├── screen.html          # Presentation output window (renderer)
├── package.json         # electron-builder config + dependencies
├── icon.ico             # Windows icon
├── icon.icns            # macOS icon
├── icon.png             # Fallback icon
├── 사용자_매뉴얼.md      # Korean end-user manual
├── WINDOWS_TEST_CHECKLIST.md   # QA checklist for Windows builds
└── .github/
    └── workflows/
        └── release.yml  # Tag-triggered CI/CD
```

---

## PPTX Conversion Engines (Windows)

| Priority | Engine | Requirement | Quality |
|----------|--------|-------------|---------|
| 1 | Microsoft PowerPoint (COM) | PowerPoint installed | ⭐⭐⭐ Best |
| 2 | LibreOffice headless | LibreOffice installed | ⭐⭐ Good |
| 3 | Windows.Data.Pdf (fallback) | Windows 10/11 built-in | ⭐ Limited |

---

## License

MIT
