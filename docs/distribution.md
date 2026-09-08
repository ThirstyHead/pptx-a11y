# Trusted Platform Distribution & Packaging Guide

This guide details the release packaging, code signing, and notarization procedures required for distributing `pptx-a11y` as a trusted native desktop application on macOS, Windows, and Linux.

---

## 1. macOS Distribution & Notarization

macOS Gatekeeper blocks unnotarized applications or marks them as damaged. To produce a seamless double-click installer DMG, developers must sign binaries with an Apple Developer ID Certificate and submit them to Apple's Notary Service.

### Prerequisites & Developer Account Requirements
- **Apple Developer Program Enrollment:** An active Apple Developer account ($99/year) at [developer.apple.com](https://developer.apple.com).
- **Certificate:** "Developer ID Application: Your Organization Name (TEAM_ID)" installed in your macOS Keychain.
- **App-Specific Password:** Generated via [appleid.apple.com](https://appleid.apple.com) for CLI notary submissions.
- **Xcode Command Line Tools:** `xcode-select --install` provides `codesign`, `notarytool`, and `stapler`.

### Environment Variables
Configure the following in your deployment environment or CI secrets:
```bash
export DEVELOPER_ID_APPLICATION="Developer ID Application: Your Name (XXXXXXXXXX)"
export APPLE_ID="developer@example.com"
export APPLE_TEAM_ID="XXXXXXXXXX"
export APPLE_APP_SPECIFIC_PASSWORD="xxxx-xxxx-xxxx-xxxx"
```

### Build & Notarize Pipeline
```bash
# 1. Build native arm64 / x86_64 bundle
./packaging/mac/build_mac.sh

# 2. Sign, package into DMG, submit to notarytool, and staple ticket
./packaging/mac/sign_and_notarize.sh
```

### Local Testing Without Developer Account
For internal development or local testing on Apple Silicon (M-series):
```bash
# Ad-hoc sign the bundle
codesign --deep --force --sign - dist/pptx-a11y.app

# Strip quarantine attribute if copied from another machine
xattr -cr dist/pptx-a11y.app
```

---

## 2. Windows Distribution & SmartScreen Trust

Windows Defender SmartScreen displays an untrusted warning ("Windows protected your PC") on unsigned or newly published executables.

### Prerequisites & Code Signing Options
- **EV (Extended Validation) Code Signing Certificate:** Provides immediate reputation trust in SmartScreen without download ramp-up periods (requires hardware token or cloud HSM).
- **Standard Code Signing Certificate:** Signs the `.exe` / `.msi`. SmartScreen trust builds progressively as users download without incident.
- **Microsoft Azure Trusted Signing (formerly Azure Code Signing):** Microsoft's fully managed PKI service for code signing in CI/CD pipelines without physical HSM dongles.

### Packaging with Inno Setup or WiX
1. Build the standalone executable:
   ```cmd
   packaging\windows\build_win.bat
   ```
2. Sign the binary using `signtool`:
   ```cmd
   signtool sign /f certificate.pfx /p Password /tr http://timestamp.digicert.com /td sha256 /fd sha256 dist\pptx-a11y\pptx-a11y.exe
   ```
3. Compile the setup installer (`.msi` or `.exe`) using Inno Setup or WiX Toolset, and sign the installer executable.

---

## 3. Linux Distribution (AppImage & DEB)

Linux users require self-contained packages with zero shared library conflicts across distributions (Ubuntu, Fedora, Arch).

### AppImage Build
1. Build the standalone binary:
   ```bash
   ./packaging/linux/build_linux.sh
   ```
2. Use `appimagetool` to bundle into a `.AppImage`:
   ```bash
   appimagetool dist/pptx-a11y pptx-a11y-x86_64.AppImage
   ```
3. Sign with GPG key:
   ```bash
   gpg --armor --detach-sign pptx-a11y-x86_64.AppImage
   ```

---

## 4. Package Manager Distribution (Developer / Headless Users)

For developer workstations, automation pipelines, and CI/CD:
```bash
# Install CLI only (minimal dependencies)
pip install pptx-a11y

# Install with GUI support
pip install "pptx-a11y[gui]"

# Or via pipx for isolated global CLI/GUI execution
pipx install "pptx-a11y[gui]"
```
