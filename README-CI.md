# Access2SQL automated releases

The workflow in `.github/workflows/release.yml` builds:

- macOS Apple Silicon (`arm64`) `.dmg`
- macOS Intel (`x86_64`) `.dmg`
- Windows x64 `Setup.exe`
- Linux x64 `.deb`

A release is published automatically when a tag such as `v1.4.11` is pushed.

Manual `workflow_dispatch` builds the artifacts but does not publish a GitHub Release.

## Important runtime dependency

Linux packages depend on `mdbtools`.

Windows still requires the Microsoft Access Database Engine / Access ODBC driver, because that driver is not redistributable as part of the PyInstaller application.

## Release

```bash
git add .github/workflows/release.yml installer/Access2SQL.iss
git commit -m "Add automated cross-platform releases"
git push

git tag v1.4.11
git push origin v1.4.11
```

The workflow uses the GitHub-hosted macOS ARM64 and Intel runners and the Windows/Linux x64 runners.
