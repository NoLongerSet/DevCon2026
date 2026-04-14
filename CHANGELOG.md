# Changelog

All notable changes to GBLauncher will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/),
and this project adheres to [Semantic Versioning](https://semver.org/).

## [1.0.0] - 2026-04-13

### Added

- File-based deployment (`--source` with local or network paths)
- Dropbox share link support (`--dropbox-share-link`) for direct downloads without a web server
- Local publish workflow (`--publish-local`) for file share and Dropbox deployments
- Version sidecar format (`.version` JSON files with version, SHA-256 hash, and download URL)
- SHA-256 integrity verification on every download
- Local caching to `%LOCALAPPDATA%\GBLauncher\cache\`
- Automatic Microsoft Access launching after download
- Trusted location registration for cached Access files
