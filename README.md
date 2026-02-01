# Excel macOS Matlab Link

VBA ribbon add-in for Excel on macOS that generates MATLAB code snippets and
copies them to the clipboard.

## Motivation

Excel for macOS lacks the classic spreadsheet-to-MATLAB link since ActiveX /
OLE is limited to Microsoft Windows-only. As macOS adoption grows in engineering
and startups, that missing bridge becomes a real gap. This project restores
some of that convenience with range-based imports on Mac, within the platform’s
constraints.

## Install

See `docs/install.md`. For clipboard access issues, see
`docs/troubleshooting.md#clipboard-access`.

## Quick Start

1. Build the add-in: `docs/build.md`.
2. Install the add-in: `docs/install.md`.
3. Use the ribbon tab to copy a `readmatrix` snippet.

## Development

- Source VBA modules: `src/vba/`
- Ribbon XML/icons: `src/ribbon/`
- Build instructions: `docs/build.md`

## License

MIT (see `LICENSE`).
