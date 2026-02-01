# Troubleshooting

- Ribbon tab missing: ensure macros are enabled and the add-in is loaded.
- Clipboard not updated: check macOS permission prompts.
- Workbook path missing: the workbook may be unsaved; save and retry.

## Clipboard access

If you see a clipboard warning, Excel may be blocking AppleScript clipboard
access. Try:

- macOS `System Settings` -> `Privacy & Security` -> `Automation` and allow
  Excel to control `System Events`.
- Restart Excel after changing permissions.
