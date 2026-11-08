# External Viewers

[日本語版](README.ja.md)

A Windows XeFM extension that routes selected files to external viewers. XeFM
1.13.0 or later is required. Install with:

```text
xefm ext install github:SinisterSoftware/xefm-external-viewer
```

Accept the offered dependency sync, or run `xefm deps sync` afterwards. The
extension requires the Python package `pywin32` for read-only Office automation;
it no longer depends on `extract-msg`, Mail Archive, or SQLite Archive.

The default rules use **SumatraPDF** (`sumatrapdf` on PATH) for PDF and EPUB,
**mpv** (`mpv` on PATH) for the listed audio/video formats, and Microsoft
Excel, Word, or PowerPoint for legacy `.xls`, `.doc`, and `.ppt` respectively.
Office must be installed and registered for COM automation; the Office
commands request read-only documents. Newer Office formats are not routed.
These applications are optional: they are not installed by this extension.
They are deliberately not listed under `requires.programs`, because XeFM
would otherwise disable the whole extension when any one is missing. Use
absolute executable paths in your configuration if an application is not on
PATH.

## Usage and fallback

Enter runs matching rules. The extension also contributes the action
`external-viewer.view-file`, which suggests `V`; an existing binding may
prevent the suggestion from being applied. To bind it explicitly, add the
following entry to the existing `KEY_BINDINGS` dictionary in `Config`:

```python
'external-viewer.view-file': ['V'],
```

Remove `V` from a conflicting binding if necessary. The action is also
available in the command palette. A successful launch claims the file; if
the program cannot be started, Office automation fails, a rule is invalid,
or a virtual file cannot be copied, the error is logged and the extension
declines the open. Enter then continues through XeFM's other hooks and normal
file handling if no archive action handles it; the action uses XeFM's normal
`view_file` action, which can itself honor `FILE_ASSOCIATIONS`. Whether XeFM
has a built-in viewer for the file depends on its format. Failures after an
external process starts cannot be detected or trigger fallback.

Mail and SQLite files are intentionally absent from the default rules. Their
extensions are optional. For both Enter and `V`, External Viewers first tries
any configured viewer, then invokes Mail Archive's `mail_archive.open` for
`.eml` and `.msg` or SQLite Archive's `sqlite_archive.open` for `.db`,
`.sqlite`, `.sqlite3`, and `.db3` when the action is available. The archive
action handles the file only if it changes the pane's location. Otherwise,
Enter falls through to XeFM's normal file handling and `V` uses the normal
viewer. XeFM 1.13.1 accepts only one extension manifest contribution for a
given hook key, so this delegation lets the archive actions work when External
Viewers owns the `file_open` hook.

Remote and archive files are copied to a uniquely named temporary directory
before an external program is launched. The copy is removed when XeFM exits;
it is not written back, and may disappear while an external viewer is still
open. Large or remote files can take time to copy.

## Configuration

Add a top-level class named `ExternalViewerConfig` to `~/.xefm/config.py`,
outside `class Config`. `VIEWERS` replaces the **whole** default list. Rules
are matched in order, case-insensitively; the first matching pattern wins.
Use `xefm ext settings external-viewer` to see the defaults. For example:

```python
class ExternalViewerConfig:
    VIEWERS = [
        {'pattern': ['*.pdf', '*.epub'],
         'command': [r'C:\Program Files\SumatraPDF\SumatraPDF.exe']},
        {'pattern': '*.xls', 'command': 'office:excel'},
        {'pattern': '*.doc', 'command': 'office:word'},
        {'pattern': '*.ppt', 'command': 'office:powerpoint'},
        {'pattern': ['*.ts', '*.mp4', '*.mpg', '*.vob', '*.mov',
                     '*.flac', '*.wv', '*.m4a'], 'command': ['mpv']},
    ]
```

`pattern` is a glob or a nonempty list of globs. `command` is a nonempty
argument list (the file path is appended as one argument), one of the three
`office:` values above, `None` for the Windows file association, or a callable
`fn(ctx, path)`. A callable returning `False` declines the open. Programs
launched via argv should be graphical/non-interactive; they do not receive
the XeFM terminal. Remove a rule to leave that format to XeFM or another
extension. `VIEWERS = []` disables all routing. No installed source files
need editing.

After changing configuration, use **Tools > Reload Configuration** or restart
XeFM. The extension uses XeFM's Preview event-hook and settings APIs, which may
change in a future release.

## Development

Run `xefm ext check .` and `python -m pytest tests`. Tests use fake contexts
and launchers and do not start XeFM or external applications.
