"""Defaults for ExternalViewerConfig in the user's config.py."""


class Settings:
    VIEWERS = [
        {"pattern": ["*.pdf", "*.epub"], "command": ["sumatrapdf"]},
        {"pattern": "*.xls", "command": "office:excel"},
        {"pattern": "*.doc", "command": "office:word"},
        {"pattern": "*.ppt", "command": "office:powerpoint"},
        {"pattern": ["*.ts", "*.mp4", "*.mpg", "*.vob", "*.mov",
                     "*.flac", "*.wv", "*.m4a"], "command": ["mpv"]},
    ]
