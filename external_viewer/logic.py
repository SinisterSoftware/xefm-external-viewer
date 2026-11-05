"""外部ビューアへの振り分け (~/.xefm/extensions/external_viewer.py)

V / Enter で開くファイルを、拡張子ごとに外部プログラムへ渡す。
zip や ssh:// の中のファイルは、一時フォルダへ書き出してから渡す。

  - 設定は EXTERNAL_VIEWERS だけ。command は次のどれか
      ['prog', ...]  プログラムを起動 (末尾にファイルパスが付く)
      callable       fn(ctx, path) を呼ぶ
      None           os.startfile (関連付けで開く)
  - config.py 側は ACTIONS / EVENT_HOOKS をマージするだけ
"""
import atexit
import fnmatch
import os
import shutil
import subprocess
import tempfile
from pathlib import Path



# ---- Office を読み取り専用で開く (COM) -------------------------------------------

def _office_readonly(prog_id, open_doc):
    def opener(ctx, path):
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        app = win32com.client.Dispatch(prog_id)
        app.Visible = True
        open_doc(app, str(path))
    return opener


open_excel_readonly = _office_readonly(
    "Excel.Application", lambda app, p: app.Workbooks.Open(p, 0, True))
open_word_readonly = _office_readonly(
    "Word.Application", lambda app, p: app.Documents.Open(p, False, True))
open_powerpoint_readonly = _office_readonly(
    "PowerPoint.Application", lambda app, p: app.Presentations.Open(p, True, False, True))


# ---- 設定: ここを編集する ---------------------------------------------------------

def _open_eml(ctx, path):
    from mail_archive import open_eml
    return open_eml(ctx, path)

def _open_msg(ctx, path):
    from mail_archive import open_msg
    return open_msg(ctx, path)

def _open_sqlite(ctx, path):
    from sqlite_archive import open_sqlite
    return open_sqlite(ctx, path)


EXTERNAL_VIEWERS = [
    {"pattern": ["*.pdf", "*.epub"],                "command": ["sumatrapdf"]},
    # 新形式 (xlsx/xlsm/docx/docm/pptx/pptm) は内蔵ビューア (doc-files)。旧形式だけ Office で開く。
    {"pattern": ["*.xls"],                          "command": open_excel_readonly},
    {"pattern": ["*.doc"],                          "command": open_word_readonly},
    {"pattern": ["*.ppt"],                          "command": open_powerpoint_readonly},
    {"pattern": ["*.ts", "*.mp4", "*.mpg",
                 "*.vob", "*.mov",
                 "*.flac", "*.wv", "*.m4a"],        "command": ["mpv"]},
    {"pattern": ["*.eml"],                          "command": _open_eml},
    {"pattern": ["*.msg"],                          "command": _open_msg},
    {"pattern": ["*.db", "*.sqlite", "*.sqlite3", "*.db3"], "command": _open_sqlite},
]


# ---- 仮想パスの実体化 -------------------------------------------------------------

_tmp_root = None


def _tmp_dir():
    """一時フォルダは最初に必要になった時に作る (起動フック不要・config 再読込に強い)。"""
    global _tmp_root
    if _tmp_root is None or not _tmp_root.exists():
        _tmp_root = Path(tempfile.mkdtemp(prefix="xefm_viewer_"))
        atexit.register(shutil.rmtree, _tmp_root, ignore_errors=True)
    return _tmp_root


def _materialize(path):
    if "://" not in str(path):
        return path                      # 本物のローカルファイル
    target = Path(tempfile.mkdtemp(dir=_tmp_dir())) / path.name   # 同名衝突を避ける
    target.write_bytes(path.read_bytes())
    return target


def _find_viewer(filename):
    name = filename.lower()
    for entry in EXTERNAL_VIEWERS:
        if any(fnmatch.fnmatch(name, pat.lower()) for pat in entry["pattern"]):
            return entry
    return None


# ---- XeFM 連携 --------------------------------------------------------------------

def open_with_external_viewer(ctx, event):
    """file_open フック。担当外なら False を返して XeFM に任せる。"""
    path = getattr(event, "path", event)     # フックのイベントでもパスそのものでも受ける
    entry = _find_viewer(path.name)
    if entry is None:
        return False
    try:
        target = _materialize(path)
        command = entry["command"]
        if command is None:
            os.startfile(str(target))
        elif callable(command):
            if command(ctx, target) is False:
                return False
        else:
            subprocess.Popen(list(command) + [str(target)])
    except Exception as e:
        ctx.message(f"外部ビューア起動失敗: {e}")
    return True


def view_file_external(ctx):
    """V: 対象外のファイルは従来の内蔵ビューアへ。"""
    e = ctx.pane.focused
    if e is not None and e.is_file and open_with_external_viewer(ctx, e.path):
        return
    ctx.invoke("view_file")


ACTIONS = {"view_file": {"func": view_file_external, "override": True}}
EVENT_HOOKS = {"file_open": [open_with_external_viewer]}



