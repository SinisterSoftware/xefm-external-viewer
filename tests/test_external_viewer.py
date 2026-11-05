from external_viewer.logic import _find_viewer

def test_viewer_matches_case_insensitively():
    assert _find_viewer("REPORT.PDF")["pattern"] == ["*.pdf", "*.epub"]
    assert _find_viewer("unknown.xyz") is None
