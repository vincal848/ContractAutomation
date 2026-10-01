"""Static checks on src/ReportAutomation.bas text -- VBA itself cannot run in CI."""
import os
import re
import sys
from collections import Counter

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

BAS_PATH = os.path.join(ROOT, "src", "ReportAutomation.bas")

SUB_FUNCTION_RE = re.compile(
    r"^\s*(?:Public\s+|Private\s+)?(?:Sub|Function)\s+(\w+)",
    re.IGNORECASE | re.MULTILINE,
)


def _read_bas():
    with open(BAS_PATH, "r", encoding="utf-8") as f:
        return f.read()


def test_bas_file_exists():
    assert os.path.isfile(BAS_PATH)


def test_no_continue_do_which_is_not_valid_vba():
    text = _read_bas()
    assert "Continue Do" not in text


def test_no_duplicate_sub_or_function_names():
    text = _read_bas()
    names = [name.lower() for name in SUB_FUNCTION_RE.findall(text)]
    assert names, "expected at least one Sub or Function definition"
    duplicates = [name for name, count in Counter(names).items() if count > 1]
    assert duplicates == []


def test_option_explicit_is_present():
    text = _read_bas()
    assert "Option Explicit" in text


def test_no_hardcoded_c_drive_paths():
    text = _read_bas()
    assert "C:\\" not in text
