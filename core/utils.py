"""
core/utils.py
─────────────
Pure helper functions with no Qt dependency.
"""
import os
import re
from threading import Lock

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string

# ── Encoding-aware file reader ───────────────────────────────────────────────
# Japanese automotive C/C++ files are typically Shift-JIS (CP932) or UTF-8.
# Strategy:
#   1. Honour BOM markers.
#   2. Strict-probe the common Japanese + Western encodings in order.
#      "latin-1" is intentionally EXCLUDED from the strict probe list because
#      it never raises UnicodeDecodeError (accepts every byte), which would
#      short-circuit the correct Japanese codec and produce mojibake.
#   3. Use chardet (if available) with a high-confidence gate (≥ 0.80) as a
#      smart fallback before resorting to cp932 with replacement chars.
_PROBE_ENCODINGS = ["cp932", "shift_jis", "euc_jp", "euc_jisx0213", "utf-8"]
def _is_valid_text(text: str) -> bool:
    # Reject mojibake
    if "�" in text:
        return False

    # Detect Japanese characters (Hiragana, Katakana, Kanji)
    japanese_chars = sum(
        1 for c in text
        if '\u3040' <= c <= '\u30ff' or '\u4e00' <= c <= '\u9faf'
    )
    if japanese_chars > 0:
        return True

    # Accept if mostly readable
    printable_ratio = sum(c.isprintable() for c in text) / max(len(text), 1)
    return printable_ratio > 0.95


def read_source_file(file_path: str) -> str:
    """Read a source file, auto-detecting encoding (UTF-8 / Shift-JIS / EUC-JP)."""
    try:
        raw = open(file_path, "rb").read()
    except Exception:
        return ""
    if not raw:
        return ""

    # Check for BOM markers first
    if raw.startswith(b"\xef\xbb\xbf"):
        return raw[3:].decode("utf-8", errors="replace")
    if raw.startswith(b"\xff\xfe"):
        return raw.decode("utf-16-le", errors="replace")
    if raw.startswith(b"\xfe\xff"):
        return raw.decode("utf-16-be", errors="replace")

    # Strict probe — only accept an encoding that decodes the WHOLE file without error.
    # latin-1 is NOT in this list; it silently accepts any byte and would mask
    # the correct Japanese codec, turning Shift-JIS text into mojibake.
    for enc in _PROBE_ENCODINGS:
        try:
            text = raw.decode(enc)
            if _is_valid_text(text):
                return text
        except (UnicodeDecodeError, LookupError):
            continue


    # chardet fallback — use only when confident (≥ 80 %) to avoid wrong guesses
    try:
        import chardet
        detected = chardet.detect(raw)
        enc = (detected or {}).get("encoding") or ""
        confidence = (detected or {}).get("confidence") or 0.0
        if enc and confidence >= 0.80:
            try:
                return raw.decode(enc)
            except (UnicodeDecodeError, LookupError):
                pass
    except ImportError:
        pass

    # Absolute last resort — cp932 with replacement chars; never raises.
    return raw.decode("cp932", errors="replace")


# ── Scan cache ──────────────────────────────────────────────────────────────
SCAN_CACHE: dict = {}
SCAN_CACHE_LOCK = Lock()


# ── Path helpers ─────────────────────────────────────────────────────────────
def normalize_path(path_str: str) -> str:
    return os.path.normpath(path_str.strip().strip('"'))


def resource_path(relative_path: str) -> str:
    import sys
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)


def summarize_paths(paths: list, kind_singular: str = "item") -> str:
    if not paths:
        return f"No {kind_singular}s selected"
    if len(paths) == 1:
        return paths[0]
    preview = paths[:3]
    remainder = len(paths) - len(preview)
    lines = [f"{len(paths)} {kind_singular}s selected"] + preview
    if remainder > 0:
        lines.append(f"... and {remainder} more")
    return "\n".join(lines)


def iter_source_files(root_folder: str, valid_exts: set = None):
    valid_exts = valid_exts or {".c", ".cpp", ".cc", ".cxx", ".h", ".hpp"}
    for root, _, files in os.walk(root_folder):
        for file_name in sorted(files):
            ext = os.path.splitext(file_name)[1].lower()
            if ext in valid_exts:
                yield os.path.join(root, file_name), file_name


def relative_from_src(full_path: str) -> str:
    norm = normalize_path(full_path)
    lower = norm.lower()
    idx = lower.find(f"{os.sep}src{os.sep}")
    if idx != -1:
        return norm[idx + len(f"{os.sep}src{os.sep}"):]
    return os.path.basename(norm)


def resolve_real_file(selected_src_root: str, raw_file_path: str) -> str:
    raw_norm = normalize_path(raw_file_path)
    if os.path.isfile(raw_norm):
        return raw_norm
    rel = relative_from_src(raw_norm)
    return os.path.join(selected_src_root, rel)


# ── Text helpers ─────────────────────────────────────────────────────────────
def clean_text(value) -> str:
    if value is None:
        return ""
    return str(value).strip()


def normalize_name(value: str) -> str:
    return re.sub(r"\s+", " ", clean_text(value)).strip().lower()


# ── Function name validation ─────────────────────────────────────────────────
INVALID_FUNCTION_TOKENS = {
    "if", "else", "elseif", "for", "while", "do", "switch", "case", "return",
    "defined", "block", "num_channels", "null", "true", "false", "function"
}


def is_probable_function_name(name: str) -> bool:
    s = clean_text(name)
    if not s:
        return False
    n = normalize_name(s)
    if n in INVALID_FUNCTION_TOKENS:
        return False
    if re.fullmatch(r"c", n):
        return False
    if re.fullmatch(r"c_[A-Za-z0-9_]+", s):
        return False
    return bool(re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]*", s))


def extract_function_name_from_cell(s: str) -> str:
    s = clean_text(s)
    if not s:
        return ""
    fn = re.sub(r"\(.*?\)", "", s).strip()
    if is_probable_function_name(fn):
        return fn
    tokens = re.findall(r"[A-Za-z_][A-Za-z0-9_]*", s)
    for token in reversed(tokens):
        if is_probable_function_name(token):
            return token.strip()
    return ""


# ── Excel reference helpers ──────────────────────────────────────────────────
def normalize_excel_reference(value: str) -> str:
    value = clean_text(value).upper()
    if not value:
        return ""
    m = re.fullmatch(r"([A-Z]+)(\d+)?", value)
    if m:
        col, row = m.groups()
        return f"{col}{row or ''}"
    m = re.fullmatch(r"(\d+)([A-Z]+)", value)
    if m:
        row, col = m.groups()
        return f"{col}{row}"
    return value


def is_valid_excel_reference(value: str) -> bool:
    value = normalize_excel_reference(value)
    return bool(re.fullmatch(r"[A-Z]+(\d+)?", value))


def extract_excel_column_letters(value: str) -> str:
    value = normalize_excel_reference(value)
    m = re.fullmatch(r"([A-Z]+)(\d+)?", value)
    return m.group(1) if m else ""


def extract_excel_row_number(value: str) -> int:
    value = normalize_excel_reference(value)
    m = re.fullmatch(r"([A-Z]+)(\d+)?", value)
    return int(m.group(2)) if m and m.group(2) else 1


def excel_col_to_index(col_letters: str) -> int:
    if not col_letters:
        return -1
    return column_index_from_string(col_letters)


# ── File signature for cache ─────────────────────────────────────────────────
def _file_signature(file_path: str):
    try:
        stat = os.stat(file_path)
        return (file_path, stat.st_mtime_ns, stat.st_size)
    except OSError:
        return None


# ── C/C++ function detector ──────────────────────────────────────────────────
def detect_functions_in_file(file_path: str) -> list:
    signature = _file_signature(file_path)
    if signature is None:
        return []
    with SCAN_CACHE_LOCK:
        cached = SCAN_CACHE.get(signature)
    if cached is not None:
        return list(cached)

    # Skip files over 2 MB — regex hangs on large generated C files
    try:
        if os.path.getsize(file_path) > 2 * 1024 * 1024:
            return []
    except OSError:
        return []


    text = read_source_file(file_path)
    if not text:
        return f"Could not read file:\n{file_path}"

    text_no_comments = re.sub(r'/\*.*?\*/', ' ', text, flags=re.DOTALL)
    text_no_comments = re.sub(r'//[^\n]*', ' ', text_no_comments)
    text_no_comments = re.sub(
        r'#\s*if\s+0\b.*?#\s*endif\b',
        lambda m: '\n' * m.group(0).count('\n'),
        text_no_comments, flags=re.DOTALL
    )
    text_no_comments = re.sub(r'"(?:[^"\\]|\\.)*"', '""', text_no_comments)
    text_no_comments = re.sub(r"'(?:[^'\\]|\\.)*'", "''", text_no_comments)

    # Primary pattern: standard function definitions (single or multi-line params,
    # no braces/semicolons inside the parameter list)
    pattern = re.compile(
        r'(?:(?:^|\n)[ \t]{0,4})'
        r'(?:(?:static|extern|inline|__inline__|__forceinline)\s+)?'
        r'(?:(?:const|volatile)\s+)?'
        r'(?:(?:unsigned|signed|long|short)\s+)*'
        r'[A-Za-z_]\w*'
        r'(?:\s+[A-Za-z_]\w*)*?'
        r'\s*\*{0,3}\s*'
        r'\b([A-Za-z_]\w*)\b'
        r'\s*\('
        r'([^;{}]{0,512}?)'          # params: no braces/semicolons, up to 512 chars (allows newlines)
        r'\)'
        r'[ \t\r\n]*'                # optional whitespace/newlines between ) and {
        r'\{'
    )

    # Supplemental pattern: catches multi-line signatures where params span
    # many lines. Looser — validated by _accept() guards below.
    pattern_multiline = re.compile(
        r'(?:^|\n)[ \t]{0,4}'
        r'\b([A-Za-z_]\w*)\b'
        r'\s*\('
        r'([\s\S]{0,1024}?)'         # allow newlines in params
        r'\)'
        r'[ \t\r\n]{0,80}'
        r'\{',
        re.MULTILINE
    )

    IGNORED_NAMES = {
        "if", "else", "for", "while", "do", "switch", "case", "default",
        "break", "continue", "return", "goto",
        "sizeof", "typeof", "typedef", "struct", "union", "enum", "class",
        "static", "extern", "inline", "const", "volatile", "register", "auto",
        "unsigned", "signed", "long", "short", "void", "int", "char", "float",
        "double", "bool", "_Bool", "__attribute__",
        "SETBIT", "CLRBIT", "ASSERT", "assert", "NULL", "TRUE", "FALSE",
        "defined", "offsetof", "alignof", "decltype",
        "c", "f", "i", "j", "k", "n", "p", "s", "t", "x", "y", "z",
    }

    # Build a brace-depth map so we can reject matches inside function bodies.
    # A real top-level function definition always starts at depth 0.
    #
    # Resilience against unbalanced/truncated functions:
    # When a '}' would drop depth below 0, it means we already reached the
    # apparent top-level — clamp to 0.  Additionally, a closing '}' that
    # brings depth from 1 → 0 while at the start of a line is a strong signal
    # that the enclosing function just ended.  If depth never reached 0 between
    # two such function-signature patterns (i.e. the previous function had
    # missing closing braces), we force-reset depth to 0 when we encounter a
    # new top-level-style function header (void/type at column 0 followed by
    # name then '(').  This prevents one malformed function from hiding all
    # subsequent functions.
    _top_level_fn_re = re.compile(
        r'(?:^|\n)(?:(?:static|extern|inline|__inline__|__forceinline)\s+)?'
        r'(?:(?:const|volatile)\s+)?'
        r'(?:(?:unsigned|signed|long|short)\s+)*'
        r'[A-Za-z_]\w*(?:\s+[A-Za-z_]\w*)*?\s*\*{0,3}\s*'
        r'\b([A-Za-z_]\w*)\b\s*\([^;{}]{0,512}?\)\s*\{'
    )
    _top_level_positions = set()
    for _m in _top_level_fn_re.finditer(text_no_comments):
        _top_level_positions.add(_m.start())

    _depth_at = []
    _d = 0
    _prev_was_top_level = False
    for _pos, _ch in enumerate(text_no_comments):
        # If we are at a known top-level function header position and depth > 0,
        # the previous function must have had unbalanced braces — reset depth.
        if _pos in _top_level_positions and _d > 0:
            _d = 0
        _depth_at.append(_d)
        if _ch == '{':
            _d += 1
        elif _ch == '}':
            _d = max(0, _d - 1)

    found = []
    seen = set()

    def _accept(name, params, match_start):
        """Return True if this match is a valid top-level function definition."""
        if name in IGNORED_NAMES:
            return False
        if name in seen:
            return False
        # Params must not contain block-scope characters (means we matched inside code)
        if '{' in params or ';' in params:
            return False
        if not is_probable_function_name(name):
            return False
        # Only accept definitions that begin at brace depth 0 (top-level scope).
        # This prevents nested function-like signatures (e.g. inside #if 0 blocks
        # or another function body) from being listed as separate functions.
        if match_start < len(_depth_at) and _depth_at[match_start] != 0:
            return False
        return True

    # First pass: primary pattern (fast, handles 99% of cases)
    for match in pattern.finditer(text_no_comments):
        name   = match.group(1)
        params = match.group(2).strip()
        if _accept(name, params, match.start()):
            seen.add(name)
            found.append(name)

    # Second pass: multi-line pattern catches signatures missed above
    for match in pattern_multiline.finditer(text_no_comments):
        name   = match.group(1)
        params = match.group(2).strip()
        if _accept(name, params, match.start()):
            seen.add(name)
            found.append(name)

    with SCAN_CACHE_LOCK:
        SCAN_CACHE[signature] = list(found)
    return found


# ── Function body extractor ──────────────────────────────────────────────────
def extract_function_body(file_path: str, function_name: str) -> str:
    if not os.path.isfile(file_path):
        return f"File not found:\n{file_path}"
    text = read_source_file(file_path)
    if not text:
        return f"Could not read file:\n{file_path}"

    def _find_opening_brace_and_extract(match_start: int, match_end: int):
        brace_start = match_end - 1
        depth = 0
        end_pos = None
        i = brace_start
        while i < len(text):
            ch = text[i]
            if ch == "{":
                depth += 1
            elif ch == "}":
                depth -= 1
                if depth == 0:
                    end_pos = i + 1
                    break
            i += 1
        if end_pos is None:
            line_start = match_start
            if line_start < len(text) and text[line_start] == '\n':
                line_start += 1
            return text[line_start:].strip() + "\n/* [incomplete: closing brace not found] */"
        line_start = match_start
        if line_start < len(text) and text[line_start] == '\n':
            line_start += 1
        return text[line_start:end_pos].strip()

    primary_pattern = re.compile(
        r'(?:(?:^|\n)[ \t]{0,4})'
        r'(?:(?:static|extern|inline|__inline__|__forceinline)\s+)?'
        r'(?:(?:const|volatile)\s+)?'
        r'(?:(?:unsigned|signed|long|short)\s+)*'
        r'[A-Za-z_]\w*'
        r'(?:\s+[A-Za-z_]\w*)*?'
        r'\s*\*{0,3}\s*'
        r'\b(' + re.escape(function_name) + r')\b'
        r'\s*\([^;{}]*\)'
        r'[^;{}]*'
        r'\{'
    )

    for match in primary_pattern.finditer(text):
        body = _find_opening_brace_and_extract(match.start(), match.end())
        if body:
            return body

    return f"Body not found for function '{function_name}' in:\n{file_path}"