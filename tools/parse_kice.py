"""평가원 2027 수능 문제지 PDF를 문항별로 나눠 pdf/questions.json을 만든다.

입력: pdf/27DaeSuneung_<번호>.pdf (평가원 '문제 및 정답' 페이지의 파일 이름 그대로, 홀수형만 읽는다)
출력: pdf/questions.json — 시트별 문항의 영역(쪽·좌표)과 텍스트, 공통 지문
  questions.json에는 문제 원문이 들어가므로 pdf/ 폴더째 git에 올리지 않는다.

실행: python -I tools/parse_kice.py pdf

배치 규칙(2027 수능 문제지에서 확인):
  - 쪽 842×1191pt, 두 단(왼쪽 x 80~419, 오른쪽 423~770). 문항 번호 'N.'과 공통 지문 '[a～b]'는
    각 단의 왼쪽 끝(x 87.9 / 436.6)에서 시작한다.
  - 머리말은 쪽 위 가로 줄(폭 600pt 넘는 선) 위, 꼬리말은 y 1076 아래.
  - 국어·수학·영어·한국사 파일은 홀수형 뒤에 짝수형이 붙어 있다. 머리말에 '짝수형'이 나오면 멈춘다.
  - 수식은 한글 수식 글꼴(HyhwpEQ)이 숫자·기울임 영문·기호를 사용자 정의 코드(U+E000~)로 넣는다.
    대응표는 글자 모양을 보고 만들었다(EQ_MAP). 분수 막대·근호 윗줄은 U+E06D 글자이며,
    위아래에 글자가 있으면 분수 '(분자)/(분모)'로, 아니면 버린다. 작은 글자가 올라가 있으면 '^', 내려가 있으면 '_'.
"""
from __future__ import annotations

import json
import re
import statistics
import sys
from datetime import date
from pathlib import Path

import fitz  # PyMuPDF

SOC = ["생활과윤리", "윤리와사상", "한국지리", "세계지리", "동아시아사", "세계사", "경제", "정치와법", "사회문화"]
SCI = ["물리1", "화학1", "생명1", "지구1", "물리2", "화학2", "생명2", "지구2"]
FILES = [("1", "국어"), ("2", "수학"), ("3", "영어"), ("41", "한국사")] + \
        [(f"42_{i + 1}", s) for i, s in enumerate(SOC)] + [(f"43_{i + 1}", s) for i, s in enumerate(SCI)]
# 선택과목 시작 쪽의 머리말(공백 제거)로 시트를 바꾼다. 앞 쪽의 시트는 다음 머리말이 나올 때까지 이어진다.
SECTION_MARKS = {
    "국어": [("(화법과작문)", "국어-화작"), ("(언어와매체)", "국어-언매")],
    "수학": [("(확률과통계)", "수학-확통"), ("(미적분)", "수학-미적"), ("(기하)", "수학-기하")],
}
FIRST_SHEET = {"국어": "국어-공통", "수학": "수학-공통"}
EXPECTED = {"국어-공통": range(1, 35), "국어-화작": range(35, 46), "국어-언매": range(35, 46),
            "수학-공통": range(1, 23), "수학-확통": range(23, 31), "수학-미적": range(23, 31), "수학-기하": range(23, 31),
            "영어": range(1, 46)}

COLS = [(80.0, 419.0), (423.0, 770.0)]
ANCHOR_X = [87.9, 436.6]
TEXT_RIGHT = [406.5, 754.5]
CONTENT_BOTTOM = 1076.0
NUM_RE = re.compile(r"^\s*(\d{1,2})\s*\.(?!\d)")
RNG_RE = re.compile(r"^\s*\[\s*(\d{1,2})\s*[～~∼〜\-–]\s*(\d{1,2})\s*\]")
CHOICE = "①②③④⑤"
SPEAKER = r"(?:[갑을병정무]|학생\s?\d?|교사|선생님|사회자)\s?:"
BULLET_RE = re.compile(r"^([ㄱ-ㅎ]\.\s|[◦•∙·]|" + SPEAKER + ")")  # 앞 줄이 끝까지 차도 이어 붙이지 않는 줄머리
ITEM_RE = re.compile(r"^(\d{1,2}\.(?!\d)|[ㄱ-ㅎ]\.|\([가-하]\)|[◦•∙·*]|" + SPEAKER + ")")  # 다음 줄이 내어쓰기인 줄머리
POINTS_RE = re.compile(r"^\[\d점\]")
HEADER_RE = re.compile(r"^<[^<>]{1,12}>(\s+<[^<>]{1,12}>)*$")  # '<보 기>', '<조 건>'처럼 한 줄 전체가 상자 제목인 줄
PLACE_RE = re.compile(r"^\([가-하]\)")

BAR = 0xE06D
EQ_SYM = {0xE044: "(", 0xE045: ")", 0xE046: "−", 0xE047: "=", 0xE048: "+", 0xE049: "[", 0xE04A: "]",
          0xE04B: "{", 0xE04C: "}", 0xE04D: "|", 0xE04F: ":", 0xE052: ",", 0xE053: ".", 0xE054: "/",
          0xE055: "<", 0xE056: ">", 0xE05B: "∫", 0xE05C: "√", 0xE067: "Σ", 0xE06E: "→", 0xE088: "Δ",
          0xE079: "{", 0xE078: "", 0xE07A: "", 0xE07B: "", 0xE101: "|"}
GREEK = "αβγδεζηθικλμνξοπρστυφχψω"


def is_pua(cp: int) -> bool:
    return 0xE000 <= cp <= 0xF8FF or cp == 0xFFFD


def map_eq(cp: int) -> str | None:
    """HyhwpEQ 사용자 정의 코드 → 문자. 모르는 코드는 None."""
    if 0xE000 <= cp <= 0xE019:
        return chr(ord("A") + cp - 0xE000)
    if 0xE0E5 <= cp <= 0xE0FE:
        return chr(ord("a") + cp - 0xE0E5)
    if 0xE034 <= cp <= 0xE03D:
        return "1234567890"[cp - 0xE034]
    if 0xE09D <= cp < 0xE09D + len(GREEK):
        return GREEK[cp - 0xE09D]
    return EQ_SYM.get(cp)


class Tok:
    __slots__ = ("t", "x0", "y0", "x1", "y1", "size", "eq", "bar", "space", "frac")

    def __init__(self, t, bbox, size, eq=False, bar=False, frac=False):
        self.t = t
        self.x0, self.y0, self.x1, self.y1 = bbox
        self.size = size
        self.eq = eq
        self.bar = bar
        self.frac = frac
        self.space = t.isspace() if t else False

    @property
    def cx(self):
        return (self.x0 + self.x1) / 2

    @property
    def cy(self):
        return (self.y0 + self.y1) / 2


def page_chars(page) -> tuple[list[Tok], int]:
    """쪽의 모든 글자를 Tok으로. 수식 글꼴 코드는 대응표로 바꾸고, 못 바꾼 글자는 □."""
    toks, unknown = [], 0
    for b in page.get_text("rawdict")["blocks"]:
        if b["type"] != 0:
            continue
        for ln in b["lines"]:
            for sp in ln["spans"]:
                eqfont = "EQ" in sp["font"].upper()
                for ch in sp["chars"]:
                    c, cp = ch["c"], ord(ch["c"])
                    if cp == BAR and eqfont:
                        toks.append(Tok("", ch["bbox"], sp["size"], eq=True, bar=True))
                        continue
                    if is_pua(cp):
                        m = map_eq(cp) if eqfont else None
                        if m is None:
                            m, unknown = "□", unknown + 1
                        c = m
                    if c == "":
                        continue
                    toks.append(Tok(c, ch["bbox"], sp["size"], eq=eqfont))
    return toks, unknown


def wrap(s: str) -> str:
    return s if len(s) <= 1 or re.fullmatch(r"[0-9]+|[A-Za-z]", s) else f"({s})"


def script_kind(t: Tok, m: float, base: float) -> str:
    """'n' 보통, '^' 위첨자, '_' 아래첨자. 분수 덩어리는 막대 높이로만 판단한다."""
    if t.size >= 0.85 * m:
        return "n"
    d = base - t.cy
    lim = 0.3 * m if t.frac else 0.18 * m
    return "^" if d > lim else "_" if d < -lim else "n"


def inline(toks: list[Tok]) -> str:
    """한 줄 안의 글자들을 x 순으로 이어 붙이며 위·아래 첨자를 표시한다.
    lim·Σ처럼 글자 바로 아래·위에 겹쳐 놓인 첨자 묶음은 그 글자들 뒤에 '_( )', '^( )'로 붙인다."""
    toks = sorted((t for t in toks if not t.space), key=lambda t: t.x0)
    if not toks:
        return ""
    m = max(t.size for t in toks)
    big = [t.cy for t in toks if t.size >= 0.85 * m]
    base = statistics.median(big) if big else statistics.median(t.cy for t in toks)
    kinds = {}
    for t in toks:
        k = script_kind(t, m, base)
        kinds[id(t)] = "n" if (t.t == "→" and k == "^") else k  # 벡터 화살표는 글자 위에 그대로
    normal = [t for t in toks if kinds[id(t)] == "n"]
    scripts = []
    for t in toks:
        if kinds[id(t)] == "n":
            continue
        host = next((n for n in normal if n.t.startswith("√") and n.x0 - 0.5 <= t.cx <= n.x1 + 0.5), None)
        if host is not None and kinds[id(t)] == "^" and host.t == "√":
            host.t = f"√[{t.t}]"  # 근호 지표(세제곱근 등)
            continue
        scripts.append(t)
    clusters: list[list[Tok]] = []
    for t in scripts:
        if clusters and kinds[id(clusters[-1][-1])] == kinds[id(t)] and t.x0 - max(u.x1 for u in clusters[-1]) < 0.3 * m:
            clusters[-1].append(t)
        else:
            clusters.append([t])
    items = [(t.x0, 1, t.t) for t in normal]
    for c in clusters:
        cx0, cx1 = min(t.x0 for t in c), max(t.x1 for t in c)
        ov = sum(max(0.0, min(cx1, n.x1) - max(cx0, n.x0)) for n in normal)
        text = kinds[id(c[0])] + wrap("".join(t.t for t in c))
        if ov >= 0.4 * (cx1 - cx0):
            nxt = [n.x0 for n in normal if n.x0 >= cx1 - 0.3 * m]
            items.append(((min(nxt) - 1e-3) if nxt else float("inf"), 0 if text[0] == "_" else 0.5, text))
        else:
            items.append((cx0, 1, text))
    items.sort(key=lambda it: (it[0], it[1]))
    return "".join(it[2] for it in items)


def resolve_fractions(toks: list[Tok]) -> list[Tok]:
    """분수 막대(U+E06D) 위아래의 수식 글자를 '(분자)/(분모)' 한 덩어리로 바꾼다.
    막대 글리프 상자와 겹치는 수식 글자를 세로로 묶어, 사이에 빈 틈이 있는 두 묶음이면 분수로 본다
    (작은 분수는 상자가 분자·분모에 걸치고, 큰 분수는 상자가 둘을 덮는다 — 1번·미적분 23번에서 확인).
    한 묶음뿐이면 근호 윗줄·선분 기호로 보고 막대를 버린다. 본문 글자는 넣지 않는다
    (윗줄 '점근선'의 글자를 분자로 잘못 가져온 일이 있었다). 분자 지수처럼 상자 밖으로 나간 글자는 세로로 이어진 것만 더한다."""
    bars = sorted((t for t in toks if t.bar), key=lambda t: t.x1 - t.x0)
    used: set[int] = set()
    rest = [t for t in toks if not t.bar]
    for b in bars:
        xin = [t for t in rest if id(t) not in used and t.eq and not t.space and b.x0 - 0.8 <= t.cx <= b.x1 + 0.8]
        pool = [t for t in xin if t.y1 > b.y0 + 0.5 and t.y0 < b.y1 - 0.5]
        if len(pool) < 2:
            continue
        merged: list[list[float]] = []
        for a0, a1 in sorted((t.y0, t.y1) for t in pool):
            if merged and a0 <= merged[-1][1] + 0.3:
                merged[-1][1] = max(merged[-1][1], a1)
            else:
                merged.append([a0, a1])
        if len(merged) < 2:
            continue
        _, gi = max((merged[i + 1][0] - merged[i][1], i) for i in range(len(merged) - 1))
        stroke = (merged[gi][1] + merged[gi + 1][0]) / 2
        above = [t for t in pool if t.cy < stroke]
        below = [t for t in pool if t.cy > stroke]
        for side, up in ((above, True), (below, False)):
            grown = True
            while grown:
                grown = False
                lo, hi = min(t.y0 for t in side), max(t.y1 for t in side)
                for t in xin:
                    if t in side or t in pool or (up and t.cy >= stroke) or (not up and t.cy <= stroke):
                        continue
                    if (up and t.y1 >= lo - 0.2 * t.size) or (not up and t.y0 <= hi + 0.2 * t.size):
                        side.append(t)
                        grown = True
        cand = above + below
        for t in cand:
            used.add(id(t))
        size = max(t.size for t in cand)
        rest.append(Tok(f"{wrap(inline(above))}/{wrap(inline(below))}", (b.x0, stroke - size / 2, b.x1, stroke + size / 2),
                        size, eq=True, frac=True))
    return [t for t in rest if id(t) not in used]


def make_rows(toks: list[Tok]) -> list[list[Tok]]:
    """보통 크기 글자로 줄을 먼저 만들고, 작은 글자(첨자·분수 안 글자)는 가장 가까운 줄에 붙인다."""
    sizes = [t.size for t in toks if not t.space]
    if not sizes:
        return []
    med = statistics.median(sizes)
    rows: list[dict] = []
    for t in sorted((t for t in toks if t.space or t.size >= 0.8 * med), key=lambda t: t.cy):
        if rows and abs(t.cy - rows[-1]["cy"]) <= 0.5 * max(rows[-1]["size"], t.size):
            r = rows[-1]
            r["toks"].append(t)
            if not t.space:
                r["n"] += 1
                r["cy"] += (t.cy - r["cy"]) / r["n"]
                r["size"] = max(r["size"], t.size)
            continue
        rows.append({"toks": [t], "cy": t.cy, "size": t.size, "n": 0 if t.space else 1})
    for t in (t for t in toks if not t.space and t.size < 0.8 * med):
        best = min(rows, key=lambda r: abs(t.cy - r["cy"]), default=None)
        if best is not None and abs(t.cy - best["cy"]) <= 1.3 * best["size"]:
            best["toks"].append(t)
        else:
            rows.append({"toks": [t], "cy": t.cy, "size": t.size, "n": 1})
    return [sorted(r["toks"], key=lambda t: t.x0) for r in sorted(rows, key=lambda r: r["cy"])]


NOTE_LINES = {"단답형", "5지선다형"}


def row_gaps(row: list[Tok], m: float) -> list[tuple[float, float]]:
    ns = [t for t in row if not t.space]
    return [(a.x1, b.x0) for a, b in zip(ns, ns[1:]) if b.x0 - a.x1 >= 1.0 * m]


def split_side(rows: list[list[Tok]], med: float) -> list[tuple[list[list[Tok]], bool]]:
    """본문 옆에 표·그림 글자가 붙은 구간을 찾아 왼쪽 묶음과 오른쪽 묶음으로 나눈다.
    같은 x 틈이 3줄 넘게 이어지고, 한쪽이 산문(줄당 평균 10자 이상)일 때만 나눈다. 선택지 줄은 나누지 않는다.
    반환: (줄 묶음, 나뉜 묶음인지) 목록. 각 묶음은 위에서 아래로 읽는 줄들."""
    def ok_row(r):
        txt = "".join(t.t for t in r)
        return not any(c in txt for c in CHOICE)
    blocks, pending, i = [], [], 0
    while i < len(rows):
        best = None
        if ok_row(rows[i]):
            for g0, g1 in row_gaps(rows[i], med):
                lo, hi, j, two = g0, g1, i + 1, 1
                while j < len(rows) and ok_row(rows[j]):
                    ns = [t for t in rows[j] if not t.space]
                    if any(t.x0 < hi - 1 and t.x1 > lo + 1 for t in ns):
                        nlo, nhi = None, None
                        for a0, a1 in row_gaps(rows[j], med):
                            if min(hi, a1) - max(lo, a0) >= 4:
                                nlo, nhi = max(lo, a0), min(hi, a1)
                                break
                        if nlo is None:
                            break
                        lo, hi = nlo, nhi
                    if any(t.cx < lo for t in ns) and any(t.cx > hi for t in ns):
                        two += 1
                    j += 1
                if two >= 3 and (best is None or j > best[2]):
                    best = (lo, hi, j)
        if best:
            lo, hi, j = best
            split = (lo + hi) / 2
            left = [[t for t in r if t.cx < split] for r in rows[i:j]]
            right = [[t for t in r if t.cx >= split] for r in rows[i:j]]
            chars = lambda part: [len([t for t in r if not t.space]) for r in part if any(not t.space for t in r)]
            lc, rc = chars(left), chars(right)
            prose = (lc and sum(lc) / len(lc) >= 10) or (rc and sum(rc) / len(rc) >= 10)
            if prose:
                if pending:
                    blocks.append((pending, False))
                    pending = []
                blocks.append(([r for r in left if r], True))
                blocks.append(([r for r in right if r], True))
                i = j
                continue
        pending.append(rows[i])
        i += 1
    if pending:
        blocks.append((pending, False))
    return blocks


def block_lines(rows: list[list[Tok]]) -> list[tuple[str, float, float, float]]:
    lines = []
    for row in rows:
        m = max(t.size for t in row)
        parts, prev, seg = [], None, []

        def flush():
            if seg:
                parts.append(inline(seg))
                seg.clear()
        for t in row:
            if t.space:
                flush()
                if parts and not parts[-1].endswith(" "):
                    parts.append(" ")
                prev = t
                continue
            if prev is not None:
                gap = t.x0 - prev.x1
                if t.t in CHOICE and (seg or parts):
                    flush()  # 한 줄에 선택지가 여러 개면 사이를 넉넉히 띄운다
                    parts.append("   ")
                elif not prev.space and (gap > 1.6 * m or (prev.t in CHOICE and gap > 0.1 * m)):
                    flush()
                    parts.append("   " if gap > 1.6 * m else " ")
            seg.append(t)
            prev = t
        flush()
        s = re.sub(r" {4,}", "   ", "".join(parts)).strip()
        if s and s.replace(" ", "") not in NOTE_LINES:
            ink = [t for t in row if not t.space]
            lines.append((s, max(t.x1 for t in ink) if ink else 0, m, min(t.x0 for t in ink) if ink else 0))
    return lines


def clip_text(toks: list[Tok], text_right: float) -> str:
    rows = make_rows(resolve_fractions(toks))
    sizes = [t.size for r in rows for t in r if not t.space]
    if not sizes:
        return ""
    blocks = split_side(rows, statistics.median(sizes))
    outs = []
    for rws, is_split in blocks:
        lines = block_lines(rws)
        if not lines:
            continue
        # 오른쪽 끝까지 찬 줄은 줄바꿈이 아니라 이어지는 문장으로 본다(나뉜 묶음은 그 묶음의 오른쪽 끝 기준)
        right_edge = max(ln[1] for ln in lines) if is_split else text_right
        out, joined = "", False
        for k, (s, right, m, left) in enumerate(lines):
            out += s
            if k + 1 < len(lines):
                nxt, nright, nm, nleft = lines[k + 1]
                full = right >= right_edge - 1.2 * m
                # 끝까지 찬 줄 뒤라도 다음 줄이 들여 쓰여 있으면 새 문단이다. 단, 이 줄이 문단의 첫 줄이거나
                # (번호·보기·'갑:' 다음의 내어쓰기) 그다음 줄도 같은 위치에서 이어지면(내어쓰기) 문단이 아니다.
                # '(가)' 빈칸 상자로 시작하는 줄은 상자 여백 때문에 들여 쓴 것처럼 보이므로 뺀다.
                hang = k + 2 < len(lines) and nright >= right_edge - 1.2 * nm and abs(lines[k + 2][3] - nleft) < 0.3 * m
                indent = (nleft > left + 0.6 * m and joined and not hang and not PLACE_RE.match(nxt) and not POINTS_RE.match(nxt)
                          and not (s[0] in CHOICE or ITEM_RE.match(s)))
                joined = full and not indent and nxt[0] not in CHOICE and not RNG_RE.match(nxt) and not BULLET_RE.match(nxt) and not HEADER_RE.match(nxt)
                out += " " if joined else "\n"
        outs.append(out)
    return "\n".join(outs)


def content_top(page) -> float:
    """머리말 아래 가로 줄(폭 600pt 넘는 선) 중 가장 아래 것. 없으면 150."""
    ys = []
    for d in page.get_drawings():
        for it in d["items"]:
            if it[0] == "l":
                r = fitz.Rect(it[1], it[2]).normalize()
            elif it[0] == "re":
                r = fitz.Rect(it[1]).normalize()
            else:
                continue
            if r.width > 600 and r.height < 3 and r.y0 < 300:
                ys.append(r.y1)
    return (max(ys) + 2) if ys else 150.0


def parse_file(path: Path, subject: str) -> tuple[dict, list[str]]:
    doc = fitz.open(path)
    warn = []
    sheet = FIRST_SHEET.get(subject, subject)
    segs: list[dict] = []
    seen: dict[str, set] = {}
    current = None
    for pno, page in enumerate(doc):
        top = content_top(page)
        lines = []
        for b in page.get_text("dict")["blocks"]:
            if b["type"] != 0:
                continue
            for ln in b["lines"]:
                t = "".join(sp["text"] for sp in ln["spans"])
                # 줄 상자는 큰 근호·분수 때문에 위로 커질 수 있어, 첫 글자 묶음(번호)의 상자를 쓴다
                lines.append((ln["spans"][0]["bbox"] if ln["spans"] else ln["bbox"], t, ln["bbox"]))
        header = "".join(t for bb, t, _ in lines if bb[1] < top).replace(" ", "")
        if "짝수형" in header:
            break
        for mark, sh in SECTION_MARKS.get(subject, []):
            if mark in header and sh != sheet:
                sheet, current = sh, None
        toks, unknown = page_chars(page)
        draws = [d["rect"] for d in page.get_drawings()]
        imgs = [fitz.Rect(i["bbox"]) for i in page.get_image_info()]
        for ci, (cx0, cx1) in enumerate(COLS):
            anchors = []
            for bb, t, lbb in lines:
                # 경계: 번호 줄의 위첨자·분자·큰 근호는 번호보다 위로 올라가므로 줄 상자 위와 번호 위 8pt 중 높은 쪽
                cut = min(lbb[1] - 2, bb[1] - 8)
                if not (top <= bb[1] < CONTENT_BOTTOM):
                    continue
                # 선택과목 끝의 '* 확인 사항' 안내 상자, '단답형'·'5지선다형' 표시부터는 앞 문항에 넣지 않는다
                if cx0 <= bb[0] < cx1 and ((t.lstrip().startswith("*") and "확인" in t) or t.replace(" ", "") in NOTE_LINES):
                    anchors.append((cut, "X", None))
                    continue
                if abs(bb[0] - ANCHOR_X[ci]) > 2.5:
                    continue
                m = RNG_RE.match(t)
                if m:
                    anchors.append((cut, "P", (int(m.group(1)), int(m.group(2)))))
                    continue
                m = NUM_RE.match(t)
                if m:
                    n = int(m.group(1))
                    rng = EXPECTED.get(sheet, range(1, 21))
                    if n in rng and n not in seen.setdefault(sheet, set()):
                        anchors.append((cut, "Q", n))
            anchors.sort(key=lambda a: a[0])
            bounds = [top] + [max(top, a[0]) for a in anchors] + [CONTENT_BOTTOM]
            for k in range(len(bounds) - 1):
                y0, y1 = bounds[k], bounds[k + 1]
                if k == 0:
                    seg = current
                elif anchors[k - 1][1] == "X":
                    seg = current = None
                else:
                    _, kind, key = anchors[k - 1]
                    seg = {"kind": kind, "key": key, "sheet": sheet, "clips": [], "toks": [], "eq": False, "unknown": 0}
                    if kind == "Q":
                        seen[sheet].add(key)
                    segs.append(seg)
                    current = seg
                if seg is None:
                    continue
                inside = lambda r: cx0 <= (r.x0 + r.x1) / 2 <= cx1 and y0 <= (r.y0 + r.y1) / 2 < y1
                ct = [t for t in toks if cx0 <= t.cx <= cx1 and y0 <= t.cy < y1]
                boxes = [fitz.Rect(t.x0, t.y0, t.x1, t.y1) for t in ct if not t.space] + \
                        [r for r in draws if inside(r) and r.width < 400] + [r for r in imgs if inside(r)]
                if not boxes:
                    continue
                by0 = max(y0, min(r.y0 for r in boxes) - 4)
                by1 = min(y1, max(r.y1 for r in boxes) + 4)
                seg["clips"].append([pno + 1, cx0, round(by0, 1), cx1, round(by1, 1)])
                seg["texts"] = seg.get("texts", []) + [clip_text(ct, TEXT_RIGHT[ci])]
                seg["eq"] = seg["eq"] or any(t.eq for t in ct)
                seg["unknown"] += sum(1 for t in ct if t.t == "□")
    sheets: dict[str, dict] = {}
    for s in segs:
        sh = sheets.setdefault(s["sheet"], {"file": path.name, "questions": {}, "passages": {}})
        text = "\n".join(x for x in s.get("texts", []) if x)
        if s["kind"] == "P":
            a, b = s["key"]
            sh["passages"][f"{a}-{b}"] = {"range": [a, b], "clips": s["clips"], "text": text}
        else:
            sh["questions"][str(s["key"])] = {"clips": s["clips"], "text": text, "eq": s["eq"], "unknown": s["unknown"]}
    for name, sh in sheets.items():
        for q in sh["questions"]:
            n = int(q)
            for pid, p in sh["passages"].items():
                if p["range"][0] <= n <= p["range"][1]:
                    sh["questions"][q]["passage"] = pid
        exp = EXPECTED.get(name, range(1, 21))
        missing = [n for n in exp if str(n) not in sh["questions"]]
        if missing:
            warn.append(f"{name}: 못 찾은 문항 {missing}")
        empty = [q for q, v in sh["questions"].items() if not v["clips"]]
        if empty:
            warn.append(f"{name}: 영역이 빈 문항 {empty}")
    return sheets, warn


def main():
    pdf_dir = Path(sys.argv[1] if len(sys.argv) > 1 else "pdf")
    out = {"version": 1, "generated": date.today().isoformat(),
           "source": "한국교육과정평가원 2027학년도 대학수학능력시험 문제지(홀수형)", "sheets": {}}
    warns = []
    for code, subject in FILES:
        p = pdf_dir / f"27DaeSuneung_{code}.pdf"
        if not p.exists():
            warns.append(f"파일 없음: {p.name}")
            continue
        sheets, w = parse_file(p, subject)
        out["sheets"].update(sheets)
        warns += w
    dst = pdf_dir / "questions.json"
    dst.write_text(json.dumps(out, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    sys.stdout.reconfigure(encoding="utf-8")
    nq = sum(len(s["questions"]) for s in out["sheets"].values())
    np_ = sum(len(s["passages"]) for s in out["sheets"].values())
    print(f"{dst}: 시트 {len(out['sheets'])}개, 문항 {nq}개, 공통 지문 {np_}개")
    for w in warns:
        print("경고:", w)


if __name__ == "__main__":
    main()
