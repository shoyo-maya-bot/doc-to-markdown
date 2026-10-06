"""Excel のセルの表示形式（number_format）に従って値を文字列にする.

`render_value(値, 表示形式)` は `(文字列, 再現できたか)` を返す。再現できない表示形式（指数・分数・条件付き・
システムの長い日付 等）は書式を付けない値（従来の出力）を返し、2 つ目を False にする（呼び出し側がロスとして報告する）。

| 表示形式の例 | 値 | 出力 |
|---|---|---|
| `000` / `0000` / `000-0000` | 1 / 42 / 1234567 | `001` / `0042` / `123-4567` |
| `0.00` / `#,##0` / `#,##0.00` | 1.5 / 1234567 / 1234.5 | `1.50` / `1,234,567` / `1,234.50` |
| `0%` / `0.0%` | 0.125 | `13%` / `12.5%` |
| `"¥"#,##0` / `#,##0"円"` | 1500 | `¥1,500` / `1,500円` |
| `#,##0;[Red]-#,##0` / `#,##0_);(#,##0)` | -1500 | `-1,500` / `(1,500)` |
| `@`（文字列） | `007` | `007` |
| `yyyy/mm/dd` / `m/d` / `yyyy"年"m"月"d"日"` | 2026-04-05 | `2026/04/05` / `4/5` / `2026年4月5日` |
| `h:mm` / `hh:mm:ss` / `h:mm AM/PM` | 9:07:03 | `9:07` / `09:07:03` / `9:07 AM` |
| `ggge"年"m"月"d"日"` / `ge.m.d` / `m月d日(aaa)` | 2026-04-05 | `令和8年4月5日` / `R8.4.5` / `4月5日(日)` |
| `General` / `mm-dd-yy`（地域設定に従う既定の日付） | — | 従来どおり（日付は ISO 形式） |
"""

from __future__ import annotations

import datetime as dt
import math
import re
from decimal import ROUND_HALF_UP, Decimal

# 地域設定（Windows の短い日付形式）に従って表示が変わる組み込み書式. 決まった見た目が無いので ISO 形式で出す
LOCALE_DEFAULT = {"mm-dd-yy", "m/d/yy h:mm", "m/d/yy"}

_COLORS = {"black", "blue", "cyan", "green", "magenta", "red", "white", "yellow"}
_DATE_LETTERS = set("ymdhsge")
_MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October",
           "November", "December"]  # fmt: skip
_WEEKDAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
_WEEKDAYS_JA = "月火水木金土日"
# (開始日, 名前, 略称, ローマ字)
_ERAS = [
    (dt.date(2019, 5, 1), "令和", "令", "R"),
    (dt.date(1989, 1, 8), "平成", "平", "H"),
    (dt.date(1926, 12, 25), "昭和", "昭", "S"),
    (dt.date(1912, 7, 30), "大正", "大", "T"),
    (dt.date(1868, 10, 23), "明治", "明", "M"),
]


class Unsupported(Exception):
    """この表示形式は再現できない."""


def general(v) -> str:
    """表示形式「標準」の出力（書式を再現できないときの既定もこれ）."""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, float):
        if math.isnan(v) or math.isinf(v):
            return str(v)
        return str(int(v)) if v.is_integer() else repr(v)
    if isinstance(v, dt.datetime):
        return v.date().isoformat() if v.time() == dt.time(0) else v.isoformat(sep=" ", timespec="minutes")
    if isinstance(v, (dt.date, dt.time)):
        return v.isoformat()
    return str(v)


def render_value(v, number_format: str = "") -> tuple[str, bool]:
    if v is None:
        return "", True
    fmt = number_format or ""
    key = fmt.strip().lower()
    if isinstance(v, bool) or not key or key == "general" or key in LOCALE_DEFAULT:
        return general(v), True
    try:
        sections = _split_sections(fmt)
        if isinstance(v, str):
            return _render_text(v, sections), True
        if isinstance(v, (int, float)):
            if isinstance(v, float) and (math.isnan(v) or math.isinf(v)):
                return general(v), True
            return _render_number(v, sections), True
        if isinstance(v, (dt.datetime, dt.date, dt.time)):
            return _render_datetime(v, sections[0]), True
    except Unsupported:
        pass
    return general(v), False


# ---------------------------------------------------------------------------
# 字句解析
# ---------------------------------------------------------------------------
def _split_sections(fmt: str) -> list[str]:
    """`;` で区切る（引用符・\\ のエスケープ・[] の中の ; は区切りではない）."""
    out, cur, i = [], "", 0
    while i < len(fmt):
        ch = fmt[i]
        if ch == '"':
            j = fmt.find('"', i + 1)
            j = len(fmt) - 1 if j < 0 else j
            cur += fmt[i : j + 1]
            i = j + 1
            continue
        if ch in "\\_*" and i + 1 < len(fmt):
            cur += fmt[i : i + 2]
            i += 2
            continue
        if ch == "[":
            j = fmt.find("]", i)
            j = len(fmt) - 1 if j < 0 else j
            cur += fmt[i : j + 1]
            i = j + 1
            continue
        if ch == ";":
            out.append(cur)
            cur = ""
        else:
            cur += ch
        i += 1
    out.append(cur)
    return out


def _bracket(body: str) -> tuple[str, str] | None:
    """[...] の中身. 表示に影響しないもの（色・言語）は None、通貨記号は ("lit", 記号)."""
    low = body.lower()
    if low in _COLORS or re.fullmatch(r"color\d{1,2}", low):
        return None
    if body.startswith("$"):
        sym, _, locale = body[1:].partition("-")
        loc = locale.lower()
        if loc.startswith(("f800", "f400")) or "x-sys" in loc:  # システムの長い日付・時刻（地域設定次第）
            raise Unsupported(body)
        if "gannen" in loc:
            return ("gannen", "")
        return ("lit", sym) if sym else None
    raise Unsupported(body)  # 条件 [>=100]、経過時間 [h]、[DBNum1] 等


def _tokenize(sec: str) -> list[tuple[str, str]]:
    toks: list[tuple[str, str]] = []
    i, n = 0, len(sec)
    while i < n:
        ch = sec[i]
        low = ch.lower()
        if ch == '"':
            j = sec.find('"', i + 1)
            j = n if j < 0 else j
            toks.append(("lit", sec[i + 1 : j]))
            i = j + 1
        elif ch == "\\" and i + 1 < n:
            toks.append(("lit", sec[i + 1]))
            i += 2
        elif ch in "_*":  # 幅合わせの空白・繰り返し文字は出力しない
            i += 2
        elif ch == "[":
            j = sec.find("]", i)
            j = n if j < 0 else j
            t = _bracket(sec[i + 1 : j])
            if t:
                toks.append(t)
            i = j + 1
        elif ch == "@":
            toks.append(("text", "@"))
            i += 1
        elif sec[i : i + 7].lower() == "general":
            toks.append(("general", ""))
            i += 7
        elif sec[i : i + 5].upper() == "AM/PM":
            toks.append(("ampm", "AM/PM"))
            i += 5
        elif sec[i : i + 3].upper() == "A/P":
            toks.append(("ampm", "A/P"))
            i += 3
        elif low == "e" and i + 1 < n and sec[i + 1] in "+-":
            raise Unsupported("指数")
        elif low in _DATE_LETTERS or (low == "a" and sec[i : i + 3].lower() == "aaa"):
            j = i
            while j < n and sec[j].lower() == low:
                j += 1
            toks.append(("date", low * (j - i)))
            i = j
        elif ch in "0#?.,%/":
            toks.append(("num", ch))
            i += 1
        else:
            toks.append(("lit", ch))
            i += 1
    return toks


# ---------------------------------------------------------------------------
# 文字列・数値
# ---------------------------------------------------------------------------
def _render_text(v: str, sections: list[str]) -> str:
    sec = sections[3] if len(sections) >= 4 else sections[0] if len(sections) == 1 else ""
    toks = _tokenize(sec) if "@" in sec else []
    if not any(k == "text" for k, _ in toks):
        return v
    return "".join(v if k == "text" else t for k, t in toks if k in ("text", "lit"))


def _render_number(v: float, sections: list[str]) -> str:
    sign = ""
    if v < 0 and len(sections) >= 2 and sections[1] != "":
        sec, v = sections[1], -v
    elif v == 0 and len(sections) >= 3 and sections[2] != "":
        sec = sections[2]
    else:
        sec = sections[0]
        if v < 0:
            sign, v = "-", -v
    toks = _tokenize(sec)
    kinds = {k for k, _ in toks}
    if kinds & {"date", "ampm", "gannen"}:
        raise Unsupported("数値に日付の書式")
    if "general" in kinds:
        body = "".join(general(v) if k == "general" else t for k, t in toks if k in ("general", "lit"))
        return sign + body
    if "text" in kinds:
        if all(k == "text" for k, _ in toks):  # 「文字列」書式の数値は標準と同じ見た目
            return sign + general(v)
        raise Unsupported("数値に文字列の書式")
    if any(k == "num" and t == "/" for k, t in toks):
        raise Unsupported("分数")
    text = _digits(Decimal(repr(float(v))) if isinstance(v, float) else Decimal(v), toks)
    if sign and not re.search(r"[1-9]", text):  # 丸めて 0 になった負数に - は付けない
        sign = ""
    return sign + text


def _digits(val: Decimal, toks: list[tuple[str, str]]) -> str:
    ph = "0#?"
    val *= Decimal(100) ** sum(1 for k, t in toks if k == "num" and t == "%")
    dot = next((i for i, (k, t) in enumerate(toks) if k == "num" and t == "."), len(toks))
    int_toks, frac_toks = toks[:dot], toks[dot + 1 :]
    # 整数部の , : 後ろに桁の記号が続けば桁区切り、続かなければ 1000 で割る
    int_ph = [i for i, (k, t) in enumerate(int_toks) if k == "num" and t in ph]
    grouping = False
    for i, (k, t) in enumerate(int_toks):
        if k == "num" and t == ",":
            if any(p > i for p in int_ph):
                grouping = grouping or any(p < i for p in int_ph)
            elif int_ph:
                val /= 1000
    n_frac = sum(1 for k, t in frac_toks if k == "num" and t in ph)
    q = val.quantize(Decimal(1).scaleb(-n_frac), rounding=ROUND_HALF_UP)
    int_s, _, frac_s = f"{q:f}".partition(".")
    int_s = int_s.lstrip("0")
    if int_ph and grouping and any(k == "lit" for k, _ in int_toks[int_ph[0] : int_ph[-1]]):
        raise Unsupported("桁区切りと文字の混在")

    out_int = [t if k == "lit" else "" for k, t in int_toks]
    if int_ph:
        marks = [int_toks[i][1] for i in int_ph]
        if grouping:
            s = int_s.zfill(marks.count("0"))
            out_int[int_ph[0]] = _group(s) if s else ""
        else:
            digits = list(int_s)
            for i in reversed(int_ph):
                m = int_toks[i][1]
                if i == int_ph[0]:
                    rest = "".join(digits)
                    out_int[i] = rest or ("0" if m == "0" else " " if m == "?" else "")
                elif digits:
                    out_int[i] = digits.pop()
                else:
                    out_int[i] = "0" if m == "0" else " " if m == "?" else ""
    elif int_s and dot < len(toks):
        # 整数部に桁の記号が無い（`.00` 等）: 整数部を小数点の前に出す
        out_int.append(int_s)

    out_frac: list[str] = []
    if dot < len(toks):
        fd = list(frac_s)
        marks = [t for k, t in frac_toks if k == "num" and t in ph]
        # 末尾の # は 0 を省き、? は空白にする
        for j in range(len(marks) - 1, -1, -1):
            if fd[j] != "0" or marks[j] == "0":
                break
            fd[j] = "" if marks[j] == "#" else " "
        it = iter(fd)
        for k, t in frac_toks:
            if k == "num" and t in ph:
                out_frac.append(next(it))
            elif k == "lit" or (k == "num" and t == "%"):
                out_frac.append(t)
        out_frac.insert(0, ".")
    for i, (k, t) in enumerate(int_toks):
        if k == "num" and t == "%":
            out_int[i] = "%"
    return "".join(out_int) + "".join(out_frac)


def _group(s: str) -> str:
    head = len(s) % 3 or 3
    return ",".join([s[:head]] + [s[i : i + 3] for i in range(head, len(s), 3)])


# ---------------------------------------------------------------------------
# 日付・時刻
# ---------------------------------------------------------------------------
def _era(d: dt.date) -> tuple[int, str, str, str]:
    for start, name, short, roman in _ERAS:
        if d >= start:
            return d.year - start.year + 1, name, short, roman
    raise Unsupported("明治より前の和暦")


def _render_datetime(v, sec: str) -> str:
    toks = _tokenize(sec)
    kinds = {k for k, _ in toks}
    if "date" not in kinds:
        raise Unsupported("日付に数値の書式")
    if isinstance(v, dt.datetime):
        d, t = v.date(), v.time()
    elif isinstance(v, dt.date):
        d, t = v, dt.time(0)
    else:
        d, t = None, v
    gannen = "gannen" in kinds
    twelve = "ampm" in kinds
    date_idx = [i for i, (k, _) in enumerate(toks) if k == "date"]
    out: list[str] = []
    i = 0
    while i < len(toks):
        k, s = toks[i]
        if k == "lit":
            out.append(s)
        elif k == "num":
            if s == "." and i > 0 and toks[i - 1][0] == "date" and toks[i - 1][1].startswith("s"):
                # 秒の小数（ss.00）
                j = i + 1
                while j < len(toks) and toks[j] == ("num", "0"):
                    j += 1
                n = j - i - 1
                if n:
                    out.append("." + f"{t.microsecond:06d}"[:n])
                    i = j
                    continue
                out.append(s)
            elif s in ".,/%":
                out.append(s)
            else:
                raise Unsupported("日付に数字の書式")
        elif k == "ampm":
            am = t.hour < 12
            out.append(("AM" if am else "PM") if s == "AM/PM" else ("A" if am else "P"))
        elif k == "date":
            c = s[0]
            if c in "ymdgea" and d is None and not (c == "m" and _is_minute(toks, date_idx, i)):
                raise Unsupported("時刻だけの値に日付の書式")
            if c == "y":
                out.append(f"{d.year % 100:02d}" if len(s) <= 2 else f"{d.year:04d}")
            elif c == "m" and _is_minute(toks, date_idx, i):
                out.append(f"{t.minute:02d}" if len(s) >= 2 else str(t.minute))
            elif c == "m":
                out.append(
                    {1: str(d.month), 2: f"{d.month:02d}", 3: _MONTHS[d.month - 1][:3], 4: _MONTHS[d.month - 1]}.get(
                        len(s), _MONTHS[d.month - 1][0]
                    )
                )
            elif c == "d":
                wd = _WEEKDAYS[d.weekday()]
                out.append({1: str(d.day), 2: f"{d.day:02d}", 3: wd[:3]}.get(len(s), wd))
            elif c == "a":
                out.append(_WEEKDAYS_JA[d.weekday()] + ("曜日" if len(s) >= 4 else ""))
            elif c == "h":
                h = (t.hour % 12 or 12) if twelve else t.hour
                out.append(f"{h:02d}" if len(s) >= 2 else str(h))
            elif c == "s":
                out.append(f"{t.second:02d}" if len(s) >= 2 else str(t.second))
            elif c == "e":
                year = _era(d)[0]
                out.append("元" if gannen and year == 1 else f"{year:02d}" if len(s) >= 2 else str(year))
            elif c == "g":
                _, name, short, roman = _era(d)
                out.append({1: roman, 2: short}.get(len(s), name))
        elif k in ("text", "general"):
            raise Unsupported("日付に文字列の書式")
        i += 1
    return "".join(out)


def _is_minute(toks: list[tuple[str, str]], date_idx: list[int], i: int) -> bool:
    """m / mm が「分」か: 直前の日付要素が時、または直後が秒."""
    if len(toks[i][1]) > 2:
        return False
    pos = date_idx.index(i)
    prev = toks[date_idx[pos - 1]][1] if pos > 0 else ""
    nxt = toks[date_idx[pos + 1]][1] if pos + 1 < len(date_idx) else ""
    return prev.startswith("h") or nxt.startswith("s")
