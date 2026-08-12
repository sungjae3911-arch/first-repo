"""두 엑셀 문서의 차이를 '개발자 diff 뷰처럼' 색칠된 엑셀 파일로 내보낸다.

git diff / 머지 툴(예: VS Code, Beyond Compare)이 소스코드를 비교하듯,
행을 정렬(LCS)해 좌우로 나란히 놓고 색으로 표시한다.

  · 왼쪽 블록  = 기존 파일 (A)
  · 오른쪽 블록 = 새 파일 (B)
  · 가운데 Δ열 = 각 행의 상태 기호

색상/기호
  · 흰색        = 동일한 행
  · 노랑  ~     = 변경된 행 (달라진 셀만 노랑 강조)
  · 빨강  -     = 삭제된 행 (A에만 있음)
  · 초록  +     = 추가된 행 (B에만 있음)

행 정렬에는 표준 라이브러리 difflib.SequenceMatcher 를 사용한다.
이 덕분에 중간에 행이 삽입/삭제되어도 diff 도구처럼 줄이 알아서 밀려 맞춰진다.

사용법(CLI):
    python excel_diff.py 기존.xlsx 새.xlsx -o 결과.xlsx
"""

from __future__ import annotations

import argparse
import difflib

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from comparator import _normalize, _read_sheets

# 색상 (ARGB)
FILL_CHANGED = PatternFill("solid", fgColor="FFF2CC")  # 노랑
FILL_ADDED = PatternFill("solid", fgColor="C6EFCE")    # 초록
FILL_REMOVED = PatternFill("solid", fgColor="FFC7CE")  # 빨강
FILL_HEADER = PatternFill("solid", fgColor="4F46E5")   # 인디고
FILL_GUTTER = PatternFill("solid", fgColor="F2F2F2")

FONT_HEADER = Font(color="FFFFFF", bold=True)
FONT_REMOVED = Font(color="9C0006")
FONT_ADDED = Font(color="006100")
FONT_CHANGED = Font(color="7F6000")
FONT_MUTED = Font(color="BFBFBF")

THIN = Side(style="thin", color="E0E0E0")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal="center", vertical="center")


def _grid_rows(df, cols: int) -> list[list[str]]:
    """DataFrame 을 cols 폭에 맞춘 정규화 문자열 2차원 리스트로 변환."""
    rows = []
    n = df.shape[0] if df is not None else 0
    for r in range(n):
        row = []
        for c in range(cols):
            if df is not None and c < df.shape[1]:
                row.append(_normalize(df.iat[r, c]))
            else:
                row.append("")
        rows.append(row)
    return rows


def _sheet_cols(df_a, df_b) -> int:
    a = df_a.shape[1] if df_a is not None else 0
    b = df_b.shape[1] if df_b is not None else 0
    return max(a, b, 1)


def export_diff_xlsx(path_a: str, path_b: str, out_path: str) -> dict:
    """두 엑셀을 좌우 diff 뷰로 비교해 결과 워크북을 저장하고 집계를 반환한다."""
    sheets_a = _read_sheets(path_a)
    sheets_b = _read_sheets(path_b)

    names = list(sheets_a.keys())
    for name in sheets_b.keys():
        if name not in names:
            names.append(name)

    wb = Workbook()
    summary_ws = wb.active
    summary_ws.title = "요약"

    total = {"same": 0, "changed": 0, "added": 0, "removed": 0}
    per_sheet: list[tuple[str, dict, str | None]] = []

    for name in names:
        df_a = sheets_a.get(name)
        df_b = sheets_b.get(name)
        only_in = "B" if df_a is None else ("A" if df_b is None else None)

        ws = wb.create_sheet(title=_safe_sheet_title(name, wb))
        counts = _render_side_by_side(ws, df_a, df_b)

        for k in total:
            total[k] += counts[k]
        per_sheet.append((name, counts, only_in))

    _write_summary(summary_ws, total, per_sheet, path_a, path_b)
    wb.save(out_path)
    return {"total": total, "sheets": [{"name": n, "counts": c, "only_in": o} for n, c, o in per_sheet]}


def _align_replace(a_rows: list[list[str]], b_rows: list[list[str]]):
    """diff 의 replace 구간을 첫 열(식별자)로 2차 정렬해 세부 op 로 나눈다.

    yield ("change", a_row, b_row) | ("del", a_row, None) | ("ins", None, b_row)

    첫 열 값이 같은 행끼리 '변경'으로 짝짓고, 짝이 없는 행은 삭제/추가로 본다.
    diff 도구가 비슷한 줄을 정렬해 보여주는 것과 같은 효과를 낸다.
    """
    anchors_a = [r[0] for r in a_rows]
    anchors_b = [r[0] for r in b_rows]
    sm = difflib.SequenceMatcher(a=anchors_a, b=anchors_b, autojunk=False)
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            for k in range(i2 - i1):
                yield ("change", a_rows[i1 + k], b_rows[j1 + k])
        elif tag == "delete":
            for k in range(i1, i2):
                yield ("del", a_rows[k], None)
        elif tag == "insert":
            for k in range(j1, j2):
                yield ("ins", None, b_rows[k])
        else:  # replace — 첫 열도 다르면 index 로 짝짓고 나머지는 삭제/추가
            la, lb = i2 - i1, j2 - j1
            for k in range(min(la, lb)):
                yield ("change", a_rows[i1 + k], b_rows[j1 + k])
            for k in range(min(la, lb), la):
                yield ("del", a_rows[i1 + k], None)
            for k in range(min(la, lb), lb):
                yield ("ins", None, b_rows[j1 + k])


def _render_side_by_side(ws, df_a, df_b) -> dict:
    """한 시트를 좌(A)·우(B) diff 뷰로 그린다. 셀 단위 집계를 반환."""
    cols = _sheet_cols(df_a, df_b)
    rows_a = _grid_rows(df_a, cols)
    rows_b = _grid_rows(df_b, cols)
    keys_a = ["\x01".join(r) for r in rows_a]
    keys_b = ["\x01".join(r) for r in rows_b]

    left0 = 1                 # 왼쪽 블록 시작 열
    gutter = cols + 1         # 가운데 Δ열
    right0 = cols + 2         # 오른쪽 블록 시작 열

    # 헤더 (1행)
    ws.cell(row=1, column=left0, value="기존 파일 (A)").font = Font(bold=True, color="9C0006")
    ws.cell(row=1, column=gutter, value="Δ").font = Font(bold=True)
    ws.cell(row=1, column=right0, value="새 파일 (B)").font = Font(bold=True, color="006100")
    if cols > 1:
        ws.merge_cells(start_row=1, start_column=left0, end_row=1, end_column=left0 + cols - 1)
        ws.merge_cells(start_row=1, start_column=right0, end_row=1, end_column=right0 + cols - 1)
    for cc in (left0, gutter, right0):
        ws.cell(row=1, column=cc).alignment = CENTER
    ws.freeze_panes = "A2"

    counts = {"same": 0, "changed": 0, "added": 0, "removed": 0}
    xr = 2  # 현재 엑셀 출력 행

    def put(row_idx, start_col, values, fill, font):
        for i, val in enumerate(values):
            cell = ws.cell(row=row_idx, column=start_col + i)
            cell.value = val if val != "" else None
            cell.border = BORDER
            if fill is not None:
                cell.fill = fill
            if font is not None:
                cell.font = font

    def gutter_mark(row_idx, sym, fill):
        cell = ws.cell(row=row_idx, column=gutter, value=sym)
        cell.alignment = CENTER
        cell.fill = fill
        cell.font = Font(bold=True)

    def count_nonempty(row):
        return sum(1 for v in row if v != "")

    sm = difflib.SequenceMatcher(a=keys_a, b=keys_b, autojunk=False)
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            for k in range(i2 - i1):
                put(xr, left0, rows_a[i1 + k], None, None)
                put(xr, right0, rows_b[j1 + k], None, None)
                counts["same"] += count_nonempty(rows_a[i1 + k])
                xr += 1
        elif tag == "delete":
            for k in range(i1, i2):
                put(xr, left0, rows_a[k], FILL_REMOVED, FONT_REMOVED)
                gutter_mark(xr, "-", FILL_REMOVED)
                counts["removed"] += count_nonempty(rows_a[k])
                xr += 1
        elif tag == "insert":
            for k in range(j1, j2):
                put(xr, right0, rows_b[k], FILL_ADDED, FONT_ADDED)
                gutter_mark(xr, "+", FILL_ADDED)
                counts["added"] += count_nonempty(rows_b[k])
                xr += 1
        elif tag == "replace":
            # 변경 구간 안에서 첫 열(이름/ID)을 기준으로 2차 정렬해
            # 삽입(+)·삭제(-)·변경(~)을 정확히 구분한다 (diff 도구와 동일한 방식).
            for op, a_row, b_row in _align_replace(rows_a[i1:i2], rows_b[j1:j2]):
                if op == "change":
                    for i in range(cols):
                        differ = a_row[i] != b_row[i]
                        lc = ws.cell(row=xr, column=left0 + i)
                        lc.value = a_row[i] or None
                        lc.border = BORDER
                        rc = ws.cell(row=xr, column=right0 + i)
                        rc.value = b_row[i] or None
                        rc.border = BORDER
                        if differ:
                            lc.fill = FILL_CHANGED
                            lc.font = FONT_CHANGED
                            rc.fill = FILL_CHANGED
                            rc.font = FONT_CHANGED
                            if a_row[i] and b_row[i]:
                                counts["changed"] += 1
                            elif b_row[i]:
                                counts["added"] += 1
                            else:
                                counts["removed"] += 1
                        elif a_row[i] != "":
                            counts["same"] += 1
                    gutter_mark(xr, "~", FILL_CHANGED)
                elif op == "del":
                    put(xr, left0, a_row, FILL_REMOVED, FONT_REMOVED)
                    gutter_mark(xr, "-", FILL_REMOVED)
                    counts["removed"] += count_nonempty(a_row)
                else:  # ins
                    put(xr, right0, b_row, FILL_ADDED, FONT_ADDED)
                    gutter_mark(xr, "+", FILL_ADDED)
                    counts["added"] += count_nonempty(b_row)
                xr += 1

    _autofit(ws, cols, gutter, right0)
    return counts


def _safe_sheet_title(name: str, wb: Workbook) -> str:
    for ch in r"[]:*?/\\":
        name = name.replace(ch, "_")
    name = name[:28]
    base, i, title = name or "Sheet", 1, name or "Sheet"
    while title in wb.sheetnames:
        title = f"{base[:25]}_{i}"
        i += 1
    return title


def _autofit(ws, cols: int, gutter: int, right0: int) -> None:
    ws.column_dimensions[get_column_letter(gutter)].width = 4
    for block_start in (1, right0):
        for c in range(block_start, block_start + cols):
            letter = get_column_letter(c)
            width = 10
            for cell in ws[letter]:
                if cell.value is not None:
                    width = max(width, min(45, len(str(cell.value)) + 2))
            ws.column_dimensions[letter].width = width


def _write_summary(ws, total: dict, per_sheet: list, path_a: str, path_b: str) -> None:
    ws["A1"] = "엑셀 문서 비교 결과 (개발자 diff 뷰)"
    ws["A1"].font = Font(size=14, bold=True)
    ws["A2"] = f"기존 파일(A): {path_a}"
    ws["A3"] = f"새 파일(B): {path_b}"
    for r in (2, 3):
        ws[f"A{r}"].font = Font(color="666666")

    diff_n = total["changed"] + total["added"] + total["removed"]
    same_all = diff_n == 0
    ws["A5"] = "판정"
    ws["B5"] = "✅ 두 파일의 데이터가 완전히 동일합니다" if same_all else f"⚠️ 데이터가 다릅니다 — 차이 {diff_n}곳"
    ws["A5"].font = Font(bold=True)
    ws["B5"].font = Font(bold=True, color="006100" if same_all else "9C5700")

    ws["A7"] = "각 시트 탭에서 좌(기존 A) · 우(새 B)를 나란히 비교할 수 있습니다."
    ws["A8"] = "기호: ~ 변경(노랑) · - 삭제(빨강) · + 추가(초록)"
    for r in (7, 8):
        ws[f"A{r}"].font = Font(color="666666")

    header = ["시트", "변경", "추가", "삭제", "비고"]
    row0 = 10
    for j, h in enumerate(header, start=1):
        cell = ws.cell(row=row0, column=j, value=h)
        cell.fill = FILL_HEADER
        cell.font = FONT_HEADER
        cell.alignment = CENTER

    r = row0 + 1
    for name, counts, only_in in per_sheet:
        note = ""
        if only_in == "A":
            note = "기존 파일에만 있음 (삭제됨)"
        elif only_in == "B":
            note = "새 파일에만 있음 (추가됨)"
        ws.cell(row=r, column=1, value=name)
        ws.cell(row=r, column=2, value=counts["changed"])
        ws.cell(row=r, column=3, value=counts["added"])
        ws.cell(row=r, column=4, value=counts["removed"])
        ws.cell(row=r, column=5, value=note)
        r += 1

    ws.cell(row=r, column=1, value="합계").font = Font(bold=True)
    for col, key in ((2, "changed"), (3, "added"), (4, "removed")):
        ws.cell(row=r, column=col, value=total[key]).font = Font(bold=True)

    ws.column_dimensions["A"].width = 26
    for col in ("B", "C", "D"):
        ws.column_dimensions[col].width = 8
    ws.column_dimensions["E"].width = 30


def _value_set(df) -> set[str]:
    """시트의 비어있지 않은 셀 값 집합."""
    values: set[str] = set()
    if df is None:
        return values
    for r in range(df.shape[0]):
        for c in range(df.shape[1]):
            v = _normalize(df.iat[r, c])
            if v != "":
                values.add(v)
    return values


def export_presence_xlsx(path_a: str, path_b: str, out_path: str) -> dict:
    """위치(행·열)를 무시하고 '있음/없음'을 색으로 칠한 엑셀을 만든다.

    두 파일을 좌(A)·우(B)로 나란히 두고, 각 셀 값이 상대 파일에 존재하면
    초록, 없으면 빨강으로 칠한다.
    """
    sheets_a = _read_sheets(path_a)
    sheets_b = _read_sheets(path_b)

    names = list(sheets_a.keys())
    for name in sheets_b.keys():
        if name not in names:
            names.append(name)

    wb = Workbook()
    summary_ws = wb.active
    summary_ws.title = "요약"

    total = {"green": 0, "red": 0}
    per_sheet = []

    for name in names:
        df_a, df_b = sheets_a.get(name), sheets_b.get(name)
        only_in = "B" if df_a is None else ("A" if df_b is None else None)
        cols_a = df_a.shape[1] if df_a is not None else 0
        cols_b = df_b.shape[1] if df_b is not None else 0
        rows_a = _grid_rows(df_a, cols_a)
        rows_b = _grid_rows(df_b, cols_b)
        set_a, set_b = _value_set(df_a), _value_set(df_b)

        ws = wb.create_sheet(title=_safe_sheet_title(name, wb))
        ca, cb = max(cols_a, 1), max(cols_b, 1)
        left0, right0 = 1, ca + 2  # ca+1 열은 좌우 사이 여백

        ws.cell(row=1, column=left0, value="기존 파일 (A)").font = Font(bold=True)
        ws.cell(row=1, column=right0, value="새 파일 (B)").font = Font(bold=True)
        if ca > 1:
            ws.merge_cells(start_row=1, start_column=left0, end_row=1, end_column=left0 + ca - 1)
        if cb > 1:
            ws.merge_cells(start_row=1, start_column=right0, end_row=1, end_column=right0 + cb - 1)
        ws.cell(row=1, column=left0).alignment = CENTER
        ws.cell(row=1, column=right0).alignment = CENTER
        ws.freeze_panes = "A2"

        green = red = 0

        def paint(excel_row, start_col, row, cols, other_set):
            nonlocal green, red
            for c in range(cols):
                val = row[c] if c < len(row) else ""
                if val == "":
                    continue
                cell = ws.cell(row=excel_row, column=start_col + c, value=val)
                cell.border = BORDER
                if val in other_set:
                    cell.fill = FILL_ADDED
                    cell.font = FONT_ADDED
                    green += 1
                else:
                    cell.fill = FILL_REMOVED
                    cell.font = FONT_REMOVED
                    red += 1

        for i in range(max(len(rows_a), len(rows_b))):
            xr = i + 2
            if i < len(rows_a):
                paint(xr, left0, rows_a[i], ca, set_b)
            if i < len(rows_b):
                paint(xr, right0, rows_b[i], cb, set_a)

        for c in range(1, right0 + cb):
            ws.column_dimensions[get_column_letter(c)].width = 3 if c == ca + 1 else 12

        total["green"] += green
        total["red"] += red
        per_sheet.append((name, green, red, only_in))

    _write_presence_summary(summary_ws, total, per_sheet, path_a, path_b)
    wb.save(out_path)
    return {"total": total, "sheets": [{"name": n, "green": g, "red": r, "only_in": o} for n, g, r, o in per_sheet]}


def _write_presence_summary(ws, total: dict, per_sheet: list, path_a: str, path_b: str) -> None:
    ws["A1"] = "엑셀 문서 비교 결과 (위치 무시 · 있음/없음)"
    ws["A1"].font = Font(size=14, bold=True)
    ws["A2"] = f"기존 파일(A): {path_a}"
    ws["A3"] = f"새 파일(B): {path_b}"
    for r in (2, 3):
        ws[f"A{r}"].font = Font(color="666666")

    ok = total["red"] == 0
    ws["A5"] = "판정"
    ws["B5"] = "✅ 모든 내용이 상대 파일에도 존재합니다" if ok else f"⚠️ 상대 파일에 없는 값 {total['red']}곳"
    ws["A5"].font = Font(bold=True)
    ws["B5"].font = Font(bold=True, color="006100" if ok else "9C5700")
    ws["A7"] = "색: 초록 = 상대 파일에 있음 · 빨강 = 상대 파일에 없음"
    ws["A7"].font = Font(color="666666")

    header = ["시트", "있음(초록)", "없음(빨강)", "비고"]
    row0 = 9
    for j, h in enumerate(header, start=1):
        cell = ws.cell(row=row0, column=j, value=h)
        cell.fill = FILL_HEADER
        cell.font = FONT_HEADER
        cell.alignment = CENTER

    r = row0 + 1
    for name, green, red, only_in in per_sheet:
        note = "기존 파일에만 있음" if only_in == "A" else ("새 파일에만 있음" if only_in == "B" else "")
        ws.cell(row=r, column=1, value=name)
        ws.cell(row=r, column=2, value=green)
        ws.cell(row=r, column=3, value=red)
        ws.cell(row=r, column=4, value=note)
        r += 1

    ws.cell(row=r, column=1, value="합계").font = Font(bold=True)
    ws.cell(row=r, column=2, value=total["green"]).font = Font(bold=True)
    ws.cell(row=r, column=3, value=total["red"]).font = Font(bold=True)
    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["D"].width = 24


def main() -> None:
    parser = argparse.ArgumentParser(description="두 엑셀 파일의 차이를 색칠된 엑셀로 내보냅니다.")
    parser.add_argument("file_a", help="기존 파일 (A)")
    parser.add_argument("file_b", help="새 파일 (B)")
    parser.add_argument("-o", "--out", default="비교결과.xlsx", help="결과 파일 경로 (기본: 비교결과.xlsx)")
    parser.add_argument(
        "-m", "--mode", choices=["diff", "presence"], default="diff",
        help="diff=개발자 diff 뷰(기본) · presence=위치 무시 있음/없음(초록/빨강)",
    )
    args = parser.parse_args()

    if args.mode == "presence":
        result = export_presence_xlsx(args.file_a, args.file_b, args.out)
        t = result["total"]
        print(f"결과 저장: {args.out}")
        print(f"  있음(초록) {t['green']} · 없음(빨강) {t['red']}")
    else:
        result = export_diff_xlsx(args.file_a, args.file_b, args.out)
        t = result["total"]
        print(f"결과 저장: {args.out}")
        print(f"  변경 {t['changed']} · 추가 {t['added']} · 삭제 {t['removed']}")


if __name__ == "__main__":
    main()
