"""엑셀 문서 비교 핵심 로직.

두 개의 엑셀 파일(.xlsx/.xls)을 시트별로 읽어 셀 단위로 비교한다.
비교 결과는 아래 4가지 상태로 분류한다.

- same     : 두 파일의 값이 동일
- changed  : 두 파일 모두에 존재하지만 값이 다름
- added    : 새 파일(B)에만 존재 (추가됨)
- removed  : 기존 파일(A)에만 존재 (삭제됨)

pandas / openpyxl 에 의존한다.
"""

from __future__ import annotations

import math
from dataclasses import dataclass, field

import pandas as pd


def _read_sheets(path: str) -> dict[str, pd.DataFrame]:
    """엑셀 파일을 읽어 {시트이름: DataFrame} 형태로 반환한다.

    헤더를 별도로 해석하지 않고 모든 셀을 문자열 격자로 다룬다.
    이렇게 하면 헤더 위치가 다르거나 병합된 셀이 있어도 안전하게 비교할 수 있다.
    """
    excel = pd.read_excel(path, sheet_name=None, header=None, dtype=object)
    return excel


def _normalize(value: object) -> str:
    """셀 값을 비교 가능한 문자열로 정규화한다.

    - 결측치(NaN/None)는 빈 문자열로 처리
    - 정수처럼 표현되는 실수(3.0)는 정수 문자열(3)로 통일
    - 앞뒤 공백은 제거
    """
    if value is None:
        return ""
    if isinstance(value, float):
        if math.isnan(value):
            return ""
        if value.is_integer():
            return str(int(value))
        return str(value)
    return str(value).strip()


@dataclass
class CellDiff:
    row: int
    col: int
    status: str  # same | changed | added | removed
    old: str
    new: str


@dataclass
class SheetDiff:
    name: str
    rows: int
    cols: int
    cells: list[CellDiff] = field(default_factory=list)
    only_in: str | None = None  # 'A' 또는 'B' — 한쪽에만 존재하는 시트일 때

    @property
    def counts(self) -> dict[str, int]:
        c = {"same": 0, "changed": 0, "added": 0, "removed": 0}
        for cell in self.cells:
            c[cell.status] += 1
        return c


def _compare_sheet(name: str, df_a: pd.DataFrame | None, df_b: pd.DataFrame | None) -> SheetDiff:
    only_in = None
    if df_a is None:
        only_in = "B"
        df_a = pd.DataFrame()
    elif df_b is None:
        only_in = "A"
        df_b = pd.DataFrame()

    rows = max(df_a.shape[0], df_b.shape[0])
    cols = max(df_a.shape[1], df_b.shape[1])

    sheet = SheetDiff(name=name, rows=rows, cols=cols, only_in=only_in)

    for r in range(rows):
        for c in range(cols):
            a_val = _normalize(df_a.iat[r, c]) if r < df_a.shape[0] and c < df_a.shape[1] else ""
            b_val = _normalize(df_b.iat[r, c]) if r < df_b.shape[0] and c < df_b.shape[1] else ""

            if a_val == "" and b_val == "":
                continue  # 양쪽 모두 빈 셀은 건너뛴다

            if a_val == b_val:
                status = "same"
            elif a_val == "":
                status = "added"
            elif b_val == "":
                status = "removed"
            else:
                status = "changed"

            sheet.cells.append(CellDiff(row=r, col=c, status=status, old=a_val, new=b_val))

    return sheet


def compare_files(path_a: str, path_b: str) -> dict:
    """두 엑셀 파일을 비교하여 결과를 dict(JSON 직렬화 가능)로 반환한다."""
    sheets_a = _read_sheets(path_a)
    sheets_b = _read_sheets(path_b)

    all_names = list(sheets_a.keys())
    for name in sheets_b.keys():
        if name not in all_names:
            all_names.append(name)

    sheet_results = []
    total = {"same": 0, "changed": 0, "added": 0, "removed": 0}

    for name in all_names:
        diff = _compare_sheet(name, sheets_a.get(name), sheets_b.get(name))
        counts = diff.counts
        for k in total:
            total[k] += counts[k]

        sheet_results.append(
            {
                "name": diff.name,
                "rows": diff.rows,
                "cols": diff.cols,
                "only_in": diff.only_in,
                "counts": counts,
                "cells": [
                    {
                        "row": cell.row,
                        "col": cell.col,
                        "status": cell.status,
                        "old": cell.old,
                        "new": cell.new,
                    }
                    for cell in diff.cells
                ],
            }
        )

    return {
        "sheets": sheet_results,
        "total": total,
        "changed": total["changed"] + total["added"] + total["removed"] > 0,
    }


def column_label(index: int) -> str:
    """0-based 컬럼 인덱스를 엑셀식 열 문자(A, B, ..., Z, AA, ...)로 변환한다."""
    label = ""
    index += 1
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        label = chr(65 + remainder) + label
    return label
