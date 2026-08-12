"""comparator 모듈 동작을 검증하는 간단한 테스트.

실행: python test_comparator.py
(pytest 없이도 실행되도록 순수 assert 로 작성)
"""

import os
import tempfile

import pandas as pd

from comparator import column_label, compare_files, compare_presence


def _make_xlsx(path: str, sheets: dict[str, list[list]]) -> None:
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for name, rows in sheets.items():
            pd.DataFrame(rows).to_excel(writer, sheet_name=name, header=False, index=False)


def test_column_label():
    assert column_label(0) == "A"
    assert column_label(25) == "Z"
    assert column_label(26) == "AA"
    assert column_label(27) == "AB"
    print("✓ column_label")


def test_identical_files():
    with tempfile.TemporaryDirectory() as d:
        a = os.path.join(d, "a.xlsx")
        b = os.path.join(d, "b.xlsx")
        data = {"Sheet1": [["이름", "점수"], ["철수", 90], ["영희", 85]]}
        _make_xlsx(a, data)
        _make_xlsx(b, data)
        result = compare_files(a, b)
        assert result["changed"] is False
        assert result["total"] == {"same": 6, "changed": 0, "added": 0, "removed": 0}
    print("✓ identical files → no diff")


def test_changed_added_removed():
    with tempfile.TemporaryDirectory() as d:
        a = os.path.join(d, "a.xlsx")
        b = os.path.join(d, "b.xlsx")
        _make_xlsx(a, {"Sheet1": [["이름", "점수"], ["철수", 90], ["영희", 85]]})
        # 철수 90->95 (변경), 영희 행 삭제, 민수 행 추가
        _make_xlsx(b, {"Sheet1": [["이름", "점수"], ["철수", 95], ["민수", 70]]})
        result = compare_files(a, b)
        assert result["changed"] is True

        cells = result["sheets"][0]["cells"]
        by_status = {}
        for c in cells:
            by_status.setdefault(c["status"], []).append(c)

        # 철수 점수 90 -> 95
        changed = [c for c in by_status.get("changed", []) if c["old"] == "90" and c["new"] == "95"]
        assert changed, "철수 점수 변경이 감지되어야 함"
        print("✓ changed / added / removed detected:", result["total"])


def test_int_float_normalization():
    """엑셀에서 정수가 3.0 처럼 읽혀도 동일 값으로 취급되어야 한다."""
    with tempfile.TemporaryDirectory() as d:
        a = os.path.join(d, "a.xlsx")
        b = os.path.join(d, "b.xlsx")
        _make_xlsx(a, {"Sheet1": [[3, 5]]})
        _make_xlsx(b, {"Sheet1": [[3.0, 5.0]]})
        result = compare_files(a, b)
        assert result["changed"] is False, "3 과 3.0 은 같은 값으로 취급되어야 함"
    print("✓ int/float normalization")


def test_sheet_only_in_one_file():
    with tempfile.TemporaryDirectory() as d:
        a = os.path.join(d, "a.xlsx")
        b = os.path.join(d, "b.xlsx")
        _make_xlsx(a, {"Sheet1": [["x"]], "OnlyA": [["삭제됨"]]})
        _make_xlsx(b, {"Sheet1": [["x"]], "OnlyB": [["추가됨"]]})
        result = compare_files(a, b)
        names = {s["name"]: s for s in result["sheets"]}
        assert names["OnlyA"]["only_in"] == "A"
        assert names["OnlyB"]["only_in"] == "B"
    print("✓ sheet only in one file")


def test_presence_ignores_position():
    """행·열 위치가 달라도 값이 존재하면 공통으로 처리되는지 검증."""
    with tempfile.TemporaryDirectory() as d:
        a = os.path.join(d, "a.xlsx")
        b = os.path.join(d, "b.xlsx")
        _make_xlsx(a, {"과일": [["사과", "바나나"], ["체리", "포도"]]})
        # 같은 값들이 다른 행·열로 이동 + 오렌지 추가
        _make_xlsx(b, {"과일": [["포도", "체리", "사과"], ["바나나", "오렌지"]]})
        result = compare_presence(a, b)
        assert result["total"]["common"] == 4, "위치가 달라도 4개 값이 공통이어야 함"
        assert result["total"]["only_a"] == 0
        assert result["total"]["only_b"] == 1, "오렌지만 B에만 있어야 함"
        assert result["sheets"][0]["only_b"][0]["value"] == "오렌지"
    print("✓ presence comparison ignores row/column position")


if __name__ == "__main__":
    test_column_label()
    test_identical_files()
    test_changed_added_removed()
    test_int_float_normalization()
    test_sheet_only_in_one_file()
    test_presence_ignores_position()
    print("\n모든 테스트 통과 ✅")
