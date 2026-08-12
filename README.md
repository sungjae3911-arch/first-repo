# First Repo

git 연습을 위한 나의 첫 레포

[See Demo](https://www.google.com/)

## Prerequisites

- git
- python^3.13
- pandas==1.0.0

## How to start(Quickstart)

```shell
$ git clone {addr}
$ cd first-repo
$ python main.py
```
## Euler's Formula

e^{i*theta} = \\cos(theta) + i*sin(theta)

## Installation

There is no installation.

## Features

- Greeting
- Counting fizzbuzz

![](https://images.pexels.com/photos/104827/cat-pet-animal-domestic-104827.jpeg)

---

## 📊 엑셀 문서 비교 앱 (Excel Document Comparison)

서로 다른 두 엑셀 파일을 업로드하면 **시트별 · 셀 단위**로 차이를 비교해
색으로 하이라이트하여 보여주는 웹 앱입니다.

### 무엇을 감지하나요?

| 상태 | 의미 | 색상 |
|------|------|------|
| 변경됨 (changed) | 두 파일 모두에 값이 있지만 서로 다름 | 노랑 |
| 추가됨 (added) | 새 파일(B)에만 존재 | 초록 |
| 삭제됨 (removed) | 기존 파일(A)에만 존재 | 빨강 |

- 시트 단위 비교 (한쪽 파일에만 있는 시트도 표시)
- **3가지 비교 방식** (웹 앱에서 선택)
  1. **행 정렬 diff** (기본) — git/머지 툴처럼 행을 정렬(LCS)해 좌(A)·우(B)를 나란히 표시.
     중간에 행이 삽입/삭제되어도 줄이 알아서 밀려 맞춰집니다.
  2. **행 순서 무시** — 행 순서가 달라도 같은 내용의 행이면 동일로 처리
  3. **위치(행·열) 무시 · 내용 존재만 확인** — 값이 다른 행·열로 이동했더라도
     상대 파일 어딘가에 존재하면 "있음(공통)"으로 처리. 정말로 한쪽에만 있는 값만 차이로 표시
- **동일/다름 판정 배너** + `3`/`3.0`·앞뒤 공백 정규화로 오탐 방지
- **결과 엑셀 다운로드**: 바뀐 셀이 색으로 칠해지고 `이전값 → 새값` 메모가 달린
  `.xlsx` 를 만들어, 엑셀에서 바로 열어볼 수 있습니다.

### 실행 방법 (웹 앱)

```shell
$ pip install -r requirements.txt
$ python app.py
# 브라우저에서 http://127.0.0.1:5000 접속
```

`기존 파일(A)`와 `새 파일(B)`을 각각 드래그하거나 클릭하여 올린 뒤
`비교하기`로 화면에서 확인하거나, `결과 엑셀 다운로드`로 색칠된 엑셀을 받습니다.
(지원 형식: `.xlsx`, `.xls`, `.xlsm`)

### CLI 로 결과 엑셀 만들기

웹 없이 명령줄에서 바로 색칠된 diff 엑셀을 만들 수 있습니다.

```shell
$ python excel_diff.py 기존.xlsx 새.xlsx -o 비교결과.xlsx
```

결과 워크북 구성: `[요약]` 시트(전체/시트별 집계 + 동일/다름 판정) +
원본 시트별 좌우 diff 뷰(`~` 변경·노랑 / `-` 삭제·빨강 / `+` 추가·초록).

### 프로젝트 구조

```
app.py               # Flask 웹 서버 (/compare = 결과 JSON, /export = 색칠 엑셀)
comparator.py        # 엑셀 비교 핵심 로직 (셀 단위)
excel_diff.py        # 개발자 diff 뷰 엑셀 생성 (모듈 + CLI)
templates/index.html # 업로드 UI + 결과 화면 + 다운로드 버튼
test_comparator.py   # 비교 로직 테스트
test_excel_diff.py   # 결과 엑셀 생성 테스트
requirements.txt     # 의존성
```

### 테스트

```shell
$ python test_comparator.py
$ python test_excel_diff.py
```
