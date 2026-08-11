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
- 셀 좌표(예: `B3`)와 기존 값 → 새 값을 표 형태로 제공
- `3` 과 `3.0`, 앞뒤 공백 등은 같은 값으로 정규화하여 오탐 방지

### 실행 방법

```shell
$ pip install -r requirements.txt
$ python app.py
# 브라우저에서 http://127.0.0.1:5000 접속
```

`기존 파일(A)`와 `새 파일(B)`을 각각 드래그하거나 클릭하여 올린 뒤
`비교하기` 버튼을 누르면 결과가 표시됩니다. (지원 형식: `.xlsx`, `.xls`, `.xlsm`)

### 프로젝트 구조

```
app.py               # Flask 웹 서버 (업로드 → 비교 → 결과 JSON)
comparator.py        # 엑셀 비교 핵심 로직
templates/index.html # 업로드 UI + 결과 화면
test_comparator.py   # 비교 로직 테스트 (python test_comparator.py)
requirements.txt     # 의존성
```

### 테스트

```shell
$ python test_comparator.py
```
