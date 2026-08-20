# scourt-disclose

대한민국 법원 공시송달(scourt.go.kr) 조회 프로그램. 선택한 법원(및 산하 지원/시군법원)의 공시송달 목록을 조회한 뒤, `keywords.txt`에 등록한 회사명이 포함된 공시대상자만 걸러서 엑셀 파일로 저장한다.

## 동작 방식

1. `keywords.txt`에서 필터링할 회사명 목록을 읽는다. (없으면 예시 템플릿을 자동 생성)
2. 법원 목록(XML)을 불러온다. (24시간 캐시)
3. 사용자가 입력한 법원명으로 대상 법원을 찾고, 산하 지원/시군법원까지 함께 조회 대상에 포함한다.
4. 게시기간 0~21일 전체 페이지를 순회하며 공시송달 목록을 수집한다. (법원 단위 6시간 캐시)
5. 공시대상자 문구에 키워드가 포함된 건만 추려서 `공시송달_법원명_날짜시간.xlsx`로 저장한다.

## 요구사항

- Python 3.11
- `requirements.txt`: `requests`, `openpyxl`, `pyinstaller`

## 로컬 빌드 (Windows)

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe -m PyInstaller --onefile --console --name scourt_disclose scourt_disclose.py
```

빌드 결과물은 `dist/scourt_disclose.exe`. 실행 시 같은 폴더에 `keywords.txt`, `cache/` 가 함께 생성/사용된다.

## CI 빌드 (GitHub Actions)

`.github/workflows/build.yml`이 push / 수동 실행(`workflow_dispatch`) 시 `windows-latest`에서 위와 동일한 순서로 빌드하고, 결과 exe를 `scourt_disclose_windows` 아티팩트로 업로드한다.

## 실행

```
scourt_disclose.exe [--court 법원명]
```

`--court`를 생략하면 실행 중에 법원명을 입력받는다(부분 일치 검색, `목록` 입력 시 전체 법원 확인 가능).
