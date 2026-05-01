# SQLite 자동화 정리

이 프로젝트는 월별 엑셀 파일을 원본 그대로 유지하면서 SQLite DB로 적재하는 자동화 도구입니다.  
현재 기본 설정은 `수축률` 파트용이며, 다른 파트도 `config` 파일만 추가해서 같은 프로그램 구조로 확장할 수 있습니다.

## 현재 구조

- Python 소스는 `src` 폴더에 있습니다.
- 파트별 설정은 `config/parts` 폴더에서 관리합니다.
- 배치 파일은 루트에서 실행합니다.
- 원본 엑셀은 수정하거나 삭제하지 않습니다.
- 적재 대상 DB는 연도별 파일 1개로 생성합니다.

## 폴더 설명

- `src/import_excel_to_sqlite.py`
  - 실제 엑셀 → SQLite 적재 로직
- `src/run_part_import.py`
  - 설정 파일을 읽어 적재를 실행하고 로그를 남기는 실행기
- `config/parts/shrinkage.json`
  - 로컬 실사용 수축률 파트 설정
- `config/parts/shrinkage.example.json`
  - 공개 저장소용 예시 설정
- `run_import_single_db_shrinkage_status.bat`
  - 수축률 파트용 실사용 배치 파일
- `debug_run_shrinkage_status_import.bat`
  - 수축률 파트용 진단 배치 파일
- `run_import_with_config.bat`
  - 다른 파트 설정 파일을 넘겨 실행할 수 있는 공용 배치 파일
- `logs/`
  - 실행 로그 저장 폴더

## 현재 기본 동작

- 원본 경로는 설정 파일에서 읽습니다.
- 연도 폴더는 실행 시 자동으로 붙습니다.
  - 예: `...\start\2026`
- DB 파일명도 설정에서 관리합니다.
  - 예: `접수현황_수축률DB_2026.db`
- 메인 테이블은 `receipt_status` 입니다.

## 파트 관리 방식

현재는 `receipt_status` 테이블에 아래 컬럼이 자동으로 들어갑니다.

- `part_code`
- `part_name`
- `part_description`

즉 `수축률` 데이터는 적재 시 자동으로 아래 값이 들어갑니다.

- `part_code = shrinkage`
- `part_name = 수축률`
- `part_description = 가공성능평가팀 수축률 파트 접수 현황 데이터`

이 구조라서 다른 파트와 `접수번호`가 겹쳐도 함께 적재할 수 있습니다.  
`receipt_no` 단독 고유 제약은 두지 않았고, 조회 성능을 위해 `(part_code, receipt_no)` 인덱스를 추가했습니다.

## 다른 파트로 확장하는 방법

다른 파트를 추가할 때는 `config/parts` 아래에 새 JSON 파일만 만들면 됩니다.

예시:

```json
{
  "part_code": "ph",
  "part_name": "pH",
  "part_description": "pH 파트 접수 현황 데이터",
  "source_base_dirs": [
    "\\\\서버\\경로\\start"
  ],
  "target_dir": "\\\\서버\\경로\\DB\\접수현황_pHDB",
  "sheet_name": "접수내역",
  "table_name": "receipt_status",
  "db_mode": "single_db",
  "table_mode": "fixed",
  "pattern": "*.xlsx",
  "header_row": 3,
  "data_start_row": 4,
  "date_column": "접수일자",
  "append_year_to_source_dir": true,
  "single_db_name_template": "접수현황_pHDB_{year}.db"
}
```

그 다음 이렇게 실행하면 됩니다.

```bat
run_import_with_config.bat config\parts\ph.json 2026
```

공개 저장소에서는 실제 사내 경로가 들어간 `config/parts/shrinkage.json` 대신 `config/parts/shrinkage.example.json`만 포함하고, 로컬에서 복사해서 사용하면 됩니다.

## 실행 방법

수축률 파트 실사용:

```bat
run_import_single_db_shrinkage_status.bat 2026
```

현재 연도를 자동 사용:

```bat
run_import_single_db_shrinkage_status.bat
```

진단 실행:

```bat
debug_run_shrinkage_status_import.bat 2026
```

## 샘플 검증 예시

로컬 `sample` 폴더로 테스트할 때:

```powershell
python src\run_part_import.py --config config\parts\shrinkage.json --source-dir . --target-dir output_db_standard --year sample --single-db-name shrinkage_2026_standard.db --replace-db
```

## DB 구조

### 메인 테이블: `receipt_status`

주요 컬럼:

- `part_code`
- `part_name`
- `part_description`
- `receipt_no`
- `test_item`
- `department`
- `client_name`
- `received_date`
- `due_date`
- `status`
- `source_file`
- `source_sheet`
- `source_row`
- `created_at`
- `updated_at`

### 파일 입력 로그 테이블: `file_import_logs`

주요 컬럼:

- `part_code`
- `part_name`
- `file_name`
- `file_path`
- `file_hash`
- `imported_at`
- `total_rows`
- `inserted_rows`
- `failed_rows`
- `status`

### 적재 메타 테이블: `import_metadata`

주요 컬럼:

- `part_code`
- `part_name`
- `source_file`
- `source_sheet`
- `table_name`
- `imported_at_utc`
- `row_count`

## 로그

- 실행 로그는 `logs` 폴더에 저장됩니다.
- 파일명은 `설정파일명_연도_날짜시간.log` 형식입니다.
- 실패 시 가장 먼저 `logs` 폴더의 최신 파일을 확인하면 됩니다.

## Python 준비

배치 파일은 실제 Python이 설치되어 있어야 합니다.

권장:

- Python 3.11 이상 설치
- 설치 시 `Add python.exe to PATH` 체크

직접 경로를 지정해야 하면:

```bat
set PYTHON_EXE=C:\Users\사용자명\AppData\Local\Programs\Python\Python313\python.exe
```

## VS Code에서 보기

- 이 폴더를 VS Code로 엽니다.
- 추천 확장 `SQLite Viewer` 또는 `SQLTools`를 설치합니다.
- 생성된 `.db` 파일을 열면 됩니다.
- 주요 테이블은 `receipt_status`, `file_import_logs`, `import_metadata` 입니다.

## VBA / AutoHotkey 연동

DB 파일이 생성되면 VBA나 AutoHotkey에서도 바로 읽을 수 있습니다.

- VBA 예제: `examples/read_sqlite_vba.bas`
- AHK 예제: `examples/read_sqlite_ahk.ahk`

## 참고

- 현재 수축률 설정은 `config/parts/shrinkage.json`에서 관리합니다.
- 경로, 파일명, 파트명, 설명, DB 이름 규칙도 모두 설정 파일에서 바꿀 수 있습니다.
