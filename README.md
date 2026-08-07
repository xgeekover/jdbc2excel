# jdbc2excel

[![Python](https://img.shields.io/badge/python-3.9%2B-3776AB?logo=python&logoColor=white)](https://www.python.org/)
[![Java](https://img.shields.io/badge/java-8%2B-ED8B00?logo=openjdk&logoColor=white)](https://adoptium.net/)
[![JDBC](https://img.shields.io/badge/JDBC-jaydebeapi-4B8BBE)](https://pypi.org/project/JayDeBeApi/)
[![Excel](https://img.shields.io/badge/Excel-openpyxl-217346?logo=microsoftexcel&logoColor=white)](https://openpyxl.readthedocs.io/)
[![Last Commit](https://img.shields.io/github/last-commit/xgeekover/jdbc2excel)](https://github.com/xgeekover/jdbc2excel/commits/main)

**리포지토리**: https://github.com/xgeekover/jdbc2excel

JDBC로 여러 DBMS에 조회 쿼리를 실행하고, 결과를 서식 있는 엑셀(.xlsx)로 저장하는 도구.

- DBMS별로 엑셀 파일 1개, 쿼리별로 시트(탭) 1개 생성
- 헤더: **볼드** + 회색 배경 + 자동필터, 첫 행 고정(틀 고정)
- 전체 셀: 검정 실선 테두리 + 자동 줄바꿈, 열 너비 자동 조정(한글은 2칸 폭으로 계산)
- 한글·유니코드 기호 안전 (JDBC → 파이썬 str → xlsx 전 구간 유니코드)

## 요구 사항

- Python 3.9+
- Java(JRE/JDK) 8+ — JPype가 JVM을 띄워 JDBC 드라이버를 실행한다
- 각 DBMS의 JDBC 드라이버 jar (`drivers/` 폴더에 넣는 것을 권장)

## 설치

```bash
git clone https://github.com/xgeekover/jdbc2excel.git
cd jdbc2excel
python3 -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

## 실행

```bash
python main.py                   # ./config.json 사용
python main.py -c other.json     # 다른 설정 파일
python main.py --dry-run         # DB 접속 없이 설정·쿼리 파싱만 검증
```

결과는 `output/{DB이름}_{YYYYMMDD_HHMMSS}.xlsx`로 저장된다.
종료 코드: 전부 성공 0, 접속 실패나 실패한 쿼리가 하나라도 있으면 1.
실패한 쿼리는 해당 시트에 오류 메시지와 SQL이 기록되고 나머지 쿼리는 계속 실행된다.

## config.json

```json
{
  "output_dir": "output",          // 엑셀 저장 폴더 (설정 파일 기준 상대 경로)
  "script_file": "script.sql",     // 기본 스크립트 (DB별로 override 가능)
  "fetch_size": 1000,              // fetchmany 배치 크기
  "databases": [
    {
      "name": "postgres_dev",              // 파일명·로그에 쓰이는 식별자
      "enabled": true,                     // false면 건너뜀
      "driver_class": "org.postgresql.Driver",
      "jdbc_url": "jdbc:postgresql://localhost:5432/mydb",
      "user": "postgres",
      "password": "postgres",              // 또는 "password_env": "PG_PASSWORD" (환경변수에서 읽음)
      "jar_files": ["drivers/postgresql-42.7.4.jar"],
      "script_file": "script_pg.sql",      // (선택) 이 DB 전용 스크립트
      "properties": {"encrypt": "false"}   // (선택) 드라이버 추가 접속 속성
    }
  ]
}
```

`password_env`가 있으면 `password`보다 우선한다. 운영 비밀번호는 환경변수 사용을 권장.

## script.sql

```sql
-- @separator: ;
-- @tab: 사용자 목록
SELECT id, name, email FROM users;

-- @tab: 최근 주문
SELECT * FROM orders WHERE created_at >= CURRENT_DATE - 7;

SELECT COUNT(*) FROM users;   -- @tab이 없으면 시트 이름은 Query3처럼 자동 번호
```

- `-- @separator: <구분자>` — 파일 전체에 적용할 쿼리 구분자. 생략 시 `;`.
  PL/SQL처럼 본문에 `;`가 들어가면 `-- @separator: @@` 등으로 바꾸고 쿼리 끝마다 `@@`를 붙인다.
- `-- @tab: <시트이름>` — 바로 뒤 쿼리의 엑셀 탭 제목. 쿼리마다 지정.
- 구분자는 **줄 끝(또는 단독 줄)**에 있어야 쿼리가 분리된다.
- 파일 인코딩은 UTF-8 (BOM 허용).

## 주의 사항

- **jar 통합 로딩**: JVM classpath는 최초 접속 시 한 번만 고정되므로, 활성화된 모든 DB의
  jar를 합쳐서 로딩한다. 실행 중 config의 jar를 바꾸면 프로세스를 새로 시작해야 반영된다.
- **시트 이름 제약**: 엑셀 규칙에 따라 31자로 자르고 `\ / * ? : [ ]`는 `_`로 치환한다.
  이름이 겹치면 `_2`, `_3` 접미사가 붙는다.
- **행 수 한도**: 엑셀 시트 한도(1,048,576행) 초과분은 잘리고 경고 로그가 남는다.
- **문자열 리터럴 안의 구분자**: SQL 문자열 안에 구분자가 줄 끝 형태로 들어가면 잘못
  분리될 수 있다. 그런 경우 `@separator`를 `@@` 같은 특수 구분자로 바꿔 사용한다.
- **Windows 콘솔 한글 깨짐**: 프로그램이 stdout/stderr를 UTF-8로 재설정하지만,
  필요하면 `set PYTHONUTF8=1`을 함께 사용한다 (엑셀 파일 자체는 콘솔과 무관하게 안전).

## 폴더 구조

```
jdbc2excel/
├── main.py            # 진입점 (설정 로딩, 실행 흐름)
├── db.py              # jaydebeapi 접속·쿼리 실행
├── excel_writer.py    # openpyxl 서식·저장
├── sql_script.py      # script.sql 파서 (@separator / @tab)
├── config.json        # DBMS 설정
├── script.sql         # 실행할 쿼리 목록
├── drivers/           # JDBC 드라이버 jar 두는 곳
└── output/            # 생성된 엑셀 (자동 생성)
```
