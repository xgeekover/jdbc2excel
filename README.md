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
- JDBC 드라이버 jar만 있으면 어떤 DBMS든 접속 가능 (Oracle, MSSQL, PostgreSQL, MySQL, DB2, Tibero, ...)

## 요구 사항

| 항목 | 버전 | 비고 |
|---|---|---|
| Python | 3.9+ | |
| Java (JRE/JDK) | 8+ | JPype가 JVM을 띄워 JDBC 드라이버를 실행 |
| JDBC 드라이버 jar | DBMS별 | `drivers/` 폴더에 배치 |

파이썬 모듈 의존성은 `requirements.txt`에 정의되어 있다: `jaydebeapi`, `JPype1`, `openpyxl`.

## 빠른 시작

```bash
# 1. 클론 및 의존성 설치
git clone https://github.com/xgeekover/jdbc2excel.git
cd jdbc2excel
python3 -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate
pip install -r requirements.txt

# 2. JDBC 드라이버 jar를 drivers/에 넣기 (아래 "드라이버 준비" 참고)

# 3. config.json에서 접속 정보 수정, 사용할 DB만 "enabled": true

# 4. script.sql에 실행할 쿼리 작성

# 5. 접속 없이 설정·쿼리 파싱 검증
python main.py --dry-run

# 6. 실행 → output/{DB이름}_{타임스탬프}.xlsx 생성
python main.py
```

## 사용법 상세

### 1. JDBC 드라이버 준비

사용할 DBMS의 드라이버 jar를 내려받아 `drivers/`에 넣는다. 전부 Maven Central에서 받을 수 있다.

| DBMS | 아티팩트 | 다운로드 |
|---|---|---|
| Oracle | `com.oracle.database.jdbc:ojdbc11` | [Maven Central](https://central.sonatype.com/artifact/com.oracle.database.jdbc/ojdbc11) |
| MSSQL | `com.microsoft.sqlserver:mssql-jdbc` (`*.jre11` 버전) | [Maven Central](https://central.sonatype.com/artifact/com.microsoft.sqlserver/mssql-jdbc) |
| PostgreSQL | `org.postgresql:postgresql` | [Maven Central](https://central.sonatype.com/artifact/org.postgresql/postgresql) |
| MySQL | `com.mysql:mysql-connector-j` | [Maven Central](https://central.sonatype.com/artifact/com.mysql/mysql-connector-j) |
| MariaDB | `org.mariadb.jdbc:mariadb-java-client` | [Maven Central](https://central.sonatype.com/artifact/org.mariadb.jdbc/mariadb-java-client) |
| DB2 | `com.ibm.db2:jcc` | [Maven Central](https://central.sonatype.com/artifact/com.ibm.db2/jcc) |

### 2. config.json 작성

최상위 필드:

| 필드 | 필수 | 기본값 | 설명 |
|---|---|---|---|
| `output_dir` | | `output` | 엑셀 저장 폴더 (config.json 위치 기준 상대 경로) |
| `script_file` | | `script.sql` | 기본 스크립트 파일 (DB별로 override 가능) |
| `fetch_size` | | `1000` | 한 번에 가져올 행 수 (대용량 결과의 메모리 완충) |
| `databases` | ✔ | | DB 접속 정보 배열 |

`databases[]` 항목 필드:

| 필드 | 필수 | 설명 |
|---|---|---|
| `name` | ✔ | 식별자. 출력 파일명(`{name}_{타임스탬프}.xlsx`)과 로그에 사용 |
| `enabled` | ✔ | `false`면 이 DB는 건너뜀 |
| `driver_class` | ✔ | JDBC 드라이버 클래스 (이 값이 사실상 DBMS 구분) |
| `jdbc_url` | ✔ | JDBC 접속 URL |
| `jar_files` | ✔ | 드라이버 jar 경로 배열 (config.json 위치 기준 상대 경로) |
| `user` | | 접속 계정 |
| `password` | | 접속 비밀번호 (평문) |
| `password_env` | | 비밀번호를 읽을 환경변수 이름. `password`보다 우선. 운영 환경 권장 |
| `script_file` | | 이 DB 전용 스크립트 (SQL 방언이 다를 때 사용) |
| `properties` | | 드라이버 추가 접속 속성 (예: MSSQL `{"encrypt": "false"}`) |

DBMS별 `driver_class` / `jdbc_url` 값:

| DBMS | driver_class | jdbc_url 패턴 |
|---|---|---|
| Oracle | `oracle.jdbc.OracleDriver` | `jdbc:oracle:thin:@//host:1521/서비스명` (SID 방식: `@host:1521:SID`) |
| MSSQL | `com.microsoft.sqlserver.jdbc.SQLServerDriver` | `jdbc:sqlserver://host:1433;databaseName=DB명` |
| PostgreSQL | `org.postgresql.Driver` | `jdbc:postgresql://host:5432/DB명` |
| MySQL | `com.mysql.cj.jdbc.Driver` | `jdbc:mysql://host:3306/DB명` |
| MariaDB | `org.mariadb.jdbc.Driver` | `jdbc:mariadb://host:3306/DB명` |
| DB2 | `com.ibm.db2.jcc.DB2Driver` | `jdbc:db2://host:50000/DB명` |
| Tibero | `com.tmax.tibero.jdbc.TbDriver` | `jdbc:tibero:thin:@host:8629:SID` |

작성 예 (전체 예시는 리포의 `config.json` 참고):

```json
{
  "output_dir": "output",
  "script_file": "script.sql",
  "databases": [
    {
      "name": "erp_oracle",
      "enabled": true,
      "driver_class": "oracle.jdbc.OracleDriver",
      "jdbc_url": "jdbc:oracle:thin:@//db.example.com:1521/ORCLPDB1",
      "user": "scott",
      "password_env": "ORACLE_PASSWORD",
      "jar_files": ["drivers/ojdbc11-23.26.3.0.0.jar"]
    },
    {
      "name": "dw_mssql",
      "enabled": true,
      "driver_class": "com.microsoft.sqlserver.jdbc.SQLServerDriver",
      "jdbc_url": "jdbc:sqlserver://dw.example.com:1433;databaseName=DW",
      "user": "reader",
      "password": "secret",
      "jar_files": ["drivers/mssql-jdbc-13.4.0.jre11.jar"],
      "properties": {"encrypt": "false"},
      "script_file": "script_mssql.sql"
    }
  ]
}
```

- 같은 DBMS를 여러 항목으로 등록해도 된다 (`oracle_운영`, `oracle_개발` 등). jar는 공유된다.
- `password_env` 사용 시 실행 전에 환경변수를 설정한다:
  `export ORACLE_PASSWORD='실제비밀번호'` (Windows: `set ORACLE_PASSWORD=실제비밀번호`)

### 3. script.sql 작성

```sql
-- @separator: ;
-- @tab: 사용자 목록
SELECT id, name, email FROM users;

-- @tab: 최근 주문
SELECT * FROM orders WHERE created_at >= CURRENT_DATE - 7;

SELECT COUNT(*) FROM users;   -- @tab이 없으면 시트 이름은 Query3처럼 자동 번호
```

규칙:

- `-- @separator: <구분자>` — 파일 전체에 적용할 쿼리 구분자. 생략하면 `;`.
- `-- @tab: <시트이름>` — 바로 뒤 쿼리의 엑셀 탭 제목. 쿼리마다 지정.
- 구분자는 **줄 끝(또는 단독 줄)**에 있어야 쿼리가 분리된다.
- `--` 주석은 그대로 둬도 된다 (실행할 문장이 없는 조각은 자동으로 건너뜀).
- 파일 인코딩은 UTF-8 (BOM 허용).
- 조회(SELECT) 쿼리만 지원한다. DML/DDL은 오류로 기록된다.

본문에 `;`가 들어가는 PL/SQL 등은 구분자를 바꿔서 쓴다:

```sql
-- @separator: @@

-- @tab: 프로시저 결과
SELECT status, err_msg
  FROM batch_log
 WHERE run_date = TRUNC(SYSDATE)
@@

-- @tab: 커서 블록
SELECT CASE WHEN cnt > 0 THEN '실행됨;정상' ELSE '미실행' END AS 상태
  FROM (SELECT COUNT(*) cnt FROM job_history)
@@
```

DBMS마다 SQL 방언이 다르면(예: Oracle `SYSDATE` vs MSSQL `GETDATE()`) DB 항목에
`"script_file": "script_mssql.sql"`을 넣어 그 DB만 다른 쿼리 세트를 실행한다.

### 4. 실행

```bash
python main.py                   # ./config.json 사용
python main.py -c other.json     # 다른 설정 파일 지정
python main.py --dry-run         # DB 접속 없이 설정·쿼리 파싱만 검증
python main.py -v                # DEBUG 로그 출력
```

`--dry-run`은 접속 정보를 채우기 전에 config·스크립트 문법을 미리 확인하는 용도다.
DB별 쿼리 목록과 탭 이름, jar 존재 여부를 출력한다.

종료 코드:

| 코드 | 의미 |
|---|---|
| 0 | 전체 성공 |
| 1 | 접속 실패 또는 실패한 쿼리가 1개 이상 (성공한 결과는 정상 저장됨) |

배치 스케줄러(cron, Jenkins 등)에 넣을 때 종료 코드로 성패를 판정하면 된다.

### 5. 출력 확인

- 저장 위치: `output/{DB이름}_{YYYYMMDD_HHMMSS}.xlsx` — 실행할 때마다 새 파일 생성 (덮어쓰지 않음)
- 시트 구성: 쿼리 순서대로 시트 1개씩. `@tab` 이름이 겹치면 `_2`, `_3` 접미사
- 실패한 쿼리: 해당 시트에 **오류 메시지와 SQL**이 기록되고 나머지 쿼리는 계속 실행
- 서식: 헤더 볼드·회색 배경·자동필터·첫 행 고정, 전체 셀 검정 테두리·자동 줄바꿈, 열 너비 자동

## 문제 해결

| 증상 | 원인 / 해결 |
|---|---|
| `JVMNotFoundException` | Java 미설치 또는 JPype가 JVM을 못 찾음. `java -version` 확인, 필요 시 `JAVA_HOME` 설정 |
| `ClassNotFoundException: <driver_class>` | jar 경로 오류 또는 오타. `--dry-run`으로 jar 존재 경고 확인 |
| MSSQL SSL 오류 (`encrypt`) | 드라이버 12.x부터 기본 `encrypt=true`. `"properties": {"encrypt": "false"}` 또는 `{"encrypt": "true", "trustServerCertificate": "true"}` |
| Oracle `ORA-12514` (리스너가 서비스명을 모름) | 서비스명 방식(`@//host:1521/SERVICE`)과 SID 방식(`@host:1521:SID`)을 바꿔서 시도 |
| 환경변수 오류로 즉시 종료 | `password_env`에 지정한 환경변수가 미설정. `export 변수명=값` 후 재실행 |
| jar 교체가 반영 안 됨 | JVM classpath는 최초 접속 시 고정. 프로세스를 새로 시작해야 반영 |
| Windows 콘솔 한글 깨짐 | 프로그램이 stdout을 UTF-8로 재설정하지만, 필요 시 `set PYTHONUTF8=1` (엑셀 파일 자체는 무관하게 안전) |
| 대용량 결과 메모리 부족 | `fetch_size`는 전송 완충일 뿐 결과 전체는 메모리에 적재됨. 쿼리에서 행수를 줄이거나 분할 조회 권장 |

## 주의 사항

- **jar 통합 로딩**: 활성화된 모든 DB의 jar를 합쳐서 JVM에 한 번에 로딩한다.
- **시트 이름 제약**: 엑셀 규칙에 따라 31자로 자르고 `\ / * ? : [ ]`는 `_`로 치환한다.
- **행 수 한도**: 엑셀 시트 한도(1,048,576행) 초과분은 잘리고 경고 로그가 남는다.
- **문자열 리터럴 안의 구분자**: SQL 문자열 안에 구분자가 줄 끝 형태로 들어가면 잘못
  분리될 수 있다. 그런 경우 `@separator`를 `@@` 같은 특수 구분자로 바꿔 사용한다.
- **DATE 컬럼**: jaydebeapi 기본 변환에 따라 `'2024-01-15'` 형태의 문자열로 기록된다.

## 폴더 구조

```
jdbc2excel/
├── main.py            # 진입점 (설정 로딩, 실행 흐름)
├── db.py              # jaydebeapi 접속·쿼리 실행
├── excel_writer.py    # openpyxl 서식·저장
├── sql_script.py      # script.sql 파서 (@separator / @tab)
├── config.json        # DBMS 설정
├── script.sql         # 실행할 쿼리 목록
├── requirements.txt   # 파이썬 의존성 (jaydebeapi, JPype1, openpyxl)
├── drivers/           # JDBC 드라이버 jar 두는 곳
└── output/            # 생성된 엑셀 (자동 생성)
```
