"""jaydebeapi(JDBC) 기반 쿼리 실행."""

import datetime
import decimal
import logging
import time

import jaydebeapi

log = logging.getLogger("jdbc2excel")

_PASSTHROUGH_TYPES = (
    str,
    int,
    float,
    bool,
    datetime.datetime,
    datetime.date,
    datetime.time,
    decimal.Decimal,
)


def _coerce(value):
    """openpyxl이 그대로 쓸 수 있는 파이썬 타입만 통과시키고, 나머지(자바 객체 등)는 문자열화한다."""
    if value is None or isinstance(value, _PASSTHROUGH_TYPES):
        return value
    return str(value)


class QueryResult:
    def __init__(self, tab, sql):
        self.tab = tab
        self.sql = sql
        self.columns = []
        self.rows = []
        self.error = None
        self.elapsed = 0.0


def run_queries(db_conf, password, queries, jars, fetch_size=1000):
    """DB 하나에 접속해 queries를 순서대로 실행하고 QueryResult 목록을 돌려준다.

    접속 실패는 예외로 올리고, 개별 쿼리 실패는 QueryResult.error에 담고 계속 진행한다.
    jars는 (JVM classpath가 최초 접속 시 고정되므로) 활성화된 모든 DB의 jar 합집합이어야 한다.
    """
    name = db_conf["name"]
    props = {}
    if db_conf.get("user") is not None:
        props["user"] = str(db_conf["user"])
    if password is not None:
        props["password"] = str(password)
    props.update({str(k): str(v) for k, v in db_conf.get("properties", {}).items()})

    conn = jaydebeapi.connect(
        db_conf["driver_class"],
        db_conf["jdbc_url"],
        props,
        jars=jars or None,
    )
    results = []
    try:
        try:
            conn.jconn.setAutoCommit(True)
        except Exception:  # noqa: BLE001 - 드라이버가 지원하지 않아도 치명적이지 않다
            pass

        cur = conn.cursor()
        try:
            for q in queries:
                res = QueryResult(q.tab, q.sql)
                results.append(res)
                start = time.monotonic()
                try:
                    cur.execute(q.sql)
                    if cur.description is None:
                        res.error = "결과 집합이 없는 문장입니다 (조회(SELECT) 쿼리만 지원)"
                        continue
                    res.columns = [str(d[0]) for d in cur.description]
                    while True:
                        batch = cur.fetchmany(fetch_size)
                        if not batch:
                            break
                        res.rows.extend([_coerce(v) for v in row] for row in batch)
                except Exception as e:  # noqa: BLE001 - 쿼리 하나가 실패해도 나머지는 계속
                    res.error = f"{type(e).__name__}: {e}"
                    log.error("[%s] 쿼리 '%s' 실행 실패: %s", name, res.tab, res.error)
                finally:
                    res.elapsed = time.monotonic() - start
                if res.error is None:
                    log.info(
                        "[%s] '%s' %d행 조회 (%.2fs)",
                        name, res.tab, len(res.rows), res.elapsed,
                    )
        finally:
            cur.close()
    finally:
        conn.close()
    return results
