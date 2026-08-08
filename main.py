#!/usr/bin/env python3
"""JDBC로 여러 DBMS에 조회 쿼리를 실행하고, 결과를 서식 있는 엑셀로 저장한다.

사용법:
    python main.py                      # ./config.json 사용
    python main.py -c other/config.json
    python main.py --dry-run            # DB 접속 없이 설정·쿼리 파싱만 검증
"""

import argparse
import json
import logging
import os
import sys
from datetime import datetime
from pathlib import Path

from excel_writer import write_workbook
from sql_script import parse_script

log = logging.getLogger("jdbc2excel")

REQUIRED_DB_KEYS = ("name", "driver_class", "jdbc_url", "jar_files")


def _parse_args():
    parser = argparse.ArgumentParser(description="JDBC 조회 결과를 엑셀로 내보내기")
    parser.add_argument("-c", "--config", default="config.json", help="설정 파일 경로 (기본 config.json)")
    parser.add_argument("--dry-run", action="store_true", help="접속 없이 설정과 스크립트 파싱만 확인")
    parser.add_argument("-v", "--verbose", action="store_true", help="DEBUG 로그 출력")
    return parser.parse_args()


def _resolve(base, path_str):
    """설정 파일 안의 상대 경로를 설정 파일 위치 기준으로 해석한다."""
    p = Path(path_str).expanduser()
    return p if p.is_absolute() else (base / p)


def _resolve_password(db):
    env_key = db.get("password_env")
    if env_key:
        value = os.environ.get(env_key)
        if value is None:
            raise SystemExit(f"[{db['name']}] 환경변수 {env_key}가 설정되어 있지 않습니다")
        return value
    return db.get("password", "")


def _validate_databases(databases):
    for db in databases:
        missing = [k for k in REQUIRED_DB_KEYS if k not in db]
        if missing:
            raise SystemExit(f"DB 설정 오류 ({db.get('name', '이름 없음')}): 필수 항목 누락 {missing}")


def main():
    args = _parse_args()
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, ValueError):
            pass
    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(asctime)s %(levelname)-7s %(message)s",
        datefmt="%H:%M:%S",
    )

    config_path = Path(args.config).resolve()
    if not config_path.exists():
        log.error("설정 파일이 없습니다: %s", config_path)
        return 1
    conf = json.loads(config_path.read_text(encoding="utf-8"))
    base = config_path.parent

    databases = [d for d in conf.get("databases", []) if d.get("enabled")]
    if not databases:
        log.error("enabled=true인 DBMS가 없습니다 (%s)", config_path)
        return 1
    _validate_databases(databases)

    default_script = conf.get("script_file", "script.sql")
    output_dir = _resolve(base, conf.get("output_dir", "output"))
    fetch_size = int(conf.get("fetch_size", 1000))
    verify = bool(conf.get("verify", True))

    scripts = {}
    for db in databases:
        script_path = _resolve(base, db.get("script_file", default_script))
        if script_path not in scripts:
            if not script_path.exists():
                log.error("스크립트 파일이 없습니다: %s", script_path)
                return 1
            scripts[script_path] = parse_script(script_path)
            log.info("%s: 쿼리 %d개 파싱", script_path.name, len(scripts[script_path]))

    # JVM classpath는 최초 접속 시 한 번만 고정되므로 모든 활성 DB의 jar를 합쳐서 넘긴다.
    all_jars = []
    for db in databases:
        for jar in db["jar_files"]:
            jar_path = _resolve(base, jar)
            if not jar_path.exists():
                log.warning("[%s] jar 파일이 없습니다: %s", db["name"], jar_path)
            if str(jar_path) not in all_jars:
                all_jars.append(str(jar_path))

    if args.dry_run:
        for db in databases:
            script_path = _resolve(base, db.get("script_file", default_script))
            log.info("[%s] %s → %d개 쿼리", db["name"], db["jdbc_url"], len(scripts[script_path]))
            for i, q in enumerate(scripts[script_path], 1):
                first_line = next(
                    (ln.strip() for ln in q.sql.splitlines()
                     if ln.strip() and not ln.strip().startswith("--")),
                    "",
                )
                log.info("    %d. [%s] %s", i, q.tab, first_line)
        log.info("dry-run 완료 (DB 접속·엑셀 생성 안 함)")
        return 0

    from db import run_queries  # jaydebeapi(JPype) 로딩 비용이 있어 실제 실행 시에만 임포트

    output_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    had_error = False

    for db in databases:
        name = db["name"]
        queries = scripts[_resolve(base, db.get("script_file", default_script))]
        if not queries:
            log.warning("[%s] 실행할 쿼리가 없어 건너뜁니다", name)
            continue
        log.info("[%s] 접속: %s", name, db["jdbc_url"])
        try:
            results = run_queries(db, _resolve_password(db), queries, all_jars, fetch_size)
        except Exception as e:  # noqa: BLE001 - DB 하나가 실패해도 나머지는 계속
            log.error("[%s] 접속 실패: %s", name, e)
            had_error = True
            continue

        out_path = output_dir / f"{name}_{timestamp}.xlsx"
        sheet_map = write_workbook(out_path, results)
        ok = sum(1 for r in results if r.error is None)
        if ok < len(results):
            had_error = True
        log.info("[%s] 쿼리 %d/%d개 성공 → %s", name, ok, len(results), out_path)

        if verify:
            from verifier import verify_workbook

            mismatches = verify_workbook(out_path, sheet_map)
            if mismatches:
                had_error = True
                for m in mismatches[:10]:
                    log.error("[%s] 검증 불일치: %s", name, m)
                if len(mismatches) > 10:
                    log.error("[%s] ... 외 %d건", name, len(mismatches) - 10)
            else:
                total_rows = sum(len(r.rows) for r in results if r.error is None)
                log.info(
                    "[%s] 엑셀 검증 통과: 시트 %d개, 총 %d행 셀 단위 일치",
                    name, ok, total_rows,
                )

    return 1 if had_error else 0


if __name__ == "__main__":
    sys.exit(main())
