"""script.sql 파서.

문법:
    -- @separator: ;          파일 전체에 적용할 쿼리 구분자 (기본값 ;)
    -- @tab: 시트이름          바로 뒤 쿼리의 엑셀 탭 제목 (없으면 Query1, Query2, ...)

구분자는 반드시 줄 끝(또는 단독 줄)에 있어야 쿼리가 분리된다.
PL/SQL처럼 본문에 ;가 들어가는 쿼리는 @separator를 @@ 등으로 바꿔 사용한다.
"""

import re
from dataclasses import dataclass
from pathlib import Path

_DIRECTIVE_RE = re.compile(
    r"^\s*--\s*@(?P<key>separator|tab)\s*[:=]\s*(?P<value>.+?)\s*$",
    re.IGNORECASE,
)


@dataclass
class Query:
    tab: str
    sql: str


def parse_script(path):
    """script.sql을 파싱해 Query 목록을 돌려준다."""
    text = Path(path).read_text(encoding="utf-8-sig")

    separator = ";"
    kept_lines = []
    for line in text.splitlines():
        m = _DIRECTIVE_RE.match(line)
        if m and m.group("key").lower() == "separator":
            separator = m.group("value")
        else:
            kept_lines.append(line)
    body = "\n".join(kept_lines)

    chunks = re.split(rf"{re.escape(separator)}[ \t]*$", body, flags=re.MULTILINE)

    queries = []
    for chunk in chunks:
        tab = None
        sql_lines = []
        for line in chunk.splitlines():
            m = _DIRECTIVE_RE.match(line)
            if m and m.group("key").lower() == "tab":
                if tab is None:
                    tab = m.group("value")
            else:
                sql_lines.append(line)
        sql = "\n".join(sql_lines).strip()

        has_statement = any(
            ln.strip() and not ln.strip().startswith("--") for ln in sql_lines
        )
        if not has_statement:
            continue

        queries.append(Query(tab=tab or f"Query{len(queries) + 1}", sql=sql))

    return queries
