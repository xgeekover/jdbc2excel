-- @separator: ;
-- 구분자는 파일 상단에서 한 번 지정한다 (생략 시 기본값 ;).
-- PL/SQL처럼 본문에 ;가 들어가면 "-- @separator: @@" 등으로 바꾸고 쿼리 끝마다 @@를 붙인다.
-- 구분자는 반드시 줄 끝(또는 단독 줄)에 있어야 한다.

-- @tab: 사용자 목록
SELECT id, name, email, created_at
  FROM users
 ORDER BY id;

-- @tab: 최근 7일 주문
SELECT o.id       AS 주문번호,
       u.name     AS 주문자,
       o.amount   AS 금액,
       o.created_at AS 주문일시
  FROM orders o
  JOIN users u ON u.id = o.user_id
 WHERE o.created_at >= CURRENT_DATE - 7
 ORDER BY o.created_at DESC;

-- @tab 지시어가 없으면 시트 이름은 Query3 처럼 자동 번호가 붙는다.
SELECT COUNT(*) AS total_users FROM users;
