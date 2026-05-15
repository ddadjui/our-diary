"""
누나와 가계부.xlsx → Firebase Realtime Database 마이그레이션 스크립트

Firebase DB 경로:
  ourdiary/expenses/{boy|girl|shared}/{YYYY-MM}/{DD}/{id}
    - desc: string
    - amount: number
    - cat: string
    - ts: number (Unix timestamp ms)
"""

import re
import time
import uuid
import requests
import pandas as pd
from datetime import datetime

# ── 설정 ──────────────────────────────────────────────────────────────────────
EXCEL_PATH = "/Users/user/Downloads/누나와 가계부.xlsx"
FIREBASE_BASE = "https://our-diary-240f6-default-rtdb.firebaseio.com/ourdiary/expenses"

# 시트 이름 → Firebase 탭 이름
SHEET_TAB_MAP = {
    "주머니쥐": "boy",
    "누나양":   "girl",
    "공금":     "shared",
}

# 연도 (엑셀에 연도 정보 없음 → 2026년으로 가정)
YEAR = 2026

# ── 카테고리 분류 ─────────────────────────────────────────────────────────────
CATEGORY_RULES = [
    ("쿠팡카드", ["쿠팡카드", "쿠팡 카드"]),
    ("식비",     ["밥", "점심", "저녁", "아침", "식사", "카페", "커피", "치킨", "피자",
                  "떡볶이", "분식", "편의점", "GS", "CU", "세븐", "스타벅스", "맥도날드",
                  "음식", "식당", "국밥", "김밥", "냉면", "삼겹살", "술", "맥주", "소주",
                  "빵", "음료", "케이크", "아이스크림", "버터떡", "에그타르트",
                  "샤브샤브", "막창", "족발", "찜닭", "순대국", "된장", "갈비",
                  "배달", "쿠팡 이츠", "로또"]),
    ("교통",     ["버스", "지하철", "택시", "주유", "기름", "KTX", "기차", "교통",
                  "티머니"]),
    ("쇼핑",     ["쇼핑", "옷", "의류", "신발", "가방", "아마존", "쿠팡", "무신사",
                  "원피스", "속옷", "케이스", "필름", "화장품", "앰플", "컨실러",
                  "설화수", "수분 크림", "수분크림", "바디로션", "마그네슘",
                  "침낭", "믹싱볼", "면도기", "다이소"]),
    ("의료",     ["병원", "약국", "의원", "치과", "약", "의료", "상비약"]),
    ("구독",     ["넷플릭스", "유튜브", "스포티파이", "구독", "멜론", "왓챠",
                  "클로드", "톡서랍", "톡클라우드", "배달의 민족"]),
    ("여가",     ["영화", "노래방", "놀이공원", "여행", "숙박", "호텔", "게임", "공연",
                  "인터파크", "PC방"]),
    ("생활비",   ["마트", "이마트", "홈플러스", "롯데마트", "생활", "청소", "세탁",
                  "울샴푸", "액상세제", "탈취제", "고춧가루", "쌀", "양파", "감자",
                  "식빵", "우유", "부추", "당면", "콩나물", "두부", "대패 삼겹살",
                  "탄산수", "팽이버섯", "와사비", "모닝캄 생수", "면봉", "우동면",
                  "깻잎", "소시지", "목전지", "새우", "콘칩", "튀김가루", "라이스페이퍼",
                  "식혜", "진간장", "다진마늘", "고추장", "커피(배달)", "전기세",
                  "가스비", "통신비", "관리비", "쓰레기 봉투"]),
    ("선물",     ["선물", "생일", "기념일"]),
]

def classify(desc: str) -> str:
    text = desc.lower()
    for cat, keywords in CATEGORY_RULES:
        for kw in keywords:
            if kw.lower() in text:
                return cat
    return "기타"

# ── 금액 파싱 ─────────────────────────────────────────────────────────────────
# 지원 형식:
#   "점심 식사 11,900원(쿠팡)"
#   "점심 -10,965"  / "점심 -10,965원"
#   "점심 11,900(쿠팡카드)"
#   "점심 11900"
#   "다이소 21,200원"

AMOUNT_PATTERN = re.compile(
    r"(-?[\d,]+(?:\.\d+)?)\s*원?"  # 숫자 (쉼표 포함, 소수점 가능)
    r"(?:\s*\(.*?\))?$"            # 선택적 괄호 메모
)

def parse_entry(raw: str):
    """
    '설명 금액원' 형태의 문자열을 (desc, amount) 로 분해.
    파싱 실패 시 None 반환.
    """
    raw = raw.strip()

    # "신분증 발급 15,000원" 처럼 금액이 끝에 있는 경우
    # 마지막 숫자 토큰을 찾아 분리
    m = re.search(
        r"(-?[\d,]+)\s*원(?:\([^)]*\))?$",
        raw
    )
    if not m:
        # "점심 -10,965" 또는 "점심 10,965(쿠팡카드)" 형태
        m = re.search(
            r"\s+(-?[\d,]+)(?:\([^)]*\))?$",
            raw
        )
    if not m:
        return None

    amount_str = m.group(1).replace(",", "")
    try:
        amount = int(float(amount_str))
    except ValueError:
        return None

    # 금액이 0이면 스킵
    if amount == 0:
        return None

    amount = abs(amount)  # 누나양 시트는 -로 표기

    # desc = 금액 부분을 제거한 나머지
    desc = raw[:m.start()].strip()
    # 괄호 속 메모 제거 (쿠팡 / 데이트 / 개인 등)
    desc = re.sub(r"\([^)]*\)", "", desc).strip()
    # 금액 앞의 공백만 남긴 경우 처리
    if not desc:
        desc = raw

    return desc, amount

# ── 달력 파싱 ─────────────────────────────────────────────────────────────────
MONTH_MAP = {
    "1월": 1, "2월": 2,  "3월": 3,  "4월": 4,
    "5월": 5, "6월": 6,  "7월": 7,  "8월": 8,
    "9월": 9, "10월": 10, "11월": 11, "12월": 12,
}
DAY_COLS = [1, 2, 3, 4, 5, 6, 7]  # 일~토 (column indices in df)

def is_month_header(val):
    """셀 값이 월 헤더이면 해당 월 숫자 반환, 아니면 None."""
    if pd.isna(val):
        return None
    s = str(val).strip()
    return MONTH_MAP.get(s)

def extract_day(val):
    """셀 값에서 날짜 숫자만 추출 (예: '5 어린이날' → 5, '9(쿠팡 카드 정산 완료)' → 9)."""
    if pd.isna(val):
        return None
    s = str(val).strip()
    m = re.match(r"^(\d{1,2})", s)
    if m:
        d = int(m.group(1))
        if 1 <= d <= 31:
            return d
    return None

def is_day_header_row(row) -> bool:
    """행이 요일 헤더(일 월 화 수 목 금 토)인지 확인."""
    days = {"일", "월", "화", "수", "목", "금", "토"}
    vals = [str(row[c]).strip() for c in DAY_COLS if not pd.isna(row[c])]
    return len(vals) >= 5 and all(v in days for v in vals)

def is_expense_cell(cell_str: str) -> bool:
    """셀이 지출 항목(금액 포함)인지 빠르게 확인."""
    return bool(re.search(r"\d[\d,]*\s*원", cell_str) or
                re.search(r"\s-\s*\d[\d,]+", cell_str) or
                re.search(r"\s\d[\d,]+(?:\([^)]*\))?$", cell_str))

def extract_records_from_cell(cell_str, tab, year, month, day):
    """셀 문자열에서 지출 항목 목록을 추출."""
    result = []
    lines = cell_str.split("\n")
    for line in lines:
        line = line.strip()
        if not line:
            continue
        parsed = parse_entry(line)
        if parsed is None:
            continue
        desc, amount = parsed
        # desc 끝의 불필요한 '-' 제거 (누나양 시트 "-" 잔류 현상)
        desc = desc.rstrip(" -").strip()
        if not desc:
            desc = line.split()[0] if line.split() else line
        cat = classify(line)
        try:
            dt = datetime(year, month, day)
            ts = int(dt.timestamp() * 1000)
        except ValueError:
            continue
        result.append({
            "tab":    tab,
            "month":  month,
            "day":    day,
            "desc":   desc,
            "amount": amount,
            "cat":    cat,
            "ts":     ts,
        })
    return result

def parse_sheet(df, tab, year):
    """
    달력형 시트에서 지출 항목 추출.
    - 월 헤더(4월, 5월 등)가 있으면 해당 월로 전환
    - 월 헤더가 없는 시트(공금)는 날짜 감소 시 다음 달로 자동 전환
    반환: [{"tab", "month", "day", "desc", "amount", "cat", "ts"}, ...]
    """
    records = []
    rows = df.values

    # 공금 시트처럼 월 헤더 없이 시작하는 경우 초기 월을 추론
    # 첫 번째 날짜 행에서 가장 작은 날짜를 보고 시트 전체의 시작 월을 판단할 수 없으므로
    # 기본값을 None으로 두되, 첫 날짜 행 발견 시 엑셀 파일 생성 맥락(4월 시작)으로 초기화
    current_month = None
    prev_week_min_day = None  # 이전 주의 최솟값 (월 전환 감지용)
    has_month_header = False

    # 사전 스캔: 월 헤더가 존재하는지 확인
    for row in rows:
        for col in DAY_COLS:
            if col < len(row) and is_month_header(row[col]):
                has_month_header = True
                break
        if has_month_header:
            break

    i = 0
    while i < len(rows):
        row = rows[i]

        # 월 헤더 감지 (모든 열 검사)
        found_month = None
        for col in range(min(len(row), 8)):
            m = is_month_header(row[col])
            if m:
                found_month = m
                break
        if found_month:
            current_month = found_month
            prev_week_min_day = None
            i += 1
            continue

        # 요일 헤더 행 스킵
        row_series = pd.Series(row)
        if is_day_header_row(row_series):
            i += 1
            continue

        # 날짜 행 감지
        day_map = {}
        for col in DAY_COLS:
            if col < len(row):
                d = extract_day(row[col])
                if d is not None:
                    day_map[col] = d

        if not day_map:
            i += 1
            continue

        week_days = sorted(day_map.values())

        # 월 헤더 없는 시트: 날짜가 이전 주보다 작으면 다음 달로 전환
        if not has_month_header:
            if current_month is None:
                current_month = 4  # 공금 시트 시작 월 (4월)
                prev_week_min_day = week_days[0]
            else:
                if prev_week_min_day is not None and week_days[0] < prev_week_min_day:
                    current_month = current_month + 1
                    if current_month > 12:
                        current_month = 1
                prev_week_min_day = week_days[0]

        if current_month is None:
            i += 1
            continue

        # 날짜 행(offset=0)과 다음 행(offset=1) 모두 지출 데이터 검사
        for data_row_offset in [0, 1]:
            data_row_idx = i + data_row_offset
            if data_row_idx >= len(rows):
                break
            data_row = rows[data_row_idx]

            for col in DAY_COLS:
                if col >= len(data_row):
                    continue
                day = day_map.get(col)
                if day is None:
                    continue

                cell_val = data_row[col]
                if pd.isna(cell_val):
                    continue
                cell_str = str(cell_val).strip()
                if not cell_str:
                    continue

                # offset=0(날짜 행)이면 순수 날짜 셀 스킵
                if data_row_offset == 0:
                    # 순수 숫자 날짜
                    if re.match(r"^\d{1,2}$", cell_str):
                        continue
                    # 날짜 + 메모 (예: "9(쿠팡 카드 정산 완료)", "1 노동절", "9💵 쿠팡카드 정산완료")
                    if re.match(r"^\d{1,2}[\s\(一-鿿ÿ-￿]", cell_str):
                        continue

                if not is_expense_cell(cell_str):
                    continue

                records.extend(
                    extract_records_from_cell(cell_str, tab, year, current_month, day)
                )

        i += 1

    return records

# ── Firebase 업로드 ───────────────────────────────────────────────────────────
def push_to_firebase(records):
    """업로드 성공/실패 건수 반환."""
    ok = 0
    fail = 0
    for rec in records:
        tab   = rec["tab"]
        mm    = f"{YEAR}-{rec['month']:02d}"
        dd    = f"{rec['day']:02d}"
        uid   = uuid.uuid4().hex[:8]

        payload = {
            "desc":   rec["desc"],
            "amount": rec["amount"],
            "cat":    rec["cat"],
            "ts":     rec["ts"],
        }

        url = f"{FIREBASE_BASE}/{tab}/{mm}/{dd}/{uid}.json"
        try:
            resp = requests.put(url, json=payload, timeout=10)
            if resp.status_code == 200:
                ok += 1
            else:
                fail += 1
                print(f"  [WARN] PUT {url} → {resp.status_code}: {resp.text[:80]}")
        except requests.RequestException as e:
            fail += 1
            print(f"  [ERROR] {e}")

    return ok, fail

# ── 메인 ─────────────────────────────────────────────────────────────────────
def main():
    print("엑셀 파일 로딩 중...")
    all_sheets = pd.read_excel(EXCEL_PATH, sheet_name=None, header=None)

    all_records = []
    for sheet_name, tab in SHEET_TAB_MAP.items():
        if sheet_name not in all_sheets:
            print(f"  [SKIP] 시트 없음: {sheet_name}")
            continue
        df = all_sheets[sheet_name]
        records = parse_sheet(df, tab, YEAR)
        print(f"  {sheet_name:8s} ({tab:6s}) → {len(records):3d}건 파싱")
        all_records.extend(records)

    print(f"\n총 파싱된 항목: {len(all_records)}건")
    print("Firebase에 업로드 중...\n")

    ok, fail = push_to_firebase(all_records)

    print(f"\n{'='*40}")
    print(f"마이그레이션 완료")
    print(f"  성공: {ok}건")
    if fail:
        print(f"  실패: {fail}건")
    print(f"  합계: {ok + fail}건")
    print(f"{'='*40}")

if __name__ == "__main__":
    main()
