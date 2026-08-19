#!/usr/bin/env python3
"""
대한민국 법원 공시송달(scourt.go.kr) 조회 프로그램
"""
import argparse
import datetime
import json
import os
import random
import re
import sys
import time

import requests
from openpyxl import Workbook

sys.stdout.reconfigure(line_buffering=True)

BASE = "https://ssgo.scourt.go.kr"
ENTRY_URL = f"{BASE}/ssgo/ssgo910/igong.on"
COURT_LIST_XML_URL = f"{BASE}/ssgo/ui/ssgo900/ssgo910/SSGO911F02.xml"
SEARCH_URL = f"{BASE}/ssgo/ssgo910/selectPbancLst.on"

HEADERS = {
    "User-Agent": "Mozilla/5.0",
    "Content-Type": "application/json;charset=UTF-8",
    "X-Requested-With": "XMLHttpRequest",
    "Referer": f"{BASE}/ssgo/ui/ssgo900/ssgo910/SSGO911F02.xml",
}

DAY_RANGE = range(0, 22)  # 0=오늘 ~ 21=21일이상
PAGE_SIZE = 50

# 서버에 부담을 주지 않기 위한 호출 간격 (초). 매 요청마다 이 값에 임의의 지터를 더해 사용한다.
BASE_DELAY = 0.6
JITTER = 0.4
MAX_RETRIES = 4

# 법원 하나를 조회한 결과를 이 시간(초) 동안 캐시로 재사용한다. 캐시가 있으면 네트워크 호출 없이 바로 사용한다.
CACHE_TTL_SECONDS = 6 * 60 * 60  # 6시간
COURT_LIST_CACHE_TTL_SECONDS = 24 * 60 * 60  # 24시간

KEYWORDS_TEMPLATE = """# 공시대상자 필터용 회사명 목록
# 한 줄에 회사명을 하나씩 입력하세요. '#'으로 시작하는 줄은 무시됩니다.
# 회사명이 공시대상자 문구에 "포함"되어 있으면 결과에 나옵니다.
#
# 예시:
# 삼성물산
# 현대건설
# 주식회사 OOO
"""


def get_base_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def get_cache_dir(base_dir):
    path = os.path.join(base_dir, "cache")
    os.makedirs(path, exist_ok=True)
    return path


def sleep_politely():
    time.sleep(BASE_DELAY + random.uniform(0, JITTER))


def load_keywords(base_dir):
    path = os.path.join(base_dir, "keywords.txt")
    if not os.path.exists(path):
        with open(path, "w", encoding="utf-8-sig") as f:
            f.write(KEYWORDS_TEMPLATE)
        print(f"'keywords.txt' 파일이 없어서 새로 만들었습니다: {path}")
        print("메모장으로 열어서 찾고 싶은 회사명을 한 줄에 하나씩 입력한 뒤 다시 실행해주세요.")
        return []

    keywords = []
    with open(path, "r", encoding="utf-8-sig") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            keywords.append(line)
    return keywords


def new_session():
    s = requests.Session()
    s.get(ENTRY_URL, headers={"User-Agent": HEADERS["User-Agent"]}, timeout=15)
    return s


def fetch_with_retry(fn, *args, **kwargs):
    last_err = None
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            return fn(*args, **kwargs)
        except requests.exceptions.RequestException as e:
            last_err = e
            wait = min(2 ** attempt, 20) + random.uniform(0, 1)
            print(f"  (네트워크 오류, {wait:.1f}초 후 재시도 {attempt}/{MAX_RETRIES}: {e})")
            time.sleep(wait)
    raise last_err if last_err is not None else RuntimeError("요청에 반복적으로 실패했습니다.")


def fetch_court_list(session, base_dir, force=False):
    cache_path = os.path.join(get_cache_dir(base_dir), "court_list.json")
    if not force and os.path.exists(cache_path):
        age = time.time() - os.path.getmtime(cache_path)
        if age < COURT_LIST_CACHE_TTL_SECONDS:
            with open(cache_path, "r", encoding="utf-8") as f:
                return json.load(f)

    def _do():
        resp = session.get(COURT_LIST_XML_URL, headers={"User-Agent": HEADERS["User-Agent"]}, timeout=15)
        resp.encoding = "utf-8"
        return resp.text

    text = fetch_with_retry(_do)

    block_match = re.search(
        r'<w2:dataList[^>]*id="dlt_cortCd"[^>]*>(.*?)</w2:dataList>', text, re.S
    )
    if not block_match:
        raise RuntimeError("법원 목록을 불러오지 못했습니다. (사이트 구조가 변경되었을 수 있습니다)")
    block = block_match.group(1)

    courts = []
    for row in re.findall(r"<w2:row>(.*?)</w2:row>", block, re.S):
        cd = re.search(r"<cortCd><!\[CDATA\[(.*?)\]\]></cortCd>", row)
        nm = re.search(r"<cortNm><!\[CDATA\[(.*?)\]\]></cortNm>", row)
        area = re.search(r"<area><!\[CDATA\[(.*?)\]\]></area>", row)
        if cd and nm:
            courts.append({
                "cortCd": cd.group(1),
                "cortNm": nm.group(1),
                "area": area.group(1) if area else "",
            })

    with open(cache_path, "w", encoding="utf-8") as f:
        json.dump(courts, f, ensure_ascii=False)
    return courts


def print_all_courts(courts):
    current_area = None
    for i, c in enumerate(courts, 1):
        if c["area"] != current_area:
            current_area = c["area"]
            print(f"\n[{current_area}]")
        print(f"  {i:>3}. {c['cortNm']}")


def choose_court(courts):
    while True:
        query = input(
            "\n법원명을 입력하세요 (예: 서울중앙 / 전체 목록을 보려면 '목록' 입력): "
        ).strip()

        if query in ("목록", "list", "리스트"):
            print_all_courts(courts)
            continue

        if query.isdigit():
            idx = int(query)
            if 1 <= idx <= len(courts):
                return courts[idx - 1]
            print("범위를 벗어난 번호입니다. 다시 입력해주세요.")
            continue

        if not query:
            print("법원명을 입력해주세요.")
            continue

        exact = [c for c in courts if c["cortNm"] == query]
        matches = exact if exact else [c for c in courts if query in c["cortNm"]]

        if not matches:
            print(f"'{query}'에 해당하는 법원을 찾을 수 없습니다. 다시 입력해주세요.")
            continue

        if len(matches) == 1:
            return matches[0]

        print(f"\n'{query}'와(과) 일치하는 법원이 여러 개 있습니다:")
        for i, c in enumerate(matches, 1):
            print(f"  {i}. {c['cortNm']} ({c['area']})")
        sel = input("번호를 선택하세요: ").strip()
        if sel.isdigit() and 1 <= int(sel) <= len(matches):
            return matches[int(sel) - 1]
        print("잘못된 선택입니다. 처음부터 다시 입력해주세요.")


def get_court_family(courts, base_court):
    prefix = base_court["cortNm"] + " "
    return [
        c for c in courts
        if c["cortNm"] == base_court["cortNm"] or c["cortNm"].startswith(prefix)
    ]


def search_page(session, cort_cd, crtr_day, page_no):
    payload = {
        "dma_search": {
            "cortCd": cort_cd,
            "csNo": "",
            "crtrDay": str(crtr_day),
            "searchTyp": "ymd",
            "pageNo": str(page_no),
        }
    }

    def _do():
        resp = session.post(SEARCH_URL, json=payload, headers=HEADERS, timeout=20)
        resp.raise_for_status()
        return resp.json()

    body = fetch_with_retry(_do)
    return body.get("data", {}).get("dlt_pbanc", []) or []


def crawl_court_live(session, cort_cd, cort_nm):
    all_rows = {}
    for day in DAY_RANGE:
        page_no = 1
        while True:
            rows = search_page(session, cort_cd, day, page_no)
            for r in rows:
                r["_cortNm"] = cort_nm
                key = r.get("dlvrRgstNo") or (r.get("csNo", ""), r.get("prtDlvrrNm", ""))
                all_rows[key] = r
            print(f"  [{cort_nm}] 게시기간 {day:>2}일 - 페이지 {page_no} 조회 ({len(rows)}건, 누적 {len(all_rows)}건)")
            if len(rows) < PAGE_SIZE:
                break
            page_no += 1
            sleep_politely()
        sleep_politely()
    return list(all_rows.values())


def crawl_court_cached(session, cort_cd, cort_nm, cache_dir):
    cache_path = os.path.join(cache_dir, f"{cort_cd}.json")
    if os.path.exists(cache_path):
        age = time.time() - os.path.getmtime(cache_path)
        if age < CACHE_TTL_SECONDS:
            with open(cache_path, "r", encoding="utf-8") as f:
                rows = json.load(f)
            print(f"  [{cort_nm}] 캐시된 결과 사용 ({len(rows)}건, {int(age / 60)}분 전 조회)")
            return rows

    rows = crawl_court_live(session, cort_cd, cort_nm)
    with open(cache_path, "w", encoding="utf-8") as f:
        json.dump(rows, f, ensure_ascii=False)
    return rows


def crawl_court_family(session, family_courts, cache_dir):
    all_rows = []
    for c in family_courts:
        all_rows.extend(crawl_court_cached(session, c["cortCd"], c["cortNm"], cache_dir))
    return all_rows


def filter_by_keywords(rows, keywords):
    matched = []
    for r in rows:
        target = r.get("prtDlvrrNm", "")
        hit = [kw for kw in keywords if kw in target]
        if hit:
            r2 = dict(r)
            r2["_matchedKeywords"] = ", ".join(hit)
            matched.append(r2)
    return matched


def save_excel(rows, output):
    fields = ["_cortNm", "userCsNo", "jdbnCdNm", "prtDlvrrNm", "_matchedKeywords", "dlvrCrtYmd", "pntcTermYmd", "dlvrblNm"]
    labels = ["법원명", "사건번호", "재판부", "공시대상자", "매칭된 키워드", "게시일", "게시만료일", "송달물"]

    wb = Workbook()
    ws = wb.active
    ws.title = "공시송달 조회결과"
    ws.append(labels)
    for r in rows:
        ws.append([r.get(k, "") for k in fields])

    widths = [22, 18, 16, 30, 20, 12, 12, 30]
    for col_idx, width in enumerate(widths, start=1):
        ws.column_dimensions[ws.cell(row=1, column=col_idx).column_letter].width = width

    wb.save(output)


def run():
    parser = argparse.ArgumentParser(description="법원 공시송달 조회 (특정 법원, 공시대상자 회사명 필터)")
    parser.add_argument("--court", help="법원명 (부분 일치, 예: 서울중앙)")
    args = parser.parse_args()

    base_dir = get_base_dir()
    cache_dir = get_cache_dir(base_dir)

    print("=" * 50)
    print(" 법원 공시송달 조회 프로그램")
    print("=" * 50)

    keywords = load_keywords(base_dir)
    if not keywords:
        return
    print(f"\n필터 키워드 {len(keywords)}개 로드 완료: {', '.join(keywords)}")

    print("\n법원 목록을 불러오는 중...")
    session = new_session()
    courts = fetch_court_list(session, base_dir)
    print(f"총 {len(courts)}개 법원 로드 완료.")

    if args.court:
        exact = [c for c in courts if c["cortNm"] == args.court]
        matches = exact if exact else [c for c in courts if args.court in c["cortNm"]]
        if not matches:
            print(f"'{args.court}'에 해당하는 법원을 찾을 수 없습니다.")
            return
        court = matches[0]
    else:
        court = choose_court(courts)

    family = get_court_family(courts, court)
    print(f"\n선택한 법원: {court['cortNm']} ({court['area']})")
    if len(family) > 1:
        print(f"산하 지원/시군법원 {len(family) - 1}곳을 포함하여 총 {len(family)}곳을 조회합니다:")
        for c in family:
            print(f"  - {c['cortNm']}")
    print(f"게시기간 0~21일 전체를 조회합니다. (같은 법원은 {CACHE_TTL_SECONDS // 3600}시간 동안 캐시를 재사용합니다)\n")

    all_rows = crawl_court_family(session, family, cache_dir)
    matched = filter_by_keywords(all_rows, keywords)

    print(f"\n전체 {len(all_rows)}건 중 키워드에 매칭된 공시대상자: {len(matched)}건\n")
    for r in matched:
        print(f"- [{r.get('_cortNm')}] 사건번호={r.get('userCsNo')} 재판부={r.get('jdbnCdNm')} "
              f"공시대상자={r.get('prtDlvrrNm')} 매칭키워드={r.get('_matchedKeywords')} "
              f"게시일={r.get('dlvrCrtYmd')}")

    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_cort_nm = re.sub(r"[\\/:*?\"<>|]", "_", court["cortNm"])
    output_path = os.path.join(base_dir, f"공시송달_{safe_cort_nm}_{ts}.xlsx")
    save_excel(matched, output_path)
    print(f"\n결과를 엑셀 파일로 저장했습니다: {output_path}")


def main():
    try:
        run()
    except requests.exceptions.RequestException:
        print("\n[오류] 법원 사이트에 접속하는 중 문제가 발생했습니다. 인터넷 연결을 확인하고 다시 시도해주세요.")
    except KeyboardInterrupt:
        print("\n사용자에 의해 중단되었습니다.")
    except Exception as e:
        print(f"\n[오류] 예상치 못한 문제가 발생했습니다: {e}")
    finally:
        try:
            input("\n종료하려면 Enter 키를 누르세요...")
        except EOFError:
            pass


if __name__ == "__main__":
    main()
