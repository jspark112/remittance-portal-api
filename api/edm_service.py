"""미결번호 → 확정전표번호 → EDM 증빙 조회/다운로드 (SAMSAPI)

흐름
1. ntstl/ntstlinfo : 미결번호로 발생일자·거래처코드 조회
2. jrn/list        : 발생일 이후 확정된 해당 거래처 전표 중 not_settled_number 가 미결번호와 '정확히' 같은 행 검색
3. edm/list        : 찾은 journal_number 로 EDM 증빙 목록 조회 → downloadurl 로 파일 다운로드

주의: jrn/list 의 fixed_journal_number 는 EDM 증빙과 연결되지 않는다. 반드시 journal_number 를 사용.
"""
import time
from datetime import datetime, timedelta
from typing import Dict, List, Optional

import requests

SAMSAPI_BASE_URL = "http://samsapi.sinokor.co.kr:8400"
DEFAULT_SAMSAPI_KEY = "kEM2f1JJ3c20Tu0Z9O47gkqe5DlwX87uu1p80GaBXE0"
COMPANY_CODE = "HASL"

# 확정일 검색 범위: 발생일부터 우선 60일, 못 찾으면 400일까지 확장
SEARCH_WINDOWS_DAYS = (60, 400)


class SamsApi:
    def __init__(self, api_key: Optional[str] = None, base_url: str = SAMSAPI_BASE_URL):
        key = (api_key or "").strip() or DEFAULT_SAMSAPI_KEY
        self.base_url = base_url
        self.session = requests.Session()
        self.session.headers.update({"X-API-Key": key, "Content-Type": "application/json"})

    def _request(self, method: str, url: str, retries: int = 4, **kwargs) -> requests.Response:
        # SAMSAPI 는 연속 호출 시 429(Too Many Requests)를 반환하므로 대기 후 재시도
        for attempt in range(retries + 1):
            res = self.session.request(method, url, **kwargs)
            if res.status_code != 429 or attempt == retries:
                return res
            time.sleep(1.5 * (attempt + 1))
        return res

    def post(self, path: str, body: dict, timeout: int = 30, **params) -> List[dict]:
        res = self._request("POST", self.base_url + path, json=body, params=params or None, timeout=timeout)
        if res.status_code != 200:
            raise RuntimeError(f"{path} HTTP {res.status_code}: {res.text[:200]}")
        js = res.json()
        if not js.get("success"):
            raise RuntimeError(f"{path} 실패: {js.get('message')}")
        return js.get("data") or []

    # 1) 미결 건별 조회
    def get_pending_info(self, pending_nos: List[str]) -> Dict[str, dict]:
        rows = self.post("/api/v1/ntstl/ntstlinfo",
                         {"company_code": COMPANY_CODE, "type_not_settled_number": pending_nos},
                         page=1, pageSize=2000)
        return {str(r.get("not_settled_number") or "").strip(): r for r in rows}

    # 2) 미결번호와 정확히 일치하는 확정전표 행 검색
    def find_journal(self, pending_no: str, occur_date: str, customer_code: str) -> Optional[dict]:
        start = datetime.strptime(occur_date, "%Y%m%d")
        today = datetime.today()
        for days in SEARCH_WINDOWS_DAYS:
            body = {
                "company_code": COMPANY_CODE,
                "from_date": start.strftime("%Y%m%d"),
                "to_date": min(start + timedelta(days=days), today).strftime("%Y%m%d"),
                "type_customer_code": [customer_code] if customer_code else [],
                "type_account_code": [],
                "type_journal_type": [],
            }
            rows = self.post("/api/v1/jrn/list", body, page=1, pageSize=2000)
            hits = [r for r in rows if str(r.get("not_settled_number") or "").strip() == pending_no]
            if hits:
                # 같은 미결번호가 이후 상계전표에도 찍힐 수 있으므로 가장 먼저 확정된 (발생)전표를 사용
                hits.sort(key=lambda r: (str(r.get("fixed_journal_date") or ""), str(r.get("journal_number") or "")))
                return hits[0]
            if start + timedelta(days=days) >= today:
                break
        return None

    # 3) EDM 증빙 목록
    def list_edm(self, journal_number: str) -> List[dict]:
        return self.post("/api/v1/edm/list",
                         {"company_code": COMPANY_CODE, "journal_number": journal_number, "language_gubun": "L"})

    def download(self, url: str) -> bytes:
        if url.startswith("/"):
            url = self.base_url + url
        res = self._request("GET", url, timeout=60)
        if res.status_code != 200:
            raise RuntimeError(f"다운로드 실패 HTTP {res.status_code}: {url}")
        return res.content


def resolve_pending(api: SamsApi, pending_nos: List[str]) -> List[dict]:
    """미결번호 목록 → [{pending_no, vendor_name, journal_number, edm_files, error}]"""
    pending_nos = [p.strip() for p in pending_nos if p and p.strip()]
    results = []
    try:
        infos = api.get_pending_info(pending_nos)
    except Exception as e:
        return [{"pending_no": p, "error": f"미결 조회 실패: {e}"} for p in pending_nos]

    for p_no in pending_nos:
        item = {"pending_no": p_no, "vendor_name": "", "journal_number": None, "edm_files": [], "error": None}
        info = infos.get(p_no)
        if not info:
            item["error"] = "SAMSAPI 에서 미결번호를 찾을 수 없음"
            results.append(item)
            continue
        item["vendor_name"] = str(info.get("customer_name") or "").strip()
        try:
            jrn = api.find_journal(p_no, str(info.get("occur_date") or ""), str(info.get("customer_code") or ""))
            if not jrn:
                item["error"] = "확정전표에서 이 미결번호를 가진 전표를 찾지 못함"
            else:
                item["journal_number"] = str(jrn.get("journal_number") or "").strip()
                item["journal_date"] = jrn.get("fixed_journal_date")
                item["edm_files"] = api.list_edm(item["journal_number"])
        except Exception as e:
            item["error"] = str(e)
        results.append(item)
    return results


def safe_name(text: str) -> str:
    # 윈도우는 끝에 마침표/공백이 있는 폴더명을 만들지 못함 (예: "... LTD.")
    return "".join("_" if c in '\\/:*?"<>|' else c for c in text).strip().rstrip(". ") or "unknown"
