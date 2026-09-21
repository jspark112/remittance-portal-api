import io
import os
import base64
import requests
import zipfile
from datetime import datetime, timedelta
from typing import List, Optional, Dict, Any
from fastapi import FastAPI, HTTPException, Response
from fastapi.responses import HTMLResponse
from pydantic import BaseModel, Field
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, borders

app = FastAPI(
    title="Remittance Portal API",
    description="SamsApi 실시간 연동 (결재용 송금신청서 엑셀 다운로드 탑재)",
    version="14.0.0"
)

SAMSAPI_BASE_URL = "http://samsapi.sinokor.co.kr:8400"
DEFAULT_SAMSAPI_KEY = "kEM2f1JJ3c20Tu0Z9O47gkqe5DlwX87uu1p80GaBXE0"

manual_payment_dates_db: Dict[str, str] = {}

REGULAR_SUPPLIERS = {
    "한라시스템", "해동구명설비(주)", "한라레벨(주)-한라IMS(주)", "정양엔지니어링",
    "해양금속(주)", "(주)마린테크니컬엔지니어링", "대림엔지니어링(주)", "씨스펙",
    "화진기업(주)", "주식회사 금강항해통신", "(주)해양무선", "주식회사 선양",
    "세연테크", "(주)종합해사청학공장", "(주)이강공사", "(주)매일마린",
    "(주)화일통상", "(주)와이에이치마린", "비아이산업주식회사", "세인베스텍(미우라서비스코리아)",
    "베스트마린코프", "금양엔지니어링", "(주)상봉코포레이션", "(주)디에스케이",
    "주식회사 케이피에스", "(주)해바다", "(주)케이씨", "충무전기공업사", "(주)그린-씨"
}

def get_payment_days_by_amount(amount: float) -> int:
    if amount <= 2_000_000: return 90
    elif amount <= 5_000_000: return 60
    elif amount <= 10_000_000: return 70
    elif amount <= 20_000_000: return 80
    elif amount <= 50_000_000: return 90
    elif amount <= 100_000_000: return 100
    else: return 110

def calculate_payment_date(occur_date_str: str, request_date_str: str, vendor_name: str, amount: float) -> str:
    if not occur_date_str or len(occur_date_str) < 8: occur_date_str = datetime.today().strftime("%Y-%m-%d")
    is_regular = any(supplier in vendor_name for supplier in REGULAR_SUPPLIERS)
    if is_regular: base_target_dt = datetime.strptime(occur_date_str, "%Y-%m-%d") + timedelta(days=get_payment_days_by_amount(amount))
    else:
        if request_date_str and len(request_date_str) >= 8: base_target_dt = datetime.strptime(request_date_str, "%Y-%m-%d")
        else: base_target_dt = datetime.strptime(occur_date_str, "%Y-%m-%d") + timedelta(days=30)
    weekday = base_target_dt.weekday()
    days_to_add_map = {0: 1, 1: 0, 2: 2, 3: 1, 4: 0, 5: 3, 6: 2}
    return (base_target_dt + timedelta(days=days_to_add_map[weekday])).strftime("%Y-%m-%d")

class PendingSearchQuery(BaseModel):
    branch_code: str = Field(default="본사")
    start_date: Optional[str] = None
    end_date: Optional[str] = None
    pay_start_date: Optional[str] = None
    pay_end_date: Optional[str] = None
    account_codes: Optional[List[str]] = Field(default_factory=list)
    vendor_code: Optional[str] = None
    pending_no: Optional[str] = None
    unsettled_only: bool = True
    use_mock: Optional[bool] = False
    api_key: Optional[str] = None
    selected_pending_nos: Optional[List[str]] = None
    saved_dates: Optional[Dict[str, str]] = Field(default_factory=dict)

class PaymentDateSaveRequest(BaseModel):
    pending_no: str
    target_payment_date: str

def get_mock_pending_data(payload: PendingSearchQuery) -> List[dict]:
    items = [
        {"pending_no": "APS202606250008-0001", "account_code": "2002", "account_name": "외상매입금(외화)", "vendor_code": "007003", "vendor_name": "TIME MARINE CO., LTD", "occur_date": "2026-06-03", "acc_date": "2026-06-22", "payment_request_date": "2026-07-03", "currency": "USD", "exchange_rate": 1511.30, "occur_amount": 130.00, "balance_amount": 130.00, "krw_balance": 196469.0, "confirmed_voucher_no": "VC20260622-0045"},
        {"pending_no": "APS202607090021-0002", "account_code": "2103", "account_name": "미지급금(원화)", "vendor_code": "003143", "vendor_name": "(주)케이씨", "occur_date": "2026-06-03", "acc_date": "2026-06-07", "payment_request_date": "", "currency": "KRW", "exchange_rate": 1.0, "occur_amount": 6711000.00, "balance_amount": 6711000.00, "krw_balance": 6711000.0, "confirmed_voucher_no": "VC20260607-0012"},
        {"pending_no": "LN20260901-001", "account_code": "LOAN", "account_name": "운전자금차입금(차입금)", "vendor_code": "001001", "vendor_name": "KB국민은행", "occur_date": "2026-03-01", "acc_date": "2026-03-01", "payment_request_date": "2026-09-30", "currency": "KRW", "exchange_rate": 1.0, "occur_amount": 500000000.0, "balance_amount": 500000000.0, "krw_balance": 500000000.0, "confirmed_voucher_no": "LN-001"}
    ]
    for item in items:
        auto_date = calculate_payment_date(item["occur_date"], item["payment_request_date"], item["vendor_name"], item["krw_balance"])
        item["auto_payment_date"] = auto_date
        item["scheduled_payment_date"] = payload.saved_dates.get(item["pending_no"]) or manual_payment_dates_db.get(item["pending_no"], item["payment_request_date"] or auto_date)
    return items

def fetch_real_loan_data(payload: PendingSearchQuery) -> List[dict]:
    active_key = payload.api_key.strip() if payload.api_key and payload.api_key.strip() else DEFAULT_SAMSAPI_KEY
    headers = {"X-API-Key": active_key, "Authorization": f"Bearer {active_key}", "Content-Type": "application/json"}
    api_url = f"{SAMSAPI_BASE_URL}/api/v1/loan/currentstatelist"
    target_dt = payload.end_date if (payload.end_date and payload.end_date != "string") else datetime.today().strftime("%Y-%m-%d")

    req_body = {
        "company_code": "HASL",
        "target_date": target_dt.replace("-", ""),
        "ploantp": "", "currency_code": "", "pkindtp": "", "pcocd4direct": "", "pcocd4financial": "", "query_type": "", "language_gubun": "KO"
    }

    try:
        res = requests.post(api_url, headers=headers, json=req_body, timeout=10)
        if res.status_code == 200:
            json_data = res.json()
            if json_data.get("success"):
                raw_list = json_data.get("data", [])
                parsed_items = []
                for raw in raw_list:
                    def parse_float(val):
                        try: return float(val) if val else 0.0
                        except: return 0.0

                    balance_amount = parse_float(raw.get("balance_amount"))
                    loan_amount = parse_float(raw.get("loan_amount"))
                    if balance_amount <= 0: continue
                    
                    loan_id = str(raw.get("loand_id") or raw.get("group_settled_number") or "").strip()
                    if not loan_id: continue

                    def format_date(d_str): return f"{d_str[:4]}-{d_str[4:6]}-{d_str[6:]}" if d_str and len(d_str)==8 else d_str
                    from_dt = format_date(raw.get("from_date", ""))
                    to_dt = format_date(raw.get("to_date", ""))
                    
                    vendor_name = raw.get("financial_customer_name") or raw.get("direct_customer_name") or ""
                    if not vendor_name: continue
                    
                    loan_type_name = raw.get("loan_type_name") or raw.get("kind_type_name") or "차입금"
                    if "차입금" not in loan_type_name: loan_type_name += "(차입금)"
                    
                    auto_date = calculate_payment_date(from_dt, to_dt, vendor_name, balance_amount)
                    scheduled_date = payload.saved_dates.get(loan_id) or manual_payment_dates_db.get(loan_id, to_dt or auto_date)

                    parsed_items.append({
                        "pending_no": loan_id, "account_code": "LOAN", "account_name": loan_type_name,
                        "vendor_code": raw.get("financial_customer_code", ""), "vendor_name": vendor_name, 
                        "occur_date": from_dt, "acc_date": from_dt,
                        "payment_request_date": to_dt, "currency": raw.get("currency_code", "KRW"), "exchange_rate": 1.0,
                        "occur_amount": loan_amount, "balance_amount": balance_amount, "krw_balance": balance_amount,
                        "auto_payment_date": auto_date, "scheduled_payment_date": scheduled_date, "confirmed_voucher_no": raw.get("group_settled_number", "")
                    })
                return parsed_items
            else: return []
        else: return []
    except Exception: return []

def fetch_real_pending_data(payload: PendingSearchQuery) -> List[dict]:
    active_key = payload.api_key.strip() if payload.api_key and payload.api_key.strip() else DEFAULT_SAMSAPI_KEY
    headers = {"X-API-Key": active_key, "Authorization": f"Bearer {active_key}", "Content-Type": "application/json"}
    api_url = f"{SAMSAPI_BASE_URL}/api/v1/ntstl/list"
    target_dt = payload.end_date if (payload.end_date and payload.end_date != "string") else datetime.today().strftime("%Y-%m-%d")
    
    acc_code_param = []
    if payload.account_codes:
        for ac in payload.account_codes:
            if ac.isdigit(): acc_code_param.append(ac)

    req_body = {
        "company_code": "HASL", "target_date": target_dt.replace("-", ""),
        "type_account_code": acc_code_param, "type_customer_code": [] 
    }

    try:
        res = requests.post(api_url, headers=headers, params={"page": 1, "pageSize": 2000}, json=req_body, timeout=10)
        if res.status_code == 200:
            json_data = res.json()
            if json_data.get("success"):
                raw_list = json_data.get("data", [])
                parsed_items = []
                for raw in raw_list:
                    def parse_float(val):
                        try: return float(val) if val else 0.0
                        except: return 0.0
                        
                    krw_balance = parse_float(raw.get("local_amount_bal") or raw.get("functional_amount_bal"))
                    balance_amount = parse_float(raw.get("occur_amount_bal") or raw.get("local_amount_bal"))
                    if balance_amount <= 0 or krw_balance <= 0: continue
                    
                    pending_no = str(raw.get("not_settled_number") or "").strip()
                    if not pending_no: continue

                    def format_date(d_str): return f"{d_str[:4]}-{d_str[4:6]}-{d_str[6:]}" if d_str and len(d_str)==8 else d_str
                    occur_date = format_date(raw.get("occur_date", "")); due_date = format_date(raw.get("due_date", ""))
                    
                    vendor_name = raw.get("customer_name") or ""
                    if not vendor_name: continue

                    auto_date = calculate_payment_date(occur_date, due_date, vendor_name, krw_balance)
                    scheduled_date = payload.saved_dates.get(pending_no) or manual_payment_dates_db.get(pending_no, auto_date)

                    parsed_items.append({
                        "pending_no": pending_no, "account_code": raw.get("account_code", ""), "account_name": raw.get("account_name", ""),
                        "vendor_code": raw.get("customer_code", ""), "vendor_name": vendor_name, "occur_date": occur_date, "acc_date": occur_date,
                        "payment_request_date": due_date, "currency": raw.get("currency_code", "KRW"), "exchange_rate": parse_float(raw.get("occur_exchange_rate")),
                        "occur_amount": parse_float(raw.get("occur_amount_ocr")), "balance_amount": balance_amount, "krw_balance": krw_balance,
                        "auto_payment_date": auto_date, "scheduled_payment_date": scheduled_date, "confirmed_voucher_no": raw.get("group_settled_number", "")
                    })
                return parsed_items
            else: return []
        else: return []
    except Exception: return []

def fetch_combined_dataset(payload: PendingSearchQuery) -> List[dict]:
    selected = [a.strip() for a in payload.account_codes] if payload.account_codes else []
    has_loan = ("차입금" in selected) or ("LOAN" in selected) or (len(selected) == 0)
    has_pending = any(a in selected for a in ["2001", "2002", "미지급금"]) or (len(selected) == 0) or (has_loan and len(selected) > 1)

    if payload.use_mock: return get_mock_pending_data(payload)

    combined = []
    if has_pending: combined.extend(fetch_real_pending_data(payload))
    if has_loan: combined.extend(fetch_real_loan_data(payload))
    return combined

def filter_data(payload: PendingSearchQuery, data: List[dict]) -> List[dict]:
    if data and data[0].get("error_msg"): return data
    if payload.start_date and payload.start_date.strip() not in ["", "string"]: data = [item for item in data if item["occur_date"] >= payload.start_date]
    if payload.end_date and payload.end_date.strip() not in ["", "string"]: data = [item for item in data if item["occur_date"] <= payload.end_date]
    if payload.pay_start_date and payload.pay_start_date.strip() not in ["", "string"]: data = [item for item in data if item["scheduled_payment_date"] >= payload.pay_start_date]
    if payload.pay_end_date and payload.pay_end_date.strip() not in ["", "string"]: data = [item for item in data if item["scheduled_payment_date"] <= payload.pay_end_date]
        
    if payload.account_codes and len(payload.account_codes) > 0:
        filtered_by_acc = []
        for item in data:
            item_code = str(item.get("account_code", "")).lower()
            item_name = str(item.get("account_name", "")).lower()
            match = False
            for target in payload.account_codes:
                t = target.strip().lower()
                if t == "차입금" or t == "loan":
                    if "loan" in item_code or "차입" in item_name or "loan" in item_name:
                        match = True; break
                elif t in item_code or t in item_name:
                    match = True; break
            if match: filtered_by_acc.append(item)
        data = filtered_by_acc

    if payload.vendor_code and payload.vendor_code.strip() not in ["", "string"]:
        v_code = payload.vendor_code.strip().lower()
        data = [item for item in data if v_code in item.get("vendor_code", "").lower() or v_code in item.get("vendor_name", "").lower()]
    if payload.pending_no and payload.pending_no.strip() not in ["", "string"]:
        p_no = payload.pending_no.strip().lower()
        data = [item for item in data if p_no in item.get("pending_no", "").lower()]
    return data

@app.get("/portal", response_class=HTMLResponse, tags=["0. 사용자 포털 UI"])
def render_portal_ui():
    return f"""
    <!DOCTYPE html>
    <html lang="ko">
    <head>
        <meta charset="UTF-8">
        <title>흥아해운 송금 관리 포털</title>
        <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
        <style>
            body {{ background-color: #f4f6f9; font-family: 'Malgun Gothic', sans-serif; }}
            .navbar {{ background-color: #1a365d; }}
            .card {{ border-radius: 8px; border: none; box-shadow: 0 2px 8px rgba(0,0,0,0.06); }}
            .table-responsive {{ overflow-x: auto; }}
            .resizable-table {{ table-layout: fixed; width: 100%; min-width: 1200px; border-collapse: collapse; }}
            .resizable-table th {{ position: relative; background-color: #2b4c7e; color: white; padding: 10px; border: 1px solid #dee2e6; user-select: none; }}
            .resizer {{ width: 6px; height: 100%; position: absolute; right: 0; top: 0; cursor: col-resize; z-index: 1; }}
            .resizer:hover, .resizer.resizing {{ background-color: #ffc107; border-right: 2px solid #e0a800; }}
            .btn-excel {{ background-color: #1d6f42; color: white; }}
            .btn-remit {{ background-color: #e67e22; color: white; }} /* 송금신청서 전용 주황색 버튼 */
            .btn-zip {{ background-color: #6f42c1; color: white; }}
            .btn-mock {{ background-color: #6c757d; color: white; border-color: #6c757d; }}
            .bg-summary {{ background-color: #fffbeb; }}
            .date-input {{ background-color: #e8f5e9; border: 1px solid #4caf50; color: #1b5e20; font-weight: bold; width: 100%; }}
            .chk-group {{ background-color: #f8f9fa; padding: 8px 15px; border-radius: 6px; border: 1px solid #dee2e6; box-shadow: inset 0 1px 3px rgba(0,0,0,0.05); }}
        </style>
    </head>
    <body class="p-3">
        <nav class="navbar navbar-dark px-4 py-3 rounded mb-4 d-flex justify-content-between">
            <span class="navbar-brand mb-0 h1 fw-bold">🚢 흥아해운 미결 포털 <span class="badge bg-primary fs-6 ms-2">결재용 송금신청서 탑재 🟢</span></span>
        </nav>
        
        <div class="card p-3 mb-4 border-primary">
            <h5 class="fw-bold text-primary mb-3">🔑 인증 정보 설정</h5>
            <div class="input-group">
                <span class="input-group-text bg-primary text-white fw-bold">SAMSAPI Key</span>
                <input type="text" class="form-control fw-bold text-secondary" id="customApiKey" value="{DEFAULT_SAMSAPI_KEY}" placeholder="API Key 입력">
                <button class="btn btn-warning fw-bold" onclick="loadPendingData(false)">조회(API)</button>
            </div>
        </div>

        <div class="card p-3 mb-4">
            <h5 class="fw-bold text-secondary mb-3">🔍 상세 미결 & 차입금 조회 조건 (회사코드: HASL)</h5>
            <div class="row g-3 mb-2">
                <div class="col-md-2">
                    <label class="form-label text-secondary fw-bold">발생 시작일</label>
                    <input type="date" class="form-control" id="startDate" value="2026-01-01">
                </div>
                <div class="col-md-2">
                    <label class="form-label text-secondary fw-bold">발생 종료일</label>
                    <input type="date" class="form-control" id="endDate" value="2026-12-31">
                </div>
                <div class="col-md-2">
                    <label class="form-label fw-bold text-success">지불예정 시작일</label>
                    <input type="date" class="form-control border-success" id="payStartDate" value="">
                </div>
                <div class="col-md-2">
                    <label class="form-label fw-bold text-success">지불예정 종료일</label>
                    <input type="date" class="form-control border-success" id="payEndDate" value="">
                </div>
            </div>

            <div class="row g-3 mb-2 align-items-end">
                <div class="col-md-6">
                    <label class="form-label text-secondary fw-bold">계정과목 선택 (스위치 - 다중 선택 가능)</label>
                    <div class="chk-group d-flex gap-4 align-items-center">
                        <div class="form-check form-switch">
                            <input class="form-check-input acc-chk" type="checkbox" value="2001" id="acc2001">
                            <label class="form-check-label fw-bold" for="acc2001">2001 (원화)</label>
                        </div>
                        <div class="form-check form-switch">
                            <input class="form-check-input acc-chk" type="checkbox" value="2002" id="acc2002">
                            <label class="form-check-label fw-bold" for="acc2002">2002 (외화)</label>
                        </div>
                        <div class="form-check form-switch">
                            <input class="form-check-input acc-chk" type="checkbox" value="미지급금" id="accUnpaid">
                            <label class="form-check-label fw-bold" for="accUnpaid">미지급금</label>
                        </div>
                        <div class="form-check form-switch">
                            <input class="form-check-input acc-chk" type="checkbox" value="차입금" id="accLoan">
                            <label class="form-check-label fw-bold text-primary" for="accLoan">🏦 차입금</label>
                        </div>
                    </div>
                </div>

                <div class="col-md-3">
                    <label class="form-label text-secondary fw-bold">거래처/금융기관 (코드/명)</label>
                    <input type="text" class="form-control" id="vendorCode" placeholder="예: KB국민은행">
                </div>
                <div class="col-md-3">
                    <label class="form-label text-secondary fw-bold">미결/차입 번호</label>
                    <input type="text" class="form-control" id="pendingNo" placeholder="예: APS2026...">
                </div>
            </div>

            <div class="d-flex justify-content-end gap-2 mt-3">
                <button class="btn btn-mock fw-bold px-4" onclick="loadPendingData(true)">MOCK조회(테스트)</button>
                <button class="btn btn-primary fw-bold px-5" onclick="loadPendingData(false)">조회(API)</button>
                <button class="btn btn-excel fw-bold px-4" onclick="downloadExcel()">리스트(엑셀)</button>
                <!-- 🚨 [추가] 송금신청서 엑셀 다운로드 버튼 -->
                <button class="btn btn-remit fw-bold px-4" onclick="downloadRemittanceForm()">선택건 송금신청서(엑셀)</button>
                <button class="btn btn-zip fw-bold px-4" onclick="downloadEdmZip()">선택항목 증빙 ZIP</button>
            </div>
        </div>

        <div class="card p-3 mb-4">
            <div class="table-responsive">
                <table class="table table-hover align-middle border text-center resizable-table" style="font-size: 0.9rem;" id="pendingTable">
                    <thead>
                        <tr>
                            <th style="width: 50px;"><input type="checkbox" id="selectAll" onclick="toggleAll(this)" title="전체 선택"></th>
                            <th style="width: 160px;">미결/차입번호</th>
                            <th style="width: 140px;">계정/차입구분</th>
                            <th style="width: 180px;">거래처/금융기관</th>
                            <th style="width: 80px;">통화</th>
                            <th style="width: 120px;">외화(원화잔액)</th>
                            <th style="width: 140px;">원화환산액</th>
                            <th style="width: 110px;">자동산정일</th>
                            <th style="width: 150px; background-color: #1b5e20;">지불/만기예정일(수정)</th>
                            <th style="width: 80px;">저장</th>
                        </tr>
                    </thead>
                    <tbody id="pendingTableBody">
                        <tr><td colspan="10" class="py-4 text-muted">조회 버튼을 눌러 데이터를 불러오세요.</td></tr>
                    </tbody>
                </table>
            </div>
        </div>

        <div class="card p-3 border-warning">
            <h5 class="fw-bold text-warning mb-3">📊 지불 / 만기 계획 요약 (통화별 합계)</h5>
            <div class="table-responsive">
                <table class="table table-bordered text-center align-middle">
                    <thead class="bg-light"><tr><th>통화 (Currency)</th><th>통화별 원화 환산 합계 (KRW Converted)</th></tr></thead>
                    <tbody id="summaryTableBody"></tbody>
                    <tfoot><tr class="bg-summary fw-bold fs-5 text-danger"><td class="text-end pe-4">총 원화 환산 지불 계획 금액 :</td><td id="grandTotalKrw">0 원</td></tr></tfoot>
                </table>
            </div>
        </div>

        <script>
            function initResizableTable() {{
                const table = document.getElementById('pendingTable');
                const ths = table.querySelectorAll('th');
                ths.forEach(th => {{
                    if(th.querySelector('.resizer')) return; 
                    const resizer = document.createElement('div');
                    resizer.classList.add('resizer');
                    th.appendChild(resizer);
                    
                    let startX, startWidth;
                    resizer.addEventListener('mousedown', function(e) {{
                        startX = e.pageX; startWidth = th.offsetWidth;
                        resizer.classList.add('resizing');
                        function mouseMoveHandler(e) {{
                            const newWidth = startWidth + (e.pageX - startX);
                            if (newWidth > 40) {{ th.style.width = newWidth + 'px'; th.style.minWidth = newWidth + 'px'; }}
                        }}
                        function mouseUpHandler() {{
                            resizer.classList.remove('resizing');
                            document.removeEventListener('mousemove', mouseMoveHandler);
                            document.removeEventListener('mouseup', mouseUpHandler);
                        }}
                        document.addEventListener('mousemove', mouseMoveHandler);
                        document.addEventListener('mouseup', mouseUpHandler);
                        e.stopPropagation();
                    }});
                }});
            }}

            function buildSearchPayload(useMock = false) {{
                const savedDatesObj = JSON.parse(localStorage.getItem('manualPaymentDates') || '{{}}');
                const selectedAccs = [];
                document.querySelectorAll('.acc-chk:checked').forEach(chk => selectedAccs.push(chk.value));

                return {{
                    branch_code: "본사",
                    start_date: document.getElementById("startDate").value,
                    end_date: document.getElementById("endDate").value,
                    pay_start_date: document.getElementById("payStartDate").value,
                    pay_end_date: document.getElementById("payEndDate").value,
                    account_codes: selectedAccs,
                    vendor_code: document.getElementById("vendorCode").value,
                    pending_no: document.getElementById("pendingNo").value,
                    unsettled_only: true,
                    use_mock: useMock,
                    api_key: document.getElementById("customApiKey").value,
                    saved_dates: savedDatesObj
                }};
            }}

            function toggleAll(source) {{
                const checkboxes = document.querySelectorAll('.row-chk');
                checkboxes.forEach(chk => chk.checked = source.checked);
            }}

            async function loadPendingData(useMock = false) {{
                const payload = buildSearchPayload(useMock);
                const tbody = document.getElementById("pendingTableBody");
                const summaryBody = document.getElementById("summaryTableBody");
                
                tbody.innerHTML = '<tr><td colspan="10" class="py-4 text-primary fw-bold">데이터 조회 중입니다...</td></tr>';
                summaryBody.innerHTML = ""; document.getElementById("grandTotalKrw").innerText = "0 원";
                
                try {{
                    const res = await fetch('/api/pending/search-and-schedule', {{
                        method: 'POST', headers: {{'Content-Type': 'application/json'}}, body: JSON.stringify(payload)
                    }});
                    const text = await res.text();
                    let data = JSON.parse(text);
                    
                    if(data.length > 0 && data[0].error_msg) {{
                         tbody.innerHTML = `<tr><td colspan="10" class="py-4 text-danger fw-bold fs-5">${{data[0].error_msg}}</td></tr>`;
                         return;
                    }}
                    if(data.length === 0) {{
                        tbody.innerHTML = '<tr><td colspan="10" class="py-4 text-muted fw-bold">검색 조건에 일치하는 데이터가 없습니다.</td></tr>';
                        return;
                    }}

                    let summary = {{}}; let grandTotalKrw = 0; tbody.innerHTML = "";
                    data.forEach(item => {{
                        let curr = item.currency;
                        if(!summary[curr]) summary[curr] = {{ krw: 0 }};
                        summary[curr].krw += item.krw_balance;
                        grandTotalKrw += item.krw_balance;

                        const tr = document.createElement("tr");
                        tr.innerHTML = `
                            <td><input type="checkbox" class="row-chk form-check-input" value="${{item.pending_no}}"></td>
                            <td class="text-primary fw-bold text-break">${{item.pending_no}}</td>
                            <td class="text-break"><span class="badge bg-light text-dark border">${{item.account_name || item.account_code}}</span></td>
                            <td class="fw-bold text-break">${{item.vendor_name}}</td>
                            <td><span class="badge ${{curr === 'KRW' ? 'bg-secondary' : 'bg-danger'}}">${{curr}}</span></td>
                            <td class="text-end pe-3 text-break">${{Number(item.balance_amount).toLocaleString()}}</td>
                            <td class="fw-bold text-end pe-3 text-break">${{Number(item.krw_balance).toLocaleString()}} 원</td>
                            <td><span class="text-muted">${{item.auto_payment_date}}</span></td>
                            <td><input type="date" class="form-control form-control-sm text-center date-input" id="date-${{item.pending_no}}" value="${{item.scheduled_payment_date}}"></td>
                            <td><button class="btn btn-sm btn-success fw-bold w-100" onclick="saveDate('${{item.pending_no}}')">저장</button></td>
                        `;
                        tbody.appendChild(tr);
                    }});
                    for(const [curr, amounts] of Object.entries(summary)) {{
                        const tr = document.createElement("tr");
                        tr.innerHTML = `<td class="fw-bold text-primary" colspan="2">${{curr}}</td><td class="text-end pe-4 fw-bold" colspan="8">${{Number(amounts.krw).toLocaleString()}} 원</td>`;
                        summaryBody.appendChild(tr);
                    }}
                    document.getElementById("grandTotalKrw").innerText = Number(grandTotalKrw).toLocaleString() + " 원";
                }} catch(e) {{ tbody.innerHTML = `<tr><td colspan="10" class="py-4 text-danger fw-bold">🚨 오류 발생: ${{e.message}}</td></tr>`; }}
            }}
            
            async function saveDate(pendingNo) {{
                const newDate = document.getElementById(`date-${{pendingNo}}`).value;
                const savedObj = JSON.parse(localStorage.getItem('manualPaymentDates') || '{{}}');
                savedObj[pendingNo] = newDate;
                localStorage.setItem('manualPaymentDates', JSON.stringify(savedObj));

                const res = await fetch('/api/pending/save-payment-date', {{ method: 'POST', headers: {{'Content-Type': 'application/json'}}, body: JSON.stringify({{ pending_no: pendingNo, target_payment_date: newDate }}) }});
                alert((await res.json()).message + "\\n(수정된 지불예정일이 정상 고정되었습니다.)");
            }}

            function downloadExcel() {{
                fetch('/api/pending/export-plan-excel', {{ method: 'POST', headers: {{'Content-Type': 'application/json'}}, body: JSON.stringify(buildSearchPayload(false)) }})
                .then(res => res.blob()).then(blob => {{ const a = document.createElement('a'); a.href = window.URL.createObjectURL(blob); a.download = `전체_지불_계획리스트.xlsx`; a.click(); }});
            }}

            // 🚨 [추가] 송금신청서 다운로드 JS 로직
            function downloadRemittanceForm() {{
                const checkedBoxes = document.querySelectorAll('.row-chk:checked');
                if (checkedBoxes.length === 0) {{ alert("송금신청서를 작성할 건을 체크박스로 먼저 선택해주세요."); return; }}
                
                const payload = buildSearchPayload(false);
                payload.selected_pending_nos = Array.from(checkedBoxes).map(cb => cb.value);

                fetch('/api/pending/export-remittance-form', {{ method: 'POST', headers: {{'Content-Type': 'application/json'}}, body: JSON.stringify(payload) }})
                .then(res => res.blob()).then(blob => {{ const a = document.createElement('a'); a.href = window.URL.createObjectURL(blob); a.download = `송금신청서(기안용).xlsx`; a.click(); }});
            }}
            
            function downloadEdmZip() {{
                const checkedBoxes = document.querySelectorAll('.row-chk:checked');
                if (checkedBoxes.length === 0) {{ alert("증빙 자료를 다운로드할 건을 선택해주세요."); return; }}
                const payload = buildSearchPayload(false);
                payload.selected_pending_nos = Array.from(checkedBoxes).map(cb => cb.value);

                const btn = document.querySelector('.btn-zip');
                const originalText = btn.innerText;
                btn.innerText = "원본 다운로드 중...";
                btn.disabled = true;

                fetch('/api/pending/export-edm-zip', {{ method: 'POST', headers: {{'Content-Type': 'application/json'}}, body: JSON.stringify(payload) }})
                .then(res => res.blob()).then(blob => {{ 
                    const a = document.createElement('a'); a.href = window.URL.createObjectURL(blob); a.download = `실제_EDM_원본증빙.zip`; a.click(); 
                    btn.innerText = originalText; btn.disabled = false;
                }}).catch(() => {{ 
                    alert("다운로드 중 오류가 발생했습니다."); 
                    btn.innerText = originalText; btn.disabled = false;
                }});
            }}
            
            window.onload = function() {{ 
                initResizableTable(); 
                loadPendingData(false); 
            }};
        </script>
    </body>
    </html>
    """

@app.get("/")
def read_root(): return {"status": "online"}

@app.post("/api/pending/search-and-schedule")
def search_and_schedule_pending(payload: PendingSearchQuery):
    try: 
        raw_data = fetch_combined_dataset(payload)
        return filter_data(payload, raw_data)
    except Exception as e: return [{"error_msg": f"백엔드 처리 오류: {str(e)}"}]

@app.post("/api/pending/save-payment-date")
def save_payment_date(payload: PaymentDateSaveRequest): 
    manual_payment_dates_db[payload.pending_no] = payload.target_payment_date
    return {"status": "success", "message": f"[{payload.pending_no}] 지불예정일이 {payload.target_payment_date}로 정상 저장되었습니다."}

@app.post("/api/pending/export-plan-excel")
def export_plan_excel(payload: PendingSearchQuery):
    filtered_data = filter_data(payload, fetch_combined_dataset(payload))
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "전체_리스트"
    ws.append(["지불/만기예정일", "미결/차입번호", "계정/차입구분", "거래처/금융기관", "통화", "환율", "발생/차입금액", "원화환산액", "자동산정일"])
    for cell in ws[1]: cell.fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid"); cell.font = Font(color="FFFFFF", bold=True); cell.alignment = Alignment(horizontal="center")
    for row in filtered_data:
        if row.get("error_msg"): continue
        ws.append([row["scheduled_payment_date"], row["pending_no"], row.get("account_name", ""), row["vendor_name"], row["currency"], row["exchange_rate"], row["balance_amount"], row["krw_balance"], row["auto_payment_date"]])
    stream = io.BytesIO(); wb.save(stream); stream.seek(0)
    return Response(content=stream.getvalue(), media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers={"Content-Disposition": "attachment; filename=List.xlsx"})

# 🚨 [신규 추가] 송금신청서 엑셀 다운로드 API
@app.post("/api/pending/export-remittance-form")
def export_remittance_form(payload: PendingSearchQuery):
    filtered_data = filter_data(payload, fetch_combined_dataset(payload))
    
    # 체크박스로 선택한 건만 필터링
    if payload.selected_pending_nos:
        filtered_data = [item for item in filtered_data if item["pending_no"] in payload.selected_pending_nos]

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "송금신청서"

    # 스타일 세팅
    title_font = Font(size=18, bold=True)
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    center_align = Alignment(horizontal="center", vertical="center")
    right_align = Alignment(horizontal="right", vertical="center")
    thin_border = borders.Border(
        left=borders.Side(style='thin'), right=borders.Side(style='thin'),
        top=borders.Side(style='thin'), bottom=borders.Side(style='thin')
    )

    # 1. 엑셀 상단 제목 (병합)
    ws.merge_cells("A1:I2")
    title_cell = ws["A1"]
    title_cell.value = "송 금 신 청 서 (지 급 품 의)"
    title_cell.font = title_font
    title_cell.alignment = center_align

    # 2. 작성일자
    ws["A4"] = f"작성일자 : {datetime.today().strftime('%Y-%m-%d')}"
    ws["A4"].font = Font(bold=True)

    # 3. 테이블 헤더
    headers = ["순번", "지불예정일", "수취인(거래처명)", "통화", "청구금액", "원화금액(KRW)", "은행명(수기)", "계좌번호(수기)", "적요(미결번호)"]
    ws.append([]) # 5행 공란
    ws.append(headers) # 6행 헤더
    for col_idx, cell in enumerate(ws[6], 1):
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center_align
        cell.border = thin_border

    # 4. 데이터 삽입
    total_krw = 0
    for i, row in enumerate(filtered_data, 1):
        if row.get("error_msg"): continue
        
        krw_bal = row.get("krw_balance", 0)
        total_krw += krw_bal
        
        data_row = [
            i,
            row.get("scheduled_payment_date", ""),
            row.get("vendor_name", ""),
            row.get("currency", ""),
            row.get("balance_amount", 0),
            krw_bal,
            "", # 은행명 (엑셀에서 수기 입력하도록 공란 처리)
            "", # 계좌번호 (엑셀에서 수기 입력하도록 공란 처리)
            row.get("pending_no", "")
        ]
        ws.append(data_row)
        
        for cell in ws[ws.max_row]:
            cell.border = thin_border
            if cell.column in [1, 2, 4, 7, 8, 9]: cell.alignment = center_align # 텍스트 중앙정렬
            elif cell.column in [5, 6]: # 금액 우측정렬 및 콤마
                cell.number_format = '#,##0'
                cell.alignment = right_align

    # 5. 합계 행
    ws.append(["합계", "", "", "", "", total_krw, "", "", ""])
    total_row = ws.max_row
    ws.merge_cells(start_row=total_row, start_column=1, end_row=total_row, end_column=5)
    
    sum_cell = ws.cell(row=total_row, column=1)
    sum_cell.value = "합 계 (Total)"
    sum_cell.alignment = center_align
    sum_cell.font = Font(bold=True)
    
    for cell in ws[total_row]:
        cell.border = thin_border
        cell.fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
        if cell.column == 6:
            cell.number_format = '#,##0'
            cell.font = Font(bold=True, color="FF0000") # 합계는 빨간색 볼드체

    # 6. 컬럼 너비 조정
    widths = {'A': 6, 'B': 14, 'C': 30, 'D': 8, 'E': 15, 'F': 18, 'G': 15, 'H': 25, 'I': 25}
    for col, width in widths.items():
        ws.column_dimensions[col].width = width

    stream = io.BytesIO()
    wb.save(stream)
    stream.seek(0)
    return Response(content=stream.getvalue(), media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers={"Content-Disposition": "attachment; filename=Remittance_Form.xlsx"})

@app.post("/api/pending/export-edm-zip")
def export_edm_zip(payload: PendingSearchQuery):
    filtered_data = filter_data(payload, fetch_combined_dataset(payload))
    zip_buffer = io.BytesIO()
    
    active_key = payload.api_key.strip() if payload.api_key and payload.api_key.strip() else DEFAULT_SAMSAPI_KEY
    headers = {"X-API-Key": active_key, "Authorization": f"Bearer {active_key}", "Content-Type": "application/json"}
    edm_api_url = f"{SAMSAPI_BASE_URL}/api/v1/edm/list"
    
    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
        counter = 1
        for item in filtered_data:
            if item.get("error_msg"): continue
            if payload.selected_pending_nos and item["pending_no"] not in payload.selected_pending_nos: continue
            
            if payload.use_mock:
                dummy_pdf = base64.b64decode("JVBERi0xLjQKJfbkwzEAMCBvYmoKPDwgL1R5cGUgL0NhdGFsb2cgL1BhZ2VzIDIgMCBSID4+CmVuZG9iaiAyIDAgb2JqCjw8IC9UeXBlIC9QYWdlcyAvS2lkcyBbMyAwIFJdIC9Db3VudCAxID4+CmVuZG9iaiAzIDAgb2JqCjw8IC9UeXBlIC9QYWdlIC9QYXJlbnQgMiAwIFIgL01lZGlhQm94IFswIDAgNjEyIDc5Ml0gL0NvbnRlbnRzIDQgMCBSIC9SZXNvdXJjZXMgPDwgL0ZvbnQgPDwgL0YxIDw8IC9UeXBlIC9Gb250IC9TdWJ0eXBlIC9UeXBlMSAvQmFzZUZvbnQgL0hlbHZldGljYSA+PiA+PiA+PiA+PgplbmRvYmogNCAwIG9iago8PCAvTGVuZ3RoIDUzID4+CnN0cmVhbQpCVCAvRjEgMTIgVGYgNTAgNzAwIFRkIChFRE0gU2FtcGxlIERvY3VtZW50KSBUaiBFVAplbmRzdHJlYW0KZW5kb2JqCnhyZWYKMCA1CjAwMDAwMDAwMDAgNjU1MzUgZiAKMDAwMDAwMDAxNSA0MDAwMCBuIAowMDAwMDAwMDY4IDAwMDAwIG4gCjAwMDAwMDAxMjUgMDAwMDAgbiAKMDAwMDAwMDI3MyAwMDAwMCBuIAp0cmFpbGVyCjw8IC9TaXplIDUgL1Jvb3QgMSAwIFIgPj4Kc3RhcnR4cmVmCjM3NgolJUVPRgo=")
                clean_vendor = item['vendor_name'].replace('/', '_').replace('\\', '_').replace('(', '').replace(')', '')
                zip_file.writestr(f"{counter}_{clean_vendor}_MOCK_FILE.pdf", dummy_pdf)
                counter += 1
                continue
                
            journal_number = item.get("pending_no", "")
            confirmed_voucher = item.get("confirmed_voucher_no", "")
            clean_vendor = item.get('vendor_name', '알수없음').replace('/', '_').replace('\\', '_').replace('(', '').replace(')', '')
            
            req_body = {"company_code": "HASL", "journal_number": journal_number, "language_gubun": "KO"}
            
            try:
                res = requests.post(edm_api_url, headers=headers, json=req_body, timeout=10)
                edm_list = res.json().get("data", []) if res.status_code == 200 and res.json().get("success") else []
                
                if not edm_list and confirmed_voucher and confirmed_voucher != journal_number:
                    req_body["journal_number"] = confirmed_voucher
                    res2 = requests.post(edm_api_url, headers=headers, json=req_body, timeout=10)
                    if res2.status_code == 200 and res2.json().get("success"):
                        edm_list = res2.json().get("data", [])

                if not edm_list:
                    zip_file.writestr(f"{counter}_{clean_vendor}_증빙없음.txt", f"SAMSAPI 서버에 {item['pending_no']} 에 대한 증빙 파일이 없습니다.".encode('utf-8'))
                else:
                    for edm in edm_list:
                        real_filename = edm.get("filename", f"document_{counter}.pdf")
                        download_url = edm.get("downloadurl", "")
                        
                        if download_url:
                            if download_url.startswith("/"):
                                download_url = SAMSAPI_BASE_URL + download_url
                                
                            file_res = requests.get(download_url, headers=headers, timeout=20)
                            if file_res.status_code == 200:
                                new_filename = f"{counter}_{clean_vendor}_{real_filename}"
                                zip_file.writestr(new_filename, file_res.content)
                        counter += 1
            except Exception as e:
                zip_file.writestr(f"{counter}_{clean_vendor}_다운로드오류.txt", f"EDM API 연결 실패: {str(e)}".encode('utf-8'))
                counter += 1
                
    zip_buffer.seek(0)
    return Response(content=zip_buffer.getvalue(), media_type="application/zip", headers={"Content-Disposition": "attachment; filename=Real_EDM_Documents.zip"})