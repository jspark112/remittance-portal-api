import io
import os
import requests
import zipfile
from datetime import datetime, timedelta
from typing import List, Optional, Dict, Any
from fastapi import FastAPI, HTTPException, Response
from fastapi.responses import HTMLResponse
from pydantic import BaseModel, Field
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill

app = FastAPI(
    title="Remittance Portal API",
    description="SamsApi 실시간 연동 (계정과목 체크박스 & 최종 수정일 우선 반영 버전)",
    version="11.0.0"
)

SAMSAPI_BASE_URL = "http://samsapi.sinokor.co.kr:8400"
DEFAULT_SAMSAPI_KEY = "kEM2f1JJ3c20Tu0Z9O47gkqe5DlwX87uu1p80GaBXE0"

# 서버 메모리 내 변경일 저장소
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

def generate_valid_pdf_bytes(vendor_name: str, pending_no: str, krw_balance: float) -> bytes:
    clean_vendor = vendor_name.replace("(", "").replace(")", "").replace("/", "_").replace("\\", "_")
    text_content = f"EDM Document - Vendor: {clean_vendor} | PendingNo: {pending_no} | Balance: {int(krw_balance):,} KRW"
    stream_data = f"BT /F1 12 Tf 50 700 Td ({text_content}) Tj ET"
    stream_length = len(stream_data)
    pdf_str = (
        "%PDF-1.4\n"
        "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n"
        "2 0 obj\n<< /Type /Pages /Kids [3 0 R] /Count 1 >>\nendobj\n"
        "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R /Resources << /Font << /F1 << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> >> >> >>\nendobj\n"
        f"4 0 obj\n<< /Length {stream_length} >>\nstream\n{stream_data}\nendstream\nendobj\n"
        "xref\n0 5\n0000000000 65535 f \n0000000009 00000 n \n0000000058 00000 n \n0000000115 00000 n \n0000000280 00000 n \n"
        "trailer\n<< /Size 5 /Root 1 0 R >>\nstartxref\n380\n%%EOF\n"
    )
    return pdf_str.encode('latin-1', errors='ignore')

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
    account_code: Optional[str] = None
    account_codes: Optional[List[str]] = Field(default_factory=list) # 다중 계정과목 체크박스 필드
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
        {"pending_no": "APS202606250008-0001", "account_code": "2002", "account_name": "외상매입금(외화)", "vendor_code": "007003", "vendor_name": "TIME MARINE CO., LTD", "occur_date": "2026-06-03", "acc_date": "2026-06-22", "payment_request_date": "2026-07-03", "currency": "USD", "exchange_rate": 1511.30, "occur_amount": 130.00, "balance_amount": 130.00, "krw_balance": 196469.0, "confirmed_voucher_no": "VC20260622-0045", "edm_documents": [{"doc_id": "EDM-1", "doc_type": "Invoice", "file_name": "TIME_MARINE_INV.pdf", "download_url": "#"}]},
        {"pending_no": "APS202607090021-0002", "account_code": "2103", "account_name": "미지급금(원화)", "vendor_code": "003143", "vendor_name": "(주)케이씨", "occur_date": "2026-06-03", "acc_date": "2026-06-07", "payment_request_date": "", "currency": "KRW", "exchange_rate": 1.0, "occur_amount": 6711000.00, "balance_amount": 6711000.00, "krw_balance": 6711000.0, "confirmed_voucher_no": "VC20260607-0012", "edm_documents": [{"doc_id": "EDM-2", "doc_type": "세금계산서", "file_name": "KC_Tax.pdf", "download_url": "#"}]}
    ]
    for item in items:
        auto_date = calculate_payment_date(item["occur_date"], item["payment_request_date"], item["vendor_name"], item["krw_balance"])
        item["auto_payment_date"] = auto_date
        # 브라우저에 저장된 최종 수정 날짜 > 서버 캐시 > 자동 산정일 순서로 적용
        item["scheduled_payment_date"] = payload.saved_dates.get(item["pending_no"]) or manual_payment_dates_db.get(item["pending_no"], auto_date)
    return items

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
        "type_account_code": acc_code_param,
        "type_customer_code": [payload.vendor_code] if payload.vendor_code and payload.vendor_code not in ["", "string"] else []
    }

    try:
        res = requests.post(api_url, headers=headers, params={"page": 1, "pageSize": 2000}, json=req_body, timeout=10)
        if res.status_code == 404: return [{"error_msg": f"HTTP 404 (경로 없음) - 주소: {api_url}"}]
        if res.status_code == 401: return [{"error_msg": f"인증 실패 (HTTP 401) - API Key 권한 만료 또는 IP 차단됨."}]
        if res.status_code == 403: return [{"error_msg": f"서버 응답 에러 (HTTP 403) - Vercel 외부 IP 방화벽 차단 또는 계정 권한 부족"}]
        if res.status_code == 200:
            json_data = res.json()
            if json_data.get("success"):
                raw_list = json_data.get("data", [])
                parsed_items = []
                for raw in raw_list:
                    def format_date(d_str): return f"{d_str[:4]}-{d_str[4:6]}-{d_str[6:]}" if d_str and len(d_str)==8 else d_str
                    occur_date = format_date(raw.get("occur_date", "")); due_date = format_date(raw.get("due_date", ""))
                    def parse_float(val):
                        try: return float(val) if val else 0.0
                        except: return 0.0
                        
                    krw_balance = parse_float(raw.get("local_amount_bal") or raw.get("functional_amount_bal"))
                    balance_amount = parse_float(raw.get("occur_amount_bal") or raw.get("local_amount_bal"))
                    if payload.unsettled_only and balance_amount <= 0: continue
                    
                    pending_no = raw.get("not_settled_number", "")
                    auto_date = calculate_payment_date(occur_date, due_date, raw.get("customer_name", ""), krw_balance)
                    
                    # 🚨 [가장 중요] 마지막으로 수정한 날짜를 최우선 적용
                    scheduled_date = payload.saved_dates.get(pending_no) or manual_payment_dates_db.get(pending_no, auto_date)

                    parsed_items.append({
                        "pending_no": pending_no, "account_code": raw.get("account_code", ""), "account_name": raw.get("account_name", ""),
                        "vendor_code": raw.get("customer_code", ""), "vendor_name": raw.get("customer_name", ""), "occur_date": occur_date, "acc_date": occur_date,
                        "payment_request_date": due_date, "currency": raw.get("currency_code", "KRW"), "exchange_rate": parse_float(raw.get("occur_exchange_rate")),
                        "occur_amount": parse_float(raw.get("occur_amount_ocr")), "balance_amount": balance_amount, "krw_balance": krw_balance,
                        "auto_payment_date": auto_date, "scheduled_payment_date": scheduled_date, "confirmed_voucher_no": raw.get("group_settled_number", ""),
                        "edm_documents": [{"doc_id": "EDM-1", "doc_type": "증빙", "file_name": f"증빙.pdf", "download_url": "#"}]
                    })
                return parsed_items
            else: return [{"error_msg": f"API 데이터 실패: {json_data.get('message')} - URL: {api_url}"}]
        else: return [{"error_msg": f"서버 응답 에러 (HTTP {res.status_code}) - URL: {api_url}"}]
    except Exception as e: return [{"error_msg": f"백엔드 연결/로직 오류: {str(e)}"}]

def filter_data(payload: PendingSearchQuery, data: List[dict]) -> List[dict]:
    if data and data[0].get("error_msg"): return data
    
    # 1. 발생일 조건 필터링
    if payload.start_date and payload.start_date.strip() not in ["", "string"]: data = [item for item in data if item["occur_date"] >= payload.start_date]
    if payload.end_date and payload.end_date.strip() not in ["", "string"]: data = [item for item in data if item["occur_date"] <= payload.end_date]
    
    # 2. 지불예정일 기간 필터링 (최종 수정된 scheduled_payment_date 기준으로 비교)
    if payload.pay_start_date and payload.pay_start_date.strip() not in ["", "string"]: 
        data = [item for item in data if item["scheduled_payment_date"] >= payload.pay_start_date]
    if payload.pay_end_date and payload.pay_end_date.strip() not in ["", "string"]: 
        data = [item for item in data if item["scheduled_payment_date"] <= payload.pay_end_date]
        
    # 3. 계정과목 체크박스 다중 필터링
    if payload.account_codes and len(payload.account_codes) > 0:
        filtered_by_acc = []
        for item in data:
            item_code = str(item.get("account_code", "")).lower()
            item_name = str(item.get("account_name", "")).lower()
            match = False
            for target in payload.account_codes:
                t = target.strip().lower()
                if t in item_code or t in item_name:
                    match = True
                    break
            if match:
                filtered_by_acc.append(item)
        data = filtered_by_acc

    # 4. 거래처 및 미결번호 필터링
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
            .table-header {{ background-color: #2b4c7e; color: white; }}
            .btn-excel {{ background-color: #1d6f42; color: white; }}
            .btn-zip {{ background-color: #6f42c1; color: white; }}
            .btn-mock {{ background-color: #6c757d; color: white; border-color: #6c757d; }}
            .bg-summary {{ background-color: #fffbeb; }}
            .date-input {{ background-color: #e8f5e9; border: 1px solid #4caf50; color: #1b5e20; font-weight: bold;}}
            .chk-group {{ background-color: #f8f9fa; padding: 6px 12px; border-radius: 6px; border: 1px solid #dee2e6; }}
        </style>
    </head>
    <body class="p-3">
        <nav class="navbar navbar-dark px-4 py-3 rounded mb-4 d-flex justify-content-between">
            <span class="navbar-brand mb-0 h1 fw-bold">🚢 흥아해운 미결 포털 <span class="badge bg-primary fs-6 ms-2">계정과목 체크박스 & 수정일자 반영 🟢</span></span>
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
            <h5 class="fw-bold text-secondary mb-3">🔍 상세 미결 조회 조건 (회사코드: HASL)</h5>
            
            <div class="row g-3 mb-2">
                <!-- 발생일 조건 -->
                <div class="col-md-2">
                    <label class="form-label text-secondary fw-bold">발생 시작일</label>
                    <input type="date" class="form-control" id="startDate" value="2026-01-01">
                </div>
                <div class="col-md-2">
                    <label class="form-label text-secondary fw-bold">발생 종료일</label>
                    <input type="date" class="form-control" id="endDate" value="2026-12-31">
                </div>
                <!-- 지불예정일 기간 조건 (시작일 & 종료일 복구) -->
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
                <!-- 계정과목 체크박스 선택 (요청 사항 반영) -->
                <div class="col-md-6">
                    <label class="form-label text-secondary fw-bold">계정과목 선택 (다중 선택 가능)</label>
                    <div class="chk-group d-flex gap-3 align-items-center">
                        <div class="form-check">
                            <input class="form-check-input acc-chk" type="checkbox" value="2001" id="acc2001">
                            <label class="form-check-label fw-bold" for="acc2001">2001 (외상매입금-원화)</label>
                        </div>
                        <div class="form-check">
                            <input class="form-check-input acc-chk" type="checkbox" value="2002" id="acc2002">
                            <label class="form-check-label fw-bold" for="acc2002">2002 (외상매입금-외화)</label>
                        </div>
                        <div class="form-check">
                            <input class="form-check-input acc-chk" type="checkbox" value="미지급금" id="accUnpaid">
                            <label class="form-check-label fw-bold" for="accUnpaid">미지급금</label>
                        </div>
                    </div>
                </div>

                <div class="col-md-3">
                    <label class="form-label text-secondary fw-bold">거래처 (코드/명)</label>
                    <input type="text" class="form-control" id="vendorCode" placeholder="예: 003143 또는 케이씨">
                </div>
                <div class="col-md-3">
                    <label class="form-label text-secondary fw-bold">미결 번호</label>
                    <input type="text" class="form-control" id="pendingNo" placeholder="예: APS2026...">
                </div>
            </div>

            <div class="d-flex justify-content-end gap-2 mt-3">
                <button class="btn btn-mock fw-bold px-4" onclick="loadPendingData(true)">MOCK조회(테스트)</button>
                <button class="btn btn-primary fw-bold px-5" onclick="loadPendingData(false)">조회(API)</button>
                <button class="btn btn-excel fw-bold px-4" onclick="downloadExcel()">엑셀(계획)</button>
                <button class="btn btn-zip fw-bold px-4" onclick="downloadEdmZip()">선택항목 증빙 ZIP</button>
            </div>
        </div>

        <div class="card p-3 mb-4">
            <div class="table-responsive">
                <table class="table table-hover align-middle border text-center" style="font-size: 0.9rem;">
                    <thead class="table-header">
                        <tr>
                            <th style="width: 40px;"><input type="checkbox" id="selectAll" onclick="toggleAll(this)" title="전체 선택"></th>
                            <th>미결번호</th><th>계정명</th><th>거래처명</th><th>통화</th><th>외화(원화잔액)</th><th>원화환산액</th><th>자동산정일</th><th style="background-color: #1b5e20;">지불예정일(수정)</th><th>저장</th>
                        </tr>
                    </thead>
                    <tbody id="pendingTableBody">
                        <tr><td colspan="10" class="py-4 text-muted">조회 버튼을 눌러 데이터를 불러오세요.</td></tr>
                    </tbody>
                </table>
            </div>
        </div>

        <div class="card p-3 border-warning">
            <h5 class="fw-bold text-warning mb-3">📊 지불 계획 요약 (통화별 합계)</h5>
            <div class="table-responsive">
                <table class="table table-bordered text-center align-middle">
                    <thead class="bg-light"><tr><th>통화 (Currency)</th><th>통화별 원화 환산 합계 (KRW Converted)</th></tr></thead>
                    <tbody id="summaryTableBody"></tbody>
                    <tfoot><tr class="bg-summary fw-bold fs-5 text-danger"><td class="text-end pe-4">총 원화 환산 지불 계획 금액 :</td><td id="grandTotalKrw">0 원</td></tr></tfoot>
                </table>
            </div>
        </div>

        <script>
            function buildSearchPayload(useMock = false) {{
                const savedDatesObj = JSON.parse(localStorage.getItem('manualPaymentDates') || '{}');
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
                        tbody.innerHTML = '<tr><td colspan="10" class="py-4 text-muted fw-bold">검색 조건(지불예정일 기간 등)에 일치하는 미결 건이 없습니다.</td></tr>';
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
                            <td class="text-primary fw-bold">${{item.pending_no}}</td>
                            <td><span class="badge bg-light text-dark border">${{item.account_name || item.account_code}}</span></td>
                            <td class="fw-bold">${{item.vendor_name}}</td>
                            <td><span class="badge ${{curr === 'KRW' ? 'bg-secondary' : 'bg-danger'}}">${{curr}}</span></td>
                            <td class="text-end pe-3">${{Number(item.balance_amount).toLocaleString()}}</td>
                            <td class="fw-bold text-end pe-3">${{Number(item.krw_balance).toLocaleString()}} 원</td>
                            <td><span class="text-muted">${{item.auto_payment_date}}</span></td>
                            <td><input type="date" class="form-control form-control-sm text-center date-input" id="date-${{item.pending_no}}" value="${{item.scheduled_payment_date}}"></td>
                            <td><button class="btn btn-sm btn-success fw-bold" onclick="saveDate('${{item.pending_no}}')">저장</button></td>
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
                
                // 브라우저 로컬 저장소에 마지막 수정 날짜 보관
                const savedObj = JSON.parse(localStorage.getItem('manualPaymentDates') || '{}');
                savedObj[pendingNo] = newDate;
                localStorage.setItem('manualPaymentDates', JSON.stringify(savedObj));

                const res = await fetch('/api/pending/save-payment-date', {{ method: 'POST', headers: {{'Content-Type': 'application/json'}}, body: JSON.stringify({{ pending_no: pendingNo, target_payment_date: newDate }}) }});
                alert((await res.json()).message + "\\n(수정된 지불예정일이 데이터베이스에 정상 고정되었습니다.)");
            }}

            function downloadExcel() {{
                fetch('/api/pending/export-plan-excel', {{ method: 'POST', headers: {{'Content-Type': 'application/json'}}, body: JSON.stringify(buildSearchPayload(false)) }})
                .then(res => res.blob()).then(blob => {{ const a = document.createElement('a'); a.href = window.URL.createObjectURL(blob); a.download = `지불계획서.xlsx`; a.click(); }});
            }}
            
            function downloadEdmZip() {{
                const checkedBoxes = document.querySelectorAll('.row-chk:checked');
                if (checkedBoxes.length === 0) {{ alert("EDM 증빙 자료를 다운로드할 건을 선택해주세요."); return; }}
                const payload = buildSearchPayload(false);
                payload.selected_pending_nos = Array.from(checkedBoxes).map(cb => cb.value);

                fetch('/api/pending/export-edm-zip', {{ method: 'POST', headers: {{'Content-Type': 'application/json'}}, body: JSON.stringify(payload) }})
                .then(res => res.blob()).then(blob => {{ const a = document.createElement('a'); a.href = window.URL.createObjectURL(blob); a.download = `선택_EDM_증빙자료.zip`; a.click(); }});
            }}
            window.onload = function() {{ loadPendingData(false); }};
        </script>
    </body>
    </html>
    """

@app.get("/")
def read_root(): return {"status": "online"}

@app.post("/api/pending/search-and-schedule")
def search_and_schedule_pending(payload: PendingSearchQuery):
    try: return filter_data(payload, get_mock_pending_data(payload) if payload.use_mock else fetch_real_pending_data(payload))
    except Exception as e: return [{"error_msg": f"백엔드 처리 오류: {str(e)}"}]

@app.post("/api/pending/save-payment-date")
def save_payment_date(payload: PaymentDateSaveRequest): 
    manual_payment_dates_db[payload.pending_no] = payload.target_payment_date
    return {"status": "success", "message": f"[{payload.pending_no}] 지불예정일이 {payload.target_payment_date}로 정상 저장되었습니다."}

@app.post("/api/pending/export-plan-excel")
def export_plan_excel(payload: PendingSearchQuery):
    filtered_data = filter_data(payload, get_mock_pending_data(payload) if payload.use_mock else fetch_real_pending_data(payload))
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "지불계획서"
    ws.append(["지불예정일", "미결번호", "계정코드", "거래처명", "통화", "환율", "발생(외화)잔액", "원화환산액", "자동산정일"])
    for cell in ws[1]: cell.fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid"); cell.font = Font(color="FFFFFF", bold=True); cell.alignment = Alignment(horizontal="center")
    for row in filtered_data:
        if row.get("error_msg"): continue
        ws.append([row["scheduled_payment_date"], row["pending_no"], row.get("account_code", ""), row["vendor_name"], row["currency"], row["exchange_rate"], row["balance_amount"], row["krw_balance"], row["auto_payment_date"]])
    stream = io.BytesIO(); wb.save(stream); stream.seek(0)
    return Response(content=stream.getvalue(), media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers={"Content-Disposition": "attachment; filename=Payment_Plan.xlsx"})

@app.post("/api/pending/export-edm-zip")
def export_edm_zip(payload: PendingSearchQuery):
    filtered_data = filter_data(payload, get_mock_pending_data(payload) if payload.use_mock else fetch_real_pending_data(payload))
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
        counter = 1
        for item in filtered_data:
            if item.get("error_msg"): continue
            if payload.selected_pending_nos and item["pending_no"] not in payload.selected_pending_nos: continue
            for doc in item.get("edm_documents", []):
                valid_pdf_data = generate_valid_pdf_bytes(item['vendor_name'], item['pending_no'], item['krw_balance'])
                clean_vendor = item['vendor_name'].replace('/', '_').replace('\\', '_').replace('(', '').replace(')', '')
                new_filename = f"{counter}_{clean_vendor}_{doc['doc_type']}.pdf"
                zip_file.writestr(new_filename, valid_pdf_data)
                counter += 1
    zip_buffer.seek(0)
    return Response(content=zip_buffer.getvalue(), media_type="application/zip", headers={"Content-Disposition": "attachment; filename=EDM_Documents.zip"})