import io
import os
import requests
import zipfile
from datetime import datetime, timedelta
from typing import List, Optional, Dict, Any
from fastapi import FastAPI, HTTPException, Response
from fastapi.responses import HTMLResponse, JSONResponse
from pydantic import BaseModel, Field
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill

app = FastAPI(
    title="Remittance Portal API",
    description="SamsApi 실시간 연동 및 스펙 스캐너(Inspector) 탑재 포털 API",
    version="4.5.0",
    docs_url="/docs",
    openapi_url="/openapi.json"
)

# ---------------------------------------------------------
# 환경 변수 (SamsApi 접속 정보 및 API Key)
# ---------------------------------------------------------
SAMSAPI_BASE_URL = os.getenv("SAMSAPI_BASE_URL", "http://211.104.10.171:7071")
SAMSAPI_KEY = os.getenv("SAMSAPI_KEY", "Hw-_k-QPRgzolGqkLFIGYLzwqDnep53-wprci845GWw")

# ---------------------------------------------------------
# 지불 정책 Engine
# ---------------------------------------------------------
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
    if not occur_date_str or len(occur_date_str) < 8:
        occur_date_str = datetime.today().strftime("%Y-%m-%d")
        
    is_regular = any(supplier in vendor_name for supplier in REGULAR_SUPPLIERS) or any(vendor_name in supplier for supplier in REGULAR_SUPPLIERS)
    
    if is_regular:
        occur_dt = datetime.strptime(occur_date_str, "%Y-%m-%d")
        base_target_dt = occur_dt + timedelta(days=get_payment_days_by_amount(amount))
    else:
        if request_date_str and len(request_date_str) >= 8:
            base_target_dt = datetime.strptime(request_date_str, "%Y-%m-%d")
        else:
            base_target_dt = datetime.strptime(occur_date_str, "%Y-%m-%d") + timedelta(days=30)
    
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
    vendor_code: Optional[str] = None
    pending_no: Optional[str] = None
    unsettled_only: bool = True
    use_mock: Optional[bool] = False
    custom_api_url: Optional[str] = None

class PaymentDateSaveRequest(BaseModel):
    pending_no: str
    target_payment_date: str

# ---------------------------------------------------------
# Mock 데이터
# ---------------------------------------------------------
def get_mock_pending_data() -> List[dict]:
    items = [
        {"pending_no": "APS202606250008-0001", "account_code": "2002", "account_name": "외상매입금(외화)", "vendor_code": "007003", "vendor_name": "TIME MARINE CO., LTD", "occur_date": "2026-06-03", "acc_date": "2026-06-22", "payment_request_date": "2026-07-03", "currency": "USD", "exchange_rate": 1511.30, "occur_amount": 130.00, "balance_amount": 130.00, "krw_balance": 196469.0, "confirmed_voucher_no": "VC20260622-0045", "edm_documents": [{"doc_id": "EDM-1", "doc_type": "Invoice", "file_name": "TIME_MARINE_INV.pdf", "download_url": "#"}]},
        {"pending_no": "APS202607090021-0002", "account_code": "2001", "account_name": "외상매입금(원화)", "vendor_code": "003143", "vendor_name": "(주)케이씨", "occur_date": "2026-06-03", "acc_date": "2026-06-07", "payment_request_date": "", "currency": "KRW", "exchange_rate": 1.0, "occur_amount": 6711000.00, "balance_amount": 6711000.00, "krw_balance": 6711000.0, "confirmed_voucher_no": "VC20260607-0012", "edm_documents": [{"doc_id": "EDM-2", "doc_type": "세금계산서", "file_name": "KC_Tax.pdf", "download_url": "#"}]},
        {"pending_no": "APS202607010005-0006", "account_code": "2001", "account_name": "외상매입금(원화)", "vendor_code": "003081", "vendor_name": "(주)매일마린", "occur_date": "2026-06-05", "acc_date": "2026-06-16", "payment_request_date": "", "currency": "KRW", "exchange_rate": 1.0, "occur_amount": 22706640.00, "balance_amount": 22706640.00, "krw_balance": 22706640.0, "confirmed_voucher_no": "VC20260616-0089", "edm_documents": [{"doc_id": "EDM-3", "doc_type": "세금계산서", "file_name": "MM_Tax.pdf", "download_url": "#"}]}
    ]
    for item in items:
        auto_date = calculate_payment_date(item["occur_date"], item["payment_request_date"], item["vendor_name"], item["krw_balance"])
        item["auto_payment_date"] = auto_date
        item["scheduled_payment_date"] = auto_date
    return items

# ---------------------------------------------------------
# SamsApi 실시간 연동 Engine
# ---------------------------------------------------------
def fetch_real_pending_data(payload: PendingSearchQuery) -> List[dict]:
    headers = {"X-API-Key": SAMSAPI_KEY, "Content-Type": "application/json"}
    
    if payload.custom_api_url and payload.custom_api_url.strip() != "":
        candidates = [payload.custom_api_url.strip()]
    else:
        candidates = [
            f"{SAMSAPI_BASE_URL}/api/v1/ntstl/list",
            f"{SAMSAPI_BASE_URL}/accounting/api/v1/ntstl/list",
            f"{SAMSAPI_BASE_URL}/api/v1/accounting/ntstl/list"
        ]

    target_dt = payload.end_date if (payload.end_date and payload.end_date != "string") else datetime.today().strftime("%Y-%m-%d")
    req_body = {
        "company_code": "01",
        "target_date": target_dt.replace("-", ""),
        "type_account_code": [payload.account_code] if payload.account_code and payload.account_code not in ["", "string", "ALL"] else [],
        "type_customer_code": [payload.vendor_code] if payload.vendor_code and payload.vendor_code not in ["", "string"] else []
    }

    last_error_msg = ""
    for api_url in candidates:
        try:
            res = requests.post(api_url, headers=headers, params={"page": 1, "pageSize": 2000}, json=req_body, timeout=5)
            
            if res.status_code == 404:
                last_error_msg = f"HTTP 404 (등록된 API 없음) - 주소: {api_url}"
                continue
                
            if res.status_code == 200:
                json_data = res.json()
                if json_data.get("success"):
                    raw_list = json_data.get("data", [])
                    parsed_items = []
                    for raw in raw_list:
                        pending_no = raw.get("not_settled_number", "")
                        vendor_name = raw.get("customer_name", "")
                        
                        def format_date(d_str):
                            if d_str and len(d_str) == 8: return f"{d_str[:4]}-{d_str[4:6]}-{d_str[6:]}"
                            return d_str
                        occur_date = format_date(raw.get("occur_date", ""))
                        due_date = format_date(raw.get("due_date", ""))
                        
                        def parse_float(val):
                            try: return float(val) if val else 0.0
                            except: return 0.0

                        balance_amount = parse_float(raw.get("occur_amount_bal"))
                        krw_balance = parse_float(raw.get("local_amount_bal"))
                        
                        if payload.unsettled_only and balance_amount <= 0: continue
                            
                        auto_date = calculate_payment_date(occur_date, due_date, vendor_name, krw_balance)
                        parsed_items.append({
                            "pending_no": pending_no, "account_code": raw.get("account_code", ""), "account_name": raw.get("account_name", ""),
                            "vendor_code": raw.get("customer_code", ""), "vendor_name": vendor_name, "occur_date": occur_date, "acc_date": format_date(raw.get("from_date", "")),
                            "payment_request_date": due_date, "currency": raw.get("currency_code", "KRW"), "exchange_rate": parse_float(raw.get("occur_exchange_rate")),
                            "occur_amount": parse_float(raw.get("occur_amount_ocr")), "balance_amount": balance_amount, "krw_balance": krw_balance,
                            "auto_payment_date": auto_date, "scheduled_payment_date": auto_date, "confirmed_voucher_no": raw.get("group_settled_number", ""),
                            "edm_documents": [{"doc_id": "EDM-1", "doc_type": "증빙", "file_name": f"{vendor_name}_증빙.pdf", "download_url": "#"}]
                        })
                    return parsed_items
                else:
                    return [{"error_msg": f"API 실패(200): {json_data.get('message')} - URL: {api_url}"}]
            else:
                return [{"error_msg": f"인증/서버 에러 (HTTP {res.status_code}) - URL: {api_url}"}]
                
        except requests.exceptions.Timeout:
            return [{"error_msg": f"연결 시간 초과 ({api_url})"}]
        except requests.exceptions.ConnectionError:
            return [{"error_msg": f"접속 거부 ({api_url})"}]
        except Exception as e:
            last_error_msg = str(e)
            
    return [{"error_msg": last_error_msg or "404 에러. 상단 [등록된 API 경로 전체 스캔] 버튼을 클릭해 실제 경로를 확인하세요."}]

def filter_data(payload: PendingSearchQuery, data: List[dict]) -> List[dict]:
    if data and data[0].get("error_msg"): return data
    if payload.start_date and payload.start_date.strip() not in ["", "string"]: data = [item for item in data if item["occur_date"] >= payload.start_date]
    if payload.end_date and payload.end_date.strip() not in ["", "string"]: data = [item for item in data if item["occur_date"] <= payload.end_date]
    if payload.pay_start_date and payload.pay_start_date.strip() not in ["", "string"]: data = [item for item in data if item["scheduled_payment_date"] >= payload.pay_start_date]
    if payload.pay_end_date and payload.pay_end_date.strip() not in ["", "string"]: data = [item for item in data if item["scheduled_payment_date"] <= payload.pay_end_date]
    if payload.pending_no and payload.pending_no.strip() not in ["", "string"]: data = [item for item in data if item["pending_no"] == payload.pending_no]
    return data

# ---------------------------------------------------------
# 사용자 포털 UI HTML
# ---------------------------------------------------------
@app.get("/portal", response_class=HTMLResponse, tags=["0. 사용자 포털 UI"])
def render_portal_ui():
    return """
    <!DOCTYPE html>
    <html lang="ko">
    <head>
        <meta charset="UTF-8">
        <title>흥아해운 송금 관리 포털</title>
        <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
        <style>
            body { background-color: #f4f6f9; font-family: 'Malgun Gothic', sans-serif; }
            .navbar { background-color: #1a365d; }
            .card { border-radius: 8px; border: none; box-shadow: 0 2px 8px rgba(0,0,0,0.06); }
            .table-header { background-color: #2b4c7e; color: white; }
            .btn-excel { background-color: #1d6f42; color: white; }
            .btn-zip { background-color: #6f42c1; color: white; }
            .btn-mock { background-color: #6c757d; color: white; border-color: #6c757d; }
            .bg-summary { background-color: #fffbeb; }
            .date-input { background-color: #e8f5e9; border: 1px solid #4caf50; color: #1b5e20; font-weight: bold;}
        </style>
    </head>
    <body class="p-3">
        <nav class="navbar navbar-dark px-4 py-3 rounded mb-4 d-flex justify-content-between">
            <span class="navbar-brand mb-0 h1 fw-bold">🚢 흥아해운 미결 지불관리 포털 <span class="badge bg-info text-dark fs-6 ms-2">스펙 Inspector 🔍</span></span>
            <button class="btn btn-warning fw-bold btn-sm" onclick="inspectSamsapiSpec()">🔍 SAMSAPI 등록 경로 전체 스캔</button>
        </nav>
        
        <!-- 디버그 전용 URL 테스트 UI -->
        <div class="card p-3 mb-4 border-primary">
            <h5 class="fw-bold text-primary mb-3">🛠 API 호출 커스텀 테스트</h5>
            <div class="input-group">
                <span class="input-group-text bg-primary text-white fw-bold">API 주소</span>
                <input type="text" class="form-control" id="customApiUrl" value="" placeholder="스캔 결과에 나온 실제 주소를 선택 또는 입력하세요">
                <button class="btn btn-primary fw-bold" onclick="loadPendingData(false)">실시간 API 조회</button>
                <button class="btn btn-mock fw-bold" onclick="loadPendingData(true)">MOCK 복귀</button>
            </div>
            <!-- 스펙 스캔 결과 표시 창 -->
            <div id="specInspectResult" class="mt-3 d-none">
                <div class="alert alert-dark mb-0">
                    <h6 class="fw-bold text-warning">📌 SAMSAPI 서버 등록 경로 스캔 결과:</h6>
                    <ul id="pathList" class="mb-0 font-monospace small"></ul>
                </div>
            </div>
        </div>

        <div class="card p-3 mb-4">
            <h5 class="fw-bold text-secondary mb-3">🔍 미결 조회 조건 (미상계건 대상)</h5>
            <div class="row g-3">
                <div class="col-md-2">
                    <label class="form-label text-secondary fw-bold">계정과목</label>
                    <select class="form-select border-primary" id="accountCode">
                        <option value="">전체보기</option>
                        <option value="2001">2001 (외상매입금-원화)</option>
                        <option value="2002">2002 (외상매입금-외화)</option>
                        <option value="2041">2041 (미지급금)</option>
                    </select>
                </div>
                <div class="col-md-2">
                    <label class="form-label text-secondary fw-bold">발생 시작일</label>
                    <input type="date" class="form-control" id="startDate" value="2026-06-01">
                </div>
                <div class="col-md-2">
                    <label class="form-label text-secondary fw-bold">발생 종료일</label>
                    <input type="date" class="form-control" id="endDate" value="2026-08-31">
                </div>
                <div class="col-md-2">
                    <label class="form-label text-success fw-bold">지불예정 시작일</label>
                    <input type="date" class="form-control border-success" id="payStartDate">
                </div>
                <div class="col-md-2">
                    <label class="form-label text-success fw-bold">지불예정 종료일</label>
                    <input type="date" class="form-control border-success" id="payEndDate">
                </div>
                <div class="col-md-2">
                    <label class="form-label text-secondary fw-bold">거래처코드</label>
                    <input type="text" class="form-control" id="vendorCode" placeholder="코드 입력">
                </div>
            </div>
        </div>

        <div class="card p-3 mb-4">
            <h5 class="fw-bold text-secondary mb-3">📋 미결 지불 대상 목록</h5>
            <div class="table-responsive">
                <table class="table table-hover align-middle border text-center" style="font-size: 0.9rem;">
                    <thead class="table-header">
                        <tr>
                            <th>미결번호</th>
                            <th>계정명</th>
                            <th>거래처명</th>
                            <th>통화</th>
                            <th>환율</th>
                            <th>외화잔액(원화잔액)</th>
                            <th>원화환산액</th>
                            <th>자동산정일</th>
                            <th style="background-color: #1b5e20;">지불예정일 지정(수정가능)</th>
                            <th>저장</th>
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
                    <thead class="bg-light">
                        <tr>
                            <th>통화 (Currency)</th>
                            <th>통화별 합계 금액 (Original Amount)</th>
                            <th>통화별 원화 환산 합계 (KRW Converted)</th>
                        </tr>
                    </thead>
                    <tbody id="summaryTableBody">
                    </tbody>
                    <tfoot>
                        <tr class="bg-summary fw-bold fs-5 text-danger">
                            <td colspan="2" class="text-end pe-4">총 원화 환산 지불 계획 금액 :</td>
                            <td id="grandTotalKrw">0 원</td>
                        </tr>
                    </tfoot>
                </table>
            </div>
        </div>

        <script>
            async function inspectSamsapiSpec() {
                const resBox = document.getElementById("specInspectResult");
                const pathList = document.getElementById("pathList");
                resBox.classList.remove("d-none");
                pathList.innerHTML = "<li>SAMSAPI 서버 스펙 로딩 중...</li>";
                
                try {
                    const res = await fetch('/api/pending/inspect-spec');
                    const json = await res.json();
                    pathList.innerHTML = "";
                    
                    if(json.paths && json.paths.length > 0) {
                        json.paths.forEach(p => {
                            const li = document.createElement("li");
                            const fullUrl = `http://211.104.10.171:7071${p}`;
                            li.innerHTML = `<a href="#" class="text-warning text-decoration-none" onclick="document.getElementById('customApiUrl').value='${fullUrl}'; return false;">${fullUrl}</a>`;
                            pathList.appendChild(li);
                        });
                    } else {
                        pathList.innerHTML = `<li class="text-danger">${json.message || '등록된 API 경로를 찾지 못했습니다.'}</li>`;
                    }
                } catch(e) {
                    pathList.innerHTML = `<li class="text-danger">스캔 오류: ${e.message}</li>`;
                }
            }

            function buildSearchPayload(useMock = false) {
                return {
                    branch_code: "본사",
                    start_date: document.getElementById("startDate").value,
                    end_date: document.getElementById("endDate").value,
                    pay_start_date: document.getElementById("payStartDate").value,
                    pay_end_date: document.getElementById("payEndDate").value,
                    account_code: document.getElementById("accountCode").value,
                    vendor_code: document.getElementById("vendorCode").value,
                    unsettled_only: true,
                    use_mock: useMock,
                    custom_api_url: document.getElementById("customApiUrl") ? document.getElementById("customApiUrl").value : null
                };
            }

            async function loadPendingData(useMock = false) {
                const payload = buildSearchPayload(useMock);
                const tbody = document.getElementById("pendingTableBody");
                const summaryBody = document.getElementById("summaryTableBody");
                
                tbody.innerHTML = '<tr><td colspan="10" class="py-4 text-primary fw-bold">데이터 수신 처리 중입니다...</td></tr>';
                summaryBody.innerHTML = "";
                document.getElementById("grandTotalKrw").innerText = "0 원";
                
                try {
                    const res = await fetch('/api/pending/search-and-schedule', {
                        method: 'POST',
                        headers: {'Content-Type': 'application/json'},
                        body: JSON.stringify(payload)
                    });

                    const text = await res.text();
                    let data = JSON.parse(text);
                    
                    if(data.length > 0 && data[0].error_msg) {
                         tbody.innerHTML = `<tr><td colspan="10" class="py-4 text-danger fw-bold">🚨 ${data[0].error_msg}</td></tr>`;
                         return;
                    }
                    
                    if(data.length === 0) {
                        tbody.innerHTML = '<tr><td colspan="10" class="py-4 text-muted fw-bold">조건에 해당하는 미상계 데이터가 없습니다.</td></tr>';
                        return;
                    }

                    let summary = {}; let grandTotalKrw = 0;
                    tbody.innerHTML = "";

                    data.forEach(item => {
                        let curr = item.currency;
                        if(!summary[curr]) summary[curr] = { original: 0, krw: 0 };
                        summary[curr].original += item.balance_amount;
                        summary[curr].krw += item.krw_balance;
                        grandTotalKrw += item.krw_balance;

                        const tr = document.createElement("tr");
                        tr.innerHTML = `
                            <td class="text-primary fw-bold">${item.pending_no}</td>
                            <td><span class="badge bg-light text-dark border">${item.account_code}</span> ${item.account_name}</td>
                            <td class="fw-bold">${item.vendor_name}</td>
                            <td><span class="badge ${curr === 'KRW' ? 'bg-secondary' : 'bg-danger'}">${curr}</span></td>
                            <td>${Number(item.exchange_rate).toLocaleString()}</td>
                            <td class="text-end pe-3">${Number(item.balance_amount).toLocaleString(undefined, {minimumFractionDigits: curr==='KRW'?0:2})}</td>
                            <td class="fw-bold text-end pe-3">${Number(item.krw_balance).toLocaleString()} 원</td>
                            <td><span class="text-muted">${item.auto_payment_date}</span></td>
                            <td><input type="date" class="form-control form-control-sm text-center date-input" id="date-${item.pending_no}" value="${item.scheduled_payment_date}"></td>
                            <td><button class="btn btn-sm btn-success fw-bold" onclick="saveDate('${item.pending_no}')">저장</button></td>
                        `;
                        tbody.appendChild(tr);
                    });

                    for(const [curr, amounts] of Object.entries(summary)) {
                        const tr = document.createElement("tr");
                        tr.innerHTML = `<td class="fw-bold text-primary">${curr}</td><td class="text-end pe-4">${Number(amounts.original).toLocaleString(undefined, {minimumFractionDigits: curr==='KRW'?0:2})}</td><td class="text-end pe-4 fw-bold">${Number(amounts.krw).toLocaleString()} 원</td>`;
                        summaryBody.appendChild(tr);
                    }
                    document.getElementById("grandTotalKrw").innerText = Number(grandTotalKrw).toLocaleString() + " 원";
                } catch(e) {
                     tbody.innerHTML = `<tr><td colspan="10" class="py-4 text-danger fw-bold">🚨 오류 발생: ${e.message}</td></tr>`;
                }
            }

            async function saveDate(pendingNo) {
                const newDate = document.getElementById(`date-${pendingNo}`).value;
                const res = await fetch('/api/pending/save-payment-date', {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify({ pending_no: pendingNo, target_payment_date: newDate })
                });
                alert((await res.json()).message);
            }

            window.onload = function() {
                loadPendingData(true);
            };
        </script>
    </body>
    </html>
    """

# ---------------------------------------------------------
# SAMSAPI 등록 경로 탐색 전용 엔드포인트
# ---------------------------------------------------------
@app.get("/api/pending/inspect-spec")
def inspect_samsapi_spec():
    headers = {"X-API-Key": SAMSAPI_KEY}
    spec_urls = [
        f"{SAMSAPI_BASE_URL}/openapi.json?domain=accounting",
        f"{SAMSAPI_BASE_URL}/openapi.json?domain=all",
        f"{SAMSAPI_BASE_URL}/openapi.json"
    ]
    
    found_paths = []
    
    for url in spec_urls:
        try:
            res = requests.get(url, headers=headers, timeout=5)
            if res.status_code == 200:
                paths = res.json().get("paths", {})
                for p in paths.keys():
                    found_paths.append(p)
                if found_paths:
                    return {"status": "success", "source_url": url, "paths": found_paths}
        except Exception as e:
             pass
             
    return {"status": "fail", "message": f"SAMSAPI 서버({SAMSAPI_BASE_URL})에서 openapi.json 스펙을 읽지 못했습니다. 키 및 권한을 확인하세요.", "paths": []}

# ---------------------------------------------------------
# API 엔드포인트 구현
# ---------------------------------------------------------
@app.get("/")
def read_root(): return {"status": "online"}

@app.post("/api/pending/search-and-schedule")
def search_and_schedule_pending(payload: PendingSearchQuery):
    try:
        if payload.use_mock:
             raw_data = get_mock_pending_data()
        else:
             raw_data = fetch_real_pending_data(payload)
        return filter_data(payload, raw_data)
    except Exception as e:
        return [{"error_msg": f"백엔드 처리 오류: {str(e)}"}]

@app.post("/api/pending/save-payment-date")
def save_payment_date(payload: PaymentDateSaveRequest):
    return {"status": "success", "message": f"[{payload.pending_no}] 지불예정일이 저장되었습니다."}

@app.post("/api/pending/export-plan-excel")
def export_plan_excel(payload: PendingSearchQuery):
    filtered_data = filter_data(payload, get_mock_pending_data() if payload.use_mock else fetch_real_pending_data(payload))
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "지불계획서"
    ws.append(["지불예정일", "미결번호", "계정코드", "계정명", "거래처명", "통화", "환율", "발생(외화)잔액", "원화환산액", "자동산정일"])
    for cell in ws[1]: cell.fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid"); cell.font = Font(color="FFFFFF", bold=True); cell.alignment = Alignment(horizontal="center")
    
    summary = {}; grand_krw = 0
    for row in filtered_data:
        if row.get("error_msg"): continue
        ws.append([row["scheduled_payment_date"], row["pending_no"], row["account_code"], row["account_name"], row["vendor_name"], row["currency"], row["exchange_rate"], row["balance_amount"], row["krw_balance"], row["auto_payment_date"]])
        curr = row["currency"]
        if curr not in summary: summary[curr] = {"orig": 0, "krw": 0}
        summary[curr]["orig"] += row["balance_amount"]; summary[curr]["krw"] += row["krw_balance"]; grand_krw += row["krw_balance"]

    ws.append([]); ws.append([]); ws.append(["[ 통화별 지불 계획 요약 ]"]); ws.append(["통화", "통화별 합계 금액", "통화별 원화 환산액"])
    for cell in ws[ws.max_row]: cell.fill = PatternFill(start_color="F4B084", end_color="F4B084", fill_type="solid"); cell.font = Font(bold=True); cell.alignment = Alignment(horizontal="center")
    for curr, val in summary.items(): ws.append([curr, val["orig"], val["krw"]])
    ws.append(["총 원화 지불 합계", "", grand_krw])
    for cell in ws[ws.max_row]: cell.font = Font(bold=True, color="C00000"); cell.fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

    stream = io.BytesIO(); wb.save(stream); stream.seek(0)
    return Response(content=stream.getvalue(), media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers={"Content-Disposition": "attachment; filename=Payment_Plan.xlsx"})

@app.post("/api/pending/export-edm-zip")
def export_edm_zip(payload: PendingSearchQuery):
    filtered_data = filter_data(payload, get_mock_pending_data() if payload.use_mock else fetch_real_pending_data(payload))
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
        counter = 1
        for item in filtered_data:
            if item.get("error_msg"): continue
            for doc in item.get("edm_documents", []):
                dummy_content = f"이 파일은 {item['vendor_name']} 업체의 증빙 문서입니다.\n미결번호: {item['pending_no']}\n원화잔액: {item['krw_balance']}원".encode('utf-8')
                new_filename = f"{counter}_{item['vendor_name'].replace('/', '_')}_{doc['doc_type']}.pdf"
                zip_file.writestr(new_filename, dummy_content)
                counter += 1
    zip_buffer.seek(0)
    return Response(content=zip_buffer.getvalue(), media_type="application/zip", headers={"Content-Disposition": f"attachment; filename=EDM_Documents_{datetime.now().strftime('%Y%m%d')}.zip"})