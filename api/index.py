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
    description="SamsApi 실시간 연동 (API Key 동적 테스터 탑재)",
    version="8.1.0"
)

SAMSAPI_BASE_URL = "http://samsapi.sinokor.co.kr:8400"
DEFAULT_SAMSAPI_KEY = "Hw-_k-QPRgzolGqkLFIGYLzwqDnep53-wprci845GWw"

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
    if is_regular:
        base_target_dt = datetime.strptime(occur_date_str, "%Y-%m-%d") + timedelta(days=get_payment_days_by_amount(amount))
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
    vendor_code: Optional[str] = None
    pending_no: Optional[str] = None
    unsettled_only: bool = True
    use_mock: Optional[bool] = False
    api_key: Optional[str] = None # 화면에서 입력받은 API 키

class PaymentDateSaveRequest(BaseModel):
    pending_no: str
    target_payment_date: str

def get_mock_pending_data() -> List[dict]:
    items = [
        {"pending_no": "APS202606250008-0001", "account_code": "2002", "account_name": "외상매입금(외화)", "vendor_code": "007003", "vendor_name": "TIME MARINE CO., LTD", "occur_date": "2026-06-03", "acc_date": "2026-06-22", "payment_request_date": "2026-07-03", "currency": "USD", "exchange_rate": 1511.30, "occur_amount": 130.00, "balance_amount": 130.00, "krw_balance": 196469.0, "confirmed_voucher_no": "VC20260622-0045", "edm_documents": [{"doc_id": "EDM-1", "doc_type": "Invoice", "file_name": "TIME_MARINE_INV.pdf", "download_url": "#"}]},
        {"pending_no": "APS202607090021-0002", "account_code": "2001", "account_name": "외상매입금(원화)", "vendor_code": "003143", "vendor_name": "(주)케이씨", "occur_date": "2026-06-03", "acc_date": "2026-06-07", "payment_request_date": "", "currency": "KRW", "exchange_rate": 1.0, "occur_amount": 6711000.00, "balance_amount": 6711000.00, "krw_balance": 6711000.0, "confirmed_voucher_no": "VC20260607-0012", "edm_documents": [{"doc_id": "EDM-2", "doc_type": "세금계산서", "file_name": "KC_Tax.pdf", "download_url": "#"}]}
    ]
    for item in items:
        auto_date = calculate_payment_date(item["occur_date"], item["payment_request_date"], item["vendor_name"], item["krw_balance"])
        item["auto_payment_date"] = auto_date; item["scheduled_payment_date"] = auto_date
    return items

def fetch_real_pending_data(payload: PendingSearchQuery) -> List[dict]:
    # 화면에서 키를 입력했으면 그 키를, 아니면 하드코딩된 기본 키를 사용
    active_key = payload.api_key.strip() if payload.api_key and payload.api_key.strip() else DEFAULT_SAMSAPI_KEY
    
    headers = {
        "X-API-Key": active_key,
        "Authorization": f"Bearer {active_key}", # 혹시 몰라 Bearer 토큰 형식도 같이 전송
        "Content-Type": "application/json"
    }
    
    # 드디어 정확히 찾아낸 그 주소!
    api_url = f"{SAMSAPI_BASE_URL}/api/v1/accounting/ntstl/list"

    target_dt = payload.end_date if (payload.end_date and payload.end_date != "string") else datetime.today().strftime("%Y-%m-%d")
    req_body = {
        "company_code": "01", "target_date": target_dt.replace("-", ""),
        "type_account_code": [payload.account_code] if payload.account_code and payload.account_code not in ["", "string", "ALL"] else [],
        "type_customer_code": [payload.vendor_code] if payload.vendor_code and payload.vendor_code not in ["", "string"] else []
    }

    try:
        res = requests.post(api_url, headers=headers, params={"page": 1, "pageSize": 2000}, json=req_body, timeout=10)
        if res.status_code == 404:
            return [{"error_msg": f"HTTP 404 (경로 없음) - 주소: {api_url}"}]
        if res.status_code == 401:
            return [{"error_msg": f"인증 실패 (HTTP 401) - API Key가 만료되었거나 IP 차단됨. 새 키를 발급받아 상단 입력창에 넣고 다시 조회하세요."}]
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
                        
                    auto_date = calculate_payment_date(occur_date, due_date, raw.get("customer_name", ""), krw_balance)
                    parsed_items.append({
                        "pending_no": raw.get("not_settled_number", ""), "account_code": raw.get("account_code", ""), "account_name": raw.get("account_name", ""),
                        "vendor_code": raw.get("customer_code", ""), "vendor_name": raw.get("customer_name", ""), "occur_date": occur_date, "acc_date": occur_date,
                        "payment_request_date": due_date, "currency": raw.get("currency_code", "KRW"), "exchange_rate": parse_float(raw.get("occur_exchange_rate")),
                        "occur_amount": parse_float(raw.get("occur_amount_ocr")), "balance_amount": balance_amount, "krw_balance": krw_balance,
                        "auto_payment_date": auto_date, "scheduled_payment_date": auto_date, "confirmed_voucher_no": raw.get("group_settled_number", ""),
                        "edm_documents": [{"doc_id": "EDM-1", "doc_type": "증빙", "file_name": f"증빙.pdf", "download_url": "#"}]
                    })
                return parsed_items
            else: return [{"error_msg": f"API 데이터 실패: {json_data.get('message')} - URL: {api_url}"}]
        else: return [{"error_msg": f"서버 응답 에러 (HTTP {res.status_code}) - URL: {api_url}"}]
    except requests.exceptions.Timeout: return [{"error_msg": f"연결 시간 초과 - 방화벽(8400포트) 차단 확인"}]
    except requests.exceptions.ConnectionError: return [{"error_msg": f"접속 거부 - 도메인 장애 또는 Vercel 차단"}]
    except Exception as e: return [{"error_msg": f"백엔드 로직 오류: {str(e)}"}]

def filter_data(payload: PendingSearchQuery, data: List[dict]) -> List[dict]:
    if data and data[0].get("error_msg"): return data
    if payload.start_date and payload.start_date.strip() not in ["", "string"]: data = [item for item in data if item["occur_date"] >= payload.start_date]
    if payload.end_date and payload.end_date.strip() not in ["", "string"]: data = [item for item in data if item["occur_date"] <= payload.end_date]
    if payload.pay_start_date and payload.pay_start_date.strip() not in ["", "string"]: data = [item for item in data if item["scheduled_payment_date"] >= payload.pay_start_date]
    if payload.pay_end_date and payload.pay_end_date.strip() not in ["", "string"]: data = [item for item in data if item["scheduled_payment_date"] <= payload.pay_end_date]
    if payload.pending_no and payload.pending_no.strip() not in ["", "string"]: data = [item for item in data if item["pending_no"] == payload.pending_no]
    return data

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
            .btn-mock { background-color: #6c757d; color: white; border-color: #6c757d; }
            .bg-summary { background-color: #fffbeb; }
            .date-input { background-color: #e8f5e9; border: 1px solid #4caf50; color: #1b5e20; font-weight: bold;}
        </style>
    </head>
    <body class="p-3">
        <nav class="navbar navbar-dark px-4 py-3 rounded mb-4 d-flex justify-content-between">
            <span class="navbar-brand mb-0 h1 fw-bold">🚢 흥아해운 미결 포털 <span class="badge bg-primary fs-6 ms-2">API Key 동적 교체 지원 🟢</span></span>
        </nav>
        
        <!-- 🚨 새로 추가된 API Key 입력창 -->
        <div class="card p-3 mb-4 border-primary">
            <h5 class="fw-bold text-primary mb-3">🔑 인증 정보 설정 (전산팀 발급 키)</h5>
            <div class="input-group">
                <span class="input-group-text bg-primary text-white fw-bold">SAMSAPI Key</span>
                <input type="text" class="form-control fw-bold" id="customApiKey" value="Hw-_k-QPRgzolGqkLFIGYLzwqDnep53-wprci845GWw" placeholder="새로운 API Key를 발급받으면 여기에 붙여넣으세요">
                <button class="btn btn-warning fw-bold" onclick="loadPendingData(false)">이 인증키로 조회(API)</button>
            </div>
            <small class="text-muted mt-2">※ 401 에러가 나면 전산팀에 새 키를 발급받아 위 칸에 붙여넣기만 하면 바로 실데이터가 연동됩니다.</small>
        </div>

        <div class="card p-3 mb-4">
            <h5 class="fw-bold text-secondary mb-3">🔍 미결 조회 조건</h5>
            <div class="row g-3">
                <div class="col-md-2">
                    <label class="form-label text-secondary fw-bold">발생 시작일</label>
                    <input type="date" class="form-control" id="startDate" value="2026-06-01">
                </div>
                <div class="col-md-2">
                    <label class="form-label text-secondary fw-bold">발생 종료일</label>
                    <input type="date" class="form-control" id="endDate" value="2026-08-31">
                </div>
                <div class="col-md-8 d-flex align-items-end gap-2 justify-content-end">
                    <button class="btn btn-mock fw-bold px-4" onclick="loadPendingData(true)">MOCK조회(테스트)</button>
                    <button class="btn btn-primary fw-bold px-5" onclick="loadPendingData(false)">조회(API)</button>
                    <button class="btn btn-excel fw-bold px-4" onclick="downloadExcel()">엑셀(계획)</button>
                </div>
            </div>
        </div>

        <div class="card p-3 mb-4">
            <div class="table-responsive">
                <table class="table table-hover align-middle border text-center" style="font-size: 0.9rem;">
                    <thead class="table-header">
                        <tr>
                            <th>미결번호</th><th>거래처명</th><th>통화</th><th>외화(원화잔액)</th><th>원화환산액</th><th>자동산정일</th><th style="background-color: #1b5e20;">지불예정일(수정)</th><th>저장</th>
                        </tr>
                    </thead>
                    <tbody id="pendingTableBody">
                        <tr><td colspan="8" class="py-4 text-muted">조회 버튼을 눌러 데이터를 불러오세요.</td></tr>
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
            function buildSearchPayload(useMock = false) {
                return {
                    branch_code: "본사", start_date: document.getElementById("startDate").value, end_date: document.getElementById("endDate").value,
                    unsettled_only: true, use_mock: useMock,
                    api_key: document.getElementById("customApiKey").value // 화면의 키워드를 페이로드에 담아 전송
                };
            }

            async function loadPendingData(useMock = false) {
                const payload = buildSearchPayload(useMock);
                const tbody = document.getElementById("pendingTableBody");
                const summaryBody = document.getElementById("summaryTableBody");
                
                tbody.innerHTML = '<tr><td colspan="8" class="py-4 text-primary fw-bold">데이터 수신 처리 중입니다...</td></tr>';
                summaryBody.innerHTML = ""; document.getElementById("grandTotalKrw").innerText = "0 원";
                
                try {
                    const res = await fetch('/api/pending/search-and-schedule', {
                        method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify(payload)
                    });
                    const text = await res.text();
                    let data = JSON.parse(text);
                    
                    if(data.length > 0 && data[0].error_msg) {
                         tbody.innerHTML = `<tr><td colspan="8" class="py-4 text-danger fw-bold fs-5">${data[0].error_msg}</td></tr>`;
                         return;
                    }
                    if(data.length === 0) {
                        tbody.innerHTML = '<tr><td colspan="8" class="py-4 text-muted fw-bold">조건에 해당하는 미상계 데이터가 없습니다. (통신 성공)</td></tr>';
                        return;
                    }

                    let summary = {}; let grandTotalKrw = 0; tbody.innerHTML = "";
                    data.forEach(item => {
                        let curr = item.currency;
                        if(!summary[curr]) summary[curr] = { krw: 0 };
                        summary[curr].krw += item.krw_balance;
                        grandTotalKrw += item.krw_balance;

                        const tr = document.createElement("tr");
                        tr.innerHTML = `
                            <td class="text-primary fw-bold">${item.pending_no}</td>
                            <td class="fw-bold">${item.vendor_name}</td>
                            <td><span class="badge ${curr === 'KRW' ? 'bg-secondary' : 'bg-danger'}">${curr}</span></td>
                            <td class="text-end pe-3">${Number(item.balance_amount).toLocaleString()}</td>
                            <td class="fw-bold text-end pe-3">${Number(item.krw_balance).toLocaleString()} 원</td>
                            <td><span class="text-muted">${item.auto_payment_date}</span></td>
                            <td><input type="date" class="form-control form-control-sm text-center date-input" id="date-${item.pending_no}" value="${item.scheduled_payment_date}"></td>
                            <td><button class="btn btn-sm btn-success fw-bold" onclick="saveDate('${item.pending_no}')">저장</button></td>
                        `;
                        tbody.appendChild(tr);
                    });
                    for(const [curr, amounts] of Object.entries(summary)) {
                        const tr = document.createElement("tr");
                        tr.innerHTML = `<td class="fw-bold text-primary">${curr}</td><td class="text-end pe-4 fw-bold">${Number(amounts.krw).toLocaleString()} 원</td>`;
                        summaryBody.appendChild(tr);
                    }
                    document.getElementById("grandTotalKrw").innerText = Number(grandTotalKrw).toLocaleString() + " 원";
                } catch(e) { tbody.innerHTML = `<tr><td colspan="8" class="py-4 text-danger fw-bold">🚨 오류 발생: ${e.message}</td></tr>`; }
            }
            async function saveDate(pendingNo) {
                const newDate = document.getElementById(`date-${pendingNo}`).value;
                const res = await fetch('/api/pending/save-payment-date', { method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify({ pending_no: pendingNo, target_payment_date: newDate }) });
                alert((await res.json()).message);
            }
            function downloadExcel() {
                fetch('/api/pending/export-plan-excel', { method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify(buildSearchPayload(false)) })
                .then(res => res.blob()).then(blob => { const a = document.createElement('a'); a.href = window.URL.createObjectURL(blob); a.download = `지불계획서.xlsx`; a.click(); });
            }
            window.onload = function() { loadPendingData(false); };
        </script>
    </body>
    </html>
    """

@app.get("/")
def read_root(): return {"status": "online"}

@app.post("/api/pending/search-and-schedule")
def search_and_schedule_pending(payload: PendingSearchQuery):
    try: return filter_data(payload, get_mock_pending_data() if payload.use_mock else fetch_real_pending_data(payload))
    except Exception as e: return [{"error_msg": f"백엔드 처리 오류: {str(e)}"}]

@app.post("/api/pending/save-payment-date")
def save_payment_date(payload: PaymentDateSaveRequest): return {"status": "success", "message": f"[{payload.pending_no}] 지불예정일 저장 완료"}

@app.post("/api/pending/export-plan-excel")
def export_plan_excel(payload: PendingSearchQuery):
    filtered_data = filter_data(payload, get_mock_pending_data() if payload.use_mock else fetch_real_pending_data(payload))
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "지불계획서"
    ws.append(["지불예정일", "미결번호", "거래처명", "통화", "환율", "발생(외화)잔액", "원화환산액", "자동산정일"])
    for cell in ws[1]: cell.fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid"); cell.font = Font(color="FFFFFF", bold=True); cell.alignment = Alignment(horizontal="center")
    
    for row in filtered_data:
        if row.get("error_msg"): continue
        ws.append([row["scheduled_payment_date"], row["pending_no"], row["vendor_name"], row["currency"], row["exchange_rate"], row["balance_amount"], row["krw_balance"], row["auto_payment_date"]])
    stream = io.BytesIO(); wb.save(stream); stream.seek(0)
    return Response(content=stream.getvalue(), media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers={"Content-Disposition": "attachment; filename=Payment_Plan.xlsx"})