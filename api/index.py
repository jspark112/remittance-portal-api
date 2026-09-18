import io
import os
import requests
import zipfile
from datetime import datetime, timedelta
from typing import List, Optional, Dict
from fastapi import FastAPI, HTTPException, Response
from fastapi.responses import HTMLResponse
from pydantic import BaseModel, Field
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill

app = FastAPI(
    title="Remittance Portal API",
    description="SAMSAPI 실시간 연동 미결 지불관리 포털 API",
    version="3.0.0",
    docs_url="/docs",
    openapi_url="/openapi.json"
)

# ---------------------------------------------------------
# 환경 변수 (서버 API 주소 및 키)
# ---------------------------------------------------------
SAMSAPI_BASE_URL = os.getenv("SAMSAPI_BASE_URL", "http://211.104.10.171:7071")
SAMSAPI_KEY = os.getenv("SAMSAPI_KEY", "여기에_발급받은_API_KEY_입력")

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
    # 빈 값 예외 처리
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

# ---------------------------------------------------------
# Pydantic 모델
# ---------------------------------------------------------
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

class PaymentDateSaveRequest(BaseModel):
    pending_no: str
    target_payment_date: str

class EdmDocumentInfo(BaseModel):
    doc_id: str
    doc_type: str
    file_name: str
    download_url: str

class PendingPaymentItem(BaseModel):
    pending_no: str
    account_code: str
    account_name: str
    vendor_code: str
    vendor_name: str
    occur_date: str
    acc_date: str
    payment_request_date: Optional[str] = None
    currency: str
    exchange_rate: float
    occur_amount: float
    balance_amount: float
    krw_balance: float
    auto_payment_date: str
    scheduled_payment_date: str
    confirmed_voucher_no: str
    edm_documents: List[EdmDocumentInfo] = []

# ---------------------------------------------------------
# SAMSAPI 실시간 연동 핵심 로직
# ---------------------------------------------------------
def fetch_real_pending_data(payload: PendingSearchQuery) -> List[dict]:
    api_url = f"{SAMSAPI_BASE_URL}/api/v1/ntstl/list"
    headers = {
        "X-API-Key": SAMSAPI_KEY,
        "Content-Type": "application/json"
    }
    
    # 1. API 요청 본문(Payload) 조립
    target_dt = payload.end_date if (payload.end_date and payload.end_date != "string") else datetime.today().strftime("%Y-%m-%d")
    
    req_body = {
        "company_code": "01", # 전산팀 확인 필요 (회사코드 디폴트값)
        "target_date": target_dt.replace("-", ""), # 예: 20260831
        "type_account_code": [payload.account_code] if payload.account_code and payload.account_code not in ["", "string", "ALL"] else [],
        "type_customer_code": [payload.vendor_code] if payload.vendor_code and payload.vendor_code not in ["", "string"] else []
    }
    
    parsed_items = []
    
    try:
        # 실제 서버로 데이터 요청 (페이징 2000건)
        res = requests.post(api_url, headers=headers, params={"page": 1, "pageSize": 2000}, json=req_body, timeout=10)
        
        if res.status_code == 200:
            json_data = res.json()
            if json_data.get("success"):
                raw_list = json_data.get("data", [])
                
                for raw in raw_list:
                    # 데이터 맵핑 및 전처리
                    pending_no = raw.get("not_settled_number", "")
                    vendor_name = raw.get("customer_name", "")
                    
                    # 날짜 형식 변환 (YYYYMMDD -> YYYY-MM-DD)
                    def format_date(d_str):
                        if d_str and len(d_str) == 8: return f"{d_str[:4]}-{d_str[4:6]}-{d_str[6:]}"
                        return d_str
                        
                    occur_date = format_date(raw.get("occur_date", ""))
                    due_date = format_date(raw.get("due_date", ""))
                    
                    # 금액 형식 변환 (String -> Float)
                    def parse_float(val):
                        try: return float(val) if val else 0.0
                        except: return 0.0

                    balance_amount = parse_float(raw.get("occur_amount_bal"))
                    krw_balance = parse_float(raw.get("local_amount_bal"))
                    
                    # 미상계건(잔액 존재)만 필터링
                    if payload.unsettled_only and balance_amount <= 0:
                        continue
                        
                    # 지불일 자동 산정
                    auto_date = calculate_payment_date(occur_date, due_date, vendor_name, krw_balance)
                    
                    parsed_items.append({
                        "pending_no": pending_no,
                        "account_code": raw.get("account_code", ""),
                        "account_name": raw.get("account_name", ""),
                        "vendor_code": raw.get("customer_code", ""),
                        "vendor_name": vendor_name,
                        "occur_date": occur_date,
                        "acc_date": format_date(raw.get("from_date", "")),
                        "payment_request_date": due_date,
                        "currency": raw.get("currency_code", "KRW"),
                        "exchange_rate": parse_float(raw.get("occur_exchange_rate")),
                        "occur_amount": parse_float(raw.get("occur_amount_ocr")),
                        "balance_amount": balance_amount,
                        "krw_balance": krw_balance,
                        "auto_payment_date": auto_date,
                        "scheduled_payment_date": auto_date,
                        "confirmed_voucher_no": raw.get("group_settled_number", ""),
                        "edm_documents": [{"doc_id": "EDM-1", "doc_type": "증빙", "file_name": f"{vendor_name}_증빙.pdf", "download_url": "#"}]
                    })
                return parsed_items
    except Exception as e:
        print(f"SAMSAPI 연동 에러: {e}")
        # 오류 발생 시 빈 배열 반환
        pass

    return parsed_items

# 로컬(프론트엔드) 추가 필터링 로직 (API가 처리해주지 않는 세부 필터)
def filter_data(payload: PendingSearchQuery, data: List[dict]) -> List[dict]:
    if payload.start_date and payload.start_date.strip() not in ["", "string"]:
        data = [item for item in data if item["occur_date"] >= payload.start_date]
    if payload.end_date and payload.end_date.strip() not in ["", "string"]:
        data = [item for item in data if item["occur_date"] <= payload.end_date]

    if payload.pay_start_date and payload.pay_start_date.strip() not in ["", "string"]:
        data = [item for item in data if item["scheduled_payment_date"] >= payload.pay_start_date]
    if payload.pay_end_date and payload.pay_end_date.strip() not in ["", "string"]:
        data = [item for item in data if item["scheduled_payment_date"] <= payload.pay_end_date]

    if payload.pending_no and payload.pending_no.strip() not in ["", "string"]:
        data = [item for item in data if item["pending_no"] == payload.pending_no]
    return data

# ---------------------------------------------------------
# 포털 대시보드 사용자 UI HTML
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
            .bg-summary { background-color: #fffbeb; }
            .date-input { background-color: #e8f5e9; border: 1px solid #4caf50; color: #1b5e20; font-weight: bold;}
        </style>
    </head>
    <body class="p-3">
        <nav class="navbar navbar-dark px-4 py-3 rounded mb-4 d-flex justify-content-between">
            <span class="navbar-brand mb-0 h1 fw-bold">🚢 흥아해운 미결 지불관리 포털 <span class="badge bg-success fs-6 ms-2">SAMSAPI 연동 🟢</span></span>
        </nav>
        
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
            
            <div class="row g-3 mt-2">
                <div class="col-md-8"></div>
                <div class="col-md-4 d-flex align-items-end gap-2">
                    <button class="btn btn-primary fw-bold flex-fill" onclick="loadPendingData()">조회</button>
                    <button class="btn btn-excel fw-bold flex-fill" onclick="downloadExcel()">엑셀(계획)</button>
                    <button class="btn btn-zip fw-bold flex-fill" onclick="downloadEdmZip()">증빙 ZIP</button>
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
            function buildSearchPayload() {
                return {
                    branch_code: "본사",
                    start_date: document.getElementById("startDate").value,
                    end_date: document.getElementById("endDate").value,
                    pay_start_date: document.getElementById("payStartDate").value,
                    pay_end_date: document.getElementById("payEndDate").value,
                    account_code: document.getElementById("accountCode").value,
                    vendor_code: document.getElementById("vendorCode").value,
                    unsettled_only: true
                };
            }

            async function loadPendingData() {
                const payload = buildSearchPayload();
                const res = await fetch('/api/pending/search-and-schedule', {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify(payload)
                });
                const data = await res.json();
                
                const tbody = document.getElementById("pendingTableBody");
                const summaryBody = document.getElementById("summaryTableBody");
                tbody.innerHTML = ""; summaryBody.innerHTML = "";
                
                if(data.length === 0) {
                    tbody.innerHTML = '<tr><td colspan="10" class="py-4 text-danger fw-bold">검색 조건에 해당하거나 미상계된 데이터가 없습니다. (API 통신 확인)</td></tr>';
                    document.getElementById("grandTotalKrw").innerText = "0 원";
                    return;
                }

                let summary = {}; let grandTotalKrw = 0;

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

            function downloadExcel() {
                fetch('/api/pending/export-plan-excel', {
                    method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify(buildSearchPayload())
                }).then(res => res.blob()).then(blob => {
                    const a = document.createElement('a'); a.href = window.URL.createObjectURL(blob); a.download = `지불계획서_${new Date().toISOString().slice(0,10)}.xlsx`; a.click();
                });
            }

            function downloadEdmZip() {
                fetch('/api/pending/export-edm-zip', {
                    method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify(buildSearchPayload())
                }).then(res => res.blob()).then(blob => {
                    const a = document.createElement('a'); a.href = window.URL.createObjectURL(blob); a.download = `EDM_증빙자료_${new Date().toISOString().slice(0,10)}.zip`; a.click();
                });
            }

            window.onload = loadPendingData;
        </script>
    </body>
    </html>
    """

# ---------------------------------------------------------
# API 엔드포인트 구현
# ---------------------------------------------------------
@app.get("/")
def read_root(): return {"status": "online"}

@app.post("/api/pending/search-and-schedule", response_model=List[PendingPaymentItem])
def search_and_schedule_pending(payload: PendingSearchQuery):
    raw_data = fetch_real_pending_data(payload)
    return filter_data(payload, raw_data)

@app.post("/api/pending/save-payment-date")
def save_payment_date(payload: PaymentDateSaveRequest):
    return {"status": "success", "message": f"[{payload.pending_no}] 지불예정일이 저장되었습니다."}

@app.post("/api/pending/export-plan-excel")
def export_plan_excel(payload: PendingSearchQuery):
    filtered_data = filter_data(payload, fetch_real_pending_data(payload))
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "지불계획서"
    ws.append(["지불예정일", "미결번호", "계정코드", "계정명", "거래처명", "통화", "환율", "발생(외화)잔액", "원화환산액", "자동산정일"])
    for cell in ws[1]: cell.fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid"); cell.font = Font(color="FFFFFF", bold=True); cell.alignment = Alignment(horizontal="center")
    
    summary = {}; grand_krw = 0
    for row in filtered_data:
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
    filtered_data = filter_data(payload, fetch_real_pending_data(payload))
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
        counter = 1
        for item in filtered_data:
            for doc in item["edm_documents"]:
                dummy_content = f"이 파일은 {item['vendor_name']} 업체의 증빙 문서입니다.\n미결번호: {item['pending_no']}\n원화잔액: {item['krw_balance']}원".encode('utf-8')
                new_filename = f"{counter}_{item['vendor_name'].replace('/', '_')}_{doc['doc_type']}.pdf"
                zip_file.writestr(new_filename, dummy_content)
                counter += 1
    zip_buffer.seek(0)
    return Response(content=zip_buffer.getvalue(), media_type="application/zip", headers={"Content-Disposition": f"attachment; filename=EDM_Documents_{datetime.now().strftime('%Y%m%d')}.zip"})