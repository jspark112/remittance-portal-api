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
    description="지불계획 요약 및 EDM 증빙자료 일괄 압축(ZIP) 다운로드 기능이 포함된 포털 API",
    version="2.1.0",
    docs_url="/docs",
    openapi_url="/openapi.json"
)

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
    is_regular = any(supplier in vendor_name for supplier in REGULAR_SUPPLIERS) or any(vendor_name in supplier for supplier in REGULAR_SUPPLIERS)
    if is_regular:
        occur_dt = datetime.strptime(occur_date_str, "%Y-%m-%d")
        base_target_dt = occur_dt + timedelta(days=get_payment_days_by_amount(amount))
    else:
        if request_date_str:
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
    start_date: str = Field(..., example="2026-06-01")
    end_date: str = Field(..., example="2026-06-30")
    account_code: Optional[str] = None
    vendor_code: Optional[str] = None
    pending_no: Optional[str] = None

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

def get_mock_pending_data() -> List[dict]:
    items = [
        {"pending_no": "APS202606250008-0001", "account_code": "2002", "account_name": "외상매입금(외화)", "vendor_code": "007003", "vendor_name": "TIME MARINE CO., LTD", "occur_date": "2026-06-03", "acc_date": "2026-06-22", "payment_request_date": "2026-07-03", "currency": "USD", "exchange_rate": 1511.30, "occur_amount": 130.00, "balance_amount": 130.00, "krw_balance": 196469.0, "confirmed_voucher_no": "VC20260622-0045", "edm_documents": [{"doc_id": "EDM-1", "doc_type": "Invoice", "file_name": "TIME_MARINE_INV.pdf", "download_url": "#"}]},
        {"pending_no": "APS202607090021-0002", "account_code": "2001", "account_name": "외상매입금(원화)", "vendor_code": "003143", "vendor_name": "(주)케이씨", "occur_date": "2026-06-03", "acc_date": "2026-06-07", "payment_request_date": "", "currency": "KRW", "exchange_rate": 1.0, "occur_amount": 6711000.00, "balance_amount": 6711000.00, "krw_balance": 6711000.0, "confirmed_voucher_no": "VC20260607-0012", "edm_documents": [{"doc_id": "EDM-2", "doc_type": "TaxInvoice", "file_name": "KC_Tax.pdf", "download_url": "#"}]},
        {"pending_no": "APS202607010005-0006", "account_code": "2001", "account_name": "외상매입금(원화)", "vendor_code": "003081", "vendor_name": "(주)매일마린", "occur_date": "2026-06-05", "acc_date": "2026-06-16", "payment_request_date": "", "currency": "KRW", "exchange_rate": 1.0, "occur_amount": 22706640.00, "balance_amount": 22706640.00, "krw_balance": 22706640.0, "confirmed_voucher_no": "VC20260616-0089", "edm_documents": [{"doc_id": "EDM-3", "doc_type": "TaxInvoice", "file_name": "MM_Tax.pdf", "download_url": "#"}]},
        {"pending_no": "APS202606260013-0003", "account_code": "2001", "account_name": "외상매입금(원화)", "vendor_code": "003130", "vendor_name": "주식회사 선양", "occur_date": "2026-06-05", "acc_date": "2026-06-15", "payment_request_date": "", "currency": "KRW", "exchange_rate": 1.0, "occur_amount": 819000.00, "balance_amount": 819000.00, "krw_balance": 819000.0, "confirmed_voucher_no": "VC20260615-0034", "edm_documents": [{"doc_id": "EDM-4", "doc_type": "TaxInvoice", "file_name": "Sunyang_Tax.pdf", "download_url": "#"}]},
        {"pending_no": "APS202607060017-0002", "account_code": "2001", "account_name": "외상매입금(원화)", "vendor_code": "002947", "vendor_name": "주식회사 금강항해통신", "occur_date": "2026-06-07", "acc_date": "2026-06-22", "payment_request_date": "", "currency": "KRW", "exchange_rate": 1.0, "occur_amount": 7143930.00, "balance_amount": 7143930.00, "krw_balance": 7143930.0, "confirmed_voucher_no": "VC20260622-0091", "edm_documents": [{"doc_id": "EDM-5", "doc_type": "TaxInvoice", "file_name": "Geumgang_Tax.pdf", "download_url": "#"}]}
    ]
    for item in items:
        auto_date = calculate_payment_date(item["occur_date"], item["payment_request_date"], item["vendor_name"], item["krw_balance"])
        item["auto_payment_date"] = auto_date
        item["scheduled_payment_date"] = auto_date
    return items

def filter_data(payload: PendingSearchQuery, data: List[dict]) -> List[dict]:
    if payload.start_date and payload.start_date.strip() != "string":
        data = [item for item in data if item["occur_date"] >= payload.start_date]
    if payload.end_date and payload.end_date.strip() != "string":
        data = [item for item in data if item["occur_date"] <= payload.end_date]
    if payload.account_code and payload.account_code.strip() not in ["", "string", "ALL"]:
        data = [item for item in data if item["account_code"] == payload.account_code]
    if payload.pending_no and payload.pending_no.strip() not in ["", "string"]:
        data = [item for item in data if item["pending_no"] == payload.pending_no]
    if payload.vendor_code and payload.vendor_code.strip() not in ["", "string"]:
        data = [item for item in data if item["vendor_code"] == payload.vendor_code]
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
            .btn-zip:hover { background-color: #59359a; color: white; }
            .bg-summary { background-color: #fffbeb; }
            .date-input { background-color: #e8f5e9; border: 1px solid #4caf50; color: #1b5e20; font-weight: bold;}
        </style>
    </head>
    <body class="p-3">
        <nav class="navbar navbar-dark px-4 py-3 rounded mb-4">
            <span class="navbar-brand mb-0 h1 fw-bold">🚢 흥아해운 미결 지불관리 포털</span>
        </nav>
        
        <div class="card p-3 mb-4">
            <h5 class="fw-bold text-secondary mb-3">🔍 미결 조회 조건</h5>
            <div class="row g-3">
                <div class="col-md-2">
                    <label class="form-label text-secondary fw-bold">계정과목</label>
                    <select class="form-select border-primary" id="accountCode">
                        <option value="">전체보기</option>
                        <option value="2001">2001 (원화)</option>
                        <option value="2002">2002 (외화)</option>
                    </select>
                </div>
                <div class="col-md-2">
                    <label class="form-label text-secondary fw-bold">발생 시작일</label>
                    <input type="date" class="form-control" id="startDate" value="2026-06-01">
                </div>
                <div class="col-md-2">
                    <label class="form-label text-secondary fw-bold">발생 종료일</label>
                    <input type="date" class="form-control" id="endDate" value="2026-06-30">
                </div>
                <div class="col-md-2">
                    <label class="form-label text-secondary fw-bold">거래처코드</label>
                    <input type="text" class="form-control" id="vendorCode" placeholder="코드 입력">
                </div>
                <!-- 버튼 3종 세트 -->
                <div class="col-md-4 d-flex align-items-end gap-2">
                    <button class="btn btn-primary fw-bold flex-fill" onclick="loadPendingData()">조회</button>
                    <button class="btn btn-excel fw-bold flex-fill" onclick="downloadExcel()">엑셀(계획)</button>
                    <button class="btn btn-zip fw-bold flex-fill" onclick="downloadEdmZip()">증빙 ZIP</button>
                </div>
            </div>
        </div>

        <div class="card p-3 mb-4">
            <h5 class="fw-bold text-secondary mb-3">📋 미결 지불 대상 목록 (자동산정 및 수동저장)</h5>
            <div class="table-responsive">
                <table class="table table-hover align-middle border text-center" style="font-size: 0.9rem;">
                    <thead class="table-header">
                        <tr>
                            <th>미결번호</th>
                            <th>거래처명</th>
                            <th>통화</th>
                            <th>환율</th>
                            <th>외화잔액(원화잔액)</th>
                            <th>원화환산액</th>
                            <th>자동산정일(참고용)</th>
                            <th style="background-color: #1b5e20;">지불예정일 지정(수정가능)</th>
                            <th>저장</th>
                            <th>증빙</th>
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
                    account_code: document.getElementById("accountCode").value,
                    vendor_code: document.getElementById("vendorCode").value
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
                tbody.innerHTML = "";
                summaryBody.innerHTML = "";
                
                if(data.length === 0) {
                    tbody.innerHTML = '<tr><td colspan="10" class="py-4 text-danger fw-bold">검색된 데이터가 없습니다.</td></tr>';
                    document.getElementById("grandTotalKrw").innerText = "0 원";
                    return;
                }

                let summary = {};
                let grandTotalKrw = 0;

                data.forEach(item => {
                    let curr = item.currency;
                    if(!summary[curr]) summary[curr] = { original: 0, krw: 0 };
                    summary[curr].original += item.balance_amount;
                    summary[curr].krw += item.krw_balance;
                    grandTotalKrw += item.krw_balance;

                    const tr = document.createElement("tr");
                    const origAmount = Number(item.balance_amount).toLocaleString(undefined, {minimumFractionDigits: curr==='KRW'?0:2});
                    const krwAmount = Number(item.krw_balance).toLocaleString();
                    const fxRate = Number(item.exchange_rate).toLocaleString();
                    
                    tr.innerHTML = `
                        <td class="text-primary">${item.pending_no}</td>
                        <td class="fw-bold">${item.vendor_name}</td>
                        <td><span class="badge ${curr === 'KRW' ? 'bg-secondary' : 'bg-danger'}">${curr}</span></td>
                        <td>${fxRate}</td>
                        <td class="text-end pe-3">${origAmount}</td>
                        <td class="fw-bold text-end pe-3">${krwAmount} 원</td>
                        <td><span class="text-muted">${item.auto_payment_date}</span></td>
                        <td>
                            <input type="date" class="form-control form-control-sm text-center date-input" id="date-${item.pending_no}" value="${item.scheduled_payment_date}">
                        </td>
                        <td><button class="btn btn-sm btn-success fw-bold" onclick="saveDate('${item.pending_no}')">저장</button></td>
                        <td><a href="#" class="btn btn-sm btn-outline-secondary">EDM</a></td>
                    `;
                    tbody.appendChild(tr);
                });

                for(const [curr, amounts] of Object.entries(summary)) {
                    const tr = document.createElement("tr");
                    tr.innerHTML = `
                        <td class="fw-bold text-primary">${curr}</td>
                        <td class="text-end pe-4">${Number(amounts.original).toLocaleString(undefined, {minimumFractionDigits: curr==='KRW'?0:2})}</td>
                        <td class="text-end pe-4 fw-bold">${Number(amounts.krw).toLocaleString()} 원</td>
                    `;
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
                const result = await res.json();
                alert(result.message);
            }

            function downloadExcel() {
                fetch('/api/pending/export-plan-excel', {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify(buildSearchPayload())
                })
                .then(res => res.blob())
                .then(blob => {
                    const url = window.URL.createObjectURL(blob);
                    const a = document.createElement('a');
                    a.href = url;
                    a.download = `지불계획서_${new Date().toISOString().slice(0,10)}.xlsx`;
                    a.click();
                });
            }

            function downloadEdmZip() {
                alert("현재 화면에 조회된 건들의 증빙서류를 압축(ZIP)하여 다운로드합니다.");
                fetch('/api/pending/export-edm-zip', {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify(buildSearchPayload())
                })
                .then(res => res.blob())
                .then(blob => {
                    const url = window.URL.createObjectURL(blob);
                    const a = document.createElement('a');
                    a.href = url;
                    a.download = `EDM_증빙자료_${new Date().toISOString().slice(0,10)}.zip`;
                    a.click();
                });
            }

            window.onload = loadPendingData;
        </script>
    </body>
    </html>
    """

# ---------------------------------------------------------
# API 엔드포인트
# ---------------------------------------------------------
@app.get("/", tags=["1. 상태"])
def read_root(): return {"status": "online"}

@app.post("/api/pending/search-and-schedule", response_model=List[PendingPaymentItem], tags=["2. 미결조회"])
def search_and_schedule_pending(payload: PendingSearchQuery):
    return filter_data(payload, get_mock_pending_data())

@app.post("/api/pending/save-payment-date", tags=["3. 지불일 수정"])
def save_payment_date(payload: PaymentDateSaveRequest):
    return {"status": "success", "message": f"미결번호 [{payload.pending_no}]의 지불예정일이 [{payload.target_payment_date}]로 저장되었습니다."}

@app.post("/api/pending/export-plan-excel", tags=["4. 지불계획 엑셀 다운로드"])
def export_plan_excel(payload: PendingSearchQuery):
    filtered_data = filter_data(payload, get_mock_pending_data())
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "지불계획서"

    headers = ["지불예정일", "미결번호", "거래처명", "통화", "환율", "발생(외화)잔액", "원화환산액", "자동산정일"]
    ws.append(headers)
    
    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    for cell in ws[1]:
        cell.fill = header_fill; cell.font = header_font; cell.alignment = Alignment(horizontal="center", vertical="center")

    summary = {}
    grand_krw = 0
    for row in filtered_data:
        ws.append([row["scheduled_payment_date"], row["pending_no"], row["vendor_name"], row["currency"], row["exchange_rate"], row["balance_amount"], row["krw_balance"], row["auto_payment_date"]])
        curr = row["currency"]
        if curr not in summary: summary[curr] = {"orig": 0, "krw": 0}
        summary[curr]["orig"] += row["balance_amount"]
        summary[curr]["krw"] += row["krw_balance"]
        grand_krw += row["krw_balance"]

    ws.append([]); ws.append([])
    ws.append(["[ 통화별 지불 계획 요약 ]"])
    ws.append(["통화", "통화별 합계 금액", "통화별 원화 환산액"])
    
    summary_header_fill = PatternFill(start_color="F4B084", end_color="F4B084", fill_type="solid")
    for cell in ws[ws.max_row]:
        cell.fill = summary_header_fill; cell.font = Font(bold=True); cell.alignment = Alignment(horizontal="center")

    for curr, val in summary.items(): ws.append([curr, val["orig"], val["krw"]])
    ws.append(["총 원화 지불 합계", "", grand_krw])
    
    for cell in ws[ws.max_row]:
        cell.font = Font(bold=True, color="C00000"); cell.fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

    stream = io.BytesIO()
    wb.save(stream)
    stream.seek(0)
    return Response(content=stream.getvalue(), media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers={"Content-Disposition": f"attachment; filename=Payment_Plan.xlsx"})

# [신규] 5. EDM 증빙자료 일괄 ZIP 다운로드
@app.post("/api/pending/export-edm-zip", tags=["5. 증빙자료(ZIP) 일괄 다운로드"])
def export_edm_zip(payload: PendingSearchQuery):
    filtered_data = filter_data(payload, get_mock_pending_data())
    
    zip_buffer = io.BytesIO()
    
    # zipfile 객체를 생성하여 파일들을 메모리에 압축
    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
        counter = 1
        for item in filtered_data:
            for doc in item["edm_documents"]:
                # 실무 환경: content = requests.get(doc["download_url"]).content
                # 현재는 테스트를 위해 파일 내용(텍스트)을 동적으로 생성
                dummy_content = f"이 파일은 {item['vendor_name']} 업체의 {doc['doc_type']} 증빙 문서입니다.\n미결번호: {item['pending_no']}\n발생금액: {item['krw_balance']}원".encode('utf-8')
                
                # 파일명 생성 규칙: [순번]_[거래처명]_[증빙종류].pdf
                # 파일명에 들어갈 수 없는 특수기호 변환
                safe_vendor_name = item['vendor_name'].replace("/", "_").replace("\\", "_")
                new_filename = f"{counter}_{safe_vendor_name}_{doc['doc_type']}.pdf"
                
                zip_file.writestr(new_filename, dummy_content)
                counter += 1

    zip_buffer.seek(0)
    
    return Response(
        content=zip_buffer.getvalue(),
        media_type="application/zip",
        headers={"Content-Disposition": f"attachment; filename=EDM_Documents_{datetime.now().strftime('%Y%m%d')}.zip"}
    )