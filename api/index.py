import io
import os
import requests
from datetime import datetime, timedelta
from typing import List, Optional
from fastapi import FastAPI, HTTPException, Response
from pydantic import BaseModel, Field
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill

app = FastAPI(
    title="Remittance Portal API",
    description="전체 업체 미결 지불예정일(일반:요청일기준 / 정기물품대:금액별) 자동산정, 수정 저장 및 엑셀 다운로드 API",
    version="1.6.0",
    docs_url="/docs",
    openapi_url="/openapi.json"
)

# ---------------------------------------------------------
# 정기 물품대 업체 목록 (29개사) 및 금액별 지불 정책 Engine
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
    """정기 물품대 업체의 합산 금액 구간별 지급 일수 산정"""
    if amount <= 2_000_000: return 90
    elif amount <= 5_000_000: return 60
    elif amount <= 10_000_000: return 70
    elif amount <= 20_000_000: return 80
    elif amount <= 50_000_000: return 90
    elif amount <= 100_000_000: return 100
    else: return 110

def calculate_payment_date(occur_date_str: str, request_date_str: str, vendor_name: str, amount: float) -> str:
    """
    [지불일 산정 핵심 로직]
    - 정기 물품대 업체: 발생일자(occur_date) + 금액별 유예일수(60~110일) 적용 후 화/금 이월
    - 일반 업체: 지불 요청일자(request_date) 기준 화/금 이월
    """
    is_regular = any(supplier in vendor_name for supplier in REGULAR_SUPPLIERS) or vendor_name in REGULAR_SUPPLIERS

    if is_regular:
        # 정기 물품대 업체: 발생일 기준 금액 조건 적용
        occur_dt = datetime.strptime(occur_date_str, "%Y-%m-%d")
        payment_days = get_payment_days_by_amount(amount)
        base_target_dt = occur_dt + timedelta(days=payment_days)
    else:
        # 일반 업체: 지불 요청일 기준 (요청일이 없으면 발생일+30일 대체)
        if request_date_str:
            base_target_dt = datetime.strptime(request_date_str, "%Y-%m-%d")
        else:
            base_target_dt = datetime.strptime(occur_date_str, "%Y-%m-%d") + timedelta(days=30)

    # 화/금 집행일 자동 보정 (0:월, 1:화, 2:수, 3:목, 4:금, 5:토, 6:일)
    weekday = base_target_dt.weekday()
    days_to_add_map = {0: 1, 1: 0, 2: 2, 3: 1, 4: 0, 5: 3, 6: 2}
    
    final_payment_dt = base_target_dt + timedelta(days=days_to_add_map[weekday])
    return final_payment_dt.strftime("%Y-%m-%d")

# ---------------------------------------------------------
# Pydantic 데이터 모델
# ---------------------------------------------------------
class PendingSearchQuery(BaseModel):
    branch_code: str = Field(default="본사", description="지사코드")
    start_date: str = Field(..., description="발생 시작일자 (YYYY-MM-DD)", example="2026-06-01")
    end_date: str = Field(..., description="발생 종료일자 (YYYY-MM-DD)", example="2026-06-30")
    vendor_code: Optional[str] = Field(None, description="거래처코드")
    pending_no: Optional[str] = Field(None, description="미결번호")

class SinglePendingQuery(BaseModel):
    pending_no: str = Field(..., description="조회할 미결번호", example="APS202606250008-0001")

class PaymentDateSaveRequest(BaseModel):
    pending_no: str = Field(..., description="미결번호", example="APS202606250008-0001")
    target_payment_date: str = Field(..., description="수정/저장할 지불예정일 (YYYY-MM-DD)", example="2026-09-08")
    remarks: Optional[str] = Field(None, description="수정 사유")

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
    payment_request_date: str  # 신규 추가: 지불 요청일자
    currency: str
    occur_amount: float
    balance_amount: float
    krw_balance: float
    auto_payment_date: str
    scheduled_payment_date: str
    confirmed_voucher_no: str
    edm_documents: List[EdmDocumentInfo] = []

def get_mock_pending_data() -> List[dict]:
    items = [
        {
            "pending_no": "APS202606250008-0001",
            "account_code": "2002",
            "account_name": "외상매입금(외화)",
            "vendor_code": "007003",
            "vendor_name": "TIME MARINE CO., LTD", # 일반 업체 (정기물품대 아님)
            "occur_date": "2026-06-03",
            "acc_date": "2026-06-22",
            "payment_request_date": "2026-07-05", # 지불 요청일자
            "currency": "USD",
            "occur_amount": 130.00,
            "balance_amount": 130.00,
            "krw_balance": 196469.0,
            "confirmed_voucher_no": "VC20260622-0045",
            "edm_documents": [
                {"doc_id": "EDM-001", "doc_type": "Invoice", "file_name": "INV_130.pdf", "download_url": "https://edm/1"}
            ]
        },
        {
            "pending_no": "APS202607090021-0002",
            "account_code": "2001",
            "account_name": "외상매입금(원화)",
            "vendor_code": "001201",
            "vendor_name": "한라시스템", # 정기 물품대 업체 (금액 정책 70일 대상)
            "occur_date": "2026-06-03",
            "acc_date": "2026-06-07",
            "payment_request_date": "2026-06-15", # 정기물품대이므로 요청일 대신 발생일 기반 계산됨
            "currency": "KRW",
            "occur_amount": 6711000.00,
            "balance_amount": 6711000.00,
            "krw_balance": 6711000.0,
            "confirmed_voucher_no": "VC20260607-0012",
            "edm_documents": [
                {"doc_id": "EDM-002", "doc_type": "TaxInvoice", "file_name": "Tax_671.pdf", "download_url": "https://edm/2"}
            ]
        }
    ]
    # 지불 정책 계산 로직 태우기
    for item in items:
        auto_date = calculate_payment_date(
            item["occur_date"], 
            item["payment_request_date"], 
            item["vendor_name"], 
            item["krw_balance"]
        )
        item["auto_payment_date"] = auto_date
        item["scheduled_payment_date"] = auto_date # 초기 세팅은 자동산정일과 동일
    return items

# ---------------------------------------------------------
# API 엔드포인트 구현
# ---------------------------------------------------------
@app.get("/")
def read_root(): return {"status": "online"}

# 1. 모든 업체 미결 목록 조회 (자동 지불일 세팅 포함)
@app.post("/api/pending/search-and-schedule", response_model=List[PendingPaymentItem])
def search_and_schedule_pending(payload: PendingSearchQuery):
    data = get_mock_pending_data()
    if payload.pending_no and payload.pending_no.strip() not in ["", "string"]:
        data = [item for item in data if item["pending_no"] == payload.pending_no]
    if payload.vendor_code and payload.vendor_code.strip() not in ["", "string"]:
        data = [item for item in data if item["vendor_code"] == payload.vendor_code]
    return data

# 2. 특정 미결건 수정/저장 기능 (수동으로 지불일자 변경 시)
@app.post("/api/pending/save-payment-date")
def save_payment_date(payload: PaymentDateSaveRequest):
    """
    전달받은 특정 미결번호의 최종 지불예정일을 DB에 덮어쓰기 저장하는 역할
    """
    return {
        "status": "success",
        "message": f"미결번호 [{payload.pending_no}]의 지불예정일이 [{payload.target_payment_date}]로 정상 변경/저장 되었습니다.",
        "pending_no": payload.pending_no,
        "saved_payment_date": payload.target_payment_date,
        "remarks": payload.remarks
    }

# 3. 엑셀 다운로드 API (지불요청일자 컬럼 추가)
@app.post("/api/pending/export-excel")
def export_pending_to_excel(payload: PendingSearchQuery):
    data = get_mock_pending_data()
    if payload.pending_no and payload.pending_no.strip() not in ["", "string"]:
        data = [item for item in data if item["pending_no"] == payload.pending_no]

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "지불예정 미결목록"

    headers = [
        "미결번호", "계정코드", "계정과목명", "거래처코드", "거래처명",
        "발생일자", "회계일자", "지불요청일자", "통화", "발생금액", "발생잔액",
        "원화잔액", "자동산정 지불일", "최종 지불예정일", "확정전표번호"
    ]
    ws.append(headers)

    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for row in data:
        ws.append([
            row["pending_no"], row["account_code"], row["account_name"],
            row["vendor_code"], row["vendor_name"], row["occur_date"],
            row["acc_date"], row["payment_request_date"], row["currency"], 
            row["occur_amount"], row["balance_amount"], row["krw_balance"], 
            row["auto_payment_date"], row["scheduled_payment_date"], row["confirmed_voucher_no"]
        ])

    stream = io.BytesIO()
    wb.save(stream)
    stream.seek(0)
    filename = f"pending_payments_{datetime.now().strftime('%Y%m%d')}.xlsx"

    return Response(
        content=stream.getvalue(),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )