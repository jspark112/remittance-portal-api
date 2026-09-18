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
    description="미결 지불예정일(정기물품대/금액별) 자동산정, 단건조회/저장, EDM증빙 및 엑셀 다운로드 API",
    version="1.5.1",
    docs_url="/docs",
    openapi_url="/openapi.json"
)

# 환경 변수 로드
SAMSAPI_BASE_URL = os.getenv("SAMSAPI_BASE_URL", "http://211.104.10.171:7071")
SAMSAPI_KEY = os.getenv("SAMSAPI_KEY", "")

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
    """합산 금액 구간별 지급 일수 산정"""
    if amount <= 2_000_000:
        return 90
    elif amount <= 5_000_000:
        return 60
    elif amount <= 10_000_000:
        return 70
    elif amount <= 20_000_000:
        return 80
    elif amount <= 50_000_000:
        return 90
    elif amount <= 100_000_000:
        return 100
    else:
        return 110

def calculate_payment_date(occur_date_str: str, vendor_name: str, amount: float, currency: str = "KRW") -> str:
    """발생일자 기준 지불예정일(화/금 이월 적용) 자동 계산"""
    occur_dt = datetime.strptime(occur_date_str, "%Y-%m-%d")
    is_regular = any(supplier in vendor_name for supplier in REGULAR_SUPPLIERS) or vendor_name in REGULAR_SUPPLIERS

    if is_regular:
        payment_days = get_payment_days_by_amount(amount)
    else:
        payment_days = 30  # 일반 업체 기본 30일

    base_target_dt = occur_dt + timedelta(days=payment_days)
    weekday = base_target_dt.weekday()  # 0:월, 1:화, 2:수, 3:목, 4:금, 5:토, 6:일
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
    account_codes: Optional[str] = Field(default="2001;2002;2041;", description="계정과목(다중)")
    vendor_code: Optional[str] = Field(None, description="거래처코드")
    pending_no: Optional[str] = Field(None, description="미결번호")
    unsettled_only: bool = Field(default=True, description="미상계건만 조회 여부")

class SinglePendingQuery(BaseModel):
    pending_no: str = Field(..., description="조회할 특정 미결번호", example="APS202606250008-0001")

class PaymentDateSaveRequest(BaseModel):
    pending_no: str = Field(..., description="미결번호", example="APS202606250008-0001")
    target_payment_date: str = Field(..., description="변경 저장할 지불예정일 (YYYY-MM-DD)", example="2026-09-08")
    remarks: Optional[str] = Field(None, description="변경 사유 및 메모")

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
            "vendor_code": "003143",
            "vendor_name": "(주)케이씨",
            "occur_date": "2026-06-03",
            "acc_date": "2026-06-22",
            "currency": "USD",
            "occur_amount": 130.00,
            "balance_amount": 130.00,
            "krw_balance": 196469.0,
            "confirmed_voucher_no": "VC20260622-0045",
            "edm_documents": [
                {
                    "doc_id": "EDM-202606-8801",
                    "doc_type": "외화인보이스(Invoice)",
                    "file_name": "KC_INV_130USD.pdf",
                    "download_url": "https://edm.heunga.com/view/VC20260622-0045/doc1"
                }
            ]
        },
        {
            "pending_no": "APS202607090021-0002",
            "account_code": "2001",
            "account_name": "외상매입금(원화)",
            "vendor_code": "001201",
            "vendor_name": "한라시스템",
            "occur_date": "2026-06-03",
            "acc_date": "2026-06-07",
            "currency": "KRW",
            "occur_amount": 6711000.00,
            "balance_amount": 6711000.00,
            "krw_balance": 6711000.0,
            "confirmed_voucher_no": "VC20260607-0012",
            "edm_documents": [
                {
                    "doc_id": "EDM-202606-9901",
                    "doc_type": "전자세금계산서",
                    "file_name": "Halla_TaxInvoice_6711000.pdf",
                    "download_url": "https://edm.heunga.com/view/VC20260607-0012/doc1"
                }
            ]
        }
    ]
    for item in items:
        auto_date = calculate_payment_date(item["occur_date"], item["vendor_name"], item["krw_balance"], item["currency"])
        item["auto_payment_date"] = auto_date
        item["scheduled_payment_date"] = auto_date
    return items

# ---------------------------------------------------------
# API 엔드포인트 구현
# ---------------------------------------------------------
@app.get("/")
def read_root():
    return {"status": "online", "message": "송금 포털 미결/지불예정/EDM 연동 API 구동 중", "docs": "/docs"}

# 1. 미결 목록 조회 및 지불일 자동 생성 API
@app.post("/api/pending/search-and-schedule", response_model=List[PendingPaymentItem])
def search_and_schedule_pending(payload: PendingSearchQuery):
    data = get_mock_pending_data()
    # Swagger 기본 입력값 "string" 및 공백 필터 예외 처리
    if payload.pending_no and payload.pending_no.strip() not in ["", "string"]:
        data = [item for item in data if item["pending_no"] == payload.pending_no]
    if payload.vendor_code and payload.vendor_code.strip() not in ["", "string"]:
        data = [item for item in data if item["vendor_code"] == payload.vendor_code]
    return data

# 2. 특정 미결번호 단건 상세 조회 API
@app.post("/api/pending/detail", response_model=PendingPaymentItem)
def get_single_pending_detail(payload: SinglePendingQuery):
    data = get_mock_pending_data()
    target_item = next((item for item in data if item["pending_no"] == payload.pending_no), None)
    if not target_item:
        raise HTTPException(status_code=404, detail=f"미결번호 [{payload.pending_no}]를 찾을 수 없습니다.")
    return target_item

# 3. 특정 미결번호 지불일자 변경 및 저장 API
@app.post("/api/pending/save-payment-date")
def save_payment_date(payload: PaymentDateSaveRequest):
    return {
        "status": "success",
        "message": f"미결번호 [{payload.pending_no}]의 지불예정일이 [{payload.target_payment_date}]로 저장되었습니다.",
        "pending_no": payload.pending_no,
        "saved_payment_date": payload.target_payment_date,
        "remarks": payload.remarks,
        "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }

# 4. 미결 지불예정 목록 엑셀 다운로드 API
@app.post("/api/pending/export-excel")
def export_pending_to_excel(payload: PendingSearchQuery):
    data = get_mock_pending_data()
    # Swagger 기본 입력값 "string" 및 공백 필터 예외 처리
    if payload.pending_no and payload.pending_no.strip() not in ["", "string"]:
        data = [item for item in data if item["pending_no"] == payload.pending_no]
    if payload.vendor_code and payload.vendor_code.strip() not in ["", "string"]:
        data = [item for item in data if item["vendor_code"] == payload.vendor_code]

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "지불예정 미결목록"

    headers = [
        "미결번호", "계정코드", "계정과목명", "거래처코드", "거래처명",
        "발생일자", "회계일자", "통화", "발생금액", "발생잔액",
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
            row["acc_date"], row["currency"], row["occur_amount"],
            row["balance_amount"], row["krw_balance"], row["auto_payment_date"],
            row["scheduled_payment_date"], row["confirmed_voucher_no"]
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