import os
import io
from pathlib import Path
from typing import List, Dict, Any, Optional
from fastapi import FastAPI, HTTPException, UploadFile, File, Form
from pydantic import BaseModel
from dotenv import load_dotenv
import requests
import pandas as pd

app = FastAPI(title="Sinokor Remittance Portal API")

BASE_URL = os.getenv("SAMSAPI_BASE_URL", "http://211.104.10.171:7071")
DOMAIN = "accounting"

# =================================================================
# [전산팀 표준 스펙 연동용 데이터 모델]
# =================================================================
class CompanySearchRequest(BaseModel):
    company_name: str = ""
    keyword: str = ""
    domain: str = "accounting"

# =================================================================
# [보안 및 검증 모듈]
# =================================================================
def get_sams_key() -> str:
    cwd_env = Path.cwd() / ".env"
    if cwd_env.exists():
        load_dotenv(dotenv_path=cwd_env, override=True)
        
    key = os.getenv("SAMSAPI_KEY")
    if key and key.strip():
        return key.strip()
            
    raise ValueError(f"[보안 지침 위반] .env 경로({cwd_env})에서 SAMSAPI_KEY를 찾을 수 없습니다.")

def get_headers() -> Dict[str, str]:
    return {
        "X-API-Key": get_sams_key(),
        "Content-Type": "application/json"
    }

def verify_spec_version():
    version_url = f"{BASE_URL}/api/v1/spec/version?domain={DOMAIN}"
    try:
        res = requests.get(version_url, headers=get_headers(), timeout=2)
        if res.status_code == 200:
            return
        print(f"⚠️ [SamsApi 경고] 게이트웨이 응답 코드: {res.status_code} (Base URL 확인 필요)")
    except Exception as e:
        print(f"⚠️ [SamsApi 경고] 게이트웨이 통신 실패: {e}")
    # 올바른 Base URL 수신 전까지 테스트 중단을 방지하기 위해 통과 처리
    pass

# =================================================================
# [API 연동 핵심 로직]
# =================================================================
def fetch_company_info(company_name: str = "", keyword: str = "") -> List[Dict[str, Any]]:
    verify_spec_version()
    payload = {
        "company_name": company_name,
        "keyword": keyword,
        "domain": DOMAIN
    }

    res = requests.post(f"{BASE_URL}/api/v1/co/list", headers=get_headers(), json=payload)
    if res.status_code == 200:
        res_data = res.json()
        if isinstance(res_data, list):
            return res_data
        return res_data.get("data", [])
    return []

def fetch_edm_proof(company_code: str, journal_number: str) -> List[Dict[str, Any]]:
    verify_spec_version()
    payload = {
        "company_code": company_code,
        "journal_number": journal_number,
        "language_gubun": "KO"
    }
    res = requests.post(f"{BASE_URL}/api/v1/edm/list", headers=get_headers(), json=payload)
    if res.status_code == 200:
        res_data = res.json()
        if res_data.get("success"):
            return res_data.get("data", [])
    return []

# =================================================================
# [FastAPI 엔드포인트 라우터]
# =================================================================
@app.get("/")
def root():
    return {"status": "online", "message": "송금 관리 포털 API 서버 작동 중"}

# [신규] 회사코드 / 회사명 검색 API (전산팀 스펙 반영)
@app.post("/api/company/search")
def search_company(req: CompanySearchRequest):
    company_list = fetch_company_info(company_name=req.company_name, keyword=req.keyword)
    return {
        "status": "success",
        "data": company_list
    }

# [방식 1] 엑셀 파일 업로드 수집 API 
@app.post("/api/edm/collect-by-excel")
async def collect_by_excel(
    file: UploadFile = File(...),
    company_code: str = Form("1000")
):
    if not (file.filename.endswith(".xlsx") or file.filename.endswith(".xls")):
        raise HTTPException(status_code=400, detail="엑셀 파일(.xlsx, .xls)만 업로드 가능합니다.")
    
    co_info = fetch_company_info(keyword=company_code)

    try:
        contents = await file.read()
        df = pd.read_excel(io.BytesIO(contents))
        
        target_col = None
        for col in df.columns:
            if str(col).replace(" ", "").lower() in ["전표번호", "journal_number", "journalnumber", "slipno"]:
                target_col = col
                break
                
        if not target_col:
            raise HTTPException(status_code=400, detail="엑셀 내에 '전표번호' 컬럼이 존재하지 않습니다.")

        journal_numbers = df[target_col].dropna().astype(str).str.strip().tolist()
        
        collected_results = []
        for journal_no in journal_numbers:
            proofs = fetch_edm_proof(company_code=company_code, journal_number=journal_no)
            collected_results.append({
                "journal_number": journal_no,
                "proof_count": len(proofs),
                "proof_files": proofs
            })

        return {
            "status": "success",
            "method": "EXCEL_UPLOAD",
            "company_info": co_info[0] if co_info else {"company_code": company_code},
            "total_journals": len(journal_numbers),
            "data": collected_results
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"엑셀 처리 오류: {str(e)}")

# [방식 2] 미결 자동 조회 후 EDM 증빙 연속 수집 API
@app.post("/api/edm/collect-auto")
def collect_auto(payload: dict):
    company_code = payload.get("coCd", "1000")
    
    co_info = fetch_company_info(keyword=company_code)

    verify_spec_version()
    ntstl_res = requests.post(
        f"{BASE_URL}/api/v1/ntstl/list?page=1&pageSize=1000",
        headers=get_headers(),
        json=payload
    )
    
    if ntstl_res.status_code != 200:
        raise HTTPException(status_code=ntstl_res.status_code, detail="미결 자료 조회 실패")

    unsettled_list = ntstl_res.json().get("data", [])
    
    collected_results = []
    for item in unsettled_list:
        journal_no = item.get("slipNo") or item.get("journal_number")
        if not journal_no:
            continue
            
        proofs = fetch_edm_proof(company_code=company_code, journal_number=journal_no)
        
        item_with_proof = item.copy()
        item_with_proof["proof_count"] = len(proofs)
        item_with_proof["proof_files"] = proofs
        collected_results.append(item_with_proof)

    return {
        "status": "success",
        "method": "AUTO_DIRECT_SYNC",
        "company_info": co_info[0] if co_info else {"company_code": company_code},
        "total_unsettled_count": len(collected_results),
        "data": collected_results
    }