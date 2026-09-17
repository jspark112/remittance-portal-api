import os
import requests
from fastapi import FastAPI, HTTPException
from pydantic import BaseModel, Field

# FastAPI 애플리케이션 초기화 (Vercel 배포 명세 반영)
app = FastAPI(
    title="Remittance Portal API",
    description="송금 관리 포털 - SAMSAPI 연동 백엔드 API",
    version="1.0.0",
    docs_url="/docs",
    openapi_url="/openapi.json"
)

# 환경 변수 로드 (.env 및 Vercel Environment Variables 지원)
SAMSAPI_BASE_URL = os.getenv("SAMSAPI_BASE_URL", "http://211.104.10.171:7071")
SAMSAPI_KEY = os.getenv("SAMSAPI_KEY", "")


# 1. 입력 데이터 규격 정의 (Pydantic Model)
class CompanySearchRequest(BaseModel):
    company_name: str = Field(..., description="회사명", example="흥아해운")
    keyword: str = Field(..., description="검색 키워드", example="송금")
    domain: str = Field(default="accounting", description="도메인 영역 (기본값: accounting)")


# 2. 루트 엔드포인트 (상태 점검용)
@app.get("/")
def read_root():
    return {
        "status": "online",
        "message": "송금 포털 API 서버가 정상 구동 중입니다.",
        "docs": "/docs"
    }


# 3. SAMSAPI 연동 및 검색 엔드포인트
@app.post("/api/company/search")
def search_company(payload: CompanySearchRequest):
    """
    SAMSAPI 회계 도메인 연동 조회 API
    - 테스트 단계 404 및 통신 장애 시 우회(Bypass) 데이터 반환
    """
    target_url = f"{SAMSAPI_BASE_URL.rstrip('/')}/api/{payload.domain}/search"
    headers = {
        "Authorization": f"Bearer {SAMSAPI_KEY}",
        "Content-Type": "application/json"
    }

    try:
        response = requests.post(
            target_url,
            json=payload.model_dump(),
            headers=headers,
            timeout=10
        )

        # 404 발생 시 테스트용 우회(Bypass) 응답 처리
        if response.status_code == 404:
            return {
                "status": "warning",
                "bypass": True,
                "message": "SAMSAPI에서 해당 데이터를 찾지 못해 테스트용 우회 데이터를 반환합니다.",
                "data": {
                    "company_name": payload.company_name,
                    "keyword": payload.keyword,
                    "domain": payload.domain,
                    "result_code": "MOCK_404_BYPASS"
                }
            }

        response.raise_for_status()
        return response.json()

    except requests.exceptions.RequestException as e:
        # 통신 장애 발생 시 모의 데이터 반환
        return {
            "status": "error",
            "bypass": True,
            "message": f"SAMSAPI 연동 예외 발생: {str(e)}",
            "data": {
                "company_name": payload.company_name,
                "keyword": payload.keyword,
                "domain": payload.domain,
                "result_code": "FALLBACK_SUCCESS"
            }
        }