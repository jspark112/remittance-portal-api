"""미결번호로 확정전표를 찾아 EDM 증빙을 내려받는 프로그램

사용법
  python download_edm.py APS202609220005-0001 APS202609210015-0001
  python download_edm.py --file 미결번호목록.txt        (한 줄에 미결번호 하나)
  python download_edm.py                                 (실행 후 직접 입력)

결과: edm_downloads/<오늘날짜>/<미결번호>_<거래처>/ 아래에 PDF 저장
"""
import argparse
import os
import sys
from datetime import datetime

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "api"))
from edm_service import SamsApi, resolve_pending, safe_name  # noqa: E402


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description="미결번호 → 전표번호 → EDM 증빙 다운로드")
    parser.add_argument("pending_nos", nargs="*", help="미결번호 (여러 개 가능)")
    parser.add_argument("--file", help="미결번호 목록 텍스트 파일")
    parser.add_argument("--out", default="edm_downloads", help="저장 폴더 (기본: edm_downloads)")
    parser.add_argument("--key", help="SAMSAPI Key (생략 시 기본 키 사용)")
    args = parser.parse_args()

    pending_nos = list(args.pending_nos)
    if args.file:
        with open(args.file, encoding="utf-8-sig") as f:
            pending_nos += [line.strip() for line in f if line.strip()]
    if not pending_nos:
        typed = input("미결번호를 입력하세요 (여러 개는 공백/쉼표로 구분): ")
        pending_nos = [p for p in typed.replace(",", " ").split() if p]
    if not pending_nos:
        print("입력된 미결번호가 없습니다.")
        return

    api = SamsApi(args.key)
    out_root = os.path.join(args.out, datetime.today().strftime("%Y-%m-%d"))
    print(f"미결번호 {len(pending_nos)}건 조회 중...\n")

    ok = 0
    for item in resolve_pending(api, pending_nos):
        p_no = item["pending_no"]
        if item["error"]:
            print(f"✗ {p_no} : {item['error']}")
            continue
        files = item["edm_files"]
        print(f"● {p_no}  {item['vendor_name']}  → 전표 {item['journal_number']} (확정일 {item.get('journal_date')})  증빙 {len(files)}개")
        if not files:
            print("    (EDM 에 등록된 증빙 없음)")
            continue
        folder = os.path.join(out_root, safe_name(f"{p_no}_{item['vendor_name'][:40]}"))
        os.makedirs(folder, exist_ok=True)
        for f in files:
            name = safe_name(f"{f.get('order', '')}_{f.get('filename') or 'file.pdf'}")
            try:
                with open(os.path.join(folder, name), "wb") as fp:
                    fp.write(api.download(f["downloadurl"]))
                print(f"    ✓ {name}")
            except Exception as e:
                print(f"    ✗ {name} : {e}")
        ok += 1

    print(f"\n완료: {ok}/{len(pending_nos)}건  저장 위치: {os.path.abspath(out_root)}")


if __name__ == "__main__":
    main()
