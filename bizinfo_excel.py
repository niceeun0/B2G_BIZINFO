# -*- coding: utf-8 -*-
import os
import subprocess
import requests
import pandas as pd
from datetime import datetime, timedelta
import urllib3

# HTTPS 보안 경고 숨김
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# 기업마당 API 정보
BIZINFO_API_URL = "https://www.bizinfo.go.kr/uss/rss/bizinfoApi.do"
CRTFC_KEY = "4vc2gy"

def fetch_all_bizinfo_data():
    """기업마당 API에서 데이터를 대량으로 수집합니다."""
    target_url = f"{BIZINFO_API_URL}?crtfcKey={CRTFC_KEY}&dataType=json&searchCnt=1000"
    
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36",
        "Accept": "application/json"
    }

    try:
        print("⏳ 기업마당 API에서 데이터를 불러오는 중입니다...")
        res = requests.get(target_url, headers=headers, timeout=30, verify=False)
        
        if res.status_code == 200:
            data = res.json()
            items = data.get("jsonArray") or data.get("item") or data.get("items") or []
            print(f"🎯 [수집 성공] 총 {len(items)}건의 원본 데이터를 가져왔습니다.")
            return items
        else:
            print(f"⚠️ [API 응답 오류] 상태 코드: {res.status_code}")
    except Exception as e:
        print(f"❌ [통신 에러 발생]: {str(e)}")
        
    return []

def git_commit_and_push(file_path):
    """생성된 엑셀 파일을 깃허브 저장소에 자동으로 커밋 및 푸시합니다."""
    try:
        print("🔄 깃허브 저장소로 엑셀 파일 업로드 중...")
        subprocess.run(["git", "config", "--global", "user.name", "github-actions[bot]"], check=True)
        subprocess.run(["git", "config", "--global", "user.email", "github-actions[bot]@users.noreply.github.com"], check=True)
        
        subprocess.run(["git", "add", file_path], check=True)
        
        # 커밋 메시지에 오늘 날짜 반영
        commit_msg = f"Auto-update Excel report: {datetime.now().strftime('%Y-%m-%d')}"
        res = subprocess.run(["git", "commit", "-m", commit_msg], capture_output=True, text=True)
        
        if "nothing to commit" in res.stdout:
            print("ℹ️ 변경된 엑셀 내용이 없어 커밋을 생략합니다.")
            return

        subprocess.run(["git", "push"], check=True)
        print("🎉 깃허브 저장소로 엑셀 파일 업로드 완료!")
    except Exception as e:
        print(f"❌ 깃허브 자동 업로드 실패: {str(e)}")

def process_and_save_excel():
    raw_items = fetch_all_bizinfo_data()
    if not raw_items:
        print("❌ 처리할 데이터가 없습니다.")
        return

    # 기준일 계산: 오늘부터 정확히 1년 전
    today = datetime.now()
    one_year_ago = today - timedelta(days=365)
    print(f"📅 필터링 기간: {one_year_ago.strftime('%Y-%m-%d')} ~ {today.strftime('%Y-%m-%d')}")

    parsed_rows = []

    for item in raw_items:
        # 1. 등록일자 추출 및 표준화
        reg_date_str = str(item.get("pblancDe") or item.get("regDt") or item.get("creatDt") or "").strip()
        clean_date_str = reg_date_str.replace("-", "").replace(".", "")[:8]
        
        if len(clean_date_str) == 8:
            try:
                item_date = datetime.strptime(clean_date_str, "%Y%m%d")
                # 2. 지금으로부터 1년 전 ~ 오늘 사이 데이터만 필터링
                if not (one_year_ago <= item_date <= today):
                    continue
            except ValueError:
                continue
        else:
            continue

        # 3. 요청하신 컬럼 매핑
        row = {
            "소관기관명": item.get("author") or item.get("jrsdInsttNm") or "정보 없음",
            "사업수행기관명": item.get("excInsttNm") or "정보 없음",
            "공고명": item.get("pblancNm") or item.get("title") or "제목 없음",
            "담당자": item.get("managingEditor") or item.get("chargerNm") or "정보 없음",
            "관리자": item.get("webMaster") or "정보 없음",
            "문의처": item.get("inquiryTel") or item.get("telNo") or item.get("excInsttTel") or "문의처 참조",
            "기관담당자정보": item.get("chargerInfo") or item.get("deptNm") or "정보 없음",
            "등록일자": reg_date_str
        }
        parsed_rows.append(row)

    if not parsed_rows:
        print("⚠️ 조건(최근 1년 이내)에 일치하는 공고 데이터가 없습니다.")
        return

    # 4. DataFrame 생성 및 엑셀 저장
    df = pd.DataFrame(parsed_rows)
    file_name = f"B2G_Bizinfo_Report_{today.strftime('%Y%m%d')}.xlsx"
    
    try:
        df.to_excel(file_name, index=False, engine='openpyxl')
        print(f"🎉 성공적으로 엑셀 파일이 저장되었습니다! 파일명: {file_name}")
        print(f"📊 총 수집 및 저장된 공고 건수: {len(df)}건")
        
        # 5. 깃허브에 자동 업로드 실행
        git_commit_and_push(file_name)
    except Exception as e:
        print(f"❌ 엑셀 저장 실패: {str(e)}")

if __name__ == "__main__":
    process_and_save_excel()
