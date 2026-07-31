# -*- coding: utf-8 -*-
import os
import json
import subprocess
import pandas as pd
from datetime import datetime
import urllib3

# HTTPS 보안 경고 숨김
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# 기업마당 API 정보
BIZINFO_API_URL = "https://www.bizinfo.go.kr/uss/rss/bizinfoApi.do"
CRTFC_KEY = "4vc2gy"

def fetch_all_bizinfo_data():
    """시스템 curl 명령어를 사용하여 기업마당 API 방화벽을 우회하고 데이터를 수집합니다."""
    target_url = f"{BIZINFO_API_URL}?crtfcKey={CRTFC_KEY}&dataType=json&searchCnt=1000"
    
    max_retries = 3
    for attempt in range(1, max_retries + 1):
        try:
            print(f"⏳ [시도 {attempt}/{max_retries}] curl 명령어로 기업마당 API 우회 호출 중...")
            
            curl_cmd = ["curl", "-s", "-k", "--max-time", "40", target_url]
            result = subprocess.run(curl_cmd, capture_output=True, text=True, timeout=45)
            
            if result.returncode == 0 and result.stdout:
                res_body = result.stdout.strip()
                if res_body.startswith("{") or res_body.startswith("["):
                    data = json.loads(res_body)
                    items = data.get("jsonArray") or data.get("item") or data.get("items") or []
                    print(f"🎯 [수집 성공] 총 {len(items)}건의 원본 데이터를 가져왔습니다.")
                    return items
                else:
                    print(f"⚠️ [API 응답 형식 오류]: {res_body[:100]}")
        except Exception as e:
            print(f"⚠️ [통신 경고 (시도 {attempt})]: {str(e)}")
            if attempt < max_retries:
                subprocess.run(["sleep", "3"])
                
    return []

def git_commit_and_push(file_path):
    """생성된 엑셀 파일을 깃허브 저장소에 자동으로 커밋 및 푸시합니다."""
    try:
        print("🔄 깃허브 저장소로 엑셀 파일 업로드 중...")
        subprocess.run(["git", "config", "--global", "user.name", "github-actions[bot]"], check=True)
        subprocess.run(["git", "config", "--global", "user.email", "github-actions[bot]@users.noreply.github.com"], check=True)
        
        subprocess.run(["git", "add", file_path], check=True)
        
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

    parsed_rows = []

    for item in raw_items:
        # 날짜 검증 없이 API에서 가져온 원본 등록일자 그대로 사용
        reg_date_str = str(item.get("pblancDe") or item.get("regDt") or item.get("creatDt") or "정보 없음").strip()

        # 요청하신 컬럼 매핑
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
        print("⚠️ 변환된 데이터가 없습니다.")
        return

    df = pd.DataFrame(parsed_rows)
    today = datetime.now()
    file_name = f"B2G_Bizinfo_Report_{today.strftime('%Y%m%d')}.xlsx"
    
    try:
        df.to_excel(file_name, index=False, engine='openpyxl')
        print(f"🎉 성공적으로 엑셀 파일이 저장되었습니다! 파일명: {file_name}")
        print(f"📊 총 수집 및 저장된 공고 건수: {len(df)}건")
        
        git_commit_and_push(file_name)
    except Exception as e:
        print(f"❌ 엑셀 저장 실패: {str(e)}")

if __name__ == "__main__":
    process_and_save_excel()
