# -*- coding: utf-8 -*-
import os
import json
import re
import subprocess
import pandas as pd
from datetime import datetime
import urllib3

# HTTPS 보안 경고 숨김
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# 기업마당 API 정보
BIZINFO_API_URL = "https://www.bizinfo.go.kr/uss/rss/bizinfoApi.do"
CRTFC_KEY = "4vc2gy"

# 전화번호 및 이메일 정규식 패턴
PHONE_PATTERN = re.compile(r"0\d{1,2}[-.\s]?\d{3,4}[-.\s]?\d{4}")
EMAIL_PATTERN = re.compile(r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}")

def parse_contact_info(raw_text):
    """
    문의처 텍스트에서 전화번호와 이메일을 정확히 추출하고,
    나머지 텍스트를 다듬어 부서명으로 분리합니다.
    """
    if not raw_text or raw_text == "정보 없음" or raw_text == "문의처 참조":
        return "정보 없음", "정보 없음", "정보 없음"

    # 1. 이메일 추출
    emails = EMAIL_PATTERN.findall(raw_text)
    email_str = ", ".join(sorted(set(emails))) if emails else "정보 없음"

    # 2. 전화번호 추출
    phones = PHONE_PATTERN.findall(raw_text)
    phone_str = ", ".join(sorted(set(phones))) if phones else "정보 없음"

    # 3. 부서명 추출 (전화번호와 이메일 문자열을 제거한 나머지 영역)
    dept_text = raw_text
    for p in phones:
        dept_text = dept_text.replace(p, "")
    for e in emails:
        dept_text = dept_text.replace(e, "")
    
    # 불필요한 기호나 공백 정리
    dept_text = re.sub(r"[,/:\-\(\)]+", " ", dept_text)
    dept_text = re.sub(r"\s+", " ", dept_text).strip()
    
    if not dept_text or len(dept_text) < 2:
        dept_text = raw_text[:50]

    return dept_text, phone_str, email_str

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
        reg_date_str = str(
            item.get("pblancDe") or 
            item.get("regDt") or 
            item.get("creatPnttm") or 
            item.get("creatDt") or 
            "정보 없음"
        ).strip()

        # 원본 문의처 텍스트 확보
        raw_inquiry = item.get("refrncNm") or item.get("inquiryTel") or item.get("telNo") or item.get("excInsttTel") or "문의처 참조"
        
        # 문의처를 부서명, 전화번호, 이메일로 각각 분할
        dept_name, phone_num, email_addr = parse_contact_info(raw_inquiry)

        # 최종 엑셀 컬럼 매핑 구조
        row = {
            "소관기관명": item.get("author") or item.get("jrsdInsttNm") or "정보 없음",
            "사업수행기관명": item.get("excInsttNm") or "정보 없음",
            "공고명": item.get("pblancNm") or item.get("title") or "제목 없음",
            "담당자": item.get("managingEditor") or item.get("chargerNm") or "정보 없음",
            "관리자": item.get("webMaster") or "정보 없음",
            "부서명": dept_name,
            "전화번호": phone_num,
            "이메일": email_addr,
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
