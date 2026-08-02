# -*- coding: utf-8 -*-
import os
import json
import re
import subprocess
import pandas as pd
from datetime import datetime, timedelta
import urllib3

# HTTPS 보안 경고 숨김
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# 기업마당 API 정보
BIZINFO_API_URL = "https://www.bizinfo.go.kr/uss/rss/bizinfoApi.do"
CRTFC_KEY = "4vc2gy"

PHONE_PATTERN = re.compile(r"0\d{1,2}[-.\s]?\d{3,4}[-.\s]?\d{4}")
EMAIL_PATTERN = re.compile(r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}")

def parse_contact_info(raw_text):
    if not raw_text or raw_text == "정보 없음" or raw_text == "문의처 참조":
        return "정보 없음", "정보 없음", "정보 없음"
    emails = EMAIL_PATTERN.findall(raw_text)
    email_str = ", ".join(sorted(set(emails))) if emails else "정보 없음"
    phones = PHONE_PATTERN.findall(raw_text)
    phone_str = ", ".join(sorted(set(phones))) if phones else "정보 없음"
    dept_text = raw_text
    for p in phones: dept_text = dept_text.replace(p, "")
    for e in emails: dept_text = dept_text.replace(e, "")
    dept_text = re.sub(r"[,/:\-\(\)]+", " ", dept_text)
    dept_text = re.sub(r"\s+", " ", dept_text).strip()
    if not dept_text or len(dept_text) < 2: dept_text = raw_text[:50]
    return dept_text, phone_str, email_str

def parse_apply_method(raw_text):
    if not raw_text or raw_text == "정보 없음": return "정보 없음", "정보 없음"
    emails = EMAIL_PATTERN.findall(raw_text)
    email_str = ", ".join(sorted(set(emails))) if emails else "정보 없음"
    method_text = raw_text
    for e in emails: method_text = method_text.replace(e, "")
    method_text = re.sub(r"[,/:\-]+$", "", method_text).strip()
    if not method_text: method_text = "신청방법 참조"
    return method_text, email_str

def fetch_bizinfo_data():
    """안정적으로 데이터를 수집합니다."""
    target_url = f"{BIZINFO_API_URL}?crtfcKey={CRTFC_KEY}&dataType=json&searchCnt=1000"
    
    curl_cmd = [
        "curl", "-s", "-k", 
        "--max-time", "45", 
        "-A", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36", 
        target_url
    ]
    
    try:
        result = subprocess.run(curl_cmd, capture_output=True, text=True, timeout=50)
        if result.returncode == 0 and result.stdout:
            res_body = result.stdout.strip()
            if res_body.startswith("{") or res_body.startswith("["):
                data = json.loads(res_body)
                items = data.get("jsonArray") or data.get("item") or data.get("items") or []
                print(f"🎯 [수집 성공] 총 {len(items)}건의 데이터를 가져왔습니다.")
                return items
    except Exception as e:
        print(f"❌ [통신 에러]: {str(e)}")
    return []

def git_commit_and_push(file_path):
    """생성된 엑셀 파일을 깃허브 저장소에 자동으로 커밋 및 푸시합니다."""
    try:
        print("🔄 깃허브 저장소로 엑셀 파일 업로드 중...")
        subprocess.run(["git", "config", "--global", "user.name", "github-actions[bot]"], check=True)
        subprocess.run(["git", "config", "--global", "user.email", "github-actions[bot]@users.noreply.github.com"], check=True)
        subprocess.run(["git", "add", file_path], check=True)
        commit_msg = f"Auto-accumulate Excel report: {datetime.now().strftime('%Y-%m-%d %H:%M')}"
        res = subprocess.run(["git", "commit", "-m", commit_msg], capture_output=True, text=True)
        if "nothing to commit" in res.stdout:
            print("ℹ️ 변경된 엑셀 내용이 없어 커밋을 생략합니다.")
            return
        subprocess.run(["git", "push"], check=True)
        print("🎉 깃허브 업로드 완료!")
    except Exception as e:
        print(f"❌ 깃허브 업로드 실패: {str(e)}")

def process_and_save_excel():
    raw_items = fetch_bizinfo_data()
    if not raw_items:
        print("❌ 처리할 데이터가 없습니다.")
        return

    custom_start_date = "2025-07-31"
    custom_end_date = "2026-07-31"
    start_date = datetime.strptime(custom_start_date, "%Y-%m-%d")
    end_date = datetime.strptime(custom_end_date, "%Y-%m-%d")

    # 고정된 파일명 사용 (누적 저장을 위해 날짜별이 아닌 고정 파일명 추천, 혹은 기존 파일 탐색)
    file_name = "B2G_Bizinfo_Accumulated_Report.xlsx"
    
    # 1. 기존에 저장되어 있던 엑셀 파일이 있다면 불러오기
    existing_df = pd.DataFrame()
    if os.path.exists(file_name):
        try:
            existing_df = pd.read_excel(file_name)
            print(f"📂 기존에 누적된 엑셀 데이터 {len(existing_df)}건을 불러왔습니다.")
        except Exception as e:
            print(f"⚠️ 기존 엑셀 파일 읽기 실패 (새로 생성합니다): {str(e)}")

    # 2. 이번에 새로 가져온 데이터 파싱
    parsed_rows = []
    for item in raw_items:
        reg_date_str = str(item.get("pblancDe") or item.get("regDt") or item.get("creatPnttm") or item.get("creatDt") or "").strip()
        clean_date_str = reg_date_str.replace("-", "").replace(".", "")[:8]
        if len(clean_date_str) == 8:
            try:
                item_date = datetime.strptime(clean_date_str, "%Y%m%d")
                if not (start_date <= item_date <= end_date): continue
            except ValueError: continue
        else: continue

        raw_inquiry = item.get("refrncNm") or item.get("inquiryTel") or item.get("telNo") or item.get("excInsttTel") or "문의처 참조"
        dept_name, phone_num, email_addr = parse_contact_info(raw_inquiry)
        raw_apply = item.get("reqstMthPapersCn") or item.get("reqstMth") or "정보 없음"
        apply_method_text, apply_email = parse_apply_method(raw_apply)
        pblanc_url = item.get("pblancUrl") or item.get("rceptEngnHmpgUrl") or "링크 없음"

        row = {
            "소관기관명": item.get("author") or item.get("jrsdInsttNm") or "정보 없음",
            "사업수행기관명": item.get("excInsttNm") or "정보 없음",
            "공고명": item.get("pblancNm") or item.get("title") or "제목 없음",
            "문의처": dept_name,
            "문의처(전화번호)": phone_num,
            "문의처(이메일주소)": email_addr,
            "사업신청방법": apply_method_text,
            "사업신청 방법(이메일주소)": apply_email,
            "등록일자": reg_date_str,
            "공고URL": pblanc_url
        }
        parsed_rows.append(row)

    new_df = pd.DataFrame(parsed_rows)

    # 3. 기존 데이터와 신규 데이터를 합치기 (공고명 기준 중복 제거)
    if not existing_df.empty:
        combined_df = pd.concat([existing_df, new_df]).drop_duplicates(subset=["공고명", "등록일자"], keep="last")
    else:
        combined_df = new_df

    # 4. 엑셀로 저장 및 업로드
    try:
        combined_df.to_excel(file_name, index=False, engine='openpyxl')
        print(f"🎉 누적 저장 완료! 파일명: {file_name}")
        print(f"📊 총 누적 공고 건수: {len(combined_df)}건 (신규 추가: {len(new_df)}건 중 중복 제외 반영)")
        
        git_commit_and_push(file_name)
    except Exception as e:
        print(f"❌ 엑셀 저장 실패: {str(e)}")

if __name__ == "__main__":
    process_and_save_excel()
