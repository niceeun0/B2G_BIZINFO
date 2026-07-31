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

# 기업마당 API 정보 (페이징 없이 한 번에 넉넉하게 대량 호출)
BIZINFO_API_URL = "https://www.bizinfo.go.kr/uss/rss/bizinfoApi.do"
CRTFC_KEY = "4vc2gy"

# 전화번호 및 이메일 정규식 패턴
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
    for p in phones:
        dept_text = dept_text.replace(p, "")
    for e in emails:
        dept_text = dept_text.replace(e, "")
    
    dept_text = re.sub(r"[,/:\-\(\)]+", " ", dept_text)
    dept_text = re.sub(r"\s+", " ", dept_text).strip()
    
    if not dept_text or len(dept_text) < 2:
        dept_text = raw_text[:50]

    return dept_text, phone_str, email_str

def parse_apply_method(raw_text):
    if not raw_text or raw_text == "정보 없음":
        return "정보 없음", "정보 없음"

    emails = EMAIL_PATTERN.findall(raw_text)
    email_str = ", ".join(sorted(set(emails))) if emails else "정보 없음"

    method_text = raw_text
    for e in emails:
        method_text = method_text.replace(e, "")
    
    method_text = re.sub(r"[,/:\-]+$", "", method_text).strip()
    if not method_text:
        method_text = "신청방법 참조"

    return method_text, email_str

def fetch_all_bizinfo_data():
    """페이징 없이 단일 호출로 최대 3000건의 데이터를 안전하게 수집합니다."""
    target_url = f"{BIZINFO_API_URL}?crtfcKey={CRTFC_KEY}&dataType=json&searchCnt=3000"
    
    max_retries = 3
    for attempt in range(1, max_retries + 1):
        try:
            print(f"⏳ [시도 {attempt}/{max_retries}] 기업마당 API 대량 호출 중 (searchCnt=3000)...")
            
            curl_cmd = [
                "curl", "-s", "-k", 
                "--max-time", "60", 
                "-A", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36", 
                target_url
            ]
            result = subprocess.run(curl_cmd, capture_output=True, text=True, timeout=70)
            
            if result.returncode == 0 and result.stdout:
                res_body = result.stdout.strip()
                if res_body.startswith("{") or res_body.startswith("["):
                    data = json.loads(res_body)
                    items = data.get("jsonArray") or data.get("item") or data.get("items") or []
                    print(f"🎯 [수집 성공] 총 {len(items)}건의 원본 데이터를 가져왔습니다.")
                    return items
                else:
                    print(f"⚠️ [API 응답 오류]: {res_body[:100]}")
            else:
                print(f"⚠️ [통신 실패]: {result.stderr}")
        except Exception as e:
            print(f"⚠️ [통신 에러 (시도 {attempt})]: {str(e)}")
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

    # =========================================================================
    # 📌 [원하시는 기간을 직접 입력하세요] 형식: "YYYY-MM-DD"
    # =========================================================================
    custom_start_date = "2025-07-31"  # 시작일 (예: 1년 전 또는 특정 시작일)
    custom_end_date = "2026-07-31"    # 종료일 (예: 오늘 또는 특정 종료일)
    
    start_date = datetime.strptime(custom_start_date, "%Y-%m-%d")
    end_date = datetime.strptime(custom_end_date, "%Y-%m-%d")
    
    print(f"📅 직접 입력한 필터링 기간: {start_date.strftime('%Y-%m-%d')} ~ {end_date.strftime('%Y-%m-%d')}")

    parsed_rows = []

    for item in raw_items:
        reg_date_str = str(
            item.get("pblancDe") or 
            item.get("regDt") or 
            item.get("creatPnttm") or 
            item.get("creatDt") or 
            ""
        ).strip()
        
        clean_date_str = reg_date_str.replace("-", "").replace(".", "")[:8]
        
        if len(clean_date_str) == 8:
            try:
                item_date = datetime.strptime(clean_date_str, "%Y%m%d")
                # 직접 입력한 시작일 ~ 종료일 사이 데이터만 통과
                if not (start_date <= item_date <= end_date):
                    continue
            except ValueError:
                continue
        else:
            continue

        # (이하 문의처 및 사업신청방법 파싱 로직 동일)
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

    if not parsed_rows:
        print("⚠️ 입력하신 기간에 일치하는 공고 데이터가 없습니다.")
        return

    df = pd.DataFrame(parsed_rows)
    # 파일명에 오늘 날짜나 지정일 표기
    file_name = f"B2G_Bizinfo_Report_{datetime.now().strftime('%Y%m%d')}.xlsx"
    
    try:
        df.to_excel(file_name, index=False, engine='openpyxl')
        print(f"🎉 성공적으로 엑셀 파일이 저장되었습니다! 파일명: {file_name}")
        print(f"📊 총 수집 및 저장된 공고 건수: {len(df)}건")
        
        git_commit_and_push(file_name)
    except Exception as e:
        print(f"❌ 엑셀 저장 실패: {str(e)}")

if __name__ == "__main__":
    process_and_save_excel()
