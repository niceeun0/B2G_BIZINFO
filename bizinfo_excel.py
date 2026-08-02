# -*- coding: utf-8 -*-
import os
import json
import re
import subprocess
import time
import pandas as pd
from datetime import datetime, timedelta
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

def fetch_data_by_period(start_str, end_str):
    """특정 기간(시작일~종료일)을 지정하여 API 데이터를 호출합니다."""
    target_url = (
        f"{BIZINFO_API_URL}?crtfcKey={CRTFC_KEY}&dataType=json&searchCnt=1000"
        f"&searchBeginDe={start_str}&searchEndDe={end_str}"
    )
    
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
                return items
    except Exception as e:
        print(f"⚠️ 통신 에러 ({start_str} ~ {end_str}): {str(e)}")
    return []

def fetch_all_bizinfo_data(custom_start_str, custom_end_str):
    """
    5000개 이상의 대량 데이터를 누락 없이 수집하기 위해, 
    전체 기간을 1개월(30일) 단위로 아주 잘게 쪼개어 각각 수집 후 합칩니다.
    """
    start_dt = datetime.strptime(custom_start_str, "%Y-%m-%d")
    end_dt = datetime.strptime(custom_end_str, "%Y-%m-%d")
    
    all_items = []
    seen_ids = set()
    
    current_start = start_dt
    while current_start <= end_dt:
        # 1개월(30일) 단위로 구간 설정
        current_end = min(current_start + timedelta(days=30), end_dt)
        
        s_str = current_start.strftime("%Y%m%d")
        e_str = current_end.strftime("%Y%m%d")
        
        print(f"⏳ 월별 구간 수집 중: {current_start.strftime('%Y-%m-%d')} ~ {current_end.strftime('%Y-%m-%d')}")
        
        items = fetch_data_by_period(s_str, e_str)
        added_count = 0
        for item in items:
            item_id = item.get("pblancId") or item.get("pblancNm")
            if item_id not in seen_ids:
                seen_ids.add(item_id)
                all_items.append(item)
                added_count += 1
                
        print(f"➕ 해당 구간에서 {added_count}건 수집 (누적 총 {len(all_items)}건)")
        
        # 다음 구간으로 이동
        current_start = current_end + timedelta(days=1)
        time.sleep(2) # 서버 부하 방지
        
    print(f"🎯 [최종 수집 완료] 총 {len(all_items)}건의 고유 데이터를 확보했습니다.")
    return all_items

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
    # =========================================================================
    # 📌 [원하시는 기간을 직접 입력하세요] 형식: "YYYY-MM-DD"
    # =========================================================================
    custom_start_date = "2025-07-31"  # 시작일
    custom_end_date = "2026-07-31"    # 종료일
    
    print(f"📅 전체 목표 기간: {custom_start_date} ~ {custom_end_date}")
    
    raw_items = fetch_all_bizinfo_data(custom_start_date, custom_end_date)
    if not raw_items:
        print("❌ 처리할 데이터가 없습니다.")
        return

    start_date = datetime.strptime(custom_start_date, "%Y-%m-%d")
    end_date = datetime.strptime(custom_end_date, "%Y-%m-%d")

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
                if not (start_date <= item_date <= end_date):
                    continue
            except ValueError:
                continue
        else:
            continue

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
