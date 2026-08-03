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
    """API 서버에서 기본 1000건의 데이터를 가져옵니다."""
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
                return items
    except Exception as e:
        print(f"⚠️ 통신 에러: {str(e)}")
    return []

def git_commit_and_push(file_path):
    """생성된 엑셀 파일을 깃허브 저장소에 자동으로 커밋 및 푸시합니다."""
    try:
        subprocess.run(["git", "config", "--global", "user.name", "github-actions[bot]"], check=True)
        subprocess.run(["git", "config", "--global", "user.email", "github-actions[bot]@users.noreply.github.com"], check=True)
        subprocess.run(["git", "add", file_path], check=True)
        commit_msg = f"Auto-update report: {os.path.basename(file_path)}"
        res = subprocess.run(["git", "commit", "-m", commit_msg], capture_output=True, text=True)
        if "nothing to commit" in res.stdout:
            return
        subprocess.run(["git", "push"], check=True)
        print(f"   🔄 깃허브 업로드 완료: {os.path.basename(file_path)}")
    except Exception as e:
        print(f"   ❌ 깃허브 업로드 실패: {str(e)}")

def process_and_save_split_and_accumulate():
    current_dir = os.getcwd()
    
    # =========================================================================
    # 📌 [전체 목표 기간 설정]
    # =========================================================================
    total_start_str = "2025-07-31"
    total_end_str = "2026-07-31"
    
    start_dt = datetime.strptime(total_start_str, "%Y-%m-%d")
    end_dt = datetime.strptime(total_end_str, "%Y-%m-%d")

    print(f"🚀 전체 기간({total_start_str} ~ {total_end_str}) 데이터 수집 및 분할/통합 작업을 시작합니다...\n")

    # 1. API에서 원본 데이터 전체 당겨오기
    raw_items = fetch_bizinfo_data()
    if not raw_items:
        print("❌ 수집된 데이터가 없습니다.")
        return

    # 2. 전체 데이터 파싱
    all_parsed_rows = []
    for item in raw_items:
        reg_date_str = str(item.get("pblancDe") or item.get("regDt") or item.get("creatPnttm") or item.get("creatDt") or "").strip()
        clean_date_str = reg_date_str.replace("-", "").replace(".", "")[:8]
        
        if len(clean_date_str) == 8:
            try:
                item_date = datetime.strptime(clean_date_str, "%Y%m%d")
                if not (start_dt <= item_date <= end_dt): 
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
            "공고URL": pblanc_url,
            "_parsed_date": item_date
        }
        all_parsed_rows.append(row)

    if not all_parsed_rows:
        print("⚠️ 조건에 일치하는 파싱된 데이터가 없습니다.")
        return

    df_all = pd.DataFrame(all_parsed_rows)

    # 3. 1개월(30일) 단위 구간별로 쪼개서 개별 파일(part) 저장
    current_start = start_dt
    part_num = 1
    generated_files = []
    
    while current_start <= end_dt:
        current_end = min(current_start + timedelta(days=30), end_dt)
        
        s_str = current_start.strftime("%Y-%m-%d")
        e_str = current_end.strftime("%Y-%m-%d")
        
        mask = (df_all["_parsed_date"] >= current_start) & (df_all["_parsed_date"] <= current_end)
        df_chunk = df_all.loc[mask].copy()
        
        if not df_chunk.empty:
            df_chunk = df_chunk.drop(columns=["_parsed_date"])
            file_name = f"B2G_Bizinfo_{s_str}_to_{e_str}_part{part_num}.xlsx"
            full_file_path = os.path.join(current_dir, file_name)
            
            df_chunk.to_excel(full_file_path, index=False, engine='openpyxl')
            print(f"📁 [분할 파일 생성] {file_name} (공고 {len(df_chunk)}건)")
            
            git_commit_and_push(full_file_path)
            generated_files.append(full_file_path)
            part_num += 1
            
        current_start = current_end + timedelta(days=1)

    # 4. 마지막 단계: 생성된 모든 파트 파일들을 모아서 하나의 거대한 누적 파일로 통합
    if generated_files:
        print("\n⏳ 생성된 모든 분할 파일을 취합하여 최종 누적 파일을 구축하는 중...")
        accumulated_list = []
        for f_path in generated_files:
            try:
                temp_df = pd.read_excel(f_path)
                accumulated_list.append(temp_df)
            except Exception as e:
                print(f"⚠️ 파일 읽기 경고 ({os.path.basename(f_path)}): {str(e)}")
                
        if accumulated_list:
            # 전체 합치기 및 공고명 + 등록일자 기준 중복 제거
            final_accumulated_df = pd.concat(accumulated_list, ignore_index=True)
            final_accumulated_df = final_accumulated_df.drop_duplicates(subset=["공고명", "등록일자"], keep="last")
            
            accumulated_file_name = "B2G_Bizinfo_Accumulated_Report.xlsx"
            accumulated_file_path = os.path.join(current_dir, accumulated_file_name)
            
            final_accumulated_df.to_excel(accumulated_file_path, index=False, engine='openpyxl')
            print(f"🎉 [최종 통합 누적 파일 완성] {accumulated_file_name}")
            print(f"📊 총 통합 공고 건수: {len(final_accumulated_df)}건")
            
            # 누적 파일도 깃허브에 업로드
            git_commit_and_push(accumulated_file_path)

    print(f"\n✨ 모든 작업이 성공적으로 완료되었습니다!")

if __name__ == "__main__":
    process_and_save_split_and_accumulate()
