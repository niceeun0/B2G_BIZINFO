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

def fetch_data_by_period(start_str, end_str):
    """API 서버에 명시적인 날짜 구간(YYYYMMDD)을 전달하여 과거 데이터를 강제로 끌어옵니다."""
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
    # 📌 [전체 목표 기간 설정] 2025년 8월 1일 ~ 2026년 8월 3일
    # =========================================================================
    total_start_str = "2025-08-01"
    total_end_str = "2026-08-03"
    
    start_dt = datetime.strptime(total_start_str, "%Y-%m-%d")
    end_dt = datetime.strptime(total_end_str, "%Y-%m-%d")

    print(f"🚀 전체 기간({total_start_str} ~ {total_end_str}) 월별 구간별 API 집중 수집을 시작합니다...\n")

    current_start = start_dt
    part_num = 1
    generated_files = []
    all_parsed_rows = []

    # 월별(30일 단위)로 쪼개서 API에 직접 날짜 조건을 주고 가져오기
    while current_start <= end_dt:
        current_end = min(current_start + timedelta(days=30), end_dt)
        
        s_date_str = current_start.strftime("%Y-%m-%d")
        e_date_str = current_end.strftime("%Y-%m-%d")
        
        s_param = current_start.strftime("%Y%m%d")
        e_param = current_end.strftime("%Y%m%d")
        
        print(f"⏳ 구간 수집 중 ({s_date_str} ~ {e_date_str})...")
        raw_items = fetch_data_by_period(s_param, e_param)
        print(f"   👉 API 응답: {len(raw_items)}건 수집됨")

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
                "공고URL": pblanc_url
            }
            all_parsed_rows.append(row)

        # 구간별 개별 파일 저장
        if raw_items:
            df_chunk = pd.DataFrame([r for r in all_parsed_rows if s_date_str <= str(r["등록일자"]).replace("-", "").replace(".", "")[:4] + "-" + str(r["등록일자"]).replace("-", "").replace(".", "")[4:6] + "-" + str(r["등록일자"]).replace("-", "").replace(".", "")[6:8] <= e_date_str])
            # 위 필터링 대신 해당 구간 아이템들로만 척척 담기 위해 임시 데이터프레임 생성
            chunk_rows = []
            for item in raw_items:
                reg_date_str = str(item.get("pblancDe") or item.get("regDt") or item.get("creatPnttm") or item.get("creatDt") or "").strip()
                clean_date_str = reg_date_str.replace("-", "").replace(".", "")[:8]
                if len(clean_date_str) == 8:
                    try:
                        item_date = datetime.strptime(clean_date_str, "%Y%m%d")
                        if not (current_start <= item_date <= current_end): continue
                    except ValueError: continue
                else: continue

                raw_inquiry = item.get("refrncNm") or item.get("inquiryTel") or item.get("telNo") or item.get("excInsttTel") or "문의처 참조"
                dept_name, phone_num, email_addr = parse_contact_info(raw_inquiry)
                raw_apply = item.get("reqstMthPapersCn") or item.get("reqstMth") or "정보 없음"
                apply_method_text, apply_email = parse_apply_method(raw_apply)
                pblanc_url = item.get("pblancUrl") or item.get("rceptEngnHmpgUrl") or "링크 없음"

                chunk_rows.append({
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
                })
            
            if chunk_rows:
                df_chunk = pd.DataFrame(chunk_rows)
                file_name = f"B2G_Bizinfo_{s_date_str}_to_{e_date_str}_part{part_num}.xlsx"
                full_file_path = os.path.join(current_dir, file_name)
                df_chunk.to_excel(full_file_path, index=False, engine='openpyxl')
                print(f"   📁 [분할 파일 저장] {file_name} ({len(df_chunk)}건)")
                git_commit_and_push(full_file_path)
                generated_files.append(full_file_path)
                part_num += 1

        current_start = current_end + timedelta(days=1)
        time.sleep(1)

    # 최종 통합 누적 파일 생성
    if all_parsed_rows:
        final_df = pd.DataFrame(all_parsed_rows)
        final_df = final_df.drop_duplicates(subset=["공고명", "등록일자"], keep="last")
        
        accumulated_file_name = "B2G_Bizinfo_Accumulated_Report.xlsx"
        accumulated_file_path = os.path.join(current_dir, accumulated_file_name)
        
        final_df.to_excel(accumulated_file_path, index=False, engine='openpyxl')
        print(f"\n🎉 [최종 통합 누적 파일 완성] {accumulated_file_name}")
        print(f"📊 총 통합 공고 건수: {len(final_df)}건")
        git_commit_and_push(accumulated_file_path)
    else:
        print("\n❌ 수집된 최종 데이터가 없습니다.")

    print(f"\n✨ 모든 작업이 완료되었습니다!")

if __name__ == "__main__":
    process_and_save_split_and_accumulate()
