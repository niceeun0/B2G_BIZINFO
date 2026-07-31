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

def fetch_all_bizinfo_data():
    """서버 과부하를 막기 위해 1000개씩 페이지를 나누어 안전하게 전체 데이터를 수집합니다."""
    all_items = []
    display_count = 1000  # 한 번에 1000개씩 요청
    max_pages = 3         # 최대 3페이지(총 3000개)까지 순차 탐색
    
    for page in range(1, max_pages + 1):
        target_url = f"{BIZINFO_API_URL}?crtfcKey={CRTFC_KEY}&dataType=json&searchCnt={display_count}&pageIndex={page}"
        print(f"⏳ [페이지 {page}/{max_pages}] 기업마당 API 안전 호출 중...")
        
        curl_cmd = [
            "curl", "-s", "-k", 
            "--max-time", "40", 
            "-A", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36", 
            target_url
        ]
        
        try:
            result = subprocess.run(curl_cmd, capture_output=True, text=True, timeout=45)
            if result.returncode == 0 and result.stdout:
                res_body = result.stdout.strip()
                if res_body.startswith("{") or res_body.startswith("["):
                    data = json.loads(res_body)
                    items = data.get("jsonArray") or data.get("item") or data.get("items") or []
                    
                    if not items:
                        print(f"🎯 [수집 완료] 더 이상 가져올 데이터가 없습니다. (총 누적: {len(all_items)}건)")
                        break
                        
                    all_items.extend(items)
                    print(f"➕ [성공] {len(items)}건 추가 (누적 총 {len(all_items)}건)")
                    
                    if len(items) < display_count:
                        print("🎯 [수집 완료] 마지막 페이지 도달.")
                        break
                else:
                    print(f"⚠️ [API 응답 오류]: {res_body[:100]}")
                    break
            else:
                print(f"⚠️ [통신 실패]: {result.stderr}")
                break
        except Exception as e:
            print(f"⚠️ [통신 에러]: {str(e)}")
            break
            
        # 서버 차단 방지를 위해 페이지 간 3초 대기
        if page < max_pages:
            print("💤 서버 안정화를 위해 3초 대기 중...")
            time.sleep(3)
            
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
    raw_items = fetch_all_bizinfo_data()
    if not raw_items:
        print("❌ 처리할 데이터가 없습니다.")
        return

    # =========================================================================
    # 📌 [원하시는 기간을 직접 입력하세요] 형식: "YYYY-MM-DD"
    # =========================================================================
    custom_start_date = "2025-07-31"  # 시작일
    custom_end_date = "2026-07-31"    # 종료일 (현재)
    
    start_date = datetime.strptime(custom_start_date, "%Y-%m-%d")
    end_date = datetime.strptime(custom_end_date, "%Y-%m-%d")
    
    print(f"📅 필터링 기간 설정: {start_date.strftime('%Y-%m-%d')} ~ {end_date.strftime('%Y-%m-%d')}")

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
                # 입력하신 기간 내에 포함된 데이터만 통과
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
