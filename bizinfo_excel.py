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
    """페이지네이션을 적용하여 한도 없이 전체 데이터를 모두 수집합니다."""
    all_items = []
    page_index = 1
    display_count = 1000  # 한 번에 요청할 최대 갯수
    
    while True:
        target_url = f"{BIZINFO_API_URL}?crtfcKey={CRTFC_KEY}&dataType=json&searchCnt={display_count}&pageIndex={page_index}"
        print(f"⏳ [페이지 {page_index}] curl 명령어로 데이터 수집 중...")
        
        curl_cmd = ["curl", "-s", "-k", "--max-time", "40", target_url]
        try:
            result = subprocess.run(curl_cmd, capture_output=True, text=True, timeout=45)
            if result.returncode == 0 and result.stdout:
                res_body = result.stdout.strip()
                if res_body.startswith("{") or res_body.startswith("["):
                    data = json.loads(res_body)
                    items = data.get("jsonArray") or data.get("item") or data.get("items") or []
                    
                    if not items:
                        print("🎯 더 이상 가져올 데이터가 없습니다. 수집을 완료합니다.")
                        break
                        
                    all_items.extend(items)
                    print(f"➕ 현재 페이지에서 {len(items)}건 수집 (누적 총 {len(all_items)}건)")
                    
                    # 가져온 개수가 요청한 단위보다 적으면 마지막 페이지임
                    if len(items) < display_count:
                        break
                        
                    page_index += 1
                else:
                    print(f"⚠️ [API 응답 형식 오류]: {res_body[:100]}")
                    break
            else:
                print("⚠️ 통신 실패 또는 응답 없음")
                break
        except Exception as e:
            print(f"❌ [통신 에러]: {str(e)}")
            break
            
    print(f"🎉 [전체 수집 성공] 총 {len(all_items)}건의 데이터를 확보했습니다.")
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

    parsed_rows = []

    for item in raw_items:
        reg_date_str = str(
            item.get("pblancDe") or 
            item.get("regDt") or 
            item.get("creatPnttm") or 
            item.get("creatDt") or 
            "정보 없음"
        ).strip()

        # 문의처 파싱
        raw_inquiry = item.get("refrncNm") or item.get("inquiryTel") or item.get("telNo") or item.get("excInsttTel") or "문의처 참조"
        dept_name, phone_num, email_addr = parse_contact_info(raw_inquiry)

        # 사업신청방법 파싱
        raw_apply = item.get("reqstMthPapersCn") or item.get("reqstMth") or "정보 없음"
        apply_method_text, apply_email = parse_apply_method(raw_apply)

        # 공고 URL 추출
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
