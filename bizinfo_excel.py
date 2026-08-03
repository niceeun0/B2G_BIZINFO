import requests
import pandas as pd
from datetime import datetime, timedelta

# API 설정
url = "https://apis.data.go.kr/1421000/bizinfo/sptngInfoList" # 또는 실제 사용하는 세부 서비스 명칭
service_key = "2caa4e5037386a3f7ba07e36cfd04811967b31ce15d0ccfeec0bc07a5a1e6eb0"

# 조회 기간 설정 (2025년 8월 1일 ~ 2026년 8월 31일)
start_date = "20250801"
end_date = "20260831"

all_items = []
page_index = 1
display_count = 100 # 한 번에 가져올 데이터 수

print(f"[{start_date} ~ {end_date}] 기간 동안의 기업마당 지원사업 데이터 수집을 시작합니다...")

while True:
    params = {
        'serviceKey': service_key,
        'pageIndex': page_index,
        'display': display_count,
        'schRegYmdFrom': start_date, # 등록일자 시작
        'schRegYmdTo': end_date      # 등록일자 종료
        # 필요시 API 명세에 따른 추가 파라미터(예: dataType=JSON 등) 기입
    }
    
    try:
        response = requests.get(url, params=params)
        
        # 응답 상태 확인
        if response.status_code != 200:
            print(f"API 요청 실패 (Status Code: {response.status_code})")
            break
            
        # JSON 또는 XML 파싱 (공공데이터포털은 XML로 올 수도 있으므로 xmltodict 사용 고려)
        # 여기서는 JSON 응답 가정
        data = response.json()
        
        # 데이터 구조에 맞게 리스트 추출 (API 응답 스키마에 따라 경로 수정 필요)
        items = data.get('response', {}).get('body', {}).get('items', {}).get('item', [])
        
        if not items:
            print(f"더 이상 조회할 데이터가 없습니다. (마지막 페이지: {page_index})")
            break
            
        if isinstance(items, dict): # 단건일 경우
            items = [items]
            
        all_items.extend(items)
        print(f"현재 {page_index}페이지 수집 완료 (누적 건수: {len(all_items)}건)")
        
        # 다음 페이지로 이동
        page_index += 1
        
        # 안전장치 (무한 루프 방지용 최대 페이지 제한, 예: 100페이지)
        if page_index > 100:
            break
            
    except Exception as e:
        print(f"에러 발생: {e}")
        break

# 수집된 데이터를 DataFrame으로 변환 후 엑셀 저장
if all_items:
    df_result = pd.DataFrame(all_items)
    output_filename = "기업마당_25년8월_26년8월_전수데이터.xlsx"
    df_result.to_excel(output_filename, index=False)
    print(f"수집 완료! '{output_filename}' 파일로 저장되었습니다. 총 {len(df_result)}건")
else:
    print("수집된 데이터가 없습니다. 파라미터명이나 응답 포맷(XML/JSON)을 확인해 주세요.")
