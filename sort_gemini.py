import os

import re

import pandas as pd


# 1. 파일 경로 설정

current_dir = os.path.dirname(os.path.abspath(__file__))

file_path = os.path.join(current_dir, 'data', 'consult_log.txt')

output_excel_path = os.path.join(current_dir, 'outputs', 'consult_summary.xlsx')


# 2. 분류 매핑 사전 정의
jls_extract_var = "불량"
reason_map = {
    jls_extract_var: "D", "재고소진": "S", "오배달": "W", "배달지연": "EL", "누락": "M",

    "배차지연": "DD", "소비기한": "E", "배민클럽": "C", "재고노출": "PD", "단순변심": "CM"

}


request_map = {

    "주문취소": "C", "문의": "Q", "환불": "R", "재배달": "RD",

    "정정배달": "CD", "새상품재배달": "ND", "확인": "I",

    "상급자요청": "E", "재판매불가상품보상": "NC"

}


result_map = {

    "주문취소": "C", "불가": "N", "전달": "P", "환불": "R",

    "재배달": "RD", "정정배달": "CD", "새상품재배달": "ND",

    "확인": "I", "보상": "CP", "수긍": "A", "쿠폰케어": "CC"

}


final_data = []

current_date = "미지정"

valid_consult_count = 0


# 3. 파일 줄 단위 읽기 및 처리

with open(file_path, 'r', encoding='utf-8-sig') as file:

    lines = file.readlines()


for raw_line in lines:

    line = raw_line.strip()


    # 1차 필터: 완전 빈 줄 제외

    if not line:

        continue


    # 날짜 줄 감지 ('2026.'으로 시작하는 경우)

    if line.startswith('2026.'):

        current_date = line

        continue


    # 줄 맨 앞의 번호 형태 제거 (예: '1. ', '1) ', '[1] ', '1- ' 등 모두 제거)

    cleaned_line = re.sub(r'^[\[\(]?\d+[\.\)\-\]\s]+', '', line).strip()


    # ★ 2차 필터(가장 확실한 조건): 한글 또는 영문이 최소 1글자도 없으면 무조건 건너뛰기!

    # (공백, 번호만 남은 줄, '/', '-', 특수문자만 남은 줄을 100% 차단)

    if not re.search(r'[가-힣a-zA-Z]', cleaned_line):

        continue


    # 전각 슬래시(／) 정규화 및 빈 항목 제외하고 분리

    normalized_line = cleaned_line.replace('／', '/')

    items = [item.strip() for item in normalized_line.split('/') if item.strip()]


    # 분리된 리스트에 실제 데이터가 없으면 건너뛰기

    if not items:

        continue


    # 정상적인 상담 데이터만 순번 카운트

    valid_consult_count += 1

    log_key = f"log_{valid_consult_count}"


    # -------------------------------------------------------------

    # 1) 대상 구분 (CPR) - 완전 일치 검사

    # -------------------------------------------------------------

    CPR = 'O'

    for item in items:

        clean = item.replace(" ", "")

        if clean == '고객':

            CPR = 'C'

            break

        elif clean == '파트너':

            CPR = 'P'

            break

        elif clean == '라이더':

            CPR = 'R'

            break


    # -------------------------------------------------------------

    # 2) 서비스 구분 (BSFO) - 완전 일치 검사

    # -------------------------------------------------------------

    BSFO = 'O'

    for item in items:

        clean = item.replace(" ", "")

        if clean == 'B마트':

            BSFO = 'B'

            break

        elif clean == '장보기':

            BSFO = 'S'

            break

        elif clean == '푸드':

            BSFO = 'F'

            break


    # -------------------------------------------------------------

    # 3) 인입 이유 (REASON)

    # -------------------------------------------------------------

    REASON = "O"

    for item in items:

        clean = item.replace(" ", "")

        if clean in reason_map:

            REASON = reason_map[clean]

            break

        elif "소비기한" in clean:

            REASON = "E"

            break


    # -------------------------------------------------------------

    # 4) 요청 사항 (REQUEST)

    # -------------------------------------------------------------

    REQUEST = "O"

    request_idx = -1

    for idx, item in enumerate(items):

        clean = item.replace(" ", "")

        if clean in request_map:

            REQUEST = request_map[clean]

            request_idx = idx

            break

        elif "새상품" in clean and "재배달" in clean:

            REQUEST = "ND_R"; request_idx = idx; break

        elif "정정" in clean and "배달" in clean:

            REQUEST = "CD_R"; request_idx = idx; break

        elif "재배달" in clean:

            REQUEST = "RD_R"; request_idx = idx; break

        elif "재판매불가" in clean or "보상" in clean:

            REQUEST = "NC_R"; request_idx = idx; break

        elif "환불" in clean:

            REQUEST = "R_R"; request_idx = idx; break

        elif "취소" in clean:

            REQUEST = "C_R"; request_idx = idx; break

        elif "문의" in clean:

            REQUEST = "Q_R"; request_idx = idx; break

        elif "확인" in clean:

            REQUEST = "I_R"; request_idx = idx; break


    # -------------------------------------------------------------

    # 5) 처리 결과 (RESULT)

    # -------------------------------------------------------------

    RESULT = "O"

    target_items = items[request_idx + 1:] if request_idx != -1 else items

    for item in target_items:

        clean = item.replace(" ", "")

        if clean in result_map:

            RESULT = result_map[clean]

            break

        elif "새상품" in clean and "재배달" in clean:

            RESULT = "ND_R"; break

        elif "정정" in clean and "배달" in clean:

            RESULT = "CD_R"; break

        elif "재배달" in clean:

            RESULT = "RD_R"; break

        elif "쿠폰" in clean:

            RESULT = "CC_R"; break

        elif "보상" in clean:

            RESULT = "CP_R"; break

        elif "환불" in clean:

            RESULT = "R_R"; break

        elif "취소" in clean:

            RESULT = "C_R"; break

        elif "불가" in clean or "거절" in clean:

            RESULT = "N_R"; break

        elif "전달" in clean:

            RESULT = "P_R"; break

        elif "수긍" in clean or "인정" in clean:

            RESULT = "A_R"; break

        elif "확인" in clean:

            RESULT = "I_R"; break


    # -------------------------------------------------------------

    # 6) 여부 체크 (ADVANCE, COUPON, TRANSFER)

    # -------------------------------------------------------------

    ADVANCE, COUPON, TRANSFER = "X", "X", "X"

    for item in items:

        clean = item.replace(" ", "")

        if "선조치" in clean:

            ADVANCE = "O"

        if "쿠폰" in clean:

            COUPON = "O"

        if "이관" in clean:

            TRANSFER = "O"


    # -------------------------------------------------------------

    # 7) 행 추가

    # -------------------------------------------------------------

    row = {

        "순번": log_key,

        "상담일자": current_date,

        "대상(CPR)": CPR,

        "서비스(BSFO)": BSFO,

        "인입이유(REASON)": REASON,

        "요청사항(REQUEST)": REQUEST,

        "처리결과(RESULT)": RESULT,

        "선조치여부": ADVANCE,

        "쿠폰발급여부": COUPON,

        "이관여부": TRANSFER,

        "원본내용": " / ".join(items)

    }

    final_data.append(row)


# 4. DataFrame 생성 및 엑셀 저장

df = pd.DataFrame(final_data)

df.to_excel(output_excel_path, index=False)


# 5. 터미널 통계 요약 출력

print("\n" + "=" * 45)

print("              📊 상담 데이터 분류 요약"
)
print("=" * 45)

print(f"✔ 총 유효 상담 건수: {len(df):,}건")

print(f"✔ 엑셀 파일 저장 완료: {output_excel_path}")

print("-" * 45)

print("📅 [날짜별 상담 건수]"
)
for date_str, count in df['상담일자'].value_counts(sort=False).items():

    print(f" - {date_str:<15} : {count:>4}건")


print("-" * 45)

print("👥 [대상별(CPR) 비중]"
)
for cpr, count in df['대상(CPR)'].value_counts().items():

    print(f" - {cpr} : {count:>4}건 ({count/len(df)*100:.1f}%)")


print("-" * 45)

print("📦 [서비스별(BSFO) 비중]"
)
for bsfo, count in df['서비스(BSFO)'].value_counts().items():

    print(f" - {bsfo} : {count:>4}건 ({count/len(df)*100:.1f}%)")

print("=" * 45 + "\n")