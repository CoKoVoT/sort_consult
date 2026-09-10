import os
import re
from collections import Counter
import pandas as pd

# =============================================================
# 1. 인입 이유별 프로파일 클래스 (ReasonProfile)
# =============================================================
class ReasonProfile:
    """
    특정 인입 이유를 대표하는 클래스.
    데이터가 들어올 때마다 속성별 빈도수와 매칭된 로그 번호를 동적으로 누적/학습합니다.
    """
    def __init__(self, reason_name):
        self.reason_name = reason_name
        self.total_cases = 0
        
        # 동적으로 학습/기록하는 카운터
        self.cpr_counter = Counter()       # 대상
        self.service_counter = Counter()   # 서비스
        self.request_counter = Counter()   # 요청 사항
        self.process_counter = Counter()   # 처리 과정
        self.result_counter = Counter()    # 처리 결과
        
        self.three_way_count = 0           # 3자 통화 진행 횟수
        self.misroute_count = 0            # 고객 오인입 횟수
        self.manual_class_count = 0        # 수동 분류 필요 횟수
        self.advance_count = 0             # 선조치 여부 횟수
        self.coupon_count = 0              # 쿠폰발급 여부 횟수
        self.transfer_count = 0            # 이관 여부 횟수

        # 각 항목에 매칭된 로그 번호 목록 저장
        self.matched_logs = []

    def learn(self, log_key, cpr, service, request, process, result, 
              three_way, misroute, manual_class, advance, coupon, transfer):
        """데이터 1건을 전달받아 속성별 통계 및 로그 번호를 갱신(학습)합니다."""
        self.total_cases += 1
        self.matched_logs.append(log_key)
        
        self.cpr_counter[cpr] += 1
        self.service_counter[service] += 1
        self.request_counter[request] += 1
        if process: self.process_counter[process] += 1
        self.result_counter[result] += 1
        
        if three_way == 'O': self.three_way_count += 1
        if misroute == 'O': self.misroute_count += 1
        if manual_class == 'O': self.manual_class_count += 1
        if advance == 'O': self.advance_count += 1
        if coupon == 'O': self.coupon_count += 1
        if transfer == 'O': self.transfer_count += 1

    def get_summary(self):
        """Sheet 1(summary)에 기록할 대표 속성 요약 딕셔너리 반환"""
        top_cpr = self.cpr_counter.most_common(1)[0][0] if self.cpr_counter else "-"
        top_service = self.service_counter.most_common(1)[0][0] if self.service_counter else "-"
        top_request = self.request_counter.most_common(1)[0][0] if self.request_counter else "-"
        top_process = self.process_counter.most_common(1)[0][0] if self.process_counter else "-"
        top_result = self.result_counter.most_common(1)[0][0] if self.result_counter else "-"

        return {
            "인입이유": self.reason_name,
            "총발생건수": self.total_cases,
            "주요대상": f"{top_cpr} ({self.cpr_counter[top_cpr]}건)" if self.cpr_counter else "-",
            "주요서비스": f"{top_service} ({self.service_counter[top_service]}건)" if self.service_counter else "-",
            "최다요청사항": f"{top_request} ({self.request_counter[top_request]}건)" if self.request_counter else "-",
            "최다처리과정": f"{top_process} ({self.process_counter[top_process]}건)" if self.process_counter else "-",
            "최다처리결과": f"{top_result} ({self.result_counter[top_result]}건)" if self.result_counter else "-",
            "3자통화비율": f"{(self.three_way_count / self.total_cases) * 100:.1f}%",
            "고객오인입비율": f"{(self.misroute_count / self.total_cases) * 100:.1f}%",
            "수동분류필요비율": f"{(self.manual_class_count / self.total_cases) * 100:.1f}%",
            "선조치비율": f"{(self.advance_count / self.total_cases) * 100:.1f}%",
            "쿠폰발급비율": f"{(self.coupon_count / self.total_cases) * 100:.1f}%",
            "이관비율": f"{(self.transfer_count / self.total_cases) * 100:.1f}%",
            "매칭_로그번호": ", ".join(self.matched_logs)
        }


# =============================================================
# 2. 자가 학습 및 클래스 관리자 (AutoProfileLearner)
# =============================================================
class AutoProfileLearner:
    def __init__(self):
        self.profiles = {}
        self.detailed_records = []

    def extract_features(self, items):
        """위치 기반 및 조건 기반 속성 추출"""
        
        # 1. 대상(CPR) 파싱: 첫 번째 위치 (items[0])
        cpr = "기타(O)"
        if len(items) > 0:
            clean = items[0].replace(" ", "")
            if clean == '고객': cpr = '고객'
            elif clean == '파트너': cpr = '파트너'
            elif clean == '라이더': cpr = '라이더'

        # 2. 서비스(BSFO) 파싱: 두 번째 위치 (items[1])
        service = "기타(O)"
        if len(items) > 1:
            clean = items[1].replace(" ", "")
            if clean == 'B마트': service = 'B마트'
            elif clean == '장보기': service = '장보기'
            elif clean == '푸드': service = '푸드'

        # 3. 인입 이유, 요청 사항 (3번째, 4번째 고정)
        reason = items[2].strip() if len(items) > 2 else ""
        request = items[3].strip() if len(items) > 3 else ""

        # 4. 처리 과정과 처리 결과 동적 분리
        # 규칙: 맨 마지막 항목은 결과, 4번째와 마지막 항목 사이는 전부 '처리 과정'
        if len(items) > 5:
            process = " / ".join(items[4:-1]).strip()
            result_raw = items[-1].strip()
        elif len(items) == 5:
            process = ""
            result_raw = items[4].strip()
        else:
            process = ""
            result_raw = ""

        # 5. 새 변수 (3자 통화, 고객 오인입)
        full_text = " ".join(items)
        
        three_way = "X"
        if "파트너 " in process or "고객" in process:
            three_way = "O"
            
        misroute = "X"
        if "고객 오인입" in full_text or "고객오인입" in full_text.replace(" ", ""):
            misroute = "O"

        # 6. '동의', '수긍' 규칙 적용 (수동 분류 여부)
        clean_consent_text = re.sub(r'제공\s*동의', '', full_text)
        
        manual_class = "O"   # 기본값: 수동 분류 필요
        result = result_raw  # 기본값: 원문 기록
        
        if "동의" in clean_consent_text or "수긍" in clean_consent_text:
            result = request      # 동의/수긍 시 요청사항과 동일하게 매핑
            manual_class = "X"    # 자동 처리되었으므로 수동분류 불필요

        # 7. 여부 플래그 (선조치, 쿠폰, 이관)
        advance = "O" if any("선조치" in it.replace(" ", "") for it in items) else "X"
        coupon = "O" if any("쿠폰" in it.replace(" ", "") for it in items) else "X"
        transfer = "O" if any("이관" in it.replace(" ", "") for it in items) else "X"

        return (reason, cpr, service, request, process, result, 
                three_way, misroute, manual_class, advance, coupon, transfer)

    def process_file(self, file_path):
        """파일 전처리 및 한 줄씩 순회하며 학습 및 로그 누적"""
        with open(file_path, 'r', encoding='utf-8-sig') as f:
            lines = f.readlines()

        valid_count = 0

        for raw_line in lines:
            line = raw_line.strip()

            if not line or line.startswith('2026.'):
                continue

            cleaned_line = re.sub(r'^[\[\(]?\d+[\.\)\-\]\s]+', '', line).strip()

            if not re.search(r'[가-힣a-zA-Z]', cleaned_line):
                continue

            normalized_line = cleaned_line.replace('／', '/')
            items = [item.strip() for item in normalized_line.split('/')]

            valid_count += 1
            log_key = f"log_{valid_count}"

            # 속성 추출
            reason, cpr, service, request, process, result, three_way, misroute, manual_class, adv, cpn, trf = self.extract_features(items)

            # 인입 이유 프로파일 자동 생성 및 누적 학습
            if reason not in self.profiles:
                self.profiles[reason] = ReasonProfile(reason)
            self.profiles[reason].learn(
                log_key, cpr, service, request, process, result, 
                three_way, misroute, manual_class, adv, cpn, trf
            )

            # 개별 상세 행 데이터 저장
            self.detailed_records.append({
                "순번": log_key,
                "대상(CPR)": cpr,
                "서비스(BSFO)": service,
                "인입이유": reason,
                "요청사항": request,
                "처리과정": process,
                "처리결과": result,
                "3자통화여부": three_way,
                "고객오인입": misroute,
                "수동분류필요": manual_class,
                "선조치여부": adv,
                "쿠폰발급여부": cpn,
                "이관여부": trf,
                "원본내용": " / ".join(items)
            })

        return valid_count


# =============================================================
# 3. 다중 시트(Sheet 1, Sheet 2) 엑셀 저장
# =============================================================
if __name__ == "__main__":
    current_dir = os.path.dirname(os.path.abspath(__file__))

    # 파일 입출력 경로 설정
    input_file_path = os.path.join(current_dir, 'data', 'consult_log.txt')
    if not os.path.exists(input_file_path):
        alt_path = os.path.join(current_dir, 'data', 'consult_log')
        if os.path.exists(alt_path):
            input_file_path = alt_path

    output_dir = os.path.join(current_dir, 'outputs')
    os.makedirs(output_dir, exist_ok=True)
    output_excel_path = os.path.join(output_dir, 'class_summary.xlsx')

    # 데이터 학습 및 파싱 실행
    learner = AutoProfileLearner()
    processed_count = learner.process_file(input_file_path)

    # DataFrame 생성
    summary_list = [profile.get_summary() for profile in learner.profiles.values()]
    df_summary = pd.DataFrame(summary_list).sort_values(by="총발생건수", ascending=False)
    df_details = pd.DataFrame(learner.detailed_records)

    # 엑셀 저장 (다중 시트)
    with pd.ExcelWriter(output_excel_path, engine='openpyxl') as writer:
        df_summary.to_excel(writer, sheet_name='summary', index=False)
        df_details.to_excel(writer, sheet_name='detail_logs', index=False)

    print("\n" + "=" * 65)
    print("      📊 다중 시트 엑셀 생성 완료 (class_summary.xlsx)")
    print("=" * 65)
    print(f"✔ 총 처리 건수: {processed_count:,}건")
    print(f"✔ 생성된 인입 이유 클래스: {len(learner.profiles)}개")
    print(f"✔ 저장 파일: {output_excel_path}")
    print("  └─ [Sheet 1: summary]     대표 속성 요약 + 매칭_로그번호")
    print("  └─ [Sheet 2: detail_logs] 개별 상담 로그 데이터프레임")
    print("=" * 65 + "\n")