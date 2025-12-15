# -*- coding: utf-8 -*-
"""
Created on Mon Dec 15 17:06:15 2025

@author: SW25
"""

# -*- coding: utf-8 -*-
"""
Created on Sun Dec 14 21:26:07 2025
@author: algus
"""

import streamlit as st
import pandas as pd
import io

# 페이지 설정
st.set_page_config(page_title="SW 기초교과목 이수자 분석 도구", layout="wide")

st.title("SW기초교과목 이수자 분석 프로그램")
st.markdown("""
엑셀(CSV) 파일을 업로드하면 다음 내용을 분석합니다:
1. **통계 분석**: 과목별 분반 수, 수강생 수, 1학년 수강생 수 (합계 포함)
2. **데이터 정리**: 전체 이수자, 1학년 이수자, 개설 분반 리스트 엑셀 다운로드
""")

# 1. 파일 업로드
uploaded_file = st.file_uploader("CSV 파일을 업로드하세요 (.CSV)", type=['csv'])

if uploaded_file is not None:
    try:
        # 데이터 로드
        df = pd.read_csv(uploaded_file)
        
        # --- 필수 컬럼 설정 ---
        col_id = '학번'
        col_grade = '학년(수강시점)'
        col_subject = '교과목명'
        col_class = '분반'
        col_semester = '학기'
        
        # 필수 컬럼 확인
        required_cols = [col_id, col_grade, col_subject, col_class, col_semester]
        if not all(col in df.columns for col in required_cols):
            st.error(f"엑셀 파일에 다음 컬럼이 반드시 포함되어야 합니다: {required_cols}")
            st.stop()

        # --- 데이터 전처리 ---
        # 학년 데이터 숫자형 변환 (예: "1학년" -> 1)
        df[col_grade] = df[col_grade].astype(str).str.extract(r'(\d+)').astype(float).fillna(0).astype(int)

        # 2. 정렬 로직 (사용자 지정 순서)
        custom_order = ['컴퓨팅사고와인공지능', '기초컴퓨터프로그래밍','IT환경에서의개인정보보호','멀티미디어의이해와활용','디지털리터러시의 이해와 활용','컴퓨터 시뮬레이션', '컴퓨터프로그래밍입문']
        
        existing_subjects = df[col_subject].unique()
        remaining_subjects = [x for x in existing_subjects if x not in custom_order]
        final_order = custom_order + remaining_subjects
        
        # Categorical 변환 (정렬 순서 적용)
        df[col_subject] = pd.Categorical(df[col_subject], categories=final_order, ordered=True)
        
        # 전체 데이터 정렬 (학기 -> 학년 -> 과목 순)
        #df_sorted = df.sort_values(by=[col_semester, col_grade, col_subject], ascending=[True, True, True])
        df_sorted = df.sort_values(by=[col_grade, col_subject], ascending=[True, True])

        # ---------------------------------------------------------
        # [분석 1] 개설 분반 분석 (개설분반리스트 시트용)
        # ---------------------------------------------------------
        # [과목, 학기, 분반] 고유 조합 추출 = 실제 개설된 강좌 목록
        unique_sections = df_sorted[[col_subject, col_semester, col_class]].drop_duplicates()
        
        # 엑셀 저장용 리스트 (보기 좋게 정렬)
        section_list_df = unique_sections.sort_values(by=[col_subject, col_semester, col_class])
        
        # 통계용: 과목별 분반 개수 카운트
        class_counts_df = unique_sections.groupby(col_subject, observed=True).size().to_frame(name='개설분반수')

        # ---------------------------------------------------------
        # [분석 2] 수강생 분석 (전체이수자 시트용)
        # ---------------------------------------------------------
        # [학번] 기준으로 중복 제거 -> 한 학생이 여러 번(재수강 등) 데이터에 있어도 1번으로 처리
        df_dedup = df_sorted.drop_duplicates(subset=[col_id])
        
        # 인덱스 재설정 (1번부터)
        df_dedup = df_dedup.reset_index(drop=True)
        df_dedup.index = df_dedup.index + 1
        
        # ---------------------------------------------------------
        # [분석 3] 1학년 수강생 분석 (1학년이수자 시트용)
        # ---------------------------------------------------------
        # 위에서 정제된 df_dedup에서 학년이 1인 경우만 추출
        df_freshman = df_dedup[df_dedup[col_grade] == 1].copy()
        df_freshman = df_freshman.reset_index(drop=True)
        df_freshman.index = df_freshman.index + 1

        # ---------------------------------------------------------
        # [분석 4] 통계 테이블 생성
        # ---------------------------------------------------------
        # 과목별 수강생 수 집계
        student_counts_df = df_dedup.groupby(col_subject, observed=True).agg(
            전체수강생=(col_id, 'count'),
            일학년수강생=(col_grade, lambda x: (x == 1).sum())
        )

        # 분반 수 + 수강생 수 병합
        final_stats = pd.merge(
            class_counts_df, 
            student_counts_df, 
            left_index=True, 
            right_index=True, 
            how='outer'
        )
        final_stats = final_stats.fillna(0).astype(int)
        
        # 컬럼명 설정 (요청사항 반영)
        final_stats.columns = ['개설분반수', '전체수강생(중복자제거후)', '일학년수강생(중복자제거후)']

        # ---------------------------------------------------------
        # [계산] 상단 5대 주요 지표
        # ---------------------------------------------------------
        # 1. 총 수강 건수: 통계표의 전체 수강생 합 (학생수 * 과목수)
        #stat_total_enrollments = final_stats['전체수강생(중복자제거후)'].sum()
        stat_total_enrollments = len(df)
        
        # 2. 중복을 제거한 이수자 건수: 순수 학번(ID)의 개수
        stat_unique_students = len(df_dedup)
        
        # 3. 중복을 제거한 1학년 이수자 건수: 순수 1학년 학번(ID)의 개수
        # (전체 데이터 중 1학년인 행들의 학번 유니크 카운트)
        stat_unique_freshmen = len(df_freshman)
        
        # 4. 분석된 과목 수
        stat_subject_count = len(final_stats)
        
        # 5. 총 개설 분반 수
        stat_total_sections = final_stats['개설분반수'].sum()

        # ---------------------------------------------------------
        # [가공] 합계 행 추가 (테이블 표시용)
        # ---------------------------------------------------------
        # 1. 인덱스(교과목명)를 컬럼으로 변환하여 데이터프레임화
        final_stats_display = final_stats.reset_index().rename(columns={'index': col_subject})
        if col_subject not in final_stats_display.columns: 
            final_stats_display.rename(columns={final_stats_display.columns[0]: col_subject}, inplace=True)

        # 2. 합계 행 계산 (숫자형 컬럼만 더함)
        sum_row = final_stats.sum(numeric_only=True).to_frame().T
        sum_row[col_subject] = '합계' # 교과목명 컬럼에 '합계'라고 표시
        
        # 3. 원본 데이터 맨 아래에 합계 행 붙이기
        final_stats_with_sum = pd.concat([final_stats_display, sum_row], ignore_index=True)
        
        # 4. 컬럼 순서 재배치
        cols = [col_subject, '개설분반수', '전체수강생(중복자제거후)', '일학년수강생(중복자제거후)']
        final_stats_with_sum = final_stats_with_sum[cols]

        # 5. 인덱스 정리 (1, 2, 3... 하고 마지막 합계는 빈칸으로)
        new_index = list(range(1, len(final_stats_with_sum))) + ['']
        final_stats_with_sum.index = new_index

        # ---------------------------------------------------------
        # [출력] 화면 구성
        # ---------------------------------------------------------
        st.divider()
        st.subheader("1. 분석 요약 (Metrics)")

        # 5개 지표 출력
        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("총 수강건수(학생)", f"{stat_total_enrollments}건")
        m2.metric("실 이수자(중복제거)", f"{stat_unique_students}명")
        m3.metric("1학년 실 이수자", f"{stat_unique_freshmen}명")
        m4.metric("분석된 과목수", f"{stat_subject_count}개")
        m5.metric("총 개설분반", f"{stat_total_sections}개")

        st.divider()
        st.subheader("2. 과목별 상세 분석 결과 (합계 포함)")
        
        # 합계가 포함된 테이블 출력
        st.dataframe(final_stats_with_sum, use_container_width=True)

        # ---------------------------------------------------------
        # [출력] 엑셀 다운로드 (4개 시트)
        # ---------------------------------------------------------
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            # 시트 1: 전체 이수자 명단 (정렬됨)
            df_dedup.to_excel(writer, index=True, sheet_name='전체이수자')
            
            # 시트 2: 1학년 이수자 명단 (NEW)
            df_freshman.to_excel(writer, index=True, sheet_name='1학년이수자')
            
            # 시트 3: 개설 분반 리스트 (NEW)
            section_list_df.to_excel(writer, index=False, sheet_name='개설분반리스트')

            # 시트 4: 통계 분석 (합계 포함)
            final_stats_with_sum.to_excel(writer, index=False, sheet_name='통계분석')
        
        st.download_button(
            label="결과 엑셀 다운로드 (시트 4개 포함)",
            data=output.getvalue(),
            file_name="SW기초교과목_이수자_분석결과_최종.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        # 참고용 미리보기 (접을 수 있음)
        with st.expander("데이터 미리보기 (전체 이수자 명단)"):
             st.dataframe(df_dedup)

    except Exception as e:
        st.error(f"오류가 발생했습니다: {e}")
        st.warning("데이터 파일의 컬럼명('학번', '학년', '교과목명', '분반', '학기')을 확인해주세요.")

else:
    st.info("CSV 파일을 업로드하면 자동으로 분석이 시작됩니다.(파일 비밀번호 제거) ")



