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
st.markdown("엑셀(CSV) 파일을 업로드하면 정렬 및 분석 결과를 보여줍니다. (중복 제거 통계 및 합계 행 포함)")

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
        # 학년 데이터 숫자형 변환
        df[col_grade] = df[col_grade].astype(str).str.extract(r'(\d+)', expand=False).fillna(0).astype(int)

        # 2. 정렬 로직
        custom_order = ['컴퓨팅사고와인공지능', '기초컴퓨터프로그래밍','IT환경에서의개인정보보호','멀티미디어의이해와활용','디지털리터러시의 이해와 활용','컴퓨터 시뮬레이션', '컴퓨터프로그래밍입문']
        
        existing_subjects = df[col_subject].unique()
        remaining_subjects = [x for x in existing_subjects if x not in custom_order]
        final_order = custom_order + remaining_subjects
        
        # Categorical 변환
        df[col_subject] = pd.Categorical(df[col_subject], categories=final_order, ordered=True)
        
        # 전체 데이터 정렬
        df_sorted = df.sort_values(by=[col_grade, col_subject, col_semester], ascending=[True, True, True])

        # --- 분석 1: 분반 수 계산 (전체 데이터 기준) ---
        # [과목, 학기, 분반] 고유 조합 추출
        unique_sections = df_sorted[[col_subject, col_semester, col_class]].drop_duplicates()
        class_counts_df = unique_sections.groupby(col_subject, observed=True).size().to_frame(name='개설분반수')

        # --- 분석 2: 학생 수 계산 (학생 중복 제거 기준) ---
        # [학번] 기준으로 중복 제거 (한 과목을 여러번 들어도 1명으로 카운트)
        df_dedup = df_sorted.drop_duplicates(subset=[col_id])
        
        # 인덱스 재설정
        df_dedup = df_dedup.reset_index(drop=True)
        df_dedup.index = df_dedup.index + 1
        
        # 정제된 명단 출력
        st.subheader("1. 정렬 및 중복 제거 완료 데이터")
        st.dataframe(df_dedup, use_container_width=True)

        # 학생 수 집계
        student_counts_df = df_dedup.groupby(col_subject, observed=True).agg(
            전체수강생=(col_id, 'count'),
            일학년수강생=(col_grade, lambda x: (x == 1).sum())
        )

        # --- 결과 병합 ---
        final_stats = pd.merge(
            class_counts_df, 
            student_counts_df, 
            left_index=True, 
            right_index=True, 
            how='outer'
        )
        final_stats = final_stats.fillna(0).astype(int)

        # 컬럼명 변경 (요청하신 대로)
        final_stats.columns = ['개설분반수', '전체수강생(중복후제거)', '일학년수강생(중복후제거)']

        # --- [상단 5대 주요 지표 계산] ---
        
        # 1. 총 수강 건수 (각 과목별 수강생 합계)
        stat_total_enrollments = final_stats['전체수강생(중복후제거)'].sum()
        
        # 2. 중복 제거한 이수자 건수 (전체 데이터에서 고유 학번 수)
        # df_dedup은 [학번, 과목] 유니크이므로, 순수 학번 유니크는 df_dedup[col_id].nunique()로 계산
        stat_unique_students = df_dedup[col_id].nunique()
        
        # 3. 중복 제거한 1학년 이수자 건수 (전체 데이터에서 1학년인 고유 학번 수)
        # 전체 이력 중 1학년인 행만 뽑아서 학번 중복 제거
        stat_unique_freshmen = df_sorted[df_sorted[col_grade] == 1][col_id].nunique()
        
        # 4. 분석된 과목 수
        stat_subject_count = len(final_stats)
        
        # 5. 총 개설 분반 수
        stat_total_sections = final_stats['개설분반수'].sum()

        # --- [합계 행 생성 로직] ---
        # 1. 인덱스(교과목명)를 컬럼으로 변환
        final_stats_display = final_stats.reset_index().rename(columns={'index': col_subject})
        if col_subject not in final_stats_display.columns: 
            final_stats_display.rename(columns={final_stats_display.columns[0]: col_subject}, inplace=True)

        # 2. 합계 행 계산
        sum_row = final_stats.sum(numeric_only=True).to_frame().T
        sum_row[col_subject] = '합계' # 교과목명 자리에 '합계' 입력
        
        # 3. 합치기
        final_stats_with_sum = pd.concat([final_stats_display, sum_row], ignore_index=True)
        
        # 4. 컬럼 순서 재배치
        cols = [col_subject, '개설분반수', '전체수강생(중복후제거)', '일학년수강생(중복후제거)']
        final_stats_with_sum = final_stats_with_sum[cols]

        # 5. 인덱스 정리 (1번부터, 합계는 빈칸)
        new_index = list(range(1, len(final_stats_with_sum))) + ['']
        final_stats_with_sum.index = new_index

        # --- 결과 화면 출력 ---
        st.divider()
        st.subheader("2. 과목별 상세 분석 결과")

        # 5개 지표 출력
        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("총 수강건수(학생)", f"{stat_total_enrollments}건")
        m2.metric("실 이수자(중복제거)", f"{stat_unique_students}명")
        m3.metric("1학년 실 이수자(중복제거)", f"{stat_unique_freshmen}명")
        m4.metric("분석된 과목수", f"{stat_subject_count}개")
        m5.metric("총 개설분반", f"{stat_total_sections}개")

        # 합계가 포함된 테이블 출력
        st.dataframe(final_stats_with_sum, use_container_width=True)

       # 엑셀 다운로드
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            # 시트 1: 전체 이수자 명단
            df_dedup.to_excel(writer, index=True, sheet_name='전체이수자(정렬됨)')
            
            # 시트 2: 1학년 이수자 명단
            df_freshman.to_excel(writer, index=True, sheet_name='1학년이수자')
            
            # 시트 3: [NEW] 개설 분반 리스트 (과목/학기/분반)
            section_list_df.to_excel(writer, index=False, sheet_name='개설분반리스트')

            # 시트 4: 통계 분석 (합계 포함)
            final_stats_with_sum.to_excel(writer, index=False, sheet_name='통계분석')
        
        st.download_button(
            label="결과 엑셀 다운로드",
            data=output.getvalue(),
            file_name="수강생분석결과_최종.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"오류가 발생했습니다: {e}")
        st.warning("데이터 파일의 컬럼명과 형식을 확인해주세요.")

else:
    st.info("파일을 업로드하면 분석이 시작됩니다.")

