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
st.markdown("엑셀(CSV) 파일을 업로드하면 정렬 및 분석 결과를 보여줍니다. (개설 분반 리스트 시트 추가)")

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
        
        # Categorical 변환 (정렬 순서 적용)
        df[col_subject] = pd.Categorical(df[col_subject], categories=final_order, ordered=True)
        
        # 전체 데이터 정렬 (학기 -> 학년 -> 과목 순)
        df_sorted = df.sort_values(by=[col_semester, col_grade, col_subject], ascending=[True, True, True])

        # --- 분석 1: 분반 정보 추출 ---
        # [과목, 학기, 분반] 고유 조합 추출 (순수 개설 강좌 목록)
        unique_sections = df_sorted[[col_subject, col_semester, col_class]].drop_duplicates()
        
        # [NEW] 엑셀 저장용 분반 리스트 정렬 (과목 -> 학기 -> 분반 순서로 보기 좋게)
        section_list_df = unique_sections.sort_values(by=[col_subject, col_semester, col_class])
        
        # 분반 수 카운트 (통계용)
        class_counts_df = unique_sections.groupby(col_subject, observed=True).size().to_frame(name='개설분반수')

        # --- 분석 2: 학생 수 계산 ---
        # [학번, 과목] 기준으로 중복 제거
        df_dedup = df_sorted.drop_duplicates(subset=[col_id, col_subject])
        
        # 인덱스 재설정
        df_dedup = df_dedup.reset_index(drop=True)
        df_dedup.index = df_dedup.index + 1
        
        # 1학년 데이터 별도 추출
        df_freshman = df_dedup[df_dedup[col_grade] == 1].copy()
        df_freshman = df_freshman.reset_index(drop=True)
        df_freshman.index = df_freshman.index + 1

        # 정제된 명단 화면 출력
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
        final_stats.columns = ['개설분반수', '전체수강생(중복후제거)', '일학년수강생(중복후제거)']

        # --- 상단 5대 주요 지표 계산 ---
        stat_total_enrollments = final_stats['전체수강생(중복후제거)'].sum()
        stat_unique_students = df_dedup[col_id].nunique()
        stat_unique_freshmen = df_sorted[df_sorted[col_grade] == 1][col_id].nunique()
        stat_subject_count = len(final_stats)
        stat_total_sections = final_stats['개설분반수'].sum()

        # --- 합계 행 생성 ---
        final_stats_display = final_stats.reset_index().rename(columns={'index': col_subject})
        if col_subject not in final_stats_display.columns: 
            final_stats_display.rename(columns={final_stats_display.columns[0]: col_subject}, inplace=True)

        sum_row = final_stats.sum(numeric_only=True).to_frame().T
        sum_row[col_subject] = '합계'
        
        final_stats_with_sum = pd.concat([final_stats_display, sum_row], ignore_index=True)
        cols = [col_subject, '개설분반수', '전체수강생(중복후제거)', '일학년수강생(중복후제거)']
        final_stats_with_sum = final_stats_with_sum[cols]

        new_index = list(range(1, len(final_stats_with_sum))) + ['']
        final_stats_with_sum.index = new_index

        # --- 결과 화면 출력 ---
        st.divider()
        st.subheader("2. 과목별 상세 분석 결과")

        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("총 수강건수(학생)", f"{stat_total_enrollments}건")
        m2.metric("실 이수자(중복제거)", f"{stat_unique_students}명")
        m3.metric("1학년 실 이수자(중복제거)", f"{stat_unique_freshmen}명")
        m4.metric("분석된 과목수", f"{stat_subject_count}개")
        m5.metric("총 개설분반", f"{stat_total_sections}개")

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
