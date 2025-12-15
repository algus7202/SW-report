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
st.markdown("엑셀(CSV) 파일을 업로드하면 정렬 및 분석 결과를 보여줍니다. (데이터 병합 오류 수정 버전)")

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
        
        # Categorical 변환 (순서 적용)
        df[col_subject] = pd.Categorical(df[col_subject], categories=final_order, ordered=True)
        
        # 전체 데이터 정렬
        df_sorted = df.sort_values(by=[col_semester, col_grade, col_subject], ascending=[True, True, True])

        # --- 분석 1: 분반 수 계산 (전체 데이터 기준) ---
        # [과목, 학기, 분반] 3가지 키의 고유 조합 추출
        unique_sections = df_sorted[[col_subject, col_semester, col_class]].drop_duplicates()
        
        # 과목별 그룹화 -> 사이즈 계산 -> [중요] DataFrame으로 즉시 변환
        class_counts_df = unique_sections.groupby(col_subject, observed=True).size().to_frame(name='개설분반수')

        # --- 분석 2: 학생 수 계산 (학생 중복 제거 기준) ---
        # [학번, 과목] 기준으로 중복 제거
        df_dedup = df_sorted.drop_duplicates(subset=[col_id, col_subject])
        
        # 인덱스 재설정 (1번부터)
        df_dedup = df_dedup.reset_index(drop=True)
        df_dedup.index = df_dedup.index + 1
        
        # 화면 출력 1
        st.subheader("1. 정렬 및 중복 제거 완료 데이터")
        st.dataframe(df_dedup, use_container_width=True)

        # 학생 수 집계
        student_counts_df = df_dedup.groupby(col_subject, observed=True).agg(
            전체수강생=(col_id, 'count'),
            일학년수강생=(col_grade, lambda x: (x == 1).sum())
        )

        # --- [수정된 병합 로직] ---
        # concat 대신 merge 사용 (인덱스 기준, outer join)
        # 이렇게 하면 어느 한쪽에만 있는 과목도 누락되지 않고 안전하게 합쳐짐
        final_stats = pd.merge(
            class_counts_df, 
            student_counts_df, 
            left_index=True, 
            right_index=True, 
            how='outer'
        )

        # 결측치(NaN)가 있다면 0으로 채우고 정수형으로 변환 (깔끔한 출력을 위해)
        final_stats = final_stats.fillna(0).astype(int)

        # 컬럼 순서 지정
        final_stats = final_stats[['개설분반수', '전체수강생', '일학년수강생']]
        
        # 통계표 인덱스 재설정 (1번부터 표시)
        final_stats = final_stats.reset_index()
        final_stats.index = final_stats.index + 1

        # --- 결과 출력 ---
        st.divider()
        st.subheader("2. 과목별 상세 분석 결과")

        # 메트릭 표시
        total_students = len(df_dedup)
        total_classes = final_stats['개설분반수'].sum()
        
        c1, c2, c3 = st.columns(3)
        c1.metric("총 수강 건수 (학생)", f"{total_students}명")
        c2.metric("총 개설 분반 (학기포함)", f"{total_classes}개")
        c3.metric("분석된 과목 수", f"{len(final_stats)}개")

        st.dataframe(final_stats, use_container_width=True)

        # 엑셀 다운로드
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_dedup.to_excel(writer, index=True, sheet_name='정렬된데이터')
            final_stats.to_excel(writer, index=False, sheet_name='통계분석')
        
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
