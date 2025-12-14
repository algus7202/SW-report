# -*- coding: utf-8 -*-
"""
Created on Sun Dec 14 21:26:07 2025

@author: algus
"""

import streamlit as st
import pandas as pd
import io

# 페이지 설정
st.set_page_config(page_title="수강생 분석 도구", layout="wide")

st.title("📊 수강생 정렬 및 분석 시스템")
st.markdown("엑셀 파일을 업로드하면 정렬 규칙에 따라 데이터를 정리하고 분석 결과를 보여줍니다.")

# 1. 파일 업로드
uploaded_file = st.file_uploader("엑셀 파일을 업로드하세요 (.xlsx)", type=['xlsx', 'xls'])

if uploaded_file is not None:
    try:
        # 데이터 로드
        df = pd.read_excel(uploaded_file)
        
        # --- 필수 컬럼 확인 및 매핑 (사용자 파일에 맞게 수정 필요) ---
        # 예시: 사용자의 엑셀 컬럼명이 다를 경우를 대비해 변수로 관리
        col_id = '학번'
        col_grade = '학년'
        col_subject = '수강과목'
        
        # 필수 컬럼이 있는지 확인
        required_cols = [col_id, col_grade, col_subject]
        if not all(col in df.columns for col in required_cols):
            st.error(f"엑셀 파일에 다음 컬럼이 반드시 포함되어야 합니다: {required_cols}")
            st.stop()

        # --- 데이터 전처리 ---
        # 학년 데이터 숫자형 변환 (예: "1학년" -> 1) 시도, 이미 숫자면 패스
        df[col_grade] = df[col_grade].astype(str).str.extract(r'(\d+)').astype(int)

        # 2. 정렬 로직 구현
        # 1순위: 학년 (1 -> 4)
        # 2순위: 수강과목 (컴퓨팅사고와 인공지능 -> 기초컴퓨터프로그래밍)
        
        # 과목 정렬 순서 지정 (Categorical 타입 활용)
        custom_order = ["컴퓨팅사고와 인공지능", "기초컴퓨터프로그래밍"]
        
        # 지정된 과목 외의 과목이 있을 경우를 대비해 순서 목록 확장
        existing_subjects = df[col_subject].unique()
        remaining_subjects = [x for x in existing_subjects if x not in custom_order]
        final_order = custom_order + remaining_subjects
        
        # Categorical 데이터로 변환하여 정렬 순서 강제
        df[col_subject] = pd.Categorical(df[col_subject], categories=final_order, ordered=True)
        
        # 정렬 실행
        df_sorted = df.sort_values(by=[col_grade, col_subject], ascending=[True, True])

        # 3. 중복 제거
        # 기준: 학번과 수강과목이 같으면 중복으로 간주 (동일 학생이 동일 과목 중복 수강 신청된 경우)
        # 만약 과목 상관없이 학생 자체의 중복을 제거하려면 subset=[col_id] 로 변경
        df_dedup = df_sorted.drop_duplicates(subset=[col_id, col_subject], keep='first')
        
        # 결과 보여주기
        st.subheader("1. 정렬 및 중복 제거 완료 데이터")
        st.dataframe(df_dedup, use_container_width=True)

        # --- 분석 로직 ---
        st.divider()
        st.subheader("2. 수강생 분석 결과")

        # 분석 1: 전체 수강자 수 (중복 제거된 데이터 기준 행의 개수)
        total_students = len(df_dedup)
        # 분석 2: 1학년 수강자 수
        freshman_students = len(df_dedup[df_dedup[col_grade] == 1])

        # 메트릭 표시
        c1, c2 = st.columns(2)
        c1.metric("총 수강 건수", f"{total_students}명")
        c2.metric("1학년 수강 건수", f"{freshman_students}명")

        # 분석 3 & 4: 각 과목별 전체 수강자 수 및 1학년 수강자 수
        st.markdown("##### 과목별 상세 현황")
        
        # 과목별 그룹화
        subject_stats = df_dedup.groupby(col_subject, observed=True).agg(
            전체수강생=(col_id, 'count'),
            일학년수강생=(col_grade, lambda x: (x == 1).sum())
        ).reset_index()

        st.dataframe(subject_stats, use_container_width=True)

        # (선택) 엑셀 다운로드 버튼
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_dedup.to_excel(writer, index=False, sheet_name='정렬된데이터')
            subject_stats.to_excel(writer, index=False, sheet_name='통계분석')
        
        st.download_button(
            label="결과 엑셀 다운로드",
            data=output.getvalue(),
            file_name="수강생분석결과.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"오류가 발생했습니다: {e}")
        st.warning("엑셀 파일의 컬럼명('학번', '학년', '수강과목')과 데이터 형식을 확인해주세요.")

else:
    st.info("파일을 업로드하면 분석이 시작됩니다.")