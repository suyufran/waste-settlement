import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io
import plotly.express as px

# 1. 페이지 기본 설정 및 스타일 주입
st.set_page_config(page_title="종량제 배부 정산 시스템", layout="centered")

st.markdown("""
    <style>
    .stApp { background-color: #F8FAFC; }
    h1 {
        color: #0F172A !important;
        font-family: 'Malgun Gothic', sans-serif;
        font-weight: 700 !important;
        border-bottom: 2px solid #1E3A8A;
        padding-bottom: 10px;
    }
    h3 { color: #1E3A8A !important; font-weight: 600 !important; margin-top: 20px; }
    .main-header {
        background: linear-gradient(135deg, #1E3A8A 0%, #0F172A 100%);
        padding: 22px; border-radius: 10px; color: white; margin-bottom: 20px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
    }
    .step-container {
        display: flex; justify-content: space-between; background-color: #FFFFFF;
        padding: 12px 20px; border-radius: 8px; border: 1px solid #E2E8F0;
        margin-bottom: 25px; box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }
    .step-item { font-size: 14px; font-weight: 600; color: #94A3B8; }
    .step-item.active { color: #1E3A8A; border-bottom: 2px solid #1E3A8A; padding-bottom: 2px; }
    .step-item.complete { color: #10B981; }
    
    .stButton>button {
        background-color: #1E3A8A !important; color: white !important;
        border-radius: 6px !important; border: none !important;
        padding: 10px 24px !important; font-weight: bold !important;
        transition: all 0.2s; width: 100%;
    }
    .stButton>button:hover {
        background-color: #1D4ED8 !important;
        box-shadow: 0 4px 12px rgba(29, 78, 216, 0.3) !important;
    }
    .warning-box {
        background-color: #FEF2F2; border-left: 4px solid #EF4444;
        padding: 15px; border-radius: 4px; color: #991B1B; margin-bottom: 15px;
        font-size: 14.5px; line-height: 1.6;
    }
    .row-grid-box {
        background-color: #FFFFFF; border: 1px solid #E2E8F0;
        border-radius: 6px; padding: 12px; margin-top: 10px;
    }
    </style>
    """, unsafe_allow_html=True)

st.markdown("""
    <div class="main-header">
        <h2 style='margin:0; font-size:24px; font-weight:700;'>🏢 쓰레기 종량제 배부 정산 시스템 v1.25</h2>
        <p style='margin:5px 0 0 0; opacity:0.8; font-size:14px;'>세대별 사용량을 기준으로 계량 오차를 비례 배부하여 정산 엑셀을 생성합니다.</p>
    </div>
    """, unsafe_allow_html=True)

is_file_uploaded = False
has_bad_rows = False
step1_class = "active"
step2_class = ""
step3_class = ""

# 2. 데이터 입력 섹션
st.subheader("📋 1. 정산 데이터 입력 및 행 선택")
col_input1, col_input2 = st.columns(2)

with col_input1:
    uploaded_file = st.file_uploader("세대별 최초 사용량 엑셀 파일 (.xlsx)", type=["xlsx"])
with col_input2:
    total_weight = st.number_input("전체 계량 종량 (측정치, kg)", min_value=0.0, value=500.0, step=10.0)

header_row = 1
if uploaded_file is not None:
    is_file_uploaded = True
    step1_class = "complete"
    step2_class = "active"
    
    st.write("📌 **[시각적 행 선택기]**")
    try:
        preview_raw = pd.read_excel(uploaded_file, header=None, nrows=7)
        preview_raw = preview_raw.iloc[:, :6]
        preview_raw.columns = [f"{i+1}열" for i in range(preview_raw.shape[1])]
        preview_raw.index = [f"{i+1}행" for i in range(preview_raw.shape[0])]
        
        st.markdown('<div class="row-grid-box">', unsafe_allow_html=True)
        st.dataframe(preview_raw, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)
        
        row_options = [i for i in range(1, len(preview_raw) + 1)]
        header_row = st.radio(
            "실제 데이터가 시작되는 행은 몇 번째인가요? 선택해 주세요:",
            options=row_options, 
            format_func=lambda x: f"{x}행",
            index=0, horizontal=True
        )
    except Exception as e:
        st.warning(f"행 선택기 초기화 중 일시적 에러가 발생하여 기본값(1행)을 사용합니다: {e}")

st.markdown(f"""
    <div class="step-container">
        <div class="step-item {step1_class}">📁 Step 1. 파일 업로드 및 열 지정</div>
        <div class="step-item {step2_class}">⚠️ Step 2. 데이터 검증 및 수정</div>
        <div class="step-item {step3_class}">🚀 Step 3. 대시보드 분석 및 정산</div>
    </div>
    """, unsafe_allow_html=True)

if uploaded_file is not None:
    try:
        file_key = f"df_{uploaded_file.name}_{header_row}"
        if 'file_key' not in st.session_state or st.session_state.file_key != file_key or st.sidebar.button("🔄 원본 파일 다시 읽기"):
            st.session_state.file_key = file_key
            base_df = pd.read_excel(uploaded_file, header=header_row - 1)
            base_df = base_df.reset_index(drop=True)
            base_df['__fixed_row_id'] = base_df.index
            st.session_state.edited_df = base_df
        
        df = st.session_state.edited_df.copy()
        
        raw_columns = [c for c in df.columns if c != '__fixed_row_id']
        dropdown_options = []
        display_columns_map = {}
        
        for idx, col in enumerate(raw_columns):
            col_str = str(col).strip()
            if "Unnamed:" in col_str or col_str == "" or pd.isna(col):
                friendly_name = f"{idx + 1}열 (내용 없음)"
            else:
                friendly_name = f"{idx + 1}열 ({col_str})"
            
            display_columns_map[friendly_name] = col
            dropdown_options.append(friendly_name)
        
        st.write("")
        st.markdown("<h3 style='font-size:18px; margin-bottom:5px;'>🔍 2. 각 항목에 해당되는 열을 선택해 주세요.</h3>", unsafe_allow_html=True)
        
        col1, col2, col3, col4, col5 = st.columns(5)
        with col1: dong_sel = st.selectbox("📍 '동'", dropdown_options, index=0)
        with col2: ho_sel = st.selectbox("🚪 '호수'", dropdown_options, index=1 if len(dropdown_options)>1 else 0)
        with col3: usage_sel = st.selectbox("⚖️ '최초사용량'", dropdown_options, index=2 if len(dropdown_options)>2 else 0)
        with col4: eco_home_sel = st.selectbox("🏡 '에코홈'", dropdown_options, index=3 if len(dropdown_options)>3 else 0)
        with col5: eco_park_sel = st.selectbox("🌳 '에코파크'", dropdown_options, index=4 if len(dropdown_options)>4 else 0)
        
        dong_col = display_columns_map[dong_sel]
        ho_col = display_columns_map[ho_sel]
        usage_col = display_columns_map[usage_sel]
        eco_home_col = display_columns_map[eco_home_sel]
        eco_park_col = display_columns_map[eco_park_sel]
            
        for c in [usage_col, eco_home_col, eco_park_col]:
            df[c] = df[c].astype(str).str.replace(',', '').str.strip()
            df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0.0)

        df['_max_allowed'] = (df[eco_home_col] * 1.5) + (df[eco_park_col] * 5.0)
        df['검증결과'] = "정상"
        
        is_bad = (df[usage_col] > df['_max_allowed']) | ((df[eco_home_col] == 0) & (df[eco_park_col] == 0) & (df[usage_col] > 0))
        df.loc[is_bad, '검증결과'] = "비정상 (최대 허용량 초과)"

        def get_sort_keys(value):
            if pd.isna(value): return (1, 0)
            val_str = str(value).strip()
            cleaned = ''.join(filter(str.isdigit, val_str))
            if cleaned: return (0, int(cleaned))
            else: return (1, val_str)

        df['_sort_dong_type'], df['_sort_dong_val'] = zip(*df[dong_col].apply(get_sort_keys))
        df['_sort_ho_type'], df['_sort_ho_val'] = zip(*df[ho_col].apply(get_sort_keys))
        df = df.sort_values(by=['_sort_dong_type', '_sort_dong_val', '_sort_ho_type', '_sort_ho_val']).reset_index(drop=True)

        st.write("---")
        st.subheader("⚠️ 3. 비정상 사용량 세대 및 데이터 교정")
        
        bad_rows = df[df['검증결과'] != "정상"]
        display_cols = ['__fixed_row_id', dong_col, ho_col, usage_col, eco_home_col, eco_park_col, '검증결과']
        
        if len(bad_rows) > 0:
            has_bad_rows = True
            st.markdown(f"""
            <div class="warning-box">
                <strong>🚨 검증 경고 ({len(bad_rows)}건 발견)</strong><br>
                기기 입력 오류나 허용 한도를 초과한 세대가 발견되었습니다. 아래 테이블에서 <strong>최초사용량</strong> 수치를 수정한 후 반드시 <strong>💾 수정사항 저장 및 데이터 동기화</strong> 버튼을 눌러주세요.
            </div>
            """, unsafe_allow_html=True)
            
            cell_styles = {
                "__fixed_row_id": st.column_config.NumberColumn("고유번호", disabled=True, width="small"),
                usage_col: st.column_config.NumberColumn(
                    f"✏️ {usage_col} (교정 필수)",
                    help="수정할 수치를 더블클릭하여 입력하세요.",
                    required=True, format="%.2f"
                )
            }
            
            edited_bad_df = st.data_editor(
                bad_rows[display_cols].copy(),
                column_config=cell_styles,
                use_container_width=True, num_rows="fixed",
                disabled=[dong_col, ho_col, eco_home_col, eco_park_col, '검증결과'], key="bad_df_editor"
            )
            
            if st.button("💾 수정사항 저장 및 데이터 동기화"):
                for row_idx in range(len(edited_bad_df)):
                    target_id = int(edited_bad_df.iloc[row_idx]['__fixed_row_id'])
                    raw_val = str(edited_bad_df.iloc[row_idx][usage_col]).replace(',', '').strip()
                    try:
                        final_val = float(raw_val)
                    except ValueError:
                        final_val = 0.0
                    
                    st.session_state.edited_df.loc[st.session_state.edited_df['__fixed_row_id'] == target_id, usage_col] = final_val
                
                st.success("🎉 데이터 동기화가 안전하게 완료되었습니다! 상단 및 엑셀 수치에 정상 반영됩니다.")
                st.rerun()
        else:
            step2_class = "complete"
            step3_class = "active"
            st.success("✅ 현재 모든 세대의 중량이 기기별 허용 범위 내에 존재합니다. (정상)")
            with st.expander("🔍 정렬된 전체 세대 데이터 확인"):
                st.dataframe(df[[c for c in display_cols if c != '__fixed_row_id']], use_container_width=True)

        st.write("---")
        st.subheader("📊 4. 동별 배출량 분석 대시보드")
        
        df_chart = df.copy()
        df_chart['차트용_동'] = df_chart[dong_col].astype(str).str.replace('.0', '', regex=False).str.strip()
        df_chart[usage_col] = pd.to_numeric(df_chart[usage_col], errors='coerce').fillna(0.0)
        
        dong_summary = df_chart.groupby('차트용_동')[usage_col].sum().reset_index()
        dong_summary['_sort_type'], dong_summary['_sort_val'] = zip(*dong_summary['차트용_동'].apply(get_sort_keys))
        dong_summary = dong_summary.sort_values(by=['_sort_type', '_sort_val']).reset_index(drop=True)
        
        fig_bar = px.bar(
            dong_summary, x='차트용_동', y=usage_col,
            labels={'차트용_동': '동 (Dong)', usage_col: '최초 배출 총합 (kg)'},
            title="🏢 동별 최초 사용량 총합 통계 (kg)", color_discrete_sequence=['#1E3A8A']
        )
        fig_bar.update_layout(margin=dict(l=20, r=20, t=40, b=20), plot_bgcolor='rgba(0,0,0,0)')
        st.plotly_chart(fig_bar, use_container_width=True)
            
        st.write("---")
        st.subheader("🚀 5. 비례 배부 정산 및 결과 출력")
        
        bypass_button = False
        if has_bad_rows:
            st.markdown("""
            <div style="background-color: #FFF9E6; border-left: 4px solid #D97706; padding: 15px; color: #92400E; font-size: 14px; margin-bottom: 12px; border-radius:4px;">
                ⚠️ <strong>알림:</strong> 현재 미수정되거나 허용치를 초과한 비정상 데이터 세대가 포함되어 있습니다.<br>
                정상적인 정산 처리를 위해 데이터를 교정하는 것을 추천하지만, <strong>실무상 무시하고 그대로 정산을 밀고 나가야 한다면</strong> 아래 확인 상자에 체크해 주세요.
            </div>
            """, unsafe_allow_html=True)
            bypass_approval = st.checkbox("⚠️ 예, 일부 비정상 세대가 존재하더라도 수치를 무시하고 정산을 진행하겠습니다.")
            if bypass_approval:
                bypass_button = True
        else:
            bypass_button = True
            
        if st.button("정산 실행 및 엑셀 파일 생성", disabled=not bypass_button):
            step3_class = "complete"
            wb = openpyxl.Workbook()
            ws_summary = wb.active
            ws_summary.title = "정산 요약"
            ws_data = wb.create_sheet(title="세대별 상세 데이터")
            
            font_title = Font(name="Malgun Gothic", size=16, bold=True, color="1B365D")
            font_header = Font(name="Malgun Gothic", size=11, bold=True, color="FFFFFF")
            font_bold = Font(name="Malgun Gothic", size=10, bold=True, color="000000")
            font_regular = Font(name="Malgun Gothic", size=10, color="333333")
            fill_header = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
            fill_summary_hdr = PatternFill(start_color="334155", end_color="334155", fill_type="solid")
            fill_zebra = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
            fill_total = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
            thin_border = Border(left=Side(style="thin", color="CBD5E1"), right=Side(style="thin", color="CBD5E1"), top=Side(style="thin", color="CBD5E1"), bottom=Side(style="thin", color="CBD5E1"))
            total_border = Border(top=Side(style="thin", color="000000"), bottom=Side(style="double", color="000000"))
            
            headers_data = ["동", "호수", "최초 사용량 (kg)", "분할 배부량 (kg)", "최종 조정 사용량 (kg)"]
            for col_num, header in enumerate(headers_data, 1):
                cell = ws_data.cell(row=3, column=col_num, value=header)
                cell.font = font_header; cell.fill = fill_header; cell.alignment = Alignment(horizontal="center", vertical="center")
            
            start_row = 4
            total_rows = len(df)
            sum_row = start_row + total_rows
            
            for i, row in df.iterrows():
                current_row = start_row + i
                ws_data.cell(row=current_row, column=1, value=str(row[dong_col]).replace('.0', '')).alignment = Alignment(horizontal="center")
                ws_data.cell(row=current_row, column=2, value=str(row[ho_col]).replace('.0', '')).alignment = Alignment(horizontal="center")
                
                c_orig = ws_data.cell(row=current_row, column=3, value=float(row[usage_col]))
                c_orig.number_format = '#,##0.00'
                
                c_dist = ws_data.cell(row=current_row, column=4, value=f"='정산 요약'!$B$4*(C{current_row}/C${sum_row})")
                c_dist.number_format = '#,##0.00'
                c_final = ws_data.cell(row=current_row, column=5, value=f"=C{current_row}+D{current_row}")
                c_final.number_format = '#,##0.00'
                
                for c in range(1, 6):
                    cell = ws_data.cell(row=current_row, column=c)
                    cell.font = font_regular; cell.border = thin_border
                    if c >= 3: cell.alignment = Alignment(horizontal="right")
                    if i % 2 == 1: cell.fill = fill_zebra
            
            ws_data.cell(row=sum_row, column=1, value="합계").alignment = Alignment(horizontal="center")
            ws_data.cell(row=sum_row, column=3, value=f"=SUM(C{start_row}:C{sum_row-1})").number_format = '#,##0.00'
            ws_data.cell(row=sum_row, column=4, value=f"=SUM(D{start_row}:D{sum_row-1})").number_format = '#,##0.00'
            ws_data.cell(row=sum_row, column=5, value=f"=SUM(E{start_row}:E{sum_row-1})").number_format = '#,##0.00'
            for c in range(1, 6):
                cell = ws_data.cell(row=sum_row, column=c)
                cell.font = font_bold; cell.fill = fill_total; cell.border = total_border
                if c >= 3: cell.alignment = Alignment(horizontal="right")
                
            ws_summary.cell(row=2, column=1, value="쓰레기 종량제 배부 정산 요약").font = font_title
            ws_summary.cell(row=3, column=1, value="항목").fill = fill_summary_hdr
            ws_summary.cell(row=3, column=2, value="수량 (kg)").fill = fill_summary_hdr
            for c in [1, 2]:
                ws_summary.cell(row=3, column=c).font = font_header; ws_summary.cell(row=3, column=c).alignment = Alignment(horizontal="center"); ws_summary.cell(row=3, column=c).border = thin_border
                
            summary_items = [
                ("전체 계량 종량 (측정치)", float(total_weight)),
                ("세대별 최초 사용량 합계", f"='세대별 상세 데이터'!C{sum_row}"),
                ("정산 대상 차이 (오차)", "=B4-B5"),
                ("배부 후 최종 확인 합계", f"='세대별 상세 데이터'!E{sum_row}")
            ]
            for idx, (label, val) in enumerate(summary_items, 4):
                c1 = ws_summary.cell(row=idx, column=1, value=label)
                c2 = ws_summary.cell(row=idx, column=2, value=val)
                
                is_bold = "합계" in label or "차이" in label
                
                c1.font = font_bold if is_bold else font_regular
                c2.font = font_bold if is_bold else font_regular
                c2.number_format = '#,##0.00'
                c1.alignment = Alignment(horizontal="left")
                c2.alignment = Alignment(horizontal="right")
                c1.border = thin_border
                c2.border = thin_border
                
                if "차이" in label:
                    c1.fill = fill_total
                    c2.fill = fill_total
                
            ws_summary.column_dimensions['A'].width = 30
            ws_summary.column_dimensions['B'].width = 18
            for col in ws_data.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                col_letter = get_column_letter(col[0].column)
                ws_data.column_dimensions[col_letter].width = max(max_len * 1.6, 15)
                
            excel_data = io.BytesIO()
            wb.save(excel_data)
            excel_data.seek(0)
            
            st.download_button(
                label="📁 정산 완료 엑셀 다운로드",
                data=excel_data,
                file_name="종량제_배부_정산_완료.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    except Exception as e:
        st.error(f"파일을 처리하는 중 오류가 발생했습니다: {e}")