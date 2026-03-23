import streamlit as st
import pandas as pd
import io
import os

# ==========================================
# 1. 网页界面设置
# ==========================================
st.set_page_config(page_title="EV 充电账单生成系统", page_icon="⚡", layout="centered")

st.title("⚡ EV 充电月度账单自动生成系统")
st.markdown("""
**小白专属操作指南：**
请将本月的 **4 个必需文件**（包括包含各个 Sheet 的 Excel 总表、GoParkin 明细以及两个 CRM 匹配表）一次性全选，拖拽到下方区域中。系统会自动嗅探文件、精准提取数据并生成报表。
""")

# 文件上传组件（支持同时拖入多个 CSV 和 Excel）
uploaded_files = st.file_uploader("📂 请拖入数据文件 (支持 .csv 和 .xlsx)", accept_multiple_files=True)

# 基础读取函数
def read_file(file_obj):
    filename = file_obj.name.lower()
    try:
        if filename.endswith('.csv'):
            return pd.read_csv(file_obj)
        elif filename.endswith(('.xlsx', '.xls')):
            return pd.read_excel(file_obj)
    except Exception as e:
        return None
    return None

# ==========================================
# 2. 核心计算逻辑
# ==========================================
if st.button("🚀 一键生成合并报表"):
    if len(uploaded_files) < 4:
        st.warning("⚠️ 似乎少传了文件！请确保上传了包含 CRM 和明细的至少 4 个文件。")
    else:
        with st.spinner("系统正在疯狂计算中，自动提取 Sheet 中..."):
            try:
                # --- A. 智能识别上传的文件 ---
                file_gp_tx = None
                file_gp_crm = None
                file_sp_tx = None
                file_sp_crm = None

                for file in uploaded_files:
                    if "charging_transactions_goparkin" in file.name:
                        file_gp_tx = file
                    elif "GoParkin Corporate Vehicles" in file.name:
                        file_gp_crm = file
                    # 🌟 只要名字里带 "EVOne Report breakdown" 或 "EVOne Corporate fleet" 都认作 SP 数据源
                    elif "EVOne Report breakdown" in file.name or "EVOne Corporate fleet" in file.name:
                        file_sp_tx = file
                    elif "SP Corporate Vehicles" in file.name:
                        file_sp_crm = file

                if not all([file_gp_tx, file_gp_crm, file_sp_tx, file_sp_crm]):
                    st.error("❌ 文件名称不匹配！请确保你拖入的文件中包含了那 4 个核心文件。")
                    st.stop()

                # --- B. 处理 GoParkin 数据 (通过车牌) ---
                crm_gp = read_file(file_gp_crm)[['Vehicle No.', 'Company']].dropna()
                crm_gp['Vehicle No.'] = crm_gp['Vehicle No.'].astype(str).str.strip().str.upper()

                gp_tx = read_file(file_gp_tx)
                gp_tx = gp_tx[gp_tx['payment_status'] == 'Success'].copy()
                gp_tx['vehicle_plate_number'] = gp_tx['vehicle_plate_number'].astype(str).str.strip().str.upper()
                gp_tx['Year-Month'] = gp_tx['start_date_time'].astype(str).str[0:7]

                gp_merged = pd.merge(gp_tx, crm_gp, left_on='vehicle_plate_number', right_on='Vehicle No.', how='left')
                gp_merged['Company'] = gp_merged['Company'].fillna('Unmatched GoParkin') 
                gp_summary = gp_merged.groupby(['Company', 'Year-Month'])['total_energy_supplied_kwh'].sum().reset_index()
                gp_summary.rename(columns={'total_energy_supplied_kwh': 'GoParkin(kWh)'}, inplace=True)

                # --- C. 处理 SP 数据 (通过邮箱) ---
                crm_sp = read_file(file_sp_crm)[['Email', 'Company']].dropna()
                crm_sp['Email'] = crm_sp['Email'].astype(str).str.strip().str.lower()

                # 🌟 核心升级：如果是 Excel，精准读取指定的 Sheet
                if file_sp_tx.name.endswith('.xlsx') or file_sp_tx.name.endswith('.xls'):
                    sp_tx = pd.read_excel(file_sp_tx, sheet_name='EVOne Corporate fleet')
                else:
                    sp_tx = read_file(file_sp_tx) # 兼容同事直接上传 CSV 的情况
                    
                sp_tx['Driver Email'] = sp_tx['Driver Email'].astype(str).str.strip().str.lower()
                sp_tx['Year-Month'] = sp_tx['Date'].astype(str).str[0:7]
                sp_tx['CDR Total Energy'] = pd.to_numeric(sp_tx['CDR Total Energy'], errors='coerce').fillna(0)

                sp_merged = pd.merge(sp_tx, crm_sp, left_on='Driver Email', right_on='Email', how='left')
                sp_merged['Company'] = sp_merged['Company'].fillna('Unmatched SP Email')
                sp_summary = sp_merged.groupby(['Company', 'Year-Month'])['CDR Total Energy'].sum().reset_index()
                sp_summary.rename(columns={'CDR Total Energy': 'SP(kWh)'}, inplace=True)

                # --- D. 合并终极账单 ---
                final_df = pd.merge(gp_summary, sp_summary, on=['Company', 'Year-Month'], how='outer').fillna(0)
                if 'GoParkin(kWh)' not in final_df.columns: final_df['GoParkin(kWh)'] = 0
                if 'SP(kWh)' not in final_df.columns: final_df['SP(kWh)'] = 0
                final_df['Total(kWh)'] = final_df['GoParkin(kWh)'] + final_df['SP(kWh)']

                months = sorted([m for m in final_df['Year-Month'].dropna().unique() if len(str(m)) == 7]) 

                # --- E. 写入内存中的 Excel ---
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    for month in months:
                        month_df = final_df[final_df['Year-Month'] == month].copy()
                        month_df = month_df.sort_values(by='Company').reset_index(drop=True)
                        month_df.insert(0, 'S/N', month_df.index + 1)
                        month_df.to_excel(writer, sheet_name=month, index=False)
                
                st.success("✅ 处理大功告成！系统已成功跨 Sheet 提取数据并完成合并。")
                
                # --- F. 提供下载按钮 ---
                st.download_button(
                    label="📥 点击下载最终 Excel 报表",
                    data=output.getvalue(),
                    file_name="Monthly_Billing_Report_Final.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"❌ 处理过程中出现错误，请检查上传的文件内容是否被更改。错误详情：{str(e)}")