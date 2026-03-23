import streamlit as st
import pandas as pd
import io
import warnings

warnings.filterwarnings('ignore')

st.set_page_config(page_title="充电数据合并工具", layout="wide")

st.title("🚗 充电账单自动合并系统")
st.markdown("请上传对应的源文件，系统将自动进行数据清洗与合并。")

# --- 1. 文件上传区 ---
with st.sidebar:
    st.header("文件上传")
    gp_tx_file = st.file_uploader("上传 GoParkin 交易明细 (.csv)", type=["csv"])
    gp_crm_file = st.file_uploader("上传 GoParkin CRM (.xlsx)", type=["xlsx"])
    st.divider()
    sp_tx_file = st.file_uploader("上传 SP 交易明细 (.xlsx)", type=["xlsx"])
    sp_crm_file = st.file_uploader("上传 SP CRM (.xlsx)", type=["xlsx"])

# --- 2. 核心处理逻辑 ---
if st.button("开始合并数据"):
    if not all([gp_tx_file, gp_crm_file, sp_tx_file, sp_crm_file]):
        st.error("❌ 请先上传所有四个必要文件！")
    else:
        try:
            with st.status("正在处理数据...", expanded=True) as status:
                # 处理 GoParkin
                st.write("读取 GoParkin 数据...")
                crm_gp = pd.read_excel(gp_crm_file)
                crm_gp = crm_gp[['Vehicle No.', 'Company']].dropna()
                crm_gp['Vehicle No.'] = crm_gp['Vehicle No.'].astype(str).str.strip().str.upper()

                gp_tx = pd.read_csv(gp_tx_file)
                gp_tx = gp_tx[gp_tx['payment_status'] == 'Success'].copy()
                gp_tx['vehicle_plate_number'] = gp_tx['vehicle_plate_number'].astype(str).str.strip().str.upper()
                gp_tx['Year-Month'] = gp_tx['start_date_time'].astype(str).str[0:7]

                gp_merged = pd.merge(gp_tx, crm_gp, left_on='vehicle_plate_number', right_on='Vehicle No.', how='left')
                gp_merged['Company'] = gp_merged['Company'].fillna('Unmatched GoParkin')
                gp_summary = gp_merged.groupby(['Company', 'Year-Month'])['total_energy_supplied_kwh'].sum().reset_index()
                gp_summary.rename(columns={'total_energy_supplied_kwh': 'GoParkin(kWh)'}, inplace=True)

                # 处理 SP
                st.write("读取 SP 数据...")
                crm_sp = pd.read_excel(sp_crm_file)
                crm_sp = crm_sp[['Email', 'Company']].dropna()
                crm_sp['Email'] = crm_sp['Email'].astype(str).str.strip().str.lower()

                sp_tx = pd.read_excel(sp_tx_file, sheet_name='EVOne Corporate fleet')
                sp_tx['Driver Email'] = sp_tx['Driver Email'].astype(str).str.strip().str.lower()
                sp_tx['Year-Month'] = sp_tx['Date'].astype(str).str[0:7]
                sp_tx['CDR Total Energy'] = pd.to_numeric(sp_tx['CDR Total Energy'], errors='coerce').fillna(0)

                sp_merged = pd.merge(sp_tx, crm_sp, left_on='Driver Email', right_on='Email', how='left')
                sp_merged['Company'] = sp_merged['Company'].fillna('Unmatched SP Email')
                sp_summary = sp_merged.groupby(['Company', 'Year-Month'])['CDR Total Energy'].sum().reset_index()
                sp_summary.rename(columns={'CDR Total Energy': 'SP(kWh)'}, inplace=True)

                # 合并
                st.write("合并并生成报表...")
                final_df = pd.merge(gp_summary, sp_summary, on=['Company', 'Year-Month'], how='outer').fillna(0)
                final_df['Total(kWh)'] = final_df.get('GoParkin(kWh)', 0) + final_df.get('SP(kWh)', 0)

                # 写入内存缓冲区
                output = io.BytesIO()
                months = sorted([m for m in final_df['Year-Month'].dropna().unique() if len(str(m)) == 7])
                
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    for month in months:
                        month_df = final_df[final_df['Year-Month'] == month].copy()
                        month_df = month_df.sort_values(by='Company').reset_index(drop=True)
                        month_df.insert(0, 'S/N', month_df.index + 1)
                        month_df.to_excel(writer, sheet_name=month, index=False)
                
                status.update(label="✅ 处理完成!", state="complete", expanded=False)

            # --- 3. 下载区 ---
            st.success("报表生成成功！")
            st.download_button(
                label="📥 下载最终账单报表 (Excel)",
                data=output.getvalue(),
                file_name="Monthly_Billing_Report_Final.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.dataframe(final_df) # 在页面预览数据

        except Exception as e:
            st.error(f"处理过程中出错: {e}")