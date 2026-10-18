import streamlit as st
import pandas as pd
import plotly.express as px

# ---------------------------------------------------------
# 1. Page Configuration
# ---------------------------------------------------------
st.set_page_config(
    page_title="Daily Sales & Stock Dashboard (ALL)",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.title("📊 Comprehensive Sales & Stock Dashboard")
st.markdown("---")

# Number formatting function (M for Millions, K for Thousands)
def format_num(val, is_currency=True):
    suffix = " EGP" if is_currency else ""
    if abs(val) >= 1_000_000:
        return f"{val / 1_000_000:,.2f}M{suffix}"
    elif abs(val) >= 1_000:
        return f"{val / 1_000:,.1f}K{suffix}"
    else:
        return f"{val:,.0f}{suffix}"

# ---------------------------------------------------------
# 2. File Uploader & Data Loading
# ---------------------------------------------------------
st.sidebar.header("⚙️ Data Settings")
uploaded_file = st.sidebar.file_uploader("Upload your Excel Sales file", type=["xlsx", "xls"])

@st.cache_data
def load_and_clean_data(file):
    df_raw = pd.read_excel(file)
    data = df_raw.iloc[3:].copy().reset_index(drop=True)
    
    col_mapping = {
        df_raw.columns[0]: "Line",
        df_raw.columns[1]: "Product_ID",
        df_raw.columns[2]: "Product_Name",
        df_raw.columns[3]: "Private_Sales_Q",
        df_raw.columns[4]: "Private_Sales_V",
        df_raw.columns[5]: "Private_Target_Q",
        df_raw.columns[6]: "Private_Target_V",
        df_raw.columns[7]: "Private_Ach_Pct",
        df_raw.columns[8]: "IbnSina_Sales_Q",
        df_raw.columns[9]: "IbnSina_Sales_V",
        df_raw.columns[10]: "Pharma_Sales_Q",
        df_raw.columns[11]: "Pharma_Sales_V",
        df_raw.columns[12]: "Ramco_Sales_Q",
        df_raw.columns[13]: "Ramco_Sales_V",
        df_raw.columns[14]: "UCP_Sales_Q",
        df_raw.columns[15]: "UCP_Sales_V",
        df_raw.columns[16]: "MEC_Sales_Q",
        df_raw.columns[17]: "MEC_Sales_V",
        df_raw.columns[18]: "Sofico_Sales_Q",
        df_raw.columns[19]: "Sofico_Sales_V",
        df_raw.columns[20]: "ETC_Sales_Q",
        df_raw.columns[21]: "ETC_Sales_V",
        df_raw.columns[42]: "Stock_Days",
        df_raw.columns[43]: "Stock_Qty_NoBadr",
        df_raw.columns[44]: "Stock_Val_NoBadr",
        df_raw.columns[45]: "Total_Stock_Qty",
        df_raw.columns[58]: "Tender_Sales_Q",
        df_raw.columns[59]: "Tender_Sales_V",
        df_raw.columns[60]: "Total_Sales_Q",
        df_raw.columns[61]: "Total_Sales_V",
        df_raw.columns[62]: "Total_Target_Q",
        df_raw.columns[63]: "Total_Target_V",
        df_raw.columns[64]: "Total_Ach_Pct"
    }
    
    data = data.rename(columns=col_mapping)
    
    data_clean = data[
        ~data['Line'].astype(str).str.contains("Total", case=False, na=False) &
        data['Product_Name'].notna()
    ].copy()
    
    for col in data_clean.columns:
        if col not in ["Line", "Product_ID", "Product_Name"]:
            data_clean[col] = pd.to_numeric(data_clean[col], errors='coerce').fillna(0)
            
    return data_clean, df_raw

if uploaded_file is not None:
    df, df_raw = load_and_clean_data(uploaded_file)
    
    # ---------------------------------------------------------
    # 3. Sidebar Filters
    # ---------------------------------------------------------
    st.sidebar.subheader("🔍 Filters")
    
    # Product Line Filter
    lines = ["All"] + sorted([str(l) for l in df['Line'].dropna().unique()])
    selected_line = st.sidebar.selectbox("Select Product Line:", lines)
    
    # Distributor Filter
    distributor_map = {
        "All Distributors": "Total_Sales_V",
        "IBN SINA": "IbnSina_Sales_V",
        "PHARMA (POS)": "Pharma_Sales_V",
        "RAMCO": "Ramco_Sales_V",
        "UCP": "UCP_Sales_V",
        "MEC": "MEC_Sales_V",
        "SOFICO": "Sofico_Sales_V",
        "ETC": "ETC_Sales_V"
    }
    selected_distributor = st.sidebar.selectbox("Select Distributor:", list(distributor_map.keys()))
    sales_col = distributor_map[selected_distributor]

    # Applying Line Filter
    if selected_line != "All":
        df_filtered = df[df['Line'] == selected_line].copy()
    else:
        df_filtered = df.copy()

    # ---------------------------------------------------------
    # 4. KPIs Summary Cards
    # ---------------------------------------------------------
    tot_sales_v = df_filtered[sales_col].sum()
    tot_target_v = df_filtered['Total_Target_V'].sum()
    ach_pct = (tot_sales_v / tot_target_v * 100) if tot_target_v > 0 else 0
    
    priv_sales_v = df_filtered['Private_Sales_V'].sum()
    tender_sales_v = df_filtered['Tender_Sales_V'].sum()
    
    stock_val = df_filtered['Stock_Val_NoBadr'].sum()
    stock_qty = df_filtered['Total_Stock_Qty'].sum()
    
    ibnsina_sales_v = df_filtered['IbnSina_Sales_V'].sum()
    ibnsina_share = (ibnsina_sales_v / priv_sales_v * 100) if priv_sales_v > 0 else 0
    
    total_skus = len(df_filtered)
    active_skus = len(df_filtered[df_filtered[sales_col] > 0])
    zero_sales_skus = len(df_filtered[df_filtered[sales_col] == 0])

    st.subheader(f"🎯 Sales Performance KPIs ({selected_distributor})")
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric(f"Sales ({selected_distributor})", format_num(tot_sales_v))
    c2.metric("Total Target", format_num(tot_target_v))
    c3.metric("Achievement %", f"{ach_pct:.2f}%")
    c4.metric("Private Sales", format_num(priv_sales_v))
    c5.metric("Tender Sales", format_num(tender_sales_v))

    st.markdown("---")
    st.subheader("📦 Inventory & Distribution KPIs")
    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("Stock Value", format_num(stock_val))
    k2.metric("Stock Quantity", format_num(stock_qty, is_currency=False))
    k3.metric("Ibn Sina Share (Private)", f"{ibnsina_share:.1f}%")
    k4.metric("Active SKUs", f"{active_skus} / {total_skus}")
    k5.metric("Zero Sales SKUs", f"{zero_sales_skus}", delta=f"-{(zero_sales_skus/total_skus*100):.1f}%", delta_color="inverse")

    st.markdown("---")

    # ---------------------------------------------------------
    # 5. Tabs & Visualizations
    # ---------------------------------------------------------
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "📈 Line Performance",
        "🏆 Top Products",
        "🤝 Top Product per Distributor",
        "⚖️ Private & Stock Analysis",
        "📋 Detailed Data Table"
    ])

    with tab1:
        st.subheader(f"Sales ({selected_distributor}) vs Target by Product Line")
        line_summary = df_filtered.groupby('Line')[[sales_col, 'Total_Target_V']].sum().reset_index()
        line_summary['Ach_%'] = (line_summary[sales_col] / line_summary['Total_Target_V'] * 100).round(2)
        line_summary = line_summary.sort_values(by=sales_col, ascending=False)
        
        fig_line = px.bar(line_summary, x='Line', y=[sales_col, 'Total_Target_V'], barmode='group', title=f"Sales ({selected_distributor}) vs Target per Line")
        st.plotly_chart(fig_line, use_container_width=True)
        st.dataframe(line_summary.style.format({sales_col: '{:,.0f}', 'Total_Target_V': '{:,.0f}', 'Ach_%': '{:.2f}%'}), use_container_width=True)

    with tab2:
        col_top, col_zero = st.columns(2)
        with col_top:
            st.subheader(f"🏆 Top 10 Best-Selling Products ({selected_distributor})")
            top_products = df_filtered.nlargest(10, sales_col)[['Product_Name', 'Line', sales_col]]
            fig_top = px.bar(top_products, x=sales_col, y='Product_Name', orientation='h', color='Line', title=f"Top Products by Value ({selected_distributor})")
            fig_top.update_layout(yaxis={'categoryorder': 'total ascending'})
            st.plotly_chart(fig_top, use_container_width=True)
            
        with col_zero:
            st.subheader(f"⚠️ Zero Sales Products ({selected_distributor})")
            low_p = df_filtered[(df_filtered[sales_col] == 0) & (df_filtered['Total_Target_V'] > 0)]
            st.dataframe(low_p[['Product_Name', 'Line', 'Total_Target_V', 'Total_Stock_Qty']], use_container_width=True)

    with tab3:
        st.subheader("🤝 Top Product per Distributor Overview")
        dist_cols = {
            "IBN SINA": "IbnSina_Sales_V",
            "PHARMA (POS)": "Pharma_Sales_V",
            "RAMCO": "Ramco_Sales_V",
            "UCP": "UCP_Sales_V",
            "MEC": "MEC_Sales_V",
            "SOFICO": "Sofico_Sales_V",
            "ETC": "ETC_Sales_V"
        }
        
        dist_top_list = []
        for d_name, d_col in dist_cols.items():
            if d_col in df.columns:
                top_row = df_filtered.loc[df_filtered[d_col].idxmax()] if not df_filtered.empty and df_filtered[d_col].max() > 0 else None
                if top_row is not None:
                    dist_top_list.append({
                        "Distributor": d_name,
                        "Top Product": top_row['Product_Name'],
                        "Product Line": top_row['Line'],
                        "Sales Value (EGP)": top_row[d_col]
                    })
        
        if dist_top_list:
            dist_top_df = pd.DataFrame(dist_top_list)
            st.dataframe(dist_top_df.style.format({'Sales Value (EGP)': '{:,.2f}'}), use_container_width=True)

    with tab4:
        col_pie1, col_pie2 = st.columns(2)
        with col_pie1:
            sales_split = pd.DataFrame({'Category': ['Private Sales', 'Tender Sales'], 'Value': [priv_sales_v, tender_sales_v]})
            fig_pie = px.pie(sales_split, names='Category', values='Value', title="Private vs Tender Sales Split", hole=0.4)
            st.plotly_chart(fig_pie, use_container_width=True)
        with col_pie2:
            top_stock = df_filtered.nlargest(10, 'Total_Stock_Qty')[['Product_Name', 'Line', 'Total_Stock_Qty']]
            fig_stock = px.bar(top_stock, x='Total_Stock_Qty', y='Product_Name', orientation='h', color='Line', title="Top 10 Stock Quantities")
            fig_stock.update_layout(yaxis={'categoryorder': 'total ascending'})
            st.plotly_chart(fig_stock, use_container_width=True)

    with tab5:
        st.subheader("View & Export Filtered Data")
        st.dataframe(df_filtered, use_container_width=True)
        csv = df_filtered.to_csv(index=False).encode('utf-8-sig')
        st.download_button("📥 Download Filtered Data (CSV)", data=csv, file_name="Filtered_Daily_Sales.csv", mime="text/csv")

else:
    st.info("👈 Please upload your Excel sales file from the sidebar to launch the Dashboard.")