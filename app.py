import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

# ---------------------------------------------------------
# 1. Page Configuration
# ---------------------------------------------------------
st.set_page_config(
    page_title="Sales Dashboard",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Sales & Performance Dashboard")
st.markdown("---")

# ---------------------------------------------------------
# 2. Data Loading & Cleaning
# ---------------------------------------------------------
@st.cache_data
def load_and_clean_data(file):
    raw_df = pd.read_excel(file, header=None)
    df = raw_df.iloc[3:].copy()
    
    data = pd.DataFrame()
    data['Cluster'] = df.iloc[:, 0].astype(str).str.strip()
    data['Product_ID'] = df.iloc[:, 1]
    data['Product_Name'] = df.iloc[:, 2].astype(str).str.strip()
    
    # Private Sales & Target
    data['Private_Sales_Qty'] = pd.to_numeric(df.iloc[:, 3], errors='coerce').fillna(0)
    data['Private_Sales_Val'] = pd.to_numeric(df.iloc[:, 4], errors='coerce').fillna(0)
    data['Private_Target_Qty'] = pd.to_numeric(df.iloc[:, 5], errors='coerce').fillna(0)
    data['Private_Target_Val'] = pd.to_numeric(df.iloc[:, 6], errors='coerce').fillna(0)
    
    # Distributors Sales
    data['IbnSina_Sales_Val'] = pd.to_numeric(df.iloc[:, 9], errors='coerce').fillna(0)
    data['Pharma_Sales_Val'] = pd.to_numeric(df.iloc[:, 11], errors='coerce').fillna(0)
    data['Ramco_Sales_Val'] = pd.to_numeric(df.iloc[:, 13], errors='coerce').fillna(0)
    data['UCP_Sales_Val'] = pd.to_numeric(df.iloc[:, 15], errors='coerce').fillna(0)
    data['MEC_Sales_Val'] = pd.to_numeric(df.iloc[:, 17], errors='coerce').fillna(0)
    data['Sofico_Sales_Val'] = pd.to_numeric(df.iloc[:, 19], errors='coerce').fillna(0)
    data['ETC_Sales_Val'] = pd.to_numeric(df.iloc[:, 21], errors='coerce').fillna(0)
    data['Direct_Sales_Val'] = pd.to_numeric(df.iloc[:, 23], errors='coerce').fillna(0)
    data['Stores_Sales_Val'] = pd.to_numeric(df.iloc[:, 25], errors='coerce').fillna(0)
    
    # Stock
    data['Stock_Days'] = pd.to_numeric(df.iloc[:, 42], errors='coerce').fillna(0)
    data['Total_Stock_Qty'] = pd.to_numeric(df.iloc[:, 45], errors='coerce').fillna(0)
    data['Total_Stock_Val'] = pd.to_numeric(df.iloc[:, 46], errors='coerce').fillna(0)
    
    # Tender & Totals
    data['Tender_Sales_Qty'] = pd.to_numeric(df.iloc[:, 58], errors='coerce').fillna(0)
    data['Tender_Sales_Val'] = pd.to_numeric(df.iloc[:, 59], errors='coerce').fillna(0)
    data['Total_Sales_Qty'] = pd.to_numeric(df.iloc[:, 60], errors='coerce').fillna(0)
    data['Total_Sales_Val'] = pd.to_numeric(df.iloc[:, 61], errors='coerce').fillna(0)
    data['Total_Target_Qty'] = pd.to_numeric(df.iloc[:, 62], errors='coerce').fillna(0)
    data['Total_Target_Val'] = pd.to_numeric(df.iloc[:, 63], errors='coerce').fillna(0)
    
    data = data[data['Product_Name'].str.lower() != 'nan']
    return data

# ---------------------------------------------------------
# 3. Sidebar & Filters
# ---------------------------------------------------------
st.sidebar.header("📁 File Upload")
uploaded_file = st.sidebar.file_uploader("Choose Daily Sales ALL Excel File", type=["xlsx", "xls"])

if uploaded_file is not None:
    df = load_and_clean_data(uploaded_file)
    
    st.sidebar.header("🔍 Interactive Filters")
    selected_clusters = st.sidebar.multiselect("Select Line (Cluster):", options=df['Cluster'].unique(), default=df['Cluster'].unique())
    filtered_df = df[df['Cluster'].isin(selected_clusters)]
    
    # ---------------------------------------------------------
    # 4. Executive KPIs Overview
    # ---------------------------------------------------------
    st.subheader("🎯 Executive Overview & Target KPIs")
    
    total_sales = filtered_df['Total_Sales_Val'].sum()
    total_target = filtered_df['Total_Target_Val'].sum()
    total_ach = (total_sales / total_target * 100) if total_target > 0 else 0
    target_gap = total_target - total_sales
    
    private_sales = filtered_df['Private_Sales_Val'].sum()
    tender_sales = filtered_df['Tender_Sales_Val'].sum()
    
    active_skus = (filtered_df['Total_Sales_Val'] > 0).sum()
    total_skus = len(filtered_df)
    zero_sales_skus = total_skus - active_skus
    
    kpi1, kpi2, kpi3, kpi4, kpi5, kpi6 = st.columns(6)
    
    kpi1.metric("Total Sales", f"{total_sales/1e6:.2f}M EGP")
    kpi2.metric("Total Target", f"{total_target/1e6:.2f}M EGP")
    kpi3.metric("Total Achievement", f"{total_ach:.1f}%")
    kpi4.metric("Target Gap", f"{target_gap/1e6:.2f}M EGP")
    kpi5.metric("Private Sales", f"{private_sales/1e6:.2f}M EGP")
    kpi6.metric("Zero Sales SKUs", f"{zero_sales_skus} / {total_skus}")
    
    st.markdown("---")

    # ---------------------------------------------------------
    # 5. Visuals 1: Treemap & Radar Chart
    # ---------------------------------------------------------
    st.subheader("🗺️ Structural Breakdown & Cluster Radar Analysis")
    col_vis1, col_vis2 = st.columns([3, 2])
    
    with col_vis1:
        fig_tree = px.treemap(
            filtered_df,
            path=['Cluster', 'Product_Name'],
            values='Total_Sales_Val',
            color='Total_Sales_Val',
            color_continuous_scale='Blues',
            title="Sales Treemap Distribution by Cluster & Product"
        )
        st.plotly_chart(fig_tree, use_container_width=True)
        
    with col_vis2:
        cluster_summary = filtered_df.groupby('Cluster').agg({
            'Total_Sales_Val': 'sum',
            'Total_Target_Val': 'sum'
        }).reset_index()
        cluster_summary['Ach_Pct'] = (cluster_summary['Total_Sales_Val'] / cluster_summary['Total_Target_Val']) * 100
        
        fig_radar = go.Figure(data=go.Scatterpolar(
            r=cluster_summary['Ach_Pct'],
            theta=cluster_summary['Cluster'],
            fill='toself',
            name='Achievement %',
            line_color='#2ecc71'
        ))
        fig_radar.update_layout(
            polar=dict(radialaxis=dict(visible=True, range=[0, max(cluster_summary['Ach_Pct'].max(), 100)])),
            title="Cluster Achievement % Radar Chart",
            showlegend=False
        )
        st.plotly_chart(fig_radar, use_container_width=True)

    st.markdown("---")

    # ---------------------------------------------------------
    # 6. Visuals 2: Private vs Tender & Distributor Breakdown
    # ---------------------------------------------------------
    st.subheader("🏢 Private vs Tender Sales & Distributor Share")
    col_vis3, col_vis4 = st.columns(2)
    
    with col_vis3:
        cluster_pt = filtered_df.groupby('Cluster').agg({
            'Private_Sales_Val': 'sum',
            'Tender_Sales_Val': 'sum'
        }).reset_index()
        
        fig_pt = px.bar(
            cluster_pt,
            x='Cluster',
            y=['Private_Sales_Val', 'Tender_Sales_Val'],
            title="Private vs Tender Sales per Cluster",
            barmode='stack',
            labels={'value': 'Value (EGP)', 'variable': 'Sector'},
            color_discrete_map={'Private_Sales_Val': '#3498db', 'Tender_Sales_Val': '#f39c12'}
        )
        st.plotly_chart(fig_pt, use_container_width=True)
        
    with col_vis4:
        dist_cols = {
            'Ibn Sina': filtered_df['IbnSina_Sales_Val'].sum(),
            'Pharma': filtered_df['Pharma_Sales_Val'].sum(),
            'Ramco': filtered_df['Ramco_Sales_Val'].sum(),
            'UCP': filtered_df['UCP_Sales_Val'].sum(),
            'MEC': filtered_df['MEC_Sales_Val'].sum(),
            'Sofico': filtered_df['Sofico_Sales_Val'].sum(),
            'ETC': filtered_df['ETC_Sales_Val'].sum(),
            'Direct': filtered_df['Direct_Sales_Val'].sum(),
            'Stores': filtered_df['Stores_Sales_Val'].sum()
        }
        dist_df = pd.DataFrame(list(dist_cols.items()), columns=['Distributor', 'Sales_Value'])
        dist_df = dist_df[dist_df['Sales_Value'] > 0]
        
        fig_donut = px.pie(
            dist_df,
            names='Distributor',
            values='Sales_Value',
            title="Distributor Private Sales Market Share",
            hole=0.4,
            color_discrete_sequence=px.colors.qualitative.Pastel
        )
        st.plotly_chart(fig_donut, use_container_width=True)

    st.markdown("---")

    # ---------------------------------------------------------
    # 7. Visuals 3: Top Products & Stock Days Analysis
    # ---------------------------------------------------------
    st.subheader("🏆 Top Performing Products & Stock Days Analysis")
    col_vis5, col_vis6 = st.columns(2)
    
    with col_vis5:
        top_10 = filtered_df.nlargest(10, 'Total_Sales_Val')
        fig_top = px.bar(
            top_10,
            x='Total_Sales_Val',
            y='Product_Name',
            orientation='h',
            title="Top 10 Products by Sales Value",
            color='Total_Sales_Val',
            color_continuous_scale='Teal'
        )
        fig_top.update_layout(yaxis={'categoryorder':'total ascending'})
        st.plotly_chart(fig_top, use_container_width=True)
        
    with col_vis6:
        stock_cluster = filtered_df.groupby('Cluster').agg({
            'Stock_Days': 'mean',
            'Total_Stock_Val': 'sum'
        }).reset_index()
        
        fig_stock = px.scatter(
            stock_cluster,
            x='Total_Stock_Val',
            y='Stock_Days',
            size='Total_Stock_Val',
            color='Cluster',
            text='Cluster',
            title="Average Stock Cover Days vs Total Stock Value",
            labels={'Total_Stock_Val': 'Total Stock Value', 'Stock_Days': 'Average Stock Days'}
        )
        st.plotly_chart(fig_stock, use_container_width=True)

    st.markdown("---")

    # ---------------------------------------------------------
    # 8. Detailed Data Table
    # ---------------------------------------------------------
    st.subheader("📋 Detailed Data Table")
    
    search_term = st.text_input("🔎 Search Product Name:")
    if search_term:
        display_df = filtered_df[filtered_df['Product_Name'].str.contains(search_term, case=False, na=False)]
    else:
        display_df = filtered_df
        
    st.dataframe(
        display_df[[
            'Cluster', 'Product_ID', 'Product_Name', 
            'Total_Sales_Qty', 'Total_Sales_Val', 
            'Total_Target_Val', 'Total_Stock_Qty', 'Stock_Days'
        ]].sort_values(by='Total_Sales_Val', ascending=False),
        use_container_width=True
    )

else:
    st.info("👋 Please upload your Daily Sales ALL Excel file from the sidebar to view the updated dashboard.")
