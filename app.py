import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

# ---------------------------------------------------------
# 1. Page Setup & Theme Handling
# ---------------------------------------------------------
st.set_page_config(
    page_title="Sales & Distribution Dashboard",
    page_icon="📊",
    layout="wide"
)

# Responsive styling compatible with Dark & Light Mode
st.markdown("""
    <style>
    div[data-testid="stMetric"] {
        background-color: rgba(255, 255, 255, 0.05);
        border: 1px solid rgba(255, 255, 255, 0.1);
        padding: 15px;
        border-radius: 10px;
        border-left: 5px solid #007bff;
    }
    </style>
""", unsafe_allow_html=True)

st.title("📊 Sales & Distribution Dashboard")
st.markdown("---")

# ---------------------------------------------------------
# 2. Data Processing Function
# ---------------------------------------------------------
@st.cache_data
def load_and_clean_data(file):
    raw_df = pd.read_excel(file, header=None)
    df = raw_df.iloc[3:].copy()
    
    data = pd.DataFrame()
    data['Cluster'] = df.iloc[:, 0].astype(str).str.strip()
    data['Product_ID'] = df.iloc[:, 1]
    data['Product_Name'] = df.iloc[:, 2].astype(str).str.strip()
    
    # Filter out empty or header rows
    data = data[data['Product_Name'].notna()]
    data = data[~data['Product_Name'].str.lower().str.contains("crm product name|total|sum|nan")]
    
    # Private Sales & Target
    data['Private_Sales_Qty'] = pd.to_numeric(df.iloc[:, 3], errors='coerce').fillna(0)
    data['Private_Sales_Val'] = pd.to_numeric(df.iloc[:, 4], errors='coerce').fillna(0)
    data['Private_Target_Qty'] = pd.to_numeric(df.iloc[:, 5], errors='coerce').fillna(0)
    data['Private_Target_Val'] = pd.to_numeric(df.iloc[:, 6], errors='coerce').fillna(0)
    
    # Distributors Sales Values
    data['Ibn Sina'] = pd.to_numeric(df.iloc[:, 9], errors='coerce').fillna(0)
    data['Pharma'] = pd.to_numeric(df.iloc[:, 11], errors='coerce').fillna(0)
    data['Ramco'] = pd.to_numeric(df.iloc[:, 13], errors='coerce').fillna(0)
    data['UCP'] = pd.to_numeric(df.iloc[:, 15], errors='coerce').fillna(0)
    data['MEC'] = pd.to_numeric(df.iloc[:, 17], errors='coerce').fillna(0)
    data['Sofico'] = pd.to_numeric(df.iloc[:, 19], errors='coerce').fillna(0)
    data['ETC'] = pd.to_numeric(df.iloc[:, 21], errors='coerce').fillna(0)
    data['Direct'] = pd.to_numeric(df.iloc[:, 23], errors='coerce').fillna(0)
    data['Stores'] = pd.to_numeric(df.iloc[:, 25], errors='coerce').fillna(0)
    
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
    
    return data

# ---------------------------------------------------------
# 3. Sidebar Upload & Filters
# ---------------------------------------------------------
st.sidebar.header("📁 Control Panel")
uploaded_file = st.sidebar.file_uploader("Upload Excel File", type=["xlsx", "xls"])

if uploaded_file is not None:
    df = load_and_clean_data(uploaded_file)
    
    st.sidebar.subheader("Filter Options")
    selected_clusters = st.sidebar.multiselect(
        "Select Cluster / Line:", 
        options=df['Cluster'].unique(), 
        default=df['Cluster'].unique()
    )
    
    filtered_df = df[df['Cluster'].isin(selected_clusters)]
    
    # Handling Empty Selection Error Safely
    if filtered_df.empty:
        st.warning("⚠️ Please select at least one Cluster / Line from the sidebar filter options.")
    else:
        # ---------------------------------------------------------
        # 4. Top Executive Metrics (KPIs)
        # ---------------------------------------------------------
        st.subheader("📌 Performance Summary")
        
        total_sales = filtered_df['Total_Sales_Val'].sum()
        total_target = filtered_df['Total_Target_Val'].sum()
        total_ach = (total_sales / total_target * 100) if total_target > 0 else 0
        private_sales = filtered_df['Private_Sales_Val'].sum()
        tender_sales = filtered_df['Tender_Sales_Val'].sum()
        
        k1, k2, k3, k4, k5 = st.columns(5)
        k1.metric("Total Sales", f"{total_sales/1e6:.2f}M EGP")
        k2.metric("Target", f"{total_target/1e6:.2f}M EGP")
        k3.metric("Achievement %", f"{total_ach:.1f}%")
        k4.metric("Private Sales", f"{private_sales/1e6:.2f}M EGP")
        k5.metric("Tender Sales", f"{tender_sales/1e6:.2f}M EGP")
        
        # Progress gauge bar
        st.progress(min(max(int(total_ach), 0), 100), text=f"Overall Target Achievement: {total_ach:.1f}%")
        st.markdown("---")

        # ---------------------------------------------------------
        # 5. Top Product Per Distributor Section
        # ---------------------------------------------------------
        st.subheader("🌟 Top Product for Each Distributor")
        
        dist_names = ['Ibn Sina', 'Pharma', 'Ramco', 'UCP', 'MEC', 'Sofico', 'ETC', 'Direct', 'Stores']
        top_products_data = []
        
        for dist in dist_names:
            dist_total = filtered_df[dist].sum()
            if dist_total > 0:
                top_row = filtered_df.nlargest(1, dist).iloc[0]
                top_products_data.append({
                    'Distributor': dist,
                    'Top Product': top_row['Product_Name'],
                    'Sales Value (EGP)': top_row[dist],
                    'Distributor Total Sales': dist_total,
                    '% of Distributor Sales': (top_row[dist] / dist_total * 100)
                })
                
        top_dist_df = pd.DataFrame(top_products_data)
        
        if not top_dist_df.empty:
            col_dist_cards, col_dist_selector = st.columns([3, 2])
            
            with col_dist_cards:
                fig_top_dist = px.bar(
                    top_dist_df,
                    x='Sales Value (EGP)',
                    y='Distributor',
                    text='Top Product',
                    orientation='h',
                    title="Top Product Sales per Distributor",
                    color='Sales Value (EGP)',
                    color_continuous_scale='Blues'
                )
                fig_top_dist.update_traces(textposition='inside', texttemplate='%{text}: %{x:,.0f} EGP')
                fig_top_dist.update_layout(yaxis={'categoryorder':'total ascending'}, showlegend=False)
                st.plotly_chart(fig_top_dist, use_container_width=True)
                
            with col_dist_selector:
                st.markdown("##### 🔍 Deep Dive by Distributor")
                selected_dist = st.selectbox("Choose a Distributor:", dist_names)
                
                top_5_dist = filtered_df.nlargest(5, selected_dist)[['Product_Name', selected_dist]]
                top_5_dist = top_5_dist[top_5_dist[selected_dist] > 0]
                
                if not top_5_dist.empty:
                    fig_dist_deep = px.bar(
                        top_5_dist,
                        x=selected_dist,
                        y='Product_Name',
                        orientation='h',
                        title=f"Top 5 Products in {selected_dist}",
                        text_auto=',.0f',
                        color_discrete_sequence=['#2ecc71']
                    )
                    fig_dist_deep.update_layout(yaxis={'categoryorder':'total ascending'})
                    st.plotly_chart(fig_dist_deep, use_container_width=True)
                else:
                    st.info(f"No sales recorded for {selected_dist} in selected filter(s).")
        else:
            st.info("No distributor sales available for current filter selection.")

        st.markdown("---")

        # ---------------------------------------------------------
        # 6. Cluster Performance & Distributor Breakdown
        # ---------------------------------------------------------
        st.subheader("📈 Cluster Performance & Market Share")
        c1, c2 = st.columns(2)
        
        with c1:
            cluster_acc = filtered_df.groupby('Cluster').agg({
                'Total_Sales_Val': 'sum',
                'Total_Target_Val': 'sum'
            }).reset_index()
            cluster_acc['Ach %'] = (cluster_acc['Total_Sales_Val'] / cluster_acc['Total_Target_Val'] * 100).round(1)
            
            fig_cluster = px.bar(
                cluster_acc,
                x='Cluster',
                y='Total_Sales_Val',
                text='Ach %',
                title="Sales by Cluster (Label shows Achievement %)",
                color='Ach %',
                color_continuous_scale='Greens',
                labels={'Total_Sales_Val': 'Sales Value (EGP)'}
            )
            fig_cluster.update_traces(texttemplate='Ach: %{text}%', textposition='outside')
            st.plotly_chart(fig_cluster, use_container_width=True)

        with c2:
            dist_totals = pd.DataFrame([
                {'Distributor': d, 'Sales': filtered_df[d].sum()} for d in dist_names
            ])
            dist_totals = dist_totals[dist_totals['Sales'] > 0]
            
            if not dist_totals.empty:
                fig_dist_pie = px.pie(
                    dist_totals,
                    names='Distributor',
                    values='Sales',
                    title="Distributor Share Breakdown",
                    hole=0.4,
                    color_discrete_sequence=px.colors.qualitative.Set2
                )
                fig_dist_pie.update_traces(textinfo='percent+label')
                st.plotly_chart(fig_dist_pie, use_container_width=True)

        st.markdown("---")

        # ---------------------------------------------------------
        # 7. Overall Top & Bottom Performing Products
        # ---------------------------------------------------------
        st.subheader("🏆 Overall Top & Bottom Performing Products")
        p1, p2 = st.columns(2)
        
        with p1:
            top10_overall = filtered_df.nlargest(10, 'Total_Sales_Val')
            if not top10_overall.empty:
                fig_top10 = px.bar(
                    top10_overall,
                    x='Total_Sales_Val',
                    y='Product_Name',
                    orientation='h',
                    title="Top 10 Products (Sales Value)",
                    text_auto=',.0f',
                    color_discrete_sequence=['#3498db']
                )
                fig_top10.update_layout(yaxis={'categoryorder':'total ascending'})
                st.plotly_chart(fig_top10, use_container_width=True)

        with p2:
            bottom10_overall = filtered_df[filtered_df['Total_Target_Val'] > 0].nsmallest(10, 'Total_Sales_Val')
            if not bottom10_overall.empty:
                fig_bot10 = px.bar(
                    bottom10_overall,
                    x='Total_Sales_Val',
                    y='Product_Name',
                    orientation='h',
                    title="Lowest 10 Target-Assigned Products",
                    text_auto=',.0f',
                    color_discrete_sequence=['#e74c3c']
                )
                fig_bot10.update_layout(yaxis={'categoryorder':'total descending'})
                st.plotly_chart(fig_bot10, use_container_width=True)

        st.markdown("---")

        # ---------------------------------------------------------
        # 8. Detailed Data Table
        # ---------------------------------------------------------
        st.subheader("📋 Product Sales Summary Table")
        
        search_q = st.text_input("🔍 Search product name:")
        if search_q:
            table_df = filtered_df[filtered_df['Product_Name'].str.contains(search_q, case=False, na=False)]
        else:
            table_df = filtered_df

        st.dataframe(
            table_df[[
                'Cluster', 'Product_Name', 'Total_Sales_Val', 
                'Total_Target_Val', 'Private_Sales_Val', 'Tender_Sales_Val', 'Stock_Days'
            ]].sort_values(by='Total_Sales_Val', ascending=False),
            use_container_width=True
        )

else:
    st.info("👋 Please upload your `Daily Sales ALL.xlsx` file from the sidebar to view the dashboard.")
