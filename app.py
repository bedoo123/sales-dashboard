import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os
import io
import base64

st.set_page_config(
    page_title="Sales & Distribution Dashboard",
    page_icon="📊",
    layout="wide"
)

st.markdown("""
    <style>
    div[data-testid="stSidebarNav"] {display: none;}
    
    div[data-testid="stMetric"] {
        background-color: rgba(255, 255, 255, 0.05);
        border: 1px solid rgba(255, 255, 255, 0.1);
        padding: 15px;
        border-radius: 10px;
        border-left: 5px solid #007bff;
    }
    
    .stButton>button {
        border-radius: 8px;
        font-weight: bold;
        transition: all 0.3s ease;
    }
    .stButton>button:hover {
        border-color: #007bff;
        color: #007bff;
    }
    </style>
""", unsafe_allow_html=True)

# قائمة المستخدمين والصلاحيات
USERS = {
    "admin": {"password": "admin", "role": "admin", "line": "All"},
    "chc1": {"password": "123", "role": "user", "line": "Mash CHC 1"},
    "chc2": {"password": "123", "role": "user", "line": "Mash CHC 2"},
    "cns1": {"password": "123", "role": "user", "line": "Mash CNS 1"},
    "cns2": {"password": "123", "role": "user", "line": "Mash CNS 2"},
    "cns3": {"password": "123", "role": "user", "line": "Mash CNS 3"},
    "cns4": {"password": "123", "role": "user", "line": "Mash CNS 4"},
    "git1": {"password": "123", "role": "user", "line": "Mash GIT 1"},
    "git2": {"password": "123", "role": "user", "line": "Mash GIT 2"},
    "git3": {"password": "123", "role": "user", "line": "Mash GIT 3"},
    "cvm1": {"password": "123", "role": "user", "line": "Mash CVM 1"},
    "cvm2": {"password": "123", "role": "user", "line": "Mash CVM 2 New"},
    "ortho": {"password": "123", "role": "user", "line": "Mash Ortho"},
    "gyna": {"password": "123", "role": "user", "line": "Mash Gyna"},
    "dnu": {"password": "123", "role": "user", "line": "Mash DNU"},
    "oxa": {"password": "123", "role": "user", "line": "Mash Oxaleptal-Atten"},
    "derma": {"password": "123", "role": "user", "line": "Mash DERMA"},
    "primary1": {"password": "123", "role": "user", "line": "Mash Primary Care-1"},
    "primary2": {"password": "123", "role": "user", "line": "Mash Primary Care -2"}
}

if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.username = ""
    st.session_state.role = ""
    st.session_state.assigned_line = ""

def get_image_base64(path):
    if os.path.exists(path):
        with open(path, "rb") as image_file:
            return base64.b64encode(image_file.read()).decode()
    return ""

def login_page():
    _, col2, _ = st.columns([1, 1.2, 1])
    with col2:
        st.markdown("<div style='height: 20px;'></div>", unsafe_allow_html=True)
        
        logo_b64 = get_image_base64("logo.png")
        logo_html = f'''
            <div style="text-align: center; margin-bottom: 15px;">
                <img src="data:image/png;base64,{logo_b64}" width="110" style="border-radius: 8px; box-shadow: 0 4px 10px rgba(0,0,0,0.3);" />
            </div>
        ''' if logo_b64 else ''
        
        # Clean unified login card container
        st.markdown(f'''
            <div style="
                background: linear-gradient(145deg, rgba(30, 30, 35, 0.8), rgba(20, 20, 25, 0.95));
                padding: 30px;
                border-radius: 15px;
                border: 1px solid rgba(255, 255, 255, 0.1);
                box-shadow: 0 8px 32px 0 rgba(0, 0, 0, 0.5);
            ">
                {logo_html}
                <h3 style='color: #ffffff; text-align: center; font-size: 20px; margin-bottom: 5px; font-weight: 600;'>Welcome Back</h3>
                <p style='color: #a0a0a0; text-align: center; font-size: 13px; margin-bottom: 20px;'>Sales & Distribution Executive Dashboard</p>
        ''', unsafe_allow_html=True)
        
        with st.form("login_form"):
            username = st.text_input("👤 Username", placeholder="Enter your username")
            password = st.text_input("🔑 Password", type="password", placeholder="Enter your password")
            st.markdown("<div style='margin-top: 10px;'></div>", unsafe_allow_html=True)
            submit = st.form_submit_button("Sign In", use_container_width=True)
            
            if submit:
                if username in USERS and USERS[username]["password"] == password:
                    st.session_state.logged_in = True
                    st.session_state.username = username
                    st.session_state.role = USERS[username]["role"]
                    st.session_state.assigned_line = USERS[username]["line"]
                    st.success("تم تسجيل الدخول بنجاح!")
                    st.rerun()
                else:
                    st.error("اسم المستخدم أو كلمة المرور غير صحيحة.")
                    
        st.markdown("</div>", unsafe_allow_html=True)

if not st.session_state.logged_in:
    login_page()
else:
    # Sidebar Logo & User Info
    if os.path.exists("logo.png"):
        sb_c1, sb_c2, sb_c3 = st.sidebar.columns([1, 2, 1])
        with sb_c2:
            st.sidebar.image("logo.png", width=120)
            
    st.sidebar.markdown("<br>", unsafe_allow_html=True)
    st.sidebar.title(f"👤 Welcome, {st.session_state.username}")
    if st.sidebar.button("Logout", use_container_width=True):
        st.session_state.logged_in = False
        st.rerun()

    st.title("📊 Sales & Distribution Executive Dashboard")
    st.markdown("---")

    def format_currency(val):
        if abs(val) >= 1_000_000:
            return f"{val / 1_000_000:.2f}M EGP"
        elif abs(val) >= 1_000:
            return f"{val:,.0f} EGP"
        else:
            return f"{val:,.0f} EGP"

    @st.cache_data
    def load_and_clean_data(file_path):
        raw_df = pd.read_excel(file_path, header=None)
        df = raw_df.iloc[3:].copy()
        
        data = pd.DataFrame()
        data['Cluster'] = df.iloc[:, 0].astype(str).str.strip()
        data['Product_ID'] = df.iloc[:, 1]
        data['Product_Name'] = df.iloc[:, 2].astype(str).str.strip()
        
        data = data[data['Product_Name'].notna()]
        data = data[~data['Product_Name'].str.lower().str.contains("crm product name|total|sum|nan")]
        
        data['Private_Sales_Qty'] = pd.to_numeric(df.iloc[:, 3], errors='coerce').fillna(0)
        data['Private_Sales_Val'] = pd.to_numeric(df.iloc[:, 4], errors='coerce').fillna(0)
        data['Private_Target_Qty'] = pd.to_numeric(df.iloc[:, 5], errors='coerce').fillna(0)
        data['Private_Target_Val'] = pd.to_numeric(df.iloc[:, 6], errors='coerce').fillna(0)
        
        data['Ibn Sina'] = pd.to_numeric(df.iloc[:, 9], errors='coerce').fillna(0)
        data['Pharma'] = pd.to_numeric(df.iloc[:, 11], errors='coerce').fillna(0)
        data['Ramco'] = pd.to_numeric(df.iloc[:, 13], errors='coerce').fillna(0)
        data['UCP'] = pd.to_numeric(df.iloc[:, 15], errors='coerce').fillna(0)
        data['MEC'] = pd.to_numeric(df.iloc[:, 17], errors='coerce').fillna(0)
        data['Sofico'] = pd.to_numeric(df.iloc[:, 19], errors='coerce').fillna(0)
        data['ETC'] = pd.to_numeric(df.iloc[:, 21], errors='coerce').fillna(0)
        data['Direct'] = pd.to_numeric(df.iloc[:, 23], errors='coerce').fillna(0)
        data['Stores'] = pd.to_numeric(df.iloc[:, 25], errors='coerce').fillna(0)
        
        data['Stock_Days'] = pd.to_numeric(df.iloc[:, 42], errors='coerce').fillna(0)
        data['Total_Stock_Qty'] = pd.to_numeric(df.iloc[:, 45], errors='coerce').fillna(0)
        data['Total_Stock_Val'] = pd.to_numeric(df.iloc[:, 46], errors='coerce').fillna(0)
        
        data['Tender_Sales_Qty'] = pd.to_numeric(df.iloc[:, 58], errors='coerce').fillna(0)
        data['Tender_Sales_Val'] = pd.to_numeric(df.iloc[:, 59], errors='coerce').fillna(0)
        data['Total_Sales_Qty'] = pd.to_numeric(df.iloc[:, 60], errors='coerce').fillna(0)
        data['Total_Sales_Val'] = pd.to_numeric(df.iloc[:, 61], errors='coerce').fillna(0)
        data['Total_Target_Qty'] = pd.to_numeric(df.iloc[:, 62], errors='coerce').fillna(0)
        data['Total_Target_Val'] = pd.to_numeric(df.iloc[:, 63], errors='coerce').fillna(0)
        
        return data

    st.sidebar.header("📁 Control Panel")
    
    saved_file_path = "shared_daily_sales.xlsx"

    if st.session_state.role == "admin":
        uploaded_file = st.sidebar.file_uploader("Upload Excel File (Admin Only)", type=["xlsx", "xls"])
        if uploaded_file is not None:
            with open(saved_file_path, "wb") as f:
                f.write(uploaded_file.getbuffer())
            st.sidebar.success("تم رفع وحفظ الشيت بنجاح!")
    else:
        st.sidebar.info(f"👤 Line User: {st.session_state.assigned_line}")

    target_file = None
    if os.path.exists(saved_file_path):
        target_file = saved_file_path
    elif os.path.exists("Daily Sales 15-10-2026- ALL.xlsx"):
        target_file = "Daily Sales 15-10-2026- ALL.xlsx"

    if target_file is not None:
        df = load_and_clean_data(target_file)
        
        if st.session_state.role == "admin":
            st.sidebar.subheader("Filter Options (Admin)")
            selected_clusters = st.sidebar.multiselect(
                "Select Cluster / Line:", 
                options=df['Cluster'].unique(), 
                default=df['Cluster'].unique()
            )
            filtered_df = df[df['Cluster'].isin(selected_clusters)]
            
            st.sidebar.markdown("---")
            with open(saved_file_path, "rb") as f:
                st.sidebar.download_button(
                    label="📥 تحميل الشيت الكامل (Admin)",
                    data=f,
                    file_name="Daily_Sales_All.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
        else:
            assigned_line = st.session_state.assigned_line
            filtered_df = df[df['Cluster'] == assigned_line]
            st.sidebar.warning(f"Viewing Line: {assigned_line}")
            
            if not filtered_df.empty:
                st.sidebar.markdown("---")
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    filtered_df.to_excel(writer, index=False, sheet_name="Sales_Data")
                buffer.seek(0)
                st.sidebar.download_button(
                    label=f"📥 تحميل شيت خط {assigned_line}",
                    data=buffer,
                    file_name=f"{assigned_line.replace(' ', '_')}_sales.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

        if filtered_df.empty:
            st.warning("⚠️ لا توجد بيانات مطابقة للخط الخاص بك في هذا الشيت. تأكد من تطابق اسم الخط (Cluster) مع اسم المستخدم.")
        else:
            st.subheader("📌 Performance Summary")
            
            total_sales = filtered_df['Total_Sales_Val'].sum()
            total_target = filtered_df['Total_Target_Val'].sum()
            total_ach = (total_sales / total_target * 100) if total_target > 0 else 0
            private_sales = filtered_df['Private_Sales_Val'].sum()
            tender_sales = filtered_df['Tender_Sales_Val'].sum()
            
            zero_target_count = (filtered_df['Total_Target_Val'] == 0).sum()
            total_products_count = len(filtered_df)
            
            k1, k2, k3, k4, k5, k6 = st.columns(6)
            k1.metric("Total Sales", format_currency(total_sales))
            k2.metric("Target", format_currency(total_target))
            k3.metric("Achievement %", f"{total_ach:.1f}%")
            k4.metric("Private Sales", format_currency(private_sales))
            k5.metric("Tender Sales", format_currency(tender_sales))
            k6.metric("Zero Target SKUs", f"{zero_target_count} / {total_products_count}")
            
            st.progress(min(max(int(total_ach), 0), 100), text=f"Overall Target Achievement: {total_ach:.1f}%")
            st.markdown("---")

            st.subheader("⚠️ Products Without Target (Zero Target Analysis)")
            no_target_df = filtered_df[filtered_df['Total_Target_Val'] == 0]
            
            if not no_target_df.empty:
                no_target_with_sales = no_target_df[no_target_df['Total_Sales_Val'] > 0]
                no_target_no_sales = no_target_df[no_target_df['Total_Sales_Val'] == 0]
                
                nt_col1, nt_col2 = st.columns(2)
                
                with nt_col1:
                    st.markdown(f"##### 🟢 No Target BUT Has Sales ({len(no_target_with_sales)} SKUs)")
                    if not no_target_with_sales.empty:
                        fig_nt_sales = px.bar(
                            no_target_with_sales.nlargest(10, 'Total_Sales_Val'),
                            x='Total_Sales_Val',
                            y='Product_Name',
                            orientation='h',
                            title="Top 10 Zero-Target Products by Sales",
                            text_auto=',.0f',
                            color_discrete_sequence=['#2ecc71']
                        )
                        fig_nt_sales.update_layout(yaxis={'categoryorder':'total ascending'})
                        st.plotly_chart(fig_nt_sales, use_container_width=True)
                    else:
                        st.info("No sales recorded for any zero-target product.")

                with nt_col2:
                    st.markdown(f"##### 🔴 No Target AND No Sales ({len(no_target_no_sales)} SKUs)")
                    st.dataframe(
                        no_target_no_sales[['Cluster', 'Product_Name', 'Total_Stock_Qty', 'Stock_Days']], 
                        use_container_width=True, 
                        height=300
                    )

                with st.expander("📄 View All Products Without Target"):
                    st.dataframe(
                        no_target_df[['Cluster', 'Product_ID', 'Product_Name', 'Total_Sales_Val', 'Total_Stock_Qty', 'Stock_Days']].sort_values(by='Total_Sales_Val', ascending=False),
                        use_container_width=True
                    )
            else:
                st.success("🎉 All selected products have an assigned target!")

            st.markdown("---")

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

            st.subheader("📈 Cluster Performance & Market Share")
            c1, c2 = st.columns(2)
            
            with c1:
                cluster_acc = filtered_df.groupby('Cluster').agg({
                    'Total_Sales_Val': 'sum',
                    'Total_Target_Val': 'sum'
                }).reset_index()
                cluster_acc['Dict_Ach'] = cluster_acc['Total_Sales_Val'] / cluster_acc['Total_Target_Val'] * 100
                cluster_acc['Ach %'] = cluster_acc['Dict_Ach'].round(1)
                
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
                    fig_dist_pie.update_traces(
                        textinfo='percent+label',
                        texttemplate='%{label}<br>%{percent:.1%}',
                        insidetextorientation='radial'
                    )
                    st.plotly_chart(fig_dist_pie, use_container_width=True)

            st.markdown("---")

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

            st.subheader("📋 Product Sales Summary Table")
            search_q = st.text_input("🔍 Search product name:")
            if search_q:
                table_dir = filtered_df[filtered_df['Product_Name'].str.contains(search_q, case=False, na=False)]
            else:
                table_dir = filtered_df

            disp_df = table_dir[[
                'Cluster', 'Product_Name', 'Total_Sales_Val', 
                'Total_Target_Val', 'Private_Sales_Val', 'Tender_Sales_Val', 'Stock_Days'
            ]].copy().sort_values(by='Total_Sales_Val', ascending=False)
            
            for col in ['Total_Sales_Val', 'Total_Target_Val', 'Private_Sales_Val', 'Tender_Sales_Val', 'Stock_Days']:
                if col in disp_df.columns:
                    disp_df[col] = disp_df[col].apply(lambda x: f"{x:,.2f}" if pd.notnull(x) else "0.00")

            st.dataframe(disp_df, use_container_width=True)
    else:
        st.info("👋 يُرجى من الأدمن رفع ملف الشيت أولاً من لوحة التحكم ليتمكن باقي المستخدمين من استعراض البيانات.")
