"""
Telecom Customer Analytics Dashboard
------------------------------------
Installation:
pip install streamlit pandas plotly fpdf openpyxl

Run locally:
streamlit run app.py
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import io
import base64
from datetime import datetime
import numpy as np
try:
    from fpdf import FPDF
    FPDF_AVAILABLE = True
except ImportError:
    FPDF_AVAILABLE = False

# --- Page Configuration ---
st.set_page_config(
    page_title="Telecom Analytics Dashboard",
    page_icon="📶",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- CSS Injection for Corporate Telecom Theme ---
st.markdown("""
<style>
    /* Professional colors: Blue (#0066CC) and Green (#00CC66) */
    .stApp {
        background-color: #F8F9FA;
    }
    h1, h2, h3, h4 {
        color: #003366 !important;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    }
    .metric-card {
        background-color: #FFFFFF;
        border-radius: 10px;
        padding: 15px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.05);
        border-left: 5px solid #0066CC;
        margin-bottom: 1rem;
    }
    .metric-value {
        font-size: 26px;
        font-weight: 700;
        color: #003366;
    }
    .metric-label {
        font-size: 14px;
        color: #6C757D;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    /* Enhance buttons */
    .stButton>button {
        background-color: #0066CC;
        color: white;
        border-radius: 5px;
        border: none;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        transition: all 0.2s ease-in-out;
    }
    .stButton>button:hover {
        background-color: #004C99;
        box-shadow: 0 4px 8px rgba(0,0,0,0.15);
    }
    /* Better table fonts */
    [data-testid="stDataFrame"] {
        font-size: 14px;
    }
</style>
""", unsafe_allow_html=True)

@st.cache_data
def load_sample_data():
    """Generates realistic embedded sample data (around 60 rows)"""
    np.random.seed(42)  # For consistent sample generation
    months = ['Jan-2025', 'Feb-2025', 'Mar-2025', 'Apr-2025', 'May-2025', 'Jun-2025']
    regions = ['Lagos', 'Abuja', 'Kano', 'Port Harcourt', 'Enugu', 'Ibadan', 'Other']
    plans = ['Prepaid', 'Postpaid']
    complaints = ['Network', 'Billing', 'Coverage', 'Service', 'Device', 'None', 'Others']
    
    data = []
    for i in range(1001, 1061):
        region = np.random.choice(regions, p=[0.3, 0.2, 0.15, 0.15, 0.1, 0.05, 0.05])
        plan = np.random.choice(plans, p=[0.75, 0.25])
        month = np.random.choice(months)
        
        # Logic to make data realistic
        data_usage = round(np.random.uniform(1.0, 40.0) if plan == 'Prepaid' else np.random.uniform(20.0, 150.0), 2)
        voice = int(np.random.uniform(50, 400) if plan == 'Prepaid' else np.random.uniform(300, 1200))
        sms = int(np.random.uniform(10, 150))
        complaint = np.random.choice(complaints, p=[0.25, 0.15, 0.1, 0.1, 0.05, 0.3, 0.05])
        recharge = round(np.random.uniform(500, 3000) if plan == 'Prepaid' else np.random.uniform(5000, 25000), 2)
        
        data.append([i, region, plan, data_usage, voice, sms, complaint, recharge, month])
        
    df = pd.DataFrame(data, columns=[
        'Customer_ID', 'Region', 'Plan_Type', 'Data_Usage_GB', 
        'Voice_Minutes', 'SMS_Count', 'Complaint_Category', 
        'Last_Recharge_Amount_NGN', 'Report_Month'
    ])
    
    # Sort chronologically by converting the month strings to datetime
    # Use coerce to handle any edge cases where a format might be wrong, rather than crashing
    df['Month_DT'] = pd.to_datetime(df['Report_Month'], format='%b-%Y', errors='coerce')
    df = df.sort_values(by=['Month_DT']).drop(columns=['Month_DT'])
    
    return df

def get_pdf_report(df):
    """Generates a simple PDF report using FPDF"""
    if not FPDF_AVAILABLE:
        return None
        
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", 'B', 16)
    pdf.cell(200, 10, txt="Telecom Customer Analytics Report", ln=True, align='C')
    pdf.set_font("Arial", '', 12)
    pdf.cell(200, 10, txt=f"Generated date: {datetime.now().strftime('%Y-%m-%d %H:%M')}", ln=True, align='C')
    pdf.ln(10)
    
    pdf.set_font("Arial", 'B', 14)
    pdf.cell(200, 10, txt="Executive Summary KPIs", ln=True)
    pdf.set_font("Arial", '', 12)
    
    pdf.cell(200, 8, txt=f"Total Customers Filtered: {len(df)}", ln=True)
    if len(df) > 0:
        pdf.cell(200, 8, txt=f"Average Data Usage: {df['Data_Usage_GB'].mean():.2f} GB", ln=True)
        pdf.cell(200, 8, txt=f"Average Voice Minutes: {df['Voice_Minutes'].mean():.0f} min", ln=True)
        pdf.cell(200, 8, txt=f"Percentage Prepaid: {(df['Plan_Type'] == 'Prepaid').mean()*100:.1f}%", ln=True)
        
    return pdf.output(dest='S').encode('latin-1')

def main():
    # --- Header ---
    st.markdown("""
    <div style="display: flex; align-items: center; justify-content: space-between;">
        <h1 style="margin:0;">📶 Business Intelligence Operations</h1>
        <span style="color: #6C757D; font-weight: bold;">Executive Dashboard</span>
    </div>
    <p style="color: #6C757D; margin-bottom: 2rem;">Real-time interactive insights and performance analytics for telecommunication networks.</p>
    """, unsafe_allow_html=True)
    
    # --- Sidebar ---
    with st.sidebar:
        st.header("⚙️ Data Source")
        uploaded_file = st.file_uploader("Override with custom data (CSV/Excel)", type=["csv", "xlsx"])
        
        st.markdown("---")
        st.header("🔍 Global Filters")
        
        # Load data logic
        df = load_sample_data()
        
        if uploaded_file is not None:
            try:
                if uploaded_file.name.endswith('.csv'):
                    user_df = pd.read_csv(uploaded_file)
                else:
                    user_df = pd.read_excel(uploaded_file)
                required_cols = ['Customer_ID', 'Region', 'Plan_Type', 'Data_Usage_GB']
                if all(c in user_df.columns for c in required_cols):
                    df = user_df
                    st.success("Custom data loaded!")
                else:
                    st.error(f"Missing required columns. Expected: {required_cols}. Using sample data.")
            except Exception as e:
                st.error(f"Error reading file: {e}. Using sample data.")

        # Multi-select filters
        regions = st.multiselect("Region", options=df['Region'].unique(), default=df['Region'].unique())
        plans = st.multiselect("Plan Type", options=df['Plan_Type'].unique(), default=df['Plan_Type'].unique())
        complaints = st.multiselect("Complaint Category", options=df['Complaint_Category'].unique(), default=df['Complaint_Category'].unique())
        
        # Try to sort months correctly
        months_list = list(df['Report_Month'].unique())
        try:
            months_list.sort(key=lambda x: datetime.strptime(x, "%b-%Y"))
        except:
            pass
        selected_months = st.multiselect("Report Month", options=months_list, default=months_list)

    # --- Apply Filters ---
    filtered_df = df[
        (df['Region'].isin(regions)) & 
        (df['Plan_Type'].isin(plans)) & 
        (df['Complaint_Category'].isin(complaints)) &
        (df['Report_Month'].isin(selected_months))
    ]
    
    if filtered_df.empty:
        st.warning("No data points available for the current filter criteria. Please adjust your selections.")
        st.stop()
        
    # --- Compute KPIs ---
    total_cust = len(filtered_df)
    avg_data = filtered_df['Data_Usage_GB'].mean()
    prepaid_pct = (filtered_df['Plan_Type'] == 'Prepaid').mean() * 100
    avg_recharge = filtered_df['Last_Recharge_Amount_NGN'].mean()
    
    real_complaints = filtered_df[filtered_df['Complaint_Category'] != 'None']
    top_complaint = real_complaints['Complaint_Category'].mode()[0] if not real_complaints.empty else "N/A"

    # --- Responsive KPI Row ---
    kpi1, kpi2, kpi3, kpi4, kpi5 = st.columns(5)
    
    metrics = [
        (kpi1, "👥 Total Customers", f"{total_cust:,}"),
        (kpi2, "📊 Avg Data Usage", f"{avg_data:.1f} GB"),
        (kpi3, "📱 Prepaid Ratio", f"{prepaid_pct:.1f}%"),
        (kpi4, "⚠️ Top Complaint", top_complaint),
        (kpi5, "💳 Avg Recharge", f"₦{avg_recharge:,.0f}")
    ]
    
    for col, title, val in metrics:
        with col:
            st.markdown(f"""
            <div class="metric-card">
                <div class="metric-label">{title}</div>
                <div class="metric-value">{val}</div>
            </div>
            """, unsafe_allow_html=True)
            
    st.markdown("<br>", unsafe_allow_html=True)

    # --- Visualizations ---
    chart_config = {'displayModeBar': False} # Cleaner UI on mobile
    
    # Row 1
    col_chart1, col_chart2 = st.columns([1, 1])
    
    with col_chart1:
        st.markdown("**Average Data Usage by Region**")
        region_avg = filtered_df.groupby('Region')['Data_Usage_GB'].mean().reset_index().sort_values('Data_Usage_GB')
        fig1 = px.bar(region_avg, x='Data_Usage_GB', y='Region', orientation='h',
                     color_discrete_sequence=['#0066CC'])
        fig1.update_layout(margin=dict(l=0, r=0, t=20, b=0), plot_bgcolor='rgba(0,0,0,0)', 
                           xaxis_title="Avg Data (GB)", yaxis_title="")
        st.plotly_chart(fig1, use_container_width=True, config=chart_config)
        
    with col_chart2:
        st.markdown("**Complaint Category Distribution**")
        if not real_complaints.empty:
            fig2 = px.pie(real_complaints, names='Complaint_Category', hole=0.45,
                         color_discrete_sequence=px.colors.sequential.Tealgrn)
            fig2.update_traces(textposition='inside', textinfo='percent+label')
            fig2.update_layout(margin=dict(l=0, r=0, t=20, b=0), showlegend=False)
            st.plotly_chart(fig2, use_container_width=True, config=chart_config)
        else:
            st.info("No active complaints to display.")

    # Row 2
    col_chart3, col_chart4 = st.columns([1, 1])
    
    with col_chart3:
        st.markdown("**Incident Volumes by Region**")
        if not real_complaints.empty:
            top_cats = real_complaints['Complaint_Category'].value_counts().nlargest(3).index
            stacked = real_complaints[real_complaints['Complaint_Category'].isin(top_cats)]
            fig3 = px.histogram(stacked, x='Region', color='Complaint_Category',
                               barmode='group', color_discrete_sequence=['#0066CC', '#00CC66', '#80ADD6'])
            fig3.update_layout(margin=dict(l=0, r=0, t=20, b=0), plot_bgcolor='rgba(0,0,0,0)',
                              xaxis_title="", legend_title="")
            st.plotly_chart(fig3, use_container_width=True, config=chart_config)
        else:
            st.info("No incident data.")
            
    with col_chart4:
        st.markdown("**Monthly Usage Trends**")
        trend = filtered_df.groupby('Report_Month')[['Data_Usage_GB', 'Voice_Minutes']].mean().reset_index()
        try:
            trend['Month_DT'] = pd.to_datetime(trend['Report_Month'], format='%b-%Y')
            trend = trend.sort_values(by='Month_DT')
        except:
            pass
            
        fig4 = go.Figure()
        fig4.add_trace(go.Scatter(x=trend['Report_Month'], y=trend['Data_Usage_GB'], 
                      mode='lines+markers', name='Data (GB)', line=dict(color='#0066CC', width=3)))
        fig4.add_trace(go.Scatter(x=trend['Report_Month'], y=trend['Voice_Minutes'], 
                      mode='lines+markers', name='Voice (Min)', yaxis='y2', line=dict(color='#00CC66', width=3)))
        
        fig4.update_layout(
            margin=dict(l=0, r=0, t=20, b=0), plot_bgcolor='rgba(0,0,0,0)',
            yaxis=dict(title='Data (GB)', titlefont=dict(color='#0066CC'), tickfont=dict(color='#0066CC')),
            yaxis2=dict(title='Voice (Min)', titlefont=dict(color='#00CC66'), tickfont=dict(color='#00CC66'), overlaying='y', side='right'),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
        )
        st.plotly_chart(fig4, use_container_width=True, config=chart_config)

    # --- Data & AI Recommendations ---
    st.markdown("---")
    res_col1, res_col2 = st.columns([1.5, 1])
    
    with res_col1:
        st.markdown("### 🗃️ Raw Data Explorer")
        st.dataframe(filtered_df, use_container_width=True, height=280)
        
    with res_col2:
        st.markdown("### 💡 AI Recommendations")
        recs = []
        if top_complaint in ['Network', 'Coverage']:
            recs.append(f"🛠️ **Infrastructure Action**: High frequency of '{top_complaint}' incidents indicates immediate cell-site optimization is required.")
        if prepaid_pct > 80:
            recs.append("🎯 **Marketing Focus**: Dominant prepaid user base detected. Introduce micro-data bundles to stimulate daily recharges.")
        elif prepaid_pct < 40:
            recs.append("💼 **Retention Strategy**: Postpaid-heavy segment. Prioritize white-glove customer service and loyalty tiering.")
            
        if avg_data < 10:
            recs.append("📺 **Revenue Growth**: Average data usage is critically low. Cross-sell streaming or OTT service bundles.")
            
        if not real_complaints.empty:
            worst_region = real_complaints['Region'].value_counts().idxmax()
            recs.append(f"📍 **Regional Alert**: Deploy rapid response engineering units to **{worst_region}** due to peak localized complaint volumes.")
            
        # Guarantee 4 recommendations
        while len(recs) < 4:
            recs.append("🔄 **Operations**: Automate proactive SMS alerts during regional network downtimes to reduce support call volume.")
            recs.append("📱 **App Engagement**: Boost user self-service adoption to alleviate basic billing complaints.")
            break
            
        for r in recs[:5]:
            st.info(r)

    # --- Data Export ---
    st.markdown("---")
    st.markdown("### 📥 Export Capabilities")
    exp_col1, exp_col2, exp_col3 = st.columns(3)
    
    with exp_col1:
        csv = filtered_df.to_csv(index=False).encode('utf-8')
        st.download_button("Download CSV Data", data=csv, file_name="telecom_export.csv", mime="text/csv", use_container_width=True)
        
    with exp_col2:
        out_excel = io.BytesIO()
        with pd.ExcelWriter(out_excel, engine='openpyxl') as writer:
            filtered_df.to_excel(writer, index=False, sheet_name='Data')
        b_excel = out_excel.getvalue()
        st.download_button("Download Excel Report", data=b_excel, file_name="telecom_export.xlsx", 
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                           
    with exp_col3:
        if FPDF_AVAILABLE:
            pdf_data = get_pdf_report(filtered_df)
            st.download_button("Download Executive PDF", data=pdf_data, file_name="telecom_executive.pdf", 
                               mime="application/pdf", use_container_width=True)
        else:
            st.button("Download Executive PDF", disabled=True, help="fpdf library is required for this feature. Install via pip install fpdf", use_container_width=True)

    # --- Footer ---
    st.markdown("<br><br>", unsafe_allow_html=True)
    st.markdown(f"<div style='text-align: center; color: #ADB5BD; font-size: 13px;'>Telecom Customer Insights • © {datetime.now().strftime('%Y')} • Designed for Executive Operations</div>", unsafe_allow_html=True)

if __name__ == "__main__":
    main()
