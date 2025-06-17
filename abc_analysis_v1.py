import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import io

st.set_page_config(page_title="ABC Analysis Dashboard", layout="wide")
st.title("📊 ABC Analysis Dashboard")

# Upload Excel file
uploaded_file = st.file_uploader("Upload your sales Excel file", type=["xlsx"])

if uploaded_file:
    try:
        # Read Excel file
        df_sales = pd.read_excel(uploaded_file)

        st.subheader("Raw Data Preview")
        st.dataframe(df_sales.head())

        # Clean and prepare data
        df_sales = df_sales.dropna(subset=['qty'])
        df_sales = df_sales.sort_values(by='description', ascending=True)
        df_sales.rename(columns={
            'ProductName': 'item_code',
            'LINeSales': 'value'
        }, inplace=True)

        # ABC Classification thresholds
        A, B, C = 80, 15, 5

        df1 = df_sales.groupby(['item_id'], as_index=False)['value'].sum()
        df1 = df1.sort_values(by='value', ascending=False).reset_index(drop=True)
        df1['perc'] = df1['value'] / df1['value'].sum() * 100
        df1['cumu'] = df1['perc'].cumsum()
        df1['abc'] = df1['cumu'].apply(lambda x: 'A' if x < A else ('B' if x < A + B else 'C'))

        # ABC Summary Table
        df_abc_summary = pd.pivot_table(df1, index='abc', values=['perc', 'value', 'item_id'],
                                        aggfunc={'perc': np.sum, 'value': np.sum, 'item_id': 'count'}).reset_index()
        df_abc_summary['item_perc'] = df_abc_summary['item_id'] / df_abc_summary['item_id'].sum() * 100

        st.subheader("ABC Summary Table")
        st.dataframe(df_abc_summary)

        # Excel download
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df1.to_excel(writer, index=False, sheet_name='ABC Table')
            df_abc_summary.to_excel(writer, index=False, sheet_name='Summary')
        st.download_button(
            label="📥 Download ABC Tables as Excel",
            data=output.getvalue(),
            file_name="abc_analysis.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # Pareto Chart
        st.subheader("Pareto Distribution")
        fig_pareto = px.bar(df1, x='item_id', y='cumu', color='abc', title='Pareto Distribution', template='simple_white')
        st.plotly_chart(fig_pareto, use_container_width=True)

        # ABC Distribution Chart
        st.subheader("ABC Distribution by Number of Items")
        fig_summary1 = px.bar(df_abc_summary, x='abc', y='item_perc', title='ABC distribution by number of items')
        st.plotly_chart(fig_summary1, use_container_width=True)

        # Tabs for ABC Categories
        tab_a, tab_b, tab_c = st.tabs(["Category A", "Category B", "Category C"])

        with tab_a:
            st.subheader("Category A Items")
            df_cat_a = df1[df1['abc'] == 'A']
            st.dataframe(df_cat_a)
            fig_cat_a = px.bar(df_cat_a, x='item_id', y='value', title='Value Distribution - Category A', template='simple_white')
            st.plotly_chart(fig_cat_a, use_container_width=True)

        with tab_b:
            st.subheader("Category B Items")
            df_cat_b = df1[df1['abc'] == 'B']
            st.dataframe(df_cat_b)
            fig_cat_b = px.bar(df_cat_b, x='item_id', y='value', title='Value Distribution - Category B', template='simple_white')
            st.plotly_chart(fig_cat_b, use_container_width=True)

        with tab_c:
            st.subheader("Category C Items")
            df_cat_c = df1[df1['abc'] == 'C']
            st.dataframe(df_cat_c)
            fig_cat_c = px.bar(df_cat_c, x='item_id', y='value', title='Value Distribution - Category C', template='simple_white')
            st.plotly_chart(fig_cat_c, use_container_width=True)

    except Exception as e:
        st.error(f"❌ Error processing file: {e}")

else:
    st.info("📂 Please upload an Excel file to begin.")

# Footer
st.markdown("---")
st.markdown(
    """
    <div style='text-align: center; font-size: 14px;'>
        <strong>ABC Analysis Dashboard – with compliments</strong><br>
        📧 Email: <a href='mailto:promotions@realanalytics101.co.za'>promotions@realanalytics101.co.za</a><br>
        🌐 Website: <a href='https://www.realanalytics101.co.za' target='_blank'>www.realanalytics101.co.za</a>
    </div>
    """,
    unsafe_allow_html=True
)
