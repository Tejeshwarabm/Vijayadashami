# Save this as streamlit_dashboard.py
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go


def create_streamlit_dashboard(file_path):
    """Create an interactive Streamlit dashboard"""

    # Set page config
    st.set_page_config(
        page_title="Vijayadashami 2025 Dashboard",
        page_icon="🎉",
        layout="wide",
        initial_sidebar_state="expanded"
    )

    # Read data
    @st.cache_data
    def load_data():
        excel_data = pd.read_excel(file_path, sheet_name=None, engine='openpyxl')
        df_summary = excel_data['Sheet3'].dropna(how='all')
        df_detailed = excel_data['Sheet8'].dropna(how='all')
        df_summary.columns = df_summary.columns.str.strip()
        df_detailed.columns = df_detailed.columns.str.strip()
        return df_summary, df_detailed

    try:
        df_summary, df_detailed = load_data()
    except Exception as e:
        st.error(f"Error loading data: {e}")
        return

    # Sidebar
    st.sidebar.title("🎛️ Dashboard Controls")
    selected_tab = st.sidebar.radio("Select View:",
                                    ["📊 Summary Overview", "📈 Detailed Analysis", "💡 Insights"])

    # Main content
    st.title("🎉 Vijayadashami 2025 Dashboard")

    if selected_tab == "📊 Summary Overview":
        st.header("Summary Overview - Sheet3 Data")

        # Key metrics
        col1, col2, col3, col4 = st.columns(4)

        with col1:
            total_attendance = df_summary['Grand Total'].sum()
            st.metric("Total Attendance", f"{total_attendance:,}")

        with col2:
            total_nagaras = len(df_summary)
            st.metric("Number of Nagaras", total_nagaras)

        with col3:
            avg_attendance = int(total_attendance / total_nagaras)
            st.metric("Average per Nagara", f"{avg_attendance:,}")

        with col4:
            total_ghosh = df_summary['Ghosh'].sum()
            st.metric("Total Ghosh Participants", total_ghosh)

        # Charts
        col1, col2 = st.columns(2)

        with col1:
            if 'Grand Total' in df_summary.columns:
                fig = px.bar(df_summary, x='Nagara', y='Grand Total',
                             title='Total Attendance by Nagara',
                             color='Grand Total')
                st.plotly_chart(fig, use_container_width=True)

        with col2:
            if 'Ghosh' in df_summary.columns:
                fig = px.bar(df_summary, x='Nagara', y='Ghosh',
                             title='Ghosh Participants by Nagara',
                             color='Ghosh')
                st.plotly_chart(fig, use_container_width=True)

        col3, col4 = st.columns(2)

        with col3:
            if 'Total Vasati' in df_summary.columns:
                fig = px.pie(df_summary, values='Total Vasati', names='Nagara',
                             title='Vasati Distribution')
                st.plotly_chart(fig, use_container_width=True)

        with col4:
            if all(col in df_summary.columns for col in ['Total Booths', 'Represented Booths']):
                df_summary['Representation_Rate'] = (
                            df_summary['Represented Booths'] / df_summary['Total Booths'] * 100).round(1)
                fig = px.bar(df_summary, x='Nagara', y='Representation_Rate',
                             title='Booth Representation Rate (%)')
                st.plotly_chart(fig, use_container_width=True)

    elif selected_tab == "📈 Detailed Analysis":
        st.header("Detailed Analysis - Sheet8 Data")

        # Filters
        col1, col2 = st.columns(2)

        with col1:
            selected_nagara = st.selectbox("Select Nagara:",
                                           ['All'] + df_detailed['Nagara'].unique().tolist())

        with col2:
            top_n = st.slider("Number of Top Vasatis:", 5, 20, 10)

        # Filter data
        if selected_nagara != 'All':
            filtered_data = df_detailed[df_detailed['Nagara'] == selected_nagara]
        else:
            filtered_data = df_detailed

        # Charts
        col1, col2 = st.columns(2)

        with col1:
            if 'Grand Total' in df_detailed.columns:
                top_vasatis = filtered_data.nlargest(top_n, 'Grand Total')
                fig = px.bar(top_vasatis, x='Vasati', y='Grand Total',
                             title=f'Top {top_n} Vasatis by Attendance')
                st.plotly_chart(fig, use_container_width=True)

        with col2:
            if 'Grade' in df_detailed.columns:
                grade_counts = filtered_data['Grade'].value_counts()
                fig = px.pie(values=grade_counts.values, names=grade_counts.index,
                             title='Grade Distribution')
                st.plotly_chart(fig, use_container_width=True)

        # Additional detailed charts
        if all(col in df_detailed.columns for col in ['Tarun', 'Balak', 'Women']):
            st.subheader("Participant Categories")
            sample_data = filtered_data.head(8)
            fig = go.Figure(data=[
                go.Bar(name='Tarun', x=sample_data['Vasati'], y=sample_data['Tarun']),
                go.Bar(name='Balak', x=sample_data['Vasati'], y=sample_data['Balak']),
                go.Bar(name='Women', x=sample_data['Vasati'], y=sample_data['Women'])
            ])
            fig.update_layout(barmode='stack', title='Participant Categories by Vasati')
            st.plotly_chart(fig, use_container_width=True)

    else:  # Insights tab
        st.header("💡 Key Insights & Analytics")

        # Performance analysis
        col1, col2 = st.columns(2)

        with col1:
            if 'Grand Total' in df_summary.columns:
                # Performance ranking
                df_sorted = df_summary.sort_values('Grand Total', ascending=False)
                fig = px.bar(df_sorted, x='Nagara', y='Grand Total',
                             title='Performance Ranking by Nagara',
                             color='Grand Total')
                st.plotly_chart(fig, use_container_width=True)

        with col2:
            # Efficiency analysis
            if all(col in df_summary.columns for col in ['Grand Total', 'Total Vasati']):
                df_summary['Efficiency'] = (df_summary['Grand Total'] / df_summary['Total Vasati']).round(1)
                fig = px.bar(df_summary, x='Nagara', y='Efficiency',
                             title='Efficiency (Attendance per Vasati)')
                st.plotly_chart(fig, use_container_width=True)

        # Data table
        st.subheader("Raw Data Preview")
        tab1, tab2 = st.tabs(["Summary Data", "Detailed Data"])

        with tab1:
            st.dataframe(df_summary, use_container_width=True)

        with tab2:
            st.dataframe(df_detailed, use_container_width=True)

# To run the Streamlit app, use: streamlit run streamlit_dashboard.py

# Main execution
if __name__ == "__main__":
    file_path = "Vijayadashami_VIJ_2025.xlsx"

    # Create complete dashboard solution
    create_streamlit_dashboard(file_path)