import streamlit as st
import pandas as pd
import plotly.express as px

# Page config for wide layout
st.set_page_config(layout="wide")

# Read the Excel file, focusing only on Sheet8 and Sheet3
file_path = 'Vijayadashami_VIJ_2025.xlsx'


@st.cache_data
def load_sheets():
    sheets = pd.read_excel(file_path, sheet_name=['Sheet8', 'Sheet3'])
    return sheets


sheets = load_sheets()

# Streamlit app title
st.title('Vijayadashami 2025 (VIJAYNAGARA BHAGA)')

# Create main tabs for each sheet
tab8, tab3 = st.tabs(['Vijayadashami Utsava', 'Patha Sanchalana'])

with tab8:
    df8 = sheets['Sheet8']

    # Summary: Group by Nagara for Grand Total
    df8_summary = df8.groupby('Nagara')['Grand Total'].sum().reset_index()
    # Sort in ascending order by Grand Total
    df8_summary = df8_summary.sort_values('Grand Total', ascending=True).reset_index(drop=True)

    # Plot: Interactive bar chart with Plotly
    fig8 = px.bar(df8_summary,
                  x='Nagara',
                  y='Grand Total',
                  title='Grand Total per Nagara (Utsava)',
                  labels={'Nagara': 'Nagara', 'Grand Total': 'Total'},
                  color='Grand Total',
                  color_continuous_scale='Viridis')

    # Increase y-axis by 5% padding
    y_max = df8_summary['Grand Total'].max()
    y_padding = y_max * 0.25
    fig8.update_layout(yaxis=dict(range=[0, y_max + y_padding]), height=500, showlegend=False)
    # Customize x-axis: Increase font size and make bold
    fig8.update_xaxes(tickfont=dict(size=14, family="Arial Black, sans-serif"))
    # Add total count on each bar
    fig8.update_traces(texttemplate='%{y}', textposition='outside', textfont=dict(size=16))

    # Plot 2: Top 5 Grand Total Attendance
    df8_top5 = df8.nlargest(5, 'Grand Total')[['Nagara', 'Vasati', 'Grand Total']]
    fig8_top5 = px.bar(df8_top5,
                       x='Vasati',
                       y='Grand Total',
                       title='Top 5 (Utsava)',
                       labels={'Vasati': 'Vasati', 'Nagara': 'Nagara', 'Grand Total': 'Total'},
                       color='Grand Total',
                       color_continuous_scale='Reds')

    # Increase y-axis by 5% padding
    y_max_top5 = df8_top5['Grand Total'].max()
    y_padding_top5 = y_max_top5 * 0.25
    fig8_top5.update_layout(yaxis=dict(range=[0, y_max_top5 + y_padding_top5]), height=400, showlegend=False)
    # Customize x-axis: Increase font size and make bold
    fig8_top5.update_xaxes(tickfont=dict(size=14, family="Arial Black, sans-serif"))
    # Add total count on each bar with increased font size
    fig8_top5.update_traces(texttemplate='%{y}', textposition='outside', textfont=dict(size=16))


    # Interactive Selection: Select Nagara and show Vasati details
    unique_nagara = sorted(df8['Nagara'].unique())
    selected_nagara = st.selectbox(
        'Select a Nagara to view Vasati Details:',
        options=unique_nagara,
        index=0
    )

    # Filter data for selected Nagara
    df_selected = df8[df8['Nagara'] == selected_nagara]

    # Plot 3: Grand Total per Vasati for selected Nagara
    if not df_selected.empty:
        fig8_vasati = px.bar(df_selected,
                             x='Vasati',
                             y='Grand Total',
                             title=f'Grand Total per Vasati in {selected_nagara}',
                             labels={'Vasati': 'Vasati', 'Grand Total': 'Total'},
                             color='Grand Total',
                             color_continuous_scale='Greens')
        # Increase y-axis by 5% padding
        y_max_vasati = df_selected['Grand Total'].max()
        y_padding_vasati = y_max_vasati * 0.25
        fig8_vasati.update_layout(yaxis=dict(range=[0, y_max_vasati + y_padding_vasati]),
                                  height=500, showlegend=False)
        # Customize x-axis: Increase font size and make bold
        fig8_vasati.update_xaxes(tickfont=dict(size=14, family="Arial Black, sans-serif"))
        # Add total count on each bar with increased font size
        fig8_vasati.update_traces(texttemplate='%{y}', textposition='outside', textfont=dict(size=16))
    else:
        fig8_vasati = None
        st.warning(f"No data found for {selected_nagara}")


    # Sub-tabs for plot and table
    subtab_plot, subtab_data = st.tabs(['Summary Plot', 'Detailed Data'])

    with subtab_plot:
        if fig8_vasati:
            st.plotly_chart(fig8_vasati)
        # st.dataframe(df_selected)  # Show filtered table below the plot

        st.plotly_chart(fig8)
        st.plotly_chart(fig8_top5)


    with subtab_data:
        st.subheader('Vijayadashami Utsava')
        st.dataframe(df8)

with tab3:
    df3 = sheets['Sheet3']
    # df3 = df3.dropna()
    # Clean: Remove first row (header) and strip strings
    df3 = df3.iloc[:].reset_index(drop=True)
    df3['Nagara'] = df3['Nagara'].str.strip()

    # Ensure 'Grand Total' column is numeric (handle any non-numeric)
    df3['Grand Total'] = pd.to_numeric(df3['Grand Total'], errors='coerce').fillna(0)

    # Sort in ascending order by Grand Total
    df3 = df3.sort_values('Grand Total', ascending=True).reset_index(drop=True)

    # Plot: Interactive bar chart with Plotly
    fig3 = px.bar(df3,
                  x='Nagara',
                  y='Grand Total',
                  title='Grand Total per Nagara (Patha Sanchalana)',
                  labels={'Nagara': 'Nagara', 'Grand Total': 'Grand Total'},
                  color='Grand Total',
                  color_continuous_scale='Viridis')

    # Increase y-axis by 5% padding
    y_max3 = df3['Grand Total'].max()
    y_padding3 = y_max3 * 0.25
    fig3.update_layout(yaxis=dict(range=[0, y_max3 + y_padding3]), height=800, showlegend=False)
    # Customize x-axis: Increase font size and make bold (for consistency)
    fig3.update_xaxes(tickfont=dict(size=14, family="Arial Black, sans-serif"))
    # Add total count on each bar with increased font size
    fig3.update_traces(texttemplate='%{y}', textposition='outside', textfont=dict(size=16))


    # Sub-tabs for plot and table
    subtab_plot, subtab_data = st.tabs(['Summary Plot', 'Detailed Data'])

    with subtab_plot:
        st.plotly_chart(fig3)

    with subtab_data:
        st.subheader('Patha Sanchalana')
        st.dataframe(df3)

# Footer
st.markdown('---')
st.caption('RSS@100 | Sangha Shatabdi | Vijayanagara Bhaga | Bengaluru Dakshina')