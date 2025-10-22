import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import plotly.io as pio


def create_plotly_tabs_dashboard(file_path):
    """Create a Plotly dashboard with proper tabs for both sheets"""

    # Read Excel file
    try:
        excel_data = pd.read_excel(file_path, sheet_name=None, engine='openpyxl')
        print("Available sheets:", list(excel_data.keys()))
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return None

    # Process Sheet3 - Summary Data
    df_summary = excel_data['Sheet3'].copy()
    df_summary = df_summary.dropna(how='all')
    df_summary.columns = df_summary.columns.str.strip()

    # Process Sheet8 - Detailed Data
    df_detailed = excel_data['Sheet8'].copy()
    df_detailed = df_detailed.dropna(how='all')
    df_detailed.columns = df_detailed.columns.str.strip()

    # Create HTML with tabs using Plotly figures
    html_content = """
    <!DOCTYPE html>
    <html>
    <head>
        <title>Vijayadashami 2025 Dashboard</title>
        <script src="https://cdn.plot.ly/plotly-latest.min.js"></script>
        <style>
            body {
                font-family: Arial, sans-serif;
                margin: 20px;
                background-color: #f5f5f5;
            }
            .dashboard-container {
                max-width: 1400px;
                margin: 0 auto;
                background: white;
                padding: 20px;
                border-radius: 10px;
                box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            }
            .tabs {
                display: flex;
                margin-bottom: 20px;
                border-bottom: 2px solid #ddd;
            }
            .tab-button {
                padding: 12px 24px;
                background: #f1f1f1;
                border: none;
                cursor: pointer;
                font-size: 16px;
                margin-right: 5px;
                border-radius: 5px 5px 0 0;
                transition: background 0.3s;
            }
            .tab-button:hover {
                background: #ddd;
            }
            .tab-button.active {
                background: #4CAF50;
                color: white;
            }
            .tab-content {
                display: none;
                padding: 20px 0;
            }
            .tab-content.active {
                display: block;
            }
            .chart-grid {
                display: grid;
                grid-template-columns: repeat(2, 1fr);
                gap: 20px;
                margin-bottom: 20px;
            }
            .chart-container {
                background: white;
                padding: 15px;
                border-radius: 8px;
                box-shadow: 0 1px 5px rgba(0,0,0,0.1);
            }
            h1 {
                color: #2c3e50;
                text-align: center;
                margin-bottom: 30px;
            }
            h2 {
                color: #34495e;
                border-bottom: 2px solid #3498db;
                padding-bottom: 10px;
            }
            .summary-stats {
                display: grid;
                grid-template-columns: repeat(4, 1fr);
                gap: 15px;
                margin-bottom: 30px;
            }
            .stat-card {
                background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                color: white;
                padding: 20px;
                border-radius: 8px;
                text-align: center;
            }
            .stat-number {
                font-size: 24px;
                font-weight: bold;
                margin: 10px 0;
            }
            .stat-label {
                font-size: 14px;
                opacity: 0.9;
            }
        </style>
    </head>
    <body>
        <div class="dashboard-container">
            <h1>🎉 Vijayadashami 2025 Dashboard</h1>

            <div class="tabs">
                <button class="tab-button active" onclick="openTab(event, 'summary')">Summary Overview</button>
                <button class="tab-button" onclick="openTab(event, 'detailed')">Detailed Analysis</button>
                <button class="tab-button" onclick="openTab(event, 'insights')">Key Insights</button>
            </div>

            <!-- Summary Tab -->
            <div id="summary" class="tab-content active">
                <h2>📊 Summary Overview - Pathasanchalana</h2>

                <div class="summary-stats" id="summary-stats">
                    <!-- Stats will be populated by JavaScript -->
                </div>

                <div class="chart-grid">
                    <div class="chart-container">
                        <div id="attendance-chart"></div>
                    </div>
                    <div class="chart-container">
                        <div id="ghosh-chart"></div>
                    </div>
                    <div class="chart-container">
                        <div id="vasati-chart"></div>
                    </div>
                    <div class="chart-container">
                        <div id="booth-chart"></div>
                    </div>
                </div>
            </div>

            <!-- Detailed Tab -->
            <div id="detailed" class="tab-content">
                <h2>📈 Detailed Analysis - Utsava</h2>

                <div class="chart-grid">
                    <div class="chart-container">
                        <div id="top-vasatis-chart"></div>
                    </div>
                    <div class="chart-container">
                        <div id="categories-chart"></div>
                    </div>
                    <div class="chart-container">
                        <div id="grade-chart"></div>
                    </div>
                    <div class="chart-container">
                        <div id="booth-detailed-chart"></div>
                    </div>
                </div>
            </div>

            <!-- Insights Tab -->
            <div id="insights" class="tab-content">
                <h2>💡 Key Insights & Analytics</h2>

                <div class="chart-grid">
                    <div class="chart-container">
                        <div id="performance-chart"></div>
                    </div>
                    <div class="chart-container">
                        <div id="comparison-chart"></div>
                    </div>
                </div>

                <div class="chart-container" style="grid-column: span 2;">
                    <div id="trend-chart"></div>
                </div>
            </div>
        </div>

        <script>
            function openTab(evt, tabName) {
                var i, tabcontent, tabbuttons;
                tabcontent = document.getElementsByClassName("tab-content");
                for (i = 0; i < tabcontent.length; i++) {
                    tabcontent[i].classList.remove("active");
                }
                tabbuttons = document.getElementsByClassName("tab-button");
                for (i = 0; i < tabbuttons.length; i++) {
                    tabbuttons[i].classList.remove("active");
                }
                document.getElementById(tabName).classList.add("active");
                evt.currentTarget.classList.add("active");
            }
        </script>
    </body>
    </html>
    """

    # Create Plotly charts for Sheet3 (Summary)
    summary_charts_js = ""

    # 1. Attendance Chart
    if 'Grand Total' in df_summary.columns and 'Nagara' in df_summary.columns:
        fig1 = px.bar(df_summary, x='Nagara', y='Grand Total',
                      title='Total Attendance by Nagara',
                      color='Grand Total',
                      color_continuous_scale='viridis')
        summary_charts_js += f"Plotly.newPlot('attendance-chart', {fig1.to_json()});\n"

    # 2. Ghosh Chart
    if 'Ghosh' in df_summary.columns and 'Nagara' in df_summary.columns:
        fig2 = px.bar(df_summary, x='Nagara', y='Ghosh',
                      title='Ghosh Participants by Nagara',
                      color='Ghosh',
                      color_continuous_scale='oranges')
        summary_charts_js += f"Plotly.newPlot('ghosh-chart', {fig2.to_json()});\n"

    # 3. Vasati Chart
    if 'Total Vasati' in df_summary.columns and 'Nagara' in df_summary.columns:
        fig3 = px.pie(df_summary, values='Total Vasati', names='Nagara',
                      title='Vasati Distribution by Nagara')
        summary_charts_js += f"Plotly.newPlot('vasati-chart', {fig3.to_json()});\n"

    # 4. Booth Chart
    if all(col in df_summary.columns for col in ['Total Booths', 'Represented Booths']):
        df_summary['Representation_Rate'] = (df_summary['Represented Booths'] / df_summary['Total Booths'] * 100).round(
            1)
        fig4 = px.bar(df_summary, x='Nagara', y='Representation_Rate',
                      title='Booth Representation Rate (%)',
                      color='Representation_Rate',
                      color_continuous_scale='blues')
        summary_charts_js += f"Plotly.newPlot('booth-chart', {fig4.to_json()});\n"

    # Create Plotly charts for Sheet8 (Detailed)
    detailed_charts_js = ""

    # 1. Top Vasatis
    if 'Grand Total' in df_detailed.columns and 'Vasati' in df_detailed.columns:
        top_vasatis = df_detailed.nlargest(10, 'Grand Total')
        fig5 = px.bar(top_vasatis, x='Vasati', y='Grand Total',
                      title='Top 10 Vasatis by Attendance',
                      color='Grand Total',
                      color_continuous_scale='plasma')
        detailed_charts_js += f"Plotly.newPlot('top-vasatis-chart', {fig5.to_json()});\n"

    # 2. Categories Chart
    if all(col in df_detailed.columns for col in ['Tarun', 'Balak', 'Women']):
        sample_data = df_detailed.head(8)
        categories_fig = go.Figure(data=[
            go.Bar(name='Tarun', x=sample_data['Vasati'], y=sample_data['Tarun']),
            go.Bar(name='Balak', x=sample_data['Vasati'], y=sample_data['Balak'])
        ])
        categories_fig.update_layout(title='Participant Categories by Vasati', barmode='stack')
        detailed_charts_js += f"Plotly.newPlot('categories-chart', {categories_fig.to_json()});\n"

    # 3. Grade Distribution
    if 'Grade' in df_detailed.columns:
        grade_counts = df_detailed['Grade'].value_counts()
        fig6 = px.pie(values=grade_counts.values, names=grade_counts.index,
                      title='Grade Distribution Across Vasatis')
        detailed_charts_js += f"Plotly.newPlot('grade-chart', {fig6.to_json()});\n"

    # 4. Booth Detailed
    if all(col in df_detailed.columns for col in ['Total Booths', 'Represented Booths']):
        sample_booth = df_detailed.head(10)
        booth_fig = go.Figure(data=[
            go.Bar(name='Total Booths', x=sample_booth['Vasati'], y=sample_booth['Total Booths']),
            go.Bar(name='Represented Booths', x=sample_booth['Vasati'], y=sample_booth['Represented Booths'])
        ])
        booth_fig.update_layout(title='Booth Representation by Vasati', barmode='group')
        detailed_charts_js += f"Plotly.newPlot('booth-detailed-chart', {booth_fig.to_json()});\n"

    # Create summary statistics
    if 'Grand Total' in df_summary.columns:
        total_attendance = df_summary['Grand Total'].sum()
        total_nagaras = len(df_summary)
        avg_attendance = int(total_attendance / total_nagaras)
        total_ghosh = df_summary['Ghosh'].sum() if 'Ghosh' in df_summary.columns else 0

        stats_html = f"""
        <div class="stat-card">
            <div class="stat-label">Total Attendance</div>
            <div class="stat-number">{total_attendance:,}</div>
        </div>
        <div class="stat-card">
            <div class="stat-label">Number of Nagaras</div>
            <div class="stat-number">{total_nagaras}</div>
        </div>
        <div class="stat-card">
            <div class="stat-label">Average per Nagara</div>
            <div class="stat-number">{avg_attendance:,}</div>
        </div>
        <div class="stat-card">
            <div class="stat-label">Total Ghosh</div>
            <div class="stat-number">{total_ghosh}</div>
        </div>
        """
        html_content = html_content.replace('<!-- Stats will be populated by JavaScript -->', stats_html)

    # Add all JavaScript charts to the HTML
    charts_js = f"""
    <script>
        // Summary Tab Charts
        {summary_charts_js}

        // Detailed Tab Charts  
        {detailed_charts_js}

        // Insights Tab Charts (you can add more analytical charts here)
        // Performance comparison, trends, etc.
    </script>
    """

    html_content += charts_js

    # Save the complete dashboard
    with open("vijayadashami_plotly_dashboard.html", "w", encoding='utf-8') as f:
        f.write(html_content)

    print("✅ Plotly Dashboard with Tabs saved as 'vijayadashami_plotly_dashboard.html'")
    return html_content


# Main execution
if __name__ == "__main__":
    file_path = "Vijayadashami_VIJ_2025.xlsx"

    # Create complete dashboard solution
    create_plotly_tabs_dashboard(file_path)