import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np


def create_matplotlib_dashboard(file_path):
    """Create a static dashboard using Matplotlib and Seaborn"""

    # Read data
    excel_data = pd.read_excel(file_path, sheet_name=None, engine='openpyxl')
    df_summary = excel_data['Sheet3'].dropna(how='all')
    df_detailed = excel_data['Sheet8'].dropna(how='all')
    df_summary.columns = df_summary.columns.str.strip()
    df_detailed.columns = df_detailed.columns.str.strip()

    # Set style
    plt.style.use('seaborn-v0_8')
    sns.set_palette("husl")

    # Create figure with subplots
    fig = plt.figure(figsize=(20, 15))

    # Summary plots
    # 1. Total Attendance
    ax1 = plt.subplot2grid((3, 4), (0, 0), colspan=2)
    if 'Grand Total' in df_summary.columns:
        bars = ax1.bar(df_summary['Nagara'], df_summary['Grand Total'])
        ax1.set_title('Total Attendance by Nagara', fontsize=14, fontweight='bold')
        ax1.tick_params(axis='x', rotation=45)
        # Add value labels on bars
        for bar in bars:
            height = bar.get_height()
            ax1.text(bar.get_x() + bar.get_width() / 2., height,
                     f'{int(height):,}', ha='center', va='bottom')

    # 2. Ghosh Participants
    ax2 = plt.subplot2grid((3, 4), (0, 2), colspan=2)
    if 'Ghosh' in df_summary.columns:
        bars = ax2.bar(df_summary['Nagara'], df_summary['Ghosh'], color='orange')
        ax2.set_title('Ghosh Participants by Nagara', fontsize=14, fontweight='bold')
        ax2.tick_params(axis='x', rotation=45)

    # 3. Vasati Distribution
    ax3 = plt.subplot2grid((3, 4), (1, 0), colspan=2)
    if 'Total Vasati' in df_summary.columns:
        ax3.pie(df_summary['Total Vasati'], labels=df_summary['Nagara'], autopct='%1.1f%%')
        ax3.set_title('Vasati Distribution', fontsize=14, fontweight='bold')

    # 4. Booth Representation
    ax4 = plt.subplot2grid((3, 4), (1, 2), colspan=2)
    if all(col in df_summary.columns for col in ['Total Booths', 'Represented Booths']):
        df_summary['Representation_Rate'] = (df_summary['Represented Booths'] / df_summary['Total Booths'] * 100).round(
            1)
        bars = ax4.bar(df_summary['Nagara'], df_summary['Representation_Rate'], color='green')
        ax4.set_title('Booth Representation Rate (%)', fontsize=14, fontweight='bold')
        ax4.tick_params(axis='x', rotation=45)
        ax4.set_ylim(0, 100)

    # Detailed plots
    # 5. Top Vasatis
    ax5 = plt.subplot2grid((3, 4), (2, 0), colspan=2)
    if 'Grand Total' in df_detailed.columns:
        top_vasatis = df_detailed.nlargest(8, 'Grand Total')
        bars = ax5.bar(top_vasatis['Vasati'], top_vasatis['Grand Total'], color='purple')
        ax5.set_title('Top Vasatis by Attendance', fontsize=14, fontweight='bold')
        ax5.tick_params(axis='x', rotation=45)

    # 6. Grade Distribution
    ax6 = plt.subplot2grid((3, 4), (2, 2))
    if 'Grade' in df_detailed.columns:
        grade_counts = df_detailed['Grade'].value_counts()
        ax6.pie(grade_counts.values, labels=grade_counts.index, autopct='%1.1f%%')
        ax6.set_title('Grade Distribution', fontsize=14, fontweight='bold')

    # 7. Participant Categories
    ax7 = plt.subplot2grid((3, 4), (2, 3))
    if all(col in df_detailed.columns for col in ['Tarun', 'Balak', 'Women']):
        sample_data = df_detailed.head(6)
        categories = ['Tarun', 'Balak', 'Women']
        bottom = np.zeros(len(sample_data))

        for i, category in enumerate(categories):
            ax7.bar(sample_data['Vasati'], sample_data[category], bottom=bottom, label=category)
            bottom += sample_data[category]

        ax7.set_title('Participant Categories', fontsize=14, fontweight='bold')
        ax7.tick_params(axis='x', rotation=45)
        ax7.legend()

    plt.tight_layout()
    plt.suptitle('Vijayadashami 2025 - Comprehensive Dashboard', fontsize=16, fontweight='bold', y=1.02)

    # Save the dashboard
    plt.savefig('vijayadashami_matplotlib_dashboard.png', dpi=300, bbox_inches='tight')
    plt.savefig('vijayadashami_matplotlib_dashboard.pdf', bbox_inches='tight')

    print("✅ Matplotlib Dashboard saved as PNG and PDF")
    plt.show()


# Main execution
if __name__ == "__main__":
    file_path = "Vijayadashami_VIJ_2025.xlsx"

    print("Creating Vijayadashami 2025 Dashboards...")
    print("=" * 50)


    # Create Matplotlib Dashboard
    print("2. Creating Matplotlib Static Dashboard...")
    create_matplotlib_dashboard(file_path)

