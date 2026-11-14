import streamlit as st
import pandas as pd
import plotly.express as px
from pathlib import Path

# Page configuration
st.set_page_config(layout="wide")

# Find the workbook either in the existing Design1 folder or beside this script.
BASE_DIR = Path(__file__).resolve().parent
CANDIDATE_PATHS = [
    BASE_DIR / "Design1" / "Vijayadashami_VIJ_2025.xlsx",
    BASE_DIR / "Vijayadashami_VIJ_2025.xlsx",
    BASE_DIR / "Vijayadashami_VIJ_2025(1).xlsx",
]
file_path = next((path for path in CANDIDATE_PATHS if path.exists()), CANDIDATE_PATHS[0])


@st.cache_data
def load_sheets(path):
    return pd.read_excel(
        path,
        sheet_name=["Sheet8", "2024", "Sheet3", "2026"]
    )


if not Path(file_path).exists():
    st.error(
        f"Excel file not found: {file_path}. "
        "Place the workbook beside design3.py or inside the Design1 folder."
    )
    st.stop()

sheets = load_sheets(str(file_path))

# Dashboard title
st.title("Vijayadashami 2025–2026 (VIJAYANAGARA BHAGA)")

# Main tabs
tab8, tab3 = st.tabs(["Vijayadashami Utsava", "Patha Sanchalana"])

with tab8:
    df8 = sheets["Sheet8"].copy()

    # Clean and prepare data
    df8["Nagara"] = df8["Nagara"].astype(str).str.strip()
    df8["Grand Total"] = pd.to_numeric(df8["Grand Total"], errors="coerce").fillna(0)

    df8_summary = (
        df8.groupby("Nagara", as_index=False)["Grand Total"]
        .sum()
        .sort_values("Grand Total", ascending=True)
    )

    fig8 = px.bar(
        df8_summary,
        x="Nagara",
        y="Grand Total",
        title="Grand Total per Nagara (Utsava)",
        labels={"Nagara": "Nagara", "Grand Total": "Total"},
        color="Grand Total",
        color_continuous_scale="Viridis",
    )
    y_max = df8_summary["Grand Total"].max()
    fig8.update_layout(
        yaxis=dict(range=[0, y_max * 1.25 if y_max else 1]),
        height=500,
        showlegend=False,
    )
    fig8.update_xaxes(tickfont=dict(size=14, family="Arial Black, sans-serif"))
    fig8.update_traces(texttemplate="%{y}", textposition="outside", textfont=dict(size=16))

    df8_top5 = df8.nlargest(5, "Grand Total")[["Nagara", "Vasati", "Grand Total"]]
    fig8_top5 = px.bar(
        df8_top5,
        x="Vasati",
        y="Grand Total",
        title="Top 5 Vasati (Utsava)",
        labels={"Vasati": "Vasati", "Grand Total": "Total"},
        color="Grand Total",
        color_continuous_scale="Reds",
    )
    top5_max = df8_top5["Grand Total"].max()
    fig8_top5.update_layout(
        yaxis=dict(range=[0, top5_max * 1.25 if top5_max else 1]),
        height=400,
        showlegend=False,
    )
    fig8_top5.update_xaxes(tickfont=dict(size=14, family="Arial Black, sans-serif"))
    fig8_top5.update_traces(texttemplate="%{y}", textposition="outside", textfont=dict(size=16))

    unique_nagara = sorted(df8["Nagara"].dropna().unique())
    selected_nagara = st.selectbox(
        "Select a Nagara to view Vasati Details:",
        options=unique_nagara,
        index=0,
        key="utsava_nagara_select",
    )
    df_selected = df8[df8["Nagara"] == selected_nagara]

    if not df_selected.empty:
        fig8_vasati = px.bar(
            df_selected,
            x="Vasati",
            y="Grand Total",
            title=f"Grand Total per Vasati in {selected_nagara}",
            labels={"Vasati": "Vasati", "Grand Total": "Total"},
            color="Grand Total",
            color_continuous_scale="Greens",
        )
        vasati_max = df_selected["Grand Total"].max()
        fig8_vasati.update_layout(
            yaxis=dict(range=[0, vasati_max * 1.25 if vasati_max else 1]),
            height=500,
            showlegend=False,
        )
        fig8_vasati.update_xaxes(tickfont=dict(size=14, family="Arial Black, sans-serif"))
        fig8_vasati.update_traces(
            texttemplate="%{y}", textposition="outside", textfont=dict(size=16)
        )
    else:
        fig8_vasati = None
        st.warning(f"No data found for {selected_nagara}")

    # Distribution of Utsava attendance
    bin_edges = [0, 10, 25, 51, 75, 100, 501]
    bin_labels = ["0-9", "10-24", "25-50", "51-74", "75-99", "100-500"]
    df8["Bin"] = pd.cut(
        df8["Grand Total"],
        bins=bin_edges,
        labels=bin_labels,
        right=False,
        include_lowest=True,
    )
    binned_df = (
        df8.groupby("Bin", observed=True)
        .size()
        .reset_index(name="Frequency")
        .sort_values("Bin")
    )
    fig8_hist = px.bar(
        binned_df,
        x="Bin",
        y="Frequency",
        title="Utsava Attendance Grouping",
        labels={"Bin": "Upasthiti", "Frequency": "Utsava Sankhye"},
        color="Frequency",
        color_continuous_scale="Oranges",
    )
    hist_max = binned_df["Frequency"].max()
    fig8_hist.update_layout(
        yaxis=dict(range=[0, hist_max * 1.25 if hist_max else 1]),
        height=400,
        showlegend=False,
        margin=dict(l=100),
    )
    fig8_hist.update_xaxes(tickfont=dict(size=14, family="Arial Black, sans-serif"))
    fig8_hist.update_traces(texttemplate="%{y}", textposition="outside", textfont=dict(size=20))

    nagara_count = df8["Nagara"].value_counts().reset_index()
    nagara_count.columns = ["Nagara", "Count"]
    fig8_pie = px.pie(
        nagara_count,
        values="Count",
        names="Nagara",
        title="Total Utsava",
        color_discrete_sequence=px.colors.qualitative.Set3,
    )
    fig8_pie.update_traces(textinfo="label+value", textfont_size=14)
    fig8_pie.update_layout(height=400, showlegend=True)

    subtab_plot, subtab_data = st.tabs(["Summary Plot", "Detailed Data"])
    with subtab_plot:
        if fig8_vasati is not None:
            st.plotly_chart(fig8_vasati, use_container_width=True)
        st.plotly_chart(fig8, use_container_width=True)
        st.plotly_chart(fig8_top5, use_container_width=True)
        st.plotly_chart(fig8_hist, use_container_width=True)
        st.plotly_chart(fig8_pie, use_container_width=True)

    with subtab_data:
        st.subheader("Vijayadashami Utsava")
        st.dataframe(df8, use_container_width=True)


with tab3:
    st.subheader("Patha Sanchalana – 2024, 2025 and 2026")

    # Read the three year-wise datasets.
    year_sheets = {
        "2024": sheets["2024"],
        "2025": sheets["Sheet3"],
        "2026": sheets["2026"],
    }

    required_columns = {"Nagara", "Grand Total"}
    prepared_years = []

    for year, frame in year_sheets.items():
        missing = required_columns - set(frame.columns)
        if missing:
            st.error(
                f"{year} worksheet is missing required column(s): "
                f"{', '.join(sorted(missing))}"
            )
            st.stop()

        year_data = frame[["Nagara", "Grand Total"]].copy()
        year_data["Nagara"] = year_data["Nagara"].astype(str).str.strip()
        year_data["Grand Total"] = pd.to_numeric(
            year_data["Grand Total"], errors="coerce"
        )
        year_data = year_data.dropna(subset=["Nagara", "Grand Total"])
        year_data = year_data[
            year_data["Nagara"].ne("") & year_data["Nagara"].ne("nan")
        ]
        year_data = year_data.groupby("Nagara", as_index=False)["Grand Total"].sum()
        year_data = year_data.rename(columns={"Grand Total": year})
        prepared_years.append(year_data)

    # Outer joins retain any Nagara that appears in only one of the years.
    comparison = prepared_years[0]
    for year_data in prepared_years[1:]:
        comparison = comparison.merge(year_data, on="Nagara", how="outer")

    comparison = comparison.fillna({"2024": 0, "2025": 0, "2026": 0})
    comparison["Change 2024–2025"] = comparison["2025"] - comparison["2024"]
    comparison["Change 2025–2026"] = comparison["2026"] - comparison["2025"]
    comparison["Change 2024–2026"] = comparison["2026"] - comparison["2024"]

    comparison["Change 2024–2026 (%)"] = comparison.apply(
        lambda row: (
            (row["Change 2024–2026"] / row["2024"]) * 100
            if row["2024"] != 0
            else (100.0 if row["2026"] > 0 else 0.0)
        ),
        axis=1,
    )
    comparison = comparison.sort_values("Nagara").reset_index(drop=True)

    # Overall year-wise totals and growth.
    total_2024 = comparison["2024"].sum()
    total_2025 = comparison["2025"].sum()
    total_2026 = comparison["2026"].sum()

    overall_change_24_25 = total_2025 - total_2024
    overall_change_25_26 = total_2026 - total_2025
    overall_change_24_26 = total_2026 - total_2024
    overall_pct_24_26 = (
        overall_change_24_26 / total_2024 * 100 if total_2024 else 0
    )

    metric1, metric2, metric3 = st.columns(3)
    metric1.metric("2024 Total", f"{total_2024:,.0f}")
    metric2.metric(
        "2025 Total",
        f"{total_2025:,.0f}",
        delta=f"{overall_change_24_25:+,.0f} vs 2024",
    )
    metric3.metric(
        "2026 Total",
        f"{total_2026:,.0f}",
        delta=f"{overall_change_25_26:+,.0f} vs 2025",
    )

    st.caption(
        f"Overall change from 2024 to 2026: "
        f"{overall_change_24_26:+,.0f} participants "
        f"({overall_pct_24_26:+.1f}%)."
    )

    # Convert to long format for a three-series grouped bar chart.
    comparison_long = comparison.melt(
        id_vars="Nagara",
        value_vars=["2024", "2025", "2026"],
        var_name="Year",
        value_name="Grand Total",
    )

    fig_compare = px.bar(
        comparison_long,
        x="Nagara",
        y="Grand Total",
        color="Year",
        barmode="group",
        text="Grand Total",
        title="Nagara-wise Patha Sanchalana: 2024 vs 2025 vs 2026",
        labels={
            "Nagara": "Nagara",
            "Grand Total": "Participants",
            "Year": "Year",
        },
        color_discrete_map={
            "2024": "#2E86AB",
            "2025": "#3366CC",
            "2026": "#F28E2B",
        },
    )
    fig_compare.update_traces(textposition="outside", cliponaxis=False)
    chart_max = comparison_long["Grand Total"].max()
    fig_compare.update_layout(
        height=650,
        yaxis=dict(
            range=[0, chart_max * 1.2 if chart_max else 1],
            title="Grand Total",
        ),
        xaxis=dict(tickfont=dict(size=13, family="Arial Black, sans-serif")),
        legend_title_text="Year",
        bargap=0.22,
    )

    # Show the change over the full three-year period.
    fig_trend = px.bar(
        comparison,
        x="Nagara",
        y="Change 2024–2026",
        title="Nagara-wise Change: 2026 compared with 2024",
        labels={
            "Nagara": "Nagara",
            "Change 2024–2026": "Change in Participants",
        },
        color="Change 2024–2026",
        color_continuous_scale="RdYlGn",
        text="Change 2024–2026",
    )
    fig_trend.update_traces(textposition="outside", cliponaxis=False)
    fig_trend.update_layout(height=450, showlegend=False)

    subtab_plot, subtab_data = st.tabs(
        ["Three-Year Comparative Charts", "Detailed Comparison"]
    )

    with subtab_plot:
        st.plotly_chart(fig_compare, use_container_width=True)
        st.plotly_chart(fig_trend, use_container_width=True)

    with subtab_data:
        st.subheader("Nagara-wise Three-Year Comparison")
        display_comparison = comparison.copy()
        display_comparison["Change 2024–2026 (%)"] = (
            display_comparison["Change 2024–2026 (%)"].map(
                lambda value: f"{value:+.1f}%"
            )
        )
        st.dataframe(
            display_comparison,
            use_container_width=True,
            hide_index=True,
        )

        st.download_button(
            "Download Three-Year Comparison CSV",
            data=comparison.to_csv(index=False).encode("utf-8-sig"),
            file_name="Patha_Sanchalana_2024_2025_2026.csv",
            mime="text/csv",
        )

# Footer
st.markdown("---")
st.caption("RSS@100 | Sangha Shatabdi | Vijayanagara Bhaga | Bengaluru Dakshina")
