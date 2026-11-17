import streamlit as st
import pandas as pd
import plotly.express as px
from pathlib import Path

# Page configuration

# Page configuration
st.set_page_config(layout="wide")

# Custom CSS – Professional Blue and Saffron Theme

# Custom CSS – Blue and Saffron Theme
st.markdown("""
<style>

/* Tab navigation container */
.stTabs [data-testid="stTab"] {
    background-color: #17365D !important;
    color: white !important;
    border-radius: 8px !important;
    padding: 12px 22px !important;
    font-weight: 600 !important;
}

/* Active tab */
.stTabs [data-testid="stTab"][aria-selected="true"] {
    background-color: #E98224 !important;
    color: white !important;
}

/* Tab label */
.stTabs [data-testid="stTab"] p {
    color: white !important;
    font-size: 15px !important;
    font-weight: bold !important;
}

/* Hover */
.stTabs [data-testid="stTab"]:hover {
    background-color: #315B85 !important;
}

</style>
""", unsafe_allow_html=True)



# Set Pan mode as the default for all Plotly charts
def display_plotly_chart(fig, **kwargs):
    fig.update_layout(dragmode="pan")
    st.plotly_chart(fig, **kwargs)


# Find the workbook either in the existing Design1 folder or beside this script.
BASE_DIR = Path(__file__).resolve().parent
CANDIDATE_PATHS = [
    BASE_DIR / "Design1" / "Vijayadashami_VIJ_2025.xlsx",
    BASE_DIR / "Vijayadashami_VIJ_2025.xlsx",
    BASE_DIR / "Vijayadashami_VIJ_2025(3).xlsx",
    BASE_DIR / "Vijayadashami_VIJ_2025(1).xlsx",
]
file_path = next((path for path in CANDIDATE_PATHS if path.exists()), CANDIDATE_PATHS[0])


@st.cache_data
def load_sheets(path):
    # Most worksheets have a single header row. New_Horizon is a key-value
    # list and Yuva has two header rows, so load those two without headers.
    loaded = pd.read_excel(
        path,
        sheet_name=[
            "Sheet8", "2024", "Sheet3", "2026",
            "Gruha", "HinduSammelan"
        ]
    )
    loaded["New_Horizon"] = pd.read_excel(
        path, sheet_name="New_Horizon", header=None
    )
    loaded["Yuva"] = pd.read_excel(
        path, sheet_name="Yuva", header=None
    )
    return loaded


if not Path(file_path).exists():
    st.error(
        f"Excel file not found: {file_path}. "
        "Place the workbook beside design3.py or inside the Design1 folder."
    )
    st.stop()

sheets = load_sheets(str(file_path))

# Dashboard title
st.title("Centenary Year: RSS@100 – Vijayanagara Bhaga")

# Main tabs
tab8, tab3, tab_gruha, tab_hindu, tab_horizon, tab_yuva = st.tabs([
    "Vijayadashami Utsava",
    "Patha Sanchalana",
    "Gruha Samparka Abhiyana",
    "Hindu Sammelan",
    "Pramukha Nagarika Ghoshti",
    "Yuva",
])

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
            display_plotly_chart(fig8_vasati, use_container_width=True)
        display_plotly_chart(fig8, use_container_width=True)
        display_plotly_chart(fig8_top5, use_container_width=True)
        display_plotly_chart(fig8_hist, use_container_width=True)
        display_plotly_chart(fig8_pie, use_container_width=True)

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
        display_plotly_chart(fig_compare, use_container_width=True)
        display_plotly_chart(fig_trend, use_container_width=True)

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


# -------------------------------------------------------------------
# Gruha – Centenary year house outreach
# -------------------------------------------------------------------
with tab_gruha:
    st.subheader("Gruha – Centenary Year House Outreach")

    raw_gruha = sheets["Gruha"].copy()
    raw_gruha.columns = ["Metric"] + list(raw_gruha.columns[1:])
    raw_gruha["Metric"] = raw_gruha["Metric"].astype(str).str.strip()

    nagara_columns = [c for c in raw_gruha.columns if c != "Metric" and c != "Total"]
    for col in nagara_columns + (["Total"] if "Total" in raw_gruha.columns else []):
        raw_gruha[col] = pd.to_numeric(raw_gruha[col], errors="coerce").fillna(0)

    gruha_total = raw_gruha.set_index("Metric")
    total_houses = float(gruha_total.loc["Reaches to houses", "Total"])
    total_upavasati = float(gruha_total.loc["Total Upavasati", "Total"])
    reached_upavasati = float(gruha_total.loc["Reached Upavasati", "Total"])
    outreach_teams = float(gruha_total.loc["Total Teams", "Total"])

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Houses Reached", f"{total_houses:,.0f}")
    c2.metric("Total Upavasati", f"{total_upavasati:,.0f}")
    c3.metric("Reached Upavasati", f"{reached_upavasati:,.0f}")
    c4.metric("Outreach Teams", f"{outreach_teams:,.0f}")

    st.markdown("#### Nagara-wise House Outreach")
    houses_chart = raw_gruha[["Metric"] + nagara_columns + ["Total"]]
    houses_row = houses_chart[houses_chart["Metric"].str.casefold() == "reaches to houses"]
    if not houses_row.empty:
        houses_by_nagara = houses_row.melt(
            id_vars="Metric", value_vars=nagara_columns,
            var_name="Nagara", value_name="Houses Reached"
        )
        fig_houses = px.bar(
            houses_by_nagara, x="Nagara", y="Houses Reached",
            text="Houses Reached", color="Houses Reached",
            color_continuous_scale="Blues",
            title="Number of Houses Reached in Each Nagara"
        )
        fig_houses.update_traces(textposition="outside", cliponaxis=False)
        fig_houses.update_layout(height=500, showlegend=False)
        display_plotly_chart(fig_houses, use_container_width=True)

    outreach_metrics = [
        "Reaches to houses",
        "Total Book Sold",
        "Total Bharat Mata Photos Sold",
        "Total Number of Outreach Teams",
        "Total Number of Karyakarta",
        "Sagha Karyakarta",
        "Women Participants",
        "Various Sectors Karyakarta",
    ]
    available_outreach_metrics = [
        m for m in outreach_metrics if m in raw_gruha["Metric"].values
    ]
    selected_gruha_metric = st.selectbox(
        "Select a metric for Nagara-wise analysis",
        available_outreach_metrics,
        key="gruha_metric_select",
    )
    selected_row = raw_gruha[raw_gruha["Metric"] == selected_gruha_metric]
    metric_chart_data = selected_row.melt(
        id_vars="Metric", value_vars=nagara_columns,
        var_name="Nagara", value_name="Count"
    )
    fig_gruha_metric = px.bar(
        metric_chart_data, x="Nagara", y="Count", text="Count",
        title=f"{selected_gruha_metric} – Nagara-wise",
        color="Count", color_continuous_scale="Teal"
    )
    fig_gruha_metric.update_traces(textposition="outside", cliponaxis=False)
    fig_gruha_metric.update_layout(height=450, showlegend=False)
    display_plotly_chart(fig_gruha_metric, use_container_width=True)

    with st.expander("View complete Gruha data"):
        st.dataframe(raw_gruha, use_container_width=True, hide_index=True)


# -------------------------------------------------------------------
# Hindu Sammelan – Milestone event
# -------------------------------------------------------------------
with tab_hindu:
    st.subheader("Hindu Sammelan – Milestone Event")

    raw_hindu = sheets["HinduSammelan"].copy()
    raw_hindu.columns = ["Metric"] + list(raw_hindu.columns[1:])
    raw_hindu["Metric"] = raw_hindu["Metric"].astype(str).str.strip()
    hindu_nagara_columns = [
        c for c in raw_hindu.columns if c not in ("Metric", "Total")
    ]
    for col in hindu_nagara_columns + ["Total"]:
        raw_hindu[col] = pd.to_numeric(raw_hindu[col], errors="coerce").fillna(0)

    hindu_index = raw_hindu.set_index("Metric")
    total_hindu_participants = float(hindu_index.loc["Total Participated", "Total"])
    men_total = float(hindu_index.loc["Participated Men", "Total"])
    women_total = float(hindu_index.loc["participated Women", "Total"])
    sammelan_count = float(hindu_index.loc["Total Hindu Sammelan", "Total"])

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Participation", f"{total_hindu_participants:,.0f}")
    c2.metric("Men", f"{men_total:,.0f}")
    c3.metric("Women", f"{women_total:,.0f}")
    c4.metric("Hindu Sammelans", f"{sammelan_count:,.0f}")

    participation_rows = raw_hindu[
        raw_hindu["Metric"].isin(["Participated Men", "participated Women"])
    ]
    participation_long = participation_rows.melt(
        id_vars="Metric", value_vars=hindu_nagara_columns,
        var_name="Nagara", value_name="Participants"
    )
    participation_long["Category"] = participation_long["Metric"].map({
        "Participated Men": "Men",
        "participated Women": "Women",
    })
    fig_hindu_gender = px.bar(
        participation_long, x="Nagara", y="Participants",
        color="Category", barmode="group", text="Participants",
        title="Nagara-wise Hindu Sammelan Participation",
        color_discrete_map={"Men": "#3366CC", "Women": "#E07A9A"}
    )
    fig_hindu_gender.update_traces(textposition="outside", cliponaxis=False)
    fig_hindu_gender.update_layout(height=550)
    display_plotly_chart(fig_hindu_gender, use_container_width=True)

    total_row = raw_hindu[raw_hindu["Metric"] == "Total Participated"]
    if not total_row.empty:
        total_long = total_row.melt(
            id_vars="Metric", value_vars=hindu_nagara_columns,
            var_name="Nagara", value_name="Participants"
        )
        fig_hindu_total = px.bar(
            total_long, x="Nagara", y="Participants", text="Participants",
            title="Total Participation by Nagara",
            color="Participants", color_continuous_scale="Purples"
        )
        fig_hindu_total.update_traces(textposition="outside", cliponaxis=False)
        fig_hindu_total.update_layout(height=450, showlegend=False)
        display_plotly_chart(fig_hindu_total, use_container_width=True)

    with st.expander("View complete Hindu Sammelan data"):
        st.dataframe(raw_hindu, use_container_width=True, hide_index=True)


# -------------------------------------------------------------------
# Pramukha Nagarika Ghoshti – New Horizon
# -------------------------------------------------------------------
with tab_horizon:
    st.subheader("Pramukha Nagarika Ghoshti")

    raw_horizon = sheets["New_Horizon"].copy()
    raw_horizon.columns = ["Metric", "Value"]
    raw_horizon["Metric"] = raw_horizon["Metric"].astype(str).str.strip()
    raw_horizon["Value"] = pd.to_numeric(raw_horizon["Value"], errors="coerce").fillna(0)
    horizon_values = raw_horizon.set_index("Metric")["Value"].to_dict()

    planned = float(horizon_values.get("Planned", 0))
    men = float(horizon_values.get("Participated Men", 0))
    women = float(horizon_values.get("Participated Women", 0))
    participated = float(horizon_values.get("Total Participated", men + women))
    divisions = float(horizon_values.get("Divisions", 0))
    achievement = participated / planned * 100 if planned else 0

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Planned", f"{planned:,.0f}")
    c2.metric("Total Participated", f"{participated:,.0f}")
    c3.metric("Divisions", f"{divisions:,.0f}")
    c4.metric("Plan Achievement", f"{achievement:.1f}%")

    gender_data = pd.DataFrame({
        "Category": ["Men", "Women"],
        "Participants": [men, women],
    })
    fig_horizon_gender = px.pie(
        gender_data, names="Category", values="Participants",
        title="Participation Composition",
        color="Category",
        color_discrete_map={"Men": "#3366CC", "Women": "#E07A9A"},
        hole=0.35
    )
    fig_horizon_gender.update_traces(textinfo="label+value+percent")
    display_plotly_chart(fig_horizon_gender, use_container_width=True)

    plan_data = pd.DataFrame({
        "Status": ["Planned", "Participated"],
        "Count": [planned, participated],
    })
    fig_horizon_plan = px.bar(
        plan_data, x="Status", y="Count", text="Count",
        title="Planned vs Actual Participation",
        color="Status",
        color_discrete_map={"Planned": "#9AA0A6", "Participated": "#2E8B57"}
    )
    fig_horizon_plan.update_traces(textposition="outside", cliponaxis=False)
    fig_horizon_plan.update_layout(height=400, showlegend=False)
    display_plotly_chart(fig_horizon_plan, use_container_width=True)

    st.markdown("#### Source Sheet Details")
    st.dataframe(raw_horizon, use_container_width=True, hide_index=True)


# -------------------------------------------------------------------
# Yuva – Yuva Conclave and participation
# -------------------------------------------------------------------
with tab_yuva:
    st.subheader("Yuva – Yuva Conclave and Youth Participation")

    raw_yuva = sheets["Yuva"].copy()

    # The worksheet uses two header rows. Assign readable column names
    # from the two header levels, then exclude the final grand-total row.
    yuva_columns = [
        "Nagara",
        "Yuva Conclave",
        "Total College",
        "Participated College",
        "Total Hostel",
        "Participated Hostel",
        "Participated VV",
        "Participated Premier Institutions",
        "Corners Meeting",
        "College Students",
        "Young Employees",
        "Total Participants",
    ]
    yuva = raw_yuva.iloc[2:].copy()
    yuva.columns = yuva_columns
    yuva["Nagara"] = yuva["Nagara"].astype(str).str.strip()
    numeric_yuva_columns = yuva_columns[1:]
    for col in numeric_yuva_columns:
        yuva[col] = pd.to_numeric(yuva[col], errors="coerce").fillna(0)

    # The final row in the worksheet is the aggregate total, not a Nagara.
    yuva = yuva[yuva["Nagara"].ne("nan") & yuva["Nagara"].ne("")]
    yuva_nagara = yuva[yuva["Nagara"].ne("IT Milan")].copy()
    yuva_it_milan = yuva[yuva["Nagara"].eq("IT Milan")].copy()

    total_participants = float(raw_yuva.iloc[-1, 11])
    total_college_students = float(raw_yuva.iloc[-1, 9])
    total_young_employees = float(raw_yuva.iloc[-1, 10])
    total_conclaves = float(raw_yuva.iloc[-1, 1])
    total_corners_meetings = float(raw_yuva.iloc[-1, 8])

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Total Participants", f"{total_participants:,.0f}")
    c2.metric("College Students", f"{total_college_students:,.0f}")
    c3.metric("Young Employees", f"{total_young_employees:,.0f}")
    c4.metric("Yuva Conclaves", f"{total_conclaves:,.0f}")
    c5.metric("Corners Meetings", f"{total_corners_meetings:,.0f}")

    yuva_participation_long = yuva_nagara.melt(
        id_vars="Nagara",
        value_vars=["College Students", "Young Employees"],
        var_name="Participant Group",
        value_name="Participants",
    )
    fig_yuva_groups = px.bar(
        yuva_participation_long,
        x="Nagara",
        y="Participants",
        color="Participant Group",
        barmode="group",
        text="Participants",
        title="Nagara-wise Youth Participation",
        color_discrete_map={
            "College Students": "#3366CC",
            "Young Employees": "#F28E2B",
        },
    )
    fig_yuva_groups.update_traces(textposition="outside", cliponaxis=False)
    fig_yuva_groups.update_layout(height=550)
    display_plotly_chart(fig_yuva_groups, use_container_width=True)

    fig_yuva_total = px.bar(
        yuva_nagara,
        x="Nagara",
        y="Total Participants",
        text="Total Participants",
        title="Total Yuva Participation by Nagara",
        color="Total Participants",
        color_continuous_scale="Viridis",
    )
    fig_yuva_total.update_traces(textposition="outside", cliponaxis=False)
    fig_yuva_total.update_layout(height=450, showlegend=False)
    display_plotly_chart(fig_yuva_total, use_container_width=True)

    if not yuva_it_milan.empty:
        st.markdown("#### IT Milan – Separate Contribution")
        st.dataframe(yuva_it_milan, use_container_width=True, hide_index=True)

    with st.expander("View complete Yuva data"):
        st.dataframe(yuva, use_container_width=True, hide_index=True)

# Footer
st.markdown("---")
st.caption("RSS@100 | Sangha Shatabdi | Vijayanagara Bhaga | Bengaluru Dakshina")
