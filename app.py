from __future__ import annotations

from datetime import datetime
from io import BytesIO
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st


APP_DIR = Path(__file__).resolve().parent
DATA_DIR = APP_DIR / "data"
EXPECTED_COLUMNS = [
    "dscomment",
    "source_name",
    "location",
    "trx_time",
    "trx_name",
    "user_first_name",
    "user_name",
    "carrier_key",
    "equip_name",
]
SHIFT_ORDER = ["Shift 1", "Shift 2", "Shift 3"]
SHIFT_COLORS = {
    "Shift 1": "#4C78A8",
    "Shift 2": "#F58518",
    "Shift 3": "#54A24B",
}
PLOT_TEMPLATE = "plotly_white"
OVERVIEW_PLOT_CONFIG = {
    "displayModeBar": True,
    "displaylogo": False,
    "scrollZoom": True,
    "modeBarButtonsToRemove": ["toImage"],
}

st.set_page_config(
    page_title="Carrier Mapping Dashboard",
    page_icon="📊",
    layout="wide",
)


def discover_files(folder: Path) -> tuple[list[Path], tuple[tuple[str, int], ...]]:
    """Return supported Excel files and a stable name/modified-time cache key."""
    folder.mkdir(parents=True, exist_ok=True)
    files = sorted(
        path
        for path in folder.iterdir()
        if path.is_file()
        and path.suffix.lower() in {".xlsx", ".xls"}
        and not path.name.startswith("~$")
    )
    signature = tuple(
        (path.name, path.stat().st_mtime_ns)
        for path in files
    )
    return files, signature


@st.cache_data(show_spinner=False)
def load_excel_files(
    folder_path: str,
    file_signature: tuple[tuple[str, int], ...],
) -> tuple[pd.DataFrame, list[str], str]:
    """Load, combine, and deduplicate workbooks for the given folder signature."""
    folder = Path(folder_path)
    frames: list[pd.DataFrame] = []
    errors: list[str] = []

    for file_name, _ in file_signature:
        file_path = folder / file_name
        try:
            frame = pd.read_excel(file_path, engine="openpyxl" if file_path.suffix.lower() == ".xlsx" else "xlrd")
            frame.columns = [str(column).strip() for column in frame.columns]
            for column in EXPECTED_COLUMNS:
                if column not in frame.columns:
                    frame[column] = pd.NA
            frames.append(frame)
        except (PermissionError, OSError, ValueError, ImportError) as exc:
            errors.append(
                f"{file_name}: couldn't read this workbook. Close it in Excel, check that it is a valid "
                f".xls/.xlsx file, and try again. ({exc})"
            )
        except Exception as exc:  # Excel readers can raise several library-specific corruption errors.
            errors.append(
                f"{file_name}: couldn't read this workbook. Close it in Excel, check that it is a valid "
                f".xls/.xlsx file, and try again. ({exc})"
            )

    if frames:
        combined = pd.concat(frames, ignore_index=True, sort=False)
        combined = combined.drop_duplicates().reset_index(drop=True)
    else:
        combined = pd.DataFrame(columns=EXPECTED_COLUMNS)

    for column in EXPECTED_COLUMNS:
        if column not in combined.columns:
            combined[column] = pd.NA
    combined["trx_time"] = pd.to_datetime(combined["trx_time"], errors="coerce")

    loaded_at = datetime.now().astimezone().strftime("%Y-%m-%d %H:%M:%S %Z")
    return combined, errors, loaded_at


def prepare_data(raw: pd.DataFrame) -> pd.DataFrame:
    """Add reporting-day, shift, and display fields without changing source filters."""
    data = raw.copy()
    data["window_date"] = (data["trx_time"] - pd.Timedelta(hours=7)).dt.date
    hour = data["trx_time"].dt.hour
    data["Shift"] = pd.Series(pd.NA, index=data.index, dtype="object")
    data.loc[hour.between(7, 14), "Shift"] = "Shift 1"
    data.loc[hour.between(15, 22), "Shift"] = "Shift 2"
    data.loc[(hour >= 23) | (hour < 7), "Shift"] = "Shift 3"
    data["Month"] = data["window_date"].map(
        lambda value: value.strftime("%Y-%m") if pd.notna(value) else pd.NA
    )
    data["Hour"] = hour
    data["Weekday"] = data["trx_time"].dt.day_name()
    data["User"] = data["user_name"].astype("string").str.strip()
    missing_user = data["User"].isna() | data["User"].eq("")
    data.loc[missing_user, "User"] = data.loc[missing_user, "user_first_name"].astype("string").str.strip()
    data["User"] = data["User"].fillna("Unknown")
    data.loc[data["User"].eq(""), "User"] = "Unknown"
    return data


def apply_filters(data: pd.DataFrame) -> pd.DataFrame:
    """Build the sidebar and apply its selections to every dashboard section."""
    st.sidebar.header("Filters")
    valid_dates = data["window_date"].dropna()
    if valid_dates.empty:
        st.sidebar.info("Date filters are unavailable until a readable transaction time is loaded.")
        return data.iloc[0:0].copy()

    minimum_date = min(valid_dates)
    maximum_date = max(valid_dates)
    selected_range = st.sidebar.date_input(
        "Reporting date range",
        value=(minimum_date, maximum_date),
        min_value=minimum_date,
        max_value=maximum_date,
    )
    if isinstance(selected_range, (tuple, list)) and len(selected_range) == 2:
        start_date, end_date = selected_range
    elif isinstance(selected_range, (tuple, list)) and len(selected_range) == 1:
        start_date = end_date = selected_range[0]
    else:
        start_date = end_date = selected_range

    date_match = data["window_date"].between(start_date, end_date, inclusive="both")
    result = data.loc[date_match | data["window_date"].isna()].copy()

    source_search = st.sidebar.text_input(
        "Source",
        placeholder="Enter text to find in source names",
    ).strip()
    if source_search:
        source_match = result["source_name"].astype("string").str.contains(
            source_search,
            case=False,
            regex=False,
            na=False,
        )
        result = result.loc[source_match]

    def multiselect_filter(label: str, column: str) -> None:
        nonlocal result
        values = sorted(result[column].dropna().astype(str).unique().tolist())
        chosen = st.sidebar.multiselect(label, options=values, default=values)
        if len(chosen) != len(values):
            result = result.loc[result[column].astype("string").isin(chosen)]

    multiselect_filter("User (first name)", "user_first_name")
    multiselect_filter("Shift", "Shift")
    return result


def make_daily_summary(data: pd.DataFrame) -> pd.DataFrame:
    """Create one row per reporting day with all three shift counts."""
    valid = data.dropna(subset=["window_date"])
    if valid.empty:
        return pd.DataFrame(
            columns=["window_date", "window_range", *SHIFT_ORDER, "Total"]
        )

    counts = (
        valid.groupby(["window_date", "Shift"], observed=True)
        .size()
        .unstack(fill_value=0)
        .reindex(columns=SHIFT_ORDER, fill_value=0)
    )
    counts["Total"] = counts.sum(axis=1)
    summary = counts.reset_index()
    summary["window_range"] = summary["window_date"].map(
        lambda day: f"{day:%Y-%m-%d} 07:00 to {(pd.Timestamp(day) + pd.Timedelta(days=1)):%Y-%m-%d} 07:00"
    )
    return summary[["window_date", "window_range", *SHIFT_ORDER, "Total"]]


def styled_chart(fig: go.Figure, height: int = 400) -> go.Figure:
    fig.update_layout(
        template=PLOT_TEMPLATE,
        height=height,
        margin=dict(l=20, r=20, t=55, b=20),
        legend_title_text="",
    )
    return fig


def labeled_heatmap(
    values: pd.DataFrame,
    title: str,
    x_title: str,
    y_title: str,
    height: int,
) -> go.Figure:
    """Create a color-scaled heatmap with readable counts in each cell."""
    matrix = values.to_numpy(dtype="int64")
    maximum = max(int(matrix.max()), 1)
    x_positions = list(range(len(values.columns)))
    y_positions = list(range(len(values.index)))
    cell_labels = [
        [
            [str(values.index[row_index]), str(values.columns[column_index])]
            for column_index in x_positions
        ]
        for row_index in y_positions
    ]
    fig = go.Figure(
        data=go.Heatmap(
            z=matrix,
            x=x_positions,
            y=y_positions,
            customdata=cell_labels,
            colorscale="Blues",
            zmin=0,
            zmax=maximum,
            colorbar=dict(title="Mapping"),
            hovertemplate=(
                f"{y_title}: %{{customdata[0]}}<br>"
                f"{x_title}: %{{customdata[1]}}<br>"
                "Mapping: %{z}<extra></extra>"
            ),
        )
    )
    fig.update_layout(
        title=title,
        xaxis=dict(
            title=x_title,
            tickmode="array",
            tickvals=x_positions,
            ticktext=[str(column) for column in values.columns],
        ),
        yaxis=dict(
            title=y_title,
            tickmode="array",
            tickvals=y_positions,
            ticktext=[str(row) for row in values.index],
            autorange="reversed",
        ),
    )
    for row_index in y_positions:
        for column_index in x_positions:
            value = matrix[row_index, column_index]
            fig.add_annotation(
                x=column_index,
                y=row_index,
                text=f"{value:,}",
                showarrow=False,
                font=dict(
                    size=10,
                    color="white" if value >= maximum * 0.55 else "#1F2937",
                ),
            )
    return styled_chart(fig, height=height)


def show_kpis(data: pd.DataFrame) -> None:
    valid = data.dropna(subset=["window_date"])
    total = len(valid)
    daily = valid.groupby("window_date").size()
    average = total / daily.index.nunique() if not daily.empty else 0
    if daily.empty:
        busiest_text = "—"
    else:
        busiest_day = daily.idxmax()
        busiest_text = f"{busiest_day:%Y-%m-%d} ({daily.max():,})"
    locations = valid["location"].dropna().astype(str).str.strip()
    location_count = locations[locations.ne("")].nunique()

    columns = st.columns(4)
    columns[0].metric("Total Mappings", f"{total:,}")
    columns[1].metric("Average per active day", f"{average:,.1f}")
    columns[2].metric("Busiest day (count)", busiest_text)
    columns[3].metric("Locations", f"{location_count:,}")


def build_overview_export_figure(data: pd.DataFrame) -> go.Figure:
    valid = data.dropna(subset=["window_date"])
    daily_counts = valid.groupby("window_date").size()
    daily = daily_counts.rename("Mapping")
    calendar = pd.date_range(daily.index.min(), daily.index.max(), freq="D").date
    daily = daily.reindex(calendar, fill_value=0).rename_axis("window_date").reset_index()
    daily["7-day average"] = daily["Mapping"].rolling(7, min_periods=1).mean()
    summary = make_daily_summary(data)

    total = len(valid)
    locations = valid["location"].dropna().astype(str).str.strip()
    location_count = locations[locations.ne("")].nunique()
    busiest_day = daily_counts.idxmax()
    table_height = max(150, 55 + len(summary) * 24)
    figure_height = 150 + 400 + table_height + 150
    row_heights = [
        150 / figure_height,
        400 / figure_height,
        table_height / figure_height,
    ]

    fig = make_subplots(
        rows=3,
        cols=4,
        specs=[
            [{"type": "indicator"} for _ in range(4)],
            [{"type": "xy", "colspan": 4}, None, None, None],
            [{"type": "table", "colspan": 4}, None, None, None],
        ],
        row_heights=row_heights,
        vertical_spacing=0.08,
    )
    kpis = [
        ("Total Mappings", total, ",d"),
        ("Average per active day", total / daily_counts.size, ",.1f"),
        (f"Busiest day<br>{busiest_day:%Y-%m-%d}", int(daily_counts.max()), ",d"),
        ("Locations", location_count, ",d"),
    ]
    for column, (label, value, number_format) in enumerate(kpis, start=1):
        fig.add_trace(
            go.Indicator(
                mode="number",
                value=value,
                title={"text": label, "font": {"size": 14}},
                number={"valueformat": number_format, "font": {"size": 34}},
            ),
            row=1,
            col=column,
        )

    fig.add_trace(
        go.Scatter(
            x=daily["window_date"],
            y=daily["Mapping"],
            mode="lines+markers",
            name="Mapping",
            line=dict(color="#4C78A8"),
            hovertemplate="Window date: %{x}<br>Mapping: %{y:,}<extra></extra>",
        ),
        row=2,
        col=1,
    )
    fig.add_trace(
        go.Scatter(
            x=daily["window_date"],
            y=daily["7-day average"],
            mode="lines",
            name="7-day moving average",
            line=dict(color="#F58518", dash="dash"),
            hovertemplate="Window date: %{x}<br>7-day average: %{y:,.1f}<extra></extra>",
        ),
        row=2,
        col=1,
    )
    fig.add_trace(
        go.Table(
            columnwidth=[1, 3, 1, 1, 1, 1],
            header=dict(
                values=["Window date", "Window range", *SHIFT_ORDER, "Total"],
                fill_color="#4C78A8",
                font=dict(color="white"),
                align="left",
            ),
            cells=dict(
                values=[
                    summary["window_date"].astype(str),
                    summary["window_range"],
                    *[summary[shift] for shift in SHIFT_ORDER],
                    summary["Total"],
                ],
                fill_color="#F2F5FA",
                align="left",
                height=24,
            ),
        ),
        row=3,
        col=1,
    )
    fig.update_xaxes(title_text="Window date", row=2, col=1)
    fig.update_yaxes(title_text="Mapping", row=2, col=1)
    fig.update_layout(
        template=PLOT_TEMPLATE,
        width=1600,
        height=figure_height,
        title="Overview",
        margin=dict(l=45, r=45, t=90, b=50),
        legend_title_text="",
    )
    return fig


def render_overview(data: pd.DataFrame) -> None:
    st.subheader("Key metrics")
    show_kpis(data)
    valid = data.dropna(subset=["window_date"])
    if valid.empty:
        st.info("No transactions with a valid trx_time match the selected filters.")
        return

    daily = valid.groupby("window_date").size().rename("Mapping")
    calendar = pd.date_range(daily.index.min(), daily.index.max(), freq="D").date
    daily = daily.reindex(calendar, fill_value=0).rename_axis("window_date").reset_index()
    daily["7-day average"] = daily["Mapping"].rolling(7, min_periods=1).mean()
    st.download_button(
        "📸 Download full Overview as PNG",
        data=lambda: build_overview_export_figure(data).to_image(
            format="png",
            width=1600,
            scale=1,
        ),
        file_name="mapping_overview.png",
        mime="image/png",
        help="Downloads the KPIs, daily trend, and daily summary together as one image.",
    )
    fig = go.Figure()
    fig.add_trace(
        go.Scatter(
            x=daily["window_date"],
            y=daily["Mapping"],
            mode="lines+markers",
            name="Mapping",
            line=dict(color="#4C78A8"),
        )
    )
    fig.add_trace(
        go.Scatter(
            x=daily["window_date"],
            y=daily["7-day average"],
            mode="lines",
            name="7-day moving average",
            line=dict(color="#F58518", dash="dash"),
        )
    )
    fig.update_layout(
        title="Daily Mapping by reporting day",
        xaxis_title="Window date",
        yaxis_title="Mapping",
    )
    st.plotly_chart(
        styled_chart(fig),
        use_container_width=True,
        config=OVERVIEW_PLOT_CONFIG,
    )
    st.subheader("Daily summary")
    st.dataframe(make_daily_summary(data), use_container_width=True, hide_index=True)


def render_shifts(data: pd.DataFrame) -> None:
    valid = data.dropna(subset=["window_date", "Shift"])
    if valid.empty:
        st.info("No shift data matches the selected filters.")
        return
    counts = valid["Shift"].value_counts().reindex(SHIFT_ORDER, fill_value=0).rename_axis("Shift").reset_index(name="Transactions")
    left, right = st.columns(2)
    with left:
        fig = px.bar(
            counts,
            x="Shift",
            y="Transactions",
            color="Shift",
            category_orders={"Shift": SHIFT_ORDER},
            color_discrete_map=SHIFT_COLORS,
            text="Transactions",
            title="Mapping by shift",
            labels={"Transactions": "Mapping"},
        )
        fig.update_traces(
            textposition="outside",
            texttemplate="%{y:,}",
            cliponaxis=False,
        )
        st.plotly_chart(styled_chart(fig), use_container_width=True)
    with right:
        by_day = valid.groupby(["window_date", "Shift"], observed=True).size().reset_index(name="Transactions")
        fig = px.bar(
            by_day,
            x="window_date",
            y="Transactions",
            color="Shift",
            category_orders={"Shift": SHIFT_ORDER},
            color_discrete_map=SHIFT_COLORS,
            text="Transactions",
            title="Daily Mapping split by shift",
            labels={"Transactions": "Mapping"},
        )
        fig.update_traces(
            textposition="outside",
            texttemplate="%{y:,}",
            cliponaxis=False,
        )
        st.plotly_chart(styled_chart(fig), use_container_width=True)


def render_monthly(data: pd.DataFrame) -> None:
    valid = data.dropna(subset=["Month"])
    valid = valid.loc[valid["Month"] >= "2026-01"]
    if valid.empty:
        st.info("No transactions from January 2026 onward match the selected filters.")
        return
    latest_month = valid["Month"].max()
    months = pd.period_range("2026-01", latest_month, freq="M").astype(str)
    totals = valid.groupby("Month").size().reindex(months, fill_value=0).rename("Transactions")
    summary = totals.rename_axis("Month").reset_index()
    st.dataframe(summary, use_container_width=True, hide_index=True)
    fig = px.bar(
        summary,
        x="Month",
        y="Transactions",
        title="Monthly Mapping from January 2026",
        labels={"Transactions": "Mapping"},
    )
    st.plotly_chart(styled_chart(fig), use_container_width=True)


def render_users(data: pd.DataFrame) -> None:
    valid = data.dropna(subset=["Shift"])
    if valid.empty:
        st.info("No user or shift data matches the selected filters.")
        return
    user_counts = (
        valid.groupby("User")
        .size()
        .sort_values(ascending=False)
        .head(30)
        .rename("Transactions")
        .rename_axis("User")
        .reset_index()
    )
    left, right = st.columns(2)
    with left:
        fig = px.bar(
            user_counts.sort_values("Transactions"),
            x="Transactions",
            y="User",
            orientation="h",
            text="Transactions",
            title="Mapping per user (top 30)",
            labels={"Transactions": "Mapping"},
        )
        fig.update_traces(
            textposition="outside",
            texttemplate="%{x:,}",
            cliponaxis=False,
        )
        st.plotly_chart(styled_chart(fig, height=max(400, len(user_counts) * 24)), use_container_width=True)
    with right:
        top_users = user_counts["User"].tolist()
        heat = (
            valid.loc[valid["User"].isin(top_users)]
            .groupby(["User", "Shift"], observed=True)
            .size()
            .unstack(fill_value=0)
            .reindex(index=top_users, columns=SHIFT_ORDER, fill_value=0)
        )
        fig = labeled_heatmap(
            heat,
            title="User vs. shift (top 30 users)",
            x_title="Shift",
            y_title="User",
            height=max(400, len(top_users) * 24),
        )
        st.plotly_chart(fig, use_container_width=True)


def render_location_source(data: pd.DataFrame) -> None:
    data = data.dropna(subset=["window_date"])
    left, right = st.columns(2)
    for container, column, title in (
        (left, "location", "Top 15 locations"),
        (right, "source_name", "Top 15 sources"),
    ):
        values = data[column].dropna().astype(str).str.strip()
        values = values[values.ne("")]
        counts = values.value_counts().head(15).sort_values().rename_axis(column).reset_index(name="Transactions")
        with container:
            if counts.empty:
                st.info(f"No {column} values match the selected filters.")
                continue
            fig = px.bar(
                counts,
                x="Transactions",
                y=column,
                orientation="h",
                title=title,
                labels={"Transactions": "Mapping"},
            )
            st.plotly_chart(styled_chart(fig, height=max(400, len(counts) * 30)), use_container_width=True)


def render_hourly(data: pd.DataFrame) -> None:
    valid = data.dropna(subset=["Hour", "Weekday"])
    if valid.empty:
        st.info("No hourly data matches the selected filters.")
        return
    weekday_order = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
    heat = (
        valid.groupby(["Weekday", "Hour"], observed=True)
        .size()
        .unstack(fill_value=0)
        .reindex(index=weekday_order, columns=range(24), fill_value=0)
    )
    fig = labeled_heatmap(
        heat,
        title="Mapping by hour and weekday",
        x_title="Hour of day",
        y_title="Weekday",
        height=400,
    )
    st.plotly_chart(fig, use_container_width=True)


def excel_bytes(data: pd.DataFrame) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        data.to_excel(writer, index=False, sheet_name="Filtered data")
    return output.getvalue()


def render_data(data: pd.DataFrame) -> None:
    st.subheader("Filtered transaction details")
    search = st.text_input("Search rows", placeholder="Type a value to find in the table")
    display_data = data
    if search.strip():
        term = search.strip().casefold()
        mask = display_data.astype("string").apply(
            lambda column: column.str.casefold().str.contains(term, regex=False, na=False)
        ).any(axis=1)
        display_data = display_data.loc[mask]

    download_columns = st.columns(2)
    csv_data = display_data.to_csv(index=False).encode("utf-8-sig")
    download_columns[0].download_button(
        "Download filtered data (CSV)",
        data=csv_data,
        file_name="filtered_transactions.csv",
        mime="text/csv",
    )
    download_columns[1].download_button(
        "Download filtered data (Excel)",
        data=excel_bytes(display_data),
        file_name="filtered_transactions.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    st.caption(f"{len(display_data):,} rows shown. Use the column headers to sort.")
    st.dataframe(display_data, use_container_width=True, hide_index=True)


def main() -> None:
    st.title("Carrier Maping Dashboard")
    st.caption("Reporting day: 07:00 to 07:00 the following day")

    files, signature = discover_files(DATA_DIR)
    if st.sidebar.button("Refresh data", use_container_width=True):
        load_excel_files.clear()
        st.rerun()

    if not files:
        st.info(f"No Excel files found in `{DATA_DIR}`. Add .xlsx or .xls files to the data folder.")
        st.caption("Last data update: not loaded | Files: 0 | Rows: 0")
        return

    with st.spinner("Loading Excel files..."):
        raw, errors, loaded_at = load_excel_files(str(DATA_DIR), signature)
    st.caption(f"Last data update: {loaded_at} | Files: {len(files)} | Rows: {len(raw):,}")
    for error in errors:
        st.warning(error)
    if errors and not len(raw):
        st.error("No workbook could be loaded. Close open files or replace damaged files, then click Refresh data.")
        return

    data = prepare_data(raw)
    invalid_times = int(data["trx_time"].isna().sum())
    if invalid_times:
        st.warning(
            f"{invalid_times:,} row(s) have a blank or unreadable trx_time. "
            "They remain available in the detailed table but are not included in date-based summaries."
        )

    filtered = apply_filters(data)
    if filtered.empty:
        st.info("No rows match the selected filters.")
    tabs = st.tabs(
        ["Overview", "Shift analysis", "Monthly", "User performance", "Locations & sources", "Hourly pattern", "Data"]
    )
    with tabs[0]:
        render_overview(filtered)
    with tabs[1]:
        render_shifts(filtered)
    with tabs[2]:
        render_monthly(filtered)
    with tabs[3]:
        render_users(filtered)
    with tabs[4]:
        render_location_source(filtered)
    with tabs[5]:
        render_hourly(filtered)
    with tabs[6]:
        render_data(filtered)


if __name__ == "__main__":
    main()
