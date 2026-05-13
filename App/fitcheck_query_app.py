import os
from io import BytesIO
from pathlib import Path
from typing import Dict, List

import pandas as pd
import streamlit as st
from sqlalchemy import create_engine, text

st.set_page_config(page_title="FitCheck Query Portal", layout="wide")

DB_URL = os.getenv("POSTGRES_URL", "postgresql+psycopg2://postgres:postgresDARA1@localhost:5432/fitcheck_main")

# Path to the cleaned export workbook produced by the ETL pipeline
_APP_DIR = Path(__file__).resolve().parent
CLEANED_EXPORT_PATH = Path(os.getenv("FITCHECK_EXPORT_PATH", str(_APP_DIR.parent / "FitCheck_os" / "fitcheck_cleaned_export.xlsx")))

TABLE_CONFIG: Dict[str, Dict[str, List[str]]] = {
    "participants": {
        "default_columns": ["participant_id"],
        "filters": ["participant_id"],
    },
    "demographics_health": {
        "default_columns": [
            "participant_id", "sex_assigned_at_birth", "ethnicity", "age_years",
            "daily_physical_activity", "recurring_chronic_pain", "general_health_rating",
            "ergonomic_familiarity", "weekly_desk_hours_time",
        ],
        "filters": [
            "participant_id", "sex_assigned_at_birth", "ethnicity", "age_years",
            "recurring_chronic_pain", "general_health_rating", "ergonomic_familiarity",
        ],
    },
    "anthropometry": {
        "default_columns": [
            "participant_id", "session_date", "body_mass_kg", "bmi_kg_m2",
            "hip_breadth_sitting_cm", "elbow_height_sitting_cm",
            "buttock_popliteal_length_cm", "popliteal_height_sitting_cm",
        ],
        "filters": [
            "participant_id", "session_date", "body_mass_kg", "bmi_kg_m2",
            "body_fat_percent", "skeletal_muscle_percent",
        ],
    },
    "chair_selection_pre": {
        "default_columns": [
            "participant_id", "pre_selected_chair_code", "pre_selection_basis",
            "pre_selection_confidence_3_3",
            "pre_selection_basis_1", "pre_selection_basis_2", "pre_selection_basis_3",
        ],
        "filters": [
            "participant_id", "pre_selected_chair_code", "pre_selection_confidence_3_3",
            "pre_chair_familiarity_check_1", "pre_chair_familiarity_check_2",
        ],
    },
    "chair_selection_post": {
        "default_columns": [
            "participant_id", "post_selected_chair_code", "post_selection_basis",
            "post_selection_confidence_3_3", "selection_changed", "selection_change_reason",
            "post_selection_basis_1", "post_selection_basis_2",
        ],
        "filters": [
            "participant_id", "post_selected_chair_code", "post_selection_confidence_3_3",
            "selection_changed",
        ],
    },
    "chair_lookup": {
        "default_columns": ["chair_code", "chair_id", "chair_model", "chair_variant", "chair_name_display"],
        "filters": ["chair_code", "chair_id", "chair_model", "chair_variant"],
    },
    "chair_feature_bounds_wide": {
        "default_columns": [
            "chair_code", "chair_id", "chair_model", "chair_variant", "chair_name_display",
            "seat_height_min_cm", "seat_height_max_cm", "seat_depth_min_cm", "seat_depth_max_cm",
            "armrest_height_min_cm", "armrest_height_max_cm", "lumbar_min_unit", "lumbar_max_unit",
        ],
        "filters": ["chair_code", "chair_id", "chair_model", "chair_variant"],
    },
    "chair_trials": {
        "default_columns": [
            "participant_id", "session_date", "chair_code", "trial_order",
            "overall_comfort_3_3", "discomfort_seat_3_3", "discomfort_back_3_3",
            "expected_vs_experienced_overall_3_3", "prolonged_willingness_3_3",
            "geometric_fit_1", "geometric_fit_2", "geometric_fit_3",
            "headrest_dependency", "interfere",
        ],
        "filters": [
            "participant_id", "session_date", "chair_code", "trial_order",
            "overall_comfort_3_3", "prolonged_willingness_3_3",
            "headrest_dependency", "interfere",
        ],
    },
}

JOIN_VIEWS = {
    "chair_trials_analysis": """
        SELECT
            ct.*,
            cl.chair_id, cl.chair_model, cl.chair_variant, cl.chair_name_display
        FROM chair_trials ct
        LEFT JOIN chair_lookup cl ON ct.chair_code = cl.chair_code
    """,
    "participant_overview": """
        SELECT
            p.participant_id,
            d.sex_assigned_at_birth,
            d.ethnicity,
            d.age_years,
            d.daily_physical_activity,
            d.recurring_chronic_pain,
            d.general_health_rating,
            d.ergonomic_familiarity,
            d.weekly_desk_hours_time,
            d.work_before_break_time,
            d.seat_continuously_change_freq,
            d.posture_most_used,
            d.posture_change_frequency,
            a.body_mass_kg,
            a.bmi_kg_m2,
            a.body_fat_percent,
            a.skeletal_muscle_percent,
            a.hip_breadth_sitting_cm,
            a.popliteal_height_sitting_cm,
            a.buttock_popliteal_length_cm,
            a.elbow_height_sitting_cm,
            a.sitting_height_erect_cm,
            pre.pre_selected_chair_code,
            pre.pre_selection_basis,
            pre.pre_selection_confidence_3_3,
            pre.pre_selection_basis_1,
            pre.pre_selection_basis_2,
            pre.pre_selection_basis_3,
            pre.pre_selection_basis_4,
            pre.pre_selection_basis_5,
            pre.pre_chair_familiarity_check_1,
            pre.pre_chair_familiarity_check_2,
            post.post_selected_chair_code,
            post.post_selection_basis,
            post.post_selection_confidence_3_3,
            post.post_selection_basis_1,
            post.post_selection_basis_2,
            post.post_selection_basis_3,
            post.post_selection_basis_4,
            post.post_selection_basis_5,
            post.selection_changed,
            post.selection_change_reason
        FROM participants p
        LEFT JOIN demographics_health d ON p.participant_id = d.participant_id
        LEFT JOIN anthropometry a ON p.participant_id = a.participant_id
        LEFT JOIN chair_selection_pre pre ON p.participant_id = pre.participant_id
        LEFT JOIN chair_selection_post post ON p.participant_id = post.participant_id
    """,
}

@st.cache_resource
def get_engine():
    try:
        engine = create_engine(DB_URL)
        # Test connection
        with engine.connect() as conn:
            conn.execute(text("SELECT 1"))
        return engine
    except Exception as e:
        st.error(f"❌ Database connection failed: {e}\n\nPlease ensure the ETL pipeline has been run and Neon is populated.")
        return None

@st.cache_data(show_spinner=False)
def get_columns(table_name: str) -> List[str]:
    engine = get_engine()
    if engine is None:
        # Return fallback columns from TABLE_CONFIG
        return TABLE_CONFIG.get(table_name, {}).get("default_columns", ["error: no data"])
    
    try:
        if table_name in JOIN_VIEWS:
            sql = text(f"SELECT * FROM ({JOIN_VIEWS[table_name]}) q LIMIT 0")
        else:
            sql = text(f'SELECT * FROM "{table_name}" LIMIT 0')
        with engine.connect() as conn:
            df = pd.read_sql(sql, conn)
        return list(df.columns)
    except Exception:
        # Fallback to configured columns
        return TABLE_CONFIG.get(table_name, {}).get("default_columns", [])

@st.cache_data(show_spinner=False)
def get_distinct_values(table_name: str, column_name: str, limit: int = 200):
    engine = get_engine()
    if engine is None:
        return []
    
    try:
        if table_name in JOIN_VIEWS:
            sql = text(f'''
                SELECT DISTINCT "{column_name}"
                FROM ({JOIN_VIEWS[table_name]}) q
                WHERE "{column_name}" IS NOT NULL
                ORDER BY 1
                LIMIT {limit}
            ''')
        else:
            sql = text(f'''
                SELECT DISTINCT "{column_name}"
                FROM "{table_name}"
                WHERE "{column_name}" IS NOT NULL
                ORDER BY 1
                LIMIT {limit}
            ''')
        with engine.connect() as conn:
            result = conn.execute(sql).fetchall()
        return [r[0] for r in result]
    except Exception:
        return []


def build_query(table_name: str, selected_columns: List[str], filters: dict, row_limit: int):
    where_clauses = []
    params = {}

    for idx, (col, meta) in enumerate(filters.items()):
        if meta["mode"] == "multiselect" and meta["value"]:
            placeholders = []
            for j, value in enumerate(meta["value"]):
                key = f"p_{idx}_{j}"
                placeholders.append(f":{key}")
                params[key] = value
            where_clauses.append(f'"{col}" IN ({", ".join(placeholders)})')
        elif meta["mode"] == "contains" and meta["value"]:
            key = f"p_{idx}"
            where_clauses.append(f'CAST("{col}" AS TEXT) ILIKE :{key}')
            params[key] = f'%{meta["value"]}%'
        elif meta["mode"] == "minmax":
            min_val = meta.get("min")
            max_val = meta.get("max")
            if min_val is not None:
                key = f"pmin_{idx}"
                where_clauses.append(f'"{col}" >= :{key}')
                params[key] = min_val
            if max_val is not None:
                key = f"pmax_{idx}"
                where_clauses.append(f'"{col}" <= :{key}')
                params[key] = max_val

    select_clause = ", ".join([f'"{c}"' for c in selected_columns]) if selected_columns else "*"
    where_sql = f"WHERE {' AND '.join(where_clauses)}" if where_clauses else ""

    if table_name in JOIN_VIEWS:
        sql = text(f'''
            SELECT {select_clause}
            FROM ({JOIN_VIEWS[table_name]}) q
            {where_sql}
            LIMIT {row_limit}
        ''')
    else:
        sql = text(f'''
            SELECT {select_clause}
            FROM "{table_name}"
            {where_sql}
            LIMIT {row_limit}
        ''')

    return sql, params


def dataframe_to_excel_bytes(df: pd.DataFrame) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="query_results", index=False)
    return output.getvalue()


st.title("FitCheck Query Portal")
st.caption("Filter FitCheck study data, preview results, and download CSV or Excel files.")

with st.sidebar:
    # st.header("Connection")
    # st.code(DB_URL.replace(DB_URL.split("@")[-1], "***@***") if "@" in DB_URL else DB_URL)
    st.header("Query Setup")
    data_source = st.selectbox(
        "Choose data source",
        options=[
            "participants",
            "demographics_health",
            "anthropometry",
            "chair_selection_pre",
            "chair_selection_post",
            "chair_lookup",
            "chair_feature_bounds_wide",
            "chair_trials",
            "participant_overview",
            "chair_trials_analysis",
        ],
    )
    st.markdown(
        "**Column scale notes**\n"
        "- `_3_3` columns: −3 to +3 (e.g. comfort, confidence, expected vs experienced)\n"
        "- `geometric_fit_*` columns: −3 to +3 (too small → too large)\n"
        "- `*_familiarity*` / `*_basis_*` columns: binary 0/1 (unchecked / checked)\n"
        "- Measurement columns: cm, kg, %, or minutes as indicated in the name"
    )

    all_columns = get_columns(data_source)
    default_columns = TABLE_CONFIG.get(data_source, {}).get("default_columns", all_columns[: min(8, len(all_columns))])
    selected_columns = st.multiselect("Columns to include", all_columns, default=default_columns)
    row_limit = st.number_input("Max rows to return", min_value=10, max_value=100000, value=1000, step=10)

    st.divider()
    st.header("Cleaned Export")
    if CLEANED_EXPORT_PATH.exists():
        _export_mtime = CLEANED_EXPORT_PATH.stat().st_mtime
        import datetime
        _export_ts = datetime.datetime.fromtimestamp(_export_mtime).strftime("%Y-%m-%d %H:%M")
        st.caption(f"Last generated: {_export_ts}")
        with open(CLEANED_EXPORT_PATH, "rb") as _f:
            st.download_button(
                label="Download fitcheck_cleaned_export.xlsx",
                data=_f.read(),
                file_name="fitcheck_cleaned_export.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
    else:
        st.warning(
            "fitcheck_cleaned_export.xlsx not found.\n\n"
            f"Expected at: `{CLEANED_EXPORT_PATH}`\n\n"
            "Run the ETL pipeline to generate it."
        )

st.subheader("Filters")
filterable_columns = TABLE_CONFIG.get(data_source, {}).get("filters", all_columns[: min(6, len(all_columns))])
filters = {}

cols = st.columns(2)
for i, column_name in enumerate(filterable_columns):
    with cols[i % 2]:
        st.markdown(f"**{column_name}**")
        mode = st.selectbox(
            f"Filter type for {column_name}",
            ["none", "multiselect", "contains", "minmax"],
            key=f"mode_{column_name}",
        )
        if mode == "multiselect":
            options = get_distinct_values(data_source, column_name)
            value = st.multiselect(f"Values for {column_name}", options, key=f"vals_{column_name}")
            filters[column_name] = {"mode": mode, "value": value}
        elif mode == "contains":
            value = st.text_input(f"Contains text for {column_name}", key=f"txt_{column_name}")
            filters[column_name] = {"mode": mode, "value": value}
        elif mode == "minmax":
            min_val = st.text_input(f"Min for {column_name}", key=f"min_{column_name}")
            max_val = st.text_input(f"Max for {column_name}", key=f"max_{column_name}")
            filters[column_name] = {
                "mode": mode,
                "min": None if min_val == "" else min_val,
                "max": None if max_val == "" else max_val,
            }

run_query = st.button("Run Query", type="primary")

if run_query:
    engine = get_engine()
    if engine is None:
        st.error("❌ Cannot run query: Database not connected. Please wait for the ETL pipeline to complete and refresh the page.")
    else:
        try:
            sql, params = build_query(data_source, selected_columns, filters, row_limit)
            with engine.connect() as conn:
                result_df = pd.read_sql(sql, conn, params=params)

            st.success(f"Query returned {len(result_df)} rows.")
            st.dataframe(result_df, use_container_width=True, height=500)

            c1, c2, c3 = st.columns(3)
            with c1:
                st.download_button(
                    label="Download CSV",
                    data=result_df.to_csv(index=False).encode("utf-8"),
                    file_name=f"{data_source}_query_results.csv",
                    mime="text/csv",
                )
            with c2:
                st.download_button(
                    label="Download Excel",
                    data=dataframe_to_excel_bytes(result_df),
                    file_name=f"{data_source}_query_results.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            with c3:
                full_sql_preview = str(sql)
                st.download_button(
                    label="Download SQL",
                    data=full_sql_preview.encode("utf-8"),
                    file_name=f"{data_source}_query.sql",
                    mime="text/plain",
                )

        except Exception as e:
            st.error(f"Query failed: {e}")

st.markdown("---")
# st.subheader("Suggested next improvements")
# st.write(
#     "Add login/authentication, save common queries, include charts for summary views, and offer cross-table participant filtering."
# )

# st.code(
#     """
# # Example run command
# streamlit run fitcheck_query_app.py

# # Example environment variable (Windows PowerShell)
# $env:POSTGRES_URL = 'postgresql+psycopg2://username:password@localhost:5432/fitcheck'
#     """.strip(),
#     language="bash",
# )
