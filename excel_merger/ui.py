import pandas as pd
import streamlit as st


def render_column_picker(
    title: str,
    columns: list[str],
    key_prefix: str,
    default_selected: list[str],
) -> list[str]:
    st.markdown(f"**{title}**")
    action_col_1, action_col_2 = st.columns(2)
    if action_col_1.button("Select all", key=f"{key_prefix}_select_all"):
        for idx in range(len(columns)):
            st.session_state[f"{key_prefix}_{idx}"] = True
    if action_col_2.button("Clear all", key=f"{key_prefix}_clear_all"):
        for idx in range(len(columns)):
            st.session_state[f"{key_prefix}_{idx}"] = False

    selected: list[str] = []
    per_row = 4
    for row_start in range(0, len(columns), per_row):
        row = st.columns(per_row)
        for idx, col in enumerate(columns[row_start : row_start + per_row]):
            key = f"{key_prefix}_{row_start + idx}"
            checked = row[idx].checkbox(col, value=col in default_selected, key=key)
            if checked:
                selected.append(col)
    return selected


def filter_results_by_status(result_df: pd.DataFrame, filter_mode: str) -> pd.DataFrame:
    if "Match_Status" not in result_df.columns:
        return result_df
    if filter_mode == "Matched only":
        return result_df[result_df["Match_Status"] == "Matched"].copy()
    if filter_mode == "No match only":
        return result_df[result_df["Match_Status"] == "No match"].copy()
    return result_df.copy()
