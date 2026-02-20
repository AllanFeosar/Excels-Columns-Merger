import pandas as pd
import streamlit as st

from .io import file_to_bytes, get_sheet_names, read_excel_sheet, to_excel_bytes
from .matching import MatchResult, run_matching
from .presets import delete_preset, load_presets, upsert_preset
from .state import apply_pending_preset, build_preset_payload
from .ui import filter_results_by_status, render_column_picker


@st.cache_data(show_spinner=False)
def cached_sheet_names(file_bytes: bytes) -> list[str]:
    return get_sheet_names(file_bytes)


@st.cache_data(show_spinner=False)
def cached_read_sheet(file_bytes: bytes, sheet_name: str) -> pd.DataFrame:
    return read_excel_sheet(file_bytes, sheet_name)


def _render_uploaders() -> tuple[object | None, object | None]:
    upload_col_1, upload_col_2 = st.columns(2)
    with upload_col_1:
        left_file = st.file_uploader(
            "Upload Left Excel File",
            type=["xlsx", "xls", "xlsm", "xlsb"],
            key="left_file",
        )
    with upload_col_2:
        right_file = st.file_uploader(
            "Upload Right Excel File",
            type=["xlsx", "xls", "xlsm", "xlsb"],
            key="right_file",
        )
    return left_file, right_file


def _render_presets() -> None:
    presets = load_presets()
    preset_options = ["(None)"] + sorted(presets.keys())

    st.subheader("Preset")
    preset_col_1, preset_col_2, preset_col_3 = st.columns([2, 1, 1])
    selected_preset = preset_col_1.selectbox("Saved Preset", preset_options, key="selected_preset")
    if preset_col_2.button("Apply", key="apply_preset"):
        if selected_preset == "(None)":
            st.warning("Select a preset to apply.")
        else:
            st.session_state["pending_preset_name"] = selected_preset
            st.session_state["pending_preset_data"] = presets[selected_preset]
            st.rerun()
    if preset_col_3.button("Delete", key="delete_preset"):
        if selected_preset == "(None)":
            st.warning("Select a preset to delete.")
        else:
            delete_preset(selected_preset)
            st.session_state["preset_notice"] = f"Deleted preset: {selected_preset}"
            st.rerun()


def _render_run_merge(
    left_df: pd.DataFrame,
    right_df: pd.DataFrame,
    left_output_cols: list[str],
    right_output_cols: list[str],
    left_match_cols: list[str],
    right_match_cols: list[str],
    threshold: float,
    include_unmatched: bool,
    prefer_rapidfuzz: bool,
) -> None:
    if not left_output_cols:
        st.error("Select at least one left output column.")
        st.stop()
    if not right_output_cols:
        st.error("Select at least one right output column.")
        st.stop()

    similarity_requested = bool(left_match_cols and right_match_cols)
    action_word = "Matching" if similarity_requested else "Merging"

    if not similarity_requested:
        st.info("Similarity is disabled because similarity columns are not selected on both sides.")

    progress_bar = st.progress(0.0, text=f"{action_word} rows... 0/0")

    def on_progress(done_rows: int, total_rows: int) -> None:
        total = total_rows if total_rows > 0 else 1
        ratio = min(done_rows / total, 1.0)
        progress_bar.progress(ratio, text=f"{action_word} rows... {done_rows}/{total_rows}")

    match_result = run_matching(
        left_df=left_df,
        right_df=right_df,
        left_output_cols=left_output_cols,
        right_output_cols=right_output_cols,
        left_match_cols=left_match_cols,
        right_match_cols=right_match_cols,
        threshold=threshold,
        include_unmatched=include_unmatched,
        prefer_rapidfuzz=prefer_rapidfuzz,
        progress_callback=on_progress,
    )
    progress_bar.progress(1.0, text=f"{action_word} complete")
    _store_result_state(match_result, threshold)


def _store_result_state(match_result: MatchResult, threshold: float) -> None:
    st.session_state["last_result_df"] = match_result.result_df
    st.session_state["last_best_scores"] = match_result.best_scores
    st.session_state["last_threshold"] = threshold
    st.session_state["last_similarity_enabled"] = match_result.similarity_enabled
    st.session_state["last_similarity_engine"] = match_result.similarity_engine
    st.session_state["last_exact_match_count"] = match_result.exact_match_count
    st.session_state["last_candidate_comparisons"] = match_result.candidate_comparisons
    st.session_state["last_left_match_cols_used"] = match_result.left_match_columns_used
    st.session_state["last_right_match_cols_used"] = match_result.right_match_columns_used


def _render_results() -> None:
    result_df = st.session_state["last_result_df"]
    best_scores = st.session_state["last_best_scores"]
    threshold = st.session_state["last_threshold"]
    similarity_enabled = bool(st.session_state.get("last_similarity_enabled", True))
    similarity_engine = st.session_state.get("last_similarity_engine", "difflib")
    exact_match_count = int(st.session_state.get("last_exact_match_count", 0))
    candidate_comparisons = int(st.session_state.get("last_candidate_comparisons", 0))
    left_match_cols_used = st.session_state.get("last_left_match_cols_used", [])
    right_match_cols_used = st.session_state.get("last_right_match_cols_used", [])

    st.subheader("Result")
    if similarity_enabled:
        matched_count = sum(score >= threshold for score in best_scores)
        no_match_count = int((result_df["Match_Status"] == "No match").sum()) if "Match_Status" in result_df.columns else 0
        low_confidence = sum(threshold <= score < min(threshold + 0.05, 1.0) for score in best_scores)

        summary_col_1, summary_col_2, summary_col_3, summary_col_4, summary_col_5 = st.columns(5)
        summary_col_1.metric("Matched Left Rows", matched_count)
        summary_col_2.metric("No-Match Left Rows", no_match_count)
        summary_col_3.metric("Low-Confidence Matches", low_confidence)
        summary_col_4.metric("Exact Fast-Path", exact_match_count)
        summary_col_5.metric("Output Rows", len(result_df))

        st.caption(
            f"Similarity engine: `{similarity_engine}` | Candidate comparisons: `{candidate_comparisons}` | "
            f"Match columns used -> Left: `{', '.join(left_match_cols_used) if left_match_cols_used else '(none)'}` | "
            f"Right: `{', '.join(right_match_cols_used) if right_match_cols_used else '(none)'}` | "
            "Only one scale is used: `Similarity_Score`."
        )

        filter_mode = st.radio(
            "Rows to display",
            options=["All", "Matched only", "No match only"],
            horizontal=True,
            key="filter_mode",
        )
        filtered_result_df = filter_results_by_status(result_df, filter_mode)
    else:
        summary_col_1, summary_col_2 = st.columns(2)
        summary_col_1.metric("Output Rows", len(result_df))
        summary_col_2.metric("Right Rows Filled", int(result_df.filter(regex=r"^Right_").notna().any(axis=1).sum()))
        filtered_result_df = result_df.copy()
    if filtered_result_df.empty:
        st.warning("No rows for the selected filter.")

    st.dataframe(filtered_result_df, width="stretch", height=450)
    st.download_button(
        label="Download merged Excel",
        data=to_excel_bytes(filtered_result_df),
        file_name="merged_with_similarity.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


def render_app() -> None:
    st.set_page_config(page_title="Excel Columns Merger", layout="wide")
    st.title("Excel Columns Merger")
    st.caption("Select columns, run matching, filter rows, and export one merged Excel file.")

    notice = st.session_state.pop("preset_notice", None)
    if notice:
        st.success(notice)

    left_file, right_file = _render_uploaders()
    if not left_file or not right_file:
        st.info("Upload both files to start.")
        st.stop()

    left_bytes = file_to_bytes(left_file)
    right_bytes = file_to_bytes(right_file)

    try:
        left_sheets = cached_sheet_names(left_bytes)
        right_sheets = cached_sheet_names(right_bytes)
    except Exception as exc:
        st.error(f"Unable to read sheet names. Error: {exc}")
        st.stop()

    sheet_col_1, sheet_col_2 = st.columns(2)
    with sheet_col_1:
        left_sheet = st.selectbox("Left Sheet", left_sheets, key="left_sheet")
    with sheet_col_2:
        right_sheet = st.selectbox("Right Sheet", right_sheets, key="right_sheet")

    try:
        left_df = cached_read_sheet(left_bytes, left_sheet).copy()
        right_df = cached_read_sheet(right_bytes, right_sheet).copy()
    except Exception as exc:
        st.error(f"Unable to read selected sheets. Error: {exc}")
        st.stop()

    left_df.columns = left_df.columns.map(str)
    right_df.columns = right_df.columns.map(str)
    left_columns = left_df.columns.tolist()
    right_columns = right_df.columns.tolist()

    apply_pending_preset(
        left_sheets=left_sheets,
        right_sheets=right_sheets,
        left_sheet=left_sheet,
        right_sheet=right_sheet,
        left_columns=left_columns,
        right_columns=right_columns,
    )

    info_col_1, info_col_2 = st.columns(2)
    with info_col_1:
        st.metric("Left Rows", len(left_df))
        st.metric("Left Columns", len(left_columns))
    with info_col_2:
        st.metric("Right Rows", len(right_df))
        st.metric("Right Columns", len(right_columns))

    with st.expander("Preview Input Data", expanded=False):
        preview_col_1, preview_col_2 = st.columns(2)
        with preview_col_1:
            st.markdown("**Left Preview**")
            st.dataframe(left_df.head(10), width="stretch")
        with preview_col_2:
            st.markdown("**Right Preview**")
            st.dataframe(right_df.head(10), width="stretch")

    _render_presets()

    st.subheader("Select Columns")
    selectors_left, selectors_right = st.columns(2)
    with selectors_left:
        left_output_cols = render_column_picker(
            "Left columns in output",
            left_columns,
            "left_output",
            default_selected=left_columns[: min(3, len(left_columns))],
        )
        left_match_cols = render_column_picker(
            "Left columns for similarity (optional)",
            left_columns,
            "left_match",
            default_selected=left_columns[:1],
        )

    with selectors_right:
        right_output_cols = render_column_picker(
            "Right columns in output",
            right_columns,
            "right_output",
            default_selected=right_columns[: min(3, len(right_columns))],
        )
        right_match_cols = render_column_picker(
            "Right columns for similarity (optional)",
            right_columns,
            "right_match",
            default_selected=right_columns[:1],
        )

    similarity_requested = bool(left_match_cols and right_match_cols)

    st.subheader("Matching Settings")
    if similarity_requested:
        settings_col_1, settings_col_2, settings_col_3 = st.columns(3)
        with settings_col_1:
            threshold = st.slider(
                "Similarity threshold",
                min_value=0.0,
                max_value=1.0,
                value=0.75,
                step=0.01,
                key="threshold",
            )
        with settings_col_2:
            include_unmatched = st.checkbox("Include No match rows", value=True, key="include_unmatched")
        with settings_col_3:
            prefer_rapidfuzz = st.checkbox("Use rapidfuzz acceleration", value=True, key="prefer_rapidfuzz")
    else:
        st.info("Similarity settings are hidden. Select similarity columns on both sides to enable matching mode.")
        threshold = float(st.session_state.get("threshold", 0.75))
        include_unmatched = True
        prefer_rapidfuzz = bool(st.session_state.get("prefer_rapidfuzz", True))

    save_col_1, save_col_2 = st.columns([3, 1])
    preset_name_input = save_col_1.text_input("Preset name", key="preset_name_input", placeholder="e.g. default_setup")
    if save_col_2.button("Save preset", key="save_preset"):
        preset_name = preset_name_input.strip()
        if not preset_name:
            st.error("Enter a preset name.")
        else:
            payload = build_preset_payload(
                left_sheet=left_sheet,
                right_sheet=right_sheet,
                left_output_cols=left_output_cols,
                left_match_cols=left_match_cols,
                right_output_cols=right_output_cols,
                right_match_cols=right_match_cols,
                threshold=threshold,
                include_unmatched=include_unmatched,
                prefer_rapidfuzz=prefer_rapidfuzz,
                filter_mode=st.session_state.get("filter_mode", "All"),
            )
            upsert_preset(preset_name, payload)
            st.session_state["preset_notice"] = f"Saved preset: {preset_name}"
            st.rerun()

    run_merge = st.button("Run Merge", type="primary")
    if run_merge:
        _render_run_merge(
            left_df=left_df,
            right_df=right_df,
            left_output_cols=left_output_cols,
            right_output_cols=right_output_cols,
            left_match_cols=left_match_cols,
            right_match_cols=right_match_cols,
            threshold=threshold,
            include_unmatched=include_unmatched,
            prefer_rapidfuzz=prefer_rapidfuzz,
        )

    if "last_result_df" in st.session_state:
        _render_results()
