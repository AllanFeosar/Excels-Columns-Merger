import streamlit as st


def ensure_streamlit_context() -> None:
    if not st.runtime.exists():
        print("This is a Streamlit app.")
        print("Run it with: python -m streamlit run app.py")
        raise SystemExit(0)


def set_column_selection_state(columns: list[str], key_prefix: str, selected_cols: list[str]) -> None:
    selected_set = set(selected_cols)
    for idx, col in enumerate(columns):
        st.session_state[f"{key_prefix}_{idx}"] = col in selected_set


def apply_pending_preset(
    left_sheets: list[str],
    right_sheets: list[str],
    left_sheet: str,
    right_sheet: str,
    left_columns: list[str],
    right_columns: list[str],
) -> None:
    pending_data = st.session_state.pop("pending_preset_data", None)
    pending_name = st.session_state.pop("pending_preset_name", "")
    if not isinstance(pending_data, dict):
        return

    sheet_changed = False
    target_left_sheet = pending_data.get("left_sheet")
    if isinstance(target_left_sheet, str) and target_left_sheet in left_sheets and target_left_sheet != left_sheet:
        st.session_state["left_sheet"] = target_left_sheet
        sheet_changed = True

    target_right_sheet = pending_data.get("right_sheet")
    if isinstance(target_right_sheet, str) and target_right_sheet in right_sheets and target_right_sheet != right_sheet:
        st.session_state["right_sheet"] = target_right_sheet
        sheet_changed = True

    if sheet_changed:
        st.session_state["pending_preset_data"] = pending_data
        st.session_state["pending_preset_name"] = pending_name
        st.rerun()

    if isinstance(pending_data.get("left_output_cols"), list):
        set_column_selection_state(left_columns, "left_output", [str(x) for x in pending_data["left_output_cols"]])
    if isinstance(pending_data.get("left_match_cols"), list):
        set_column_selection_state(left_columns, "left_match", [str(x) for x in pending_data["left_match_cols"]])
    if isinstance(pending_data.get("right_output_cols"), list):
        set_column_selection_state(right_columns, "right_output", [str(x) for x in pending_data["right_output_cols"]])
    if isinstance(pending_data.get("right_match_cols"), list):
        set_column_selection_state(right_columns, "right_match", [str(x) for x in pending_data["right_match_cols"]])

    if isinstance(pending_data.get("threshold"), (int, float)):
        st.session_state["threshold"] = max(0.0, min(1.0, float(pending_data["threshold"])))
    if "include_unmatched" in pending_data:
        st.session_state["include_unmatched"] = bool(pending_data["include_unmatched"])
    if "prefer_rapidfuzz" in pending_data:
        st.session_state["prefer_rapidfuzz"] = bool(pending_data["prefer_rapidfuzz"])
    if pending_data.get("filter_mode") in {"All", "Matched only", "No match only"}:
        st.session_state["filter_mode"] = pending_data["filter_mode"]

    notice_name = pending_name if pending_name else "preset"
    st.session_state["preset_notice"] = f"Applied preset: {notice_name}"


def build_preset_payload(
    left_sheet: str,
    right_sheet: str,
    left_output_cols: list[str],
    left_match_cols: list[str],
    right_output_cols: list[str],
    right_match_cols: list[str],
    threshold: float,
    include_unmatched: bool,
    prefer_rapidfuzz: bool,
    filter_mode: str,
) -> dict:
    return {
        "left_sheet": left_sheet,
        "right_sheet": right_sheet,
        "left_output_cols": left_output_cols,
        "left_match_cols": left_match_cols,
        "right_output_cols": right_output_cols,
        "right_match_cols": right_match_cols,
        "threshold": float(threshold),
        "include_unmatched": bool(include_unmatched),
        "prefer_rapidfuzz": bool(prefer_rapidfuzz),
        "filter_mode": filter_mode,
    }
