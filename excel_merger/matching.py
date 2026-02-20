import re
from collections import defaultdict
from dataclasses import dataclass
from difflib import SequenceMatcher
from typing import Callable

import pandas as pd

try:
    from rapidfuzz import fuzz as rapidfuzz_fuzz
except ModuleNotFoundError:
    rapidfuzz_fuzz = None


@dataclass
class MatchResult:
    result_df: pd.DataFrame
    best_scores: list[float]
    best_positions: list[int | None]
    left_text: pd.Series
    right_text: pd.Series
    similarity_enabled: bool
    similarity_engine: str
    exact_match_count: int
    candidate_comparisons: int
    left_match_columns_used: list[str]
    right_match_columns_used: list[str]


def normalize_text(value: object) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip().lower()
    return re.sub(r"\s+", " ", text)


def combine_columns(df: pd.DataFrame, selected_cols: list[str]) -> pd.Series:
    if not selected_cols:
        return pd.Series([""] * len(df), index=df.index, dtype="object")
    return (
        df[selected_cols]
        .fillna("")
        .astype(str)
        .agg(" ".join, axis=1)
        .map(normalize_text)
    )


def build_token_index(text_values: list[str]) -> dict[str, set[int]]:
    token_index: dict[str, set[int]] = defaultdict(set)
    for pos, text in enumerate(text_values):
        for token in text.split():
            token_index[token].add(pos)
    return token_index


def build_exact_index(text_values: list[str]) -> dict[str, list[int]]:
    exact_index: dict[str, list[int]] = defaultdict(list)
    for pos, text in enumerate(text_values):
        if text:
            exact_index[text].append(pos)
    return exact_index


def _similarity_ratio(left_text: str, right_text: str, use_rapidfuzz: bool) -> float:
    if use_rapidfuzz and rapidfuzz_fuzz is not None:
        return float(rapidfuzz_fuzz.ratio(left_text, right_text) / 100.0)
    return SequenceMatcher(None, left_text, right_text).ratio()


def run_matching(
    left_df: pd.DataFrame,
    right_df: pd.DataFrame,
    left_output_cols: list[str],
    right_output_cols: list[str],
    left_match_cols: list[str],
    right_match_cols: list[str],
    threshold: float,
    include_unmatched: bool,
    prefer_rapidfuzz: bool = True,
    progress_callback: Callable[[int, int], None] | None = None,
    progress_update_every: int | None = None,
) -> MatchResult:
    similarity_enabled = bool(left_match_cols and right_match_cols)
    effective_left_match_cols = left_match_cols if similarity_enabled else []
    effective_right_match_cols = right_match_cols if similarity_enabled else []

    left_text = combine_columns(left_df, effective_left_match_cols)
    right_text = combine_columns(right_df, effective_right_match_cols)

    left_output_records = left_df[left_output_cols].to_dict("records")
    right_output_records = right_df[right_output_cols].to_dict("records")

    if not similarity_enabled:
        rows: list[dict[str, object]] = []
        total_rows = len(left_output_records)
        if progress_callback is not None:
            progress_callback(0, total_rows)
        if progress_update_every is None:
            progress_update_every = max(total_rows // 200, 1)

        for left_pos in range(total_rows):
            row: dict[str, object] = {}
            for col, value in left_output_records[left_pos].items():
                row[f"Left_{col}"] = value

            has_right_row = left_pos < len(right_output_records)
            if has_right_row:
                for col, value in right_output_records[left_pos].items():
                    row[f"Right_{col}"] = value
            else:
                for col in right_output_cols:
                    row[f"Right_{col}"] = pd.NA

            if has_right_row or include_unmatched:
                rows.append(row)

            if progress_callback is not None:
                done_rows = left_pos + 1
                if done_rows % progress_update_every == 0 or done_rows == total_rows:
                    progress_callback(done_rows, total_rows)

        return MatchResult(
            result_df=pd.DataFrame(rows),
            best_scores=[],
            best_positions=[],
            left_text=left_text,
            right_text=right_text,
            similarity_enabled=False,
            similarity_engine="disabled",
            exact_match_count=0,
            candidate_comparisons=0,
            left_match_columns_used=[],
            right_match_columns_used=[],
        )

    left_text_values = left_text.tolist()
    right_text_values = right_text.tolist()
    token_index = build_token_index(right_text_values)
    exact_index = build_exact_index(right_text_values)
    use_rapidfuzz = bool(prefer_rapidfuzz and rapidfuzz_fuzz is not None)
    similarity_engine = "rapidfuzz" if use_rapidfuzz else "difflib"

    rows: list[dict[str, object]] = []
    best_scores: list[float] = []
    best_positions: list[int | None] = []
    exact_match_count = 0
    candidate_comparisons = 0

    total_rows = len(left_text_values)
    if progress_callback is not None:
        progress_callback(0, total_rows)
    if progress_update_every is None:
        progress_update_every = max(total_rows // 200, 1)

    for left_pos, ltext in enumerate(left_text_values):
        best_score = 0.0
        best_right_pos: int | None = None

        if ltext in exact_index:
            best_right_pos = exact_index[ltext][0]
            best_score = 1.0
            exact_match_count += 1
        else:
            ltokens = set(ltext.split())
            if ltokens:
                candidate_positions: set[int] = set()
                for token in ltokens:
                    candidate_positions.update(token_index.get(token, set()))

                for right_pos in candidate_positions:
                    candidate_comparisons += 1
                    score = _similarity_ratio(ltext, right_text_values[right_pos], use_rapidfuzz)
                    if score > best_score:
                        best_score = score
                        best_right_pos = right_pos

        best_scores.append(best_score)
        best_positions.append(best_right_pos)

        matched = best_right_pos is not None and best_score >= threshold
        if matched or include_unmatched:
            row: dict[str, object] = {}
            for col, value in left_output_records[left_pos].items():
                row[f"Left_{col}"] = value

            if matched and best_right_pos is not None:
                for col, value in right_output_records[best_right_pos].items():
                    row[f"Right_{col}"] = value
            else:
                for col in right_output_cols:
                    row[f"Right_{col}"] = pd.NA

            row["Similarity_Score"] = round(best_score, 4)
            row["Match_Status"] = "Matched" if matched else "No match"
            rows.append(row)

        if progress_callback is not None:
            done_rows = left_pos + 1
            if done_rows % progress_update_every == 0 or done_rows == total_rows:
                progress_callback(done_rows, total_rows)

    return MatchResult(
        result_df=pd.DataFrame(rows),
        best_scores=best_scores,
        best_positions=best_positions,
        left_text=left_text,
        right_text=right_text,
        similarity_enabled=True,
        similarity_engine=similarity_engine,
        exact_match_count=exact_match_count,
        candidate_comparisons=candidate_comparisons,
        left_match_columns_used=effective_left_match_cols,
        right_match_columns_used=effective_right_match_cols,
    )
