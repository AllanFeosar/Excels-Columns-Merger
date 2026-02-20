import io

import pandas as pd


def file_to_bytes(uploaded_file) -> bytes:
    uploaded_file.seek(0)
    return uploaded_file.read()


def get_sheet_names(file_bytes: bytes) -> list[str]:
    with io.BytesIO(file_bytes) as buffer:
        excel = pd.ExcelFile(buffer)
    return [str(name) for name in excel.sheet_names]


def read_excel_sheet(file_bytes: bytes, sheet_name: str) -> pd.DataFrame:
    with io.BytesIO(file_bytes) as buffer:
        return pd.read_excel(buffer, sheet_name=sheet_name)


def to_excel_bytes(df: pd.DataFrame, sheet_name: str = "Merged_Result") -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    buffer.seek(0)
    return buffer.getvalue()
