from pathlib import Path

app_code = r'''import io
import re
import time
import zipfile
import tempfile
from pathlib import Path
from datetime import datetime, date, time as dt_time

import numpy as np
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils.datetime import to_excel


# ============================================================
# APP CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="Learning Agility Assessment",
    page_icon="📊",
    layout="wide",
)

APP_DIR = Path(__file__).resolve().parent
MASTER_FILE = APP_DIR / "templates" / "master.xlsx"

DIMENSIONS = {
    "mental_agility": "mental agility",
    "people_agility": "people agility",
    "change_agility": "change agility",
    "result_agility": "result agility",
    "self_awareness": "self awareness",
}

IDENTITY_INPUT_CELLS = [
    "D2",   # Nama
    "D3",   # Nama Atasan
    "D4",   # Level Saat Ini
    "D5",   # Level Tujuan
    "D6",   # Tanggal Assessment
    "D7",   # Unit
    "D8",   # Departemen
    "D9",   # Divisi
    "D10",  # Direktorat
    "D12",  # Level Next Path
]

ASSESSMENT_INPUT_COLUMNS = {
    "self_rating": 9,       # I
    "self_comment": 10,     # J
    "superior_rating": 11,  # K
    "superior_comment": 12, # L
}

ASSESSMENT_ROWS = list(range(13, 23))
ASSESSMENT_SHEETS = list(DIMENSIONS.values())

# Hanya cell berikut yang boleh diubah oleh Python.
# Semua cell lain harus tetap berasal dari master.
ALLOWED_OUTPUT_CELLS = {
    "Identitas": set(IDENTITY_INPUT_CELLS),
}

for _sheet in ASSESSMENT_SHEETS:
    ALLOWED_OUTPUT_CELLS[_sheet] = {
        f"{column_letter}{row}"
        for row in ASSESSMENT_ROWS
        for column_letter in ("I", "J", "K", "L")
    }


# ============================================================
# HELPERS
# ============================================================

def normalize_text(value):
    """Normalisasi teks untuk pencocokan yang konsisten."""
    if pd.isna(value):
        return ""

    value = str(value)
    value = value.replace("\xa0", " ")
    value = value.replace("\n", " ")
    value = value.replace("\r", " ")
    value = re.sub(r"\s+", " ", value)

    return value.strip().lower()


def clean_value(value):
    """Membersihkan value dari Excel. NaN / None -> string kosong."""
    if pd.isna(value):
        return ""

    return str(value).strip()


def clean_excel_value(value):
    """
    Mengubah NaN/empty menjadi None dan Timestamp menjadi
    datetime Python.
    """
    if value is None:
        return None

    try:
        if pd.isna(value):
            return None
    except Exception:
        pass

    if isinstance(value, pd.Timestamp):
        return value.to_pydatetime()

    if isinstance(value, np.generic):
        return value.item()

    if isinstance(value, str):
        value = value.strip()

        if value == "":
            return None

    return value


def find_column(df, candidates):
    """Mencari nama kolom berdasarkan nama yang sudah dinormalisasi."""
    if df is None or df.empty:
        return None

    normalized_columns = {
        normalize_text(col): col
        for col in df.columns
    }

    for candidate in candidates:
        key = normalize_text(candidate)

        if key in normalized_columns:
            return normalized_columns[key]

    return None


def safe_filename(name):
    """Membuat nama file aman untuk Windows/Linux."""
    value = re.sub(r'[\\/*?:"<>|]', "_", str(name)).strip()

    if not value:
        value = "Unknown"

    return value[:150]


# ============================================================
# MASTER PROCESSING
# ============================================================

def extract_indicator_master(master_wb):
    indicator_master = []

    for dimension_key, sheet_name in DIMENSIONS.items():

        if sheet_name not in master_wb.sheetnames:
            raise ValueError(
                f"Sheet master tidak ditemukan: '{sheet_name}'"
            )

        ws = master_wb[sheet_name]

        for indicator_no, row in enumerate(
            range(13, 23),
            start=1,
        ):
            indicator = ws.cell(row, 2).value
            question = ws.cell(row, 3).value

            levels = {
                1: ws.cell(row, 4).value,
                2: ws.cell(row, 5).value,
                3: ws.cell(row, 6).value,
                4: ws.cell(row, 7).value,
                5: ws.cell(row, 8).value,
            }

            indicator_master.append({
                "dimension": dimension_key,
                "sheet_name": sheet_name,
                "indicator_no": indicator_no,
                "indicator_id": f"{dimension_key}_{indicator_no:02d}",
                "indicator": indicator,
                "question": question,
                "question_normalized": normalize_text(question),
                "level_1": levels[1],
                "level_2": levels[2],
                "level_3": levels[3],
                "level_4": levels[4],
                "level_5": levels[5],
                "level_1_normalized": normalize_text(levels[1]),
                "level_2_normalized": normalize_text(levels[2]),
                "level_3_normalized": normalize_text(levels[3]),
                "level_4_normalized": normalize_text(levels[4]),
                "level_5_normalized": normalize_text(levels[5]),
            })

    indicator_master_df = pd.DataFrame(indicator_master)

    if len(indicator_master_df) != 50:
        raise ValueError(
            "Master harus memiliki 50 indikator. "
            f"Ditemukan {len(indicator_master_df)}."
        )

    if indicator_master_df["indicator_id"].nunique() != 50:
        raise ValueError("Duplicate indicator_id ditemukan.")

    if indicator_master_df["question_normalized"].nunique() != 50:
        raise ValueError("Duplicate questions ditemukan.")

    for col in [
        "level_1",
        "level_2",
        "level_3",
        "level_4",
        "level_5",
    ]:
        missing = indicator_master_df[col].isna().sum()

        if missing:
            raise ValueError(
                f"{col} memiliki {missing} nilai kosong."
            )

    return indicator_master_df


def detect_question_columns(df, master_questions):
    master_question_set = {
        normalize_text(q)
        for q in master_questions
        if pd.notna(q)
    }

    question_columns = []

    for col in df.columns:
        if normalize_text(col) in master_question_set:
            question_columns.append(col)

    return question_columns


def build_answer_mapping(indicator_master_df):
    answer_mapping = {}

    for _, row in indicator_master_df.iterrows():

        question_key = row["question_normalized"]

        answer_mapping[question_key] = {
            normalize_text(row["level_1"]): 1,
            normalize_text(row["level_2"]): 2,
            normalize_text(row["level_3"]): 3,
            normalize_text(row["level_4"]): 4,
            normalize_text(row["level_5"]): 5,
        }

    if len(answer_mapping) != 50:
        raise ValueError(
            "Answer mapping hanya memiliki "
            f"{len(answer_mapping)} question."
        )

    return answer_mapping


# ============================================================
# RAW -> SCORED DATA
# ============================================================

def process_assessment(
    df,
    question_columns,
    rater_type,
    indicator_master_df,
    answer_mapping,
):
    results = []

    for _, row in df.iterrows():

        row_dict = {
            normalize_text(k): v
            for k, v in row.items()
        }

        if rater_type == "Self":
            employee_name = clean_value(
                row.get(
                    "Nama",
                    row_dict.get("nama", ""),
                )
            )

            superior_name = ""

        else:
            employee_name = clean_value(
                row.get(
                    "Nama bawahan yang dinilai",
                    row_dict.get(
                        "nama bawahan yang dinilai",
                        "",
                    ),
                )
            )

            superior_name = clean_value(
                row.get(
                    "Nama",
                    row_dict.get("nama", ""),
                )
            )

        email = clean_value(
            row.get(
                "Email",
                row_dict.get("email", ""),
            )
        )

        division = clean_value(
            row.get(
                "Divisi",
                row_dict.get("divisi", ""),
            )
        )

        unit = clean_value(
            row.get(
                "Unit",
                row_dict.get("unit", ""),
            )
        )

        department = clean_value(
            row.get(
                "Departemen",
                row_dict.get("departemen", ""),
            )
        )

        directorate = clean_value(
            row.get(
                "Direktorat",
                row_dict.get("direktorat", ""),
            )
        )

        current_level = clean_value(
            row.get(
                "Level Saat Ini",
                row_dict.get("level saat ini", ""),
            )
        )

        target_level = clean_value(
            row.get(
                "Level Tujuan",
                row_dict.get("level tujuan", ""),
            )
        )

        next_path = clean_value(
            row.get(
                "Level Next Path",
                row_dict.get("level next path", ""),
            )
        )

        completion_time = row.get(
            "Completion time",
            row_dict.get("completion time", None),
        )

        last_modified_time = row.get(
            "Last modified time",
            row_dict.get("last modified time", None),
        )

        for question_col in question_columns:

            question_normalized = normalize_text(question_col)
            answer = row.get(question_col, "")

            if pd.isna(answer):
                continue

            answer_str = str(answer).strip()

            if not answer_str:
                continue

            master_rows = indicator_master_df[
                indicator_master_df["question_normalized"]
                == question_normalized
            ]

            if master_rows.empty:
                results.append({
                    "rater_type": rater_type,
                    "employee_name": employee_name,
                    "superior_name": superior_name,
                    "email": email,
                    "division": division,
                    "unit": unit,
                    "department": department,
                    "directorate": directorate,
                    "current_level": current_level,
                    "target_level": target_level,
                    "next_path": next_path,
                    "completion_time": completion_time,
                    "last_modified_time": last_modified_time,
                    "question": question_col,
                    "answer_text": answer_str,
                    "rating": np.nan,
                    "mapping_status": "QUESTION_NOT_FOUND",
                    "indicator_id": "",
                    "dimension": "",
                    "indicator_no": np.nan,
                    "indicator": "",
                })
                continue

            master_row = master_rows.iloc[0]

            rating = answer_mapping[
                question_normalized
            ].get(
                normalize_text(answer_str),
                np.nan,
            )

            mapping_status = (
                "ANSWER_NOT_FOUND"
                if pd.isna(rating)
                else "MATCHED"
            )

            results.append({
                "rater_type": rater_type,
                "employee_name": employee_name,
                "superior_name": superior_name,
                "email": email,
                "division": division,
                "unit": unit,
                "department": department,
                "directorate": directorate,
                "current_level": current_level,
                "target_level": target_level,
                "next_path": next_path,
                "completion_time": completion_time,
                "last_modified_time": last_modified_time,
                "question": question_col,
                "answer_text": answer_str,
                "rating": rating,
                "mapping_status": mapping_status,
                "indicator_id": master_row["indicator_id"],
                "dimension": master_row["dimension"],
                "indicator_no": master_row["indicator_no"],
                "indicator": master_row["indicator"],
            })

    return pd.DataFrame(results)


def calculate_scores(
    self_scored_df,
    superior_scored_df,
    indicator_master_df,
):
    scored_answers_df = pd.concat(
        [
            self_scored_df,
            superior_scored_df,
        ],
        ignore_index=True,
    )

    total_records = len(scored_answers_df)

    matched_records = (
        scored_answers_df[
            scored_answers_df["mapping_status"] == "MATCHED"
        ].shape[0]
    )

    unmatched_records = total_records - matched_records

    coverage = (
        matched_records / total_records
        if total_records > 0
        else 0
    )

    if (
        not scored_answers_df.empty
        and scored_answers_df["employee_name"]
        .fillna("")
        .astype(str)
        .str.strip()
        .eq("")
        .any()
    ):
        raise ValueError(
            "Ada record dengan employee_name kosong."
        )

    indicator_score_df = (
        scored_answers_df[
            scored_answers_df["mapping_status"] == "MATCHED"
        ]
        .groupby(
            [
                "rater_type",
                "employee_name",
                "indicator_id",
            ],
            as_index=False,
        )["rating"]
        .mean()
        .rename(
            columns={
                "rating": "indicator_score",
            }
        )
    )

    indicator_score_df = indicator_score_df.merge(
        indicator_master_df[
            [
                "indicator_id",
                "dimension",
                "indicator_no",
                "indicator",
            ]
        ],
        on="indicator_id",
        how="left",
    )

    dimension_score_df = (
        indicator_score_df
        .groupby(
            [
                "rater_type",
                "employee_name",
                "dimension",
            ],
            as_index=False,
        )["indicator_score"]
        .mean()
        .rename(
            columns={
                "indicator_score": "dimension_score",
            }
        )
    )

    dimension_gap_df = (
        dimension_score_df
        .pivot_table(
            index=[
                "employee_name",
                "dimension",
            ],
            columns="rater_type",
            values="dimension_score",
        )
        .reset_index()
    )

    if "Self" not in dimension_gap_df.columns:
        dimension_gap_df["Self"] = np.nan

    if "Superior" not in dimension_gap_df.columns:
        dimension_gap_df["Superior"] = np.nan

    dimension_gap_df["gap"] = (
        dimension_gap_df["Self"]
        - dimension_gap_df["Superior"]
    )

    return (
        scored_answers_df,
        indicator_score_df,
        dimension_score_df,
        dimension_gap_df,
        total_records,
        matched_records,
        unmatched_records,
        coverage,
    )


# ============================================================
# EXACT MASTER PRESERVATION
# ============================================================

def _excel_value_xml(value):
    """
    Membuat XML value untuk cell target.

    Penting:
    - Tidak menyentuh style cell.
    - Tidak menyentuh formula.
    - Hanya mengganti isi cell target.
    """
    value = clean_excel_value(value)

    if value is None:
        return ""

    if isinstance(value, bool):
        return f"<v>{1 if value else 0}</v>"

    if isinstance(value, (int, float, np.integer, np.floating)):
        if isinstance(value, float) and np.isnan(value):
            return ""

        return f"<v>{value}</v>"

    if isinstance(value, (datetime, date)):
        serial = to_excel(value)
        return f"<v>{serial}</v>"

    if isinstance(value, dt_time):
        fraction = (
            value.hour * 3600
            + value.minute * 60
            + value.second
            + value.microsecond / 1_000_000
        ) / 86400

        return f"<v>{fraction}</v>"

    text_value = str(value)

    # Excel XML inline string.
    escaped = (
        text_value
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )

    if (
        text_value.startswith(" ")
        or text_value.endswith(" ")
        or "\n" in text_value
        or "\r" in text_value
        or "\t" in text_value
    ):
        preserve = ' xml:space="preserve"'
    else:
        preserve = ""

    return (
        f'<is><t{preserve}>{escaped}</t></is>'
    )


def _patch_cell_xml(cell_xml, value):
    """
    Patch satu cell XML tanpa mengubah atribut cell.

    Contoh:
    <c r="D2" s="15"><v>old</v></c>

    menjadi:
    <c r="D2" s="15" t="inlineStr"><is><t>new</t></is></c>

    atau angka:
    <c r="I13" s="20"><v>4</v></c>

    Atribut style (s="...") tetap dipertahankan.
    """

    # Ambil opening <c ...>
    opening_match = re.match(
        rb"(<c\b[^>]*)(>)",
        cell_xml,
    )

    if not opening_match:
        raise ValueError(
            f"Format cell XML tidak dikenali: {cell_xml[:200]!r}"
        )

    opening = opening_match.group(1)
    closing = opening_match.group(2)

    # Hapus type lama jika ada.
    opening = re.sub(
        rb'\s+t="[^"]*"',
        b"",
        opening,
    )

    value = clean_excel_value(value)

    if value is None:
        # Cell kosong: pertahankan cell node dan style,
        # tetapi hilangkan value.
        new_opening = opening
        new_body = b""

    elif isinstance(value, bool):
        new_opening = opening
        new_body = (
            b"<v>"
            + (b"1" if value else b"0")
            + b"</v>"
        )

    elif isinstance(value, (int, float, np.integer, np.floating)):
        if isinstance(value, float) and np.isnan(value):
            new_opening = opening
            new_body = b""
        else:
            new_opening = opening
            number_text = str(value).encode("utf-8")
            new_body = b"<v>" + number_text + b"</v>"

    elif isinstance(value, (datetime, date)):
        new_opening = opening
        number_text = str(to_excel(value)).encode("utf-8")
        new_body = b"<v>" + number_text + b"</v>"

    elif isinstance(value, dt_time):
        new_opening = opening
        fraction = (
            value.hour * 3600
            + value.minute * 60
            + value.second
            + value.microsecond / 1_000_000
        ) / 86400

        new_body = (
            b"<v>"
            + str(fraction).encode("utf-8")
            + b"</v>"
        )

    else:
        text_value = str(value)
        escaped = (
            text_value
            .replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
        )

        preserve = (
            ' xml:space="preserve"'
            if (
                text_value.startswith(" ")
                or text_value.endswith(" ")
                or "\n" in text_value
                or "\r" in text_value
                or "\t" in text_value
            )
            else ""
        )

        new_opening = opening + b' t="inlineStr"'

        new_body = (
            f'<is><t{preserve}>{escaped}</t></is>'
            .encode("utf-8")
        )

    return (
        new_opening
        + closing
        + new_body
        + b"</c>"
    )


def _patch_sheet_xml(
    xml_bytes,
    cell_values,
):
    """
    Mengubah hanya cell target pada sheet XML.

    Tidak melakukan parse + reserialize seluruh XML.
    Ini penting karena kita ingin perubahan sekecil mungkin
    terhadap master.
    """

    for coordinate, value in cell_values.items():

        coordinate_bytes = coordinate.encode("utf-8")

        # Cell normal:
        # <c r="D2"...>...</c>
        pattern = re.compile(
            rb"<c\b(?=[^>]*\br=\""
            + re.escape(coordinate_bytes)
            + rb"\")[^>]*>.*?</c>",
            re.DOTALL,
        )

        match = pattern.search(xml_bytes)

        if match:
            old_cell = match.group(0)
            new_cell = _patch_cell_xml(
                old_cell,
                value,
            )

            xml_bytes = (
                xml_bytes[:match.start()]
                + new_cell
                + xml_bytes[match.end():]
            )

            continue

        # Jika cell target benar-benar kosong dan tidak mempunyai
        # node <c>, kita perlu membuat node baru.
        # Ini tetap dilakukan hanya pada target cell.
        row_number_match = re.match(
            r"([A-Z]+)(\d+)$",
            coordinate,
        )

        if not row_number_match:
            raise ValueError(
                f"Coordinate Excel tidak valid: {coordinate}"
            )

        row_number = row_number_match.group(2)

        row_pattern = re.compile(
            rb"<row\b(?=[^>]*\br=\""
            + re.escape(row_number.encode("utf-8"))
            + rb"\")[^>]*>.*?</row>",
            re.DOTALL,
        )

        row_match = row_pattern.search(xml_bytes)

        if not row_match:
            raise ValueError(
                f"Row {row_number} tidak ditemukan "
                "di worksheet XML."
            )

        row_xml = row_match.group(0)

        # Untuk cell yang belum ada, kita tambahkan node tanpa
        # style karena memang tidak ada cell master sebelumnya.
        # Dalam master yang normal, input cells seharusnya sudah
        # memiliki style/node.
        new_cell = (
            f'<c r="{coordinate}" '
            f't="inlineStr">'
            f'</c>'
        ).encode("utf-8")

        # Sisipkan sebelum </row>.
        new_row_xml = (
            row_xml[:-6]
            + new_cell
            + b"</row>"
        )

        xml_bytes = (
            xml_bytes[:row_match.start()]
            + new_row_xml
            + xml_bytes[row_match.end():]
        )

    return xml_bytes


def get_sheet_xml_paths(xlsx_path):
    """
    Mendapatkan mapping nama sheet Excel -> path XML worksheet
    berdasarkan workbook.xml + relationships.

    Tidak mengubah file.
    """
    with zipfile.ZipFile(
        xlsx_path,
        "r",
    ) as zf:

        workbook_xml = zf.read(
            "xl/workbook.xml"
        )

        rels_xml = zf.read(
            "xl/_rels/workbook.xml.rels"
        )

    workbook_match = re.search(
        rb"<workbook[^>]*xmlns=\"([^\"]+)\"",
        workbook_xml,
    )

    if workbook_match:
        workbook_xml_clean = workbook_xml
    else:
        workbook_xml_clean = workbook_xml

    # Extract sheet name + relationship ID.
    sheet_pattern = re.compile(
        rb"<sheet\b[^>]*\bname=\"([^\"]+)\""
        rb"[^>]*\br:id=\"([^\"]+)\"[^>]*/?>"
    )

    sheets = sheet_pattern.findall(
        workbook_xml_clean
    )

    # Extract relationship ID + target.
    rel_pattern = re.compile(
        rb"<Relationship\b[^>]*\bId=\"([^\"]+)\""
        rb"[^>]*\bTarget=\"([^\"]+)\"[^>]*/?>"
    )

    relationships = {
        rel_id.decode("utf-8"): target.decode("utf-8")
        for rel_id, target in rel_pattern.findall(rels_xml)
    }

    result = {}

    for sheet_name_bytes, rel_id_bytes in sheets:
        sheet_name = sheet_name_bytes.decode(
            "utf-8"
        )
        rel_id = rel_id_bytes.decode(
            "utf-8"
        )

        target = relationships.get(rel_id)

        if target is None:
            raise ValueError(
                f"Relationship sheet '{sheet_name}' "
                f"tidak ditemukan."
            )

        if target.startswith("/"):
            target = target.lstrip("/")

        elif not target.startswith("xl/"):
            target = "xl/" + target

        result[sheet_name] = target

    return result


def patch_master_exactly(
    master_file,
    output_file,
    changes_by_sheet,
):
    """
    CORE FUNCTION.

    Membuat output berdasarkan ZIP master dan hanya mengganti
    value pada cell yang memang diizinkan.

    Semua entry ZIP lain disalin dari master tanpa disentuh.
    Ini jauh lebih aman untuk mempertahankan:
    - formula
    - style
    - merged cells
    - chart
    - image
    - drawing
    - conditional formatting
    - named range
    - workbook settings
    - print settings
    - page layout
    - hidden sheets
    - metadata
    - dan struktur XML lain

    dibanding membuka lalu save ulang seluruh workbook dengan
    openpyxl.
    """

    sheet_paths = get_sheet_xml_paths(
        master_file
    )

    # Pastikan semua target sheet tersedia.
    for sheet_name in changes_by_sheet:

        if sheet_name not in sheet_paths:
            raise ValueError(
                f"Sheet '{sheet_name}' tidak ditemukan "
                "di master."
            )

    with zipfile.ZipFile(
        master_file,
        "r",
    ) as source_zip:

        with zipfile.ZipFile(
            output_file,
            "w",
            compression=zipfile.ZIP_DEFLATED,
            compresslevel=6,
        ) as target_zip:

            for item in source_zip.infolist():

                data = source_zip.read(
                    item.filename
                )

                # Hanya worksheet XML target yang dipatch.
                affected_sheets = [
                    sheet_name
                    for sheet_name, sheet_path
                    in sheet_paths.items()
                    if sheet_path == item.filename
                    and sheet_name in changes_by_sheet
                ]

                if affected_sheets:
                    for sheet_name in affected_sheets:
                        data = _patch_sheet_xml(
                            data,
                            changes_by_sheet[sheet_name],
                        )

                target_zip.writestr(
                    item,
                    data,
                )


# ============================================================
# BUILD EMPLOYEE CHANGES
# ============================================================

def build_employee_changes(
    employee_name,
    scored_answers,
    df_self,
    df_superior,
):
    """
    Menghasilkan daftar cell yang memang boleh berubah.

    Fungsi ini TIDAK menyentuh workbook.
    """

    target_employee = normalize_text(
        employee_name
    )

    employee_data = scored_answers[
        scored_answers["employee_name"]
        .apply(normalize_text)
        == target_employee
    ].copy()

    if employee_data.empty:
        raise ValueError(
            f"Tidak ada data assessment untuk employee: "
            f"{employee_name}"
        )

    changes = {
        "Identitas": {},
    }

    for sheet_name in ASSESSMENT_SHEETS:
        changes[sheet_name] = {}

    # --------------------------------------------------------
    # IDENTITAS
    # --------------------------------------------------------

    changes["Identitas"]["D2"] = employee_name

    self_raw = pd.DataFrame()

    if df_self is not None and not df_self.empty:

        self_name_col = find_column(
            df_self,
            [
                "Nama",
                "Name",
            ],
        )

        if self_name_col is not None:
            self_raw = df_self[
                df_self[self_name_col]
                .apply(normalize_text)
                == target_employee
            ].copy()

    # --------------------------------------------------------
    # NAMA ATASAN
    # --------------------------------------------------------

    superior_name = None

    if (
        df_superior is not None
        and not df_superior.empty
    ):

        subordinate_col = find_column(
            df_superior,
            [
                "Nama bawahan yang dinilai",
                "Nama Bawahan",
                "Bawahan",
            ],
        )

        if subordinate_col is not None:

            superior_employee_data = df_superior[
                df_superior[subordinate_col]
                .apply(normalize_text)
                == target_employee
            ].copy()

            if not superior_employee_data.empty:

                superior_name_col = find_column(
                    df_superior,
                    [
                        "Nama",
                        "Name",
                        "Nama atasan yang menilai",
                        "Nama atasan",
                        "Nama superior",
                        "Superior Name",
                    ],
                )

                if superior_name_col is not None:
                    superior_name = clean_excel_value(
                        superior_employee_data.iloc[0][
                            superior_name_col
                        ]
                    )

    changes["Identitas"]["D3"] = superior_name

    # --------------------------------------------------------
    # IDENTITY MAPPING
    # --------------------------------------------------------

    identity_mapping = {
        "D4": [
            "Level Saat Ini",
            "Level saat ini",
            "Current Level",
        ],
        "D5": [
            "Level Tujuan",
            "Level tujuan",
            "Target Level",
        ],
        "D6": [
            "Tanggal Assesment",
            "Tanggal Assessment",
            "Assessment Date",
            "Completion time",
            "Start time",
            "Last modified time",
        ],
        "D7": [
            "Unit",
        ],
        "D8": [
            "Departemen",
            "Department",
        ],
        "D9": [
            "Divisi",
            "Division",
        ],
        "D10": [
            "Direktorat",
            "Directorate",
        ],
        "D12": [
            "Level Next Path",
            "Next Path",
            "Level Next",
        ],
    }

    if not self_raw.empty:

        self_row = self_raw.iloc[0]

        for target_cell, candidates in identity_mapping.items():

            source_col = find_column(
                self_raw,
                candidates,
            )

            if source_col is None:
                continue

            value = clean_excel_value(
                self_row[source_col]
            )

            if value is None:
                continue

            changes["Identitas"][
                target_cell
            ] = value

    # --------------------------------------------------------
    # 5 DIMENSIONS
    # --------------------------------------------------------

    for dimension_key, sheet_name in DIMENSIONS.items():

        dimension_data = employee_data[
            employee_data["dimension"]
            == dimension_key
        ].copy()

        for indicator_no, row in enumerate(
            ASSESSMENT_ROWS,
            start=1,
        ):

            indicator_id = (
                f"{dimension_key}_{indicator_no:02d}"
            )

            indicator_data = dimension_data[
                dimension_data["indicator_id"]
                == indicator_id
            ].copy()

            # ------------------------------------------------
            # SELF
            # ------------------------------------------------

            self_data = indicator_data[
                indicator_data["rater_type"] == "Self"
            ].copy()

            self_rating = None

            if not self_data.empty:

                ratings = pd.to_numeric(
                    self_data["rating"],
                    errors="coerce",
                ).dropna()

                if not ratings.empty:
                    self_rating = ratings.iloc[0]

            changes[sheet_name][
                f"I{row}"
            ] = self_rating

            self_comment = None

            if not self_data.empty:

                self_answers = (
                    self_data["answer_text"]
                    .dropna()
                    .astype(str)
                    .str.strip()
                )

                self_answers = self_answers[
                    self_answers != ""
                ]

                if not self_answers.empty:
                    self_comment = "\n".join(
                        self_answers.tolist()
                    )

            changes[sheet_name][
                f"J{row}"
            ] = self_comment

            # ------------------------------------------------
            # SUPERIOR
            # ------------------------------------------------

            superior_data = indicator_data[
                indicator_data["rater_type"]
                == "Superior"
            ].copy()

            superior_rating = None

            if not superior_data.empty:

                ratings = pd.to_numeric(
                    superior_data["rating"],
                    errors="coerce",
                ).dropna()

                if not ratings.empty:
                    superior_rating = ratings.iloc[0]

            changes[sheet_name][
                f"K{row}"
            ] = superior_rating

            superior_comment = None

            if not superior_data.empty:

                superior_answers = (
                    superior_data["answer_text"]
                    .dropna()
                    .astype(str)
                    .str.strip()
                )

                superior_answers = superior_answers[
                    superior_answers != ""
                ]

                if not superior_answers.empty:
                    superior_comment = "\n".join(
                        superior_answers.tolist()
                    )

            changes[sheet_name][
                f"L{row}"
            ] = superior_comment

    return changes


# ============================================================
# VALIDATION
# ============================================================

def get_master_structure(master_file):
    wb = load_workbook(
        master_file,
        data_only=False,
        read_only=False,
    )

    try:
        structure = {}

        for sheet_name in wb.sheetnames:

            ws = wb[sheet_name]

            structure[sheet_name] = {
                "max_row": ws.max_row,
                "max_column": ws.max_column,
                "merged_ranges": tuple(
                    sorted(
                        str(rng)
                        for rng in ws.merged_cells.ranges
                    )
                ),
            }

        return structure

    finally:
        wb.close()


def get_formula_snapshot(file_path):
    wb = load_workbook(
        file_path,
        data_only=False,
        read_only=False,
    )

    try:
        formulas = {}

        for sheet_name in wb.sheetnames:

            ws = wb[sheet_name]
            sheet_formulas = {}

            for row in ws.iter_rows():

                for cell in row:

                    if (
                        isinstance(cell.value, str)
                        and cell.value.startswith("=")
                    ):
                        sheet_formulas[
                            cell.coordinate
                        ] = cell.value

            formulas[sheet_name] = sheet_formulas

        return formulas

    finally:
        wb.close()


def normalize_formula(formula):
    if not isinstance(formula, str):
        return formula

    formula = formula.replace("\n", "")
    formula = formula.replace("\r", "")
    formula = re.sub(r"\s+", "", formula)

    return formula


def validate_formula_preservation(
    output_file,
    master_formulas,
):
    output_formulas = get_formula_snapshot(
        output_file
    )

    if set(output_formulas.keys()) != set(
        master_formulas.keys()
    ):
        return {
            "status": "FAILED",
            "reason": "Formula sheet structure changed.",
        }

    differences = []

    for sheet_name in master_formulas:

        master_sheet = master_formulas[
            sheet_name
        ]

        output_sheet = output_formulas[
            sheet_name
        ]

        if len(master_sheet) != len(
            output_sheet
        ):
            return {
                "status": "FAILED",
                "reason": (
                    f"{sheet_name}: formula count changed. "
                    f"Master={len(master_sheet)}, "
                    f"Output={len(output_sheet)}"
                ),
                "differences": differences,
            }

        for coordinate, master_formula in (
            master_sheet.items()
        ):

            output_formula = output_sheet.get(
                coordinate
            )

            if output_formula is None:
                differences.append({
                    "sheet": sheet_name,
                    "cell": coordinate,
                    "master": master_formula,
                    "output": None,
                })
                continue

            if (
                normalize_formula(master_formula)
                != normalize_formula(output_formula)
            ):
                differences.append({
                    "sheet": sheet_name,
                    "cell": coordinate,
                    "master": master_formula,
                    "output": output_formula,
                })

    if differences:
        return {
            "status": "FAILED",
            "reason": (
                f"Ditemukan {len(differences)} "
                "formula yang berbeda."
            ),
            "differences": differences,
        }

    return {
        "status": "PASSED",
        "reason": "All master formulas preserved.",
        "differences": [],
    }


def validate_output_structure(
    output_file,
    master_structure,
):
    wb = load_workbook(
        output_file,
        data_only=False,
        read_only=False,
    )

    try:

        if wb.sheetnames != list(
            master_structure.keys()
        ):
            return {
                "status": "FAILED",
                "reason": "Sheet structure changed.",
            }

        for sheet_name, master_info in (
            master_structure.items()
        ):

            ws = wb[sheet_name]

            if ws.max_row != master_info["max_row"]:
                return {
                    "status": "FAILED",
                    "reason": (
                        f"{sheet_name}: max_row changed."
                    ),
                }

            if (
                ws.max_column
                != master_info["max_column"]
            ):
                return {
                    "status": "FAILED",
                    "reason": (
                        f"{sheet_name}: "
                        "max_column changed."
                    ),
                }

            output_merged = tuple(
                sorted(
                    str(rng)
                    for rng in ws.merged_cells.ranges
                )
            )

            if (
                output_merged
                != master_info["merged_ranges"]
            ):
                return {
                    "status": "FAILED",
                    "reason": (
                        f"{sheet_name}: "
                        "merged cells changed."
                    ),
                }

        return {
            "status": "PASSED",
            "reason": "Structure preserved.",
        }

    finally:
        wb.close()


def get_input_font_snapshot(file_path):
    wb = load_workbook(
        file_path,
        data_only=False,
        read_only=False,
    )

    result = {}

    try:

        ws = wb["Identitas"]

        for cell_address in IDENTITY_INPUT_CELLS:

            cell = ws[cell_address]

            result[
                ("Identitas", cell_address)
            ] = {
                "font_name": cell.font.name,
                "font_size": cell.font.sz,
                "bold": cell.font.bold,
                "italic": cell.font.italic,
                "underline": cell.font.underline,
                "strike": cell.font.strike,
            }

        for sheet_name in ASSESSMENT_SHEETS:

            ws = wb[sheet_name]

            for row in ASSESSMENT_ROWS:

                for column_number in (
                    ASSESSMENT_INPUT_COLUMNS.values()
                ):

                    cell = ws.cell(
                        row=row,
                        column=column_number,
                    )

                    result[
                        (
                            sheet_name,
                            cell.coordinate,
                        )
                    ] = {
                        "font_name": cell.font.name,
                        "font_size": cell.font.sz,
                        "bold": cell.font.bold,
                        "italic": cell.font.italic,
                        "underline": cell.font.underline,
                        "strike": cell.font.strike,
                    }

        return result

    finally:
        wb.close()


def validate_input_font_preservation(
    output_file,
    master_fonts,
):
    output_fonts = get_input_font_snapshot(
        output_file
    )

    differences = []

    if set(output_fonts.keys()) != set(
        master_fonts.keys()
    ):
        return {
            "status": "FAILED",
            "reason": "Jumlah / daftar input cell berubah.",
            "differences": [],
        }

    for key, master_font in (
        master_fonts.items()
    ):

        output_font = output_fonts.get(key)

        if output_font != master_font:
            differences.append({
                "sheet": key[0],
                "cell": key[1],
                "master": master_font,
                "output": output_font,
            })

    if differences:
        return {
            "status": "FAILED",
            "reason": (
                f"Ditemukan {len(differences)} "
                "input cell dengan font berbeda."
            ),
            "differences": differences,
        }

    return {
        "status": "PASSED",
        "reason": (
            "Semua font input cell sesuai dengan master."
        ),
        "differences": [],
    }


def validate_only_allowed_cells_changed(
    master_file,
    output_file,
    allowed_cells,
):
    """
    Validasi paling penting.

    Membandingkan value/formula setiap cell master vs output.

    Perbedaan hanya boleh terjadi pada cell yang memang
    dikerjakan Python.
    """

    master_wb = load_workbook(
        master_file,
        data_only=False,
        read_only=False,
    )

    output_wb = load_workbook(
        output_file,
        data_only=False,
        read_only=False,
    )

    differences = []

    try:

        if master_wb.sheetnames != output_wb.sheetnames:
            return {
                "status": "FAILED",
                "reason": "Daftar sheet berubah.",
                "differences": [],
            }

        for sheet_name in master_wb.sheetnames:

            master_ws = master_wb[sheet_name]
            output_ws = output_wb[sheet_name]

            max_row = max(
                master_ws.max_row,
                output_ws.max_row,
            )

            max_col = max(
                master_ws.max_column,
                output_ws.max_column,
            )

            allowed = allowed_cells.get(
                sheet_name,
                set(),
            )

            for row in range(
                1,
                max_row + 1,
            ):
                for col in range(
                    1,
                    max_col + 1,
                ):

                    master_value = master_ws.cell(
                        row=row,
                        column=col,
                    ).value

                    output_value = output_ws.cell(
                        row=row,
                        column=col,
                    ).value

                    if master_value != output_value:

                        coordinate = (
                            output_ws.cell(
                                row=row,
                                column=col,
                            ).coordinate
                        )

                        if coordinate not in allowed:

                            differences.append({
                                "sheet": sheet_name,
                                "cell": coordinate,
                                "master": master_value,
                                "output": output_value,
                            })

        if differences:
            return {
                "status": "FAILED",
                "reason": (
                    f"Ada {len(differences)} cell "
                    "di luar area yang diizinkan berubah."
                ),
                "differences": differences[:100],
            }

        return {
            "status": "PASSED",
            "reason": (
                "Tidak ada cell di luar area Python "
                "yang berubah."
            ),
            "differences": [],
        }

    finally:
        master_wb.close()
        output_wb.close()


def validate_employee_output(
    output_file,
    employee_name,
):
    wb = load_workbook(
        output_file,
        data_only=False,
        read_only=False,
    )

    try:

        ws_identity = wb["Identitas"]
        errors = []

        output_name = ws_identity["D2"].value

        if normalize_text(output_name) != normalize_text(
            employee_name
        ):
            errors.append(
                "Identitas!D2 tidak sesuai employee."
            )

        for sheet_name in ASSESSMENT_SHEETS:

            if sheet_name not in wb.sheetnames:
                errors.append(
                    f"Sheet {sheet_name} tidak ditemukan."
                )

        if errors:
            return {
                "status": "FAILED",
                "errors": errors,
            }

        return {
            "status": "PASSED",
            "errors": [],
        }

    finally:
        wb.close()


# ============================================================
# INITIALIZE MASTER
# ============================================================

@st.cache_resource
def load_master_metadata():

    if not MASTER_FILE.exists():
        raise FileNotFoundError(
            "Template master tidak ditemukan. "
            f"Letakkan file Excel master di:\n{MASTER_FILE}"
        )

    wb = load_workbook(
        MASTER_FILE,
        data_only=False,
    )

    indicator_master_df = extract_indicator_master(
        wb
    )

    wb.close()

    master_structure = get_master_structure(
        MASTER_FILE
    )

    master_formulas = get_formula_snapshot(
        MASTER_FILE
    )

    master_fonts = get_input_font_snapshot(
        MASTER_FILE
    )

    answer_mapping = build_answer_mapping(
        indicator_master_df
    )

    return (
        indicator_master_df,
        answer_mapping,
        master_structure,
        master_formulas,
        master_fonts,
    )


# ============================================================
# GENERATE ALL EMPLOYEE FILES
# ============================================================

def generate_all_files(
    df_self,
    df_superior,
    master_file,
    indicator_master_df,
    answer_mapping,
    master_structure,
    master_formulas,
    master_fonts,
    validation_enabled=True,
):

    master_questions = indicator_master_df[
        "question"
    ].tolist()

    # --------------------------------------------------------
    # DETECT QUESTIONS
    # --------------------------------------------------------

    self_question_cols = detect_question_columns(
        df_self,
        master_questions,
    )

    superior_question_cols = detect_question_columns(
        df_superior,
        master_questions,
    )

    if len(self_question_cols) != 50:
        raise ValueError(
            "Self Assessment memiliki "
            f"{len(self_question_cols)} pertanyaan yang cocok "
            "dengan master. Expected: 50."
        )

    if len(superior_question_cols) != 50:
        raise ValueError(
            "Superior Assessment memiliki "
            f"{len(superior_question_cols)} pertanyaan yang cocok "
            "dengan master. Expected: 50."
        )

    # --------------------------------------------------------
    # PROCESS SELF
    # --------------------------------------------------------

    self_scored_df = process_assessment(
        df=df_self,
        question_columns=self_question_cols,
        rater_type="Self",
        indicator_master_df=indicator_master_df,
        answer_mapping=answer_mapping,
    )

    # --------------------------------------------------------
    # PROCESS SUPERIOR
    # --------------------------------------------------------

    superior_scored_df = process_assessment(
        df=df_superior,
        question_columns=superior_question_cols,
        rater_type="Superior",
        indicator_master_df=indicator_master_df,
        answer_mapping=answer_mapping,
    )

    # --------------------------------------------------------
    # COMBINE + SCORE
    # --------------------------------------------------------

    (
        scored_answers_df,
        indicator_score_df,
        dimension_score_df,
        dimension_gap_df,
        total_records,
        matched_records,
        unmatched_records,
        coverage,
    ) = calculate_scores(
        self_scored_df,
        superior_scored_df,
        indicator_master_df,
    )

    # --------------------------------------------------------
    # EMPLOYEE LIST
    # --------------------------------------------------------

    employees = sorted(
        scored_answers_df[
            scored_answers_df["rater_type"] == "Self"
        ]["employee_name"]
        .dropna()
        .astype(str)
        .str.strip()
        .loc[lambda s: s != ""]
        .unique()
    )

    if len(employees) == 0:
        raise ValueError(
            "Tidak ada employee dari Self Assessment."
        )

    generated_files = []
    failed_files = []

    temp_dir = Path(
        tempfile.mkdtemp(
            prefix="learning_agility_"
        )
    )

    output_dir = temp_dir / "results"

    output_dir.mkdir(
        parents=True,
        exist_ok=True,
    )

    total_employees = len(employees)

    for index, employee_name in enumerate(
        employees,
        start=1,
    ):

        employee_start = time.time()

        output_file = (
            output_dir
            / (
                "Learning_Agility_"
                f"{safe_filename(employee_name)}.xlsx"
            )
        )

        try:

            # ------------------------------------------------
            # BUILD ONLY THE CHANGES
            # ------------------------------------------------

            changes = build_employee_changes(
                employee_name=employee_name,
                scored_answers=scored_answers_df,
                df_self=df_self,
                df_superior=df_superior,
            )

            # ------------------------------------------------
            # COPY MASTER + PATCH TARGET CELLS ONLY
            # ------------------------------------------------

            patch_master_exactly(
                master_file=master_file,
                output_file=output_file,
                changes_by_sheet=changes,
            )

            if (
                not output_file.exists()
                or output_file.stat().st_size == 0
            ):
                raise IOError(
                    "Output Excel gagal dibuat atau kosong."
                )

            # ------------------------------------------------
            # VALIDATION
            # ------------------------------------------------

            if validation_enabled:

                structure_result = (
                    validate_output_structure(
                        output_file,
                        master_structure,
                    )
                )

                if (
                    structure_result["status"]
                    != "PASSED"
                ):
                    raise ValueError(
                        "Structure validation failed: "
                        + structure_result["reason"]
                    )

                formula_result = (
                    validate_formula_preservation(
                        output_file,
                        master_formulas,
                    )
                )

                if (
                    formula_result["status"]
                    != "PASSED"
                ):
                    raise ValueError(
                        "Formula validation failed: "
                        + formula_result["reason"]
                    )

                changed_cells_result = (
                    validate_only_allowed_cells_changed(
                        master_file,
                        output_file,
                        ALLOWED_OUTPUT_CELLS,
                    )
                )

                if (
                    changed_cells_result["status"]
                    != "PASSED"
                ):
                    raise ValueError(
                        "Unexpected workbook change: "
                        + changed_cells_result["reason"]
                    )

                employee_result = (
                    validate_employee_output(
                        output_file,
                        employee_name,
                    )
                )

                if (
                    employee_result["status"]
                    != "PASSED"
                ):
                    raise ValueError(
                        "Employee validation failed: "
                        + str(
                            employee_result["errors"]
                        )
                    )

                font_result = (
                    validate_input_font_preservation(
                        output_file,
                        master_fonts,
                    )
                )

                if (
                    font_result["status"]
                    != "PASSED"
                ):
                    raise ValueError(
                        "Input font validation failed: "
                        + font_result["reason"]
                    )

            generated_files.append(
                output_file
            )

        except Exception as exc:

            failed_files.append({
                "employee": employee_name,
                "error": repr(exc),
                "processing_time": (
                    f"{time.time() - employee_start:.2f} sec"
                ),
            })

        yield {
            "index": index,
            "total": total_employees,
            "employee": employee_name,
            "generated_files": generated_files,
            "failed_files": failed_files,
            "scored_answers_df": scored_answers_df,
            "indicator_score_df": indicator_score_df,
            "dimension_score_df": dimension_score_df,
            "dimension_gap_df": dimension_gap_df,
            "total_records": total_records,
            "matched_records": matched_records,
            "unmatched_records": unmatched_records,
            "coverage": coverage,
            "temp_dir": temp_dir,
        }


# ============================================================
# ZIP
# ============================================================

def create_zip(generated_files):
    zip_buffer = io.BytesIO()

    with zipfile.ZipFile(
        zip_buffer,
        mode="w",
        compression=zipfile.ZIP_DEFLATED,
    ) as zip_file:

        for file_path in generated_files:

            zip_file.write(
                file_path,
                arcname=file_path.name,
            )

    zip_buffer.seek(0)

    return zip_buffer.getvalue()


# ============================================================
# UI
# ============================================================

st.title(
    "📊 Learning Agility Assessment System"
)

st.write(
    "Upload hasil **Self Assessment** dan "
    "**Superior Assessment**. Sistem akan memproses "
    "jawaban, melakukan scoring berdasarkan master, "
    "mengisi hanya cell yang memang menjadi area input "
    "Python pada template Excel, lalu menggabungkan "
    "seluruh hasil employee ke dalam satu file ZIP."
)

if not MASTER_FILE.exists():

    st.error(
        "Template master belum ditemukan."
    )

    st.code(
        str(MASTER_FILE),
        language="text",
    )

    st.info(
        "Buat folder `templates` di lokasi yang sama "
        "dengan app.py, lalu simpan template Excel "
        "sebagai `master.xlsx`."
    )

    st.stop()


try:

    (
        indicator_master_df,
        answer_mapping,
        master_structure,
        master_formulas,
        master_fonts,
    ) = load_master_metadata()

except Exception as exc:

    st.error(
        "Master template gagal dibaca."
    )

    st.exception(exc)

    st.stop()


with st.sidebar:

    st.header("Configuration")

    st.success(
        f"Master OK — "
        f"{len(indicator_master_df)} indicators"
    )

    st.write(
        f"• Dimensions: {len(DIMENSIONS)}"
    )

    st.write(
        "• Indicators: 50"
    )

    st.write(
        "• Ratings: 1–5"
    )

    st.write(
        "• Output mode: exact master patch"
    )

    validation_enabled = st.checkbox(
        "Enable output validation",
        value=True,
        help=(
            "Validasi structure, formula, employee data, "
            "font, dan memastikan tidak ada cell di luar "
            "area Python yang berubah."
        ),
    )


col1, col2 = st.columns(2)

with col1:

    self_file = st.file_uploader(
        "1. Upload Self Assessment",
        type=["xlsx"],
        key="self_upload",
    )

with col2:

    superior_file = st.file_uploader(
        "2. Upload Superior Assessment",
        type=["xlsx"],
        key="superior_upload",
    )


process_button = st.button(
    "🚀 Process Assessment",
    type="primary",
    use_container_width=True,
    disabled=(
        self_file is None
        or superior_file is None
    ),
)


if process_button:

    progress_bar = st.progress(0)
    status_box = st.empty()

    try:

        # ----------------------------------------------------
        # LOAD UPLOADED FILES
        # ----------------------------------------------------

        status_box.info(
            "Membaca Self Assessment..."
        )

        df_self = pd.read_excel(
            io.BytesIO(
                self_file.getvalue()
            ),
            engine="openpyxl",
        )

        status_box.info(
            "Membaca Superior Assessment..."
        )

        df_superior = pd.read_excel(
            io.BytesIO(
                superior_file.getvalue()
            ),
            engine="openpyxl",
        )

        st.session_state["input_shapes"] = {
            "self": df_self.shape,
            "superior": df_superior.shape,
        }

        # ----------------------------------------------------
        # PROCESS
        # ----------------------------------------------------

        status_box.info(
            "Memproses scoring dan membuat file employee..."
        )

        final_state = None

        generator = generate_all_files(
            df_self=df_self,
            df_superior=df_superior,
            master_file=MASTER_FILE,
            indicator_master_df=indicator_master_df,
            answer_mapping=answer_mapping,
            master_structure=master_structure,
            master_formulas=master_formulas,
            master_fonts=master_fonts,
            validation_enabled=validation_enabled,
        )

        for state in generator:

            final_state = state

            progress = (
                state["index"]
                / state["total"]
            )

            progress_bar.progress(
                progress
            )

            status_box.info(
                f"Processing "
                f"{state['index']}/{state['total']}: "
                f"{state['employee']}"
            )

        if final_state is None:
            raise RuntimeError(
                "Tidak ada hasil processing."
            )

        # ----------------------------------------------------
        # RESULT SUMMARY
        # ----------------------------------------------------

        generated_files = final_state[
            "generated_files"
        ]

        failed_files = final_state[
            "failed_files"
        ]

        progress_bar.progress(1.0)

        status_box.success(
            "Processing selesai."
        )

        # ----------------------------------------------------
        # KPI
        # ----------------------------------------------------

        k1, k2, k3, k4 = st.columns(4)

        k1.metric(
            "Employees",
            final_state["total"],
        )

        k2.metric(
            "Generated",
            len(generated_files),
        )

        k3.metric(
            "Failed",
            len(failed_files),
        )

        k4.metric(
            "Mapping Coverage",
            f"{final_state['coverage']:.2%}",
        )

        # ----------------------------------------------------
        # MAPPING STATUS
        # ----------------------------------------------------

        st.subheader(
            "Processing Summary"
        )

        st.write(
            f"Total records: "
            f"**{final_state['total_records']:,}**"
        )

        st.write(
            f"Matched records: "
            f"**{final_state['matched_records']:,}**"
        )

        st.write(
            f"Unmatched records: "
            f"**{final_state['unmatched_records']:,}**"
        )

        # ----------------------------------------------------
        # FAILED EMPLOYEES
        # ----------------------------------------------------

        if failed_files:

            st.warning(
                f"{len(failed_files)} employee gagal dibuat."
            )

            st.dataframe(
                pd.DataFrame(failed_files),
                use_container_width=True,
            )

        # ----------------------------------------------------
        # ZIP DOWNLOAD
        # ----------------------------------------------------

        if generated_files:

            status_box.info(
                "Membuat ZIP hasil..."
            )

            zip_bytes = create_zip(
                generated_files
            )

            st.download_button(
                label=(
                    "📦 Download All Results (.ZIP)"
                ),
                data=zip_bytes,
                file_name=(
                    "Learning_Agility_Results.zip"
                ),
                mime="application/zip",
                type="primary",
                use_container_width=True,
            )

            st.success(
                f"{len(generated_files)} file Excel "
                "berhasil dibuat dan siap di-download."
            )

        # ----------------------------------------------------
        # DIMENSION SCORE PREVIEW
        # ----------------------------------------------------

        st.subheader(
            "Dimension Score Preview"
        )

        preview_df = final_state[
            "dimension_score_df"
        ].copy()

        if not preview_df.empty:

            preview_df["dimension_score"] = (
                preview_df["dimension_score"]
                .round(2)
            )

            st.dataframe(
                preview_df,
                use_container_width=True,
            )

        # ----------------------------------------------------
        # GAP PREVIEW
        # ----------------------------------------------------

        st.subheader(
            "Self vs Superior Gap Preview"
        )

        gap_df = final_state[
            "dimension_gap_df"
        ].copy()

        if not gap_df.empty:

            for col in [
                "Self",
                "Superior",
                "gap",
            ]:

                if col in gap_df.columns:

                    gap_df[col] = (
                        pd.to_numeric(
                            gap_df[col],
                            errors="coerce",
                        )
                        .round(2)
                    )

            st.dataframe(
                gap_df,
                use_container_width=True,
            )

    except Exception as exc:

        progress_bar.empty()

        status_box.error(
            "Processing gagal."
        )

        st.exception(exc)


# ============================================================
# FOOTER
# ============================================================

st.divider()

st.caption(
    "Learning Agility Assessment System — "
    "Python + Streamlit + Pandas + OpenPyXL"
)
'''

out = Path("/mnt/data/app.py")
out.write_text(app_code, encoding="utf-8")

# Syntax validation
import py_compile
py_compile.compile(str(out), doraise=True)

print(f"Created: {out}")
print(f"Lines: {len(app_code.splitlines())}")
print("Syntax check: PASSED")
