from __future__ import annotations

import io
import zipfile

import pandas as pd

from fastapi import (
    FastAPI,
    File,
    UploadFile,
    HTTPException,
)

from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse

from .processor import generate_reports


# ============================================================
# APP
# ============================================================

app = FastAPI(
    title="Learning Agility API",
    version="1.0.0",
)


# ============================================================
# CORS
# ============================================================

app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "http://localhost:5173",
        "http://127.0.0.1:5173",
    ],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


# ============================================================
# ROOT
# ============================================================

@app.get("/api")
def root():
    return {
        "status": "ok",
        "service": "Learning Agility API",
    }


# ============================================================
# HEALTH CHECK
# ============================================================

@app.get("/api/health")
def health():
    return {
        "status": "healthy",
    }


# ============================================================
# READ EXCEL
# ============================================================

def read_excel_upload(
    file_bytes: bytes,
    filename: str,
):
    """
    Membaca file Excel upload menjadi pandas DataFrame.

    .xlsx -> openpyxl
    .xls  -> xlrd
    """

    filename = filename.lower()

    if filename.endswith(".xlsx"):
        return pd.read_excel(
            io.BytesIO(file_bytes),
            engine="openpyxl",
        )

    if filename.endswith(".xls"):
        return pd.read_excel(
            io.BytesIO(file_bytes),
            engine="xlrd",
        )

    raise ValueError(
        f"Format file tidak didukung: {filename}"
    )


# ============================================================
# GENERATE
# ============================================================

@app.post("/api/generate")
async def generate(
    self_assessment: UploadFile = File(...),
    superior_assessment: UploadFile = File(...),
):

    # ========================================================
    # VALIDATE EXTENSION
    # ========================================================

    self_filename = (
        self_assessment.filename or ""
    ).lower()

    superior_filename = (
        superior_assessment.filename or ""
    ).lower()

    allowed_extensions = (
        ".xlsx",
        ".xls",
    )

    # ========================================================
    # VALIDATE SELF
    # ========================================================

    if not self_filename.endswith(
        allowed_extensions
    ):
        raise HTTPException(
            status_code=400,
            detail=(
                "Self Assessment harus "
                "berupa file Excel (.xlsx atau .xls)."
            ),
        )

    # ========================================================
    # VALIDATE SUPERIOR
    # ========================================================

    if not superior_filename.endswith(
        allowed_extensions
    ):
        raise HTTPException(
            status_code=400,
            detail=(
                "Superior Assessment harus "
                "berupa file Excel (.xlsx atau .xls)."
            ),
        )

    try:

        # ====================================================
        # READ UPLOAD
        # ====================================================

        self_bytes = await self_assessment.read()

        superior_bytes = await (
            superior_assessment.read()
        )

        # ====================================================
        # VALIDATE EMPTY FILE
        # ====================================================

        if not self_bytes:
            raise HTTPException(
                status_code=400,
                detail="Self Assessment kosong.",
            )

        if not superior_bytes:
            raise HTTPException(
                status_code=400,
                detail="Superior Assessment kosong.",
            )

        # ====================================================
        # EXCEL -> DATAFRAME
        # ====================================================

        try:

            df_self = read_excel_upload(
                file_bytes=self_bytes,
                filename=self_filename,
            )

        except Exception as exc:

            raise HTTPException(
                status_code=400,
                detail={
                    "message": (
                        "Self Assessment "
                        "tidak dapat dibaca."
                    ),
                    "error": repr(exc),
                },
            )

        try:

            df_superior = read_excel_upload(
                file_bytes=superior_bytes,
                filename=superior_filename,
            )

        except Exception as exc:

            raise HTTPException(
                status_code=400,
                detail={
                    "message": (
                        "Superior Assessment "
                        "tidak dapat dibaca."
                    ),
                    "error": repr(exc),
                },
            )

        # ====================================================
        # VALIDATE DATAFRAME
        # ====================================================

        if df_self.empty:
            raise HTTPException(
                status_code=400,
                detail=(
                    "Self Assessment tidak "
                    "memiliki data."
                ),
            )

        if df_superior.empty:
            raise HTTPException(
                status_code=400,
                detail=(
                    "Superior Assessment tidak "
                    "memiliki data."
                ),
            )

        # ====================================================
        # PROCESS
        # ====================================================

        result = generate_reports(
            df_self=df_self,
            df_superior=df_superior,
        )

        # ====================================================
        # GET GENERATED FILES
        # ====================================================

        generated_files = result.get(
            "generated",
            [],
        )

        failed_files = result.get(
            "failed",
            [],
        )

        # ====================================================
        # NO REPORT SUCCESS
        # ====================================================

        if len(generated_files) == 0:

            raise HTTPException(
                status_code=422,
                detail={
                    "message": (
                        "Tidak ada report yang "
                        "berhasil dibuat."
                    ),
                    "employee_count": result.get(
                        "employee_count",
                        0,
                    ),
                    "failed": failed_files,
                    "master_file": result.get(
                        "master_file",
                        "",
                    ),
                },
            )

        # ====================================================
        # CREATE ZIP
        # ====================================================

        zip_buffer = io.BytesIO()

        with zipfile.ZipFile(
            zip_buffer,
            mode="w",
            compression=zipfile.ZIP_DEFLATED,
        ) as zip_file:

            for item in generated_files:

                zip_file.writestr(
                    item["filename"],
                    item["content"],
                )

        zip_buffer.seek(0)

        # ====================================================
        # RESPONSE
        # ====================================================

        filename = (
            "Learning_Agility_Reports.zip"
        )

        response = StreamingResponse(
            zip_buffer,
            media_type="application/zip",
        )

        response.headers[
            "Content-Disposition"
        ] = (
            f'attachment; filename="{filename}"'
        )

        # ====================================================
        # DEBUG HEADERS
        # ====================================================

        response.headers[
            "X-Generated-Reports"
        ] = str(
            len(generated_files)
        )

        response.headers[
            "X-Failed-Reports"
        ] = str(
            len(failed_files)
        )

        return response

    # ========================================================
    # HTTP ERROR
    # ========================================================

    except HTTPException:
        raise

    # ========================================================
    # GENERAL ERROR
    # ========================================================

    except Exception as exc:

        raise HTTPException(
            status_code=500,
            detail={
                "message": (
                    "Terjadi error saat "
                    "memproses assessment."
                ),
                "error": repr(exc),
            },
        )