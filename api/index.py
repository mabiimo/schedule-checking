from __future__ import annotations

import io
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from fastapi import FastAPI, File, Form, UploadFile
from fastapi.responses import JSONResponse, StreamingResponse

from checker import (
    DataValidationError,
    FILTER_SEMUA,
    build_summary,
    compare_plan_actual,
    export_to_excel,
    export_to_pptx,
    filter_by_status,
    generate_output_filename,
    generate_pptx_filename,
)

app = FastAPI(title="Cek Plan vs Actual")


@app.get("/api/health")
async def health():
    return {"status": "ok"}


@app.post("/api/check")
async def check_data(
    plan_file: UploadFile = File(..., description="File Excel Plan (Data1)"),
    actual_file: UploadFile = File(..., description="File Excel Actual (Data2)"),
    status_filter: str = Form(FILTER_SEMUA),
):
    plan_bytes = await plan_file.read()
    actual_bytes = await actual_file.read()

    try:
        result_df = compare_plan_actual(plan_bytes, actual_bytes)
        result_df = filter_by_status(result_df, status_filter)
    except DataValidationError as exc:
        return JSONResponse(status_code=400, content={"detail": str(exc)})
    except Exception as exc:  # pragma: no cover
        return JSONResponse(
            status_code=500,
            content={"detail": f"Terjadi kesalahan saat memproses file: {exc}"},
        )

    summary_df = build_summary(result_df)
    excel_bytes = export_to_excel(result_df, summary_df)
    filename = generate_output_filename()

    return StreamingResponse(
        io.BytesIO(excel_bytes),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


@app.post("/api/check/pptx")
async def check_data_pptx(
    plan_file: UploadFile = File(..., description="File Excel Plan (Data1)"),
    actual_file: UploadFile = File(..., description="File Excel Actual (Data2)"),
    status_filter: str = Form(FILTER_SEMUA),
):
    plan_bytes = await plan_file.read()
    actual_bytes = await actual_file.read()

    try:
        result_df = compare_plan_actual(plan_bytes, actual_bytes)
        result_df = filter_by_status(result_df, status_filter)
    except DataValidationError as exc:
        return JSONResponse(status_code=400, content={"detail": str(exc)})
    except Exception as exc:  # pragma: no cover
        return JSONResponse(
            status_code=500,
            content={"detail": f"Terjadi kesalahan saat memproses file: {exc}"},
        )

    summary_df = build_summary(result_df)
    pptx_bytes = export_to_pptx(result_df, summary_df)
    filename = generate_pptx_filename()

    return StreamingResponse(
        io.BytesIO(pptx_bytes),
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


@app.post("/api/check/preview")
async def check_data_preview(
    plan_file: UploadFile = File(...),
    actual_file: UploadFile = File(...),
    status_filter: str = Form(FILTER_SEMUA),
):
    """Mengembalikan hasil cek dalam bentuk JSON (untuk preview di browser sebelum download)."""
    plan_bytes = await plan_file.read()
    actual_bytes = await actual_file.read()

    try:
        result_df = compare_plan_actual(plan_bytes, actual_bytes)
        result_df = filter_by_status(result_df, status_filter)
    except DataValidationError as exc:
        return JSONResponse(status_code=400, content={"detail": str(exc)})
    except Exception as exc:  # pragma: no cover
        return JSONResponse(
            status_code=500,
            content={"detail": f"Terjadi kesalahan saat memproses file: {exc}"},
        )

    summary_df = build_summary(result_df)

    return {
        "rows": result_df.to_dict(orient="records"),
        "summary": summary_df.to_dict(orient="records"),
        "total_rows": len(result_df),
    }
