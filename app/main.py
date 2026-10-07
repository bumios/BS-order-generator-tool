"""FastAPI app: read orders from Google Sheets (or an uploaded file) and generate BigSeller xlsx files."""
import io
import uuid
from datetime import datetime
from pathlib import Path

import openpyxl

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import RedirectResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from . import core, gsheet

BASE = Path(__file__).resolve().parent.parent
TEMPLATE_DIR = BASE / "template"
STATIC_DIR = BASE / "static"
UPLOAD_DIR = BASE / "uploads"
UPLOAD_DIR.mkdir(exist_ok=True)

app = FastAPI(title="BS Order Generator")


def find_template():
    """Most recently modified .xlsx inside template/."""
    candidates = sorted(
        TEMPLATE_DIR.glob("*.xlsx"), key=lambda p: p.stat().st_mtime, reverse=True
    )
    candidates = [p for p in candidates if not p.name.startswith("~$")]
    if not candidates:
        raise HTTPException(
            400, f"Không tìm thấy file template .xlsx trong thư mục {TEMPLATE_DIR.name}/"
        )
    return candidates[0]


# --- Auth --------------------------------------------------------------------

@app.get("/api/status")
def status():
    try:
        template = find_template().name
    except HTTPException:
        template = None
    return {
        "google_ready": gsheet.credentials_ready(),
        "authenticated": gsheet._load_credentials() is not None,
        "template": template,
    }


@app.get("/api/auth/login")
def auth_login():
    try:
        return {"url": gsheet.start_login()}
    except gsheet.AuthError as e:
        raise HTTPException(400, str(e))


@app.get("/oauth/callback")
def oauth_callback(code: str):
    try:
        gsheet.finish_login(code)
    except Exception as e:  # noqa: BLE001 - surface OAuth errors to the UI
        raise HTTPException(400, f"Đăng nhập Google thất bại: {e}")
    return RedirectResponse("/", status_code=302)


@app.post("/api/auth/logout")
def auth_logout():
    gsheet.clear_token()
    return {"ok": True}


# --- Source: upload -----------------------------------------------------------

@app.post("/api/upload")
async def upload(file: UploadFile = File(...)):
    ext = Path(file.filename or "data.xlsx").suffix.lower()
    if ext not in (".xlsx", ".csv"):
        raise HTTPException(400, "Chỉ hỗ trợ file .xlsx hoặc .csv")
    fid = uuid.uuid4().hex[:10]
    path = UPLOAD_DIR / f"{fid}{ext}"
    path.write_bytes(await file.read())
    return {
        "fileId": fid,
        "filename": file.filename,
        "sheets": core.list_upload_sheets(path),
    }


def _upload_path(file_id):
    for ext in (".xlsx", ".csv"):
        p = UPLOAD_DIR / f"{file_id}{ext}"
        if p.exists():
            return p
    raise HTTPException(404, "File upload không tồn tại hoặc đã bị xóa")


# --- Source: Google Sheets ----------------------------------------------------

@app.get("/api/sheets")
def sheets(spreadsheetId: str = "", fileId: str = ""):
    if spreadsheetId:
        sid = gsheet.extract_spreadsheet_id(spreadsheetId)
        try:
            return {"sheets": gsheet.list_sheets(sid)}
        except gsheet.AuthError as e:
            raise HTTPException(401, str(e))
        except Exception as e:  # noqa: BLE001
            raise HTTPException(400, f"Không đọc được file Google Sheets: {e}")
    if fileId:
        return {"sheets": core.list_upload_sheets(_upload_path(fileId))}
    raise HTTPException(400, "Cần spreadsheetId hoặc fileId")


# --- Preview ------------------------------------------------------------------

@app.get("/api/preview")
def preview(spreadsheetId: str = "", fileId: str = "", sheet: str = ""):
    if not sheet:
        raise HTTPException(400, "Thiếu tên sheet")
    if spreadsheetId:
        sid = gsheet.extract_spreadsheet_id(spreadsheetId)
        try:
            values = gsheet.read_sheet(sid, sheet)
        except gsheet.AuthError as e:
            raise HTTPException(401, str(e))
        except Exception as e:  # noqa: BLE001
            raise HTTPException(400, f"Không đọc được sheet {sheet!r}: {e}")
    elif fileId:
        values = core.read_upload_sheet(_upload_path(fileId), sheet)
    else:
        raise HTTPException(400, "Cần spreadsheetId hoặc fileId")
    return core.parse_order_sheet(sheet, values)


# --- Generate -----------------------------------------------------------------

class GenerateRequest(BaseModel):
    spreadsheetId: str = ""
    fileId: str = ""
    sheets: list[str]


@app.post("/api/generate")
def generate(req: GenerateRequest):
    if not req.sheets:
        raise HTTPException(400, "Chưa chọn sheet nào")
    now = datetime.now()
    template_path = find_template()
    wb = openpyxl.load_workbook(template_path)
    ws = wb.worksheets[0]
    try:
        row, cols = core.locate_fields(ws)
    except ValueError as e:
        raise HTTPException(400, str(e))

    seen_ids = set()
    summary = []
    for sheet in req.sheets:
        if req.spreadsheetId:
            sid = gsheet.extract_spreadsheet_id(req.spreadsheetId)
            try:
                values = gsheet.read_sheet(sid, sheet)
            except gsheet.AuthError as e:
                raise HTTPException(401, str(e))
            except Exception as e:  # noqa: BLE001
                raise HTTPException(400, f"Không đọc được sheet {sheet!r}: {e}")
        elif req.fileId:
            values = core.read_upload_sheet(_upload_path(req.fileId), sheet)
        else:
            raise HTTPException(400, "Cần spreadsheetId hoặc fileId")

        order = core.parse_order_sheet(sheet, values, now=now)
        if not order["hasHeader"]:
            raise HTTPException(
                400, f"Sheet {sheet!r} không có cột 'BS - SKU' / 'BS - Số lượng'"
            )
        if not order["rows"]:
            raise HTTPException(400, f"Sheet {sheet!r} không có dòng dữ liệu hợp lệ")
        if order["orderId"] in seen_ids:
            order["orderId"] = f"{order['orderId']}_2"
        seen_ids.add(order["orderId"])

        row = core.append_order(ws, order, row, cols)
        summary.append(
            {
                "sheet": sheet,
                "orderId": order["orderId"],
                "rows": len(order["rows"]),
                "skipped": order["skipped"],
            }
        )

    out = io.BytesIO()
    wb.save(out)
    name = f"BS-orders_{now:%Y%m%d_%H%M}.xlsx"
    return StreamingResponse(
        io.BytesIO(out.getvalue()),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{name}"'},
    )


app.mount("/", StaticFiles(directory=STATIC_DIR, html=True), name="static")
