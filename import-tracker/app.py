"""
Трекер импортных поставок — первая рабочая версия.

Запуск:
    pip install -r requirements.txt
    python app.py
Затем открой в браузере http://127.0.0.1:8000
Первый вход: admin / admin  (поменяй пароль коллегам в разделе «Пользователи»).
"""

import os
import secrets
import shutil
import tempfile

from fastapi import FastAPI, Request, Form, UploadFile, File, Depends, HTTPException
from fastapi.responses import HTMLResponse, RedirectResponse, FileResponse
from fastapi.templating import Jinja2Templates
from starlette.middleware.sessions import SessionMiddleware

import db
import xlsx_import

db.init_db()
created_admin = db.ensure_admin()
if created_admin:
    print(f"[init] Создан первый пользователь: {created_admin} / (пароль из ADMIN_PASSWORD или 'admin')")

def _lan_ip() -> str:
    """Локальный IP компьютера в офисной сети (для адреса коллегам)."""
    import socket
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
    except Exception:
        ip = "127.0.0.1"
    finally:
        s.close()
    return ip


_port = os.getenv("PORT", "8000")
print("\n" + "=" * 54)
print("  📦  Трекер поставок — сервер запускается")
print(f"  На ЭТОМ компьютере:   http://127.0.0.1:{_port}")
print(f"  Коллегам в сети дать: http://{_lan_ip()}:{_port}")
print("  (чтобы остановить — закрой это окно)")
print("=" * 54 + "\n")

app = FastAPI(title="Трекер поставок")
app.add_middleware(
    SessionMiddleware,
    secret_key=os.getenv("SECRET_KEY", secrets.token_hex(32)),
)
templates = Jinja2Templates(directory=str(db.BASE_DIR / "templates"))


# --- Авторизация -----------------------------------------------------------
def current_user(request: Request):
    return request.session.get("user")


def require_user(request: Request):
    user = current_user(request)
    if not user:
        raise HTTPException(status_code=307, headers={"Location": "/login"})
    return user


@app.exception_handler(HTTPException)
async def redirect_login(request: Request, exc: HTTPException):
    if exc.status_code == 307 and exc.headers and "Location" in exc.headers:
        return RedirectResponse(exc.headers["Location"], status_code=307)
    raise exc


@app.get("/login", response_class=HTMLResponse)
def login_form(request: Request):
    return templates.TemplateResponse(request, "login.html", {"error": None})


@app.post("/login")
def login(request: Request, username: str = Form(...), password: str = Form(...)):
    row = db.get_user(username.strip())
    if row and db.verify_password(password, row["password_hash"]):
        request.session["user"] = {"name": row["username"], "role": row["role"]}
        return RedirectResponse("/", status_code=302)
    return templates.TemplateResponse(
        request, "login.html", {"error": "Неверный логин или пароль"}
    )


@app.get("/logout")
def logout(request: Request):
    request.session.clear()
    return RedirectResponse("/login", status_code=302)


# --- Поставки --------------------------------------------------------------
@app.get("/", response_class=HTMLResponse)
def index(request: Request, user=Depends(require_user)):
    shipments = db.list_shipments()
    return templates.TemplateResponse(
        request, "shipments.html", {"user": user, "shipments": shipments, "nav": "shipments"}
    )


@app.get("/invoices", response_class=HTMLResponse)
def all_invoices(request: Request, user=Depends(require_user)):
    """Общий экран-матрица: строки — инвойсы, столбцы — этапы, в клетках галочки."""
    rows = []
    for inv in db.list_all_invoices():
        steps = db.get_steps(inv["id"])
        rows.append({
            "inv": inv,
            "statuses": [s["status"] for s in steps],  # по порядку этапов
            "progress": db.progress(inv["id"]),
        })
    return templates.TemplateResponse(
        request, "all_invoices.html",
        {"user": user, "rows": rows, "stages": db.STAGE_SHORT,
         "stage_full": [n for n, _ in db.DEFAULT_STEPS], "nav": "invoices"},
    )


@app.post("/shipments")
def add_shipment(request: Request, number: str = Form(...), title: str = Form(""),
                 note: str = Form(""), user=Depends(require_user)):
    if number.strip():
        try:
            db.create_shipment(number, title, note)
        except Exception as e:  # дубликат номера и т.п.
            print("create_shipment error:", e)
    return RedirectResponse("/", status_code=302)


@app.get("/shipments/{shipment_id}", response_class=HTMLResponse)
def shipment_detail(request: Request, shipment_id: int, user=Depends(require_user)):
    shipment = db.get_shipment(shipment_id)
    if not shipment:
        raise HTTPException(404, "Поставка не найдена")
    invoices = db.list_invoices(shipment_id)
    rows = [{"inv": inv, "progress": db.progress(inv["id"])} for inv in invoices]
    return templates.TemplateResponse(
        request, "shipment.html",
        {"user": user, "shipment": shipment, "rows": rows, "nav": "shipments"},
    )


@app.post("/shipments/{shipment_id}/invoices")
def add_invoice(request: Request, shipment_id: int, number: str = Form(...),
                supplier: str = Form(""), amount: str = Form(""),
                note: str = Form(""), user=Depends(require_user)):
    if number.strip():
        try:
            db.create_invoice(shipment_id, number, supplier, amount, note)
        except Exception as e:
            print("create_invoice error:", e)
    return RedirectResponse(f"/shipments/{shipment_id}", status_code=302)


def _save_original(invoice_id: int, shipment, invoice, filename: str,
                   data: bytes, step_id, uploaded_by: str) -> None:
    """Сохраняет оригинал файла в папку инвойса и регистрирует в базе."""
    folder = db.invoice_folder(shipment["number"], invoice["number"])
    dest = folder / filename
    if dest.exists():
        dest = folder / f"{secrets.token_hex(4)}_{filename}"
    with open(dest, "wb") as out:
        out.write(data)
    db.add_file(invoice_id, step_id, filename, str(dest), len(data), uploaded_by)


@app.post("/shipments/{shipment_id}/import")
async def import_invoice(request: Request, shipment_id: int,
                         f: UploadFile = File(...), user=Depends(require_user)):
    """Создать инвойс, разобрав присланный Excel-файл поставщика."""
    shipment = db.get_shipment(shipment_id)
    if not shipment:
        raise HTTPException(404, "Поставка не найдена")

    data = await f.read()
    # openpyxl читает с диска — пишем во временный файл
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        tmp.write(data)
        tmp_path = tmp.name
    try:
        parsed = xlsx_import.parse_invoice(tmp_path)
    finally:
        try:
            os.remove(tmp_path)
        except OSError:
            pass

    fields = parsed["fields"]
    number = fields.get("number") or os.path.splitext(f.filename)[0]

    try:
        invoice_id = db.create_invoice(shipment_id, number,
                                       supplier=fields.get("supplier", ""),
                                       amount=fields.get("amount", ""))
    except Exception:
        # инвойс с таким номером уже есть — открываем существующий и обновляем
        existing = next((i for i in db.list_invoices(shipment_id)
                         if i["number"] == number), None)
        if not existing:
            return RedirectResponse(f"/shipments/{shipment_id}", status_code=302)
        invoice_id = existing["id"]

    invoice = db.get_invoice(invoice_id)
    db.set_invoice_fields(invoice_id, {**fields, "source_file": f.filename})
    if parsed["items"]:
        db.replace_items(invoice_id, parsed["items"])

    # сохранить оригинал (привязан к этапу «Инвойс»)
    steps = db.get_steps(invoice_id)
    first_step_id = steps[0]["id"] if steps else None
    _save_original(invoice_id, shipment, invoice, f.filename, data,
                   first_step_id, user["name"])

    # авто-отметка: входящие документы (Инвойс, Пакинг лист, Перевод) пришли в этом файле
    for s in steps:
        if s["kind"] == "входящий" and s["position"] <= 2:
            db.update_step(s["id"], "готово", "получено из Excel-файла")

    return RedirectResponse(f"/invoices/{invoice_id}", status_code=302)


# --- Инвойс (шаги + файлы) -------------------------------------------------
@app.get("/invoices/{invoice_id}", response_class=HTMLResponse)
def invoice_detail(request: Request, invoice_id: int, user=Depends(require_user)):
    invoice = db.get_invoice(invoice_id)
    if not invoice:
        raise HTTPException(404, "Инвойс не найден")
    shipment = db.get_shipment(invoice["shipment_id"])
    steps = db.get_steps(invoice_id)
    files = db.list_files(invoice_id)
    files_by_step = {}
    for f in files:
        files_by_step.setdefault(f["step_id"], []).append(f)
    items = db.get_items(invoice_id)
    return templates.TemplateResponse(
        request, "invoice.html",
        {
            "user": user, "invoice": invoice,
            "shipment": shipment, "steps": steps, "files": files,
            "files_by_step": files_by_step, "progress": db.progress(invoice_id),
            "statuses": db.STATUSES, "items": items, "nav": "invoices",
        },
    )


@app.post("/invoices/{invoice_id}/steps/{step_id}")
def set_step(request: Request, invoice_id: int, step_id: int,
             status: str = Form(...), note: str = Form(""),
             user=Depends(require_user)):
    db.update_step(step_id, status, note)
    return RedirectResponse(f"/invoices/{invoice_id}", status_code=302)


@app.post("/invoices/{invoice_id}/files")
async def upload_file(request: Request, invoice_id: int,
                      step_id: str = Form(""), f: UploadFile = File(...),
                      user=Depends(require_user)):
    invoice = db.get_invoice(invoice_id)
    if not invoice:
        raise HTTPException(404, "Инвойс не найден")
    shipment = db.get_shipment(invoice["shipment_id"])
    folder = db.invoice_folder(shipment["number"], invoice["number"])

    # Уникальное имя на диске, но храним и оригинальное имя
    dest = folder / f.filename
    if dest.exists():
        dest = folder / f"{secrets.token_hex(4)}_{f.filename}"
    with open(dest, "wb") as out:
        shutil.copyfileobj(f.file, out)

    step_id_val = int(step_id) if step_id.strip().isdigit() else None
    db.add_file(invoice_id, step_id_val, f.filename, str(dest),
                dest.stat().st_size, user["name"])
    return RedirectResponse(f"/invoices/{invoice_id}", status_code=302)


@app.get("/files/{file_id}/download")
def download_file(request: Request, file_id: int, user=Depends(require_user)):
    row = db.get_file(file_id)
    if not row or not os.path.exists(row["stored_path"]):
        raise HTTPException(404, "Файл не найден")
    return FileResponse(row["stored_path"], filename=row["original_name"])


@app.post("/files/{file_id}/delete")
def remove_file(request: Request, file_id: int, invoice_id: int = Form(...),
                user=Depends(require_user)):
    db.delete_file(file_id)
    return RedirectResponse(f"/invoices/{invoice_id}", status_code=302)


# --- Пользователи (только админ) -------------------------------------------
@app.get("/users", response_class=HTMLResponse)
def users_page(request: Request, user=Depends(require_user)):
    if user["role"] != "admin":
        raise HTTPException(403, "Только для администратора")
    return templates.TemplateResponse(
        request, "users.html", {"user": user, "users": db.list_users(), "error": None, "nav": "users"}
    )


@app.post("/users")
def add_user(request: Request, username: str = Form(...), password: str = Form(...),
             role: str = Form("user"), user=Depends(require_user)):
    if user["role"] != "admin":
        raise HTTPException(403, "Только для администратора")
    error = None
    try:
        db.create_user(username.strip(), password, role)
    except Exception:
        error = "Такой логин уже существует"
    return templates.TemplateResponse(
        request, "users.html", {"user": user, "users": db.list_users(), "error": error, "nav": "users"}
    )


if __name__ == "__main__":
    import uvicorn
    # host 0.0.0.0 — чтобы коллеги в офисной сети тоже могли подключиться
    uvicorn.run(app, host="0.0.0.0", port=int(_port))
