"""
Трекер импортных поставок — первая рабочая версия.

Запуск:
    pip install -r requirements.txt
    python app.py
Затем открой в браузере http://127.0.0.1:8000
Первый вход: admin / admin  (поменяй пароль коллегам в разделе «Пользователи»).
"""

import os
import json
import secrets
import shutil
import tempfile

from fastapi import FastAPI, Request, Form, UploadFile, File, Depends, HTTPException
from fastapi.responses import HTMLResponse, RedirectResponse, FileResponse, Response, JSONResponse
from fastapi.templating import Jinja2Templates
from starlette.middleware.sessions import SessionMiddleware

import db
import xlsx_import
import ref_import
import camion_export

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
def shipment_detail(request: Request, shipment_id: int, err: str = "",
                    user=Depends(require_user)):
    shipment = db.get_shipment(shipment_id)
    if not shipment:
        raise HTTPException(404, "Поставка не найдена")
    invoices = db.list_invoices(shipment_id)
    rows = [{"inv": inv, "progress": db.progress(inv["id"])} for inv in invoices]
    return templates.TemplateResponse(
        request, "shipment.html",
        {"user": user, "shipment": shipment, "rows": rows, "nav": "shipments",
         "err": err},
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


def _parse_xlsx(data: bytes) -> dict:
    """Разбор Excel из байтов: сам определяет формат (инвойс поставщика
    или готовый/переделанный Камьён)."""
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        tmp.write(data)
        tmp_path = tmp.name
    try:
        return xlsx_import.parse_any(tmp_path)
    finally:
        try:
            os.remove(tmp_path)
        except OSError:
            pass


def _source_items(invoice_id: int) -> list:
    """Товары из сохранённого Исходного Excel — БЕЗ записи в БД.
    Нужно для предварительного Камьёна, пока проверенный файл не на этапе 4."""
    steps = db.get_steps(invoice_id)
    if not steps:
        return []
    src_id = steps[0]["id"]
    srcs = [f for f in db.list_files(invoice_id) if f["step_id"] == src_id]
    if not srcs:
        return []
    try:
        with open(srcs[0]["stored_path"], "rb") as fh:
            data = fh.read()
    except OSError:
        return []
    parsed = _parse_xlsx(data)
    out = []
    for it in parsed.get("items", []):
        row = {k: (it.get(k, "") or "") for k in ITEM_FIELDS}
        row["set_count"] = row.get("set_count") or "1"
        out.append(row)
    return out


def _apply_parsed(invoice_id: int, shipment, invoice, parsed: dict,
                  filename: str, data: bytes, user, replace_source: bool) -> None:
    """Сохраняет Исходный Excel и поля инвойса, отмечает входящие этапы.
    Таблицу товаров НЕ заполняем: она строится только из проверенного Камьёна,
    который загружается на этап 4. До этого из исходника собирается лишь
    предварительный Камьён (на лету, без записи в БД)."""
    fields = parsed["fields"]
    db.set_invoice_fields(invoice_id, {**fields, "source_file": filename})
    steps = db.get_steps(invoice_id)
    first = steps[0]["id"] if steps else None
    if replace_source and first:
        db.delete_step_files(first)  # убрать прежний исходник
    if first:
        _save_original(invoice_id, shipment, invoice, filename, data, first, user["name"])
    # Инвойс + Пакинг лист + Перевод приходят в одном файле → отмечаем все три
    for s in steps:
        if s["kind"] == "входящий" and s["position"] <= 2:
            db.update_step(s["id"], "готово", "получено из Excel-файла")


def _err_redirect(shipment_id: int, msg: str):
    from urllib.parse import quote
    return RedirectResponse(f"/shipments/{shipment_id}?err={quote(msg)}", status_code=302)


@app.post("/shipments/{shipment_id}/import")
async def import_invoice(request: Request, shipment_id: int,
                         f: UploadFile = File(...), user=Depends(require_user)):
    """Создать инвойс, разобрав присланный Excel-файл поставщика.
    В поставку загружается ТОЛЬКО инвойс поставщика (не Камьён и не прочее)."""
    shipment = db.get_shipment(shipment_id)
    if not shipment:
        raise HTTPException(404, "Поставка не найдена")

    name = (f.filename or "").lower()
    if not name.endswith((".xlsx", ".xlsm")):
        return _err_redirect(shipment_id,
            "Поддерживаются только файлы Excel (.xlsx/.xlsm) — инвойс поставщика.")

    data = await f.read()
    try:
        parsed = _parse_xlsx(data)
    except Exception as e:
        print("import parse error:", repr(e))
        return _err_redirect(shipment_id,
            "Не удалось прочитать файл. Убедись, что это инвойс поставщика в Excel "
            "(.xlsx) и он не защищён паролем.")

    if parsed.get("format") == "camion":
        return _err_redirect(shipment_id,
            "Это похоже на Камьён, а в поставку загружается только ИНВОЙС поставщика. "
            "Камьён загружается внутри инвойса, на этапе 4.")
    if not parsed.get("items") and not parsed.get("fields"):
        return _err_redirect(shipment_id,
            "Файл не распознан как инвойс поставщика. Проверь, что это исходный файл "
            "инвойса (FAC-IN / Packing list / Перевод).")

    number = parsed["fields"].get("number") or os.path.splitext(f.filename)[0]

    try:
        invoice_id = db.create_invoice(shipment_id, number,
                                       supplier=parsed["fields"].get("supplier", ""),
                                       amount=parsed["fields"].get("amount", ""))
    except Exception:
        existing = next((i for i in db.list_invoices(shipment_id)
                         if i["number"] == number), None)
        if not existing:
            return RedirectResponse(f"/shipments/{shipment_id}", status_code=302)
        invoice_id = existing["id"]

    invoice = db.get_invoice(invoice_id)
    _apply_parsed(invoice_id, shipment, invoice, parsed, f.filename, data, user,
                  replace_source=False)
    return RedirectResponse(f"/invoices/{invoice_id}", status_code=302)


@app.post("/invoices/{invoice_id}/reimport")
async def reimport_invoice(request: Request, invoice_id: int,
                           f: UploadFile = File(...), user=Depends(require_user)):
    """Заменить исходный Excel (инвойс поставщика) и пересчитать данные/поля."""
    from urllib.parse import quote
    invoice = db.get_invoice(invoice_id)
    if not invoice:
        raise HTTPException(404, "Инвойс не найден")
    shipment = db.get_shipment(invoice["shipment_id"])
    name = (f.filename or "").lower()
    if not name.endswith((".xlsx", ".xlsm")):
        return RedirectResponse(
            f"/invoices/{invoice_id}?err={quote('Нужен файл инвойса в Excel (.xlsx/.xlsm).')}",
            status_code=302)
    data = await f.read()
    try:
        parsed = _parse_xlsx(data)
    except Exception as e:
        print("reimport parse error:", repr(e))
        return RedirectResponse(
            f"/invoices/{invoice_id}?err={quote('Не удалось прочитать файл (проверь, что это инвойс в Excel без пароля).')}",
            status_code=302)
    if parsed.get("format") == "camion":
        return RedirectResponse(
            f"/invoices/{invoice_id}?err={quote('Это Камьён. Его загружай на этапе 4, а здесь — исходный инвойс поставщика.')}",
            status_code=302)
    _apply_parsed(invoice_id, shipment, invoice, parsed, f.filename, data, user,
                  replace_source=True)
    return RedirectResponse(f"/invoices/{invoice_id}", status_code=302)


# --- Инвойс (шаги + файлы) -------------------------------------------------
@app.get("/invoices/{invoice_id}", response_class=HTMLResponse)
def invoice_detail(request: Request, invoice_id: int, camion: int = 0,
                   err: str = "", user=Depends(require_user)):
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
    totals = _item_totals(items)
    vitems, vsummary = validate_items(items)
    kits = _build_kits(invoice_id, items)
    # Данные для встроенного редактора товаров: значения из сверки
    # (обогащённые страной/рус-названием/производителем) + статусы ячеек,
    # чтобы прямо в редактируемой таблице были подсвечены проблемы.
    _ed_keys = ITEM_FIELDS + ["producer", "hs_st", "ean_st", "origin_st",
                              "prod_st", "rus_st", "hs_ref"]
    items_json = json.dumps(
        [{k: (v.get(k, "") or "") for k in _ed_keys} for v in vitems],
        ensure_ascii=False,
    )
    # Этапность показа:
    #  - has_source  — загружен Исходный Excel (есть разобранные товары);
    #  - camion_done — Камьён уже на этапе 4 (приложен файл ИЛИ этап «готово»).
    # Данные инвойса / Товары / Разбивка показываем только после Камьёна на этапе 4.
    # Пока Камьёна нет — предлагаем скачать предварительный Камьён.
    step4 = steps[3] if len(steps) > 3 else None
    has_source = bool(invoice["source_file"]) or bool(items)
    # Товары показываем только когда проверенный Камьён на этапе 4 И есть строки.
    camion_done = bool(items) and bool(step4) and (
        step4["status"] == "готово" or bool(files_by_step.get(step4["id"]))
    )
    return templates.TemplateResponse(
        request, "invoice.html",
        {
            "user": user, "invoice": invoice,
            "shipment": shipment, "steps": steps,
            "files_by_step": files_by_step, "progress": db.progress(invoice_id),
            "statuses": db.STATUSES, "items": items, "totals": totals,
            "vsummary": vsummary, "camion_updated": camion,
            "kits": kits, "items_json": items_json, "nav": "invoices",
            "has_source": has_source, "camion_done": camion_done, "err": err,
        },
    )


ITEM_FIELDS = ["hs_code", "item_code", "barcode", "name", "name_en", "origin",
               "qty", "unit_price", "total_price", "net_weight", "gross_weight",
               "box_count", "set_count"]


def _has_cyr(s) -> bool:
    import re
    return bool(re.search(r"[А-Яа-яЁё]", str(s or "")))


def _digits(s) -> str:
    import re
    return re.sub(r"\D", "", str(s or ""))


def validate_items(items):
    """Сверяет товары инвойса со справочником (TRANSIT/Новинки).
    Возвращает список строк с данными для показа + статусами ячеек
    ('ok' | 'warn' | 'bad') и сводку проблем."""
    refmap = db.ref_get_many([ref_import.norm_art(it["item_code"]) for it in items])
    out, bad, warn = [], 0, 0
    for it in items:
        ref = refmap.get(ref_import.norm_art(it["item_code"]))
        row = {k: it[k] for k in ITEM_FIELDS}

        # --- ТН ВЭД ---
        if ref is None:
            row["hs_st"] = "warn"; row["hs_ref"] = ""
        elif ref["hs_code"] and str(ref["hs_code"]).strip() == str(it["hs_code"]).strip():
            row["hs_st"] = "ok"; row["hs_ref"] = ""
        elif ref["hs_code"]:
            row["hs_st"] = "bad"; row["hs_ref"] = ref["hs_code"]
        else:
            row["hs_st"] = "warn"; row["hs_ref"] = ""

        # --- EAN ---
        inv_ean = _digits(it["barcode"])
        ref_ean = _digits(ref["ean"]) if ref else ""
        if not inv_ean:
            row["ean_st"] = "bad"
        elif ref is None or not ref_ean:
            row["ean_st"] = "warn"
        elif inv_ean == ref_ean:
            row["ean_st"] = "ok"
        else:
            row["ean_st"] = "bad"

        # --- Страна (TRANSIT -> инвойс(жёлтый) -> нет(красный)) ---
        if ref and ref["origin"]:
            row["origin"] = ref["origin"]; row["origin_st"] = "ok"
        elif it["origin"]:
            row["origin_st"] = "warn"
        else:
            row["origin_st"] = "bad"

        # --- Производитель (только из справочника) ---
        row["producer"] = ref["producer"] if ref else ""
        row["prod_st"] = "ok" if row["producer"] else "bad"

        # --- Рус. название ---
        if ref and _has_cyr(ref["name_rus"]):
            if not _has_cyr(it["name"]):
                row["name"] = ref["name_rus"]
            row["rus_st"] = "ok"
        elif _has_cyr(it["name"]):
            row["rus_st"] = "warn"
        else:
            row["rus_st"] = "bad"

        has_bad = any(row[k] == "bad" for k in ("hs_st", "ean_st", "origin_st", "prod_st", "rus_st"))
        has_warn = any(row[k] == "warn" for k in ("hs_st", "ean_st", "origin_st", "prod_st", "rus_st"))
        if has_bad:
            bad += 1
        elif has_warn:
            warn += 1
        row["problem"] = has_bad or has_warn
        out.append(row)
    summary = {"bad": bad, "warn": warn, "checked": bool(refmap) or db.ref_stats()["total"] > 0}
    return out, summary


@app.post("/invoices/{invoice_id}/items")
def save_items(request: Request, invoice_id: int, data: str = Form(...),
               user=Depends(require_user)):
    """Автосохранение таблицы товаров (вызывается из JS, возвращает JSON)."""
    try:
        rows = json.loads(data)
    except (ValueError, TypeError):
        rows = []
    clean = []
    for r in rows:
        if not isinstance(r, dict):
            continue
        it = {k: str(r.get(k, "") or "").strip() for k in ITEM_FIELDS}
        if any(it.values()):  # пропускаем полностью пустые строки
            clean.append(it)
    db.replace_items(invoice_id, clean)
    return JSONResponse({"ok": True, "count": len(clean)})


@app.post("/invoices/{invoice_id}/fields")
def save_invoice_fields(request: Request, invoice_id: int, data: str = Form(...),
                        user=Depends(require_user)):
    """Автосохранение вручную правленых полей инвойса (JSON из JS)."""
    try:
        fields = json.loads(data)
    except (ValueError, TypeError):
        fields = {}
    if isinstance(fields, dict):
        db.update_invoice_fields(invoice_id, fields)
    return JSONResponse({"ok": True})


def _num(x) -> float:
    try:
        return float(str(x).replace(" ", "").replace(",", "."))
    except (ValueError, TypeError):
        return 0.0


def _fmt_num(x: float) -> str:
    if x == int(x):
        return f"{int(x):,}".replace(",", " ")
    return f"{x:,.2f}".replace(",", " ")


def _set_count(v) -> int:
    try:
        return int(float(str(v or "1").replace(",", ".")))
    except (ValueError, TypeError):
        return 1


def _build_kits(invoice_id: int, items) -> list:
    """Разбивка: для товаров с «Кол-во сост.» > 1 — блок из N строк составляющих."""
    kitmap = db.get_kit_rows(invoice_id)
    kits = []
    for it in items:
        n = _set_count(it["set_count"])
        if n <= 1:
            continue
        rows = []
        for slot in range(1, n + 1):
            ex = kitmap.get((it["item_code"], slot))
            rows.append({"slot": slot, **{k: (ex[k] if ex else "") for k in db.KIT_FIELDS}})
        kits.append({
            "hs_code": it["hs_code"], "item_code": it["item_code"],
            "name_en": it["name_en"] or "", "name": it["name"] or "",
            "barcode": it["barcode"] or "", "n": n, "rows": rows,
        })
    return kits


def _item_totals(items) -> dict:
    return {
        "qty": _fmt_num(sum(_num(it["qty"]) for it in items)),
        "total_price": _fmt_num(sum(_num(it["total_price"]) for it in items)),
        "net": _fmt_num(sum(_num(it["net_weight"]) for it in items)),
        "gross": _fmt_num(sum(_num(it["gross_weight"]) for it in items)),
        "boxes": _fmt_num(sum(_num(it["box_count"]) for it in items)),
    }


@app.post("/invoices/{invoice_id}/kit")
def save_kit(request: Request, invoice_id: int, data: str = Form(...),
             user=Depends(require_user)):
    """Сохранить составляющие наборов (разбивку)."""
    try:
        rows = json.loads(data)
    except (ValueError, TypeError):
        rows = []
    db.save_kit_rows(invoice_id, rows if isinstance(rows, list) else [])
    return RedirectResponse(f"/invoices/{invoice_id}", status_code=302)


@app.get("/invoices/{invoice_id}/razbivka.xlsx")
def export_razbivka(request: Request, invoice_id: int, user=Depends(require_user)):
    """Разбивка наборов в Excel (формат листа «Разбивка»)."""
    from urllib.parse import quote
    import razbivka_export
    invoice = db.get_invoice(invoice_id)
    if not invoice:
        raise HTTPException(404, "Инвойс не найден")
    kits = _build_kits(invoice_id, db.get_items(invoice_id))
    content = razbivka_export.build_razbivka_xlsx(kits)
    fname = f"Разбивка_{invoice['number']}.xlsx"
    return Response(
        content=content,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename*=UTF-8''{quote(fname)}"},
    )


@app.post("/invoices/{invoice_id}/steps/{step_id}")
async def set_step(request: Request, invoice_id: int, step_id: int,
                   status: str = Form(...), note: str = Form(""),
                   f: UploadFile = File(None), user=Depends(require_user)):
    """Сохранить шаг: статус + заметка, и (опционально) приложить файл."""
    db.update_step(step_id, status, note)
    if f is not None and f.filename:
        invoice = db.get_invoice(invoice_id)
        shipment = db.get_shipment(invoice["shipment_id"])
        data = await f.read()
        _save_original(invoice_id, shipment, invoice, f.filename, data, step_id, user["name"])
        # Файл приложен к этапу «Камьён» (поз. 3) и это Excel -> обновляем таблицу из него
        step = db.get_step(step_id)
        if step and step["position"] == 3 and f.filename.lower().endswith((".xlsx", ".xlsm")):
            parsed = _parse_xlsx(data)
            if parsed.get("items"):
                db.set_invoice_fields(invoice_id, parsed["fields"])
                db.replace_items(invoice_id, parsed["items"])
                return RedirectResponse(f"/invoices/{invoice_id}?camion=1", status_code=302)
    return RedirectResponse(f"/invoices/{invoice_id}", status_code=302)


@app.post("/invoices/{invoice_id}/steps/{step_id}/delfile")
def delete_step_file(request: Request, invoice_id: int, step_id: int,
                     user=Depends(require_user)):
    """Удалить файл(ы) этапа и снять отметку (этап снова «не начато»)."""
    db.delete_step_files(step_id)
    db.set_step_status(step_id, "не начато")
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


@app.get("/invoices/{invoice_id}/camion.xlsx")
def export_camion(request: Request, invoice_id: int, user=Depends(require_user)):
    """Готовый Камьён-реестр в Excel: метаданные + 15 колонок + подсветка проблем."""
    from urllib.parse import quote
    invoice = db.get_invoice(invoice_id)
    if not invoice:
        raise HTTPException(404, "Инвойс не найден")
    # Если проверенный Камьён уже на этапе 4 — берём товары из БД;
    # иначе собираем ПРЕДВАРИТЕЛЬНЫЙ из Исходного Excel (без записи в БД).
    items = db.get_items(invoice_id) or _source_items(invoice_id)
    vitems, _ = validate_items(items)
    content = camion_export.build_camion_xlsx(invoice, vitems)
    fname = f"Камьён_{invoice['number']}.xlsx"
    return Response(
        content=content,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename*=UTF-8''{quote(fname)}"},
    )


# --- Справочники (TRANSIT / НОВИНКИ) --------------------------------------
@app.get("/references", response_class=HTMLResponse)
def references_page(request: Request, user=Depends(require_user), msg: str = "",
                    q: str = "", source: str = "", brand: str = "", page: int = 1):
    if user["role"] != "admin":
        raise HTTPException(403, "Только для администратора")
    per = 50
    page = max(1, page)
    rows, total = db.ref_search(q.strip(), source, brand, limit=per, offset=(page - 1) * per)
    pages = max(1, (total + per - 1) // per)
    return templates.TemplateResponse(
        request, "references.html",
        {"user": user, "brands": db.ref_brands(), "msg": msg, "nav": "references",
         "rows": rows, "total": total, "q": q, "source": source, "brand": brand,
         "page": page, "pages": pages, "per": per},
    )


def _brand_from_name(filename: str) -> str:
    """Бренд из имени файла новинок (убираем годы, даты, лишние символы)."""
    import re
    name = os.path.splitext(os.path.basename(filename))[0]
    name = re.sub(r"^[0-9a-f]{6,}-", "", name)          # хэш-префикс загрузки, если есть
    name = name.replace("_", " ")
    name = re.sub(r"\b(?:19|20)\d{2}\b", " ", name)      # годы
    name = re.sub(r"\d+", " ", name)                     # прочие числа/даты
    name = re.sub(r"[^\w\sА-Яа-яЁё&.-]", " ", name)
    name = re.sub(r"\s+", " ", name).strip(" _-.")
    return name or "Новинки"


def _save_tmp(data: bytes, suffix: str) -> str:
    with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
        tmp.write(data)
        return tmp.name


@app.post("/references/transit")
async def references_transit(request: Request, f: UploadFile = File(...),
                             password: str = Form(""), user=Depends(require_user)):
    if user["role"] != "admin":
        raise HTTPException(403)
    data = await f.read()
    path = _save_tmp(data, ".xlsx")
    dec = None
    try:
        # пробуем расшифровать (файл под паролем); если не зашифрован — читаем как есть
        try:
            import msoffcrypto
            with open(path, "rb") as fh:
                off = msoffcrypto.OfficeFile(fh)
                off.load_key(password=password or "2208")
                dec = _save_tmp(b"", ".xlsx")
                with open(dec, "wb") as out:
                    off.decrypt(out)
            parse_path = dec
        except Exception:
            parse_path = path
        rows = ref_import.parse_transit(parse_path)
        n = db.ref_upsert_transit(rows)
        msg = f"TRANSIT загружен: {n} артикулов (по самой свежей поставке)."
    except Exception as e:
        msg = f"Ошибка загрузки TRANSIT: {e}"
    finally:
        for p in (path, dec):
            if p:
                try:
                    os.remove(p)
                except OSError:
                    pass
    return RedirectResponse(f"/references?msg={msg}", status_code=302)


@app.post("/references/novinki")
async def references_novinki(request: Request, files: list[UploadFile] = File(...),
                             user=Depends(require_user)):
    if user["role"] != "admin":
        raise HTTPException(403)
    total, done = 0, 0
    for f in files:
        if not f.filename:
            continue
        data = await f.read()
        suffix = ".xlsb" if f.filename.lower().endswith(".xlsb") else ".xlsx"
        path = _save_tmp(data, suffix)
        try:
            rows = ref_import.parse_novinki(path)
            db.ref_upsert_novinki(rows, source_name=_brand_from_name(f.filename))
            total += len(rows); done += 1
        except Exception as e:
            print("novinki error", f.filename, e)
        finally:
            try:
                os.remove(path)
            except OSError:
                pass
    msg = f"Новинки загружены: {done} файлов, {total} артикулов (добавлены недостающие)."
    return RedirectResponse(f"/references?msg={msg}", status_code=302)


if __name__ == "__main__":
    import uvicorn
    # host 0.0.0.0 — чтобы коллеги в офисной сети тоже могли подключиться
    uvicorn.run(app, host="0.0.0.0", port=int(_port))
