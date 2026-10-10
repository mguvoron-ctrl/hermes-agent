"""
Разбор входящего Excel-инвойса (формат поставщика: листы FAC-IN /
Packing list / Инвойс перевод) → структурированные поля + таблица товаров.

Парсим по ПОДПИСЯМ полей (а не по жёстким адресам ячеек), чтобы небольшие
сдвиги в шаблоне не ломали импорт. Если какое-то поле не нашлось — оставляем
пустым, пользователь дозаполнит руками.

Используется в app.py при загрузке файла на инвойс.
"""

from __future__ import annotations
import re
import datetime as _dt

try:
    from openpyxl import load_workbook
except Exception:  # openpyxl не установлен
    load_workbook = None


def _norm(v) -> str:
    """Нормализация текста ячейки: только буквы/цифры, нижний регистр."""
    if v is None:
        return ""
    s = str(v).lower()
    return re.sub(r"[^0-9a-zа-яё]", "", s)


def _fmt(v) -> str:
    """Человекочитаемое значение ячейки."""
    if v is None:
        return ""
    if isinstance(v, (_dt.datetime, _dt.date)):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, float):
        if v == int(v):
            return str(int(v))
        return f"{round(v, 2):.2f}".rstrip("0").rstrip(".")
    if isinstance(v, int):
        return str(v)
    return str(v).strip()


def _matrix(ws) -> list[list]:
    return [list(row) for row in ws.iter_rows(values_only=True)]


def _cell(m: list[list], i: int, j: int):
    if 0 <= i < len(m) and 0 <= j < len(m[i]):
        return m[i][j]
    return None


def _find_right(m, keywords, need_digit=False) -> str:
    """Находит ячейку-подпись и возвращает первое непустое значение справа."""
    for i, row in enumerate(m):
        for j, cell in enumerate(row):
            nc = _norm(cell)
            if nc and any(k in nc for k in keywords):
                for jj in range(j + 1, len(row)):
                    val = _fmt(row[jj])
                    if val:
                        if need_digit and not re.search(r"\d", val):
                            break  # не то поле, пробуем следующую подпись
                        return val
    return ""


def _find_below(m, keywords) -> str:
    """Находит подпись и возвращает первое непустое значение НИЖЕ в том же столбце."""
    for i, row in enumerate(m):
        for j, cell in enumerate(row):
            nc = _norm(cell)
            if nc and any(k in nc for k in keywords):
                for ii in range(i + 1, len(m)):
                    val = _fmt(_cell(m, ii, j))
                    if val:
                        return val
    return ""


def _find_cell_text(m, keywords) -> str:
    """Возвращает полный текст первой ячейки, содержащей ключевое слово."""
    for row in m:
        for cell in row:
            if cell and any(k in _norm(cell) for k in keywords):
                return str(cell).strip()
    return ""


# Штрихкод (EAN-13) и страна происхождения внутри описания товара
_BARCODE = re.compile(r"\b(\d{12,14})\b")
_ORIGIN = re.compile(r"(?:origin|страна\s*происхождения)\s*:?\s*(.+)$",
                     re.IGNORECASE | re.DOTALL)


def _split_description(desc: str) -> tuple[str, str, str]:
    """Из описания достаём: наименование, штрихкод, страну происхождения."""
    if not desc:
        return "", "", ""
    text = str(desc)
    barcode = ""
    mb = _BARCODE.search(text)
    if mb:
        barcode = mb.group(1)
    origin = ""
    mo = _ORIGIN.search(text)
    if mo:
        origin = re.sub(r"\s+", " ", mo.group(1)).strip()
    # Наименование = текст до штрихкода / до "origin"
    name = text
    if mb:
        name = text[:mb.start()]
    name = re.split(r"origin|страна\s*происхождения", name, flags=re.IGNORECASE)[0]
    name = re.sub(r"\s+", " ", name).strip()
    return name, barcode, origin


def _parse_items(m_inv, m_rus, m_pack) -> list[dict]:
    """Таблица товаров из FAC-IN + вес/коробки из Packing list + рус. название."""
    # найти строку заголовка таблицы в FAC-IN
    header_i = None
    for i, row in enumerate(m_inv):
        joined = " ".join(_norm(c) for c in row)
        # заголовок: столбец кода (Custom/Customs Code, Код ТН ВЭД, HS Code) + столбец кол-ва
        has_code = ("custom" in joined or "тнвэд" in joined or "hscode" in joined)
        has_qty = ("qty" in joined or "колво" in joined or "quantity" in joined)
        if has_code and has_qty:
            header_i = i
            break
    if header_i is None:
        return []

    def col_of(m, i, keys):
        for j, c in enumerate(m[i]):
            if any(k in _norm(c) for k in keys):
                return j
        return None

    c_hs = col_of(m_inv, header_i, ["customcode", "кодтнвэд", "customscode"])
    c_item = col_of(m_inv, header_i, ["itemcode", "артикул"])
    c_desc = col_of(m_inv, header_i, ["description", "наименование"])
    c_qty = col_of(m_inv, header_i, ["qty", "колво", "quantity"])
    c_price = col_of(m_inv, header_i, ["unitprice", "ценазашт", "цена"])
    c_total = col_of(m_inv, header_i, ["totalprice", "стоимость", "total"])

    # карта веса/коробок из Packing list по артикулу
    pack = {}
    if m_pack:
        for row in m_pack:
            cells = [_fmt(c) for c in row]
            codes = [c for c in cells if re.fullmatch(r"[0-9A-Za-z]{5,}", c) and re.search(r"[A-Za-z]", c)]
            # ищем артикул в строке и три последних числовых столбца
            nums = [c for c in cells if re.fullmatch(r"-?\d+(\.\d+)?", c)]
            if codes and len(nums) >= 3:
                code = None
                for c in cells:
                    if re.fullmatch(r"\d?[A-Za-z]{1,3}\d{3,}", c) or re.search(r"[A-Za-z]\d{3,}", c):
                        code = c
                        break
                if code:
                    pack[code] = {"net_weight": nums[-3], "gross_weight": nums[-2], "box_count": nums[-1]}

    # рус. названия по артикулу
    rus = {}
    if m_rus:
        for row in m_rus:
            cells = row
            code = None
            name = ""
            for j, c in enumerate(cells):
                if c and re.search(r"[A-Za-z]\d{3,}", str(c)):
                    code = str(c).strip()
                    # наименование — следующая содержательная ячейка
                    for jj in range(j + 1, len(cells)):
                        if cells[jj] and len(str(cells[jj])) > 4:
                            name = str(cells[jj])
                            break
                    break
            if code and name:
                rus[code] = _split_description(name)[0]

    items = []
    for i in range(header_i + 1, len(m_inv)):
        hs = _fmt(_cell(m_inv, i, c_hs)) if c_hs is not None else ""
        item_code = _fmt(_cell(m_inv, i, c_item)) if c_item is not None else ""
        desc = _cell(m_inv, i, c_desc) if c_desc is not None else ""
        qty = _fmt(_cell(m_inv, i, c_qty)) if c_qty is not None else ""
        # явный конец таблицы — строка итогов
        if _norm(item_code).startswith(("total", "итого", "grand")) or \
           _norm(hs).startswith(("total", "итого", "grand")):
            break
        # строка товара = есть код ТН ВЭД, артикул и числовое количество
        if not hs or not item_code or \
           not re.fullmatch(r"\d+([.,]\d+)?", qty.replace(" ", "")):
            continue
        name, barcode, origin = _split_description(desc)
        if item_code in rus:
            name = rus[item_code] or name
        it = {
            "hs_code": hs, "item_code": item_code, "name": name,
            "barcode": barcode, "origin": origin, "qty": qty,
            "unit_price": _fmt(_cell(m_inv, i, c_price)) if c_price is not None else "",
            "total_price": _fmt(_cell(m_inv, i, c_total)) if c_total is not None else "",
            "net_weight": "", "gross_weight": "", "box_count": "",
        }
        if item_code in pack:
            it.update(pack[item_code])
        items.append(it)
    return items


def parse_invoice(path: str) -> dict:
    """Главная функция. Возвращает {'fields': {...}, 'items': [...], 'warnings': [...]}."""
    if load_workbook is None:
        return {"fields": {}, "items": [], "warnings": ["openpyxl не установлен"]}

    wb = load_workbook(path, data_only=True)

    def sheet(*names):
        for n in wb.sheetnames:
            nn = _norm(n)
            if any(k in nn for k in names):
                return _matrix(wb[n])
        return None

    m_inv = sheet("facin", "invoice", "инвойс") or _matrix(wb[wb.sheetnames[0]])
    m_pack = sheet("packing", "упаковочн")
    m_rus = sheet("перевод", "rus")

    fields: dict = {}
    fields["number"] = _find_right(m_inv, ["nnvoice", "ninvoice", "invoiceno", "номеринвойса"])
    fields["supplier"] = (_find_below(m_inv, ["exporter", "отправитель", "shipper"]) or "")
    # первая строка блока экспортёра = название компании
    if fields["supplier"]:
        fields["supplier"] = fields["supplier"].splitlines()[0].strip()
    fields["inv_date"] = _find_right(m_inv, ["invoicedate", "датаинвойса", "датаinvoice"])
    fields["incoterms"] = _find_right(m_inv, ["incoterms", "условияпоставки"])
    fields["currency"] = _find_right(m_inv, ["invoicecurrency", "ivoicecurrency", "валютаинвойса", "currency", "валюта"])
    fields["amount"] = _find_right(m_inv, ["totaltopay", "ксумме", "коплате", "итогокоплате"])
    fields["total_qty"] = _find_right(m_inv, ["totalofgoods", "итоготоваров"])
    fields["contract"] = _find_cell_text(m_inv, ["contract", "контракт"])
    fields["consignee"] = _find_below(m_inv, ["invoicingaddress", "юридическийадрес", "invoiceaddress"])
    if fields["consignee"]:
        fields["consignee"] = fields["consignee"].splitlines()[0].strip()
    inn_cell = _find_cell_text(m_inv, ["inn", "инн"])
    m_inn = re.search(r"\d{9,12}", inn_cell)
    fields["inn"] = m_inn.group(0) if m_inn else ""

    # из Packing list — вес, паллеты, коробки, номер PL
    if m_pack:
        fields["packing_list_no"] = _find_right(m_pack, ["packinglist", "упаковочныйлист"], need_digit=True)
        fields["net_weight"] = _find_right(m_pack, ["totalnetweight", "неттовес", "вескгнетто"])
        fields["gross_weight"] = _find_right(m_pack, ["totalgrossweight", "бруттовес"])
        fields["boxes"] = _find_right(m_pack, ["totalboxes", "всегокоробок"])
        fields["pallets"] = _find_right(m_pack, ["totalpallets", "всегопаллет"])

    # запасной вариант паллет/коробок из текстовой строки инвойса
    if not fields.get("pallets") or not fields.get("boxes"):
        txt = _find_cell_text(m_inv, ["pallet", "паллет"])
        mp = re.search(r"(\d+)\s*pallet", txt, re.IGNORECASE)
        mb = re.search(r"(\d+)\s*box", txt, re.IGNORECASE)
        if mp and not fields.get("pallets"):
            fields["pallets"] = mp.group(1)
        if mb and not fields.get("boxes"):
            fields["boxes"] = mb.group(1)

    items = _parse_items(m_inv, m_rus, m_pack)

    warnings = []
    if not fields.get("number"):
        warnings.append("Не найден номер инвойса — заполни вручную.")
    if not items:
        warnings.append("Не распознана таблица товаров — проверь формат файла.")

    return {"fields": {k: v for k, v in fields.items() if v},
            "items": items, "warnings": warnings}
