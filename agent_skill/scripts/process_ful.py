#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Детерминированное распределение ФУЛ между логистами. Python 3 + openpyxl."""
import argparse
import json
import re
import sys
import unicodedata
from collections import Counter, defaultdict
from datetime import datetime, date
from pathlib import Path
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

UNKNOWN = "не определено"

ALIASES = {
    "ful": {
        "Дата": ["дата", "дата фул", "дата и время фул", "дата получения фул"],
        "Cut-off date": ["cut-off date", "cut off date", "cutoff date", "cut-off", "cut off"],
        "Номер": ["номер", "номер фул", "номер упаковочного листа"],
        "Статус": ["статус"],
        "Комментарий": ["комментарий"],
        "Способ доставки": ["способ доставки", "доставка", "тип доставки"],
        "Бизнес единица": ["бизнес единица", "бизнес-единица", "бренд"],
        "Логист": ["логист"],
        "Контрагент": ["контрагент", "поставщик"],
        "Страна": ["страна", "страна отправления"],
    },
    "cloud": {
        "Заказ": ["заказ", "номер заказа"],
        "Поставщик": ["поставщик", "контрагент"],
        "Бренд": ["бренд", "бизнес единица", "бизнес-единица"],
        "Логист": ["логист"],
        "Статус": ["статус"],
        "Дата подтверждения": ["дата подтверждения"],
        "cut-off": ["cut-off", "cut-off date", "cut off", "cutoff"],
        "Дата получения УЛ": ["дата получения ул", "дата получения упаковочного листа"],
        "Страна отправления": ["страна отправления", "страна"],
        "Способ доставки": ["способ доставки", "доставка"],
        "Документы получены": ["документы получены"],
        "Комментарий": ["комментарий"],
    },
    "employees": {
        "Наименование": ["наименование", "фио", "сотрудник", "фамилия имя"],
        "Текущее состояние": ["текущее состояние", "состояние", "статус"],
        "Должность": ["должность", "роль"],
        "Код": ["код", "таб. №", "табельный номер", "табельный №"],
    },
}

ROLE_TO_GROUP = {
    "специалист по авиаперевозкам": "air",
    "специалист по работе с вэд документацией": "ved",
    "специалист по мультимодальным перевозкам": "multi",
}
DELIVERY = {"air": "air", "sea": "multi", "truck": "multi",
            "direct truck": "multi", "railway": "multi"}

def norm(value):
    if value is None:
        return ""
    s = unicodedata.normalize("NFKC", str(value)).strip().casefold()
    s = re.sub(r"[.,;:()'\"`]", "", s)
    s = re.sub(r"[\s\-_]+", " ", s)
    return s.strip()

def clean(value):
    if value is None or str(value).strip() == "":
        return UNKNOWN
    if isinstance(value, (datetime, date)):
        return value.strftime("%d.%m.%Y")
    return str(value).strip()

def parse_date(value):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if value is None or str(value).strip() == "":
        return None
    s = str(value).strip()
    for fmt in ("%d.%m.%Y", "%Y-%m-%d", "%d/%m/%Y", "%d.%m.%Y %H:%M", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return None

def parse_order(comment):
    if comment is None:
        return None
    m = re.search(r"\b(?:Номер\s*:\s*)?([A-ZА-Я]\d{5,})\b", str(comment), re.I)
    return m.group(1).upper() if m else None

def find_header_map(ws, kind):
    aliases = ALIASES[kind]
    rows_to_scan = min(ws.max_row, 12)
    for row_idx in range(1, rows_to_scan + 1):
        values = [norm(c.value) for c in ws[row_idx]]
        mapping = {}
        for canonical, variants in aliases.items():
            candidates = {norm(canonical), *(norm(v) for v in variants)}
            for col_idx, value in enumerate(values, start=1):
                if value in candidates:
                    mapping[canonical] = col_idx
                    break
        required = {
            "ful": {"Дата", "Номер", "Комментарий", "Способ доставки", "Контрагент", "Страна"},
            "cloud": {"Заказ", "Поставщик", "Логист", "Статус"},
            "employees": {"Наименование", "Текущее состояние", "Должность"},
        }[kind]
        if len(required.intersection(mapping)) >= max(3, len(required)-1):
            return row_idx, mapping
    raise ValueError(f"Не удалось определить строку заголовков для файла типа {kind}")

def read_records(path, kind):
    wb = load_workbook(path, data_only=True, read_only=True)
    for ws in wb.worksheets:
        try:
            header_row, mapping = find_header_map(ws, kind)
        except ValueError:
            continue
        records = []
        for row in ws.iter_rows(min_row=header_row+1, values_only=True):
            if not any(v is not None and str(v).strip() for v in row):
                continue
            rec = {}
            for canonical, col in mapping.items():
                rec[canonical] = row[col-1] if col-1 < len(row) else None
            records.append(rec)
        wb.close()
        return records
    wb.close()
    raise ValueError(f"В файле {path} не найден лист с ожидаемыми заголовками ({kind})")

def group_for(delivery, country):
    d = norm(delivery)
    c = norm(country)
    if not d:
        return None
    if d == "air":
        return "air"
    if not c:
        return None
    if c in ("китай", "china", "cn"):
        return "ved"
    if d in DELIVERY and DELIVERY[d] == "multi":
        return "multi"
    return None

def state_available(state):
    return norm(state) in ("работает", "работает удаленно", "работает удалённо")

def role_group(role):
    return ROLE_TO_GROUP.get(norm(role))

def load_input_files(ful_path, cloud_path, employees_path):
    fuls = read_records(ful_path, "ful")
    cloud = read_records(cloud_path, "cloud")
    employees = read_records(employees_path, "employees")
    return fuls, cloud, employees

def distribute(fuls, cloud, employees):
    cloud_orders = {clean(r.get("Заказ")).upper() for r in cloud if clean(r.get("Заказ")) != UNKNOWN}
    seen_batch = set()
    history = Counter()
    workload = Counter()
    for r in cloud:
        supplier = norm(r.get("Поставщик"))
        logistic = clean(r.get("Логист"))
        if supplier and logistic != UNKNOWN:
            history[(supplier, norm(logistic))] += 1
        if norm(r.get("Статус")) == "в работе" and logistic != UNKNOWN:
            workload[norm(logistic)] += 1

    staff = []
    for e in employees:
        name = clean(e.get("Наименование"))
        role = role_group(e.get("Должность"))
        if name == UNKNOWN or role is None:
            continue
        staff.append({
            "name": name, "key": norm(name), "group": role,
            "available": state_available(e.get("Текущее состояние")),
            "state": clean(e.get("Текущее состояние")),
        })

    latest_cloud_date = max((parse_date(r.get("Дата получения УЛ")) for r in cloud
                             if parse_date(r.get("Дата получения УЛ"))), default=None)

    new_rows, repeats, attention, journal = [], [], [], []
    for idx, f in enumerate(fuls, start=1):
        order = parse_order(f.get("Комментарий"))
        supplier_raw = clean(f.get("Контрагент"))
        supplier_key = norm(f.get("Контрагент"))
        ful_date = parse_date(f.get("Дата"))
        delivery = clean(f.get("Способ доставки"))
        country = clean(f.get("Страна"))
        brand = clean(f.get("Бизнес единица"))
        cutoff = clean(f.get("Cut-off date"))
        reasons = []
        decision = "Новая строка"
        group = group_for(f.get("Способ доставки"), f.get("Страна"))
        logistic = UNKNOWN
        is_repeat = False

        if not order:
            reasons.append("Не указан номер заказа: невозможно надёжно проверить дубликат")
            attention.append([idx, clean(f.get("Номер")), UNKNOWN, supplier_raw,
                              "Не создавать строку", "Номер заказа отсутствует"])
            journal.append([idx, UNKNOWN, supplier_raw, UNKNOWN, UNKNOWN, "Не создавать строку",
                            "Не определено", "; ".join(reasons)])
            continue

        if order in cloud_orders or order in seen_batch:
            is_repeat = True
            decision = "Повторная отправка ФУЛ"
            date_text = clean(f.get("Дата"))
            repeats.append([idx, order, supplier_raw, date_text,
                            f"Повторная отправка ФУЛа, {date_text}"])
            journal.append([idx, order, supplier_raw, UNKNOWN, UNKNOWN, decision,
                            "Не создавать новую строку", "Заказ уже есть в Cloud или ранее встречен в этой пачке"])
            seen_batch.add(order)
            continue

        seen_batch.add(order)

        if group is None:
            reasons.append("Группу нельзя определить: отсутствует или не распознан способ доставки/страна")
            group_label = UNKNOWN
        else:
            group_label = {"air": "Авиаперевозки", "ved": "ВЭД", "multi": "Мультимодальные перевозки"}[group]
            candidates = [s for s in staff if s["group"] == group and s["available"]]
            if not candidates:
                reasons.append("В целевой группе нет доступных сотрудников")
            else:
                group_staff = [s for s in staff if s["group"] == group]
                all_history = [(history[(supplier_key, s["key"])], s) for s in group_staff]
                max_history = max((n for n, _ in all_history), default=0)
                preferred_all = [s for n, s in all_history if n == max_history and n > 0]
                preferred_available = [s for s in preferred_all if s["available"]]
                if preferred_available:
                    logistic = sorted(preferred_available, key=lambda s: (workload[s["key"]], s["key"]))[0]["name"]
                    reasons.append(f"Выбран по истории поставщика: {max_history} заказ(а/ов)")
                else:
                    logistic = sorted(candidates, key=lambda s: (workload[s["key"]], s["key"]))[0]["name"]
                    if max_history > 0:
                        reasons.append("Предпочтительный логист недоступен: выбран наименее загруженный доступный сотрудник")
                    else:
                        reasons.append("Новый поставщик/нет истории: выбран наименее загруженный доступный сотрудник")

        if supplier_raw == UNKNOWN:
            reasons.append("Поставщик отсутствует")
        if delivery == UNKNOWN:
            reasons.append("Способ доставки отсутствует")
        if country == UNKNOWN:
            reasons.append("Страна отправления отсутствует")
        if latest_cloud_date and ful_date and ful_date < latest_cloud_date:
            reasons.append(f"Дата ФУЛ раньше последней даты в Cloud ({latest_cloud_date.strftime('%d.%m.%Y')})")

        if reasons and any(("отсутствует" in r or "нельзя определить" in r or "раньше" in r or "нет доступных" in r) for r in reasons):
            attention.append([idx, clean(f.get("Номер")), order, supplier_raw,
                              logistic if logistic != UNKNOWN else "Оставить пустым", "; ".join(reasons)])

        new_rows.append([
            order, supplier_raw, brand, logistic, "в работе", UNKNOWN, cutoff,
            clean(f.get("Дата")), country, delivery, UNKNOWN, UNKNOWN
        ])
        journal.append([idx, order, supplier_raw, group_label, logistic, decision,
                        "Создать строку" if logistic != UNKNOWN else "Создать строку без логиста",
                        "; ".join(reasons) if reasons else "Стандартное распределение"])

    return new_rows, repeats, attention, journal

def style_sheet(ws):
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="244062")
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[1].height = 32
    for col in ws.columns:
        max_len = max((len(str(c.value)) if c.value is not None else 0 for c in col), default=10)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max(max_len + 2, 12), 42)
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)

def write_output(output_path, new_rows, repeats, attention, journal):
    wb = Workbook()
    ws = wb.active
    ws.title = "Новые строки Cloud"
    headers = ["Заказ", "Поставщик", "Бренд", "Логист", "Статус", "Дата подтверждения",
               "cut-off", "Дата получения УЛ", "Страна отправления", "Способ доставки",
               "Документы получены", "Комментарий"]
    ws.append(headers)
    for r in new_rows:
        ws.append(r)
    ws2 = wb.create_sheet("Повторные ФУЛ")
    ws2.append(["№ строки ФУЛ", "Заказ", "Поставщик", "Дата ФУЛ", "Комментарий для Cloud"])
    for r in repeats: ws2.append(r)
    ws3 = wb.create_sheet("Требует внимания")
    ws3.append(["№ строки ФУЛ", "Номер ФУЛ", "Заказ", "Поставщик", "Действие", "Причина"])
    for r in attention: ws3.append(r)
    ws4 = wb.create_sheet("Журнал решений")
    ws4.append(["№ строки ФУЛ", "Заказ", "Поставщик", "Группа", "Выбранный логист", "Тип решения", "Действие", "Основание"])
    for r in journal: ws4.append(r)
    for sheet in wb.worksheets: style_sheet(sheet)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)

def main():
    p = argparse.ArgumentParser()
    p.add_argument("--ful", required=True)
    p.add_argument("--cloud", required=True)
    p.add_argument("--employees", required=True)
    p.add_argument("--output-dir", required=True)
    args = p.parse_args()
    fuls, cloud, employees = load_input_files(args.ful, args.cloud, args.employees)
    new_rows, repeats, attention, journal = distribute(fuls, cloud, employees)
    out = Path(args.output_dir) / "result.xlsx"
    write_output(out, new_rows, repeats, attention, journal)
    print(json.dumps({
        "status": "ok", "ful_count": len(fuls), "new_rows": len(new_rows),
        "repeats": len(repeats), "attention": len(attention),
        "output": str(out)
    }, ensure_ascii=False))

if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(json.dumps({"status": "error", "error": str(exc)}, ensure_ascii=False), file=sys.stderr)
        raise
