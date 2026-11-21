from fastapi import FastAPI, HTTPException, BackgroundTasks
from fastapi.responses import HTMLResponse, StreamingResponse
from pydantic import BaseModel
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import os
import io
import json
import math
import re
from datetime import date, timedelta, datetime
import requests
import asyncio
import subprocess
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, landscape
import psycopg2
import psycopg2.extras
from apscheduler.schedulers.background import BackgroundScheduler

app = FastAPI(title="Catering Management API", version="1.1.0")

_MENU_CACHE = {}
_STAFF_MENU_CACHE = {}

TELEGRAM_BOT_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN", "")
TELEGRAM_CHAT_ID = os.environ.get("TELEGRAM_CHAT_ID", "")
TELEGRAM_ADMIN_CHAT_ID = os.environ.get("TELEGRAM_ADMIN_CHAT_ID", "")
MONTH_CLOSE_PASSWORD = os.environ.get("MONTH_CLOSE_PASSWORD", "close2026")


class CloseMonthRequest(BaseModel):
    month: str
    region: str = "Город"
    password: str


def get_db_connection():
    database_url = os.environ.get("DATABASE_URL")
    if database_url and database_url.startswith("postgres://"):
        database_url = database_url.replace("postgres://", "postgresql://", 1)

    if database_url:
        conn = psycopg2.connect(database_url, cursor_factory=psycopg2.extras.RealDictCursor)
    else:
        import sqlite3
        conn = sqlite3.connect("warehouse.db")
        conn.row_factory = sqlite3.Row
    return conn


def _is_postgres(conn) -> bool:
    return conn.__class__.__module__.startswith("psycopg2")


def init_db():
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS app_settings (
            date_str VARCHAR(20) PRIMARY KEY,
            data JSONB
        );
    """)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS stock_acts (
            date_str VARCHAR(20),
            region VARCHAR(50),
            data JSONB,
            PRIMARY KEY (date_str, region)
        );
    """)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS fact_children (
            date_str VARCHAR(20),
            region VARCHAR(50),
            data JSONB,
            PRIMARY KEY (date_str, region)
        );
    """)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS fact_staff (
            date_str VARCHAR(20),
            region VARCHAR(50),
            data JSONB,
            PRIMARY KEY (date_str, region)
        );
    """)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS closed_months (
            month_str VARCHAR(10),
            region VARCHAR(50),
            closed_at VARCHAR(40),
            snapshot JSONB,
            PRIMARY KEY (month_str, region)
        );
    """)
    conn.commit()
    cursor.close()
    conn.close()


scheduler = BackgroundScheduler()


@app.on_event("startup")
def startup_event():
    init_db()
    scheduler.start()


@app.on_event("shutdown")
def shutdown_event():
    scheduler.shutdown()


def is_month_closed(date_str: str, region: str) -> bool:
    if not date_str or len(date_str) < 7:
        return False
    month_str = date_str[:7]
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        ph = "%s" if _is_postgres(conn) else "?"
        cursor.execute(f"SELECT 1 FROM closed_months WHERE month_str = {ph} AND region = {ph};", (month_str, region))
        row = cursor.fetchone()
        cursor.close()
        conn.close()
        return row is not None
    except Exception as e:
        print(f"Ошибка проверки закрытия месяца: {e}")
        return False


def get_closed_month_snapshot(month_str: str, region: str) -> dict | None:
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        ph = "%s" if _is_postgres(conn) else "?"
        cursor.execute(f"SELECT snapshot, closed_at FROM closed_months WHERE month_str = {ph} AND region = {ph};", (month_str, region))
        row = cursor.fetchone()
        cursor.close()
        conn.close()
        if row:
            raw_snap = row["snapshot"]
            snap = raw_snap if isinstance(raw_snap, dict) else json.loads(raw_snap)
            snap["closed_at"] = row["closed_at"]
            return snap
    except Exception as e:
        print(f"Ошибка чтения снимка закрытого месяца: {e}")
    return None


def get_closed_day_snapshot(date_str: str, region: str) -> dict | None:
    if not date_str or len(date_str) < 7:
        return None
    snap = get_closed_month_snapshot(date_str[:7], region)
    if not snap:
        return None
    return snap.get("days", {}).get(date_str)


def assert_month_open(date_str: str, region: str):
    if is_month_closed(date_str, region):
        raise HTTPException(
            status_code=423,
            detail=f"Период за {date_str} ({region}) закрыт и заморожен. Изменение и удаление запрещены."
        )


def load_manual_settings() -> dict:
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        ph = "%s" if _is_postgres(conn) else "?"
        cursor.execute(f"SELECT data FROM app_settings WHERE date_str = {ph}", ("calendar_config",))
        row = cursor.fetchone()
        cursor.close()
        conn.close()
        if row:
            return row["data"] if isinstance(row["data"], dict) else json.loads(row["data"])
    except Exception as e:
        print(f"Ошибка загрузки календаря из БД: {e}")
    return {"days": {}}


def save_manual_settings(settings: dict):
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        json_val = json.dumps(settings, ensure_ascii=False)
        if _is_postgres(conn):
            cursor.execute("""
                INSERT INTO app_settings (date_str, data) VALUES ('calendar_config', %s)
                ON CONFLICT (date_str) DO UPDATE SET data = EXCLUDED.data;
            """, (json_val,))
        else:
            cursor.execute("""
                INSERT INTO app_settings (date_str, data) VALUES ('calendar_config', ?)
                ON CONFLICT (date_str) DO UPDATE SET data = excluded.data;
            """, (json_val,))
        conn.commit()
        cursor.close()
        conn.close()
    except Exception as e:
        print(f"Ошибка сохранения календаря в БД: {e}")


def load_stock_settings() -> dict:
    stock_map = {}
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT date_str, region, data FROM stock_acts;")
        rows = cursor.fetchall()
        cursor.close()
        conn.close()
        for r in rows:
            d = str(r["date_str"]).strip()
            reg = str(r["region"]).strip()
            val = r["data"] if isinstance(r["data"], dict) else json.loads(r["data"])
            if d not in stock_map:
                stock_map[d] = {}
            stock_map[d][reg] = val
    except Exception as e:
        print(f"Ошибка загрузки инвентаризации из БД: {e}")
    return stock_map


def save_stock_settings_for_date(date_str: str, region: str, data: dict):
    conn = get_db_connection()
    cursor = conn.cursor()
    json_val = json.dumps(data, ensure_ascii=False)
    if _is_postgres(conn):
        cursor.execute("""
            INSERT INTO stock_acts (date_str, region, data) VALUES (%s, %s, %s)
            ON CONFLICT (date_str, region) DO UPDATE SET data = EXCLUDED.data;
        """, (date_str, region, json_val))
    else:
        cursor.execute("""
            INSERT INTO stock_acts (date_str, region, data) VALUES (?, ?, ?)
            ON CONFLICT (date_str, region) DO UPDATE SET data = excluded.data;
        """, (date_str, region, json_val))
    conn.commit()
    cursor.close()
    conn.close()


def delete_stock_settings_for_date(date_str: str, region: str):
    conn = get_db_connection()
    cursor = conn.cursor()
    ph = "%s" if _is_postgres(conn) else "?"
    cursor.execute(f"DELETE FROM stock_acts WHERE date_str = {ph} AND region = {ph};", (date_str, region))
    conn.commit()
    cursor.close()
    conn.close()


def load_fact_children_settings() -> dict:
    fact_map = {}
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT date_str, region, data FROM fact_children;")
        rows = cursor.fetchall()
        cursor.close()
        conn.close()
        for r in rows:
            d = str(r["date_str"]).strip()
            reg = str(r["region"]).strip()
            val = r["data"] if isinstance(r["data"], dict) else json.loads(r["data"])
            if d not in fact_map:
                fact_map[d] = {}
            fact_map[d][reg] = val
    except Exception as e:
        print(f"Ошибка загрузки факта детей из БД: {e}")
    return fact_map


def save_fact_children_for_date(date_str: str, region: str, data: dict):
    conn = get_db_connection()
    cursor = conn.cursor()
    json_val = json.dumps(data, ensure_ascii=False)
    if _is_postgres(conn):
        cursor.execute("""
            INSERT INTO fact_children (date_str, region, data) VALUES (%s, %s, %s)
            ON CONFLICT (date_str, region) DO UPDATE SET data = EXCLUDED.data;
        """, (date_str, region, json_val))
    else:
        cursor.execute("""
            INSERT INTO fact_children (date_str, region, data) VALUES (?, ?, ?)
            ON CONFLICT (date_str, region) DO UPDATE SET data = excluded.data;
        """, (date_str, region, json_val))
    conn.commit()
    cursor.close()
    conn.close()


def delete_fact_children_for_date(date_str: str, region: str):
    conn = get_db_connection()
    cursor = conn.cursor()
    ph = "%s" if _is_postgres(conn) else "?"
    cursor.execute(f"DELETE FROM fact_children WHERE date_str = {ph} AND region = {ph};", (date_str, region))
    conn.commit()
    cursor.close()
    conn.close()


def load_fact_staff_settings() -> dict:
    staff_map = {}
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT date_str, region, data FROM fact_staff;")
        rows = cursor.fetchall()
        cursor.close()
        conn.close()
        for r in rows:
            d = str(r["date_str"]).strip()
            reg = str(r["region"]).strip()
            val = r["data"] if isinstance(r["data"], dict) else json.loads(r["data"])
            if d not in staff_map:
                staff_map[d] = {}
            staff_map[d][reg] = val
    except Exception as e:
        print(f"Ошибка загрузки факта сотрудников из БД: {e}")
    return staff_map


def save_fact_staff_for_date(date_str: str, region: str, data: dict):
    conn = get_db_connection()
    cursor = conn.cursor()
    json_val = json.dumps(data, ensure_ascii=False)
    if _is_postgres(conn):
        cursor.execute("""
            INSERT INTO fact_staff (date_str, region, data) VALUES (%s, %s, %s)
            ON CONFLICT (date_str, region) DO UPDATE SET data = EXCLUDED.data;
        """, (date_str, region, json_val))
    else:
        cursor.execute("""
            INSERT INTO fact_staff (date_str, region, data) VALUES (?, ?, ?)
            ON CONFLICT (date_str, region) DO UPDATE SET data = excluded.data;
        """, (date_str, region, json_val))
    conn.commit()
    cursor.close()
    conn.close()


def delete_fact_staff_for_date(date_str: str, region: str):
    conn = get_db_connection()
    cursor = conn.cursor()
    ph = "%s" if _is_postgres(conn) else "?"
    cursor.execute(f"DELETE FROM fact_staff WHERE date_str = {ph} AND region = {ph};", (date_str, region))
    conn.commit()
    cursor.close()
    conn.close()


def load_purchases_settings() -> dict:
    purchases_by_region = {"Город": {}, "Район": {}}
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        query = """
            SELECT d.date, c.category, c.client_number, i.name as product_name, si.quantity
            FROM sales_docs d
            JOIN clients c ON d.client_id = c.id
            JOIN sales_items si ON si.doc_id = d.id
            JOIN items i ON si.item_id = i.id
            WHERE d.organization = 'MUXSIN-FAYZ'
        """
        cursor.execute(query)
        rows = cursor.fetchall()
        cursor.close()
        conn.close()

        points_gorod = load_points_from_excel("Город")
        points_rayon = load_points_from_excel("Район")

        for r in rows:
            date_str = str(r["date"]).split("T")[0] if r["date"] else ""
            category_raw = str(r["category"]).strip().lower()

            region = "Город" if ("город" in category_raw or "текущие" in category_raw) else "Район"

            points_dict = points_gorod if region == "Город" else points_rayon
            db_client_num = int(r["client_number"])

            target_point_id = None
            for pid, pinfo in points_dict.items():
                name_str = str(pinfo["name"])
                numbers = re.findall(r'\d+', name_str)
                if numbers and int(numbers[0]) == db_client_num:
                    target_point_id = pid
                    break

            if target_point_id:
                if region not in purchases_by_region:
                    purchases_by_region[region] = {}
                if target_point_id not in purchases_by_region[region]:
                    purchases_by_region[region][target_point_id] = []

                purchases_by_region[region][target_point_id].append({
                    "date": date_str,
                    "product": r["product_name"],
                    "qty": float(r["quantity"])
                })
    except Exception as e:
        print(f"Ошибка чтения приходов из базы отгрузок MUXSIN-FAYZ: {e}")

    return purchases_by_region


def calculate_calendar_dynamic(start_date: date, end_date: date) -> dict:
    settings = load_manual_settings()
    manual_days = settings.get("days", {})
    global_start = date(2026, 1, 1)
    calc_start = min(global_start, start_date)

    current_menu = 1
    calendar_map = {}
    current = calc_start

    while current <= end_date:
        d_str = current.isoformat()
        day_conf = manual_days.get(d_str, {})
        is_holiday = day_conf.get("is_holiday", False)
        if current.weekday() >= 5 and "is_holiday" not in day_conf:
            is_holiday = True
        forced_menu = day_conf.get("menu", None)

        if is_holiday:
            if current >= start_date:
                calendar_map[d_str] = {
                    "date": d_str,
                    "is_holiday": True,
                    "menu": None,
                    "forced": False
                }
        else:
            if forced_menu is not None:
                current_menu = int(forced_menu)
                if current >= start_date:
                    calendar_map[d_str] = {
                        "date": d_str,
                        "is_holiday": False,
                        "menu": current_menu,
                        "forced": True
                    }
            else:
                if current >= start_date:
                    calendar_map[d_str] = {
                        "date": d_str,
                        "is_holiday": False,
                        "menu": current_menu,
                        "forced": False
                    }
            current_menu = (current_menu % 10) + 1
        current += timedelta(days=1)
    return calendar_map


def get_menu_number_for_date_custom(target_date: date) -> int | None:
    closed_day = get_closed_day_snapshot(target_date.isoformat(), "Город")
    if closed_day and "menu_num" in closed_day:
        return closed_day["menu_num"]
    closed_day_r = get_closed_day_snapshot(target_date.isoformat(), "Район")
    if closed_day_r and "menu_num" in closed_day_r:
        return closed_day_r["menu_num"]

    start = date(2026, 1, 1)
    cal = calculate_calendar_dynamic(start, target_date + timedelta(days=30))
    d_str = target_date.isoformat()
    if d_str in cal:
        day_info = cal[d_str]
        if day_info["is_holiday"]:
            return None
        return day_info["menu"]
    return None


def load_points_from_excel(region: str = "Город") -> dict:
    try:
        df = pd.read_excel("Location.xlsx", sheet_name=region, header=None)
        points = {}
        for idx, row in df.iterrows():
            mtt_name = str(row.iloc[0]).strip()
            if mtt_name and mtt_name != 'nan':
                children = 0
                if len(row) > 1 and pd.notna(row.iloc[1]):
                    try:
                        children = int(float(row.iloc[1]))
                    except ValueError:
                        children = 0
                points[idx + 1] = {"name": mtt_name, "children_count": children}
        return points
    except Exception as e:
        print(f"Ошибка чтения точек ({region}): {e}")
        return {}


def load_menu_from_excel(menu_num: int) -> dict:
    if menu_num in _MENU_CACHE:
        return _MENU_CACHE[menu_num]

    sheet_name = f"{menu_num}-Меню"
    try:
        df = pd.read_excel("Menu.xlsx", sheet_name=sheet_name)
        menu_items = {}
        for _, row in df.iterrows():
            prod_name = str(row.iloc[0]).strip()
            val = float(row.iloc[1]) if pd.notna(row.iloc[1]) else 0.0
            unit = str(row.iloc[2]).strip().lower() if len(row) > 2 and pd.notna(row.iloc[2]) else "граммов"

            if "тухум" in prod_name.lower() or "яйц" in prod_name.lower():
                qty_in_units = val / 55.0
                display_unit = "шт"
                menu_items[prod_name] = {"qty_per_child": qty_in_units, "unit": display_unit}
            elif "грам" in unit or "гр" in unit:
                qty_in_kg = val / 1000.0
                display_unit = "кг"
                menu_items[prod_name] = {"qty_per_child": qty_in_kg, "unit": display_unit}
            else:
                qty_in_kg = val
                display_unit = unit if unit else "шт"
                menu_items[prod_name] = {"qty_per_child": qty_in_kg, "unit": display_unit}

        _MENU_CACHE[menu_num] = menu_items
        return menu_items
    except Exception as e:
        print(f"Ошибка чтения листа {sheet_name}: {e}")
        return {}


def load_staff_menu_from_excel(menu_num: int) -> dict:
    if menu_num in _STAFF_MENU_CACHE:
        return _STAFF_MENU_CACHE[menu_num]

    sheet_name = f"{menu_num}-Меню"
    try:
        df = pd.read_excel("Menu_staff.xlsx", sheet_name=sheet_name)
        menu_items = {}
        for _, row in df.iterrows():
            prod_name = str(row.iloc[0]).strip()
            val = float(row.iloc[1]) if pd.notna(row.iloc[1]) else 0.0
            unit = str(row.iloc[2]).strip().lower() if len(row) > 2 and pd.notna(row.iloc[2]) else "граммов"

            if "тухум" in prod_name.lower() or "яйц" in prod_name.lower():
                qty_in_units = val / 55.0
                display_unit = "шт"
                menu_items[prod_name] = {"qty_per_staff": qty_in_units, "unit": display_unit}
            elif "грам" in unit or "гр" in unit:
                qty_in_kg = val / 1000.0
                display_unit = "кг"
                menu_items[prod_name] = {"qty_per_staff": qty_in_kg, "unit": display_unit}
            else:
                qty_in_kg = val
                display_unit = unit if unit else "шт"
                menu_items[prod_name] = {"qty_per_staff": qty_in_kg, "unit": display_unit}

        _STAFF_MENU_CACHE[menu_num] = menu_items
        return menu_items
    except Exception as e:
        print(f"Ошибка чтения листа сотрудников {sheet_name}: {e}")
        return {}


def load_all_nomenclature_periodicities() -> dict:
    mapping = {}
    try:
        df = pd.read_excel("Menu.xlsx", sheet_name="Все номенклатуры")
        for _, row in df.iterrows():
            prod = str(row.iloc[0]).strip()
            period = str(row.iloc[1]).strip()
            if prod and prod != 'nan':
                mapping[prod] = period
    except Exception as e:
        print(f"Ошибка чтения листа 'Все номенклатуры': {e}")
    return mapping


def get_all_products_with_units() -> dict:
    products = {}
    for i in range(1, 11):
        for items in (load_menu_from_excel(i), load_staff_menu_from_excel(i)):
            for p, info in items.items():
                if p not in products:
                    products[p] = info["unit"]
    return dict(sorted(products.items()))


def get_closed_units_map(region: str) -> dict:
    """Единицы измерения продуктов из снимков закрытых месяцев.

    Нужна, чтобы продукты старого меню (которых уже нет в Excel) в отчётах
    остатков показывались с правильной единицей измерения.
    """
    units = {}
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        ph = "%s" if _is_postgres(conn) else "?"
        cursor.execute(
            f"SELECT snapshot FROM closed_months WHERE region = {ph} ORDER BY month_str DESC;",
            (region,),
        )
        rows = cursor.fetchall()
        cursor.close()
        conn.close()
    except Exception as e:
        print(f"Ошибка чтения единиц из снимков закрытых месяцев: {e}")
        return units

    for r in rows:
        try:
            raw = r["snapshot"]
            snap = raw if isinstance(raw, dict) else json.loads(raw)
        except Exception:
            continue
        for day in snap.get("days", {}).values():
            for key in ("children_consumption", "staff_consumption"):
                for prod_name, pdata in day.get(key, {}).items():
                    if prod_name not in units:
                        units[prod_name] = pdata.get("unit", "кг")
    return units


def _safe_int_or_zero(val) -> int:
    if val is None or val == "":
        return 0
    try:
        v = int(float(val))
        return v if v > 0 else 0
    except Exception:
        return 0


def _close_single_month(month: str, region: str) -> dict:
    try:
        parts = month.split("-")
        year, m = int(parts[0]), int(parts[1])
        if not (1 <= m <= 12):
            raise ValueError("Месяц вне диапазона")
        start_d = date(year, m, 1)
        end_d = date(year + 1, 1, 1) - timedelta(days=1) if m == 12 else date(year, m + 1, 1) - timedelta(days=1)
    except Exception:
        raise HTTPException(status_code=400, detail="Неверный формат месяца (ожидается YYYY-MM)")

    if get_closed_month_snapshot(month, region) is not None:
        raise HTTPException(status_code=400, detail=f"Месяц {month} ({region}) уже закрыт и заморожен навсегда.")

    points_dict = load_points_from_excel(region)
    fact_children_all = load_fact_children_settings()
    fact_staff_all = load_fact_staff_settings()

    days_snapshot = {}
    curr = start_d
    while curr <= end_d:
        d_str = curr.isoformat()
        menu_num = get_menu_number_for_date_custom(curr)
        saved_ch = fact_children_all.get(d_str, {}).get(region, {})
        saved_st = fact_staff_all.get(d_str, {}).get(region, {})

        children_counts = {}
        staff_counts = {}
        for pid in points_dict.keys():
            raw_ch = saved_ch.get(str(pid), saved_ch.get(pid, 0))
            raw_st = saved_st.get(str(pid), saved_st.get(pid, 0))
            children_counts[str(pid)] = _safe_int_or_zero(raw_ch)
            staff_counts[str(pid)] = _safe_int_or_zero(raw_st)

        children_consumption = {}
        staff_consumption = {}
        if menu_num is not None:
            menu_ch = load_menu_from_excel(menu_num)
            for prod_name, pdata in menu_ch.items():
                unit = pdata.get("unit", "кг")
                q_per = pdata.get("qty_per_child", 0.0)
                pts_q = {}
                total_q = 0.0
                for pid in points_dict.keys():
                    cnt = children_counts[str(pid)]
                    q = float(math.ceil(cnt * q_per)) if unit == "шт" else round(cnt * q_per, 2)
                    pts_q[str(pid)] = q
                    total_q += q
                children_consumption[prod_name] = {
                    "unit": unit,
                    "points_qty": pts_q,
                    "total": float(math.ceil(total_q)) if unit == "шт" else round(total_q, 2)
                }

            menu_st = load_staff_menu_from_excel(menu_num)
            for prod_name, pdata in menu_st.items():
                unit = pdata.get("unit", "кг")
                q_per = pdata.get("qty_per_staff", 0.0)
                pts_q = {}
                total_q = 0.0
                for pid in points_dict.keys():
                    cnt = staff_counts[str(pid)]
                    q = float(math.ceil(cnt * q_per)) if unit == "шт" else round(cnt * q_per, 2)
                    pts_q[str(pid)] = q
                    total_q += q
                staff_consumption[prod_name] = {
                    "unit": unit,
                    "points_qty": pts_q,
                    "total": float(math.ceil(total_q)) if unit == "шт" else round(total_q, 2)
                }

        days_snapshot[d_str] = {
            "menu_num": menu_num,
            "children_counts": children_counts,
            "staff_counts": staff_counts,
            "children_consumption": children_consumption,
            "staff_consumption": staff_consumption
        }
        curr += timedelta(days=1)

    closed_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    snapshot_payload = {
        "month": month,
        "region": region,
        "closed_at": closed_at,
        "days": days_snapshot
    }

    conn = get_db_connection()
    cursor = conn.cursor()
    json_val = json.dumps(snapshot_payload, ensure_ascii=False)
    if _is_postgres(conn):
        cursor.execute("""
            INSERT INTO closed_months (month_str, region, closed_at, snapshot)
            VALUES (%s, %s, %s, %s)
            ON CONFLICT (month_str, region) DO NOTHING;
        """, (month, region, closed_at, json_val))
    else:
        cursor.execute("""
            INSERT INTO closed_months (month_str, region, closed_at, snapshot)
            VALUES (?, ?, ?, ?)
            ON CONFLICT (month_str, region) DO NOTHING;
        """, (month, region, closed_at, json_val))
    conn.commit()
    cursor.close()
    conn.close()

    return {
        "month": month,
        "region": region,
        "closed_at": closed_at,
        "days_count": len(days_snapshot)
    }


def generate_menu_pdf_for_point(report_date_str: str, region: str, point_id: int, point_name: str, children_count: int, menu_num: int) -> str:
    template_path = "Menu shablon.xlsx"
    if not os.path.exists(template_path):
        raise FileNotFoundError("Шаблон Menu shablon.xlsx не найден")

    try:
        d_obj = date.fromisoformat(report_date_str)
        formatted_date = d_obj.strftime("%d.%m.%Y")
    except Exception:
        formatted_date = report_date_str

    wb = openpyxl.load_workbook(template_path)
    target_sheet_name = str(menu_num)

    if target_sheet_name not in wb.sheetnames:
        raise ValueError(f"Лист меню {target_sheet_name} не найден в шаблоне")

    for sheetname in wb.sheetnames:
        if sheetname != target_sheet_name:
            del wb[sheetname]

    ws = wb[target_sheet_name]
    ws['A1'] = formatted_date
    ws['A4'] = children_count
    ws['F1'] = point_name

    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.horizontalCentered = True
    ws.page_setup.verticalCentered = True
    ws.page_margins.left = 0.4
    ws.page_margins.right = 0.4
    ws.page_margins.top = 1.5
    ws.page_margins.bottom = 0.3
    ws.page_margins.header = 0.0
    ws.page_margins.footer = 0.0

    temp_excel_path = f"temp_{region}_{point_id}_{report_date_str}.xlsx"
    final_pdf_path = f"Menu_{point_name}_{report_date_str}.pdf"

    wb.save(temp_excel_path)
    wb.close()

    subprocess.run(["libreoffice", "--headless", "--convert-to", "pdf", temp_excel_path, "--outdir", "."], check=True)

    generated_pdf = temp_excel_path.replace(".xlsx", ".pdf")
    if os.path.exists(generated_pdf):
        if os.path.exists(final_pdf_path):
            os.remove(final_pdf_path)
        os.rename(generated_pdf, final_pdf_path)
        if os.path.exists(temp_excel_path):
            os.remove(temp_excel_path)
        return final_pdf_path
    return temp_excel_path


def generate_staff_menu_pdf_for_point(report_date_str: str, region: str, point_id: int, point_name: str, staff_count: int, menu_num: int) -> str:
    template_path = "Menu shablon_staff.xlsx"
    if not os.path.exists(template_path):
        raise FileNotFoundError("Шаблон Menu shablon_staff.xlsx не найден")

    try:
        d_obj = date.fromisoformat(report_date_str)
        formatted_date = d_obj.strftime("%d.%m.%Y")
    except Exception:
        formatted_date = report_date_str

    wb = openpyxl.load_workbook(template_path)
    target_sheet_name = str(menu_num)

    if target_sheet_name not in wb.sheetnames:
        raise ValueError(f"Лист меню {target_sheet_name} не найден в шаблоне сотрудников")

    for sheetname in wb.sheetnames:
        if sheetname != target_sheet_name:
            del wb[sheetname]

    ws = wb[target_sheet_name]
    ws['A1'] = formatted_date
    ws['A4'] = staff_count
    ws['F1'] = point_name

    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.horizontalCentered = True
    ws.page_setup.verticalCentered = True
    ws.page_margins.left = 0.4
    ws.page_margins.right = 0.4
    ws.page_margins.top = 1.5
    ws.page_margins.bottom = 0.3
    ws.page_margins.header = 0.0
    ws.page_margins.footer = 0.0

    temp_excel_path = f"temp_staff_{region}_{point_id}_{report_date_str}.xlsx"
    final_pdf_path = f"Menu_Staff_{point_name}_{report_date_str}.pdf"

    wb.save(temp_excel_path)
    wb.close()

    subprocess.run(["libreoffice", "--headless", "--convert-to", "pdf", temp_excel_path, "--outdir", "."], check=True)

    generated_pdf = temp_excel_path.replace(".xlsx", ".pdf")
    if os.path.exists(generated_pdf):
        if os.path.exists(final_pdf_path):
            os.remove(final_pdf_path)
        os.rename(generated_pdf, final_pdf_path)
        if os.path.exists(temp_excel_path):
            os.remove(temp_excel_path)
        return final_pdf_path
    return temp_excel_path


def send_pdf_to_telegram(pdf_path: str, caption: str):
    target_chat_id = TELEGRAM_ADMIN_CHAT_ID or TELEGRAM_CHAT_ID
    if not TELEGRAM_BOT_TOKEN or not target_chat_id:
        print("❌ ОШИБКА: Не задан TELEGRAM_BOT_TOKEN или TELEGRAM_ADMIN_CHAT_ID")
        return None

    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendDocument"
    try:
        with open(pdf_path, "rb") as f:
            files = {"document": f}
            data = {"chat_id": target_chat_id, "caption": caption}
            response = requests.post(url, data=data, files=files)
        res_json = response.json()
        if not res_json.get("ok"):
            print(f"❌ ОШИБКА от Telegram API для файла {pdf_path}: {res_json}")
        else:
            print(f"✅ Успешно отправлено в Telegram: {pdf_path}")
        return res_json
    except Exception as e:
        print(f"❌ Исключение при отправке в Telegram файла {pdf_path}: {e}")
        return None


async def process_and_send_menu_in_background(report_date: str, region: str, saved_fact: dict, menu_num: int):
    points_dict = load_points_from_excel(region)
    for pid, pinfo in points_dict.items():
        child_cnt = saved_fact.get(str(pid), saved_fact.get(pid))
        if child_cnt is None:
            continue
        p_name = pinfo["name"]
        try:
            pdf_file = generate_menu_pdf_for_point(
                report_date_str=report_date,
                region=region,
                point_id=pid,
                point_name=p_name,
                children_count=int(child_cnt),
                menu_num=menu_num
            )
            caption = f"📋 Официальное меню № {menu_num} для {p_name} ({region})\n📅 Дата: {report_date}\n👶 Детей: {child_cnt}"
            send_pdf_to_telegram(pdf_file, caption)
            if pdf_file and os.path.exists(pdf_file):
                os.remove(pdf_file)
            await asyncio.sleep(2)
        except Exception as e:
            print(f"❌ Ошибка при обработке точки {p_name}: {e}")


async def process_and_send_staff_menu_in_background(report_date: str, region: str, saved_fact: dict, menu_num: int):
    points_dict = load_points_from_excel(region)
    for pid, pinfo in points_dict.items():
        staff_cnt = saved_fact.get(str(pid), saved_fact.get(pid))
        if staff_cnt is None or int(staff_cnt) <= 0:
            continue
        p_name = pinfo["name"]
        try:
            pdf_file = generate_staff_menu_pdf_for_point(
                report_date_str=report_date,
                region=region,
                point_id=pid,
                point_name=p_name,
                staff_count=int(staff_cnt),
                menu_num=menu_num
            )
            caption = f"📋 Меню сотрудников № {menu_num} для {p_name} ({region})\n📅 Дата: {report_date}\n👥 Сотрудников: {staff_cnt}"
            send_pdf_to_telegram(pdf_file, caption)
            if pdf_file and os.path.exists(pdf_file):
                os.remove(pdf_file)
            await asyncio.sleep(2)
        except Exception as e:
            print(f"❌ Ошибка при обработке точки сотрудников {p_name}: {e}")


@app.api_route("/", methods=["GET", "HEAD"], response_class=HTMLResponse)
def get_dashboard():
    html_content = """
    <!DOCTYPE html>
    <html lang="ru">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Разнарядка Кетеринг</title>
        <script src="https://cdn.tailwindcss.com"></script>
    </head>
    <body class="bg-gray-50 text-gray-800 font-sans">
        <div class="max-w-[95%] mx-auto px-4 py-8">
            <nav class="bg-white shadow rounded-lg p-4 mb-6 flex justify-between items-center flex-wrap gap-3">
                <div class="flex items-center space-x-3">
                    <span class="text-xl font-bold text-blue-600">🍽 Кетеринг ERP</span>
                    <span class="text-xs bg-blue-100 text-blue-800 px-2.5 py-1 rounded-full font-medium">MUXSIN-FAYZ</span>
                </div>
                <div class="flex flex-wrap gap-2">
                    <button onclick="switchTab('matrix')" id="btnTabMatrix" class="px-3 py-2 rounded-md text-sm font-medium bg-blue-600 text-white transition">📊 Матрица разнарядки</button>
                    <button onclick="switchTab('fact')" id="btnTabFact" class="px-3 py-2 rounded-md text-sm font-medium bg-gray-200 text-gray-700 hover:bg-gray-300 transition">📋 Ежедневный факт (Дети)</button>
                    <button onclick="switchTab('staffFact')" id="btnTabStaffFact" class="px-3 py-2 rounded-md text-sm font-medium bg-gray-200 text-gray-700 hover:bg-gray-300 transition">👥 Факт (Сотрудники)</button>
                    <button onclick="switchTab('balance')" id="btnTabBalance" class="px-3 py-2 rounded-md text-sm font-medium bg-gray-200 text-gray-700 hover:bg-gray-300 transition">📦 Остатки по МТТ</button>
                    <button onclick="switchTab('calendar')" id="btnTabCalendar" class="px-3 py-2 rounded-md text-sm font-medium bg-gray-200 text-gray-700 hover:bg-gray-300 transition">📅 Календарь меню</button>
                    <button onclick="switchTab('stock')" id="btnTabStock" class="px-3 py-2 rounded-md text-sm font-medium bg-gray-200 text-gray-700 hover:bg-gray-300 transition">⚙️ Инвентаризация</button>
                    <button onclick="switchTab('monthClose')" id="btnTabMonthClose" class="px-3 py-2 rounded-md text-sm font-medium bg-amber-100 text-amber-900 hover:bg-amber-200 transition">🔒 Закрытие месяца</button>
                </div>
            </nav>

            <!-- МАТРИЦА РАЗНАРЯДКИ -->
            <div id="tabMatrix" class="space-y-6">
                <header class="bg-white shadow rounded-lg p-6 flex flex-col justify-between">
                    <div>
                        <h1 class="text-xl font-bold text-gray-800">Матричная разнарядка (Чистая потребность)</h1>
                        <p class="text-sm text-gray-500 mt-1">Автоматический учет остатков, факта детей и сотрудников за последние дни и приходов от MUXSIN-FAYZ</p>
                    </div>
                </header>

                <div class="bg-white shadow rounded-lg p-4 flex flex-wrap items-center justify-between gap-4">
                    <div class="flex flex-wrap items-center gap-2">
                        <span class="text-xs font-bold text-gray-500 mr-1">Номенклатура:</span>
                        <button onclick="setMatrixFilter('ALL')" id="filterBtnALL" class="px-3 py-2 rounded-md text-xs font-bold bg-blue-600 text-white transition">Все позиции</button>
                        <button onclick="setMatrixFilter('Ежедневно')" id="filterBtnЕжедневно" class="px-3 py-2 rounded-md text-xs font-bold bg-gray-200 text-gray-700 hover:bg-gray-300 transition">Ежедневно</button>
                        <button onclick="setMatrixFilter('Неделья')" id="filterBtnНеделья" class="px-3 py-2 rounded-md text-xs font-bold bg-gray-200 text-gray-700 hover:bg-gray-300 transition">Неделья</button>
                        <button onclick="setMatrixFilter('Месяц')" id="filterBtnМесяц" class="px-3 py-2 rounded-md text-xs font-bold bg-gray-200 text-gray-700 hover:bg-gray-300 transition">Месяц</button>
                    </div>

                    <div class="flex flex-wrap items-center gap-3">
                        <div>
                            <label class="text-xs font-medium text-gray-600 block">Регион:</label>
                            <select id="regionSelect" class="border border-gray-300 rounded-md px-3 py-2 text-sm focus:ring-2 focus:ring-blue-500">
                                <option value="Город" selected>Город (Текущие)</option>
                                <option value="Район">Район (Тендер / Будущие)</option>
                            </select>
                        </div>
                        <div>
                            <label class="text-xs font-medium text-gray-600 block">Начальная дата:</label>
                            <input type="date" id="targetDate" class="border border-gray-300 rounded-md px-3 py-2 text-sm focus:ring-2 focus:ring-blue-500">
                        </div>
                        <div>
                            <label class="text-xs font-medium text-gray-600 block">Кол-во дней:</label>
                            <input type="number" id="daysCountInput" value="2" min="1" max="30" class="border border-gray-300 rounded-md px-3 py-2 text-sm w-20 text-center font-bold">
                        </div>
                        <div class="self-end flex gap-2">
                            <button onclick="loadMatrixRaznaryadka()" class="bg-blue-600 text-white px-4 py-2 rounded-md text-sm font-medium hover:bg-blue-700 transition">Сформировать</button>
                            <button onclick="downloadRaznaryadkaExcel()" class="bg-emerald-600 text-white px-4 py-2 rounded-md text-sm font-medium hover:bg-emerald-700 transition">📥 Скачать Excel</button>
                        </div>
                    </div>
                </div>

                <div id="matrixResultContainer" class="space-y-6">
                    <div class="bg-white shadow rounded-lg p-8 text-center text-gray-400">Выберите параметры и нажмите «Сформировать».</div>
                </div>
            </div>

            <!-- ФАКТ ДЕТИ -->
            <div id="tabFact" class="hidden space-y-6">
                <header class="bg-white shadow rounded-lg p-6 flex flex-col md:flex-row justify-between items-center">
                    <div>
                        <h1 class="text-xl font-bold text-gray-800">📋 Фактический отчет расхода продуктов (Дети)</h1>
                        <p class="text-sm text-gray-500 mt-1">Утренний ввод детей по факту с автоматическим списанием по календарю меню</p>
                    </div>
                    <div class="mt-4 md:mt-0 flex flex-wrap items-center gap-3">
                        <div>
                            <label class="text-xs font-medium text-gray-600 block">Регион:</label>
                            <select id="factRegionSelect" onchange="loadFactData()" class="border border-gray-300 rounded-md px-3 py-2 text-sm">
                                <option value="Город" selected>Город</option>
                                <option value="Район">Район</option>
                            </select>
                        </div>
                        <div>
                            <label class="text-xs font-medium text-gray-600 block">Дата отчета:</label>
                            <input type="date" id="factDate" onchange="loadFactData()" class="border border-gray-300 rounded-md px-3 py-2 text-sm">
                        </div>
                        <div class="self-end flex flex-wrap gap-2">
                            <button onclick="saveFactChildren()" id="btnSaveFactChildren" class="bg-green-600 text-white px-4 py-2 rounded-md text-sm font-medium hover:bg-green-700 transition">💾 Сохранить факт</button>
                            <button onclick="downloadFactAllRegionsExcel()" class="bg-emerald-600 text-white px-4 py-2 rounded-md text-sm font-medium hover:bg-emerald-700 transition">📥 Скачать факт</button>
                            <button onclick="sendMenuTelegramManual()" class="bg-indigo-600 text-white px-4 py-2 rounded-md text-sm font-medium hover:bg-indigo-700 transition">📤 Отправить меню в Telegram</button>
                        </div>
                    </div>
                </header>

                <div class="bg-white shadow rounded-lg p-6 flex flex-col md:flex-row justify-between items-center border-l-4 border-purple-500">
                    <div>
                        <h3 class="font-bold text-gray-800 text-md">📊 Выгрузка отчетов за месяц</h3>
                        <p class="text-xs text-gray-500 mt-0.5">Выберите месяц для формирования посещаемости или детального расхода по МТТ</p>
                    </div>
                    <div class="mt-3 md:mt-0 flex items-center gap-3">
                        <input type="month" id="childrenMonthPicker" class="border border-gray-300 rounded-md px-3 py-2 text-sm">
                        <button onclick="downloadChildrenMatrixExcel()" class="bg-purple-600 text-white px-4 py-2 rounded-md text-sm font-medium hover:bg-purple-700 transition">📥 Скачать детей</button>
                        <button onclick="downloadMonthlyMttExcel()" class="bg-blue-600 text-white px-4 py-2 rounded-md text-sm font-medium hover:bg-blue-700 transition">📥 Скачать расход</button>
                    </div>
                </div>

                <div id="factContainer" class="space-y-6">
                    <p class="p-8 text-center text-gray-400 bg-white shadow rounded-lg">Загрузка данных...</p>
                </div>
            </div>

            <!-- ФАКТ СОТРУДНИКИ -->
            <div id="tabStaffFact" class="hidden space-y-6">
                <header class="bg-white shadow rounded-lg p-6 flex flex-col md:flex-row justify-between items-center">
                    <div>
                        <h1 class="text-xl font-bold text-gray-800">👥 Фактический расход продуктов (Сотрудники МТТ)</h1>
                        <p class="text-sm text-gray-500 mt-1">Ввод сотрудников по факту с автоматическим списанием остатков и отправкой меню</p>
                    </div>
                    <div class="mt-4 md:mt-0 flex flex-wrap items-center gap-3">
                        <div>
                            <label class="text-xs font-medium text-gray-600 block">Регион:</label>
                            <select id="staffFactRegionSelect" onchange="loadStaffFactData()" class="border border-gray-300 rounded-md px-3 py-2 text-sm">
                                <option value="Город" selected>Город</option>
                                <option value="Район">Район</option>
                            </select>
                        </div>
                        <div>
                            <label class="text-xs font-medium text-gray-600 block">Дата отчета:</label>
                            <input type="date" id="staffFactDate" onchange="loadStaffFactData()" class="border border-gray-300 rounded-md px-3 py-2 text-sm">
                        </div>
                        <div class="self-end flex flex-wrap gap-2">
                            <button onclick="saveStaffFact()" id="btnSaveStaffFact" class="bg-green-600 text-white px-4 py-2 rounded-md text-sm font-medium hover:bg-green-700 transition">💾 Сохранить факт</button>
                            <button onclick="downloadStaffFactAllRegionsExcel()" class="bg-emerald-600 text-white px-4 py-2 rounded-md text-sm font-medium hover:bg-emerald-700 transition">📥 Скачать факт</button>
                            <button onclick="sendStaffMenuTelegramManual()" class="bg-indigo-600 text-white px-4 py-2 rounded-md text-sm font-medium hover:bg-indigo-700 transition">📤 Отправить меню в Telegram</button>
                        </div>
                    </div>
                </header>

                <div class="bg-white shadow rounded-lg p-6 flex flex-col md:flex-row justify-between items-center border-l-4 border-blue-500">
                    <div>
                        <h3 class="font-bold text-gray-800 text-md">📊 Выгрузка отчетов за месяц (Сотрудники)</h3>
                        <p class="text-xs text-gray-500 mt-0.5">Выберите месяц для формирования количества сотрудников или детального расхода по МТТ</p>
                    </div>
                    <div class="mt-3 md:mt-0 flex items-center gap-3">
                        <input type="month" id="staffMonthPicker" class="border border-gray-300 rounded-md px-3 py-2 text-sm">
                        <button onclick="downloadStaffMatrixExcel()" class="bg-purple-600 text-white px-4 py-2 rounded-md text-sm font-medium hover:bg-purple-700 transition">📥 Скачать сотрудников</button>
                        <button onclick="downloadStaffMonthlyMttExcel()" class="bg-blue-600 text-white px-4 py-2 rounded-md text-sm font-medium hover:bg-blue-700 transition">📥 Скачать расход</button>
                    </div>
                </div>

                <div id="staffFactContainer" class="space-y-6">
                    <p class="p-8 text-center text-gray-400 bg-white shadow rounded-lg">Загрузка данных...</p>
                </div>
            </div>

            <!-- ОСТАТКИ -->
            <div id="tabBalance" class="hidden space-y-6">
                <header class="bg-white shadow rounded-lg p-6 flex flex-col md:flex-row justify-between items-center">
                    <div>
                        <h1 class="text-xl font-bold text-gray-800">📦 Актуальные остатки продуктов по МТТ</h1>
                        <p class="text-sm text-gray-500 mt-1">Формула: Остаток = Начальный остаток + Приход (MUXSIN-FAYZ) - Расход фактический (дети и сотрудники)</p>
                    </div>
                    <div class="mt-4 md:mt-0 flex flex-wrap items-center gap-3">
                        <div>
                            <label class="text-xs font-medium text-gray-600 block">Регион:</label>
                            <select id="balanceRegionSelect" onchange="loadBalanceData()" class="border border-gray-300 rounded-md px-3 py-2 text-sm">
                                <option value="Город" selected>Город</option>
                                <option value="Район">Район</option>
                            </select>
                        </div>
                        <div>
                            <label class="text-xs font-medium text-gray-600 block">На дату:</label>
                            <input type="date" id="balanceDate" onchange="loadBalanceData()" class="border border-gray-300 rounded-md px-3 py-2 text-sm">
                        </div>
                        <div class="self-end flex gap-2">
                            <button onclick="loadBalanceData()" class="bg-indigo-600 text-white px-4 py-2 rounded-md text-sm font-medium hover:bg-indigo-700 transition">Обновить остатки</button>
                            <button onclick="downloadBalanceExcel()" class="bg-emerald-600 text-white px-4 py-2 rounded-md text-sm font-medium hover:bg-emerald-700 transition">📥 Скачать Excel</button>
                        </div>
                    </div>
                </header>
                <div id="balanceContainer" class="space-y-6">
                    <p class="p-8 text-center text-gray-400 bg-white shadow rounded-lg">Загрузка остатков...</p>
                </div>
            </div>

            <!-- КАЛЕНДАРЬ МЕНЮ -->
            <div id="tabCalendar" class="hidden space-y-6">
                <header class="bg-white shadow rounded-lg p-6 flex flex-col md:flex-row justify-between items-center">
                    <div>
                        <h1 class="text-xl font-bold text-gray-800">Интерактивный календарь меню</h1>
                        <p class="text-sm text-gray-500 mt-1">Кликните на любой день, чтобы изменить его статус или номер меню</p>
                    </div>
                    <div class="flex items-center space-x-3 mt-4 md:mt-0">
                        <button onclick="changeMonth(-1)" class="bg-gray-200 px-3 py-1.5 rounded text-sm font-bold hover:bg-gray-300">◄ Месяц</button>
                        <span id="currentMonthLabel" class="font-bold text-lg text-blue-600"></span>
                        <button onclick="changeMonth(1)" class="bg-gray-200 px-3 py-1.5 rounded text-sm font-bold hover:bg-gray-300">Месяц ►</button>
                    </div>
                </header>
                <div class="bg-white shadow rounded-lg p-6">
                    <div class="grid grid-cols-7 gap-2 mb-3 text-center font-bold text-sm text-gray-600">
                        <div>Пн</div><div>Вт</div><div>Ср</div><div>Чт</div><div>Пт</div><div class="text-red-500">Сб</div><div class="text-red-500">Вс</div>
                    </div>
                    <div id="calendarGrid" class="grid grid-cols-7 gap-3"></div>
                </div>
            </div>

            <!-- ИНВЕНТАРИЗАЦИЯ -->
            <div id="tabStock" class="hidden space-y-6">
                <header class="bg-white shadow rounded-lg p-6 flex flex-col md:flex-row justify-between items-center">
                    <div>
                        <h1 class="text-xl font-bold text-gray-800">Акт инвентаризации склада</h1>
                        <p class="text-sm text-gray-500 mt-1">Укажите дату ревизии и фактические остатки продуктов по точкам</p>
                    </div>
                    <div class="flex flex-wrap items-center gap-3 mt-4 md:mt-0">
                        <div>
                            <label class="text-xs font-medium text-gray-600 block">Дата инвентаризации:</label>
                            <input type="date" id="stockDate" onchange="loadStockMatrix()" class="border border-gray-300 rounded-md px-3 py-2 text-sm">
                        </div>
                        <div>
                            <label class="text-xs font-medium text-gray-600 block">Регион:</label>
                            <select id="stockRegionSelect" onchange="loadStockMatrix()" class="border border-gray-300 rounded-md px-3 py-2 text-sm">
                                <option value="Город" selected>Город</option>
                                <option value="Район">Район</option>
                            </select>
                        </div>
                        <div class="self-end flex flex-wrap gap-2">
                            <button onclick="saveStockMatrix()" class="bg-green-600 text-white px-4 py-2 rounded-md text-sm font-medium hover:bg-green-700 transition">💾 Сохранить акт</button>
                            <button onclick="downloadStockExcel()" class="bg-emerald-600 text-white px-4 py-2 rounded-md text-sm font-medium hover:bg-emerald-700 transition">📥 Скачать инвентаризацию</button>
                        </div>
                    </div>
                </header>
                <div class="bg-white shadow rounded-lg overflow-hidden mb-6">
                    <div id="stockMatrixContainer" class="overflow-x-auto" style="max-height: 700px;">
                        <p class="p-8 text-center text-gray-400">Загрузка матрицы инвентаризации...</p>
                    </div>
                </div>

                <div class="bg-white shadow rounded-lg p-6">
                    <h3 class="text-lg font-bold text-gray-800 mb-4">📅 История сохраненных актов инвентаризации (<span id="stockRegionLabel">Город</span>)</h3>
                    <div class="overflow-x-auto">
                        <table class="min-w-full divide-y divide-gray-200 text-sm">
                            <thead class="bg-gray-100">
                                <tr>
                                    <th class="px-4 py-2 text-left font-bold text-gray-700">Дата акта</th>
                                    <th class="px-4 py-2 text-center font-bold text-gray-700">Регион</th>
                                    <th class="px-4 py-2 text-center font-bold text-gray-700">Статус</th>
                                    <th class="px-4 py-2 text-right font-bold text-gray-700">Действие</th>
                                </tr>
                            </thead>
                            <tbody id="savedStockDaysTableBody" class="divide-y divide-gray-200">
                                <tr><td colspan="4" class="text-center py-4 text-gray-400">Загрузка сохраненных актов...</td></tr>
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>

            <!-- ЗАКРЫТИЕ МЕСЯЦА -->
            <div id="tabMonthClose" class="hidden space-y-6">
                <header class="bg-white shadow rounded-lg p-6 border-l-4 border-amber-500">
                    <h1 class="text-xl font-bold text-gray-800">🔒 Закрытие месяца и фиксация расходов</h1>
                    <p class="text-sm text-gray-500 mt-1">
                        После закрытия месяца количество детей, сотрудников и расходы замораживаются навсегда. Редактирование за закрытый период запрещается, скачивание отчётов сохраняется, а обновление норм в Excel/GitHub влияет только на открытые и будущие дни.
                    </p>
                </header>

                <div class="bg-white shadow rounded-lg p-6">
                    <div class="grid grid-cols-1 md:grid-cols-4 gap-4 items-end">
                        <div>
                            <label class="text-xs font-medium text-gray-600 block mb-1">Месяц закрытия:</label>
                            <input type="month" id="closeMonthPicker" onchange="checkCloseMonthStatus()" class="border border-gray-300 rounded-md px-3 py-2 text-sm w-full">
                        </div>
                        <div>
                            <label class="text-xs font-medium text-gray-600 block mb-1">Регион:</label>
                            <select id="closeRegionSelect" onchange="checkCloseMonthStatus()" class="border border-gray-300 rounded-md px-3 py-2 text-sm w-full">
                                <option value="Город" selected>Город</option>
                                <option value="Район">Район</option>
                                <option value="Оба">Оба региона</option>
                            </select>
                        </div>
                        <div>
                            <label class="text-xs font-medium text-gray-600 block mb-1">Пароль доступа:</label>
                            <input type="password" id="closeMonthPassword" placeholder="Введите пароль" class="border border-gray-300 rounded-md px-3 py-2 text-sm w-full">
                        </div>
                        <div>
                            <button onclick="submitCloseMonth()" id="btnCloseMonthAction" class="bg-amber-600 text-white px-4 py-2 rounded-md text-sm font-bold hover:bg-amber-700 transition w-full">
                                🔒 Закрыть месяц
                            </button>
                        </div>
                    </div>

                    <div id="closeMonthStatusBox" class="mt-4 p-4 rounded-md bg-gray-50 border text-sm text-gray-700">
                        Проверка статуса периода...
                    </div>
                </div>

                <div class="bg-white shadow rounded-lg p-6">
                    <h3 class="text-lg font-bold text-gray-800 mb-4">📜 Закрытые месяцы (выгрузка доступна)</h3>
                    <div class="overflow-x-auto">
                        <table class="min-w-full divide-y divide-gray-200 text-sm">
                            <thead class="bg-gray-100">
                                <tr>
                                    <th class="px-4 py-2 text-left font-bold text-gray-700">Месяц</th>
                                    <th class="px-4 py-2 text-center font-bold text-gray-700">Регион</th>
                                    <th class="px-4 py-2 text-center font-bold text-gray-700">Статус</th>
                                    <th class="px-4 py-2 text-center font-bold text-gray-700">Закрыт</th>
                                    <th class="px-4 py-2 text-right font-bold text-gray-700">Скачать отчёты</th>
                                </tr>
                            </thead>
                            <tbody id="closedMonthsTableBody" class="divide-y divide-gray-200">
                                <tr><td colspan="5" class="text-center py-4 text-gray-400">Загрузка списка...</td></tr>
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>
        </div>

        <div id="dayModal" class="hidden fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50">
            <div class="bg-white rounded-lg p-6 w-96 shadow-xl space-y-4">
                <h3 id="modalDateTitle" class="font-bold text-lg text-gray-800">Настройка дня</h3>
                <div>
                    <label class="block text-xs font-medium text-gray-700 mb-1">Статус дня:</label>
                    <select id="modalIsHoliday" class="w-full border rounded px-3 py-2 text-sm">
                        <option value="false">Рабочий день</option>
                        <option value="true">День отдыха / Выходной (Красный)</option>
                    </select>
                </div>
                <div>
                    <label class="block text-xs font-medium text-gray-700 mb-1">Принудительное меню:</label>
                    <select id="modalMenu" class="w-full border rounded px-3 py-2 text-sm">
                        <option value="">Автоматически по цепочке</option>
                        <option value="1">Меню № 1</option><option value="2">Меню № 2</option><option value="3">Меню № 3</option><option value="4">Меню № 4</option><option value="5">Меню № 5</option><option value="6">Меню № 6</option><option value="7">Меню № 7</option><option value="8">Меню № 8</option><option value="9">Меню № 9</option><option value="10">Меню № 10</option>
                    </select>
                </div>
                <div class="flex justify-end space-x-2 pt-2">
                    <button onclick="closeModal()" class="bg-gray-200 px-4 py-2 rounded text-sm font-medium">Отмена</button>
                    <button onclick="saveDaySettings()" class="bg-blue-600 text-white px-4 py-2 rounded text-sm font-medium hover:bg-blue-700">Сохранить</button>
                </div>
            </div>
        </div>

        <script>
            const SHIPPING_SERVICE_URL = "https://muxsin-fayz.onrender.com";
            let currentMatrixData = null;

            document.getElementById('targetDate').valueAsDate = new Date();
            document.getElementById('stockDate').valueAsDate = new Date();
            document.getElementById('factDate').valueAsDate = new Date();
            document.getElementById('staffFactDate').valueAsDate = new Date();
            document.getElementById('balanceDate').valueAsDate = new Date();

            const now = new Date();
            const yearStr = now.getFullYear();
            const monthStr = String(now.getMonth() + 1).padStart(2, '0');
            document.getElementById('childrenMonthPicker').value = `${yearStr}-${monthStr}`;
            document.getElementById('staffMonthPicker').value = `${yearStr}-${monthStr}`;
            document.getElementById('closeMonthPicker').value = `${yearStr}-${monthStr}`;

            let viewYear = now.getFullYear();
            let viewMonth = now.getMonth();
            let selectedDateStr = '';
            let currentMatrixFilter = 'ALL';

            function switchTab(tab) {
                ['matrix', 'fact', 'staffFact', 'balance', 'calendar', 'stock', 'monthClose'].forEach(t => {
                    const el = document.getElementById('tab' + capitalize(t));
                    const btn = document.getElementById('btnTab' + capitalize(t));
                    if (el) el.classList.add('hidden');
                    if (btn) {
                        btn.className = (t === 'monthClose')
                            ? "px-3 py-2 rounded-md text-sm font-medium bg-amber-100 text-amber-900 hover:bg-amber-200 transition"
                            : "px-3 py-2 rounded-md text-sm font-medium bg-gray-200 text-gray-700 hover:bg-gray-300 transition";
                    }
                });
                const activeTab = document.getElementById('tab' + capitalize(tab));
                const activeBtn = document.getElementById('btnTab' + capitalize(tab));
                if (activeTab) activeTab.classList.remove('hidden');
                if (activeBtn) activeBtn.className = "px-3 py-2 rounded-md text-sm font-medium bg-blue-600 text-white transition";

                if (tab === 'calendar') loadCalendarGrid();
                if (tab === 'stock') loadStockMatrix();
                if (tab === 'fact') loadFactData();
                if (tab === 'staffFact') loadStaffFactData();
                if (tab === 'balance') loadBalanceData();
                if (tab === 'monthClose') {
                    checkCloseMonthStatus();
                    loadClosedMonthsList();
                }
            }

            function capitalize(s) {
                if (s === 'staffFact') return 'StaffFact';
                if (s === 'monthClose') return 'MonthClose';
                return s.charAt(0).toUpperCase() + s.slice(1);
            }

            function setMatrixFilter(filterVal) {
                currentMatrixFilter = filterVal;
                ['ALL', 'Ежедневно', 'Неделья', 'Месяц'].forEach(f => {
                    const btn = document.getElementById('filterBtn' + f);
                    if (btn) {
                        btn.className = (f === filterVal)
                            ? "px-3 py-2 rounded-md text-xs font-bold bg-blue-600 text-white transition"
                            : "px-3 py-2 rounded-md text-xs font-bold bg-gray-200 text-gray-700 hover:bg-gray-300 transition";
                    }
                });
                const daysInput = document.getElementById('daysCountInput');
                if (filterVal === 'Ежедневно') daysInput.value = 1;
                else if (filterVal === 'Неделья') daysInput.value = 5;
                else if (filterVal === 'Месяц') daysInput.value = 10;
            }

            function changeMonth(dir) {
                viewMonth += dir;
                if (viewMonth > 11) { viewMonth = 0; viewYear++; }
                if (viewMonth < 0) { viewMonth = 11; viewYear--; }
                loadCalendarGrid();
            }

            async function loadCalendarGrid() {
                const monthNames = ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"];
                document.getElementById('currentMonthLabel').innerText = `${monthNames[viewMonth]} ${viewYear}`;
                const lastDay = new Date(viewYear, viewMonth + 1, 0);
                const startStr = `${viewYear}-${String(viewMonth + 1).padStart(2, '0')}-01`;
                const endStr = `${viewYear}-${String(viewMonth + 1).padStart(2, '0')}-${String(lastDay.getDate()).padStart(2, '0')}`;

                const response = await fetch(`/api/calendar/range?start=${startStr}&end=${endStr}`);
                const data = await response.json();
                const grid = document.getElementById('calendarGrid');
                grid.innerHTML = '';

                const firstDay = new Date(viewYear, viewMonth, 1);
                let startWeekDay = (firstDay.getDay() + 6) % 7;
                for (let i = 0; i < startWeekDay; i++) {
                    grid.innerHTML += `<div class="bg-gray-50 rounded-lg min-h-[105px] opacity-30"></div>`;
                }

                for (let d = 1; d <= lastDay.getDate(); d++) {
                    const dStr = `${viewYear}-${String(viewMonth + 1).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
                    const info = data[dStr] || { is_holiday: false, menu: 1, forced: false, is_closed: false };
                    let bgClass = "bg-white border-gray-200 hover:border-blue-400";
                    let badgeBg = "bg-blue-100 text-blue-800";
                    let statusText = `Меню № ${info.menu}`;

                    if (info.is_holiday) {
                        bgClass = "bg-red-50 border-red-300 hover:border-red-500";
                        badgeBg = "bg-red-200 text-red-900";
                        statusText = "День отдыха";
                    } else if (info.forced) {
                        badgeBg = "bg-purple-100 text-purple-800";
                        statusText += " (Руч.)";
                    }
                    if (info.is_closed) {
                        statusText += " 🔒";
                    }

                    grid.innerHTML += `
                        <div onclick="openDayModal('${dStr}', ${info.is_holiday}, ${info.menu !== null ? info.menu : 'null'})" class="border ${bgClass} rounded-lg p-3 min-h-[105px] flex flex-col justify-between cursor-pointer shadow-sm transition">
                            <div class="flex justify-between items-center"><span class="font-bold text-sm text-gray-700">${d}</span></div>
                            <div class="text-center my-auto"><span class="inline-block px-2.5 py-1 rounded-md text-xs font-bold ${badgeBg}">${statusText}</span></div>
                        </div>
                    `;
                }
            }

            function openDayModal(dateStr, isHoliday, menuNum) {
                selectedDateStr = dateStr;
                document.getElementById('modalDateTitle').innerText = `Настройка дня: ${dateStr}`;
                document.getElementById('modalIsHoliday').value = isHoliday ? "true" : "false";
                document.getElementById('modalMenu').value = (menuNum !== 'null' && menuNum !== null) ? menuNum : "";
                document.getElementById('dayModal').classList.remove('hidden');
            }

            function closeModal() { document.getElementById('dayModal').classList.add('hidden'); }

            async function saveDaySettings() {
                const is_holiday = document.getElementById('modalIsHoliday').value === 'true';
                const menuVal = document.getElementById('modalMenu').value;
                const menu = menuVal ? parseInt(menuVal) : null;
                const res = await fetch('/api/calendar/settings', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ date: selectedDateStr, is_holiday, menu })
                });
                const resData = await res.json();
                if (!res.ok) {
                    alert(resData.detail || 'Ошибка сохранения календаря');
                    return;
                }
                closeModal();
                loadCalendarGrid();
            }

            async function loadFactData() {
                const dateVal = document.getElementById('factDate').value;
                const region = document.getElementById('factRegionSelect').value;
                const container = document.getElementById('factContainer');
                container.innerHTML = '<div class="text-center py-12 text-gray-500 bg-white rounded-lg shadow">Загрузка данных факта...</div>';

                try {
                    const res = await fetch(`/api/fact/report?report_date=${dateVal}&region=${region}`);
                    const data = await res.json();
                    const isClosed = !!data.is_closed;

                    let html = '';
                    if (isClosed) {
                        html += `<div class="bg-amber-50 border border-amber-300 text-amber-900 px-4 py-3 rounded-lg text-sm font-bold">🔒 Месяц закрыт и заморожен. Редактирование и удаление недоступны, данные не пересчитываются по новым нормам.</div>`;
                    }

                    html += `
                        <div class="bg-white shadow rounded-lg p-6 border-l-4 border-green-500 mb-6">
                            <div class="flex justify-between items-center mb-4">
                                <h3 class="font-bold text-gray-800 text-lg">👶 Фактическое присутствие детей на ${data.date}</h3>
                                <span class="text-xs bg-purple-100 text-purple-800 font-bold px-3 py-1 rounded-full">${data.menu_num ? 'Меню № ' + data.menu_num : 'Выходной / Нет меню'}</span>
                            </div>
                            <div class="grid grid-cols-2 sm:grid-cols-4 md:grid-cols-6 lg:grid-cols-8 gap-3">
                    `;

                    data.points.forEach(p => {
                        const childVal = data.fact_children[p.id] !== undefined ? data.fact_children[p.id] : p.default_children;
                        html += `
                            <div class="bg-gray-50 p-2.5 border rounded-md text-center">
                                <div class="text-xs font-bold text-gray-700 truncate" title="${p.name}">${p.name}</div>
                                <div class="text-[10px] text-gray-400 mb-1">(План: ${p.plan_children !== undefined ? p.plan_children : p.default_children})</div>
                                <input type="number" ${isClosed ? 'disabled' : ''} data-point-fact="${p.id}" value="${childVal}" class="fact-child-cell border border-gray-300 rounded px-2 py-1 text-center text-sm w-full font-bold text-green-700 focus:ring-2 focus:ring-green-500 disabled:bg-gray-200 disabled:text-gray-500">
                            </div>
                        `;
                    });

                    html += `</div></div>`;

                    if (!data.menu_num) {
                        html += `<div class="bg-white p-8 text-center text-red-500 font-bold rounded-lg shadow">В этот день по календарю выходной — списание не производится.</div>`;
                    } else {
                        html += `
                            <div class="bg-white shadow rounded-lg overflow-hidden mb-6">
                                <div class="bg-gray-50 px-6 py-4 border-b border-gray-200">
                                    <h3 class="text-lg font-bold text-gray-800">📦 Фактический расход продуктов за ${data.date}</h3>
                                </div>
                                <div class="overflow-x-auto" style="max-height: 500px;">
                                    <table class="min-w-full divide-y divide-gray-200 text-xs border-collapse">
                                        <thead class="bg-gray-100 sticky top-0 z-10">
                                            <tr>
                                                <th class="px-4 py-3 text-left font-bold text-gray-700 border-r bg-gray-100 sticky left-0 z-20 min-w-[180px]">Продукт</th>
                                                <th class="px-3 py-3 text-center font-bold text-gray-700 border-r bg-gray-100 min-w-[70px]">Ед.</th>
                        `;

                        data.points.forEach(p => {
                            const childVal = data.fact_children[p.id] !== undefined ? data.fact_children[p.id] : p.default_children;
                            html += `<th class="px-3 py-3 text-center font-medium text-gray-700 border-r min-w-[90px]">${p.name}<br><span class="text-[10px] text-green-600">(${childVal} факт)</span></th>`;
                        });

                        html += `<th class="px-4 py-3 text-center font-bold text-green-800 bg-green-50">ИТОГО ФАКТ</th></tr></thead><tbody class="divide-y divide-gray-200 bg-white">`;

                        for (const [prodName, rowData] of Object.entries(data.fact_consumption)) {
                            html += `
                                <tr class="hover:bg-gray-50">
                                    <td class="px-4 py-3 font-medium text-gray-900 border-r bg-white sticky left-0 z-10">${prodName}</td>
                                    <td class="px-3 py-3 text-center text-gray-500 border-r bg-gray-50">${rowData.unit}</td>
                            `;
                            data.points.forEach(p => {
                                const val = rowData.points_qty[p.id] || 0;
                                html += `<td class="px-3 py-3 text-center border-r text-gray-700 ${val > 0 ? 'font-semibold text-green-700' : 'text-gray-300'}">${val > 0 ? val : '-'}</td>`;
                            });
                            html += `<td class="px-4 py-3 text-center font-bold text-green-900 bg-green-50">${rowData.total_fact}</td></tr>`;
                        }

                        html += `</tbody></table></div></div>`;
                    }

                    html += `
                        <div class="bg-white shadow rounded-lg p-6">
                            <h3 class="text-lg font-bold text-gray-800 mb-4">📅 История сохраненных дней по фактическому вводу (${region})</h3>
                            <div class="overflow-x-auto">
                                <table class="min-w-full divide-y divide-gray-200 text-sm">
                                    <thead class="bg-gray-100">
                                        <tr>
                                            <th class="px-4 py-2 text-left font-bold text-gray-700">Дата</th>
                                            <th class="px-4 py-2 text-center font-bold text-gray-700">Регион</th>
                                            <th class="px-4 py-2 text-center font-bold text-gray-700">Статус</th>
                                            <th class="px-4 py-2 text-right font-bold text-gray-700">Действие</th>
                                        </tr>
                                    </thead>
                                    <tbody id="savedDaysTableBody" class="divide-y divide-gray-200">
                                        <tr><td colspan="4" class="text-center py-4 text-gray-400">Загрузка сохраненных дней...</td></tr>
                                    </tbody>
                                </table>
                            </div>
                        </div>
                    `;

                    container.innerHTML = html;
                    loadSavedDaysList(region);
                } catch (err) {
                    container.innerHTML = '<div class="text-center py-12 text-red-500 bg-white rounded-lg shadow">Ошибка загрузки данных факта.</div>';
                }
            }

            async function loadSavedDaysList(region) {
                const res = await fetch(`/api/fact/saved-days?region=${region}`);
                const dates = await res.json();
                const tbody = document.getElementById('savedDaysTableBody');
                if (!tbody) return;

                if (dates.length === 0) {
                    tbody.innerHTML = `<tr><td colspan="4" class="text-center py-4 text-gray-400">Нет сохраненных дней для региона ${region}.</td></tr>`;
                    return;
                }

                let rowsHtml = '';
                dates.forEach(item => {
                    const dStr = typeof item === 'string' ? item : item.date;
                    const closed = typeof item === 'object' && item.is_closed;
                    rowsHtml += `
                        <tr class="hover:bg-gray-50">
                            <td class="px-4 py-3 font-semibold text-gray-800">${dStr}</td>
                            <td class="px-4 py-3 text-center text-gray-600">${region}</td>
                            <td class="px-4 py-3 text-center">${closed ? '<span class="bg-amber-100 text-amber-900 px-2.5 py-1 rounded-full text-xs font-bold">🔒 Закрыт</span>' : '<span class="bg-green-100 text-green-800 px-2.5 py-1 rounded-full text-xs font-bold">Сохранено</span>'}</td>
                            <td class="px-4 py-3 text-right space-x-2">
                                <button onclick="selectSavedDate('${dStr}')" class="bg-blue-600 text-white px-3 py-1.5 rounded text-xs font-medium hover:bg-blue-700 transition">👁 Открыть</button>
                                ${closed ? '' : `<button onclick="deleteFactDate('${dStr}', '${region}')" class="bg-red-600 text-white px-3 py-1.5 rounded text-xs font-medium hover:bg-red-700 transition">🗑 Удалить</button>`}
                            </td>
                        </tr>
                    `;
                });
                tbody.innerHTML = rowsHtml;
            }

            function selectSavedDate(dStr) {
                document.getElementById('factDate').value = dStr;
                loadFactData();
                window.scrollTo({ top: 0, behavior: 'smooth' });
            }

            async function deleteFactDate(dStr, region) {
                if (!confirm(`Точно удалить сохраненный факт за ${dStr} (${region})?`)) return;
                const res = await fetch(`/api/fact/children?report_date=${dStr}&region=${region}`, { method: 'DELETE' });
                const resData = await res.json();
                if (res.ok) {
                    loadFactData();
                } else {
                    alert(resData.detail || 'Ошибка при удалении.');
                }
            }

            async function loadStaffFactData() {
                const dateVal = document.getElementById('staffFactDate').value;
                const region = document.getElementById('staffFactRegionSelect').value;
                const container = document.getElementById('staffFactContainer');
                container.innerHTML = '<div class="text-center py-12 text-gray-500 bg-white rounded-lg shadow">Загрузка данных сотрудников...</div>';

                try {
                    const res = await fetch(`/api/fact-staff/report?report_date=${dateVal}&region=${region}`);
                    const data = await res.json();
                    const isClosed = !!data.is_closed;

                    let html = '';
                    if (isClosed) {
                        html += `<div class="bg-amber-50 border border-amber-300 text-amber-900 px-4 py-3 rounded-lg text-sm font-bold">🔒 Месяц закрыт и заморожен. Редактирование и удаление недоступны, данные не пересчитываются по новым нормам.</div>`;
                    }

                    html += `
                        <div class="bg-white shadow rounded-lg p-6 border-l-4 border-blue-500 mb-6">
                            <div class="flex justify-between items-center mb-4">
                                <h3 class="font-bold text-gray-800 text-lg">👥 Фактическое присутствие сотрудников на ${data.date}</h3>
                                <span class="text-xs bg-purple-100 text-purple-800 font-bold px-3 py-1 rounded-full">${data.menu_num ? 'Меню № ' + data.menu_num : 'Выходной / Нет меню'}</span>
                            </div>
                            <div class="grid grid-cols-2 sm:grid-cols-4 md:grid-cols-6 lg:grid-cols-8 gap-3">
                    `;

                    data.points.forEach(p => {
                        const staffVal = data.fact_staff[p.id] !== undefined ? data.fact_staff[p.id] : p.default_staff;
                        html += `
                            <div class="bg-gray-50 p-2.5 border rounded-md text-center">
                                <div class="text-xs font-bold text-gray-700 truncate mb-1" title="${p.name}">${p.name}</div>
                                <input type="number" ${isClosed ? 'disabled' : ''} data-point-staff-fact="${p.id}" value="${staffVal}" class="fact-staff-cell border border-gray-300 rounded px-2 py-1 text-center text-sm w-full font-bold text-blue-700 focus:ring-2 focus:ring-blue-500 disabled:bg-gray-200 disabled:text-gray-500">
                            </div>
                        `;
                    });

                    html += `</div></div>`;

                    if (!data.menu_num) {
                        html += `<div class="bg-white p-8 text-center text-red-500 font-bold rounded-lg shadow">В этот день по календарю выходной — списание не производится.</div>`;
                    } else {
                        html += `
                            <div class="bg-white shadow rounded-lg overflow-hidden mb-6">
                                <div class="bg-gray-50 px-6 py-4 border-b border-gray-200">
                                    <h3 class="text-lg font-bold text-gray-800">📦 Фактический расход продуктов (Сотрудники) за ${data.date}</h3>
                                </div>
                                <div class="overflow-x-auto" style="max-height: 500px;">
                                    <table class="min-w-full divide-y divide-gray-200 text-xs border-collapse">
                                        <thead class="bg-gray-100 sticky top-0 z-10">
                                            <tr>
                                                <th class="px-4 py-3 text-left font-bold text-gray-700 border-r bg-gray-100 sticky left-0 z-20 min-w-[180px]">Продукт</th>
                                                <th class="px-3 py-3 text-center font-bold text-gray-700 border-r bg-gray-100 min-w-[70px]">Ед.</th>
                        `;

                        data.points.forEach(p => {
                            const staffVal = data.fact_staff[p.id] !== undefined ? data.fact_staff[p.id] : p.default_staff;
                            html += `<th class="px-3 py-3 text-center font-medium text-gray-700 border-r min-w-[90px]">${p.name}<br><span class="text-[10px] text-blue-600">(${staffVal} сотр)</span></th>`;
                        });

                        html += `<th class="px-4 py-3 text-center font-bold text-blue-800 bg-blue-50">ИТОГО ФАКТ</th></tr></thead><tbody class="divide-y divide-gray-200 bg-white">`;

                        for (const [prodName, rowData] of Object.entries(data.fact_consumption)) {
                            html += `
                                <tr class="hover:bg-gray-50">
                                    <td class="px-4 py-3 font-medium text-gray-900 border-r bg-white sticky left-0 z-10">${prodName}</td>
                                    <td class="px-3 py-3 text-center text-gray-500 border-r bg-gray-50">${rowData.unit}</td>
                            `;
                            data.points.forEach(p => {
                                const val = rowData.points_qty[p.id] || 0;
                                html += `<td class="px-3 py-3 text-center border-r text-gray-700 ${val > 0 ? 'font-semibold text-blue-700' : 'text-gray-300'}">${val > 0 ? val : '-'}</td>`;
                            });
                            html += `<td class="px-4 py-3 text-center font-bold text-blue-900 bg-blue-50">${rowData.total_fact}</td></tr>`;
                        }

                        html += `</tbody></table></div></div>`;
                    }

                    html += `
                        <div class="bg-white shadow rounded-lg p-6">
                            <h3 class="text-lg font-bold text-gray-800 mb-4">📅 История сохраненных дней по сотрудникам (${region})</h3>
                            <div class="overflow-x-auto">
                                <table class="min-w-full divide-y divide-gray-200 text-sm">
                                    <thead class="bg-gray-100">
                                        <tr>
                                            <th class="px-4 py-2 text-left font-bold text-gray-700">Дата</th>
                                            <th class="px-4 py-2 text-center font-bold text-gray-700">Регион</th>
                                            <th class="px-4 py-2 text-center font-bold text-gray-700">Статус</th>
                                            <th class="px-4 py-2 text-right font-bold text-gray-700">Действие</th>
                                        </tr>
                                    </thead>
                                    <tbody id="savedStaffDaysTableBody" class="divide-y divide-gray-200"></tbody>
                                </table>
                            </div>
                        </div>
                    `;

                    container.innerHTML = html;
                    loadSavedStaffDaysList(region);
                } catch (err) {
                    container.innerHTML = '<div class="text-center py-12 text-red-500 bg-white rounded-lg shadow">Ошибка загрузки данных сотрудников.</div>';
                }
            }

            async function saveStaffFact() {
                const dateVal = document.getElementById('staffFactDate').value;
                const region = document.getElementById('staffFactRegionSelect').value;
                const cells = document.querySelectorAll('.fact-staff-cell');
                const payload = {};
                cells.forEach(cell => {
                    const pid = cell.getAttribute('data-point-staff-fact');
                    payload[pid] = parseInt(cell.value) || 0;
                });

                const res = await fetch(`/api/fact-staff/staff?report_date=${dateVal}&region=${region}`, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(payload)
                });
                const resData = await res.json();
                if (res.ok) {
                    alert('Факт сотрудников сохранен и учтен в остатках!');
                    loadStaffFactData();
                } else {
                    alert(resData.detail || 'Ошибка сохранения.');
                }
            }

            async function loadSavedStaffDaysList(region) {
                const res = await fetch(`/api/fact-staff/saved-days?region=${region}`);
                const dates = await res.json();
                const tbody = document.getElementById('savedStaffDaysTableBody');
                if (!tbody) return;
                if (dates.length === 0) {
                    tbody.innerHTML = `<tr><td colspan="4" class="text-center py-4 text-gray-400">Нет сохраненных дней для региона ${region}.</td></tr>`;
                    return;
                }
                tbody.innerHTML = dates.map(item => {
                    const dStr = typeof item === 'string' ? item : item.date;
                    const closed = typeof item === 'object' && item.is_closed;
                    return `
                    <tr class="hover:bg-gray-50">
                        <td class="px-4 py-3 font-semibold text-gray-800">${dStr}</td>
                        <td class="px-4 py-3 text-center text-gray-600">${region}</td>
                        <td class="px-4 py-3 text-center">${closed ? '<span class="bg-amber-100 text-amber-900 px-2.5 py-1 rounded-full text-xs font-bold">🔒 Закрыт</span>' : '<span class="bg-green-100 text-green-800 px-2.5 py-1 rounded-full text-xs font-bold">Сохранено</span>'}</td>
                        <td class="px-4 py-3 text-right space-x-2">
                            <button onclick="document.getElementById('staffFactDate').value='${dStr}'; loadStaffFactData(); window.scrollTo({top:0, behavior:'smooth'});" class="bg-blue-600 text-white px-3 py-1.5 rounded text-xs font-medium hover:bg-blue-700 transition">👁 Открыть</button>
                            ${closed ? '' : `<button onclick="deleteStaffFactDate('${dStr}', '${region}')" class="bg-red-600 text-white px-3 py-1.5 rounded text-xs font-medium hover:bg-red-700 transition">🗑 Удалить</button>`}
                        </td>
                    </tr>
                `}).join('');
            }

            async function deleteStaffFactDate(dStr, region) {
                if (!confirm(`Удалить сохраненный факт сотрудников за ${dStr} (${region})?`)) return;
                const res = await fetch(`/api/fact-staff/staff?report_date=${dStr}&region=${region}`, { method: 'DELETE' });
                const resData = await res.json();
                if (res.ok) loadStaffFactData();
                else alert(resData.detail || 'Ошибка удаления');
            }

            function downloadStaffFactAllRegionsExcel() {
                const dateVal = document.getElementById('staffFactDate').value;
                window.location.href = `/api/fact-staff/export/all-regions-excel?report_date=${dateVal}`;
            }

            function downloadStaffMatrixExcel() {
                const monthVal = document.getElementById('staffMonthPicker').value;
                const region = document.getElementById('staffFactRegionSelect').value;
                if (!monthVal) { alert('Выберите месяц'); return; }
                window.location.href = `/api/fact-staff/export/staff-matrix?month=${monthVal}&region=${region}`;
            }

            function downloadStaffMonthlyMttExcel() {
                const monthVal = document.getElementById('staffMonthPicker').value;
                const region = document.getElementById('staffFactRegionSelect').value;
                if (!monthVal) { alert('Выберите месяц'); return; }
                window.location.href = `/api/fact-staff/export/monthly-mtt-excel?month=${monthVal}&region=${region}`;
            }

            async function sendStaffMenuTelegramManual() {
                const dateVal = document.getElementById('staffFactDate').value;
                const region = document.getElementById('staffFactRegionSelect').value;
                if(!confirm(`Отправить меню сотрудников за ${dateVal} (${region}) в Telegram?`)) return;
                const res = await fetch(`/api/menu-staff/generate-and-send?report_date=${dateVal}&region=${region}`, { method: 'POST' });
                const data = await res.json();
                alert(res.ok ? data.message : ('Ошибка: ' + (data.detail || 'Не удалось запустить отправку')));
            }

            async function sendMenuTelegramManual() {
                const dateVal = document.getElementById('factDate').value;
                const region = document.getElementById('factRegionSelect').value;
                if(!confirm(`Запустить отправку официального меню с печатью за ${dateVal} (${region}) в Telegram в фоновом режиме?`)) return;
                const res = await fetch(`/api/menu/generate-and-send?report_date=${dateVal}&region=${region}`, { method: 'POST' });
                const data = await res.json();
                if(res.ok) alert('Процесс отправки запущен! Файлы будут поступать в Telegram по очереди.');
                else alert('Ошибка: ' + (data.detail || 'Не удалось запустить отправку'));
            }

            function downloadFactAllRegionsExcel() {
                const dateVal = document.getElementById('factDate').value;
                window.location.href = `/api/fact/export/all-regions-excel?report_date=${dateVal}`;
            }

            function downloadChildrenMatrixExcel() {
                const monthVal = document.getElementById('childrenMonthPicker').value;
                const region = document.getElementById('factRegionSelect').value;
                if (!monthVal) { alert('Пожалуйста, выберите месяц'); return; }
                window.location.href = `/api/fact/export/children-matrix?month=${monthVal}&region=${region}`;
            }

            function downloadMonthlyMttExcel() {
                const monthVal = document.getElementById('childrenMonthPicker').value;
                const region = document.getElementById('factRegionSelect').value;
                if (!monthVal) { alert('Пожалуйста, выберите месяц'); return; }
                window.location.href = `/api/fact/export/monthly-mtt-excel?month=${monthVal}&region=${region}`;
            }

            async function loadBalanceData() {
                const dateVal = document.getElementById('balanceDate').value;
                const region = document.getElementById('balanceRegionSelect').value;
                const container = document.getElementById('balanceContainer');
                container.innerHTML = '<div class="text-center py-12 text-gray-500 bg-white rounded-lg shadow">Расчет актуальных остатков по МТТ...</div>';

                try {
                    const res = await fetch(`/api/balance/report?report_date=${dateVal}&region=${region}`);
                    const data = await res.json();

                    let html = `
                        <div class="bg-white shadow rounded-lg overflow-hidden">
                            <div class="bg-gray-50 px-6 py-4 border-b border-gray-200 flex justify-between items-center">
                                <h3 class="text-lg font-bold text-gray-800">📦 Остатки на ${data.date} (Регион: ${data.region})</h3>
                                <span class="text-xs text-gray-500 bg-gray-200 px-3 py-1 rounded">Инв. база: ${data.inv_date ? data.inv_date : 'Нет актов'}</span>
                            </div>
                            <div class="overflow-x-auto" style="max-height: 700px;">
                                <table class="min-w-full divide-y divide-gray-200 text-xs border-collapse">
                                    <thead class="bg-gray-100 sticky top-0 z-10">
                                        <tr>
                                            <th class="px-4 py-3 text-left font-bold text-gray-700 border-r bg-gray-100 sticky left-0 z-20 min-w-[180px]">Продукт</th>
                                            <th class="px-3 py-3 text-center font-bold text-gray-700 border-r bg-gray-100 min-w-[70px]">Ед.</th>
                    `;

                    data.points.forEach(p => {
                        const countDisplay = p.is_saved ? `${p.children_count} дет` : `<span class="text-red-600 font-extrabold text-sm" title="Факт не сохранен (План)">П</span>`;
                        html += `<th class="px-3 py-3 text-center font-medium text-gray-700 border-r min-w-[90px]">${p.name}<br><span class="text-[10px] text-indigo-600">(${countDisplay})</span></th>`;
                    });

                    html += `<th class="px-4 py-3 text-center font-bold text-indigo-800 bg-indigo-50">ИТОГО ОСТАТОК</th></tr></thead><tbody class="divide-y divide-gray-200 bg-white">`;

                    for (const [prodName, rowData] of Object.entries(data.balance_report)) {
                        html += `
                            <tr class="hover:bg-gray-50">
                                <td class="px-4 py-3 font-medium text-gray-900 border-r bg-white sticky left-0 z-10">${prodName}</td>
                                <td class="px-3 py-3 text-center text-gray-500 border-r bg-gray-50">${rowData.unit}</td>
                        `;
                        data.points.forEach(p => {
                            const val = rowData.points_balance[p.id] || 0;
                            const colorClass = val < 0 ? 'text-red-600 font-extrabold bg-red-50' : (val > 0 ? 'text-indigo-600 font-bold' : 'text-gray-400');
                            html += `<td class="px-3 py-3 text-center border-r ${colorClass}">${val}</td>`;
                        });
                        const totalColorClass = rowData.total_balance < 0 ? 'text-red-700 bg-red-100 font-extrabold' : 'text-indigo-900 font-bold';
                        html += `<td class="px-4 py-3 text-center ${totalColorClass} bg-indigo-50">${rowData.total_balance}</td></tr>`;
                    }

                    html += `</tbody></table></div></div>`;
                    container.innerHTML = html;
                } catch (err) {
                    container.innerHTML = '<div class="text-center py-12 text-red-500 bg-white rounded-lg shadow">Ошибка загрузки остатков.</div>';
                }
            }

            async function saveFactChildren() {
                const dateVal = document.getElementById('factDate').value;
                const region = document.getElementById('factRegionSelect').value;
                const cells = document.querySelectorAll('.fact-child-cell');
                const payload = {};
                cells.forEach(cell => {
                    const pid = cell.getAttribute('data-point-fact');
                    payload[pid] = parseInt(cell.value) || 0;
                });

                const res = await fetch(`/api/fact/children?report_date=${dateVal}&region=${region}`, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(payload)
                });
                const resData = await res.json();
                if (res.ok) {
                    alert('Фактическое присутствие детей успешно сохранено!');
                    loadFactData();
                } else {
                    alert(resData.detail || 'Ошибка сохранения.');
                }
            }

            async function loadStockMatrix() {
                const dateVal = document.getElementById('stockDate').value;
                const region = document.getElementById('stockRegionSelect').value;
                document.getElementById('stockRegionLabel').innerText = region;
                const container = document.getElementById('stockMatrixContainer');
                container.innerHTML = '<div class="text-center py-12 text-gray-500">Загрузка акта инвентаризации...</div>';

                const res = await fetch(`/api/stock/matrix/${region}?inv_date=${dateVal}`);
                const data = await res.json();
                let html = `
                    <table class="min-w-full divide-y divide-gray-200 text-xs border-collapse">
                        <thead class="bg-gray-100 sticky top-0 z-10">
                            <tr>
                                <th class="px-4 py-3 text-left font-bold text-gray-700 border-r bg-gray-100 sticky left-0 z-20 min-w-[180px]">Продукт</th>
                                <th class="px-3 py-3 text-center font-bold text-gray-700 border-r bg-gray-100 min-w-[70px]">Ед.</th>
                `;
                data.points.forEach(p => {
                    html += `<th class="px-3 py-3 text-center font-medium text-gray-700 border-r min-w-[90px]">${p.name}<br><span class="text-[10px] text-blue-600">(${p.children_count} дет)</span></th>`;
                });
                html += `</tr></thead><tbody class="divide-y divide-gray-200 bg-white">`;

                for (const [prodName, pInfo] of Object.entries(data.products)) {
                    html += `
                        <tr class="hover:bg-gray-50">
                            <td class="px-4 py-2 font-medium text-gray-900 border-r bg-white sticky left-0 z-10">${prodName}</td>
                            <td class="px-3 py-2 text-center text-gray-500 border-r bg-gray-50">${pInfo.unit}</td>
                    `;
                    data.points.forEach(p => {
                        const val = pInfo.points_stock[p.id] || 0;
                        html += `
                            <td class="px-2 py-1.5 text-center border-r">
                                <input type="number" step="0.01" data-point="${p.id}" data-product="${prodName}" value="${val}" class="stock-cell border border-gray-200 rounded px-1.5 py-1 text-center text-xs w-20 focus:ring-2 focus:ring-blue-500">
                            </td>
                        `;
                    });
                    html += `</tr>`;
                }
                html += `</tbody></table>`;
                container.innerHTML = html;
                loadSavedStockDaysList(region);
            }

            async function loadSavedStockDaysList(region) {
                const res = await fetch(`/api/stock/saved-days?region=${region}`);
                const dates = await res.json();
                const tbody = document.getElementById('savedStockDaysTableBody');
                if (!tbody) return;

                if (dates.length === 0) {
                    tbody.innerHTML = `<tr><td colspan="4" class="text-center py-4 text-gray-400">Нет сохраненных актов для региона ${region}.</td></tr>`;
                    return;
                }

                let rowsHtml = '';
                dates.forEach(dStr => {
                    rowsHtml += `
                        <tr class="hover:bg-gray-50">
                            <td class="px-4 py-3 font-semibold text-gray-800">${dStr}</td>
                            <td class="px-4 py-3 text-center text-gray-600">${region}</td>
                            <td class="px-4 py-3 text-center"><span class="bg-green-100 text-green-800 px-2.5 py-1 rounded-full text-xs font-bold">Сохранено</span></td>
                            <td class="px-4 py-3 text-right space-x-2">
                                <button onclick="selectSavedStockDate('${dStr}')" class="bg-blue-600 text-white px-3 py-1.5 rounded text-xs font-medium hover:bg-blue-700 transition">✏️ Редактировать</button>
                                <button onclick="deleteStockDate('${dStr}', '${region}')" class="bg-red-600 text-white px-3 py-1.5 rounded text-xs font-medium hover:bg-red-700 transition">🗑 Удалить</button>
                            </td>
                        </tr>
                    `;
                });
                tbody.innerHTML = rowsHtml;
            }

            function selectSavedStockDate(dStr) {
                document.getElementById('stockDate').value = dStr;
                loadStockMatrix();
                window.scrollTo({ top: 0, behavior: 'smooth' });
            }

            async function deleteStockDate(dStr, region) {
                if (!confirm(`Точно удалить акт инвентаризации за ${dStr} (${region})?`)) return;
                const res = await fetch(`/api/stock/matrix/${region}?inv_date=${dStr}`, { method: 'DELETE' });
                if (res.ok) loadStockMatrix();
                else alert('Ошибка при удалении акта.');
            }

            async function saveStockMatrix() {
                const dateVal = document.getElementById('stockDate').value;
                const region = document.getElementById('stockRegionSelect').value;
                const cells = document.querySelectorAll('.stock-cell');
                const stockPayload = {};
                cells.forEach(cell => {
                    const pid = cell.getAttribute('data-point');
                    const prod = cell.getAttribute('data-product');
                    const val = parseFloat(cell.value) || 0.0;
                    if (!stockPayload[pid]) stockPayload[pid] = {};
                    stockPayload[pid][prod] = val;
                });

                const res = await fetch(`/api/stock/matrix/${region}?inv_date=${dateVal}`, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(stockPayload)
                });
                if (res.ok) {
                    alert('Акт инвентаризации успешно сохранен!');
                    loadStockMatrix();
                } else alert('Ошибка сохранения акта.');
            }

            function downloadStockExcel() {
                const dateVal = document.getElementById('stockDate').value;
                const region = document.getElementById('stockRegionSelect').value;
                if (!dateVal) { alert('Выберите дату инвентаризации'); return; }
                window.location.href = `/api/stock/export/excel?inv_date=${dateVal}&region=${region}`;
            }

            async function checkCloseMonthStatus() {
                const monthVal = document.getElementById('closeMonthPicker').value;
                const regionVal = document.getElementById('closeRegionSelect').value;
                const box = document.getElementById('closeMonthStatusBox');
                const btn = document.getElementById('btnCloseMonthAction');
                if (!monthVal) return;

                const regionsToCheck = regionVal === 'Оба' ? ['Город', 'Район'] : [regionVal];
                const statuses = [];
                for (const reg of regionsToCheck) {
                    const res = await fetch(`/api/month-close/status?month=${monthVal}&region=${encodeURIComponent(reg)}`);
                    statuses.push(await res.json());
                }

                const allClosed = statuses.every(s => s.is_closed);
                if (allClosed) {
                    box.className = "mt-4 p-4 rounded-md bg-amber-50 border border-amber-300 text-sm text-amber-900 font-medium";
                    box.innerHTML = statuses.map(s => `🔒 <b>${s.month} (${s.region})</b>: Закрыт навсегда (${s.closed_at || 'дата зафиксирована'}, дней в снимке: ${s.days_count})`).join('<br>');
                    btn.disabled = true;
                    btn.className = "bg-gray-400 text-white px-4 py-2 rounded-md text-sm font-bold cursor-not-allowed w-full";
                } else {
                    box.className = "mt-4 p-4 rounded-md bg-emerald-50 border border-emerald-300 text-sm text-emerald-900";
                    box.innerHTML = statuses.map(s => s.is_closed
                        ? `🔒 <b>${s.month} (${s.region})</b>: Уже закрыт (${s.closed_at})`
                        : `🟢 <b>${s.month} (${s.region})</b>: Открыт для редактирования. Пустые/нулевые дни по МТТ будут зафиксированы как <b>0</b>.`
                    ).join('<br>');
                    btn.disabled = false;
                    btn.className = "bg-amber-600 text-white px-4 py-2 rounded-md text-sm font-bold hover:bg-amber-700 transition w-full";
                }
            }

            async function submitCloseMonth() {
                const monthVal = document.getElementById('closeMonthPicker').value;
                const regionVal = document.getElementById('closeRegionSelect').value;
                const passwordVal = document.getElementById('closeMonthPassword').value;

                if (!monthVal) { alert('Выберите месяц'); return; }
                if (!passwordVal) { alert('Введите пароль для закрытия месяца'); return; }
                if (!confirm(`ВНИМАНИЕ! Закрыть месяц ${monthVal} (${regionVal}) навсегда?\nПосле закрытия редактирование и пересчёт за этот период будут невозможны.`)) return;

                const res = await fetch('/api/month-close/close', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ month: monthVal, region: regionVal, password: passwordVal })
                });
                const data = await res.json();
                if (res.ok) {
                    alert('✅ ' + data.message);
                    document.getElementById('closeMonthPassword').value = '';
                    checkCloseMonthStatus();
                    loadClosedMonthsList();
                } else {
                    alert('❌ ' + (data.detail || 'Ошибка закрытия месяца'));
                }
            }

            async function loadClosedMonthsList() {
                const res = await fetch('/api/month-close/list');
                const items = await res.json();
                const tbody = document.getElementById('closedMonthsTableBody');
                if (!tbody) return;

                if (!items.length) {
                    tbody.innerHTML = `<tr><td colspan="5" class="text-center py-4 text-gray-400">Пока нет закрытых месяцев.</td></tr>`;
                    return;
                }

                tbody.innerHTML = items.map(it => `
                    <tr class="hover:bg-gray-50">
                        <td class="px-4 py-3 font-bold text-gray-800">${it.month}</td>
                        <td class="px-4 py-3 text-center text-gray-700">${it.region}</td>
                        <td class="px-4 py-3 text-center"><span class="bg-amber-100 text-amber-900 px-2.5 py-1 rounded-full text-xs font-bold">🔒 Закрыт</span></td>
                        <td class="px-4 py-3 text-center text-xs text-gray-500">${it.closed_at || ''}</td>
                        <td class="px-4 py-3 text-right space-x-2">
                            <a href="/api/fact/export/children-matrix?month=${it.month}&region=${encodeURIComponent(it.region)}" class="inline-block bg-purple-600 text-white px-2.5 py-1 rounded text-xs font-medium hover:bg-purple-700">Дети</a>
                            <a href="/api/fact/export/monthly-mtt-excel?month=${it.month}&region=${encodeURIComponent(it.region)}" class="inline-block bg-blue-600 text-white px-2.5 py-1 rounded text-xs font-medium hover:bg-blue-700">Расход дети</a>
                            <a href="/api/fact-staff/export/staff-matrix?month=${it.month}&region=${encodeURIComponent(it.region)}" class="inline-block bg-indigo-600 text-white px-2.5 py-1 rounded text-xs font-medium hover:bg-indigo-700">Сотрудники</a>
                            <a href="/api/fact-staff/export/monthly-mtt-excel?month=${it.month}&region=${encodeURIComponent(it.region)}" class="inline-block bg-emerald-600 text-white px-2.5 py-1 rounded text-xs font-medium hover:bg-emerald-700">Расход сотр.</a>
                        </td>
                    </tr>
                `).join('');
            }

            async function loadMatrixRaznaryadka() {
                const dateStr = document.getElementById('targetDate').value;
                const region = document.getElementById('regionSelect').value;
                const daysCount = document.getElementById('daysCountInput').value || 2;
                const container = document.getElementById('matrixResultContainer');
                if (!dateStr) { alert('Выберите дату'); return; }
                container.innerHTML = '<div class="text-center py-12 text-gray-500">Формирование точной закупки (с учетом фильтра и дней)...</div>';

                try {
                    const response = await fetch(`/api/raznaryadka/matrix/${dateStr}?region=${region}&days_count=${daysCount}&filter_period=${currentMatrixFilter}`);
                    const data = await response.json();
                    currentMatrixData = data;

                    let html = `
                        <div class="grid grid-cols-1 md:grid-cols-3 gap-6 mb-6">
                            <div class="bg-white shadow rounded-lg p-6 border-t-4 border-blue-500">
                                <h3 class="text-sm font-medium text-gray-500">Период закупки</h3>
                                <p class="text-lg font-bold text-gray-800 mt-1">${data.week_start} - ${data.week_end}</p>
                            </div>
                            <div class="bg-white shadow rounded-lg p-6 border-t-4 border-green-500">
                                <h3 class="text-sm font-medium text-gray-500">Инвентаризация (база)</h3>
                                <p class="text-lg font-bold text-gray-800 mt-1">${data.inv_date ? data.inv_date : 'Нет актов'}</p>
                            </div>
                            <div class="bg-white shadow rounded-lg p-6 border-t-4 border-purple-500">
                                <h3 class="text-sm font-medium text-gray-500">Учтено рабочих дней</h3>
                                <p class="text-lg font-bold text-purple-600 mt-1">${data.working_days_count} дней</p>
                            </div>
                        </div>

                        <div class="bg-white shadow rounded-lg overflow-hidden">
                            <div class="bg-gray-50 px-6 py-4 border-b border-gray-200 flex flex-wrap justify-between items-center gap-3">
                                <h3 class="text-lg font-medium text-gray-900">📊 Чистая потребность к закупке <span class="text-xs text-blue-600 font-bold">(Фильтр: ${currentMatrixFilter})</span></h3>
                                <div class="flex items-center gap-2">
                                    <select id="createDraftMttSelect" class="border border-gray-300 rounded-md px-3 py-1.5 text-xs font-bold focus:ring-2 focus:ring-amber-500 bg-white"></select>
                                    <button onclick="createDraftSaleDocument()" class="bg-amber-600 text-white px-3.5 py-1.5 rounded-md text-xs font-bold hover:bg-amber-700 transition flex items-center gap-1 shadow">📄 Создать накладную</button>
                                </div>
                            </div>
                            <div class="overflow-x-auto" style="max-height: 700px;">
                                <table class="min-w-full divide-y divide-gray-200 text-xs border-collapse">
                                    <thead class="bg-gray-100 sticky top-0 z-10">
                                        <tr>
                                            <th class="px-4 py-3 text-left font-bold text-gray-700 border-r bg-gray-100 sticky left-0 z-20 min-w-[180px]">Продукт</th>
                                            <th class="px-3 py-3 text-center font-bold text-gray-700 border-r bg-gray-100 min-w-[70px]">Ед.</th>
                    `;
                    data.points.forEach(point => {
                        html += `<th class="px-3 py-3 text-center font-medium text-gray-700 border-r min-w-[90px]">${point.name}<br><span class="text-[10px] text-blue-600">(по факту)</span></th>`;
                    });
                    html += `<th class="px-4 py-3 text-center font-bold text-blue-800 bg-blue-50">ИТОГО К ЗАКУПКЕ</th></tr></thead><tbody class="divide-y divide-gray-200 bg-white">`;

                    for (const [prodName, rowData] of Object.entries(data.matrix)) {
                        html += `
                            <tr class="hover:bg-gray-50">
                                <td class="px-4 py-3 font-medium text-gray-900 border-r bg-white sticky left-0 z-10">${prodName}</td>
                                <td class="px-3 py-3 text-center text-gray-500 border-r bg-gray-50">${rowData.unit}</td>
                        `;
                        data.points.forEach(point => {
                            const val = rowData.points_qty[point.id] || 0;
                            html += `<td class="px-3 py-3 text-center border-r text-gray-700 ${val > 0 ? 'font-semibold text-blue-600' : 'text-gray-300'}">${val > 0 ? val : '-'}</td>`;
                        });
                        html += `<td class="px-4 py-3 text-center font-bold text-blue-900 bg-green-50">${rowData.weekly_total}</td></tr>`;
                    }
                    html += `</tbody></table></div></div>`;
                    container.innerHTML = html;

                    const mttSelect = document.getElementById('createDraftMttSelect');
                    if (mttSelect && data.points) {
                        mttSelect.innerHTML = data.points.map(p => `<option value="${p.id}" data-name="${p.name}">${p.name}</option>`).join('');
                    }
                } catch (err) {
                    container.innerHTML = '<div class="text-center py-12 text-red-500">Ошибка загрузки матрицы.</div>';
                }
            }

            async function createDraftSaleDocument() {
                if (!currentMatrixData) { alert('Сначала сформируйте разнарядку!'); return; }
                const selectEl = document.getElementById('createDraftMttSelect');
                if (!selectEl || !selectEl.value) { alert('Выберите МТТ для создания накладной'); return; }

                const pointId = parseInt(selectEl.value);
                const selectedOption = selectEl.options[selectEl.selectedIndex];
                const mttName = selectedOption.getAttribute('data-name') || selectedOption.text;
                const targetDate = document.getElementById('targetDate').value;

                const items = [];
                for (const [prodName, rowData] of Object.entries(currentMatrixData.matrix)) {
                    const qty = rowData.points_qty[pointId] || 0;
                    if (qty > 0) items.push({ product: prodName, qty: qty });
                }

                if (items.length === 0) {
                    alert(`Для ${mttName} нет продуктов к закупке (все значения равны 0).`);
                    return;
                }
                if (!confirm(`Создать черновик накладной на дату ${targetDate} для ${mttName} (${items.length} позиций)?`)) return;

                try {
                    const response = await fetch(`${SHIPPING_SERVICE_URL}/api/external/sales/create-draft`, {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({ date: targetDate, mtt_name: mttName, organization: 'MUXSIN-FAYZ', items: items })
                    });
                    const resData = await response.json();
                    if (response.ok) alert(`✅ ${resData.message}`);
                    else alert(`❌ Ошибка: ${resData.error || 'Не удалось создать накладную'}`);
                } catch (err) {
                    alert(`❌ Ошибка соединения со второй системой отгрузок: ${err.message}`);
                }
            }

            function downloadRaznaryadkaExcel() {
                const dateStr = document.getElementById('targetDate').value;
                const region = document.getElementById('regionSelect').value;
                const daysCount = document.getElementById('daysCountInput').value || 2;
                if (!dateStr) { alert('Выберите начальную дату'); return; }
                window.location.href = `/api/raznaryadka/export/excel?target_date=${dateStr}&region=${region}&days_count=${daysCount}&filter_period=${currentMatrixFilter}`;
            }

            window.onload = loadMatrixRaznaryadka;

            function downloadBalanceExcel() {
                const dateVal = document.getElementById('balanceDate').value;
                const region = document.getElementById('balanceRegionSelect').value;
                if (!dateVal) { alert('Выберите дату'); return; }
                window.location.href = `/api/balance/export/excel?report_date=${dateVal}&region=${region}`;
            }
        </script>
    </body>
    </html>
    """
    return HTMLResponse(content=html_content)


@app.get("/api/month-close/status")
def get_month_close_status_api(month: str, region: str = "Город"):
    snap = get_closed_month_snapshot(month, region)
    return {
        "month": month,
        "region": region,
        "is_closed": snap is not None,
        "closed_at": snap.get("closed_at") if snap else None,
        "days_count": len(snap.get("days", {})) if snap else 0
    }


@app.get("/api/month-close/list")
def list_closed_months_api():
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT month_str, region, closed_at FROM closed_months ORDER BY month_str DESC, region ASC;")
        rows = cursor.fetchall()
        cursor.close()
        conn.close()
        return [{"month": r["month_str"], "region": r["region"], "closed_at": r["closed_at"]} for r in rows]
    except Exception as e:
        print(f"Ошибка загрузки списка закрытых месяцев: {e}")
        return []


@app.post("/api/month-close/close")
def close_month_api(req: CloseMonthRequest):
    if req.password != MONTH_CLOSE_PASSWORD:
        raise HTTPException(status_code=403, detail="Неверный пароль закрытия месяца")

    if req.region == "Оба":
        results = [_close_single_month(req.month, "Город"), _close_single_month(req.month, "Район")]
        return {
            "status": "closed",
            "month": req.month,
            "regions": results,
            "message": f"Месяц {req.month} закрыт и заморожен по обоим регионам (Город и Район)."
        }

    res = _close_single_month(req.month, req.region)
    return {
        "status": "closed",
        "month": req.month,
        "region": req.region,
        "closed_at": res["closed_at"],
        "message": f"Месяц {req.month} ({req.region}) закрыт и заморожен навсегда."
    }


@app.get("/api/fact/report")
def get_fact_report(report_date: str, region: str = "Город"):
    try:
        parsed_date = date.fromisoformat(report_date)
    except Exception:
        parsed_date = date.today()

    points_dict = load_points_from_excel(region)
    closed_day = get_closed_day_snapshot(report_date, region)

    if closed_day is not None:
        snap_ch = closed_day.get("children_counts", {})
        points_list = []
        fact_children_out = {}
        for pid, pinfo in points_dict.items():
            c_cnt = _safe_int_or_zero(snap_ch.get(str(pid), snap_ch.get(pid, 0)))
            fact_children_out[str(pid)] = c_cnt
            points_list.append({
                "id": pid,
                "name": pinfo["name"],
                "default_children": c_cnt,
                "plan_children": pinfo["children_count"]
            })

        menu_num = closed_day.get("menu_num")
        fact_consumption = {}
        snap_cons = closed_day.get("children_consumption", {})
        if menu_num is not None:
            for prod_name, pdata in snap_cons.items():
                pts_q = {}
                for pid in points_dict.keys():
                    pts_q[pid] = pdata.get("points_qty", {}).get(str(pid), pdata.get("points_qty", {}).get(pid, 0.0))
                fact_consumption[prod_name] = {
                    "unit": pdata.get("unit", "кг"),
                    "points_qty": pts_q,
                    "total_fact": pdata.get("total", 0.0)
                }
        return {
            "date": report_date,
            "region": region,
            "menu_num": menu_num,
            "is_closed": True,
            "points": points_list,
            "fact_children": fact_children_out,
            "fact_consumption": fact_consumption
        }

    fact_storage = load_fact_children_settings()
    saved_fact_children = fact_storage.get(report_date, {}).get(region, {})

    points_list = []
    for pid, pinfo in points_dict.items():
        child_cnt = saved_fact_children.get(str(pid))
        if child_cnt is None:
            child_cnt = saved_fact_children.get(pid, 0)
        else:
            child_cnt = int(child_cnt)

        points_list.append({
            "id": pid,
            "name": pinfo["name"],
            "default_children": child_cnt,
            "plan_children": pinfo["children_count"]
        })

    menu_num = get_menu_number_for_date_custom(parsed_date)
    fact_consumption = {}

    if menu_num is not None:
        menu_items = load_menu_from_excel(menu_num)
        for prod_name, pdata in menu_items.items():
            kg_per_child = pdata.get("qty_per_child", 0.0)
            unit = pdata.get("unit", "кг")
            points_qty = {}
            row_sum = 0.0
            for p in points_list:
                pid = p["id"]
                child_cnt = p["default_children"]
                qty = math.ceil(child_cnt * kg_per_child) if unit == "шт" else round(child_cnt * kg_per_child, 2)
                points_qty[pid] = qty
                row_sum += qty

            fact_consumption[prod_name] = {
                "unit": unit,
                "points_qty": points_qty,
                "total_fact": math.ceil(row_sum) if unit == "шт" else round(row_sum, 2)
            }
    return {
        "date": report_date,
        "region": region,
        "menu_num": menu_num,
        "is_closed": False,
        "points": points_list,
        "fact_children": saved_fact_children,
        "fact_consumption": fact_consumption
    }


@app.get("/api/fact-staff/report")
def get_fact_staff_report(report_date: str, region: str = "Город"):
    try:
        parsed_date = date.fromisoformat(report_date)
    except Exception:
        parsed_date = date.today()

    points_dict = load_points_from_excel(region)
    closed_day = get_closed_day_snapshot(report_date, region)

    if closed_day is not None:
        snap_st = closed_day.get("staff_counts", {})
        points_list = []
        fact_staff_out = {}
        for pid, pinfo in points_dict.items():
            s_cnt = _safe_int_or_zero(snap_st.get(str(pid), snap_st.get(pid, 0)))
            fact_staff_out[str(pid)] = s_cnt
            points_list.append({
                "id": pid,
                "name": pinfo["name"],
                "default_staff": s_cnt
            })

        menu_num = closed_day.get("menu_num")
        fact_consumption = {}
        snap_cons = closed_day.get("staff_consumption", {})
        if menu_num is not None:
            for prod_name, pdata in snap_cons.items():
                pts_q = {}
                for pid in points_dict.keys():
                    pts_q[pid] = pdata.get("points_qty", {}).get(str(pid), pdata.get("points_qty", {}).get(pid, 0.0))
                fact_consumption[prod_name] = {
                    "unit": pdata.get("unit", "кг"),
                    "points_qty": pts_q,
                    "total_fact": pdata.get("total", 0.0)
                }
        return {
            "date": report_date,
            "region": region,
            "menu_num": menu_num,
            "is_closed": True,
            "points": points_list,
            "fact_staff": fact_staff_out,
            "fact_consumption": fact_consumption
        }

    fact_storage = load_fact_staff_settings()
    saved_fact_staff = fact_storage.get(report_date, {}).get(region, {})

    points_list = []
    for pid, pinfo in points_dict.items():
        staff_cnt = saved_fact_staff.get(str(pid))
        if staff_cnt is None:
            staff_cnt = saved_fact_staff.get(pid, 0)
        else:
            staff_cnt = int(staff_cnt)

        points_list.append({
            "id": pid,
            "name": pinfo["name"],
            "default_staff": staff_cnt
        })

    menu_num = get_menu_number_for_date_custom(parsed_date)
    fact_consumption = {}

    if menu_num is not None:
        menu_items = load_staff_menu_from_excel(menu_num)
        for prod_name, pdata in menu_items.items():
            kg_per_staff = pdata.get("qty_per_staff", 0.0)
            unit = pdata.get("unit", "кг")
            points_qty = {}
            row_sum = 0.0
            for p in points_list:
                pid = p["id"]
                staff_cnt = p["default_staff"]
                qty = math.ceil(staff_cnt * kg_per_staff) if unit == "шт" else round(staff_cnt * kg_per_staff, 2)
                points_qty[pid] = qty
                row_sum += qty

            fact_consumption[prod_name] = {
                "unit": unit,
                "points_qty": points_qty,
                "total_fact": math.ceil(row_sum) if unit == "шт" else round(row_sum, 2)
            }
    return {
        "date": report_date,
        "region": region,
        "menu_num": menu_num,
        "is_closed": False,
        "points": points_list,
        "fact_staff": saved_fact_staff,
        "fact_consumption": fact_consumption
    }


@app.get("/api/fact/saved-days")
def get_saved_days_list(region: str = "Город"):
    storage = load_fact_children_settings()
    saved_dates = []
    for d_str, regions_data in storage.items():
        if regions_data and region in regions_data and regions_data[region]:
            saved_dates.append({"date": d_str, "is_closed": is_month_closed(d_str, region)})
    return sorted(saved_dates, key=lambda x: x["date"], reverse=True)


@app.get("/api/fact-staff/saved-days")
def get_saved_staff_days_list(region: str = "Город"):
    storage = load_fact_staff_settings()
    saved_dates = []
    for d_str, regions_data in storage.items():
        if regions_data and region in regions_data and regions_data[region]:
            saved_dates.append({"date": d_str, "is_closed": is_month_closed(d_str, region)})
    return sorted(saved_dates, key=lambda x: x["date"], reverse=True)


@app.get("/api/stock/saved-days")
def get_saved_stock_days_list(region: str = "Город"):
    storage = load_stock_settings()
    saved_dates = []
    for d_str, regions_data in storage.items():
        if regions_data and region in regions_data and regions_data[region]:
            saved_dates.append(d_str)
    return sorted(saved_dates, reverse=True)


@app.get("/api/balance/report")
def get_balance_report(report_date: str, region: str = "Город"):
    try:
        parsed_date = date.fromisoformat(report_date)
    except Exception:
        parsed_date = date.today()

    points_dict = load_points_from_excel(region)
    fact_storage = load_fact_children_settings()
    fact_staff_storage = load_fact_staff_settings()
    stock_storage = load_stock_settings()

    closed_target_day = get_closed_day_snapshot(report_date, region)
    saved_fact_children_target = closed_target_day.get("children_counts", {}) if closed_target_day else fact_storage.get(report_date, {}).get(region, {})

    points_list = []
    for pid, pinfo in points_dict.items():
        is_saved = True if closed_target_day is not None else ((str(pid) in saved_fact_children_target) or (pid in saved_fact_children_target))
        if is_saved:
            c_cnt = int(saved_fact_children_target.get(str(pid), saved_fact_children_target.get(pid, 0)) or 0)
        else:
            c_cnt = pinfo["children_count"]

        points_list.append({
            "id": pid,
            "name": pinfo["name"],
            "children_count": c_cnt,
            "is_saved": is_saved
        })

    available_inv_dates = sorted([d for d in stock_storage.keys() if d <= report_date])
    inv_date_used = None
    initial_stock = {}
    if available_inv_dates:
        inv_date_used = available_inv_dates[-1]
        initial_stock = stock_storage[inv_date_used].get(region, {})

    purchases_storage = load_purchases_settings().get(region, {})
    purchases_between = {}
    inv_dt = date.fromisoformat(inv_date_used) if inv_date_used else date(2000, 1, 1)

    for pid, p_list in purchases_storage.items():
        for p in p_list:
            p_date = p.get("date")
            if p_date and inv_dt <= date.fromisoformat(p_date) <= parsed_date:
                if pid not in purchases_between:
                    purchases_between[pid] = {}
                prod = p.get("product")
                purchases_between[pid][prod] = purchases_between[pid].get(prod, 0.0) + p.get("qty", 0.0)

    consumed_between = {}
    if inv_date_used:
        inv_dt_real = date.fromisoformat(inv_date_used)
        if inv_dt_real <= parsed_date:
            curr_d = inv_dt_real
            while curr_d <= parsed_date:
                d_iso = curr_d.isoformat()
                closed_snap = get_closed_day_snapshot(d_iso, region)
                if closed_snap is not None:
                    for prod_name, pdata in closed_snap.get("children_consumption", {}).items():
                        for pid in points_dict.keys():
                            if pid not in consumed_between:
                                consumed_between[pid] = {}
                            q_val = float(pdata.get("points_qty", {}).get(str(pid), pdata.get("points_qty", {}).get(pid, 0.0)) or 0.0)
                            consumed_between[pid][prod_name] = consumed_between[pid].get(prod_name, 0.0) + q_val

                    for prod_name, pdata in closed_snap.get("staff_consumption", {}).items():
                        for pid in points_dict.keys():
                            if pid not in consumed_between:
                                consumed_between[pid] = {}
                            q_val = float(pdata.get("points_qty", {}).get(str(pid), pdata.get("points_qty", {}).get(pid, 0.0)) or 0.0)
                            consumed_between[pid][prod_name] = consumed_between[pid].get(prod_name, 0.0) + q_val
                else:
                    m_num = get_menu_number_for_date_custom(curr_d)
                    if m_num is not None:
                        day_menu = load_menu_from_excel(m_num)
                        day_fact_children = fact_storage.get(d_iso, {}).get(region, {})
                        for pid, pinfo in points_dict.items():
                            child_cnt = day_fact_children.get(str(pid))
                            if child_cnt is None:
                                child_cnt = day_fact_children.get(pid, pinfo["children_count"])
                            else:
                                child_cnt = int(child_cnt)

                            if pid not in consumed_between:
                                consumed_between[pid] = {}
                            for prod_name, pdata in day_menu.items():
                                unit_check = pdata.get("unit", "кг")
                                raw_consumption = pdata.get("qty_per_child", 0.0) * child_cnt
                                day_qty = math.ceil(raw_consumption) if unit_check == "шт" else raw_consumption
                                consumed_between[pid][prod_name] = consumed_between[pid].get(prod_name, 0.0) + day_qty

                        day_staff_menu = load_staff_menu_from_excel(m_num)
                        day_fact_staff = fact_staff_storage.get(d_iso, {}).get(region, {})
                        for pid in points_dict.keys():
                            staff_cnt = day_fact_staff.get(str(pid), day_fact_staff.get(pid, 0))
                            staff_cnt = int(staff_cnt) if staff_cnt is not None else 0
                            if staff_cnt > 0:
                                if pid not in consumed_between:
                                    consumed_between[pid] = {}
                                for prod_name, pdata in day_staff_menu.items():
                                    unit_check = pdata.get("unit", "кг")
                                    raw_consumption = pdata.get("qty_per_staff", 0.0) * staff_cnt
                                    day_qty = math.ceil(raw_consumption) if unit_check == "шт" else raw_consumption
                                    consumed_between[pid][prod_name] = consumed_between[pid].get(prod_name, 0.0) + day_qty

                curr_d += timedelta(days=1)

    balance_report = {}
    current_products = get_all_products_with_units()

    # Продукты старого меню (их уже нет в Excel), по которым в системе есть данные:
    # остаток из акта инвентаризации или расход, зафиксированный в закрытых месяцах.
    legacy_products = set()
    for p_data in initial_stock.values():
        if isinstance(p_data, dict):
            legacy_products.update(p_data.keys())
    for p_cons in consumed_between.values():
        if isinstance(p_cons, dict):
            legacy_products.update(p_cons.keys())
    legacy_products -= set(current_products.keys())
    legacy_units = get_closed_units_map(region) if legacy_products else {}

    all_products_set = set(current_products.keys()) | legacy_products

    for prod_name in sorted(all_products_set):
        unit = legacy_units.get(prod_name, "кг")
        for i in range(1, 11):
            m_items = load_menu_from_excel(i)
            if prod_name in m_items:
                unit = m_items[prod_name]["unit"]
                break

        points_balance = {}
        prod_total_balance = 0.0

        for pid in points_dict.keys():
            p_init = initial_stock.get(str(pid), {}) if str(pid) in initial_stock else initial_stock.get(pid, {})
            init_val = p_init.get(prod_name, 0.0) if isinstance(p_init, dict) else float(p_init or 0.0)

            p_purch = purchases_between.get(pid, {})
            purch_val = p_purch.get(prod_name, 0.0) if isinstance(p_purch, dict) else float(p_purch or 0.0)

            p_cons = consumed_between.get(pid, {})
            cons_val = p_cons.get(prod_name, 0.0) if isinstance(p_cons, dict) else float(p_cons or 0.0)

            bal = (init_val + purch_val) - cons_val
            bal_rounded = float(math.ceil(bal)) if unit == "шт" else round(bal, 2)

            points_balance[pid] = bal_rounded
            prod_total_balance += bal_rounded

        # Продукт старого меню с нулевым остатком везде в отчёте не нужен
        if prod_name in legacy_products and all(v == 0 for v in points_balance.values()):
            continue

        balance_report[prod_name] = {
            "unit": unit,
            "points_balance": points_balance,
            "total_balance": float(math.ceil(prod_total_balance)) if unit == "шт" else round(prod_total_balance, 2)
        }

    return {
        "date": report_date,
        "region": region,
        "inv_date": inv_date_used,
        "points": points_list,
        "balance_report": balance_report
    }


def _apply_excel_table_style(worksheet, df: pd.DataFrame, col_width_mid: float = 10.0, row_height_pt: float = 20.0):
    thin_border = Border(
        left=Side(style='thin', color='CCCCCC'),
        right=Side(style='thin', color='CCCCCC'),
        top=Side(style='thin', color='CCCCCC'),
        bottom=Side(style='thin', color='CCCCCC')
    )
    header_fill = PatternFill(start_color='E0E0E0', end_color='E0E0E0', fill_type='solid')
    header_font = Font(bold=True)

    for row_idx, row in enumerate(worksheet.iter_rows(min_row=1, max_row=len(df) + 1, min_col=1, max_col=len(df.columns)), start=1):
        for cell in row:
            cell.border = thin_border
            if row_idx == 1:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            else:
                cell.alignment = Alignment(horizontal="left" if cell.column == 1 else "center", vertical="center")

    worksheet.row_dimensions[1].height = 25
    for r_i in range(2, len(df) + 2):
        worksheet.row_dimensions[r_i].height = row_height_pt

    worksheet.column_dimensions['A'].width = 21.4
    num_columns = len(df.columns)
    for col_idx in range(2, num_columns):
        col_letter = openpyxl.utils.get_column_letter(col_idx)
        worksheet.column_dimensions[col_letter].width = col_width_mid
    if num_columns >= 2:
        last_col_letter = openpyxl.utils.get_column_letter(num_columns)
        worksheet.column_dimensions[last_col_letter].width = 17.14


@app.get("/api/balance/export/excel")
def export_balance_excel(report_date: str, region: str = "Город"):
    try:
        data = get_balance_report(report_date, region)
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Ошибка расчета остатков: {str(e)}")

    points = data.get("points", [])
    balance_report = data.get("balance_report", {})

    rows_data = []
    for prod_name, prod_info in balance_report.items():
        row = {"Продукт": prod_name, "Ед. изм.": prod_info["unit"]}
        for p in points:
            row[p["name"]] = prod_info["points_balance"].get(p["id"], 0.0)
        row["ИТОГО ОСТАТОК"] = prod_info["total_balance"]
        rows_data.append(row)

    df = pd.DataFrame(rows_data)
    if df.empty:
        columns = ["Продукт", "Ед. изм."] + [p["name"] for p in points] + ["ИТОГО ОСТАТОК"]
        df = pd.DataFrame(columns=columns)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Остатки")
        _apply_excel_table_style(writer.sheets["Остатки"], df, col_width_mid=10.0)

    output.seek(0)
    region_safe = "Gorod" if region == "Город" else "Rayon"
    filename = f"Balance_Report_{region_safe}_{report_date}.xlsx"

    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )


@app.post("/api/fact/children")
def save_fact_children_api(data: dict, report_date: str, region: str = "Город"):
    assert_month_open(report_date, region)
    save_fact_children_for_date(report_date, region, data)
    return {"status": "success"}


@app.delete("/api/fact/children")
def delete_fact_children_api(report_date: str, region: str = "Город"):
    assert_month_open(report_date, region)
    delete_fact_children_for_date(report_date, region)
    return {"status": "success"}


@app.post("/api/fact-staff/staff")
def save_fact_staff_api(data: dict, report_date: str, region: str = "Город"):
    assert_month_open(report_date, region)
    save_fact_staff_for_date(report_date, region, data)
    return {"status": "success"}


@app.delete("/api/fact-staff/staff")
def delete_fact_staff_api(report_date: str, region: str = "Город"):
    assert_month_open(report_date, region)
    delete_fact_staff_for_date(report_date, region)
    return {"status": "success"}


@app.get("/api/fact/export/all-regions-excel")
def export_fact_all_regions_excel(report_date: str):
    try:
        date.fromisoformat(report_date)
    except Exception:
        raise HTTPException(status_code=400, detail="Неверный формат даты")

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for region in ["Город", "Район"]:
            report_data = get_fact_report(report_date, region)
            points_dict = load_points_from_excel(region)

            rows_data = []
            fact_cons = report_data.get("fact_consumption", {})
            for prod_name, pdata in fact_cons.items():
                row = {"Продукт": prod_name, "Ед. изм.": pdata["unit"]}
                for pid, pinfo in points_dict.items():
                    row[pinfo["name"]] = pdata["points_qty"].get(pid, 0.0)
                row["ИТОГО ФАКТ"] = pdata["total_fact"]
                rows_data.append(row)

            df = pd.DataFrame(rows_data)
            if df.empty:
                df = pd.DataFrame(columns=["Продукт", "Ед. изм.", "ИТОГО ФАКТ"])
            df.to_excel(writer, index=False, sheet_name=region)

    output.seek(0)
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename=Fact_Consumption_All_Regions_{report_date}.xlsx"}
    )


@app.get("/api/fact-staff/export/all-regions-excel")
def export_fact_staff_all_regions_excel(report_date: str):
    try:
        date.fromisoformat(report_date)
    except Exception:
        raise HTTPException(status_code=400, detail="Неверный формат даты")

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for region in ["Город", "Район"]:
            report_data = get_fact_staff_report(report_date, region)
            points_dict = load_points_from_excel(region)

            rows_data = []
            fact_cons = report_data.get("fact_consumption", {})
            for prod_name, pdata in fact_cons.items():
                row = {"Продукт": prod_name, "Ед. изм.": pdata["unit"]}
                for pid, pinfo in points_dict.items():
                    row[pinfo["name"]] = pdata["points_qty"].get(pid, 0.0)
                row["ИТОГО ФАКТ"] = pdata["total_fact"]
                rows_data.append(row)

            df = pd.DataFrame(rows_data)
            if df.empty:
                df = pd.DataFrame(columns=["Продукт", "Ед. изм.", "ИТОГО ФАКТ"])
            df.to_excel(writer, index=False, sheet_name=region)

    output.seek(0)
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename=Fact_Staff_Consumption_{report_date}.xlsx"}
    )


@app.get("/api/fact/export/children-matrix")
def export_children_matrix_excel(month: str, region: str = "Город"):
    try:
        parts = month.split("-")
        year = int(parts[0])
        m = int(parts[1])
    except Exception:
        raise HTTPException(status_code=400, detail="Неверный формат месяца. Используйте YYYY-MM")

    start_date = date(year, m, 1)
    end_date = date(year + 1, 1, 1) - timedelta(days=1) if m == 12 else date(year, m + 1, 1) - timedelta(days=1)

    try:
        points_dict = load_points_from_excel(region)
        fact_storage = load_fact_children_settings()

        rows_data = []
        current = start_date
        while current <= end_date:
            d_str = current.isoformat()
            closed_day = get_closed_day_snapshot(d_str, region)

            row = {"Дата": d_str}
            if closed_day is not None:
                snap_ch = closed_day.get("children_counts", {})
                for pid, pinfo in points_dict.items():
                    row[pinfo["name"]] = _safe_int_or_zero(snap_ch.get(str(pid), snap_ch.get(pid, 0)))
            else:
                saved_day_fact = fact_storage.get(d_str, {}).get(region, {})
                for pid, pinfo in points_dict.items():
                    val = saved_day_fact.get(str(pid))
                    if val is None:
                        val = saved_day_fact.get(pid, pinfo["children_count"])
                    try:
                        num_val = int(float(val)) if val != "" and val is not None else int(pinfo["children_count"])
                    except Exception:
                        num_val = 0
                    row[pinfo["name"]] = num_val

            rows_data.append(row)
            current += timedelta(days=1)

        df = pd.DataFrame(rows_data)
        if df.empty:
            df = pd.DataFrame(columns=["Дата"])

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name=f"Дети_{region}_{month}")
        output.seek(0)

        region_safe = "Gorod" if region == "Город" else "Rayon"
        return StreamingResponse(
            output,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f"attachment; filename=Children_Attendance_{region_safe}_{month}.xlsx"}
        )
    except Exception as e:
        print(f"❌ Ошибка в export_children_matrix_excel: {e}")
        raise HTTPException(status_code=500, detail=f"Ошибка генерации файла: {str(e)}")


@app.get("/api/fact-staff/export/staff-matrix")
def export_staff_matrix_excel(month: str, region: str = "Город"):
    try:
        parts = month.split("-")
        year = int(parts[0])
        m = int(parts[1])
    except Exception:
        raise HTTPException(status_code=400, detail="Неверный формат месяца. Используйте YYYY-MM")

    start_date = date(year, m, 1)
    end_date = date(year + 1, 1, 1) - timedelta(days=1) if m == 12 else date(year, m + 1, 1) - timedelta(days=1)

    try:
        points_dict = load_points_from_excel(region)
        staff_storage = load_fact_staff_settings()

        rows_data = []
        current = start_date
        while current <= end_date:
            d_str = current.isoformat()
            closed_day = get_closed_day_snapshot(d_str, region)
            row = {"Дата": d_str}

            if closed_day is not None:
                snap_st = closed_day.get("staff_counts", {})
                for pid, pinfo in points_dict.items():
                    row[pinfo["name"]] = _safe_int_or_zero(snap_st.get(str(pid), snap_st.get(pid, 0)))
            else:
                saved_day_fact = staff_storage.get(d_str, {}).get(region, {})
                for pid, pinfo in points_dict.items():
                    val = saved_day_fact.get(str(pid), saved_day_fact.get(pid, 0))
                    row[pinfo["name"]] = _safe_int_or_zero(val)

            rows_data.append(row)
            current += timedelta(days=1)

        df = pd.DataFrame(rows_data)
        if df.empty:
            df = pd.DataFrame(columns=["Дата"])

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name=f"Сотрудники_{region}_{month}")
        output.seek(0)

        region_safe = "Gorod" if region == "Город" else "Rayon"
        return StreamingResponse(
            output,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f"attachment; filename=Staff_Attendance_{region_safe}_{month}.xlsx"}
        )
    except Exception as e:
        print(f"❌ Ошибка в export_staff_matrix_excel: {e}")
        raise HTTPException(status_code=500, detail=f"Ошибка генерации файла: {str(e)}")


@app.get("/api/fact/export/monthly-mtt-excel")
def export_monthly_mtt_consumption_excel(month: str, region: str = "Город"):
    try:
        parts = month.split("-")
        year = int(parts[0])
        m = int(parts[1])
    except Exception:
        raise HTTPException(status_code=400, detail="Неверный формат месяца. Используйте YYYY-MM")

    start_date = date(year, m, 1)
    end_date = date(year + 1, 1, 1) - timedelta(days=1) if m == 12 else date(year, m + 1, 1) - timedelta(days=1)

    try:
        points_dict = load_points_from_excel(region)
        fact_storage = load_fact_children_settings()

        rows_data = []
        current = start_date
        while current <= end_date:
            d_str = current.isoformat()
            closed_day = get_closed_day_snapshot(d_str, region)

            if closed_day is not None:
                snap_cons = closed_day.get("children_consumption", {})
                for prod_name, pdata in snap_cons.items():
                    unit = pdata.get("unit", "кг")
                    pts_q = pdata.get("points_qty", {})
                    for pid, pinfo in points_dict.items():
                        qty = float(pts_q.get(str(pid), pts_q.get(pid, 0.0)) or 0.0)
                        if qty > 0:
                            rows_data.append({
                                "Дата": d_str,
                                "МТТ": pinfo["name"],
                                "Номенклатура": prod_name,
                                "ед. изм.": unit,
                                "кол-во": qty
                            })
            else:
                menu_num = get_menu_number_for_date_custom(current)
                if menu_num is not None:
                    menu_items = load_menu_from_excel(menu_num)
                    saved_day_fact = fact_storage.get(d_str, {}).get(region, {})

                    for pid, pinfo in points_dict.items():
                        p_name = pinfo["name"]
                        val = saved_day_fact.get(str(pid))
                        if val is None:
                            val = saved_day_fact.get(pid, pinfo["children_count"])

                        try:
                            child_cnt = int(float(val)) if val != "" and val is not None else int(pinfo["children_count"])
                        except Exception:
                            child_cnt = 0

                        for prod_name, pdata in menu_items.items():
                            kg_per_child = pdata.get("qty_per_child", 0.0)
                            unit = pdata.get("unit", "кг")
                            raw_val = kg_per_child * child_cnt
                            qty = float(math.ceil(raw_val)) if unit == "шт" else round(raw_val, 2)
                            if qty > 0:
                                rows_data.append({
                                    "Дата": d_str,
                                    "МТТ": p_name,
                                    "Номенклатура": prod_name,
                                    "ед. изм.": unit,
                                    "кол-во": qty
                                })

            current += timedelta(days=1)

        df = pd.DataFrame(rows_data)
        if df.empty:
            df = pd.DataFrame(columns=["Дата", "МТТ", "Номенклатура", "ед. изм.", "кол-во"])

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name=f"Месячный_расход_{month}")
        output.seek(0)

        region_safe = "Gorod" if region == "Город" else "Rayon"
        filename = f"Monthly_Consumption_MTT_{region_safe}_{month}.xlsx"

        return StreamingResponse(
            output,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f"attachment; filename={filename}"}
        )
    except Exception as e:
        print(f"❌ Ошибка в export_monthly_mtt_consumption_excel: {e}")
        raise HTTPException(status_code=500, detail=f"Ошибка генерации файла: {str(e)}")


@app.get("/api/fact-staff/export/monthly-mtt-excel")
def export_staff_monthly_mtt_consumption_excel(month: str, region: str = "Город"):
    try:
        parts = month.split("-")
        year = int(parts[0])
        m = int(parts[1])
    except Exception:
        raise HTTPException(status_code=400, detail="Неверный формат месяца. Используйте YYYY-MM")

    start_date = date(year, m, 1)
    end_date = date(year + 1, 1, 1) - timedelta(days=1) if m == 12 else date(year, m + 1, 1) - timedelta(days=1)

    try:
        points_dict = load_points_from_excel(region)
        staff_storage = load_fact_staff_settings()

        rows_data = []
        current = start_date
        while current <= end_date:
            d_str = current.isoformat()
            closed_day = get_closed_day_snapshot(d_str, region)

            if closed_day is not None:
                snap_cons = closed_day.get("staff_consumption", {})
                for prod_name, pdata in snap_cons.items():
                    unit = pdata.get("unit", "кг")
                    pts_q = pdata.get("points_qty", {})
                    for pid, pinfo in points_dict.items():
                        qty = float(pts_q.get(str(pid), pts_q.get(pid, 0.0)) or 0.0)
                        if qty > 0:
                            rows_data.append({
                                "Дата": d_str,
                                "МТТ": pinfo["name"],
                                "Номенклатура": prod_name,
                                "ед. изм.": unit,
                                "кол-во": qty
                            })
            else:
                menu_num = get_menu_number_for_date_custom(current)
                if menu_num is not None:
                    menu_items = load_staff_menu_from_excel(menu_num)
                    saved_day_fact = staff_storage.get(d_str, {}).get(region, {})

                    for pid, pinfo in points_dict.items():
                        p_name = pinfo["name"]
                        val = saved_day_fact.get(str(pid), saved_day_fact.get(pid, 0))
                        staff_cnt = _safe_int_or_zero(val)

                        for prod_name, pdata in menu_items.items():
                            kg_per_staff = pdata.get("qty_per_staff", 0.0)
                            unit = pdata.get("unit", "кг")
                            raw_val = kg_per_staff * staff_cnt
                            qty = float(math.ceil(raw_val)) if unit == "шт" else round(raw_val, 2)
                            if qty > 0:
                                rows_data.append({
                                    "Дата": d_str,
                                    "МТТ": p_name,
                                    "Номенклатура": prod_name,
                                    "ед. изм.": unit,
                                    "кол-во": qty
                                })

            current += timedelta(days=1)

        df = pd.DataFrame(rows_data)
        if df.empty:
            df = pd.DataFrame(columns=["Дата", "МТТ", "Номенклатура", "ед. изм.", "кол-во"])

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name=f"Расход_сотр_{month}")
        output.seek(0)

        region_safe = "Gorod" if region == "Город" else "Rayon"
        filename = f"Monthly_Staff_Consumption_MTT_{region_safe}_{month}.xlsx"

        return StreamingResponse(
            output,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f"attachment; filename={filename}"}
        )
    except Exception as e:
        print(f"❌ Ошибка в export_staff_monthly_mtt_consumption_excel: {e}")
        raise HTTPException(status_code=500, detail=f"Ошибка генерации файла: {str(e)}")


@app.post("/api/menu/generate-and-send")
def generate_and_send_menu_api(report_date: str, background_tasks: BackgroundTasks, region: str = "Город"):
    try:
        parsed_date = date.fromisoformat(report_date)
    except Exception:
        raise HTTPException(status_code=400, detail="Неверная дата")

    fact_storage = load_fact_children_settings()
    saved_fact = fact_storage.get(report_date, {}).get(region, {})

    if not saved_fact:
        raise HTTPException(status_code=400, detail=f"Нет сохраненной проводки факта за {report_date} ({region}). Отправка PDF невозможна.")

    menu_num = get_menu_number_for_date_custom(parsed_date)
    if not menu_num:
        raise HTTPException(status_code=400, detail="На указанную дату по календарю выходной день.")

    background_tasks.add_task(
        process_and_send_menu_in_background,
        report_date=report_date,
        region=region,
        saved_fact=saved_fact,
        menu_num=menu_num
    )
    return {"status": "success", "message": "Процесс отправки запущен в фоновом режиме! Файлы будут поступать в Telegram по очереди."}


@app.post("/api/menu-staff/generate-and-send")
def generate_and_send_staff_menu_api(report_date: str, background_tasks: BackgroundTasks, region: str = "Город"):
    try:
        parsed_date = date.fromisoformat(report_date)
    except Exception:
        raise HTTPException(status_code=400, detail="Неверная дата")

    fact_storage = load_fact_staff_settings()
    saved_fact = fact_storage.get(report_date, {}).get(region, {})

    if not saved_fact:
        raise HTTPException(status_code=400, detail=f"Нет сохраненного факта сотрудников за {report_date} ({region}).")

    menu_num = get_menu_number_for_date_custom(parsed_date)
    if not menu_num:
        raise HTTPException(status_code=400, detail="На указанную дату по календарю выходной день.")

    background_tasks.add_task(
        process_and_send_staff_menu_in_background,
        report_date=report_date,
        region=region,
        saved_fact=saved_fact,
        menu_num=menu_num
    )
    return {"status": "success", "message": "Отправка меню сотрудников запущена!"}


# --- ИНВЕНТАРИЗАЦИЯ И ЭКСПОРТ ИНВЕНТАРИЗАЦИИ ---

def _resolve_stock_date(inv_date: str | None = None, target_date: str | None = None) -> str:
    raw = inv_date or target_date
    if raw:
        return str(raw).strip()
    return date.today().isoformat()


@app.get("/api/stock/matrix/{region}")
def get_stock_matrix_api(region: str, inv_date: str = None, target_date: str = None):
    points_dict = load_points_from_excel(region)
    points_list = [{"id": pid, "name": p["name"], "children_count": p["children_count"]} for pid, p in points_dict.items()]
    all_products = dict(get_all_products_with_units())
    stock_storage = load_stock_settings()
    resolved_date = _resolve_stock_date(inv_date, target_date)
    region_stock = stock_storage.get(resolved_date, {}).get(region, {})

    # Если в уже сохранённом акте есть продукты старого меню с остатком, показываем их,
    # иначе при повторном сохранении акта эти данные молча пропадут.
    legacy_names = set()
    for p_data in region_stock.values():
        if not isinstance(p_data, dict):
            continue
        for prod, val in p_data.items():
            if prod in all_products:
                continue
            try:
                if float(val or 0) != 0:
                    legacy_names.add(prod)
            except (TypeError, ValueError):
                continue
    if legacy_names:
        legacy_units = get_closed_units_map(region)
        for prod in legacy_names:
            all_products[prod] = legacy_units.get(prod, "кг")
        all_products = dict(sorted(all_products.items()))

    products_matrix = {}
    for prodName, unit in all_products.items():
        points_stock = {}
        for pid in points_dict.keys():
            p_data = region_stock.get(str(pid))
            if p_data is None:
                p_data = region_stock.get(pid, {})
            val = p_data.get(prodName, 0.0) if isinstance(p_data, dict) else 0.0
            points_stock[pid] = val
        products_matrix[prodName] = {"unit": unit, "points_stock": points_stock}

    return {"points": points_list, "products": products_matrix}


@app.post("/api/stock/matrix/{region}")
def save_stock_matrix_api(region: str, data: dict, inv_date: str = None, target_date: str = None):
    resolved_date = _resolve_stock_date(inv_date, target_date)
    assert_month_open(resolved_date, region)
    save_stock_settings_for_date(resolved_date, region, data)
    return {"status": "success"}


@app.delete("/api/stock/matrix/{region}")
def delete_stock_matrix_api(region: str, inv_date: str = None, target_date: str = None):
    resolved_date = _resolve_stock_date(inv_date, target_date)
    assert_month_open(resolved_date, region)
    delete_stock_settings_for_date(resolved_date, region)
    return {"status": "success"}


@app.get("/api/stock/export/excel")
def export_stock_excel(region: str = "Город", inv_date: str = None, target_date: str = None):
    resolved_date = _resolve_stock_date(inv_date, target_date)
    try:
        date.fromisoformat(resolved_date)
    except Exception:
        raise HTTPException(status_code=400, detail="Неверный формат даты")

    matrix_data = get_stock_matrix_api(region=region, inv_date=resolved_date, target_date=resolved_date)
    points = matrix_data.get("points", [])
    products = matrix_data.get("products", {})

    rows_data = []
    for prod_name, p_info in products.items():
        row = {"Продукт": prod_name, "Ед. изм.": p_info["unit"]}
        total_val = 0.0
        for p in points:
            val = float(p_info["points_stock"].get(p["id"], 0.0) or 0.0)
            row[p["name"]] = val
            total_val += val
        row["ИТОГО ОСТАТОК"] = float(math.ceil(total_val)) if p_info["unit"] == "шт" else round(total_val, 2)
        rows_data.append(row)

    df = pd.DataFrame(rows_data)
    if df.empty:
        columns = ["Продукт", "Ед. изм."] + [p["name"] for p in points] + ["ИТОГО ОСТАТОК"]
        df = pd.DataFrame(columns=columns)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Инвентаризация")
        _apply_excel_table_style(writer.sheets["Инвентаризация"], df, col_width_mid=10.0)

    output.seek(0)
    region_safe = "Gorod" if region == "Город" else "Rayon"
    filename = f"Inventory_{region_safe}_{resolved_date}.xlsx"

    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )


@app.get("/api/calendar/range")
def get_calendar_range_api(start: str, end: str):
    try:
        s_date = date.fromisoformat(start)
        e_date = date.fromisoformat(end)
    except ValueError:
        raise HTTPException(status_code=400, detail="Неверный формат дат")
    cal_map = calculate_calendar_dynamic(s_date, e_date)
    for d_str, info in cal_map.items():
        info["is_closed"] = is_month_closed(d_str, "Город") or is_month_closed(d_str, "Район")
    return cal_map


@app.post("/api/calendar/settings")
def save_setting_api(data: dict):
    date_str = data.get("date")
    if not date_str:
        raise HTTPException(status_code=400, detail="Дата обязательна")
    if is_month_closed(date_str, "Город") or is_month_closed(date_str, "Район"):
        raise HTTPException(status_code=423, detail=f"Дата {date_str} входит в закрытый месяц. Изменение календаря запрещено.")
    settings = load_manual_settings()
    settings["days"][date_str] = {
        "is_holiday": data.get("is_holiday", False),
        "menu": data.get("menu")
    }
    save_manual_settings(settings)
    return {"status": "success"}


@app.get("/api/raznaryadka/matrix/{target_date}")
def get_matrix_raznaryadka(target_date: str, region: str = "Город", days_count: int = 2, filter_period: str = "ALL"):
    try:
        parsed_date = date.fromisoformat(target_date)
    except ValueError:
        raise HTTPException(status_code=400, detail="Неверный формат даты")

    target_days = []
    curr = parsed_date
    while len(target_days) < days_count:
        m_num = get_menu_number_for_date_custom(curr)
        if m_num is not None:
            target_days.append((curr, m_num))
        curr += timedelta(days=1)
        if (curr - parsed_date).days > 90:
            break

    if not target_days:
        raise HTTPException(status_code=400, detail="В выбранном периоде нет рабочих дней")

    week_start_str = target_days[0][0].isoformat()
    week_end_str = target_days[-1][0].isoformat()

    points_dict = load_points_from_excel(region)
    points_list = [{"id": pid, "name": pinfo["name"], "children_count": pinfo["children_count"]} for pid, pinfo in points_dict.items()]

    stock_storage = load_stock_settings()
    available_inv_dates = sorted([d for d in stock_storage.keys() if d <= week_start_str])
    inv_date_used = None
    initial_stock = {}
    if available_inv_dates:
        inv_date_used = available_inv_dates[-1]
        initial_stock = stock_storage[inv_date_used].get(region, {})

    purchases_storage = load_purchases_settings().get(region, {})
    purchases_between = {}
    inv_dt = date.fromisoformat(inv_date_used) if inv_date_used else date(2000, 1, 1)
    first_target_date = target_days[0][0]
    end_prev_period = first_target_date - timedelta(days=1)

    for pid, p_list in purchases_storage.items():
        for p in p_list:
            p_date = p.get("date")
            if p_date and inv_dt <= date.fromisoformat(p_date) <= end_prev_period:
                if pid not in purchases_between:
                    purchases_between[pid] = {}
                prod = p.get("product")
                purchases_between[pid][prod] = purchases_between[pid].get(prod, 0.0) + p.get("qty", 0.0)

    fact_storage = load_fact_children_settings()
    fact_staff_storage = load_fact_staff_settings()
    consumed_between = {}
    if inv_date_used:
        inv_dt_real = date.fromisoformat(inv_date_used)
        if inv_dt_real < first_target_date:
            curr_d = inv_dt_real
            while curr_d <= end_prev_period:
                d_iso = curr_d.isoformat()
                closed_snap = get_closed_day_snapshot(d_iso, region)
                if closed_snap is not None:
                    for prod_name, pdata in closed_snap.get("children_consumption", {}).items():
                        for pid in points_dict.keys():
                            if pid not in consumed_between:
                                consumed_between[pid] = {}
                            q_val = float(pdata.get("points_qty", {}).get(str(pid), pdata.get("points_qty", {}).get(pid, 0.0)) or 0.0)
                            consumed_between[pid][prod_name] = consumed_between[pid].get(prod_name, 0.0) + q_val

                    for prod_name, pdata in closed_snap.get("staff_consumption", {}).items():
                        for pid in points_dict.keys():
                            if pid not in consumed_between:
                                consumed_between[pid] = {}
                            q_val = float(pdata.get("points_qty", {}).get(str(pid), pdata.get("points_qty", {}).get(pid, 0.0)) or 0.0)
                            consumed_between[pid][prod_name] = consumed_between[pid].get(prod_name, 0.0) + q_val
                else:
                    menu_num = get_menu_number_for_date_custom(curr_d)
                    if menu_num is not None:
                        day_menu = load_menu_from_excel(menu_num)
                        day_fact = fact_storage.get(d_iso, {}).get(region, {})
                        for pid, pinfo in points_dict.items():
                            val_ch = day_fact.get(str(pid), day_fact.get(pid))
                            children = int(val_ch) if val_ch is not None else pinfo["children_count"]

                            if pid not in consumed_between:
                                consumed_between[pid] = {}
                            for prod_name, pdata in day_menu.items():
                                unit_check = pdata.get("unit", "кг")
                                raw_cons = pdata.get("qty_per_child", 0.0) * children
                                cons_val = math.ceil(raw_cons) if unit_check == "шт" else raw_cons
                                consumed_between[pid][prod_name] = consumed_between[pid].get(prod_name, 0.0) + cons_val

                        day_staff_menu = load_staff_menu_from_excel(menu_num)
                        day_fact_staff = fact_staff_storage.get(d_iso, {}).get(region, {})
                        for pid in points_dict.keys():
                            val_st = day_fact_staff.get(str(pid), day_fact_staff.get(pid, 0))
                            staff_cnt = int(val_st) if val_st is not None else 0
                            if staff_cnt > 0:
                                if pid not in consumed_between:
                                    consumed_between[pid] = {}
                                for prod_name, pdata in day_staff_menu.items():
                                    unit_check = pdata.get("unit", "кг")
                                    raw_cons = pdata.get("qty_per_staff", 0.0) * staff_cnt
                                    cons_val = math.ceil(raw_cons) if unit_check == "шт" else raw_cons
                                    consumed_between[pid][prod_name] = consumed_between[pid].get(prod_name, 0.0) + cons_val

                curr_d += timedelta(days=1)

    effective_stock = {}
    for pid in points_dict.keys():
        effective_stock[pid] = {}
        p_init = initial_stock.get(str(pid), {})
        p_purch = purchases_between.get(pid, {})
        p_cons = consumed_between.get(pid, {})
        all_prods = set(list(p_init.keys()) + list(p_purch.keys()) + list(p_cons.keys()))
        for prod in all_prods:
            init_val = float(p_init.get(prod, 0.0) or 0.0)
            purch_val = float(p_purch.get(prod, 0.0) or 0.0)
            cons_val = float(p_cons.get(prod, 0.0) or 0.0)
            effective_stock[pid][prod] = (init_val + purch_val) - cons_val

    matrix_data = {}
    working_days_count = 0

    running_stock = {}
    for pid, prods in effective_stock.items():
        running_stock[pid] = prods.copy()

    for day_obj, menu_num in target_days:
        working_days_count += 1
        d_str = day_obj.isoformat()
        day_fact = fact_storage.get(d_str, {}).get(region, {})
        day_staff_fact = fact_staff_storage.get(d_str, {}).get(region, {})
        ttk_items = load_menu_from_excel(menu_num)
        staff_ttk_items = load_staff_menu_from_excel(menu_num)

        for point_id, point_info in points_dict.items():
            val_ch = day_fact.get(str(point_id), day_fact.get(point_id))

            if val_ch is None:
                last_known_fact = None
                for prev_d in sorted(fact_storage.keys(), reverse=True):
                    if prev_d <= d_str and region in fact_storage[prev_d]:
                        p_val = fact_storage[prev_d][region].get(str(point_id), fact_storage[prev_d][region].get(point_id))
                        if p_val is not None:
                            last_known_fact = int(p_val)
                            break
                children = last_known_fact if last_known_fact is not None else point_info["children_count"]
            else:
                children = int(val_ch)

            for prod_name, data in ttk_items.items():
                kg_per_child = data.get("qty_per_child", 0.0)
                unit = data["unit"]
                gross_needed = math.ceil(kg_per_child * children) if unit == "шт" else kg_per_child * children

                if prod_name not in matrix_data:
                    matrix_data[prod_name] = {
                        "unit": unit,
                        "points_qty": {pid: 0.0 for pid in points_dict.keys()},
                        "weekly_total": 0.0
                    }

                point_stocks = running_stock.get(point_id, {})
                available = 0.0
                st_key_matched = None
                for st_prod, st_val in point_stocks.items():
                    if st_prod.strip().lower() == prod_name.strip().lower():
                        available = st_val
                        st_key_matched = st_prod
                        break

                if available >= gross_needed:
                    net_needed = 0.0
                    if st_key_matched:
                        running_stock[point_id][st_key_matched] -= gross_needed
                else:
                    net_needed = gross_needed - available
                    if st_key_matched:
                        running_stock[point_id][st_key_matched] = 0.0

                if unit == "шт":
                    matrix_data[prod_name]["points_qty"][point_id] += float(math.ceil(net_needed))
                else:
                    matrix_data[prod_name]["points_qty"][point_id] += round(net_needed, 2)

            val_st = day_staff_fact.get(str(point_id), day_staff_fact.get(point_id))
            if val_st is None:
                last_known_staff = 0
                for prev_d in sorted(fact_staff_storage.keys(), reverse=True):
                    if prev_d <= d_str and region in fact_staff_storage[prev_d]:
                        p_val = fact_staff_storage[prev_d][region].get(str(point_id), fact_staff_storage[prev_d][region].get(point_id))
                        if p_val is not None:
                            last_known_staff = int(p_val)
                            break
                staff_cnt = last_known_staff
            else:
                staff_cnt = int(val_st)

            if staff_cnt > 0:
                for prod_name, data in staff_ttk_items.items():
                    kg_per_staff = data.get("qty_per_staff", 0.0)
                    unit = data["unit"]
                    gross_needed = math.ceil(kg_per_staff * staff_cnt) if unit == "шт" else kg_per_staff * staff_cnt

                    if prod_name not in matrix_data:
                        matrix_data[prod_name] = {
                            "unit": unit,
                            "points_qty": {pid: 0.0 for pid in points_dict.keys()},
                            "weekly_total": 0.0
                        }

                    point_stocks = running_stock.get(point_id, {})
                    available = 0.0
                    st_key_matched = None
                    for st_prod, st_val in point_stocks.items():
                        if st_prod.strip().lower() == prod_name.strip().lower():
                            available = st_val
                            st_key_matched = st_prod
                            break

                    if available >= gross_needed:
                        net_needed = 0.0
                        if st_key_matched:
                            running_stock[point_id][st_key_matched] -= gross_needed
                    else:
                        net_needed = gross_needed - available
                        if st_key_matched:
                            running_stock[point_id][st_key_matched] = 0.0

                    if unit == "шт":
                        matrix_data[prod_name]["points_qty"][point_id] += float(math.ceil(net_needed))
                    else:
                        matrix_data[prod_name]["points_qty"][point_id] += round(net_needed, 2)

    for prod_name, prod_info in matrix_data.items():
        unit = prod_info["unit"]
        row_sum = 0.0
        for pid in points_dict.keys():
            val = prod_info["points_qty"][pid]
            val = float(math.ceil(val)) if unit == "шт" else round(val, 2)
            prod_info["points_qty"][pid] = val
            row_sum += val

        prod_info["weekly_total"] = float(math.ceil(row_sum)) if unit == "шт" else round(row_sum, 2)

    if filter_period and filter_period != "ALL":
        periodicities = load_all_nomenclature_periodicities()
        filtered_matrix = {}
        for prod_name, prod_info in matrix_data.items():
            prod_mapped_period = None
            for p_key, p_val in periodicities.items():
                if p_key.strip().lower() == prod_name.strip().lower():
                    prod_mapped_period = p_val
                    break

            if prod_mapped_period and prod_mapped_period.lower() == filter_period.lower():
                filtered_matrix[prod_name] = prod_info
        matrix_data = filtered_matrix

    return {
        "week_start": week_start_str,
        "week_end": week_end_str,
        "region": region,
        "inv_date": inv_date_used,
        "working_days_count": working_days_count,
        "points": points_list,
        "matrix": matrix_data
    }


@app.get("/api/raznaryadka/export/excel")
def export_raznaryadka_excel(target_date: str, region: str = "Город", days_count: int = 2, filter_period: str = "ALL"):
    try:
        data = get_matrix_raznaryadka(target_date, region, days_count, filter_period)
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Ошибка расчета данных: {str(e)}")

    points = data.get("points", [])
    matrix = data.get("matrix", {})

    rows_data = []
    for prod_name, prod_info in matrix.items():
        row = {"Продукт": prod_name, "Ед. изм.": prod_info["unit"]}
        for p in points:
            row[p["name"]] = prod_info["points_qty"].get(p["id"], 0.0)
        row["ИТОГО К ЗАКУПКЕ"] = prod_info["weekly_total"]
        rows_data.append(row)

    df = pd.DataFrame(rows_data)
    if df.empty:
        columns = ["Продукт", "Ед. изм."] + [p["name"] for p in points] + ["ИТОГО К ЗАКУПКЕ"]
        df = pd.DataFrame(columns=columns)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Разнарядка")
        _apply_excel_table_style(writer.sheets["Разнарядка"], df, col_width_mid=8.57, row_height_pt=23 * 0.75)

    output.seek(0)
    region_safe = "Gorod" if region == "Город" else "Rayon"
    filename = f"Raznaryadka_{region_safe}_{target_date}_days_{days_count}.xlsx"

    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )
