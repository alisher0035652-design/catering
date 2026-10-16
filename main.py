from datetime import date, timedelta
import io
import json
import os
import re
from fastapi import FastAPI, HTTPException
from fastapi.responses import HTMLResponse, StreamingResponse
import pandas as pd
import psycopg2
import psycopg2.extras

app = FastAPI(title="Catering Management API", version="1.0.0")

SETTINGS_FILE = "calendar_settings.json"
STOCK_FILE = "stock_settings.json"
FACT_CHILDREN_FILE = "fact_children_settings.json"

_MENU_CACHE = {}


def get_db_connection():
  database_url = os.environ.get("DATABASE_URL")
  if database_url and database_url.startswith("postgres://"):
    database_url = database_url.replace("postgres://", "postgresql://", 1)

  if database_url:
    conn = psycopg2.connect(
        database_url, cursor_factory=psycopg2.extras.RealDictCursor
    )
  else:
    import sqlite3

    conn = sqlite3.connect("warehouse.db")
    conn.row_factory = sqlite3.Row
  return conn


def load_manual_settings() -> dict:
  if os.path.exists(SETTINGS_FILE):
    try:
      with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
        return json.load(f)
    except Exception:
      return {}
  return {"days": {}}


def save_manual_settings(settings: dict):
  with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
    json.dump(settings, f, ensure_ascii=False, indent=4)


def load_stock_settings() -> dict:
  if os.path.exists(STOCK_FILE):
    try:
      with open(STOCK_FILE, "r", encoding="utf-8") as f:
        return json.load(f)
    except Exception:
      return {}
  return {}


def save_stock_settings(settings: dict):
  with open(STOCK_FILE, "w", encoding="utf-8") as f:
    json.dump(settings, f, ensure_ascii=False, indent=4)


def load_fact_children_settings() -> dict:
  if os.path.exists(FACT_CHILDREN_FILE):
    try:
      with open(FACT_CHILDREN_FILE, "r", encoding="utf-8") as f:
        return json.load(f)
    except Exception:
      return {}
  return {}


def save_fact_children_settings(settings: dict):
  with open(FACT_CHILDREN_FILE, "w", encoding="utf-8") as f:
    json.dump(settings, f, ensure_ascii=False, indent=4)


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

      region = (
          "Город" if "город" in category_raw or "текущие" in category_raw else "Район"
      )

      points_dict = points_gorod if region == "Город" else points_rayon
      db_client_num = int(r["client_number"])

      target_point_id = None
      for pid, pinfo in points_dict.items():
        name_str = str(pinfo["name"])
        numbers = re.findall(r"\d+", name_str)
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
            "qty": float(r["quantity"]),
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
            "forced": False,
        }
    else:
      if forced_menu is not None:
        current_menu = int(forced_menu)
        if current >= start_date:
          calendar_map[d_str] = {
              "date": d_str,
              "is_holiday": False,
              "menu": current_menu,
              "forced": True,
          }
      else:
        if current >= start_date:
          calendar_map[d_str] = {
              "date": d_str,
              "is_holiday": False,
              "menu": current_menu,
              "forced": False,
          }
      current_menu = (current_menu % 10) + 1
    current += timedelta(days=1)
  return calendar_map


def get_menu_number_for_date_custom(target_date: date) -> int | None:
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
      if mtt_name and mtt_name != "nan":
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
    for _,
