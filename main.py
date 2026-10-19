from datetime import date, timedelta
import io
import json
import os
import re
import subprocess
import asyncio
from apscheduler.schedulers.background import BackgroundScheduler
from fastapi import BackgroundTasks, FastAPI, HTTPException
from fastapi.responses import HTMLResponse, StreamingResponse
import openpyxl
import pandas as pd
import psycopg2
import psycopg2.extras
from PyPDF2 import PdfReader, PdfWriter
from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfgen import canvas
import requests

app = FastAPI(title="Catering Management API", version="1.0.0")

_MENU_CACHE = {}

TELEGRAM_BOT_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN", "")
TELEGRAM_CHAT_ID = os.environ.get("TELEGRAM_CHAT_ID", "")
TELEGRAM_ADMIN_CHAT_ID = os.environ.get("TELEGRAM_ADMIN_CHAT_ID", "")  

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
  conn.commit()
  cursor.close()
  conn.close()


# Планировщик задач для ежедневной отправки меню
scheduler = BackgroundScheduler()

@app.on_event("startup")
def startup_event():
  init_db()
  scheduler.start()


@app.on_event("shutdown")
def shutdown_event():
  scheduler.shutdown()


def load_manual_settings() -> dict:
  try:
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute(
        "SELECT data FROM app_settings WHERE date_str = 'calendar_config'"
    )
    row = cursor.fetchone()
    cursor.close()
    conn.close()
    if row:
      return (
          row["data"] if isinstance(row["data"], dict) else json.loads(row["data"])
      )
  except Exception as e:
    print(f"Ошибка загрузки календаря из БД: {e}")
  return {"days": {}}


def save_manual_settings(settings: dict):
  try:
    conn = get_db_connection()
    cursor = conn.cursor()
    json_val = json.dumps(settings, ensure_ascii=False)
    cursor.execute(
        """
            INSERT INTO app_settings (date_str, data) VALUES ('calendar_config', %s)
            ON CONFLICT (date_str) DO UPDATE SET data = EXCLUDED.data;
        """,
        (json_val,),
    )
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
      d = r["date_str"]
      reg = r["region"]
      val = r["data"] if isinstance(r["data"], dict) else json.loads(r["data"])
      if d not in stock_map:
        stock_map[d] = {}
      stock_map[d][reg] = val
  except Exception as e:
    print(f"Ошибка загрузки инвентаризации из БД: {e}")
  return stock_map


def save_stock_settings_for_date(date_str: str, region: str, data: dict):
  try:
    conn = get_db_connection()
    cursor = conn.cursor()
    json_val = json.dumps(data, ensure_ascii=False)
    cursor.execute(
        """
            INSERT INTO stock_acts (date_str, region, data) VALUES (%s, %s, %s)
            ON CONFLICT (date_str, region) DO UPDATE SET data = EXCLUDED.data;
        """,
        (date_str, region, json_val),
    )
    conn.commit()
    cursor.close()
    conn.close()
  except Exception as e:
    print(f"Ошибка сохранения инвентаризации в БД: {e}")


def delete_stock_settings_for_date(date_str: str, region: str):
  try:
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute(
        "DELETE FROM stock_acts WHERE date_str = %s AND region = %s;",
        (date_str, region),
    )
    conn.commit()
    cursor.close()
    conn.close()
  except Exception as e:
    print(f"Ошибка удаления инвентаризации из БД: {e}")


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
      d = r["date_str"]
      reg = r["region"]
      val = r["data"] if isinstance(r["data"], dict) else json.loads(r["data"])
      if d not in fact_map:
        fact_map[d] = {}
      fact_map[d][reg] = val
  except Exception as e:
    print(f"Ошибка загрузки факта детей из БД: {e}")
  return fact_map


def save_fact_children_for_date(date_str: str, region: str, data: dict):
  try:
    conn = get_db_connection()
    cursor = conn.cursor()
    json_val = json.dumps(data, ensure_ascii=False)
    cursor.execute(
        """
            INSERT INTO fact_children (date_str, region, data) VALUES (%s, %s, %s)
            ON CONFLICT (date_str, region) DO UPDATE SET data = EXCLUDED.data;
        """,
        (date_str, region, json_val),
    )
    conn.commit()
    cursor.close()
    conn.close()
  except Exception as e:
    print(f"Ошибка сохранения факта детей в БД: {e}")


def delete_fact_children_for_date(date_str: str, region: str):
  try:
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute(
        "DELETE FROM fact_children WHERE date_str = %s AND region = %s;",
        (date_str, region),
    )
    conn.commit()
    cursor.close()
    conn.close()
  except Exception as e:
    print(f"Ошибка удаления факта детей из БД: {e}")


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
    for _, row in df.iterrows():
      prod_name = str(row.iloc[0]).strip()
      val = float(row.iloc[1]) if pd.notna(row.iloc[1]) else 0.0
      unit = (
          str(row.iloc[2]).strip().lower()
          if len(row) > 2 and pd.notna(row.iloc[2])
          else "граммов"
      )
      if "грам" in unit or "гр" in unit:
        qty_in_kg = val / 1000.0
        display_unit = "кг"
      else:
        qty_in_kg = val
        display_unit = unit if unit else "шт"
      menu_items[prod_name] = {"qty_per_child": qty_in_kg, "unit": display_unit}
    _MENU_CACHE[menu_num] = menu_items
    return menu_items
  except Exception as e:
    print(f"Ошибка чтения листа {sheet_name}: {e}")
    return {}


def get_all_products_with_units() -> dict:
  products = {}
  for i in range(1, 11):
    items = load_menu_from_excel(i)
    for p, info in items.items():
      if p not in products:
        products[p] = info["unit"]
  return dict(sorted(products.items()))


# --- Генерация PDF по шаблону Menu shablon.xlsx с 3 штампами/подписями ---
# --- ИСПРАВЛЕННАЯ ФУНКЦИЯ ГЕНЕРАЦИИ PDF ---
# --- ИСПРАВЛЕННАЯ ФУНКЦИЯ ГЕНЕРАЦИИ PDF ---
def generate_menu_pdf_for_point(
    report_date_str: str,
    region: str,
    point_id: int,
    point_name: str,
    children_count: int,
    menu_num: int,
) -> str:
  template_path = "Menu shablon.xlsx"
  if not os.path.exists(template_path):
    raise FileNotFoundError("Шаблон Menu shablon.xlsx не найден")

  # Формат даты ДД.ММ.ГГГГ (например, 19.10.2026)
  try:
    d_obj = date.fromisoformat(report_date_str)
    formatted_date = d_obj.strftime("%d.%m.%Y")
  except Exception:
    formatted_date = report_date_str

  wb = openpyxl.load_workbook(template_path)
  target_sheet_name = str(menu_num)

  if target_sheet_name not in wb.sheetnames:
    raise ValueError(f"Лист меню {target_sheet_name} не найден в шаблоне")

  # Удаляем все лишние листы, оставляем только нужный по календарю день меню
  for sheetname in wb.sheetnames:
    if sheetname != target_sheet_name:
      del wb[sheetname]

  ws = wb[target_sheet_name]

  # Заполняем данные
  ws["A1"] = formatted_date  # Дата в формате ДД.ММ.ГГГГ
  ws["A4"] = children_count  # Кол-во детей
  ws["F1"] = point_name  # Номер МТТ

  # Принудительная подгонка под 1 страницу A4 Landscape
  ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
  ws.page_setup.paperSize = ws.PAPERSIZE_A4
  ws.page_setup.fitToWidth = 1
  ws.page_setup.fitToHeight = 1
  ws.sheet_properties.pageSetUpPr.fitToPage = True

  temp_excel_path = f"temp_{region}_{point_id}_{report_date_str}.xlsx"
  temp_pdf_path = f"temp_{region}_{point_id}_{report_date_str}.pdf"
  final_pdf_path = f"Menu_{point_name}_{report_date_str}.pdf"

  wb.save(temp_excel_path)
  wb.close()

  # Конвертация в PDF через LibreOffice
  subprocess.run(
      [
          "libreoffice",
          "--headless",
          "--convert-to",
          "pdf",
          temp_excel_path,
          "--outdir",
          ".",
      ],
      check=True,
  )

  stamp1_path = "stamp_1.png"  # Подпись технолога
  stamp2_path = "stamp_2.png"  # Круглая печать
  stamp3_path = "stamp_3.png"  # Подпись Ярманова

  if os.path.exists(temp_pdf_path):
    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=landscape(A4))

    # --- ТОЧНАЯ РАССТАНОВКА ШТАМПОВ И ПОДПИСЕЙ ---
    # Лист A4 landscape = 842 x 595 точек

    # 1. Подпись Ярманова (stamp_3.png) — Вверху справа над черточкой Р. Ярманова
    if os.path.exists(stamp3_path):
      can.drawImage(stamp3_path, 680, 500, width=90, height=45, mask="auto")

    # 2. Круглая печать (stamp_2.png) — Вверху справа поверх ТАСДИҚЛАЙМАН
    if os.path.exists(stamp2_path):
      can.drawImage(stamp2_path, 640, 480, width=90, height=90, mask="auto")

    # 3. Подпись технолога (stamp_1.png) — Внизу ровно над черточкой "Технолог"
    if os.path.exists(stamp1_path):
      can.drawImage(
          stamp1_path, 435, 62, width=90, height=38, mask="auto"
      )  # Скорректировано точно на линию технолога

    can.save()
    packet.seek(0)

    stamp_pdf = PdfReader(packet)
    original_pdf = PdfReader(temp_pdf_path)
    output_pdf = PdfWriter()

    for page in original_pdf.pages:
      page.merge_page(stamp_pdf.pages[0])
      output_pdf.add_page(page)

    with open(final_pdf_path, "wb") as f:
      output_pdf.write(f)

    if os.path.exists(temp_excel_path):
      os.remove(temp_excel_path)
    if os.path.exists(temp_pdf_path):
      os.remove(temp_pdf_path)

    return final_pdf_path
  else:
    return temp_pdf_path

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


# --- Асинхронная фоновая задача для генерации и отправки по очереди ---
async def process_and_send_menu_in_background(
    report_date: str, region: str, saved_fact: dict, menu_num: int
):
  points_dict = load_points_from_excel(region)
  print(
      f"🚀 Старт фоновой генерации для {len(points_dict)} точек в регионе"
      f" '{region}'..."
  )

  for pid, pinfo in points_dict.items():
    child_cnt = saved_fact.get(str(pid), saved_fact.get(pid))
    if child_cnt is None:
      continue

    p_name = pinfo["name"]

    try:
      # 1. Генерируем PDF с печатями
      pdf_file = generate_menu_pdf_for_point(
          report_date_str=report_date,
          region=region,
          point_id=pid,
          point_name=p_name,
          children_count=int(child_cnt),
          menu_num=menu_num,
      )

      # 2. Отправляем в Telegram
      caption = (
          f"📋 Официальное меню № {menu_num} для {p_name} ({region})\n📅 Дата:"
          f" {report_date}\n👶 Детей: {child_cnt}"
      )
      send_pdf_to_telegram(pdf_file, caption)

      # 3. Удаляем временный файл
      if pdf_file and os.path.exists(pdf_file):
        os.remove(pdf_file)

      # 4. Пауза 2 секунды, чтобы бесплатный сервер Render успевал освобождать CPU/RAM
      await asyncio.sleep(2)

    except Exception as e:
      print(f"❌ Ошибка при обработке точки {p_name}: {e}")

  print("✅ Все меню успешно сгенерированы и отправлены в фоновом режиме!")


@app.get("/", response_class=HTMLResponse)
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
                    <button onclick="switchTab('fact')" id="btnTabFact" class="px-3 py-2 rounded-md text-sm font-medium bg-gray-200 text-gray-700 hover:bg-gray-300 transition">📋 Ежедневный факт</button>
                    <button onclick="switchTab('balance')" id="btnTabBalance" class="px-3 py-2 rounded-md text-sm font-medium bg-gray-200 text-gray-700 hover:bg-gray-300 transition">📦 Остатки по МТТ</button>
                    <button onclick="switchTab('calendar')" id="btnTabCalendar" class="px-3 py-2 rounded-md text-sm font-medium bg-gray-200 text-gray-700 hover:bg-gray-300 transition">📅 Календарь меню</button>
                    <button onclick="switchTab('stock')" id="btnTabStock" class="px-3 py-2 rounded-md text-sm font-medium bg-gray-200 text-gray-700 hover:bg-gray-300 transition">⚙️ Инвентаризация</button>
                </div>
            </nav>

            <div id="tabMatrix" class="space-y-6">
                <header class="bg-white shadow rounded-lg p-6 flex flex-col md:flex-row justify-between items-center">
                    <div>
                        <h1 class="text-xl font-bold text-gray-800">Матричная разнарядка (Чистая потребность)</h1>
                        <p class="text-sm text-gray-500 mt-1">Автоматический учет инвентаризации, приходов от MUXSIN-FAYZ и расхода</p>
                    </div>
                    <div class="mt-4 md:mt-0 flex flex-wrap items-center gap-3">
                        <div>
                            <label class="text-xs font-medium text-gray-600 block">Регион:</label>
                            <select id="regionSelect" class="border border-gray-300 rounded-md px-3 py-2 text-sm focus:ring-2 focus:ring-blue-500">
                                <option value="Город" selected>Город (Текущие)</option>
                                <option value="Район">Район (Тендер / Будущие)</option>
                            </select>
                        </div>
                        <div>
                            <label class="text-xs font-medium text-gray-600 block">Дата недели:</label>
                            <input type="date" id="targetDate" class="border border-gray-300 rounded-md px-3 py-2 text-sm focus:ring-2 focus:ring-blue-500">
                        </div>
                        <div class="self-end">
                            <button onclick="loadMatrixRaznaryadka()" class="bg-blue-600 text-white px-4 py-2 rounded-md text-sm font-medium hover:bg-blue-700 transition">Сформировать</button>
                        </div>
                    </div>
                </header>
                <div id="matrixResultContainer" class="space-y-6">
                    <div class="bg-white shadow rounded-lg p-8 text-center text-gray-400">Выберите дату и нажмите «Сформировать».</div>
                </div>
            </div>

            <div id="tabFact" class="hidden space-y-6">
                <header class="bg-white shadow rounded-lg p-6 flex flex-col md:flex-row justify-between items-center">
                    <div>
                        <h1 class="text-xl font-bold text-gray-800">📋 Фактический отчет расхода продуктов</h1>
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
                            <button onclick="saveFactChildren()" class="bg-green-600 text-white px-4 py-2 rounded-md text-sm font-medium hover:bg-green-700 transition">💾 Сохранить факт</button>
                            <button onclick="downloadFactExcel()" class="bg-blue-600 text-white px-4 py-2 rounded-md text-sm font-medium hover:bg-blue-700 transition">📥 Скачать в Excel</button>
                            <button onclick="sendMenuTelegramManual()" class="bg-indigo-600 text-white px-4 py-2 rounded-md text-sm font-medium hover:bg-indigo-700 transition">📤 Отправить меню в Telegram</button>
                        </div>
                    </div>
                </header>
                <div id="factContainer" class="space-y-6">
                    <p class="p-8 text-center text-gray-400 bg-white shadow rounded-lg">Загрузка данных...</p>
                </div>
            </div>

            <div id="tabBalance" class="hidden space-y-6">
                <header class="bg-white shadow rounded-lg p-6 flex flex-col md:flex-row justify-between items-center">
                    <div>
                        <h1 class="text-xl font-bold text-gray-800">📦 Актуальные остатки продуктов по МТТ</h1>
                        <p class="text-sm text-gray-500 mt-1">Формула: Остаток = Начальный остаток + Приход (MUXSIN-FAYZ) - Расход фактический</p>
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
                        <div class="self-end">
                            <button onclick="loadBalanceData()" class="bg-indigo-600 text-white px-4 py-2 rounded-md text-sm font-medium hover:bg-indigo-700 transition">Обновить остатки</button>
                        </div>
                    </div>
                </header>
                <div id="balanceContainer" class="space-y-6">
                    <p class="p-8 text-center text-gray-400 bg-white shadow rounded-lg">Загрузка остатков...</p>
                </div>
            </div>

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
                        <div class="self-end">
                            <button onclick="saveStockMatrix()" class="bg-green-600 text-white px-4 py-2 rounded-md text-sm font-medium hover:bg-green-700 transition">💾 Сохранить акт</button>
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
            document.getElementById('targetDate').valueAsDate = new Date();
            document.getElementById('stockDate').valueAsDate = new Date();
            document.getElementById('factDate').valueAsDate = new Date();
            document.getElementById('balanceDate').valueAsDate = new Date();

            let viewYear = new Date().getFullYear();
            let viewMonth = new Date().getMonth();
            let selectedDateStr = '';

            function switchTab(tab) {
                ['matrix', 'fact', 'balance', 'calendar', 'stock'].forEach(t => {
                    const el = document.getElementById('tab' + capitalize(t));
                    const btn = document.getElementById('btnTab' + capitalize(t));
                    if (el) el.classList.add('hidden');
                    if (btn) btn.className = "px-3 py-2 rounded-md text-sm font-medium bg-gray-200 text-gray-700 hover:bg-gray-300 transition";
                });
                const activeTab = document.getElementById('tab' + capitalize(tab));
                const activeBtn = document.getElementById('btnTab' + capitalize(tab));
                if (activeTab) activeTab.classList.remove('hidden');
                if (activeBtn) activeBtn.className = "px-3 py-2 rounded-md text-sm font-medium bg-blue-600 text-white transition";

                if (tab === 'calendar') loadCalendarGrid();
                if (tab === 'stock') loadStockMatrix();
                if (tab === 'fact') loadFactData();
                if (tab === 'balance') loadBalanceData();
            }

            function capitalize(s) { return s.charAt(0).toUpperCase() + s.slice(1); }

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
                    const info = data[dStr] || { is_holiday: false, menu: 1, forced: false };
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
                await fetch('/api/calendar/settings', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ date: selectedDateStr, is_holiday, menu })
                });
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

                    let html = `
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
                                <div class="text-[10px] text-gray-400 mb-1">(План: ${p.default_children})</div>
                                <input type="number" data-point-fact="${p.id}" value="${childVal}" class="fact-child-cell border border-gray-300 rounded px-2 py-1 text-center text-sm w-full font-bold text-green-700 focus:ring-2 focus:ring-green-500">
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
                dates.forEach(dStr => {
                    rowsHtml += `
                        <tr class="hover:bg-gray-50">
                            <td class="px-4 py-3 font-semibold text-gray-800">${dStr}</td>
                            <td class="px-4 py-3 text-center text-gray-600">${region}</td>
                            <td class="px-4 py-3 text-center"><span class="bg-green-100 text-green-800 px-2.5 py-1 rounded-full text-xs font-bold">Сохранено</span></td>
                            <td class="px-4 py-3 text-right space-x-2">
                                <button onclick="selectSavedDate('${dStr}')" class="bg-blue-600 text-white px-3 py-1.5 rounded text-xs font-medium hover:bg-blue-700 transition">✏️ Редактировать</button>
                                <button onclick="deleteFactDate('${dStr}', '${region}')" class="bg-red-600 text-white px-3 py-1.5 rounded text-xs font-medium hover:bg-red-700 transition">🗑 Удалить</button>
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
                if (res.ok) {
                    loadFactData();
                } else {
                    alert('Ошибка при удалении.');
                }
            }

            async function sendMenuTelegramManual() {
                const dateVal = document.getElementById('factDate').value;
                const region = document.getElementById('factRegionSelect').value;
                if(!confirm(`Запустить отправку официального меню с печатью за ${dateVal} (${region}) в Telegram в фоновом режиме?`)) return;

                const res = await fetch(`/api/menu/generate-and-send?report_date=${dateVal}&region=${region}`, { method: 'POST' });
                const data = await res.json();
                if(res.ok) {
                    alert('Процесс отправки запущен! Файлы будут поступать в Telegram по очереди.');
                } else {
                    alert('Ошибка: ' + (data.detail || 'Не удалось запустить отправку'));
                }
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
                    const val = parseInt(cell.value) || 0;
                    payload[pid] = val;
                });

                const res = await fetch(`/api/fact/children?report_date=${dateVal}&region=${region}`, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(payload)
                });

                if (res.ok) {
                    alert('Фактическое присутствие детей успешно сохранено!');
                    loadFactData();
                } else alert('Ошибка сохранения.');
            }

            function downloadFactExcel() {
                const dateVal = document.getElementById('factDate').value;
                const region = document.getElementById('factRegionSelect').value;
                window.location.href = `/api/fact/export/excel?report_date=${dateVal}&region=${region}`;
            }

            async function loadStockMatrix() {
                const dateVal = document.getElementById('stockDate').value;
                const region = document.getElementById('stockRegionSelect').value;
                document.getElementById('stockRegionLabel').innerText = region;
                const container = document.getElementById('stockMatrixContainer');
                container.innerHTML = '<div class="text-center py-12 text-gray-500">Загрузка акта инвентаризации...</div>';

                const res = await fetch(`/api/stock/matrix/${region}?date=${dateVal}`);
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
                const res = await fetch(`/api/stock/matrix/${region}?date=${dStr}`, { method: 'DELETE' });
                if (res.ok) {
                    loadStockMatrix();
                } else {
                    alert('Ошибка при удалении акта.');
                }
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

                const res = await fetch(`/api/stock/matrix/${region}?date=${dateVal}`, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(stockPayload)
                });
                if (res.ok) {
                    alert('Акт инвентаризации успешно сохранен!');
                    loadStockMatrix();
                } else alert('Ошибка сохранения акта.');
            }

            async function loadMatrixRaznaryadka() {
                const dateStr = document.getElementById('targetDate').value;
                const region = document.getElementById('regionSelect').value;
                const container = document.getElementById('matrixResultContainer');
                if (!dateStr) { alert('Выберите дату'); return; }
                container.innerHTML = '<div class="text-center py-12 text-gray-500">Формирование точной закупки (с учетом MUXSIN-FAYZ)...</div>';

                try {
                    const response = await fetch(`/api/raznaryadka/matrix/${dateStr}?region=${region}`);
                    const data = await response.json();
                    let html = `
                        <div class="grid grid-cols-1 md:grid-cols-3 gap-6 mb-6">
                            <div class="bg-white shadow rounded-lg p-6 border-t-4 border-blue-500">
                                <h3 class="text-sm font-medium text-gray-500">Период закупки</h3>
                                <p class="text-lg font-bold text-gray-800 mt-1">${data.week_start} — ${data.week_end}</p>
                            </div>
                            <div class="bg-white shadow rounded-lg p-6 border-t-4 border-green-500">
                                <h3 class="text-sm font-medium text-gray-500">Инвентаризация (база)</h3>
                                <p class="text-lg font-bold text-gray-800 mt-1">${data.inv_date ? data.inv_date : 'Нет актов'}</p>
                            </div>
                            <div class="bg-white shadow rounded-lg p-6 border-t-4 border-purple-500">
                                <h3 class="text-sm font-medium text-gray-500">Рабочих дней</h3>
                                <p class="text-lg font-bold text-purple-600 mt-1">${data.working_days_count} дней</p>
                            </div>
                        </div>

                        <div class="bg-white shadow rounded-lg overflow-hidden">
                            <div class="bg-gray-50 px-6 py-4 border-b border-gray-200">
                                <h3 class="text-lg font-medium text-gray-900">📊 Чистая потребность к закупке</h3>
                            </div>
                            <div class="overflow-x-auto" style="max-height: 700px;">
                                <table class="min-w-full divide-y divide-gray-200 text-xs border-collapse">
                                    <thead class="bg-gray-100 sticky top-0 z-10">
                                        <tr>
                                            <th class="px-4 py-3 text-left font-bold text-gray-700 border-r bg-gray-100 sticky left-0 z-20 min-w-[180px]">Продукт</th>
                                            <th class="px-3 py-3 text-center font-bold text-gray-700 border-r bg-gray-100 min-w-[70px]">Ед.</th>
                    `;
                    data.points.forEach(point => {
                        html += `<th class="px-3 py-3 text-center font-medium text-gray-700 border-r min-w-[90px]">${point.name}<br><span class="text-[10px] text-blue-600">(${point.children_count} дет)</span></th>`;
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
                        html += `<td class="px-4 py-3 text-center font-bold text-blue-900 bg-blue-50">${rowData.weekly_total}</td></tr>`;
                    }
                    html += `</tbody></table></div></div>`;
                    container.innerHTML = html;
                } catch (err) {
                    container.innerHTML = '<div class="text-center py-12 text-red-500">Ошибка загрузки матрицы.</div>';
                }
            }

            window.onload = loadMatrixRaznaryadka;
        </script>
    </body>
    </html>
    """
  return HTMLResponse(content=html_content)


@app.get("/api/fact/report")
def get_fact_report(report_date: str, region: str = "Город"):
  try:
    parsed_date = date.fromisoformat(report_date)
  except Exception:
    parsed_date = date.today()

  points_dict = load_points_from_excel(region)
  fact_storage = load_fact_children_settings()
  saved_fact_children = fact_storage.get(report_date, {}).get(region, {})

  points_list = []
  for pid, pinfo in points_dict.items():
    child_cnt = saved_fact_children.get(str(pid))
    if child_cnt is None:
      child_cnt = saved_fact_children.get(pid, pinfo["children_count"])
    else:
      child_cnt = int(child_cnt)

    points_list.append({
        "id": pid,
        "name": pinfo["name"],
        "default_children": child_cnt,
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
        qty = round(child_cnt * kg_per_child, 2)
        points_qty[pid] = qty
        row_sum += qty

      fact_consumption[prod_name] = {
          "unit": unit,
          "points_qty": points_qty,
          "total_fact": round(row_sum, 2),
      }

  return {
      "date": report_date,
      "region": region,
      "menu_num": menu_num,
      "points": points_list,
      "fact_children": saved_fact_children,
      "fact_consumption": fact_consumption,
  }


@app.get("/api/fact/saved-days")
def get_saved_days_list(region: str = "Город"):
  storage = load_fact_children_settings()
  saved_dates = []
  for d_str, regions_data in storage.items():
    if regions_data and region in regions_data and regions_data[region]:
      saved_dates.append(d_str)
  return sorted(saved_dates, reverse=True)


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
  stock_storage = load_stock_settings()

  saved_fact_children_target = (
      fact_storage.get(report_date, {}).get(region, {})
  )
  points_list = []
  for pid, pinfo in points_dict.items():
    is_saved = (str(pid) in saved_fact_children_target) or (
        pid in saved_fact_children_target
    )
    if is_saved:
      c_cnt = int(
          saved_fact_children_target.get(
              str(pid), saved_fact_children_target.get(pid)
          )
      )
    else:
      c_cnt = pinfo["children_count"]

    points_list.append({
        "id": pid,
        "name": pinfo["name"],
        "children_count": c_cnt,
        "is_saved": is_saved,
    })

  available_inv_dates = sorted(
      [d for d in stock_storage.keys() if d <= report_date]
  )
  inv_date_used = None
  initial_stock = {}
  if available_inv_dates:
    inv_date_used = available_inv_dates[-1]
    initial_stock = stock_storage[inv_date_used].get(region, {})

  purchases_storage = load_purchases_settings().get(region, {})
  purchases_between = {}
  inv_dt = (
      date.fromisoformat(inv_date_used) if inv_date_used else date(2000, 1, 1)
  )

  for pid, p_list in purchases_storage.items():
    for p in p_list:
      p_date = p.get("date")
      if p_date and inv_dt <= date.fromisoformat(p_date) <= parsed_date:
        if pid not in purchases_between:
          purchases_between[pid] = {}
        prod = p.get("product")
        purchases_between[pid][prod] = (
            purchases_between[pid].get(prod, 0.0) + p.get("qty", 0.0)
        )

  consumed_between = {}
  if inv_date_used:
    inv_dt_real = date.fromisoformat(inv_date_used)
    if inv_dt_real <= parsed_date:
      curr_d = inv_dt_real
      while curr_d <= parsed_date:
        m_num = get_menu_number_for_date_custom(curr_d)
        if m_num is not None:
          day_menu = load_menu_from_excel(m_num)
          day_fact_children = (
              fact_storage.get(curr_d.isoformat(), {}).get(region, {})
          )
          for pid, pinfo in points_dict.items():
            child_cnt = day_fact_children.get(str(pid))
            if child_cnt is None:
              child_cnt = day_fact_children.get(pid, pinfo["children_count"])
            else:
              child_cnt = int(child_cnt)

            if pid not in consumed_between:
              consumed_between[pid] = {}
            for prod_name, pdata in day_menu.items():
              consumed_between[pid][prod_name] = consumed_between[pid].get(
                  prod_name, 0.0
              ) + (pdata.get("qty_per_child", 0.0) * child_cnt)
        curr_d += timedelta(days=1)

  balance_report = {}
  all_products_set = set(get_all_products_with_units().keys())

  for prod_name in sorted(all_products_set):
    unit = "кг"
    for i in range(1, 11):
      m_items = load_menu_from_excel(i)
      if prod_name in m_items:
        unit = m_items[prod_name]["unit"]
        break

    points_balance = {}
    prod_total_balance = 0.0

    for pid in points_dict.keys():
      p_init = (
          initial_stock.get(str(pid), {})
          if str(pid) in initial_stock
          else initial_stock.get(pid, {})
      )
      init_val = (
          p_init.get(prod_name, 0.0)
          if isinstance(p_init, dict)
          else float(p_init or 0.0)
      )

      p_purch = purchases_between.get(pid, {})
      purch_val = (
          p_purch.get(prod_name, 0.0)
          if isinstance(p_purch, dict)
          else float(p_purch or 0.0)
      )

      p_cons = consumed_between.get(pid, {})
      cons_val = (
          p_cons.get(prod_name, 0.0)
          if isinstance(p_cons, dict)
          else float(p_cons or 0.0)
      )

      bal = (init_val + purch_val) - cons_val
      bal_rounded = round(bal, 2)
      points_balance[pid] = bal_rounded
      prod_total_balance += bal_rounded

    balance_report[prod_name] = {
        "unit": unit,
        "points_balance": points_balance,
        "total_balance": round(prod_total_balance, 2),
    }

  return {
      "date": report_date,
      "region": region,
      "inv_date": inv_date_used,
      "points": points_list,
      "balance_report": balance_report,
  }


@app.post("/api/fact/children")
def save_fact_children_api(data: dict, report_date: str, region: str = "Город"):
  save_fact_children_for_date(report_date, region, data)
  return {"status": "success"}


@app.delete("/api/fact/children")
def delete_fact_children_api(report_date: str, region: str = "Город"):
  delete_fact_children_for_date(report_date, region)
  return {"status": "success"}


@app.get("/api/fact/export/excel")
def export_fact_excel(report_date: str, region: str = "Город"):
  try:
    parsed_date = date.fromisoformat(report_date)
  except Exception:
    parsed_date = date.today()

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

  output = io.BytesIO()
  with pd.ExcelWriter(output, engine="openpyxl") as writer:
    df.to_excel(writer, index=False, sheet_name=f"Факт_{report_date}")
  output.seek(0)

  return StreamingResponse(
      output,
      media_type=(
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      ),
      headers={
          "Content-Disposition": (
              f"attachment; filename=Fact_Consumption_{region}_{report_date}.xlsx"
          )
      },
  )


# --- Эндпоинт ручного запуска отправки через фоновую задачу ---
@app.post("/api/menu/generate-and-send")
def generate_and_send_menu_api(
    report_date: str, background_tasks: BackgroundTasks, region: str = "Город"
):
  try:
    parsed_date = date.fromisoformat(report_date)
  except Exception:
    raise HTTPException(status_code=400, detail="Неверная дата")

  fact_storage = load_fact_children_settings()
  saved_fact = fact_storage.get(report_date, {}).get(region, {})

  if not saved_fact:
    raise HTTPException(
        status_code=400,
        detail=(
            f"Нет сохраненной проводки факта за {report_date} ({region})."
            " Отправка PDF невозможна."
        ),
    )

  menu_num = get_menu_number_for_date_custom(parsed_date)
  if not menu_num:
    raise HTTPException(
        status_code=400, detail="На указанную дату по календарю выходной день."
    )

  # Запуск в фоновом режиме, чтобы UI не зависал
  background_tasks.add_task(
      process_and_send_menu_in_background,
      report_date=report_date,
      region=region,
      saved_fact=saved_fact,
      menu_num=menu_num,
  )

  return {
      "status": "success",
      "message": (
          "Процесс отправки запущен в фоновом режиме! Файлы будут поступать в"
          " Telegram по очереди."
      ),
  }


@app.get("/api/stock/matrix/{region}")
def get_stock_matrix_api(region: str, date: str = None):
  points_dict = load_points_from_excel(region)
  points_list = [
      {
          "id": pid,
          "name": p["name"],
          "children_count": p["children_count"],
      }
      for pid, p in points_dict.items()
  ]
  all_products = get_all_products_with_units()
  stock_storage = load_stock_settings()
  target_inv_date = date if date else str(date.today())
  region_stock = stock_storage.get(target_inv_date, {}).get(region, {})

  products_matrix = {}
  for prodName, unit in all_products.items():
    points_stock = {}
    for pid in points_dict.keys():
      points_stock[pid] = region_stock.get(str(pid), {}).get(prodName, 0.0)
    products_matrix[prodName] = {"unit": unit, "points_stock": points_stock}

  return {"points": points_list, "products": products_matrix}


@app.post("/api/stock/matrix/{region}")
def save_stock_matrix_api(region: str, data: dict, date: str = "2026-10-01"):
  save_stock_settings_for_date(date, region, data)
  return {"status": "success"}


@app.delete("/api/stock/matrix/{region}")
def delete_stock_matrix_api(date: str, region: str):
  delete_stock_settings_for_date(date, region)
  return {"status": "success"}


@app.get("/api/calendar/range")
def get_calendar_range_api(start: str, end: str):
  try:
    s_date = date.fromisoformat(start)
    e_date = date.fromisoformat(end)
  except ValueError:
    raise HTTPException(status_code=400, detail="Неверный формат дат")
  return calculate_calendar_dynamic(s_date, e_date)


@app.post("/api/calendar/settings")
def save_setting_api(data: dict):
  date_str = data.get("date")
  if not date_str:
    raise HTTPException(status_code=400, detail="Дата обязательна")
  settings = load_manual_settings()
  settings["days"][date_str] = {
      "is_holiday": data.get("is_holiday", False),
      "menu": data.get("menu"),
  }
  save_manual_settings(settings)
  return {"status": "success"}


@app.get("/api/raznaryadka/matrix/{target_date}")
def get_matrix_raznaryadka(target_date: str, region: str = "Город"):
  try:
    parsed_date = date.fromisoformat(target_date)
  except ValueError:
    raise HTTPException(status_code=400, detail="Неверный формат дат")

  start_of_week = parsed_date - timedelta(days=parsed_date.weekday())
  week_days = [start_of_week + timedelta(days=i) for i in range(5)]

  points_dict = load_points_from_excel(region)
  points_list = [
      {
          "id": pid,
          "name": pinfo["name"],
          "children_count": pinfo["children_count"],
      }
      for pid, pinfo in points_dict.items()
  ]

  stock_storage = load_stock_settings()
  available_inv_dates = sorted(
      [d for d in stock_storage.keys() if d <= start_of_week.isoformat()]
  )
  inv_date_used = None
  initial_stock = {}
  if available_inv_dates:
    inv_date_used = available_inv_dates[-1]
    initial_stock = stock_storage[inv_date_used].get(region, {})

  purchases_storage = load_purchases_settings().get(region, {})
  purchases_between = {}

  inv_dt = (
      date.fromisoformat(inv_date_used) if inv_date_used else date(2000, 1, 1)
  )
  end_prev_period = start_of_week - timedelta(days=1)

  for pid, p_list in purchases_storage.items():
    for p in p_list:
      p_date = p.get("date")
      if p_date and inv_dt <= date.fromisoformat(p_date) <= end_prev_period:
        if pid not in purchases_between:
          purchases_between[pid] = {}
        prod = p.get("product")
        purchases_between[pid][prod] = (
            purchases_between[pid].get(prod, 0.0) + p.get("qty", 0.0)
        )

  consumed_between = {}
  if inv_date_used:
    inv_dt_real = date.fromisoformat(inv_date_used)
    if inv_dt_real < start_of_week:
      curr_d = inv_dt_real
      while curr_d <= end_prev_period:
        menu_num = get_menu_number_for_date_custom(curr_d)
        if menu_num is not None:
          day_menu = load_menu_from_excel(menu_num)
          for pid, pinfo in points_dict.items():
            children = pinfo["children_count"]
            if pid not in consumed_between:
              consumed_between[pid] = {}
            for prod_name, pdata in day_menu.items():
              consumed_between[pid][prod_name] = consumed_between[pid].get(
                  prod_name, 0.0
              ) + (pdata.get("qty_per_child", 0.0) * children)
        curr_d += timedelta(days=1)

  effective_stock = {}
  for pid in points_dict.keys():
    effective_stock[pid] = {}
    p_init = initial_stock.get(str(pid), {})
    p_purch = purchases_between.get(pid, {})
    p_cons = consumed_between.get(pid, {})
    all_prods = set(
        list(p_init.keys()) + list(p_purch.keys()) + list(p_cons.keys())
    )
    for prod in all_prods:
      init_val = p_init.get(prod, 0.0)
      purch_val = p_purch.get(prod, 0.0)
      cons_val = p_cons.get(prod, 0.0)
      net_stock = (init_val + purch_val) - cons_val
      effective_stock[pid][prod] = net_stock

  matrix_data = {}
  working_days_count = 0
  for day in week_days:
    menu_num = get_menu_number_for_date_custom(day)
    if menu_num is None:
      continue
    working_days_count += 1
    ttk_items = load_menu_from_excel(menu_num)
    for point_id, point_info in points_dict.items():
      children = point_info["children_count"]
      for prod_name, data in ttk_items.items():
        kg_per_child = data.get("qty_per_child", 0.0)
        unit = data["unit"]
        needed = kg_per_child * children
        if prod_name not in matrix_data:
          matrix_data[prod_name] = {
              "unit": unit,
              "points_qty": {pid: 0.0 for pid in points_dict.keys()},
              "weekly_total": 0.0,
          }
        matrix_data[prod_name]["points_qty"][point_id] += needed

  for prod_name, prod_info in matrix_data.items():
    weekly_sum = 0.0
    for pid in points_dict.keys():
      gross_qty = prod_info["points_qty"][pid]
      stock_at_start = 0.0
      p_stocks = effective_stock.get(pid, {})
      for st_prod, st_qty in p_stocks.items():
        if st_prod.strip().lower() == prod_name.strip().lower():
          stock_at_start += st_qty

      net_qty = max(0.0, gross_qty - stock_at_start)
      prod_info["points_qty"][pid] = round(net_qty, 2)
      weekly_sum += net_qty
    prod_info["weekly_total"] = round(weekly_sum, 2)

  return {
      "week_start": week_days[0].isoformat(),
      "week_end": week_days[4].isoformat(),
      "region": region,
      "inv_date": inv_date_used,
      "working_days_count": working_days_count,
      "points": points_list,
      "matrix": matrix_data,
  }
