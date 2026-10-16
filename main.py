from datetime import date, timedelta
import json
import os
from fastapi import FastAPI, HTTPException
from fastapi.responses import HTMLResponse
import pandas as pd
import psycopg2
import psycopg2.extras

app = FastAPI(title="Catering Management API", version="1.0.0")

SETTINGS_FILE = "calendar_settings.json"
STOCK_FILE = "stock_settings.json"


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


def load_purchases_settings() -> dict:
  """Автоматически подтягивает отгрузки организации 'MUXSIN-FAYZ' из базы данных

  и конвертирует их в приходы продуктов по точкам МТТ.
  """
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

    for r in rows:
      date_str = str(r["date"]).split("T")[0] if r["date"] else ""
      category = str(r["category"]).strip().lower()
      region = (
          "Город" if "город" in category or "текущие" in category else "Район"
      )

      points_dict = load_points_from_excel(region)
      target_point_id = None
      for pid, pinfo in points_dict.items():
        if str(r["client_number"]) in str(pinfo["name"]):
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
            <nav class="bg-white shadow rounded-lg p-4 mb-6 flex justify-between items-center">
                <div class="flex items-center space-x-3">
                    <span class="text-xl font-bold text-blue-600">🍽 Кетеринг ERP</span>
                    <span class="text-xs bg-blue-100 text-blue-800 px-2.5 py-1 rounded-full font-medium">Интегрировано с MUXSIN-FAYZ</span>
                </div>
                <div class="flex space-x-2">
                    <button onclick="switchTab('matrix')" id="btnTabMatrix" class="px-4 py-2 rounded-md text-sm font-medium bg-blue-600 text-white transition">📊 Матрица разнарядки</button>
                    <button onclick="switchTab('calendar')" id="btnTabCalendar" class="px-4 py-2 rounded-md text-sm font-medium bg-gray-200 text-gray-700 hover:bg-gray-300 transition">📅 Календарь меню</button>
                    <button onclick="switchTab('stock')" id="btnTabStock" class="px-4 py-2 rounded-md text-sm font-medium bg-gray-200 text-gray-700 hover:bg-gray-300 transition">📦 Инвентаризация</button>
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
                <div class="bg-white shadow rounded-lg overflow-hidden">
                    <div id="stockMatrixContainer" class="overflow-x-auto" style="max-height: 700px;">
                        <p class="p-8 text-center text-gray-400">Загрузка матрицы инвентаризации...</p>
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
            let viewYear = new Date().getFullYear();
            let viewMonth = new Date().getMonth();
            let selectedDateStr = '';

            function switchTab(tab) {
                ['matrix', 'calendar', 'stock'].forEach(t => {
                    document.getElementById('tab' + capitalize(t)).classList.add('hidden');
                    document.getElementById('btnTab' + capitalize(t)).className = "px-4 py-2 rounded-md text-sm font-medium bg-gray-200 text-gray-700 hover:bg-gray-300 transition";
                });
                document.getElementById('tab' + capitalize(tab)).classList.remove('hidden');
                document.getElementById('btnTab' + capitalize(tab)).className = "px-4 py-2 rounded-md text-sm font-medium bg-blue-600 text-white transition";
                if (tab === 'calendar') loadCalendarGrid();
                if (tab === 'stock') loadStockMatrix();
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

            async function loadStockMatrix() {
                const dateVal = document.getElementById('stockDate').value;
                const region = document.getElementById('stockRegionSelect').value;
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
                if (res.ok) alert('Акт инвентаризации успешно сохранен!');
                else alert('Ошибка сохранения акта.');
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
  target_inv_date = (
      date
      if date
      else str(date.today())
      if "date" in globals()
      else "2026-10-01"
  )
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
  stock_storage = load_stock_settings()
  if date not in stock_storage:
    stock_storage[date] = {}
  stock_storage[date][region] = data
  save_stock_settings(stock_storage)
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

  # 1. Инвентаризация
  stock_storage = load_stock_settings()
  available_inv_dates = sorted(
      [d for d in stock_storage.keys() if d <= start_of_week.isoformat()]
  )
  inv_date_used = None
  initial_stock = {}
  if available_inv_dates:
    inv_date_used = available_inv_dates[-1]
    initial_stock = stock_storage[inv_date_used].get(region, {})

  # 2. Автоматические приходы от MUXSIN-FAYZ
  purchases_storage = load_purchases_settings().get(region, {})
  purchases_between = {}  # { point_id: { prod: qty } }

  if inv_date_used:
    inv_dt = date.fromisoformat(inv_date_used)
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

  # 3. Расход за прошлые дни
  consumed_between = {}
  if inv_date_used:
    inv_dt = date.fromisoformat(inv_date_used)
    if inv_dt < start_of_week:
      curr_d = inv_dt
      end_prev_period = start_of_week - timedelta(days=1)
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

  # 4. Эффективный остаток на начало недели
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
      effective_stock[pid][prod] = max(0.0, net_stock)

  # 5. Валовая потребность недели
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

  # 6. Расчет чистой потребности (Нетто)
  for prod_name, prod_info in matrix_data.items():
    weekly_sum = 0.0
    for pid in points_dict.keys():
      gross_qty = prod_info["points_qty"][pid]
      stock_at_start = effective_stock.get(pid, {}).get(prod_name, 0.0)
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
