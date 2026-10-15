from datetime import date, timedelta
import json
import os
from fastapi import FastAPI, HTTPException
from fastapi.responses import HTMLResponse
import pandas as pd

app = FastAPI(title="Catering Management API", version="1.0.0")

SETTINGS_FILE = "calendar_settings.json"
STOCK_FILE = "stock_settings.json"


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
  """Структура: { region: { point_id: { product_name: qty } } }"""
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
      menu_items[prod_name] = {"unit": display_unit}
    return menu_items
  except Exception as e:
    print(f"Ошибка чтения листа {sheet_name}: {e}")
    return {}


def get_all_products_with_units() -> dict:
  """Собирает все продукты и их единицы измерения со всех листов меню."""
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
                    <span class="text-xs bg-blue-100 text-blue-800 px-2.5 py-1 rounded-full font-medium">Система разнарядки</span>
                </div>
                <div class="flex space-x-2">
                    <button onclick="switchTab('matrix')" id="btnTabMatrix" class="px-4 py-2 rounded-md text-sm font-medium bg-blue-600 text-white transition">📊 Матрица разнарядки</button>
                    <button onclick="switchTab('calendar')" id="btnTabCalendar" class="px-4 py-2 rounded-md text-sm font-medium bg-gray-200 text-gray-700 hover:bg-gray-300 transition">📅 Календарь меню</button>
                    <button onclick="switchTab('stock')" id="btnTabStock" class="px-4 py-2 rounded-md text-sm font-medium bg-gray-200 text-gray-700 hover:bg-gray-300 transition">📦 Начальные остатки</button>
                </div>
            </nav>

            <!-- Вкладка 1: Матрица -->
            <div id="tabMatrix" class="space-y-6">
                <header class="bg-white shadow rounded-lg p-6 flex flex-col md:flex-row justify-between items-center">
                    <div>
                        <h1 class="text-xl font-bold text-gray-800">Матричная разнарядка (Продукты × МТТ)</h1>
                        <p class="text-sm text-gray-500 mt-1">Автоматический расчет по календарю с учетом дней отдыха</p>
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

            <!-- Вкладка 2: Календарь -->
            <div id="tabCalendar" class="hidden space-y-6">
                <header class="bg-white shadow rounded-lg p-6 flex flex-col md:flex-row justify-between items-center">
                    <div>
                        <h1 class="text-xl font-bold text-gray-800">Интерактивный календарь меню</h1>
                        <p class="text-sm text-gray-500 mt-1">Кликните на любой день, чтобы изменить его статус или номер меню — цепочка пересчитается автоматически!</p>
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

            <!-- Вкладка 3: Начальные остатки (Матрица: Продукты х МТТ) -->
            <div id="tabStock" class="hidden space-y-6">
                <header class="bg-white shadow rounded-lg p-6 flex flex-col md:flex-row justify-between items-center">
                    <div>
                        <h1 class="text-xl font-bold text-gray-800">Матрица начальных остатков (Продукты × МТТ)</h1>
                        <p class="text-sm text-gray-500 mt-1">Введите текущие остатки продуктов по всем точкам региона</p>
                    </div>
                    <div class="flex items-center space-x-3 mt-4 md:mt-0">
                        <div>
                            <label class="text-xs font-medium text-gray-600 block">Регион:</label>
                            <select id="stockRegionSelect" onchange="loadStockMatrix()" class="border border-gray-300 rounded-md px-3 py-2 text-sm">
                                <option value="Город" selected>Город</option>
                                <option value="Район">Район</option>
                            </select>
                        </div>
                        <div class="self-end">
                            <button onclick="saveStockMatrix()" class="bg-green-600 text-white px-4 py-2 rounded-md text-sm font-medium hover:bg-green-700 transition">💾 Сохранить все остатки</button>
                        </div>
                    </div>
                </header>

                <div class="bg-white shadow rounded-lg overflow-hidden">
                    <div id="stockMatrixContainer" class="overflow-x-auto" style="max-height: 700px;">
                        <p class="p-8 text-center text-gray-400">Загрузка матрицы остатков...</p>
                    </div>
                </div>
            </div>
        </div>

        <!-- Модальное окно календаря -->
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
                        <option value="1">Меню № 1</option>
                        <option value="2">Меню № 2</option>
                        <option value="3">Меню № 3</option>
                        <option value="4">Меню № 4</option>
                        <option value="5">Меню № 5</option>
                        <option value="6">Меню № 6</option>
                        <option value="7">Меню № 7</option>
                        <option value="8">Меню № 8</option>
                        <option value="9">Меню № 9</option>
                        <option value="10">Меню № 10</option>
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
                document.getElementById('currentMonthLabel').innerText = `${monthNames[viewMonth]}${viewYear}`;

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
                        <div onclick="openDayModal('${dStr}',${info.is_holiday}, ${info.menu !== null ? info.menu : 'null'})" class="border ${bgClass} rounded-lg p-3 min-h-[105px] flex flex-col justify-between cursor-pointer shadow-sm transition">
                            <div class="flex justify-between items-center">
                                <span class="font-bold text-sm text-gray-700">${d}</span>${info.forced ? '<span class="text-[10px] bg-purple-200 text-purple-800 px-1 rounded">Fix</span>' : ''}
                            </div>
                            <div class="text-center my-auto">
                                <span class="inline-block px-2.5 py-1 rounded-md text-xs font-bold ${badgeBg}">${statusText}</span>
                            </div>
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
                const region = document.getElementById('stockRegionSelect').value;
                const container = document.getElementById('stockMatrixContainer');
                container.innerHTML = '<div class="text-center py-12 text-gray-500">Загрузка матрицы остатков...</div>';

                const res = await fetch(`/api/stock/matrix/${region}`);
                const data = await res.json(); // { points: [...], products: { prodName: { unit: 'кг', points_stock: {point_id: qty} } } }

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
                                <input
