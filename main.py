from datetime import date, timedelta
import json
import os
from fastapi import FastAPI, HTTPException
from fastapi.responses import HTMLResponse
import pandas as pd

app = FastAPI(title="Catering Management API", version="1.0.0")

SETTINGS_FILE = "calendar_settings.json"


def load_manual_settings() -> dict:
  """Загружает ручные настройки календаря (какой день — выходной, или какое меню

  принудительно).
  """
  if os.path.exists(SETTINGS_FILE):
    try:
      with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
        return json.load(f)
    except Exception:
      return {}
  return {"days": {}}  # Формат: {"2026-11-01": {"is_holiday": True, "menu": null}}


def save_manual_settings(settings: dict):
  with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
    json.dump(settings, f, ensure_ascii=False, indent=4)


def get_menu_number_for_date_custom(target_date: date) -> int | None:
  """Определяет номер меню на дату с учетом ручных настроек и стандартных выходных."""
  date_str = target_date.isoformat()
  settings = load_manual_settings()

  # 1. Проверяем ручные настройки для конкретной даты
  if date_str in settings.get("days", {}):
    day_conf = settings["days"][date_str]
    if day_conf.get("is_holiday", False):
      return None  # Вручную отмечено как выходной / нерабочий день
    if "menu" in day_conf and day_conf["menu"] is not None:
      return int(day_conf["menu"])  # Принудительно заданный номер меню (от 1 до 10)

  # 2. Если ручных настроек нет, смотрим стандартные правила: Сб и Вс — выходные
  if target_date.weekday() >= 5:
    return None

  # 3. Автоматический расчет по циклу (если нет ручных переопределений)
  anchor_date = date(2026, 1, 5)
  anchor_menu = 1

  if target_date < anchor_date:
    return 1

  current = anchor_date
  work_days_passed = 0
  while current < target_date:
    current += timedelta(days=1)
    # Проверяем, не был ли промежуточный день отмечен как выходной вручную или Сб/Вс
    curr_str = current.isoformat()
    is_hol = False
    if curr_str in settings.get("days", {}):
      is_hol = settings["days"][curr_str].get("is_holiday", False)

    if not is_hol and current.weekday() < 5:
      work_days_passed += 1

  menu_number = ((anchor_menu - 1 + work_days_passed) % 10) + 1
  return menu_number


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


@app.get("/", response_class=HTMLResponse)
def get_dashboard():
  """Главная страница с переключением между Матрицей и Настройками Календаря."""
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
            <!-- Навигация -->
            <nav class="bg-white shadow rounded-lg p-4 mb-6 flex justify-between items-center">
                <div class="flex items-center space-x-3">
                    <span class="text-xl font-bold text-blue-600">🍽 Кетеринг ERP</span>
                    <span class="text-xs bg-blue-100 text-blue-800 px-2.5 py-1 rounded-full font-medium">Система разнарядки</span>
                </div>
                <div class="flex space-x-2">
                    <button onclick="switchTab('matrix')" id="btnTabMatrix" class="px-4 py-2 rounded-md text-sm font-medium bg-blue-600 text-white transition">📊 Матрица разнарядки</button>
                    <button onclick="switchTab('calendar')" id="btnTabCalendar" class="px-4 py-2 rounded-md text-sm font-medium bg-gray-200 text-gray-700 hover:bg-gray-300 transition">⚙️ Настройка календаря меню</button>
                </div>
            </nav>

            <!-- Вкладка 1: Матрица -->
            <div id="tabMatrix" class="space-y-6">
                <header class="bg-white shadow rounded-lg p-6 flex flex-col md:flex-row justify-between items-center">
                    <div>
                        <h1 class="text-xl font-bold text-gray-800">Матричная разнарядка (Продукты × МТТ)</h1>
                        <p class="text-sm text-gray-500 mt-1">Учитывает ручные переопределения дней и праздников</p>
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

            <!-- Вкладка 2: Настройка календаря -->
            <div id="tabCalendar" class="hidden space-y-6">
                <header class="bg-white shadow rounded-lg p-6">
                    <h1 class="text-xl font-bold text-gray-800">Ручная настройка календаря и меню</h1>
                    <p class="text-sm text-gray-500 mt-1">Здесь вы можете указать для конкретного дня статус выходного (отдыха) или принудительно назначить номер меню (от 1 до 10).</p>
                </header>

                <div class="bg-white shadow rounded-lg p-6 space-y-6">
                    <div class="grid grid-cols-1 md:grid-cols-4 gap-4 items-end">
                        <div>
                            <label class="block text-xs font-medium text-gray-700 mb-1">Дата:</label>
                            <input type="date" id="calDate" class="w-full border border-gray-300 rounded-md px-3 py-2 text-sm">
                        </div>
                        <div>
                            <label class="block text-xs font-medium text-gray-700 mb-1">Тип дня:</label>
                            <select id="calIsHoliday" class="w-full border border-gray-300 rounded-md px-3 py-2 text-sm">
                                <option value="false">Рабочий день</option>
                                <option value="true">Выходной / Праздник (Отдых)</option>
                            </select>
                        </div>
                        <div>
                            <label class="block text-xs font-medium text-gray-700 mb-1">Принудительное меню (1-10):</label>
                            <select id="calMenuNum" class="w-full border border-gray-300 rounded-md px-3 py-2 text-sm">
                                <option value="">Автоматически (по циклу)</option>
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
                        <div>
                            <button onclick="saveCalendarSetting()" class="w-full bg-green-600 text-white px-4 py-2 rounded-md text-sm font-medium hover:bg-green-700 transition">Сохранить настройку</button>
                        </div>
                    </div>

                    <hr class="my-4">

                    <div>
                        <h3 class="text-md font-bold text-gray-800 mb-3">Сохраненные ручные настройки дней:</h3>
                        <div id="settingsListContainer" class="overflow-x-auto">
                            <p class="text-sm text-gray-500">Загрузка...</p>
                        </div>
                    </div>
                </div>
            </div>
        </div>

        <script>
            document.getElementById('targetDate').valueAsDate = new Date();
            document.getElementById('calDate').valueAsDate = new Date();

            function switchTab(tab) {
                if (tab === 'matrix') {
                    document.getElementById('tabMatrix').classList.remove('hidden');
                    document.getElementById('tabCalendar').classList.add('hidden');
                    document.getElementById('btnTabMatrix').className = "px-4 py-2 rounded-md text-sm font-medium bg-blue-600 text-white transition";
                    document.getElementById('btnTabCalendar').className = "px-4 py-2 rounded-md text-sm font-medium bg-gray-200 text-gray-700 hover:bg-gray-300 transition";
                } else {
                    document.getElementById('tabMatrix').classList.add('hidden');
                    document.getElementById('tabCalendar').classList.remove('hidden');
                    document.getElementById('btnTabCalendar').className = "px-4 py-2 rounded-md text-sm font-medium bg-blue-600 text-white transition";
                    document.getElementById('btnTabMatrix').className = "px-4 py-2 rounded-md text-sm font-medium bg-gray-200 text-gray-700 hover:bg-gray-300 transition";
                    loadCalendarSettingsList();
                }
            }

            async function loadMatrixRaznaryadka() {
                const dateStr = document.getElementById('targetDate').value;
                const region = document.getElementById('regionSelect').value;
                const container = document.getElementById('matrixResultContainer');
                
                if (!dateStr) { alert('Выберите дату'); return; }
                container.innerHTML = '<div class="text-center py-12 text-gray-500">Формирование матрицы...</div>';

                try {
                    const response = await fetch(`/api/raznaryadka/matrix/${dateStr}?region=${region}`);
                    const data = await response.json();

                    let html = `
                        <div class="grid grid-cols-1 md:grid-cols-3 gap-6 mb-6">
                            <div class="bg-white shadow rounded-lg p-6 border-t-4 border-blue-500">
                                <h3 class="text-sm font-medium text-gray-500">Период недели</h3>
                                <p class="text-lg font-bold text-gray-800 mt-1">${data.week_start} — ${data.week_end}</p>
                            </div>
                            <div class="bg-white shadow rounded-lg p-6 border-t-4 border-green-500">
                                <h3 class="text-sm font-medium text-gray-500">Регион</h3>
                                <p class="text-lg font-bold text-gray-800 mt-1">${data.region} (${data.points.length} точек)</p>
                            </div>
                            <div class="bg-white shadow rounded-lg p-6 border-t-4 border-purple-500">
                                <h3 class="text-sm font-medium text-gray-500">Активные дни</h3>
                                <p class="text-lg font-bold text-purple-600 mt-1">${data.working_days_count} рабочих дней</p>
                            </div>
                        </div>

                        <div class="bg-white shadow rounded-lg overflow-hidden">
                            <div class="bg-gray-50 px-6 py-4 border-b border-gray-200 flex justify-between items-center">
                                <h3 class="text-lg font-medium text-gray-900">📊 Сводная матрица закупки на неделю</h3>
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

                    html += `
                                            <th class="px-4 py-3 text-center font-bold text-blue-800 bg-blue-50">ИТОГО ЗА НЕДЕЛЮ</th>
                                        </tr>
                                    </thead>
                                    <tbody class="divide-y divide-gray-200 bg-white">
                    `;

                    for (const [prodName, rowData] of Object.entries(data.matrix)) {
                        html += `
                            <tr class="hover:bg-gray-50">
                                <td class="px-4 py-3 font-medium text-gray-900 border-r bg-white sticky left-0 z-10">${prodName}</td>
                                <td class="px-3 py-3 text-center text-gray-500 border-r bg-gray-50">${rowData.unit}</td>
                        `;
                        data.points.forEach(point => {
                            const val = rowData.points_qty[point.id] || 0;
                            html += `<td class="px-3 py-3 text-center border-r text-gray-700 ${val > 0 ? 'font-semibold' : 'text-gray-300'}">${val > 0 ? val : '-'}</td>`;
                        });
                        html += `
                                <td class="px-4 py-3 text-center font-bold text-blue-900 bg-blue-50">${rowData.weekly_total}</td>
                            </tr>
                        `;
                    }

                    html += `</tbody></table></div></div>`;
                    container.innerHTML = html;
                } catch (err) {
                    container.innerHTML = '<div class="text-center py-12 text-red-500">Ошибка загрузки матрицы.</div>';
                }
            }

            async function saveCalendarSetting() {
                const date = document.getElementById('calDate').value;
                const is_holiday = document.getElementById('calIsHoliday').value === 'true';
                const menuVal = document.getElementById('calMenuNum').value;
                const menu = menuVal ? parseInt(menuVal) : null;

                if (!date) { alert('Выберите дату'); return; }

                const response = await fetch('/api/calendar/settings', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ date, is_holiday, menu })
                });

                if (response.ok) {
                    alert('Настройка успешно сохранена!');
                    loadCalendarSettingsList();
                } else {
                    alert('Ошибка сохранения');
                }
            }

            async function loadCalendarSettingsList() {
                const container = document.getElementById('settingsListContainer');
                const response = await fetch('/api/calendar/settings');
                const data = await response.json();
                
                const days = data.days || {};
                if (Object.keys(days).length === 0) {
                    container.innerHTML = '<p class="text-sm text-gray-500">Ручных настроек пока нет. Все дни рассчитываются автоматически.</p>';
                    return;
                }

                let html = '<table class="min-w-full divide-y divide-gray-200 text-sm"><thead><tr><th class="px-4 py-2 text-left">Дата</th><th class="px-4 py-2 text-left">Статус</th><th class="px-4 py-2 text-left">Меню</th><th class="px-4 py-2 text-right">Действие</th></tr></thead><tbody class="divide-y divide-gray-200">';
                for (const [d, conf] of Object.entries(days)) {
                    html += `<tr>
                        <td class="px-4 py-2 font-medium">${d}</td>
                        <td class="px-4 py-2">${conf.is_holiday ? '<span class="text-red-600 font-bold">Выходной / Отдых</span>' : '<span class="text-green-600 font-bold">Рабочий день</span>'}</td>
                        <td class="px-4 py-2">${conf.menu !== null ? 'Меню № ' + conf.menu : 'Авто по циклу'}</td>
                        <td class="px-4 py-2 text-right"><button onclick="deleteSetting('${d}')" class="text-red-600 hover:text-red-800 text-xs font-semibold">Удалить</button></td>
                    </tr>`;
                }
                html += '</tbody></table>';
                container.innerHTML = html;
            }

            async function deleteSetting(dateStr) {
                await fetch(`/api/calendar/settings/${dateStr}`, { method: 'DELETE' });
                loadCalendarSettingsList();
            }

            window.onload = loadMatrixRaznaryadka;
        </script>
    </body>
    </html>
    """
  return HTMLResponse(content=html_content)


@app.get("/api/calendar/settings")
def get_settings_api():
  return load_manual_settings()


@app.post("/api/calendar/settings")
def save_setting_api(data: dict):
  """Сохраняет настройку дня: {"date": "2026-11-01", "is_holiday": false, "menu":

  3}
  """
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


@app.delete("/api/calendar/settings/{date_str}")
def delete_setting_api(date_str: str):
  settings = load_manual_settings()
  if date_str in settings["days"]:
    del settings["days"][date_str]
    save_manual_settings(settings)
  return {"status": "success"}


@app.get("/api/raznaryadka/matrix/{target_date}")
def get_matrix_raznaryadka(target_date: str, region: str = "Город"):
  try:
    parsed_date = date.fromisoformat(target_date)
  except ValueError:
    raise HTTPException(status_code=400, detail="Неверный формат даты")

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

  matrix_data = {}
  working_days_count = 0

  for day in week_days:
    menu_num = get_menu_number_for_date_custom(day)
    if menu_num is None:
      continue  # Пропускаем выходные и дни отдыха

    working_days_count += 1
    ttk_items = load_menu_from_excel(menu_num)

    for point_id, point_info in points_dict.items():
      children = point_info["children_count"]
      for prod_name, data in ttk_items.items():
        kg_per_child = data["qty_per_child"]
        unit = data["unit"]
        needed = kg_per_child * children

        if prod_name not in matrix_data:
          matrix_data[prod_name] = {
              "unit": unit,
              "points_qty": {pid: 0.0 for pid in points_dict.keys()},
              "weekly_total": 0.0,
          }
        matrix_data[prod_name]["points_qty"][point_id] += needed
        matrix_data[prod_name]["weekly_total"] += needed

  for prod_name in matrix_data:
    matrix_data[prod_name]["weekly_total"] = round(
        matrix_data[prod_name]["weekly_total"], 2
    )
    for pid in matrix_data[prod_name]["points_qty"]:
      matrix_data[prod_name]["points_qty"][pid] = round(
          matrix_data[prod_name]["points_qty"][pid], 2
      )

  return {
      "week_start": week_days[0].isoformat(),
      "week_end": week_days[4].isoformat(),
      "region": region,
      "working_days_count": working_days_count,
      "points": points_list,
      "matrix": matrix_data,
  }
