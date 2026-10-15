from datetime import date, timedelta
from fastapi import FastAPI, HTTPException
from fastapi.responses import HTMLResponse
import pandas as pd
from calendar_logic import get_menu_number_for_date

app = FastAPI(title="Catering Management API", version="1.0.0")


def load_points_from_excel(region: str = "Город") -> dict:
  """Загружает список точек (МТТ) и точное количество детей из Location.xlsx."""
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

        points[idx + 1] = {
            "name": mtt_name,
            "children_count": children,
        }
    return points
  except Exception as e:
    print(f"Ошибка чтения точек из Location.xlsx ({region}): {e}")
    return {}


def load_menu_from_excel(menu_num: int) -> dict:
  """Загружает техкарту из листа Menu.xlsx (Грамм -> Кг)."""
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

      menu_items[prod_name] = {
          "qty_per_child": qty_in_kg,
          "unit": display_unit,
      }
    return menu_items
  except Exception as e:
    print(f"Ошибка чтения листа {sheet_name}: {e}")
    return {}


@app.get("/", response_class=HTMLResponse)
def get_dashboard():
  """Дашборд логиста с разнарядкой на неделю."""
  html_content = """
    <!DOCTYPE html>
    <html lang="ru">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Разнарядка Кетеринг на неделю</title>
        <script src="https://cdn.tailwindcss.com"></script>
    </head>
    <body class="bg-gray-50 text-gray-800 font-sans">
        <div class="max-w-7xl mx-auto px-4 py-8">
            <header class="bg-white shadow rounded-lg p-6 mb-8 flex flex-col md:flex-row justify-between items-center">
                <div>
                    <h1 class="text-2xl font-bold text-blue-600">🍽 Разнарядка Кетеринг — Планирование на Неделю</h1>
                    <p class="text-sm text-gray-500 mt-1">Расчет продуктов на всю рабочую неделю (Пн-Пт) из Menu.xlsx и Location.xlsx</p>
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
                        <label class="text-xs font-medium text-gray-600 block">Любой день недели:</label>
                        <input type="date" id="targetDate" class="border border-gray-300 rounded-md px-3 py-2 text-sm focus:ring-2 focus:ring-blue-500">
                    </div>
                    <div class="self-end">
                        <button onclick="loadWeeklyRaznaryadka()" class="bg-blue-600 text-white px-4 py-2 rounded-md text-sm font-medium hover:bg-blue-700 transition">Сформировать на неделю</button>
                    </div>
                </div>
            </header>

            <div id="resultContainer" class="space-y-8">
                <div class="bg-white shadow rounded-lg p-8 text-center text-gray-400">
                    Выберите дату, регион и нажмите «Сформировать на неделю».
                </div>
            </div>
        </div>

        <script>
            document.getElementById('targetDate').valueAsDate = new Date();

            async function loadWeeklyRaznaryadka() {
                const dateStr = document.getElementById('targetDate').value;
                const region = document.getElementById('regionSelect').value;
                const container = document.getElementById('resultContainer');
                
                if (!dateStr) {
                    alert('Выберите дату');
                    return;
                }

                container.innerHTML = '<div class="text-center py-12 text-gray-500">Расчет разнарядки на всю неделю по 5 рабочим дням...</div>';

                try {
                    const response = await fetch(`/api/raznaryadka/week/${dateStr}?region=${region}`);
                    const data = await response.json();

                    let html = `
                        <div class="grid grid-cols-1 md:grid-cols-3 gap-6">
                            <div class="bg-white shadow rounded-lg p-6 border-t-4 border-blue-500">
                                <h3 class="text-sm font-medium text-gray-500">Период недели</h3>
                                <p class="text-lg font-bold text-gray-800 mt-1">${data.week_start} — ${data.week_end}</p>
                            </div>
                            <div class="bg-white shadow rounded-lg p-6 border-t-4 border-green-500">
                                <h3 class="text-sm font-medium text-gray-500">Регион / Точки</h3>
                                <p class="text-lg font-bold text-gray-800 mt-1">${data.region} (${data.points_count} точек)</p>
                            </div>
                            <div class="bg-white shadow rounded-lg p-6 border-t-4 border-purple-500">
                                <h3 class="text-sm font-medium text-gray-500">Рабочих дней в расчете</h3>
                                <p class="text-lg font-bold text-purple-600 mt-1">${data.working_days.length} дня (Пн-Пт)</p>
                            </div>
                        </div>

                        <!-- Сводный список закупки на всю неделю -->
                        <div class="bg-white shadow rounded-lg overflow-hidden">
                            <div class="bg-gray-50 px-6 py-4 border-b border-gray-200">
                                <h3 class="text-lg font-medium text-gray-900">📦 Итоговый список закупки НА ВСЮ НЕДЕЛЮ для сектора "${data.region}"</h3>
                            </div>
                            <div class="p-6 overflow-x-auto">
                                <table class="min-w-full divide-y divide-gray-200">
                                    <thead>
                                        <tr>
                                            <th class="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Продукт</th>
                                            <th class="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Общий объем за неделю</th>
                                            <th class="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Ед. изм.</th>
                                        </tr>
                                    </thead>
                                    <tbody class="divide-y divide-gray-200">
                    `;

                    for (const [prodName, item] of Object.entries(data.weekly_purchase_list)) {
                        html += `
                            <tr>
                                <td class="px-6 py-4 whitespace-nowrap text-sm font-medium text-gray-900">${prodName}</td>
                                <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-700 font-bold">${item.total_qty}</td>
                                <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-500">${item.unit}</td>
                            </tr>
                        `;
                    }

                    html += `
                                    </tbody>
                                </table>
                            </div>
                        </div>

                        <!-- Детализация по дням недели -->
                        <div class="space-y-6">
                            <h3 class="text-xl font-bold text-gray-800">📅 Детализация по дням недели</h3>
                    `;

                    data.working_days.forEach(day => {
                        html += `
                            <div class="bg-white shadow rounded-lg overflow-hidden">
                                <div class="bg-blue-50 px-6 py-3 border-b border-gray-200 flex justify-between items-center">
                                    <h4 class="font-bold text-blue-900">Дата: ${day.date} (Меню № ${day.menu_number})</h4>
                                    <span class="text-xs text-gray-600">Всего точек: ${data.points_count}</span>
                                </div>
                                <div class="p-4 grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
                        `;

                        day.points_breakdown.forEach(point => {
                            html += `
                                <div class="border border-gray-200 rounded-lg p-3 bg-gray-50">
                                    <div class="flex justify-between items-center mb-2">
                                        <span class="font-bold text-xs text-gray-800">${point.point_name}</span>
                                        <span class="bg-blue-100 text-blue-800 text-[10px] font-semibold px-2 py-0.5 rounded">Детей: ${point.children_count}</span>
                                    </div>
                                    <ul class="divide-y divide-gray-200 text-xs">
                            `;
                            for (const [prodName, need] of Object.entries(point.needs)) {
                                html += `
                                    <li class="py-1 flex justify-between">
                                        <span class="text-gray-600">${prodName}:</span>
                                        <span class="font-medium text-gray-900">${need.qty} ${need.unit}</span>
                                    </li>
                                `;
                            }
                            html += `</ul></div>`;
                        });

                        html += `</div></div>`;
                    });

                    html += `</div>`;
                    container.innerHTML = html;

                } catch (error) {
                    container.innerHTML = '<div class="text-center py-12 text-red-500">Ошибка при расчете недели. Проверьте данные.</div>';
                    console.error(error);
                }
            }

            window.onload = loadWeeklyRaznaryadka;
        </script>
    </body>
    </html>
    """
  return HTMLResponse(content=html_content)


@app.get("/api/raznaryadka/week/{target_date}")
def get_weekly_raznaryadka(target_date: str, region: str = "Город"):
  """Считает разнарядку на всю рабочую неделю (Пн-Пт), в которую входит

  target_date.
  """
  try:
    parsed_date = date.fromisoformat(target_date)
  except ValueError:
    raise HTTPException(
        status_code=400, detail="Неверный формат даты. Используйте ГГГГ-ММ-ДД"
    )

  # Находим понедельник текущей недели
  start_of_week = parsed_date - timedelta(days=parsed_date.weekday())
  # Рабочая неделя: Пн (0) - Пт (4)
  week_days = [start_of_week + timedelta(days=i) for i in range(5)]

  points_dict = load_points_from_excel(region)
  weekly_summary = {}
  daily_reports = []

  for day in week_days:
    menu_num = get_menu_number_for_date(day)
    if menu_num is None:
      continue  # Пропускаем выходные/праздники если попали

    ttk_items = load_menu_from_excel(menu_num)
    points_report = []

    for point_id, point_info in points_dict.items():
      children = point_info["children_count"]
      point_needs = {}

      for prod_name, data in ttk_items.items():
        kg_per_child = data["qty_per_child"]
        unit = data["unit"]

        total_needed = kg_per_child * children
        rounded_qty = round(total_needed, 2)

        point_needs[prod_name] = {
            "qty": rounded_qty,
            "unit": unit,
        }

        # Суммируем в общий недельный закуп
        if prod_name not in weekly_summary:
          weekly_summary[prod_name] = {"total_qty": 0.0, "unit": unit}
        weekly_summary[prod_name]["total_qty"] += total_needed

      points_report.append({
          "point_id": point_id,
          "point_name": point_info["name"],
          "children_count": children,
          "needs": point_needs,
      })

    daily_reports.append({
        "date": day.isoformat(),
        "menu_number": menu_num,
        "points_breakdown": points_report,
    })

  weekly_summary_rounded = {
      prod: {
          "total_qty": round(info["total_qty"], 2),
          "unit": info["unit"],
      }
      for prod, info in weekly_summary.items()
  }

  return {
      "week_start": week_days[0].isoformat(),
      "week_end": week_days[4].isoformat(),
      "region": region,
      "points_count": len(points_dict),
      "working_days": daily_reports,
      "weekly_purchase_list": weekly_summary_rounded,
  }
