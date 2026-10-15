from datetime import date
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
  """Загружает техкарту из листа Menu.xlsx.

  Колонка 0: Название продукта Колонка 1: Расход на 1 ребенка Колонка 2: Ед. изм.
  (Грамм -> переводим в кг, деля на 1000)
  """
  sheet_name = f"{menu_num}-Меню"
  try:
    df = pd.read_excel("Menu.xlsx", sheet_name=sheet_name)
    menu_items = {}
    for _, row in df.iterrows():
      prod_name = str(row.iloc[0]).strip()
      val = float(row.iloc[1]) if pd.notna(row.iloc[1]) else 0.0

      # Проверяем 3-ю колонку на единицу измерения (если есть)
      unit = str(row.iloc[2]).strip().lower() if len(row) > 2 and pd.notna(row.iloc[2]) else "граммов"

      # Если указано в граммах, переводим в килограммы для закупа
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
  """Дашборд логиста с выбором региона и даты."""
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
        <div class="max-w-7xl mx-auto px-4 py-8">
            <header class="bg-white shadow rounded-lg p-6 mb-8 flex flex-col md:flex-row justify-between items-center">
                <div>
                    <h1 class="text-2xl font-bold text-blue-600">🍽 Управление Разнарядкой Кетеринга</h1>
                    <p class="text-sm text-gray-500 mt-1">Автоматический расчет из Menu.xlsx (с учетом ед. изм.) и Location.xlsx</p>
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
                        <label class="text-xs font-medium text-gray-600 block">Дата доставки:</label>
                        <input type="date" id="targetDate" class="border border-gray-300 rounded-md px-3 py-2 text-sm focus:ring-2 focus:ring-blue-500">
                    </div>
                    <div class="self-end">
                        <button onclick="loadRaznaryadka()" class="bg-blue-600 text-white px-4 py-2 rounded-md text-sm font-medium hover:bg-blue-700 transition">Сформировать</button>
                    </div>
                </div>
            </header>

            <div id="resultContainer" class="space-y-8">
                <div class="bg-white shadow rounded-lg p-8 text-center text-gray-400">
                    Выберите регион, дату и нажмите «Сформировать».
                </div>
            </div>
        </div>

        <script>
            document.getElementById('targetDate').valueAsDate = new Date();

            async function loadRaznaryadka() {
                const dateStr = document.getElementById('targetDate').value;
                const region = document.getElementById('regionSelect').value;
                const container = document.getElementById('resultContainer');
                
                if (!dateStr) {
                    alert('Выберите дату');
                    return;
                }

                container.innerHTML = '<div class="text-center py-12 text-gray-500">Загрузка данных из Excel файлов...</div>';

                try {
                    const response = await fetch(`/api/raznaryadka/${dateStr}?region=${region}`);
                    const data = await response.json();

                    if (!data.is_working_day) {
                        container.innerHTML = `
                            <div class="bg-yellow-50 border-l-4 border-yellow-400 p-4 rounded-md shadow">
                                <p class="text-yellow-700 font-medium">Выходной или праздничный день (${data.date})</p>
                                <p class="text-yellow-600 text-sm mt-1">${data.message}</p>
                            </div>
                        `;
                        return;
                    }

                    let html = `
                        <div class="grid grid-cols-1 md:grid-cols-4 gap-6">
                            <div class="bg-white shadow rounded-lg p-6 border-t-4 border-blue-500">
                                <h3 class="text-sm font-medium text-gray-500">Дата</h3>
                                <p class="text-xl font-bold text-gray-800 mt-1">${data.date}</p>
                            </div>
                            <div class="bg-white shadow rounded-lg p-6 border-t-4 border-green-500">
                                <h3 class="text-sm font-medium text-gray-500">Регион</h3>
                                <p class="text-xl font-bold text-gray-800 mt-1">${data.region}</p>
                            </div>
                            <div class="bg-white shadow rounded-lg p-6 border-t-4 border-purple-500">
                                <h3 class="text-sm font-medium text-gray-500">Меню (из Excel)</h3>
                                <p class="text-xl font-bold text-gray-800 mt-1">Меню № ${data.menu_number}</p>
                            </div>
                            <div class="bg-white shadow rounded-lg p-6 border-t-4 border-orange-500">
                                <h3 class="text-sm font-medium text-gray-500">Объектов МТТ</h3>
                                <p class="text-xl font-bold text-orange-600 mt-1">${data.points_count} точек</p>
                            </div>
                        </div>

                        <div class="bg-white shadow rounded-lg overflow-hidden">
                            <div class="bg-gray-50 px-6 py-4 border-b border-gray-200">
                                <h3 class="text-lg font-medium text-gray-900">📦 Сводный список закупки для сектора "${data.region}"</h3>
                            </div>
                            <div class="p-6 overflow-x-auto">
                                <table class="min-w-full divide-y divide-gray-200">
                                    <thead>
                                        <tr>
                                            <th class="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Продукт</th>
                                            <th class="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Общий объем</th>
                                            <th class="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Ед. изм.</th>
                                        </tr>
                                    </thead>
                                    <tbody class="divide-y divide-gray-200">
                    `;

                    for (const [prodName, item] of Object.entries(data.total_purchase_list)) {
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

                        <div class="bg-white shadow rounded-lg overflow-hidden">
                            <div class="bg-gray-50 px-6 py-4 border-b border-gray-200">
                                <h3 class="text-lg font-medium text-gray-900">📍 Детализация по точкам МТТ (${data.region})</h3>
                            </div>
                            <div class="p-6 grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
                    `;

                    data.points_breakdown.forEach(point => {
                        html += `
                            <div class="border border-gray-200 rounded-lg p-4 bg-gray-50 shadow-sm">
                                <div class="flex justify-between items-center mb-3">
                                    <h4 class="font-bold text-gray-800 text-sm">${point.point_name}</h4>
                                    <span class="bg-blue-100 text-blue-800 text-xs font-semibold px-2 py-0.5 rounded">Детей: ${point.children_count}</span>
                                </div>
                                <ul class="divide-y divide-gray-200 text-xs">
                        `;
                        
                        for (const [prodName, need] of Object.entries(point.needs)) {
                            html += `
                                <li class="py-1.5 flex justify-between">
                                    <span class="text-gray-600">${prodName}:</span>
                                    <span class="font-medium text-gray-900">${need.qty} ${need.unit}</span>
                                </li>
                            `;
                        }

                        html += `
                                </ul>
                            </div>
                        `;
                    });

                    html += `
                            </div>
                        </div>
                    `;

                    container.innerHTML = html;

                } catch (error) {
                    container.innerHTML = '<div class="text-center py-12 text-red-500">Ошибка при расчете. Проверьте структуру Excel файлов.</div>';
                    console.error(error);
                }
            }

            window.onload = loadRaznaryadka;
        </script>
    </body>
    </html>
    """
  return HTMLResponse(content=html_content)


@app.get("/api/raznaryadka/{target_date}")
def get_raznaryadka_by_date(target_date: str, region: str = "Город"):
  """Формирует разнарядку с переводом граммов в килограммы на основе Menu.xlsx."""
  try:
    parsed_date = date.fromisoformat(target_date)
  except ValueError:
    raise HTTPException(
        status_code=400, detail="Неверный формат даты. Используйте ГГГГ-ММ-ДД"
    )

  menu_num = get_menu_number_for_date(parsed_date)
  if menu_num is None:
    return {
        "date": parsed_date.isoformat(),
        "is_working_day": False,
        "message": "В этот день выходной или праздник. Разнарядка не нужна.",
    }

  points_dict = load_points_from_excel(region)
  ttk_items = load_menu_from_excel(menu_num)

  points_report = []
  summary_ingredients = {}

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

      if prod_name not in summary_ingredients:
        summary_ingredients[prod_name] = {"total_qty": 0.0, "unit": unit}
      summary_ingredients[prod_name]["total_qty"] += total_needed

    points_report.append({
        "point_id": point_id,
        "point_name": point_info["name"],
        "children_count": children,
        "needs": point_needs,
    })

  summary_rounded = {
      prod: {
          "total_qty": round(info["total_qty"], 2),
          "unit": info["unit"],
      }
      for prod, info in summary_ingredients.items()
  }

  return {
      "date": parsed_date.isoformat(),
      "region": region,
      "is_working_day": True,
      "menu_number": menu_num,
      "points_count": len(points_dict),
      "points_breakdown": points_report,
      "total_purchase_list": summary_rounded,
  }
