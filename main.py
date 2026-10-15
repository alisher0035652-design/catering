from datetime import date
from fastapi import FastAPI, HTTPException
from fastapi.responses import HTMLResponse
import pandas as pd
from calendar_logic import get_menu_number_for_date

app = FastAPI(title="Catering Management API", version="1.0.0")

# Справочник 24 точек (можно позже тоже вынести в базу или Excel, пока зафиксируем здесь)
POINTS = {
    1: {"name": "Школа №1", "children_count": 300},
    2: {"name": "Детсад №12", "children_count": 150},
    3: {"name": "Школа №3", "children_count": 400},
    4: {"name": "Детсад №4", "children_count": 220},
    5: {"name": "Школа №5", "children_count": 350},
    6: {"name": "Детсад №6", "children_count": 180},
    7: {"name": "Школа №7", "children_count": 500},
    8: {"name": "Детсад №8", "children_count": 210},
    9: {"name": "Школа №9", "children_count": 320},
    10: {"name": "Детсад №10", "children_count": 160},
    11: {"name": "Школа №11", "children_count": 290},
    12: {"name": "Детсад №12-б", "children_count": 190},
    13: {"name": "Школа №13", "children_count": 370},
    14: {"name": "Детсад №14", "children_count": 200},
    15: {"name": "Школа №15", "children_count": 410},
    16: {"name": "Детсад №16", "children_count": 175},
    17: {"name": "Школа №17", "children_count": 340},
    18: {"name": "Детсад №18", "children_count": 230},
    19: {"name": "Школа №19", "children_count": 380},
    20: {"name": "Детсад №20", "children_count": 195},
    21: {"name": "Школа №21", "children_count": 450},
    22: {"name": "Детсад №22", "children_count": 215},
    23: {"name": "Школа №23", "children_count": 310},
    24: {"name": "Детсад №24", "children_count": 240},
}


def load_menu_from_excel(menu_num: int) -> dict:
  """Загружает техкарту из соответствующего листа Excel-файла Menu.xlsx.

  Листы названы: '1-Меню', '2-Меню' и т.д. В файле нормы даны в граммах на 1
  ребенка. Переводим в килограммы/литры (делим на 1000).
  """
  sheet_name = f"{menu_num}-Меню"
  try:
    df = pd.read_excel("Menu.xlsx", sheet_name=sheet_name)
    menu_items = {}
    for _, row in df.iterrows():
      prod_name = str(row.iloc[0]).strip()
      grams_per_child = float(row.iloc[1]) if pd.notna(row.iloc[1]) else 0.0
      # Переводим граммы в кг/л для расхода
      menu_items[prod_name] = grams_per_child / 1000.0
    return menu_items
  except Exception as e:
    print(f"Ошибка чтения листа {sheet_name}: {e}")
    return {}


@app.get("/", response_class=HTMLResponse)
def get_dashboard():
  """Интерфейс логиста с поддержкой загрузки данных из Excel."""
  html_content = """
    <!DOCTYPE html>
    <html lang="ru">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Разнарядка Кетеринг (Excel Menu)</title>
        <script src="https://cdn.tailwindcss.com"></script>
    </head>
    <body class="bg-gray-50 text-gray-800 font-sans">
        <div class="max-w-7xl mx-auto px-4 py-8">
            <header class="bg-white shadow rounded-lg p-6 mb-8 flex flex-col md:flex-row justify-between items-center">
                <div>
                    <h1 class="text-2xl font-bold text-blue-600">🍽 Управление Разнарядкой Кетеринга</h1>
                    <p class="text-sm text-gray-500 mt-1">Данные меню загружаются автоматически из файла Menu.xlsx (10 листов)</p>
                </div>
                <div class="mt-4 md:mt-0 flex items-center space-x-3">
                    <label class="text-sm font-medium text-gray-700">Дата доставки:</label>
                    <input type="date" id="targetDate" class="border border-gray-300 rounded-md px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500">
                    <button onclick="loadRaznaryadka()" class="bg-blue-600 text-white px-4 py-2 rounded-md text-sm font-medium hover:bg-blue-700 transition">Сформировать</button>
                </div>
            </header>

            <div id="resultContainer" class="space-y-8">
                <div class="bg-white shadow rounded-lg p-8 text-center text-gray-400">
                    Выберите дату и нажмите «Сформировать».
                </div>
            </div>
        </div>

        <script>
            document.getElementById('targetDate').valueAsDate = new Date();

            async function loadRaznaryadka() {
                const dateStr = document.getElementById('targetDate').value;
                const container = document.getElementById('resultContainer');
                
                if (!dateStr) {
                    alert('Выберите дату');
                    return;
                }

                container.innerHTML = '<div class="text-center py-12 text-gray-500">Считывание Excel меню и расчет по 24 точкам...</div>';

                try {
                    const response = await fetch(`/api/raznaryadka/${dateStr}`);
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
                        <div class="grid grid-cols-1 md:grid-cols-3 gap-6">
                            <div class="bg-white shadow rounded-lg p-6 border-t-4 border-blue-500">
                                <h3 class="text-sm font-medium text-gray-500">Дата</h3>
                                <p class="text-2xl font-bold text-gray-800 mt-1">${data.date}</p>
                            </div>
                            <div class="bg-white shadow rounded-lg p-6 border-t-4 border-green-500">
                                <h3 class="text-sm font-medium text-gray-500">Активное меню (из Excel)</h3>
                                <p class="text-2xl font-bold text-gray-800 mt-1">Меню № ${data.menu_number}</p>
                            </div>
                            <div class="bg-white shadow rounded-lg p-6 border-t-4 border-purple-500">
                                <h3 class="text-sm font-medium text-gray-500">Всего точек</h3>
                                <p class="text-2xl font-bold text-blue-600 mt-1">24 объекта</p>
                            </div>
                        </div>

                        <div class="bg-white shadow rounded-lg overflow-hidden">
                            <div class="bg-gray-50 px-6 py-4 border-b border-gray-200">
                                <h3 class="text-lg font-medium text-gray-900">📦 Сводный список закупки (кг / литры)</h3>
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

                    for (const [prodName, totalQty] of Object.entries(data.total_purchase_list)) {
                        html += `
                            <tr>
                                <td class="px-6 py-4 whitespace-nowrap text-sm font-medium text-gray-900">${prodName}</td>
                                <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-700 font-bold">${totalQty}</td>
                                <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-500">кг / л</td>
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
                                <h3 class="text-lg font-medium text-gray-900">📍 Детализация по 24 точкам</h3>
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
                        
                        for (const [prodName, qty] of Object.entries(point.needs)) {
                            html += `
                                <li class="py-1.5 flex justify-between">
                                    <span class="text-gray-600">${prodName}:</span>
                                    <span class="font-medium text-gray-900">${qty} кг</span>
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
                    container.innerHTML = '<div class="text-center py-12 text-red-500">Ошибка при расчете данных. Проверьте структуру файла Menu.xlsx.</div>';
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
def get_raznaryadka_by_date(target_date: str):
  """Формирует разнарядку, считывая техкарту из нужного листа Menu.xlsx."""
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

  # Загружаем граммовки из Excel для текущего меню (1-10)
  ttk_items = load_menu_from_excel(menu_num)

  points_report = []
  summary_ingredients = {}

  for point_id, point_info in POINTS.items():
    children = point_info["children_count"]
    point_needs = {}

    for prod_name, kg_per_child in ttk_items.items():
      total_needed = kg_per_child * children
      point_needs[prod_name] = round(total_needed, 2)

      if prod_name not in summary_ingredients:
        summary_ingredients[prod_name] = 0.0
      summary_ingredients[prod_name] += total_needed

    points_report.append({
        "point_id": point_id,
        "point_name": point_info["name"],
        "children_count": children,
        "needs": point_needs,
    })

  # Округляем общие итоги
  summary_rounded = {
      prod: round(qty, 2) for prod, qty in summary_ingredients.items()
  }

  return {
      "date": parsed_date.isoformat(),
      "is_working_day": True,
      "menu_number": menu_num,
      "points_breakdown": points_report,
      "total_purchase_list": summary_rounded,
  }
