from datetime import date
from fastapi import FastAPI, HTTPException
from fastapi.responses import HTMLResponse
from calendar_logic import get_menu_number_for_date
from models_and_data import POINTS, PRODUCTS, MENUS_TTK

app = FastAPI(title="Catering Management API", version="1.0.0")


@app.get("/", response_class=HTMLResponse)
def get_dashboard():
  """Красивый веб-интерфейс для логиста (дашборд разнарядки)."""
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
            <!-- Шапка -->
            <header class="bg-white shadow rounded-lg p-6 mb-8 flex flex-col md:flex-row justify-between items-center">
                <div>
                    <h1 class="text-2xl font-bold text-blue-600">🍽 Управление Разнарядкой Кетеринга</h1>
                    <p class="text-sm text-gray-500 mt-1">Автоматический расчет продуктов для точек по меню и количеству детей</p>
                </div>
                <div class="mt-4 md:mt-0 flex items-center space-x-3">
                    <label class="text-sm font-medium text-gray-700">Выберите дату:</label>
                    <input type="date" id="targetDate" class="border border-gray-300 rounded-md px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500">
                    <button onclick="loadRaznaryadka()" class="bg-blue-600 text-white px-4 py-2 rounded-md text-sm font-medium hover:bg-blue-700 transition">Сформировать</button>
                </div>
            </header>

            <!-- Основной контент -->
            <div id="resultContainer" class="space-y-8">
                <div class="bg-white shadow rounded-lg p-8 text-center text-gray-400">
                    Выберите дату и нажмите «Сформировать», чтобы увидеть разнарядку.
                </div>
            </div>
        </div>

        <script>
            // Устанавливаем сегодняшнюю дату по умолчанию
            document.getElementById('targetDate').valueAsDate = new Date();

            async function loadRaznaryadka() {
                const dateStr = document.getElementById('targetDate').value;
                const container = document.getElementById('resultContainer');
                
                if (!dateStr) {
                    alert('Пожалуйста, выберите дату');
                    return;
                }

                container.innerHTML = '<div class="text-center py-12 text-gray-500">Загрузка данных...</div>';

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

                    // Рендерим успешный отчет
                    let html = `
                        <div class="grid grid-cols-1 md:grid-cols-3 gap-6">
                            <div class="bg-white shadow rounded-lg p-6 border-t-4 border-blue-500">
                                <h3 class="text-sm font-medium text-gray-500">Дата доставки</h3>
                                <p class="text-2xl font-bold text-gray-800 mt-1">${data.date}</p>
                            </div>
                            <div class="bg-white shadow rounded-lg p-6 border-t-4 border-green-500">
                                <h3 class="text-sm font-medium text-gray-500">Активное меню</h3>
                                <p class="text-2xl font-bold text-gray-800 mt-1">Меню № ${data.menu_number}</p>
                            </div>
                            <div class="bg-white shadow rounded-lg p-6 border-t-4 border-purple-500">
                                <h3 class="text-sm font-medium text-gray-500">Статус дня</h3>
                                <p class="text-2xl font-bold text-green-600 mt-1">Рабочий день</p>
                            </div>
                        </div>

                        <!-- Сводный список для закупа -->
                        <div class="bg-white shadow rounded-lg overflow-hidden">
                            <div class="bg-gray-50 px-6 py-4 border-b border-gray-200">
                                <h3 class="text-lg font-medium text-gray-900">📦 Сводный список для закупки (на все точки)</h3>
                            </div>
                            <div class="p-6 overflow-x-auto">
                                <table class="min-w-full divide-y divide-gray-200">
                                    <thead>
                                        <tr>
                                            <th class="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Продукт</th>
                                            <th class="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Общее количество</th>
                                            <th class="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Ед. изм.</th>
                                        </tr>
                                    </thead>
                                    <tbody class="divide-y divide-gray-200">
                    `;

                    for (const [code, item] of Object.entries(data.total_purchase_list)) {
                        html += `
                            <tr>
                                <td class="px-6 py-4 whitespace-nowrap text-sm font-medium text-gray-900">${item.product_name}</td>
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

                        <!-- Детализация по точкам -->
                        <div class="bg-white shadow rounded-lg overflow-hidden">
                            <div class="bg-gray-50 px-6 py-4 border-b border-gray-200">
                                <h3 class="text-lg font-medium text-gray-900">📍 Детализация по точкам доставки</h3>
                            </div>
                            <div class="p-6 grid grid-cols-1 md:grid-cols-2 gap-6">
                    `;

                    data.points_breakdown.forEach(point => {
                        html += `
                            <div class="border border-gray-200 rounded-lg p-4 bg-gray-50">
                                <div class="flex justify-between items-center mb-3">
                                    <h4 class="font-bold text-gray-800">${point.point_name}</h4>
                                    <span class="bg-blue-100 text-blue-800 text-xs font-semibold px-2.5 py-0.5 rounded">Детей: ${point.children_count}</span>
                                </div>
                                <ul class="divide-y divide-gray-200 text-sm">
                        `;
                        
                        for (const [prodCode, need] of Object.entries(point.needs)) {
                            html += `
                                <li class="py-1.5 flex justify-between">
                                    <span class="text-gray-600">${need.product_name}:</span>
                                    <span class="font-medium text-gray-900">${need.required_qty} ${need.unit}</span>
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
                    container.innerHTML = '<div class="text-center py-12 text-red-500">Ошибка при загрузке данных с сервера.</div>';
                    console.error(error);
                }
            }

            // Автоматически загружаем данные на сегодня при открытии
            window.onload = loadRaznaryadka;
        </script>
    </body>
    </html>
    """
  return HTMLResponse(content=html_content)


@app.get("/api/raznaryadka/{target_date}")
def get_raznaryadka_by_date(target_date: str):
  """Формирует разнарядку продуктов по всем точкам на конкретную дату (ГГГГ-ММ-ДД)."""
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

  ttk_items = MENUS_TTK.get(menu_num, {})

  points_report = []
  summary_ingredients = {}

  for point_id, point_info in POINTS.items():
    children = point_info["children_count"]
    point_needs = {}

    for prod_code, norm_per_child in ttk_items.items():
      total_needed = norm_per_child * children
      point_needs[prod_code] = {
          "product_name": PRODUCTS[prod_code]["name"],
          "unit": PRODUCTS[prod_code]["unit"],
          "required_qty": round(total_needed, 2),
      }

      if prod_code not in summary_ingredients:
        summary_ingredients[prod_code] = {
            "product_name": PRODUCTS[prod_code]["name"],
            "unit": PRODUCTS[prod_code]["unit"],
            "total_qty": 0.0,
        }
      summary_ingredients[prod_code]["total_qty"] += total_needed

    points_report.append({
        "point_id": point_id,
        "point_name": point_info["name"],
        "children_count": children,
        "needs": point_needs,
    })

  for prod_code in summary_ingredients:
    summary_ingredients[prod_code]["total_qty"] = round(
        summary_ingredients[prod_code]["total_qty"], 2
    )

  return {
      "date": parsed_date.isoformat(),
      "is_working_day": True,
      "menu_number": menu_num,
      "points_breakdown": points_report,
      "total_purchase_list": summary_ingredients,
  }
