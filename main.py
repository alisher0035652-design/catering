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
  """Дашборд логиста в матричном виде (Продукты х МТТ)."""
  html_content = """
    <!DOCTYPE html>
    <html lang="ru">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Разнарядка Кетеринг (Матрица)</title>
        <script src="https://cdn.tailwindcss.com"></script>
    </head>
    <body class="bg-gray-50 text-gray-800 font-sans">
        <div class="max-w-[95%] mx-auto px-4 py-8">
            <header class="bg-white shadow rounded-lg p-6 mb-8 flex flex-col md:flex-row justify-between items-center">
                <div>
                    <h1 class="text-2xl font-bold text-blue-600">🍽 Матричная Разнарядка Кетеринг</h1>
                    <p class="text-sm text-gray-500 mt-1">Продукты по строкам, точки МТТ по столбцам на всю неделю (Пн-Пт)</p>
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
                        <button onclick="loadMatrixRaznaryadka()" class="bg-blue-600 text-white px-4 py-2 rounded-md text-sm font-medium hover:bg-blue-700 transition">Сформировать матрицу</button>
                    </div>
                </div>
            </header>

            <div id="resultContainer" class="space-y-8">
                <div class="bg-white shadow rounded-lg p-8 text-center text-gray-400">
                    Выберите дату, регион и нажмите «Сформировать матрицу».
                </div>
            </div>
        </div>

        <script>
            document.getElementById('targetDate').valueAsDate = new Date();

            async function loadMatrixRaznaryadka() {
                const dateStr = document.getElementById('targetDate').value;
                const region = document.getElementById('regionSelect').value;
                const container = document.getElementById('resultContainer');
                
                if (!dateStr) {
                    alert('Выберите дату');
                    return;
                }

                container.innerHTML = '<div class="text-center py-12 text-gray-500">Формирование матричной разнарядки на неделю...</div>';

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
                                <p class="text-lg font-bold text-gray-800 mt-1">${data.region} (${data.points.length} точек МТТ)</p>
                            </div>
                            <div class="bg-white shadow rounded-lg p-6 border-t-4 border-purple-500">
                                <h3 class="text-sm font-medium text-gray-500">Рабочих дней</h3>
                                <p class="text-lg font-bold text-purple-600 mt-1">5 дней (Пн-Пт)</p>
                            </div>
                        </div>

                        <!-- Матричная таблица -->
                        <div class="bg-white shadow rounded-lg overflow-hidden">
                            <div class="bg-gray-50 px-6 py-4 border-b border-gray-200 flex justify-between items-center">
                                <h3 class="text-lg font-medium text-gray-900">📊 Сводная матрица закупки на неделю (Продукты × Точки МТТ)</h3>
                                <span class="text-xs text-gray-500">Объемы указаны в кг / шт</span>
                            </div>
                            <div class="overflow-x-auto" style="max-height: 700px;">
                                <table class="min-w-full divide-y divide-gray-200 text-xs border-collapse">
                                    <thead class="bg-gray-100 sticky top-0 z-10">
                                        <tr>
                                            <th class="px-4 py-3 text-left font-bold text-gray-700 border-r bg-gray-100 sticky left-0 z-20 min-w-[180px]">Продукт</th>
                                            <th class="px-3 py-3 text-center font-bold text-gray-700 border-r bg-gray-100 min-w-[70px]">Ед.</th>
                    `;

                    // Заголовки столбцов - названия точек МТТ
                    data.points.forEach(point => {
                        html += `<th class="px-3 py-3 text-center font-medium text-gray-700 border-r min-w-[90px]" title="Детей: ${point.children_count}">${point.name}<br><span class="text-[10px] text-blue-600">(${point.children_count} дет)</span></th>`;
                    });

                    html += `
                                            <th class="px-4 py-3 text-center font-bold text-blue-800 bg-blue-50">ИТОГО ЗА НЕДЕЛЮ</th>
                                        </tr>
                                    </thead>
                                    <tbody class="divide-y divide-gray-200 bg-white">
                    `;

                    // Строки матрицы по каждому продукту
                    for (const [prodName, rowData] of Object.entries(data.matrix)) {
                        html += `
                            hover:bg-gray-50
                            <tr>
                                <td class="px-4 py-3 font-medium text-gray-900 border-r bg-white sticky left-0 z-10">${prodName}</td>
                                <td class="px-3 py-3 text-center text-gray-500 border-r bg-gray-50">${rowData.unit}</td>
                        `;

                        // Значения по каждой точке МТТ
                        data.points.forEach(point => {
                            const val = rowData.points_qty[point.id] || 0;
                            html += `<td class="px-3 py-3 text-center border-r text-gray-700 ${val > 0 ? 'font-semibold' : 'text-gray-300'}">${val > 0 ? val : '-'}</td>`;
                        });

                        // Итоговая колонка по продукту
                        html += `
                                <td class="px-4 py-3 text-center font-bold text-blue-900 bg-blue-50">${rowData.weekly_total}</td>
                            </tr>
                        `;
                    }

                    html += `
                                    </tbody>
                                </table>
                            </div>
                        </div>
                    `;

                    container.innerHTML = html;

                } catch (error) {
                    container.innerHTML = '<div class="text-center py-12 text-red-500">Ошибка при формировании матрицы.</div>';
                    console.error(error);
                }
            }

            window.onload = loadMatrixRaznaryadka;
        </script>
    </body>
    </html>
    """
  return HTMLResponse(content=html_content)


@app.get("/api/raznaryadka/matrix/{target_date}")
def get_matrix_raznaryadka(target_date: str, region: str = "Город"):
  """Гформирует матрицу: Продукты в строках, точки МТТ в столбцах (суммарно за

  неделю).
  """
  try:
    parsed_date = date.fromisoformat(target_date)
  except ValueError:
    raise HTTPException(
        status_code=400, detail="Неверный формат даты. Используйте ГГГГ-ММ-ДД"
    )

  # Пн-Пт текущей недели
  start_of_week = parsed_date - timedelta(days=parsed_date.weekday())
  week_days = [start_of_week + timedelta(days=i) for i in range(5)]

  points_dict = load_points_from_excel(region)

  # Подготавливаем структуру точек для ответа
  points_list = [
      {
          "id": pid,
          "name": pinfo["name"],
          "children_count": pinfo["children_count"],
      }
      for pid, pinfo in points_dict.items()
  ]

  # Собираем данные за все рабочие дни недели
  matrix_data = {}  # { prod_name: { unit: 'кг', points_qty: {point_id: qty}, weekly_total: 0 } }

  for day in week_days:
    menu_num = get_menu_number_for_date(day)
    if menu_num is None:
      continue

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

  # Округляем значения в матрице
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
      "points": points_list,
      "matrix": matrix_data,
  }
