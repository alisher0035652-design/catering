from datetime import date
from fastapi import FastAPI, HTTPException
from calendar_logic import get_menu_number_for_date
from models_and_data import POINTS, PRODUCTS, MENUS_TTK

app = FastAPI(title="Catering Management API", version="1.0.0")


@app.get("/")
def root():
  return {"message": "Catering ERP API is running!"}


@app.get("/api/raznaryadka/{target_date}")
def get_raznaryadka_by_date(target_date: str):
  """Формирует разнарядку продуктов по всем точкам на конкретную дату (ГГГГ-ММ-ДД).

  Учитывает: - Рабочий день или выходной (через календарь меню) - Номер меню (от
  1 до 10) - Количество детей на каждой точке - Нормы продуктов из ТТК
  """
  try:
    parsed_date = date.fromisoformat(target_date)
  except ValueError:
    raise HTTPException(
        status_code=400, detail="Неверный формат даты. Используйте ГГГГ-ММ-ДД"
    )

  # 1. Узнаем номер меню на этот день
  menu_num = get_menu_number_for_date(parsed_date)
  if menu_num is None:
    return {
        "date": parsed_date.isoformat(),
        "is_working_day": False,
        "message": "В этот день выходной или праздник. Разнарядка не нужна.",
    }

  # 2. Получаем ТТК для этого меню
  ttk_items = MENUS_TTK.get(menu_num, {})

  # 3. Считаем потребность по каждой точке
  points_report = []
  summary_ingredients = {}  г  # Общий итог закупок по всем точкам

  for point_id, point_info in POINTS.items():
    children = point_info["children_count"]
    point_needs = {}

    for prod_code, norm_per_child in ttk_items.items():
      total_needed = norm_per_child * children  г  # расход на точку
      point_needs[prod_code] = {
          "product_name": PRODUCTS[prod_code]["name"],
          "unit": PRODUCTS[prod_code]["unit"],
          "required_qty": round(total_needed, 2),
      }

      # Суммируем в общий итог для закупщика/водителей
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

  # Округляем общие итоги
  for prod_code in summary_ingredients:
    summary_ingredients[prod_code]["total_qty"] = round(
        summary_ingredients[prod_code]["total_qty"], 2
    )

  return {
      "date": parsed_date.isoformat(),
      "is_working_day": True,
      "menu_number": menu_num,
      "points_breakdown": points_report,  г  # Детализация по каждой из 24 точек
      "total_purchase_list": summary_ingredients,  г  # Сводный список для закупа
  }
