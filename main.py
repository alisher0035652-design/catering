from datetime import date
from fastapi import FastAPI, HTTPException
from calendar_logic import get_menu_number_for_date

app = FastAPI(title="Catering Management API", version="1.0.0")


@app.get("/")
def root():
  return {"message": "Catering ERP API is running!"}


@app.get("/api/menu/today")
def get_today_menu():
  """Возвращает номер меню на сегодня."""
  today = date.today()
  menu_num = get_menu_number_for_date(today)

  if menu_num is None:
    return {
        "date": today.isoformat(),
        "is_working_day": False,
        "message": "Сегодня выходной или праздничный день. Доставки нет.",
    }

  return {
      "date": today.isoformat(),
      "is_working_day": True,
      "menu_number": menu_num,
  }


@app.get("/api/menu/{target_date}")
def get_menu_by_date(target_date: str):
  """Возвращает номер меню на конкретную дату (в формате ГГГГ-ММ-ДД)."""
  try:
    parsed_date = date.fromisoformat(target_date)
  except ValueError:
    raise HTTPException(
        status_code=400,
        detail="Неверный формат даты. Используйте ГГГГ-ММ-ДД (например,"
        " 2026-10-16)",
    )

  try:
    menu_num = get_menu_number_for_date(parsed_date)
  except ValueError as e:
    raise HTTPException(status_code=400, detail=str(e))

  if menu_num is None:
    return {
        "date": parsed_date.isoformat(),
        "is_working_day": False,
        "message": "В указанную дату выходной или праздник.",
    }

  return {
      "date": parsed_date.isoformat(),
      "is_working_day": True,
      "menu_number": menu_num,
  }
