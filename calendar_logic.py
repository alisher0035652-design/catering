from datetime import date, timedelta

# Пример праздничных дней
DEFAULT_HOLIDAYS = {
    date(2026, 1, 1),
    date(2026, 3, 21),  # Навруз
    date(2026, 9, 1),
}


def get_menu_number_for_date(
    target_date: date,
    anchor_date: date = date(2026, 1, 5),
    anchor_menu: int = 1,
    holidays: set = DEFAULT_HOLIDAYS,
) -> int | None:
  """Рассчитывает номер меню (от 1 до 10) для конкретной даты.

  Пропускает выходные (суббота, воскресенье) и праздники.
  """
  if target_date.weekday() >= 5 or target_date in holidays:
    return None  # Выходной день

  if target_date < anchor_date:
    raise ValueError("Целевая дата не может быть раньше даты старта системы.")

  current = anchor_date
  work_days_passed = 0

  while current < target_date:
    current += timedelta(days=1)
    if current.weekday() < 5 and current not in holidays:
      work_days_passed += 1

  menu_number = ((anchor_menu - 1 + work_days_passed) % 10) + 1
  return menu_number
