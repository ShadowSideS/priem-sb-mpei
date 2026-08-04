import asyncio
import html
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import pandas as pd
from playwright.async_api import async_playwright

from parser import build_summary, collect_main_links, format_excel, read_main_list

ROOT = Path(__file__).resolve().parent
SITE_DIR = ROOT / "site"
EXCEL_FILE = SITE_DIR / "СФ_МЭИ_Сводка_для_руководства.xlsx"
TIMEZONE = ZoneInfo("Europe/Moscow")


async def collect_summary() -> pd.DataFrame:
    main_records: list[dict] = []

    async with async_playwright() as playwright:
        browser = await playwright.chromium.launch(headless=True)
        context = await browser.new_context(
            viewport={"width": 1920, "height": 1080},
            locale="ru-RU",
        )
        page = await context.new_page()

        try:
            links = await collect_main_links(page)
            if not links:
                raise RuntimeError(
                    "На сайте не найдено ни одного направления общего конкурса."
                )

            print(f"Найдено направлений: {len(links)}")
            for index, item in enumerate(links, start=1):
                print(f"[{index}/{len(links)}] {item['Направление']}")
                data = await read_main_list(page, item["Ссылка"])
                main_records.append({**item, **data})
        finally:
            await browser.close()

    if not main_records:
        raise RuntimeError("Данные общего конкурса не получены.")

    return build_summary(main_records)


def save_excel(summary: pd.DataFrame) -> None:
    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="Сводка", index=False)
    format_excel(EXCEL_FILE)


def render_html(summary: pd.DataFrame, updated_at: datetime) -> str:
    rows = summary.where(pd.notnull(summary), None).to_dict(orient="records")
    rows_html: list[str] = []

    for row in rows:
        is_total = row.get("Направление") == "ИТОГО"
        row_class = ' class="total"' if is_total else ""
        values = [
            row.get("Направление"),
            row.get("Основной конкурс"),
            row.get("Переполнение"),
            row.get("Минимальный балл основного конкурса"),
        ]
        cells = "".join(
            f"<td>{html.escape('' if value is None else str(value))}</td>"
            for value in values
        )
        rows_html.append(f"<tr{row_class}>{cells}</tr>")

    updated_text = updated_at.strftime("%d.%m.%Y %H:%M:%S МСК")

    return f"""<!doctype html>
<html lang="ru">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta http-equiv="refresh" content="300">
<title>СФ МЭИ — сводка общего конкурса</title>
<style>
:root {{ color-scheme: light; }}
* {{ box-sizing: border-box; }}
body {{ font-family: Arial, sans-serif; margin: 0; background: #f4f6f8; color: #1f2937; }}
.wrap {{ width: 100%; max-width: 1400px; margin: 18px auto; padding: 0 10px; }}
.card {{ background: #fff; border-radius: 12px; padding: 18px; box-shadow: 0 3px 14px rgba(0,0,0,.07); }}
h1 {{ margin: 0 0 8px; font-size: 25px; line-height: 1.15; }}
.meta {{ color: #4b5563; margin-bottom: 12px; line-height: 1.35; font-size: 14px; }}
.actions {{ margin-bottom: 12px; }}
.button {{ display: inline-block; border-radius: 7px; padding: 9px 13px; font-size: 14px; font-weight: 700; text-decoration: none; background: #475569; color: #fff; }}
.table-wrap {{ width: 100%; overflow-x: auto; }}
table {{ width: 100%; table-layout: fixed; border-collapse: collapse; background: #fff; }}
th, td {{ border: 1px solid #cbd5e1; padding: 8px 9px; text-align: center; line-height: 1.2; overflow-wrap: anywhere; }}
th {{ background: #d9eaf7; font-weight: 700; }}
th:nth-child(1), td:nth-child(1) {{ width: 56%; text-align: left; }}
th:nth-child(2), td:nth-child(2) {{ width: 13%; }}
th:nth-child(3), td:nth-child(3) {{ width: 13%; }}
th:nth-child(4), td:nth-child(4) {{ width: 18%; }}
tr.total td {{ background: #fff2cc; font-weight: 700; }}
.note {{ margin-top: 10px; color: #64748b; font-size: 12px; line-height: 1.35; }}
.mobile-label {{ display: none; }}
@media (max-width: 700px) {{
  body {{ background: #fff; }}
  .wrap {{ margin: 0; padding: 0; max-width: none; }}
  .card {{ padding: 8px 5px 10px; border-radius: 0; box-shadow: none; }}
  h1 {{ margin-bottom: 5px; font-size: 17px; }}
  .meta {{ margin-bottom: 7px; font-size: 10.5px; line-height: 1.25; }}
  .actions {{ margin-bottom: 7px; }}
  .button {{ padding: 6px 9px; border-radius: 5px; font-size: 11px; }}
  .table-wrap {{ overflow-x: visible; }}
  table {{ font-size: 10px; }}
  th, td {{ padding: 4px 3px; line-height: 1.12; }}
  th:nth-child(1), td:nth-child(1) {{ width: 55%; }}
  th:nth-child(2), td:nth-child(2) {{ width: 14%; }}
  th:nth-child(3), td:nth-child(3) {{ width: 14%; }}
  th:nth-child(4), td:nth-child(4) {{ width: 17%; }}
  .desktop-label {{ display: none; }}
  .mobile-label {{ display: inline; }}
  .note {{ margin-top: 7px; font-size: 9.5px; }}
}}
@media (max-width: 390px) {{
  .card {{ padding-left: 3px; padding-right: 3px; }}
  h1 {{ font-size: 16px; }}
  table {{ font-size: 9px; }}
  th, td {{ padding: 3px 2px; }}
}}
</style>
</head>
<body>
<div class="wrap"><div class="card">
<h1>СФ МЭИ — сводка общего конкурса</h1>
<div class="meta">
Последнее успешное обновление: <strong>{html.escape(updated_text)}</strong><br>
Обновление выполняется автоматически примерно каждые 15 минут.
</div>
<div class="actions">
<a class="button" href="./СФ_МЭИ_Сводка_для_руководства.xlsx" download>Скачать Excel</a>
</div>
<div class="table-wrap"><table>
<thead><tr>
<th>Направление</th>
<th><span class="desktop-label">Основной конкурс</span><span class="mobile-label">Основной</span></th>
<th><span class="desktop-label">Переполнение</span><span class="mobile-label">Переполн.</span></th>
<th><span class="desktop-label">Минимальный балл основного конкурса</span><span class="mobile-label">Мин. балл</span></th>
</tr></thead>
<tbody>{''.join(rows_html)}</tbody>
</table></div>
<div class="note">
При ошибке сбора новая версия не публикуется, поэтому на странице остаётся последняя успешно сформированная сводка.
Страница автоматически проверяет обновление раз в 5 минут.
</div>
</div></div>
</body>
</html>"""


async def main() -> None:
    SITE_DIR.mkdir(parents=True, exist_ok=True)
    summary = await collect_summary()

    # Защита от публикации явно пустого или повреждённого результата.
    if summary.empty or "ИТОГО" not in set(summary["Направление"].astype(str)):
        raise RuntimeError("Итоговая сводка пуста или не содержит строку ИТОГО.")

    save_excel(summary)
    updated_at = datetime.now(TIMEZONE)
    (SITE_DIR / "index.html").write_text(
        render_html(summary, updated_at),
        encoding="utf-8",
    )
    (SITE_DIR / ".nojekyll").write_text("", encoding="utf-8")
    print(f"Готово: {SITE_DIR}")


if __name__ == "__main__":
    asyncio.run(main())
