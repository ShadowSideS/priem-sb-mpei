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
.wrap {{ max-width: 1220px; margin: 36px auto; padding: 0 18px; }}
.card {{ background: #fff; border-radius: 14px; padding: 24px; box-shadow: 0 4px 18px rgba(0,0,0,.08); }}
h1 {{ margin: 0 0 12px; font-size: 28px; }}
.meta {{ color: #4b5563; margin-bottom: 18px; line-height: 1.6; }}
.actions {{ margin-bottom: 18px; }}
.button {{ display: inline-block; border-radius: 8px; padding: 11px 16px; font-weight: 700; text-decoration: none; background: #475569; color: #fff; }}
.table-wrap {{ overflow-x: auto; }}
table {{ width: 100%; border-collapse: collapse; background: #fff; }}
th, td {{ border: 1px solid #cbd5e1; padding: 10px 12px; text-align: center; }}
th {{ background: #d9eaf7; font-weight: 700; }}
th:first-child, td:first-child {{ text-align: left; }}
tr.total td {{ background: #fff2cc; font-weight: 700; }}
.note {{ margin-top: 14px; color: #64748b; font-size: 13px; line-height: 1.5; }}
@media (max-width: 700px) {{
  .wrap {{ margin: 12px auto; padding: 0 8px; }}
  .card {{ padding: 14px; border-radius: 10px; }}
  h1 {{ font-size: 22px; }}
  th, td {{ padding: 8px; font-size: 14px; }}
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
<th>Основной конкурс</th>
<th>Переполнение</th>
<th>Минимальный балл основного конкурса</th>
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
