import os
import calendar
import re
import io
import sqlite3
import logging
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd
from telegram import Update, BotCommand
from telegram.ext import (
    Application, CommandHandler, MessageHandler, ConversationHandler,
    ContextTypes, filters
)

logging.basicConfig(
    format="%(asctime)s %(levelname)s %(name)s: %(message)s",
    level=logging.INFO
)
logging.getLogger("httpx").setLevel(logging.WARNING)
log = logging.getLogger(__name__)

TOKEN = os.getenv("BOT_TOKEN")
OWNER_ID = os.getenv("OWNER_ID")
DB_PATH = os.getenv("DB_PATH", "/data/analytics.sqlite3")
DATE_WAIT = 1

STORE_ORDER = [
    "Буча-Рітейл Парк", "Київ-Аладдін", "Київ-Аркадія",
    "Київ-Блокбастер", "Київ-Дрімтаун_new", "Київ-Квадрат",
    "Київ-Комод", "Київ-Лавіна Молл", "Київ-Материк",
    "Київ-Нью Вей", "Київ-Олімпійський", "Київ-Республіка Парк",
    "Київ-Ретровіль", "Київ-Хрещатик 21"
]

ALIASES = {
    "store": ["Склад", "Магазин", "ТТ"],
    "sales": ["Продажи", "Продажі", "Продажи $", "Продажі $", "Продажи, $", "Продажі, $"],
    "py": ["ПродажиPY", "ПродажіPY", "Продажи PY", "Продажі PY"],
    "direct": ["Прямые Продажи", "Прямі Продажі", "Прямые продажи"],
    "pickup": ["Пикап", "Пікап"],
    "services": ["Услуги, $", "Послуги, $", "Услуги", "Послуги"],
    "accessories": ["Акс, $", "Аксесуари, $", "Акс", "Аксесуари"],
    "plan": ["План", "План $", "ПланДень"]
}

def norm(value):
    return re.sub(r"\s+", " ", str(value).strip().lower().replace("\n", " "))

def find_col(df, aliases, required=True):
    lookup = {norm(c): c for c in df.columns}
    for alias in aliases:
        if norm(alias) in lookup:
            return lookup[norm(alias)]
    if required:
        raise ValueError("Не знайшов колонку: " + " / ".join(aliases))
    return None

def to_num(series):
    return pd.to_numeric(
        series.astype(str).str.replace(r"\s+", "", regex=True)
              .str.replace(",", ".", regex=False)
              .str.replace("$", "", regex=False),
        errors="coerce"
    ).fillna(0)

def read_report(raw: bytes, filename: str) -> pd.DataFrame:
    suffix = Path(filename).suffix.lower()
    if suffix in (".xlsx", ".xls"):
        df = pd.read_excel(io.BytesIO(raw), header=None)
    elif suffix == ".csv":
        try:
            df = pd.read_csv(io.BytesIO(raw), sep=None, engine="python", header=None, encoding="utf-8-sig")
        except UnicodeDecodeError:
            df = pd.read_csv(io.BytesIO(raw), sep=None, engine="python", header=None, encoding="cp1251")
    else:
        raise ValueError("Підтримуються лише XLSX, XLS та CSV.")
    header_row = None
    for index, row in df.iterrows():
        values = {norm(value) for value in row if pd.notna(value)}
        if (values.intersection(map(norm, ALIASES["store"]))
                and values.intersection(map(norm, ALIASES["sales"]))):
            header_row = index
            break
    if header_row is None:
        raise ValueError("Не знайшов заголовки магазину та продажів у таблиці.")
    df.columns = [str(c).strip() for c in df.loc[header_row]]
    df = df.loc[header_row + 1:].copy()
    df = df.dropna(subset=[find_col(df, ALIASES["store"])])

    cols = {
        key: find_col(df, aliases, required=(key != "plan"))
        for key, aliases in ALIASES.items()
    }
    out = pd.DataFrame()
    out["store"] = df[cols["store"]].astype(str).str.strip()
    for key in ("sales", "py", "direct", "pickup", "services", "accessories", "plan"):
        col = cols[key]
        out[key] = to_num(df[col]) if col else 0.0
    out = out[out["store"].notna() & (out["store"] != "")].copy()
    out = out.groupby("store", as_index=False).sum(numeric_only=True)
    return out

def db():
    Path(DB_PATH).parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS snapshots (
            report_date TEXT NOT NULL,
            store TEXT NOT NULL,
            sales REAL, py REAL, direct REAL, pickup REAL,
            services REAL, accessories REAL, plan REAL,
            PRIMARY KEY (report_date, store)
        )
    """)
    return conn

def get_snapshot(report_date):
    conn = db()
    df = pd.read_sql_query(
        "SELECT * FROM snapshots WHERE report_date = ?",
        conn, params=(report_date,)
    )
    conn.close()
    return df

def save_snapshot(report_date, df):
    conn = db()
    with conn:
        for _, row in df.iterrows():
            conn.execute("""
                INSERT OR REPLACE INTO snapshots
                (report_date, store, sales, py, direct, pickup, services, accessories, plan)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (report_date, row.store, row.sales, row.py, row.direct,
                  row.pickup, row.services, row.accessories, row.plan))
    conn.close()

def previous_snapshot(report_date):
    conn = db()
    row = conn.execute(
        "SELECT MAX(report_date) FROM snapshots WHERE report_date < ?",
        (report_date,)
    ).fetchone()
    conn.close()
    return row[0] if row and row[0] else None

def fmt(n):
    return f"${n:,.0f}".replace(",", " ")

def pct(n, d):
    return n / d * 100 if d else 0

def make_report(report_date, current, previous, previous_date):
    metrics = ["sales", "direct", "pickup", "py", "services", "accessories"]
    labels = {
        "sales": "Продажі", "direct": "Прямі", "pickup": "Пікап",
        "py": "Продажі PY", "services": "Послуги", "accessories": "Аксесуари"
    }
    cur = current.set_index("store")
    if previous is not None and not previous.empty:
        prev = previous.set_index("store")
        delta = cur[metrics].subtract(prev[metrics], fill_value=0)
        compare_text = f"Приріст за день: {previous_date} → {report_date}"
    else:
        prev = None
        delta = cur[metrics].copy()
        compare_text = "Перший звіт збережено як стартову базу. Денний приріст з'явиться після наступного звіту."

    totals = cur[metrics].sum()
    changes = delta.sum()
    plan = cur["plan"].sum()
    lines = [
        f"📊 РЕГІОН — {report_date}",
        compare_text,
        "",
        "ДЕННИЙ ПРИРІСТ" if prev is not None else "СТАРТОВІ НАКОПИЧУВАЛЬНІ ДАНІ"
    ]
    for m in metrics:
        lines.append(f"• {labels[m]}: {fmt(changes[m])}")
    report_day = datetime.strptime(report_date, "%Y-%m-%d")
    days_in_month = calendar.monthrange(report_day.year, report_day.month)[1]
    lines += ["", f"НАКОПИЧУВАЛЬНО З 01.{report_day.month:02d}"]
    for m in metrics:
        lines.append(f"• {labels[m]}: {fmt(totals[m])}")
    if plan:
        lines += [
            f"• План місяця: {fmt(plan)}",
            f"• Факт виконання плану: {pct(totals['sales'], plan):.2f}%",
            f"• Тренд виконання: {pct(totals['sales'] / report_day.day * days_in_month, plan):.2f}%",
            f"• Прогноз продажів за місяць: {fmt(totals['sales'] / report_day.day * days_in_month)}",
            "Тренд за поточним середньоденним темпом, включно з датою звіту.",
        ]

    if prev is not None:
        lines += ["", "ПРИРІСТ ПРОДАЖІВ ПО МАГАЗИНАХ"]
        store_delta = delta["sales"].sort_values(ascending=False)
        for store, value in store_delta.items():
            lines.append(f"• {store}: {fmt(value)}")
    else:
        lines += ["", "Наступного дня надішли новий накопичувальний звіт — бот порівняє його з цією базою."]
    return "\n".join(lines)

def is_owner(update: Update) -> bool:
    return bool(OWNER_ID) and update.effective_user and str(update.effective_user.id) == str(OWNER_ID)

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_owner(update):
        await update.message.reply_text("Цей бот приватний і доступний лише власнику.")
        return
    await update.message.reply_text(
        "Вітаю! Я бот щоденної аналітики регіону.\n\n"
        "Надішли Excel або CSV зі звітом Power BI. "
        "У підписі до файлу вкажи дату звіту у форматі ДД.ММ.РРРР "
        "(наприклад, 15.11.2026).\n\n"
        "Перший файл буде стартовою базою, наступні — порівнюватиму з попередніми."
    )

async def help_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_owner(update):
        await update.message.reply_text("Цей бот приватний і доступний лише власнику.")
        return
    await update.message.reply_text(
        "Надішли файл XLSX, XLS або CSV з накопичувальними даними з 01 числа. "
        "У підписі вкажи дату зрізу, наприклад: 16.11.2026.\n"
        "Команда /status покаже останню збережену дату."
    )

async def status(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_owner(update):
        await update.message.reply_text("Цей бот приватний і доступний лише власнику.")
        return
    conn = db()
    row = conn.execute("SELECT MAX(report_date) FROM snapshots").fetchone()
    conn.close()
    await update.message.reply_text(
        f"Останній збережений звіт: {row[0] or 'ще немає'}"
    )

async def dates_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_owner(update):
        return
    conn = db()
    try:
        rows = conn.execute(
            "SELECT report_date, COUNT(*) FROM snapshots GROUP BY report_date ORDER BY report_date DESC LIMIT 30"
        ).fetchall()
    finally:
        conn.close()
    lines = ["Останні збережені дати (до 30):"]
    lines.extend(f"{date} — {count} магазинів" for date, count in rows)
    await update.message.reply_text("\n".join(lines) if rows else "Звітів ще немає.")

async def report_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_owner(update):
        return
    conn = db()
    try:
        report_date = conn.execute("SELECT MAX(report_date) FROM snapshots").fetchone()[0]
    finally:
        conn.close()
    if not report_date:
        await update.message.reply_text("Спочатку надішли звіт.")
        return
    current = get_snapshot(report_date)
    prev_date = previous_snapshot(report_date)
    previous = get_snapshot(prev_date) if prev_date else None
    report = make_report(report_date, current, previous, prev_date)
    for offset in range(0, len(report), 4000):
        await update.message.reply_text(report[offset:offset + 4000])

async def store_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_owner(update):
        return
    conn = db()
    try:
        report_date = conn.execute("SELECT MAX(report_date) FROM snapshots").fetchone()[0]
    finally:
        conn.close()
    if not report_date:
        await update.message.reply_text("Спочатку надішли звіт.")
        return
    current = get_snapshot(report_date)
    query = norm(" ".join(context.args))
    stores = current["store"].tolist()
    if not query:
        await update.message.reply_text(
            "Введи назву, наприклад /store Ретровіль\n\nМагазини в останньому звіті:\n"
            + "\n".join(stores)
        )
        return
    exact = [name for name in stores if norm(name) == query]
    matches = exact or [name for name in stores if query in norm(name)]
    if len(matches) != 1:
        message = ("Уточни назву магазину:\n" + "\n".join(matches)
                   if matches else "Магазин не знайдено. /store покаже доступні назви.")
        await update.message.reply_text(message)
        return
    store = matches[0]
    current = current[current["store"] == store].copy()
    prev_date = previous_snapshot(report_date)
    previous = get_snapshot(prev_date) if prev_date else None
    if previous is not None:
        previous = previous[previous["store"] == store].copy()
    report = make_report(report_date, current, previous, prev_date)
    report = report.replace(f"📊 РЕГІОН — {report_date}", f"📊 {store} — {report_date}", 1)
    await update.message.reply_text(report)

WEEK_METRICS = {
    "sales": "Продажі", "direct": "Прямі", "pickup": "Пікап",
    "services": "Послуги", "accessories": "Аксесуари",
}


def period_boundaries(start, end):
    """Split cumulative monthly data into differences within each month."""
    segments = []
    while start <= end:
        month_end = start.replace(day=calendar.monthrange(start.year, start.month)[1])
        segment_end = min(end, month_end)
        baseline = start - timedelta(days=1) if start.day > 1 else None
        segments.append((baseline, segment_end))
        start = segment_end + timedelta(days=1)
    return segments


def weekly_report(end, store=None):
    # Use the latest complete Monday–Sunday week covered by the report date.
    end -= timedelta(days=(end.weekday() + 1) % 7)
    current_start = end - timedelta(days=6)
    previous_end = current_start - timedelta(days=1)
    previous_start = previous_end - timedelta(days=6)
    periods = [period_boundaries(previous_start, previous_end),
               period_boundaries(current_start, end)]
    required = sorted({day for segments in periods for pair in segments
                       for day in pair if day is not None})
    snapshots = {day: get_snapshot(day.isoformat()) for day in required}
    missing = []
    for day, frame in snapshots.items():
        if store:
            frame = frame[frame["store"] == store]
            snapshots[day] = frame
        if frame.empty:
            missing.append(day.strftime("%d.%m.%Y"))
    if missing:
        return ("Для порівняння потрібні накопичувальні звіти з початку відповідного місяця "
                "по ці дати включно" + (f" для {store}" if store else "")
                + ":\n" + "\n".join(missing))
    store_sets = [set(frame["store"]) for frame in snapshots.values()]
    if any(names != store_sets[0] for names in store_sets[1:]):
        return ("Склад магазинів у звітах на граничні дати відрізняється. "
                "Завантаж повні звіти з однаковим переліком магазинів або скористайся /week назва_магазину.")
    totals = []
    for segments in periods:
        total = pd.Series(0.0, index=list(WEEK_METRICS))
        for baseline, segment_end in segments:
            total += snapshots[segment_end][list(WEEK_METRICS)].sum()
            if baseline is not None:
                total -= snapshots[baseline][list(WEEK_METRICS)].sum()
        totals.append(total)
    previous, current = totals
    lines = [f"📊 ТИЖДЕНЬ ДО ТИЖНЯ — {store or 'РЕГІОН'}",
             f"Попередній календарний тиждень: {previous_start:%d.%m.%Y}–{previous_end:%d.%m.%Y}",
             f"Останній повний календарний тиждень: {current_start:%d.%m.%Y}–{end:%d.%m.%Y}",
             "", "Попередній → поточний | зміна"]
    for metric, label in WEEK_METRICS.items():
        old, new = previous[metric], current[metric]
        delta = new - old
        relative = f"{delta / old * 100:+.2f}%" if old > 0 else "база для % відсутня або ≤ 0"
        lines.append(f"• {label}: {fmt(old)} → {fmt(new)} | {delta:+,.0f} $ ({relative})".replace(",", " "))
    return "\n".join(lines)


async def week_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_owner(update):
        return
    conn = db()
    try:
        latest = conn.execute("SELECT MAX(report_date) FROM snapshots").fetchone()[0]
    finally:
        conn.close()
    if not latest:
        await update.message.reply_text("Спочатку надішли звіт.")
        return
    args = list(context.args)
    end = datetime.strptime(latest, "%Y-%m-%d").date()
    if args and re.fullmatch(r"\d{2}\.\d{2}\.\d{4}", args[-1]):
        try:
            end = datetime.strptime(args.pop(), "%d.%m.%Y").date()
        except ValueError:
            await update.message.reply_text("Некоректна дата. Приклад: /week Ретровіль 15.11.2026")
            return
    query = norm(" ".join(args))
    store = None
    if query and query != "регіон":
        conn = db()
        try:
            stores = [row[0] for row in conn.execute("SELECT DISTINCT store FROM snapshots ORDER BY store")]
        finally:
            conn.close()
        exact = [name for name in stores if norm(name) == query]
        matches = exact or [name for name in stores if query in norm(name)]
        if len(matches) != 1:
            message = "Уточни назву:\n" + "\n".join(matches) if matches else "Магазин не знайдено. Перелік — /store."
            await update.message.reply_text(message)
            return
        store = matches[0]
    await update.message.reply_text(weekly_report(end, store))


async def setup_commands(app):
    await app.bot.set_my_commands([
        BotCommand("week", "Порівняння календарних тижнів Пн–Нд"),
        BotCommand("store", "Показники конкретного магазину"),
        BotCommand("report", "Останній звіт і тренд регіону"),
        BotCommand("dates", "Збережені дати та кількість магазинів"),
        BotCommand("status", "Дата останнього звіту"),
        BotCommand("cancel", "Скасувати введення дати"),
        BotCommand("help", "Як завантажувати звіти"),
        BotCommand("start", "Початок роботи"),
    ])

async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_owner(update):
        await update.message.reply_text("Цей бот приватний і доступний лише власнику.")
        return ConversationHandler.END
    doc = update.message.document
    if not doc:
        return ConversationHandler.END
    filename = doc.file_name or "report.xlsx"
    if Path(filename).suffix.lower() not in (".xlsx", ".xls", ".csv"):
        await update.message.reply_text("Надішли файл XLSX, XLS або CSV.")
        return ConversationHandler.END

    context.user_data.pop("pending_file", None)
    caption = update.message.caption or ""
    match = re.search(r"(\d{2})[./-](\d{2})[./-](\d{4})", caption)
    if not match:
        context.user_data["pending_file"] = (doc.file_id, filename)
        await update.message.reply_text(
            "Вкажи дату цього звіту у форматі ДД.ММ.РРРР. Наприклад: 16.11.2026"
        )
        return DATE_WAIT

    report_date = f"{match.group(3)}-{match.group(2)}-{match.group(1)}"
    return await process_file(update, context, doc.file_id, filename, report_date)

async def date_received(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_owner(update):
        await update.message.reply_text("Цей бот приватний і доступний лише власнику.")
        return ConversationHandler.END
    match = re.search(r"(\d{2})[./-](\d{2})[./-](\d{4})", update.message.text or "")
    if not match:
        await update.message.reply_text("Не розпізнав дату. Введи, наприклад, 16.11.2026.")
        return DATE_WAIT
    report_date = f"{match.group(3)}-{match.group(2)}-{match.group(1)}"
    pending = context.user_data.pop("pending_file", None)
    if not pending:
        await update.message.reply_text("Файл не знайдено. Надішли його ще раз.")
        return ConversationHandler.END
    return await process_file(update, context, pending[0], pending[1], report_date)

async def process_file(update, context, file_id, filename, report_date):
    await update.message.reply_text("Файл отримано. Читаю таблицю…")
    try:
        datetime.strptime(report_date, "%Y-%m-%d")
        tg_file = await context.bot.get_file(file_id)
        bio = io.BytesIO()
        await tg_file.download_to_memory(out=bio)
        current = read_report(bio.getvalue(), filename)
        if current.empty:
            raise ValueError("У файлі не знайшов рядків із магазинами.")
        prev_date = previous_snapshot(report_date)
        previous = get_snapshot(prev_date) if prev_date else None
        save_snapshot(report_date, current)
        report = make_report(report_date, current, previous, prev_date)
        await update.message.reply_text(report)
    except Exception as exc:
        log.exception("Failed to process report")
        await update.message.reply_text(
            "Не вдалося обробити файл: " + str(exc) +
            "\nПеревір назви колонок і формат файлу."
        )
    return ConversationHandler.END

async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_owner(update):
        return ConversationHandler.END
    context.user_data.pop("pending_file", None)
    await update.message.reply_text("Скасовано.")
    return ConversationHandler.END

def main():
    if not TOKEN:
        raise RuntimeError("BOT_TOKEN environment variable is missing")
    if not OWNER_ID:
        raise RuntimeError("OWNER_ID environment variable is missing")
    app = Application.builder().token(TOKEN).post_init(setup_commands).build()
    conv = ConversationHandler(
        entry_points=[MessageHandler(filters.Document.ALL, handle_document)],
        states={DATE_WAIT: [MessageHandler(filters.TEXT & ~filters.COMMAND, date_received)]},
        fallbacks=[CommandHandler("cancel", cancel)],
        allow_reentry=True,
    )
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("help", help_cmd))
    app.add_handler(CommandHandler("status", status))
    app.add_handler(CommandHandler("week", week_cmd))
    app.add_handler(CommandHandler("store", store_cmd))
    app.add_handler(CommandHandler("report", report_cmd))
    app.add_handler(CommandHandler("dates", dates_cmd))
    app.add_handler(conv)
    app.add_handler(CommandHandler("cancel", cancel))
    app.run_polling()

if __name__ == "__main__":
    main()
