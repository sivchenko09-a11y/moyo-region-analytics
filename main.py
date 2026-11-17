import os
import math
import asyncio
import json
from types import SimpleNamespace
import calendar
import re
import io
import sqlite3
import logging
from datetime import datetime, timedelta, time as dt_time
from zoneinfo import ZoneInfo
from pathlib import Path

import pandas as pd
from telegram import Update, BotCommand, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    Application, CommandHandler, MessageHandler, ConversationHandler, CallbackQueryHandler,
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
    trend_frame = current.copy()
    for metric in PLAN_COLUMNS:
        trend_frame["cumulative_" + metric] = trend_frame[metric]
    trend_store = trend_frame.iloc[0].store if len(trend_frame) == 1 else None
    lines.extend(monthly_trend_lines({report_date: trend_frame}, trend_store))

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
        "Команда /status покаже останню збережену дату.\n"
        "/focus — фокуси й категорії: дата (місяць до місяця / Рік до Року), тиждень та Поточний місяць.\n"
        "/weeks — обрати тиждень за номером без порівняння.\n"
        "/week — порівняти календарні тижні."
    )

async def status(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_owner(update):
        await update.message.reply_text("Цей бот приватний і доступний лише власнику.")
        return
    frames = daily_frames()
    if frames:
        await update.message.reply_text("Денні звіти:\n" + "\n".join(f"{day} — {len(frame)} магазинів" for day, frame in list(frames.items())[-30:]))
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
    frames = daily_frames()
    if frames:
        await update.message.reply_text("Денні звіти:\n" + "\n".join(f"{day} — {len(frame)} магазинів" for day, frame in list(frames.items())[-30:]))
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
    if await daily_command(update, context, "report"):
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
        await update.message.reply_text(bold_report_percentages(report[offset:offset + 4000]), parse_mode="HTML")

async def store_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_owner(update):
        return
    if not context.args:
        await choose_store(update, "store")
        return
    if await daily_command(update, context, "store"):
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
    await update.message.reply_text(bold_report_percentages(report), parse_mode="HTML")

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
             "Попередній: " + week_title(previous_start),
             "Останній повний: " + week_title(current_start),
             "", "Попередній → поточний | зміна"]
    for metric, label in WEEK_METRICS.items():
        lines.append(comparison_line(label, previous[metric], current[metric], money=True))
    for name in sorted(store_sets[0]):
        values = []
        for segments in periods:
            total = pd.Series(0.0, index=list(WEEK_METRICS))
            for baseline, segment_end in segments:
                frame = snapshots[segment_end]
                total += frame[frame.store == name][list(WEEK_METRICS)].sum()
                if baseline is not None:
                    frame = snapshots[baseline]
                    total -= frame[frame.store == name][list(WEEK_METRICS)].sum()
            values.append(pd.DataFrame([total]))
        actions = decline_actions(values[0], values[1], end, include_quality=False)
        if actions:
            lines.extend(["", f"План дій — {name}", "Цілі: повернення до рівня попереднього тижня.", *actions])
    conn = db()
    try:
        last = conn.execute("SELECT MAX(report_date) FROM snapshots").fetchone()[0]
    finally:
        conn.close()
    if last:
        latest_frame = get_snapshot(last)
        for metric in PLAN_COLUMNS:
            latest_frame["cumulative_" + metric] = latest_frame[metric]
        lines.extend(monthly_trend_lines({last: latest_frame}, store))
    return "\n".join(lines)


async def week_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_owner(update):
        return
    if not context.args:
        await choose_store(update, "week")
        return
    if await daily_command(update, context, "week"):
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
    await update.message.reply_text(bold_report_percentages(weekly_report(end, store)), parse_mode="HTML")


def daily_db():
    conn = db()
    conn.execute("CREATE TABLE IF NOT EXISTS daily_reports (report_date TEXT PRIMARY KEY, payload TEXT NOT NULL)")
    return conn


def daily_frames():
    conn = daily_db()
    try:
        stored = {day: pd.read_json(io.StringIO(payload), orient="records")
                  for day, payload in conn.execute("SELECT report_date, payload FROM daily_reports ORDER BY report_date")}
    finally:
        conn.close()
    return derive_daily_sales(stored)


def derive_daily_sales(stored):
    # Stored sales are month-to-date snapshots; quality values are daily.
    result = {}
    metrics = list(WEEK_METRICS) + ["py"]
    for day, source in stored.items():
        frame = source.copy()
        for metric in metrics:
            frame["cumulative_" + metric] = frame[metric]
        report_day = datetime.strptime(day, "%Y-%m-%d").date()
        if report_day.day > 1:
            previous_date = (report_day - timedelta(days=1)).isoformat()
            previous = stored.get(previous_date)
            if previous is None:
                frame[metrics] = float("nan")
            else:
                baseline = previous.set_index("store")
                for metric in metrics:
                    frame[metric] = source[metric] - source.store.map(baseline[metric])
        result[day] = frame
    return result


def read_daily(raw):
    sheets = pd.read_excel(io.BytesIO(raw), sheet_name=None)
    sheets = {name.strip(): frame for name, frame in sheets.items()}
    if not {"Продажі", "Якісні"}.issubset(sheets):
        raise ValueError("Потрібні вкладки Продажі та Якісні.")
    buffer = io.BytesIO()
    sheets["Продажі"].to_excel(buffer, index=False)
    sales = read_report(buffer.getvalue(), "sales.xlsx")
    quality = sheets["Якісні"].copy()
    quality.columns = quality.columns.str.strip()
    quality = quality.dropna(subset=["Склад"])
    quality["store"] = quality["Склад"].astype(str).str.strip()
    if quality["store"].duplicated().any():
        raise ValueError("На вкладці Якісні повторюються магазини.")
    mapping = {"traffic": "Трафік загальний", "receipts": "Кіл-ть Чеків",
               "pudo": "Трафік PUDO", "ktc": "К-ть Товарів у Чеку", "quality_sales": "Сума, $"}
    for target, column in mapping.items():
        values = pd.to_numeric(quality[column].astype(str).str.replace(r"\s+", "", regex=True).str.replace(",", ".", regex=False), errors="coerce")
        if values.isna().any():
            raise ValueError(f"Порожні або некоректні числа у колонці {column}.")
        quality[target] = values
    if (quality[["traffic", "receipts", "pudo", "ktc"]] < 0).any().any():
        raise ValueError("Трафік, чеки та КТЧ мають бути невід’ємними.")
    if (quality.pudo > quality.traffic).any():
        raise ValueError("Трафік PUDO перевищує загальний трафік.")
    if set(sales.store) != set(quality.store):
        raise ValueError("Перелік магазинів на двох вкладках має збігатися.")
    quality["estimated_items"] = quality.ktc * quality.receipts
    return sales.merge(quality[["store", *mapping, "estimated_items"]], on="store", validate="one_to_one")


def daily_period(frames, start, end, store=None):
    days = [start + timedelta(days=i) for i in range((end-start).days+1)]
    missing = [day.isoformat() for day in days if day.isoformat() not in frames]
    if missing:
        raise ValueError("Бракує денних звітів: " + ", ".join(missing))
    selected = [frames[day.isoformat()] for day in days]
    if store:
        if any(store not in set(frame.store) for frame in selected):
            raise ValueError("Магазин відсутній у частині денних звітів.")
        selected = [frame[frame.store == store] for frame in selected]
    elif any(set(frame.store) != set(selected[0].store) for frame in selected):
        raise ValueError("Перелік магазинів у денних звітах відрізняється.")
    if any(frame[list(WEEK_METRICS)].isna().any().any() for frame in selected):
        raise ValueError("Для денних продажів бракує звіту за попередню дату або магазину в ньому. Завантаж також дату перед початком періоду.")
    return pd.concat(selected, ignore_index=True)


def quality_summary(frame):
    totals = frame.sum(numeric_only=True)
    traffic = totals.traffic - totals.pudo
    return {"Трафік": traffic, "Чеки": totals.receipts,
            "КОП, %": totals.receipts / traffic * 100 if traffic else None,
            "Середній чек, $": totals.quality_sales / totals.receipts if totals.receipts else None,
            "КТЧ ≈": totals.estimated_items / totals.receipts if totals.receipts else None}


def daily_plan_lines(frame, report_day, plans):
    days = calendar.monthrange(report_day.year, report_day.month)[1]
    lines = ["ПЛАН / ФАКТ ЗА ДЕНЬ", f"Денний план = місячний план ÷ {days}."]
    for metric, label in (("sales", "Оборот"), ("accessories", "Аксесуари"), ("services", "Послуги")):
        values, missing = [], []
        for _, row in frame.iterrows():
            value = plans.loc[row.store, metric] if row.store in plans.index else None
            if (value is None or pd.isna(value)) and metric == "sales":
                value = row.get("plan")
            if value is None or pd.isna(value):
                missing.append(row.store)
            else:
                values.append(float(value))
        daily_target = sum(values) / days if not missing else None
        if frame[metric].isna().any():
            target_text = f"план {fmt(daily_target)}; " if daily_target is not None else ""
            lines.append(f"⚪ {label}: {target_text}факт і виконання не визначені — потрібен звіт за {report_day-timedelta(days=1):%d.%m.%Y} для всіх магазинів.")
            continue
        actual = frame[metric].sum()
        if missing:
            lines.append(f"⚪ {label}: факт {fmt(actual)}; немає місячного плану для: " + ", ".join(missing))
            continue
        if daily_target <= 0:
            lines.append(f"⚪ {label}: план {fmt(daily_target)} / факт {fmt(actual)}; виконання у % не визначається за нульового плану.")
            continue
        fulfillment = actual / daily_target * 100
        difference = actual - daily_target
        if difference < -1e-9:
            status, icon = f"недовиконання {fmt(-difference)}", "🔴"
        elif difference > 1e-9:
            status, icon = f"перевиконання {fmt(difference)}", "🟢"
        else:
            status, icon = "план виконано", "🟢"
        lines.append(f"{icon} {label}: план {fmt(daily_target)} / факт {fmt(actual)}; виконання {fulfillment:.2f}%; {status}.")
    return lines



def daily_summary(frames, store=None):
    latest = max(frames)
    end = datetime.strptime(latest, "%Y-%m-%d").date()
    day = frames[latest]
    if store:
        day = day[day.store == store]
    if day.empty:
        raise ValueError("Магазин відсутній в останньому звіті.")
    lines = [f"📊 {store or 'РЕГІОН'} — {latest}", "ЗА ДЕНЬ"]
    plans = load_monthly_plans(latest[:7])
    lines.extend(daily_plan_lines(day, end, plans))
    lines.extend(["", "Інші показники за день:"])
    for metric, label in (("direct", "Прямі"), ("pickup", "Пікап")):
        if day[metric].isna().any():
            lines.append(f"• {label}: потрібен звіт за {end-timedelta(days=1):%d.%m.%Y}")
        else:
            lines.append(f"• {label}: {fmt(day[metric].sum())}")
    for label, value in quality_summary(day).items():
        lines.append(f"• {label}: {value:.2f}" if value is not None else f"• {label}: немає бази")
    lines.extend(monthly_trend_lines(frames, store, plans))
    return "\n".join(lines)


def week_title(start):
    year, number, _ = start.isocalendar()
    end = start + timedelta(days=6)
    return f"Тиждень {number} ({year}) · {start:%d.%m.%Y}–{end:%d.%m.%Y}"


def single_week_report(frames, monday, store=None):
    sunday = monday + timedelta(days=6)
    dates = sorted(day for day in frames if monday.isoformat() <= day <= sunday.isoformat())
    if not dates:
        return "За цей тиждень даних немає. Онови їх через /sync."
    start = datetime.strptime(dates[0], "%Y-%m-%d").date()
    end = datetime.strptime(dates[-1], "%Y-%m-%d").date()
    lines = [f"📊 {store or 'РЕГІОН'}", week_title(monday)]
    if len(dates) < 7:
        lines.append(f"Неповний тиждень: дані за {start:%d.%m}–{end:%d.%m.%Y} ({len(dates)} із 7 днів).")
    frame = daily_period(frames, start, end, store)
    for metric, label in WEEK_METRICS.items():
        lines.append(f"• {label}: {fmt(frame[metric].sum())}")
    sales = frame.sales.sum()
    for metric, label in (("direct", "Частка прямих"), ("pickup", "Частка пікап"),
                          ("accessories", "Частка аксесуарів"), ("services", "Частка послуг")):
        value = f"{frame[metric].sum() / sales * 100:.2f}%" if sales else "немає бази"
        lines.append(f"• {label}: {value}")
    for label, value in quality_summary(frame).items():
        lines.append(f"• {label}: {value:.2f}" if value is not None else f"• {label}: немає бази")
    lines.extend(monthly_trend_lines(frames, store))
    return "\n".join(lines)


async def weeks_menu(message, page=0):
    frames = daily_frames()
    mondays = sorted({(datetime.strptime(day, "%Y-%m-%d").date() -
                       timedelta(days=datetime.strptime(day, "%Y-%m-%d").weekday()))
                      for day in frames}, reverse=True)
    if not mondays:
        await message.reply_text("Спочатку онови дані через /sync.")
        return
    page = min(max(page, 0), (len(mondays)-1)//12)
    buttons = [[InlineKeyboardButton(week_title(day), callback_data=f"wsdate:{day.isoformat()}")]
               for day in mondays[page*12:(page+1)*12]]
    nav = []
    if page:
        nav.append(InlineKeyboardButton("Новіші", callback_data=f"wspage:{page-1}"))
    if (page+1)*12 < len(mondays):
        nav.append(InlineKeyboardButton("Старіші", callback_data=f"wspage:{page+1}"))
    if nav:
        buttons.append(nav)
    await message.reply_text("Обери тиждень (Пн–Нд):", reply_markup=InlineKeyboardMarkup(buttons))


async def weeks_cmd(update, context):
    if is_owner(update):
        await weeks_menu(update.message)


async def weeks_selected(update, context):
    query = update.callback_query
    if not is_owner(update):
        await query.answer("Доступ лише власнику.", show_alert=True)
        return
    await query.answer()
    action, value = query.data.split(":", 1)
    if action == "wspage":
        await weeks_menu(query.message, int(value))
        return
    frames = daily_frames()
    if action == "wsdate":
        monday = datetime.strptime(value, "%Y-%m-%d").date()
        sunday = monday + timedelta(days=6)
        names = sorted({name for day, frame in frames.items()
                        if value <= day <= sunday.isoformat() for name in frame.store})
        if not names:
            await query.message.reply_text("Даних немає. Відкрий /weeks знову.")
            return
        import uuid
        menu_id = uuid.uuid4().hex[:12]
        menus = context.user_data.setdefault("week_menus", {})
        menus[menu_id] = (value, names)
        if len(menus) > 50:
            del menus[next(iter(menus))]
        buttons = [[InlineKeyboardButton("Увесь регіон", callback_data=f"wspick:{menu_id}:-1")]]
        buttons.extend([InlineKeyboardButton(name, callback_data=f"wspick:{menu_id}:{index}")]
                       for index, name in enumerate(names))
        await query.message.reply_text(week_title(monday) + "\nОбери магазин або регіон:",
                                       reply_markup=InlineKeyboardMarkup(buttons))
        return
    menu_id, index = value.split(":")
    menu = context.user_data.get("week_menus", {}).get(menu_id)
    index = int(index)
    if menu is None or index < -1 or index >= len(menu[1]):
        await query.message.reply_text("Відкрий вибір тижня знову через /weeks.")
        return
    monday = datetime.strptime(menu[0], "%Y-%m-%d").date()
    try:
        report = single_week_report(frames, monday, menu[1][index] if index >= 0 else None)
    except ValueError as exc:
        report = str(exc)
    for offset in range(0, len(report), 3500):
        await query.message.reply_text(bold_report_percentages(report[offset:offset+3500]), parse_mode="HTML")




PLAN_COLUMNS = {"sales": "План обороту", "accessories": "План аксесуарів", "services": "План послуг"}


def parse_monthly_plans(frame):
    required = ["Місяць", "Склад", *PLAN_COLUMNS.values()]
    if any(column not in frame.columns for column in required):
        raise ValueError("Вкладка Плани: потрібні колонки " + ", ".join(required))
    records, seen = [], set()
    for _, row in frame.iterrows():
        if all(pd.isna(row[c]) or str(row[c]).strip() == "" for c in required):
            continue
        name = str(row["Склад"]).strip() if pd.notna(row["Склад"]) else ""
        month = str(row["Місяць"]).strip()
        try:
            month = datetime.strptime(month, "%m.%Y").strftime("%Y-%m")
        except ValueError:
            raise ValueError(f"Вкладка Плани: місяць для {name or 'рядка'} має бути текстом ММ.РРРР, наприклад 11.2026.")
        if not name:
            raise ValueError("Вкладка Плани: заповни Склад.")
        if (month, name) in seen:
            raise ValueError(f"Вкладка Плани: повтор магазину {name} за {month}.")
        seen.add((month, name))
        values = []
        for column in PLAN_COLUMNS.values():
            raw = row[column]
            if pd.isna(raw) or str(raw).strip() == "":
                values.append(None)
                continue
            try:
                value = float(re.sub(r"\s+", "", str(raw)).replace(",", "."))
            except ValueError:
                raise ValueError(f"Вкладка Плани: {name}, {column} — потрібне число в доларах.")
            if not math.isfinite(value) or value < 0:
                raise ValueError(f"Вкладка Плани: {name}, {column} — план має бути невід’ємним числом.")
            values.append(value)
        records.append((month, name, *values))
    return records


def monthly_plans_db(conn):
    conn.execute("CREATE TABLE IF NOT EXISTS monthly_plans (month TEXT, store TEXT, sales REAL, accessories REAL, services REAL, PRIMARY KEY(month, store))")


def load_monthly_plans(month):
    conn = db()
    try:
        monthly_plans_db(conn)
        return pd.read_sql_query("SELECT * FROM monthly_plans WHERE month = ?", conn, params=(month,)).set_index("store")
    finally:
        conn.close()


def monthly_trend_lines(frames, store=None, plans=None):
    if not frames:
        return ["Місячний тренд: немає даних."]
    latest = max(frames)
    day = datetime.strptime(latest, "%Y-%m-%d").date()
    frame = frames[latest]
    if store:
        frame = frame[frame.store == store]
    lines = ["", f"📈 Місячний тренд на {day:%d.%m.%Y} — {store or 'РЕГІОН'}",
             f"Місяць {day:%m.%Y}; минуло {day.day} із {calendar.monthrange(day.year, day.month)[1]} днів."]
    if frame.empty:
        return lines + ["Магазин відсутній у звіті за останню відому дату."]
    if plans is None:
        plans = load_monthly_plans(latest[:7])
    for metric, label in (("sales", "Оборот"), ("accessories", "Аксесуари"), ("services", "Послуги")):
        actuals = frame["cumulative_" + metric]
        if actuals.isna().any():
            lines.append(f"⚪ {label}: неповні дані факту.")
            continue
        actual = actuals.sum()
        targets, missing = [], []
        for _, row in frame.iterrows():
            value = plans.loc[row.store, metric] if row.store in plans.index else None
            # Preserve the existing turnover plan when the new tab has no override.
            if (value is None or pd.isna(value)) and metric == "sales":
                value = row.get("plan")
            if value is None or pd.isna(value):
                missing.append(row.store)
            else:
                targets.append(float(value))
        if missing:
            lines.append(f"⚪ {label}: факт {fmt(actual)}; немає плану для: " + ", ".join(missing))
            continue
        target = sum(targets)
        if target <= 0:
            lines.append(f"⚪ {label}: факт {fmt(actual)}; план {fmt(target)} — тренд не визначається.")
            continue
        forecast = actual / day.day * calendar.monthrange(day.year, day.month)[1]
        trend = forecast / target * 100
        icon = "🟢" if trend >= 100 else "🔴"
        lines.append(f"{icon} {label}: факт {fmt(actual)} / план {fmt(target)}; виконання {actual/target*100:.2f}%; тренд {trend:.2f}% (прогноз {fmt(forecast)}).")
    return lines



def comparison_line(label, old, new, money=False):
    if old is None or new is None or pd.isna(old) or pd.isna(new):
        return f"⚪ {label}: немає бази для порівняння"
    delta = new - old
    icon = "🟢" if delta > 1e-9 else "🔴" if delta < -1e-9 else "⚪"
    change = f"{delta / old * 100:+.2f}%" if old > 0 else "немає бази для %"
    values = f"{fmt(old)} → {fmt(new)}" if money else f"{old:.2f} → {new:.2f}"
    points = f"; {delta:+.2f} в.п." if label == "КОП, %" else ""
    return f"{icon} {label}: {values} ({change}{points})"


def decline_actions(before, after, end, include_quality=True):
    # Targets restore the previous week's level; they are not forecasts.
    actions = []
    checkpoint = end + timedelta(days=3)
    deadline = end + timedelta(days=7)
    control = f"Контроль {checkpoint:%d.%m}: перевірка 5 чеків і консультацій; підсумок {deadline:%d.%m.%Y}."
    tactics = {
        "direct": ("Оборот прямий", "Керуючому виділити продавця для активних консультацій у пікові години та відпрацювати з командою пропозицію комплекту."),
        "services": ("Послуги", "Керуючому відпрацювати з командою презентацію доречної послуги до основного товару; на контрольній точці розібрати причини відмов."),
        "accessories": ("Аксесуари", "Керуючому сформувати комплекти аксесуарів до ходових моделей, перевірити наявність і закріпити пропозицію комплекту на консультації."),
    }
    for metric, (label, action) in tactics.items():
        old, new = before[metric].sum(), after[metric].sum()
        if new < old - 1e-9:
            actions.append(f"🔴 {label}: {action} Ціль на наступний тиждень — {fmt(old)} (+{fmt(old-new)} до поточного). {control}")
    if include_quality:
        oldq, newq = quality_summary(before), quality_summary(after)
        tactics_q = {
            "КОП, %": ("КОП", "%", "Керуючому організувати парні консультації у пікові години та на контрольній точці розібрати 5 незавершених продажів."),
            "Середній чек, $": ("Середній чек", " $", "Керуючому відпрацювати з продавцями презентацію дорожчої альтернативи й комплекту під потребу клієнта."),
            "КТЧ ≈": ("КТЧ", "", "Керуючому підготувати по 2 доречні доповнення до ходових товарів та відпрацювати їх пропозицію в команді."),
        }
        for key, (label, unit, action) in tactics_q.items():
            old, new = oldq[key], newq[key]
            if old is not None and new is not None and new < old - 1e-9:
                gap_unit = " в.п." if key == "КОП, %" else unit
                actions.append(f"🔴 {label}: {action} Ціль на наступний тиждень — {old:.2f}{unit} (+{old-new:.2f}{gap_unit} до поточного). {control}")
    return actions


def daily_week(frames, end, store=None):
    end -= timedelta(days=(end.weekday()+1)%7)
    current = daily_period(frames, end-timedelta(days=6), end, store)
    previous = daily_period(frames, end-timedelta(days=13), end-timedelta(days=7), store)
    lines = [f"📊 {store or 'РЕГІОН'} — календарні тижні",
             "Попередній: " + week_title(end-timedelta(days=13)),
             "Поточний: " + week_title(end-timedelta(days=6)),
             "Попередній → поточний | зміна"]
    for metric, label in WEEK_METRICS.items():
        lines.append(comparison_line(label, previous[metric].sum(), current[metric].sum(), money=True))
    oldq, newq = quality_summary(previous), quality_summary(current)
    for label in oldq:
        lines.append(comparison_line(label, oldq[label], newq[label]))
    lines.extend(monthly_trend_lines(frames, store))
    for name in sorted(set(current.store)):
        before, after = previous[previous.store == name], current[current.store == name]
        if before.empty:
            continue
        actions = decline_actions(before, after, end)
        if actions:
            lines.extend(["", f"План дій — {name}",
                          "Цілі: повернення до рівня попереднього тижня.", *actions])
    return "\n".join(lines)



def monthly_store_summary(frames, store=None, plans=None):
    latest = max(frames)
    end = datetime.strptime(latest, "%Y-%m-%d").date()
    start = end.replace(day=1)
    current = frames[latest]
    if store:
        current = current[current.store == store]
    if current.empty:
        raise ValueError("Магазин відсутній у звіті за останню завантажену дату.")
    lines = [f"📊 {store or 'РЕГІОН'} — З ПОЧАТКУ МІСЯЦЯ",
             f"{start:%d.%m.%Y}–{end:%d.%m.%Y}"]
    trend = monthly_trend_lines(frames, store, plans)
    lines.extend(trend[1:])
    sales = current.cumulative_sales.sum()
    lines.extend(["", "Продажі за місяць:"])
    for metric, label in (("direct", "Прямий оборот"), ("pickup", "Пікап")):
        value = current["cumulative_" + metric].sum()
        share = f"; частка {value/sales*100:.2f}%" if sales else ""
        lines.append(f"• {label}: {fmt(value)}{share}")
    for metric, label in (("accessories", "Частка аксесуарів"), ("services", "Частка послуг")):
        value = current["cumulative_" + metric].sum()
        lines.append(f"• {label}: {value/sales*100:.2f}%" if sales else f"• {label}: немає бази")
    lines.extend(["", "Якісні показники з початку місяця:"])
    selected, missing = [], []
    expected_names = set(current.store)
    for offset in range(end.day):
        day = (start + timedelta(days=offset)).isoformat()
        frame = frames.get(day)
        if frame is None:
            missing.append(day)
            continue
        if store:
            frame = frame[frame.store == store]
        if frame.empty or set(frame.store) != expected_names:
            missing.append(day)
            continue
        selected.append(frame)
    if missing:
        lines.append("Не розраховано за повний період: бракує звітів або магазинів за " + ", ".join(missing) + ".")
    else:
        # Quality columns are daily: aggregate totals, then calculate weighted ratios.
        quality = quality_summary(pd.concat(selected, ignore_index=True))
        for label, value in quality.items():
            lines.append(f"• {label}: {value:.2f}" if value is not None else f"• {label}: немає бази")
    return "\n".join(lines)



async def daily_command(update, context, mode):
    frames=daily_frames()
    if not frames:
        return False
    args=list(context.args) if mode in ('store','week') else []
    end=datetime.strptime(max(frames), '%Y-%m-%d').date()
    try:
        if mode=='week' and args and re.fullmatch(r'\d{2}\.\d{2}\.\d{4}',args[-1]):
            end=datetime.strptime(args.pop(),'%d.%m.%Y').date()
        query=norm(' '.join(args)); store=None
        if query and query!='регіон':
            names=sorted({name for frame in frames.values() for name in frame.store})
            matches=[name for name in names if norm(name)==query] or [name for name in names if query in norm(name)]
            if len(matches)!=1:
                raise ValueError('Уточни магазин: '+', '.join(matches or names))
            store=matches[0]
        if mode=='store' and not query:
            result='Приклад: /store Материк\n'+'\n'.join(frames[max(frames)].store)
        elif mode=='week':
            result=daily_week(frames,end,store)
        elif mode=='store':
            result=monthly_store_summary(frames,store)
        else:
            result=daily_summary(frames,store)
    except ValueError as exc:
        result=str(exc)
    for offset in range(0,len(result),3500):
        await update.message.reply_text(bold_report_percentages(result[offset:offset+3500]), parse_mode="HTML")
    return True


async def choose_store(update, mode):
    frames = daily_frames()
    if frames:
        names = sorted(frames[max(frames)].store.tolist())
    else:
        conn = db()
        try:
            names = [row[0] for row in conn.execute(
                "SELECT store FROM snapshots WHERE report_date = (SELECT MAX(report_date) FROM snapshots) ORDER BY store"
            )]
        finally:
            conn.close()
    if not names:
        await update.message.reply_text("Спочатку надішли звіт.")
        return
    buttons = [[InlineKeyboardButton("Увесь регіон", callback_data=f"pick:{mode}:регіон")]]
    buttons.extend([InlineKeyboardButton(name, callback_data=f"pick:{mode}:{name}")] for name in names)
    await update.message.reply_text("Обери магазин або регіон:", reply_markup=InlineKeyboardMarkup(buttons))


async def store_selected(update, context):
    query = update.callback_query
    if not is_owner(update):
        await query.answer("Доступ лише власнику.", show_alert=True)
        return
    await query.answer()
    _, mode, name = query.data.split(":", 2)
    selected_update = SimpleNamespace(message=query.message, effective_user=update.effective_user)
    context.args = [name]
    if mode == "week":
        await week_cmd(selected_update, context)
    elif name == "регіон":
        if not await daily_command(selected_update, context, "store"):
            await report_cmd(selected_update, context)
    else:
        await store_cmd(selected_update, context)


async def day_menu(message, page=0):
    frames = daily_frames()
    dates = sorted(frames, reverse=True)
    if not dates:
        await message.reply_text("Спочатку завантаж Excel із вкладками Продажі та Якісні.")
        return
    page = min(max(page, 0), (len(dates)-1)//20)
    buttons = [[InlineKeyboardButton(datetime.strptime(day, "%Y-%m-%d").strftime("%d.%m.%Y"),
                                     callback_data=f"daydate:{day}")]
               for day in dates[page*20:(page+1)*20]]
    navigation = []
    if page:
        navigation.append(InlineKeyboardButton("Новіші дати", callback_data=f"daypage:{page-1}"))
    if (page+1)*20 < len(dates):
        navigation.append(InlineKeyboardButton("Старіші дати", callback_data=f"daypage:{page+1}"))
    if navigation:
        buttons.append(navigation)
    await message.reply_text("Обери дату звіту:", reply_markup=InlineKeyboardMarkup(buttons))


async def day_stores(message, day, context):
    frames = daily_frames()
    if day not in frames:
        await message.reply_text("За цю дату звіту немає. Вибери іншу дату через /day.")
        return
    names = sorted(frames[day].store.tolist())
    # Each menu keeps its own mapping, so older buttons remain unambiguous.
    import uuid
    menu_id = uuid.uuid4().hex[:12]
    menus = context.user_data.setdefault("day_menus", {})
    menus[menu_id] = (day, names)
    if len(menus) > 50:
        del menus[next(iter(menus))]
    buttons = [[InlineKeyboardButton("Увесь регіон", callback_data=f"daypick:{menu_id}:-1")]]
    buttons.extend([InlineKeyboardButton(name, callback_data=f"daypick:{menu_id}:{index}")]
                   for index, name in enumerate(names))
    await message.reply_text(f"{day} — обери магазин або регіон:", reply_markup=InlineKeyboardMarkup(buttons))


async def day_cmd(update, context):
    if not is_owner(update):
        return
    if not context.args:
        await day_menu(update.message)
        return
    try:
        day = datetime.strptime(context.args[0], "%d.%m.%Y").date().isoformat()
    except ValueError:
        await update.message.reply_text("Приклад: /day 03.11.2026. Або натисни /day для вибору дати.")
        return
    await day_stores(update.message, day, context)


async def day_selected(update, context):
    query = update.callback_query
    if not is_owner(update):
        await query.answer("Доступ лише власнику.", show_alert=True)
        return
    await query.answer()
    action, value = query.data.split(":", 1)
    if action == "daypage":
        await day_menu(query.message, int(value))
    elif action == "daydate":
        await day_stores(query.message, value, context)
    else:
        menu_id, index = value.split(":")
        menu = context.user_data.get("day_menus", {}).get(menu_id)
        if menu is None:
            await query.message.reply_text("Відкрий вибір дати знову через /day.")
            return
        day, names = menu
        index = int(index)
        if index < -1 or index >= len(names):
            await query.message.reply_text("Відкрий вибір дати знову через /day.")
            return
        frames = daily_frames()
        if day not in frames:
            await query.message.reply_text("За цю дату звіту немає.")
            return
        try:
            report = daily_summary({date: frame for date, frame in frames.items() if date <= day},
                                   names[index] if index >= 0 else None)
        except ValueError as exc:
            report = str(exc)
        for offset in range(0, len(report), 3500):
            await query.message.reply_text(bold_report_percentages(report[offset:offset+3500]), parse_mode="HTML")


SALES_SHEET_ID = "1JmXrIBABlzrUFG5IsOgjTD_rG7zbAzvdLBnB4FDlqnE"
QUALITY_SHEET_ID = "18snM8aRqhGmnws4Fx7UbyXzAKexpVzl5p60nEAKJaNQ"
PLANS_SHEET_ID = "1vPS13rER8_vzJxKkoKaDHcrWTcxjntQXJFEgLBu8w50"


def sheet_dates(metadata):
    dates = {}
    for sheet in metadata.get("sheets", []):
        title = sheet["properties"]["title"]
        try:
            day = datetime.strptime(title.strip(), "%d.%m.%Y").date().isoformat()
        except ValueError:
            continue
        if day in dates:
            raise ValueError(f"Дві вкладки позначають одну дату {day}. Залиш одну вкладку на дату.")
        dates[day] = title
    return dates


def google_table(session, sheet_id, title):
    from urllib.parse import quote
    escaped = "'" + title.replace("'", "''") + "'"
    response = session.get(
        f"https://sheets.googleapis.com/v4/spreadsheets/{sheet_id}/values/{quote(escaped, safe='')}",
        params={"valueRenderOption": "UNFORMATTED_VALUE"}, timeout=45)
    response.raise_for_status()
    rows = response.json().get("values", [])
    for index, row in enumerate(rows):
        if "Склад" in [str(value).strip() for value in row]:
            header = [str(value).strip() for value in row]
            body = rows[index+1:]
            if any(len(values) > len(header) for values in body):
                raise ValueError(f"Зайві колонки у вкладці {title}.")
            return pd.DataFrame([values + [None]*(len(header)-len(values)) for values in body], columns=header)
    raise ValueError(f"У вкладці {title} не знайдено заголовок Склад.")


def fetch_monthly_plans(session):
    response = session.get(f"https://sheets.googleapis.com/v4/spreadsheets/{PLANS_SHEET_ID}",
                           params={"fields": "sheets.properties.title"}, timeout=45)
    response.raise_for_status()
    months = {}
    for sheet in response.json().get("sheets", []):
        title = sheet["properties"]["title"]
        if not re.fullmatch(r"\d{2}\.\d{4}", title.strip()):
            continue
        try:
            month = datetime.strptime(title.strip(), "%m.%Y").strftime("%m.%Y")
        except ValueError:
            raise ValueError(f"Таблиця планів: некоректний місяць у назві вкладки {title}.")
        if month in months:
            raise ValueError(f"Таблиця планів: дві вкладки за {month}. Залиш одну.")
        months[month] = title
    if not months:
        raise ValueError("Таблиця планів: назви вкладки місяцями у форматі ММ.РРРР, наприклад 11.2026.")
    records = []
    for month, title in months.items():
        try:
            frame = google_table(session, PLANS_SHEET_ID, title)
            frame = frame.loc[~frame.apply(lambda row: all(pd.isna(v) or str(v).strip() == "" for v in row), axis=1)].copy()
            frame["Місяць"] = month
            records.extend(parse_monthly_plans(frame))
        except ValueError as exc:
            raise ValueError(f"Плани за {month}: {exc}") from exc
    return records


def sync_google_sheets():
    from google.oauth2.service_account import Credentials
    from google.auth.transport.requests import AuthorizedSession
    config = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON")
    if not config:
        raise ValueError("Додай GOOGLE_SERVICE_ACCOUNT_JSON у змінні Railway та надай службовому акаунту доступ до всіх трьох таблиць.")
    credentials = Credentials.from_service_account_info(json.loads(config), scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"])
    with AuthorizedSession(credentials) as session:
        calendars = []
        for sheet_id in (SALES_SHEET_ID, QUALITY_SHEET_ID):
            response = session.get(f"https://sheets.googleapis.com/v4/spreadsheets/{sheet_id}",
                                   params={"fields": "sheets.properties.title"}, timeout=45)
            response.raise_for_status()
            metadata = response.json()
            calendars.append(sheet_dates(metadata))
        sales_dates, quality_dates = calendars
        common = sorted(set(sales_dates) & set(quality_dates))
        if not common:
            raise ValueError("Немає спільних дат. Назви вкладки в обох таблицях однаково, наприклад 01.11.2026.")
        plans = fetch_monthly_plans(session)
        prepared = []
        for day in common:
            sales = google_table(session, SALES_SHEET_ID, sales_dates[day])
            quality = google_table(session, QUALITY_SHEET_ID, quality_dates[day])
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                sales.to_excel(writer, sheet_name="Продажі", index=False)
                quality.to_excel(writer, sheet_name="Якісні", index=False)
            try:
                frame = read_daily(buffer.getvalue())
                if frame.empty:
                    raise ValueError("Таблиця порожня.")
            except (ValueError, KeyError) as exc:
                raise ValueError(f"Перевір вкладки за {day}: {exc}") from exc
            prepared.append((day, frame.to_json(orient="records", force_ascii=False)))
    conn = daily_db()
    try:
        with conn:
            monthly_plans_db(conn)
            conn.executemany("INSERT OR REPLACE INTO daily_reports VALUES (?, ?)", prepared)
            conn.execute("DELETE FROM monthly_plans")
            conn.executemany("INSERT INTO monthly_plans VALUES (?, ?, ?, ?, ?)", plans)
    finally:
        conn.close()
    unpaired = sorted(set(sales_dates) ^ set(quality_dates))
    result = f"Оновлено {len(prepared)} дат. Остання дата — {common[-1]}. Перегляд — /day або /report."
    result += f"\nПлани: завантажено {len(plans)} рядків із окремої таблиці."
    if unpaired:
        result += "\nПропущені дати без пари вкладок: " + ", ".join(unpaired)
    return result


FOCUS_SHEET_ID = "1oGoXLgzy7sIp7tqg2oiziix3ImYZaFuyBnC-3CqfJ5Q"
# Accepted headers prevent copied historical blocks of unrelated products being imported.
FOCUS_CONFIG = {
    "ЕЗ": ("ЕЗ", "шт.", {"ез"}),
    "Клімат": ("Клімат", "шт.", {"клімат"}),
    "Скріни": ("Скріни", "шт.", {"скрін"}),
    "Вуха1": ("Вуха", "шт.", {"вуха"}),
    "Павери": ("Павери", "шт.", {"павери"}),
    "Носії1": ("Носії", "шт.", {"носії"}),
    "Android": ("Android", "шт.", {"андроїд"}),
    "Самсунг": ("Самсунг", "шт.", {"самсунг"}),
    "Айфон": ("Айфон", "шт.", {"айфон"}),
    "Ноутбуки": ("Ноутбуки", "шт.", {"ноути", "ноутбуки"}),
    "PlayStation": ("PlayStation", "шт.", {"ps", "playstation"}),
    "Акс PS": ("Акс PS", "шт.", {"акс ps"}),
    "Закази з магазину": ("Замовлення з магазину", "$", {"закази з магазину"}),
    "Виторг місяць + рік": ("Виторг", "$", {"оборот $"}),
}


def focus_header(value):
    return re.sub(r"^(штуки|шт\.?)\s*", "", str(value).strip().lower()).strip()


def focus_month(value):
    if isinstance(value, datetime):
        return value.year, value.month
    if isinstance(value, (int, float)) and 30000 < value < 80000:
        date = datetime(1899, 12, 30) + timedelta(days=value)
        return date.year, date.month
    for pattern in ("%m.%Y", "%d.%m.%Y", "%Y-%m-%d", "%m/%d/%Y"):
        try:
            date = datetime.strptime(str(value).strip(), pattern)
            return date.year, date.month
        except ValueError:
            pass
    return None


def parse_focus_sheets(sheets):
    result = {}
    found = set()
    for title, rows in sheets.items():
        key = title.strip()
        if key not in FOCUS_CONFIG:
            continue
        category, unit, headers = FOCUS_CONFIG[key]
        found.add(key)
        for r, row in enumerate(rows[:5]):
            for c, value in enumerate(row):
                if str(value).strip().lower() != "день" or c + 2 >= len(row):
                    continue
                if focus_header(row[c+1]) not in headers:
                    continue
                month = focus_month(row[c+2])
                if month is None:
                    raise ValueError(f"Фокуси: {title}, не розпізнав місяць у колонці {c+3}.")
                for values in rows[r+1:r+32]:
                    if c+1 >= len(values):
                        continue
                    raw_day, raw = values[c], values[c+1]
                    if raw is None or str(raw).strip() == "":
                        continue
                    try:
                        day_number = float(raw_day)
                        if not day_number.is_integer() or not 1 <= day_number <= 31:
                            continue
                        if day_number > calendar.monthrange(*month)[1]:
                            continue  # Template has 31 rows even in shorter months.
                        date = datetime(*month, int(day_number)).date().isoformat()
                        quantity = float(re.sub(r"\s+", "", str(raw)).replace(",", "."))
                    except (ValueError, TypeError):
                        raise ValueError(f"Фокуси: {title}, день {raw_day} — некоректне число.")
                    if not math.isfinite(quantity) or (unit == "шт." and not quantity.is_integer()):
                        raise ValueError(f"Фокуси: {title}, {date} — перевір кількість.")
                    identity = (category, date)
                    if identity in result:
                        raise ValueError(f"Фокуси: повтор {category} за {date}.")
                    result[identity] = (date, category, unit, quantity)
                    # The adjacent column contains the monetary amount on tabs
                    # where it is filled. Its header is the month, not "Сума".
                    amount_raw = values[c+2] if c+2 < len(values) else None
                    if unit != "$" and amount_raw is not None and str(amount_raw).strip() != "":
                        try:
                            amount = float(re.sub(r"\s+", "", str(amount_raw)).replace(",", "."))
                        except (ValueError, TypeError):
                            raise ValueError(f"Фокуси: {title}, {date} — некоректна сума.")
                        # Negative monetary values are valid: refunds/corrections
                        # can make the day's net amount negative.
                        if not math.isfinite(amount):
                            raise ValueError(f"Фокуси: {title}, {date} — сума не є скінченним числом.")
                        amount_identity = (category + " · сума", date)
                        if amount_identity in result:
                            raise ValueError(f"Фокуси: повтор суми {category} за {date}.")
                        result[amount_identity] = (date, category + " · сума", "грн", amount)
    missing = set(FOCUS_CONFIG) - found
    if missing:
        raise ValueError("Фокуси: не знайдено вкладки: " + ", ".join(sorted(missing)))
    anchor = [date for category, date in result if category == "ЕЗ"]
    if not anchor:
        raise ValueError("Фокуси: немає заповнених днів ЕЗ для визначення останньої дати.")
    cutoff = max(anchor)
    # ЕЗ is the completion marker; blank future cells there exclude padded zeros elsewhere.
    records = [row for row in result.values() if row[0] <= cutoff]
    return records, cutoff


def focus_db():
    conn = db()
    conn.execute("CREATE TABLE IF NOT EXISTS focus_reports (report_date TEXT, category TEXT, unit TEXT, value REAL, PRIMARY KEY(report_date, category))")
    return conn


def sync_focus_google():
    from google.oauth2.service_account import Credentials
    from google.auth.transport.requests import AuthorizedSession
    from urllib.parse import quote
    config = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON")
    if not config:
        raise ValueError("Додай GOOGLE_SERVICE_ACCOUNT_JSON у Railway.")
    creds = Credentials.from_service_account_info(json.loads(config), scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"])
    with AuthorizedSession(creds) as session:
        response = session.get(f"https://sheets.googleapis.com/v4/spreadsheets/{FOCUS_SHEET_ID}", params={"fields": "sheets.properties.title"}, timeout=45)
        response.raise_for_status()
        sheets = {}
        for item in response.json().get("sheets", []):
            title = item["properties"]["title"]
            if title.strip() not in FOCUS_CONFIG:
                continue
            if any(t.strip() == title.strip() for t in sheets):
                raise ValueError(f"Фокуси: дубль вкладки {title}.")
            address = "'" + title.replace("'", "''") + "'!A1:ZZ40"
            response = session.get(f"https://sheets.googleapis.com/v4/spreadsheets/{FOCUS_SHEET_ID}/values/{quote(address, safe='')}",
                                   params={"valueRenderOption": "UNFORMATTED_VALUE", "dateTimeRenderOption": "SERIAL_NUMBER"}, timeout=45)
            response.raise_for_status()
            sheets[title] = response.json().get("values", [])
    records, cutoff = parse_focus_sheets(sheets)
    conn = focus_db()
    try:
        with conn:
            conn.execute("DELETE FROM focus_reports")
            conn.executemany("INSERT INTO focus_reports VALUES (?, ?, ?, ?)", records)
    finally:
        conn.close()
    return f"Фокуси + категорії: {len(records)} записів, дані до {cutoff}. Перегляд — /focus."


def focus_data():
    conn = focus_db()
    try:
        return {(name, day): (unit, value) for day, name, unit, value in conn.execute("SELECT * FROM focus_reports")}
    finally:
        conn.close()


def focus_value(value, unit):
    if value is None:
        return "немає даних"
    return fmt(value) if unit == "$" else (f"{value:,.0f} грн".replace(",", " ") if unit == "грн" else f"{value:,.0f} шт.".replace(",", " "))


def focus_total(data, category, start, end):
    values = [data.get((category, (start+timedelta(days=i)).isoformat()))
              for i in range((end-start).days+1)]
    return sum(v[1] for v in values) if values and all(v is not None for v in values) else None


def focus_change(old, new, unit):
    if old is None or new is None:
        return "⚪ Немає повних даних для порівняння"
    delta = new-old
    icon = "🟢" if delta > 0 else "🔴" if delta < 0 else "⚪"
    change = f"{delta/old*100:+.2f}%" if old > 0 else "немає додатної бази для %"
    sign = "+" if delta > 0 else "−" if delta < 0 else ""
    return f"{icon} {focus_value(old, unit)} → {focus_value(new, unit)}; {sign}{focus_value(abs(delta), unit)} ({change})"


def focus_comparison_date(day, year, month):
    """Same calendar day; nonexistent dates are not silently substituted."""
    if day.day > calendar.monthrange(year, month)[1]:
        return None
    return day.replace(year=year, month=month)


def focus_categories(selected=None):
    """Return one selected focus or all directions, preserving sheet order."""
    values = list(FOCUS_CONFIG.values())
    if selected is None:
        return values
    if not 0 <= selected < len(values):
        raise ValueError("Невідомий напрямок фокусів.")
    return [values[selected]]



def focus_metrics(category, unit, data, start, end=None):
    """Show amount when it exists anywhere in the imported history."""
    metrics = [(category, unit, "Кількість" if unit != "$" else "Сума")]
    amount_category = category + " · сума"
    if unit != "$" and any(name == amount_category for name, _ in data):
        metrics.append((amount_category, "грн", "Сума"))
    return metrics


def focus_amount_value(value):
    """Amounts adjacent to quantity are in UAH, not USD."""
    if value is None:
        return "немає даних"
    return f"{value:,.0f} грн".replace(",", " ")


def focus_pair(data, category, unit, start, end):
    quantity = focus_total(data, category, start, end)
    if unit == "$":
        return focus_value(quantity, unit)
    amount = focus_total(data, category + " · сума", start, end)
    if amount is None:
        return focus_value(quantity, unit)
    return f"{focus_value(quantity, unit)} / {focus_amount_value(amount)}"


def focus_pair_change(data, category, unit, old_start, old_end, new_start, new_end):
    old = focus_total(data, category, old_start, old_end)
    new = focus_total(data, category, new_start, new_end)
    if old is None or new is None:
        return "⚪ Немає повних даних для порівняння"
    delta = new - old
    icon = "🟢" if delta > 0 else "🔴" if delta < 0 else "⚪"
    pct = f"{delta / old * 100:+.2f}%" if old > 0 else "немає бази для %"
    sign = "+" if delta > 0 else "−" if delta < 0 else ""
    result = (f"{icon} {focus_pair(data, category, unit, old_start, old_end)} → "
              f"{focus_pair(data, category, unit, new_start, new_end)}; "
              f"{sign}{focus_value(abs(delta), unit)} ({pct})")
    if unit != "$":
        amount_category = category + " · сума"
        old_amount = focus_total(data, amount_category, old_start, old_end)
        new_amount = focus_total(data, amount_category, new_start, new_end)
        if old_amount is not None and new_amount is not None:
            amount_delta = new_amount - old_amount
            # Yellow indicates opposing trends in units and revenue.
            # Do not flag a zero change as an opposing trend.
            if delta * amount_delta < 0:
                result = "🟡" + result[len(icon):]
            amount_pct = f"{amount_delta / old_amount * 100:+.2f}%" if old_amount > 0 else "немає бази для %"
            amount_sign = "+" if amount_delta > 0 else "−" if amount_delta < 0 else ""
            result += f" / {amount_sign}{focus_amount_value(abs(amount_delta))} ({amount_pct})"
    return result


def focus_date_report(data, day, selected=None):
    previous_month = day.replace(day=1) - timedelta(days=1)
    lines = [f"🎯 Фокуси + категорії — РЕГІОН, {day:%d.%m.%Y}"]
    for category, unit, _ in focus_categories(selected):
        lines.extend(["", category, f"Факт: {focus_pair(data, category, unit, day, day)}"])
        for label, year, month in (("місяць до місяця", previous_month.year, previous_month.month),
                                   ("Рік до Року", day.year - 1, day.month)):
            old = focus_comparison_date(day, year, month)
            lines.append(f"{label}: " + (focus_pair_change(data, category, unit, old, old, day, day)
                                        if old else "немає аналогічної календарної дати"))
    return "\n".join(lines)


def focus_mtd_report(data, day, selected=None):
    previous_month = day.replace(day=1) - timedelta(days=1)
    start = day.replace(day=1)
    lines = ["🎯 Фокуси + категорії — РЕГІОН · Поточний місяць",
             f"Поточний період: {start:%d.%m.%Y}–{day:%d.%m.%Y}"]
    for category, unit, _ in focus_categories(selected):
        lines.extend(["", "", category, f"Факт: {focus_pair(data, category, unit, start, day)}"])
        for label, year, month in (("місяць до місяця", previous_month.year, previous_month.month),
                                   ("Рік до Року", day.year - 1, day.month)):
            if day.day > calendar.monthrange(year, month)[1]:
                lines.append(f"{label}: немає рівного календарного періоду")
                continue
            old_end = day.replace(year=year, month=month)
            lines.append(f"{label}: " + focus_pair_change(
                data, category, unit, old_end.replace(day=1), old_end, start, day))
    return "\n".join(lines)


def focus_week_report(data, monday, selected=None):
    latest = datetime.strptime(max(day for _, day in data), "%Y-%m-%d").date()
    end = min(monday + timedelta(days=6), latest)
    if end < monday:
        return "За цей тиждень немає даних."
    previous_start, previous_end = monday - timedelta(days=7), end - timedelta(days=7)
    lines = ["🎯 Фокуси + категорії — РЕГІОН", week_title(monday),
             f"Порівняння: {previous_start:%d.%m.%Y}–{previous_end:%d.%m.%Y} → {monday:%d.%m.%Y}–{end:%d.%m.%Y}"]
    if (end - monday).days < 6:
        lines.append("Неповний тиждень: порівнюються однакові дні тижня.")
    for category, unit, _ in focus_categories(selected):
        lines.extend(["", category, f"Факт: {focus_pair(data, category, unit, monday, end)}",
                      "тиждень до тижня: " + focus_pair_change(
                          data, category, unit, previous_start, previous_end, monday, end)])
    return "\n".join(lines)


def bold_report_percentages(report):
    """Escape report text and bold every numeric percentage for Telegram HTML."""
    import html
    escaped = html.escape(str(report))
    return re.sub(r"(?<![\w])([+−-]?\d+(?:[.,]\d+)?%)", r"<b>\1</b>", escaped)


async def focus_send(message, report):
    # Telegram HTML: bold percentage changes only, not the rest of the comparison.
    import html
    chunks, chunk = [], ""
    for line in report.splitlines():
        escaped = html.escape(line)
        escaped = re.sub(r"(?<![\w])([+−-]?\d+(?:[.,]\d+)?%)", r"<b>\1</b>", escaped)
        if len(chunk) + len(escaped) + 1 > 3500 and chunk:
            chunks.append(chunk)
            chunk = ""
        chunk += escaped + "\n"
    if chunk:
        chunks.append(chunk)
    for part in chunks:
        await message.reply_text(part, parse_mode="HTML")


async def focus_cmd(update, context):
    if not is_owner(update):
        return
    buttons = [[InlineKeyboardButton("📋 Усі напрямки", callback_data="focus:menu:all")]]
    for i, (category, _, _) in enumerate(FOCUS_CONFIG.values()):
        buttons.append([InlineKeyboardButton(f"🎯 {category}", callback_data=f"focus:menu:{i}")])
    await update.message.reply_text("🎯 Фокуси + категорії — регіон\nОбери напрямок:",
                                    reply_markup=InlineKeyboardMarkup(buttons))


async def focus_selected(update, context):
    query = update.callback_query
    if not is_owner(update):
        await query.answer("Доступ лише власнику.", show_alert=True)
        return
    await query.answer()
    parts = query.data.split(":")
    mode = parts[1]
    data = focus_data()
    if not data:
        await query.message.reply_text("Спочатку онови фокуси через /sync.")
        return
    dates = sorted({day for _, day in data}, reverse=True)
    if mode == "menu":
        choice = parts[2]
        selected = None if choice == "all" else int(choice)
        name = "Усі напрямки" if selected is None else focus_categories(selected)[0][0]
        await query.message.reply_text(f"🎯 {name}\nОбери період:", reply_markup=InlineKeyboardMarkup([
            [InlineKeyboardButton("📅 Конкретна дата · порівняння", callback_data=f"focus:dates:{choice}:0")],
            [InlineKeyboardButton("📊 Тиждень до тижня", callback_data=f"focus:weeks:{choice}:0")],
            [InlineKeyboardButton("📈 Поточний місяць", callback_data=f"focus:mtd:{choice}:latest")],
            [InlineKeyboardButton("↩️ Інший напрямок", callback_data="focus:back:all")]]))
        return
    if mode == "back":
        await focus_cmd(SimpleNamespace(message=query.message, effective_user=update.effective_user), context)
        return
    if len(parts) != 4:
        await query.message.reply_text("Меню застаріло. Відкрий /focus знову.")
        return
    choice, value = parts[2], parts[3]
    selected = None if choice == "all" else int(choice)
    focus_categories(selected)
    if mode in ("dates", "weeks"):
        if mode == "weeks":
            choices = sorted({(datetime.strptime(day, "%Y-%m-%d").date() -
                              timedelta(days=datetime.strptime(day, "%Y-%m-%d").weekday())).isoformat()
                              for day in dates}, reverse=True)
        else:
            choices = dates
        page = min(max(int(value), 0), (len(choices)-1)//15)
        buttons = []
        for date in choices[page*15:(page+1)*15]:
            day = datetime.strptime(date, "%Y-%m-%d").date()
            label = week_title(day) if mode == "weeks" else f"{day:%d.%m.%Y}"
            action = "week" if mode == "weeks" else "date"
            buttons.append([InlineKeyboardButton(label, callback_data=f"focus:{action}:{choice}:{date}")])
        nav = []
        if page:
            nav.append(InlineKeyboardButton("Новіші", callback_data=f"focus:{mode}:{choice}:{page-1}"))
        if (page+1)*15 < len(choices):
            nav.append(InlineKeyboardButton("Старіші", callback_data=f"focus:{mode}:{choice}:{page+1}"))
        if nav:
            buttons.append(nav)
        buttons.append([InlineKeyboardButton("↩️ До періодів", callback_data=f"focus:menu:{choice}")])
        await query.message.reply_text("Обери тиждень:" if mode == "weeks" else "Обери дату:",
                                       reply_markup=InlineKeyboardMarkup(buttons))
        return
    if mode not in ("date", "week", "mtd"):
        await query.message.reply_text("Невідомий режим. Відкрий /focus.")
        return
    day = datetime.strptime(max(dates) if mode == "mtd" else value, "%Y-%m-%d").date()
    if mode == "date" and value not in dates:
        await query.message.reply_text("Цієї дати вже немає. Відкрий /focus знову.")
        return
    report = (focus_week_report(data, day, selected) if mode == "week" else
              focus_mtd_report(data, day, selected) if mode == "mtd" else
              focus_date_report(data, day, selected))
    await focus_send(query.message, report)


async def sync_cmd(update, context):
    if not is_owner(update):
        return
    if context.bot_data.get("sync_running"):
        await update.message.reply_text("Оновлення вже виконується.")
        return
    context.bot_data["sync_running"] = True
    try:
        await update.message.reply_text("Читаю таблиці Google…")
        result = await asyncio.to_thread(sync_google_sheets)
        try:
            result += "\n" + await asyncio.to_thread(sync_focus_google)
        except ValueError as exc:
            result += "\nФокуси не оновлено: " + str(exc)
        except Exception as exc:
            log.error("Focus sync failed: %s", type(exc).__name__)
            result += "\nФокуси не оновлено. Перевір доступ службового акаунта до таблиці Фокуси + категорії. Попередні дані фокусів збережені."
    except ImportError:
        result = "Додай google-auth[requests]>=2,<3 до requirements.txt та онови Railway."
    except ValueError as exc:
        result = str(exc)
    except Exception as exc:
        log.error("Google Sheets sync failed: %s", type(exc).__name__)
        result = "Оновлення не вдалося. Перевір доступ службового акаунта до всіх трьох таблиць, увімкнення Google Sheets API та з’єднання. Спробуй /sync знову."
    finally:
        context.bot_data["sync_running"] = False
    for offset in range(0, len(result), 3500):
        await update.message.reply_text(bold_report_percentages(result[offset:offset+3500]), parse_mode="HTML")


async def setup_commands(app):
    await app.bot.set_my_commands([
        BotCommand("focus", "🎯 Фокуси + категорії"),
        BotCommand("group_test", "Тест відправки у групу"),
        BotCommand("interim_test", "Тест проміжкового звіту 1С"),
        BotCommand("group_here", "Дізнатись ID групи"),
        BotCommand("sync", "Оновити дані з Google Таблиць"),
        BotCommand("day", "Обрати дату та магазин або регіон"),
        BotCommand("weeks", "Тиждень за номером без порівняння"),
        BotCommand("week", "Порівняння календарних тижнів Пн–Нд"),
        BotCommand("store", "Магазин за місяць: план, факт і тренд"),
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
        raw = bio.getvalue()
        if Path(filename).suffix.lower() in (".xlsx", ".xls") and "Якісні" in [name.strip() for name in pd.ExcelFile(io.BytesIO(raw)).sheet_names]:
            current = read_daily(raw)
            conn = daily_db()
            try:
                with conn:
                    conn.execute("INSERT OR REPLACE INTO daily_reports VALUES (?, ?)", (report_date, current.to_json(orient="records", force_ascii=False)))
            finally:
                conn.close()
            await update.message.reply_text("Звіт збережено: продажі накопичувально, якісні за день. Повторне завантаження за цю дату замінює дані.")
            result = daily_summary({day: frame for day, frame in daily_frames().items() if day <= report_date})
            for offset in range(0, len(result), 3500):
                await update.message.reply_text(bold_report_percentages(result[offset:offset+3500]), parse_mode="HTML")
            return ConversationHandler.END
        current = read_report(raw, filename)
        if current.empty:
            raise ValueError("У файлі не знайшов рядків із магазинами.")
        prev_date = previous_snapshot(report_date)
        previous = get_snapshot(prev_date) if prev_date else None
        save_snapshot(report_date, current)
        report = make_report(report_date, current, previous, prev_date)
        await update.message.reply_text(bold_report_percentages(report), parse_mode="HTML")
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


# Automated Telegram group digest (Kyiv time).
REPORT_TZ = ZoneInfo("Europe/Kyiv")
GROUP_CHAT_ID = os.getenv("REPORT_GROUP_CHAT_ID", "").strip()

def group_report_text(day):
    key = day.isoformat()
    frames = daily_frames()
    if key not in frames:
        raise ValueError(f"Немає продажів за {day:%d.%m.%Y}.")
    frame = frames[key]
    if frame.empty or frame[["sales", "direct", "pickup", "accessories", "services"]].isna().any().any():
        raise ValueError("Немає повних денних продажів (перевір також попередній день).")
    plans = load_monthly_plans(day.strftime("%Y-%m"))
    days = calendar.monthrange(day.year, day.month)[1]
    metrics = (("sales", "Оборот"), ("accessories", "Аксесуари"), ("services", "Послуги"))
    lines = [f"📊 МОЙО — РЕГІОН | {day:%d.%m.%Y}", "ВИТОРГ ЗА ВЧОРА"]
    for metric, label in (("sales", "Оборот"), ("direct", "Прямі"), ("pickup", "Пікап"),
                          ("accessories", "Аксесуари"), ("services", "Послуги")):
        lines.append(f"• {label}: {fmt(frame[metric].sum())}")
    lines.extend(["", "📈 ВИКОНАННЯ ДЕННОГО ПЛАНУ — РЕГІОН"])
    for metric, label in metrics:
        missing = [store for store in frame.store if store not in plans.index or
                   pd.isna(plans.loc[store, metric])]
        if missing:
            lines.append(f"⚪ {label}: немає плану ({', '.join(missing)})")
            continue
        target = sum(float(plans.loc[store, metric]) for store in frame.store) / days
        if target <= 0:
            lines.append(f"⚪ {label}: немає додатного плану")
        else:
            value = frame[metric].sum() / target * 100
            icon = "🟢" if value >= 100 else "🟡" if value >= 90 else "🔴"
            lines.append(f"{icon} {label}: {value:.2f}%")
    lines.extend(["", "🏬 МАГАЗИНИ — ВИКОНАННЯ ПЛАНУ", "[[TABLE]]"])
    ranking = []
    for _, row in frame.iterrows():
        result = []
        raw_values = []
        for metric, _ in metrics:
            if row.store not in plans.index or pd.isna(plans.loc[row.store, metric]):
                result.append("⚪    н/д")
                raw_values.append(None)
                continue
            target = float(plans.loc[row.store, metric]) / days
            if target <= 0:
                result.append("⚪    н/д")
                raw_values.append(None)
                continue
            value = float(row[metric]) / target * 100
            icon = "🟢" if value >= 100 else "🔴"
            result.append(f"{icon} {value:6.1f}%")
            raw_values.append(value)
        ranking.append((raw_values[0] if raw_values[0] is not None else -1,
                        str(row.store), result))
    # Short names keep all three percentage columns aligned on mobile.
    aliases = {
        "Київ-Олімпійський": "Олімпійський", "Київ-Аркадія": "Аркадія",
        "Київ-Нью Вей": "Нью Вей", "Київ-Хрещатик 21": "Хрещатик",
        "Київ-Аладдін": "Аладін", "Київ-Комод": "Комод",
        "Буча-Рітейл Парк": "Буча", "Київ-Квадрат": "Квадрат",
        "Київ-Материк": "Материк", "Київ-Ретровіль": "Ретровіль",
        "Київ-Дрімтаун_new": "Дрім Таун", "Київ-Лавіна Молл": "Лавіна",
        "Київ-Республіка Парк": "Республіка", "Київ-Блокбастер": "Блокбастер",
    }
    # All three metric cells have identical display widths, including emoji.
    # Telegram HTML <pre> renders the grid with a monospaced font.
    lines.append(f"{'Магазин':<15} {'ТО':>10} {'Акс':>10} {'Посл':>10}")
    for _, store, result in sorted(ranking, key=lambda item: item[0], reverse=True):
        short = aliases.get(store, store.removeprefix("Київ-"))
        # A colored circle occupies two terminal columns; 1 separator + 6 digits + %.
        cells = [cell.ljust(9) for cell in result]
        lines.append(f"{short[:15]:<15} {cells[0]} {cells[1]} {cells[2]}")
    lines.append("[[/TABLE]]")
    data = focus_data()
    if not data or not any(d == key for _, d in data):
        raise ValueError(f"Немає фокусів за {day:%d.%m.%Y}.")
    # Group digest: only daily ЕЗ and Клімат facts; /focus is unchanged.
    lines.extend(["", "🎯 ФОКУСИ ЗА ВЧОРА — ЕЗ / КЛІМАТ"])
    for category in ("ЕЗ", "Клімат"):
        unit = FOCUS_CONFIG[category][1]
        lines.extend(["", category, f"Факт: {focus_pair(data, category, unit, day, day)}"])
    return "\n".join(lines)


def digest_dashboard(day):
    """Render a Telegram-friendly PNG using actual stored sales, plans and focus data."""
    from PIL import Image, ImageDraw, ImageFont
    frames = daily_frames()
    key = day.isoformat()
    if key not in frames:
        raise ValueError(f"Немає продажів за {day:%d.%m.%Y}.")
    frame = frames[key]
    metrics = [('sales','Оборот'), ('accessories','Аксесуари'), ('services','Послуги')]
    if frame.empty or frame[[m for m,_ in metrics]+['direct','pickup']].isna().any().any():
        raise ValueError('Неповні продажі за звітний день')
    plans = load_monthly_plans(day.strftime('%Y-%m'))
    days = calendar.monthrange(day.year, day.month)[1]
    aliases = {'Київ-Олімпійський':'Олімпійський','Київ-Аркадія':'Аркадія',
               'Київ-Нью Вей':'Нью Вей','Київ-Хрещатик 21':'Хрещатик 21',
               'Київ-Аладдін':'Аладін','Київ-Комод':'Комод',
               'Буча-Рітейл Парк':'Буча','Київ-Квадрат':'Квадрат',
               'Київ-Материк':'Материк','Київ-Ретровіль':'Ретровіль',
               'Київ-Дрімтаун_new':'Дрім Таун','Київ-Лавіна Молл':'Лавіна',
               'Київ-Республіка Парк':'Республіка Парк','Київ-Блокбастер':'Блокбастер'}
    def target(store, metric, source=None):
        value = plans.loc[store,metric] if store in plans.index and metric in plans.columns else None
        if (value is None or pd.isna(value)) and metric=='sales' and source is not None:
            value = source.get('plan')
        return float(value) if value is not None and pd.notna(value) else None
    def percentage(actual, goal):
        return actual/goal*100 if goal is not None and goal>0 and pd.notna(actual) else None
    region=[]
    for metric,label in metrics:
        goals=[target(row.store,metric,row) for _,row in frame.iterrows()]
        region.append((label,percentage(frame[metric].sum(),sum(goals)/days) if all(g is not None for g in goals) else None))
    ranking=[]
    for _,row in frame.iterrows():
        vals=[percentage(float(row[m]), (target(row.store,m,row) or 0)/days) for m,_ in metrics]
        ranking.append((str(row.store),vals))
    ranking.sort(key=lambda r: r[1][0] if r[1][0] is not None else -1,reverse=True)
    latest_keys=[k for k in frames if k[:7]==day.strftime('%Y-%m') and k<=key]
    latest=max(latest_keys) if latest_keys else key
    trend_day=datetime.strptime(latest,'%Y-%m-%d').date()
    trend_frame=frames[latest]
    trends=[]
    for metric,label in metrics:
        col='cumulative_'+metric
        actuals=trend_frame[col] if col in trend_frame else None
        goals=[target(row.store,metric,row) for _,row in trend_frame.iterrows()]
        if actuals is None or actuals.isna().any() or not all(g is not None for g in goals):
            trends.append((label,None,None,None))
            continue
        actual=float(actuals.sum()); goal=sum(goals)
        forecast=actual/trend_day.day*calendar.monthrange(trend_day.year,trend_day.month)[1]
        trends.append((label,percentage(forecast,goal),actual,forecast))
    focus=focus_data()
    if not any(k[1]==key for k in focus):
        raise ValueError('Немає фокусів за звітний день')
    W=1120; row_h=47; H=1320+len(ranking)*row_h
    im=Image.new('RGB',(W,H),'#eaf2fb'); d=ImageDraw.Draw(im)
    # Matplotlib bundles DejaVu TTF fonts with full Ukrainian Cyrillic support.
    # Install matplotlib in requirements.txt on Railway.
    from pathlib import Path as _FontPath
    try:
        import matplotlib
        _mpl_fonts = _FontPath(matplotlib.get_data_path()) / "fonts" / "ttf"
    except ImportError as exc:
        raise RuntimeError(
            "Для українських написів на картинці додай matplotlib>=3.8 "
            "до requirements.txt та повторно розгорни Railway."
        ) from exc

    def f(size, bold=False):
        name = "DejaVuSans-Bold.ttf" if bold else "DejaVuSans.ttf"
        path = _mpl_fonts / name
        if not path.is_file():
            raise RuntimeError(f"Не знайдено шрифт з підтримкою кирилиці: {path}")
        return ImageFont.truetype(str(path), size)

    def rect(x0,y0,x1,y1,fill,r=18): d.rounded_rectangle((x0,y0,x1,y1),radius=r,fill=fill)
    def txt(x,y,t,size=22,color='#172d4a',bold=False): d.text((x,y),str(t),font=f(size,bold),fill=color)
    def money(v): return '$'+format(round(v),',').replace(',',' ')
    def bar(x,y,width,val):
        rect(x,y,x+width,y+13,'#e0e9f1',6)
        if val is not None:
            color='#1bbf66' if val>=100 else '#ef535a' if val<90 else '#f0b52e'
            if val>0: rect(x,y,x+max(5,int(width*min(val,100)/100)),y+13,color,6)
    def pct(v): return 'н/д' if v is None else f'{v:.1f}%'
    rect(18,15,W-18,93,'#11375d',19)
    txt(45,34,'MOYO  |  РЕГІОН — ПІДСУМКИ ЗА ВЧОРА',29,'white',True)
    txt(875,42,day.strftime('%d.%m.%Y'),24,'white',True)
    rect(18,110,535,445,'white'); rect(550,110,1102,445,'white')
    txt(38,128,'ФІНАНСОВІ ПОКАЗНИКИ',24,bold=True)
    total_sales=float(frame['sales'].sum())
    for i,(metric,label) in enumerate([('sales','Оборот'),('direct','Прямі продажі'),('pickup','Пікап'),('accessories','Аксесуари'),('services','Послуги')]):
        y=175+i*51
        rect(37,y,515,y+43,'#eef4fa',9)
        txt(55,y+8,label,21)
        amount=money(frame[metric].sum())
        bbox=d.textbbox((0,0),amount,font=f(22,True))
        txt(499-(bbox[2]-bbox[0]),y+7,amount,22,bold=True)
        if metric in ('accessories','services'):
            share = float(frame[metric].sum())/total_sales*100 if total_sales else None
            share_label = f'Частка {share:.2f}%' if share is not None else 'Частка н/д'
            txt(280,y+13,share_label,15,'#42658b',True)
    txt(570,128,'ВИКОНАННЯ ДЕННОГО ПЛАНУ',23,bold=True)
    for i,(label,val) in enumerate(region):
        cx=645+i*176; cy=272
        color='#18b965' if val is not None and val>=100 else '#f0b52e' if val is not None and val>=90 else '#ef535a'
        d.ellipse((cx-64,cy-64,cx+64,cy+64),outline='#e2eaf5',width=15)
        if val is not None:
            d.arc((cx-64,cy-64,cx+64,cy+64),-90,-90+max(1,min(val,100)/100*359),fill=color,width=15)
        value=pct(val); bbox=d.textbbox((0,0),value,font=f(25,True))
        txt(cx-(bbox[2]-bbox[0])/2,cy-15,value,25,bold=True)
        bbox=d.textbbox((0,0),label,font=f(19,True))
        txt(cx-(bbox[2]-bbox[0])/2,cy+83,label,19,bold=True)
    table_top=462; table_end=table_top+96+len(ranking)*row_h
    rect(18,table_top,1102,table_end,'white')
    txt(40,table_top+14,'МАГАЗИНИ — ВИКОНАННЯ ПЛАНУ (%)',26,bold=True)
    head=table_top+58; rect(38,head,1082,head+42,'#203f6a',8)
    for x,t in [(55,'Магазин'),(395,'Оборот'),(620,'Аксесуари'),(867,'Послуги')]: txt(x,head+8,t,19,'white',True)
    for i,(store,vals) in enumerate(ranking):
        y=head+48+i*row_h
        if i%2==0: rect(39,y,1081,y+43,'#f0f5fb',6)
        txt(54,y+10,aliases.get(store,store.removeprefix('Київ-'))[:23],19,bold=True)
        for j,(x,v) in enumerate(zip((395,620,867),vals)):
            col='#11834a' if v is not None and v>=100 else '#d4424b' if v is not None and v<90 else '#bb8500'
            txt(x,y+9,pct(v),19,col,True)
            bar(x+105,y+16,100,v)
    # Top / bottom 3 stores by turnover daily plan attainment.
    y=table_end+16
    rect(18,y,1102,y+224,'white')
    txt(39,y+13,'ТОП / АНТИТОП МАГАЗИНІВ — ОБОРОТ',25,bold=True)
    txt(48,y+57,'ТОП 3',22,'#11834a',True)
    txt(580,y+57,'АНТИТОП 3',22,'#d4424b',True)
    eligible=[(store,vals[0]) for store,vals in ranking if vals[0] is not None]
    leaders=eligible[:3]
    laggards=list(reversed(eligible[-3:]))
    for i in range(3):
        yy=y+94+i*39
        for x,items,color in ((48,leaders,'#11834a'),(580,laggards,'#d4424b')):
            if i>=len(items):
                continue
            store,value=items[i]
            name=aliases.get(store,store.removeprefix('Київ-'))[:21]
            txt(x,yy,f'{i+1}. {name}',19,bold=True)
            txt(x+370,yy,pct(value),19,color,True)
    y+=240
    rect(18,y,1102,y+238,'white')
    txt(39,y+14,f'ТРЕНД РЕГІОНУ | 01.{trend_day:%m.%Y}–{trend_day:%d.%m.%Y}',25,bold=True)
    for i,(label,value,actual,forecast) in enumerate(trends):
        yy=y+59+i*53
        txt(42,yy,label,20,bold=True)
        txt(235,yy,pct(value),24,'#11834a' if value is not None and value>=100 else '#d4424b',True)
        bar(375,yy+11,200,value)
        txt(600,yy,'Факт: '+money(actual) if actual is not None else 'Немає даних',17)
        txt(835,yy,'Прогноз: '+money(forecast) if forecast is not None else '—',17)
    y+=254
    rect(18,y,1102,y+145,'white')
    txt(40,y+13,'ФОКУСИ ЗА ВЧОРА',25,bold=True)
    for i,category in enumerate(('ЕЗ','Клімат')):
        unit=FOCUS_CONFIG[category][1]
        txt(55,y+59+i*39,category+':  '+focus_pair(focus,category,unit,day,day),22,bold=True)
    buf=io.BytesIO(); im.crop((0,0,W,y+165)).save(buf,format='PNG',optimize=True)
    buf.seek(0); buf.name='moyo_digest.png'
    return buf


# Additional visual reports: completed week and current-month store trends.
def performance_dashboard(mode, reference_day):
    from PIL import Image, ImageDraw, ImageFont
    import matplotlib
    frames = daily_frames()
    metrics = [('sales', 'Оборот'), ('accessories', 'Аксесуари'), ('services', 'Послуги')]
    if mode == 'week':
        end = reference_day
        start = end - timedelta(days=6)
        dates = [(start + timedelta(days=i)).isoformat() for i in range(7)]
        missing = [x for x in dates if x not in frames]
        if missing:
            raise ValueError('Немає повних даних за тиждень: ' + ', '.join(missing))
        daily = [frames[x] for x in dates]
        stores = sorted(set().union(*(set(f.store) for f in daily)))
        records = []
        for store in stores:
            rows = [f[f.store == store] for f in daily]
            if any(len(r) != 1 for r in rows):
                raise ValueError(f'Неповний тиждень для {store}')
            actual = {m: sum(float(r.iloc[0][m]) for r in rows) for m, _ in metrics}
            goals = {}
            for m, _ in metrics:
                goal = 0.0
                for day_key in dates:
                    d = datetime.strptime(day_key, '%Y-%m-%d').date()
                    plans = load_monthly_plans(d.strftime('%Y-%m'))
                    if store not in plans.index or m not in plans.columns or pd.isna(plans.loc[store,m]):
                        goal = None
                        break
                    goal += float(plans.loc[store,m]) / calendar.monthrange(d.year,d.month)[1]
                goals[m] = goal
            records.append((store, actual, goals))
        title = 'ПІДСУМКИ ТИЖНЯ'
        period = f'{start:%d.%m.%Y} — {end:%d.%m.%Y}'
        descriptor = 'Виконання плану за 7 днів'
    else:
        month = reference_day.strftime('%Y-%m')
        keys = sorted(k for k in frames if k.startswith(month) and k <= reference_day.isoformat())
        if not keys:
            raise ValueError('Немає даних за поточний місяць')
        latest = datetime.strptime(keys[-1], '%Y-%m-%d').date()
        source = frames[keys[-1]]
        plans = load_monthly_plans(month)
        records = []
        for _, r in source.iterrows():
            store = str(r.store)
            actual = {m: float(r['cumulative_'+m]) for m,_ in metrics}
            goals = {m: float(plans.loc[store,m]) if store in plans.index and m in plans.columns and pd.notna(plans.loc[store,m]) else None for m,_ in metrics}
            records.append((store,actual,goals))
        title = 'МІСЯЧНИЙ ТРЕНД ПО МАГАЗИНАХ'
        period = f'01.{latest:%m.%Y} — {latest:%d.%m.%Y}'
        descriptor = f'Прогноз виконання на кінець місяця · {latest.day} днів факту'
    def percent(actual, goal):
        if goal is None or goal <= 0 or not math.isfinite(actual):
            return None
        if mode == 'month':
            actual = actual / latest.day * calendar.monthrange(latest.year,latest.month)[1]
        return actual / goal * 100
    ranking = [(store, [percent(actual[m],goals[m]) for m,_ in metrics], actual, goals)
               for store,actual,goals in records]
    ranking.sort(key=lambda x: x[1][0] if x[1][0] is not None else -1,reverse=True)
    totals = []
    for m,label in metrics:
        valid = all(a[m] is not None and g[m] is not None and math.isfinite(a[m]) for _,a,g in records)
        total_actual = sum(a[m] for _,a,g in records) if valid else None
        total_goal = sum(g[m] for _,a,g in records) if valid else None
        totals.append((label,percent(total_actual,total_goal) if valid else None,total_actual))
    W=1100
    H=445+len(ranking)*52+250
    im=Image.new('RGB',(W,H),'#eaf2fb')
    d=ImageDraw.Draw(im)
    font_dir=Path(matplotlib.get_data_path())/'fonts'/'ttf'
    def font(sz,bold=False):
        return ImageFont.truetype(str(font_dir/('DejaVuSans-Bold.ttf' if bold else 'DejaVuSans.ttf')),sz)
    def text(x,y,value,sz=19,color='#193451',bold=False):
        d.text((x,y),str(value),font=font(sz,bold),fill=color)
    def rect(x0,y0,x1,y1,color,r=12):
        d.rounded_rectangle((x0,y0,x1,y1),radius=r,fill=color)
    def pct(v):
        return 'н/д' if v is None else f'{v:.1f}%'
    def color(v):
        return '#7b8fa6' if v is None else '#17a85b' if v>=100 else '#f0b52e' if v>=90 else '#e44f59'
    def bar(x,y,w,v):
        rect(x,y,x+w,y+13,'#e1eaf2',6)
        if v is not None and v>0:
            rect(x,y,x+max(4,int(w*min(v,100)/100)),y+13,color(v),6)
    aliases={'Київ-Олімпійський':'Олімпійський','Київ-Аркадія':'Аркадія','Київ-Нью Вей':'Нью Вей','Київ-Хрещатик 21':'Хрещатик 21','Київ-Аладдін':'Аладін','Київ-Комод':'Комод','Буча-Рітейл Парк':'Буча','Київ-Квадрат':'Квадрат','Київ-Материк':'Материк','Київ-Ретровіль':'Ретровіль','Київ-Дрімтаун_new':'Дрім Таун','Київ-Лавіна Молл':'Лавіна','Київ-Республіка Парк':'Республіка Парк','Київ-Блокбастер':'Блокбастер'}
    rect(15,15,W-15,92,'#12375c',16)
    text(35,32,'MOYO | '+title,27,'white',True)
    rect(15,107,W-15,290,'white')
    text(35,122,period,24,bold=True)
    text(35,157,descriptor,17,'#60758b')
    for i,(label,value,actual) in enumerate(totals):
        x=35+i*350
        text(x,202,label,21,bold=True)
        text(x+145,202,pct(value),25,color(value),True)
        bar(x,252,300,value)
    table_top=308
    table_end=table_top+102+len(ranking)*52
    rect(15,table_top,W-15,table_end,'white')
    text(35,table_top+12,'МАГАЗИНИ — '+('ВИКОНАННЯ ТИЖНЯ' if mode=='week' else 'ПРОГНОЗ ВИКОНАННЯ МІСЯЦЯ'),22,bold=True)
    rect(32,table_top+48,W-32,table_top+86,'#23466f',6)
    for x,label in [(45,'Магазин'),(395,'Оборот'),(630,'Аксесуари'),(865,'Послуги')]:
        text(x,table_top+54,label,18,'white',True)
    for i,(store,vals,_,_) in enumerate(ranking):
        y=table_top+92+i*52
        if i%2==0: rect(33,y,W-33,y+45,'#eff4fa',5)
        text(45,y+10,aliases.get(store,store.removeprefix('Київ-'))[:22],18,bold=True)
        for x,v in zip((395,630,865),vals):
            text(x,y+10,pct(v),18,color(v),True)
            bar(x+110,y+18,92,v)
    bottom=table_end+18
    rect(15,bottom,W-15,bottom+207,'white')
    text(35,bottom+12,'ТОП / АНТИТОП — ЗА ОБОРОТОМ',24,bold=True)
    valid=[r for r in ranking if r[1][0] is not None]
    for x,items,label,c in [(40,valid[:3],'ТОП-3','#168d51'),(565,list(reversed(valid[-3:])),'АНТИТОП-3','#c9404b')]:
        text(x,bottom+54,label,21,c,True)
        for i,(store,vals,_,_) in enumerate(items):
            y=bottom+88+i*34
            text(x,y,f'{i+1}. {aliases.get(store,store.removeprefix("Київ-"))[:20]}',17,bold=True)
            text(x+365,y,pct(vals[0]),18,c,True)
    buf=io.BytesIO()
    im.crop((0,0,W,bottom+224)).save(buf,format='PNG',optimize=True)
    buf.seek(0)
    buf.name=f'moyo_{mode}.png'
    return buf

def extra_digest_db():
    conn=db()
    conn.execute('CREATE TABLE IF NOT EXISTS extra_digest_sent (kind TEXT, period TEXT, chat_id TEXT, PRIMARY KEY(kind,period,chat_id))')
    return conn

async def send_extra_digest(bot,chat_id,kind,day,force=False):
    period=(day-timedelta(days=6)).isoformat() if kind=='week' else day.isoformat()
    conn=extra_digest_db()
    try:
        sent=conn.execute('SELECT 1 FROM extra_digest_sent WHERE kind=? AND period=? AND chat_id=?',(kind,period,str(chat_id))).fetchone()
    finally:
        conn.close()
    if sent and not force:
        return False
    picture=await asyncio.to_thread(performance_dashboard,kind,day)
    caption=('📊 МОЙО — тижневий звіт' if kind=='week' else '📈 МОЙО — місячний тренд по магазинах')
    await bot.send_photo(chat_id=int(chat_id),photo=picture,caption=caption)
    conn=extra_digest_db()
    try:
        with conn:
            conn.execute('INSERT OR REPLACE INTO extra_digest_sent VALUES (?,?,?)',(kind,period,str(chat_id)))
    finally:
        conn.close()
    return True

async def group_week_test(update,context):
    if not is_owner(update): return
    chat_id=GROUP_CHAT_ID or str(update.effective_chat.id)
    today=datetime.now(REPORT_TZ).date()
    last_sunday=today-timedelta(days=today.weekday()+1)
    try:
        await asyncio.to_thread(sync_google_sheets)
        await send_extra_digest(context.bot,chat_id,'week',last_sunday,True)
        await update.message.reply_text('Тижневий тест надіслано.')
    except Exception as exc:
        log.exception('Weekly report test failed')
        await update.message.reply_text(f'Тижневий тест: {exc}')

async def group_month_test(update,context):
    if not is_owner(update): return
    chat_id=GROUP_CHAT_ID or str(update.effective_chat.id)
    today=datetime.now(REPORT_TZ).date()
    try:
        await asyncio.to_thread(sync_google_sheets)
        await send_extra_digest(context.bot,chat_id,'month',today-timedelta(days=1),True)
        await update.message.reply_text('Тест місячного тренду надіслано.')
    except Exception as exc:
        log.exception('Monthly trend test failed')
        await update.message.reply_text(f'Тест місячного тренду: {exc}')

async def extra_digest_loop(app,kind):
    # Week: Monday 12:10 for completed Mon-Sun; month: daily 12:15 for latest data.
    hour,minute=12,(10 if kind=='week' else 15)
    while True:
        now=datetime.now(REPORT_TZ)
        target=datetime.combine(now.date(),dt_time(hour,minute),REPORT_TZ)
        if kind=='week':
            target+=timedelta(days=(7-now.weekday())%7)
        if now>=target:
            target+=timedelta(days=7 if kind=='week' else 1)
        await asyncio.sleep(max(1,(target-now).total_seconds()))
        if not GROUP_CHAT_ID: continue
        today=datetime.now(REPORT_TZ).date()
        day=today-timedelta(days=1)
        if kind=='week':
            day=today-timedelta(days=today.weekday()+1)
        try:
            await asyncio.to_thread(sync_google_sheets)
            await send_extra_digest(app.bot,GROUP_CHAT_ID,kind,day)
        except Exception:
            log.exception('Scheduled %s report failed',kind)


def digest_db():
    conn = db()
    conn.execute("CREATE TABLE IF NOT EXISTS group_digest_sent (report_date TEXT PRIMARY KEY, chat_id TEXT NOT NULL)")
    return conn

async def send_digest(bot, chat_id, day, force=False):
    conn = digest_db()
    try:
        already = conn.execute("SELECT chat_id FROM group_digest_sent WHERE report_date = ?", (day.isoformat(),)).fetchone()
    finally:
        conn.close()
    if already and not force:
        return False
    picture = await asyncio.to_thread(digest_dashboard, day)
    await bot.send_photo(chat_id=int(chat_id), photo=picture,
                         caption=f"📊 МОЙО — звіт за {day:%d.%m.%Y} | тренд регіону")
    conn = digest_db()
    try:
        with conn:
            conn.execute("INSERT OR REPLACE INTO group_digest_sent VALUES (?, ?)", (day.isoformat(), str(chat_id)))
    finally:
        conn.close()
    return True

async def group_here(update, context):
    if not is_owner(update):
        return
    if update.effective_chat.type not in ("group", "supergroup"):
        await update.message.reply_text("Додай бота до групи та виконай /group_here там.")
        return
    await update.message.reply_text(
        f"ID цієї групи: {update.effective_chat.id}\n"
        "Додай у Railway змінну REPORT_GROUP_CHAT_ID з цим значенням. "
        "Після перезапуску бот відправлятиме звіт щодня о 12:00 за Києвом."
    )

async def group_test(update, context):
    if not is_owner(update):
        return
    chat_id = GROUP_CHAT_ID or (str(update.effective_chat.id)
              if update.effective_chat.type in ("group", "supergroup") else "")
    if not chat_id:
        await context.bot.send_message(chat_id=update.effective_chat.id, text="Спочатку додай REPORT_GROUP_CHAT_ID у Railway або виконай /group_test у групі.")
        return
    day = datetime.now(REPORT_TZ).date() - timedelta(days=1)
    try:
        await context.bot.send_message(chat_id=update.effective_chat.id, text="🔄 Оновлюю продажі, плани та фокуси з Google Sheets перед тестом…")
        await asyncio.to_thread(sync_google_sheets)
        await asyncio.to_thread(sync_focus_google)
        await send_digest(context.bot, chat_id, day, force=True)
        await context.bot.send_message(chat_id=update.effective_chat.id, text=f"Тестовий звіт за {day:%d.%m.%Y} надіслано.")
    except Exception as exc:
        log.exception("Test group digest failed")
        frames = daily_frames()
        available = sorted(frames)
        recent = ", ".join(available[-5:]) if available else "жодної"
        await context.bot.send_message(chat_id=update.effective_chat.id, text=
            f"Тест не надіслано: {exc}\nОстанні дати продажів у базі: {recent}.\n"
            "Перевір, що в таблицях Продажі та Якісні є вкладки з однаковою датою, "
            "а для розрахунку денного виторгу є також попередній день."
        )

async def digest_loop(app):
    while True:
        now = datetime.now(REPORT_TZ)
        target = datetime.combine(now.date(), dt_time(12, 0), REPORT_TZ)
        if now >= target:
            target += timedelta(days=1)
        await asyncio.sleep(max(1, (target - now).total_seconds()))
        if not GROUP_CHAT_ID:
            continue
        day = datetime.now(REPORT_TZ).date() - timedelta(days=1)
        try:
            # Refresh both sources before sending, avoid stale data.
            await asyncio.to_thread(sync_google_sheets)
            await asyncio.to_thread(sync_focus_google)
            await send_digest(app.bot, GROUP_CHAT_ID, day)
        except Exception:
            log.exception("Scheduled group digest failed; will not send incomplete report")

async def automatic_startup_test(app):
    """One-off delayed send to verify the Railway scheduler without Telegram commands."""
    if os.getenv("AUTO_TEST_ON_START", "").strip() != "1":
        return
    if not GROUP_CHAT_ID:
        log.error("AUTO_TEST_ON_START enabled, but REPORT_GROUP_CHAT_ID is missing")
        return
    log.info("AUTO_TEST_ON_START enabled: scheduled a one-off report in 300 seconds")
    await asyncio.sleep(300)
    try:
        day = datetime.now(REPORT_TZ).date() - timedelta(days=1)
        await asyncio.to_thread(sync_google_sheets)
        await asyncio.to_thread(sync_focus_google)
        await send_digest(app.bot, GROUP_CHAT_ID, day, force=True)
        log.info("AUTO TEST SUCCESS: photo sent to group %s for %s", GROUP_CHAT_ID, day)
    except Exception:
        log.exception("AUTO TEST FAILED: report not sent")

# ---- Interim 1C report (Google Sheets) ----
INTERIM_SHEET_ID = os.getenv('INTERIM_SHEET_ID', '18jazHrriQXyfQxbDohZJmHPl1l6XR9706e9tuHZ1vsc').strip()

def _interim_number(value):
    if value is None or str(value).strip() == '':
        return 0.0
    try:
        return float(str(value).replace('\u00a0', '').replace(' ', '').replace(',', '.').replace('₴', ''))
    except (ValueError, TypeError):
        raise ValueError(f'Некоректна сума в 1С: {value!r}')

def parse_interim_rows(rows, report_day):
    """Read ONLY store subtotals; receipt/date rows are deliberately ignored."""
    import hashlib
    header_text = '\n'.join(str(c) for row in rows[:8] for c in row)
    match = re.search(r'курс[^\n]*?\b(?:на\s+\d{2}\.\d{2}\.\d{4}\s*[-–:]\s*)(\d+[,.]\d+)', header_text, re.I)
    if not match:
        raise ValueError('У заголовку 1С не знайдено обліковий курс. Очікується «Учетный курс ... - 45,4».')
    rate = _interim_number(match.group(1))
    if not (0 < rate < 1000):
        raise ValueError('Некоректний курс 1С.')
    header_index = next((i for i, row in enumerate(rows) if any(str(x).strip() == 'Дата продажи' for x in row)
                         and any(str(x).strip() == 'Сумма всего ТО' for x in row)), None)
    if header_index is None:
        raise ValueError('Не знайдено заголовки «Дата продажи / Сумма всего ТО».')
    data = {}
    for row in rows[header_index + 1:]:
        if not row:
            continue
        name = str(row[0]).strip()
        if name not in STORE_ORDER:
            continue
        if name in data:
            raise ValueError(f'Повторний підсумковий рядок магазину {name}.')
        data[name] = {'sales': _interim_number(row[1] if len(row)>1 else None),
                      'accessories': _interim_number(row[2] if len(row)>2 else None),
                      'services': _interim_number(row[3] if len(row)>3 else None)}
    if not data:
        raise ValueError('Не знайдено жодного підсумкового рядка магазину.')
    # Hash of meaningful source values, independent of receipt count and cell formatting.
    canonical = json.dumps({'rate': rate, 'stores': data}, sort_keys=True, ensure_ascii=False)
    signature = hashlib.sha256(canonical.encode('utf-8')).hexdigest()
    return rate, data, signature

def fetch_interim_sheet(report_day):
    from google.oauth2.service_account import Credentials
    from google.auth.transport.requests import AuthorizedSession
    from urllib.parse import quote
    config = os.getenv('GOOGLE_SERVICE_ACCOUNT_JSON')
    if not config:
        raise ValueError('Немає GOOGLE_SERVICE_ACCOUNT_JSON у Railway.')
    credentials = Credentials.from_service_account_info(json.loads(config), scopes=['https://www.googleapis.com/auth/spreadsheets.readonly'])
    with AuthorizedSession(credentials) as session:
        metadata = session.get(f'https://sheets.googleapis.com/v4/spreadsheets/{INTERIM_SHEET_ID}',
                               params={'fields':'sheets.properties.title'},timeout=45)
        metadata.raise_for_status()
        title = sheet_dates(metadata.json()).get(report_day.isoformat())
        if not title:
            raise ValueError(f'Немає вкладки {report_day:%d.%m.%Y} у проміжковій таблиці.')
        address = quote("'" + title.replace("'", "''") + "'", safe='')
        response = session.get(f'https://sheets.googleapis.com/v4/spreadsheets/{INTERIM_SHEET_ID}/values/{address}',
                               params={'valueRenderOption':'FORMATTED_VALUE'},timeout=45)
        response.raise_for_status()
        return parse_interim_rows(response.json().get('values', []), report_day)

def interim_dashboard(report_day, slot, rate, data):
    """Render the approved dense MOYO dashboard: grouped plan/fact/%/share columns."""
    from PIL import Image, ImageDraw, ImageFont
    from matplotlib import font_manager
    W,H=1700,1900
    im=Image.new('RGB',(W,H),'#e7f2ff'); d=ImageDraw.Draw(im)
    regular=font_manager.findfont(font_manager.FontProperties(family='DejaVu Sans'))
    bold=font_manager.findfont(font_manager.FontProperties(family='DejaVu Sans',weight='bold'))
    fonts={}
    def ft(n,b=False):
        key=(n,b)
        if key not in fonts: fonts[key]=ImageFont.truetype(bold if b else regular,n)
        return fonts[key]
    def rect(x0,y0,x1,y1,fill,r=16,outline=None):
        d.rounded_rectangle((x0,y0,x1,y1),radius=r,fill=fill,outline=outline,width=2 if outline else 1)
    def txt(x,y,v,n=20,c='#173455',b=False,anchor=None):
        d.text((x,y),str(v),font=ft(n,b),fill=c,anchor=anchor)
    def money(v): return f'{v:,.0f}'.replace(',',' ')
    def pct(v): return '—' if v is None else f'{v:.1f}%'
    def color(v): return '#079e5c' if v is not None and v>=100 else '#bf8700' if v is not None and v>=70 else '#d62735'
    month_days=calendar.monthrange(report_day.year,report_day.month)[1]
    plans=load_monthly_plans(report_day.strftime('%Y-%m'))
    results=[]
    for store in STORE_ORDER:
        values=data.get(store,{'sales':0,'accessories':0,'services':0})
        targets={}
        for k in ('sales','accessories','services'):
            try:
                v=plans.loc[store,k]
                targets[k]=float(v)*rate/month_days if pd.notna(v) else None
            except (KeyError,TypeError,ValueError): targets[k]=None
        results.append((store,values,targets))
    totals={k:sum(v[k] for _,v,_ in results) for k in ('sales','accessories','services')}
    targets={k:sum(p[k] for _,_,p in results) if all(p[k] is not None for _,_,p in results) else None for k in totals}
    def perf(actual,plan): return 100*actual/plan if plan is not None and plan>0 else None
    navy='#10355e'; muted='#52677f'
    rect(14,14,W-14,145,navy,24)
    txt(40,27,'MOYO  |  ПРОМІЖКОВИЙ ЗВІТ',46,'white',True)
    txt(42,91,'РЕГІОН — ЦЕНТР    •    ДЕНЬ В ДЕНЬ',23,'#c7e1fa')
    txt(W-36,33,report_day.strftime('%d.%m.%Y'),29,'white',True,'ra')
    txt(W-36,80,f'Після {slot}:00   |   Курс {rate:g} грн/$',23,'#dcecff',True,'ra')
    cards=[('ОБОРОТ','sales','#216dc4'),('АКСЕСУАРИ','accessories','#0e9962'),('ПОСЛУГИ','services','#b63f62')]
    cw=(W-56)//3
    for i,(label,k,accent) in enumerate(cards):
        x=14+i*(cw+14); rect(x,160,x+cw,495,'white',21)
        rect(x+18,180,x+65,229,accent,12)
        txt(x+78,184,label,27,navy,True)
        txt(x+31,186,{'sales':'▣','accessories':'◇','services':'✦'}[k],27,'white',True)
        txt(x+20,239,money(totals[k])+' ₴',39,accent if k!='sales' else navy,True)
        p=perf(totals[k],targets[k]); rect(x+cw-161,178,x+cw-18,228,'#e1f4e9' if p is not None and p>=100 else '#e9f0fa',12)
        txt(x+cw-90,187,pct(p),25,color(p),True,'ma')
        txt(x+22,308,'План, грн',20,muted)
        txt(x+cw-20,308,money(targets[k]) if targets[k] is not None else '—',22,navy,True,'ra')
        d.line((x+20,341,x+cw-20,341),fill='#e4edf7',width=2)
        txt(x+22,350,'Факт, грн',20,muted)
        txt(x+cw-20,350,money(totals[k]),22,navy,True,'ra')
        d.line((x+20,386,x+cw-20,386),fill='#e4edf7',width=2)
        txt(x+22,397,'Виконання',20,muted)
        txt(x+cw-20,397,pct(p),22,color(p),True,'ra')
        rect(x+20,446,x+cw-20,461,'#dce7f3',8)
        if p is not None and p>0: rect(x+20,446,x+20+(cw-40)*min(p,100)/100,461,accent,8)
        if k!='sales':
            share=100*totals[k]/totals['sales'] if totals['sales'] else 0
            txt(x+cw-22,465,f'ЧАСТКА В ОБОРОТІ  {share:.2f}%',21,navy,True,'ra')
    rect(14,511,W-14,1505,'white',20)
    txt(36,526,'МАГАЗИНИ — ВИКОНАННЯ ПЛАНУ',34,navy,True)
    # Widths: store 248; turnover 405; accessories 505; services 505
    x0=30; boundaries=[x0,280,680,1182,1670]
    rect(30,587,1670,687,navy,12)
    txt(43,621,'Магазин',22,'white',True)
    for mid,name in ((480,'ОБОРОТ'),(931,'АКСЕСУАРИ'),(1426,'ПОСЛУГИ')):
        txt(mid,595,name,23,'white',True,'ma')
    columns=[(379,'План'),(497,'Факт'),(618,'%'),(789,'План'),(906,'Факт'),(1011,'%'),(1160,'Частка'),(1287,'План'),(1395,'Факт'),(1508,'%'),(1652,'Частка')]
    for x,label in columns: txt(x,645,label,18,'#dcecff',True,'ra')
    for x in boundaries[1:-1]: d.line((x,587,x,1490),fill='#d0dce9',width=2)
    short={'Буча-Рітейл Парк':'Буча-Рітейл','Київ-Аладдін':'Аладін','Київ-Аркадія':'Аркадія','Київ-Блокбастер':'Блокбастер','Київ-Дрімтаун_new':'Дрім Таун','Київ-Квадрат':'Квадрат','Київ-Комод':'Комод','Київ-Лавіна Молл':'Лавіна','Київ-Материк':'Материк','Київ-Нью Вей':'Нью Вей','Київ-Олімпійський':'Олімпійський','Київ-Республіка Парк':'Республіка Парк','Київ-Ретровіль':'Ретровіль','Київ-Хрещатик 21':'Хрещатик 21'}
    rowh=54
    for i,(store,vals,plan) in enumerate(results):
        y=690+i*rowh
        if i%2==0: rect(30,y,1670,y+rowh-2,'#eef5fd',5)
        txt(43,y+13,short.get(store,store),21,navy,True)
        for k,xp,xf,xe,xs in (('sales',379,497,527,None),('accessories',789,906,917,1160),('services',1287,1395,1414,1652)):
            pp=perf(vals[k],plan[k]); share=100*vals[k]/vals['sales'] if vals['sales'] else 0
            txt(xp,y+14,money(plan[k]) if plan[k] is not None else '—',19,navy,True,'ra')
            txt(xf,y+14,money(vals[k]),20,navy,True,'ra')
            if pp is not None:
                badge_x=xe-13
                rect(badge_x,y+8,badge_x+90,y+44,
                     '#dff7e9' if pp>=100 else '#fff1ce' if pp>=70 else '#ffe3e8',9)
                txt(badge_x+45,y+11,pct(pp),20,color(pp),True,'ma')
            if xs is not None: txt(xs,y+14,f'{share:.1f}%',19,navy,True,'ra')
    yy=690+14*rowh+5
    rect(30,yy,1670,yy+53,navy,8)
    txt(43,yy+13,'РАЗОМ',20,'white',True)
    for k,xp,xf,xe,xs in (('sales',379,497,527,None),('accessories',789,906,917,1160),('services',1287,1395,1414,1652)):
        txt(xp,yy+13,money(targets[k]) if targets[k] is not None else '—',19,'white',True,'ra')
        txt(xf,yy+13,money(totals[k]),20,'white',True,'ra')
        txt(xe+91,yy+13,pct(perf(totals[k],targets[k])),20,'#b9f2d6',True,'ra')
        if xs is not None: txt(xs,yy+13,f'{100*totals[k]/totals["sales"] if totals["sales"] else 0:.1f}%',19,'white',True,'ra')
    rect(14,1522,837,1816,'white',20); rect(851,1522,1686,1816,'white',20)
    rect(30,1539,822,1598,'#dcf6e9',13); rect(866,1539,1670,1598,'#fbe5e8',13)
    d.polygon([(52,1565),(64,1553),(64,1560),(79,1560),(79,1570),(64,1570),(64,1577)],fill='#07854e')
    txt(94,1551,'ТОП-3  •  ОБОРОТ',29,'#07854e',True)
    txt(885,1551,'↓ АНТИТОП-3  •  ОБОРОТ',29,'#c82e43',True)
    ranking=[(store,perf(vals['sales'],plan['sales']),vals['sales']) for store,vals,plan in results]
    ranking=[v for v in ranking if v[1] is not None]
    ranking.sort(key=lambda z:z[1],reverse=True)
    if not ranking:
        txt(55,1650,'Плани ще не завантажені',22,muted)
        txt(885,1650,'Плани ще не завантажені',22,muted)
    else:
        for x,group in ((50,ranking[:3]),(870,list(reversed(ranking[-3:])))):
            for i,(store,pc,actual) in enumerate(group):
                y=1620+i*60
                txt(x,y,f'{i+1}. {short.get(store,store)}',20,navy,True)
                txt(x+530,y,pct(pc),21,color(pc),True)
                txt(x+745,y,money(actual)+' ₴',17,muted,False,'ra')
                if i<2: d.line((x,y+46,x+765,y+46),fill='#e5edf6',width=2)
    txt(26,1844,'Дані 1С  •  Плани в грн за обліковим курсом  •  Повернення враховано в обороті',18,muted)
    from PIL import ImageEnhance, ImageFilter
    im=ImageEnhance.Contrast(im).enhance(1.07)
    im=im.filter(ImageFilter.UnsharpMask(radius=1.2,percent=145,threshold=2))
    buf=io.BytesIO(); im.save(buf,format='PNG',optimize=True); buf.seek(0)
    buf.name=f'moyo_interim_{report_day:%Y%m%d}_{slot}.png'
    return buf

def interim_db():
    conn=db()
    conn.execute('CREATE TABLE IF NOT EXISTS interim_sent (day TEXT, slot INTEGER, chat_id TEXT, signature TEXT, PRIMARY KEY(day,slot,chat_id))')
    return conn

async def interim_test(update, context):
    """Manual preview for today; never marks either scheduled slot as sent."""
    if not is_owner(update):
        return
    day = datetime.now(REPORT_TZ).date()
    chat_id = update.effective_chat.id
    try:
        await context.bot.send_message(
            chat_id=chat_id,
            text=f"🔄 Читаю проміжкові продажі 1С за {day:%d.%m.%Y}…"
        )
        # Refresh monthly plans, so the image uses the newest plan figures.
        await asyncio.to_thread(sync_google_sheets)
        rate, data, _ = await asyncio.to_thread(fetch_interim_sheet, day)
        slot = 21 if datetime.now(REPORT_TZ).hour >= 21 else 18 if datetime.now(REPORT_TZ).hour >= 18 else 15
        picture = await asyncio.to_thread(interim_dashboard, day, slot, rate, data)
        await context.bot.send_photo(
            chat_id=chat_id,
            photo=picture,
            caption=f"🧪 ТЕСТ • МОЙО — проміжковий звіт за {day:%d.%m.%Y} "
                    f"(курс {rate:g} грн/$). Автоматичні відправки не змінено."
        )
    except Exception as exc:
        log.exception("Manual interim test failed")
        await context.bot.send_message(
            chat_id=chat_id,
            text=f"❌ Проміжковий тест не сформовано: {exc}"
        )


async def interim_loop(app):
    """Poll once per minute for 15:00, 18:00, and 21:00 slots."""
    previous=None
    while True:
        await asyncio.sleep(60)
        if not GROUP_CHAT_ID:
            continue
        now=datetime.now(REPORT_TZ)
        slot=21 if now.hour>=21 else 18 if now.hour>=18 else 15 if now.hour>=15 else None
        if slot is None:
            previous=None
            continue
        day=now.date()
        conn=interim_db()
        try:
            already=conn.execute('SELECT signature FROM interim_sent WHERE day=? AND slot=? AND chat_id=?',
                                 (day.isoformat(),slot,str(GROUP_CHAT_ID))).fetchone()
            earlier=conn.execute('SELECT signature FROM interim_sent WHERE day=? AND slot=? AND chat_id=?',
                                 (day.isoformat(),15 if slot==18 else 18,str(GROUP_CHAT_ID))).fetchone() if slot in (18,21) else None
            if slot==21 and not earlier:
                earlier=conn.execute('SELECT signature FROM interim_sent WHERE day=? AND slot=15 AND chat_id=?',
                                     (day.isoformat(),str(GROUP_CHAT_ID))).fetchone()
        finally: conn.close()
        if already:
            continue
        try:
            rate,data,signature=await asyncio.to_thread(fetch_interim_sheet,day)
            if earlier and earlier[0]==signature:
                continue # No new 1C numbers since afternoon; never resend stale data.
            key=(day.isoformat(),slot,signature)
            if previous!=key:
                previous=key
                continue # wait for another poll to ensure Google Sheet upload finished
            picture=await asyncio.to_thread(interim_dashboard,day,slot,rate,data)
            await app.bot.send_photo(chat_id=int(GROUP_CHAT_ID),photo=picture,
                 caption=f'📊 МОЙО — проміжковий звіт за {day:%d.%m.%Y} (після {slot}:00)')
            conn=interim_db()
            try:
                with conn: conn.execute('INSERT OR REPLACE INTO interim_sent VALUES (?,?,?,?)',
                                       (day.isoformat(),slot,str(GROUP_CHAT_ID),signature))
            finally: conn.close()
        except Exception:
            log.exception('Interim report poll failed; will retry')


async def setup_group_schedule(app):
    await setup_commands(app)
    app.create_task(digest_loop(app), name="group-digest-12-kyiv")
    if os.getenv("AUTO_TEST_ON_START", "").strip() == "1":
        app.create_task(automatic_startup_test(app), name="automatic-startup-test-5min")
    app.create_task(extra_digest_loop(app,"week"), name="weekly-digest-kyiv")
    app.create_task(extra_digest_loop(app,"month"), name="monthly-trend-kyiv")
    app.create_task(interim_loop(app), name="interim-sales-watch-kyiv")


def main():
    if not TOKEN:
        raise RuntimeError("BOT_TOKEN environment variable is missing")
    if not OWNER_ID:
        raise RuntimeError("OWNER_ID environment variable is missing")
    app = Application.builder().token(TOKEN).post_init(setup_group_schedule).build()
    conv = ConversationHandler(
        entry_points=[MessageHandler(filters.Document.ALL, handle_document)],
        states={DATE_WAIT: [MessageHandler(filters.TEXT & ~filters.COMMAND, date_received)]},
        fallbacks=[CommandHandler("cancel", cancel)],
        allow_reentry=True,
    )
    app.add_handler(CommandHandler("group_here", group_here))
    app.add_handler(CommandHandler("group_test", group_test))
    app.add_handler(CommandHandler("interim_test", interim_test))
    app.add_handler(CommandHandler("group_week_test", group_week_test))
    app.add_handler(CommandHandler("group_month_test", group_month_test))
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("help", help_cmd))
    app.add_handler(CommandHandler("status", status))
    app.add_handler(CommandHandler("sync", sync_cmd, block=False))
    app.add_handler(CommandHandler("focus", focus_cmd))
    app.add_handler(CallbackQueryHandler(focus_selected, pattern=r"^focus:(menu|back|dates|weeks|date|week|mtd):"))
    app.add_handler(CommandHandler("day", day_cmd))
    app.add_handler(CallbackQueryHandler(day_selected, pattern=r"^day(page|date|pick):"))
    app.add_handler(CommandHandler("weeks", weeks_cmd))
    app.add_handler(CallbackQueryHandler(weeks_selected, pattern=r"^ws(page|date|pick):"))
    app.add_handler(CommandHandler("week", week_cmd))
    app.add_handler(CommandHandler("store", store_cmd))
    app.add_handler(CommandHandler("report", report_cmd))
    app.add_handler(CommandHandler("dates", dates_cmd))
    app.add_handler(CallbackQueryHandler(store_selected, pattern=r"^pick:(store|week):"))
    app.add_handler(conv)
    app.add_handler(CommandHandler("cancel", cancel))
    app.run_polling()

if __name__ == "__main__":
    main()
