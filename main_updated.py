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
from datetime import datetime, timedelta
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
        await update.message.reply_text(report[offset:offset + 4000])

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
    await update.message.reply_text(weekly_report(end, store))


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
        await query.message.reply_text(report[offset:offset+3500])




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
        await update.message.reply_text(result[offset:offset+3500])
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
            await query.message.reply_text(report[offset:offset+3500])


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
                        if not math.isfinite(amount) or amount < 0:
                            raise ValueError(f"Фокуси: {title}, {date} — перевір суму.")
                        amount_identity = (category + " · сума", date)
                        if amount_identity in result:
                            raise ValueError(f"Фокуси: повтор суми {category} за {date}.")
                        result[amount_identity] = (date, category + " · сума", "$", amount)
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
    return fmt(value) if unit == "$" else f"{value:,.0f} шт.".replace(",", " ")


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
    """Quantity and, when present for the requested period, adjacent monetary amount."""
    metrics = [(category, unit, "Кількість" if unit != "$" else "Сума")]
    amount_category = category + " · сума"
    end = end or start
    if unit != "$" and any(
        (amount_category, (start + timedelta(days=i)).isoformat()) in data
        for i in range((end-start).days + 1)
    ):
        metrics.append((amount_category, "$", "Сума"))
    return metrics

def focus_date_report(data, day, selected=None):
    previous_month = day.replace(day=1) - timedelta(days=1)
    lines = [f"🎯 Фокуси + категорії — РЕГІОН, {day:%d.%m.%Y}",
             "За обрану дату"]
    for category, unit, _ in focus_categories(selected):
        lines.extend(["", category])
        for metric, metric_unit, label_metric in focus_metrics(category, unit, data, day):
            daily = data.get((metric, day.isoformat()))
            current = daily[1] if daily is not None else None
            lines.append(f"{'Факт' if label_metric == 'Кількість' else 'Сума'} - {focus_value(current, metric_unit)}")
            for label, year, month in (("місяць до місяця", previous_month.year, previous_month.month),
                                       ("Рік до Року", day.year - 1, day.month)):
                comparison = focus_comparison_date(day, year, month)
                if comparison is None:
                    lines.append(f"  {label}: аналогічної календарної дати немає")
                    continue
                old_row = data.get((metric, comparison.isoformat()))
                old = old_row[1] if old_row is not None else None
                lines.append(f"{label}: " +
                             focus_change(old, current, metric_unit))
    return "\n".join(lines)


def focus_mtd_report(data, day, selected=None):
    """Month-to-date through the latest known date, compared over equal day counts."""
    previous_month = day.replace(day=1) - timedelta(days=1)
    lines = ["🎯 Фокуси + категорії — РЕГІОН · Поточний місяць",
             f"Поточний період: {day.replace(day=1):%d.%m.%Y}–{day:%d.%m.%Y}",
             "Порівняння накопичувальних показників за однакові календарні дні."]
    for category, unit, _ in focus_categories(selected):
        lines.extend(["", category])
        for metric, metric_unit, label_metric in focus_metrics(category, unit, data, day.replace(day=1), day):
            current = focus_total(data, metric, day.replace(day=1), day)
            lines.append(f"{label_metric} за поточний місяць: {focus_value(current, metric_unit)}")
            for label, year, month in (("місяць до місяця", previous_month.year, previous_month.month),
                                       ("Рік до Року", day.year - 1, day.month)):
                if day.day > calendar.monthrange(year, month)[1]:
                    lines.append(f"  {label}: немає рівного календарного періоду в {month:02d}.{year}")
                    continue
                old_end = day.replace(year=year, month=month)
                old = focus_total(data, metric, old_end.replace(day=1), old_end)
                lines.append(f"  {label} (01.{month:02d}–{old_end:%d.%m.%Y} → 01.{day.month:02d}–{day:%d.%m.%Y}): " +
                             focus_change(old, current, metric_unit))
    return "\n".join(lines)


def focus_week_report(data, monday, selected=None):
    latest = datetime.strptime(max(day for _, day in data), "%Y-%m-%d").date()
    end = min(monday+timedelta(days=6), latest)
    if end < monday:
        return "За цей тиждень немає даних."
    previous_start, previous_end = monday-timedelta(days=7), end-timedelta(days=7)
    lines = ["🎯 Фокуси + категорії — РЕГІОН", week_title(monday),
             f"Порівняння: {previous_start:%d.%m.%Y}–{previous_end:%d.%m.%Y} → {monday:%d.%m.%Y}–{end:%d.%m.%Y}"]
    if (end-monday).days < 6:
        lines.append("Неповний тиждень: порівнюються однакові дні тижня.")
    for category, unit, _ in focus_categories(selected):
        lines.extend(["", category])
        for metric, metric_unit, label_metric in focus_metrics(category, unit, data, monday, end):
            old = focus_total(data, metric, previous_start, previous_end)
            new = focus_total(data, metric, monday, end)
            lines.extend([f"{label_metric}: факт {focus_value(new, metric_unit)}",
                          focus_change(old, new, metric_unit)])
    return "\n".join(lines)


async def focus_send(message, report):
    # Split on lines so one category comparison is not cut in half.
    chunk = ""
    for line in report.splitlines():
        if len(chunk)+len(line)+1 > 3500:
            await message.reply_text(chunk)
            chunk = ""
        chunk += line + "\n"
    if chunk:
        await message.reply_text(chunk)


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
        await update.message.reply_text(result[offset:offset+3500])


async def setup_commands(app):
    await app.bot.set_my_commands([
        BotCommand("focus", "🎯 Фокуси + категорії"),
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
                await update.message.reply_text(result[offset:offset+3500])
            return ConversationHandler.END
        current = read_report(raw, filename)
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
