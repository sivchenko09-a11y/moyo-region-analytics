# MOYO Regional Analytics Telegram Bot

Telegram bot for receiving cumulative Power BI reports (XLSX/XLS/CSV), storing snapshots and calculating daily deltas.

## Required environment variable
- `BOT_TOKEN` — token from Telegram BotFather. Never commit it to GitHub.
- `OWNER_ID` — your numeric Telegram user ID; only this account can use the bot.


## Run locally
```bash
pip install -r requirements.txt
set BOT_TOKEN=your_token
python main.py
```

## Deploy on Railway
1. Create a GitHub repository and upload all files from this folder.
2. In Railway choose **New Project → GitHub Repository** and select this repository.
3. Add a persistent Railway Volume mounted at `/data`.
4. In the service Variables add `BOT_TOKEN` with the token from BotFather and `OWNER_ID` with your numeric Telegram user ID.
5. Deploy. The `Procfile` starts the bot as a worker.

## Daily use
Send a cumulative report from the 1st of the month as an Excel/CSV document.
Add the report date in the caption, e.g. `16.11.2026`.
If no caption date is provided, the bot asks for it in the next message.

The first report is saved as the baseline. Later reports are compared with the latest earlier report.
The bot calculates deltas for amounts, not percentage columns. Monthly plan is not added across days.

## Expected columns
`Склад`, `Продажи`, `ПродажиPY`, `Прямые Продажи`, `Пикап`, `Услуги, $`, `Акс, $`, `План`

Column names may be in Russian or Ukrainian variants supported in `main.py`.
