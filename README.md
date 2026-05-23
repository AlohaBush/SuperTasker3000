# Telegram Task Planner Bot

A lightweight personal Telegram bot that uses an Excel workbook as a task planner.

The bot keeps tasks and lets you manage them from Telegram: add tasks, list tasks, move due dates, rename tasks, mark tasks as completed, delete tasks after confirmation, and receive digests.

This repository is prepared for public GitHub use. It must not contain real Telegram tokens, LLM API keys, private chat IDs, runtime state, or personal task data.

For now Bot undertands instructions only on Russian.  

<img width="1000" height="1000" alt="ChatGPT" src="https://github.com/user-attachments/assets/2d5deec7-4fa5-4ca9-8a5a-4604450c4e22" />


## What the bot can do

- Add tasks to planner.
- Show all tasks with row numbers.
- Show tasks for today, overdue tasks, upcoming tasks, nearest future tasks, and a full summary etc.
- Show tasks for a specific date or date range.
- Delete tasks by row number or by matching task text.
- Move tasks to another date.
- Rename tasks.
- Mark tasks as completed.
- Ask for confirmation before potentially destructive changes.
- Restrict access by Telegram Chat ID.

The built-in command and date parser is mainly optimized for Russian phrases such as `сегодня`, `завтра`, `перенеси`, `удали`, `отметь выполненной`. Slash commands work without an LLM.

## Repository files

| File | Purpose |
| --- | --- |
| `telegram_task_bot.py` | Main bot script. |
| `Task_Planner.xlsx` | Empty Excel task planner template. |
| `telegram_config.example.json` | Safe configuration template. Copy it to `telegram_config.json` before running the bot. |
| `telegram_config.sample.json` | Same safe sample config, kept for convenience. |
| `telegram_task_bot_state.example.json` | Example runtime state file. Do not use it as your real state file unless you know why. |
| `run_telegram_task_bot.bat` | Windows launcher. |
| `run_telegram_task_bot.sh` | Linux/macOS launcher. |
| `requirements.txt` | Notes that no mandatory third-party Python packages are required. |
| `.gitignore` | Prevents committing local config, state, cache, logs and archives. |
| `SECURITY.md` | Security notes. |
| `LICENSE` | Project license. |

## Requirements

- Python 3.10 or newer.
- A Telegram bot token from `@BotFather`.
- `ffmpeg` for Telegram voice messages. Telegram voice messages are usually OGG/Opus, and the bot converts them before transcription.
LLM API key for free-form natural language parsing, task-table questions, and voice transcription.

The script uses only the Python standard library. There are no mandatory Python packages to install.

## Quick start

### 1. Clone the repository

```bash
git clone https://github.com/YOUR_USERNAME/telegram-task-planner-bot.git
cd telegram-task-planner-bot
```

### 2. Create your local config

Copy the example config:

```bash
cp telegram_config.example.json telegram_config.json
```

Windows PowerShell:

```powershell
Copy-Item telegram_config.example.json telegram_config.json
```

Do not commit `telegram_config.json`. It is ignored by `.gitignore` because it contains private credentials.

### 3. Create a Telegram bot

1. Open Telegram.
2. Find `@BotFather`.
3. Send `/newbot`.
4. Follow BotFather’s instructions.
5. Copy the generated bot token.
6. Paste it into `telegram_config.json`.

Initial safe config:

```json
{
  "bot_token": "PASTE_TELEGRAM_BOT_TOKEN_HERE",
  "allowed_chat_ids": [],
  "notification_chat_id": "",
  "planner_file": "Task_Planner.xlsx",
  "soon_days": 7,
  "daily_digest_time": "09:00",
  "llm_api_key": "",
  "llm_base_url": "https://api.openai.com/v1",
  "llm_transcription_model": "gpt-4o-mini-transcribe",
  "llm_intent_model": "gpt-4o-mini",
  "voice_language": "ru"
}
```

When `allowed_chat_ids` is empty, the bot is locked. It will not expose your planner to anyone. It only replies to `/start` or `/help` with the sender’s Chat ID so that you can configure access.

### 4. Run the bot for the first time

Windows:

```powershell
python telegram_task_bot.py
```

or double-click:

```text
run_telegram_task_bot.bat
```

Linux/macOS:

```bash
python3 telegram_task_bot.py
```

Then send `/start` to your Telegram bot. The bot will reply with your Chat ID.

### 5. Allow your own Chat ID

Put your Chat ID into `telegram_config.json`:

```json
{
  "bot_token": "PASTE_TELEGRAM_BOT_TOKEN_HERE",
  "allowed_chat_ids": [123456789],
  "notification_chat_id": "123456789",
  "planner_file": "Task_Planner.xlsx",
  "soon_days": 7,
  "daily_digest_time": "09:00",
  "llm_api_key": "",
  "llm_base_url": "https://api.openai.com/v1",
  "llm_transcription_model": "gpt-4o-mini-transcribe",
  "llm_intent_model": "gpt-4o-mini",
  "voice_language": "ru"
}
```

Restart the bot.

From this point, the bot will respond only to chat IDs listed in `allowed_chat_ids`.

## Configuration reference

| Key | Required | Description |
| --- | --- | --- |
| `bot_token` | Yes | Telegram bot token from BotFather. |
| `allowed_chat_ids` | Yes | List of Telegram Chat IDs allowed to use the bot. Keep empty only during first setup. |
| `notification_chat_id` | No | Chat ID that receives the daily digest. Usually the same as your own Chat ID. |
| `planner_file` | Yes | Path to the Excel workbook. Default: `Task_Planner.xlsx`. |
| `soon_days` | No | Number of days used by the `/soon` command. Default: `7`. |
| `daily_digest_time` | No | Daily digest time in `HH:MM` format. Leave empty to disable. |
| `llm_api_key` | No | API key for the LLM provider. Required only for AI features and voice transcription. |
| `llm_base_url` | No | Base URL for the LLM API. Default: `https://api.openai.com/v1`. |
| `llm_transcription_model` | No | Model used for audio transcription. |
| `llm_intent_model` | No | Model used for natural-language task parsing and table questions. |
| `voice_language` | No | Voice transcription language hint. Default: `ru`. |

The script also supports environment variables:

| Environment variable | Config equivalent |
| --- | --- |
| `TELEGRAM_BOT_TOKEN` | `bot_token` |
| `TELEGRAM_ALLOWED_CHAT_IDS` | `allowed_chat_ids`, comma-separated |
| `TELEGRAM_CHAT_ID` | `notification_chat_id` |
| `LLM_API_KEY` | `llm_api_key` |
| `LLM_BASE_URL` | `llm_base_url` |
| `LLM_TRANSCRIPTION_MODEL` | `llm_transcription_model` |
| `LLM_INTENT_MODEL` | `llm_intent_model` |
| `OPENAI_API_KEY` | fallback for `llm_api_key` |

For backward compatibility, the script also accepts legacy config keys such as `openai_api_key`, `openai_base_url`, `openai_intent_model`, and `openai_transcription_model`.

```

CLI mode without Telegram:

```bash
python telegram_task_bot.py --check
python telegram_task_bot.py --add "Pay invoice" "25.05.2026"
python telegram_task_bot.py --delete 4
python telegram_task_bot.py --config path/to/telegram_config.json
```

## Russian free-form examples

The script has many built-in Russian rules, and AI mode improves this further.

```text
добавь оплатить интернет на завтра
напомни 25 мая подготовить отчет
добавь плавание на 12 мая и бег на 14 мая
покажи список задач
что у меня сегодня
какие задачи на 12.05.2026
покажи задачи с понедельника по среду
покажи просроченные
что у меня горит
перенеси задачу 4 на 15 мая
перенеси бег и плавание на 14 мая
переименуй задачу 4 в новое название
задачу 4 запиши так - новое название
отметь задачи 2 и 3 выполненными
удали задачу номер 4
удали задачи 2 3 4
```

For deletion, moving, renaming and multi-task completion, the bot normally asks for confirmation before changing the workbook.

You can cancel pending confirmations or clarifications with:

```text
стоп
отмена
cancel
```

## Date formats

The bot understands:

```text
25.05.2026
25.05.26
2026-05-25
25/05/2026
сегодня
завтра
послезавтра
понедельник
следующая среда
на следующей неделе
12 мая
двенадцатого мая
```

English `today` and `tomorrow` are supported by the basic parser, but the full free-form parser is Russian-oriented.

## LLM support

The bot can work without an LLM. Without an LLM, it still supports slash commands and many built-in Russian text commands.

LLM mode is used for:

- better natural-language intent detection;
- matching task descriptions to existing rows;
- answering questions based on the current task table;
- voice transcription.

The script calls:

```text
POST {llm_base_url}/responses
POST {llm_base_url}/audio/transcriptions
```

Therefore, you can connect any LLM provider that is compatible with these OpenAI-style endpoints and response formats.

For the default OpenAI setup:

```json
{
  "llm_api_key": "YOUR_API_KEY",
  "llm_base_url": "https://api.openai.com/v1",
  "llm_intent_model": "gpt-4o-mini",
  "llm_transcription_model": "gpt-4o-mini-transcribe"
}
```

If your provider uses another API format, adapt these functions in `telegram_task_bot.py`:

```text
openai_json_response(...)
openai_text_response(...)
transcribe_audio(...)
extract_response_text(...)
```

## Voice messages

To use voice messages:

1. Set `llm_api_key`.
2. Make sure the configured transcription endpoint supports `{llm_base_url}/audio/transcriptions`.
3. Install `ffmpeg` if the bot reports that Telegram OGG/Opus conversion is required.

Examples:

```text
добавь задачу оплатить интернет на завтра
покажи список задач
удали задачу номер четыре
перенеси задачу четыре на завтра
отметь задачи два и три выполненными
```

## Excel planner format

The included `Task_Planner.xlsx` template is expected by the script.

The main fields used by the script are:

| Column | Meaning |
| --- | --- |
| A | Task name |
| B | Due date |
| C | Formula-based status |
| D | Formula-based days remaining |
| E | Formula-based short status |
| F | Completed flag |
| G | Completed date |

The script writes directly into the XLSX file by editing its XML. Keep the workbook structure compatible with the template.

## Daily digest

Set:

```json
"notification_chat_id": "123456789",
"daily_digest_time": "09:00"
```

The bot checks the local machine time. When the configured time is reached, it sends one summary per day and stores the date in `telegram_task_bot_state.json`.

To disable the digest, leave `notification_chat_id` or `daily_digest_time` empty.

## Running continuously

### Windows Task Scheduler

1. Open Windows Task Scheduler.
2. Create a new task.
3. Trigger: at logon or at system startup.
4. Action: start `run_telegram_task_bot.bat`.
5. Make sure the working directory is the repository folder.

### Linux systemd example

Create `/etc/systemd/system/telegram-task-planner-bot.service`:

```ini
[Unit]
Description=Telegram Task Planner Bot
After=network-online.target

[Service]
WorkingDirectory=/opt/telegram-task-planner-bot
ExecStart=/usr/bin/python3 /opt/telegram-task-planner-bot/telegram_task_bot.py
Restart=always
RestartSec=5
Environment=TELEGRAM_BOT_TOKEN=PASTE_TOKEN_HERE
Environment=TELEGRAM_ALLOWED_CHAT_IDS=123456789

[Install]
WantedBy=multi-user.target
```

Then run:

```bash
sudo systemctl daemon-reload
sudo systemctl enable telegram-task-planner-bot
sudo systemctl start telegram-task-planner-bot
sudo systemctl status telegram-task-planner-bot
```

For production, prefer an environment file with restricted permissions instead of putting secrets directly into the service file.

## Security checklist before publishing to GitHub

Before pushing the repository:

- Do not commit `telegram_config.json`.
- Do not commit `telegram_task_bot_state.json`.
- Do not commit a personal `Task_Planner.xlsx` with real tasks.
- Do not commit `.env` files.
- Do not commit Telegram bot tokens.
- Do not commit LLM/OpenAI/API keys.
- Check Git history if secrets were ever committed.
- If a secret was exposed, revoke it and create a new one.

Recommended command before publishing:

```bash
git status
```

Optional manual check:

```bash
grep -R "sk-" .
grep -R "bot_token" .
grep -R "api_key" .
```

On Windows PowerShell:

```powershell
Select-String -Path * -Pattern "sk-","bot_token","api_key" -Recurse
```

## Troubleshooting

### The bot does not answer

Check:

- The script is still running.
- The Telegram token is correct.
- Your Chat ID is listed in `allowed_chat_ids`.
- You restarted the bot after editing `telegram_config.json`.

### The bot replies with a Chat ID but does not show tasks

This is normal when `allowed_chat_ids` is empty. Add your Chat ID to `allowed_chat_ids` and restart the bot.

### Voice messages do not work

Check:

- `llm_api_key` is configured.
- The transcription model is supported by your LLM provider.
- `ffmpeg` is installed and available in `PATH`.

### Free-form text is not understood

The bot works best with Russian task commands. For more flexible understanding, enable LLM mode by setting `llm_api_key`.

### Excel file errors

Use the included `Task_Planner.xlsx` template. The script expects the worksheet structure and columns to match the template.

## License

This project is released under the license included in the `LICENSE` file.
