# Beginner Installation Guide

This guide explains how to install and run the Telegram Task Planner Bot if you have never used Python, Telegram bots, or GitHub projects before.

## 1. What this bot does

This is a personal Telegram bot for managing tasks in an Excel file.

You write messages to the bot in Telegram, and the bot updates or reads `Task_Planner.xlsx`.

The bot can:

- add new tasks to Excel;
- show all tasks;
- show tasks for today;
- show overdue tasks;
- show upcoming tasks;
- show a daily summary;
- delete tasks after confirmation;
- move tasks to another date;
- rename tasks;
- mark tasks as completed;
- send a daily digest at a selected time;
- work only with your own Telegram Chat ID;
- optionally understand normal text and voice messages with an LLM.

The bot is mainly optimized for Russian natural-language commands, but basic slash commands also work.

Examples:

```text
/add Pay invoice | 25.05.2026
/list
/today
/overdue
/summary
```

Russian examples:

```text
добавь оплатить интернет на завтра
что у меня сегодня
покажи просроченные
перенеси задачу 4 на 15 мая
отметь задачу 4 выполненной
удали задачу номер 4
```

## 2. What you need before installation

You need:

1. A computer or VPS.
2. Python 3.10 or newer.
3. Telegram installed on your phone or computer.
4. A Telegram bot token from `@BotFather`.
5. This project folder.

Optional:

- an LLM API key if you want better natural-language understanding;
- `ffmpeg` if you want Telegram voice messages to work.

You do not need to install any mandatory Python libraries. The bot uses the Python standard library.

## 3. Download the project

### Simple way

1. Open the GitHub repository page.
2. Click **Code**.
3. Click **Download ZIP**.
4. Unzip the downloaded archive.
5. Open the project folder.

### Git way

```bash
git clone https://github.com/YOUR_USERNAME/telegram-task-planner-bot.git
cd telegram-task-planner-bot
```

## 4. Install Python

Download Python from the official Python website.

During installation on Windows, enable this option:

```text
Add Python to PATH
```

After installation, check Python:

```bash
python --version
```

On Linux or macOS, use:

```bash
python3 --version
```

You should see something like:

```text
Python 3.10.x
```

or newer.

## 5. Create your Telegram bot

1. Open Telegram.
2. Search for `@BotFather`.
3. Send this command:

```text
/newbot
```

4. Follow BotFather instructions.
5. BotFather will give you a bot token.

The token looks similar to this:

```text
1234567890:ABCDEF_your_token_here
```

Keep this token private. Do not publish it on GitHub.

## 6. Create your local config file

In the project folder, find this file:

```text
telegram_config.example.json
```

Make a copy of it and rename the copy to:

```text
telegram_config.json
```

On Windows, you can do it manually in File Explorer.

On Windows PowerShell:

```powershell
Copy-Item telegram_config.example.json telegram_config.json
```

On Linux or macOS:

```bash
cp telegram_config.example.json telegram_config.json
```

## 7. Put your Telegram token into the config

Open `telegram_config.json` in any text editor.

Find this line:

```json
"bot_token": "PASTE_TELEGRAM_BOT_TOKEN_HERE"
```

Replace the placeholder with your real token:

```json
"bot_token": "1234567890:ABCDEF_your_token_here"
```

At first, keep these fields empty:

```json
"allowed_chat_ids": [],
"notification_chat_id": ""
```

This is safe. The bot will not show your task table to anyone until you allow a Chat ID.

Save the file.

## 8. Run the bot for the first time

### Windows

Double-click:

```text
run_telegram_task_bot.bat
```

or open PowerShell in the project folder and run:

```powershell
python telegram_task_bot.py
```

### Linux or macOS

Open Terminal in the project folder and run:

```bash
python3 telegram_task_bot.py
```

If everything is OK, you should see:

```text
Telegram task bot started. Press Ctrl+C to stop.
```

Do not close this window. The bot works while this script is running.

## 9. Find your Telegram Chat ID

1. Open your new bot in Telegram.
2. Send:

```text
/start
```

3. The bot will reply with your Chat ID.

It will look like a number:

```text
Your Chat ID: 123456789
```

Copy this number.

## 10. Allow your own Chat ID

Stop the bot window with:

```text
Ctrl + C
```

Open `telegram_config.json` again.

Add your Chat ID here:

```json
"allowed_chat_ids": [123456789],
"notification_chat_id": "123456789"
```

Example full config:

```json
{
  "bot_token": "1234567890:ABCDEF_your_token_here",
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

Save the file.

Run the bot again.

Now the bot should answer only to you.

## 11. Test basic commands

Send these commands to the bot in Telegram:

```text
/help
```

```text
/add Test task | tomorrow
```

```text
/list
```

```text
/today
```

```text
/summary
```

The task should appear in `Task_Planner.xlsx`.

## 12. How the Excel file works

The bot stores tasks in:

```text
Task_Planner.xlsx
```

Main columns:

| Column | Meaning |
| --- | --- |
| A | Task name |
| B | Due date |
| C | Status formula |
| D | Days remaining formula |
| E | Category formula |
| F | Completed flag |
| G | Completion date |

You can open the Excel file manually, but it is safer to edit tasks through the bot.

## 13. Useful commands

Add a task:

```text
/add Pay invoice | 25.05.2026
```

Add a task for tomorrow:

```text
/add Pay internet | tomorrow
```

Show all tasks:

```text
/list
```

Show today’s tasks:

```text
/today
```

Show overdue tasks:

```text
/overdue
```

Show upcoming tasks:

```text
/soon
```

Show nearest future tasks:

```text
/nearest
```

Show full summary:

```text
/summary
```

Delete a task:

```text
/delete 4
```

The bot will ask for confirmation before deletion.

## 14. Optional LLM setup

The bot can work without an LLM.

Without an LLM, structured commands work:

```text
/add Pay invoice | 25.05.2026
/list
/today
/delete 4
```

With an LLM, the bot can better understand normal human messages and questions about your table.

For example:

```text
добавь встречу с Иваном на завтра
что у меня самое срочное?
перенеси задачу про отчет на пятницу
```

To enable LLM support, open `telegram_config.json` and add your API key:

```json
"llm_api_key": "YOUR_API_KEY"
```

The default endpoint is:

```json
"llm_base_url": "https://api.openai.com/v1"
```

You can use another LLM provider if it supports OpenAI-compatible API endpoints for:

```text
/responses
/audio/transcriptions
```

For a non-compatible provider, you need to adapt these functions in `telegram_task_bot.py`:

```text
openai_json_response(...)
openai_text_response(...)
transcribe_audio(...)
```

## 15. Optional voice messages

For voice messages, you need:

1. LLM transcription API key.
2. `ffmpeg` installed on your system.

Telegram voice messages are usually sent as OGG/Opus files. `ffmpeg` converts them into a format suitable for transcription.

If `ffmpeg` is missing, text commands will still work.

## 16. How to keep the bot running

The bot works only while the script is running.

For simple testing, keep the terminal window open.

For permanent use, you can:

- add `run_telegram_task_bot.bat` to Windows Task Scheduler;
- run the script on a VPS;
- create a Linux `systemd` service.

## 17. Very common problems

### Problem: Bot does not answer

Check:

1. Is the script running?
2. Is the Telegram token correct?
3. Did you add your Chat ID to `allowed_chat_ids`?
4. Did you restart the bot after editing `telegram_config.json`?

### Problem: Bot says access is not configured

This is normal on first launch.

Send `/start`, copy your Chat ID, add it to `allowed_chat_ids`, then restart the bot.

### Problem: Python command does not work on Windows

Try:

```powershell
py telegram_task_bot.py
```

or reinstall Python and enable:

```text
Add Python to PATH
```

### Problem: Voice messages do not work

Install `ffmpeg` and check that `llm_api_key` is set.

### Problem: I accidentally published my token

Immediately revoke the token in `@BotFather` and create a new one.

Also revoke any exposed LLM/API key.

## 18. What files should not be published

Do not publish:

```text
telegram_config.json
telegram_task_bot_state.json
Task_Planner.xlsx with your real personal tasks
```

The repository already includes `.gitignore` to help prevent this.

## 19. Safe files to publish

Safe files for GitHub:

```text
telegram_task_bot.py
telegram_config.example.json
telegram_config.sample.json
telegram_task_bot_state.example.json
Task_Planner.xlsx if it is empty
README.md
BEGINNER_INSTALLATION_GUIDE.md
SECURITY.md
LICENSE
requirements.txt
run_telegram_task_bot.bat
run_telegram_task_bot.sh
.gitignore
```

## 20. Minimum working setup

For the simplest setup, you only need:

1. Python installed.
2. `telegram_task_bot.py`.
3. `Task_Planner.xlsx`.
4. `telegram_config.json` with your bot token and Chat ID.
5. A running terminal window.

That is enough to use the bot with basic commands.
