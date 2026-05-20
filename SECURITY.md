# Security Policy

Do not commit real Telegram bot tokens, LLM API keys, personal Excel task files or runtime state files.

Before publishing a fork or archive, check for secrets:

```bash
grep -RInE "(sk-[A-Za-z0-9_-]+|[0-9]{8,}:[A-Za-z0-9_-]{20,})" . --exclude-dir=.git
```

If a real credential was exposed, revoke it and create a new one.
