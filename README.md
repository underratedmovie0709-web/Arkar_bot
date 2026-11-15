# YZ's Bot — Railway Deployment

Telegram bot for FREE and 3K Mobile Legends custom lobbies.

## Important setup before deployment

1. Create a bot with @BotFather and copy its token.
2. Add the bot as an **administrator** to `@magicchessgogo_YZ`.
   - It needs permission to view members / manage chat membership so Telegram can deliver member updates.
   - Without channel admin access, membership checks and join/leave tracking may fail.
3. Create a Railway project and add a PostgreSQL service.
4. Push this project to GitHub and deploy the repository on Railway.
5. Add the environment variables below to the bot service.

## Railway Variables

| Variable | Value |
|---|---|
| `BOT_TOKEN` | Your BotFather token |
| `CREATOR_ID` | `5188275898` |
| `CHANNEL_USERNAME` | `@magicchessgogo_YZ` |
| `CHANNEL_URL` | `https://t.me/magicchessgogo_YZ` |
| `OWNER_GAME_ID` | `81230904` |
| `KPAY_NUMBER` | `09684329592` |
| `KPAY_NAME` | `Kyaw Kyaw Soe` |
| `WAVEPAY_NUMBER` | `09699148028` |
| `WAVEPAY_NAME` | `Aye Thida Khine` |
| `MATCH_FEE` | `3000` |
| `DATABASE_URL` | Railway PostgreSQL `DATABASE_URL` reference |
| `LOG_LEVEL` | `INFO` |

For `DATABASE_URL`, use Railway's variable reference to the PostgreSQL service's `DATABASE_URL`. If Railway exposes a private connection URL, use that.

## Commands

### User
- `/start` — start the bot
- `/help` — help

### Admin
- `/check` — dashboard
- `/backup` — Excel backup
- `/expense AMOUNT DESCRIPTION` — record an expense (requires `reports`)
- `/user USER_ID` — user detail
- `/ban USER_ID`
- `/unban USER_ID`
- `/kickmember USER_ID` — Channel member ban (requires `channel`)
- `/unkickmember USER_ID` — Channel member unban (requires `channel`)

### Creator only
- `/addadmin USER_ID users,lobby,payments,reports,channel`
- `/deladmin USER_ID`
- `/restart` — creates and sends a verified Excel backup, then re-executes the bot process

Permissions are:
- `users`
- `lobby`
- `payments`
- `reports`
- `channel`

Example:
`/addadmin 123456789 payments,lobby`

## 3K fee logic

The configured rate is 3,000 MMK per player per match:
- 1 match = 3,000 MMK
- 2 matches = 6,000 MMK
- 3 matches = 9,000 MMK

Users must send a receipt image. Creator or an admin with `payments` permission approves or rejects it. Only approved registrations enter the 3K queue.

## Data and backup

The database is PostgreSQL and is not cleared by `/restart`. Excel backup exports Users, Admins, Registrations, Payments, Expenses, Matches, MatchPlayers, ActivityLogs, and ChannelEvents. The financial dashboard calculates income, expenses, and net balance.

**Keep regular off-platform backups too.** Excel files sent in Telegram are useful exports but should not be your only disaster-recovery copy.

## Notes / operational limitations

- This is a first deployable implementation; test with a private test channel and test payment before using real money.
- The bot cannot independently verify a KPay/WavePay transfer. Admins manually verify the receipt and approve/reject.
- Telegram bot membership checks require the bot to be an administrator in the channel.
- Telegram does not expose a complete historical member list or past join/leave history to bots. Join/leave counts start from when the bot is installed with the required updates.
- Dashboard list views show recent entries (not an unlimited export). Use the Excel backup for the full stored tables.
- The `/restart` command restarts the Python process without deleting database records. It is not a database wipe.
- Railway must keep the worker running and PostgreSQL attached. Do not deploy multiple bot replicas using polling; run one worker instance.
- Use a real PostgreSQL database in Railway, not the local SQLite fallback.

## Local test

```bash
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scripts\activate
pip install -r requirements.txt
cp .env.example .env
# Set variables in your shell or use python-dotenv locally (not required by this project).
python main.py
```
