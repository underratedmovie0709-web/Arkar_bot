import os
import sys
import json
import random
import logging
import asyncio
import traceback
from datetime import datetime, timezone
from zoneinfo import ZoneInfo
from functools import wraps

from telegram import (
    Update, InlineKeyboardButton, InlineKeyboardMarkup, BotCommand,
    BotCommandScopeDefault, BotCommandScopeChat, ReplyKeyboardRemove
)
from telegram.constants import ChatMemberStatus
from telegram.ext import (
    Application, CommandHandler, MessageHandler, CallbackQueryHandler,
    ChatMemberHandler, ContextTypes, ApplicationHandlerStop, filters
)
from sqlalchemy import (
    create_engine, Column, Integer, BigInteger, String, Boolean, DateTime,
    ForeignKey, Text, func, select, desc
)
from sqlalchemy.orm import declarative_base, sessionmaker, relationship
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# -------------------- CONFIG --------------------
BOT_TOKEN = os.getenv("BOT_TOKEN", "").strip()
CREATOR_ID = int(os.getenv("CREATOR_ID", "5188275898"))
CHANNEL_USERNAME = os.getenv("CHANNEL_USERNAME", "@magicchessgogo_YZ")
CHANNEL_URL = os.getenv("CHANNEL_URL", "https://t.me/magicchessgogo_YZ")
OWNER_GAME_ID = os.getenv("OWNER_GAME_ID", "81230904")
K_PAY_NUMBER = os.getenv("KPAY_NUMBER", "09684329592")
K_PAY_NAME = os.getenv("KPAY_NAME", "Kyaw Kyaw Soe")
WAVE_NUMBER = os.getenv("WAVEPAY_NUMBER", "09699148028")
WAVE_NAME = os.getenv("WAVEPAY_NAME", "Aye Thida Khine")
MATCH_FEE = int(os.getenv("MATCH_FEE", "3000"))
TIMEZONE = ZoneInfo("Asia/Yangon")
DATABASE_URL = os.getenv("DATABASE_URL", "sqlite:///yzbot_local.db")
if DATABASE_URL.startswith("postgres://"):
    DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql+psycopg2://", 1)
elif DATABASE_URL.startswith("postgresql://") and "+" not in DATABASE_URL:
    DATABASE_URL = DATABASE_URL.replace("postgresql://", "postgresql+psycopg2://", 1)

logging.basicConfig(
    format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
    level=os.getenv("LOG_LEVEL", "INFO").upper()
)
log = logging.getLogger("yzbot")
Base = declarative_base()
engine = create_engine(DATABASE_URL, pool_pre_ping=True, future=True)
Session = sessionmaker(bind=engine, expire_on_commit=False)

# -------------------- DATABASE --------------------
class User(Base):
    __tablename__ = "users"
    id = Column(BigInteger, primary_key=True)  # Telegram ID
    first_name = Column(String(255), default="")
    last_name = Column(String(255), default="")
    username = Column(String(255), default="")
    is_banned = Column(Boolean, default=False)
    bot_status = Column(String(30), default="active")
    first_seen = Column(DateTime(timezone=True), default=lambda: datetime.now(timezone.utc))
    last_seen = Column(DateTime(timezone=True), default=lambda: datetime.now(timezone.utc))
    game_name = Column(String(255), default="")
    owner_game_id = Column(String(80), default="")

class Admin(Base):
    __tablename__ = "admins"
    id = Column(BigInteger, primary_key=True)
    added_by = Column(BigInteger, nullable=False)
    permissions = Column(Text, default='["users","lobby","payments","reports","channel"]')
    created_at = Column(DateTime(timezone=True), default=lambda: datetime.now(timezone.utc))

class Activity(Base):
    __tablename__ = "activity_logs"
    id = Column(Integer, primary_key=True, autoincrement=True)
    user_id = Column(BigInteger, index=True)
    event = Column(String(100), nullable=False)
    details = Column(Text, default="")
    created_at = Column(DateTime(timezone=True), default=lambda: datetime.now(timezone.utc))

class Registration(Base):
    __tablename__ = "registrations"
    id = Column(Integer, primary_key=True, autoincrement=True)
    user_id = Column(BigInteger, ForeignKey("users.id"), index=True)
    lobby = Column(String(10), nullable=False)  # FREE / 3K
    game_name = Column(String(255), nullable=False)
    owner_game_id = Column(String(80), nullable=False)
    requested_matches = Column(Integer, nullable=False)
    status = Column(String(30), default="queued", index=True)
    created_at = Column(DateTime(timezone=True), default=lambda: datetime.now(timezone.utc))
    user = relationship("User")

class Payment(Base):
    __tablename__ = "payments"
    id = Column(Integer, primary_key=True, autoincrement=True)
    registration_id = Column(Integer, ForeignKey("registrations.id"), index=True)
    user_id = Column(BigInteger, ForeignKey("users.id"), index=True)
    amount = Column(Integer, nullable=False)
    method = Column(String(30), default="")
    receipt_file_id = Column(Text, default="")
    status = Column(String(30), default="awaiting_receipt", index=True)
    reviewed_by = Column(BigInteger, nullable=True)
    review_note = Column(Text, default="")
    created_at = Column(DateTime(timezone=True), default=lambda: datetime.now(timezone.utc))
    reviewed_at = Column(DateTime(timezone=True), nullable=True)
    registration = relationship("Registration")
    user = relationship("User")

class Expense(Base):
    __tablename__ = "expenses"
    id = Column(Integer, primary_key=True, autoincrement=True)
    amount = Column(Integer, nullable=False)
    category = Column(String(120), default="General")
    description = Column(Text, default="")
    created_by = Column(BigInteger, nullable=False)
    created_at = Column(DateTime(timezone=True), default=lambda: datetime.now(timezone.utc))

class Match(Base):
    __tablename__ = "matches"
    id = Column(Integer, primary_key=True, autoincrement=True)
    lobby = Column(String(10), nullable=False)
    leader_user_id = Column(BigInteger, nullable=False)
    status = Column(String(30), default="formed")
    created_at = Column(DateTime(timezone=True), default=lambda: datetime.now(timezone.utc))

class MatchPlayer(Base):
    __tablename__ = "match_players"
    id = Column(Integer, primary_key=True, autoincrement=True)
    match_id = Column(Integer, ForeignKey("matches.id"), index=True)
    user_id = Column(BigInteger, ForeignKey("users.id"), index=True)
    registration_id = Column(Integer, ForeignKey("registrations.id"))
    game_name = Column(String(255), default="")
    owner_game_id = Column(String(80), default="")

class ChannelEvent(Base):
    __tablename__ = "channel_events"
    id = Column(Integer, primary_key=True, autoincrement=True)
    user_id = Column(BigInteger, index=True)
    event = Column(String(20), nullable=False)
    created_at = Column(DateTime(timezone=True), default=lambda: datetime.now(timezone.utc))

Base.metadata.create_all(engine)

def now_utc():
    return datetime.now(timezone.utc)

def local_time(dt):
    if not dt:
        return "—"
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    return dt.astimezone(TIMEZONE).strftime("%Y-%m-%d %I:%M:%S %p")

def log_event(db, user_id, event, details=""):
    db.add(Activity(user_id=user_id, event=event, details=details[:4000], created_at=now_utc()))

def get_user(db, tg_user):
    u = db.get(User, tg_user.id)
    if not u:
        u = User(
            id=tg_user.id, first_name=tg_user.first_name or "",
            last_name=tg_user.last_name or "", username=tg_user.username or "",
            first_seen=now_utc(), last_seen=now_utc()
        )
        db.add(u)
    else:
        u.first_name = tg_user.first_name or ""
        u.last_name = tg_user.last_name or ""
        u.username = tg_user.username or ""
        u.last_seen = now_utc()
    return u

def admin_permissions(db, user_id):
    if user_id == CREATOR_ID:
        return {"*"}
    a = db.get(Admin, user_id)
    if not a:
        return set()
    try:
        return set(json.loads(a.permissions or "[]"))
    except Exception:
        return set()

def is_admin(db, user_id, permission=None):
    perms = admin_permissions(db, user_id)
    return "*" in perms or (permission is None and bool(perms)) or (permission in perms)

async def notify_staff(context, text, document_path=None):
    targets = {CREATOR_ID}
    with Session() as db:
        for a in db.scalars(select(Admin)).all():
            targets.add(a.id)
    for target in targets:
        try:
            if document_path:
                with open(document_path, "rb") as f:
                    await context.bot.send_document(target, f, caption=text[:1000])
            else:
                await context.bot.send_message(target, text[:4000])
        except Exception as exc:
            log.warning("Staff notification failed for %s: %s", target, exc)

async def error_handler(update: object, context: ContextTypes.DEFAULT_TYPE):
    log.error("Unhandled exception: %s", "".join(traceback.format_exception(
        None, context.error, context.error.__traceback__
    )))
    msg = "⚠️ YZ Bot Error\n" + str(context.error)[:2500]
    try:
        await notify_staff(context, msg)
    except Exception:
        log.exception("Could not notify staff")

# -------------------- CHANNEL CHECK --------------------
async def joined_channel(bot, user_id):
    try:
        member = await bot.get_chat_member(CHANNEL_USERNAME, user_id)
        return (member.status in (
            ChatMemberStatus.MEMBER, ChatMemberStatus.ADMINISTRATOR,
            ChatMemberStatus.OWNER
        ) or (member.status == ChatMemberStatus.RESTRICTED and getattr(member, "is_member", False)))
    except Exception as exc:
        log.warning("Channel membership check failed: %s", exc)
        return False

def join_keyboard():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("📢 Channel ကို Join လုပ်ရန်", url=CHANNEL_URL)],
        [InlineKeyboardButton("🔄 Join ပြီးပါပြီ - ပြန်စစ်မယ်", callback_data="verify_join")]
    ])

async def require_join(update, context):
    user = update.effective_user
    if not user:
        return False
    with Session() as db:
        u = db.get(User, user.id)
        if u and u.is_banned:
            msg = "⛔ သင့်အကောင့်ကို Bot အသုံးပြုခွင့် ပိတ်ထားပါတယ်။"
            if update.callback_query:
                await update.callback_query.answer(msg, show_alert=True)
            elif update.effective_message:
                await update.effective_message.reply_text(msg)
            return False
    if await joined_channel(context.bot, user.id):
        return True
    if update.callback_query:
        await update.callback_query.answer("Channel ကို Join လုပ်ပြီးမှ အသုံးပြုနိုင်ပါတယ်။", show_alert=True)
        await update.callback_query.message.reply_text(
            "🔒 Bot ကိုအသုံးပြုရန် အောက်ပါ Channel ကို အရင် Join လုပ်ပေးပါ။",
            reply_markup=join_keyboard()
        )
    elif update.effective_message:
        await update.effective_message.reply_text(
            "🔒 Bot ကိုအသုံးပြုရန် အောက်ပါ Channel ကို အရင် Join လုပ်ပေးပါ။",
            reply_markup=join_keyboard()
        )
    return False

# -------------------- KEYBOARDS --------------------
def main_keyboard():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("🆓 Free Lobby", callback_data="lobby:FREE")],
        [InlineKeyboardButton("💰 3K Lobby", callback_data="lobby:3K")],
        [InlineKeyboardButton("👑 Owner သို့ ဆက်သွယ်ရန်", url=f"tg://user?id={CREATOR_ID}")]
    ])

def lobby_keyboard(lobby):
    rows = [
        [InlineKeyboardButton("🎮 MCGG In-game Name", callback_data=f"field:{lobby}:game")],
        [InlineKeyboardButton("🆔 Owner In-game ID", callback_data=f"field:{lobby}:owner")],
        [InlineKeyboardButton("🏆 ဘယ်နှစ်ပွဲ ဆော့မလဲ", callback_data=f"field:{lobby}:matches")],
    ]
    # The 3K menu has exactly four requested buttons; FREE has exactly three.
    if lobby == "3K":
        rows.append([InlineKeyboardButton("💳 ငွေပေးချေရန်", callback_data="field:3K:payment")])
    return InlineKeyboardMarkup(rows)

def payment_keyboard():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("💳 KBZPay (KPay)", callback_data="paymethod:KPay")],
        [InlineKeyboardButton("💜 WavePay", callback_data="paymethod:WavePay")]
    ])


# -------------------- MATCHMAKER --------------------
async def process_queue(context, lobby):
    """Builds matches from 8 distinct users. A player can request multiple matches."""
    completed = []
    with Session() as db:
        regs = db.scalars(
            select(Registration).where(
                Registration.lobby == lobby,
                Registration.status.in_(["queued", "partially_assigned"])
            ).order_by(Registration.created_at, Registration.id)
        ).all()
        remaining = {}
        for reg in regs:
            assigned = db.scalar(select(func.count(MatchPlayer.id)).where(
                MatchPlayer.registration_id == reg.id
            )) or 0
            remaining[reg.id] = max(0, reg.requested_matches - assigned)

        while True:
            candidates = []
            used_users = set()
            for reg in regs:
                if remaining.get(reg.id, 0) > 0 and reg.user_id not in used_users:
                    candidates.append(reg)
                    used_users.add(reg.user_id)
                if len(candidates) == 8:
                    break
            if len(candidates) < 8:
                break
            leader = random.choice(candidates)
            match = Match(lobby=lobby, leader_user_id=leader.user_id, created_at=now_utc())
            db.add(match)
            db.flush()
            for reg in candidates:
                db.add(MatchPlayer(
                    match_id=match.id, user_id=reg.user_id, registration_id=reg.id,
                    game_name=reg.game_name, owner_game_id=reg.owner_game_id
                ))
                remaining[reg.id] -= 1
            completed.append((match.id, lobby, leader.user_id, [
                (r.user_id, r.game_name, r.owner_game_id,
                 r.user.first_name if r.user else "",
                 r.user.last_name if r.user else "",
                 r.user.username if r.user else "") for r in candidates
            ]))
            for reg in candidates:
                reg.status = "completed" if remaining[reg.id] == 0 else "partially_assigned"
        for reg in regs:
            if remaining.get(reg.id, 0) > 0 and reg.status == "queued":
                reg.status = "partially_assigned"
        db.commit()

    for match_id, mode, leader_id, players in completed:
        leader_info = next((p for p in players if p[0] == leader_id), None)
        lines = [f"🏆 {mode} Lobby - Match #{match_id}", "━━━━━━━━━━━━━━"]
        for idx, (uid, game, owner_id, first_name, last_name, username) in enumerate(players, 1):
            leader_tag = " 👑 LEADER" if uid == leader_id else ""
            lines.append(f"{idx}. {game}{leader_tag}\n   Telegram Name: {(first_name + ' ' + last_name).strip()}\n   Username: @{username or '—'}\n   Telegram ID: {uid}\n   Owner ID: {owner_id}")
        report = "\n".join(lines)
        try:
            await context.bot.send_message(CREATOR_ID, "📣 Match ပြည့်သွားပါပြီ!\n\n" + report)
        except Exception:
            log.exception("Creator match notification failed")
        for uid, game, owner_id, first_name, last_name, username in players:
            try:
                await context.bot.send_message(
                    uid, f"🎉 Match #{match_id} ပြည့်သွားပါပြီ!\n\n"
                    f"👑 Leader: {leader_info[1]}\n"
                    f"{'သင်ဟာ ဒီပွဲရဲ့ Leader ဖြစ်ပါတယ်။' if uid == leader_id else 'Leader ရဲ့ညွှန်ကြားချက်ကို လိုက်နာပေးပါ။'}"
                )
            except Exception:
                log.info("Could not send match result to %s", uid)

async def queue_status(context, user_id, lobby):
    with Session() as db:
        pending = db.scalars(select(Registration).where(
            Registration.user_id == user_id,
            Registration.lobby == lobby,
            Registration.status.in_(["queued", "partially_assigned"])
        )).all()
        if not pending:
            return
        waiting = sum(max(0, r.requested_matches - (db.scalar(
            select(func.count(MatchPlayer.id)).where(MatchPlayer.registration_id == r.id)
        ) or 0)) for r in pending)
        count = len({r.user_id for r in db.scalars(select(Registration).where(
            Registration.lobby == lobby,
            Registration.status.in_(["queued", "partially_assigned"])
        )).all()})
    await context.bot.send_message(
        user_id,
        f"⏳ {lobby} Lobby စာရင်းသွင်းမှု အောင်မြင်ပါတယ်။\n"
        f"📌 သင့်အတွက် ကျန်ရှိသောပွဲ: {waiting}\n"
        f"👥 Queue ထဲရှိ ကွဲပြားသော Player: {count}\n"
        "8 ယောက်ပြည့်တာနဲ့ Match ကို အလိုအလျောက်ဖွဲ့ပေးပါမယ်။"
    )

# -------------------- USER FLOW --------------------
WELCOME = (
    "✨━━━━━━━━━━━━━━━━━━━━✨\\n"
    "👑 𝐖𝐄𝐋𝐂𝐎𝐌𝐄 𝐓𝐎 𝐘𝐙'𝐒 𝐁𝐎𝐓 👑\\n"
    "✨━━━━━━━━━━━━━━━━━━━━✨\\n\\n"
    "မင်္ဂလာပါ သူငယ်ချင်းတို့ရေ 💖\\n"
    "YZ's Bot မှာ နွေးထွေးစွာ ကြိုဆိုပါတယ်ဗျာ။\\n\\n"
    "🎮 *ကိုယ်တိုင်ရွေးချယ်ပြီး Lobby ဝင်ရောက်နိုင်ပါပြီ။*\\n"
    "🆓 Free Lobby — အခမဲ့ ပါဝင်ကစားနိုင်ပါတယ်။\\n"
    f"💎 3K Lobby — တစ်ပွဲလျှင် {MATCH_FEE:,} MMK ဖြစ်ပါတယ်။\\n\\n"
    "📌 ပါဝင်ကစားလိုတဲ့ Lobby ကို အောက်က Button မှာ ရွေးချယ်ပြီး "
    "အချက်အလက်တွေ ဖြည့်သွင်းပေးပါခင်ဗျာ။\\n\\n"
    "🏆 ပျော်ရွှင်စရာပွဲကောင်းများ ဖြစ်ပါစေ။\\n"
    "✨━━━━━━━━━━━━━━━━━━━━✨"
)

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    current_time = now_utc()
    # Notify staff only for a first-time user or a genuinely returning session.
    # This prevents repeated /start updates from flooding the creator's chat.
    notify_start = False
    with Session() as db:
        existing = db.get(User, user.id)
        if existing is None:
            notify_start = True
        else:
            previous_seen = existing.last_seen or existing.first_seen
            if previous_seen:
                if previous_seen.tzinfo is None:
                    previous_seen = previous_seen.replace(tzinfo=timezone.utc)
                notify_start = (current_time - previous_seen).total_seconds() >= 6 * 60 * 60
            else:
                notify_start = True
        get_user(db, user)
        log_event(db, user.id, "BOT_START", "User started the bot")
        db.commit()
    if notify_start:
        await notify_staff(context, f"🟢 Bot အသုံးပြုသူအသစ်/ပြန်ဝင်သူ\n"
                             f"Name: {user.full_name}\nUsername: @{user.username or '—'}\n"
                             f"ID: {user.id}\nTime: {local_time(current_time)}")
    if not await require_join(update, context):
        return
    await update.effective_message.reply_text(WELCOME, reply_markup=main_keyboard(), parse_mode="Markdown")

async def home_callback(update, context):
    q = update.callback_query
    await q.answer()
    if not await require_join(update, context):
        return
    await q.message.reply_text(WELCOME, reply_markup=main_keyboard(), parse_mode="Markdown")

async def lobby_callback(update, context):
    q = update.callback_query
    await q.answer()
    if not await require_join(update, context):
        return
    lobby = q.data.split(":")[1]
    if lobby not in ("FREE", "3K"):
        await q.answer("Lobby အမျိုးအစား မမှန်ပါ။", show_alert=True)
        return
    context.user_data["draft"] = {"lobby": lobby}
    # A new lobby registration must not accidentally reuse an older payment session.
    context.user_data.pop("payment_id", None)
    title = "🆓 FREE LOBBY" if lobby == "FREE" else "💰 3K LOBBY"
    await q.message.reply_text(f"{title}\n\nစာရင်းသွင်းရန် အောက်ပါအချက်များကို ဖြည့်ပေးပါ။",
                               reply_markup=lobby_keyboard(lobby))

async def field_callback(update, context):
    q = update.callback_query
    await q.answer()
    if not await require_join(update, context):
        return
    _, lobby, field = q.data.split(":")
    if lobby not in ("FREE", "3K"):
        await q.answer("Lobby အမျိုးအစား မမှန်ပါ။", show_alert=True)
        return
    draft = context.user_data.setdefault("draft", {"lobby": lobby})
    draft["lobby"] = lobby
    if field == "payment":
        if lobby != "3K":
            await q.answer("ဒီလုပ်ဆောင်ချက်ကို 3K Lobby မှာသာ အသုံးပြုနိုင်ပါတယ်။", show_alert=True)
            return
        await q.message.reply_text(
            "💳 *ငွေပေးချေမှု နည်းလမ်းရွေးချယ်ရန်*\\n\\n"
            f"💎 တစ်ပွဲလျှင် {MATCH_FEE:,} MMK ကျသင့်ပါမယ်။\\n"
            "🏦 အောက်ပါ Payment နည်းလမ်း ၂ မျိုးထဲမှ တစ်ခုကို ရွေးချယ်ပါ။",
            reply_markup=payment_keyboard(), parse_mode="Markdown"
        )
        return
    if field not in ("game", "owner", "matches"):
        await q.answer("လုပ်ဆောင်ချက် မမှန်ပါ။", show_alert=True)
        return
    context.user_data["awaiting"] = field
    prompts = {
        "game": "🎮 *MCGG In-game Name ထည့်သွင်းရန်*\\nသင့်ရဲ့ Game Name ကို စာသားအဖြစ် ပို့ပေးပါ။",
        "owner": f"🆔 *Owner In-game ID ထည့်သွင်းရန်*\\nကျေးဇူးပြု၍ Owner Game ID ကို အတိအကျ ပို့ပေးပါ။\\n\\n🔐 Owner ID: `{OWNER_GAME_ID}`",
        "matches": "🏆 *ကစားမည့်ပွဲအရေအတွက်*\\nဥပမာ `2` သို့မဟုတ် `3` လို့ ပို့ပေးပါ။\\n💡 3K Lobby အတွက် တစ်ပွဲလျှင် 3,000 MMK ကျသင့်ပါတယ်။"
    }
    await q.message.reply_text(prompts[field], parse_mode="Markdown")

async def finalize_draft(update, context):
    user = update.effective_user
    draft = context.user_data.get("draft", {})
    lobby = draft.get("lobby")
    if not lobby:
        return False
    field = context.user_data.get("awaiting")
    value = (update.effective_message.text or "").strip()
    if field == "game":
        if len(value) < 2 or len(value) > 50:
            await update.effective_message.reply_text("⚠️ Game Name ကို 2 မှ 50 characters အတွင်း ပို့ပေးပါ။")
            return True
        draft["game_name"] = value
    elif field == "owner":
        if value != OWNER_GAME_ID:
            await update.effective_message.reply_text(f"⚠️ Owner ID မမှန်ပါ။ {OWNER_GAME_ID} ကိုသာ ထည့်ပေးပါ။")
            return True
        draft["owner_game_id"] = value
    elif field == "matches":
        try:
            n = int(value)
            if n < 1 or n > 20:
                raise ValueError
        except ValueError:
            await update.effective_message.reply_text("⚠️ ပွဲအရေအတွက်ကို 1 မှ 20 အတွင်း ကိန်းပြည့်ဖြင့် ပို့ပေးပါ။")
            return True
        draft["matches"] = n
    else:
        return False
    context.user_data.pop("awaiting", None)
    await update.effective_message.reply_text("✅ အချက်အလက်ကို လက်ခံသိမ်းဆည်းပြီးပါပြီ။ ကျန်အချက်အလက်များကို ဆက်လက်ဖြည့်ပေးပါ။")
    if all(k in draft for k in ("game_name", "owner_game_id", "matches")):
        with Session() as db:
            u = get_user(db, user)
            u.game_name = draft["game_name"]
            u.owner_game_id = draft["owner_game_id"]
            log_event(db, user.id, "LOBBY_DETAILS_SUBMITTED", json.dumps(draft, ensure_ascii=False))
            active = db.scalar(select(Registration).where(
                Registration.user_id == user.id,
                Registration.lobby == lobby,
                Registration.status.in_(["awaiting_payment","pending_review","queued","partially_assigned"])
            ))
            if active:
                await update.effective_message.reply_text("⚠️ ဒီ Lobby မှာ သင့်ရဲ့ Active စာရင်းတစ်ခု ရှိနေပြီးသားပါ။")
                db.commit()
                context.user_data.pop("draft", None)
                return True
            reg = Registration(
                user_id=user.id, lobby=lobby, game_name=draft["game_name"],
                owner_game_id=draft["owner_game_id"], requested_matches=draft["matches"],
                status="queued" if lobby == "FREE" else "awaiting_payment", created_at=now_utc()
            )
            db.add(reg)
            db.flush()
            if lobby == "FREE":
                log_event(db, user.id, "FREE_REGISTERED", f"registration={reg.id}, matches={draft['matches']}")
                db.commit()
                context.user_data.pop("draft", None)
                await update.effective_message.reply_text("🎉 🆓 FREE Lobby စာရင်းသွင်းမှု အောင်မြင်ပါပြီ။ 8 ယောက်ပြည့်ပါက Match ကို အလိုအလျောက်ဖွဲ့ပေးပါမယ်။")
                await process_queue(context, "FREE")
                await queue_status(context, user.id, "FREE")
            else:
                amount = MATCH_FEE * draft["matches"]
                selected_method = draft.get("payment_method", "")
                payment = Payment(
                    registration_id=reg.id, user_id=user.id, amount=amount,
                    method=selected_method,
                    status="awaiting_receipt", created_at=now_utc()
                )
                db.add(payment)
                log_event(db, user.id, "3K_PAYMENT_CREATED", f"registration={reg.id}, amount={amount}")
                db.commit()
                context.user_data["payment_id"] = payment.id
                context.user_data.pop("draft", None)
                summary = (
                    "💎 *3K LOBBY စာရင်းသွင်းမှု*\\n"
                    "━━━━━━━━━━━━━━━━━━\\n"
                    f"🎮 ကစားမည့်ပွဲ: {draft['matches']} ပွဲ\\n"
                    f"💵 တစ်ပွဲနှုန်း: {MATCH_FEE:,} MMK\\n"
                    f"🧾 စုစုပေါင်းကျသင့်ငွေ: *{amount:,} MMK*\\n"
                    "━━━━━━━━━━━━━━━━━━\\n\\n"
                )
                if selected_method:
                    if selected_method == "KPay":
                        pay_details = f"💳 *KPay ငွေလွှဲရန်*\\n📱 နံပါတ်: `{K_PAY_NUMBER}`\\n👤 အမည်: *{K_PAY_NAME}*"
                    else:
                        pay_details = f"💜 *WavePay ငွေလွှဲရန်*\\n📱 နံပါတ်: `{WAVE_NUMBER}`\\n👤 အမည်: *{WAVE_NAME}*"
                    await update.effective_message.reply_text(
                        summary + pay_details + "\\n\\n🧾 ငွေလွှဲပြီးပါက ပြေစာ Screenshot ကို ဒီ Chat ထဲ ပို့ပေးပါ။",
                        parse_mode="Markdown"
                    )
                else:
                    await update.effective_message.reply_text(
                        summary + "💳 ငွေပေးချေရန် နည်းလမ်းကို ရွေးချယ်ပေးပါ။",
                        reply_markup=payment_keyboard(), parse_mode="Markdown"
                    )
    return True

async def payment_method_callback(update, context):
    q = update.callback_query
    await q.answer()
    if not await require_join(update, context):
        return
    method = q.data.split(":")[1]
    if method not in ("KPay", "WavePay"):
        await q.answer("Payment နည်းလမ်း မမှန်ပါ။", show_alert=True)
        return

    payment_id = context.user_data.get("payment_id")
    if not payment_id:
        draft = context.user_data.get("draft")
        if not draft or draft.get("lobby") != "3K":
            await q.message.reply_text("⚠️ 3K Lobby စာရင်းသွင်းမှုကို အရင်စတင်ပေးပါ။")
            return
        draft["payment_method"] = method
        await q.message.reply_text(
            f"✅ {method} ကို ရွေးချယ်ပြီးပါပြီ။\\n"
            "📝 ကျန်ရှိနေသေးတဲ့ Game Name၊ Owner ID နဲ့ ပွဲအရေအတွက်ကို ဖြည့်သွင်းပြီးပါက "
            "ငွေလွှဲရန် အချက်အလက်ကို ပြပေးပါမယ်။"
        )
        return

    with Session() as db:
        p = db.get(Payment, payment_id)
        if not p or p.user_id != q.from_user.id or p.status != "awaiting_receipt":
            await q.message.reply_text("⚠️ Payment အခြေအနေ မမှန်ပါ။")
            return
        p.method = method
        amount = p.amount
        db.commit()
    if method == "KPay":
        details = f"💳 *KPay ငွေလွှဲရန်*\\n📱 နံပါတ်: `{K_PAY_NUMBER}`\\n👤 အမည်: *{K_PAY_NAME}*"
    else:
        details = f"💜 *WavePay ငွေလွှဲရန်*\\n📱 နံပါတ်: `{WAVE_NUMBER}`\\n👤 အမည်: *{WAVE_NAME}*"
    await q.message.reply_text(
        f"{details}\\n\\n🧾 ပေးချေရမည့်ငွေ: *{amount:,} MMK*\\n"
        "📸 ငွေလွှဲပြီးပါက ပြေစာ Screenshot ကို ဒီ Chat ထဲ ပို့ပေးပါ။",
        parse_mode="Markdown"
    )

async def receipt_handler(update, context):
    user = update.effective_user
    payment_id = context.user_data.get("payment_id")
    if not payment_id:
        return
    msg = update.effective_message
    file_id = None
    if msg.photo:
        file_id = msg.photo[-1].file_id
    elif msg.document and (msg.document.mime_type or "").startswith("image/"):
        file_id = msg.document.file_id
    if not file_id:
        await msg.reply_text("🧾 ပြေစာကို Screenshot သို့မဟုတ် Image File အဖြစ် ပို့ပေးပါ။")
        return
    with Session() as db:
        p = db.get(Payment, payment_id)
        if not p or p.user_id != user.id or p.status != "awaiting_receipt":
            await msg.reply_text("⚠️ ဒီ Payment ကို လက်ခံလို့မရတော့ပါ။")
            return
        p.receipt_file_id = file_id
        p.status = "pending_review"
        log_event(db, user.id, "PAYMENT_RECEIPT_SENT", f"payment={p.id}, method={p.method}, amount={p.amount}")
        db.commit()
        details = (f"🧾 Payment စစ်ဆေးရန် #{p.id}\n"
                   f"Name: {user.full_name}\nUsername: @{user.username or '—'}\n"
                   f"Telegram ID: {user.id}\nAmount: {p.amount:,} MMK\nMethod: {p.method}\n"
                   f"Time: {local_time(now_utc())}")
    await msg.reply_text("✅ ပြေစာကို လက်ခံရရှိပါပြီ။ Admin စစ်ဆေးပြီးမှ စာရင်းဝင်ပါမယ်။")
    targets = {CREATOR_ID}
    with Session() as db:
        targets.update(a.id for a in db.scalars(select(Admin)).all()
                       if is_admin(db, a.id, "payments"))
    keyboard = InlineKeyboardMarkup([[
        InlineKeyboardButton("✅ လက်ခံမယ်", callback_data=f"payment:approve:{payment_id}"),
        InlineKeyboardButton("❌ ပယ်ချမယ်", callback_data=f"payment:reject:{payment_id}")
    ]])
    for target in targets:
        try:
            await context.bot.send_photo(target, file_id, caption=details, reply_markup=keyboard)
        except Exception:
            await context.bot.send_message(target, details, reply_markup=keyboard)

async def payment_review_callback(update, context):
    q = update.callback_query
    parts = q.data.split(":")
    action, pid = parts[1], int(parts[2])
    with Session() as db:
        if not is_admin(db, q.from_user.id, "payments"):
            await q.answer("ခွင့်ပြုချက်မရှိပါ။", show_alert=True)
            return
        p = db.get(Payment, pid)
        if not p or p.status != "pending_review":
            await q.answer("ဒီ Payment ကို စစ်ဆေးပြီးသား သို့မဟုတ် မတွေ့ပါ။", show_alert=True)
            return
        p.reviewed_by = q.from_user.id
        p.reviewed_at = now_utc()
        reg = db.get(Registration, p.registration_id)
        if action == "approve":
            p.status = "approved"
            reg.status = "queued"
            log_event(db, p.user_id, "PAYMENT_APPROVED", f"payment={pid}, amount={p.amount}")
            user_id, lobby = p.user_id, reg.lobby
            response = "✅ Payment လက်ခံပြီးပါပြီ။ သင့်စာရင်းကို Lobby Queue ထဲ ထည့်သွင်းလိုက်ပါပြီ။"
        else:
            p.status = "rejected"
            reg.status = "rejected"
            p.review_note = "Rejected by admin"
            log_event(db, p.user_id, "PAYMENT_REJECTED", f"payment={pid}")
            user_id, lobby = p.user_id, reg.lobby
            response = "❌ ပြေစာကို Admin က လက်မခံပါ။ Owner ကို ဆက်သွယ်ပေးပါ။"
        db.commit()
    await q.answer("စစ်ဆေးမှု သိမ်းဆည်းပြီးပါပြီ။")
    await q.edit_message_reply_markup(reply_markup=None)
    await context.bot.send_message(user_id, response)
    if action == "approve":
        await process_queue(context, lobby)
        await queue_status(context, user_id, lobby)

async def text_router(update, context):
    if not await require_join(update, context):
        return
    if await finalize_draft(update, context):
        return
    if context.user_data.get("payment_id"):
        await receipt_handler(update, context)
        return
    if update.effective_message:
        await update.effective_message.reply_text("မီနူးမှ လုပ်ဆောင်ချက်တစ်ခုကို ရွေးချယ်ပေးပါ။", reply_markup=main_keyboard())

# -------------------- DASHBOARD / ADMIN --------------------
def dashboard_keyboard():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("👥 Total Users", callback_data="dash:users"),
         InlineKeyboardButton("🆓 FREE Players", callback_data="dash:free")],
        [InlineKeyboardButton("💰 3K Players", callback_data="dash:3k"),
         InlineKeyboardButton("🏆 Match Lists", callback_data="dash:matches")],
        [InlineKeyboardButton("💵 Income Reports", callback_data="dash:income"),
         InlineKeyboardButton("📢 Channel Stats", callback_data="dash:channel")],
        [InlineKeyboardButton("🛡️ Admin Management", callback_data="dash:admins"),
         InlineKeyboardButton("🔐 User Management", callback_data="dash:manage")],
        [InlineKeyboardButton("📦 Excel Backup", callback_data="dash:backup")]
    ])

async def check_command(update, context):
    with Session() as db:
        if not is_admin(db, update.effective_user.id):
            await update.effective_message.reply_text("⛔ Owner/Admin သာ အသုံးပြုနိုင်ပါတယ်။")
            return
    await update.effective_message.reply_text("👑 YZ BOT CONTROL PANEL\n\nလိုအပ်သော ကဏ္ဍကို ရွေးချယ်ပါ။",
                                              reply_markup=dashboard_keyboard())

async def dash_callback(update, context):
    q = update.callback_query
    section = q.data.split(":")[1]
    with Session() as db:
        if not is_admin(db, q.from_user.id):
            await q.answer("ခွင့်ပြုချက်မရှိပါ။", show_alert=True)
            return
        if section == "users" and not is_admin(db, q.from_user.id, "users"):
            await q.answer("ခွင့်ပြုချက်မရှိပါ။", show_alert=True); return
        if section in ("free","3k","matches") and not is_admin(db, q.from_user.id, "lobby"):
            await q.answer("ခွင့်ပြုချက်မရှိပါ။", show_alert=True); return
        if section == "income" and not is_admin(db, q.from_user.id, "reports"):
            await q.answer("ခွင့်ပြုချက်မရှိပါ။", show_alert=True); return
        if section == "channel" and not is_admin(db, q.from_user.id, "channel"):
            await q.answer("ခွင့်ပြုချက်မရှိပါ။", show_alert=True); return
        if section == "admins" and q.from_user.id != CREATOR_ID:
            await q.answer("Creator သာ အသုံးပြုနိုင်ပါတယ်။", show_alert=True); return
        if section == "manage" and not is_admin(db, q.from_user.id, "users"):
            await q.answer("ခွင့်ပြုချက်မရှိပါ။", show_alert=True); return
        if section == "users":
            users = db.scalars(select(User).order_by(desc(User.last_seen)).limit(30)).all()
            lines = [f"👥 Total Users: {db.scalar(select(func.count(User.id))) or 0}", ""]
            for u in users:
                lines.append(f"• {u.first_name} {u.last_name} | @{u.username or '—'} | ID: {u.id}\n  Last: {local_time(u.last_seen)} | {'BANNED' if u.is_banned else u.bot_status}")
            text = "\n".join(lines)
        elif section in ("free","3k"):
            lobby = "FREE" if section == "free" else "3K"
            regs = db.scalars(select(Registration).where(Registration.lobby == lobby).order_by(desc(Registration.created_at)).limit(30)).all()
            text = f"📋 {lobby} Players (latest 30)\n\n" + ("\n\n".join(
                f"Name: {r.user.first_name if r.user else '—'} {r.user.last_name if r.user else ''}\n"
                f"Username: @{r.user.username if r.user and r.user.username else '—'}\n"
                f"Telegram ID: {r.user_id}\nGame: {r.game_name}\nOwner ID: {r.owner_game_id}\n"
                f"Matches: {r.requested_matches}\nStatus: {r.status}\nTime: {local_time(r.created_at)}"
                for r in regs
            ) if regs else "စာရင်းမရှိသေးပါ။")
        elif section == "matches":
            matches = db.scalars(select(Match).order_by(desc(Match.created_at)).limit(20)).all()
            free_total = db.scalar(select(func.count(Match.id)).where(Match.lobby == "FREE")) or 0
            paid_total = db.scalar(select(func.count(Match.id)).where(Match.lobby == "3K")) or 0
            text = f"🏆 MATCH STATISTICS\nFREE Matches: {free_total}\n3K Matches: {paid_total}\n\nRecent Matches\n\n"
            for m in matches:
                players = db.scalars(select(MatchPlayer).where(MatchPlayer.match_id == m.id)).all()
                text += f"#{m.id} {m.lobby} | Leader: {m.leader_user_id} | {local_time(m.created_at)}\n"
                for p in players:
                    text += f"  - {p.game_name} (ID {p.user_id}){' 👑' if p.user_id == m.leader_user_id else ''}\n"
                text += "\n"
            if not matches: text += "Match မရှိသေးပါ။"
        elif section == "income":
            today = now_utc().astimezone(TIMEZONE).date()
            month = today.strftime("%Y-%m")
            approved = db.scalars(select(Payment).where(Payment.status == "approved")).all()
            daily = sum(p.amount for p in approved if p.reviewed_at and local_time(p.reviewed_at).startswith(today.strftime("%Y-%m-%d")))
            monthly = sum(p.amount for p in approved if p.reviewed_at and local_time(p.reviewed_at).startswith(month))
            total = sum(p.amount for p in approved)
            expenses = db.scalars(select(Expense)).all()
            expense_today = sum(e.amount for e in expenses if local_time(e.created_at).startswith(today.strftime("%Y-%m-%d")))
            expense_month = sum(e.amount for e in expenses if local_time(e.created_at).startswith(month))
            expense_total = sum(e.amount for e in expenses)
            text = (f"💵 FINANCIAL REPORT\n\nIncome today: {daily:,} MMK\nExpense today: {expense_today:,} MMK\nNet today: {daily-expense_today:,} MMK\n\n"
                    f"Income this month: {monthly:,} MMK\nExpense this month: {expense_month:,} MMK\nNet this month: {monthly-expense_month:,} MMK\n\n"
                    f"All-time income: {total:,} MMK\nAll-time expenses: {expense_total:,} MMK\nNet balance: {total-expense_total:,} MMK\nApproved payments: {len(approved)}")
        elif section == "channel":
            try:
                count = await context.bot.get_chat_member_count(CHANNEL_USERNAME)
            except Exception:
                count = "မရရှိနိုင်ပါ (Bot ကို Channel Admin ပေးထားပါသလား စစ်ပါ)"
            events = db.scalars(select(ChannelEvent)).all()
            today_key = now_utc().astimezone(TIMEZONE).strftime("%Y-%m-%d")
            month_key = now_utc().astimezone(TIMEZONE).strftime("%Y-%m")
            joins = sum(1 for e in events if e.event == "join")
            leaves = sum(1 for e in events if e.event == "leave")
            today_joins = sum(1 for e in events if e.event == "join" and local_time(e.created_at).startswith(today_key))
            today_leaves = sum(1 for e in events if e.event == "leave" and local_time(e.created_at).startswith(today_key))
            month_joins = sum(1 for e in events if e.event == "join" and local_time(e.created_at).startswith(month_key))
            month_leaves = sum(1 for e in events if e.event == "leave" and local_time(e.created_at).startswith(month_key))
            text = (f"📢 CHANNEL STATISTICS\n\nCurrent members: {count}\n"
                    f"All-time joins: {joins} | leaves: {leaves}\n"
                    f"Today joins: {today_joins} | leaves: {today_leaves}\n"
                    f"This month joins: {month_joins} | leaves: {month_leaves}\n\n"
                    "မှတ်ချက်: Bot ကို Channel Admin ပေးပြီး Chat Member Updates ခွင့်ပြုထားရပါမယ်။")
        elif section == "admins":
            admins = db.scalars(select(Admin)).all()
            text = "🛡️ ADMINS\n\n" + ("\n".join(f"ID: {a.id}\nPermissions: {a.permissions}" for a in admins) or "Admin မရှိသေးပါ။")
            text += "\n\nထည့်ရန်: /addadmin USER_ID users,lobby,payments,reports,channel\nဖယ်ရန်: /deladmin USER_ID"
        elif section == "manage":
            text = "🔐 USER MANAGEMENT\n\nBan: /ban USER_ID\nUnban: /unban USER_ID\nUser detail: /user USER_ID\n\nPermission: users"
        elif section == "backup":
            if not is_admin(db, q.from_user.id, "reports"):
                await q.answer("ခွင့်ပြုချက်မရှိပါ။", show_alert=True); return
            path = create_excel_backup()
            await q.answer("Backup ဖိုင် ပြင်ဆင်ပြီးပါပြီ။")
            with open(path, "rb") as f:
                await q.message.reply_document(f, filename=os.path.basename(path), caption="📦 YZ Bot Excel Backup")
            return
        else:
            text = "Unknown section"
    await q.answer()
    await q.message.reply_text(text[:4000])

async def addadmin_command(update, context):
    if update.effective_user.id != CREATOR_ID:
        await update.effective_message.reply_text("⛔ Creator သာ အသုံးပြုနိုင်ပါတယ်။"); return
    if len(context.args) < 2:
        await update.effective_message.reply_text("အသုံးပြုပုံ: /addadmin USER_ID users,lobby,payments,reports,channel"); return
    try:
        uid = int(context.args[0])
        allowed = {"users","lobby","payments","reports","channel"}
        perms = set(context.args[1].split(","))
        if not perms or not perms.issubset(allowed):
            raise ValueError
    except ValueError:
        await update.effective_message.reply_text("⚠️ User ID သို့မဟုတ် Permission မမှန်ပါ။"); return
    with Session() as db:
        if uid == CREATOR_ID:
            await update.effective_message.reply_text("Creator ကို Admin အဖြစ် ထပ်ထည့်ရန် မလိုပါ။"); return
        a = db.get(Admin, uid) or Admin(id=uid, added_by=CREATOR_ID)
        a.permissions = json.dumps(sorted(perms))
        db.add(a)
        log_event(db, uid, "ADMIN_ADDED", ",".join(sorted(perms)))
        db.commit()
    new_admin_commands = [BotCommand("start", "Bot စတင်ရန်"), BotCommand("help", "အကူအညီ"),
                          BotCommand("check", "Admin Dashboard")]
    if "reports" in perms: new_admin_commands += [BotCommand("backup", "Excel Backup"), BotCommand("expense", "Record Expense")]
    if "users" in perms:
        new_admin_commands += [BotCommand("user", "User အသေးစိတ်"), BotCommand("ban", "User Ban"), BotCommand("unban", "User Unban")]
    if "channel" in perms:
        new_admin_commands += [BotCommand("kickmember", "Channel Member Ban"), BotCommand("unkickmember", "Channel Member Unban")]
    await context.bot.set_my_commands(new_admin_commands, scope=BotCommandScopeChat(chat_id=uid))
    await update.effective_message.reply_text(f"✅ Admin {uid} ကို ထည့်ပြီးပါပြီ။")
    try: await context.bot.send_message(uid, "🛡️ သင့်ကို YZ Bot Admin အဖြစ် ထည့်သွင်းလိုက်ပါပြီ။")
    except Exception: pass

async def deladmin_command(update, context):
    if update.effective_user.id != CREATOR_ID:
        await update.effective_message.reply_text("⛔ Creator သာ အသုံးပြုနိုင်ပါတယ်။"); return
    if not context.args or not context.args[0].isdigit():
        await update.effective_message.reply_text("အသုံးပြုပုံ: /deladmin USER_ID"); return
    uid = int(context.args[0])
    with Session() as db:
        a = db.get(Admin, uid)
        if a:
            db.delete(a); log_event(db, uid, "ADMIN_REMOVED"); db.commit()
    await update.effective_message.reply_text("✅ Admin ဖယ်ရှားပြီးပါပြီ။")

async def ban_command(update, context):
    if not context.args or not context.args[0].isdigit():
        await update.effective_message.reply_text("အသုံးပြုပုံ: /ban USER_ID"); return
    uid = int(context.args[0])
    with Session() as db:
        if not is_admin(db, update.effective_user.id, "users"):
            await update.effective_message.reply_text("ခွင့်ပြုချက်မရှိပါ။"); return
        u = db.get(User, uid)
        if not u:
            await update.effective_message.reply_text("User မတွေ့ပါ။"); return
        u.is_banned = True; log_event(db, uid, "USER_BANNED", f"by={update.effective_user.id}"); db.commit()
    await update.effective_message.reply_text(f"⛔ User {uid} ကို Ban လုပ်ပြီးပါပြီ။")

async def unban_command(update, context):
    if not context.args or not context.args[0].isdigit():
        await update.effective_message.reply_text("အသုံးပြုပုံ: /unban USER_ID"); return
    uid = int(context.args[0])
    with Session() as db:
        if not is_admin(db, update.effective_user.id, "users"):
            await update.effective_message.reply_text("ခွင့်ပြုချက်မရှိပါ။"); return
        u = db.get(User, uid)
        if not u:
            await update.effective_message.reply_text("User မတွေ့ပါ။"); return
        u.is_banned = False; log_event(db, uid, "USER_UNBANNED", f"by={update.effective_user.id}"); db.commit()
    await update.effective_message.reply_text(f"✅ User {uid} ကို Unban လုပ်ပြီးပါပြီ။")

async def expense_command(update, context):
    if len(context.args) < 2 or not context.args[0].isdigit():
        await update.effective_message.reply_text("အသုံးပြုပုံ: /expense AMOUNT DESCRIPTION\nဥပမာ: /expense 5000 Tournament prize")
        return
    amount = int(context.args[0])
    if amount <= 0 or amount > 1000000000:
        await update.effective_message.reply_text("⚠️ ငွေပမာဏ မမှန်ပါ။")
        return
    description = " ".join(context.args[1:])[:500]
    with Session() as db:
        if not is_admin(db, update.effective_user.id, "reports"):
            await update.effective_message.reply_text("ခွင့်ပြုချက်မရှိပါ။")
            return
        db.add(Expense(amount=amount, category="General", description=description,
                       created_by=update.effective_user.id, created_at=now_utc()))
        log_event(db, update.effective_user.id, "EXPENSE_RECORDED", f"{amount} MMK | {description}")
        db.commit()
    await update.effective_message.reply_text(f"✅ Expense မှတ်တမ်းတင်ပြီးပါပြီ။\n💸 {amount:,} MMK\n📝 {description}")

async def user_command(update, context):
    if not context.args or not context.args[0].isdigit():
        await update.effective_message.reply_text("အသုံးပြုပုံ: /user USER_ID"); return
    uid = int(context.args[0])
    with Session() as db:
        if not is_admin(db, update.effective_user.id, "users"):
            await update.effective_message.reply_text("ခွင့်ပြုချက်မရှိပါ။"); return
        u = db.get(User, uid)
        if not u:
            await update.effective_message.reply_text("User မတွေ့ပါ။"); return
        regs = db.scalars(select(Registration).where(Registration.user_id == uid).order_by(desc(Registration.created_at)).limit(10)).all()
        logs = db.scalars(select(Activity).where(Activity.user_id == uid).order_by(desc(Activity.created_at)).limit(15)).all()
        text = (f"👤 USER DETAIL\nName: {u.first_name} {u.last_name}\nUsername: @{u.username or '—'}\n"
                f"ID: {u.id}\nBanned: {u.is_banned}\nFirst seen: {local_time(u.first_seen)}\nLast seen: {local_time(u.last_seen)}\n\n"
                "Recent registrations:\n")
        for r in regs: text += f"{r.lobby} | {r.game_name} | {r.requested_matches} matches | {r.status} | {local_time(r.created_at)}\n"
        text += "\nRecent activity:\n"
        for a in logs: text += f"{a.event} | {a.details[:100]} | {local_time(a.created_at)}\n"
    await update.effective_message.reply_text(text[:4000])

async def kickmember_command(update, context):
    if not context.args or not context.args[0].isdigit():
        await update.effective_message.reply_text("အသုံးပြုပုံ: /kickmember USER_ID"); return
    uid = int(context.args[0])
    with Session() as db:
        if not is_admin(db, update.effective_user.id, "channel"):
            await update.effective_message.reply_text("ခွင့်ပြုချက်မရှိပါ။"); return
    try:
        await context.bot.ban_chat_member(CHANNEL_USERNAME, uid)
        await update.effective_message.reply_text(f"⛔ Channel member {uid} ကို Ban လုပ်ပြီးပါပြီ။")
    except Exception as exc:
        await update.effective_message.reply_text(f"❌ Channel Ban မအောင်မြင်ပါ: {str(exc)[:300]}")

async def unkickmember_command(update, context):
    if not context.args or not context.args[0].isdigit():
        await update.effective_message.reply_text("အသုံးပြုပုံ: /unkickmember USER_ID"); return
    uid = int(context.args[0])
    with Session() as db:
        if not is_admin(db, update.effective_user.id, "channel"):
            await update.effective_message.reply_text("ခွင့်ပြုချက်မရှိပါ။"); return
    try:
        await context.bot.unban_chat_member(CHANNEL_USERNAME, uid, only_if_banned=True)
        await update.effective_message.reply_text(f"✅ Channel member {uid} ကို Unban လုပ်ပြီးပါပြီ။")
    except Exception as exc:
        await update.effective_message.reply_text(f"❌ Channel Unban မအောင်မြင်ပါ: {str(exc)[:300]}")

# -------------------- BACKUP / RESTART --------------------
def create_excel_backup():
    os.makedirs("backups", exist_ok=True)
    filename = f"backups/yzbot_backup_{datetime.now(TIMEZONE).strftime('%Y%m%d_%H%M%S')}.xlsx"
    wb = Workbook()
    wb.remove(wb.active)
    tables = [
        ("Users", User), ("Admins", Admin), ("Registrations", Registration),
        ("Payments", Payment), ("Expenses", Expense), ("Matches", Match), ("MatchPlayers", MatchPlayer),
        ("ActivityLogs", Activity), ("ChannelEvents", ChannelEvent)
    ]
    with Session() as db:
        for sheet_name, model in tables:
            ws = wb.create_sheet(sheet_name)
            rows = db.scalars(select(model)).all()
            cols = list(model.__table__.columns)
            ws.append([c.name for c in cols])
            for row in rows:
                values = []
                for c in cols:
                    value = getattr(row, c.name)
                    if isinstance(value, datetime):
                        value = local_time(value)
                    values.append(value)
                ws.append(values)
            for cell in ws[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="243B53")
                cell.alignment = Alignment(horizontal="center")
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for col in ws.columns:
                width = min(max(max(len(str(c.value or "")) for c in col) + 2, 12), 45)
                ws.column_dimensions[get_column_letter(col[0].column)].width = width
    wb.save(filename)
    if not os.path.exists(filename) or os.path.getsize(filename) < 100:
        raise RuntimeError("Excel backup verification failed")
    return filename

async def backup_command(update, context):
    with Session() as db:
        if not is_admin(db, update.effective_user.id, "reports"):
            await update.effective_message.reply_text("⛔ ခွင့်ပြုချက်မရှိပါ။"); return
    path = create_excel_backup()
    with open(path, "rb") as f:
        await update.effective_message.reply_document(f, filename=os.path.basename(path), caption="📦 Database Backup (Myanmar Time)")

async def restart_command(update, context):
    if update.effective_user.id != CREATOR_ID:
        await update.effective_message.reply_text("⛔ Creator တစ်ဦးတည်းသာ Restart လုပ်နိုင်ပါတယ်။"); return
    await update.effective_message.reply_text("📦 Excel Backup ပြုလုပ်နေပါတယ်။ Backup မအောင်မြင်ပါက Restart မလုပ်ပါ။")
    try:
        path = create_excel_backup()
        # Deliver and wait for Telegram to accept the file before restarting.
        with open(path, "rb") as f:
            await context.bot.send_document(CREATOR_ID, f, filename=os.path.basename(path),
                                            caption="✅ Restart မတိုင်မီ မဖြစ်မနေ Backup")
        await update.effective_message.reply_text("✅ Excel Backup အောင်မြင်ပါပြီ။ Bot ကို ပြန်စတင်နေပါပြီ။")
        await asyncio.sleep(2)
        os.execv(sys.executable, [sys.executable] + sys.argv)
    except Exception as exc:
        log.exception("Protected restart failed")
        await update.effective_message.reply_text(f"❌ Backup မအောင်မြင်သဖြင့် Restart မလုပ်ပါ။\n{str(exc)[:500]}")

# -------------------- CHANNEL MEMBER EVENTS --------------------
async def channel_member_update(update, context):
    cm = update.chat_member
    if not cm:
        return
    old = cm.old_chat_member.status
    new = cm.new_chat_member.status
    joined = (new in (ChatMemberStatus.MEMBER, ChatMemberStatus.ADMINISTRATOR, ChatMemberStatus.OWNER)
              or (new == ChatMemberStatus.RESTRICTED and getattr(cm.new_chat_member, "is_member", False)))
    was_joined = (old in (ChatMemberStatus.MEMBER, ChatMemberStatus.ADMINISTRATOR, ChatMemberStatus.OWNER)
                  or (old == ChatMemberStatus.RESTRICTED and getattr(cm.old_chat_member, "is_member", False)))
    if joined == was_joined:
        return
    event = "join" if joined else "leave"
    with Session() as db:
        db.add(ChannelEvent(user_id=cm.new_chat_member.user.id, event=event, created_at=now_utc()))
        log_event(db, cm.new_chat_member.user.id, f"CHANNEL_{event.upper()}")
        db.commit()

async def my_chat_member_update(update, context):
    cm = update.my_chat_member
    if not cm:
        return
    user_id = cm.chat.id
    old_status = cm.old_chat_member.status
    new_status = cm.new_chat_member.status

    # Only private-chat block/unblock transitions are relevant here.
    if new_status in (ChatMemberStatus.LEFT, ChatMemberStatus.KICKED):
        event_name = "BOT_BLOCKED"
        new_bot_status = "blocked"
    elif new_status == ChatMemberStatus.MEMBER:
        event_name = "BOT_UNBLOCKED"
        new_bot_status = "active"
    else:
        return

    with Session() as db:
        u = db.get(User, user_id)
        if u:
            # Ignore duplicate Telegram membership updates.
            if u.bot_status == new_bot_status:
                return
            u.bot_status = new_bot_status
            log_event(db, user_id, event_name)
            db.commit()
        else:
            # Do not reference an event variable conditionally; the transition
            # is valid even if the user has not yet sent /start to this bot.
            db.add(Activity(user_id=user_id, event=event_name,
                            details="Telegram private-chat membership update",
                            created_at=now_utc()))
            db.commit()
    await notify_staff(context, f"🔔 Bot {event_name}: Telegram User ID {user_id} | Time: {local_time(now_utc())}")

async def verify_join_callback(update, context):
    q = update.callback_query
    if await joined_channel(context.bot, q.from_user.id):
        await q.answer("✅ Channel Join ဖြစ်ကြောင်း အတည်ပြုပြီးပါပြီ။")
        await q.message.reply_text(WELCOME, reply_markup=main_keyboard(), parse_mode="Markdown")
    else:
        await q.answer("❌ Channel Join မတွေ့သေးပါ။ Join ပြီးမှ ပြန်နှိပ်ပါ။", show_alert=True)

# -------------------- RATE LIMIT / COMMANDS --------------------
async def rate_limit(update, context):
    user = update.effective_user
    if not user or user.id == CREATOR_ID:
        return
    now = asyncio.get_running_loop().time()
    times = context.user_data.setdefault("_rate", [])
    times[:] = [t for t in times if now - t < 10]
    if len(times) >= 10:
        if update.callback_query:
            await update.callback_query.answer("🐢 ခဏစောင့်ပြီးမှ ထပ်မံအသုံးပြုပါ။", show_alert=True)
        elif update.effective_message:
            await update.effective_message.reply_text("🐢 Spam ကာကွယ်ရေးစနစ်ကြောင့် ခဏစောင့်ပြီးမှ ထပ်ပို့ပေးပါ။")
        raise ApplicationHandlerStop
    times.append(now)

async def set_commands(app):
    user_commands = [
        BotCommand("start", "Bot စတင်ရန်"),
        BotCommand("help", "အကူအညီ")
    ]
    admin_commands = user_commands + [
        BotCommand("check", "Admin Dashboard"),
        BotCommand("backup", "Excel Backup"),
        BotCommand("user", "User အသေးစိတ်"),
        BotCommand("ban", "User Ban"),
        BotCommand("unban", "User Unban"),
        BotCommand("kickmember", "Channel Member Ban"),
        BotCommand("unkickmember", "Channel Member Unban")
    ]
    creator_commands = admin_commands + [
        BotCommand("addadmin", "Admin ထည့်ရန်"),
        BotCommand("deladmin", "Admin ဖယ်ရန်"),
        BotCommand("restart", "Backup ပြီး Restart"),
        BotCommand("expense", "Record Expense")
    ]
    await app.bot.set_my_commands(user_commands, scope=BotCommandScopeDefault())
    await app.bot.set_my_commands(creator_commands, scope=BotCommandScopeChat(chat_id=CREATOR_ID))
    with Session() as db:
        for a in db.scalars(select(Admin)).all():
            perms = set(json.loads(a.permissions or "[]"))
            cmds = user_commands + [BotCommand("check", "Admin Dashboard")]
            if "reports" in perms: cmds += [BotCommand("backup", "Excel Backup"), BotCommand("expense", "Record Expense")]
            if "users" in perms:
                cmds += [BotCommand("user", "User အသေးစိတ်"), BotCommand("ban", "User Ban"), BotCommand("unban", "User Unban")]
            if "channel" in perms:
                cmds += [BotCommand("kickmember", "Channel Member Ban"), BotCommand("unkickmember", "Channel Member Unban")]
            await app.bot.set_my_commands(cmds, scope=BotCommandScopeChat(chat_id=a.id))

async def help_command(update, context):
    await update.effective_message.reply_text(
        "🤖 YZ's Bot အသုံးပြုပုံ\n\n/start - Bot စတင်ရန်\n"
        "🆓 Free Lobby / 💰 3K Lobby မှတစ်ဆင့် စာရင်းသွင်းနိုင်ပါတယ်။\n"
        "3K Lobby အတွက် ငွေလွှဲပြေစာကို ပို့ပြီး Admin အတည်ပြုချက်ကို စောင့်ပါ။"
    )

def main():
    if not BOT_TOKEN:
        raise RuntimeError("BOT_TOKEN environment variable မသတ်မှတ်ရသေးပါ။")
    app = Application.builder().token(BOT_TOKEN).post_init(set_commands).build()
    app.add_handler(MessageHandler(filters.ALL, rate_limit), group=-1)
    app.add_handler(CallbackQueryHandler(rate_limit), group=-1)
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("help", help_command))
    app.add_handler(CommandHandler("check", check_command))
    app.add_handler(CommandHandler("addadmin", addadmin_command))
    app.add_handler(CommandHandler("deladmin", deladmin_command))
    app.add_handler(CommandHandler("ban", ban_command))
    app.add_handler(CommandHandler("unban", unban_command))
    app.add_handler(CommandHandler("kickmember", kickmember_command))
    app.add_handler(CommandHandler("unkickmember", unkickmember_command))
    app.add_handler(CommandHandler("user", user_command))
    app.add_handler(CommandHandler("backup", backup_command))
    app.add_handler(CommandHandler("expense", expense_command))
    app.add_handler(CommandHandler("restart", restart_command))
    app.add_handler(CallbackQueryHandler(verify_join_callback, pattern=r"^verify_join$"))
    app.add_handler(CallbackQueryHandler(home_callback, pattern=r"^home$"))
    app.add_handler(CallbackQueryHandler(lobby_callback, pattern=r"^lobby:"))
    app.add_handler(CallbackQueryHandler(field_callback, pattern=r"^field:"))
    app.add_handler(CallbackQueryHandler(payment_method_callback, pattern=r"^paymethod:"))
    app.add_handler(CallbackQueryHandler(payment_review_callback, pattern=r"^payment:"))
    app.add_handler(CallbackQueryHandler(dash_callback, pattern=r"^dash:"))
    app.add_handler(ChatMemberHandler(channel_member_update, ChatMemberHandler.CHAT_MEMBER))
    app.add_handler(ChatMemberHandler(my_chat_member_update, ChatMemberHandler.MY_CHAT_MEMBER))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, text_router))
    app.add_handler(MessageHandler((filters.PHOTO | filters.Document.IMAGE) & ~filters.COMMAND, text_router))
    app.add_error_handler(error_handler)
    log.info("YZ's Bot starting...")
    app.run_polling(
        allowed_updates=["message", "callback_query", "chat_member", "my_chat_member"],
        drop_pending_updates=False
    )

if __name__ == "__main__":
    main()
