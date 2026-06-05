# 🤖 Бот для записи клиентов / Booking Bot

[🇷🇺 Русский](#-русский) | [🇬🇧 English](#-english)

---

## 🇷🇺 Русский

Telegram-бот для онлайн-записи клиентов к мастеру.  
Позволяет клиентам самостоятельно выбирать дату и время,  
а владелец получает мгновенные уведомления о каждой новой или отменённой записи.

---

### ✨ Функционал

#### Для клиента
- **Запись к мастеру** — выбор даты через инлайн-календарь и времени через кнопки
- **Выбор времени кнопками** — доступные слоты на основе настроек рабочего времени
- **Ограничение**: не более 1 активной записи на клиента
- **Отмена записи** — через reply-кнопку «❌ Отменить запись» в главном меню
- **Связь с мастером** — кнопка с прямой ссылкой на Telegram мастера

#### Для владельца
- **Уведомления** о каждой новой записи (имя, дата, время, Telegram клиента)
- **Уведомления об отменах** с полными данными клиента
- **Команда `/admin`** — список всех будущих записей на ближайшие 14 дней

#### Бизнес-логика (настраивается в `config.py`)
- Рабочие часы: **09:00 – 16:00** (настраивается)
- Максимум **5 записей в день** (настраивается)
- Минимальный интервал между записями — **60 минут** (настраивается)
- Конфликтующие и прошедшие слоты **не отображаются** клиенту
- Все записи хранятся в файле **`zayavki.xlsx`**

---

### 🛠 Стек

| Компонент     | Технология              |
|---------------|-------------------------|
| Язык          | Python 3.10+            |
| Telegram API  | aiogram 3               |
| Хранилище     | openpyxl (`.xlsx`)      |
| Настройки     | `config.py`             |
| Конфигурация  | python-dotenv (`.env`)  |
| Тесты         | pytest                  |

---

### 🚀 Как запустить

**1. Клонировать репозиторий**
```bash
git clone https://github.com/ваш-username/tg-booking-bot.git
cd tg-booking-bot
```

**2. Установить зависимости**
```bash
pip install -r requirements.txt
```

**3. Создать файл `.env`**
```env
BOT_TOKEN=ваш_токен_от_BotFather
OWNER_ID=ваш_telegram_id
```

**4. Настроить `config.py`**
- `MASTER_NAME` — имя мастера
- `MASTER_USERNAME` — Telegram username (без `@`)
- `WORK_START` / `WORK_END` — рабочие часы
- `MAX_PER_DAY` — максимум записей в день
- `MIN_INTERVAL` — интервал между записями (минуты)
- `DAYS_AHEAD` — на сколько дней вперёд открыта запись
- `WELCOME_TEXT` — приветственное сообщение

**5. Запустить**
```bash
py bot.py
```

### 🧪 Тесты
```bash
py -m pytest test_bot.py -v
```

### 📋 Команды
| Команда  | Описание                                      |
|----------|-----------------------------------------------|
| `/start` | Главное меню                                  |
| `/admin` | Записи на 14 дней (только для владельца)      |

---

## 🇬🇧 English

A Telegram bot for client appointment booking.  
Clients can independently choose a date and time,  
while the owner receives instant notifications about every new or cancelled booking.

---

### ✨ Features

#### For clients
- **Book an appointment** — choose date via inline calendar and time via buttons
- **Smart time slots** — only available slots are shown based on work schedule
- **Limit**: maximum 1 active booking per client
- **Cancel booking** — via the «❌ Cancel booking» reply button in the main menu
- **Contact master** — button with a direct Telegram link to the master

#### For the owner
- **Notifications** for every new booking (name, date, time, client Telegram)
- **Cancellation notifications** with full client details
- **`/admin` command** — all upcoming bookings for the next 14 days, grouped by date

#### Business logic (configurable via `config.py`)
- Working hours: **09:00 – 16:00** (configurable)
- Maximum **5 bookings per day** (configurable)
- Minimum interval between bookings — **60 minutes** (configurable)
- Conflicting and past slots are **hidden** from clients
- All bookings stored in **`zayavki.xlsx`**

---

### 🛠 Stack

| Component     | Technology              |
|---------------|-------------------------|
| Language      | Python 3.10+            |
| Telegram API  | aiogram 3               |
| Storage       | openpyxl (`.xlsx`)      |
| Settings      | `config.py`             |
| Configuration | python-dotenv (`.env`)  |
| Tests         | pytest                  |

---

### 🚀 Getting started

**1. Clone the repository**
```bash
git clone https://github.com/your-username/tg-booking-bot.git
cd tg-booking-bot
```

**2. Install dependencies**
```bash
pip install -r requirements.txt
```

**3. Create `.env` file**
```env
BOT_TOKEN=your_token_from_BotFather
OWNER_ID=your_telegram_id
```

**4. Configure `config.py`**
- `MASTER_NAME` — master's name
- `MASTER_USERNAME` — Telegram username (without `@`)
- `WORK_START` / `WORK_END` — working hours
- `MAX_PER_DAY` — max bookings per day
- `MIN_INTERVAL` — interval between bookings (minutes)
- `DAYS_AHEAD` — how many days ahead booking is open
- `WELCOME_TEXT` — welcome message

**5. Run**
```bash
py bot.py
```

### 🧪 Tests
```bash
py -m pytest test_bot.py -v
```

### 📋 Commands
| Command  | Description                                   |
|----------|-----------------------------------------------|
| `/start` | Main menu                                     |
| `/admin` | Bookings for 14 days (owner only)             |

---

### 📁 Project structure
```
.
├── bot.py              # main bot code
├── config.py           # master settings
├── test_bot.py         # business logic tests
├── requirements.txt    # dependencies
├── .env                # token & owner ID (not in git!)
├── .gitignore
└── zayavki.xlsx        # bookings file (auto-created)
```