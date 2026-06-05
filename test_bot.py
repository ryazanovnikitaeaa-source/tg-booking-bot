"""
test_bot.py — автотесты бизнес-логики bot.py
Запуск: pytest test_bot.py -v

Тестируются только чистые функции:
  - check_slot()
  - get_time_slots_kb()
  - get_user_future_bookings()

Telegram API и xlsx не вызываются: read_all_bookings() мокается,
datetime.now() мокается через подкласс real_datetime.
"""

import sys
import os
from datetime import datetime as real_datetime
from unittest.mock import patch

# ─── Гарантируем, что bot.py находится в пути ─────────────
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import bot  # aiogram и openpyxl устанавливаются, но сетевых вызовов нет


# ══════════════════════════════════════════════════════════════════
#  Вспомогательные утилиты
# ══════════════════════════════════════════════════════════════════

DATE_FUTURE = "2099-06-10"   # гарантированно далёкое будущее
DATE_PAST   = "2000-01-01"   # гарантированное прошлое


def make_booking(time_str: str, telegram_id=111, date_str=DATE_FUTURE) -> dict:
    """Фабрика тестовых записей (минимальный набор полей)."""
    return {
        "id":          "test-id",
        "service":     "Запись",
        "name":        "Test",
        "date":        date_str,
        "time":        time_str,
        "telegram_id": str(telegram_id),
        "username":    "testuser",
    }


def fake_datetime_now(fixed_now: real_datetime):
    """
    Возвращает подкласс real_datetime, у которого now() возвращает fixed_now,
    но strptime() работает как обычно (унаследован).
    Используется как patch('bot.datetime', fake_datetime_now(now)).
    """
    class _FakeDatetime(real_datetime):
        @classmethod
        def now(cls, tz=None):
            return fixed_now

    return _FakeDatetime


def slot_texts(kb) -> list[str]:
    """Список текстов кнопок клавиатуры (без кнопки «Отмена»)."""
    return [
        btn.text
        for row in kb.inline_keyboard
        for btn in row
        if btn.text != "❌ Отмена"
    ]


# ══════════════════════════════════════════════════════════════════
#  check_slot()
# ══════════════════════════════════════════════════════════════════

class TestCheckSlot:
    """Тесты для check_slot(date_str, time_str)."""

    # ── Рабочее время ───────────────────────────────────────────

    def test_before_work_start_returns_hours(self):
        with patch("bot.read_all_bookings", return_value=[]):
            assert bot.check_slot(DATE_FUTURE, "08:59") == "hours"

    def test_exactly_work_start_is_free(self):
        with patch("bot.read_all_bookings", return_value=[]):
            assert bot.check_slot(DATE_FUTURE, "09:00") is None

    def test_exactly_work_end_16_00_is_free(self):
        with patch("bot.read_all_bookings", return_value=[]):
            assert bot.check_slot(DATE_FUTURE, "16:00") is None

    def test_after_work_end_16_01_returns_hours(self):
        with patch("bot.read_all_bookings", return_value=[]):
            assert bot.check_slot(DATE_FUTURE, "16:01") == "hours"

    def test_17_00_returns_hours(self):
        with patch("bot.read_all_bookings", return_value=[]):
            assert bot.check_slot(DATE_FUTURE, "17:00") == "hours"

    def test_midnight_returns_hours(self):
        with patch("bot.read_all_bookings", return_value=[]):
            assert bot.check_slot(DATE_FUTURE, "00:00") == "hours"

    # ── Формат ─────────────────────────────────────────────────

    def test_non_time_string_returns_format(self):
        with patch("bot.read_all_bookings", return_value=[]):
            assert bot.check_slot(DATE_FUTURE, "bad") == "format"

    def test_empty_string_returns_format(self):
        with patch("bot.read_all_bookings", return_value=[]):
            assert bot.check_slot(DATE_FUTURE, "") == "format"

    def test_only_hours_returns_format(self):
        with patch("bot.read_all_bookings", return_value=[]):
            assert bot.check_slot(DATE_FUTURE, "12") == "format"

    # ── Лимит дня ──────────────────────────────────────────────

    def test_day_full_returns_full(self):
        """При 5 записях в день → full."""
        bookings = [make_booking(f"{9 + i}:00") for i in range(5)]
        with patch("bot.read_all_bookings", return_value=bookings):
            assert bot.check_slot(DATE_FUTURE, "15:00") == "full"

    def test_four_bookings_not_full(self):
        """4 записи — ещё не лимит (лимит = 5)."""
        bookings = [make_booking(f"{9 + i}:00") for i in range(4)]
        with patch("bot.read_all_bookings", return_value=bookings):
            # 14:00 свободно и нет конфликтов (разница ровно 60 от 13:00)
            assert bot.check_slot(DATE_FUTURE, "14:00") is None

    # ── Интервал ───────────────────────────────────────────────

    def test_interval_conflict_30_min_before(self):
        """Занято 10:00 → 09:30 конфликтует (30 мин < 60)."""
        bookings = [make_booking("10:00")]
        with patch("bot.read_all_bookings", return_value=bookings):
            result = bot.check_slot(DATE_FUTURE, "09:30")
            assert result is not None
            assert result.startswith("interval:")

    def test_interval_conflict_same_time(self):
        """Попытка записаться точно на то же время."""
        bookings = [make_booking("12:00")]
        with patch("bot.read_all_bookings", return_value=bookings):
            result = bot.check_slot(DATE_FUTURE, "12:00")
            assert result is not None
            assert result.startswith("interval:")

    def test_interval_conflict_30_min_after(self):
        """Занято 10:00 → 10:30 конфликтует (30 мин < 60)."""
        bookings = [make_booking("10:00")]
        with patch("bot.read_all_bookings", return_value=bookings):
            result = bot.check_slot(DATE_FUTURE, "10:30")
            assert result is not None
            assert result.startswith("interval:")

    def test_interval_exactly_60_min_is_free(self):
        """Ровно 60 минут — допустимо (условие строгое: < 60)."""
        bookings = [make_booking("10:00")]
        with patch("bot.read_all_bookings", return_value=bookings):
            assert bot.check_slot(DATE_FUTURE, "11:00") is None

    def test_interval_exactly_60_min_before_is_free(self):
        bookings = [make_booking("11:00")]
        with patch("bot.read_all_bookings", return_value=bookings):
            assert bot.check_slot(DATE_FUTURE, "10:00") is None

    def test_interval_result_contains_conflicting_time(self):
        """Строка interval: содержит время занятой записи."""
        bookings = [make_booking("14:00")]
        with patch("bot.read_all_bookings", return_value=bookings):
            result = bot.check_slot(DATE_FUTURE, "14:30")
            assert result == "interval:14:00"

    # ── Свободный слот ─────────────────────────────────────────

    def test_free_slot_no_bookings(self):
        with patch("bot.read_all_bookings", return_value=[]):
            assert bot.check_slot(DATE_FUTURE, "12:00") is None

    def test_free_slot_far_from_existing(self):
        bookings = [make_booking("09:00"), make_booking("11:00")]
        with patch("bot.read_all_bookings", return_value=bookings):
            assert bot.check_slot(DATE_FUTURE, "13:00") is None


# ══════════════════════════════════════════════════════════════════
#  get_time_slots_kb()
# ══════════════════════════════════════════════════════════════════

class TestGetTimeSlotsKb:
    """Тесты для get_time_slots_kb(date_str)."""

    # ── Прошедшее время ────────────────────────────────────────

    def test_all_slots_past_returns_none(self):
        """Если «сейчас» 17:00, все слоты уже прошли → None."""
        now = real_datetime(2099, 6, 10, 17, 0)
        with patch("bot.read_all_bookings", return_value=[]):
            with patch("bot.datetime", fake_datetime_now(now)):
                result = bot.get_time_slots_kb("2099-06-10")
        assert result is None

    def test_past_slots_not_shown(self):
        """Слоты до «сейчас» отсутствуют, слоты после — присутствуют."""
        now = real_datetime(2099, 6, 10, 14, 1)   # 14:01 — 14:00 и раньше уже прошли
        with patch("bot.read_all_bookings", return_value=[]):
            with patch("bot.datetime", fake_datetime_now(now)):
                kb = bot.get_time_slots_kb("2099-06-10")
        assert kb is not None
        texts = slot_texts(kb)
        assert "09:00" not in texts
        assert "14:00" not in texts
        assert "14:30" in texts    # 14:30 ещё в будущем

    def test_slot_exactly_at_now_excluded(self):
        """Слот ровно в «сейчас» не показывается (условие: slot_dt <= now)."""
        now = real_datetime(2099, 6, 10, 12, 0)
        with patch("bot.read_all_bookings", return_value=[]):
            with patch("bot.datetime", fake_datetime_now(now)):
                kb = bot.get_time_slots_kb("2099-06-10")
        assert kb is not None
        assert "12:00" not in slot_texts(kb)
        assert "12:30" in slot_texts(kb)

    # ── Граница 16:30 ──────────────────────────────────────────

    def test_16_00_is_last_slot(self):
        """16:00 — присутствует, 16:30 — никогда не появляется."""
        now = real_datetime(2099, 6, 10, 8, 0)
        with patch("bot.read_all_bookings", return_value=[]):
            with patch("bot.datetime", fake_datetime_now(now)):
                kb = bot.get_time_slots_kb("2099-06-10")
        assert kb is not None
        texts = slot_texts(kb)
        assert "16:00" in texts
        assert "16:30" not in texts

    # ── Конфликт по интервалу ──────────────────────────────────

    def test_booked_10_00_excludes_09_30_and_10_30(self):
        """Занято 10:00 → 09:30 и 10:30 не показываются."""
        now = real_datetime(2099, 6, 10, 8, 0)
        b = make_booking("10:00", date_str="2099-06-10")
        with patch("bot.read_all_bookings", return_value=[b]):
            with patch("bot.datetime", fake_datetime_now(now)):
                kb = bot.get_time_slots_kb("2099-06-10")
        assert kb is not None
        texts = slot_texts(kb)
        assert "09:30" not in texts
        assert "10:00" not in texts
        assert "10:30" not in texts

    def test_booked_10_00_keeps_09_00_and_11_00(self):
        """Занято 10:00 → 09:00 и 11:00 доступны (ровно 60 мин)."""
        now = real_datetime(2099, 6, 10, 8, 0)
        b = make_booking("10:00", date_str="2099-06-10")
        with patch("bot.read_all_bookings", return_value=[b]):
            with patch("bot.datetime", fake_datetime_now(now)):
                kb = bot.get_time_slots_kb("2099-06-10")
        assert kb is not None
        texts = slot_texts(kb)
        assert "09:00" in texts
        assert "11:00" in texts

    def test_dense_bookings_eliminate_all_slots_returns_none(self):
        """Записи в каждый получасовой слот закрывают всю сетку."""
        now = real_datetime(2099, 6, 10, 8, 0)
        # Записи в каждые 30 минут с 09:30 до 15:30
        # → каждый слот от 09:00 до 16:00 находится в зоне <60 мин от одной из записей
        busy_times = ["09:30", "10:30", "11:30", "12:30", "13:30", "14:30", "15:30"]
        bookings = [make_booking(t, date_str="2099-06-10") for t in busy_times]
        with patch("bot.read_all_bookings", return_value=bookings):
            with patch("bot.datetime", fake_datetime_now(now)):
                result = bot.get_time_slots_kb("2099-06-10")
        assert result is None

    # ── Структура клавиатуры ───────────────────────────────────

    def test_has_cancel_button(self):
        """В клавиатуре всегда есть кнопка «Отмена»."""
        now = real_datetime(2099, 6, 10, 8, 0)
        with patch("bot.read_all_bookings", return_value=[]):
            with patch("bot.datetime", fake_datetime_now(now)):
                kb = bot.get_time_slots_kb("2099-06-10")
        assert kb is not None
        all_texts = [btn.text for row in kb.inline_keyboard for btn in row]
        assert "❌ Отмена" in all_texts

    def test_callback_data_format(self):
        """Кнопки слотов имеют callback_data вида timeslot:HH:MM."""
        now = real_datetime(2099, 6, 10, 8, 0)
        with patch("bot.read_all_bookings", return_value=[]):
            with patch("bot.datetime", fake_datetime_now(now)):
                kb = bot.get_time_slots_kb("2099-06-10")
        assert kb is not None
        slot_buttons = [
            btn
            for row in kb.inline_keyboard
            for btn in row
            if btn.text != "❌ Отмена"
        ]
        assert len(slot_buttons) > 0
        for btn in slot_buttons:
            assert btn.callback_data.startswith("timeslot:"), (
                f"Неверный callback_data: {btn.callback_data!r}"
            )

    def test_max_3_buttons_per_row(self):
        """В каждом ряду не более 3 кнопок."""
        now = real_datetime(2099, 6, 10, 8, 0)
        with patch("bot.read_all_bookings", return_value=[]):
            with patch("bot.datetime", fake_datetime_now(now)):
                kb = bot.get_time_slots_kb("2099-06-10")
        assert kb is not None
        for row in kb.inline_keyboard:
            assert len(row) <= 3


# ══════════════════════════════════════════════════════════════════
#  get_user_future_bookings()
# ══════════════════════════════════════════════════════════════════

class TestGetUserFutureBookings:
    """Тесты для get_user_future_bookings(telegram_id)."""

    USER_ID = 42
    NOW     = real_datetime(2025, 6, 10, 12, 0)   # «текущее» время тестов

    def _patch_now(self):
        return patch("bot.datetime", fake_datetime_now(self.NOW))

    # ── Базовые случаи ─────────────────────────────────────────

    def test_future_booking_included(self):
        """Будущая запись пользователя включается."""
        b = make_booking("14:00", telegram_id=self.USER_ID, date_str="2099-01-01")
        with patch("bot.read_all_bookings", return_value=[b]):
            with self._patch_now():
                result = bot.get_user_future_bookings(self.USER_ID)
        assert len(result) == 1

    def test_past_booking_excluded(self):
        """Прошедшая запись не включается."""
        b = make_booking("10:00", telegram_id=self.USER_ID, date_str="2000-01-01")
        with patch("bot.read_all_bookings", return_value=[b]):
            with self._patch_now():
                result = bot.get_user_future_bookings(self.USER_ID)
        assert result == []

    def test_booking_exactly_now_excluded(self):
        """Запись ровно в «сейчас» — не в будущем (dt > now, не >=)."""
        b = make_booking("12:00", telegram_id=self.USER_ID, date_str="2025-06-10")
        with patch("bot.read_all_bookings", return_value=[b]):
            with self._patch_now():
                result = bot.get_user_future_bookings(self.USER_ID)
        assert result == []

    def test_empty_bookings_returns_empty(self):
        with patch("bot.read_all_bookings", return_value=[]):
            with self._patch_now():
                result = bot.get_user_future_bookings(self.USER_ID)
        assert result == []

    # ── Чужие записи ───────────────────────────────────────────

    def test_other_user_excluded(self):
        """Запись чужого пользователя не попадает в результат."""
        b = make_booking("14:00", telegram_id=999, date_str="2099-01-01")
        with patch("bot.read_all_bookings", return_value=[b]):
            with self._patch_now():
                result = bot.get_user_future_bookings(self.USER_ID)
        assert result == []

    def test_only_own_future_returned(self):
        """Из смешанного списка возвращаются только будущие записи нужного пользователя."""
        own_future   = make_booking("15:00", telegram_id=self.USER_ID, date_str="2099-01-01")
        own_past     = make_booking("09:00", telegram_id=self.USER_ID, date_str="2000-01-01")
        other_future = make_booking("11:00", telegram_id=999,          date_str="2099-01-01")
        all_bookings = [own_future, own_past, other_future]
        with patch("bot.read_all_bookings", return_value=all_bookings):
            with self._patch_now():
                result = bot.get_user_future_bookings(self.USER_ID)
        assert len(result) == 1
        assert result[0]["time"] == "15:00"

    def test_multiple_own_future_all_returned(self):
        """Несколько будущих записей одного пользователя — все включаются."""
        b1 = make_booking("10:00", telegram_id=self.USER_ID, date_str="2099-01-01")
        b2 = make_booking("14:00", telegram_id=self.USER_ID, date_str="2099-01-02")
        b3 = make_booking("09:00", telegram_id=self.USER_ID, date_str="2099-01-03")
        with patch("bot.read_all_bookings", return_value=[b1, b2, b3]):
            with self._patch_now():
                result = bot.get_user_future_bookings(self.USER_ID)
        assert len(result) == 3

    # ── Тип telegram_id ────────────────────────────────────────

    def test_int_and_str_telegram_id_match(self):
        """Функция принимает int-id, а в данных он хранится как str."""
        b = make_booking("14:00", telegram_id=str(self.USER_ID), date_str="2099-01-01")
        # make_booking уже сохраняет telegram_id как str; проверяем сравнение
        with patch("bot.read_all_bookings", return_value=[b]):
            with self._patch_now():
                result = bot.get_user_future_bookings(self.USER_ID)
        assert len(result) == 1
