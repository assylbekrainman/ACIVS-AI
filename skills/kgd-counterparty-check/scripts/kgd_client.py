#!/usr/bin/env python3
"""
Клиент API портала КГД МФ РК «Сведения по контрагентам (Открытые данные)».

Использование:
    export KGD_PORTAL_TOKEN="<токен от Администратора КГД>"
    python3 kgd_client.py 071140005693
    python3 kgd_client.py 071140005693 990140001234 --json

Переменные окружения:
    KGD_PORTAL_HOST   хост портала, по умолчанию https://portal.kgd.gov.kz
    KGD_PORTAL_TOKEN  токен X-Portal-Token (обязателен)
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path

import requests

DEFAULT_HOST = "https://portal.kgd.gov.kz"
ENDPOINT_PATH = "/services/isnaportal/public/get-sur-data"
XIN_RE = re.compile(r"^\d{12}$")

ERROR_MESSAGES = {
    400: "Запрос содержит синтаксическую ошибку (проверьте формат ИИН/БИН — 12 цифр)",
    401: "Пользователь не авторизован — проверьте KGD_PORTAL_TOKEN",
    404: "Доступ к сервису запрещён — обратитесь к Администратору КГД (токен не активирован)",
    500: "Ошибка на сервере КГД, запрос не выполнен — повторите позже",
}

# поле в ответе -> человекочитаемое описание красного флага
FLAG_FIELDS = {
    "bankrupt": "контрагент в реестре банкротов",
    "inactive": "налогоплательщик признан бездействующим",
    "registrationInvalid": "регистрация признана недействительной",
    "reRegistrationInvalid": "перерегистрация признана недействительной",
    "operationsWOWork": "сделки без фактического выполнения работ/услуг",
    "esfRestrinctions": "ограничения на выписку ЭСФ",
    "courtDecisionRegistry": "есть решение суда в реестре",
}


class KgdApiError(Exception):
    def __init__(self, status_code: int, message: str):
        self.status_code = status_code
        super().__init__(message)


def _load_dotenv(path: Path) -> None:
    """Подхватывает KGD_PORTAL_HOST/KGD_PORTAL_TOKEN из .env, если он есть
    рядом со скриптом и переменные ещё не заданы в окружении."""
    if not path.is_file():
        return
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, _, value = line.partition("=")
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        os.environ.setdefault(key, value)


@dataclass
class KgdClient:
    host: str = field(default_factory=lambda: os.environ.get("KGD_PORTAL_HOST", DEFAULT_HOST))
    token: str | None = field(default_factory=lambda: os.environ.get("KGD_PORTAL_TOKEN"))
    timeout: int = 15

    def get_sur_data(self, xin: str) -> dict:
        if not XIN_RE.match(xin):
            raise ValueError(f"ИИН/БИН должен состоять из 12 цифр, получено: {xin!r}")
        if not self.token:
            raise RuntimeError(
                "Не задан KGD_PORTAL_TOKEN. Получите токен у Администратора портала "
                "КГД МФ РК и задайте переменную окружения KGD_PORTAL_TOKEN."
            )

        url = self.host.rstrip("/") + ENDPOINT_PATH
        response = requests.post(
            url,
            headers={"X-Portal-Token": self.token, "Content-Type": "application/json"},
            json={"xin": xin},
            timeout=self.timeout,
        )

        if response.status_code != 200:
            message = ERROR_MESSAGES.get(
                response.status_code, f"Неожиданный код ответа: {response.status_code}"
            )
            raise KgdApiError(response.status_code, message)

        return response.json()


def detect_flags(data: dict) -> list[str]:
    """Возвращает список сработавших красных флагов на основе ответа API."""
    flags: list[str] = []
    for field_name, description in FLAG_FIELDS.items():
        value = (data.get(field_name) or {}).get("ru", "")
        if value and value != "Нет данных":
            flags.append(f"{description} ({field_name}: «{value}»)")

    reg_absent = (data.get("regAddressAbsent") or {}).get("ru", "")
    if reg_absent == "Да":
        flags.append("отсутствует по юридическому адресу (regAddressAbsent: «Да»)")

    tax_debt = data.get("taxDebt") or 0
    if tax_debt and tax_debt > 0:
        flags.append(f"есть налоговая задолженность: {tax_debt:,.2f} тенге".replace(",", " "))

    return flags


def format_summary(data: dict) -> str:
    name = data.get("name", {}).get("ru") or "(наименование не указано в ответе)"
    lines = [
        f"ИИН/БИН: {data.get('xin')}",
        f"Наименование: {name}",
        f"Дата актуальности данных: {data.get('actuality')}",
        f"Дата регистрации: {data.get('regDate')}",
        f"Резидентство: {data.get('residency', {}).get('ru')}",
        f"ОКЭД: {data.get('oked', {}).get('ru')} — {data.get('okedName', {}).get('ru')}",
        f"Режим налогообложения: {data.get('taxMode', {}).get('ru')}",
        f"Статус НДС: {data.get('vatInfo', {}).get('ru')}"
        + (f" (с {data.get('vatDate')})" if data.get("vatDate") else ""),
        f"Налоговая задолженность: {data.get('taxDebt', 0):,.2f} тенге".replace(",", " "),
    ]

    flags = detect_flags(data)
    if flags:
        lines.append("")
        lines.append("⚠️  ОБНАРУЖЕНЫ ПОТЕНЦИАЛЬНЫЕ КРАСНЫЕ ФЛАГИ:")
        lines.extend(f"  - {f}" for f in flags)
    else:
        lines.append("")
        lines.append("✅ По проверенным реестрам записей не найдено (не равно подтверждению благонадёжности).")

    return "\n".join(lines)


def main(argv: list[str] | None = None) -> int:
    _load_dotenv(Path(__file__).with_name(".env"))

    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("xin", nargs="+", help="ИИН/БИН контрагента (12 цифр), можно несколько через пробел")
    parser.add_argument("--host", default=None, help="Переопределить KGD_PORTAL_HOST")
    parser.add_argument("--token", default=None, help="Переопределить KGD_PORTAL_TOKEN")
    parser.add_argument("--json", action="store_true", help="Вывести сырой JSON вместо читаемой сводки")
    args = parser.parse_args(argv)

    client = KgdClient(
        host=args.host or os.environ.get("KGD_PORTAL_HOST", DEFAULT_HOST),
        token=args.token or os.environ.get("KGD_PORTAL_TOKEN"),
    )

    exit_code = 0
    for i, xin in enumerate(args.xin):
        if i > 0:
            print("\n" + "=" * 60 + "\n")
        try:
            data = client.get_sur_data(xin)
        except (ValueError, RuntimeError) as exc:
            print(f"Ошибка: {exc}", file=sys.stderr)
            exit_code = 1
            continue
        except KgdApiError as exc:
            print(f"Ошибка API (код {exc.status_code}) для {xin}: {exc}", file=sys.stderr)
            exit_code = 1
            continue

        if args.json:
            print(json.dumps(data, ensure_ascii=False, indent=2))
        else:
            print(format_summary(data))

    return exit_code


if __name__ == "__main__":
    sys.exit(main())
