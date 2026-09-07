---
name: gp-mailing
description: "Актуальный список адресатов ДКТ-3 (25 адресов, с учётом проектов 2026 года, asylbexul@gmail.com исключён подтверждённо)"
---

# Рассылка грантополучателям ДКТ-3

Скилл для быстрого создания email-черновика на всех грантополучателей менеджера ДКТ-3.

## Алгоритм

### Шаг 1 — Получи текст сообщения
Если пользователь ещё не дал текст — спроси: тему письма и тело сообщения.
Если текст уже есть в разговоре — используй его.

**Стандартный нижний колонтитул** (добавлять всегда, если пользователь не просит иначе):
```
АО «Фонд науки», ДКТ 3.
```

Обращение к адресатам по умолчанию — «Уважаемые грантополучатели!» (не «коллеги»), если пользователь не просит иначе.

### Шаг 2 — Определи список адресатов
Используй базовый список из `references/recipients.md` (см. актуальную версию ниже — обновлён 31.08.2026).

Если пользователь говорит «не всем» или называет исключения — убери нужные адреса.
Если пользователь добавляет нового получателя — добавь в список.

### Шаг 3 — Создай черновик Gmail
Используй инструмент `gmail_create_draft` с:
- **to**: все адреса через запятую
- **subject**: тема от пользователя
- **body**: текст сообщения

### Шаг 4 — Подтверди
Сообщи пользователю:
- Ссылку на черновик
- Количество адресатов
- Напомни проверить, кто уже ответил (при необходимости убрать из списка)

---

## Актуальный список адресатов (25 адресов, обновлено 31.08.2026)

```
Dima-shum-92@mail.ru
tsenter-zerna@mail.ru
buh8282@mail.ru
vika_rose83@mail.ru
global-bee@mail.ru
info@alkaral.kz
likhatskaya.g@alkaral.kz
spk_tjj_tolesh@mail.ru
info@kaznu.edu.kz
salima.abdraimova@kaznu.edu.kz
alimhanov.meirzhan@gmail.com
a.daurenbekova@satbayev.university
zhamalbek@mail.ru
nmanabaev@mail.ru
asil83@list.ru
talgat.tan@mail.ru
P.oleg76@mail.ru
a.temirkhan@kbtu.kz
y.iskakov@satbayev.university
serdaliyev.yerdulla@gmail.com
tokmajeshvili@gmail.com
saduakassov@mail.ru
smanovruslan@mail.ru
babkenov64@mail.ru
nazerke_oraz@mail.ru
```

### Новые контакты по проектам 2026 года (добавлены 31.08.2026)
P.oleg76@mail.ru, a.temirkhan@kbtu.kz, y.iskakov@satbayev.university, serdaliyev.yerdulla@gmail.com, tokmajeshvili@gmail.com, saduakassov@mail.ru, smanovruslan@mail.ru, babkenov64@mail.ru, nazerke_oraz@mail.ru

### Исключены как неактуальные (подтверждено пользователем 31.08.2026)
Bakyt_kusy_kz@mail.com, yelzhas_90@mail.ru, gulmirakz_80@mail.ru, Tlevlessova@gmail.com, tooalala@gmail.com, asylbexul@gmail.com

## Соответствие ИРН → почта (актуализировать по мере необходимости)

| ИРН | Краткое название | Email(ы) |
|-----|-----------------|----------|
| DP25997723 | Корм из масличных | talgat.tan@mail.ru |
| DP25900609 | С/х техника | nmanabaev@mail.ru |
| DP23692015 | Переработка дыни | alimhanov.meirzhan@gmail.com, a.daurenbekova@satbayev.university |
| DP23692342 | Овощные соки | alimhanov.meirzhan@gmail.com, a.daurenbekova@satbayev.university |
| DP21681972 | Переработка фруктового сырья | spk_tjj_tolesh@mail.ru |
| DP21681724 | Пчелопитомник | vika_rose83@mail.ru |
| DP21681508 | Совместная технология | Dima-shum-92@mail.ru |
| DP23691813 | — | salima.abdraimova@kaznu.edu.kz |
| DP23691582 | — | адрес asylbexul@gmail.com исключён из рассылки (неактуален) |

---

## Обновление списка адресатов

Если пользователь говорит, что список устарел или кто-то добавился/убрался:
1. Обнови список прямо в этом файле и создай черновик с новым списком
2. При необходимости поищи в Gmail последнюю массовую рассылку (`from:snadinassylbek@gmail.com to:(грантополучатель OR ДКТ)`) для сверки

---

## Правила

- Подпись пользователя (`snadinassylbek@gmail.com`) добавляется автоматически Gmail — не дублируй в теле
- Стандартный футер «АО «Фонд науки», ДКТ 3.» — всегда, кроме случаев, когда пользователь явно просит иначе
- Обращение по умолчанию — «Уважаемые грантополучатели!»
- Не добавляй лишних вводных фраз — создавай черновик сразу после получения текста
- Если нужно убрать получателей, которые уже ответили — уточни у пользователя список ответивших