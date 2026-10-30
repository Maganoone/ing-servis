"""Мост «1С ↔ Claude»: MCP-сервер, который даёт Claude доступ к базе 1С:Предприятие 8.3.

Подключается к базе через COM-коннектор платформы (V83.COMConnector) и умеет только
читать: смотреть структуру базы и выполнять запросы на языке запросов 1С.
Язык запросов 1С не умеет изменять данные, поэтому испортить базу через мост нельзя.

Настройка через переменные окружения (их прописывает setup.ps1 в конфиг Claude Desktop):
    ONEC_BASE       путь к каталогу файловой базы
    ONEC_USER       пользователь 1С
    ONEC_PASSWORD   пароль (если есть)
    ONEC_CONNECTION полная строка соединения — если задана, предыдущие три не нужны

Проверка без Claude:  python onec_mcp.py --test
"""

import asyncio
import decimal
import json
import logging
import os
import re
import sys
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
from typing import Any, Literal

from mcp.server.fastmcp import FastMCP

logging.basicConfig(stream=sys.stderr, level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger("onec_mcp")

MAX_ROWS = 5000

INSTRUCTIONS = """\
Ты подключён к рабочей базе 1С компании ИНГ-СЕРВИС (оборудование для видеонаблюдения,
СКУД, домофоны). Конфигурация: «Агент Плюс: Управление торговлей, редакция 10.3»
(доработанная УТ 10.3), платформа 8.3.13. Доступ только на чтение.

Как работать:
1. Сначала узнай структуру: describe_object для нужного регистра/документа. Конфигурация
   доработана, поэтому не угадывай имена реквизитов — проверяй.
2. Потом run_query с текстом на языке запросов 1С (русский синтаксис). Параметры — через &Имя.
3. Для ссылочных полей бери ПРЕДСТАВЛЕНИЕ(Поле) или Поле.Наименование, иначе получишь
   только представление по умолчанию. Ставь ПЕРВЫЕ N и сортировку, не тяни лишнее.

Типовые объекты УТ 10.3 (проверяй через describe_object):
- Справочник.Номенклатура (Код, Наименование, Артикул, Родитель, ЭтоГруппа),
  Справочник.Контрагенты, Справочник.ДоговорыКонтрагентов (Владелец = контрагент),
  Справочник.Склады, Справочник.ТипыЦенНоменклатуры
- РегистрНакопления.ТоварыНаСкладах — остатки: .Остатки(&НаДату) → КоличествоОстаток
- РегистрНакопления.ТоварыВРезервеНаСкладах — резервы
- РегистрНакопления.Продажи — обороты продаж: .Обороты(&Начало, &Конец) →
  КоличествоОборот, СтоимостьОборот (измерения Номенклатура, Контрагент, ДокументПродажи…)
- РегистрНакопления.ВзаиморасчетыСКонтрагентами — долги: .Остатки() →
  СуммаВзаиморасчетовОстаток / СуммаУпрОстаток, контрагент = ДоговорКонтрагента.Владелец
  (знак: положительный остаток — обычно долг контрагента нам; уточняй по документам)
- РегистрСведений.ЦеныНоменклатуры.СрезПоследних(, ТипЦен = &ТипЦен) — цены
- Документы: РеализацияТоваровУслуг (ТЧ Товары), ПоступлениеТоваровУслуг, ЗаказПокупателя,
  ЗаказПоставщику, ПриходныйКассовыйОрдер, ПлатежноеПоручениеВходящее
Для периода «по конец дня» используй КОНЕЦПЕРИОДА(&Конец, ДЕНЬ) в запросе.
Отвечай по-русски, суммы — с разделителями тысяч и ₽.
"""


# ---------------------------------------------------------------- соединение с 1С

def _quote(value: str) -> str:
    return '"' + value.replace('"', '""') + '"'


def connection_string() -> str:
    cs = os.environ.get("ONEC_CONNECTION")
    if cs:
        return cs
    base = os.environ.get("ONEC_BASE")
    if not base:
        raise RuntimeError("Не задан путь к базе 1С (переменная ONEC_BASE). Запустите setup.ps1.")
    parts = ["File=" + _quote(base)]
    if os.environ.get("ONEC_USER"):
        parts.append("Usr=" + _quote(os.environ["ONEC_USER"]))
    if os.environ.get("ONEC_PASSWORD"):
        parts.append("Pwd=" + _quote(os.environ["ONEC_PASSWORD"]))
    return ";".join(parts) + ";"


def com_connect():
    import win32com.client

    connector = win32com.client.Dispatch("V83.COMConnector")
    log.info("Подключаюсь к 1С…")
    conn = connector.Connect(connection_string())
    log.info("Подключено: %s %s", conn.Metadata.Name, conn.Metadata.Version)
    return conn


def describe_error(exc: Exception) -> str:
    """Достаёт из ошибки COM текст, который написала 1С (например, ошибку в запросе)."""
    args = getattr(exc, "args", ())
    if len(args) >= 3 and isinstance(args[2], tuple) and len(args[2]) >= 3 and args[2][2]:
        return str(args[2][2]).strip()
    return str(exc)


class OneC:
    """Все обращения к COM идут в одном отдельном потоке: COM-объекты 1С привязаны к потоку."""

    def __init__(self, connect=com_connect):
        self._connect = connect
        self._conn = None
        self._pool = ThreadPoolExecutor(max_workers=1, initializer=self._init_thread)

    @staticmethod
    def _init_thread():
        try:
            import pythoncom

            pythoncom.CoInitialize()
        except ImportError:
            pass

    def _alive(self) -> bool:
        try:
            self._conn.Metadata.Name
            return True
        except Exception:
            return False

    def _run(self, fn, *args):
        if self._conn is None:
            self._conn = self._connect()
        try:
            return fn(self._conn, *args)
        except Exception:
            if self._alive():
                raise
            log.warning("Соединение с 1С потеряно, переподключаюсь")
            self._conn = self._connect()
            return fn(self._conn, *args)

    async def call(self, fn, *args) -> str:
        loop = asyncio.get_running_loop()
        try:
            return await loop.run_in_executor(self._pool, self._run, fn, *args)
        except Exception as exc:
            message = describe_error(exc)
            log.warning("Ошибка 1С: %s", message)
            return "Ошибка 1С: " + message


# ---------------------------------------------------------------- значения 1С ↔ Python

DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}([ T]\d{2}:\d{2}(:\d{2})?)?$")


def to_python(conn, value):
    if value is None or isinstance(value, (bool, int, str)):
        return value
    if isinstance(value, decimal.Decimal):
        value = float(value)
    if isinstance(value, float):
        return int(value) if value.is_integer() else round(value, 6)
    if isinstance(value, datetime):
        if value.year < 1800:  # пустая дата 1С
            return None
        if value.hour or value.minute or value.second:
            return value.strftime("%Y-%m-%d %H:%M:%S")
        return value.strftime("%Y-%m-%d")
    try:
        return conn.String(value)  # ссылки, перечисления и прочее — их представление
    except Exception:
        return str(value)


def to_onec(conn, value):
    if isinstance(value, str) and DATE_RE.match(value):
        dt = datetime.fromisoformat(value.replace(" ", "T"))
        try:
            import pywintypes

            return pywintypes.Time(dt)
        except ImportError:
            return dt
    if isinstance(value, list):
        arr = conn.NewObject("Array")
        for item in value:
            arr.Add(to_onec(conn, item))
        return arr
    return value


# ---------------------------------------------------------------- операции

KINDS = {
    "Справочники": "Catalogs",
    "Документы": "Documents",
    "РегистрыНакопления": "AccumulationRegisters",
    "РегистрыСведений": "InformationRegisters",
    "Перечисления": "Enums",
    "Константы": "Constants",
    "Отчеты": "Reports",
    "Обработки": "DataProcessors",
}


def op_info(conn) -> str:
    md = conn.Metadata
    q = conn.NewObject("Query")
    q.Text = "ВЫБРАТЬ ТЕКУЩАЯДАТА() КАК Сейчас"
    sel = q.Execute().Select()
    sel.Next()
    return json.dumps(
        {
            "конфигурация": md.Synonym or md.Name,
            "версия": md.Version,
            "дата_в_базе": to_python(conn, sel.Get(0)),
        },
        ensure_ascii=False,
    )


def op_list(conn, kind: str, contains: str) -> str:
    coll = getattr(conn.Metadata, KINDS[kind])
    needle = contains.lower()
    lines = []
    for i in range(coll.Count()):
        m = coll.Get(i)
        name, syn = m.FullName(), m.Synonym
        if needle and needle not in name.lower() and needle not in syn.lower():
            continue
        lines.append(f"{name} — {syn}" if syn else name)
    return "\n".join(lines) or "Ничего не найдено."


def _type_names(conn, type_description) -> str:
    names = []
    try:
        types = type_description.Types()
        for i in range(types.Count()):
            t = types.Get(i)
            m = conn.Metadata.FindByType(t)
            names.append(m.FullName() if m is not None else conn.String(t))
    except Exception:
        return ""
    return ", ".join(names)


def _fields(conn, coll, indent="  ") -> list[str]:
    lines = []
    for i in range(coll.Count()):
        f = coll.Get(i)
        line = f"{indent}{f.Name}"
        types = _type_names(conn, f.Type)
        if types:
            line += f": {types}"
        if f.Synonym and f.Synonym != f.Name:
            line += f"  ({f.Synonym})"
        lines.append(line)
    return lines


def op_describe(conn, full_name: str) -> str:
    m = conn.Metadata.FindByFullName(full_name)
    if m is None:
        return f"Объект «{full_name}» не найден. Посмотрите имена через list_metadata."
    out = [f"{m.FullName()} — {m.Synonym}"]
    for prop in ("RegisterType", "InformationRegisterPeriodicity"):
        try:
            out.append(f"{prop}: {conn.String(getattr(m, prop))}")
        except Exception:
            pass
    for title, prop in (("Измерения", "Dimensions"), ("Ресурсы", "Resources"), ("Реквизиты", "Attributes")):
        try:
            coll = getattr(m, prop)
        except Exception:
            continue
        if coll.Count():
            out.append(f"{title}:")
            out += _fields(conn, coll)
    try:
        sections = m.TabularSections
        for i in range(sections.Count()):
            ts = sections.Get(i)
            out.append(f"Табличная часть {ts.Name} ({ts.Synonym}):")
            out += _fields(conn, ts.Attributes, "    ")
    except Exception:
        pass
    try:
        values = m.EnumValues
        out.append("Значения: " + ", ".join(values.Get(i).Name for i in range(values.Count())))
    except Exception:
        pass
    return "\n".join(out)


def op_query(conn, text: str, params: dict[str, Any], limit: int) -> str:
    q = conn.NewObject("Query")
    q.Text = text
    for name, value in params.items():
        q.SetParameter(name, to_onec(conn, value))
    result = q.Execute()
    columns = [result.Columns.Get(i).Name for i in range(result.Columns.Count())]
    sel = result.Select()
    total = sel.Count()
    rows = []
    while len(rows) < limit and sel.Next():
        rows.append([to_python(conn, sel.Get(i)) for i in range(len(columns))])
    answer = {"колонки": columns, "строки": rows, "всего_строк": total}
    if total > len(rows):
        answer["внимание"] = f"Показаны первые {len(rows)} из {total}. Уточните запрос или увеличьте limit."
    return json.dumps(answer, ensure_ascii=False)


# ---------------------------------------------------------------- MCP-инструменты

mcp = FastMCP("1c", instructions=INSTRUCTIONS)
onec = OneC()


@mcp.tool()
async def info() -> str:
    """Проверить подключение к базе 1С: название конфигурации, версия и текущая дата в базе."""
    return await onec.call(op_info)


@mcp.tool()
async def list_metadata(
    kind: Literal[
        "Справочники", "Документы", "РегистрыНакопления", "РегистрыСведений",
        "Перечисления", "Константы", "Отчеты", "Обработки",
    ],
    contains: str = "",
) -> str:
    """Список объектов базы 1С указанного вида с синонимами.

    contains — необязательный фильтр по части имени или синонима (без учёта регистра).
    """
    return await onec.call(op_list, kind, contains)


@mcp.tool()
async def describe_object(full_name: str) -> str:
    """Структура объекта 1С: измерения, ресурсы, реквизиты, табличные части и их типы.

    full_name — полное имя, например «РегистрНакопления.Продажи» или «Документ.РеализацияТоваровУслуг».
    """
    return await onec.call(op_describe, full_name)


@mcp.tool()
async def run_query(text: str, params: dict[str, Any] | None = None, limit: int = 200) -> str:
    """Выполнить запрос на языке запросов 1С (только чтение) и вернуть таблицу в JSON.

    text   — текст запроса, например «ВЫБРАТЬ ПЕРВЫЕ 10 Наименование ИЗ Справочник.Номенклатура».
             Можно пакет запросов через «;» — вернётся результат последнего.
    params — значения параметров &Имя. Даты строкой «2026-10-01» или «2026-10-31 23:59:59»,
             списки (для «В (&Список)») — массивом.
    limit  — сколько строк вернуть (по умолчанию 200, максимум 5000).
    """
    limit = max(1, min(limit, MAX_ROWS))
    return await onec.call(op_query, text, params or {}, limit)


# ---------------------------------------------------------------- проверка из консоли

def self_test():
    print("Строка соединения:", re.sub(r'Pwd="[^"]*"', 'Pwd="***"', connection_string()))
    conn = com_connect()
    print("Подключение есть:", op_info(conn))
    print(op_query(conn, "ВЫБРАТЬ ПЕРВЫЕ 5 Код, Наименование ИЗ Справочник.Номенклатура ГДЕ НЕ ЭтоГруппа", {}, 5))
    print("\nВсё работает. Можно подключать к Claude Desktop.")


if __name__ == "__main__":
    if "--test" in sys.argv:
        self_test()
    else:
        mcp.run()
