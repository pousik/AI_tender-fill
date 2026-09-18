from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Iterable

from .text import clean, norm, strip_marks


# Явный словарь вместо свободного fuzzy-match. Это важно для повторяющихся
# строк «Класс точности», «Номинальная нагрузка» и т.п.
FIELD_ALIASES: dict[str, tuple[str, ...]] = {
    "manufacturer": ("изготовитель", "производитель", "завод-изготовитель", "предприятие-изготовитель"),
    "brand": ("заводской тип", "марка", "тип изделия", "тип трансформатора"),
    "internal_insulation": ("вид внутренней изоляции",),
    "external_insulation": ("тип внешней изоляции", "внешняя изоляция"),
    "external_color": ("цвет внешней изоляции", "цвет изоляции"),
    "nominal_voltage": ("номинальное напряжение",),
    "max_working_voltage": ("наибольшее рабочее напряжение",),
    "frequency": ("номинальная частота",),
    "primary_current": ("номинальный ток первичной обмотки",),
    "max_primary_current": ("наибольший рабочий первичный ток",),
    "thermal_current": ("ток термической стойкости",),
    "thermal_time": ("время протекания тока термической стойкости",),
    "secondary_current": ("номинальный вторичный ток",),
    "secondary_count": ("количество вторичных обмоток",),
    "accuracy_class": ("класс точности",),
    "secondary_load": ("номинальная нагрузка", "вторичная нагрузка"),
    "instrument_security": ("коэффициент безопасности приборов",),
    "protection_limit_multiple": ("номинальная предельная кратность вторичных обмоток",),
    "verification_capability": ("возможности", "проведения поверки", "поверки/калибровки"),
    "mechanical_load": ("суммарная механическая нагрузка",),
    "tightness": ("герметичность конструкции",),
    "gas_pressure": ("номинальное давление",),
    "gas_alarm": ("сигнализатора давления", "сигнализатор давления"),
    "gas_gauge": ("манометра", "плотномера"),
    "gas_safety_valve": ("предохранительного клапана",),
    "gas_leak": ("расход элегаза на утечки",),
    "dimensions": ("габаритные размеры",),
    "mass": ("масса трансформатора", "масса трансформатора тока"),
    "gas_mass": ("масса масла", "масса масла (элегаза)", "масса элегаза"),
    "placement_category": ("категория размещения",),
    "climate": ("климатическое исполнение",),
    "upper_temperature": ("верхнее рабочее значение температуры",),
    "lower_temperature": ("нижнее рабочее значение температуры",),
    "wind_no_ice": ("скорость ветра при отсутствии гололеда", "скорость ветра, при отсутствии гололеда"),
    "wind_ice": ("скорость ветра при наличии гололеда", "скорость ветра, при наличии гололеда"),
    "ice_thickness": ("толщина стенки гололеда", "толщина стенки гололеда, мм"),
    "installation_altitude": ("высота установки над уровнем моря",),
    "seismicity": ("сейсмостойкость",),
    "impulse_voltage": ("испытательное напряжение полного грозового импульса",),
    "ac_test_voltage": ("одноминутное испытательное напряжение 50 гц",),
    "overvoltage": ("допустимые повышения напряжения",),
    "partial_discharge": ("уровень частичных разрядов",),
    "secondary_ac_test": ("изоляция вторичных обмоток должна выдерживать",),
    "interturn_test": ("межвитковая изоляция вторичных обмоток",),
    "service_life": ("срок службы",),
    "maintenance": ("периодичность и объем технического обслуживания",),
    "repair_free": ("отсутствие необходимости ремонта",),
    "service_cost": ("доля", "стоимость капитального ремонта", "объем необходимых затрат"),
    "explosion_safety": ("взрывобезопасность",),
    "verification_interval": ("интервал между поверками", "срок периодической поверки"),
    "warranty": ("гарантийный срок эксплуатации",),
    "radio_noise": ("уровень радиопомех",),
    "certificate": ("сертификатов безопасности", "номер и дата выдачи российских сертификатов"),
    "documentation": ("эксплуатационная документация",),
    "terminal_contacts": ("контактных клемм",),
    "support_structures": ("опорных металлоконструкций",),
    "service_kit": ("приспособлений для сервисного обслуживания",),
    "anticorrosion": ("антикоррозионное покрытие",),
    "type_approval": ("утверждении типа си", "свидетельства об утверждении типа"),
    "factory_passport": ("заводского паспорта", "формуляра",),
    "initial_verification": ("первичной поверкой",),
    "marking_packaging": ("маркировка, упаковка и консервация",),
    "transport": ("условия транспортирования",),
    "shock_indicator": ("шок-индикатора",),
    "delivery": ("доставка оборудования", "растамаживание и доставка"),
    "chief_engineer": ("шеф-инженера фирмы-изготовителя",),
    "storage": ("условия хранения", "срок хранения отдельно хранящихся деталей"),
    "packaged_storage": ("срок хранения в упаковке производителя",),
    "quality_docs": ("документа подтверждающих качество изделия", "подтверждающих качество изделия"),
    "meter_certificates": ("сертификаты об утверждении типа средств измерения",),
    "factory_tests": ("заводских приемо-сдаточных испытаниях",),
    "commissioning": ("шеф-монтажные и пуско-наладочные работы",),
    "service_center": ("сервисного центра", "ремонтной базы"),
    "training": ("обучения и периодическая аттестация персонала",),
    "certified_staff": ("аттестованных производителем специалистов",),
    "consultations": ("консультации и рекомендации по эксплуатации и ремонту",),
    "arrival_72h": ("в течение 72 часов",),
    "spare_parts_20y": ("запасных частей, ремонт и/или замена",),
    "spare_parts_6m": ("срок поставки запасных частей",),
}


@dataclass(frozen=True)
class RowSemantics:
    field_key: str
    data_type: str
    preference: str


def field_key(parameter: str, parent_context: str = "") -> str:
    p = norm(parameter)
    # Специфичные строки 6.5/6.6 проверяем раньше общей строки про AC-напряжение.
    if "межвитковая изоляция вторичных обмоток" in p:
        return "interturn_test"
    if "изоляция вторичных обмоток должна выдерживать" in p:
        return "secondary_ac_test"
    # Вложенная таблица вторичных обмоток: родитель учитывается отдельно,
    # чтобы одинаковый «Класс точности» не смешивался между обмотками.
    for key, aliases in FIELD_ALIASES.items():
        if key in {"secondary_ac_test", "interturn_test"}:
            continue
        if any(norm(alias) in p for alias in aliases):
            return key
    if p == "класс точности" and parent_context:
        return "accuracy_class"
    if "номинальная нагрузка" in p:
        return "secondary_load"
    return canonical_fallback(p)


def canonical_fallback(parameter: str) -> str:
    value = re.sub(r"[^\w\u0400-\u04ff]+", "_", parameter, flags=re.UNICODE).strip("_")
    return value[:120] or "unknown"


def semantics(key: str, requirement: str = "") -> RowSemantics:
    req = norm(requirement)
    if key in {"nominal_voltage", "max_working_voltage", "frequency", "primary_current", "max_primary_current",
               "thermal_current", "thermal_time", "secondary_current", "secondary_count", "secondary_load", "instrument_security",
               "protection_limit_multiple", "mechanical_load", "gas_pressure", "gas_leak", "dimensions", "mass", "gas_mass",
               "placement_category", "upper_temperature", "lower_temperature", "wind_no_ice", "wind_ice", "ice_thickness",
               "installation_altitude", "seismicity", "impulse_voltage", "ac_test_voltage", "partial_discharge",
               "secondary_ac_test", "interturn_test", "service_life", "service_cost", "verification_interval", "warranty", "radio_noise",
               "packaged_storage"}:
        data_type = "number_or_text"
    elif key in {"verification_capability", "tightness", "gas_alarm", "gas_gauge", "gas_safety_valve", "overvoltage", "repair_free",
                 "explosion_safety", "documentation", "terminal_contacts", "support_structures", "service_kit", "anticorrosion",
                 "type_approval", "factory_passport", "initial_verification", "marking_packaging", "shock_indicator", "delivery",
                 "factory_tests", "commissioning", "training", "arrival_72h", "spare_parts_20y", "spare_parts_6m", "quality_docs", "meter_certificates"}:
        data_type = "yes_no_text"
    elif key in {"manufacturer", "brand", "internal_insulation", "external_insulation", "external_color", "climate", "maintenance",
                 "certificate", "transport", "chief_engineer", "storage", "service_center", "certified_staff", "consultations"}:
        data_type = "text"
    else:
        data_type = "text"
    if "не менее" in req or "не более" in req:
        preference = "product_value_or_threshold"
    elif req == "*":
        preference = "product_value"
    else:
        preference = "requirement_or_product_value"
    return RowSemantics(key, data_type, preference)


def requirement_value(requirement: str) -> str:
    text = clean(requirement)
    if not text or re.fullmatch(r"\*+", text):
        return ""
    low = norm(text)
    if low in {"да", "нет"}:
        return text
    if low in {"обязательно", "обязателен", "обязательна", "обязательны"}:
        return "Да"
    if low.startswith("не менее") or low.startswith("не более"):
        m = re.search(r"[-+]?\d+(?:[,.]\d+)?", low)
        return m.group(0) if m else ""
    if "в соответствии с" in low or low.startswith("в соответствии") or low.startswith("согласно"):
        return ""
    if low.startswith("указать") or low in {"или", "не оговорено"}:
        return ""
    # Если требование содержит чистое конкретное значение, разрешаем его
    # как резерв, но не как подтвержденный заводской факт.
    if re.fullmatch(r"[\w\-/.%,+]+", text, flags=re.UNICODE):
        return strip_marks(text)
    return ""


def is_constraint(requirement: str) -> bool:
    low = norm(requirement)
    return any(token in low for token in ("не менее", "не более", "должен", "должна", "в соответствии", "согласно"))


def allowed_parent(parent_context: str, requirement: str) -> str:
    """Сохраняет родительскую обмотку как часть семантического ключа."""
    p = norm(parent_context)
    if "обмотка 1" in p:
        return "winding_1"
    if "обмотка 2" in p:
        return "winding_2"
    if re.search(r"обмотк[аи]\s*3\s*[-–]\s*6", p):
        return "winding_3_6"
    return ""
