from enum import Enum, auto
from typing import Final, Tuple

# Временные форматы
TIME_FORMAT: Final[str] = "HH:mm"
DATE_FORMAT: Final[str] = "dd.MM.yyyy"

# Диапазоны значений
TEMPERATURE_RANGE: Final[Tuple[int, int]] = (500, 2000)
MAX_PARTICIPANTS: Final[int] = 4

# Участники процесса
PARTICIPANTS: Final[list[str]] = [
    "Белков", "Карасев", "Ермаков", "Рабинович",
    "Валиулин", "Волков", "Семенов", "Левин",
    "Исмаилов", "Беляев", "Политов", "Кокшин",
    "Терентьев", "отсутствует"
]

# Наименования отливок
CASTING_NAMES: Final[list[str]] = [
    "Вороток", "Ригель", "Ригель optima", "Блок-картер", 
    "Колесо РИТМ", "Накладка резьб", "Блок цилиндров", 
    "Диагональ optima", "Кольцо"
]

class ExperimentType(Enum):
    """Типы экспериментов"""
    PAPER = "Бумага"
    FIBER = "Волокно"

class SectorName(Enum):
    """Названия секторов"""
    A = auto()
    B = auto()
    C = auto()
    D = auto()

    @classmethod
    def list(cls) -> list[str]:
        """Возвращает список имен секторов"""
        return [member.name for member in cls]
