"""
Модели данных для конфигурации анализа.
"""
from dataclasses import dataclass, field, asdict
from typing import Dict, Any, List
import json
from pathlib import Path


@dataclass
class ColumnMapping:
    """Маппинг колонок из исходного файла в целевой."""
    city: str = "A"  # Колонка с городом
    convert: str = "D"  # Конверт
    minimum_1: str = "E"  # Минималка 1
    minimum_2: str = "F"  # Минималка 2
    volume: str = "O"  # Объем
    weight_100: str = "P"  # Вес до 100 кг
    weight_3000: str = "Q"  # Вес от 3000 кг


@dataclass
class RowOffsets:
    """Смещения строк для чтения данных."""
    row_app: int = 0  # Общее смещение
    row_1: int = 0
    row_2: int = 0  # Смещение для конверта
    row_3: int = 0  # Смещение для минималки 1
    row_4: int = 0  # Смещение для минималки 2
    row_5: int = 0  # Смещение для объема
    row_6: int = 0  # Смещение для веса 100
    row_7: int = 0  # Смещение для веса 3000


@dataclass
class Markups:
    """Наценки на цены конкурента."""
    convert: float = 0.0  # Наценка на конверт (%)
    minimum_1: float = 0.0  # Наценка на минималку 1 (%)
    minimum_2: float = 0.0  # Наценка на минималку 2 (%)
    volume: float = 0.0  # Наценка на объем (%)
    weight_100: float = 0.0  # Наценка на вес 100 (%)
    weight_3000: float = 0.0  # Наценка на вес 3000 (%)


@dataclass
class OwnCompany:
    """Собственная компания — строка после среднего значения."""
    name: str = "Новая Витэка"
    enabled: bool = True
    markups: Markups = field(default_factory=Markups)

    def to_dict(self) -> Dict[str, Any]:
        return {
            'name': self.name,
            'enabled': self.enabled,
            'markups': asdict(self.markups),
        }

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'OwnCompany':
        return cls(
            name=data.get('name', 'Новая Витэка'),
            enabled=data.get('enabled', True),
            markups=Markups(**data.get('markups', {})),
        )


@dataclass
class MarkupRow:
    """Строка с наценкой N% для ТК — вставляется после строки конкурента перед Средним."""
    name: str = ""          # Название строки, например «+10%»
    percent: float = 0.0    # Наценка в процентах

    def to_dict(self) -> Dict[str, Any]:
        return {'name': self.name, 'percent': self.percent}

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'MarkupRow':
        return cls(name=data.get('name', ''), percent=data.get('percent', 0.0))


@dataclass
class CompetitorConfig:
    """Конфигурация конкурента."""
    name: str
    file_path: str = ""
    enabled: bool = True
    bold: bool = False  # Выделять строку этого конкурента жирным
    source_columns: ColumnMapping = field(default_factory=ColumnMapping)
    target_columns: ColumnMapping = field(default_factory=ColumnMapping)
    row_offsets: RowOffsets = field(default_factory=RowOffsets)
    fuzzy_match_threshold: int = 95
    markups: Markups = field(default_factory=Markups)
    special_conditions: Dict[str, Any] = field(default_factory=dict)
    markup_rows: List[MarkupRow] = field(default_factory=list)  # Дополнительные строки с наценками

    def to_dict(self) -> Dict[str, Any]:
        """Преобразовать в словарь."""
        return {
            'name': self.name,
            'file_path': self.file_path,
            'enabled': self.enabled,
            'bold': self.bold,
            'source_columns': asdict(self.source_columns),
            'target_columns': asdict(self.target_columns),
            'row_offsets': asdict(self.row_offsets),
            'markups': asdict(self.markups),
            'fuzzy_match_threshold': self.fuzzy_match_threshold,
            'special_conditions': self.special_conditions,
            'markup_rows': [r.to_dict() for r in self.markup_rows],
        }

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'CompetitorConfig':
        """Создать из словаря."""
        return cls(
            name=data['name'],
            file_path=data.get('file_path', ''),
            enabled=data.get('enabled', True),
            bold=data.get('bold', False),
            source_columns=ColumnMapping(**data.get('source_columns', {})),
            target_columns=ColumnMapping(**data.get('target_columns', {})),
            row_offsets=RowOffsets(**data.get('row_offsets', {})),
            markups=Markups(**data.get('markups', {})),
            fuzzy_match_threshold=data.get('fuzzy_match_threshold', 95),
            special_conditions=data.get('special_conditions', {}),
            markup_rows=[MarkupRow.from_dict(r) for r in data.get('markup_rows', [])],
        )


@dataclass
class OutputConfig:
    """Конфигурация выходного файла."""
    title: str = "Стоимость доставки"
    subtitle: str = ""
    start_row: int = 3
    city_column: str = "A"
    include_average: bool = True
    average_row_offset: int = 1
    markups_sheet: bool = True

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'OutputConfig':
        return cls(**{k: v for k, v in data.items() if k in asdict(cls()).keys()})


@dataclass
class AppConfig:
    """Общая конфигурация приложения."""
    output_file: str = ""
    template_file: str = ""
    competitors: Dict[str, CompetitorConfig] = field(default_factory=dict)
    cities: Dict[str, int] = field(default_factory=dict)          # Город: номер строки в выходе
    city_aliases: Dict[str, List[str]] = field(default_factory=dict)  # Город: [псевдонимы]
    output_config: OutputConfig = field(default_factory=OutputConfig)
    own_company: OwnCompany = field(default_factory=OwnCompany)

    def get_city_names(self, city: str) -> List[str]:
        """Вернуть все варианты написания города (основное + псевдонимы)."""
        return [city] + self.city_aliases.get(city, [])

    def save(self, file_path: Path):
        """Сохранить конфигурацию в файл."""
        data = {
            'output_file': self.output_file,
            'template_file': self.template_file,
            'competitors': {name: comp.to_dict() for name, comp in self.competitors.items()},
            'cities': self.cities,
            'city_aliases': self.city_aliases,
            'output_config': self.output_config.to_dict(),
            'own_company': self.own_company.to_dict(),
        }
        with open(file_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

    @classmethod
    def load(cls, file_path: Path) -> 'AppConfig':
        """Загрузить конфигурацию из файла."""
        if not file_path.exists():
            return cls()

        with open(file_path, 'r', encoding='utf-8') as f:
            data = json.load(f)

        return cls(
            output_file=data.get('output_file', ''),
            template_file=data.get('template_file', ''),
            competitors={
                name: CompetitorConfig.from_dict(comp_data)
                for name, comp_data in data.get('competitors', {}).items()
            },
            cities=data.get('cities', {}),
            city_aliases=data.get('city_aliases', {}),
            output_config=OutputConfig.from_dict(data.get('output_config', {})),
            own_company=OwnCompany.from_dict(data['own_company']) if 'own_company' in data else OwnCompany(),
        )

