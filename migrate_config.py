"""
Утилита для миграции конфигурации из старого формата в новый.
"""
import json
from pathlib import Path
from src.models import AppConfig, CompetitorConfig, ColumnMapping, RowOffsets, TargetMapping


def migrate_old_config(old_data_path: Path, old_cities_path: Path, output_path: Path):
    """
    Мигрировать конфигурацию из старого формата.

    Args:
        old_data_path: путь к старому data.json
        old_cities_path: путь к старому city.json
        output_path: путь для сохранения новой конфигурации
    """
    # Загрузить старые данные
    with open(old_data_path, 'r', encoding='utf-8') as f:
        old_data = json.load(f)

    with open(old_cities_path, 'r', encoding='utf-8') as f:
        cities = json.load(f)

    # Создать новую конфигурацию
    config = AppConfig()
    config.cities = cities

    # Получить настройки шаблона
    if ' ' in old_data:
        config.template_file = old_data[' '].get('path_pattern', '')
        config.output_file = old_data[' '].get('path_pattern', '')

    # Конвертировать каждого конкурента
    for name, competitor_data in old_data.items():
        if name == ' ':
            continue

        tk_data = competitor_data.get('tk', {})
        shb_data = competitor_data.get('shb', {})

        competitor = CompetitorConfig(
            name=name,
            file_path=competitor_data.get('path', ''),
            enabled=tk_data.get('check', True),
            source_columns=ColumnMapping(
                city=tk_data.get('city', 'A'),
                convert=tk_data.get('convert', 'D'),
                minimum_1=tk_data.get('minimum_1', 'E'),
                minimum_2=tk_data.get('minimum_2', 'F'),
                volume=tk_data.get('objem', 'O'),
                weight_100=tk_data.get('ves_100', 'P'),
                weight_3000=tk_data.get('ves_3000', 'Q'),
            ),
            row_offsets=RowOffsets(
                row_app=tk_data.get('row_app', 0),
                row_1=tk_data.get('row_1', 0),
                row_2=tk_data.get('row_2', 0),
                row_3=tk_data.get('row_3', 0),
                row_4=tk_data.get('row_4', 0),
                row_5=tk_data.get('row_5', 0),
                row_6=tk_data.get('row_6', 0),
                row_7=tk_data.get('row_7', 0),
            ),
            target_columns=TargetMapping(
                convert=shb_data.get('convert_as', 'C'),
                minimum_1=shb_data.get('minimum_1_as', 'D'),
                minimum_2=shb_data.get('minimum_2_as', 'E'),
                volume=shb_data.get('objem_as', 'F'),
                weight_100=shb_data.get('ves_100_as', 'G'),
                weight_3000=shb_data.get('ves_3000_as', 'H'),
            )
        )

        config.competitors[name] = competitor

    # Сохранить новую конфигурацию
    config.save(output_path)
    print(f"✅ Конфигурация мигрирована: {output_path}")
    print(f"   Конкурентов: {len(config.competitors)}")
    print(f"   Городов: {len(config.cities)}")


if __name__ == '__main__':
    # Пути к старым файлам
    old_dir = Path('analiz_tk__old/src/json')
    old_data = old_dir / 'data.json'
    old_cities = old_dir / 'city.json'

    # Проверить наличие файлов
    if not old_data.exists():
        print(f"❌ Файл не найден: {old_data}")
        exit(1)

    if not old_cities.exists():
        print(f"❌ Файл не найден: {old_cities}")
        exit(1)

    # Мигрировать
    output = Path('config.json')
    migrate_old_config(old_data, old_cities, output)

