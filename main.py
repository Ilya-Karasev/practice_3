import pandas as pd
import re

# Возьмем данные из входного файла
df_start_adress = pd.read_excel('Сырые адреса.xlsx')

def formating_text(text):
    if not isinstance(text, str):
        return ''

    # Удалим записи, которые содержат только числовые значения
    if re.match(r'^\d+$', text.strip()):
        return ''

    # Найдем и удалим слово "адрес" и всё, что идет после него
    match = re.search(r'\bадрес\b', text, re.IGNORECASE)
    if match:
        text = text[:match.start()].strip()

    # Определим ключевые слова, указывающие на деревню или номер дома
    village_patterns = [r'\bд\.\b']
    house_patterns = [r'д\.\s*\d+\w*', r'дом\s*\d+\w*', r'корпус\s*\d+\w*', r'корп\.\s*\d+\w*', r'к\.\s*\d+\w*', r'строение\s*\d+\w*', r'стр\.\s*\d+\w*', r'литер\s*\d+\w*', r'лит\.\s*\д+\w*', r'помещение\s*\д+\w*', r'офис\s*\д+\w*', r'квартира\s*\д+\w*', r'кв\.\с*\д+\w*', r'участок\s*\д+\w*', r'уч\.\с*\д+\w*', r'здание\s*\д+\w*', r'строение\s*\д+\w*', r'подъезд\s*\д+\w*', r'под\.\с*\д+\w*']

    # Создаем шаблон для поиска ключевых слов деревни и номеров домов
    village_pattern = re.compile('|'.join(village_patterns), re.IGNORECASE)
    house_pattern = re.compile('|'.join(house_patterns), re.IGNORECASE)

    # Найдем первое упоминание деревни или номера дома и обрежем текст
    village_match = village_pattern.search(text)
    house_match = house_pattern.search(text)

    if house_match:
        text = text[:house_match.end()].strip()
    elif village_match:
        text = text[:village_match.end()].strip()

    # Удалим повторяющиеся слова
    old_text = text.split()
    new_text = []
    for word in old_text:
        if word not in new_text:
            new_text.append(word)

    return ' '.join(new_text)

# Форматируем адреса и фильтруем те, что содержат только цифры или неподходящий текст
df_start_adress['formating_adress'] = df_start_adress['Адрес, местоположение объекта'].apply(lambda x: formating_text(x))
df_start_adress = df_start_adress[df_start_adress['formating_adress'] != '']
df_start_adress = df_start_adress[df_start_adress['formating_adress'].str.lower() != 'адрес, местоположение объекта']

# Приведем формированные адреса к нижнему регистру и удалим дубликаты
df_start_adress['formating_adress_lower'] = df_start_adress['formating_adress'].str.lower()
df_start_adress = df_start_adress.drop_duplicates(subset=['formating_adress_lower'])

# Удалим вспомогательный столбец
df_start_adress = df_start_adress.drop(columns=['formating_adress_lower'])

# Сохраним результат в новый файл
df_start_adress.to_excel('Исправленные адреса.xlsx', index=False)