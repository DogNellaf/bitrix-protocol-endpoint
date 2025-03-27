import gspread
from oauth2client.service_account import ServiceAccountCredentials

scope = [
    'https://spreadsheets.google.com/feeds',
    'https://www.googleapis.com/auth/drive'
]
credentials = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
google_credentials = gspread.authorize(credentials)

SPREADSHEET_ID = '1hVgLxjvo4BKXuIXbdvzcrIRpBjjm8jW6lt5MOj6-3Go'

def parse_data(ids):
    '''
    Собирает данные из Google‑таблиц:
      - Результаты испытаний из вкладки 'показатели'
      - Методики из вкладки 'методики'
      - Оборудование из вкладки 'оборудование'
    Отбор строк производится по ID, переданным из Битрикс24.
    '''
    data = {
        'Нормативно-техническая документация': [],
        'Используемое оборудование': [],
        'Результаты испытаний': []
    }

    try:
        # Подключаемся к гугл-таблице
        spreadsheet = google_credentials.open_by_key(SPREADSHEET_ID)
        
        # Загружаем показатели
        results_worksheet = spreadsheet.worksheet('Показатели (таблица 1)')
        all_results = results_worksheet.get_all_records()
        result_list = []
        for id in ids:
            for record in all_results:
                method_ids = str(record['ID']).strip()
                if id in method_ids and record not in result_list:
                    result_list.append(record)

        data['Результаты испытаний'] = result_list
        
        # Загружаем методики
        methods_worksheet = spreadsheet.worksheet('Методики')
        all_methods = methods_worksheet.get_all_records()
        methods_list = []
        for result in data['Результаты испытаний']:
            method_test = result['Методы испытаний '].strip()  # удаляем лишние пробелы
            for method in all_methods:
                if method['Название методики'].strip() == method_test:
                    full_name = method['Полное наименование методики']
                    if full_name not in methods_list:
                        methods_list.append(full_name)

        data['Нормативно-техническая документация'] = methods_list
        
        # Загружаем оборудование
        equipments_worksheet = spreadsheet.worksheet('Оборудование')
        all_equipments = equipments_worksheet.get_all_records()
        equipments_list = []
        for equipment in all_equipments:
            equipment_method = str(equipment['Методика']).strip()
            for result in data['Результаты испытаний']:
                if equipment_method == result['Методы испытаний '].strip():
                    if equipment not in equipments_list:
                        equipments_list.append(equipment)

        data['Используемое оборудование'] = equipments_list
        
    except Exception as e:
        print('Ошибка при сборе данных из Google таблиц:', e)

    return data
