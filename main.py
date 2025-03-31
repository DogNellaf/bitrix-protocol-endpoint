import requests

from flask import Flask, request, jsonify

from utils import convert_word_to_pdf, convert_payload_to_map
from word_worker import fill_template, save_word_document
from google_parser import parse_data
from consts import TEMPLATE_FILENAME, RESULT_FILENAME, BITRIX_DOCUMENT_ENDPOINT

# ------------------------
# Настройка Flask-сервера
# ------------------------
app = Flask(__name__)
# ------------------------
# Вспомогательные функции
# ------------------------

def send_equipments_to_bitrix(equipments):
    '''
    Обновляет поле 'Оборудование' в Битрикс24.
    Формирует строку, где выбранное оборудование разделено точкой с запятой.
    Здесь реализована заглушка – фактический API‑вызов к Битрикс24 необходимо добавить.
    '''
    equipment_str = ';'.join(equipments)
    print('Обновление поля "Оборудование" в Битрикс24 значением:', equipment_str)
    # TODO: Реализовать вызов API Битрикс24 для обновления поля
    return True

def send_documents_to_bitrix(word_filename, pdf_filename):
    """
    Отправляет созданные документы на указанный эндпоинт Битрикс.
    """
    # Открываем файлы в бинарном режиме
    with open(word_filename, 'rb') as word_file, open(pdf_filename, 'rb') as pdf_file:
        files = {
            'doc_word': (word_filename, word_file, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'),
            'doc_pdf': (pdf_filename, pdf_file, 'application/pdf')
        }
        # Дополнительные данные можно передать через параметр data (например, ID сделки)
        # data = {
        #     # 'deal_id': '12345',  # Пример передачи идентификатора сделки
        # }
        response = requests.post(BITRIX_DOCUMENT_ENDPOINT, files=files)
        if response.status_code != 200:
            raise Exception("Ошибка при отправке документов в Битрикс: " + response.text)
        return response.json()

@app.route('/protocol/create', methods=['POST'])
def webhook():
    """
    Основной webhook данного проекта для создания протокола
    """
    try:
        # Получение данных от Битрикс24
        payload = request.json

        # конвертируем полученный payload в карту сопоставления шаблону
        bitrix_data = convert_payload_to_map(payload)

        # получаем id показателей
        ids = [
            id.strip()
            for id
            in payload.get('Идентификаторы (ID) показателей', '').split(',')
            if id.strip()
        ]

        # сбор данных из Google‑таблиц по заданным ID
        google_data = parse_data(ids)

        # заполнение шаблона Word (используется файл-шаблон, например, template.docx)
        doc = fill_template(TEMPLATE_FILENAME, google_data, bitrix_data)

        # Сохранение заполненного Word-документа
        word_filename = RESULT_FILENAME + '.docx'
        save_word_document(doc, word_filename)

        # Конвертация Word в PDF (не машиночитаемый, через PNG)
        pdf_filename = RESULT_FILENAME + '.pdf'
        convert_word_to_pdf(word_filename, pdf_filename)

        # TODO: обновлять поля "Оборудование" в битриксе
        send_equipments_to_bitrix(google_data['Используемое оборудование'])

        # Посылаем документы в битрикс
        send_documents_to_bitrix(word_filename, pdf_filename)

        # Возврат результата (в реальной интеграции необходимо загрузить файлы в поля сделки Битрикс24)
        return jsonify(
            {
                'status': 'success',
                'message': 'Документы успешно отправлены в Birtix',
                'word_file': word_filename,
                'pdf_file': pdf_filename
            }
        ), 200

    except Exception as e:
        print('Ошибка при обработке вебхука:', e)
        return jsonify({'status': 'error', 'message': str(e)}), 500

if __name__ == '__main__':
    app.run(port=5000)