import logging
import requests
from flask import Flask, request, jsonify

from utils import convert_word_to_pdf, convert_payload_to_map
from word_worker import fill_template, save_word_document
from google_parser import parse_data
from consts import TEMPLATE_FILENAME, RESULT_FILENAME, BITRIX_DOCUMENT_ENDPOINT

# Настройка Flask-сервера
app = Flask(__name__)

# Настройка базового логирования
logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s: %(message)s')

def send_equipments_to_bitrix(equipments):
    equipment_str = ';'.join(equipments)
    app.logger.info('Обновление поля "Оборудование" в Битрикс24 значением: %s', equipment_str)
    return True

def send_documents_to_bitrix(word_filename, pdf_filename):
    with open(word_filename, 'rb') as word_file, open(pdf_filename, 'rb') as pdf_file:
        files = {
            'doc_word': (word_filename, word_file, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'),
            'doc_pdf': (pdf_filename, pdf_file, 'application/pdf')
        }
        response = requests.post(BITRIX_DOCUMENT_ENDPOINT, files=files)
        if response.status_code != 200:
            raise Exception("Ошибка при отправке документов в Битрикс: " + response.text)
        return response.json()

@app.route('/protocol/create', methods=['POST'])
def webhook():
    try:
        payload = request.json
        # Вместо print используем логгер
        app.logger.info("Получен payload: %s", payload)

        bitrix_data = convert_payload_to_map(payload)
        ids = [
            id.strip()
            for id in payload.get('Идентификаторы (ID) показателей', '').split(',')
            if id.strip()
        ]
        google_data = parse_data(ids)
        doc = fill_template(TEMPLATE_FILENAME, google_data, bitrix_data)
        word_filename = RESULT_FILENAME + '.docx'
        save_word_document(doc, word_filename)
        pdf_filename = RESULT_FILENAME + '.pdf'
        convert_word_to_pdf(word_filename, pdf_filename)
        send_equipments_to_bitrix(google_data['Используемое оборудование'])
        # send_documents_to_bitrix(word_filename, pdf_filename)

        app.logger.info("Документы успешно отправлены в Bitrix", payload)

        return jsonify({
            'status': 'success',
            'message': 'Документы успешно отправлены в Bitrix',
            'word_file': word_filename,
            'pdf_file': pdf_filename
        }), 200

    except Exception as e:
        app.logger.error("Ошибка при обработке вебхука: %s", e)
        return jsonify({'status': 'error', 'message': str(e)}), 500

if __name__ == '__main__':
    app.run(port=5000)
