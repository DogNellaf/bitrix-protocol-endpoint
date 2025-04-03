import logging
import requests
import json
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

def send_documents_to_bitrix(order_id, word_filename, pdf_filename):
    with open(word_filename, 'rb') as word_file, open(pdf_filename, 'rb') as pdf_file:
        data = {
            "id": order_id
        }

        files = {
            'fields[UFCRM_51742925985182]': (
                word_filename, 
                word_file,
                'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            ),
            'fields[UFCRM51742926009772]': (
                pdf_filename, 
                open(pdf_filename, 'rb'),
                'application/pdf'
            )
        }

        # Отправка POST-запроса
        return requests.post(BITRIX_DOCUMENT_ENDPOINT, data=data, files=files)

@app.route('/protocol/create', methods=['POST'])
def webhook():
    try:
        app.logger.info("Headers: %s", request.headers)
        app.logger.info("Form data: %s", request.form)
        app.logger.info("Raw data: %s", request.data.decode('utf8'))
        test = list(request.args)
        print(test)
        for key in request.args:
            payload = request.args.get(key)
            app.logger.info("Получен payload: " + payload)

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
        response = send_equipments_to_bitrix(google_data['Используемое оборудование'])
        # send_documents_to_bitrix(word_filename, pdf_filename)

        if response.status_code == 200:
            app.logger.info("Документы успешно отправлены в Bitrix", payload)
        else:
            try:
                result = response.json()
                print("Ответ сервера:", json.dumps(result, indent=4, ensure_ascii=False))
            except json.JSONDecodeError:
                print("Ошибка декодирования JSON в ответе:", response.text)

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
