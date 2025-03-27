import os
import random
import tempfile

from docx2pdf import convert as docx2pdf_convert
from pdf2image import convert_from_path
import img2pdf

def convert_word_to_pdf(word_file, output_pdf):
    """
    Конвертирует Word-документ в PDF, а затем делает PDF немашиночитаемым:
    1. Сначала генерируется временный PDF с помощью docx2pdf.
    2. Затем каждая страница PDF конвертируется в PNG (с помощью pdf2image).
    3. После чего из PNG-изображений формируется единый PDF (с помощью img2pdf).
    
    Обратите внимание, что для pdf2image требуется наличие poppler.
    """
    with tempfile.TemporaryDirectory() as tmpdirname:
        temp_pdf = os.path.join(tmpdirname, "temp.pdf")
        # Конвертируем Word в PDF
        docx2pdf_convert(word_file, temp_pdf)
        
        # Конвертируем PDF-страницы в изображения (PNG)
        images = convert_from_path(temp_pdf, fmt="png")
        png_files = []
        for idx, image in enumerate(images):
            png_path = os.path.join(tmpdirname, f"page_{idx}.png")
            image.save(png_path, "PNG")
            png_files.append(png_path)
        
        # Собираем единый PDF из PNG изображений
        with open(output_pdf, "wb") as file:
            file.write(
                img2pdf.convert(png_files)
            )

    return output_pdf

def parse_range(value):
    """
    Если в ячейке указано значение с диапазоном вида "0,0010- 0,0025",
    функция заменяет запятую на точку, затем выбирает случайное значение из диапазона.
    Если диапазон не распознан, возвращает исходное значение.
    """
    value = str(value).replace(" ", "").replace(",", ".")
    if '-' in value:
        parts = value.split('-')
        try:
            low = float(parts[0])
            high = float(parts[1])
            return round(random.uniform(low, high), 4)
        except Exception as ex:
            print("Ошибка при разборе диапазона:", ex)
            return value
    else:
        return value

def convert_payload_to_map(payload):
    """
    Создает из payload карту сопоставлений для заполнения шаблона документа
    """
    return {
        "{UfCrm51741286861298}": payload['Номер протокола испытаний'],
        "{UfCrm51741286877382}": payload['Дата испытаний'],
        "{UfCrm51741285071989}": payload['Наименование объекта испытаний'],
        "{UfCrm51741285600655}": payload['Изготовитель'],
        "{UfCrm51741799909007}": payload['Юридический адрес изготовителя'],
        "{UfCrm51741285642541}": payload['Фактический адрес места осуществления деятельности изготовителя'],
        "{UfCrm51741787685Title}": payload['Заказчик'],
        "{UfCrm51741787685UfCrm1741797658252}": payload['Юридический адрес заказчика'],
        "{UfCrm51741787685UfCrm1741797676159}": payload['Фактический адрес места осуществления деятельности заказчика'],
        "{UfCrm51741798185411}": payload['Дата получения образцов'],
        "{UfCrm51741798069674}": payload['Начало испытаний'],
        "{UfCrm51741798090701}": payload['Окончание испытаний'],
        "{UfCrm51741285472446}": payload['Сопроводительная документация (номер)'],
        "{UfCrm51741285487155}": payload['Сопроводительная документация (дата)'],
        "{UfCrm51741285539701}": payload['Акт отбора образцов (номер)'],
        "{UfCrm51741285553037}": payload['Акт отбора образцов (дата)'],
        "{UfCrm51741285167890}": payload['Описание образца'],
        "{UfCrm51741286822180}": payload['Нормативно-техническая документация на продукцию']
    }
