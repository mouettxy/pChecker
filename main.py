# -*- coding: utf-8 -*-
import os
import argparse
from CheckPresentationMain import CheckPresentationAnalyze, PrintTo, CheckPresentationGetData
from pptx import Presentation

def get_slides(file_pptx):
    return CheckPresentationGetData(Presentation(file_pptx)).get_slides_length()
def get_result(file_pptx):
    return CheckPresentationAnalyze(Presentation(file_pptx)).analyze_results()


description = (
    'Укажите путь к файлу'
)

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=description)
    parser.add_argument("-o", "--output",
                        type=str,
                        help="Путь к файлу в который необходимо выгрузить результаты."
                        )
    parser.add_argument("-e", "--encoding",
                        type=str,
                        help="Если необходимо указать кодировку содержимого файла. Поддерживается только формат .csv."
                        )
    parser.add_argument("-t", "--typeof",
                        type=str,
                        help="Тип(расширение) файла в который будет происходить выгрузка. Доступные: txt, csv.")
    parser.add_argument("-m", "--mode",
                        type=str,
                        help="Как будет записываться файл? write - дописывание, rewrite - перезапись.")
    parser.add_argument("presentation_file",
                        type=str,
                        help="Укажите путь к файлу презентации."
                        )
    args = parser.parse_args()
    file = args.presentation_file
    output = args.output
    typeof = args.typeof
    mode = args.mode
    encoding = args.encoding
    if file:
        if all([output, typeof, mode]):
            if encoding:
                if typeof == 'csv':
                    Print = PrintTo(get_result(file), output, file, encoding=encoding)
                    to_file = Print.csv(mode)
                    print(to_file)
            else:
                if typeof == 'csv':
                    Print = PrintTo(get_result(file), output, file)
                    to_file = Print.csv(mode)
                    print(to_file)
                elif typeof == 'txt':
                    Print = PrintTo(get_result(file), output, file)
                    to_file = Print.txt(mode)
                    print(to_file)
        else:
            slides = get_slides(file_pptx=file)
            if slides < 3 or slides > 4:
                print(f'В презентации должно быть ровно 3 слайда. Найдено {slides}.')
            result = get_result(file_pptx=file)
            for res in result:
                if res == 'Количество слайдов':
                    print(f'{res} => {result[res]}')
                else:
                    print(f'{res} => {"Да" if result[res] else "Нет"}')


