import os

from pathlib import Path

import pandas as pd
from pandas.errors import EmptyDataError


class Print:
    """
    Класс Print реализует распечатку результатов проверки в различне типы файлов.
    :param results: Переведённый результат работы функции анализа
    :param path_to_output: Полный путь к файлу куда будет выгружен результат
    :param path_to_pptx: Полный путь к файлу презентации
    :param encoding: Кодировка которая должна получиться в выходном файле
    TODO: Добавить поддержку изменения кодировки к PrintTo.txt
    TODO: Добавить генерацию файла Excel PrintTo.excel
    """

    def __init__(self, results, path_to_output, path_to_pptx, encoding='utf-8'):
        self.results = results
        self.path_to_output = Path(path_to_output)
        self.path_to_pptx = Path(path_to_pptx)
        self.encoding = encoding
        self.output_name, self.output_extension = os.path.splitext(self.path_to_output)
        self.results_keys = []
        self.results_values = []
        for result in self.results:
            self.results_keys.append(result)
            self.results_values.append(self.results[result])
        self.results_zip = list(zip(self.results_keys, self.results_values))
        self.mode_list = ['write', 'rewrite']

    def _extension_check(self, extension):
        """
        :param extension: Расширение файла str (пример: ".txt")
        :return: None в случае когда расширение совпадает с вызванное функцией, иначе генерирует подробную ошибку
        """
        if self.output_extension != extension:
            raise Exception('Ошибка! Неверное расширение файла. \n'
                            f'Ожидалось "{extension}", введено "{self.output_extension}". \n'
                            'Проверьте правильность введённого пути к конечному файлу.')
        return

    def _write_mode_check(self, mode):
        """
        :param mode: Режим записи в файл из значений self.mode_list
        :return: Возвращает str применимую к режиму записи файла
        """
        mode_string = ', '.join(self.mode_list)
        if mode == 'write':
            return 'a'
        elif mode == 'rewrite':
            return 'w'
        else:
            raise Exception('Ошибка! Неверно указан метод открытия файла.\n'
                            f'Ожидаемые значения "{mode_string}", введено "{mode}". \n'
                            'Доступные методы: \n'
                            ' write - файл не будет перезаписан, полученный результат добавиться к предыдущему. \n'
                            ' rewrite - файл будет перезаписан, данные которые были до этого будут удалены. \n')

    def _empty(self, file_path, extension):
        """
        :param file_path: Путь к файлу (пример: 'C:/path/to/file.extension')
        :param extension: Расширение файла (пример: '.txt')
        :return: True если файл пуст, False если файл не пуст
        """
        if extension == '.txt':
            if os.stat(file_path).st_size > 0:
                return False
            return True
        elif extension == '.csv':
            try:
                file_contents = pd.read_csv(file_path)
                return file_contents.empty
            except EmptyDataError:
                return True
            except (OSError, IOError):
                create_file = open(file_path, 'w')
                create_file.close()
                return True
        elif extension == '.xlsx':
            try:
                file_contents = pd.read_excel(file_path)
                return file_contents.empty
            except EmptyDataError:
                return True
            except (OSError, IOError):
                create_file = open(file_path, 'w')
                create_file.close()
                return True

    def _write_to_csv(self, data, file_path, mode, encoding):
        """
        :param data: Список для заполнения
        :param file_path: Путьк файлу
        :param mode: Режим записи в файл
        :param encoding: Кодировка текста внутри
        :return: None или Str в случае выполнения, игаче генерирует подробные ошибки.
        """
        data_frame = pd.DataFrame(data)
        try:
            return data_frame.to_csv(file_path, mode=mode, header=True, encoding=encoding, index=False)
        except PermissionError:
            raise Exception('Недостаточно прав для открытия, создания, или записи в файл. \n'
                            'Попробуйте закрыть файл, или проверить права на запись и чтение файла. \n')

    def txt(self, mode):
        """
        :param mode: Режим записи в файл
        :return: Возвращает str в случае успешного выполнения, иначе генерирует подробные ошибки
        """
        self._extension_check('.txt')
        mode = self._write_mode_check(mode)
        txt_file = open(self.path_to_output, mode=mode, encoding='utf-8')
        if not (self._empty(self.path_to_output, '.txt')):
            txt_file.write('\n')
        txt_file.write(f'Проверка файла: {self.path_to_pptx}\n')
        for result in self.results_zip:
            if result[0] == 'Количество слайдов':
                txt_file.write(f'{result[0]} => {result[1]}\n')
            else:
                txt_file.write(f'{result[0]} => {"Да" if result[1] else "Нет"}\n')
        txt_file.close()
        return 'Успешная запись в файл'

    def csv(self, mode):
        """
        :param mode: Режим записи в файл
        :return: Возвращает str в случае успешного выполнения, иначе генерирует подробные ошибки
        """
        self._extension_check('.csv')
        mode = self._write_mode_check(mode)
        data_without_columns = [self.results_values]
        data_with_columns = [self.results]
        if self._empty(self.path_to_output, '.csv') or (not (self._empty(self.path_to_output, '.csv')) and mode == 'w'):
            self._write_to_csv(data_with_columns, self.path_to_output, mode, self.encoding)
            return 'Успешная запись в файл'
        elif not (self._empty(self.path_to_output, '.csv')) and mode == 'a':
            self._write_to_csv(data_without_columns, self.path_to_output, mode, self.encoding)
            return 'Успешная запись в файл'

    def excel(self, mode):
        pass
