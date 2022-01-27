from xls_w import Excel
import os
from GUI import GUI


def eng_to_rus_letters(text):
    return text.replace('a', 'а').replace('e', 'е').replace('o', 'о')\
        .replace('p', 'р').replace('c', 'с').replace('y', 'у').replace('x', 'х')


def hl_column_number(xxl, text):
    text = ' '.join(str(text).strip().lower().split()) if ' ' in text else str(text).strip().lower()

    for count in range(1, xxl.size_string(1) + 1):
        cell_data = str(xxl.ws[f'{Excel.number_to_letter(count)}1'].value).strip().lower()
        if ' ' in cell_data:
            cell_data = ' '.join(cell_data.split())
        else:
            cell_data = cell_data

        if eng_to_rus_letters(cell_data) == eng_to_rus_letters(text):
            return count
    return None


def data_analysis(dir_scan, xxl, files_dir):

    # ===(initializing PROGRESSBAR)===
    pg_size = xxl.size_column('A') - 1
    pg_window = GUI.progress_bar(pg_size)
    pg_window.read(timeout=10)
    pg_window.TKroot.focus_force()
    # ===(initializing PROGRESSBAR)===

    # ===(Находим столбец для гиперссылок)===
    hl_column = hl_column_number(xxl, 'скан образ документа')

    for position, string_number in enumerate(range(2, xxl.size_column('A') + 1)):
        cell_data = str(xxl.ws[f'A{string_number}'].value).strip()
        cell_data = cell_data[:cell_data.rfind('.')] if '.' in cell_data else cell_data
        if '/' in cell_data:
            cell_data = cell_data.replace(r'/', r'-')
            xxl.ws[f'A{string_number}'].value = cell_data

        miss_file_count = 0
        for file in files_dir:
            file_type = file[:file.find('.') + 1] if file[:file.find('.') + 1].lower() in ['вх.', 'исх.'] else ''
            file_hl_name = file.split()[0] if ' ' in file else file[:file.rfind('.')]
            # file_extension = file[file.rfind('.') + 1:]
            file_name_for_find = file_hl_name[file_hl_name.rfind('.') + 1:] if '.' in file_hl_name else file_hl_name
            if '-' in file_name_for_find:
                if int(file_name_for_find[file_name_for_find.rfind('-') + 1:]) > 20:
                    file_name_for_find = file_name_for_find[:file_name_for_find.rfind('-')]

            hl_name = f'{file_type}{file_name_for_find}'
            if cell_data == file_name_for_find:
                if not xxl.check_hyperlink(hl_name, file, f'{Excel.number_to_letter(hl_column)}{string_number}'):
                    xxl.create_hyperlinks(hl_name, file, f'{Excel.number_to_letter(hl_column)}{string_number}')
            else:
                miss_file_count += 1
        print(f'Не найдено сопоставление регистрационному номеру {cell_data} среди файлов.') if miss_file_count == len(files_dir) else None
        # ===(increment PROGRESSBAR)===
        pg_window['PROGRESSBAR'].update_bar(position + 1)
    # ===(close PROGRESSBAR)===
    pg_window.close()


def body(file_path, dir_scan, ws_name, settings):
    xxl = Excel(file_path, dir_scan, ws_name, settings)
    files_dir = os.listdir(path=dir_scan)
    print(f'Получен список файлов в папке {dir_scan}.')
    print('Анализ данных и формирование гиперссылок...')
    data_analysis(dir_scan, xxl, files_dir)
    print('Гиперссылки сформированы.')
    print('Complete...' + '\n' * 1)


if __name__ == '__main__':
    pass
