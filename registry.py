from xls_w import Excel
import os
from GUI import GUI
from natsort import natsorted, ns


def light_files_in_dir(list_files):
    list_files = list(str(file) for file in list_files if not str(file)[:str(file).rfind('.')].isdigit())
    list_files = list(map(lambda x: x[x.rfind('№') + 1:x.rfind('.')].lower(), list_files))
    return list_files


def miss_files(list1, list2):
    miss_list = list(set(list1) - set(list2))
    return natsorted(miss_list, alg=ns.IGNORECASE)  # or alg=ns.IC


def data_analysis(dir_scan, xxl, files_dir):
    import time
    for string_number in range(2, xxl.size_column('A') + 1):
        cell_data = str(xxl.ws[f'A{string_number}'].value).strip()
        cell_data = cell_data[:cell_data.rfind('.')] if '.' in cell_data else cell_data
        if '/' in cell_data:
            cell_data = cell_data.replace(r'/', r'-')
            xxl.ws[f'A{string_number}'].value = cell_data

        for file in files_dir:
            file_type = file[:file.find('.') + 1] if file[:file.find('.') + 1].lower() in ['вх.', 'исх.'] else ''
            file_hl_name = file.split()[0] if '' in file else file
            file_extension = file[file.rfind('.') + 1:]
            file_name_for_find = file_hl_name[file_hl_name.rfind('.') + 1:]
            if '-' in file_name_for_find:
                if int(file_name_for_find[file_name_for_find.rfind('-') + 1:]) > 20:
                    file_name_for_find = file_name_for_find[:file_name_for_find.rfind('-')]

            hl_name = f'{file_type}{file_name_for_find}'
            # hl_link = f'{dir_scan}/{file}'

            if cell_data == file_name_for_find:
                if not xxl.check_hyperlink(hl_name, file, f'M{string_number}'):
                    print('создаем')
                    xxl.create_hyperlinks(hl_name, file, f'M{string_number}')

            # print(file_type, file_hl_name, file_extension, file_name_for_find)
            # print(hl_name)
            # print(hl_link)


def body(registry_path, dir_scan, ws_name, settings):
    xxl = Excel(registry_path, dir_scan, ws_name, settings)

    # registry_path = xxl.get_path_active_book() if registry_path in '' else registry_path
    # if registry_path[registry_path.rfind('\\') + 1:registry_path.rfind('.')].lower().count('исходящ') > 0:
    #     file_pref = 'исх.№'
    #     print('Загруженный документ идентифицирован как реестр Исходящих.')
    # elif registry_path[registry_path.rfind('\\') + 1:registry_path.rfind('.')].lower().count('входящ') > 0:
    #     file_pref = 'вход.№'
    #     print('Загруженный документ идентифицирован как реестр Входящих.')
    # else:
    #     file_pref = '№'
    #     print('Реестр не идентифицирован.')

    files_column = xxl.get_column('A')
    print(f'Получен список регистрационных номеров из столбца "А".')
    # files_a_sort = list(map(lambda x: str(x).replace(r'/', r'-').strip().split()[0], files_a))
    files_dir = os.listdir(path=dir_scan)
    print(f'Получен список файлов в папке {dir_scan}.')

    # files_dir_clear = light_files_in_dir(files_dir)
    # miss_list = miss_files(files_a_sort, files_dir_clear)
    # for miss in miss_list:
    #     print(f'Не найдено сопоставление регистрационному номеру {miss} среди файлов.')

    print('Анализ данных и формирование гиперссылок...')
    data_analysis(dir_scan, xxl, files_dir)

    # import time
    # for string_number in range(2, xxl.size_column('A') + 1):
    #     cell_data = str(xxl.ws[f'A{string_number}'].value).strip()
    #     cell_data = cell_data[:cell_data.rfind('.')] if '.' in cell_data else cell_data
    #     if '/' in cell_data:
    #         cell_data = cell_data.replace(r'/', r'-')
    #         xxl.ws[f'A{string_number}'].value = cell_data
    #
    #     for file in files_dir:
    #         file_type = file[:file.find('.') + 1] if file[:file.find('.') + 1].lower() in ['вх.', 'исх.'] else ''
    #         file_hl_name = file.split()[0] if '' in file else file
    #         file_extension = file[file.rfind('.') + 1:]
    #         file_name_for_find = file_hl_name[file_hl_name.rfind('.') + 1:]
    #         if '-' in file_name_for_find:
    #             if int(file_name_for_find[file_name_for_find.rfind('-') + 1:]) > 20:
    #                 file_name_for_find = file_name_for_find[:file_name_for_find.rfind('-')]
    #
    #         hl_name = f'{file_type}{file_name_for_find}'
    #         hl_link = f'{dir_scan}/{file}'
    #
    #
    #
    #         print(file_type, file_hl_name, file_extension, file_name_for_find)
    #         print(hl_name)
    #         print(hl_link)
    #         time.sleep(5)


    # pg_size = len(files_a)
    # pg_window = GUI.progress_bar(pg_size)
    # pg_window.read(timeout=10)
    # pg_window.TKroot.focus_force()
    #
    #
    #
    # for position, file_a in enumerate(files_a, 3):
    #     file_a_clear = file_a.replace(r'/', r'-').strip().split()[0]
    #
    #     pg_window['PROGRESSBAR'].update_bar(position + 1)
    #
    #     for file_dir in files_dir:
    #         if not file_dir.isdigit():
    #             file_type = file_dir[file_dir.rfind('.'):].lower()
    #             file_dir_clear = file_dir[file_dir.rfind('№') + 1:file_dir.rfind('.')].lower()
    #
    #             if file_dir_clear == file_a_clear:
    #                 name = f'{file_a_clear}'
    #                 link_name = f'{file_a_clear}{file_type}'
    #                 if not xxl.check_hyperlink(name, link_name, position):
    #                     xxl.create_hyperlinks(name, link_name, position)
    #
    # pg_window.close()
    # print('Гиперссылки сформированы.')
    # print('Complete...' + '\n' * 1)


if __name__ == '__main__':
    pass
