import argparse
import pythoncom
import os
from win32com.client import Dispatch, gencache
import subprocess
from tkinter import Tk
from tkinter.filedialog import askopenfilenames

# pyinstaller --noconsole --onefile .\save_to_another_file_version.py

# Добавляем аргументы
parser = argparse.ArgumentParser()
parser.add_argument("--path_dir", type=str)
parser.add_argument("--vers_k", type=str)
# Парсим аргументы
args = parser.parse_args()

dict_vers_kompas_file = {"5.11": 1,
                         "6.0": 2,
                         "6+": 3,
                         "7.0": 4,
                         "7+": 5,
                         "8.0": 6,
                         "8+": 7,
                         "9.0": 8,
                         "10.0": 9,
                         "11.0": 10,
                         "12.0": 11,
                         "13.0": 12,
                         "14.0": 13,
                         "14sp1": 14,
                         "14sp2": 15,
                         "15.0": 16,
                         "15sp1": 17,
                         "15sp2": 18,
                         "16": 19,
                         "16sp1": 20,
                         "17": 21,
                         "17sp1": 22,
                         "18": 23,
                         "18sp1": 24,
                         "19": 25,
                         "20": 26,
                         "21": 27,
                         "22": 28}


# Подключение к API7 программы Kompas 3D
def get_kompas_api():
    #  Подключим константы API Компас
    kompas6_constants = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants

    #  Подключим описание интерфейсов API5
    kompas6_api5_module = gencache.EnsureModule("{0422828C-F174-495E-AC5D-D31014DBBE87}", 0, 1, 0)
    kompas_object = kompas6_api5_module.KompasObject(
        Dispatch("Kompas.Application.5")._oleobj_.QueryInterface(kompas6_api5_module.KompasObject.CLSID,
                                                                 pythoncom.IID_IDispatch))

    #  Подключим описание интерфейсов API7
    kompas_api7_module = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
    application_api7 = kompas_api7_module.IApplication(
        Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(kompas_api7_module.IApplication.CLSID,
                                                                 pythoncom.IID_IDispatch))
    return kompas6_constants, kompas_object, application_api7


# Функция проверяет, запущена ли программа Kompas-3D
def is_running():
    proc_list = subprocess.Popen('tasklist /NH /FI "IMAGENAME eq KOMPAS*"',
                                 shell=False,
                                 stdout=subprocess.PIPE).communicate()[0]
    return True if proc_list else False


def parse_kompas_files(paths, vers_k, vers_val, new_path_to_dir):
    is_run = is_running()  # True, если программа Компас уже запущена

    # Подключаемся к программе
    kompas6_constants, kompas_object, application_api7 = get_kompas_api()

    for path in paths:
        # Интерфейс IKompasDocument
        iKompasDocument = application_api7.Documents.Open(PathName=path,
                                                          Visible=False,
                                                          ReadOnly=True)  # Откроем файл в невидимом режиме без права его изменять

        # Если сборка
        if path.endswith('.a3d'):
            # Указатель на интерфейс документа трехмерной модели ksDocument3D.
            iDocument3D = kompas_object.ActiveDocument3D()
            new_path = new_path_to_dir + "\\" + os.path.basename(path).replace('.a3d', f'_v{vers_k}.a3d')
            iDocument3D.SaveAsEx(new_path, vers_val)

        # Если модель
        elif path.endswith('.m3d'):
            # Указатель на интерфейс документа трехмерной модели ksDocument3D.
            iDocument3D = kompas_object.ActiveDocument3D()
            new_path = new_path_to_dir + "\\" + os.path.basename(path).replace('.m3d', f'_v{vers_k}.m3d')
            iDocument3D.SaveAsEx(new_path, vers_val)

        # Если чертеж
        elif path.endswith('.cdw'):
            # получаем указатель на интерфейс графического документа ksDocument2D.
            ksDocument2D = kompas_object.ActiveDocument2D()
            new_path = new_path_to_dir + "\\" + os.path.basename(path).replace('.cdw', f'_v{vers_k}.cdw')
            ksDocument2D.ksSaveDocumentEx(new_path, vers_val)

        # Если спецификация
        elif path.endswith('.spw'):
            # получаем указатель на интерфейс графического документа ksDocument2D.
            ksSpcDocument = kompas_object.SpcActiveDocument()
            new_path = new_path_to_dir + "\\" + os.path.basename(path).replace('.spw', f'_v{vers_k}.spw')
            ksSpcDocument.ksSaveDocumentEx(new_path, vers_val)

        # Если фрагмент
        elif path.endswith('.frw'):
            # получаем указатель на интерфейс графического документа ksDocument2D.
            ksDocument2D = kompas_object.ActiveDocument2D()
            new_path = new_path_to_dir + "\\" + os.path.basename(path).replace('.frw', f'_v{vers_k}.frw')
            ksDocument2D.ksSaveDocumentEx(new_path, vers_val)

        iKompasDocument.Close(kompas6_constants.kdDoNotSaveChanges)  # Закроем файл без изменения

    if not is_run:
        application_api7.Quit()  # Закрываем программу при необходимости


if __name__ == "__main__":
    root = Tk()
    root.withdraw()  # Скрываем основное окно и сразу открываем окно выбора файлов

    # Проверяем наличие и правильность параметров
    vers_v = 0
    if args.vers_k:
        vers_v = dict_vers_kompas_file.get(args.vers_k)
    else:
        print("ERRRROR")
        # Уничтожаем основное окно
        root.destroy()
        root.mainloop()

    filenames = askopenfilenames(title="Выберети файлы Компас-3D",
                                 filetypes=[('Компас 3D', '*.a3d'), ('Компас 3D', '*.m3d'), ('Компас 3D', '*.cdw'),
                                            ('Компас 3D', '*.spw'), ('Компас 3D', '*.frw')])

    # ==========Вызов основной функции===============
    parse_kompas_files(filenames, args.vers_k, vers_v, args.path_dir)
    # =====================КОНЕЦ=====================
    # Уничтожаем основное окно
    root.destroy()
    root.mainloop()
