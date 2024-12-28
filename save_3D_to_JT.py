import pythoncom
from win32com.client import Dispatch, gencache
import subprocess
from tkinter import Tk
from tkinter.filedialog import askopenfilenames


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


# Функция проверяет, запущена ли программа Kompas 3D
def is_running():
    proc_list = subprocess.Popen('tasklist /NH /FI "IMAGENAME eq KOMPAS*"',
                                 shell=False,
                                 stdout=subprocess.PIPE).communicate()[0]
    return True if proc_list else False


def parse_kompas_files(paths):
    is_run = is_running()  # True, если программа Компас уже запущена

    # Подключаемся к программе
    kompas6_constants, kompas_object, application_api7 = get_kompas_api()

    for path in paths:
        kompas_document = application_api7.Documents.Open(PathName=path,
                                                          Visible=False,
                                                          ReadOnly=True)  # Откроем файл в невидимом режиме без права его изменять

        # Если сборка
        if path.endswith('.a3d'):
            new_path = path.replace('.a3d', '.jt')
            print(new_path)
            kompas_document.SaveAs(new_path)
        # Если модель
        elif path.endswith('.m3d'):
            new_path = path.replace('.m3d', '.jt')
            print(new_path)
            kompas_document.SaveAs(new_path)

        kompas_document.Close(kompas6_constants.kdDoNotSaveChanges)  # Закроем файл без изменения

    if not is_run:
        application_api7.Quit()  # Закрываем программу при необходимости


if __name__ == "__main__":
    root = Tk()
    root.withdraw()  # Скрываем основное окно и сразу открываем окно выбора файлов

    filenames = askopenfilenames(title="Выберети файлы Компас-3D",
                                 filetypes=[('Компас 3D', '*.a3d'), ('Компас 3D', '*.m3d')])
    # ==========Вызов основной функции===============
    parse_kompas_files(filenames)
    # =====================КОНЕЦ=====================
    # Уничтожаем основное окно
    root.destroy()
    root.mainloop()
