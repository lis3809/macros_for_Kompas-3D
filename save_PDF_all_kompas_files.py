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
        kompas_document_api7 = application_api7.Documents.Open(PathName=path,
                                                               Visible=False,
                                                               ReadOnly=True)  # Откроем файл в невидимом режиме без права его изменять

        # Если чертеж
        if path.endswith('.cdw'):
            new_path_to_PDF = path.replace('.cdw', '.pdf')
            print(new_path_to_PDF)
            iConverter = application_api7.Converter(kompas_object.ksSystemPath(5) + "\Pdf2d.dll")
            iConverter.Convert(kompas_document_api7.PathName, new_path_to_PDF, 0, False)
        # Если спецификация
        elif path.endswith('.spw'):
            new_path_to_PDF = path.replace('.spw', '.pdf')
            print(new_path_to_PDF)
            iConverter = application_api7.Converter(kompas_object.ksSystemPath(5) + "\Pdf2d.dll")
            iConverter.Convert(kompas_document_api7.PathName, new_path_to_PDF, 0, False)
        # Если фрагмент
        elif path.endswith('.frw'):
            new_path_to_PDF = path.replace('.frw', '.pdf')
            print(new_path_to_PDF)
            iConverter = application_api7.Converter(kompas_object.ksSystemPath(5) + "\Pdf2d.dll")
            iConverter.Convert(kompas_document_api7.PathName, new_path_to_PDF, 0, False)

        kompas_document_api7.Close(kompas6_constants.kdDoNotSaveChanges)  # Закроем файл без изменения

    if not is_run:
        application_api7.Quit()  # Закрываем программу при необходимости


if __name__ == "__main__":
    root = Tk()
    root.withdraw()  # Скрываем основное окно и сразу открываем окно выбора файлов

    filenames = askopenfilenames(title="Выберети файлы Компас-3D",
                                 filetypes=[('Компас 3D', '*.cdw'), ('Компас 3D', '*.spw'), ('Компас 3D', '*.frw')])
    # ==========Вызов основной функции===============
    parse_kompas_files(filenames)
    # =====================КОНЕЦ=====================
    # Уничтожаем основное окно
    root.destroy()
    root.mainloop()
