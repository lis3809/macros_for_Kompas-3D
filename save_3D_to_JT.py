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


def create_addition_format_param_and_save(iDocument3D, new_path):
    # additionPar - указатель на интерфейс ksAdditionFormatParam или IAdditionFormatParam,
    # определяющий параметры записи в дополнительный формат.
    ksAdditionFormatParam = iDocument3D.AdditionFormatParam()
    ksAdditionFormatParam.Init()
    ksAdditionFormatParam.SetObjectsOptions(0, True)  # Разрешение на чтение\запись твёрдых тел
    ksAdditionFormatParam.SetObjectsOptions(2, True)  # Разрешение на чтение\запись поверхностей
    ksAdditionFormatParam.SetObjectsOptions(4, True)  # Разрешение на чтение\запись кривых
    ksAdditionFormatParam.SetObjectsOptions(6, True)  # Разрешение на чтение (не применяется) \запись эскизов
    ksAdditionFormatParam.SetObjectsOptions(18, True)  # Разрешение на чтение\запись атрибутов объектов
    ksAdditionFormatParam.SetPlacement(iDocument3D.DefaultPlacement())
    ksAdditionFormatParam.format = 8  # формат JT
    # TODO
    # ksD3COInvisibleObjects  8  Разрешение на чтение (не применяется) \запись невидимых объектов
    # ksD3COPoints  10  Разрешение на чтение\запись точек
    # ksD3CODocumentProperties  12  Разрешение на чтение\запись информации о документе (автор, организация, комментарии)
    # D3COTechnicalDemand  14  Разрешение на чтение\запись технических требований
    # ksD3CODimensions  16 Разрешение на чтение\запись размеров
    # ksD3CBRep #  20 #  Разрешение на чтение\запись форм изделий в граничном представлении (только в JT)
    # ksD3CPolygonal  22  Разрешение на чтение\запись полигональных форм изделий
    # ksD3CPolygonalLOD0  24  Разрешение на чтение\запись полигональных форм изделий уровня детализации 0
    # ksD3CAssociated  26  Разрешение на чтение ассоциированной геометрии (резьбы и др)
    # ksD3COStyle  28  Разрешение на чтение\запись элементов оформления (цвет, начертание, и т.п.)
    # ksD3CODensity  30  Разрешение на чтение\запись единиц плотности
    # ksD3COValidationProperties  32  Разрешение на чтение\запись контрольных параметров - объёма, площади поверхности, центра масс


    # В этом разделе:
    #
    # angle - Максимально допустимое угловое отклонение касательных кривой или нормалей поверхности в соседних точках на расстоянии шага
    #
    # author - Автор
    #
    # configurationFileName - Путь к текущему файлу конфигурации
    #
    # comment - Комментарий
    #
    # configuration - Выбранная конфигурация
    #
    # createLocalComponents - TRUE - создавать вставки как локальные. FALSE - сохранять вставки в отдельных файлах
    #
    # format - Формат файла для записи модели
    #
    # formatBinary - Признак, определяющий тип файла (двоичный или текстовый)
    #
    # length - Максимально допустимое расстояние между соседними точками на расстоянии шага
    #
    # lengthUnits - Единицы измерения длины
    #
    # maxTeselationCellCount - Максимальное количество ячеек в строке и ряду триангуляционной сетки (если 0, то не задано)
    #
    # needCreateComponentsFiles - Создавать файлы компонентов
    #
    # organization - Организация
    #
    # password - Пароль для загрузки упрощенных вставок перед экспортом
    #
    # saveResultDocument - Сохранить полученный документ
    #
    # stepType - Способ вычисления приращения параметра при движении по объекту
    #
    # stitchSurfaces - Флаг необходимости сшивки поверхностей при импорте
    #
    # stitchPrecision - Точность сшивки поверхностей
    #
    # textExportForm - Признак, чтения\записи текстов
    #
    # topolgyIncluded - Признак, определяющий, включать ли топологию модели при экспорте

    iDocument3D.SaveAsToAdditionFormat(new_path, ksAdditionFormatParam)


def parse_kompas_files(paths):
    is_run = is_running()  # True, если программа Компас уже запущена

    # Подключаемся к программе
    kompas6_constants, kompas_object, application_api7 = get_kompas_api()

    for path in paths:
        # Интерфейс IKompasDocument
        iKompasDocument = application_api7.Documents.Open(PathName=path,
                                                          Visible=False,
                                                          ReadOnly=True)  # Откроем файл в невидимом режиме без права его изменять

        # Указатель на интерфейс документа трехмерной модели ksDocument3D.
        iDocument3D = kompas_object.ActiveDocument3D()

        # Если сборка
        if path.endswith('.a3d'):
            new_path = path.replace('.a3d', '.jt')
            print(new_path)

            create_addition_format_param_and_save(iDocument3D, new_path)

            # TODO create отдельный макрос
            print("Пробуем сохранение в др. версии")
            s = path.replace('.a3d', '_v20.a3d')
            iDocument3D.SaveAsEx(s, 12)

        # Если модель
        elif path.endswith('.m3d'):
            new_path = path.replace('.m3d', '.jt')
            print(new_path)

            create_addition_format_param_and_save(iDocument3D, new_path)

        iKompasDocument.Close(kompas6_constants.kdDoNotSaveChanges)  # Закроем файл без изменения

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
