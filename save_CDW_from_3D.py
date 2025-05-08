import argparse
import os
import subprocess
from tkinter import Tk
from tkinter.filedialog import askopenfilenames

import pythoncom
from openpyxl import Workbook
from win32com.client import Dispatch, gencache

# pyinstaller --noconsole --onefile .\save_to_PDF_all_kompas_files.py

# Добавляем аргументы
parser = argparse.ArgumentParser()
parser.add_argument("--path_dir", type=str)

# Парсим аргументы
args = parser.parse_args()


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
    # Список для хранения данных о составных частях
    list_val = []
    is_run = is_running()  # True, если программа Компас уже запущена

    # Подключаемся к программе
    kompas6_constants, kompas_object, application_api7 = get_kompas_api()

    for path in paths:

        print(path)
        iKompasDocument_api7 = application_api7.Documents.Open(PathName=path,
                                                               Visible=False,
                                                               ReadOnly=False)  # Откроем файл в невидимом режиме без права его изменять

        # Если это сборочная единица
        if path.endswith(".a3d"):
            parse_specification(kompas_object, list_val, path)
            save_CDW(kompas6_constants, kompas_object, application_api7, path)

        # Если это деталь
        elif path.endswith(".m3d"):
            save_CDW(kompas6_constants, kompas_object, application_api7, path)
        else:
            print("Неизвестный файл")

        # iKompasDocument_api7.RebuildDocument()
        iKompasDocument_api7.Close(kompas6_constants.kdSaveChanges)  # Закроем файл сохранив изменения

    if not is_run:
        application_api7.Quit()  # Закрываем программу при необходимости

    return list_val


def parse_specification(kompas_object, list_val, path):
    """Работа с АПИ 5"""
    ksDocument3D_api5 = kompas_object.ActiveDocument3D()
    # Интерфейс спецификации ksSpecification
    ksSpecification = ksDocument3D_api5.GetSpecification()
    # Создаем итератор для перебора
    ksIterator = kompas_object.GetIterator()
    if ksIterator.ksCreateSpcIterator("graphic.lyt", 1, 3):
        spcObj = ksIterator.ksMoveIterator("F")
        while spcObj != 0:
            vals = []
            numPos = ksSpecification.ksGetSpcObjectColumnText(spcObj, 3, 1, 0)  # SPC_CLM_POS 3 - позиция
            designator = ksSpecification.ksGetSpcObjectColumnText(spcObj, 4, 1,
                                                                  0)  # SPC_CLM_MARK 4 - обозначение
            name = ksSpecification.ksGetSpcObjectColumnText(spcObj, 5, 1, 0)  # SPC_CLM_NAME 5 - наименование
            count = ksSpecification.ksGetSpcObjectColumnText(spcObj, 6, 1, 0)  # SPC_CLM_COUNT 6 - количество
            # Записываем переменные в список
            vals.append(numPos)
            vals.append(designator)
            vals.append(name)
            vals.append(count)
            # Последним элементом вставляем куда входит
            vals.append(os.path.splitext(os.path.basename(path))[0])
            list_val.append(vals)
            # Получаем новое значение из итератора
            spcObj = ksIterator.ksMoveIterator("N")


def save_CDW(kompas6_constants, kompas_object, application_api7, filePath_3D):
    """Вариант с основной надписью"""
    #  Создаем новый документ
    # cdw_doc_IKompasDocument = application_api7.Documents.AddWithDefaultSettings(kompas6_constants.ksDocumentDrawing, True)
    # ksDocument2D = kompas_object.ActiveDocument2D()
    """================"""

    """Ввариант без основной надписи"""
    # Подготавливаем параметры документа
    documentParam = kompas_object.GetParamStruct(kompas6_constants.ko_DocumentParam)
    documentParam.Init()
    documentParam.type = 1  # lt_DocSheetStandart - чертеж
    # documentParam.type = 3  # lt_DocFragment 3 - фрагмент

    # указатель на интерфейс параметров оформления документа ksSheetPar
    sheetPar = documentParam.GetLayoutParam()
    # sheetPar.SetLayoutName()    #
    sheetPar.shtType = 15  # Тип документа 15 - без оформления

    # Создаем документ с параметрами
    kompas_object.Document2D().ksCreateDocument(documentParam)

    # указатель на интерфейс графического документа ksDocument2D
    ksDocument2D = kompas_object.ActiveDocument2D()
    cdw_doc_IKompasDocument = application_api7.ActiveDocument
    """================"""

    '''Структура параметров ассоциативного вида'''
    iAssociationViewParam = kompas_object.GetParamStruct(kompas6_constants.ko_AssociationViewParam)
    iAssociationViewParam.Init()
    # Разнести
    iAssociationViewParam.disassembly = True
    iAssociationViewParam.fileName = filePath_3D
    # Признак отрисовки невидимых линий
    iAssociationViewParam.hiddenLinesShow = False
    iAssociationViewParam.hiddenLinesStyle = 4
    # Признак проецирования тел
    iAssociationViewParam.projBodies = True
    # Проекционная связь
    iAssociationViewParam.projectionLink = False
    # Имя проекции (из списка проекций в документе-источнике)
    iAssociationViewParam.projectionName = "#Изометрия"
    # Признак проецирования поверхностей
    iAssociationViewParam.projSurfaces = True
    # Признак проецирования резьбы
    iAssociationViewParam.projThreads = True
    # Одинаковая штриховка всех деталей сборки
    iAssociationViewParam.sameHatch = False
    # Признак разрез/сечение
    iAssociationViewParam.section = False
    # Признак отрисовки видимых линий перехода
    iAssociationViewParam.tangentEdgesShow = True
    iAssociationViewParam.tangentEdgesStyle = 2
    iAssociationViewParam.visibleLinesStyle = 1

    '''Структура параметров вида'''
    iViewParam = kompas_object.GetParamStruct(kompas6_constants.ko_ViewParam)
    iViewParam.Init()
    # угол поворота вида
    iViewParam.angle = 0
    # цвет вида в активном состоянии
    iViewParam.color = 0
    iViewParam.name = "Вид 1"
    iViewParam.scale_ = 1
    # состояние вида (stACTIVE 0  - активный (видимый фоновый)
    # stREADONLY 1 - фоновый
    # stINVISIBLE 2 - невидимый (погашенный)
    # stCURRENT 3 - текущий
    iViewParam.state = 3

    # точка привязки вида
    iViewParam.x = 0
    iViewParam.y = 0

    # Создать произвольный ассоциативный вид
    ksDocument2D.ksCreateSheetArbitraryView(iAssociationViewParam, 0)

    # Если сборка
    if filePath_3D.endswith('.a3d'):
        new_path = filePath_3D.replace('.a3d', '.cdw')
        cdw_doc_IKompasDocument.SaveAs(new_path)
    # Если модель
    elif filePath_3D.endswith('.m3d'):
        new_path = filePath_3D.replace('.m3d', '.cdw')
        cdw_doc_IKompasDocument.SaveAs(new_path)

    cdw_doc_IKompasDocument.Close(kompas6_constants.kdSaveChanges)


def print_to_excel(list_val, directory):
    wb = Workbook()
    sheet = wb.active

    # Создаём заголовок таблицы
    header = ["Номер позиции", "Обозначение", "Наименование", "Количество в сб.ед."]
    sheet.append(header)

    # Заполняем таблицу
    for i, row in enumerate(list_val):
        sheet.append(row)

    wb.save(directory + '\\data_from_3D.xlsx')
    wb.close()


if __name__ == "__main__":
    root = Tk()
    root.withdraw()  # Скрываем основное окно и сразу открываем окно выбора файлов

    filenames = askopenfilenames(title="Выберети файлы Компас-3D",
                                 filetypes=[('Компас 3D', '*.a3d'), ('Компас 3D', '*.m3d')])
    # ==========Вызов основной функции===============
    list_data_from_3D = parse_kompas_files(filenames)
    if len(list_data_from_3D) > 0:
        print_to_excel(list_data_from_3D, os.path.dirname(filenames[0]))

    # =====================КОНЕЦ=====================
    # Уничтожаем основное окно
    root.destroy()
    root.mainloop()
