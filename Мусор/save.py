import pythoncom
from win32com.client import Dispatch, gencache

# Временные переменные
input_file_path = r"C:\Users\Slon\Desktop\Для тестирования ПО\Каталог\219-04-сб34-01.frw"
output_png_path_PNG = r"C:\Users\Slon\Desktop\Для тестирования ПО\Каталог\219-04-сб34-01.png"
output_png_path_PDF = r"C:\Users\Slon\Desktop\Для тестирования ПО\Каталог\219-04-сб34-01.pdf"

#  Подключим константы API Компас
kompas6_constants = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants
kompas6_constants_3d = gencache.EnsureModule("{2CAF168C-7961-4B90-9DA2-701419BEEFE3}", 0, 1, 0).constants

#  Подключим описание интерфейсов API5
kompas6_api5_module = gencache.EnsureModule("{0422828C-F174-495E-AC5D-D31014DBBE87}", 0, 1, 0)
kompas_object_api5 = kompas6_api5_module.KompasObject(
    Dispatch("Kompas.Application.5")._oleobj_.QueryInterface(kompas6_api5_module.KompasObject.CLSID,
                                                             pythoncom.IID_IDispatch))

#  Подключим описание интерфейсов API7
kompas_api7_module = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
application7 = kompas_api7_module.IApplication(
    Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(kompas_api7_module.IApplication.CLSID,
                                                             pythoncom.IID_IDispatch))

# Загружаем документ from api7
kompas_document = application7.Documents.Open(input_file_path, False)

iConverter = application7.Converter(kompas_object_api5.ksSystemPath(5) + "\Pdf2d.dll")
print("convert")
iConverter.Convert(kompas_document.PathName, output_png_path_PDF, 0, False)
application7.MessageBoxEx("Создан файл PDF", "Сохранение в *.pdf", 64)

#
# def save_as_png(input_file_path, output_png_path):
#
#     kompas6_api5_module = gencache.EnsureModule("{0422828C-F174-495E-AC5D-D31014DBBE87}", 0, 1, 0)
#     kompas_api7_module = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
#
#     module = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
#
#     # Создаем объект KOMPAS-3D
#     kompas_app = win32com.client.Dispatch("Kompas.Application.5")
#
#     # Загружаем документ
#     kompas_doc = kompas_app.Documents.Open(input_file_path, False)
#
#     # Получаем интерфейс 2D-чертежа
#     kompas_layouts = kompas_doc.LayoutSheets
#     kompas_layout = kompas_layouts[0]  # Предполагаем, что у документа есть хотя бы один лист
#     kompas_view = kompas_layout.Views[0]  # И предполагаем, что на листе есть хотя бы один вид
#
#     # Устанавливаем параметры экспорта
#     export_params = kompas_app.SavingInPNG
#     export_params.ColorMode = 0  # 0 - цвет, 1 - черно-белый
#     export_params.Scale = 1.0  # Масштаб
#
#     # Сохраняем в PNG
#     kompas_view.ExportToBitmap(output_png_path, export_params)
#
#     # Закрываем документ
#     kompas_doc.Close()


# save_as_png(input_file_path, output_png_path)


"""Работа с АПИ 7"""
        # # Приводим интерфейс IKompasDocument к виду IKompasDocument3D
        # iKompasDocument3D = win32com.client.CastTo(iKompasDocument_api7, "IKompasDocument3D")
        # # Получаем указатель на интерфейс IPart7
        # iPart7 = iKompasDocument3D.TopPart
        #
        # # Выводим количество компонентов и обозначение каждого компонента в консоль
        # count = iPart7.InstanceCount(None)
        # for i in range(count):
        #     print(iPart7.Parts.Part(i).Marking)
