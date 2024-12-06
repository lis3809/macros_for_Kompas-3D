import qrcode
import pythoncom
from win32com.client import Dispatch, gencache
import os

iApplication = Dispatch('KOMPAS.Application.7')
iDocument = iApplication.ActiveDocument
print(iDocument.Name)
iLayoutSheets = iDocument.LayoutSheets
iLayoutSheet = iLayoutSheets.ItemByNumber(1)
iStamp = iLayoutSheet.Stamp
iText = iStamp.Text(1)  # номер ячейки для считывания данных наименования
Str1 = iText.Str
iText = iStamp.Text(2)  # номер ячейки для считывания данных обозначения
Str2 = iText.Str
print(Str1)
print(Str2)
data = Str1, Str2
data = str(data).replace("'", "").replace(")", "").replace("(", "")

# output file name
imgname = Str2 + '.png'

# generate qr code
img = qrcode.make(data)
# save img to a file
img.save(imgname)
# Подключим константы API Компас
kompas6_constants = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants
# Подключим описание интерфейсов API5
kompas6_api5_module = gencache.EnsureModule("{0422828C-F174-495E-AC5D-D31014DBBE87}", 0, 1, 0)
kompas_object = kompas6_api5_module.KompasObject(
    Dispatch("Kompas.Application.5")._oleobj_.QueryInterface(kompas6_api5_module.KompasObject.CLSID,
                                                             pythoncom.IID_IDispatch))
iDocument2D = kompas_object.ActiveDocument2D()
iRasterParam = kompas6_api5_module.ksRasterParam(kompas_object.GetParamStruct(kompas6_constants.ko_RasterParam))
iRasterParam.Init()
iRasterParam.embeded = True
imgName = os.getcwd() + '\\' + imgname
print(imgName)
iRasterParam.fileName = imgName
iPlacementParam = kompas6_api5_module.ksPlacementParam(
    kompas_object.GetParamStruct(kompas6_constants.ko_PlacementParam))
iPlacementParam.Init()
iPlacementParam.angle = 0
iPlacementParam.scale_ = 0.5
iPlacementParam.xBase = 0
iPlacementParam.yBase = 0
iRasterParam.SetPlace(iPlacementParam)
iDocument2D.ksInsertRaster(iRasterParam)
