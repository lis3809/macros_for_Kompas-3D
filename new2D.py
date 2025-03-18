# -*- coding: utf-8 -*-
#|new2D

import pythoncom
from win32com.client import Dispatch, gencache

#  Подключим константы API Компас
kompas6_constants = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants
kompas6_constants_3d = gencache.EnsureModule("{2CAF168C-7961-4B90-9DA2-701419BEEFE3}", 0, 1, 0).constants

#  Подключим описание интерфейсов API5
kompas6_api5_module = gencache.EnsureModule("{0422828C-F174-495E-AC5D-D31014DBBE87}", 0, 1, 0)
kompas_object = kompas6_api5_module.KompasObject(Dispatch("Kompas.Application.5")._oleobj_.QueryInterface(kompas6_api5_module.KompasObject.CLSID, pythoncom.IID_IDispatch))

#  Подключим описание интерфейсов API7
kompas_api7_module = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
application = kompas_api7_module.IApplication(Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(kompas_api7_module.IApplication.CLSID, pythoncom.IID_IDispatch))

Documents = application.Documents
#  Создаем новый документ
kompas_document = Documents.AddWithDefaultSettings(kompas6_constants.ksDocumentDrawing, True)

kompas_document_2d = kompas_api7_module.IKompasDocument2D(kompas_document)
iDocument2D = kompas_object.ActiveDocument2D()

iAssociationViewParam = kompas6_api5_module.ksAssociationViewParam(kompas_object.GetParamStruct(kompas6_constants.ko_AssociationViewParam))
iAssociationViewParam.Init()
iAssociationViewParam.disassembly = False
iAssociationViewParam.fileName = "C:\Users\Slon\Desktop\Для тестирования ПО\3D\14. 675-60-сб2 - Артемьев\675-60-сб113.a3d"
iAssociationViewParam.hiddenLinesShow = False
iAssociationViewParam.hiddenLinesStyle = 4
iAssociationViewParam.projBodies = True
iAssociationViewParam.projectionLink = False
iAssociationViewParam.projectionName = "#Изометрия"
iAssociationViewParam.projSurfaces = False
iAssociationViewParam.projThreads = True
iAssociationViewParam.sameHatch = False
iAssociationViewParam.section = False
iAssociationViewParam.tangentEdgesShow = False
iAssociationViewParam.tangentEdgesStyle = 2
iAssociationViewParam.visibleLinesStyle = 1
iViewParam = kompas6_api5_module.ksViewParam(iAssociationViewParam.GetViewParam())
iViewParam.Init()
iViewParam.angle = 0
iViewParam.color = 0
iViewParam.name = "Вид 1"
iViewParam.scale_ = 1
iViewParam.state = 3
iViewParam.x = 692.522917818619
iViewParam.y = 384.72027874684
iDocument2D.ksCreateSheetArbitraryView(iAssociationViewParam, 0)
kompas_document.SaveAs(r"C:\Users\Slon\Desktop\Для тестирования ПО\3D\14. 675-60-сб2 - Артемьев\675-60-сб113.cdw")
