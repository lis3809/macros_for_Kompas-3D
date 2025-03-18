# -*- coding: utf-8 -*-
#|save_PNG_from_3D

import pythoncom
from win32com.client import Dispatch, gencache
import LDefin2D

#  Подключим константы API Компас
kompas6_constants = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants
kompas6_constants_3d = gencache.EnsureModule("{2CAF168C-7961-4B90-9DA2-701419BEEFE3}", 0, 1, 0).constants

#  Подключим описание интерфейсов API5
kompas6_api5_module = gencache.EnsureModule("{0422828C-F174-495E-AC5D-D31014DBBE87}", 0, 1, 0)
kompas_object = kompas6_api5_module.KompasObject(Dispatch("Kompas.Application.5")._oleobj_.QueryInterface(kompas6_api5_module.KompasObject.CLSID, pythoncom.IID_IDispatch))


#  Подключим описание интерфейсов API7
kompas_api7_module = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
application = kompas_api7_module.IApplication(Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(kompas_api7_module.IApplication.CLSID, pythoncom.IID_IDispatch))


doc_path = "C:\\Users\\Slon\\Desktop\\Для тестирования ПО\\3D\\15. 675-66-сб6 - Артемьев\\765-66-295.m3d"

Documents = application.Documents
#  Открываем документ
kompas_document = Documents.Open(doc_path, True, False)

kompas_document_3d = kompas_api7_module.IKompasDocument3D(kompas_document)
iDocument3D = kompas_object.ActiveDocument3D()

kompas_document_2d = kompas_api7_module.IKompasDocument2D(kompas_document)
iDocument2D = kompas_object.ActiveDocument2D()

'''Структура параметров ассоциативного вида'''
iAssociationViewParam = kompas6_api5_module.ksAssociationViewParam(kompas_object.GetParamStruct(kompas6_constants.ko_AssociationViewParam))
iAssociationViewParam.Init()
#Разнести
iAssociationViewParam.disassembly = True
iAssociationViewParam.fileName = doc_path
#Признак отрисовки невидимых линий
iAssociationViewParam.hiddenLinesShow = False
iAssociationViewParam.hiddenLinesStyle = 4
#Признак проецирования тел
iAssociationViewParam.projBodies = True
#Проекционная связь
iAssociationViewParam.projectionLink = False
#Имя проекции (из списка проекций в документе-источнике)
iAssociationViewParam.projectionName = "#Изометрия"
#Признак проецирования поверхностей
iAssociationViewParam.projSurfaces = True
#Признак проецирования резьбы
iAssociationViewParam.projThreads = True
#Одинаковая штриховка всех деталей сборки
iAssociationViewParam.sameHatch = False
#Признак разрез/сечение
iAssociationViewParam.section = False
#Признак отрисовки видимых линий перехода
iAssociationViewParam.tangentEdgesShow = True
iAssociationViewParam.tangentEdgesStyle = 2
iAssociationViewParam.visibleLinesStyle = 1

'''Структура параметров вида'''
iViewParam = kompas6_api5_module.ksViewParam(iAssociationViewParam.GetViewParam())
iViewParam.Init()
#угол поворота вида
iViewParam.angle = 0
#цвет вида в активном состоянии
iViewParam.color = 0
iViewParam.name = "Вид 1"
iViewParam.scale_ = 1
#состояние вида (stACTIVE 0  - активный (видимый фоновый)
# stREADONLY 1 - фоновый
# stINVISIBLE 2 - невидимый (погашенный)
# stCURRENT 3 - текущий
iViewParam.state = 3

#точка привязки вида
iViewParam.x = 0
iViewParam.y = 0

#Создать произвольный ассоциативный вид
iDocument2D.ksCreateSheetArbitraryView(iAssociationViewParam, 0)



kompas_document.SaveAs(doc_path.replace(".m3d", ".cdw"))
kompas_document.Save()
