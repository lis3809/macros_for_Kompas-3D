# -*- coding: utf-8 -*-

title = u'Позиции ОС'

import pythoncom, re, sys, traceback, time, win32api
from win32com.client import Dispatch, gencache
import DispatchOCX_for_PosLeader as dispatchOCX

try:
    import Tkinter as tk
    import tkMessageBox
except:
    import tkinter as tk
    import tkinter.messagebox as tkMessageBox

try:

    def INFO (message_ = '', f_ = 0, memory = 1):

        global message, f

        if memory:
            message, f = message_, f_

        info [ 'text' ] = message_

        if f_:
            info [ 'fg' ]  = '#aa0000'
        else:
            info [ 'fg' ]  = '#000000'

        root.update()

    def Run_KOMPAS():
        KAPI = gencache.EnsureModule('{0422828C-F174-495E-AC5D-D31014DBBE87}', 0, 1, 0)
        iKompasObject = KAPI.KompasObject(Dispatch('Kompas.Application.5')._oleobj_.QueryInterface(KAPI.KompasObject.CLSID, pythoncom.IID_IDispatch))

        KAPI7 = gencache.EnsureModule('{69AC2981-37C0-4379-84FD-5DD2F3C0A520}', 0, 1, 0)
        iApplication = KAPI7.IApplication(Dispatch('Kompas.Application.7')._oleobj_.QueryInterface(KAPI7.IApplication.CLSID, pythoncom.IID_IDispatch))

        return KAPI, iKompasObject, KAPI7, iApplication

############################ Ассоциативная позиция ###################################################
    def Past_pos(premier=0):

        def PosLeader(x0, y0, x, y):
            iPosLeaderParam = KAPI.ksPosLeaderParam(iKompasObject.GetParamStruct(61)) # ko_PosLeaderParam
            iPosLeaderParam.Init()
            iPosLeaderParam.arrowType = 1

            if x < x0:
                iPosLeaderParam.dirX = -1
            else:
                iPosLeaderParam.dirX = 1

            iPosLeaderParam.style = 65535
            iPosLeaderParam.x = x
            iPosLeaderParam.y = y

            iPolylineArray = KAPI.ksDynamicArray(iPosLeaderParam.GetpPolyline())
            iMathPointArray = KAPI.ksDynamicArray(iKompasObject.GetDynamicArray(2))

            iMathPointParam = KAPI.ksMathPointParam(iKompasObject.GetParamStruct(14)) # ko_MathPointParam
            iMathPointParam.Init()
            iMathPointParam.x = x0
            iMathPointParam.y = y0

            iMathPointArray.ksAddArrayItem(-1, iMathPointParam)
            iPolylineArray.ksAddArrayItem(-1, iMathPointArray)
            iPosLeaderParam.SetpPolyline(iPolylineArray)

            return ksDocument2D.ksPositionLeader(iPosLeaderParam)


        def Selection():
            global SelectedObjects_0, obj_sp

            root.update()

            if flag_stop:
                INFO(u'Выполнение прервано пользователем')
                Stop()
                iSelectionManager.UnselectAll()
                ksDocument2D.ksLightObj (iDrawingGroup.Reference, 0)
                return

            SelectedObjects  = iSelectionManager.SelectedObjects

            if  SelectedObjects and SelectedObjects_0 != SelectedObjects and not isinstance(SelectedObjects, tuple):
                if iKompasObject.ksIsKompasCommandCheck(10162):
                    iApplication.StopCurrentProcess(0, iDocument)
                obj_sp = ksSpecification.ksGetSpcObjForGeom (LayoutName, StyleID, KAPI7.IDrawingObject(SelectedObjects).Reference, 0, 1)

                if  obj_sp:
                    iSpecificationBaseObjects = iSpecificationDescription.BaseObjects
                    iSpecificationBaseObject = iSpecificationBaseObjects.Item(obj_sp)
                    Geometry = iSpecificationBaseObject.Geometry
                    ksDocument2D.ksLightObj (iDrawingGroup.Reference, 0)
                    iDrawingGroup.Open()
                    iDrawingGroup.Clear(True)
                    iDrawingGroup.AddObjects(Geometry)
                    iDrawingGroup.Close()
                    ksDocument2D.ksLightObj (iDrawingGroup.Reference, 1)
                    iView = iViews.ViewByNumber(ksDocument2D.ksGetViewNumber(KAPI7.IDrawingObject(SelectedObjects).Reference))
                    iView.Current = True
                    iView.Update()
                    INFO (u'Подтвердите выбор на Enter или выберите другую деталь')
                    SelectedObjects_0 = SelectedObjects
                else:
                    INFO(u'Указанный примитив не связан с ОС!', 1)
                    iSelectionManager.UnselectAll()
                    ksDocument2D.ksLightObj (iDrawingGroup.Reference, 0)
                    iKompasObject.ksExecuteKompasCommand(10162, 1)
                    SelectedObjects_0 = None
                    Selection()


            if isinstance(SelectedObjects, tuple):
                iSelectionManager.UnselectAll()

            if obj_sp and win32api.GetAsyncKeyState(0x0D) != 0:
                INFO (u'Укажите точку, на которую указывает линия выноска')
                iViewsAndLayersManager = iDocument2D.ViewsAndLayersManager

                if SelectedObjects:

                    if not isinstance(SelectedObjects, tuple):
                        iView = iViews.ViewByNumber(ksDocument2D.ksGetViewNumber(KAPI7.IDrawingObject(SelectedObjects).Reference))
                        iView.Current = True
                        iView.Update()

                requestInfo = iKompasObject.GetParamStruct(10)
                requestInfo.Init()
                requestInfo.prompt = u'Укажите точку, на которую указывает линия-выноска'

                koord = ksDocument2D.ksCursor (requestInfo, 0.0, 0.0, None)

                if koord[0]:
                    requestInfo.prompt = u'Укажите точку начала полки'
                    requestInfo.dynamic = 1

                    phantom = iKompasObject.GetParamStruct(6)
                    phantom.Init()
                    phantom.phantom = 1

                    type1 = phantom.GetPhantomParam()
                    type1.Init()
                    type1.xBase = 0.0
                    type1.yBase = 0.0
                    type1.scale_ = 1.0

                    dispatchOCX.x0, dispatchOCX.y0 = koord[1], koord[2]
                    type1.gr = ksDocument2D.ksNewGroup(1)
                    dispatchOCX.PosLeader(ksDocument2D, iKompasObject, 0,0, 0,0)
                    ksDocument2D.ksEndGroup()

                    ocx = Dispatch("Python.DispatchOCX")
                    ocx.Init(iKompasObject)
                    requestInfo.SetCallBackC(u"CallBackC", 0, ocx)
                    koord2 = ksDocument2D.ksGetCursorPosition(0, 0, 1)
                    ksDocument2D.ksCursor (requestInfo, koord2[1], koord2[2], phantom)
                    koord2 = ksDocument2D.ksGetCursorPosition(0, 0, 1)

                    if koord2[0]:
                        iSelectionManager.UnselectAll()
                        ksDocument2D.ksLightObj (iDrawingGroup.Reference, 0)########################################
                        obj_pos = PosLeader( koord[1], koord [2], koord2[1], koord2 [2])
                        ksSpecification.ksSpcObjectEdit(obj_sp)
                        ksSpecification.ksSpcIncludeReference(obj_pos, 0)
                        ksSpecification.ksSpcObjectEnd()
                        Past_pos()


                INFO(u'Выполнение прервано пользователем')
                Stop()
                iSelectionManager.UnselectAll()
                ksDocument2D.ksLightObj (iDrawingGroup.Reference, 0)

            else:
                time.sleep(0.25)
                Selection()


        #### Начало ####
        global iViews, flag_stop, obj_sp, KAPI, iKompasObject, KAPI7, iApplication, iSelectionManager, iDocument, ksDocument2D, iDocument2D, SelectedObjects_0

        obj_sp = 0
        if premier:
            flag_stop = 0

        try:
            pythoncom.connect('Kompas.Application.5')                           # Проверка запущен ли КОМПАС
            KAPI, iKompasObject, KAPI7, iApplication = Run_KOMPAS()
        except:
            INFO (u'КОМПАС-3D не запущен!', 1)
            return

        iDocument = iApplication.ActiveDocument

        if iDocument:

            if iDocument.DocumentType in [1, 2]:
                ksDocument2D = iKompasObject.ActiveDocument2D()
                iDocument2D = KAPI7.IKompasDocument2D(iDocument)
                iDocument2D1 = KAPI7.IKompasDocument2D1(iDocument2D)
                iViewsAndLayersManager = iDocument2D.ViewsAndLayersManager
                iViews = iViewsAndLayersManager.Views

                iSpecificationDescriptions = iDocument.SpecificationDescriptions
                iSpecificationDescription = iSpecificationDescriptions.Active

                if iSpecificationDescription:
                    LayoutName = iSpecificationDescription.LayoutName
                    StyleID = iSpecificationDescription.StyleID
                    ksSpecification = ksDocument2D.GetSpecification()

                    iSelectionManager = iDocument2D1.SelectionManager
                    iSelectionManager.UnselectAll()
                    INFO (u'Укажите геометрический примитив')
                    but_pos.pack_forget()
                    but_stop_pos.pack()

                    iKompasObject.ksExecuteKompasCommand(10162, 1)
                    SelectedObjects_0 = None
                    iDrawingGroups = iDocument2D1.DrawingGroups
                    iDrawingGroup = iDrawingGroups.Add (True, 'Ligth_group')
                    Selection()

                else:
                    INFO(u'В активном документе нет описания спецификации!', 1)
                    Stop()
            else:
                INFO(u'Нет активного чертежа или фрагмента! 666', 1)
                Stop()
        else:
            INFO(u'Нет активного чертежа или фрагмента! 777', 1)
            Stop()

    def Stop():
        iApplication.StopCurrentProcess(0, iDocument)
        but_stop_pos.pack_forget()
        but_pos.pack()
        root.update()

    def Flag_Stop():
        global flag_stop
        flag_stop = 1

    root = tk.Tk()                                                              # создаём окно
    root.title(title)                                                           # заголовок окна
    screen_size_X = root.winfo_screenwidth()                                    # получаем ширину экрана
    screen_size_Y = root.winfo_screenheight()                                   # получаем высоту экрана
    root.geometry('+%d+%d' %(screen_size_X/2, screen_size_Y/2))
    root.resizable(width = False, height = False)
    root.wm_attributes('-topmost', 1)
    root.focus_force()
    W, E, N, S, END = tk.W, tk.E, tk.N, tk.S, tk.END

    message, f = '', 0

    ####    FRAME №2    ###########
    FRAME2 = tk.Frame(root, relief='ridge', padx=4, pady=4)
    FRAME2.grid(row=1, column = 0, sticky = E+W)
    # Поставить позицию
    but_pos = tk.Button(FRAME2, text = 'Создать обозначение позиции', width = 35, takefocus = False, bd=2, command = lambda x=1: Past_pos(x), bg = '#7777ff')
    but_pos.pack()
    but_pos.config(cursor='hand2')
    # Отменить создание позиции
    but_stop_pos = tk.Button(FRAME2, text = 'Отмена', width = 35, command = Flag_Stop, takefocus = False, bd=2, bg = '#7777ff')
    but_stop_pos.pack_forget()
    but_stop_pos.config(cursor='hand2')

    ####    FRAME №3    ############
    FRAME3 = tk.Frame(root, bd = 4)
    FRAME3.grid(row=2, column = 0, sticky = W)
    # строка состояния
    info = tk.Label(FRAME3)
    info.grid(row=0, column = 0)
    ##############################

    root.mainloop()

except:
    root = tk.Tk()
    root.withdraw()
    tkMessageBox.showwarning(title, traceback.format_exc())    # показываем окно с выводом ошибки
    root.destroy()
