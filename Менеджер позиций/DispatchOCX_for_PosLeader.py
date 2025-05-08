# -*- coding: utf-8 -*-
#~ Интерфейc, в котором реализована CallBack-функция

import pythoncom
from win32com.client import Dispatch, gencache

module =  gencache.EnsureModule("{0422828C-F174-495E-AC5D-D31014DBBE87}", 0, 1, 0)
x0, y0 = 0.0, 0.0 # переменные, принимающие координаты точки на которую будет указывать линия выноска

def PosLeader(doc, kompasObject, x_, y_, x, y):

    iPosLeaderParam = module.ksPosLeaderParam(kompasObject.GetParamStruct(61)) # ko_PosLeaderParam
    iPosLeaderParam.Init()
    iPosLeaderParam.style = 0
    iPosLeaderParam.arrowType = 1

    if x < x_:
        iPosLeaderParam.dirX = -1
    else:
        iPosLeaderParam.dirX = 1

    iPosLeaderParam.style = 65535
    iPosLeaderParam.x = 0
    iPosLeaderParam.y = 0

    iPolylineArray = module.ksDynamicArray(iPosLeaderParam.GetpPolyline())
    iMathPointArray = module.ksDynamicArray(kompasObject.GetDynamicArray(2))

    iMathPointParam = module.ksMathPointParam(kompasObject.GetParamStruct(14)) # ko_MathPointParam
    iMathPointParam.Init()
    iMathPointParam.x = x_-x
    iMathPointParam.y = y_-y

    iMathPointArray.ksAddArrayItem(-1, iMathPointParam)
    iPolylineArray.ksAddArrayItem(-1, iMathPointArray)
    iPosLeaderParam.SetpPolyline(iPolylineArray)

    doc.ksPositionLeader(iPosLeaderParam)

class DispatchOCX:
    _reg_clsid_ = "{CEAB9979-789E-4979-867F-B21DF4085255}"
    _reg_desc_ = "DispatchOCX COM"
    _reg_progid_ = "Python.DispatchOCX"
    _public_methods_ = [u'CallBackC', u'Init']
    _readonly_attrs_ = ['disp', 'kompasObject', 'doc']

    def __init__(self):
       self.disp = None
       self.kompasObject  = None
       self.doc = None

    def Init(self, disp):
        self.disp = disp
        self.kompasObject = module.KompasObject(self.disp.QueryInterface(module.KompasObject.CLSID, pythoncom.IID_IDispatch))
        self.doc = self.kompasObject.ActiveDocument2D()

    def CallBackC(self, com, x, y, info, phantom, dynamic):
        info = module.ksRequestInfo(info.QueryInterface(module.ksRequestInfo.CLSID, pythoncom.IID_IDispatch))
        phantom = module.ksPhantom(phantom.QueryInterface(module.ksPhantom.CLSID, pythoncom.IID_IDispatch))
        gr = phantom.GetPhantomParam().gr
        if dynamic == 0:
            if com == -1:
                return 0
        else:
            self.doc.ksDeleteObj(gr)
            phantom.GetPhantomParam().gr = self.doc.ksNewGroup(1)
            PosLeader(self.doc,  self.kompasObject, x0, y0, x, y)
            self.doc.ksEndGroup()
            self.doc.ksChangeObjectInLibRequest(info, phantom)
        return 1

if __name__=='__main__':
    import win32com.server.register
    win32com.server.register.UseCommandLine(DispatchOCX)