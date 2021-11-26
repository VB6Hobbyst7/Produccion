Attribute VB_Name = "gFunLogistica"
Option Explicit

Public Enum LogTipoOC
    gLogOCompraDirecta = 130
    gLogOServicioDirecta = 132
    gLogOCompraProceso = 133
    gLogOServicioProceso = 134
End Enum

Global Const gLogOCDirecta = "D"
Global Const gLogOCProceso = "P"
Global gnTipCambioPonderado As Currency

