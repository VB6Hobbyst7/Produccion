Attribute VB_Name = "gFunLogistica"
'**************************************************************************
'** Nombre : gFunGeneral
'** Descripción : Módulo para manejo de Pagos Proveedores segun ERS062-2013
'** Creación : EJVG, 20131121 11:00:00 AM
'**************************************************************************
Option Explicit

Public Enum LogTipoPagoComprobante
    gPagoCuentaCMAC = 1
    gPagoTransferencia = 2
    gPagoCheque = 3
End Enum
