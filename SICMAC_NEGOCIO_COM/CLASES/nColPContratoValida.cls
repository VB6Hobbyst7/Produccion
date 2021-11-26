VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "nColPContratoValida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function nValidaRescateCredPignoraticio(ByVal psCtaCod As String, ByVal pnEstado As ColocEstado, _
            ByVal psLote As String, ByVal psMovNro As String, Optional pbEjecBatch As Boolean = False) As String

'**
'************************************

'Dim lsSQL As String
'Dim lsMensaje As String
'
''    If pnEstado <> gPigEstCanc Then
''        lsmensaje = " Contrato se encuentra " & fgEstadoCredPigDesc(AXDesCon.CodEstadoCred), vbInformation, " Aviso "
''        Limpiar
''        AXCodCta.SetFocusCuenta
''        Exit Function
''    End If
'
'    If IsCtaBlo(vNroContrato, vConexion) Then
'        MsgBox " Contrato se encuentra Bloqueado ", vbInformation, " Aviso "
'        RegCredPrend.Close
'        Set RegCredPrend = Nothing
'        Limpiar
'        AXCodCta.SetFocusNum
'        Exit Function
'    End If
'    If RegCredPrend!cAgeBoveda <> gsCodAge Then  ' Localizacion de Joya
'        MsgBox " Garantía Prendaria se encuentra en Boveda de Agencia " & Mid(RegCredPrend!cAgeBoveda, 4, 2), vbInformation, " Aviso "
'        RegCredPrend.Close
'        Set RegCredPrend = Nothing
'        Limpiar
'        AXCodCta.SetFocusNum
'        Exit Function
'    End If
End Function
