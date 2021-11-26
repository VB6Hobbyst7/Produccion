VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPrevio"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'By capi 01102008
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long



Public Sub Show(psCadImpresion As String, psTitulo As String, Optional pbCondensado As Boolean = False, Optional pnLineas As Integer = 66, Optional pImpresora As Impresoras = gEPSON)
    Dim o As clsConstImp
    
'''    'By Capi 01102008
    Dim buffMaq As String
    Dim lSizeMaq As Long
    Dim lsql As String
    Dim loCon As COMConecta.DCOMConecta
    Dim lrs As ADODB.Recordset
    Dim lsMaquina As String

    buffMaq = Space(255)
    lSizeMaq = Len(buffMaq)
    GetComputerName buffMaq, lSizeMaq
    lsMaquina = Trim(Left$(buffMaq, lSizeMaq))


    Set loCon = New COMConecta.DCOMConecta
    Set lrs = New ADODB.Recordset

    lsql = "Select nTpoCar from ImpreSetup Where cPCCod = '" & lsMaquina & "'"

    loCon.AbreConexion
    Set lrs = loCon.CargaRecordSet(lsql)

    If lrs.EOF And lrs.BOF Then
        'GetImpreSetup = -1
        pImpresora = 0
    Else
        pImpresora = lrs!nTpoCar
    End If

    Set loCon = Nothing
    
    'psCadImpresion = Replace(psCadImpresion, Chr(10), oImpresora.gPrnSaltoLinea, vbTextCompare)
    'psCadImpresion = Replace(psCadImpresion, Chr(12), oImpresora.gPrnSaltoPagina, vbTextCompare)

'''  'End By
    oImpresora.Inicia pImpresora
    lTpoImpresora = pImpresora
    Inicia pImpresora
    frmPrevio.Previo psCadImpresion, psTitulo, pbCondensado, pnLineas
End Sub

Public Sub ShowImpreSpool(ByVal psCadImpresion As String, Optional pbCondensado As Boolean = False, Optional pnLineas As Integer = 66, Optional pImpresora As Impresoras = gEPSON)
    Dim i As Long
    frmImpresora.Show 1
    If Not lbCancela Then
        For i = 1 To lnNumCopias
            ImpreBegin pbCondensado, pnLineas
                Print #ArcSal, ImpreCarEsp(psCadImpresion)
            ImpreEnd
        Next
    End If
End Sub

Public Sub PrintSpool(ByVal psLpt As String, ByVal psCadImpresion As String, Optional pbCondensado As Boolean = False, Optional pnLineas As Integer = 66, Optional pImpresora As Impresoras = gEPSON)
    sLpt = psLpt
    ImpreBegin pbCondensado, pnLineas
        Print #ArcSal, ImpreCarEsp(psCadImpresion)
    ImpreEnd
End Sub


