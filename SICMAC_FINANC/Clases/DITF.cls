VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DITF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'*** Modulo para el ITF
Public gnITFPorcent As Double
Public gbITFAplica As Boolean
Public gbITFAsumidoAho As Boolean
Public gbITFAsumidoPF As Boolean
Public gbITFAsumidocreditos As Boolean
Public gbITFAsumidoGiros As Boolean


Public gnITFMontoMin As Double
Public gnITFNumTranOrigen As Long
Public gnITFNumTran As Long

Public gTCPonderadoSBS As Currency

'*** Obtiene los parametros de ITF
Public Function fgITFParametros()
Dim oCon As DConecta
Set oCon = New DConecta
Dim lsSQL As String
Dim lr As ADODB.Recordset
Set lr = New ADODB.Recordset
    
    lsSQL = "select nParCod, nParValor FROM PARAMETRO WHERE nParProd = 1000 And nParCod In (1001,1002,1003)"
    oCon.AbreConexion
    Set lr = oCon.CargaRecordSet(lsSQL)
    
    Do While Not lr.EOF
        Select Case lr!nParCod
            Case 1001
                gbITFAplica = IIf(lr!nParValor = 0, False, True)
            Case 1003
                gnITFPorcent = lr!nParValor
            Case 1002
                gnITFMontoMin = lr!nParValor
        End Select
        lr.MoveNext
    Loop
    lr.Close
    Set lr = Nothing

    oCon.CierraConexion
    Set oCon = Nothing
End Function
Public Function CortaDosITF(ByVal lnITF As Double) As Double
Dim intpos  As Integer
Dim lnDecimal As Double
Dim lsDec As String
Dim lnEntero As Long
Dim lnPos As Long

lnEntero = Int(lnITF)
lnDecimal = Round(lnITF - Int(lnEntero), 6)
lnPos = InStr(1, Trim(Str(lnDecimal)), ".")
If lnPos > 0 Then
    lsDec = Mid(Trim(Str(lnDecimal)), lnPos + 1, 2)
    lsDec = IIf(Len(lsDec) = 1, lsDec * 10, lsDec)
    lnDecimal = Val(lsDec) / 100
    CortaDosITF = lnEntero + lnDecimal
Else
    lnDecimal = 0
    CortaDosITF = lnEntero
End If
End Function
Public Function fgDiferenciaRedondeoITF(ByVal lnITF As Double) As Double
    Dim lnPos As Integer
    Dim DifITF As Double

    lnPos = InStr(1, Trim(Str(lnITF)), ".")
    If lnPos > 0 Then
        DifITF = Round((lnITF * 100 Mod 10) / 100, 2)
        DifITF = IIf(DifITF = 0.05, 0, IIf(DifITF > 0.05, DifITF - 0.05, DifITF))
    Else
        DifITF = 0
    End If
    fgDiferenciaRedondeoITF = Round(DifITF, 2)
End Function
