Attribute VB_Name = "gITF"
Option Explicit

'*** Modulo para el ITF
Global gnITFPorcent As Double
Global gbITFAplica As Boolean
Global gbITFAsumidoAho As Boolean
Global gbITFAsumidoPF As Boolean
Global gbITFAsumidocreditos As Boolean
Global gbITFAsumidoGiros As Boolean

Global gnITFMontoMin As Double
Global gnITFNumTranOrigen As Long
Global gnITFNumTran As Long

'Tipo Exoneracion
Global Const gnITFTpoSinExoneracion As String = 0
Global Const gnITFTpoExoPlanilla As String = 3
Global Const gnITFTpoExoUniColegios As String = 2
Global Const gnITFTpoExoIntPublicas As String = 1
Global Const gnITFTpoExoIntFinanc As String = 6
Global Const gsRUCCmac As String = "20104888934"

Global Const gnITFTpoOpeVarias As String = "1"
Global Const gnITFTpoOpeCaja As String = "2"

Public gTCPonderadoSBS As Currency

'ALPA 20091117***************************************
Public Function fgITFVerificaExoneracionInteger(ByVal psCodCta As String) As Integer
    Dim sql As String
    Dim oCon As COMDConstSistema.FCOMITF
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    Set oCon = New COMDConstSistema.FCOMITF
   
        fgITFVerificaExoneracionInteger = oCon.fgITFVerificaExoneracionInteger(psCodCta)
      
    Set oCon = Nothing
End Function
Public Function fgITFVerificaExoneracion(ByVal psCodCta As String) As Boolean
    Dim sql As String
    Dim oCon As COMDConstSistema.FCOMITF
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    Set oCon = New COMDConstSistema.FCOMITF
   
        fgITFVerificaExoneracion = oCon.fgITFVerificaExoneracion(psCodCta)
      
    Set oCon = Nothing
End Function
'***************************************************

'Sin Exoneracion              0
'Planillas                    1
'Colegios/Universidades       2
'Instituciones Publicas       3
'Instituciones Financieras    4
Public Function fgITFTipoExoneracion(ByVal psCodCta As String, _
    Optional ByRef sDescripcion As String = "") As Integer

Dim oCon As COMDConstSistema.FCOMITF
    
Set oCon = New COMDConstSistema.FCOMITF
    fgITFTipoExoneracion = oCon.fgITFTipoExoneracion(psCodCta, sDescripcion)
Set oCon = Nothing
End Function

'*** Obtiene los parametros de ITF
Public Sub fgITFParametros(Optional ByVal pRs As ADODB.Recordset = Nothing)

Dim oCon As COMDConstSistema.FCOMITF
Dim lr As New ADODB.Recordset
Set oCon = New COMDConstSistema.FCOMITF

'En el caso que se envie un Recordset previamente cargado
If Not pRs Is Nothing Then
    Do While Not pRs.EOF
        Select Case pRs!nParCod
            Case 1001
                gbITFAplica = IIf(pRs!nParValor = 0, False, True)
            Case 1003
                gnITFPorcent = pRs!nParValor
            Case 1002
                gnITFMontoMin = pRs!nParValor
            End Select
        pRs.MoveNext
    Loop
    Exit Sub
End If

Set lr = oCon.fgITFParametroForm
Set oCon = Nothing

    If lr Is Nothing Then
    Else
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
    End If
    lr.Close
    Set lr = Nothing
End Sub

'*** Obtiene los parametros de ITF
Public Sub fgITFParamAsume(psAgecod As String, Optional psProducto As String)
    Dim oCon As COMDConstSistema.FCOMITF
    Set oCon = New COMDConstSistema.FCOMITF
    Dim lr As New ADODB.Recordset
    Set lr = New ADODB.Recordset
    Set lr = oCon.fgITFParamAsumeForm(psAgecod, psProducto)
    Set oCon = Nothing
    If lr Is Nothing Then
    Else
        Do While Not lr.EOF
            Select Case lr!cProducto
            Case gCapAhorros
                gbITFAsumidoAho = lr!bAsumido
            Case gCapPlazoFijo
                gbITFAsumidoPF = lr!bAsumido
            Case gGiro
                gbITFAsumidoGiros = lr!bAsumido
            Case Else
               gbITFAsumidocreditos = lr!bAsumido
            End Select
            lr.MoveNext
        Loop
     End If
     lr.Close
     
End Sub

'*** Calcula el impuesto ITF de Transaccion
Public Function fgITFCalculaImpuesto(ByVal pnMonto As Double) As Double
Dim lnValor As Double
lnValor = pnMonto
If gbITFAplica = True Then
    If pnMonto > gnITFMontoMin Then
        
        lnValor = pnMonto * gnITFPorcent
        
        Dim aux As Double
        If InStr(1, CStr(lnValor), ".", vbTextCompare) > 0 Then
            aux = CDbl(CStr(Int(lnValor)) & "." & Mid(CStr(lnValor), InStr(1, CStr(lnValor), ".", vbTextCompare) + 1, 2))
        Else
            'By capi 30032009 se modifico porque calculaba mal el itf solicitud verbal jefatura encargada
            'aux = CDbl(CStr(Int(lnValor)))
            aux = CDbl(CStr(lnValor))
            '
        End If
        lnValor = aux

        lnValor = fgTruncar(lnValor, 2)
               
    End If
End If
fgITFCalculaImpuesto = lnValor
End Function

Public Function fgITFDesembolso(ByVal pnMonto As Double) As Double
    Dim sCadena As Currency
        sCadena = Round(pnMonto * gnITFPorcent, 6)
        fgITFDesembolso = CortaDosITF(sCadena)
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
 '*** BRGO 20110907 Redondea Valor ITF a favor del cliente
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
'*** END BRGO
'*** Calcula el impuesto ITF de Transaccion
Public Function fgITFCalculaImpuestoIncluido(ByVal pnMonto As Double, Optional ByVal bCancelacion As Boolean = False) As Double
Dim lnValor As Double

Dim nMontoITF As Double 'ARCV 20-06-2007

lnValor = pnMonto
If gbITFAplica = True Then
'ARCV 20-06-2007
    If pnMonto > gnITFMontoMin Then
        If bCancelacion = True Then
            lnValor = fgTruncar((pnMonto / (1 + gnITFPorcent)), 2)
            
            Dim aux As Double
            If InStr(1, CStr(lnValor), ".", vbTextCompare) <> 0 Then
               aux = CDbl(CStr(Int(lnValor)) & "." & Mid(CStr(lnValor), InStr(1, CStr(lnValor), ".", vbTextCompare) + 1, 2))
               lnValor = aux
            End If

            lnValor = fgTruncar(lnValor, 2)
        
        Else
            'lnValor = (pnMonto / (1 + gnITFPorcent))
            nMontoITF = fgITFCalculaImpuesto(pnMonto)
            lnValor = pnMonto - nMontoITF
        End If

    End If

End If
fgITFCalculaImpuestoIncluido = lnValor
End Function

'*** Devuelve el Monto con el ITF agregado
Public Function fgITFCalculaImpuestoNOIncluido(ByVal pnMonto As Double, Optional ByVal bCancelacion As Boolean) As Double
Dim lnValor As Double
lnValor = pnMonto
If gbITFAplica = True Then
        If bCancelacion = True Then
            lnValor = fgTruncar((pnMonto * (1 + gnITFPorcent)), 2)
        Else
            lnValor = (pnMonto * (1 + gnITFPorcent))
        End If
        
        Dim aux As Double
        If bCancelacion = True Then
            If InStr(1, CStr(lnValor), ".", vbTextCompare) <> 0 Then
                aux = CDbl(CStr(Int(lnValor)) & "." & Mid(CStr(lnValor), InStr(1, CStr(lnValor), ".", vbTextCompare) + 1, 2))
            Else
                aux = lnValor
            End If
        Else
            aux = CDbl(CStr(Int(lnValor)) & "." & Mid(CStr(lnValor), InStr(1, CStr(lnValor), ".", vbTextCompare) + 1, 2))
        End If
        lnValor = aux
        lnValor = fgTruncar(lnValor, 2)

End If
fgITFCalculaImpuestoNOIncluido = lnValor
End Function

Public Function fgITFGetTitular(psctacod As String) As String
    Dim oCon As COMDConstSistema.FCOMITF
    Set oCon = New COMDConstSistema.FCOMITF
       fgITFGetTitular = oCon.fgITFGetTitular(psctacod)
    Set oCon = Nothing
End Function

Public Function fgITFGetNumtranOrigen(oConexion As ADODB.Connection) As Long
    Dim oCon As COMDConstSistema.FCOMITF
    Set oCon = New COMDConstSistema.FCOMITF
        fgITFGetNumtranOrigen = oCon.fgITFGetNumtranOrigen(oConexion)
    Set oCon = Nothing
    
End Function

Public Function VerifOpeVariasAfectaITF(psOpeCod As String) As Boolean
    Dim oCon As COMDConstSistema.FCOMITF
    Set oCon = New COMDConstSistema.FCOMITF
        VerifOpeVariasAfectaITF = oCon.VerifOpeVariasAfectaITF(psOpeCod)
    Set oCon = Nothing
    
End Function

Public Function fgTruncar(pnNumero As Double, pnDecimales As Integer) As String

    Dim i As Integer
    Dim sEnt As String
    Dim sDec As String
    Dim sNum As String
    Dim sPunto As String
    Dim sResul As String
    
    sNum = Trim(Str(pnNumero))
    sDec = ""
    sPunto = ""
    sEnt = ""
    For i = 1 To Len(Trim(sNum))
        If Mid(sNum, i, 1) = "." Then
            sPunto = "."
        Else
            If sPunto = "" Then
                sEnt = sEnt & Mid(sNum, i, 1)
            Else
                sDec = sDec & Mid(sNum, i, 1)
            End If
        End If
    Next i
    If sDec = "" Then
        sDec = "00"
    End If
    sResul = sEnt & "." & Left(sDec, 2)
    fgTruncar = sResul
    
End Function


