VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FCOMITF"
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
'ALPA 20091125**************************************
Public Function fgITFVerificaExoneracionInteger(ByVal psCodCta As String) As Integer
    Dim Sql As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    oCon.AbreConexion
    
    Sql = "select nExoTpo from dbo.ITFExoneracionCta Where cCtaCod = '" & psCodCta & "' And nExoTpo <> 0"
    Set rs = oCon.CargaRecordSet(Sql)
    
    If Not (rs.EOF And rs.EOF) Then
        fgITFVerificaExoneracionInteger = rs!nExoTpo
    Else
        fgITFVerificaExoneracionInteger = 0
    End If
    
    oCon.CierraConexion
    Set oCon = Nothing
    rs.Close
    Set rs = Nothing
End Function
'***************************************************
Public Function fgITFVerificaExoneracion(ByVal psCodCta As String) As Boolean
    Dim Sql As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    oCon.AbreConexion
    
    Sql = "select nExoTpo from dbo.ITFExoneracionCta Where cCtaCod = '" & psCodCta & "' And nExoTpo <> 0"
    Set rs = oCon.CargaRecordSet(Sql)
    
    If Not (rs.EOF And rs.EOF) Then
        fgITFVerificaExoneracion = True
    Else
        fgITFVerificaExoneracion = False
    End If
    
    oCon.CierraConexion
    Set oCon = Nothing
    rs.Close
    Set rs = Nothing
End Function


'Sin Exoneracion              0
'Planillas                    1
'Colegios/Universidades       2
'Instituciones Publicas       3
'Instituciones Financieras    4
Public Function fgITFTipoExoneracion(ByVal psCodCta As String, _
    Optional ByRef sDescripcion As String = "") As Integer
    
Dim sSql As String
Dim oCon As COMConecta.DCOMConecta
Set oCon = New COMConecta.DCOMConecta
Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    oCon.AbreConexion
    
    sSql = "select I.nExoTpo, C.cConsDescripcion from dbo.ITFExoneracionCta I JOIN Constante C ON " _
        & "I.nExoTpo = C.nConsCod Where I.cCtaCod = '" & psCodCta & "' And C.nConsCod = 1044"
    Set rs = oCon.CargaRecordSet(sSql)
    
    If Not (rs.EOF And rs.EOF) Then
        fgITFTipoExoneracion = rs("nExoTpo")
        sDescripcion = rs("cConsDescripcion")
    Else
        fgITFTipoExoneracion = 0
    End If
    
    oCon.CierraConexion
    Set oCon = Nothing
    rs.Close
    Set rs = Nothing
End Function

'*** Obtiene los parametros de ITF
Public Function fgITFParametros()
Dim oCon As COMConecta.DCOMConecta
Set oCon = New COMConecta.DCOMConecta
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

Public Function fgITFParametroForm() As ADODB.Recordset
Dim oCon As COMConecta.DCOMConecta
Set oCon = New COMConecta.DCOMConecta
Dim lsSQL As String
Dim lr As ADODB.Recordset
Set lr = New ADODB.Recordset
    
    lsSQL = "select nParCod, nParValor FROM PARAMETRO WHERE nParProd = 1000 And nParCod In (1001,1002,1003)"
    oCon.AbreConexion
    Set lr = oCon.CargaRecordSet(lsSQL)
    If Not (lr.EOF And lr.BOF) Then
        Set fgITFParametroForm = lr
    Else
        Set fgITFParametroForm = Nothing
    End If
    oCon.CierraConexion
    Set oCon = Nothing

End Function


'*** Obtiene los parametros de ITF
Public Sub fgITFParamAsume(psAgeCod As String, Optional psProducto As String)
Dim oCon As COMConecta.DCOMConecta
Set oCon = New COMConecta.DCOMConecta
Dim lsSQL As String
Dim lr As ADODB.Recordset
Set lr = New ADODB.Recordset

oCon.AbreConexion
lsSQL = " Select cProducto, bAsumido from itfagenciaproducto where cAgeCod = '" & psAgeCod & "' and cproducto='" & psProducto & "'"
Set lr = oCon.CargaRecordSet(lsSQL)

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
oCon.CierraConexion
Set oCon = Nothing
lr.Close
Set lr = Nothing
End Sub

Public Function fgITFParamAsumeForm(psAgeCod As String, Optional psProducto As String) As ADODB.Recordset
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    Dim lsSQL As String
    Dim lr As ADODB.Recordset
    Set lr = New ADODB.Recordset

    oCon.AbreConexion
    lsSQL = " Select cProducto, bAsumido from itfagenciaproducto where cAgeCod = '" & psAgeCod & "' and cproducto='" & psProducto & "'"
    Set lr = oCon.CargaRecordSet(lsSQL)

    Set fgITFParamAsumeForm = lr
    
    oCon.CierraConexion
    
End Function

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
            'ARCV 08-06-2006
            'aux = CDbl(lnValor)
        Else
            aux = CDbl(CStr(Int(lnValor)))
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
    'ARCV 26-10-2006
    lsDec = Mid(Trim(Str(lnDecimal)), lnPos + 1, 2)
    lsDec = IIf(Len(lsDec) = 1, lsDec * 10, lsDec)
    lnDecimal = Val(lsDec) / 100
    'ARCV 08-06-2006
    CortaDosITF = lnEntero + lnDecimal
Else
    lnDecimal = 0
    CortaDosITF = lnEntero
End If
End Function
 '*** BRGO 20110907 Redondea Valor ITF a favor del cliente
Public Function DiferenciaRedondeoITF(ByVal lnITF As Double) As Double
    Dim lnPos As Integer
    Dim DifITF As Double

    lnPos = InStr(1, Trim(Str(lnITF)), ".")
    If lnPos > 0 Then
        DifITF = Round((lnITF * 10 Mod 1) / 10, 2)
        DifITF = IIf(DifITF = 0.05, 0, IIf(DifITF > 0.05, DifITF - 0.05, DifITF))
    Else
        DifITF = 0
    End If
    DiferenciaRedondeoITF = Round(lnITF - DifITF, 2)
End Function
'*** END BRGO

'*** Calcula el impuesto ITF de Transaccion
Public Function fgITFCalculaImpuestoIncluido(ByVal pnMonto As Double, Optional ByVal bCancelacion As Boolean = False) As Double
Dim lnValor As Double
lnValor = pnMonto
If gbITFAplica = True Then
    If pnMonto > gnITFMontoMin Then
        If bCancelacion = True Then
            lnValor = fgTruncar(pnMonto / (1 + gnITFPorcent), 2)
        Else
            lnValor = pnMonto / (1 + gnITFPorcent)
        End If
        
        Dim aux As Double
        If InStr(1, CStr(lnValor), ".", vbTextCompare) <> 0 Then
           aux = CDbl(CStr(Int(lnValor)) & "." & Mid(CStr(lnValor), InStr(1, CStr(lnValor), ".", vbTextCompare) + 1, 2))
           'ARCV 08-06-2006
           'aux = CDbl(lnValor)
            lnValor = aux
        End If
        
        If bCancelacion = False Then
            lnValor = fgTruncar(lnValor, 2)
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
            lnValor = fgTruncar(pnMonto * (1 + gnITFPorcent), 2)
        Else
            lnValor = pnMonto * (1 + gnITFPorcent)
        End If

        Dim aux As Double
        If bCancelacion = True Then
            If InStr(1, CStr(lnValor), ".", vbTextCompare) <> 0 Then
                aux = CDbl(CStr(Int(lnValor)) & "." & Mid(CStr(lnValor), InStr(1, CStr(lnValor), ".", vbTextCompare) + 1, 2))
                'ARCV 08-06-2006
                'aux = CDbl(lnValor)
            Else
                aux = lnValor
            End If
        Else
            aux = CDbl(CStr(Int(lnValor)) & "." & Mid(CStr(lnValor), InStr(1, CStr(lnValor), ".", vbTextCompare) + 1, 2))
            'ARCV 08-06-2006
            'aux = CDbl(lnValor)
        End If
        
        lnValor = aux
        lnValor = fgTruncar(lnValor, 2)
End If
fgITFCalculaImpuestoNOIncluido = lnValor
End Function


'*** Retiro de Cuenta Ahorros ITF
Public Function fgITFACRetiroImpuesto(lpsCodCta As String, lpnMonto As Currency, lNroDoc As String, lpgsACRetEfe As String, _
                                lpdFecha As String, lpsCodUser As String, oConexion As ADODB.Connection, _
                                Optional ByVal psCodAge As String) As Long
    
    Dim cmd As ADODB.Command
    Dim prm As ADODB.Parameter
    Dim lpnNroTransaccion As Long
    Dim lpsUsuRem As String
    Dim lpbGrabaTranDiaria As Integer
    Dim lbValidaMinimo As Integer
    Dim lpgsCodAge As String
   
    lbValidaMinimo = 1
    lpbGrabaTranDiaria = 1
    
    'Consolida la Información Para el FSD
    Set cmd = New ADODB.Command
    cmd.CommandText = "spACRetiroEfectivo"
    cmd.CommandType = adCmdStoredProc
    cmd.Name = "spACRetiroEfectivo"
    Set prm = cmd.CreateParameter("psCodCta", adVarChar, adParamInput, 12)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("pnMonto", adCurrency, adParamInput)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("NroDoc", adVarChar, adParamInput, 20)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("pgsACRetEfe", adVarChar, adParamInput, 6)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("pgsCodAge", adVarChar, adParamInput, 5)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("pdFecha", adVarChar, adParamInput, 24)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("psCodUser", adVarChar, adParamInput, 4)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("psUsuRem", adVarChar, adParamInput, 4)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("pbGrabaTranDiaria", adInteger, adParamInput)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("bValidaMinimo", adInteger, adParamInput)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("pnNroTransaccion", adInteger, adParamOutput)
    cmd.Parameters.Append prm
    
    Set cmd.ActiveConnection = oConexion
    cmd.CommandTimeout = 720
    cmd.Parameters.Refresh
    oConexion.spACRetiroEfectivo lpsCodCta, lpnMonto, lNroDoc, lpgsACRetEfe, psCodAge, lpdFecha, lpsCodUser, lpsUsuRem, lpbGrabaTranDiaria, lbValidaMinimo, lpnNroTransaccion

    fgITFACRetiroImpuesto = lpnNroTransaccion
    
    Set cmd = Nothing
    Set prm = Nothing
End Function


'*** Efectivo del MontoTotal ITF
Public Function fgITFEfectivoImpuesto(lcCodOpe As String, lcCodCta As String, lcNumDoc As String, lnMonTran As Currency, _
                                     lnSaldCnt As Currency, lcCodUsuRem As String, lcCodAge As String, lnTipCambio As Currency, oConexion As ADODB.Connection, _
                                     pdFecSis As Date, psCodUser As String) As Long
    
    Dim cmd As ADODB.Command
    Dim prm As ADODB.Parameter

    Dim oGen As New COMDConstSistema.DCOMGeneral
    Dim lnTransaccion As Long
    
    'Consolida la Información Para el FSD
    Set cmd = New ADODB.Command
    cmd.CommandText = "InsertTranDiariaOut"
    cmd.CommandType = adCmdStoredProc
    cmd.Name = "InsertTranDiariaOut"
    Set prm = cmd.CreateParameter("dFecTran", adVarChar, adParamInput, 24)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("cCodUsu", adVarChar, adParamInput, 4)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("cCodOpe", adVarChar, adParamInput, 20)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("cCodCta", adVarChar, adParamInput, 12)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("cNumDoc", adVarChar, adParamInput, 20)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("nMonTran", adCurrency, adParamInput)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("nSaldCnt", adCurrency, adParamInput)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("cCodUsuRem", adVarChar, adParamInput, 4)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("cCodAge", adVarChar, adParamInput, 5)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("nTipCambio", adCurrency, adParamInput)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("lnTransaccion", adInteger, adParamOutput)
    cmd.Parameters.Append prm
    
    Set cmd.ActiveConnection = oConexion
    cmd.CommandTimeout = 720
    cmd.Parameters.Refresh
    oConexion.InsertTranDiariaOut oGen.FechaHora(pdFecSis), psCodUser, lcCodOpe, lcCodCta, lcNumDoc, lnMonTran, lnSaldCnt, lcCodUsuRem, lcCodAge, lnTipCambio, lnTransaccion

    fgITFEfectivoImpuesto = lnTransaccion
    
    Set cmd = Nothing
    Set prm = Nothing
End Function

'*** Efectivo del MontoTotal ITF
Public Sub fgITFDetalleInserta(pnNumTran As Long, pcAgeCod As String, pnMonto As Currency, pnImpuesto As Currency, pnNumTranITF As Long, oConexion As ADODB.Connection)
    Dim Sql As String
    Sql = " Insert ITFDetalle(nNumTran, cAgeCod, nMonto, nImpuesto, nNumTranITF)" _
        & " Values (" & pnNumTran & ",'" & pcAgeCod & "'," & pnMonto & "," & pnImpuesto & "," & pnNumTranITF & ")"
    oConexion.Execute Sql
End Sub

Public Function fgITFGetTitular(psCtaCod As String) As String
    Dim Sql As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    oCon.AbreConexion
    Sql = " Select Top 1 dbo.PstaNombre(cPersNombre,1) Nombre From ProductoPersona PP" _
        & " Inner Join Persona PE ON PP.cPersCod = PE.cPersCod" _
        & " Where PP.cCtaCod = '" & psCtaCod & "' Order By cPersNombre"
    Set rs = oCon.CargaRecordSet(Sql)
    
    If rs.EOF And rs.BOF Then
        fgITFGetTitular = ""
    Else
        fgITFGetTitular = rs!NOMBRE
    End If
    
    rs.Close
    Set rs = Nothing
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function fgITFGetNumtranOrigen(oConexion As ADODB.Connection) As Long
    Dim Sql As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    Sql = " SELECT @@IDENTITY "
    rs.Open Sql, oConexion, adOpenStatic, adLockReadOnly, adCmdText
    
    If rs.EOF And rs.BOF Then
        fgITFGetNumtranOrigen = 0
    Else
        fgITFGetNumtranOrigen = rs.Fields(0)
    End If
    
End Function


'*** Retiro de Cuenta Ahorros ITF
Public Function fgITFACDepositoImpuesto(lpsCodCta As String, lpnMonto As Currency, lNroDoc As String, lpgsACRetEfe As String, _
                                lpdFecha As String, lpsCodUser As String, oConexion As ADODB.Connection, psCodAge As String) As Long
    
    Dim cmd As ADODB.Command
    Dim prm As ADODB.Parameter
    Dim lpnNroTransaccion As Long
    Dim lpsUsuRem As String
    Dim lpbGrabaTranDiaria As Integer
    Dim lbValidaMinimo As Integer
    Dim lpgsCodAge As String
   
    lbValidaMinimo = 1
    lpbGrabaTranDiaria = 1
    
    'Consolida la Información Para el FSD
    Set cmd = New ADODB.Command
    cmd.CommandText = "spACDepositoEfectivo"
    cmd.CommandType = adCmdStoredProc
    cmd.Name = "spACDepositoEfectivo"
    Set prm = cmd.CreateParameter("psCodCta", adVarChar, adParamInput, 12)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("pnMonto", adCurrency, adParamInput)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("NroDoc", adVarChar, adParamInput, 20)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("pgsACRetEfe", adVarChar, adParamInput, 6)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("pgsCodAge", adVarChar, adParamInput, 5)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("pdFecha", adVarChar, adParamInput, 24)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("psCodUser", adVarChar, adParamInput, 4)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("psUsuRem", adVarChar, adParamInput, 4)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("pbGrabaTranDiaria", adInteger, adParamInput)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("bValidaMinimo", adInteger, adParamInput)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("pnNroTransaccion", adInteger, adParamOutput)
    cmd.Parameters.Append prm
    
    Set cmd.ActiveConnection = oConexion
    cmd.CommandTimeout = 720
    cmd.Parameters.Refresh
    oConexion.spACDepositoEfectivo lpsCodCta, lpnMonto, lNroDoc, lpgsACRetEfe, psCodAge, lpdFecha, lpsCodUser, lpsUsuRem, lpbGrabaTranDiaria, lbValidaMinimo, lpnNroTransaccion

    fgITFACDepositoImpuesto = lpnNroTransaccion
    
    Set cmd = Nothing
    Set prm = Nothing
End Function


Public Function GetnNumtranRemoto(oConexion As ADODB.Connection, psFecha As String, psCtaCod As String, Optional psOpeCod As String = "", Optional ByVal pbExtornoOtroDia As Boolean = False) As Long
    Dim Sql As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    If psOpeCod = "" Then
        If Not pbExtornoOtroDia Then
            Sql = " Select Top 1 nNumTran From ITFDetalle Where nNumTran in (" _
                & " Select nNumTran From Trandiaria " _
                & " WHERE dFecTran = '" & Format(psFecha, "mm/dd/yyyy hh:mm:ss AMPM") & "'" _
                & " and cCodCta = '" & psCtaCod & "')"
        Else
            Sql = " Select Top 1 nNumTran From ITFDetalle Where nNumTran in (" _
                & " Select nNumTran From TrandiariaConsol " _
                & " WHERE dFecTran = '" & Format(psFecha, "mm/dd/yyyy hh:mm:ss AMPM") & "'" _
                & " and cCodCta = '" & psCtaCod & "')"

        End If
    Else
        If Not pbExtornoOtroDia Then
            Sql = " Select Top 1 nNumTran From ITFDetalle Where nNumTran in (" _
                & " Select nNumTran From Trandiaria " _
                & " WHERE dFecTran = '" & Format(psFecha, "mm/dd/yyyy hh:mm:ss AMPM") & "'" _
                & " and cCodCta = '" & psCtaCod & "' and cCodOpe = '" & psOpeCod & "')"
        Else
            Sql = " Select Top 1 nNumTran From ITFDetalle Where nNumTran in (" _
                & " Select nNumTran From TrandiariaConsol " _
                & " WHERE dFecTran = '" & Format(psFecha, "mm/dd/yyyy hh:mm:ss AMPM") & "'" _
                & " and cCodCta = '" & psCtaCod & "' and cCodOpe = '" & psOpeCod & "')"
        End If
    End If
    rs.Open Sql, oConexion, adOpenStatic, adLockReadOnly, adCmdText
    
    If rs.EOF And rs.BOF Then
        GetnNumtranRemoto = 0
    Else
        If IsNull(rs!nNumTran) Then
            GetnNumtranRemoto = 0
        Else
            GetnNumtranRemoto = rs!nNumTran
        End If
    End If
    
    rs.Close
    Set rs = Nothing
    
End Function

Public Function VerifOpeVariasAfectaITF(psOpeCod As String) As Boolean
    Dim Sql As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    oCon.AbreConexion
    
    Sql = "Select cOpeCod From ITFOperaciones Where nTipo = 1 And cOpeCod = '" & Trim(psOpeCod) & "'"
    Set rs = oCon.CargaRecordSet(Sql)
    
    If rs.EOF And rs.BOF Then
        VerifOpeVariasAfectaITF = False
    Else
        VerifOpeVariasAfectaITF = True
    End If
End Function

'Modificado por ARCV 27-10-2006
Public Function fgTruncar(pnNumero As Double, pnDecimales As Integer) As String
'    Dim lnValor As Currency
'    Dim lnRes As Currency
'    lnValor = 10 ^ pnDecimales
'    'lnRes = CDbl(pnNumero) + CDbl((pnNumero - CDbl(pnNumero)) * lnValor) / lnValor
'    lnRes = Int(pnNumero) + Int((pnNumero - Int(pnNumero)) * lnValor) / lnValor
'    fgTruncar = lnRes
Dim intpos  As Integer
Dim lnDecimal As Double
Dim lsDec As String
Dim lnEntero As Long
Dim lnPos As Long

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

'lnEntero = Int(pnNumero)
'lnDecimal = Round(pnNumero - Int(lnEntero), 15) 'ARCV 04-11-2006
'lnDecimal = pnNumero - lnEntero
'lnPos = InStr(1, Trim(Str(lnDecimal)), ".")
'If lnPos > 0 Then
'    lsDec = Mid(Trim(Str(lnDecimal)), lnPos + 1, 2)
'    lsDec = IIf(Len(lsDec) = 1, lsDec * 10, lsDec)
'    lnDecimal = Val(lsDec) / 100
'    fgTruncar = lnEntero + lnDecimal
'Else
'    lnDecimal = 0
'    fgTruncar = lnEntero
'End If
End Function

'Tipo Exoneracion

Public Property Get gnITFTpoSinExoneracion() As Variant
    gnITFTpoSinExoneracion = 0
End Property

Public Property Get gnITFTpoExoPlanilla() As Variant
    gnITFTpoExoPlanilla = 3
End Property

Public Property Get gnITFTpoExoUniColegios() As Variant
    gnITFTpoExoUniColegios = 2
End Property

Public Property Get gnITFTpoExoIntPublicas() As Variant
    gnITFTpoExoIntPublicas = 1
End Property

Public Property Get gnITFTpoExoIntFinanc() As Variant
    gnITFTpoExoIntFinanc = 6
End Property

Public Property Get gsRUCCmac() As Variant
    gsRUCCmac = "20104888934"
End Property

Public Property Get gnITFTpoOpeVarias() As Variant
    gnITFTpoOpeVarias = "1"
End Property

Public Property Get gnITFTpoOpeCaja() As Variant
    gnITFTpoOpeCaja = "2"
End Property



