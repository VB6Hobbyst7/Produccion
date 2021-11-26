VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NGarantia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Function InsertarGarantiaLogistica(ByVal psNumGarant As String, ByVal pnPeriodo As Integer, ByVal pnTValor As Integer, ByVal pnMontValor As Double, ByVal pcDesValor As String, ByVal pnMoneda As Integer, ByVal pnEstado As Integer, ByVal psPersCod As String, ByVal pnContador As Integer, ByVal pdHoy As Date, ByVal psUser As String, ByVal psAgeCod As String)
Dim sSql As String
Dim oGarant As DGarantia
'Dim oBase As COMDCredito.DCOMCredActBD
Dim pbTran As Boolean

    On Error GoTo ErrorInsertarGarantiaLogistica
    
    'Genera cMovNro
    Dim psMovAct As String
    Dim loContFunct As NContFunciones
    
    Set loContFunct = New NContFunciones
    psMovAct = loContFunct.GeneraMovNro(pdHoy, psAgeCod, psUser) ' gsCodUser)
    Set loContFunct = Nothing
    'Fin de Generacion de cMovNro
     
    pbTran = True
    'oBase.dBeginTrans
    Set oGarant = New DGarantia
        Call oGarant.InsertarGarantiaLogistica(psNumGarant, pnPeriodo, pnTValor, pnMontValor, pcDesValor, pnMoneda, 1, psMovNro, psPersCod, pnContador)
    Set oGarant = Nothing
    
    If pnContador = 0 Then
    Set oGarant = New DGarantia
        'Set oBase = New COMDCredito.DCOMCredActBD
        Call oGarant.dUpdateGarantiasAdjudicados(psNumGarant, pdHoy, pnEstado, psMovAct) ', psUsuariAdju)
        'Call oBase.dCommitTrans
        Set oGarant = Nothing
    End If
    
    
    Exit Function

ErrorInsertarGarantiaLogistica:
    'If pbTran Then
     '   Call oBase.dRollbackTrans
    'End If
    Err.Raise Err.Number, "Valida Datos Garantia", Err.Description
End Function
'****
Public Function ActualizargarantiasxVentaLogistica(ByVal psNumGarant As String, ByVal pnEstadoAdju As Integer, ByVal pdHoy As Date, ByVal psUser As String, ByVal pdFechaCompra As Date, ByVal pnMonedaAdju As Integer, ByVal pnTipoCambio As Double, ByVal psPersCodComprador As String, ByVal pnVendido As Integer, ByVal psAgeCod As String)
Dim sSql As String
Dim oGarant As DGarantia

Dim pbTran As Boolean

    On Error GoTo ErrorActualizargarantiasxVentaLogistica
    
    'Genera cMovNro
    Dim psMovAct As String
    Dim loContFunct As NContFunciones
    
    Set loContFunct = New NContFunciones
    psMovAct = loContFunct.GeneraMovNro(pdHoy, psAgeCod, psUser) ' gsCodUser)
    Set loContFunct = Nothing
    'Fin de Generacion de cMovNro
     
    pbTran = True
    Set oGarant = New DGarantia
        Call oGarant.ActualizargarantiasxVentaLogistica(psNumGarant, pnEstadoAdju, pdHoy, psMovAct, pdFechaCompra, pnMonedaAdju, pnTipoCambio, psPersCodComprador, pnVendido)
    Set oGarant = Nothing
    Set oGarant = New DGarantia
    If pnContador = 0 Then
        Call oGarant.dUpdateGarantiasAdjudicados(psNumGarant, pdHoy, pnEstadoAdju, psMovAct) ', psUsuariAdju)
        Set oGarant = Nothing
    End If
    
    
    Exit Function

ErrorActualizargarantiasxVentaLogistica:
    Err.Raise Err.Number, "Valida Datos Garantia", Err.Description
End Function
'ALPA ********
Public Function ImpresionRetasacionGarantia(ByVal rs As ADODB.Recordset, ByVal gsNomCmac As String, ByVal gdFecSis As String, _
                                   ByVal gsNomAge As String, ByVal gsCodUser As String, Optional ByVal psImpresora As Impresoras = gEPSON, _
                                   Optional ByVal pnTipoCambio As Double = 0#, Optional ByVal pnVendido As Integer = 0, Optional ByVal psPropietario As String = "") As String
Dim sCad As String
Dim I As Integer
Dim oFunI As New ContsImp.clsConstImp
Dim nSMonME As Double
Dim nSMonMN As Double
Dim nPosc As Integer
Dim nPos1 As Integer
nSMonME = 0#
nSMonMN = 0#
nPos1 = 0

oFunI.Inicia psImpresora
    sCad = oFunI.gPrnSaltoLinea

    sCad = sCad & Space(2) & gsNomCmac & Space(100 - Len(gsNomCmac)) & gdFecSis & oFunI.gPrnSaltoLinea
    sCad = sCad & Space(2) & gsNomAge & Space(100 - Len(gsNomAge)) & gsCodUser & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
    sCad = sCad & Space(40) & " VALORIZACION DE INMUEBLE - RESUMEN" & oFunI.gPrnSaltoLinea
    sCad = sCad & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
    sCad = sCad & Space(6) & "CLIENTE: CAJA MUNICIPAL DE AHORRO Y CREDITOS MAYNAS S.A." & oFunI.gPrnSaltoLinea

    If pnVendido = 0 Then
        sCad = sCad & Space(6) & "PROPIETARIO: CAJA MUNICIPAL DE AHORRO Y CREDITOS MAYNAS SA " & oFunI.gPrnSaltoLinea
    Else
        sCad = sCad & Space(6) & "PROPIETARIO: " & psPropietario & oFunI.gPrnSaltoLinea
    End If

    sCad = sCad & Space(6) & "TIPO DE CAMBIO: " & pnTipoCambio & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
    sCad = sCad & Space(5) & String(110, "-") & oFunI.gPrnSaltoLinea
    sCad = sCad & Space(6) & ImpreFormat("CONCEPTO", 47) & ImpreFormat("VAL.COM.", 11) & ImpreFormat("VAL.REA.", 20) & oFunI.gPrnSaltoLinea
    sCad = sCad & Space(5) & String(110, "-") & oFunI.gPrnSaltoLinea

   I = 1
   nPosc = 0
   If Not (rs.EOF And rs.BOF) Then
      rs.MoveFirst
      Do Until rs.EOF
            If nPosc = 0 And Left(CStr(rs!nTipo), 1) = 2 Then
       
                sCad = sCad & Space(5) & String(110, "-") & oFunI.gPrnSaltoLinea
                sCad = sCad & Space(42) & ImpreFormat(nSMonME, 20)
                sCad = sCad & ImpreFormat(nSMonMN, 10) & oFunI.gPrnSaltoLinea
                sCad = sCad & Space(5) & String(110, "-") & oFunI.gPrnSaltoLinea
                sCad = sCad & Space(6) & ImpreFormat("AL T.CAM", 47) & ImpreFormat("VAL.COM.", 11) & ImpreFormat("VAL.REA.", 20) & oFunI.gPrnSaltoLinea
                sCad = sCad & Space(5) & String(110, "-") & oFunI.gPrnSaltoLinea
                sCad = sCad & Space(42) & ImpreFormat(nSMonME * pnTipoCambio, 20)
                sCad = sCad & ImpreFormat(nSMonMN * pnTipoCambio, 10) & oFunI.gPrnSaltoLinea
                sCad = sCad & Space(5) & String(110, "-") & oFunI.gPrnSaltoLinea
                sCad = sCad & Space(6) & ImpreFormat("LINDEROS", 47) & ImpreFormat("VALOR.", 11) & ImpreFormat("ML.", 20) & oFunI.gPrnSaltoLinea
                sCad = sCad & Space(5) & String(110, "-") & oFunI.gPrnSaltoLinea
                nPosc = nPosc + 1
                nPos1 = 1
            End If
            If rs!nSuma = 1 Then
                sCad = sCad & Space(2) & ImpreFormat(I, 4, 0, False) & ImpreFormat(rs!cDesTipo, 34)
                sCad = sCad & ImpreFormat(rs!nMonto, 20)
                sCad = sCad & ImpreFormat(rs!nMontoRea, 10)
                nSMonME = nSMonME + CDbl(rs!nMonto)
                nSMonMN = nSMonMN + CDbl(rs!nMontoRea)
            Else
                sCad = sCad & Space(2) & ImpreFormat(I, 4, 0, False) & ImpreFormat(rs!cDesTipo & "-" & rs!cDescripcion, 34)
                sCad = sCad & ImpreFormat(rs!nMonto, 20)
                sCad = sCad & ImpreFormat("", 10)
            End If
            
            sCad = sCad & oFunI.gPrnSaltoLinea
        rs.MoveNext
        I = I + 1
      Loop
      sCad = sCad & oFunI.gPrnSaltoLinea
    End If

    If nPos1 = 0 Then
        sCad = sCad & Space(5) & String(110, "-") & oFunI.gPrnSaltoLinea
        sCad = sCad & Space(42) & ImpreFormat(nSMonME, 20)
        sCad = sCad & ImpreFormat(nSMonMN, 10) & oFunI.gPrnSaltoLinea
        sCad = sCad & Space(5) & String(110, "-") & oFunI.gPrnSaltoLinea
        sCad = sCad & Space(6) & ImpreFormat("AL T.CAM", 47) & ImpreFormat("VAL.COM.", 11) & ImpreFormat("VAL.REA.", 20) & oFunI.gPrnSaltoLinea
        sCad = sCad & Space(5) & String(110, "-") & oFunI.gPrnSaltoLinea
        sCad = sCad & Space(42) & ImpreFormat(nSMonME * pnTipoCambio, 20)
        sCad = sCad & ImpreFormat(nSMonMN * pnTipoCambio, 10) & oFunI.gPrnSaltoLinea
        sCad = sCad & Space(5) & String(110, "-") & oFunI.gPrnSaltoLinea
        
    End If
    ImpresionRetasacionGarantia = sCad
    
End Function
Public Function InsertarGarantiaSaneamiento(ByVal psNumGarant As String, ByVal psCtaCod As String, ByVal pnTipoSan As Integer, ByVal pnPeriSan As Integer, ByVal pnMontSan As Double, ByVal pdFecSan As Date, ByVal psUsuariAdju As String, ByVal pnContador As Integer, ByVal pnEstadoAdj As Integer, ByVal dfechaAdj As Date, ByVal pnEstado As Integer, ByVal pdHoy As Date, ByVal psUsuarioGaran As String, ByVal psCodAge As String, ByVal pnMoneda As Integer, ByVal nTESan As Integer)
Dim sSql As String
Dim oGarant As DGarantia
'Dim oBase As DCredActBD
Dim pbTran As Boolean

    On Error GoTo ErrorInsertarGarantiaSaneamiento
    
    'Genera cMovNro
    Dim psMovAct As String
    Dim loContFunct As NContFunciones
    
    Set loContFunct = New NContFunciones
    psMovAct = loContFunct.GeneraMovNro(pdHoy, psCodAge, psUsuariAdju) ' gsCodUser)
    Set loContFunct = Nothing
    'Fin de Generacion de cMovNro
     
    pbTran = True
    'oBase.dBeginTrans
    Set oGarant = New DGarantia
        Call oGarant.InsertarGarantiaSaneamiento(psNumGarant, psCtaCod, pnTipoSan, pnPeriSan, pnMontSan, pdFecSan, psMovAct, pnContador, pnMoneda, nTESan)
    Set oGarant = Nothing
    If pnContador = 0 Then
        Set oGarant = New DGarantia
        'Set oBase = New COMDCredito.DCOMCredActBD
        Call oGarant.dUpdateGarantiasAdjudicados(psNumGarant, pdHoy, pnEstado, psMovAct)
        'Call oBase.dCommitTrans
        Set oGarant = Nothing
    End If
    
    
    Exit Function

ErrorInsertarGarantiaSaneamiento:
    'If pbTran Then
     '   Call oBase.dRollbackTrans
    'End If
    Err.Raise Err.Number, "Valida Datos Garantia", Err.Description
End Function
