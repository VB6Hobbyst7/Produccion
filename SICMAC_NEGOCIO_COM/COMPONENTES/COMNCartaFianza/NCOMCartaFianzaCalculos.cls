VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCOMCartaFianzaCalculos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'** Clase que realiza los calculos financieros asociados a
'** Credito Pignoraticio
Option Explicit
Dim oError As New COMConecta.COMErrorHandling

' Calcula Comision de Carta Fianza
Public Function nCalculaComisionCF(ByVal pnMontoCF As Currency, ByVal pnDiasVigencia As Integer, _
        ByVal pnPorcentajeComision As Double, ByVal pnMoneda As Integer) As Currency

Dim lnMontoComision As Currency
Dim pnComisionValor As Double
Dim pnMontoValor As Double
Dim pnTipoCambio As Double
Dim O As COMConecta.DCOMConecta
Dim R As ADODB.Recordset
Dim sSql As String

On Error GoTo ECalculaComision

'Buscamos porcentaje de comision
sSql = "Select nTasaTrim from ColocCfTarifario where nMontoMinimo <=" & pnMontoCF & " AND nMontoMax >= " & pnMontoCF & " AND nMoneda = " & pnMoneda
Set O = New COMConecta.DCOMConecta
O.AbreConexion
Set R = O.CargaRecordSet(sSql)
O.CierraConexion

If R.RecordCount > 0 Then
    pnComisionValor = R!nTasaTrim
Else
    pnComisionValor = 0
End If
R.Close

'Buscamos comison de Carta Fianza
sSql = "select nParamValor  from ColocParametro  where  nParamVar='4002'"
Set O = New COMConecta.DCOMConecta
O.AbreConexion
Set R = O.CargaRecordSet(sSql)
O.CierraConexion

If R.RecordCount > 0 Then
    pnMontoValor = R!nParamValor
Else
    pnMontoValor = 0
End If
R.Close

'ARCV 05-07-2007
'sSql = "select  Top 1  nValPond from TipoCambio  order  by  dFecCamb desc"
sSql = "select  Top 1  nValFijoDia from TipoCambio  order  by  dFecCamb desc"
'------
Set O = New COMConecta.DCOMConecta
O.AbreConexion
Set R = O.CargaRecordSet(sSql)
O.CierraConexion

If R.RecordCount > 0 Then
    pnTipoCambio = R!nValPond
Else
    pnTipoCambio = 0
End If
R.Close

    If pnMoneda = 1 Then
        pnMontoValor = pnMontoValor * pnTipoCambio
    End If

   ' lnMontoComision = Format(pnMontoCF * (pnPorcentajeComision / 100) * (pnDiasVigencia / 360), "####0.00")
   ' nCalculaComisionCF = lnMontoComision
    
    
    lnMontoComision = Format(pnMontoCF * (pnComisionValor / 100) * (pnDiasVigencia / 30), "####0.00") + pnMontoValor
    
    nCalculaComisionCF = lnMontoComision
    
    
Set O = Nothing

Exit Function

ECalculaComision:
    Call oError.RaiseError(Err.Number, "Calcula Comision Carta Fianza <<NCartaFianzaCalculos>>")
End Function


' Calcula Comision Trimestral de Carta Fianza
Public Function nCalculaComisionTrimestralCF(ByVal pnMontoCF As Currency, ByVal pnDiasVigencia As Integer, _
        ByVal pnModalidad As ColCFModalidad, ByVal pnMoneda As Integer, Optional pcCtaCod As String = "", Optional pnTipoGarant As Integer) As Currency

Dim lnMontoComision As Currency
Dim O As COMConecta.DCOMConecta
Dim lr As New ADODB.Recordset
Dim lsSQL As String
Dim lnMontoMinimoxTrim As Double, lnTasaInteres As Double
Dim lnNroTrim As Integer
Dim lnMontoMinimoTotal As Double, lnComisionCal As Double
Dim oDCarFianza As COMDCartaFianza.DCOMCartaFianza
Dim lnComiTriCF As Double
Dim sSql As String
Dim R As New ADODB.Recordset
Dim pnTipoCambio As Double
Dim pnMontoValor As Double
'ALPA 20090331*************
Dim pnMontoCartaFianza As Double
'**************************
On Error GoTo ECalculaComision
'    lsSQL = "Select nTasaTrim, nMontoMinimo From ColocCFTarifario Where cTarifCod = '" & Trim(Str(pnModalidad)) & Trim(Str(pnMoneda)) & "' "
'    Set loBase = New COMConecta.DCOMConecta
'        loBase.AbreConexion
'        Set lr = loBase.CargaRecordSet(lsSQL)
'            lnMontoMinimoxTrim = lr!nMontoMinimo
'            lnTasaInteres = lr!nTasaTrim
'        Set lr = Nothing
'    Set loBase = Nothing

    sSql = "select  Top 1  nValPond from TipoCambio  order  by  dFecCamb desc"
    Set O = New COMConecta.DCOMConecta
    O.AbreConexion
    Set R = O.CargaRecordSet(sSql)
    O.CierraConexion
    
    If R.RecordCount > 0 Then
        pnTipoCambio = R!nValPond
    Else
        pnTipoCambio = 0
    End If
    R.Close
    
Set oDCarFianza = New COMDCartaFianza.DCOMCartaFianza

pnMontoCartaFianza = oDCarFianza.CantidadCartaFianza(pcCtaCod, pnTipoGarant, pnTipoCambio)
'MAVM 20091028 ********* Acta Nº 107-2009
    If pnMontoCartaFianza > 10000 Then
        lnNroTrim = Int(pnDiasVigencia / 90) + IIf((pnDiasVigencia Mod 90) > 0, 1, 0)
        lnComisionCal = Format(pnMontoCF * (1.25 / 100) * lnNroTrim, "####0.00")
    Else
        
        sSql = "select nParamValor  from ColocParametro  where  nParamVar='4002'"
        Set O = New COMConecta.DCOMConecta
        O.AbreConexion
        Set R = O.CargaRecordSet(sSql)
        O.CierraConexion
        
        If R.RecordCount > 0 Then
            pnMontoValor = R!nParamValor
        Else
            pnMontoValor = 0
        End If
        R.Close
        
        If pnMoneda = 1 Then
            pnMontoValor = pnMontoValor * pnTipoCambio
        Else
            pnMontoValor = 50
        End If
        
        lnNroTrim = Int(pnDiasVigencia / 90) + IIf((pnDiasVigencia Mod 90) > 0, 1, 0)
        'lnComisionCal = Format(pnMontoCF * (1.5 / 100) * lnNroTrim, "####0.00")
        lnComisionCal = Format(pnMontoCF * (1.5 / 100), "####0.00")
        
        If lnComisionCal <= pnMontoValor Then
            lnComisionCal = Format((pnMontoValor * lnNroTrim), "####0.00")
        Else
            lnComisionCal = Format((lnComisionCal * lnNroTrim), "####0.00")
        End If
        
    End If
    
    nCalculaComisionTrimestralCF = lnComisionCal

'End MAVM 20091028 *********

'ALPA 20090331
'    sSql = "Select cTarifCod,nTasaTrim,nMontoMinimo from ColocCfTarifario where nMontoMinimo <=" & pnMontoCF & " AND nMontoMax >= " & pnMontoCF & " AND nMoneda = " & pnMoneda
'    sSql = "Select cTarifCod,nTasaTrim,nMontoMinimo from ColocCfTarifario where nMontoMinimo <=" & pnMontoCF & " AND nMontoMax >= " & pnMontoCF & " AND nMoneda = " & pnMoneda
'    Set O = New COMConecta.DCOMConecta
'    O.AbreConexion
'    Set R = O.CargaRecordSet(sSql)
'    O.CierraConexion
'    lnTasaInteres = 0
'    If R.RecordCount > 0 Then
'        lnMontoMinimoxTrim = R!nMontoMinimo
'        lnTasaInteres = R!nTasaTrim
'    Else
'        lnMontoMinimoxTrim = 0
'        lnTasaInteres = 0
'    End If
'    R.Close
'
'
'    sSql = " select comiCF=dbo.fnc_CalculaComisionTrimestralCF(" & pnMontoCartaFianza & "," & pnMontoCF & "," & pnMoneda & "," & pnTipoCambio & "," & pnDiasVigencia & ")"
'    Set O = New COMConecta.DCOMConecta
'    O.AbreConexion
'    Set R = O.CargaRecordSet(sSql)
'    O.CierraConexion
'
'    If Not (R.EOF And R.BOF) Then
'        lnComiTriCF = R!comiCF
'    Else
'        lnComiTriCF = 0
'    End If
    
'nCalculaComisionTrimestralCF = " select dbo.fnc_CalculaComisionTrimestralCF(" & pnMontoCartaFianza & "," & pnMontoCF & "," & pnMoneda & "," & pnTipoCambio & "," & pnDiasVigencia & ")"
    
    '*** PEAC 20090707
'    'ALPA 20090331************************************
'    If pnMontoCartaFianza >= 10000 Then
'    sSql = "Select cTarifCod,nTasaTrim,nMontoMinimo from ColocCfTarifario where nMontoMinimo <=" & pnMontoCF & " AND nMontoMax >= " & pnMontoCF & " AND nMoneda = " & pnMoneda & " and cTarifCod in (13,14)"
'    Set O = New COMConecta.DCOMConecta
'    O.AbreConexion
'    Set R = O.CargaRecordSet(sSql)
'    O.CierraConexion
'    lnTasaInteres = 0
'    If R.RecordCount > 0 Then
'        If R!cTarifCod = 13 Then
'            lnMontoMinimoxTrim = R!nMontoMinimo
'            lnTasaInteres = R!nTasaTrim
'        ElseIf R!cTarifCod = 14 Then
'            lnMontoMinimoxTrim = R!nMontoMinimo / pnTipoCambio
'            lnTasaInteres = R!nTasaTrim
'        End If
'
'    Else
'        lnMontoMinimoxTrim = 0
'        lnTasaInteres = 0
'    End If
'    R.Close
'        'Calculamos comison de Carta Fianza siempre y cuando sea >= 10000
'        If IIf(pnMoneda = 1, (pnMontoCF / pnTipoCambio), pnMontoCF) >= 10000 Then
'            If lnMontoMinimoxTrim <> 0 Then
'                '*** PEAC 20090413
'                lnNroTrim = Int(pnDiasVigencia / 90) + IIf((pnDiasVigencia Mod 90) > 0, 1, 0)
'                lnMontoMinimoxTrim = lnMontoMinimoxTrim * lnNroTrim
'                'lnComisionCal = Format(pnMontoCF * (1.25 / 100) * lnNroTrim, "####0.00")
'                lnComisionCal = Format(pnMontoCF * (lnTasaInteres / 100) * lnNroTrim, "####0.00")
'                If lnComisionCal < lnMontoMinimoTotal Then ' Se toma el Minimo
'                    nCalculaComisionTrimestralCF = lnMontoMinimoTotal
'                Else    ' Se tomo el monto calculado
'                    nCalculaComisionTrimestralCF = lnComisionCal
'                End If
'            Else
'                nCalculaComisionTrimestralCF = pnMontoValor
'            End If
'        End If
'    Else
'    '****************************************************
'    'Calculamos comison de Carta Fianza siempre y cuando sea < 3000
'    If pnMontoCF < 3000 Then
'        sSql = "select nParamValor  from ColocParametro  where  nParamVar='4002'"
'        Set O = New COMConecta.DCOMConecta
'        O.AbreConexion
'        Set R = O.CargaRecordSet(sSql)
'        O.CierraConexion
'
'        If R.RecordCount > 0 Then
'            pnMontoValor = R!nParamValor
'            If pnMoneda = 1 Then
'                lnNroTrim = Int(pnDiasVigencia / 90) + IIf((pnDiasVigencia Mod 90) > 0, 1, 0)
'                pnMontoValor = pnMontoValor * lnNroTrim * pnTipoCambio
'            End If
'        Else
'            pnMontoValor = 0
'        End If
'        R.Close
'        nCalculaComisionTrimestralCF = pnMontoValor
'    End If
'
'    'Calculamos comison de Carta Fianza siempre y cuando sea >= 3000
'    If pnMontoCF >= 3000 Then
'        If lnMontoMinimoxTrim <> 0 Then
'            lnNroTrim = Int(pnDiasVigencia / 90) + IIf((pnDiasVigencia Mod 90) > 0, 1, 0)
'            lnMontoMinimoxTrim = lnMontoMinimoxTrim * lnNroTrim
'            lnComisionCal = Format(pnMontoCF * (lnTasaInteres / 100) * lnNroTrim, "####0.00")
'            If lnComisionCal < lnMontoMinimoTotal Then ' Se toma el Minimo
'                nCalculaComisionTrimestralCF = lnMontoMinimoTotal
'            Else    ' Se tomo el monto calculado
'                nCalculaComisionTrimestralCF = lnComisionCal
'            End If
'        Else
'            nCalculaComisionTrimestralCF = pnMontoValor
'        End If
'   End If
'End If


'nCalculaComisionTrimestralCF = lnComiTriCF

Exit Function

ECalculaComision:
    Call oError.RaiseError(Err.Number, "Calcula Comision CF Trimestral <<NCartaFianzaCalculos>>")
End Function


 


