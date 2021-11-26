VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCartaFianzaCalculos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'** Clase que realiza los calculos financieros asociados a
'** Credito Pignoraticio
Option Explicit

' Calcula Comision de Carta Fianza
Public Function nCalculaComisionCF(ByVal pnMontoCF As Currency, ByVal pnDiasVigencia As Integer, _
        ByVal pnPorcentajeComision As Double) As Currency

Dim lnMontoComision As Currency
On Error GoTo ECalculaComision
    lnMontoComision = Format(pnMontoCF * (pnPorcentajeComision / 100) * (pnDiasVigencia / 360), "####0.00")
    nCalculaComisionCF = lnMontoComision
Exit Function

ECalculaComision:
    Call RaiseError(Err.Number, "Calcula Comision Carta Fianza <<NCartaFianzaCalculos>>")
End Function


' Calcula Comision Trimestral de Carta Fianza
Public Function nCalculaComisionTrimestralCF(ByVal pnMontoCF As Currency, ByVal pnDiasVigencia As Integer, _
        ByVal pnModalidad As ColCFModalidad, ByVal pnMoneda As Integer) As Currency

Dim lnMontoComision As Currency
Dim loBase As DConecta
Dim lr As ADODB.Recordset
Dim lsSQL As String
Dim lnMontoMinimoxTrim As Double, lnTasaInteres As Double
Dim lnNroTrim As Integer
Dim lnMontoMinimoTotal As Double, lnComisionCal As Double


On Error GoTo ECalculaComision
lsSQL = "Select nTasaTrim, nMontoMinimo From ColocCFTarifario Where cTarifCod = '" & Trim(Str(pnModalidad)) & Trim(Str(pnMoneda)) & "' "
    Set loBase = New DConecta
        loBase.AbreConexion
        Set lr = loBase.CargaRecordSet(lsSQL)
            lnMontoMinimoxTrim = lr!nMontoMinimo
            lnTasaInteres = lr!nTasaTrim
        Set lr = Nothing
    Set loBase = Nothing
    lnNroTrim = Int(pnDiasVigencia / 90) + IIf((pnDiasVigencia Mod 90) > 0, 1, 0)
    lnMontoMinimoTotal = lnMontoMinimoxTrim * lnNroTrim
    lnComisionCal = Format(pnMontoCF * (lnTasaInteres / 100) * lnNroTrim, "####0.00")
    If lnComisionCal < lnMontoMinimoTotal Then ' Se toma el Minimo
        nCalculaComisionTrimestralCF = lnMontoMinimoTotal
    Else    ' Se tomo el monto calculado
        nCalculaComisionTrimestralCF = lnComisionCal
    End If
Exit Function

ECalculaComision:
    Call RaiseError(Err.Number, "Calcula Comision CF Trimestral <<NCartaFianzaCalculos>>")
End Function


 
