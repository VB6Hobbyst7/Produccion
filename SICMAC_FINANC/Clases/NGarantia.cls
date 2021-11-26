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
Option Explicit



Public Function PorcentajeGarantia(ByVal psConsGarantia As String) As Double
Dim oConecta As DConecta
Dim sSQL As String
Dim R As ADODB.Recordset

    On Error GoTo ErrorPorcentajeGarantia
    sSQL = "Select nParamValor from ColocParametro Where nParamVar = " & psConsGarantia
    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set R = oConecta.CargaRecordSet(sSQL)
    If Not R.BOF And Not R.EOF Then
        PorcentajeGarantia = R!nParamValor / 100
    Else
        PorcentajeGarantia = 0
    End If
    oConecta.CierraConexion
    Set oConecta = Nothing
    R.Close
    Set R = Nothing
    Exit Function

ErrorPorcentajeGarantia:
    Set oConecta = Nothing
    Set R = Nothing
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Function

Public Function ValidaDatos(ByVal RelPers As Variant, ByVal pnMontoTasac As Double, _
    ByVal pnMontoReali As Double, ByVal pnMontoDisp As Double, Optional ByVal pbSoloGaran As Boolean = False) As String
Dim I As Integer
Dim nNumTit As Integer

     ValidaDatos = ""
    'Valida Numero de Titulares
    If Not pbSoloGaran Then
        nNumTit = 0
        For I = 0 To UBound(RelPers) - 1
            If CInt(Trim(Right(RelPers(I), 10))) = gPersRelGarantiaTitular Then
                nNumTit = nNumTit + 1
            End If
        Next I
        If nNumTit > 1 Then
            ValidaDatos = "No puede haber mas de un Titular"
        End If
    End If
    'Valida Montos Tasacion, Realizacion, Disponible
    If pnMontoReali > pnMontoTasac Then
        ValidaDatos = "El Monto de Realizacion No Puede Ser Mayor que el Monto de Tasacion"
        Exit Function
    End If
    
    If pnMontoDisp > pnMontoReali Then
        ValidaDatos = "El Monto Disponible No Puede Ser Mayor que el Monto de Realizacion"
        Exit Function
    End If
    
End Function

Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing

End Sub
