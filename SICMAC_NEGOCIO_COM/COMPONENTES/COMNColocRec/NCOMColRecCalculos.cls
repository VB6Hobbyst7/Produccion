VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCOMColRecCalculos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'** Clase que realiza los calculos financieros asociados a
'** Credito en Recuperaciones
Option Explicit
Dim oError As New COMConecta.COMErrorHandling



Public Function nCalculaIntCompGenerado(ByVal pnDiasCalculo As Integer, ByVal pnTasaInt As Double, _
        ByVal pnSaldoCap As Currency) As Currency
' Calcula el Interes Compensatorio Genereado por los dias transcurridos
On Error GoTo ECalcula
'********** Realiza el calculo
    If pnDiasCalculo > 0 Then
        nCalculaIntCompGenerado = (((1 + (pnTasaInt / 100)) ^ (pnDiasCalculo / 30)) - 1) * pnSaldoCap
    Else
        nCalculaIntCompGenerado = 0
    End If
Exit Function

ECalcula:
    Call oError.RaiseError(Err.Number, "Calcula Int Comp Generado <<NColRecCalculos>>")
End Function
Public Function nCalculaIntCompGeneradoICA(ByVal pnDiasCalculo As Integer, ByVal pnTasaInt As Double, _
        ByVal pnSaldoCap As Currency) As Currency
' Calcula el Interes Compensatorio Genereado por los dias transcurridos
On Error GoTo ECalcula
'********** Realiza el calculo
    If pnDiasCalculo > 0 Then
        'nCalculaIntCompGeneradoICA = (((1 + (pnTasaInt / 100)) ^ (pnDiasCalculo / 30)) - 1) * pnSaldoCap
        nCalculaIntCompGeneradoICA = Round(((pnSaldoCap * (pnTasaInt / 100)) / 30) * pnDiasCalculo, 2)
    Else
        nCalculaIntCompGeneradoICA = 0
    End If
Exit Function

ECalcula:
    Call oError.RaiseError(Err.Number, "Calcula Int Comp Generado <<NColRecCalculos>>")
End Function


Public Function nCalculaComisionAbogado(ByVal pnPorcComision As Double, _
        ByVal pnMontoPago As Currency) As Currency
' Calcula la comision del abogado de acuerdo al monto pagado
On Error GoTo ECalcula
'********** Realiza el calculo
    nCalculaComisionAbogado = pnMontoPago * ((pnPorcComision / 100) / (1 + (pnPorcComision / 100)))
Exit Function

ECalcula:
    Call oError.RaiseError(Err.Number, "Calcula Comision de Abogado <<NColRecCalculos>>")
End Function

Public Function nCalculaIntMoratorioGenerado(ByVal pnDiasCalculo As Integer, ByVal pnTasaInt As Double, _
        ByVal pnBaseCalculo As Currency) As Currency
' Calcula el Interes Moratorio Genereado por los dias transcurridos
On Error GoTo ECalcula
'********** Realiza el calculo
    If pnDiasCalculo > 0 Then
        nCalculaIntMoratorioGenerado = (((1 + (pnTasaInt / 100)) ^ (pnDiasCalculo / 30)) - 1) * pnBaseCalculo
    Else
        nCalculaIntMoratorioGenerado = 0
    End If
Exit Function

ECalcula:
    Call oError.RaiseError(Err.Number, "Calcula Int Moratorio Generado <<NColRecCalculos>>")
End Function

Public Function nCalculaIntMoratorioGeneradoICA(ByVal pnDiasCalculo As Integer, ByVal pnTasaInt As Double, _
        ByVal pnBaseCalculo As Currency) As Currency
' Calcula el Interes Moratorio Genereado por los dias transcurridos
On Error GoTo ECalcula
'********** Realiza el calculo
    If pnDiasCalculo > 0 Then
        'nCalculaIntMoratorioGeneradoICA = (((1 + (pnTasaInt / 100)) ^ (pnDiasCalculo / 30)) - 1) * pnBaseCalculo
        nCalculaIntMoratorioGeneradoICA = Round(pnBaseCalculo * (pnTasaInt / 100) * pnDiasCalculo, 2)
    Else
        nCalculaIntMoratorioGeneradoICA = 0
    End If
Exit Function

ECalcula:
    Call oError.RaiseError(Err.Number, "Calcula Int Moratorio Generado <<NColRecCalculos>>")
End Function

'**Creado por DAOR 20070421, Funci�n que Genera la Comisi�n del Abogado
'**Cuando se conoce la suma de CIM(Capital, Interes y Mora)
Public Function nCalculaComisionAbogadoSimple(ByVal pnPorcComision As Double, _
        ByVal pnMontoCIM As Currency) As Currency
On Error GoTo ECalcula
    nCalculaComisionAbogadoSimple = pnMontoCIM * (pnPorcComision / 100)
    Exit Function
ECalcula:
    Call oError.RaiseError(Err.Number, "Calculo Simple de Comision de Abogado  <<NColRecCalculos>>")
End Function



