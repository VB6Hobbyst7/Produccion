Attribute VB_Name = "gRHcacv"

Global gbBitFichaPersonalOtrasPers As Boolean
Global gnMostrarOpcion() As Integer

Function ModoPruebasProcesosRRHH() As Boolean
Dim sValorConst As String

sValorConst = LeeConstSistema(gnModoPruebaRRHH)

If sValorConst = "1" Then
    ModoPruebasProcesosRRHH = True
Else
    ModoPruebasProcesosRRHH = False
End If

End Function

Function ModoActualizarFichaPersonalRRHH() As Boolean
Dim sValorConst As String

sValorConst = LeeConstSistema(gnActualizarFichaPersonal)

If sValorConst = "1" Then
    ModoActualizarFichaPersonalRRHH = True
Else
    ModoActualizarFichaPersonalRRHH = False
End If

End Function
