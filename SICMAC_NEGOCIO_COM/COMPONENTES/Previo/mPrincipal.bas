Attribute VB_Name = "mPrincipal"
Option Explicit
Public lbCancela As Boolean
Public lnNumCopias As Integer
Public sLpt As String
Public lTpoImpresora As Impresoras
Public oImpresora As New COMFunciones.FCOMVarImpresion

Public Sub fEnfoque(ctrControl As Control)
    ctrControl.SelStart = 0
    ctrControl.SelLength = Len(ctrControl.Text)
End Sub

Public Function Letras(intTecla As Integer, Optional lbMayusculas As Boolean = True) As Integer
    If lbMayusculas Then
        Letras = Asc(UCase(Chr(intTecla)))
    Else
        Letras = Asc(LCase(Chr(intTecla)))
    End If
End Function

Public Function NumerosEnteros(intTecla As Integer, Optional pbNegativos As Boolean = False) As Integer
    Dim cValidar As String
        If pbNegativos = False Then
            cValidar = "0123456789"
        Else
            cValidar = "0123456789-"
        End If
        If intTecla > 26 Then
            If InStr(cValidar, Chr(intTecla)) = 0 Then
                intTecla = 0
                Beep
            End If
        End If
        NumerosEnteros = intTecla
End Function

