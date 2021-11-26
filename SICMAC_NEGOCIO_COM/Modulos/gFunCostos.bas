Attribute VB_Name = "gFuncCostos"
Option Explicit

Public Function JIZQ(vCadena As String, nTam As Integer, Optional vChar As String) As String
Dim s As String, xChar As String

xChar = IIf(Len(Trim(vChar)) = 1, vChar, Space(1))
vCadena = Trim(vCadena)
If nTam > Len(Trim(vCadena)) Then
   s = String(nTam - Len(Trim(vCadena)), xChar)
   JIZQ = vCadena + s
Else
   JIZQ = Mid(vCadena, 1, nTam)
End If
End Function

Public Function JDER(vCadena As String, nTam As Integer, Optional vChar As String) As String
Dim s As String, xChar As String

vCadena = Trim(vCadena)
xChar = IIf(Len(Trim(vChar)) = 1, vChar, Space(1))
If nTam >= Len(Trim(vCadena)) Then
   s = String(nTam - Len(Trim(vCadena)), xChar)
   JDER = s + vCadena
Else
   JDER = String(nTam, "*")
End If
End Function
