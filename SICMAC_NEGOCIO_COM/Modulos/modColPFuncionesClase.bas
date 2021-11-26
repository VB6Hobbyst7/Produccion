Attribute VB_Name = "modColPFuncionesClase"
'********************************************************
'** Modulo para funciones que se usaran en las clases de
'** Credito Pignoraticio
'********************************************************


'Rutina para quebrar un texto multiline en una impresión
Public Function QuiebreTexto(vTexto As String, vFila As Byte) As String
Dim vLinea As String
Dim X As Integer
Dim paso As Integer
paso = 0
vLinea = ""
'MsgBox vTexto
For X = 1 To Len(vTexto)
    If Asc(Mid(vTexto, X, 1)) = 13 Or X = Len(vTexto) Then
        paso = paso + 1
        If paso = vFila Then
            If Len(Trim(vLinea)) < 2 Then
                QuiebreTexto = " " 'IIf(paso = 1, vLinea, Right(vLinea, Len(vLinea)))
            Else
                QuiebreTexto = IIf(paso = 1, vLinea, Right(vLinea, Len(vLinea) - 2))
            End If
            Exit Function
        End If
        vLinea = ""
    End If
    vLinea = vLinea & Mid(vTexto, X, 1)
Next X
End Function


Public Function InicialBoveda(pCodAge As String) As String
InicialBoveda = ""
Select Case pCodAge
    Case "11201"
        InicialBoveda = "Pi"
    Case "11202"
        InicialBoveda = "Zo"
    Case "11203"
        InicialBoveda = "Po"
    Case "11204"
        InicialBoveda = "SD"
    Case "11205"
        InicialBoveda = "Es"
    Case "11206"
        InicialBoveda = "Ch"
    Case "11207"
        InicialBoveda = "Se"
    Case "11208"
        InicialBoveda = "Hu"
    Case "11209"
        InicialBoveda = "He"
    Case "11210"
        InicialBoveda = "Vi"
End Select

End Function

Public Function CalculaTasaEfectivaAnual(ByVal pnTasaMensual As Double) As Double
Dim lnTasa As Double
    lnTasa = (pnTasaMensual + 1) ^ 12 - 1
CalculaTasaEfectivaAnual = lnTasa
End Function

Public Function mfgEstadoCredPigDesc(ByVal pnEstado As Integer) As String
Dim lsDesc As String
    Select Case pnEstado
        Case gColPEstRegis
            lsDesc = "Registrado"
        Case gColPEstDesem
            lsDesc = "Desembolsado"
        Case gColPEstDifer
            lsDesc = "Diferido (Para rescate)"
        Case gColPEstCance
            lsDesc = "Cancelado"
        Case gColPEstVenci
            lsDesc = "Vencido"
        Case gColPEstRemat
            lsDesc = "Remate"
        Case gColPEstPRema
            lsDesc = "Para Remate"
        Case gColPEstRenov
            lsDesc = "Renovado"
        Case gColPEstAdjud
            lsDesc = "Adjudicado"
        Case gColPEstSubas
            lsDesc = "Subastada"
        Case gColPEstAnula
            lsDesc = "Anulado"
        Case gColPEstChafa
            lsDesc = "Chafaloneado"
    End Select
    mfgEstadoCredPigDesc = lsDesc
End Function

Public Function mfgEstadoColocRecupDesc(ByVal pnEstado As Integer) As String
Dim lsDesc As String
    Select Case pnEstado
        Case gColocEstRecVigJud
            lsDesc = "Vigente Judicial"
        Case gColocEstRecVigCast
            lsDesc = "Vigente Castigado"
        Case gColocEstRecCanJud
            lsDesc = "Cancelado"
        Case gColocEstRecCanCast
            lsDesc = "Cancelado"
    End Select
    mfgEstadoColocRecupDesc = lsDesc
End Function

