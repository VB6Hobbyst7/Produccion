VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NGasto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private MatGastos() As String

Public Function GeneraCalendarioGastos(ByRef MatCalend As Variant, ByRef MatDesemb As Variant, ByRef nNumGastos As Integer, ByVal psCtaCod As String) As Variant
Dim oGastos As DGastos
Dim R As ADODB.Recordset
Dim I As Integer
Dim sProd As String
    On Error GoTo ErrorGeneraCalendarioGastos
    nNumGastos = 0
    ReDim MatGastos(1000, 4)
    Set oGastos = New DGastos
    Set R = oGastos.RecuperaGastosAplicablesDesembolso
    Set oGastos = Nothing
    
    'Aplicar Gastos al Desembolso
    If IsArray(MatDesemb) Then
        ReDim Preserve MatDesemb(UBound(MatDesemb), 3)
        Do While Not R.EOF
            For I = 0 To UBound(MatDesemb) - 1
                If IsNull(R!cProducto) Then
                    sProd = ""
                Else
                    sProd = Trim(R!cProducto)
                End If
                If sProd = "" Or sProd = Mid(psCtaCod, 6, 3) Then
                    If R!nMoneda = CInt(Mid(psCtaCod, 9, 1)) Then
                        If CDbl(MatDesemb(I, 1)) >= R!nInicial And CDbl(MatDesemb(I, 1)) <= R!nFinal Then
                            'Si es Fijo
                            If R!nTpoValor = 1 Then
                                MatDesemb(I, 2) = Format(IIf(IsNull(R!nValor), 0, R!nValor), "#0.00")
                            End If
                            'Si es Porcentaje
                            If R!nTpoValor = 2 Then
                                MatDesemb(I, 2) = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.00") * CDbl(MatDesemb(I, 1)), "#0.00")
                            End If
                            nNumGastos = nNumGastos + 1
                            MatGastos(nNumGastos - 1, 0) = "Desembolso" & Space(50) & gColocCalendAplDesembolso
                            MatGastos(nNumGastos - 1, 1) = I + 1
                            MatGastos(nNumGastos - 1, 2) = R!cDescripcion & Space(150) & R!nColocConceptoCod
                            MatGastos(nNumGastos - 1, 3) = MatDesemb(I, 2)
                            If R!nAplicado = gColocConceptoAplDesembolso Then
                                Exit For
                            End If
                        End If
                    End If
                End If
            Next I
            R.MoveNext
        Loop
        R.Close
        Set R = Nothing
    End If
    
    'Aplicar Gastos a la Cuota
    Set oGastos = New DGastos
    Set R = oGastos.RecuperaGastosAplicablesCuotas
    Set oGastos = Nothing
    'ReDim Preserve MatCalend(UBound(MatCalend), 7)
    Do While Not R.EOF
        For I = 0 To UBound(MatCalend) - 1
            If R!nMoneda = CInt(Mid(psCtaCod, 9, 1)) Then
                If IsNull(R!cProducto) Then
                    sProd = ""
                Else
                    sProd = Trim(R!cProducto)
                End If
                If sProd = "" Or sProd = Mid(psCtaCod, 6, 3) Then
                    If CDbl(MatCalend(I, 2)) >= R!nInicial And CDbl(MatCalend(I, 2)) <= R!nFinal Then
                        'Si es Fijo
                        If R!nTpoValor = 1 Then
                            MatCalend(I, 6) = Format(IIf(IsNull(R!nValor), 0, R!nValor), "#0.00")
                        End If
                        'Si es Porcentaje
                        If R!nTpoValor = 2 Then
                            MatCalend(I, 6) = Format(Format(IIf(IsNull(R!nValor), 0, R!nValor) / 100, "#0.00") * CDbl(MatCalend(I, 2)), "#0.00")
                        End If
                        nNumGastos = nNumGastos + 1
                        MatGastos(nNumGastos - 1, 0) = "Cuota" & Space(50) & gColocCalendAplCuota
                        MatGastos(nNumGastos - 1, 1) = I + 1
                        MatGastos(nNumGastos - 1, 2) = R!cDescripcion & Space(150) & R!nColocConceptoCod
                        MatGastos(nNumGastos - 1, 3) = MatCalend(I, 6)
                        If R!nAplicado = gColocConceptoAplCuota Then
                            Exit For
                        End If
                    End If
                End If
            End If
        Next I
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    GeneraCalendarioGastos = MatGastos
    Exit Function

ErrorGeneraCalendarioGastos:
    Err.Raise Err.Number, "Error En Proceso", Err.Description

End Function

Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing


End Sub
