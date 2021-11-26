Attribute VB_Name = "gCOMFunciones"
Option Explicit
'
'Public Sub CargaComboConstante(ByVal pnCodCons As Integer, ByRef Combo As Variant, _
'    Optional ByVal pnFiltro As Integer = -1)
'Dim sSQL As String
'Dim R As ADODB.Recordset
'Dim oConstante As DCOMConstante
'Dim Datos() As String
'
'    Set oConstante = New DCOMConstante
'    Set R = oConstante.RecuperaConstantes(Trim(Str(pnCodCons)), pnFiltro)
'    ReDim Datos(R.RecordCount)
'    Do While Not R.EOF
'        Datos(R.Bookmark - 1) = Trim(R!cConsDescripcion) & Space(100) & Trim(Str(R!nConsValor))
'        R.MoveNext
'    Loop
'    R.Close
'    Set R = Nothing
'    Set oConstante = Nothing
'
'    Combo = Datos
'End Sub
'
'Public Function ImpreFormat(ByVal pNumero As Variant, ByVal pLongitudEntera As Integer, _
'        Optional ByVal pLongitudDecimal As Integer = 2, _
'        Optional ByVal pMoneda As Boolean = False) As String
'Dim vPosPto As Integer
'Dim vParEnt As String
'Dim vParDec As String
'Dim vLonEnt As Integer
'Dim vLonDec As Integer
'Dim X As Integer
'
'On Error GoTo ErrHandler
'vParDec = ""
'If IsNull(pNumero) Then
'    If pLongitudDecimal > 0 Then vParDec = "." & String(pLongitudDecimal, "0")
'    If pLongitudEntera <= 0 Then pLongitudEntera = 1
'    ImpreFormat = String(pLongitudEntera - 1, " ") & "0" & vParDec
'ElseIf VarType(pNumero) = 8 Then
'    pNumero = Trim(pNumero)
'    vLonEnt = Len(pNumero)
'    If vLonEnt > pLongitudEntera Then
'        pNumero = Left(pNumero, pLongitudEntera)
'        vLonEnt = pLongitudEntera
'    End If
'    ImpreFormat = String(pLongitudDecimal, " ") & pNumero & String(pLongitudEntera - vLonEnt, " ")
'Else
'    vPosPto = InStr(Trim(CStr(pNumero)), ".")
'    If vPosPto > 0 Then
'        vParEnt = Trim(CStr(Left(pNumero, vPosPto - 1)))
'        vParDec = Trim(CStr(Mid(pNumero, vPosPto + 1)))
'        vLonEnt = Len(vParEnt)
'        vLonDec = Len(vParDec)
'    Else
'        vParEnt = Trim(Str(pNumero))
'        vParDec = ""
'        vLonEnt = Len(vParEnt)
'        vLonDec = 0
'    End If
'    If pMoneda And vLonEnt > 3 Then
'        vParEnt = Format(vParEnt, "#,###,###")
'        For X = 1 To Len(vParEnt)
'            If Mid(vParEnt, X, 1) = "," Then pLongitudEntera = pLongitudEntera - 1
'        Next X
'    End If
'    If vLonEnt > pLongitudEntera Then pLongitudEntera = vLonEnt + 1
'    If vLonDec > pLongitudDecimal Then
'        vLonDec = pLongitudDecimal
'        vParDec = Left(vParDec, vLonDec)
'    End If
'    ImpreFormat = String(pLongitudEntera - vLonEnt, " ") & vParEnt
'    If pLongitudDecimal > 0 Then
'        ImpreFormat = ImpreFormat & "." & vParDec & String(pLongitudDecimal - vLonDec, "0")
'    End If
'End If
'Exit Function
'
'ErrHandler:     ' Errores obtenidos
'    MsgBox " Operación no válida " & vbCr & _
'        " Error " & Err.Number & " : " & Err.Description & vbCr & _
'        " Avise al Area de Sistemas ", vbInformation, " ! Aviso ! "
'End Function


'Public Function FechaHora(ByVal psFecha As Date) As String
'    FechaHora = Format(psFecha & Space(1) & GetHoraServer, "mm/dd/yyyy hh:mm:ss")
'End Function
