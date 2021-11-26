Attribute VB_Name = "gFunControles"
Global Const Espaciado = 60
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4
Private Const SWP_NOMOVE = &H2
Private Const SWP_DRAWFRAME = &H20
Private Const WS_DLGFRAME = &H400000
Private Const GWL_STYLE = (-16)
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

'Verifica la corrceta habilitación de la impresora
Public Function ImpreSensa() As Boolean
Dim lbArchAbierto As Boolean
On Error GoTo ControlError
    MDISicmact.SBBarra.Panels(1).Text = "Verificando Conexión con Impresora"
    ArcSal = FreeFile
    lbArchAbierto = True
    Open sLPT For Output As ArcSal
    Print #ArcSal, oImpresora.gPrnInicializa;             'Inicializa Impresora
    Close ArcSal
    lbArchAbierto = False
    ImpreSensa = True
    MDISicmact.SBBarra.Panels(1).Text = ""
    Exit Function
ControlError:   ' Rutina de control de errores.
    MDISicmact.SBBarra.Panels(1).Text = ""
    If lbArchAbierto Then
        Close ArcSal
    End If
    MsgBox "Impresora no Encontrada ó Inactiva" & vbCr & "Verifique que la Conexión sea Correcta", vbExclamation, "Aviso de Precaución"
    ImpreSensa = False
End Function
Public Sub ControlDialog(ControlName As Control, SetTrue As Boolean)
'* Fixed Dialog BorderStyle propiedad para un Control *
    Dim dwStyle As Long
    dwStyle = GetWindowLong(ControlName.hwnd, GWL_STYLE)
    If SetTrue Then
      dwStyle = dwStyle Or WS_DLGFRAME
    Else
      dwStyle = dwStyle - WS_DLGFRAME
    End If
    dwStyle = SetWindowLong(ControlName.hwnd, GWL_STYLE, dwStyle)
    SetWindowPos ControlName.hwnd, ControlName.Parent.hwnd, 0, 0, 0, 0, SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_DRAWFRAME
End Sub

Public Function IndiceListaCombo(ByVal Ctrl As Control, ByVal psValor As String) As Long
Dim I As Integer
    IndiceListaCombo = -1
    For I = 0 To Ctrl.ListCount - 1
        If Trim(Right(Ctrl.List(I), 15)) = Trim(psValor) Then
            IndiceListaCombo = I
            Exit For
        End If
    Next I
End Function

Public Sub CargaComboPersonasTipo(ByVal psConstante As PersTipo, ByRef Combo As ComboBox)
Dim oPersonas As DPersonas
Dim R As ADODB.Recordset
    On Error GoTo ERRORCargaComboPersonasTipo
    Combo.Clear
    Set oPersonas = New DPersonas
    Set R = oPersonas.RecuperaPersonasTipo(Trim(Str(psConstante)))
    Set oPersona = Nothing
    Do While Not R.EOF
        Combo.AddItem R!cPersNombre & Space(250) & R!cPersCod
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    Exit Sub

ERRORCargaComboPersonasTipo:
    MsgBox TextErr(Err.Description), vbInformation, "Aviso"
End Sub
Public Sub CargaComboConstante(ByVal pnCodCons As Integer, ByRef Combo As ComboBox)
Dim sSql As String
Dim R As ADODB.Recordset
Dim oConstante As DConstantes
    On Error GoTo CargaComboConstanteErr
    Set oConstante = New DConstantes
    Set R = oConstante.RecuperaConstantes(Trim(Str(pnCodCons)))
    Combo.Clear
    Do While Not R.EOF
        Combo.AddItem Trim(R!cConsDescripcion) & Space(100) & R!nConsValor
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    Set oConstante = Nothing
Exit Sub
CargaComboConstanteErr:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub LimpiaFlex(ByRef Flex As Control)
    Flex.Rows = 2
    For I = 0 To Flex.Cols - 1
        Flex.TextMatrix(1, I) = ""
    Next I
End Sub

Public Sub HabilitaControles(pForm As Form, ByVal pbHabilita As Boolean, Optional pLabel As Boolean = False, Optional pBotones As Boolean = False)
Dim oControl As Control
    For Each oControl In pForm.Controls
        If UCase(TypeName(oControl)) = "COMMANDBUTTON" And pBotones Then
            oControl.Enabled = pbHabilita
        Else
            If UCase(TypeName(oControl)) = "LABEL" And pLabel Then
                oControl.Enabled = pbHabilita
            Else
                If UCase(TypeName(oControl)) <> "COMMANDBUTTON" And UCase(TypeName(oControl)) <> "LABEL" And UCase(TypeName(oControl)) <> "LINE" Then
                    oControl.Enabled = pbHabilita
                End If
            End If
        End If
    Next
End Sub

Public Sub LimpiaControles(pForm As Form, Optional pListas As Boolean = False)
Dim oControl As Control
Dim I As Integer
Dim oTmp As FlexEdit

    For Each oControl In pForm.Controls
        Select Case UCase(TypeName(oControl))
            Case "LABEL"
                If oControl.BorderStyle = 1 Then
                    oControl.Caption = ""
                End If
            Case "TEXTBOX"
                oControl.Text = ""
            Case "LISTBOX"
                If pListas Then
                    oControl.Clear
                End If
        End Select
    Next
End Sub

Public Sub InicializaCombos(pForm As Form)
Dim oControl As Control
    For Each oControl In pForm.Controls
        If UCase(TypeName(oControl)) = "COMBOBOX" Then
            oControl.ListIndex = -1
        End If
    Next
End Sub

Public Function ValidaApostrofe(ByVal psCadena As String) As String
    ValidaApostrofe = Replace(psCadena, "'", "''")
End Function
Public Sub HabilitaFlexNormal(ByRef FE As FlexEdit, ByVal pnFilaAct As Integer, ByVal psColBloq As String)
    FE.row = pnFilaAct
    FE.ColumnasAEditar = psColBloq
    Call FE.BackColorRow(vbWhite)
End Sub

Public Sub HabilitaFilaFlex(ByRef pnFilaAct As Integer, ByRef FE As FlexEdit, ByVal ColBloq As Variant, Optional ByVal SelecColorFlex As Long = vbYellow, Optional ByVal bModificar As Boolean = False)
Dim I As Integer
Dim J As Integer
    
    If bModificar Then
        pnFilaAct = FE.row
    Else
        If Trim(FE.TextMatrix(1, 1)) = "" Then
            pnFilaAct = 1
        Else
            pnFilaAct = FE.Rows
        End If
        FE.AdicionaFila
        FE.row = pnFilaAct
    End If
    Call FE.BackColorRow(SelecColorFlex)
    For I = 0 To UBound(ColBloq) - 1
        FE.ColumnasAEditar = Replace(FE.ColumnasAEditar, Trim(Str(ColBloq(I))), "X")
    Next I
End Sub

Public Function ValorConstante(ByVal pnCabecera As ConstanteCabecera, ByVal pnValor As Integer) As String
Dim oConstante As DConstantes

    On Error GoTo ErrorValorConstante
    Set oConstante = New DConstantes
    ValorConstante = oConstante.DameDescripcionConstante(pnCabecera, pnValor)
    Set oConstante = Nothing
    Exit Function

ErrorValorConstante:
        MsgBox TextErr(Err.Description), vbInformation, "Aviso"
    

End Function
'EJVG20140515 ***
Public Sub EnfocaControl(Ctrl As Control)
    If Ctrl.Visible And Ctrl.Enabled Then Ctrl.SetFocus
End Sub
'END EJVG *******
'PASIERS0872014
Public Function TextBox_SoloNumeros(ByVal KeyAscii As Integer) As Integer
    If InStr("0123456789", Chr(KeyAscii)) = 0 Then
        TextBox_SoloNumeros = 0
    Else
        TextBox_SoloNumeros = KeyAscii
    End If
    If KeyAscii = 8 Then TextBox_SoloNumeros = KeyAscii
    If KeyAscii = 13 Then TextBox_SoloNumeros = KeyAscii
End Function
Public Function TextBox_SoloNumerosDecimales(ByVal KeyAscii As Integer) As Integer
    If InStr("0123456789.", Chr(KeyAscii)) = 0 Then
        TextBox_SoloNumerosDecimales = 0
    Else
        TextBox_SoloNumerosDecimales = KeyAscii
    End If
    If KeyAscii = 8 Then TextBox_SoloNumerosDecimales = KeyAscii
    If KeyAscii = 13 Then TextBox_SoloNumerosDecimales = KeyAscii
End Function
'END PASI
