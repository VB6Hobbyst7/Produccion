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
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'Public Function GeneraMovNroActualiza(pdFecha As Date, psCodUser As String, psCodCmac As String, psCodAge As String) As String
'GeneraMovNroActualiza = Format(pdFecha & " " & GetHoraServer, gsFormatoMovFechaHora) & psCodCmac & psCodAge & "00" & psCodUser
'End Function
Public Sub EnviaPrevio(psImpre As String, psTitulo As String, ByVal pnLinPage As Integer, Optional plCondensado As Boolean = False)
Dim clsPrevio As New Previo.clsPrevio
clsPrevio.Show psImpre, psTitulo, plCondensado, pnLinPage
Set clsPrevio = Nothing
End Sub

'Verifica la corrceta habilitación de la impresora
Public Function ImpreSensa() As Boolean
Dim lbArchAbierto As Boolean
On Error GoTo ControlError
    MDISicmact.staMain.Panels(1).Text = "Verificando Conexión con Impresora"
    ArcSal = FreeFile
    lbArchAbierto = True
    Open sLpt For Output As ArcSal
    Print #ArcSal, oImpresora.gPrnInicializa;             'Inicializa Impresora
    Close ArcSal
    lbArchAbierto = False
    ImpreSensa = True
    MDISicmact.staMain.Panels(1).Text = ""
    Exit Function
ControlError:   ' Rutina de control de errores.
    MDISicmact.staMain.Panels(1).Text = ""
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

Public Function IndiceListaCombo(ByVal Ctrl As Control, ByVal psValor As String, Optional ByVal pnBusq As Integer = 1) As Long
Dim i As Integer
    IndiceListaCombo = -1
    For i = 0 To Ctrl.ListCount - 1
        If pnBusq = 1 Then
            If Trim(Right(Ctrl.List(i), 15)) = Trim(psValor) Then
                IndiceListaCombo = i
                Exit For
            End If
        Else
            If Trim(Ctrl.List(i)) = Trim(psValor) Then
                IndiceListaCombo = i
                Exit For
            End If
        End If
    Next i
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
        Combo.AddItem PstaNombre(R!cPersNombre) & Space(250) & R!cPersCod
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    Exit Sub

ERRORCargaComboPersonasTipo:
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
Public Sub CargaComboConstante(ByVal pnCodCons As Integer, ByRef Combo As ComboBox)
Dim sSQL As String
Dim R As ADODB.Recordset
Dim oConstante As DConstante
    
    Set oConstante = New DConstante
    Set R = oConstante.RecuperaConstantes(Trim(Str(pnCodCons)))
    Combo.Clear
    Do While Not R.EOF
        Combo.AddItem Trim(R!cConsDescripcion) & Space(100) & R!nConsValor
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    Set oConstante = Nothing
End Sub

Public Sub LimpiaFlex(ByRef Flex As Control)
    Flex.Rows = 2
    For i = 0 To Flex.Cols - 1
        Flex.TextMatrix(1, i) = ""
    Next i
End Sub

'Traspasados
Public Function CentraSdi(frmCentra As Form) As Integer
    frmCentra.Move (Screen.Width - frmCentra.Width) / 2, (Screen.Height - frmCentra.Height) / 2, frmCentra.Width, frmCentra.Height
    CentraSdi = 1
End Function

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
Dim i As Integer
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
    FE.Row = pnFilaAct
    FE.ColumnasAEditar = psColBloq
    Call FE.BackColorRow(vbWhite)
End Sub

Public Sub HabilitaFilaFlex(ByRef pnFilaAct As Integer, ByRef FE As FlexEdit, ByVal ColBloq As Variant, Optional ByVal SelecColorFlex As Long = vbYellow, Optional ByVal bModificar As Boolean = False)
Dim i As Integer
Dim J As Integer
    
    If bModificar Then
        pnFilaAct = FE.Row
    Else
        If Trim(FE.TextMatrix(1, 1)) = "" Then
            pnFilaAct = 1
        Else
            pnFilaAct = FE.Rows
        End If
        FE.AdicionaFila
        FE.Row = pnFilaAct
    End If
    Call FE.BackColorRow(SelecColorFlex)
    For i = 0 To UBound(ColBloq) - 1
        FE.ColumnasAEditar = Replace(FE.ColumnasAEditar, Trim(Str(ColBloq(i))), "X")
    Next i
End Sub

Public Function ValorConstante(ByVal pnCabecera As ConstanteCabecera, ByVal pnValor As Integer) As String
Dim oConstante As DConstante

    On Error GoTo ErrorValorConstante
    Set oConstante = New DConstante
    ValorConstante = oConstante.DameDescripcionConstante(pnCabecera, pnValor)
    Set oConstante = Nothing
    Exit Function

ErrorValorConstante:
        MsgBox Err.Description, vbCritical, "Aviso"
    

End Function
