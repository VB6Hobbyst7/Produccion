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
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long 'EJVG20150110
Private Const LB_SETHORIZONTALEXTENT = &H194

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

Public Function IndiceListaCombo(ByVal Ctrl As Control, ByVal psValor As String, Optional ByVal pnBusq As Integer = 1, Optional ByVal pnDigitoDerecha As Integer = 15) As Long
Dim i As Integer
    IndiceListaCombo = -1
    For i = 0 To Ctrl.ListCount - 1
        If pnBusq = 1 Then
            'If Trim(Right(Ctrl.List(i), 15)) = Trim(psValor) Then
            If Trim(Right(Ctrl.List(i), pnDigitoDerecha)) = Trim(psValor) Then 'EJVG20150127
                IndiceListaCombo = i
                Exit For
            End If
            
        Else
            If pnBusq = 2 Then
                'If Trim(Mid(Ctrl.List(i), 101, 15)) = Trim(psValor) Then
                If Trim(Mid(Ctrl.List(i), 101, pnDigitoDerecha)) = Trim(psValor) Then
                    IndiceListaCombo = i
                    Exit For
                End If
            Else
            
                If Trim(Ctrl.List(i)) = Trim(psValor) Then
                    IndiceListaCombo = i
                    Exit For
                End If
            End If
        End If
    Next i
End Function

Public Sub CargaComboPersonasTipo(ByVal psConstante As PersTipo, ByRef combo As ComboBox)
Dim oPersonas As COMDPersona.DCOMPersonas
Dim R As ADODB.Recordset
    On Error GoTo ERRORCargaComboPersonasTipo
    combo.Clear
    Set oPersonas = New COMDPersona.DCOMPersonas
    Set R = oPersonas.RecuperaPersonasTipo(Trim(Str(psConstante)))
    Set oPersona = Nothing
    Do While Not R.EOF
        combo.AddItem PstaNombre(R!cPersNombre) & Space(250) & R!cPersCod
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    Exit Sub

ERRORCargaComboPersonasTipo:
    MsgBox err.Description, vbCritical, "Aviso"
End Sub


Public Sub CargaComboConstante(ByVal pnCodCons As Integer, ByRef combo As ComboBox, _
    Optional ByVal pnFiltro As Integer = -1)
    'cambiar
Dim ssql As String
Dim R As ADODB.Recordset
Dim oConstante As COMDConstantes.DCOMConstantes
    
    Set oConstante = New COMDConstantes.DCOMConstantes
    Set R = oConstante.RecuperaConstantes(Trim(Str(pnCodCons)), pnFiltro)
    combo.Clear
    Do While Not R.EOF
        combo.AddItem Trim(R!cConsDescripcion) & Space(100) & Trim(Str(R!nConsValor))
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    Set oConstante = Nothing
End Sub
Public Sub CargaComboConstanteMatriz(ByVal Datos As Variant, ByRef combo As ComboBox)
Dim i As Integer

    For i = 0 To UBound(Datos) - 1
        combo.AddItem Datos(i)
    Next i
End Sub
Public Sub LimpiaFlex(ByRef Flex As Control)
Dim i As Integer
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

Public Sub LimpiaControles(pForm As Form, Optional pListas As Boolean = False, Optional ByVal pbNoCombo As Boolean = False)
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
            Case "MASKEDBOX"
                If oControl.Mask = "##/##/####" Then
                    oControl.Text = "__/__/____"
                End If
            Case "COMBOBOX"
                If Not pbNoCombo Then
                    oControl.ListIndex = -1
                End If
            Case "LISTVIEW"
                oControl.ListItems.Clear
            Case "ACTXCODCTA"
                oControl.NroCuenta = ""
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
Dim i As Integer
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
    For i = 0 To UBound(ColBloq) - 1
        FE.ColumnasAEditar = Replace(FE.ColumnasAEditar, Trim(Str(ColBloq(i))), "X")
    Next i
End Sub

Public Function ValorConstante(ByVal pnCabecera As ConstanteCabecera, ByVal pnValor As Integer) As String
Dim oConstante As COMDConstantes.DCOMConstantes

    On Error GoTo ErrorValorConstante
    Set oConstante = New COMDConstantes.DCOMConstantes
    ValorConstante = oConstante.DameDescripcionConstante(pnCabecera, pnValor)
    Set oConstante = Nothing
    Exit Function

ErrorValorConstante:
        MsgBox err.Description, vbCritical, "Aviso"
    

End Function

'Funcion que llena un Combo con un recordset
Sub Llenar_Combo_con_Recordset(pRs As ADODB.Recordset, pcboObjeto As ComboBox)
On Error Resume Next
pcboObjeto.Clear
Do While Not pRs.EOF
    pcboObjeto.AddItem Trim(pRs!cConsDescripcion) & Space(100) & Trim((pRs!nConsValor))
    pRs.MoveNext
Loop
pRs.Close
    
End Sub

'peac 20071226 Funcion que llena un Combo con un recordset para agencias
Sub Llenar_Combo_Agencia_con_Recordset(pRs As ADODB.Recordset, pcboObjeto As ComboBox)

pcboObjeto.Clear
Do While Not pRs.EOF
    pcboObjeto.AddItem Trim(pRs!cConsDescripcion) & Space(100) & Trim(pRs!nConsValor)
    pRs.MoveNext
Loop
pRs.Close
    
End Sub



'Funciones que manejan los Graficos a Mostrar en los Procesos Largos
Sub AbrirControlAnimation(poAnimation As Animation, pnOpcion As Integer)
Dim lsFile As String
Select Case pnOpcion
    Case 0  'Buscando
        lsFile = App.Path & "\Videos\FINDCOMP.AVI"
    Case 1  'Grabando
        lsFile = App.Path & "\Videos\Grabando.AVI"
End Select

Screen.MousePointer = vbHourglass
poAnimation.Open lsFile
poAnimation.Play
poAnimation.Visible = True
DoEvents

End Sub
Sub CerrarControlAnimation(poAnimation As Animation)
    poAnimation.Stop
    poAnimation.Visible = False
    Screen.MousePointer = vbDefault
End Sub

Public Function DevuelveCantidadCheckList(ByVal lstLista As ListBox) As Integer
    Dim i As Integer
    Dim Cant As Integer
    
    For i = 1 To lstLista.ListCount
        If lstLista.Selected(i - 1) = True Then
            Cant = Cant + 1
        End If
    Next
    DevuelveCantidadCheckList = Cant
End Function
'Add By GITU 2013-06-14
Public Sub CargaComboAgencias(ByRef combo As ComboBox)
Dim oConst As COMDConstantes.DCOMAgencias
Dim R As ADODB.Recordset
    On Error GoTo ERRORCargaComboAgencias
    combo.Clear
    Set oConst = New COMDConstantes.DCOMAgencias
    Set R = oConst.ObtieneAgencias()
    Set oConst = Nothing
    Do While Not R.EOF
        combo.AddItem R!cConsDescripcion & Space(250) & R!nConsValor
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    Exit Sub

ERRORCargaComboAgencias:
    MsgBox err.Description, vbCritical, "Aviso"
End Sub

Public Sub CargaComboMeses(ByRef combo As ComboBox)
    combo.Clear
    combo.AddItem "Enero" & Space(50) & "01"
    combo.AddItem "Febrero" & Space(50) & "02"
    combo.AddItem "Marzo" & Space(50) & "03"
    combo.AddItem "Abril" & Space(50) & "04"
    combo.AddItem "Mayo" & Space(50) & "05"
    combo.AddItem "Junio" & Space(50) & "06"
    combo.AddItem "Julio" & Space(50) & "07"
    combo.AddItem "Agosto" & Space(50) & "08"
    combo.AddItem "Septiembre" & Space(50) & "09"
    combo.AddItem "Octubre" & Space(50) & "10"
    combo.AddItem "Noviembre" & Space(50) & "11"
    combo.AddItem "Diciembre" & Space(50) & "12"
End Sub
'End GITU
'EJVG20140815 ***
Public Function EnfocaControl(Ctrl As Control) As Boolean
    If Ctrl.Visible And Ctrl.Enabled Then
        Ctrl.SetFocus
        EnfocaControl = True
    End If
End Function
'END EJVG *******
'EJVG20150110 ***
Public Function AplicarScrollBarListBox(ListBox As ListBox) As Long
    Dim ret          As Long
    Dim i            As Integer
    Dim J            As Long
    Dim Ancho_Maximo As Long
    Dim Ancho_Texto  As Long
    Dim LBParent     As Object
      
    Set LBParent = ListBox.Parent
    Ancho_Maximo = -1
    J = -1
    For i = 0 To ListBox.ListCount - 1
        Ancho_Texto = LBParent.TextWidth(ListBox.List(i))
        If Ancho_Texto > Ancho_Maximo Then
            Ancho_Maximo = Ancho_Texto + (10 * Screen.TwipsPerPixelX)
            J = i
        End If
    Next
    Set LBParent = Nothing
    ' -- Establecer el Scroll
    ret = SendMessage(ListBox.hwnd, LB_SETHORIZONTALEXTENT, (Ancho_Maximo / Screen.TwipsPerPixelX), ByVal 0&)
    ' -- retornar item mas largo
    Aplicar_ScrollBar = J
End Function
Public Sub Llenar_Combo_con_Recordset_New(pRs As ADODB.Recordset, pcboObjeto As ComboBox)
    pcboObjeto.Clear
    Do While Not pRs.EOF
        pcboObjeto.AddItem Trim(pRs.Fields(1)) & Space(100) & Trim(Str(pRs.Fields(0)))
        pRs.MoveNext
    Loop
    pRs.Close
End Sub
Public Function EsCeldaFlexEditable(ByRef pflex As FlexEdit, ByVal pnCol As Long) As Boolean
    Dim Editar() As String
    
    EsCeldaFlexEditable = True
    Editar = Split(pflex.ColumnasAEditar, "-")
    If Editar(pnCol) = "X" Then
        MsgBox "Esta celda no es editable", vbInformation, "Aviso"
        EsCeldaFlexEditable = False
        Exit Function
    End If
End Function
'END EJVG *******

'***CTI3 (ferimoro)   18102018
Sub Llenar_Combo_MotivoExtorno(pRs As ADODB.Recordset, pcboObjeto As ComboBox)
pcboObjeto.Clear
Do While Not pRs.EOF
    pcboObjeto.AddItem Trim(pRs!motivos)
    pRs.MoveNext
Loop
pRs.Close
End Sub
