VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.UserControl FlexEdit 
   ClientHeight    =   2160
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3915
   PropertyPages   =   "FlexEdit.ctx":0000
   ScaleHeight     =   2160
   ScaleWidth      =   3915
   ToolboxBitmap   =   "FlexEdit.ctx":0048
   Begin SICMACT.TxtBuscar TxtBuscar 
      Height          =   330
      Left            =   1560
      TabIndex        =   4
      Top             =   1545
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ComboBox cboCelda 
      Height          =   315
      Left            =   675
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   870
      Visible         =   0   'False
      Width           =   1965
   End
   Begin MSMask.MaskEdBox txtFecha 
      Height          =   375
      Left            =   495
      TabIndex        =   2
      Top             =   1500
      Visible         =   0   'False
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin VB.TextBox txtCelda 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      Height          =   330
      Left            =   615
      TabIndex        =   1
      Top             =   1260
      Visible         =   0   'False
      Width           =   1425
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlxGd 
      Height          =   900
      Left            =   120
      TabIndex        =   0
      Top             =   510
      Width           =   2970
      _ExtentX        =   5239
      _ExtentY        =   1588
      _Version        =   393216
      FocusRect       =   0
      AllowUserResizing=   3
      RowSizingMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Image ImgOpt1 
      Height          =   225
      Left            =   2610
      Picture         =   "FlexEdit.ctx":035A
      Top             =   180
      Width           =   240
   End
   Begin VB.Image ImgOpt0 
      Height          =   225
      Left            =   2205
      Picture         =   "FlexEdit.ctx":066C
      Top             =   150
      Width           =   240
   End
   Begin VB.Image ImgCheck1 
      Height          =   240
      Left            =   1095
      Picture         =   "FlexEdit.ctx":097E
      Top             =   180
      Width           =   240
   End
   Begin VB.Image ImgCheck0 
      Height          =   240
      Left            =   1455
      Picture         =   "FlexEdit.ctx":0CC0
      Top             =   210
      Width           =   240
   End
   Begin VB.Image imgHolders 
      Height          =   165
      Index           =   1
      Left            =   705
      Picture         =   "FlexEdit.ctx":1002
      Top             =   -15
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgHolders 
      Height          =   165
      Index           =   2
      Left            =   465
      Picture         =   "FlexEdit.ctx":11FC
      Top             =   300
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgHolders 
      Height          =   105
      Index           =   0
      Left            =   810
      Picture         =   "FlexEdit.ctx":13F6
      Top             =   225
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Menu MnuFGridRows 
      Caption         =   "Rows"
      Begin VB.Menu MnuFGridAddRow 
         Caption         =   "AdicionaFila"
      End
      Begin VB.Menu MnuFGridGuion 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDeleteGridRow 
         Caption         =   "EliminarFila"
      End
   End
End
Attribute VB_Name = "FlexEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'Default Property Values:
Const m_def_SoloFila = False
Const m_def_AutoAdd = False
Const m_def_lbBuscaDuplicadoText = False
Const m_def_lbOrdenaCol = False
Const m_def_lbPuntero = False
Const m_def_lbFormatoCol = False
Const m_def_lbFlexDuplicados = True
Const m_def_lbEditarFlex = False
Const m_def_lbRsLoad = False
Const m_def_rsFlex = 0
Const m_def_SQLDataSource = ""
Const m_def_AvanceCeldas = 0
Const m_def_CantEntero = 8
Const m_def_CantDecimales = 2
Const m_def_EncabezadosAlineacion = 0
Const m_def_FormatosEdit = 0
Const m_def_ListaControles = ""
Const m_def_ColumnasAEditar = ""
Const m_def_EncabezadosNombres = ""
Const m_def_EncabezadosAnchos = ""
Const m_def_VisiblePopMenu = False
Const m_def_ScaleMode = 0

Dim m_VisiblePopMenu As Boolean
Dim m_ScaleMode As Integer
'Property Variables:
Dim m_SoloFila As Boolean
Dim m_AutoAdd As Boolean
Dim m_lbBuscaDuplicadoText As Boolean
Dim m_lbOrdenaCol As Boolean
Dim lnColChk As Long
Dim lnColOpt As Long
Dim m_lbPuntero As Boolean
Dim m_lbFormatoCol As Boolean
Dim m_lbFlexDuplicados As Boolean
Dim m_lbEditarFlex As Boolean
Dim m_lbRsLoad As Boolean
Dim m_rsFlex As ADODB.Recordset
Dim m_SQLDataSource As String
Dim m_AvanceCeldas As Integer
Dim m_CantEntero As Integer
Dim m_CantDecimales As Integer
Dim m_EncabezadosAlineacion As Variant
Dim m_FormatosEdit As Variant
Dim m_ListaControles As String
Dim m_ColumnasAEditar As String
Dim m_EncabezadosNombres As String
Dim m_EncabezadosAnchos As String
Dim lbCambio As Boolean
Dim lnFilaActual As Long
Dim m_iSortCol As Long
Dim m_iSortType As Long
Dim m_iSortCustomAscending As Boolean
Dim i As Long
Dim lbFocoFlex As Boolean

Public Enum Avance
    Horizontal = 0
    Vertical = 1
End Enum

Private Enum PictureType
    ptArrow = 0
    ptPen
    ptStar
    ptNone
End Enum

'Event Declarations:
Event OnChangeCombo() 'MappingInfo=cboCelda,cboCelda,-1,Click
Attribute OnChangeCombo.VB_Description = "Ocurre cuando el usuario presiona y libera un botón del mouse encima de un objeto."
Event OnClickTxtBuscar(psCodigo As String, psDescripcion As String) 'MappingInfo=TxtBuscar,TxtBuscar,-1,Click
Event DblClick() 'MappingInfo=FlxGd,FlxGd,-1,DblClick
Attribute DblClick.VB_Description = "Fired when the user double-clicks the mouse over the control."
Event OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
Event OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
Event OnRowAdd(pnRow As Long)
Event OnRowDelete()
Event OnRowChange(pnRow As Long, pnCol As Long)
Event OnEnterTextBuscar(psDataCod As String, pnRow As Long, pnCol As Long, pbEsDuplicado As Boolean)         'MappingInfo=TxtBuscar,TxtBuscar,-1,EmiteDatos
Event OnCellChange(pnRow As Long, pnCol As Long)
Event EnterCell() 'MappingInfo=FlxGd,FlxGd,-1,EnterCell
Attribute EnterCell.VB_Description = "Fired before the cursor enters a cell."
Event Click() 'MappingInfo=FlxGd,FlxGd,-1,Click
Attribute Click.VB_Description = "Fired when the user presses and releases the mouse button over the control."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=FlxGd,FlxGd,-1,KeyPress
Attribute KeyPress.VB_Description = "Fired when the user presses a key."
Event RowColChange() 'MappingInfo=FlxGd,FlxGd,-1,RowColChange
Attribute RowColChange.VB_Description = "Fired when the current cell changes."
Event SelChange() 'MappingInfo=FlxGd,FlxGd,-1,SelChange
Attribute SelChange.VB_Description = "Fired when the selected range of cells changes."
Private Function DeterminaControl(lnCol As Long) As Object
Dim pControles As String
Dim X As Long
Dim lControles() As String
Dim lnNroControl As Long
pControles = m_ListaControles

If Len(Trim(pControles)) > 0 Then
    For X = 0 To Cols - 1
        vPos = InStr(1, pControles, "-", vbTextCompare)
        ReDim Preserve lControles(X)
        lControles(X) = Mid(pControles, 1, IIf(vPos > 0, vPos - 1, Len(pControles)))
        If pControles <> "" Then
            pControles = Mid(pControles, IIf(vPos > 0, vPos + 1, Len(pControles)))
        End If
        If lnCol = X Then
            lnNroControl = Val(lControles(X))
            Exit For
        End If
    Next X
End If
Select Case lnNroControl
    Case 0
        Set DeterminaControl = txtCelda
    Case 1
        Set DeterminaControl = TxtBuscar
    Case 2
        Set DeterminaControl = txtFecha
    Case 3
        Set DeterminaControl = cboCelda
End Select
End Function
Private Function DeterminaFormato(lnCol As Long) As Long
Dim pFormatos As String
Dim X As Long
Dim lFormatos() As String
Dim lnNroFormato As Long
pFormatos = m_FormatosEdit
If Len(Trim(pFormatos)) > 0 Then
    For X = 0 To Cols - 1
        vPos = InStr(1, pFormatos, "-", vbTextCompare)
        ReDim Preserve lFormatos(X)
        If vPos > 0 Then
            lFormatos(X) = Mid(pFormatos, 1, IIf(vPos > 0, vPos - 1, Len(pFormatos)))
        Else
            If pFormatos <> "" Then
                lFormatos(X) = pFormatos
                pFormatos = ""
            End If
        End If
        If pFormatos <> "" Then
            pFormatos = Mid(pFormatos, IIf(vPos > 0, vPos + 1, Len(pFormatos)))
        End If
        If lnCol = X Then
            lnNroFormato = Val(lFormatos(X))
            Exit For
        End If
    Next X
End If
DeterminaFormato = lnNroFormato
End Function
Private Sub cboCelda_Click()
If cboCelda.ListIndex <> -1 Then
    FlxGd.Text = cboCelda
End If
RaiseEvent OnChangeCombo
End Sub
Private Sub cboCelda_KeyPress(KeyAscii As Integer)
Dim lsAux As String
Dim lbCancel As Boolean
If KeyAscii = 13 Then
    lbCancel = True
    lsAux = FlxGd.Text
    FlxGd.Text = cboCelda
    RaiseEvent OnValidate(FlxGd.row, FlxGd.Col, lbCancel)
    If lbCancel = True Then
        RaiseEvent OnCellChange(FlxGd.row, FlxGd.Col)
    Else
        FlxGd.Text = lsAux
        Exit Sub
    End If
    Advance_Cell
    FlxGd.SetFocus
End If
End Sub
Private Sub cboCelda_LostFocus()
cboCelda.ListIndex = -1
cboCelda.Visible = False
FlxGd.SetFocus
RaiseEvent OnCellChange(FlxGd.row, FlxGd.Col)
End Sub

Private Sub FlxGd_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)
    'On Error Resume Next
    Dim dtmRow1 As Date, dtmRow2 As Date
    Dim lnRow1 As Currency, lnRow2 As Currency
    With FlxGd
        If IsNumeric(FlxGd.TextMatrix(1, FlxGd.Col)) Then
            lnRow1 = CCur(IIf(.TextMatrix(Row1, m_iSortCol) = "", "0", IIf(IsNumeric(.TextMatrix(Row1, m_iSortCol)) = False, "0", .TextMatrix(Row1, m_iSortCol))))
            lnRow2 = CCur(IIf(.TextMatrix(Row2, m_iSortCol) = "", "0", IIf(IsNumeric(.TextMatrix(Row2, m_iSortCol)) = False, "0", .TextMatrix(Row2, m_iSortCol))))
            If lnRow1 > lnRow2 Then
                Cmp = IIf(m_iSortCustomAscending, 1, -1)
            ElseIf lnRow1 = lnRow2 Then
                Cmp = 0
            Else
                Cmp = IIf(m_iSortCustomAscending, -1, 1)
            End If
        Else
            dtmRow1 = CDate(IIf(IsDate(.TextMatrix(Row1, m_iSortCol)) = False, "01/01/2000", .TextMatrix(Row1, m_iSortCol)))
            dtmRow2 = CDate(IIf(IsDate(.TextMatrix(Row2, m_iSortCol)) = False, "01/01/2000", .TextMatrix(Row2, m_iSortCol)))
            If dtmRow1 > dtmRow2 Then
                Cmp = IIf(m_iSortCustomAscending, 1, -1)
            ElseIf dtmRow1 = dtmRow2 Then
                Cmp = 0
            Else
                Cmp = IIf(m_iSortCustomAscending, -1, 1)
            End If
        End If
        
    End With
End Sub
Private Sub FlxGd_DblClick()
If FlxGd.MouseRow < FlxGd.FixedRows Then
    If FlxGd.TextMatrix(1, 0) <> "" Then
        DoColumnSort
    End If
End If
RaiseEvent DblClick
If m_lbEditarFlex = False Then Exit Sub
If ColumnaAEditar(FlxGd.Col) = False Then
    KeyAscii = 0
    Exit Sub
End If
EnfocaTexto DeterminaControl(FlxGd.Col), 0, FlxGd
End Sub

Private Sub FlxGd_GotFocus()
On Error GoTo ErrGotFocus
Dim oObjeto As Object
lbFocoFlex = True
txtCelda.Visible = False
TxtBuscar.Visible = False
txtFecha.Visible = False
cboCelda.Visible = False
If FlxGd.TextMatrix(FlxGd.row, 0) = "" Then
    DrawIcons ptNone, FlxGd.row
Else
    DrawIcons ptArrow, FlxGd.row
    lnFilaActual = FlxGd.row
End If
Exit Sub
ErrGotFocus:
    MsgBox err.Description, vbInformation, "¡AViso!"
End Sub

Private Sub FlxGd_LostFocus()
lbFocoFlex = False
End Sub

Private Sub FlxGd_Scroll()
    txtCelda.Visible = False
    TxtBuscar.Visible = False
    txtFecha.Visible = False
    cboCelda.Visible = False
End Sub

Private Sub FlxGd_Validate(Cancel As Boolean)
lbFocoFlex = False
End Sub

Private Sub mnuDeleteGridRow_Click()
   EliminaFila FlxGd.row
End Sub
Private Sub MnuFGridAddRow_Click()
   AdicionaFila
End Sub
Private Sub FlxGd_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lbCancel As Boolean
Dim Control As Object
    
    Select Case KeyCode
        Case vbKeyC And Shift = 2   '   Copiar  [Ctrl+C]
            Clipboard.Clear
            Clipboard.SetText FlxGd.Text
            KeyCode = 0
        Case vbKeyV And Shift = 2   '   Pegar  [Ctrl+V]
            FlxGd.Text = Clipboard.GetText
            KeyCode = 0
    End Select
    
    If m_lbEditarFlex = False Then Exit Sub
    Select Case KeyCode
        Case 46                 '<Del>, clear cell
            Set Control = DeterminaControl(FlxGd.Col)
            If UCase(Control.Name) = "TXTBUSCAR" Then
                Exit Sub
            End If
            'If InStr(1, m_ColumnasAEditar, FlxGd.Col) = 0 Then
            If ColumnaAEditar(FlxGd.Col) = False Then Exit Sub
            Select Case DeterminaFormato(FlxGd.Col)
                Case 2, 3, 4
                    FlxGd = Format("0", EmiteFormatoCol(FlxGd.Col))
                Case Else
                    FlxGd = ""
            End Select
            lbCancel = True
            RaiseEvent OnValidate(FlxGd.row, FlxGd.Col, lbCancel)
            If lbCancel Then
                RaiseEvent OnCellChange(FlxGd.row, FlxGd.Col)
            End If
            
        Case 32
                If FlxGd.Col = lnColChk Then
                    SeleccionaCheck
                End If
                If FlxGd.Col = lnColOpt Then
                    SeleccionaOpt
                End If
        Case 27
            Exit Sub
        Case 113
            EnfocaTexto DeterminaControl(FlxGd.Col), 0, FlxGd
    End Select
    
End Sub
Private Sub FlxGd_KeyPress(KeyAscii As Integer)
 RaiseEvent KeyPress(KeyAscii)
 If m_lbEditarFlex = False Then Exit Sub
    'If InStr(1, m_ColumnasAEditar, FlxGd.Col) = 0 Then
    If ColumnaAEditar(FlxGd.Col) = False Then
        If KeyAscii = 13 Then
             Advance_Cell
        End If
        KeyAscii = 0
        Exit Sub
    End If
    If KeyAscii = 27 Then Exit Sub
    If KeyAscii = 8 Then KeyAscii = 13
    KeyAscii = TeclaFormato(FlxGd.Col, KeyAscii)
    If KeyAscii = 13 Then
        If FlxGd.Col = lnColChk Then
            Advance_Cell
            Exit Sub
        End If
        EnfocaTexto DeterminaControl(FlxGd.Col), 0, FlxGd
    Else
        EnfocaTexto DeterminaControl(FlxGd.Col), KeyAscii, FlxGd
    End If
End Sub
Private Sub FlxGd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim row As Long, Col As Long
    If m_VisiblePopMenu And m_lbEditarFlex Then
        row = FlxGd.MouseRow
        Col = FlxGd.MouseCol
        cboCelda.Visible = False
        txtCelda.Visible = False
        txtFecha.Visible = False
        TxtBuscar.Visible = False
        If Button = 2 And Col = 0 Then  '(Col = 0 Or Row = 0)
            FlxGd.Col = IIf(Col = 0, 1, Col)
            FlxGd.row = IIf(row = 0, 1, row)
            PopupMenu MnuFGridRows
        End If
    End If
End Sub
Public Sub AdicionaFila(Optional nItem As Long = 0, Optional lNumItem As Integer = 0, Optional pbForzar As Boolean = False)
    On Error GoTo AdicionaRowErr
    FlxGd.Redraw = False
    If nItem = 0 Then
       nItem = FlxGd.Rows
    Else
        If nItem > FlxGd.Rows Then
            FlxGd.Redraw = True
            Exit Sub
        End If
    End If
    If pbForzar = False Then
        If FlxGd.Rows = 2 And Len(Trim(FlxGd.TextMatrix(1, 1)) & "") = 0 Then
            nItem = 1
        End If
        If FlxGd.TextMatrix(nItem - 1, 1) = "" And nItem - 1 > 0 Then
            FlxGd.row = nItem - 1
            FlxGd.Col = 1
            If nItem - 1 <> 0 Then
                FlxGd.TopRow = nItem - 1
            End If
            DrawIcons ptArrow, nItem - 1
            FlxGd.Redraw = True
            SendKeys "{DOWN}"
            Exit Sub
        End If
    Else
        If FlxGd.Rows = 2 And Len(Trim(FlxGd.TextMatrix(FlxGd.row, 0))) = 0 Then
            nItem = 1
        End If
    End If
    If lnFilaActual <> -1 Then DrawIcons ptNone, lnFilaActual
    If nItem > 1 Then
        FlxGd.AddItem " ", nItem
    End If
    DrawIcons ptArrow, nItem
    FlxGd.RowHeight(nItem) = 285
    FlxGd.row = nItem
    If lnColChk <> -1 Then
        FlxGd.Col = lnColChk
        FlxGd.CellPictureAlignment = flexAlignCenterCenter
        If FlxGd.ColWidth(lnColChk) > 0 Then
            Set FlxGd.CellPicture = ImgCheck0.Picture
        Else
            Set FlxGd.CellPicture = Nothing
        End If
    End If
    If lnColOpt <> -1 Then
        FlxGd.Col = lnColOpt
        FlxGd.CellPictureAlignment = flexAlignCenterCenter
        If FlxGd.ColWidth(lnColOpt) > 0 Then
            Set FlxGd.CellPicture = ImgOpt0.Picture
        Else
            Set FlxGd.CellPicture = Nothing
        End If
    End If
    FlxGd.Col = 1
    FlxGd.TopRow = nItem
    FlxGd.TextMatrix(nItem, 0) = IIf(lNumItem <> 0, lNumItem, nItem)
    lnFilaActual = FlxGd.row
    RaiseEvent OnRowAdd(FlxGd.row)
    FlxGd.Redraw = True
    If lbFocoFlex Then
        SendKeys "{DOWN}"
    End If
Exit Sub
AdicionaRowErr:
    FlxGd.Redraw = True
    MsgBox " Error N°[" & err.Number & "] " & err.Description, vbInformation, "Aviso"
End Sub
Public Sub EnumeraItems(Optional lbRs As Boolean = False)
Dim nPos As Long
Dim nCol As Long
For nPos = 1 To FlxGd.Rows - 1
    If nPos = 1 And FlxGd.TextMatrix(nPos, 1) = "" Then
        If FlxGd.TextMatrix(nPos, nCol) <> "" Then
            FlxGd.TextMatrix(nPos, nCol) = nPos
        Else
            FlxGd.TextMatrix(nPos, nCol) = ""
        End If
        Exit Sub
    End If
    If lbRs Then
        FlxGd.ColAlignment(0) = 4
        FlxGd.ColAlignmentFixed = 4
    End If
    nCol = 0
    FlxGd.TextMatrix(nPos, nCol) = nPos
Next
End Sub
'##ModelId=3A809921000F
Public Sub EliminaFila(ByVal nItem As Long, Optional pbEnumeraItems As Boolean = True)
    On Error GoTo EliminaRowErr
    Dim nPos As Long
    Dim lnNuevaPos As Long
    nPos = nItem
    txtCelda.Visible = False
    TxtBuscar.Visible = False
    txtFecha.Visible = False
    cboCelda.Visible = False
    If FlxGd.Rows > 2 Then
        FlxGd.RemoveItem nPos
        If nPos > FlxGd.Rows - 1 Then
             lnNuevaPos = nPos - 1 ' FlxGd.Rows - 1
        Else
            If nPos - 1 = 0 Then
                lnNuevaPos = 1
            Else
                lnNuevaPos = nPos
            End If
        End If
        lnFilaActual = lnNuevaPos
        DrawIcons ptArrow, lnFilaActual
    Else
        DrawIcons ptNone, lnFilaActual
        For nPos = 0 To FlxGd.Cols - 1
            FlxGd.TextMatrix(1, nPos) = ""
        Next
        If lnColChk <> -1 Then
            FlxGd.Col = lnColChk
            Set FlxGd.CellPicture = Nothing
        End If
        If lnColOpt <> -1 Then
            FlxGd.Col = lnColOpt
            Set FlxGd.CellPicture = Nothing
        End If
        
        lnFilaActual = 1
    End If
    If pbEnumeraItems Then
        EnumeraItems
    End If
    RaiseEvent OnRowDelete
    Exit Sub
EliminaRowErr:
    FlxGd.Redraw = True
    MsgBox " Error N°[" & err.Number & "] " & err.Description, vbInformation, "Aviso"
End Sub
Private Sub Advance_Cell()                  'advance to next cell
Dim nCol As Integer
    With FlxGd
        DrawIcons ptNone, .row
        If AvanceCeldas = Horizontal Then
            If .Col < .Cols - 1 Then
              nCol = ColVisible(.Col) 'Siguiente Columna Visible
              If nCol < .Col Then  'Columna fue la ultima Visible
                 nCol = .Cols - 1
              End If
            Else
              nCol = .Col
            End If
            If nCol < .Cols - 1 Then
                .Col = nCol
            Else
              If .row < .Rows - 1 Then
                If m_SoloFila = False Then
                    .row = .row + 1                 'down 1 row
                    .Col = 1                        'first column
                Else
                    .Col = 1                        'first column
                End If
              Else
                    If m_AutoAdd And FlxGd.TextMatrix(FlxGd.row, 0) <> "" Then
                        AdicionaFila , Val(FlxGd.TextMatrix(FlxGd.Rows - 1, 0)) + 1
                    Else
                        If m_SoloFila = False Then
                            .row = 1
                            .Col = 1
                        Else
                            .Col = 1
                        End If
                    End If
              End If
            End If
            If .CellTop + .CellHeight > .Top + .Height Then
              .TopRow = .TopRow + 1             'make sure row is visible
            End If
        Else
            If .row < .Rows - 1 Then
                .row = .row + 1
            Else
              If .Col < .Cols - 1 Then
                   .Col = ColVisible(.Col)
                   .row = 1                        'first Fila
              Else
                .row = 1
                .Col = 1
              End If
            End If
            If .CellTop + .CellHeight > .Top + .Height Then
                .TopRow = .TopRow + 1             'make sure row is visible
            End If
        End If
        lnFilaActual = .row
        DrawIcons ptArrow, .row
        RaiseEvent RowColChange
    End With
End Sub
Private Function ColVisible(lnCol As Long) As Long
Dim i As Long
Dim lnColAux As Long
lnColAux = lnCol
For i = lnColAux + 1 To FlxGd.Cols - 1
    If FlxGd.ColWidth(i) > 0 And i > FlxGd.FixedCols Then
        lnColAux = i
        Exit For
    End If
Next
If lnColAux = lnCol Then
    For i = 1 To FlxGd.Cols - 1
        If FlxGd.ColWidth(i) > 0 And i > FlxGd.FixedCols - 1 Then
            lnColAux = i
            Exit For
        End If
    Next
End If
ColVisible = lnColAux
End Function


Private Sub TxtBuscar_EmiteDatos()
TxtBuscar.Visible = False
If m_lbFlexDuplicados = False Then
    If BuscaDuplicado(FlxGd.row, FlxGd.Col, TxtBuscar.Text) Then
        Advance_Cell
        RaiseEvent OnEnterTextBuscar(TxtBuscar.Text, TxtBuscar.Tag, FlxGd.Col, True)
        FlxGd.SetFocus
        Exit Sub
    End If
End If
If FlxGd.TextMatrix(Val(TxtBuscar.Tag), FlxGd.Col) <> TxtBuscar.Text Then
    FlxGd.TextMatrix(Val(TxtBuscar.Tag), FlxGd.Col) = TxtBuscar.Text
    RaiseEvent OnCellChange(Val(TxtBuscar.Tag), FlxGd.Col)
End If
If FlxGd.Col + 1 < FlxGd.Cols Then
    If ColumnaAEditar(FlxGd.Col + 1) = False Then
        FlxGd.TextMatrix(Val(TxtBuscar.Tag), FlxGd.Col + 1) = TxtBuscar.psDescripcion
    End If
End If
RaiseEvent OnEnterTextBuscar(TxtBuscar.Text, TxtBuscar.Tag, FlxGd.Col, False)
TxtBuscar.Text = ""
TxtBuscar.psDescripcion = ""
Advance_Cell
FlxGd.SetFocus

End Sub
Private Sub TxtBuscar_LostFocus()
'TxtBuscar.Visible = False
If TipoBusqueda <> BuscaLibre Then
    TxtBuscar.Text = ""
    TxtBuscar.psDescripcion = ""
End If
FlxGd.SetFocus
End Sub

Private Sub txtCelda_KeyPress(KeyAscii As Integer)
Dim lbCancel As Boolean
Dim lsAux As String
Dim lnTpoFormato As Long

KeyAscii = TeclaFormato(FlxGd.Col, KeyAscii)
If KeyAscii = 13 Then
    lnTpoFormato = DeterminaFormato(FlxGd.Col)
    If lnTpoFormato = 4 Then
       txtCelda.Text = CalEvaluaBil(txtCelda.Text)
       If Val(txtCelda.Text) = 0 Then
            MuestraFormato FlxGd.Col
            FlxGd.SetFocus
            Exit Sub
        End If
    End If
    MuestraFormato FlxGd.Col
    If m_lbBuscaDuplicadoText = False Then
        If BuscaDuplicado(FlxGd.row, FlxGd.Col, txtCelda.Text) Then
            FlxGd.SetFocus
            Exit Sub
        End If
    End If
        lbCancel = True
        lsAux = FlxGd.Text
        FlxGd.Text = txtCelda.Text
        RaiseEvent OnValidate(FlxGd.row, FlxGd.Col, lbCancel)
        If lbCancel = True Then
            RaiseEvent OnCellChange(FlxGd.row, FlxGd.Col)
        Else
            FlxGd.Text = lsAux
            Exit Sub
        End If
    Advance_Cell
    FlxGd.SetFocus
End If
End Sub
Private Function TeclaFormato(ByVal pnCol As Long, KeyAscii As Integer) As Integer
Dim lnTpoFormato As Long
lnTpoFormato = DeterminaFormato(pnCol)
Select Case lnTpoFormato
    Case 1
        TeclaFormato = SoloLetras(KeyAscii)
    Case 2, 5
        TeclaFormato = NumerosDecimales(txtCelda, KeyAscii, m_CantEntero, m_CantDecimales)
    Case 3
        TeclaFormato = NumerosEnteros(KeyAscii)
    Case 4
        If KeyAscii <> 13 Then
            TeclaFormato = ValidaIngBilletaje(KeyAscii)
        Else
            TeclaFormato = KeyAscii
        End If
    Case Else
        TeclaFormato = KeyAscii
End Select
End Function
Private Sub MuestraFormato(pnCol As Long)
Dim lnTpoFormato As Long
lnTpoFormato = DeterminaFormato(FlxGd.Col)
Select Case lnTpoFormato
    Case 2, 4
        txtCelda = Format(IIf(txtCelda = "", 0, txtCelda), "###," & String(m_CantEntero, "#") & "#0." & String(m_CantDecimales, "0"))
    Case 3
        txtCelda = Format(IIf(txtCelda = "", 0, txtCelda), "#," & String(m_CantEntero, "#") & "#0")
End Select
End Sub
Private Sub txtCelda_LostFocus()
txtCelda.Text = ""
txtCelda.Visible = False
If FlxGd.Visible And FlxGd.Enabled Then
    FlxGd.SetFocus
End If
End Sub
Private Sub txtFecha_KeyPress(KeyAscii As Integer)
Dim lbCancel  As Boolean
Dim lsAux As String
If KeyAscii = 13 Then
    If txtFecha.Mask <> "##:##:##" Then
        If ValFecha(txtFecha) = False Then Exit Sub
    Else
        If IsDate(txtFecha) = False Then
            MsgBox "Hora no válida", vbInformation, "Aviso"
            Exit Sub
        End If
    End If
    If m_lbBuscaDuplicadoText = False Then
        If BuscaDuplicado(FlxGd.row, FlxGd.Col, txtFecha.Text) Then
            FlxGd.SetFocus
            Exit Sub
        End If
    End If
    If FlxGd.Text <> txtFecha.Text Then
        lbCancel = True
        lsAux = FlxGd.Text
        FlxGd.Text = txtFecha.Text
        RaiseEvent OnValidate(FlxGd.row, FlxGd.Col, lbCancel)
        If lbCancel = True Then
            RaiseEvent OnCellChange(FlxGd.row, FlxGd.Col)
        Else
            FlxGd.Text = lsAux
            Exit Sub
        End If
    End If
    Advance_Cell
    FlxGd.SetFocus
End If
End Sub
Private Sub txtFecha_LostFocus()
If txtFecha.Mask = "##/##/#### ##:##:##" Then
    txtFecha.Text = "  /  /       :  :  "
ElseIf txtFecha.Mask = "##:##:##" Then
        txtFecha.Mask = "  :  :  "
    ElseIf txtFecha.Mask = "##/##/####" Then
        txtFecha.Text = "  /  /    "
    End If
txtFecha.Visible = False
FlxGd.SetFocus
End Sub

Private Sub UserControl_GotFocus()
lbFocoFlex = True
End Sub
Private Sub UserControl_Initialize()
cboCelda.Visible = False
txtCelda.Visible = False
txtFecha.Visible = False
TxtBuscar.Visible = False
TxtBuscar.EditFlex = True
lnFilaActual = -1
m_iSortCol = 1
TxtBuscar.Appearance = flat
txtCelda.Appearance = 0
End Sub

Private Sub UserControl_LostFocus()
lbFocoFlex = False
End Sub

Private Sub UserControl_Resize()
FlxGd.Move 0, 0, UserControl.Width - 10, UserControl.Height - 10
End Sub
'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=FlxGd,FlxGd,-1,Rows
Public Property Get Rows() As Long
Attribute Rows.VB_Description = "Determines the total number of columns or rows in the Hierarchical FlexGrid."
    Rows = FlxGd.Rows
End Property
Public Property Let Rows(ByVal New_Rows As Long)
    FlxGd.Rows() = New_Rows
    PropertyChanged "Rows"
End Property
'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=FlxGd,FlxGd,-1,FixedCols
Public Property Get FixedCols() As Long
Attribute FixedCols.VB_Description = "Returns or sets the total number of fixed (non-scrollable) columns or rows for a Hierarchical FlexGrid."
Attribute FixedCols.VB_ProcData.VB_Invoke_Property = "PageColumnas"
Attribute FixedCols.VB_MemberFlags = "400"
    FixedCols = FlxGd.FixedCols
End Property
Public Property Let FixedCols(ByVal New_FixedCols As Long)
    FlxGd.FixedCols() = New_FixedCols
    PropertyChanged "FixedCols"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=FlxGd,FlxGd,-1,FixedRows
Public Property Get FixedRows() As Long
Attribute FixedRows.VB_Description = "Returns or sets the total number of fixed (non-scrollable) columns or rows for a Hierarchical FlexGrid."
Attribute FixedRows.VB_ProcData.VB_Invoke_Property = "FlexPagePropiedades"
Attribute FixedRows.VB_MemberFlags = "400"
    FixedRows = FlxGd.FixedRows
End Property

Public Property Let FixedRows(ByVal New_FixedRows As Long)
    FlxGd.FixedRows() = New_FixedRows
    PropertyChanged "FixedRows"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=FlxGd,FlxGd,-1,ScrollBars
Public Property Get ScrollBars() As ScrollBarsSettings
Attribute ScrollBars.VB_Description = "Returns or sets whether the Hierarchical FlexGrid has horizontal or vertical scroll bars."
    ScrollBars = FlxGd.ScrollBars
End Property

Public Property Let ScrollBars(ByVal New_ScrollBars As ScrollBarsSettings)
    FlxGd.ScrollBars() = New_ScrollBars
    PropertyChanged "ScrollBars"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=FlxGd,FlxGd,-1,HighLight
Public Property Get HighLight() As HighLightSettings
Attribute HighLight.VB_Description = "Returns or sets whether selected cells appear highlighted."
    HighLight = FlxGd.HighLight
End Property

Public Property Let HighLight(ByVal New_HighLight As HighLightSettings)
    FlxGd.HighLight() = New_HighLight
    PropertyChanged "HighLight"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=FlxGd,FlxGd,-1,FocusRect
Public Property Get FocusRect() As FocusRectSettings
Attribute FocusRect.VB_Description = "Determines whether the Hierarchical FlexGrid control should draw a focus rectangle around the current cell."
    FocusRect = FlxGd.FocusRect
End Property

Public Property Let FocusRect(ByVal New_FocusRect As FocusRectSettings)
    FlxGd.FocusRect() = New_FocusRect
    PropertyChanged "FocusRect"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=FlxGd,FlxGd,-1,BandDisplay
Public Property Get BandDisplay() As BandDisplaySettings
Attribute BandDisplay.VB_Description = "Returns or sets the band display style."
    BandDisplay = FlxGd.BandDisplay
End Property

Public Property Let BandDisplay(ByVal New_BandDisplay As BandDisplaySettings)
    FlxGd.BandDisplay() = New_BandDisplay
    PropertyChanged "BandDisplay"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=FlxGd,FlxGd,-1,MousePointer
Public Property Get MousePointer() As MousePointerSettings
Attribute MousePointer.VB_Description = "Returns or sets the type of mouse pointer displayed when over part of an object."
    MousePointer = FlxGd.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerSettings)
    FlxGd.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=FlxGd,FlxGd,-1,FillStyle
Public Property Get FillStyle() As FillStyleSettings
Attribute FillStyle.VB_Description = "Determines whether setting the Text property or one of the cell formatting properties of a Hierarchical FlexGrid applies the change to all selected cells."
    FillStyle = FlxGd.FillStyle
End Property

Public Property Let FillStyle(ByVal New_FillStyle As FillStyleSettings)
    FlxGd.FillStyle() = New_FillStyle
    PropertyChanged "FillStyle"
End Property
'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=7,0,0,0
Public Property Get ScaleMode() As Integer
Attribute ScaleMode.VB_Description = "Devuelve o establece un valor que indica las unidades de medida de las coordenadas de un objeto al usar métodos gráficos o colocar controles."
Attribute ScaleMode.VB_ProcData.VB_Invoke_Property = "FlexPagePropiedades"
    ScaleMode = m_ScaleMode
End Property
Public Property Let ScaleMode(ByVal New_ScaleMode As Integer)
    m_ScaleMode = New_ScaleMode
    PropertyChanged "ScaleMode"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=FlxGd,FlxGd,-1,AllowUserResizing
Public Property Get AllowUserResizing() As AllowUserResizeSettings
Attribute AllowUserResizing.VB_Description = "Returns or sets whether the user is allowed to resize rows and columns with the mouse."
    AllowUserResizing = FlxGd.AllowUserResizing
End Property

Public Property Let AllowUserResizing(ByVal New_AllowUserResizing As AllowUserResizeSettings)
    FlxGd.AllowUserResizing() = New_AllowUserResizing
    PropertyChanged "AllowUserResizing"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=FlxGd,FlxGd,-1,RowSizingMode
Public Property Get RowSizingMode() As RowSizingSettings
Attribute RowSizingMode.VB_Description = "Returns or sets the row sizing mode."
    RowSizingMode = FlxGd.RowSizingMode
End Property
Public Property Let RowSizingMode(ByVal New_RowSizingMode As RowSizingSettings)
    FlxGd.RowSizingMode() = New_RowSizingMode
    PropertyChanged "RowSizingMode"
End Property
'Inicializar propiedades para control de usuario
Private Sub UserControl_InitProperties()
    m_ScaleMode = m_def_ScaleMode
    m_VisiblePopMenu = m_def_VisiblePopMenu
    m_EncabezadosNombres = m_def_EncabezadosNombres
    m_EncabezadosAnchos = m_def_EncabezadosAnchos
    m_ColumnasAEditar = m_def_ColumnasAEditar
    m_ListaControles = m_def_ListaControles
    m_EncabezadosAlineacion = m_def_EncabezadosAlineacion
    m_FormatosEdit = m_def_FormatosEdit
    m_CantEntero = m_def_CantEntero
    m_CantDecimales = m_def_CantDecimales
    m_AvanceCeldas = m_def_AvanceCeldas
    m_SQLDataSource = m_def_SQLDataSource
    m_lbRsLoad = m_def_lbRsLoad
    m_lbEditarFlex = m_def_lbEditarFlex
    m_lbFlexDuplicados = m_def_lbFlexDuplicados
    m_lbFormatoCol = m_def_lbFormatoCol
    m_lbPuntero = m_def_lbPuntero
    m_lbOrdenaCol = m_def_lbOrdenaCol
    m_lbBuscaDuplicadoText = m_def_lbBuscaDuplicadoText
    m_AutoAdd = m_def_AutoAdd
    m_SoloFila = m_def_SoloFila
End Sub
'Cargar valores de propiedad desde el almacén
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Dim index As Integer
    FlxGd.Rows = PropBag.ReadProperty("Rows", 2)
    FlxGd.Cols(0) = PropBag.ReadProperty("Cols" & index, 2)
    FlxGd.FixedCols = PropBag.ReadProperty("FixedCols", 1)
    FlxGd.FixedRows = PropBag.ReadProperty("FixedRows", 1)
    FlxGd.row = PropBag.ReadProperty("Row", 1)
    FlxGd.Col = PropBag.ReadProperty("Col", 1)
    FlxGd.ScrollBars = PropBag.ReadProperty("ScrollBars", 3)
    FlxGd.HighLight = PropBag.ReadProperty("HighLight", 0)
    FlxGd.FocusRect = PropBag.ReadProperty("FocusRect", 0)
    FlxGd.BandDisplay = PropBag.ReadProperty("BandDisplay", 0)
    FlxGd.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    FlxGd.FillStyle = PropBag.ReadProperty("FillStyle", 0)
    m_ScaleMode = PropBag.ReadProperty("ScaleMode", m_def_ScaleMode)
    FlxGd.AllowUserResizing = PropBag.ReadProperty("AllowUserResizing", 0)
    FlxGd.RowSizingMode = PropBag.ReadProperty("RowSizingMode", 0)
    m_VisiblePopMenu = PropBag.ReadProperty("VisiblePopMenu", m_def_VisiblePopMenu)
    m_EncabezadosNombres = PropBag.ReadProperty("EncabezadosNombres", m_def_EncabezadosNombres)
    m_EncabezadosAnchos = PropBag.ReadProperty("EncabezadosAnchos", m_def_EncabezadosAnchos)
    
    '***********************FUENTES DE CONTROLES ********************
     Set FlxGd.Font = PropBag.ReadProperty("Font", Ambient.Font)
     Set txtCelda.Font = PropBag.ReadProperty("Font", Ambient.Font)
     Set TxtBuscar.Font = PropBag.ReadProperty("Font", Ambient.Font)
     Set txtFecha.Font = PropBag.ReadProperty("Font", Ambient.Font)
     Set cboCelda.Font = PropBag.ReadProperty("Font", Ambient.Font)
     Set FlxGd.FontFixed = PropBag.ReadProperty("FontFixed", Ambient.Font)
        
    TxtBuscar.QuerySeek = PropBag.ReadProperty("QuerySeek", "")
    TxtBuscar.psRaiz = PropBag.ReadProperty("psRaiz", "")
    TxtBuscar.BackColor = PropBag.ReadProperty("BackColorControl", &H80000018)
    txtCelda.BackColor = PropBag.ReadProperty("BackColorControl", &H80000018)
    txtFecha.BackColor = PropBag.ReadProperty("BackColorControl", &H80000018)
    TxtBuscar.lbUltimaInstancia = PropBag.ReadProperty("lbUltimaInstancia", Verdadero)
    TxtBuscar.TipoBusqueda = PropBag.ReadProperty("TipoBusqueda", 1)
    
    m_ColumnasAEditar = PropBag.ReadProperty("ColumnasAEditar", m_def_ColumnasAEditar)
    FlxGd.ToolTipText = PropBag.ReadProperty("ToolTipText", "")
    FlxGd.TextStyleFixed = PropBag.ReadProperty("TextStyleFixed", 0)
    FlxGd.TextStyle = PropBag.ReadProperty("TextStyle", 0)
    m_ListaControles = PropBag.ReadProperty("ListaControles", m_def_ListaControles)
    FlxGd.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    m_EncabezadosAlineacion = PropBag.ReadProperty("EncabezadosAlineacion", m_def_EncabezadosAlineacion)
    m_FormatosEdit = PropBag.ReadProperty("FormatosEdit", m_def_FormatosEdit)
    m_CantEntero = PropBag.ReadProperty("CantEntero", m_def_CantEntero)
    m_CantDecimales = PropBag.ReadProperty("CantDecimales", m_def_CantDecimales)
    m_AvanceCeldas = PropBag.ReadProperty("AvanceCeldas", m_def_AvanceCeldas)
    m_SQLDataSource = PropBag.ReadProperty("SQLDataSource", m_def_SQLDataSource)
    FlxGd.TextMatrix(row, Col) = PropBag.ReadProperty("TextMatrix" & index, "")
    FlxGd.TextArray(index) = PropBag.ReadProperty("TextArray" & index, "")
    Set m_rsFlex = rsFlex
    m_lbRsLoad = PropBag.ReadProperty("lbRsLoad", m_def_lbRsLoad)
    FlxGd.SelectionMode = PropBag.ReadProperty("SelectionMode", 0)
    m_lbEditarFlex = PropBag.ReadProperty("lbEditarFlex", m_def_lbEditarFlex)
    FlxGd.Enabled = PropBag.ReadProperty("Enabled", True)
    m_lbFlexDuplicados = PropBag.ReadProperty("lbFlexDuplicados", m_def_lbFlexDuplicados)
    m_lbFormatoCol = PropBag.ReadProperty("lbFormatoCol", m_def_lbFormatoCol)
    cboCelda.List(index) = PropBag.ReadProperty("List" & index, "")
    m_lbPuntero = PropBag.ReadProperty("lbPuntero", m_def_lbPuntero)
    m_lbOrdenaCol = PropBag.ReadProperty("lbOrdenaCol", m_def_lbOrdenaCol)
    m_lbBuscaDuplicadoText = PropBag.ReadProperty("lbBuscaDuplicadoText", m_def_lbBuscaDuplicadoText)
    Set Recordset = PropBag.ReadProperty("Recordset", Nothing)
    txtCelda.MaxLength = PropBag.ReadProperty("MaxLength", 0)
    MnuFGridAddRow.Enabled = PropBag.ReadProperty("EnabledMnuAdd", True)
    mnuDeleteGridRow.Enabled = PropBag.ReadProperty("EnabledMnuEliminar", True)
    FlxGd.Appearance = PropBag.ReadProperty("Appearance", 1)
    FlxGd.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    '*************************************SHOW **********************
    FormaCabecera
    If Len(Trim(m_SQLDataSource)) > 0 Then
        GeneraGridDatos
    End If
    lnColChk = DeterminaColChek
    lnColOpt = DeterminaColOption
    TxtBuscar.psDH = PropBag.ReadProperty("psDH", "")
    FlxGd.ColWidth(index) = PropBag.ReadProperty("ColWidth" & index, 0)
    FlxGd.RowHeight(index) = PropBag.ReadProperty("RowHeight" & index, -1)
    TxtBuscar.TipoBusPers = PropBag.ReadProperty("TipoBusPersona", 0)
    m_AutoAdd = PropBag.ReadProperty("AutoAdd", m_def_AutoAdd)
    FlxGd.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    FlxGd.ForeColorFixed = PropBag.ReadProperty("ForeColorFixed", &H80000008)
    FlxGd.CellForeColor = PropBag.ReadProperty("CellForeColor", &H80000008)
    FlxGd.CellBackColor = PropBag.ReadProperty("CellBackColor", &H80000005)
    TxtBuscar.PersPersoneria = PropBag.ReadProperty("PersPersoneria", gPersonaNat)
    m_SoloFila = PropBag.ReadProperty("SoloFila", m_def_SoloFila)
    FlxGd.RowHeightMin = PropBag.ReadProperty("RowHeightMin", 0)
End Sub
'Escribir valores de propiedad en el almacén
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Dim index As Integer
    Call PropBag.WriteProperty("Rows", FlxGd.Rows, 2)
    Call PropBag.WriteProperty("Cols" & index, FlxGd.Cols(0), 2)
    Call PropBag.WriteProperty("FixedCols", FlxGd.FixedCols, 1)
    Call PropBag.WriteProperty("FixedRows", FlxGd.FixedRows, 1)
    Call PropBag.WriteProperty("Row", FlxGd.row, 1)
    Call PropBag.WriteProperty("Col", FlxGd.Col, 1)
    Call PropBag.WriteProperty("ScrollBars", FlxGd.ScrollBars, 3)
    Call PropBag.WriteProperty("HighLight", FlxGd.HighLight, 0)
    Call PropBag.WriteProperty("FocusRect", FlxGd.FocusRect, 0)
    Call PropBag.WriteProperty("BandDisplay", FlxGd.BandDisplay, 0)
    Call PropBag.WriteProperty("MousePointer", FlxGd.MousePointer, 0)
    Call PropBag.WriteProperty("FillStyle", FlxGd.FillStyle, 0)
    Call PropBag.WriteProperty("ScaleMode", m_ScaleMode, m_def_ScaleMode)
    Call PropBag.WriteProperty("AllowUserResizing", FlxGd.AllowUserResizing, 0)
    Call PropBag.WriteProperty("RowSizingMode", FlxGd.RowSizingMode, 0)
    Call PropBag.WriteProperty("VisiblePopMenu", m_VisiblePopMenu, m_def_VisiblePopMenu)
    Call PropBag.WriteProperty("EncabezadosNombres", m_EncabezadosNombres, m_def_EncabezadosNombres)
    Call PropBag.WriteProperty("EncabezadosAnchos", m_EncabezadosAnchos, m_def_EncabezadosAnchos)
    
    '************************** FUENTES DE LAS LETRAS DE LOS CONTROLES **********
    Call PropBag.WriteProperty("Font", FlxGd.Font, Ambient.Font)
    Call PropBag.WriteProperty("FontFixed", FlxGd.FontFixed, Ambient.Font)
    
    Call PropBag.WriteProperty("ColumnasAEditar", m_ColumnasAEditar, m_def_ColumnasAEditar)
    Call PropBag.WriteProperty("ToolTipText", FlxGd.ToolTipText, "")
    Call PropBag.WriteProperty("TextStyleFixed", FlxGd.TextStyleFixed, 0)
    Call PropBag.WriteProperty("TextStyle", FlxGd.TextStyle, 0)
    Call PropBag.WriteProperty("ListaControles", m_ListaControles, m_def_ListaControles)
    Call PropBag.WriteProperty("QuerySeek", TxtBuscar.QuerySeek, "")
    Call PropBag.WriteProperty("psRaiz", TxtBuscar.psRaiz, "")
    Call PropBag.WriteProperty("rsTextBuscar", TxtBuscar.rs, 0)
    Call PropBag.WriteProperty("BackColor", FlxGd.BackColor, &H80000005)
    Call PropBag.WriteProperty("BackColorControl", TxtBuscar.BackColor, &H80000018)
    Call PropBag.WriteProperty("BackColorControl", txtCelda.BackColor, &H80000018)
    Call PropBag.WriteProperty("BackColorControl", txtFecha.BackColor, &H80000018)
    Call PropBag.WriteProperty("EncabezadosAlineacion", m_EncabezadosAlineacion, m_def_EncabezadosAlineacion)
    Call PropBag.WriteProperty("FormatosEdit", m_FormatosEdit, m_def_FormatosEdit)
    Call PropBag.WriteProperty("CantEntero", m_CantEntero, m_def_CantEntero)
    Call PropBag.WriteProperty("CantDecimales", m_CantDecimales, m_def_CantDecimales)
    Call PropBag.WriteProperty("AvanceCeldas", m_AvanceCeldas, m_def_AvanceCeldas)
    Call PropBag.WriteProperty("SQLDataSource", m_SQLDataSource, m_def_SQLDataSource)
    Call PropBag.WriteProperty("TextMatrix" & index, FlxGd.TextMatrix(row, Col), "")
    Call PropBag.WriteProperty("TextArray" & index, FlxGd.TextArray(index), "")
    Call PropBag.WriteProperty("rsFlex", m_rsFlex, m_def_rsFlex)
    Call PropBag.WriteProperty("lbRsLoad", m_lbRsLoad, m_def_lbRsLoad)
    Call PropBag.WriteProperty("SelectionMode", FlxGd.SelectionMode, 0)
    Call PropBag.WriteProperty("lbEditarFlex", m_lbEditarFlex, m_def_lbEditarFlex)
    Call PropBag.WriteProperty("Enabled", FlxGd.Enabled, True)
    Call PropBag.WriteProperty("lbFlexDuplicados", m_lbFlexDuplicados, m_def_lbFlexDuplicados)
    Call PropBag.WriteProperty("lbUltimaInstancia", TxtBuscar.lbUltimaInstancia, Verdadero)
    Call PropBag.WriteProperty("TipoBusqueda", TxtBuscar.TipoBusqueda, 1)
    Call PropBag.WriteProperty("lbFormatoCol", m_lbFormatoCol, m_def_lbFormatoCol)
    Call PropBag.WriteProperty("List" & index, cboCelda.List(index), "")
    
    Call PropBag.WriteProperty("lbPuntero", m_lbPuntero, m_def_lbPuntero)
    Call PropBag.WriteProperty("lbOrdenaCol", m_lbOrdenaCol, m_def_lbOrdenaCol)
    Call PropBag.WriteProperty("lbBuscaDuplicadoText", m_lbBuscaDuplicadoText, m_def_lbBuscaDuplicadoText)
    Call PropBag.WriteProperty("Recordset", Recordset, Nothing)
    Call PropBag.WriteProperty("MaxLength", txtCelda.MaxLength, 0)
    Call PropBag.WriteProperty("EnabledMnuAdd", MnuFGridAddRow.Enabled, True)
    Call PropBag.WriteProperty("EnabledMnuEliminar", mnuDeleteGridRow.Enabled, True)
    Call PropBag.WriteProperty("Appearance", FlxGd.Appearance, 1)
    Call PropBag.WriteProperty("BorderStyle", FlxGd.BorderStyle, 1)
    Call PropBag.WriteProperty("psDH", TxtBuscar.psDH, "")
    Call PropBag.WriteProperty("ColWidth" & index, FlxGd.ColWidth(index), 0)
    Call PropBag.WriteProperty("RowHeight" & index, FlxGd.RowHeight(index), -1)
    Call PropBag.WriteProperty("TipoBusPersona", TxtBuscar.TipoBusPers, 0)
    Call PropBag.WriteProperty("AutoAdd", m_AutoAdd, m_def_AutoAdd)
    Call PropBag.WriteProperty("ForeColor", FlxGd.ForeColor, &H80000008)
    Call PropBag.WriteProperty("ForeColorFixed", FlxGd.ForeColorFixed, &H80000008)
    Call PropBag.WriteProperty("CellForeColor", FlxGd.CellForeColor, &H80000008)
    Call PropBag.WriteProperty("CellBackColor", FlxGd.CellBackColor, &H80000005)
    Call PropBag.WriteProperty("PersPersoneria", TxtBuscar.PersPersoneria, gPersonaNat)
    Call PropBag.WriteProperty("SoloFila", m_SoloFila, m_def_SoloFila)
    Call PropBag.WriteProperty("RowHeightMin", FlxGd.RowHeightMin, 0)
End Sub
'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=0,0,0,False
Public Property Get VisiblePopMenu() As Boolean
Attribute VisiblePopMenu.VB_ProcData.VB_Invoke_Property = "FlexPagePropiedades"
    VisiblePopMenu = m_VisiblePopMenu
End Property

Public Property Let VisiblePopMenu(ByVal New_VisiblePopMenu As Boolean)
    m_VisiblePopMenu = New_VisiblePopMenu
    PropertyChanged "VisiblePopMenu"
End Property
'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=13,0,0,""
Public Property Get EncabezadosNombres() As String
    EncabezadosNombres = m_EncabezadosNombres
End Property

Public Property Let EncabezadosNombres(ByVal New_EncabezadosNombres As String)
    m_EncabezadosNombres = New_EncabezadosNombres
    PropertyChanged "EncabezadosNombres"
    FormaCabecera
End Property
'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=13,0,0,""
Public Property Get EncabezadosAnchos() As String
    EncabezadosAnchos = m_EncabezadosAnchos
End Property
Public Property Let EncabezadosAnchos(ByVal New_EncabezadosAnchos As String)
    m_EncabezadosAnchos = New_EncabezadosAnchos
    PropertyChanged "EncabezadosAnchos"
    FormaCabecera
End Property
'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=13,0,0,
Public Property Get ColumnasAEditar() As String
    ColumnasAEditar = m_ColumnasAEditar
End Property
Public Property Let ColumnasAEditar(ByVal New_ColumnasAEditar As String)
    m_ColumnasAEditar = New_ColumnasAEditar
    PropertyChanged "ColumnasAEditar"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=FlxGd,FlxGd,-1,ToolTipText
Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_Description = "Devuelve o establece el texto mostrado cuando el mouse se sitúa sobre un control."
    ToolTipText = FlxGd.ToolTipText
End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
    FlxGd.ToolTipText() = New_ToolTipText
    PropertyChanged "ToolTipText"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=FlxGd,FlxGd,-1,TextStyleFixed
Public Property Get TextStyleFixed() As TextStyleSettings
Attribute TextStyleFixed.VB_Description = "Returns or sets 3-D effects for displaying text."
    TextStyleFixed = FlxGd.TextStyleFixed
End Property

Public Property Let TextStyleFixed(ByVal New_TextStyleFixed As TextStyleSettings)
    FlxGd.TextStyleFixed() = New_TextStyleFixed
    PropertyChanged "TextStyleFixed"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=FlxGd,FlxGd,-1,TextStyle
Public Property Get TextStyle() As TextStyleSettings
Attribute TextStyle.VB_Description = "Returns or sets 3-D effects for displaying text."
    TextStyle = FlxGd.TextStyle
End Property

Public Property Let TextStyle(ByVal New_TextStyle As TextStyleSettings)
    FlxGd.TextStyle() = New_TextStyle
    PropertyChanged "TextStyle"
End Property
Public Sub BackColorRow(psColor As ColorConstants, Optional pbBold As Boolean = False)
Dim n As Long
Dim nCol As Long
nCol = FlxGd.Col
For n = 1 To FlxGd.Cols - 1
   FlxGd.Col = n
   FlxGd.CellBackColor = psColor
   FlxGd.CellFontBold = pbBold
Next
FlxGd.Col = nCol
End Sub
Public Sub ForeColorRow(psColor As ColorConstants, Optional pbBold As Boolean = False)
Dim n As Long
Dim nCol As Long
nCol = FlxGd.Col
For n = 1 To FlxGd.Cols - 1
   FlxGd.Col = n
   FlxGd.CellForeColor = psColor
   FlxGd.CellFontBold = pbBold
Next
FlxGd.Col = nCol
End Sub
Private Function NumColsCabecera() As Long
Dim pEncabezado As String
Dim vPos As Long
Dim lnNumCol As Long
pEncabezado = m_EncabezadosNombres
lnNumCol = 0
vPos = 1
If m_EncabezadosNombres = "" Then
    NumColsCabecera = FlxGd.Cols
    Exit Function
End If
Do While vPos > 0
    vPos = InStr(1, pEncabezado, "-", vbTextCompare)
    If vPos > 0 Then
        lnNumCol = lnNumCol + 1
    Else
        If pEncabezado <> "" Then
            lnNumCol = lnNumCol + 1
        End If
    End If
    If vPos > 0 Then
        pEncabezado = Mid(pEncabezado, IIf(vPos > 0, vPos + 1, Len(pEncabezado)))
    Else
        pEncabezado = ""
        Exit Do
    End If
Loop
If lnNumCol = 0 Then lnNumCol = 2
If FlxGd.Cols > lnNumCol Then
    lnNumCol = FlxGd.Cols
End If
NumColsCabecera = lnNumCol
End Function

Public Sub FormaCabecera()
'Optional pbRs As Boolean = False
Dim X As Long, vPos As Long
Dim vAli As String
Dim pEncabezado As String
Dim pAnchoCol As String
Dim pAlineaCol As String
Dim pCol As Long
Dim lsAncho As String

FlxGd.Redraw = False
pEncabezado = m_EncabezadosNombres
pAnchoCol = m_EncabezadosAnchos
pAlineaCol = m_EncabezadosAlineacion
With FlxGd
    Cols = NumColsCabecera
    If m_lbRsLoad = False Then
        .Clear
        .Rows = Rows
        .Cols = Cols
    End If
    pCol = Cols
    If Len(Trim(pEncabezado)) > 0 Then
        vPos = 1
        X = 0
        Do While vPos > 0
            vPos = InStr(1, pEncabezado, "-", vbTextCompare)
            If X < .Cols Then
                .TextMatrix(0, X) = Mid(pEncabezado, 1, IIf(vPos > 0, vPos - 1, Len(pEncabezado)))
            End If
            If vPos > 0 Then
                pEncabezado = Mid(pEncabezado, IIf(vPos > 0, vPos + 1, Len(pEncabezado)))
                X = X + 1
            Else
                pEncabezado = ""
                Exit Do
            End If
        Loop
    End If
    If Len(Trim(pAnchoCol)) > 0 Then
        vPos = 1
        X = 0
        Do While vPos > 0
            vPos = InStr(1, pAnchoCol, "-", vbTextCompare)
            If X < .Cols Then
                .ColWidth(X) = Mid(pAnchoCol, 1, IIf(vPos > 0, vPos - 1, Len(pAnchoCol)))
            End If
            If vPos > 0 Then
                pAnchoCol = Mid(pAnchoCol, IIf(vPos > 0, vPos + 1, Len(pAnchoCol)))
                X = X + 1
            Else
                pAnchoCol = ""
                Exit Do
            End If
        Loop
    End If
    
    For X = 0 To pCol - 1
        .ColAlignmentFixed(X) = 4
        vPos = InStr(1, pAlineaCol, "-", vbTextCompare)
        vAli = UCase(Mid(pAlineaCol, 1, IIf(vPos > 0, vPos - 1, Len(pAlineaCol))))
        If Len(vAli) > 0 And (vAli = "L" Or vAli = "R" Or vAli = "C") Then
            .ColAlignment(X) = Switch(vAli = "L", 1, vAli = "R", 7, vAli = "C", 4)
            pAlineaCol = Mid(pAlineaCol, IIf(vPos > 0, vPos + 1, Len(pAlineaCol)))
        Else
            .ColAlignment(X) = 4
        End If
    Next X
    .row = IIf(.Rows <= 1, 0, 1)
    .Col = IIf(.Cols <= 1, 0, 1)
    .RowHeight(-1) = 300
End With
FlxGd.Redraw = True
End Sub
'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=FlxGd,FlxGd,-1,Cols
Public Property Get Cols() As Long
Attribute Cols.VB_Description = "Determines the total number of columns or rows in the Hierarchical FlexGrid."
    Cols = FlxGd.Cols(0)
End Property

Public Property Let Cols(ByVal New_Cols As Long)
    FlxGd.Cols(0) = New_Cols
    PropertyChanged "Cols"
End Property
'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=13,0,0,
Public Property Get ListaControles() As String
    ListaControles = m_ListaControles
End Property

Public Property Let ListaControles(ByVal New_ListaControles As String)
    m_ListaControles = New_ListaControles
    PropertyChanged "ListaControles"
    lnColChk = DeterminaColChek
    lnColOpt = DeterminaColOption
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=txtBuscar,txtBuscar,-1,QuerySeek
Public Property Get QuerySeek() As String
Attribute QuerySeek.VB_MemberFlags = "440"
    QuerySeek = TxtBuscar.QuerySeek
End Property

Public Property Let QuerySeek(ByVal New_QuerySeek As String)
    TxtBuscar.QuerySeek() = New_QuerySeek
    PropertyChanged "QuerySeek"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=txtBuscar,txtBuscar,-1,psRaiz
Public Property Get psRaiz() As String
    psRaiz = TxtBuscar.psRaiz
End Property

Public Property Let psRaiz(ByVal New_psRaiz As String)
    TxtBuscar.psRaiz() = New_psRaiz
    PropertyChanged "psRaiz"
End Property
'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=txtBuscar,txtBuscar,-1,rs
Public Property Get rsTextBuscar() As ADODB.Recordset
  Set rsTextBuscar = TxtBuscar.rs
End Property
Public Property Let rsTextBuscar(ByVal New_rsTextBuscar As ADODB.Recordset)
    TxtBuscar.rs = New_rsTextBuscar
    PropertyChanged "rsTextBuscar"
End Property
'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=FlxGd,FlxGd,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns or sets the background color of various elements of the Hierarchical FlexGrid."
    BackColor = FlxGd.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    FlxGd.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

Public Property Get BackColorBkg() As OLE_COLOR
    BackColorBkg = FlxGd.BackColorBkg
End Property

Public Property Let BackColorBkg(ByVal New_BackColorBkg As OLE_COLOR)
    FlxGd.BackColorBkg() = New_BackColorBkg
    PropertyChanged "BackColorBkg"
End Property

Public Property Get BackColorControl() As OLE_COLOR
    BackColorControl = txtCelda.BackColor
End Property
Public Property Let BackColorControl(ByVal New_BackColorControl As OLE_COLOR)
    txtCelda.BackColor() = New_BackColorControl
    PropertyChanged "BackColor"
    TxtBuscar.BackColor() = New_BackColorControl
    PropertyChanged "BackColor"
    txtFecha.BackColor() = New_BackColorControl
    PropertyChanged "BackColor"
End Property
Private Sub FlxGd_RowColChange()
    With FlxGd
        If lnFilaActual <> FlxGd.row Then
            If m_SoloFila = False Then
                If lnFilaActual <> -1 Then
                    DrawIcons ptNone, lnFilaActual
                End If
                lnFilaActual = FlxGd.row
                If FlxGd.TextMatrix(FlxGd.row, 0) = "" Then
                    DrawIcons ptNone, FlxGd.row
                Else
                    DrawIcons ptArrow, lnFilaActual
                End If
                RaiseEvent OnRowChange(FlxGd.row, FlxGd.Col)
            Else
                If lnFilaActual <> -1 Then
                    FlxGd.row = lnFilaActual
                Else
                    lnFilaActual = FlxGd.row
                End If
            End If
        End If
    End With
    RaiseEvent RowColChange
End Sub
Private Sub FlxGd_SelChange()
    RaiseEvent SelChange
End Sub

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=14,0,0,0
Public Property Get EncabezadosAlineacion() As String
    EncabezadosAlineacion = m_EncabezadosAlineacion
End Property

Public Property Let EncabezadosAlineacion(ByVal New_EncabezadosAlineacion As String)
    m_EncabezadosAlineacion = New_EncabezadosAlineacion
    PropertyChanged "EncabezadosAlineacion"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=14,0,0,0
Public Property Get FormatosEdit() As String
    FormatosEdit = m_FormatosEdit
End Property

Public Property Let FormatosEdit(ByVal New_FormatosEdit As String)
    m_FormatosEdit = New_FormatosEdit
    PropertyChanged "FormatosEdit"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=7,0,0,8
Public Property Get CantEntero() As Integer
    CantEntero = m_CantEntero
End Property

Public Property Let CantEntero(ByVal New_CantEntero As Integer)
    m_CantEntero = New_CantEntero
    PropertyChanged "CantEntero"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=7,0,0,2
Public Property Get CantDecimales() As Integer
    CantDecimales = m_CantDecimales
End Property

Public Property Let CantDecimales(ByVal New_CantDecimales As Integer)
    m_CantDecimales = New_CantDecimales
    PropertyChanged "CantDecimales"
End Property
'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=7,0,0,0
Public Property Get AvanceCeldas() As Avance
    AvanceCeldas = m_AvanceCeldas
End Property
Public Property Let AvanceCeldas(ByVal New_AvanceCeldas As Avance)
    m_AvanceCeldas = New_AvanceCeldas
    PropertyChanged "AvanceCeldas"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=13,0,0,
Public Property Get SQLDataSource() As String
    SQLDataSource = m_SQLDataSource
End Property

Public Property Let SQLDataSource(ByVal New_SQLDataSource As String)
    m_SQLDataSource = New_SQLDataSource
    PropertyChanged "SQLDataSource"
End Property
Public Sub GeneraGridDatos()
Dim rs As ADODB.Recordset
Dim oConect As COMConecta.DCOMConecta
Dim Campo As ADODB.Field
Dim lnTotalCampos As Long
Dim i As Long, j As Long

On Error GoTo ErrorGeneraGrid
m_lbRsLoad = True
Set rs = New ADODB.Recordset
Set oConect = New COMConecta.DCOMConecta
If oConect.AbreConexion = False Then Exit Sub

If Len(Trim(m_SQLDataSource)) <> 0 Then
    Set rs = oConect.CargaRecordSet(m_SQLDataSource)
Else
    If m_rsFlex Is Nothing Then Exit Sub
    Set rs = m_rsFlex
End If
FlxGd.Redraw = False
If Not rs.EOF And Not rs.BOF Then
    FlxGd.Cols = rs.Fields.count
    Set FlxGd.DataSource = rs
    lnFilaActual = -1
    FormaCabecera
    If FlxGd.FixedCols > 0 Then
        EnumeraItems True
    End If
    FormateaColumnas
    LlenaFlexCheck
End If
rs.Close: Set rs = Nothing
oConect.CierraConexion
Set oConect = Nothing
FlxGd.Redraw = True
Exit Sub
ErrorGeneraGrid:
    FlxGd.Redraw = True
    MsgBox "Error N° [" & err.Number & "] " & err.Description, vbInformation, "Aviso"
End Sub
'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=FlxGd,FlxGd,-1,TextMatrix
Public Property Get TextMatrix(ByVal row As Long, ByVal Col As Long) As String
Attribute TextMatrix.VB_Description = "Returns or sets the text content of an arbitrary cell (row/column subscripts)."

If FlxGd.TextMatrix(1, 0) <> "" Then TextMatrix = FlxGd.TextMatrix(row, Col)
End Property

Public Property Let TextMatrix(ByVal row As Long, ByVal Col As Long, ByVal New_TextMatrix As String)
    FlxGd.TextMatrix(row, Col) = New_TextMatrix
    PropertyChanged "TextMatrix"
    If lnColChk <> -1 And lnColChk = Col And row > 0 Then
        LlenaValorFlexCheck FlxGd.TextMatrix(row, Col), row
    End If
    If lnColOpt <> -1 And lnColOpt = Col And row > 0 Then
        LlenaValorFlexCheck FlxGd.TextMatrix(row, Col), row, False
    End If
End Property
Public Property Get row() As Long
Attribute row.VB_MemberFlags = "400"
    row = FlxGd.row
End Property
Public Property Let row(ByVal vNewValue As Long)
    FlxGd.row = vNewValue
    PropertyChanged "Row"
End Property

Public Property Get TopRow() As Long
    TopRow = FlxGd.TopRow
End Property
Public Property Let TopRow(ByVal vNewValue As Long)
    FlxGd.TopRow = vNewValue
    PropertyChanged "TopRow"
End Property

Public Property Get Col() As Long
Attribute Col.VB_MemberFlags = "400"
    Col = FlxGd.Col
End Property
Public Property Let Col(ByVal vNewValue As Long)
    FlxGd.Col = vNewValue
    PropertyChanged "Col"
End Property
'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=FlxGd,FlxGd,-1,TextArray
Public Property Get TextArray(ByVal index As Long) As String
Attribute TextArray.VB_Description = "Returns or sets the text content of an arbitrary cell (single subscript)."
    TextArray = FlxGd.TextArray(index)
End Property
Public Property Let TextArray(ByVal index As Long, ByVal New_TextArray As String)
    FlxGd.TextArray(index) = New_TextArray
    PropertyChanged "TextArray"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=21,0,0,0
Public Property Get rsFlex() As ADODB.Recordset
    Set rsFlex = m_rsFlex
End Property
Public Property Let rsFlex(ByVal New_rsFlex As ADODB.Recordset)
    Set m_rsFlex = New_rsFlex
    PropertyChanged "rsFlex"
    GeneraGridDatos
End Property

Private Sub EnfocaTexto(txtCelda As Variant, KeyAscii As Integer, fg As MSHFlexGrid, Optional pnTopIni As Long = 0, Optional pnLeftIni As Long = 0)
Dim nX, nY As Long
Dim lnTpoFormato As Long
Dim R As Control
'   On Error GoTo EnfocaTextoErr
   If txtCelda Is Nothing Then Exit Sub
   If fg.TextMatrix(fg.row, 0) = "" Then Exit Sub
   If txtCelda.Name = "txtFecha" Then
        lnTpoFormato = DeterminaFormato(fg.Col)
        Select Case lnTpoFormato
            Case 5
                txtCelda.Mask = "##/##/#### ##:##:##"
                txtCelda.Text = "  /  /       :  :  "
            Case 6
                txtCelda.Mask = "##:##:##"
                txtCelda.Text = "  :  :  "
            Case Else
                txtCelda.Mask = "##/##/####"
                txtCelda.Text = "  /  /    "
        End Select
   Else
        If txtCelda.Name = "cboCelda" Then
            txtCelda.ListIndex = -1
        Else
            txtCelda.Text = ""
        End If
    End If
   nX = fg.Left + fg.ColPos(fg.Col) + 10 + pnLeftIni
   nY = fg.Top + fg.RowPos(fg.row) + 10 + pnTopIni
   
   If txtCelda.Name = "txtFecha" Then
        txtCelda.Height = fg.CellHeight + 30
   Else
        If txtCelda.Name <> "cboCelda" Then
            txtCelda.Height = fg.CellHeight + 30
        Else
            fg.RowHeight(fg.row) = txtCelda.Height - 5
        End If
   End If
   txtCelda.Width = fg.CellWidth + 30
   txtCelda.Tag = fg.row
   txtCelda.Visible = True
   If txtCelda.Name <> "txtFecha" And txtCelda.Name <> "cboCelda" And txtCelda.Name <> "TxtBuscar" Then
        Select Case fg.ColAlignment(fg.Col)
            Case 0, 1, 2
                txtCelda.Alignment = 0
            Case 6, 7, 8
                txtCelda.Alignment = 1
            Case 3, 4, 5
                txtCelda.Alignment = 2
        End Select
   End If
   Select Case KeyAscii
        Case 0
            If txtCelda.Name = "txtFecha" And fg.Text = "" Then
                Select Case lnTpoFormato
                    Case 5
                        txtCelda.Text = "  /  /       :  :  "
                    Case 6
                        txtCelda.Text = "  :  :  "
                    Case Else
                        txtCelda.Text = "  /  /    "
                End Select
            Else
                If IsDate(fg.Text) = False And txtCelda.Name = "txtFecha" Then
                    'txtCelda.Text = "  /  /    "
                    Select Case lnTpoFormato
                        Case 5
                            txtCelda.Text = "  /  /       :  :  "
                        Case 6
                            txtCelda.Text = "  :  :  "
                        Case Else
                            txtCelda.Text = "  /  /    "
                    End Select
                Else
                    If txtCelda.Name = "cboCelda" Then
                        If fg.Text <> "" Then
                            Dim nCont As Integer
                            For nCont = 0 To txtCelda.ListCount - 1
                                If txtCelda.List(nCont) = fg.Text Then
                                    txtCelda.ListIndex = nCont
                                End If
                            Next
                        Else
                            txtCelda.ListIndex = -1
                        End If
'                        If fg.Text <> "" Then
'                            txtCelda.Text = fg.Text
'                        Else
'                            txtCelda.ListIndex = -1
'                        End If
                    Else
                        txtCelda.Text = fg.Text
                    End If
                End If
            End If
            If txtCelda.Name <> "cboCelda" Then
                txtCelda.SelStart = 0
                txtCelda.SelLength = Len(txtCelda.Text)
            End If
        Case Else
            If txtCelda.Name = "txtFecha" Then
                If IsNumeric(Chr(KeyAscii)) Then
                    If IsNumeric(Chr(KeyAscii)) Then
                        'txtCelda.Text = Chr(KeyAscii) & " /  /    "
                        Select Case lnTpoFormato
                            Case 5
                                txtCelda.Text = Chr(KeyAscii) & " /  /       :  :  "
                            Case 6
                                txtCelda.Text = Chr(KeyAscii) & " :  :  "
                            Case Else
                                txtCelda.Text = Chr(KeyAscii) & " /  /    "
                        End Select
                        'If lnTpoFormato = 5 Then
                        '    txtCelda.Text = Chr(KeyAscii) & " /  /       :  :  "
                        'Else
                        '    txtCelda.Text = Chr(KeyAscii) & " /  /    "
                        'End If
                    End If
                End If
                txtCelda.SelStart = 1
            Else
                If txtCelda.Name = "cboCelda" Then
                    txtCelda.ListIndex = 0
                Else
                    txtCelda.Text = Chr(KeyAscii)
                    txtCelda.SelStart = 1
                End If
            End If
        
   End Select
   txtCelda.Left = nX
   txtCelda.Top = nY
   txtCelda.SetFocus
Exit Sub
EnfocaTextoErr:
   MsgBox err.Description, vbInformation, "¡Aviso!"
End Sub
'##ModelId=3A80B1670271
Private Sub Flex_PresionaKey(Flex As MSHFlexGrid, KeyCode As Integer, Shift As Integer)
   On Error GoTo PresionaKeyErr
    Select Case KeyCode
        Case vbKeyC And Shift = 2   '   Copiar  [Ctrl+C]
            Clipboard.Clear
            Clipboard.SetText Flex.Text
            KeyCode = 0
        Case vbKeyV And Shift = 2   '   Pegar  [Ctrl+V]
            Flex.Text = Clipboard.GetText
            KeyCode = 0
        Case vbKeyX And Shift = 2   '   Cortar  [Ctrl+X]
            Clipboard.Clear
            Clipboard.SetText Flex.Text
            Flex.Text = ""
            KeyCode = 0
        Case vbKeyDelete            '   Borrar [Delete]
            Flex.Text = ""
            KeyCode = 0
    End Select
   Exit Sub
PresionaKeyErr:
   MsgBox err.Description, vbInformation, "¡Aviso!"
End Sub

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=0,0,0,False
Public Property Get lbRsLoad() As Boolean
Attribute lbRsLoad.VB_MemberFlags = "400"
    lbRsLoad = m_lbRsLoad
End Property
Public Property Let lbRsLoad(ByVal New_lbRsLoad As Boolean)
    m_lbRsLoad = New_lbRsLoad
    PropertyChanged "lbRsLoad"
End Property
Public Sub CargaCombo(ByVal rs As ADODB.Recordset)
Dim Campo As ADODB.Field
Dim lsDato As String
If rs Is Nothing Then Exit Sub
cboCelda.Clear
Do While Not rs.EOF
    lsDato = ""
    For Each Campo In rs.Fields
        lsDato = lsDato & Campo.value & Space(75)
    Next
    lsDato = Mid(lsDato, 1, Len(lsDato) - 75)
    cboCelda.AddItem lsDato
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing
End Sub
'ANDE 20170719
Public Sub LimpiarCombo()
    cboCelda.Clear
End Sub
'END ANDE

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=FlxGd,FlxGd,-1,SelectionMode
Public Property Get SelectionMode() As SelectionModeSettings
Attribute SelectionMode.VB_Description = "Returns or sets whether a Hierarchical FlexGrid should allow regular cell selection, selection by rows, or selection by columns."
    SelectionMode = FlxGd.SelectionMode
End Property

Public Property Let SelectionMode(ByVal New_SelectionMode As SelectionModeSettings)
    FlxGd.SelectionMode() = New_SelectionMode
    PropertyChanged "SelectionMode"
End Property

Private Sub FlxGd_EnterCell()
    RaiseEvent EnterCell
End Sub
Private Sub FlxGd_Click()
    SeleccionaCheck
    SeleccionaOpt
    RaiseEvent Click
End Sub
'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=0,0,0,False
Public Property Get lbEditarFlex() As Boolean
Attribute lbEditarFlex.VB_ProcData.VB_Invoke_Property = "FlexPagePropiedades"
    lbEditarFlex = m_lbEditarFlex
End Property

Public Property Let lbEditarFlex(ByVal New_lbEditarFlex As Boolean)
    m_lbEditarFlex = New_lbEditarFlex
    PropertyChanged "lbEditarFlex"
End Property
'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=0,0,0,False
Public Property Get lbFlexDuplicados() As Boolean
Attribute lbFlexDuplicados.VB_ProcData.VB_Invoke_Property = "FlexPagePropiedades"
    lbFlexDuplicados = m_lbFlexDuplicados
End Property

Public Property Let lbFlexDuplicados(ByVal New_lbFlexDuplicados As Boolean)
    m_lbFlexDuplicados = New_lbFlexDuplicados
    PropertyChanged "lbFlexDuplicados"
End Property
Private Function BuscaDuplicado(pnFila As Long, pnCol As Long, psValor As String) As Boolean
Dim i As Long
BuscaDuplicado = False
For i = 1 To FlxGd.Rows - 1
    If i <> pnFila Then
        If FlxGd.TextMatrix(i, pnCol) = psValor Then
            BuscaDuplicado = True
            Exit Function
        End If
    End If
Next
End Function
Private Function NumerosDecimales(cTexto As TextBox, intTecla As Integer, _
    Optional nLongitud As Integer = 8, Optional nDecimal As Integer = 2) As Integer
    Dim cValidar As String
    Dim cCadena As String
    cCadena = cTexto
    cValidar = "-0123456789."
    
    If InStr(".", Chr(intTecla)) <> 0 Then
        If InStr(cCadena, ".") <> 0 Then
            intTecla = 0
            Beep
        ElseIf intTecla > 26 Then
            If InStr(cValidar, Chr(intTecla)) = 0 Then
                intTecla = 0
                Beep
            End If
        End If
    ElseIf intTecla > 26 Then
        If InStr(cValidar, Chr(intTecla)) = 0 Then
            intTecla = 0
            Beep
        End If
    End If
    Dim vPosCur As Byte
    Dim vPosPto As Byte
    Dim vNumDec As Byte
    Dim vNumLon As Byte
    
    vPosPto = InStr(cTexto.Text, ".")
    vPosCur = cTexto.SelStart
    vNumLon = Len(cTexto)
    If vPosPto > 0 Then
        vNumDec = Len(Mid(cTexto, vPosPto + 1))
    End If
    If vPosPto > 0 Then
        If cTexto.SelLength <> Len(cTexto) Then
        If ((vNumDec >= nDecimal And cTexto.SelStart >= vPosPto) Or _
        (vNumLon >= nLongitud)) _
        And intTecla <> vbKeyBack And intTecla <> vbKeyDecimal And intTecla <> vbKeyReturn Then
            intTecla = 0
            Beep
        End If
        End If
    Else
        If vNumLon >= nLongitud And intTecla <> vbKeyBack _
        And intTecla <> vbKeyReturn Then
            intTecla = 0
            Beep
        End If
        If (vNumLon - cTexto.SelStart) > nDecimal And intTecla = 46 Then
            intTecla = 0
            Beep
        End If
    End If
    NumerosDecimales = intTecla
End Function
Private Function NumerosEnteros(intTecla As Integer, Optional pbNegativos As Boolean = False) As Integer
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
Private Function SoloLetras(intTecla As Integer) As Integer
    cValidar = "0123456789+:;'<>?_=+[]{}|!@#$%^&()*"
    If intTecla > 26 Then
        If InStr(cValidar, Chr(intTecla)) <> 0 Then
            intTecla = 0
            Beep
        End If
    End If
    SoloLetras = intTecla
End Function
'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=FlxGd,FlxGd,-1,Clear
Public Sub Clear()
Attribute Clear.VB_Description = "Clears the contents of the Hierarchical FlexGrid. This includes all text, pictures, and cell formatting."
    FlxGd.Clear
    lnFilaActual = -1
End Sub
Private Function ValFecha(lsControl As Control) As Boolean
   If Mid(lsControl, 1, 2) > 0 And Mid(lsControl, 1, 2) <= 31 Then
        If Mid(lsControl, 4, 2) > 0 And Mid(lsControl, 4, 2) <= 12 Then
            If Mid(lsControl, 7, 4) >= 1900 And Mid(lsControl, 7, 4) <= 9999 Then
               If IsDate(lsControl) = False Then
                    ValFecha = False
                    MsgBox "Formato de fecha no es válido", vbInformation, "Aviso"
                    lsControl.SetFocus
                    Exit Function
               Else
                    ValFecha = True
               End If
            Else
                ValFecha = False
                MsgBox "Año de Fecha no es válido", vbInformation, "Aviso"
                lsControl.SetFocus
                lsControl.SelStart = 6
                lsControl.SelLength = 4
                Exit Function
            End If
        Else
            ValFecha = False
            MsgBox "Mes de Fecha no es válido", vbInformation, "Aviso"
            lsControl.SetFocus
            lsControl.SelStart = 3
            lsControl.SelLength = 2
            Exit Function
        End If
    Else
        ValFecha = False
        MsgBox "Dia de Fecha no es válido", vbInformation, "Aviso"
        lsControl.SetFocus
        lsControl.SelStart = 0
        lsControl.SelLength = 2
        Exit Function
    End If
End Function
'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=TxtBuscar,TxtBuscar,-1,lbUltimaInstancia
Public Property Get lbUltimaInstancia() As Boolean
    lbUltimaInstancia = TxtBuscar.lbUltimaInstancia
End Property
Public Property Let lbUltimaInstancia(ByVal New_lbUltimaInstancia As Boolean)
    TxtBuscar.lbUltimaInstancia() = New_lbUltimaInstancia
    PropertyChanged "lbUltimaInstancia"
End Property
Public Sub FormateaColumnas()
Dim lsFormato As String
Dim i As Long
Dim j As Long
If m_lbFormatoCol = False Then Exit Sub
For j = 1 To FlxGd.Cols - 1
    lsFormato = EmiteFormatoCol(j)
    If lsFormato <> "" Then
        For i = 1 To FlxGd.Rows - 1
            FlxGd.TextMatrix(i, j) = Format(FlxGd.TextMatrix(i, j), lsFormato)
        Next i
    End If
Next j

End Sub
Private Function EmiteFormatoCol(ByVal lnCol As Long) As String
Dim lnTpoFormato As Long
    lnTpoFormato = DeterminaFormato(lnCol)
    Select Case lnTpoFormato
        Case 2, 4
            EmiteFormatoCol = "###," & String(m_CantEntero, "#") & "#0." & String(m_CantDecimales, "0")
        Case 3
            EmiteFormatoCol = "#," & String(m_CantEntero, "#") & "#0"
        Case Else
            EmiteFormatoCol = ""
    End Select
End Function
'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=TxtBuscar,TxtBuscar,-1,TipoBusqueda
Public Property Get TipoBusqueda() As TipoBusqueda
    TipoBusqueda = TxtBuscar.TipoBusqueda
End Property

Public Property Let TipoBusqueda(ByVal New_TipoBusqueda As TipoBusqueda)
    TxtBuscar.TipoBusqueda() = New_TipoBusqueda
    PropertyChanged "TipoBusqueda"
End Property
'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=0,0,0,False
Public Property Get lbFormatoCol() As Boolean
Attribute lbFormatoCol.VB_ProcData.VB_Invoke_Property = "FlexPagePropiedades"
    lbFormatoCol = m_lbFormatoCol
End Property

Public Property Let lbFormatoCol(ByVal New_lbFormatoCol As Boolean)
    m_lbFormatoCol = New_lbFormatoCol
    PropertyChanged "lbFormatoCol"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=cboCelda,cboCelda,-1,List
Public Property Get List(ByVal index As Integer) As String
Attribute List.VB_Description = "Devuelve o establece los elementos contenidos en la parte de lista de un control."
    List = cboCelda.List(index)
End Property

Public Property Let List(ByVal index As Integer, ByVal New_List As String)
    cboCelda.List(index) = New_List
    PropertyChanged "List"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=FlxGd,FlxGd,-1,FontFixed
Public Property Get FontFixed() As Font
Attribute FontFixed.VB_Description = "Returns or sets the default font or the font for individual cells."
    Set FontFixed = FlxGd.FontFixed
End Property

Public Property Set FontFixed(ByVal New_FontFixed As Font)
    Set FlxGd.FontFixed = New_FontFixed
    PropertyChanged "FontFixed"
End Property
'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=FlxGd,FlxGd,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns or sets the default font or the font for individual cells."
Attribute Font.VB_UserMemId = -512
    Set Font = FlxGd.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set FlxGd.Font = New_Font
    PropertyChanged "Font"
    Set TxtBuscar.Font = New_Font
    PropertyChanged "Font"
    Set txtCelda.Font = New_Font
    PropertyChanged "Font"
    Set txtFecha.Font = New_Font
    PropertyChanged "Font"
    Set cboCelda.Font = New_Font
    PropertyChanged "Font"
    
End Property
Private Sub DrawIcons(ByVal PicType As PictureType, ByVal vRow As Long, Optional UseRedraw As Boolean = True)
    On Error GoTo Err_DrawIcons
    Dim iRow&, iCol&, iRowsel&, iColSel&
    If m_lbPuntero = False Then Exit Sub
    If FlxGd.TextMatrix(FlxGd.row, 0) = "" Then Exit Sub
    With FlxGd
        If UseRedraw Then .Redraw = False
        iRow = .row
        iCol = .Col
        iRowsel = .RowSel
        iColSel = .ColSel
        .row = vRow
        .Col = 0
        .CellPictureAlignment = flexAlignLeftCenter
        If PicType = ptNone Then
            Set .CellPicture = Nothing
            If vRow = .Rows - 2 And iRow > vRow Then
                SendKeys "{DOWN}"
            End If
        Else
            Set .CellPicture = imgHolders(PicType).Picture
        End If
        .row = iRow
        .Col = iCol
        .RowSel = iRowsel
        .ColSel = iColSel
        If UseRedraw Then .Redraw = True
        If .Visible And .Enabled Then
            '.SetFocus
        End If
    End With
    Exit Sub
Err_DrawIcons:
    FlxGd.Redraw = True
    MsgBox err.Description, vbInformation, "Aviso"
End Sub
Sub DoColumnSort()
    On Error Resume Next
    If m_lbOrdenaCol Then
        If FlxGd.TextMatrix(1, 0) <> "" Then
            i = m_iSortCol
            m_iSortCol = FlxGd.Col
            If m_iSortType = 2 Then
                m_iSortType = 1
            Else
                If i = m_iSortCol Then
                    m_iSortType = 2
                Else
                    m_iSortType = 1
                End If
            End If
            If IsDate(FlxGd.TextMatrix(1, FlxGd.Col)) Then
                m_iSortType = 9
                m_iSortCustomAscending = Not m_iSortCustomAscending
            End If
            If IsNumeric(FlxGd.TextMatrix(1, FlxGd.Col)) Then
                m_iSortType = 9
                m_iSortCustomAscending = Not m_iSortCustomAscending
            End If
        End If
    End If
    With FlxGd
        .Redraw = False
        .Sort = m_iSortType
        EnumeraItems m_lbRsLoad
        lnFilaActual = DeterminaFilaPuntero
        .Redraw = True
    End With
    '
End Sub
Private Function DeterminaFilaPuntero() As Long
Dim i As Integer
DeterminaFilaPuntero = 0
Dim pnCol As Long
Dim pnRow As Long
pnCol = FlxGd.Col
pnRow = FlxGd.row
If m_lbPuntero = False Then Exit Function
For i = 1 To FlxGd.Rows - 1
    FlxGd.row = i
    FlxGd.Col = 0
    If FlxGd.CellPicture = imgHolders(ptArrow).Picture Then
        FlxGd.Col = pnCol
        FlxGd.row = pnRow
        DeterminaFilaPuntero = i
        Exit Function
    End If
Next
End Function
'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=0,0,0,True
Public Property Get lbPuntero() As Boolean
Attribute lbPuntero.VB_ProcData.VB_Invoke_Property = "FlexPagePropiedades"
    lbPuntero = m_lbPuntero
End Property

Public Property Let lbPuntero(ByVal New_lbPuntero As Boolean)
    m_lbPuntero = New_lbPuntero
    PropertyChanged "lbPuntero"
End Property
'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=0,0,0,False
Public Property Get lbOrdenaCol() As Boolean
    lbOrdenaCol = m_lbOrdenaCol
End Property

Public Property Let lbOrdenaCol(ByVal New_lbOrdenaCol As Boolean)
    m_lbOrdenaCol = New_lbOrdenaCol
    PropertyChanged "lbOrdenaCol"
End Property
Private Function SeleccionaCheck()
Dim lbCancel As Boolean
If lnColChk = -1 Then Exit Function
If m_lbEditarFlex = False Then Exit Function
If FlxGd.Col <> lnColChk Then Exit Function
If FlxGd.TextMatrix(1, 0) = "" Then Exit Function
If ColumnaAEditar(FlxGd.Col) = False Then Exit Function
With FlxGd
    lbCancel = True
    FlxGd.Col = lnColChk
    RaiseEvent OnValidate(FlxGd.row, lnColChk, lbCancel)
    .Redraw = False
    .CellAlignment = flexAlignCenterCenter
    If lbCancel = True Then
        If FlxGd.TextMatrix(FlxGd.row, lnColChk) = "" Then
            FlxGd.TextMatrix(FlxGd.row, lnColChk) = "."
            Set .CellPicture = ImgCheck1.Picture
        Else
            FlxGd.TextMatrix(FlxGd.row, lnColChk) = ""
            Set .CellPicture = ImgCheck0.Picture
        End If
    End If
    .Redraw = True
End With
RaiseEvent OnCellCheck(FlxGd.row, lnColChk)
End Function
Private Function SeleccionaOpt()
Dim i As Integer
Dim lnFilaAnt As Long
Dim lbCancel As Boolean
If lnColOpt = -1 Then Exit Function
If m_lbEditarFlex = False Then Exit Function
If FlxGd.Col <> lnColOpt Then Exit Function
If FlxGd.TextMatrix(1, 0) = "" Then Exit Function
If ColumnaAEditar(FlxGd.Col) = False Then Exit Function
lnFilaAnt = 0
With FlxGd
    lbCancel = True
    FlxGd.Col = lnColOpt
    RaiseEvent OnValidate(FlxGd.row, lnColOpt, lbCancel)
    .Redraw = False
    lnFilaAnt = FlxGd.row
    .CellAlignment = flexAlignCenterCenter
    If lbCancel = True Then
        If FlxGd.TextMatrix(FlxGd.row, lnColOpt) = "" Then
            FlxGd.TextMatrix(FlxGd.row, lnColOpt) = "."
            Set .CellPicture = ImgOpt1.Picture
        'Else
        '    FlxGd.TextMatrix(FlxGd.Row, lnColOpt) = ""
        '    Set .CellPicture = ImgOpt0.Picture
        End If
        For i = 1 To FlxGd.Rows - 1
            If i <> lnFilaAnt Then
                FlxGd.row = i
                FlxGd.TextMatrix(FlxGd.row, lnColOpt) = ""
                Set .CellPicture = ImgOpt0.Picture
            End If
        Next
        FlxGd.row = lnFilaAnt
    End If
    .Redraw = True
End With
'RaiseEvent OnCellCheck(FlxGd.Row, lnColChk)
End Function
Sub SeleccionaChekTecla()
Dim lbCancel As Boolean
Dim pnCol1 As Long
If lnColChk = -1 Then Exit Sub
If FlxGd.TextMatrix(FlxGd.row, 0) = "" Then Exit Sub
If ColumnaAEditar(lnColChk) = False Then Exit Sub
With FlxGd
    lbCancel = True
    pnCol1 = .Col
    .Col = lnColChk
     RaiseEvent OnValidate(FlxGd.row, lnColChk, lbCancel)
    .Redraw = False
    .CellAlignment = flexAlignCenterCenter
    If lbCancel = True Then
        If FlxGd.TextMatrix(FlxGd.row, .Col) = "" Then
            FlxGd.TextMatrix(FlxGd.row, .Col) = "."
            Set .CellPicture = ImgCheck1.Picture
        Else
            FlxGd.TextMatrix(FlxGd.row, .Col) = ""
            Set .CellPicture = ImgCheck0.Picture
        End If
    End If
    .Col = pnCol1
    .Redraw = True
End With
RaiseEvent OnCellCheck(FlxGd.row, lnColChk)
End Sub
Sub SeleccionaOptTecla()
Dim lbCancel As Boolean
Dim pnCol1 As Long
If lnColOpt = -1 Then Exit Sub
If FlxGd.TextMatrix(FlxGd.row, 0) = "" Then Exit Sub
If ColumnaAEditar(lnColOpt) = False Then Exit Sub
With FlxGd
    lbCancel = True
    pnCol1 = .Col
    .Col = lnColOpt
     RaiseEvent OnValidate(FlxGd.row, lnColOpt, lbCancel)
    .Redraw = False
    .CellAlignment = flexAlignCenterCenter
    If lbCancel = True Then
        If FlxGd.TextMatrix(FlxGd.row, .Col) = "" Then
            FlxGd.TextMatrix(FlxGd.row, .Col) = "."
            Set .CellPicture = ImgOpt1.Picture
        Else
            FlxGd.TextMatrix(FlxGd.row, .Col) = ""
            Set .CellPicture = ImgOpt0.Picture
        End If
    End If
    .Col = pnCol1
    .Redraw = True
End With
'RaiseEvent OnCellCheck(FlxGd.Row, lnColChk)
End Sub
Private Sub LlenaFlexCheck(Optional ByVal pbCheck As Boolean = True)
Dim j As Long
If pbCheck Then
    If lnColChk = -1 Then Exit Sub
Else
    If lnColOpt = -1 Then Exit Sub
End If
For j = 1 To FlxGd.Rows - 1
    If pbCheck Then
        FlxGd.Col = lnColChk
    Else
        FlxGd.Col = lnColOpt
    End If
    FlxGd.row = j
    FlxGd.CellPictureAlignment = flexAlignCenterCenter
    If pbCheck Then
        If Val(FlxGd.TextMatrix(j, lnColChk)) = 1 Then
            Set FlxGd.CellPicture = ImgCheck1.Picture
            FlxGd.TextMatrix(j, lnColChk) = "."
        Else
            Set FlxGd.CellPicture = ImgCheck0.Picture
            FlxGd.TextMatrix(j, lnColChk) = ""
        End If
    Else
        If Val(FlxGd.TextMatrix(j, lnColOpt)) = 1 Then
            Set FlxGd.CellPicture = ImgOpt1.Picture
            FlxGd.TextMatrix(j, lnColOpt) = "."
        Else
            Set FlxGd.CellPicture = ImgOpt0.Picture
            FlxGd.TextMatrix(j, lnColOpt) = ""
        End If
    End If
Next
End Sub
Private Sub LlenaValorFlexCheck(psValor As String, pnRow As Long, Optional ByVal pbCheck As Boolean = True)
Dim j As Long
Dim i As Integer
Dim lnFilaAnt As Long
If pbCheck Then
    If lnColChk = -1 Then Exit Sub
Else
    If lnColOpt = -1 Then Exit Sub
End If
If pbCheck Then
    FlxGd.Col = lnColChk
Else
    FlxGd.Col = lnColOpt
End If
lnFilaAnt = 1
FlxGd.row = pnRow
lnFilaAnt = pnRow
FlxGd.CellAlignment = flexAlignCenterCenter
If pbCheck Then
    If Val(psValor) = 1 Then
        FlxGd.TextMatrix(pnRow, lnColChk) = "."
        Set FlxGd.CellPicture = ImgCheck1.Picture
    Else
        FlxGd.TextMatrix(pnRow, lnColChk) = ""
        Set FlxGd.CellPicture = ImgCheck0.Picture
    End If
Else
    If Val(psValor) = 1 Then
        FlxGd.TextMatrix(pnRow, lnColOpt) = "."
        Set FlxGd.CellPicture = ImgOpt1.Picture
        For i = 1 To FlxGd.Rows - 1
            If i <> lnFilaAnt Then
                FlxGd.row = i
                FlxGd.TextMatrix(FlxGd.row, lnColOpt) = ""
                Set FlxGd.CellPicture = ImgOpt0.Picture
            End If
        Next
        FlxGd.row = lnFilaAnt
    End If
End If
End Sub

Private Function DeterminaColChek() As Long
Dim pControles As String
Dim X As Long
Dim lControles() As String
Dim lnNroControl As Long
pControles = m_ListaControles
DeterminaColChek = -1
If Len(Trim(pControles)) > 0 Then
    For X = 0 To Cols - 1
        vPos = InStr(1, pControles, "-", vbTextCompare)
        ReDim Preserve lControles(X)
        lControles(X) = Mid(pControles, 1, IIf(vPos > 0, vPos - 1, Len(pControles)))
        If Val(lControles(X)) = 4 Then
            DeterminaColChek = X
            Exit Function
        End If
        If pControles <> "" Then
            pControles = Mid(pControles, IIf(vPos > 0, vPos + 1, Len(pControles)))
        End If
    Next X
End If
End Function
Private Function DeterminaColOption() As Long
Dim pControles As String
Dim X As Long
Dim lControles() As String
Dim lnNroControl As Long
pControles = m_ListaControles
DeterminaColOption = -1
If Len(Trim(pControles)) > 0 Then
    For X = 0 To Cols - 1
        vPos = InStr(1, pControles, "-", vbTextCompare)
        ReDim Preserve lControles(X)
        lControles(X) = Mid(pControles, 1, IIf(vPos > 0, vPos - 1, Len(pControles)))
        If Val(lControles(X)) = 5 Then
            DeterminaColOption = X
            Exit Function
        End If
        If pControles <> "" Then
            pControles = Mid(pControles, IIf(vPos > 0, vPos + 1, Len(pControles)))
        End If
    Next X
End If
End Function
'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=FlxGd,FlxGd,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Devuelve o establece un valor que determina si un objeto puede responder a eventos generados por el usuario."
    Enabled = FlxGd.Enabled
End Property
Public Property Let Enabled(ByVal New_Enabled As Boolean)
    FlxGd.Enabled = New_Enabled
    PropertyChanged "Enabled"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=0,0,0,False
Public Property Get lbBuscaDuplicadoText() As Boolean
    lbBuscaDuplicadoText = m_lbBuscaDuplicadoText
End Property

Public Property Let lbBuscaDuplicadoText(ByVal New_lbBuscaDuplicadoText As Boolean)
    m_lbBuscaDuplicadoText = New_lbBuscaDuplicadoText
    PropertyChanged "lbBuscaDuplicadoText"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=FlxGd,FlxGd,-1,Recordset
Public Property Get Recordset() As IRecordset
Attribute Recordset.VB_Description = "Binds the Hierarchical FlexGrid to an ADO Recordset. Not available at design time."
    Set Recordset = FlxGd.Recordset
End Property
Public Property Set Recordset(ByVal New_Recordset As IRecordset)
    Set FlxGd.Recordset = New_Recordset
    PropertyChanged "Recordset"
    If Not New_Recordset Is Nothing Then
        m_lbRsLoad = True
        lnFilaActual = -1
        FormaCabecera
        EnumeraItems True
        FormateaColumnas
        LlenaFlexCheck
        LlenaFlexCheck False
    Else
        FlxGd.Clear
        FormaCabecera
    End If
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=txtCelda,txtCelda,-1,MaxLength
Public Property Get MaxLength() As Long
Attribute MaxLength.VB_Description = "Devuelve o establece el número máximo de caracteres que se puede escribir en un control."
    MaxLength = txtCelda.MaxLength
End Property

Public Property Let MaxLength(ByVal New_MaxLength As Long)
    txtCelda.MaxLength() = New_MaxLength
    PropertyChanged "MaxLength"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=MnuFGridAddRow,MnuFGridAddRow,-1,Enabled
Public Property Get EnabledMnuAdd() As Boolean
Attribute EnabledMnuAdd.VB_Description = "Devuelve o establece un valor que determina si un objeto puede responder a eventos generados por el usuario."
    EnabledMnuAdd = MnuFGridAddRow.Enabled
End Property

Public Property Let EnabledMnuAdd(ByVal New_EnabledMnuAdd As Boolean)
    MnuFGridAddRow.Enabled() = New_EnabledMnuAdd
    PropertyChanged "EnabledMnuAdd"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=mnuDeleteGridRow,mnuDeleteGridRow,-1,Enabled
Public Property Get EnabledMnuEliminar() As Boolean
Attribute EnabledMnuEliminar.VB_Description = "Devuelve o establece un valor que determina si un objeto puede responder a eventos generados por el usuario."
    EnabledMnuEliminar = mnuDeleteGridRow.Enabled
End Property
Public Property Let EnabledMnuEliminar(ByVal New_EnabledMnuEliminar As Boolean)
    mnuDeleteGridRow.Enabled() = New_EnabledMnuEliminar
    PropertyChanged "EnabledMnuEliminar"
End Property
Public Function GetRsNew(Optional pnColIni As Integer = 1) As ADODB.Recordset
Dim i As Long
Dim j As Long
Dim RsAUX As ADODB.Recordset
Dim lnFila As Long
Dim lnCol As Long
Dim lsTipoDato As DataTypeEnum
Dim lsTamCampo As Long
Dim lnFormatoCol As Long
If FlxGd.TextMatrix(1, 0) <> "" Then
    lnFila = 0
    'formamos generamos del recordset
    Set RsAUX = New ADODB.Recordset
    For i = pnColIni To FlxGd.Cols - 1
        If Len(Trim(FlxGd.TextMatrix(lnFila + 1, i))) >= 16 Then
            lsTipoDato = adVarChar
        Else
            lnFormatoCol = DeterminaFormato(i)
            If (lnFormatoCol = 2 Or lnFormatoCol = 3 Or lnFormatoCol = 5) And FlxGd.ColAlignment(i) >= 7 Then
                lsTipoDato = adDouble
            Else
                If ValidaFecha(FlxGd.TextMatrix(lnFila + 1, i)) = "" Then
                    lsTipoDato = adDate
                Else
                    lsTipoDato = adVarChar
                End If
            End If
        End If
        If lsTipoDato = adVarChar Then
            RsAUX.Fields.Append FlxGd.TextMatrix(lnFila, i), lsTipoDato, 400, adFldMayBeNull
        Else
            RsAUX.Fields.Append FlxGd.TextMatrix(lnFila, i), lsTipoDato, , adFldMayBeNull
        End If
    Next
    RsAUX.Open
    
    For i = 1 To FlxGd.Rows - 1
        RsAUX.AddNew
        'columnas
        For j = pnColIni To FlxGd.Cols - 1
            If j = lnColChk Or j = lnColOpt Then
                RsAUX.Fields(FlxGd.TextMatrix(0, j)) = IIf(FlxGd.TextMatrix(i, j) = ".", 1, 0)
            Else
                If RsAUX.Fields(FlxGd.TextMatrix(0, j)).Type = adDouble Then
                    RsAUX.Fields(FlxGd.TextMatrix(0, j)) = CCur(IIf(FlxGd.TextMatrix(i, j) = "", "0", FlxGd.TextMatrix(i, j)))
                Else
                    If RsAUX.Fields(FlxGd.TextMatrix(0, j)).Type = adDate Then
                        RsAUX.Fields(FlxGd.TextMatrix(0, j)) = IIf(FlxGd.TextMatrix(i, j) = "", Null, FlxGd.TextMatrix(i, j))
                    Else
                        RsAUX.Fields(FlxGd.TextMatrix(0, j)) = FlxGd.TextMatrix(i, j)
                    End If
                End If
            End If
        Next
        RsAUX.Update
    Next
    RsAUX.MoveFirst
    Set GetRsNew = RsAUX
End If
End Function
'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=FlxGd,FlxGd,-1,Appearance
Public Property Get Appearance() As AppearanceSettings
Attribute Appearance.VB_Description = "Returns or sets whether a control should be painted with 3-D effects."
    Appearance = FlxGd.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As AppearanceSettings)
    FlxGd.Appearance() = New_Appearance
    PropertyChanged "Appearance"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=FlxGd,FlxGd,-1,BorderStyle
Public Property Get BorderStyle() As BorderStyleSettings
Attribute BorderStyle.VB_Description = "Returns or sets the border style for an object."
    BorderStyle = FlxGd.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As BorderStyleSettings)
    FlxGd.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property
'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=TxtBuscar,TxtBuscar,-1,rsDebe
Public Property Get rsDebe() As ADODB.Recordset
    Set rsDebe = TxtBuscar.rsDebe
End Property
Public Property Let rsDebe(ByVal New_rsDebe As ADODB.Recordset)
    Set TxtBuscar.rsDebe = New_rsDebe
    PropertyChanged "rsDebe"
End Property
'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=TxtBuscar,TxtBuscar,-1,rsHaber
Public Property Get rsHaber() As ADODB.Recordset
    Set rsHaber = TxtBuscar.rsHaber
End Property
Public Property Let rsHaber(ByVal New_rsHaber As ADODB.Recordset)
    Set TxtBuscar.rsHaber = New_rsHaber
    PropertyChanged "rsHaber"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=TxtBuscar,TxtBuscar,-1,psDH
Public Property Get psDH() As String
    psDH = TxtBuscar.psDH
End Property

Public Property Let psDH(ByVal New_psDH As String)
    TxtBuscar.psDH() = New_psDH
    PropertyChanged "psDH"
End Property
Public Function SumaRow(ByVal pnCol As Long) As Currency
Dim i As Integer
Dim lnSuma As Currency
lnSuma = 0
If pnCol > FlxGd.Cols - 1 Then Exit Function
If FlxGd.TextMatrix(1, 0) = "" Then Exit Function
For i = 1 To FlxGd.Rows - 1
    If IsNumeric(TextMatrix(i, pnCol)) Then
        lnSuma = lnSuma + CCur(IIf(TextMatrix(i, pnCol) = "", "0", TextMatrix(i, pnCol)))
    End If
Next
SumaRow = lnSuma
End Function
'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=FlxGd,FlxGd,-1,ColWidth
Public Property Get ColWidth(ByVal index As Long) As Long
Attribute ColWidth.VB_Description = "Determines the width of the specified column, in Twips. Not available at design time."
    ColWidth = FlxGd.ColWidth(index)
End Property

Public Property Let ColWidth(ByVal index As Long, ByVal New_ColWidth As Long)
Dim lnCol As Long
    FlxGd.ColWidth(index) = New_ColWidth
    PropertyChanged "ColWidth"
    If index = lnColChk Then
        lnCol = FlxGd.Col
        FlxGd.Redraw = False
        If New_ColWidth = 0 Then
            If FlxGd.TextMatrix(1, 0) <> "" Then
                For i = 1 To FlxGd.Rows - 1
                    FlxGd.row = i
                    FlxGd.Col = index
                    Set FlxGd.CellPicture = Nothing
                Next
            End If
        Else
            If FlxGd.TextMatrix(1, 0) <> "" Then
                For i = 1 To FlxGd.Rows - 1
                    FlxGd.row = i
                    FlxGd.Col = index
                    If TextMatrix(i, index) = "." Then
                        Set FlxGd.CellPicture = ImgCheck1
                    Else
                        Set FlxGd.CellPicture = ImgCheck0
                    End If
                Next
            End If
        End If
        FlxGd.Redraw = True
        FlxGd.Col = lnCol
    End If
End Property
'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=FlxGd,FlxGd,-1,RowHeight
Public Property Get RowHeight(ByVal index As Long) As Long
Attribute RowHeight.VB_Description = "Returns or sets the height of the specified row, in Twips. Not available at design time."
    RowHeight = FlxGd.RowHeight(index)
End Property

Public Property Let RowHeight(ByVal index As Long, ByVal New_RowHeight As Long)
    FlxGd.RowHeight(index) = New_RowHeight
    PropertyChanged "RowHeight"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=TxtBuscar,TxtBuscar,-1,TipoBusPers
Public Property Get TipoBusPersona() As TipoBusquedaPersona
    TipoBusPersona = TxtBuscar.TipoBusPers
End Property
Public Property Let TipoBusPersona(ByVal New_TipoBusPersona As TipoBusquedaPersona)
    TxtBuscar.TipoBusPers() = New_TipoBusPersona
    PropertyChanged "TipoBusPersona"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=0,0,0,False
Public Property Get AutoAdd() As Boolean
    AutoAdd = m_AutoAdd
End Property

Public Property Let AutoAdd(ByVal New_AutoAdd As Boolean)
    m_AutoAdd = New_AutoAdd
    PropertyChanged "AutoAdd"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=FlxGd,FlxGd,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Determines the color used to draw text on each part of the Hierarchical FlexGrid."
    ForeColor = FlxGd.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    FlxGd.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property
'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=FlxGd,FlxGd,-1,ForeColorFixed
Public Property Get ForeColorFixed() As OLE_COLOR
Attribute ForeColorFixed.VB_Description = "Determines the color used to draw text on each part of the Hierarchical FlexGrid."
    ForeColorFixed = FlxGd.ForeColorFixed
End Property

Public Property Let ForeColorFixed(ByVal New_ForeColorFixed As OLE_COLOR)
    FlxGd.ForeColorFixed() = New_ForeColorFixed
    PropertyChanged "ForeColorFixed"
End Property
Public Property Get CellForeColor() As OLE_COLOR
Attribute CellForeColor.VB_MemberFlags = "400"
    CellForeColor = FlxGd.CellForeColor
End Property

Public Property Let CellForeColor(ByVal New_CellForeColor As OLE_COLOR)
    FlxGd.CellForeColor() = New_CellForeColor
    PropertyChanged "CellForeColor"
End Property
Public Sub FormatoPersNom(ByVal pnCol As Long, Optional ByVal pbNomApell As Boolean = False)
Dim i As Integer
If FlxGd.TextMatrix(1, 0) = "" Then Exit Sub
For i = 1 To FlxGd.Rows - 1
    FlxGd.TextMatrix(i, pnCol) = PstaNombre(FlxGd.TextMatrix(i, pnCol), pbNomApell)
Next
End Sub
Public Sub TamañoCombo(Optional nTamaño As Long = 200)
    SendMessage cboCelda.hwnd, CB_SETDROPPEDWIDTH, nTamaño, 0
End Sub
'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=TxtBuscar,TxtBuscar,-1,PersPersoneria
Public Property Get PersPersoneria() As PersPersoneria
Attribute PersPersoneria.VB_MemberFlags = "400"
    PersPersoneria = TxtBuscar.PersPersoneria
End Property
Public Property Let PersPersoneria(ByVal New_PersPersoneria As PersPersoneria)
    TxtBuscar.PersPersoneria() = New_PersPersoneria
    PropertyChanged "PersPersoneria"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=0,0,0,False
Public Property Get SoloFila() As Boolean
    SoloFila = m_SoloFila
End Property

Public Property Let SoloFila(ByVal New_SoloFila As Boolean)
    m_SoloFila = New_SoloFila
    PropertyChanged "SoloFila"
     
End Property
''ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
''MappingInfo=FlxGd,FlxGd,-1,ForeColorFixed
Public Property Get CellBackColor() As OLE_COLOR
CellBackColor = FlxGd.CellBackColor
End Property
Public Property Let CellBackColor(ByVal New_CellBackColor As OLE_COLOR)
    FlxGd.CellBackColor() = New_CellBackColor
    PropertyChanged "CellBackColor"
End Property

Private Sub TxtBuscar_Click(psCodigo As String, psDescripcion As String)
    RaiseEvent OnClickTxtBuscar(psCodigo, psDescripcion)
End Sub
Private Function ColumnaAEditar(ByVal pnCol As Long) As Boolean
Dim psColumnas As String
Dim X As Long
Dim lsColumnas() As String
Dim lsColumna As Long
psColumnas = m_ColumnasAEditar
ColumnaAEditar = False
If Len(Trim(psColumnas)) > 0 Then
    For X = 0 To Cols - 1
        vPos = InStr(1, psColumnas, "-", vbTextCompare)
        ReDim Preserve lsColumnas(X)
        lsColumnas(X) = Mid(psColumnas, 1, IIf(vPos > 0, vPos - 1, Len(psColumnas)))
        If X = pnCol Then
            ColumnaAEditar = IIf(lsColumnas(X) = "X", False, True)
            Exit Function
        End If
        If psColumnas <> "" Then
            psColumnas = Mid(psColumnas, IIf(vPos > 0, vPos + 1, Len(psColumnas)))
        End If
    Next X
End If
End Function
Public Function ValidaIngBilletaje(KeyAscii As Integer) As Integer
     If InStr(1, "0123456789+-.", Chr$(KeyAscii)) = 0 And KeyAscii <> 8 Then
        ValidaIngBilletaje = 0
     Else
        ValidaIngBilletaje = KeyAscii
     End If
End Function
Public Function CalEvaluaBil(psCadena As String, Optional pbSoloPositivos As Boolean = True) As Currency
    Dim lsCadena As String
    Dim lsOpe As String
    Dim lsNum1 As String
    Dim lsNum2 As String
    Dim lsResp As String
    
    Dim lnNum1 As String
    Dim lnNum2 As String
    
    Dim lnPosMas As Integer
    Dim lnPosMenos As Integer
    Dim lnPosMasN As Integer
    Dim lnPosMenosN As Integer
    
    psCadena = CalPreCadenaBil(psCadena)
    
    If psCadena = "" Then
        CalEvaluaBil = Format(0, "#0.00")
        'MsgBox "Error de validación.", vbInformation, "Aviso"
        Exit Function
    End If
    
    lsCadena = psCadena
    
    lnPosMas = InStr(1, psCadena, "+")
    lnPosMenos = InStr(1, psCadena, "-")

    If lnPosMas = 0 And lnPosMenos = 0 Then
        If IsNumeric(psCadena) Then
            lnNum1 = CCur(psCadena)
            lsCadena = ""
        Else
            GoTo NOVALIDO
        End If
    ElseIf (lnPosMas < lnPosMenos) And lnPosMas <> 0 Or lnPosMenos = 0 Then
        If IsNumeric(Mid(psCadena, 1, lnPosMas - 1)) Then
            lnNum1 = CCur(Mid(psCadena, 1, lnPosMas - 1))
            lsCadena = Mid(psCadena, lnPosMas)
        Else
            GoTo NOVALIDO
        End If
    ElseIf (lnPosMas > lnPosMenos) And lnPosMenos <> 0 Or lnPosMas = 0 Then
        If IsNumeric(Mid(psCadena, 1, lnPosMenos - 1)) Then
            lnNum1 = CCur(Mid(psCadena, 1, lnPosMenos - 1))
            lsCadena = Mid(psCadena, lnPosMenos)
        Else
            GoTo NOVALIDO
        End If
    End If
    
    While lsCadena <> ""
        CalObtNumBil lsCadena, lsNum2, lsOpe
        If IsNumeric(lsNum2) Then
            lnNum1 = CalGetResOpeBil(Str(lnNum1), lsNum2, lsOpe)
        Else
            GoTo NOVALIDO
        End If
    Wend
    
    If pbSoloPositivos Then
        If lnNum1 < 0 Then
            CalEvaluaBil = Format(0, "#0.00")
            MsgBox "El resultado no puede ser Negativo.", vbInformation, "Aviso"
            Exit Function
        End If
    End If
    
    CalEvaluaBil = Format(lnNum1, "#,#00.00")
    
    Exit Function
    
NOVALIDO:
    CalEvaluaBil = 0
    MsgBox "Error Ud. ha ingresado un valor no Valido", vbInformation, "Aviso"
End Function

Public Function CalObtNumBil(psCadena As String, psNum2 As String, psOpe As String) As String
    Dim lnPosMas As Integer
    Dim lnPosMenos As Integer
    Dim lnPosMasN As Integer
    Dim lnPosMenosN As Integer
    
    psOpe = Mid(psCadena, 1, 1)
    psCadena = Mid(psCadena, 2)
    
    lnPosMas = InStr(1, psCadena, "+")
    lnPosMenos = InStr(1, psCadena, "-")

    If lnPosMas = 0 And lnPosMenos = 0 Then
        psNum2 = psCadena
        psCadena = ""
    ElseIf (lnPosMas < lnPosMenos) And lnPosMas <> 0 Or lnPosMenos = 0 Then
        psNum2 = Mid(psCadena, 1, lnPosMas - 1)
        psCadena = Mid(psCadena, lnPosMas)
    ElseIf (lnPosMas > lnPosMenos) And lnPosMenos <> 0 Or lnPosMas = 0 Then
        psNum2 = Mid(psCadena, 1, lnPosMenos - 1)
        psCadena = Mid(psCadena, lnPosMenos)
    End If
    
End Function

Public Function CalPreCadenaBil(psCadena As String) As String
    Dim i As Integer
    Dim lsCadena As String
 
    psCadena = Trim(psCadena)
    lsCadena = ""
    For i = 1 To Len(psCadena)
        If Mid(psCadena, i, 1) <> " " Then
            lsCadena = lsCadena & Mid(psCadena, i, 1)
        End If
    Next i
    
    If InStr(1, psCadena, "++") = 0 _
    And InStr(1, psCadena, "+-") = 0 _
    And InStr(1, psCadena, "-+") = 0 _
    And InStr(1, psCadena, "--") = 0 Then
        CalPreCadenaBil = lsCadena
    Else
        CalPreCadenaBil = ""
    End If
End Function

Public Function CalGetResOpeBil(psNum1 As String, psNum2 As String, psOpe As String) As Currency
    If psOpe = "+" Then
        CalGetResOpeBil = CCur(psNum1) + CCur(psNum2)
    Else
        CalGetResOpeBil = CCur(psNum1) - CCur(psNum2)
    End If
End Function
'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=FlxGd,FlxGd,-1,RowHeightMin
Public Property Get RowHeightMin() As Long
Attribute RowHeightMin.VB_Description = "Returns or sets a minimum row height for the entire control, in Twips."
    RowHeightMin = FlxGd.RowHeightMin
End Property

Public Property Let RowHeightMin(ByVal New_RowHeightMin As Long)
    If New_RowHeightMin < 0 Then New_RowHeightMin = 0
    FlxGd.RowHeightMin() = New_RowHeightMin
    PropertyChanged "RowHeightMin"
End Property

