VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmGarantiaValorDirectoDetallado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VALOR DIRECTO DETALLADO"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8640
   Icon            =   "frmGarantiaValorDirectoDetallado.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   8640
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   3240
      TabIndex        =   6
      ToolTipText     =   "Aceptar"
      Top             =   5005
      Width           =   1000
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4305
      TabIndex        =   7
      ToolTipText     =   "Cancelar"
      Top             =   5005
      Width           =   1000
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4815
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   8493
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Monto Detallado"
      TabPicture(0)   =   "frmGarantiaValorDirectoDetallado.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraDetalle"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtGlosa"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmbMoneda"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      Begin VB.ComboBox cmbMoneda 
         Height          =   315
         Left            =   6840
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox txtGlosa 
         Height          =   285
         Left            =   720
         TabIndex        =   0
         Tag             =   "txtPrincipal"
         Top             =   480
         Width           =   5130
      End
      Begin VB.Frame fraDetalle 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   3855
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   8175
         Begin VB.CommandButton cmdEliminar 
            Caption         =   "&Eliminar"
            Height          =   350
            Left            =   7165
            TabIndex        =   3
            Top             =   3420
            Width           =   900
         End
         Begin VB.CommandButton cmdAgregar 
            Caption         =   "A&gregar"
            Height          =   350
            Left            =   6220
            TabIndex        =   2
            Top             =   3420
            Width           =   900
         End
         Begin SICMACT.FlexEdit feDetalle 
            Height          =   3135
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   7935
            _ExtentX        =   13996
            _ExtentY        =   5530
            Cols0           =   7
            HighLight       =   1
            EncabezadosNombres=   "#-Descripción-Cantidad-Precio Unit.-Tipo Doc.-N° Doc-Aux"
            EncabezadosAnchos=   "400-2000-1000-1200-1800-1300-0"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-1-2-3-4-5-X"
            ListaControles  =   "0-0-0-0-3-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-C-R-L-L-C"
            FormatosEdit    =   "0-0-3-2-0-0-0"
            CantEntero      =   15
            TextArray0      =   "#"
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   3
            lbBuscaDuplicadoText=   -1  'True
            ColWidth0       =   405
            RowHeight0      =   300
         End
         Begin VB.CommandButton cmdCancel 
            Caption         =   "&Cancelar"
            Height          =   350
            Left            =   7165
            TabIndex        =   5
            Top             =   3420
            Visible         =   0   'False
            Width           =   900
         End
         Begin VB.CommandButton cmdOK 
            Caption         =   "&Aceptar"
            Height          =   350
            Left            =   6220
            TabIndex        =   4
            Top             =   3420
            Visible         =   0   'False
            Width           =   900
         End
         Begin VB.Label txtMonto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   600
            TabIndex        =   14
            Top             =   3420
            Width           =   1380
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "VRA:"
            Height          =   195
            Left            =   120
            TabIndex        =   10
            Top             =   3450
            Width           =   375
         End
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Moneda:"
         Height          =   195
         Left            =   6120
         TabIndex        =   12
         Top             =   495
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Glosa:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   450
      End
   End
End
Attribute VB_Name = "frmGarantiaValorDirectoDetallado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************************************************************
'** Nombre : frmGarantiaValorDirectoDetallado
'** Descripción : Para registro/edición/consulta de Valorización Directa Detallada creado segun TI-ERS063-2014
'** Creación : EJVG, 20130115 04:17:43 PM
'*************************************************************************************************************
Option Explicit

Dim fbRegistrar As Boolean
Dim fbEditar As Boolean
Dim fbConsultar As Boolean

Dim fbOk As Boolean
Dim fbPrimero As Boolean
Dim fnMoneda As Moneda
Dim fsGlosa As String
Dim fvDetalle() As tValorDirectoDetallado
Dim fvDetalle_ULT() As tValorDirectoDetallado
Dim fnNoMoverFila As Integer
Dim fbFocoGrilla As Boolean

Dim fRsDocs As ADODB.Recordset

Public Function Registrar(ByVal pbPrimero As Boolean, ByRef pnMoneda As Moneda, ByRef psGlosa As String, ByRef pvDetalle() As tValorDirectoDetallado, ByRef pvDetalle_ULT() As tValorDirectoDetallado) As Boolean
    fbRegistrar = True
    fbPrimero = pbPrimero
    fnMoneda = pnMoneda
    fsGlosa = psGlosa
    fvDetalle = pvDetalle
    fvDetalle_ULT = pvDetalle_ULT
    Show 1
    pnMoneda = fnMoneda
    psGlosa = fsGlosa
    pvDetalle = fvDetalle
    
    Registrar = fbOk
End Function
Public Function Editar(ByVal pbPrimero As Boolean, ByRef pnMoneda As Moneda, ByRef psGlosa As String, ByRef pvDetalle() As tValorDirectoDetallado, ByRef pvDetalle_ULT() As tValorDirectoDetallado) As Boolean
    fbEditar = True
    fnMoneda = pnMoneda
    fsGlosa = psGlosa
    fbPrimero = pbPrimero
    fvDetalle = pvDetalle
    fvDetalle_ULT = pvDetalle_ULT
    Show 1
    pnMoneda = fnMoneda
    psGlosa = fsGlosa
    pvDetalle = fvDetalle
    
    Editar = fbOk
End Function
Public Sub Consultar(ByVal pnMoneda As Moneda, ByVal pnMonto As Currency, ByVal psGlosa As String, ByRef pvDetalle() As tValorDirectoDetallado)
    fnMoneda = pnMoneda
    fsGlosa = psGlosa
    fvDetalle = pvDetalle
    fbConsultar = True
    Show 1
End Sub
Private Sub cmdCancel_Click()
    SetFlexDetalle
    EditarDetalle False
    fnNoMoverFila = -1
    
    cmdAceptar.Enabled = True
End Sub
Private Sub cmdOK_Click()
    Dim index As Integer
    Dim lsDocTpo As String
    If Not validarDetalle Then Exit Sub
    
    index = UBound(fvDetalle) + 1
    ReDim Preserve fvDetalle(index)
    fvDetalle(index).sDescripcion = feDetalle.TextMatrix(fnNoMoverFila, 1)
    fvDetalle(index).nCantidad = feDetalle.TextMatrix(fnNoMoverFila, 2)
    fvDetalle(index).nValor = feDetalle.TextMatrix(fnNoMoverFila, 3)
    lsDocTpo = feDetalle.TextMatrix(fnNoMoverFila, 4)
    fvDetalle(index).nDocTpo = CInt(Trim(Right(lsDocTpo, 3)))
    fvDetalle(index).sDocTpo = lsDocTpo
    fvDetalle(index).sDocNro = feDetalle.TextMatrix(fnNoMoverFila, 5)
    
    SetFlexDetalle
    EditarDetalle False
    fnNoMoverFila = -1
    
    cmdAceptar.Enabled = True
End Sub
Private Sub feDetalle_GotFocus()
    fbFocoGrilla = True
End Sub
Private Sub feDetalle_LostFocus()
    fbFocoGrilla = False
End Sub

Private Sub feDetalle_OnCellChange(pnRow As Long, pnCol As Long)
    If pnCol = 2 Or pnCol = 3 Then
        SumarizarDetalle
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If fbFocoGrilla Then
        If KeyCode = 86 And Shift = 2 Then
            KeyCode = 10
        End If
    End If
End Sub
Private Sub Form_Load()
    fbOk = False
    Screen.MousePointer = 11
    
    CargarControles
    CargarVariables
    LimpiarControles

    If fbEditar Or fbConsultar Then
        cmbMoneda.ListIndex = IndiceListaCombo(cmbMoneda, fnMoneda)
        txtGlosa.Text = fsGlosa
        SetFlexDetalle
        
        If fbConsultar Then
            cmbMoneda.Enabled = False
            txtGlosa.Enabled = False
            fraDetalle.Enabled = False
            cmdAceptar.Enabled = False
        End If
    End If
    
    If fbPrimero Then
        fnMoneda = gMonedaNacional
    End If
    cmbMoneda.ListIndex = IndiceListaCombo(cmbMoneda, fnMoneda)
    cmbMoneda.Enabled = False
    
    If fbRegistrar Then
        If UBound(fvDetalle_ULT) > 0 Then
            fvDetalle = fvDetalle_ULT
        End If
        SetFlexDetalle
    End If
    
    If fbRegistrar Then
        Caption = "VALOR DIRECTO DETALLADO [ NUEVO ]"
    End If
    If fbConsultar Then
        Caption = "VALOR DIRECTO DETALLADO [ CONSULTAR ]"
    End If
    If fbEditar Then
        Caption = "VALOR DIRECTO DETALLADO [ EDITAR ]"
    End If
    
    Screen.MousePointer = 0
End Sub
Private Sub CargarVariables()
    Dim oCons As New COMDConstantes.DCOMConstantes
    Set fRsDocs = oCons.RecuperaConstantes(gColocPigTipoDocumento)
    Set oCons = Nothing
End Sub
Private Sub CargarControles()
    Dim oCons As New COMDConstantes.DCOMConstantes
    Dim rsMoneda As New ADODB.Recordset
    
    Set rsMoneda = oCons.RecuperaConstantes(1011)
    Call Llenar_Combo_con_Recordset(rsMoneda, cmbMoneda)
    
    RSClose rsMoneda
    Set oCons = Nothing
End Sub
Private Sub LimpiarControles()
    cmbMoneda.ListIndex = -1
    txtMonto.Caption = "0.00"
    txtGlosa.Text = ""
    FormateaFlex feDetalle
End Sub
Private Sub txtGlosa_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii, True)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If cmbMoneda.Enabled And cmbMoneda.Visible Then
            EnfocaControl cmbMoneda
        Else
            EnfocaControl cmdAgregar
        End If
    End If
End Sub
Private Sub cmbMoneda_Click()
    Dim lnMoneda As Moneda
    lnMoneda = val(Trim(Right(cmbMoneda.Text, 3)))
    If lnMoneda = gMonedaNacional Then
        txtMonto.BackColor = &H80000005
    ElseIf lnMoneda = gMonedaExtranjera Then
        txtMonto.BackColor = &HC0FFC0
    End If
End Sub
Private Sub cmbMoneda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EnfocaControl cmdAgregar
    End If
End Sub
Private Sub cmdCancelar_Click()
    fbOk = False
    Unload Me
End Sub
Private Sub cmdAceptar_Click()
    If Len(Trim(txtGlosa.Text)) = 0 Then
        MsgBox "Ud. debe ingresar la Glosa", vbInformation, "Aviso"
        EnfocaControl txtGlosa
        Exit Sub
    End If
    If cmbMoneda.ListIndex = -1 Then
        MsgBox "Ud. debe seleccionar la Moneda", vbInformation, "Aviso"
        EnfocaControl cmbMoneda
        Exit Sub
    End If
    If Not validarDetalle Then Exit Sub
    
    fsGlosa = Trim(txtGlosa.Text)
    fnMoneda = Trim(Right(cmbMoneda.Text, 3))
    
    fbOk = True
    Unload Me
End Sub
Private Sub cmdEliminar_Click()
    Dim lvDetalleTmp() As tValorDirectoDetallado
    Dim index As Integer, indexTmp As Integer
    
    On Error GoTo ErrEliminar
    If feDetalle.TextMatrix(1, 0) = "" Then Exit Sub
    If MsgBox("¿Desea quitar [" & fvDetalle(feDetalle.row).sDescripcion & "] del detalle?", vbYesNo + vbInformation, "Aviso") = vbNo Then Exit Sub
    
    ReDim lvDetalleTmp(0)
    For index = 1 To UBound(fvDetalle)
        ReDim Preserve lvDetalleTmp(index)
        lvDetalleTmp(index) = fvDetalle(index)
    Next
    
    index = 0
    ReDim fvDetalle(0)
    For indexTmp = 1 To UBound(lvDetalleTmp)
        If indexTmp <> feDetalle.row Then
            index = index + 1
            ReDim Preserve fvDetalle(index)
            fvDetalle(index) = lvDetalleTmp(indexTmp)
        End If
    Next
    Erase lvDetalleTmp
    
    SetFlexDetalle
    Exit Sub
ErrEliminar:
    Screen.MousePointer = 0
    MsgBox err.Description, vbCritical, "Aviso"
End Sub
Private Sub cmdAgregar_Click()
    If feDetalle.TextMatrix(1, 0) <> "" Then
        If Not validarDetalle Then Exit Sub
    End If
    
    feDetalle.AdicionaFila
    feDetalle.SetFocus
    SendKeys "{ENTER}"
    
    EditarDetalle True
    fnNoMoverFila = feDetalle.row
    
    cmdAceptar.Enabled = False
End Sub
Public Function validarDetalle() As Boolean
    Dim i As Integer, j As Integer
    
    If feDetalle.TextMatrix(1, 0) = "" Then
        MsgBox "Ud. debe ingresar el Detalle de la Valorización", vbInformation, "Aviso"
        EnfocaControl cmdAgregar
        Exit Function
    End If
    For i = 1 To feDetalle.Rows - 1
        For j = 1 To feDetalle.Cols - 1
            If feDetalle.ColWidth(j) > 0 Then
                If Len(Trim(feDetalle.TextMatrix(i, j))) = 0 Then
                    MsgBox "El campo " & UCase(feDetalle.TextMatrix(0, j)) & " está vacio, verifique..", vbInformation, "Aviso"
                    EnfocaControl feDetalle
                    feDetalle.TopRow = i
                    feDetalle.row = i
                    feDetalle.Col = j
                    Exit Function
                End If
            End If
        Next
    Next
    For i = 1 To feDetalle.Rows - 1
        For j = 1 To feDetalle.Cols - 1
            If j = 2 Or j = 3 Then 'Valida se hayan ingresado Cantidad
                If val(feDetalle.TextMatrix(i, j)) = 0 Then
                    MsgBox "Ud. debe ingresar un valor mayor a cero", vbInformation, "Aviso"
                    EnfocaControl feDetalle
                    feDetalle.TopRow = i
                    feDetalle.row = i
                    feDetalle.Col = j
                    Exit Function
                End If
            End If
        Next
    Next
    validarDetalle = True
End Function
Private Sub SetFlexDetalle()
    Dim index As Integer
    
    FormateaFlex feDetalle
    For index = 1 To UBound(fvDetalle)
        feDetalle.AdicionaFila
        feDetalle.TextMatrix(index, 1) = fvDetalle(index).sDescripcion
        feDetalle.TextMatrix(index, 2) = fvDetalle(index).nCantidad
        feDetalle.TextMatrix(index, 3) = Format(fvDetalle(index).nValor, "#,##0.00")
        feDetalle.TextMatrix(index, 4) = fvDetalle(index).sDocTpo
        feDetalle.TextMatrix(index, 5) = fvDetalle(index).sDocNro
    Next
    SumarizarDetalle
End Sub
Private Sub EditarDetalle(ByVal pbEditar As Boolean)
    cmdAgregar.Visible = Not pbEditar
    cmdEliminar.Visible = Not pbEditar
    cmdOK.Visible = pbEditar
    cmdCancel.Visible = pbEditar
    
    feDetalle.lbEditarFlex = pbEditar
End Sub
Private Sub feDetalle_RowColChange()
    Dim rs As ADODB.Recordset
    If feDetalle.lbEditarFlex Then
        If feDetalle.Col = 4 Then
            Set rs = fRsDocs.Clone
            feDetalle.CargaCombo rs
        End If
        feDetalle.row = fnNoMoverFila
    End If
    feDetalle.TextMatrix(feDetalle.row, 1) = UCase(feDetalle.TextMatrix(feDetalle.row, 1))
    'If feDetalle.Col = 2 Or feDetalle.Col = 3 Then
    '    SumarizarDetalle
    'End If
    RSClose rs
End Sub
Private Sub txtGlosa_LostFocus()
    txtGlosa.Text = UCase(Trim(txtGlosa.Text))
End Sub
Private Sub SumarizarDetalle()
    Dim i As Integer
    Dim lnCantidad As Integer
    Dim lnPrecioUnit As Currency
    Dim lnTotal As Currency
    
    If feDetalle.TextMatrix(1, 0) <> "" Then
        For i = 1 To feDetalle.Rows - 1
            lnCantidad = IIf(IsNumeric(feDetalle.TextMatrix(i, 2)), feDetalle.TextMatrix(i, 2), 0)
            lnPrecioUnit = IIf(IsNumeric(feDetalle.TextMatrix(i, 3)), feDetalle.TextMatrix(i, 3), 0)
            lnTotal = lnTotal + lnCantidad * lnPrecioUnit
        Next
    End If
    
    txtMonto.Caption = Format(lnTotal, "#,##0.00")
End Sub
