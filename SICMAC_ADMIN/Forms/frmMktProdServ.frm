VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMktProdServ 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Productos y Servicios"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8295
   Icon            =   "frmMktProdServ.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   8295
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnCerrar 
      Cancel          =   -1  'True
      Caption         =   "&Cerrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   7250
      TabIndex        =   7
      Top             =   4800
      Width           =   1000
   End
   Begin VB.CommandButton btnEditar 
      Caption         =   "&Editar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   1155
      TabIndex        =   6
      Top             =   4800
      Width           =   1000
   End
   Begin VB.CommandButton btnNuevo 
      Caption         =   "&Nuevo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   75
      TabIndex        =   5
      Top             =   4800
      Width           =   1000
   End
   Begin TabDlg.SSTab TabGasto 
      Height          =   4695
      Left            =   40
      TabIndex        =   0
      Top             =   40
      Width           =   8250
      _ExtentX        =   14552
      _ExtentY        =   8281
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   6
      TabHeight       =   617
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Productos y Servicios"
      TabPicture(0)   =   "frmMktProdServ.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "feProductoServicio"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraRegistro"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin VB.Frame fraRegistro 
         Caption         =   "Registro"
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
         Height          =   1540
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   8055
         Begin VB.TextBox txtProducto 
            Height          =   300
            Left            =   1920
            MaxLength       =   199
            TabIndex        =   2
            Top             =   720
            Width           =   3255
         End
         Begin VB.CommandButton btnCancelar 
            Caption         =   "&Cancelar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   6960
            TabIndex        =   9
            Top             =   1080
            Width           =   1000
         End
         Begin VB.CommandButton btnGuardar 
            Caption         =   "&Guardar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   5880
            TabIndex        =   4
            Top             =   1080
            Width           =   1000
         End
         Begin VB.ComboBox cboCategoria 
            Height          =   315
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   340
            Width           =   3225
         End
         Begin VB.ComboBox cboUnidad 
            Height          =   315
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   1080
            Width           =   2025
         End
         Begin VB.Label Label1 
            Caption         =   "Categoría:"
            Height          =   255
            Left            =   480
            TabIndex        =   12
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label2 
            Caption         =   "Producto/Servicio:"
            Height          =   255
            Left            =   480
            TabIndex        =   11
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label Label3 
            Caption         =   "Unidad:"
            Height          =   255
            Left            =   480
            TabIndex        =   10
            Top             =   1110
            Width           =   615
         End
      End
      Begin MSComctlLib.ListView lstAhorros 
         Height          =   2790
         Left            =   -74910
         TabIndex        =   13
         Top             =   495
         Width           =   9390
         _ExtentX        =   16563
         _ExtentY        =   4921
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         Enabled         =   0   'False
         NumItems        =   12
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nro."
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Fecha"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Producto"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Agencia"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Nro. Cuenta"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Nro. Cta Antigua"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Estado"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Participación"
            Object.Width           =   2470
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "SaldoCont"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   9
            Text            =   "SaldoDisp"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Motivo de Bloque"
            Object.Width           =   7231
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Moneda"
            Object.Width           =   2540
         EndProperty
      End
      Begin Sicmact.FlexEdit feProductoServicio 
         Height          =   2730
         Left            =   120
         TabIndex        =   14
         Top             =   1920
         Width           =   8025
         _ExtentX        =   14155
         _ExtentY        =   4815
         Cols0           =   7
         HighLight       =   1
         AllowUserResizing=   3
         VisiblePopMenu  =   -1  'True
         EncabezadosNombres=   "#-Categoría-Producto-Unidad-ProductoId-CategoriaId-UnidadId"
         EncabezadosAnchos=   "350-2100-4100-1200-0-0-0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
         ColumnasAEditar =   "X-1-2-X-X-X-X"
         TextStyleFixed  =   3
         ListaControles  =   "0-0-0-0-0-0-0"
         EncabezadosAlineacion=   "C-L-L-C-C-C-C"
         FormatosEdit    =   "0-1-1-0-0-0-0"
         TextArray0      =   "#"
         lbFlexDuplicados=   0   'False
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   345
         RowHeight0      =   300
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "SOLES"
         Height          =   195
         Left            =   -71445
         TabIndex        =   19
         Top             =   3465
         Width           =   525
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "TOTAL AHORROS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -73185
         TabIndex        =   18
         Top             =   3465
         Width           =   1590
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "DOLARES"
         Height          =   195
         Left            =   -68475
         TabIndex        =   17
         Top             =   3465
         Width           =   765
      End
      Begin VB.Label lblSolesAho 
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
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   -70815
         TabIndex        =   16
         Top             =   3375
         Width           =   2145
      End
      Begin VB.Label lblDolaresAho 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
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
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   -67680
         TabIndex        =   15
         Top             =   3375
         Width           =   2145
      End
   End
End
Attribute VB_Name = "frmMktProdServ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fbNuevo As Boolean
Dim fnCodProductoServicio As Long

Private Sub Form_Load()
    CentraForm Me
    ListarCategoriasActivas
    CargaUnidades
    MuestraProductosServicios
    fnCodProductoServicio = 0
    fbNuevo = True
End Sub
Private Sub btnEditar_Click()
    If feProductoServicio.TextMatrix(feProductoServicio.Row, 0) = "" Then
        MsgBox "Ud. debe seleccionar el Producto/Servicio a Editar", vbInformation, "Aviso"
        feProductoServicio.SetFocus
        Exit Sub
    End If
    
    fnCodProductoServicio = Me.feProductoServicio.TextMatrix(Me.feProductoServicio.Row, 4)
    Me.cboCategoria.ListIndex = IndiceListaCombo(cboCategoria, CLng(Trim(feProductoServicio.TextMatrix(feProductoServicio.Row, 5))))
    Me.txtProducto.Text = feProductoServicio.TextMatrix(feProductoServicio.Row, 2)
    Me.cboUnidad.ListIndex = IndiceListaCombo(cboUnidad, CLng(Trim(feProductoServicio.TextMatrix(feProductoServicio.Row, 6))))
    fbNuevo = False
End Sub
Private Sub btnGuardar_Click()
    Dim oGasto As DGastosMarketing
    Dim lsNombreProducto As String
    Dim lnCategoriaGasto As Long
    Dim lnUnidad As Integer

    If validaGrabar = False Then Exit Sub

    lsNombreProducto = UCase(Trim(Me.txtProducto.Text))
    lnCategoriaGasto = CLng(Trim(Right(Me.cboCategoria.Text, 5)))
    lnUnidad = CInt(Trim(Right(Me.cboUnidad.Text, 5)))

    Set oGasto = New DGastosMarketing

    If MsgBox("Esta seguro de guardar el Producto y/o Servicio?", vbInformation + vbYesNo, "Aviso") = vbNo Then
        Exit Sub
    End If

    If fbNuevo Then
        Call oGasto.InsertaProductoServicio(lsNombreProducto, lnCategoriaGasto, lnUnidad)
    Else
        If fnCodProductoServicio = 0 Then
            MsgBox "Ud. debe seleccionar el Producto o Servicio a editar", vbInformation, "Aviso"
            Me.feProductoServicio.SetFocus
            Exit Sub
        End If
        Call oGasto.ActualizaProductoServicio(fnCodProductoServicio, lsNombreProducto, lnCategoriaGasto, lnUnidad)
    End If

    MsgBox "Se ha grabado con éxito el Producto/Servicio", vbInformation, "Aviso"
    limpiar
    MuestraProductosServicios
End Sub
Private Sub ListarCategoriasActivas()
    Dim oGasto As DGastosMarketing
    Dim rs As ADODB.Recordset
    Set oGasto = New DGastosMarketing
    Set rs = New ADODB.Recordset
    
    Set rs = oGasto.RecuperaCategoriaGastoxEstado(True)
    Do While Not rs.EOF
        cboCategoria.AddItem Trim(rs!cNombre) & Space(200) & Trim(rs!nId)
        rs.MoveNext
    Loop
End Sub
Private Sub MuestraProductosServicios()
    Dim oGasto As DGastosMarketing
    Dim rsProdServ As ADODB.Recordset
    Set oGasto = New DGastosMarketing
    Set rsProdServ = New ADODB.Recordset

    Call FormateaFlex(feProductoServicio)
    Set rsProdServ = oGasto.RecuperaProductoServicio
    If Not RSVacio(rsProdServ) Then
        Do While Not rsProdServ.EOF
            feProductoServicio.AdicionaFila
            feProductoServicio.TextMatrix(feProductoServicio.Row, 1) = rsProdServ!cCatGastoNombre
            feProductoServicio.TextMatrix(feProductoServicio.Row, 2) = rsProdServ!cNombre
            feProductoServicio.TextMatrix(feProductoServicio.Row, 3) = rsProdServ!cUnidad
            feProductoServicio.TextMatrix(feProductoServicio.Row, 4) = rsProdServ!nProdServId
            feProductoServicio.TextMatrix(feProductoServicio.Row, 5) = rsProdServ!nCatGastoId
            feProductoServicio.TextMatrix(feProductoServicio.Row, 6) = rsProdServ!nUnidad
            rsProdServ.MoveNext
        Loop
    End If
    Set oGasto = Nothing
    Set rsProdServ = Nothing
End Sub
Private Sub btnCancelar_Click()
    limpiar
End Sub
Private Sub btnNuevo_Click()
    limpiar
    Me.cboCategoria.SetFocus
End Sub
Private Sub limpiar()
    Me.cboCategoria.ListIndex = -1
    Me.txtProducto.Text = ""
    Me.cboUnidad.ListIndex = -1
    Me.cboCategoria.SetFocus
    fnCodProductoServicio = 0
    fbNuevo = True
End Sub
Private Function validaGrabar() As Boolean
    Dim i As Integer
    validaGrabar = True
    If Me.cboCategoria.ListIndex = -1 Then
        MsgBox "Falta seleccionar la Categoría de Gasto", vbInformation, "Aviso"
        Me.cboCategoria.SetFocus
        validaGrabar = False
        Exit Function
    End If
    If Len(Trim(Me.txtProducto.Text)) = 0 Then
        MsgBox "Falta ingresar el Nombre del Producto/Servicio", vbInformation, "Aviso"
        Me.txtProducto.SetFocus
        validaGrabar = False
        Exit Function
    End If
    If Me.cboUnidad.ListIndex = -1 Then
        MsgBox "Falta seleccionar la Unidad del Producto", vbInformation, "Aviso"
        Me.cboUnidad.SetFocus
        validaGrabar = False
        Exit Function
    End If
    If Not (feProductoServicio.Rows - 1 = 1 And feProductoServicio.TextMatrix(1, 0) = "") Then
        For i = 1 To feProductoServicio.Rows - 1
            If CLng(Trim(feProductoServicio.TextMatrix(i, 5))) = CLng(Trim(Right(Me.cboCategoria.Text, 5))) Then
                If Trim(feProductoServicio.TextMatrix(i, 2)) = Trim(Me.txtProducto.Text) And CLng(Trim(feProductoServicio.TextMatrix(i, 6))) = CLng(Trim(Right(Me.cboUnidad.Text, 5))) Then
                    If fbNuevo Then
                        MsgBox "El Producto/Servicio que se está creando ya existe en la categoria", vbInformation, "Aviso"
                        feProductoServicio.SetFocus
                        feProductoServicio.Row = i
                        feProductoServicio.Col = 2
                        validaGrabar = False
                        Exit Function
                    Else
                        If CLng(Trim(feProductoServicio.TextMatrix(i, 4))) <> fnCodProductoServicio Then
                            MsgBox "El Producto/Servicio que se está editando ya existe en la categoria", vbInformation, "Aviso"
                            feProductoServicio.SetFocus
                            feProductoServicio.Row = i
                            feProductoServicio.Col = 2
                            validaGrabar = False
                            Exit Function
                        End If
                    End If
                End If
            End If
        Next
    End If
End Function
Private Sub btnCerrar_Click()
    Unload Me
End Sub
Private Sub cboCategoria_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtProducto.SetFocus
    End If
End Sub
Private Sub txtProducto_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii, True)
    If KeyAscii = 13 Then
        Me.cboUnidad.SetFocus
    End If
End Sub
Private Sub cboUnidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        btnGuardar.SetFocus
    End If
End Sub
Private Sub CargaUnidades()
    Dim oGasto As DGastosMarketing
    Dim rs As ADODB.Recordset
    Set oGasto = New DGastosMarketing
    Set rs = New ADODB.Recordset
    Set rs = oGasto.RecuperaUnidadxEstado(True)
    Do While Not rs.EOF
        Me.cboUnidad.AddItem Trim(rs!cUnidad) & Space(200) & Trim(rs!nId)
        rs.MoveNext
    Loop
    Set oGasto = Nothing
End Sub
