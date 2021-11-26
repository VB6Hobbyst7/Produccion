VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{DFDE2506-090D-11D5-BEF8-C11EAA34970C}#2.0#0"; "Spinner.ocx"
Begin VB.Form frmMktCompras 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro Compras"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6420
   Icon            =   "frmMktCompras.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   6420
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab TabGasto 
      Height          =   2960
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   6330
      _ExtentX        =   11165
      _ExtentY        =   5212
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
      TabCaption(0)   =   "Productos"
      TabPicture(0)   =   "frmMktCompras.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraRegistro"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "btnCerrar"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "btnLimpiar"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "btnRegistrar"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      Begin VB.CommandButton btnRegistrar 
         Caption         =   "&Registrar"
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
         Left            =   3045
         TabIndex        =   6
         Top             =   2550
         Width           =   1000
      End
      Begin VB.CommandButton btnLimpiar 
         Caption         =   "&Limpiar"
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
         Left            =   4140
         TabIndex        =   7
         Top             =   2550
         Width           =   1000
      End
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
         Left            =   5220
         TabIndex        =   8
         Top             =   2550
         Width           =   1000
      End
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
         Height          =   2115
         Left            =   120
         TabIndex        =   16
         Top             =   420
         Width           =   6135
         Begin Spinner.uSpinner usCantidad 
            Height          =   255
            Left            =   4875
            TabIndex        =   5
            Top             =   1200
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            Max             =   9999999
            MaxLength       =   7
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontName        =   "MS Sans Serif"
            FontSize        =   9.75
         End
         Begin VB.ComboBox cboMoneda 
            Height          =   315
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   360
            Width           =   1935
         End
         Begin MSMask.MaskEdBox txtFechaCompra 
            Height          =   300
            Left            =   4800
            TabIndex        =   2
            Top             =   360
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   529
            _Version        =   393216
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
            PromptChar      =   "_"
         End
         Begin Sicmact.TxtBuscar TxtProductoCod 
            Height          =   285
            Left            =   1200
            TabIndex        =   3
            Top             =   790
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   503
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
            psRaiz          =   "PRODUCTOS Y/O SERVICIOS"
            sTitulo         =   ""
            lbUltimaInstancia=   0   'False
         End
         Begin Sicmact.EditMoney txtPrecioUnitario 
            Height          =   255
            Left            =   1200
            TabIndex        =   4
            Top             =   1200
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   450
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483640
            Text            =   "0.00"
            Enabled         =   -1  'True
         End
         Begin Sicmact.EditMoney txtTotal 
            Height          =   255
            Left            =   4080
            TabIndex        =   9
            Top             =   1680
            Width           =   1920
            _ExtentX        =   3387
            _ExtentY        =   450
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483640
            Text            =   "0.00"
         End
         Begin VB.Label lblSimboloMoneda 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3600
            TabIndex        =   24
            Top             =   1680
            Width           =   375
         End
         Begin VB.Label lblProductoNombre 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   2400
            TabIndex        =   23
            Top             =   790
            Width           =   3585
         End
         Begin VB.Label Label10 
            Caption         =   "Total:"
            Height          =   255
            Left            =   3000
            TabIndex        =   22
            Top             =   1700
            Width           =   495
         End
         Begin VB.Label Label9 
            Caption         =   "Moneda:"
            Height          =   255
            Left            =   240
            TabIndex        =   21
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label7 
            Caption         =   "Cantidad:"
            Height          =   255
            Left            =   3360
            TabIndex        =   20
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Producto:"
            Height          =   255
            Left            =   240
            TabIndex        =   19
            Top             =   810
            Width           =   855
         End
         Begin VB.Label Label3 
            Caption         =   "Fecha de Compra:"
            Height          =   255
            Left            =   3360
            TabIndex        =   18
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label8 
            Caption         =   "P. Unitario:"
            Height          =   255
            Left            =   240
            TabIndex        =   17
            Top             =   1200
            Width           =   855
         End
      End
      Begin MSComctlLib.ListView lstAhorros 
         Height          =   2790
         Left            =   -74910
         TabIndex        =   10
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
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "SOLES"
         Height          =   195
         Left            =   -71445
         TabIndex        =   15
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
         TabIndex        =   14
         Top             =   3465
         Width           =   1590
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "DOLARES"
         Height          =   195
         Left            =   -68475
         TabIndex        =   13
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
         TabIndex        =   12
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
         TabIndex        =   11
         Top             =   3375
         Width           =   2145
      End
   End
End
Attribute VB_Name = "frmMktCompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim oGasto As New DGastosMarketing
    Set oGasto = New DGastosMarketing
    CentraForm Me
    CargaMoneda
    Me.TxtProductoCod.lbUltimaInstancia = True
    Me.TxtProductoCod.rs = oGasto.RecuperaProductoPaArbol
End Sub
Private Sub btnRegistrar_Click()
    Dim oGasto As DGastosMarketing
    Dim lnProdServId As Long
    Dim ldFechaCompra As Date
    Dim lnMoneda As Integer
    Dim lnPrecioUnit As Double, lnTotal As Double
    Dim lnCantidad As Long
    
    If validaRegistrar = False Then Exit Sub
    
    On Error GoTo ErrorRegistrar

    Set oGasto = New DGastosMarketing
    lnProdServId = CLng(Trim(Me.TxtProductoCod.Text))
    ldFechaCompra = CDate(Me.txtFechaCompra.Text)
    lnPrecioUnit = CDbl(Trim(Me.txtPrecioUnitario.Text))
    lnMoneda = CInt(Trim(Right(Me.cboMoneda.Text, 5)))
    lnCantidad = CLng(Trim(Me.usCantidad.Valor))
    lnTotal = CDbl(Trim(Me.txtTotal.Text))
    
    If MsgBox("Esta seguro de registrar la Compra?", vbInformation + vbYesNo, "Aviso") = vbNo Then
        Exit Sub
    End If
    
    Call oGasto.InsertaCompra(lnProdServId, ldFechaCompra, lnPrecioUnit, lnMoneda, lnCantidad, lnTotal, gsCodUser)
    MsgBox "Se ha registrado con éxito la Compra", vbInformation, "Aviso"
    btnLimpiar_Click
    Exit Sub
ErrorRegistrar:
    Err.Raise Err.Number, "Error Registrar Compra", Err.Description
End Sub
Private Sub btnLimpiar_Click()
    Me.TxtProductoCod.Text = ""
    Me.lblProductoNombre.Caption = ""
    Me.txtFechaCompra.Text = "__/__/____"
    Me.txtPrecioUnitario.Text = "0.00"
    Me.cboMoneda.ListIndex = -1
    Me.usCantidad.Valor = "0"
    Me.txtTotal.Text = "0.00"
    Me.lblSimboloMoneda.Caption = ""
    Me.txtTotal.BackColor = &HFFFFFF
    Me.cboMoneda.SetFocus
End Sub
Private Sub btnCerrar_Click()
    Unload Me
End Sub
Private Sub CargaMoneda()
    Me.cboMoneda.AddItem "SOLES" & Space(200) & "1"
    Me.cboMoneda.AddItem "DOLARES" & Space(200) & "2"
End Sub
Private Sub txtFechaCompra_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii, True)
    If KeyAscii = 13 Then
        Me.TxtProductoCod.SetFocus
    End If
End Sub
Private Sub txtFechaCompra_LostFocus()
    Dim sCad As String
    sCad = ValidaFecha(txtFechaCompra.Text)
    If Not Trim(sCad) = "" Then
        MsgBox sCad, vbInformation, "Aviso"
        txtFechaCompra.SetFocus
        Exit Sub
    End If
End Sub
Private Sub txtPrecioUnitario_Change()
    Me.txtTotal.Text = Format(CDbl(Trim(Me.usCantidad.Valor)) * CDbl(Trim(Me.txtPrecioUnitario.Text)), "##,##0.00")
End Sub
Private Sub txtPrecioUnitario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.usCantidad.SetFocus
    End If
End Sub
Private Sub TxtProductoCod_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtPrecioUnitario.SetFocus
    End If
End Sub
Private Sub TxtProductoCod_EmiteDatos()
    If Me.TxtProductoCod.Text = "" Then Exit Sub
    Me.lblProductoNombre.Caption = Me.TxtProductoCod.psDescripcion
End Sub
Private Sub cboMoneda_click()
    Dim nMoneda As Moneda
    If cboMoneda.ListIndex = -1 Then Exit Sub
    nMoneda = CLng(Trim(Right(cboMoneda.Text, 2)))
    
    If nMoneda = COMDConstantes.gMonedaNacional Then
        Me.txtTotal.BackColor = &HC0FFFF
        '''Me.lblSimboloMoneda.Caption = "S/." 'marg ers044-2016
        Me.lblSimboloMoneda.Caption = gcPEN_SIMBOLO
    Else
        Me.txtTotal.BackColor = &HC0FFC0
        Me.lblSimboloMoneda.Caption = "US$"
    End If
End Sub
Private Sub cboMoneda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtFechaCompra.SetFocus
    End If
End Sub
Private Sub usCantidad_Change()
    Me.txtTotal.Text = Format(CDbl(IIf(Trim(Me.usCantidad.Valor) = "", 0, Trim(Me.usCantidad.Valor))) * CDbl(Trim(Me.txtPrecioUnitario.Text)), "##,##0.00")
End Sub
Private Sub usCantidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.btnRegistrar.SetFocus
    End If
End Sub

'Private Sub txtCantidad_Change()
'    Dim i As Integer
'    'Me.txtCantidad.Text = IIf(Trim(Me.txtCantidad.Text) = "", 0, Trim(Me.txtCantidad.Text))
''    If txtCantidad.SelStart = 1 And Trim(txtCantidad.Text) = "0" Then
''        'i = Len(Mid(txtApePat.Text, 1, txtApePat.SelStart))
''        txtCantidad.SelStart = 0
''    End If
''    'txtApePat.SelStart = i
''    txtCantidad.Text = CLng(txtCantidad.Text)
'    Me.txtTotal.Text = Format(CDbl(IIf(Trim(Me.txtCantidad.Text) = "", 0, Trim(Me.txtCantidad.Text))) * CDbl(Trim(Me.txtPrecioUnitario.Text)), "##,##0.00")
'End Sub
'Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
'    Dim i As Integer
'    KeyAscii = NumerosEnteros(KeyAscii, False)
'    If KeyAscii = 13 Then
'        Me.btnRegistrar.SetFocus
'    Else
'        If txtCantidad.SelStart > 0 And Trim(txtCantidad.Text) = "0" Then
'            txtCantidad.Text = ""
'        End If
'    End If
'End Sub
Private Function validaRegistrar() As Boolean
    validaRegistrar = True
    If Me.cboMoneda.ListIndex = -1 Then
        MsgBox "Falta seleccionar la Moneda de la Compra", vbInformation, "Aviso"
        Me.cboMoneda.SetFocus
        validaRegistrar = False
        Exit Function
    End If
    If Not IsDate(Trim(Me.txtFechaCompra.Text)) Then
        MsgBox "Falta la Fecha de la Compra", vbInformation, "Aviso"
        Me.txtFechaCompra.SetFocus
        validaRegistrar = False
        Exit Function
    End If
    If Len(Trim(Me.TxtProductoCod.Text)) = 0 Then
        MsgBox "Falta seleccionar el Producto y/o Servicio", vbInformation, "Aviso"
        Me.TxtProductoCod.SetFocus
        validaRegistrar = False
        Exit Function
    ElseIf Len(Trim(Me.TxtProductoCod.Text)) <> 6 Then
        MsgBox "Debe seleccionar un Producto y/o Servicio a ultimo nivel", vbInformation, "Aviso"
        Me.TxtProductoCod.SetFocus
        validaRegistrar = False
        Exit Function
    End If
    If CDbl(Trim(Me.txtPrecioUnitario.Text)) = 0 Then
        MsgBox "Falta ingresar el Precio Unitario del Producto", vbInformation, "Aviso"
        Me.txtPrecioUnitario.SetFocus
        validaRegistrar = False
        Exit Function
    End If
    If CDbl(Trim(Me.usCantidad.Valor)) = 0 Then
        MsgBox "Falta ingresar la Cantidad del Producto", vbInformation, "Aviso"
        Me.usCantidad.SetFocus
        validaRegistrar = False
        Exit Function
    End If
End Function
