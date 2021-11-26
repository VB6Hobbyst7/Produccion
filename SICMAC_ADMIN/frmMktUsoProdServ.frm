VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{DFDE2506-090D-11D5-BEF8-C11EAA34970C}#2.0#0"; "Spinner.ocx"
Begin VB.Form frmMktUsoProdServ 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Uso de Productos y Servicios"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8460
   Icon            =   "frmMktUsoProdServ.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   8460
   Begin TabDlg.SSTab TabActividad 
      Height          =   6330
      Left            =   45
      TabIndex        =   9
      Top             =   45
      Width           =   8370
      _ExtentX        =   14764
      _ExtentY        =   11165
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
      TabCaption(0)   =   "Uso"
      TabPicture(0)   =   "frmMktUsoProdServ.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraActividad"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      Begin VB.Frame Frame2 
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
         Height          =   2805
         Left            =   120
         TabIndex        =   24
         Top             =   3480
         Width           =   8175
         Begin VB.CommandButton btnEliminarItem 
            Caption         =   "&Eliminar Uso"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   120
            TabIndex        =   7
            Top             =   2400
            Width           =   1485
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
            Height          =   330
            Left            =   6600
            TabIndex        =   8
            Top             =   2400
            Width           =   1485
         End
         Begin Sicmact.FlexEdit feProductos 
            Height          =   2130
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Width           =   7905
            _extentx        =   13944
            _extenty        =   3757
            cols0           =   7
            highlight       =   1
            allowuserresizing=   3
            visiblepopmenu  =   -1
            encabezadosnombres=   "#-Producto-Cant.-Fecha-Glosa-ProductoId-Id General"
            encabezadosanchos=   "350-3500-1000-1200-3000-0-0"
            font            =   "frmMktUsoProdServ.frx":0326
            font            =   "frmMktUsoProdServ.frx":034E
            font            =   "frmMktUsoProdServ.frx":0376
            font            =   "frmMktUsoProdServ.frx":039E
            font            =   "frmMktUsoProdServ.frx":03C6
            fontfixed       =   "frmMktUsoProdServ.frx":03EE
            columnasaeditar =   "X-1-2-X-X-X-X"
            textstylefixed  =   3
            listacontroles  =   "0-0-0-0-0-0-0"
            encabezadosalineacion=   "C-L-C-C-L-C-C"
            formatosedit    =   "0-1-0-0-0-0-0"
            textarray0      =   "#"
            lbflexduplicados=   0
            lbbuscaduplicadotext=   -1
            colwidth0       =   345
            rowheight0      =   300
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Productos"
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
         Height          =   1890
         Left            =   120
         TabIndex        =   19
         Top             =   1560
         Width           =   8175
         Begin VB.TextBox txtGlosa 
            Height          =   660
            Left            =   1080
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   5
            Top             =   1080
            Width           =   4170
         End
         Begin VB.CommandButton btnAgregar 
            Caption         =   "&Agregar Uso"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   6600
            TabIndex        =   6
            Top             =   1440
            Width           =   1485
         End
         Begin Sicmact.TxtBuscar TxtProductoCod 
            Height          =   285
            Left            =   1080
            TabIndex        =   2
            Top             =   330
            Width           =   1260
            _extentx        =   1588
            _extenty        =   503
            appearance      =   1
            appearance      =   1
            font            =   "frmMktUsoProdServ.frx":0414
            psraiz          =   "PRODUCTOS Y/O SERVICIOS"
            appearance      =   1
            stitulo         =   ""
            lbultimainstancia=   0
         End
         Begin MSMask.MaskEdBox txtFechaUso 
            Height          =   300
            Left            =   4080
            TabIndex        =   4
            Top             =   720
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
         Begin Spinner.uSpinner usCantidad 
            Height          =   255
            Left            =   1080
            TabIndex        =   3
            Top             =   720
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
         Begin VB.Label Label8 
            Caption         =   "Glosa:"
            Height          =   255
            Left            =   240
            TabIndex        =   26
            Top             =   1200
            Width           =   495
         End
         Begin VB.Label Label7 
            Caption         =   "Fecha Uso:"
            Height          =   255
            Left            =   3120
            TabIndex        =   23
            Top             =   735
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "Cantidad:"
            Height          =   255
            Left            =   240
            TabIndex        =   22
            Top             =   730
            Width           =   735
         End
         Begin VB.Label Label3 
            Caption         =   "Producto:"
            Height          =   255
            Left            =   240
            TabIndex        =   21
            Top             =   360
            Width           =   735
         End
         Begin VB.Label lblProductoNombre 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   2400
            TabIndex        =   20
            Top             =   330
            Width           =   5460
         End
      End
      Begin VB.Frame fraActividad 
         Caption         =   "Actividad"
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
         Height          =   1125
         Left            =   120
         TabIndex        =   10
         Top             =   430
         Width           =   8175
         Begin Sicmact.TxtBuscar TxtActividadCod 
            Height          =   285
            Left            =   1080
            TabIndex        =   1
            Top             =   690
            Width           =   1260
            _extentx        =   1588
            _extenty        =   503
            appearance      =   1
            appearance      =   1
            font            =   "frmMktUsoProdServ.frx":0440
            psraiz          =   "ACTIVIDADES"
            appearance      =   1
            stitulo         =   ""
            lbultimainstancia=   0
         End
         Begin Sicmact.TxtBuscar txtAgencia 
            Height          =   285
            Left            =   1080
            TabIndex        =   0
            Top             =   315
            Width           =   1260
            _extentx        =   2223
            _extenty        =   503
            appearance      =   1
            appearance      =   1
            font            =   "frmMktUsoProdServ.frx":046C
            appearance      =   1
            stitulo         =   ""
            lbultimainstancia=   0
         End
         Begin VB.Label Label10 
            Caption         =   "Agencia:"
            Height          =   255
            Left            =   240
            TabIndex        =   28
            Top             =   345
            Width           =   735
         End
         Begin VB.Label lblAgenciaNombre 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   2400
            TabIndex        =   27
            Top             =   315
            Width           =   5460
         End
         Begin VB.Label lblActividadNombre 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   2400
            TabIndex        =   18
            Top             =   690
            Width           =   5460
         End
         Begin VB.Label Label1 
            Caption         =   "Actividad:"
            Height          =   255
            Left            =   240
            TabIndex        =   11
            Top             =   720
            Width           =   735
         End
      End
      Begin MSComctlLib.ListView lstAhorros 
         Height          =   2790
         Left            =   -74910
         TabIndex        =   12
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
         TabIndex        =   17
         Top             =   3375
         Width           =   2145
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
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "DOLARES"
         Height          =   195
         Left            =   -68475
         TabIndex        =   15
         Top             =   3465
         Width           =   765
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
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "SOLES"
         Height          =   195
         Left            =   -71445
         TabIndex        =   13
         Top             =   3465
         Width           =   525
      End
   End
End
Attribute VB_Name = "frmMktUsoProdServ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim oArAg As New DActualizaDatosArea
    Dim oGasto As New DGastosMarketing
    CentraForm Me
    
    Me.txtAgencia.rs = oArAg.GetAgencias(, , True)
    Me.TxtActividadCod.lbUltimaInstancia = True
    Me.TxtActividadCod.rs = oGasto.RecuperaActividadPaArbol
    Me.TxtProductoCod.lbUltimaInstancia = True
    Me.TxtProductoCod.rs = oGasto.RecuperaProductoPaArbol
End Sub
Private Sub btnAgregar_Click()
    Dim oNGasto As New NGastosMarketing
    Dim rsCompraDisp As New ADODB.Recordset
    Dim lnActividadId As Long, lnProductoId As Long, lnCantidad As Long
    Dim ldFecha As Date
    Dim lsGlosa As String, lsMsgErr As String, lsAgeCod As String

    On Error GoTo ErrorAgregar

    If validaAgregar = False Then Exit Sub
    
    lsAgeCod = Trim(Me.txtAgencia.Text)
    lnActividadId = CLng(Trim(Me.TxtActividadCod.Text))
    lnProductoId = CLng(Trim(Me.TxtProductoCod.Text))
    lnCantidad = CLng(Trim(Me.usCantidad.Valor))
    ldFecha = CDate(Me.txtFechaUso.Text)
    lsGlosa = Trim(Me.txtGlosa.Text)
    
    If validaStock(lnProductoId, lnCantidad) = False Then Exit Sub
    If MsgBox("Esta seguro de agregar el Producto a la Lista?", vbInformation + vbYesNo, "Aviso") = vbNo Then
        Exit Sub
    End If

    Call oNGasto.InsertaProductoEnActividad(lsAgeCod, lnActividadId, lnProductoId, lnCantidad, ldFecha, lsGlosa, lsMsgErr)
    If lsMsgErr <> "" Then
        MsgBox lsMsgErr, vbCritical, "Aviso"
    End If
    limpiar
    Call MuestraProductosDeActividad(lsAgeCod, lnActividadId)
    Exit Sub
ErrorAgregar:
     Err.Raise Err.Number, "Error Agregar Producto a Actividad", Err.Description
End Sub
Private Sub btnEliminarItem_Click()
    Dim oGasto As New DGastosMarketing
    On Error GoTo ErrorEliminar
    If feProductos.TextMatrix(feProductos.Row, 0) = "" Then
        MsgBox "Ud. debe seleccionar el Producto/Servicio a Eliminar", vbInformation, "Aviso"
        feProductos.SetFocus
        Exit Sub
    End If
    
    If MsgBox("Esta seguro de eliminar el Producto de la Lista?", vbInformation + vbYesNo, "Aviso") = vbNo Then
        Exit Sub
    End If
    
    Call oGasto.EliminaProductoEnActividad(CLng(feProductos.TextMatrix(feProductos.Row, 6)))
    limpiar
    Call MuestraProductosDeActividad(Trim(Me.txtAgencia.Text), CLng(Trim(Me.TxtActividadCod.Text)))
    Exit Sub
ErrorEliminar:
    Err.Raise Err.Number, "Error Quitar Producto a Actividad", Err.Description
End Sub
Private Sub btnCerrar_Click()
    Unload Me
End Sub
Private Sub TxtActividadCod_EmiteDatos()
    Call FormateaFlex(feProductos)
    limpiar
    If Me.TxtActividadCod.Text = "" Then Exit Sub
    Me.lblActividadNombre.Caption = Me.TxtActividadCod.psDescripcion
    If Len(Trim(Me.txtAgencia.Text)) = 2 Then
        Call MuestraProductosDeActividad(Me.txtAgencia.Text, CLng(Me.TxtActividadCod.Text))
    End If
End Sub
Private Sub TxtActividadCod_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.TxtProductoCod.SetFocus
    End If
End Sub
Private Sub txtAgencia_EmiteDatos()
    Call FormateaFlex(feProductos)
    limpiar
    If Me.txtAgencia.Text = "" Then Exit Sub
    Me.lblAgenciaNombre.Caption = Me.txtAgencia.psDescripcion
    If Len(Trim(Me.TxtActividadCod.Text)) = 6 Then
        Call MuestraProductosDeActividad(Me.txtAgencia.Text, CLng(Me.TxtActividadCod.Text))
    End If
End Sub
Private Sub txtAgencia_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.TxtActividadCod.SetFocus
    End If
End Sub
Private Sub usCantidad_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        Me.txtFechaUso.SetFocus
    End If
End Sub
Private Sub txtFechaUso_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtGlosa.SetFocus
    End If
End Sub
Private Sub txtFechaUso_LostFocus()
    Dim sCad As String
    sCad = ValidaFecha(txtFechaUso.Text)
    If Not Trim(sCad) = "" Then
        MsgBox sCad, vbInformation, "Aviso"
        txtFechaUso.SetFocus
        Exit Sub
    End If
End Sub
Private Sub txtGlosa_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloLetras2(KeyAscii, True)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Me.btnAgregar.SetFocus
    End If
End Sub
Private Sub TxtProductoCod_EmiteDatos()
    If Me.TxtProductoCod.Text = "" Then Exit Sub
    Me.lblProductoNombre.Caption = Me.TxtProductoCod.psDescripcion
End Sub
Private Sub TxtProductoCod_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.usCantidad.SetFocus
    End If
End Sub
Private Function validaAgregar() As Boolean
    Dim oNGasto As New NGastosMarketing
    Dim i As Integer

    validaAgregar = True
    If Len(Trim(Me.txtAgencia.Text)) = 0 Then
        MsgBox "Falta seleccionar la Agencia", vbInformation, "Aviso"
        Me.txtAgencia.SetFocus
        validaAgregar = False
        Exit Function
    ElseIf Len(Trim(Me.txtAgencia.Text)) <> 2 Then
        MsgBox "Ud. debe seleccionar una Agencia a ultimo Nivel", vbInformation, "Aviso"
        Me.txtAgencia.SetFocus
        validaAgregar = False
        Exit Function
    End If
    If Len(Trim(TxtActividadCod.Text)) = 0 Then
        MsgBox "Falta seleccionar la Actividad", vbInformation, "Aviso"
        Me.TxtActividadCod.SetFocus
        validaAgregar = False
        Exit Function
    ElseIf Len(Trim(TxtActividadCod.Text)) <> 6 Then
        MsgBox "Ud. debe seleccionar una Actividad a ultimo Nivel", vbInformation, "Aviso"
        Me.TxtActividadCod.SetFocus
        validaAgregar = False
        Exit Function
    End If
    If Len(Trim(TxtProductoCod.Text)) = 0 Then
        MsgBox "Falta seleccionar el Producto que se usó", vbInformation, "Aviso"
        Me.TxtProductoCod.SetFocus
        validaAgregar = False
        Exit Function
    ElseIf Len(Trim(TxtProductoCod.Text)) <> 6 Then
        MsgBox "Ud. debe seleccionar un Producto a ultimo Nivel", vbInformation, "Aviso"
        Me.TxtProductoCod.SetFocus
        validaAgregar = False
        Exit Function
    End If
    If CDbl(Trim(Me.usCantidad.Valor)) = 0 Then
        MsgBox "Falta ingresar la cantidad que se uso", vbInformation, "Aviso"
        Me.usCantidad.SetFocus
        validaAgregar = False
        Exit Function
    End If
    If Not IsDate(Trim(Me.txtFechaUso.Text)) Then
        MsgBox "La Fecha de Uso es incorrecta, favor revise", vbInformation, "Aviso"
        Me.txtFechaUso.SetFocus
        validaAgregar = False
        Exit Function
    Else
        If Not oNGasto.FechaEstaEntreFechaActividad(CLng(Trim(Me.TxtActividadCod.Text)), CDate(Trim(Me.txtFechaUso.Text))) Then
            MsgBox "La Fecha de Uso no esta entre la Fecha de Inicio y Fin de la Actividad", vbInformation, "Aviso"
            Me.txtFechaUso.SetFocus
            validaAgregar = False
            Exit Function
        End If
    End If
    If Len(Trim(Me.txtGlosa.Text)) = 0 Then
        MsgBox "Falta ingresar la glosa del Producto usado", vbInformation, "Aviso"
        Me.txtGlosa.SetFocus
        validaAgregar = False
        Exit Function
    End If
    If Not FlexVacio(feProductos) Then
        For i = 1 To feProductos.Rows - 1
            If CLng(Trim(feProductos.TextMatrix(i, 5))) = CLng(Trim(Me.TxtProductoCod.Text)) And DateDiff("D", CDate(Trim(Me.txtFechaUso.Text)), CDate((Trim(feProductos.TextMatrix(i, 3))))) = 0 Then
                MsgBox "El Producto/Servicio que se está agregando a la Lista ya existe con fecha " & Format(Trim(Me.txtFechaUso.Text), "dd/mm/yyyy"), vbInformation, "Aviso"
                feProductos.SetFocus
                feProductos.Row = i
                feProductos.Col = 1
                validaAgregar = False
                Exit Function
            End If
        Next
    End If
End Function
Private Sub MuestraProductosDeActividad(ByVal psAgeCod As String, ByVal pnActividadId As Long)
    Dim oGasto As New DGastosMarketing
    Dim rs As New ADODB.Recordset
    
    Call FormateaFlex(feProductos)
    Set rs = oGasto.RecuperaProductoxActividad(psAgeCod, pnActividadId)
    If Not RSVacio(rs) Then
        Do While Not rs.EOF
            feProductos.AdicionaFila
            feProductos.TextMatrix(feProductos.Row, 1) = rs!cProductoNombre
            feProductos.TextMatrix(feProductos.Row, 2) = rs!nCantidad
            feProductos.TextMatrix(feProductos.Row, 3) = Format(rs!dfecha, "dd/mm/yyyy")
            feProductos.TextMatrix(feProductos.Row, 4) = rs!cComentario
            feProductos.TextMatrix(feProductos.Row, 5) = rs!nProdServId 'Id Producto
            feProductos.TextMatrix(feProductos.Row, 6) = rs!nId 'Id General
            rs.MoveNext
        Loop
    End If
    Set oGasto = Nothing
End Sub
Private Sub limpiar()
    Me.TxtProductoCod.Text = ""
    Me.lblProductoNombre.Caption = ""
    Me.usCantidad.Valor = "0"
    Me.txtFechaUso.Text = "__/__/____"
    Me.txtGlosa.Text = ""
End Sub
Private Function validaStock(ByVal pnProdServId As Long, ByVal pnCantidadSolicitada As Long) As Boolean
    Dim oGasto As New DGastosMarketing
    Dim rs As ADODB.Recordset
    Dim lnCantidadDisponible As Long, lnCantidadSolGrilla As Long
    Dim lsNomProducto As String
    Dim i As Integer
    Set rs = oGasto.RecuperaStockProducto(pnProdServId)

    validaStock = True
    lnCantidadDisponible = rs!nStock
    lsNomProducto = rs!cNombre
    
    lnCantidadSolGrilla = 0
    If lnCantidadDisponible <= 0 Then
        MsgBox "No existe stock disponible para satisfacer el Prod/Serv: " & lsNomProducto, vbInformation, "Aviso"
        Me.usCantidad.SetFocus
        validaStock = False
        Exit Function
    End If
    If pnCantidadSolicitada > lnCantidadDisponible Then
        MsgBox "No existe stock disponible para satisfacer el Prod/Serv: " & lsNomProducto, vbInformation, "Aviso"
        Me.usCantidad.SetFocus
        validaStock = False
        Exit Function
    End If
End Function
