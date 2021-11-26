VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCapRegVouDep_NEW 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Voucher de Depósito"
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7920
   Icon            =   "frmCapRegVouDep_NEW.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   7920
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   345
      Left            =   6840
      TabIndex        =   17
      Top             =   6780
      Width           =   1050
   End
   Begin VB.CommandButton cmdLimpiar 
      Caption         =   "&Limpiar"
      Height          =   345
      Left            =   4740
      TabIndex        =   16
      Top             =   6780
      Width           =   1050
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   345
      Left            =   5790
      TabIndex        =   15
      Top             =   6780
      Width           =   1050
   End
   Begin TabDlg.SSTab SSTab2 
      Height          =   3135
      Left            =   60
      TabIndex        =   30
      Top             =   3600
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   5530
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "&Datos del Depósito"
      TabPicture(0)   =   "frmCapRegVouDep_NEW.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdEditar"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdAgregar"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdCancelar"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdAceptar"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "feOperaciones"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdQuitar"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtOpeTotal"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      Begin VB.TextBox txtOpeTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   6120
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   14
         Text            =   "0.00"
         Top             =   2770
         Width           =   1580
      End
      Begin VB.CommandButton cmdQuitar 
         Caption         =   "&Quitar"
         Height          =   300
         Left            =   1890
         TabIndex        =   13
         ToolTipText     =   "WSASASASASAS"
         Top             =   2770
         Width           =   885
      End
      Begin SICMACT.FlexEdit feOperaciones 
         Height          =   2250
         Left            =   120
         TabIndex        =   11
         ToolTipText     =   "ESTO ES UNA PRUEBA"
         Top             =   480
         Width           =   7620
         _ExtentX        =   13441
         _ExtentY        =   3969
         Cols0           =   9
         HighLight       =   2
         AllowUserResizing=   3
         EncabezadosNombres=   "N°-Código-Operación-Monto-Estado-Detalle de la Operación-OperacionTmp-Glosa-Aux"
         EncabezadosAnchos=   "0-1200-3000-1200-1200-2800-0-2500-0"
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
         ColumnasAEditar =   "X-X-2-3-X-5-6-7-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-0-3-0-0-1-0-0-0"
         EncabezadosAlineacion=   "C-L-L-R-C-L-C-L-C"
         FormatosEdit    =   "0-0-0-2-0-0-0-0-0"
         CantEntero      =   10
         TextArray0      =   "N°"
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   6
         lbFormatoCol    =   -1  'True
         lbPuntero       =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         RowHeight0      =   300
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   300
         Left            =   3240
         TabIndex        =   34
         Top             =   2770
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   300
         Left            =   4125
         TabIndex        =   33
         Top             =   2770
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "&Agregar"
         Height          =   300
         Left            =   120
         TabIndex        =   12
         Top             =   2770
         Width           =   885
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "&Editar"
         Height          =   300
         Left            =   1000
         TabIndex        =   32
         Top             =   2770
         Width           =   885
      End
      Begin VB.Label Label4 
         Caption         =   "Total:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5640
         TabIndex        =   31
         Top             =   2805
         Width           =   495
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3135
      Left            =   60
      TabIndex        =   19
      Top             =   360
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   5530
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "&Datos del Voucher"
      TabPicture(0)   =   "frmCapRegVouDep_NEW.frx":0326
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      Begin VB.Frame Frame3 
         Caption         =   "Moneda"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1020
         Left            =   5040
         TabIndex        =   27
         Top             =   360
         Width           =   2655
         Begin VB.ComboBox cboMoneda 
            Height          =   315
            ItemData        =   "frmCapRegVouDep_NEW.frx":0342
            Left            =   960
            List            =   "frmCapRegVouDep_NEW.frx":0344
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   240
            Width           =   1600
         End
         Begin VB.TextBox txtVoucherMonto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Left            =   960
            MaxLength       =   15
            TabIndex        =   4
            Text            =   "0.00"
            Top             =   600
            Width           =   1580
         End
         Begin VB.Label Label6 
            Caption         =   "Moneda:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label9 
            Caption         =   "Monto:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   600
            Width           =   615
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Titular del Voucher"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1020
         Left            =   80
         TabIndex        =   24
         Top             =   360
         Width           =   4935
         Begin VB.TextBox TxtTitPersNombre 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   2
            Top             =   600
            Width           =   3825
         End
         Begin SICMACT.TxtBuscar TxtTitPersCod 
            Height          =   285
            Left            =   960
            TabIndex        =   1
            Top             =   240
            Width           =   1860
            _ExtentX        =   3281
            _ExtentY        =   503
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   0
            TipoBusqueda    =   3
            sTitulo         =   ""
         End
         Begin VB.Label Label3 
            Caption         =   "Nombre:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label2 
            Caption         =   "Persona:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Datos del Depósito"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1650
         Left            =   80
         TabIndex        =   20
         Top             =   1365
         Width           =   7620
         Begin VB.TextBox txtCtaIFDesc 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   7
            Top             =   600
            Width           =   6555
         End
         Begin VB.TextBox txtCtaIFBancoNombre 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2760
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   240
            Width           =   4755
         End
         Begin VB.TextBox txtVoucherNro 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3240
            TabIndex        =   9
            Top             =   960
            Width           =   4275
         End
         Begin VB.CheckBox chkConfirmar 
            Caption         =   "&Depósito Confirmado (web)"
            Height          =   255
            Left            =   5280
            TabIndex        =   10
            Top             =   1320
            Width           =   2295
         End
         Begin SICMACT.TxtBuscar txtCtaIFCod 
            Height          =   285
            Left            =   960
            TabIndex        =   5
            Top             =   240
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   503
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   0
            sTitulo         =   ""
            EnabledText     =   0   'False
         End
         Begin MSMask.MaskEdBox txtVoucherFecha 
            Height          =   285
            Left            =   960
            TabIndex        =   8
            Top             =   960
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            ForeColor       =   -2147483630
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
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
         Begin VB.Label Label10 
            Caption         =   "Banco:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label11 
            Caption         =   "Fecha Depósito:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   22
            Top             =   840
            Width           =   735
         End
         Begin VB.Label Label5 
            Caption         =   "N° Voucher:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2280
            TabIndex        =   21
            Top             =   975
            Width           =   1095
         End
      End
   End
   Begin MSMask.MaskEdBox txtFechaMov 
      Height          =   285
      Left            =   6600
      TabIndex        =   0
      Top             =   75
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
      ForeColor       =   -2147483630
      Enabled         =   0   'False
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
   Begin VB.Label Label1 
      Caption         =   "Fecha:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6000
      TabIndex        =   18
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frmCapRegVouDep_NEW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************
'** Nombre : frmCapRegVouDep_NEW
'** Descripción : Para registro de Vouchers creado segun TI-ERS044-2013
'** Creación : EJVG, 20130904 11:45:00 AM
'**********************************************************************
Option Explicit
Public fnMovNroPen As Long
Dim fnId As Long
Dim fnFilaActual As Integer
Dim fbNuevo As Boolean
Dim fMatOperaciones As Variant

Dim fsTpoProgramaAho As String
Dim fsTpoProgramaDPF As String
Dim fsTpoProgramaCTS As String

Private Sub Form_Load()
    Dim oConsSis As New COMDConstSistema.NCOMConstSistema
    fsTpoProgramaAho = oConsSis.LeeConstSistema(439)
    fsTpoProgramaDPF = oConsSis.LeeConstSistema(440)
    fsTpoProgramaCTS = oConsSis.LeeConstSistema(441)
    cargarControles
    LimpiarControles
    Set oConsSis = Nothing
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If fbNuevo Then
        If MsgBox("¿Esta seguro de salir del Registro de Voucher de Depósito?", vbYesNo + vbInformation, "Aviso") = vbNo Then
            Cancel = 1
            Exit Sub
        End If
    End If
End Sub
Public Sub Nuevo()
    fbNuevo = True
    feOperaciones.EncabezadosAnchos = "0-1200-2300-1200-0-2800-0-2500-0"
    feOperaciones.ColumnasAEditar = "X-X-2-3-X-5-6-7-X"
    Show 1
End Sub
Public Sub Editar(ByVal pnId As Long)
    fbNuevo = False
    fnId = pnId
    feOperaciones.EncabezadosAnchos = "0-1200-2300-1200-1200-2800-0-2500-0"
    'feOperaciones.ColumnasAEditar = "X-X-2-X-X-5-6-7-X"
    feOperaciones.ColumnasAEditar = "X-X-2-3-X-5-6-7-X"
    If Not cargarDatos(fnId) Then
        MsgBox "No se pudo cargar los datos del Voucher, si el problema persiste comuniquese con el Dpto. de TI", vbInformation, "Aviso"
        Unload Me
    End If
    BloqueaControles (False)
    cmdAgregar.Visible = False
    cmdQuitar.Visible = False
    cmdLimpiar.Visible = False
    Show 1
End Sub
Private Sub cargarControles()
    CargaMoneda
End Sub
Private Function cargarDatos(ByVal pnId As Long) As Boolean
    Dim oCap As New COMNCaptaGenerales.NCOMCaptaGenerales
    Dim rsCab As New ADODB.Recordset
    Dim rsDet As New ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo ErrCargarDatos
    Set rsCab = oCap.RecuperaVoucherxId(pnId)
    If Not RSVacio(rsCab) Then
        txtFechaMov.Text = Format(rsCab!dfecReg)
        TxtTitPersCod.Text = rsCab!cPersCodCli
        TxtTitPersNombre.Text = rsCab!cPersNombreCli
        cboMoneda.ListIndex = IndiceListaCombo(cboMoneda, rsCab!cmoneda)
        txtCtaIFCod.Text = rsCab!cIFTpo & "." & rsCab!cPersCodIF & "." & rsCab!cCtaIFCod
        txtCtaIFCod_EmiteDatos
        txtVoucherNro.Text = rsCab!cNroVou
        txtVoucherMonto.Text = Format(rsCab!nMonVou, gsFormatoNumeroView)
        txtVoucherFecha.Text = Format(rsCab!dFecVou, gsFormatoFechaView)
        chkConfirmar.value = IIf(rsCab!bConfirmado = True, 1, 0)
        
        Set rsDet = oCap.RecuperaVoucherOpexId(pnId)
        Set fMatOperaciones = Nothing
        ReDim fMatOperaciones(6, 0)
        For i = 1 To rsDet.RecordCount
            ReDim Preserve fMatOperaciones(6, i)
            fMatOperaciones(1, i) = rsDet!cNroVou 'Nro Voucher
            fMatOperaciones(2, i) = rsDet!cTipMot & space(75) & rsDet!nTipMot 'Tipo Operación
            fMatOperaciones(3, i) = Format(rsDet!nMonVou, gsFormatoNumeroView) 'Monto Operación
            fMatOperaciones(4, i) = rsDet!cEstado 'Estado Operación
            fMatOperaciones(5, i) = rsDet!lsDetalle 'Detalle Operación
            fMatOperaciones(6, i) = rsDet!lsGlosa 'Glosa
            rsDet.MoveNext
        Next
        SetFlexOperaciones 'Mostramos en Pantalla
        cargarDatos = True
    Else
        cargarDatos = False
    End If
    Exit Function
ErrCargarDatos:
    cargarDatos = False
End Function
Private Sub LimpiarControles()
    txtFechaMov.Text = Format(gdFecSis, "dd/mm/yyyy")
    TxtTitPersCod.Text = ""
    TxtTitPersNombre.Text = ""
    cboMoneda.ListIndex = -1
    txtVoucherMonto.Text = "0.00"
    txtCtaIFCod.Text = ""
    txtCtaIFBancoNombre.Text = ""
    txtCtaIFDesc.Text = ""
    txtVoucherFecha.Text = Format(gdFecSis, "dd/mm/yyyy")
    txtVoucherNro.Text = ""
    chkConfirmar.value = 0
    'Limpia Operaciones ***
    Call LimpiaFlex(feOperaciones)
    SetMatrizOperaciones
    '**********************
    txtVoucherMonto.Locked = False
    txtVoucherNro.Locked = False
    cboMoneda.Locked = False
    txtOpeTotal.Text = "0.00"
    fnFilaActual = -1
    cmdCancelar_Click
End Sub
Private Sub txtFechaMov_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If TxtTitPersCod.Visible And TxtTitPersCod.Enabled Then
            TxtTitPersCod.SetFocus
        End If
    End If
End Sub
Private Sub TxtTitPersCod_EmiteDatos()
    TxtTitPersNombre.Text = ""
    If TxtTitPersCod.Text = gsCodPersUser Then
        MsgBox "No se puede registrar un Voucher de si mismo", vbInformation, "Aviso"
        TxtTitPersCod.Text = ""
        Exit Sub
    End If
    If TxtTitPersCod.psDescripcion <> "" Then
        TxtTitPersNombre.Text = TxtTitPersCod.psDescripcion
        If cboMoneda.Visible And cboMoneda.Enabled Then
            cboMoneda.SetFocus
        End If
    End If
End Sub
Private Sub cboMoneda_Click()
    Dim oOpe As New clases.DOperacion
    Dim lnMoneda As Integer
    Dim lsColor As String
    
    lsColor = &H80000005
    txtCtaIFCod.Text = ""
    txtCtaIFBancoNombre.Text = ""
    txtCtaIFDesc.Text = ""
    If cboMoneda.ListIndex <> -1 Then
        lnMoneda = CInt(Right(cboMoneda.Text, 2))
        txtCtaIFCod.psRaiz = "Cuentas de Instituciones Financieras"
        txtCtaIFCod.rs = oOpe.listarCuentasEntidadesFinacieras("_1_[12]" & CStr(lnMoneda) & "%", CStr(lnMoneda))
        If lnMoneda = 1 Then
            lsColor = &H80000005
        Else
            lsColor = &HC0FFC0
        End If
    End If
    txtCtaIFCod.BackColor = lsColor
    txtCtaIFBancoNombre.BackColor = lsColor
    txtCtaIFDesc.BackColor = lsColor
    txtVoucherMonto.BackColor = lsColor
    Set oOpe = Nothing
End Sub
Private Sub cboMoneda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtVoucherMonto.Visible And txtVoucherMonto.Enabled Then
            txtVoucherMonto.SetFocus
        End If
    End If
End Sub

'RIRO20140430 ERS017 ***
Private Sub txtVoucherMonto_GotFocus()
    txtVoucherMonto.SelStart = 0
    txtVoucherMonto.SelLength = Len(txtVoucherMonto.Text)
    txtVoucherMonto.SetFocus
End Sub
'END RIRO **************

Private Sub txtVoucherMonto_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtVoucherMonto, KeyAscii, 15)
    If KeyAscii <> 13 Then Exit Sub
    If KeyAscii = 13 Then
        If txtCtaIFCod.Visible And txtCtaIFCod.Enabled Then
            txtCtaIFCod.SetFocus
        End If
    End If
End Sub
Private Sub txtVoucherMonto_LostFocus()
    txtVoucherMonto.Text = Format(txtVoucherMonto.Text, gsFormatoNumeroView)
    If Not IsNumeric(txtVoucherMonto.Text) Then
        MsgBox "Ingrese un monto válido", vbInformation, "Aviso"
        If txtVoucherMonto.Visible And txtVoucherMonto.Enabled Then
            txtVoucherMonto.SetFocus
        End If
    End If
End Sub
Private Sub txtCtaIFCod_GotFocus()
    If cboMoneda.ListIndex = -1 Then
        MsgBox "Ud. debe de seleccionar primero la moneda", vbInformation, "Aviso"
        If cboMoneda.Visible And cboMoneda.Enabled Then
            cboMoneda.SetFocus
        End If
        Exit Sub
    End If
End Sub
Private Sub txtCtaIFCod_EmiteDatos()
    Dim oNCajaCtaIF As New clases.NCajaCtaIF
    Dim oDOperacion As New clases.DOperacion
    
    txtCtaIFBancoNombre.Text = ""
    txtCtaIFDesc.Text = ""
    If txtCtaIFCod.Text <> "" Then
        txtCtaIFBancoNombre.Text = oNCajaCtaIF.NombreIF(Mid(txtCtaIFCod.Text, 4, 13))
        txtCtaIFDesc.Text = oDOperacion.recuperaTipoCuentaEntidadFinaciera(Mid(txtCtaIFCod.Text, 18, 10)) & " " & txtCtaIFCod.psDescripcion
        If txtVoucherFecha.Visible And txtVoucherFecha.Enabled Then
            txtVoucherFecha.SetFocus
        End If
    End If
    Set oNCajaCtaIF = Nothing
    Set oDOperacion = Nothing
End Sub
Private Sub txtVoucherFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim lsmensaje As String
        lsmensaje = ValidaFecha(txtVoucherFecha.Text)
    
        If Trim(lsmensaje) <> "" Then
            MsgBox lsmensaje, vbInformation, "Aviso"
            If txtVoucherFecha.Visible And txtVoucherFecha.Enabled Then
                txtVoucherFecha.SetFocus
            End If
            Exit Sub
        ElseIf Trim(lsmensaje) = "" Then
            If txtVoucherNro.Visible And txtVoucherNro.Enabled Then
                txtVoucherNro.SetFocus
            End If
        End If
    End If
End Sub
Private Sub txtVoucherNro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If chkConfirmar.Visible And chkConfirmar.Enabled Then
            chkConfirmar.SetFocus
        End If
    End If
End Sub
Private Sub txtVoucherNro_LostFocus()
    txtVoucherNro.Text = Trim(txtVoucherNro.Text)
End Sub
Private Sub chkConfirmar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cmdAgregar.Visible And cmdAgregar.Enabled Then
            cmdAgregar.SetFocus
        End If
    End If
End Sub
Private Sub feOperaciones_OnClickTxtBuscar(psCodigo As String, psDescripcion As String)
    Dim row As Integer
    Dim lnOperacion As Integer

    row = feOperaciones.row
    If Trim(feOperaciones.TextMatrix(row, 2)) = "" Then
        MsgBox "Ud.debe seleccionar primero la Operación a realizar", vbInformation, "Aviso"
        feOperaciones.TopRow = row
        feOperaciones.row = row
        feOperaciones.Col = 2
        Exit Sub
    End If
    lnOperacion = CInt(Trim(Right(feOperaciones.TextMatrix(row, 2), 8)))
    Select Case lnOperacion
        Case 1, 3, 5 'Aperturas
            Dim frm As New frmCapRegVouDepDetApert
            feOperaciones.TextMatrix(row, 5) = frm.Inicio(feOperaciones.TextMatrix(row, 5))
            psCodigo = feOperaciones.TextMatrix(row, 5)
            psDescripcion = feOperaciones.TextMatrix(row, 5)
            Set frm = Nothing
        Case 2, 4, 6 'Depósitos y Aumento de Capital
            Dim oPersona As New COMDPersona.UCOMPersona
            Dim frmBP As New frmBuscaPersona
            
            Dim lnProducto As Producto
            Dim lsTpoPrograma As String
            Dim lsNroCuenta As String
            
            Set oPersona = frmBP.Inicio
            lsNroCuenta = feOperaciones.TextMatrix(row, 5)
            If Not oPersona Is Nothing Then
                If oPersona.sPersCod <> "" Then
                    If oPersona.sPersCod <> gsCodPersUser Then
                        Dim clsCap As New COMNCaptaGenerales.NCOMCaptaGenerales
                        Dim clsCuenta As New UCapCuenta
                        Dim rsPers As New ADODB.Recordset
                        Dim frmMntCap As New frmCapMantenimientoCtas
                                        
                        If lnOperacion = 2 Then
                            lnProducto = gCapAhorros
                            lsTpoPrograma = fsTpoProgramaAho
                        ElseIf lnOperacion = 4 Then
                            lnProducto = gCapPlazoFijo
                            lsTpoPrograma = fsTpoProgramaDPF
                        ElseIf lnOperacion = 6 Then
                            lnProducto = gCapCTS
                            lsTpoPrograma = fsTpoProgramaCTS
                        End If
                        Set rsPers = clsCap.GetCuentasPersona(oPersona.sPersCod, lnProducto, True, , CInt(Trim(Right(cboMoneda.Text, 2))), , , lsTpoPrograma)
                        If Not RSVacio(rsPers) Then
                            Do While Not rsPers.EOF
                                If lnProducto = gCapAhorros And rsPers("nPrdPersRelac") = 10 Then 'And Not EsHaberes(rsPers("cCtaCod")) Then
                                    frmMntCap.lstCuentas.AddItem rsPers("cCtaCod") & space(2) & rsPers("cRelacion") & space(2) & Trim(rsPers("cEstado"))
                                End If
                                If lnProducto = gCapPlazoFijo And rsPers("nPrdPersRelac") = 10 Then
                                    frmMntCap.lstCuentas.AddItem rsPers("cCtaCod") & space(2) & rsPers("cRelacion") & space(2) & Trim(rsPers("cEstado"))
                                End If
                                If lnProducto = gCapCTS And rsPers("nPrdPersRelac") = 10 Then
                                    frmMntCap.lstCuentas.AddItem rsPers("cCtaCod") & space(2) & rsPers("cRelacion") & space(2) & Trim(rsPers("cEstado"))
                                End If
                                rsPers.MoveNext
                            Loop
                            Set clsCuenta = frmMntCap.Inicia
                            If Not clsCuenta Is Nothing Then
                                If clsCuenta.sCtaCod <> "" Then
                                    lsNroCuenta = clsCuenta.sCtaCod
                                Else
                                    MsgBox "No se ha seleccionada ninguna cuenta", vbInformation, "Aviso"
                                End If
                            End If
                        Else
                            MsgBox "Persona no posee ninguna cuenta", vbInformation, "Aviso"
                        End If
                        
                        Set clsCap = Nothing
                        Set clsCuenta = Nothing
                        Set rsPers = Nothing
                        Set frmMntCap = Nothing
                    Else
                        MsgBox "No se puede registrar un Voucher de si mismo", vbInformation, "Aviso"
                    End If
                End If
            End If

            feOperaciones.TextMatrix(row, 5) = lsNroCuenta
            psCodigo = lsNroCuenta
            psDescripcion = lsNroCuenta
            Set oPersona = Nothing
            Set frmBP = Nothing
        Case 7 'Lote
            Dim frmLote As New frmCapRegVouDepDetLote
            feOperaciones.TextMatrix(row, 5) = frmLote.Inicio(Val(feOperaciones.TextMatrix(row, 5)))
            psCodigo = feOperaciones.TextMatrix(row, 5)
            psDescripcion = feOperaciones.TextMatrix(row, 5)
            Set frmLote = Nothing
        'RIRO20140407 ERS017 **************
        Case 7, 8, 9, 14, 15, 16, 17, 22, 23, 24 'CTI7 OPEv2
            Dim sNroOpe As String
            If lnOperacion <> 22 Then
                sNroOpe = InputBox("Ingrese el Nro de Operaciones", "Operaciones")
                If IsNumeric(sNroOpe) Then
                    If CDbl(sNroOpe) <= 100000 Then
                        psCodigo = Round(CDbl(sNroOpe), 0)
                        psDescripcion = Round(CDbl(sNroOpe), 0)
                    Else
                        psCodigo = "0"
                        psDescripcion = "0"
                    End If
                Else
                    psCodigo = "0"
                    psDescripcion = "0"
                End If
            Else
                sNroOpe = 1
                psCodigo = "1"
                psDescripcion = "1"
            End If
            
        Case 10, 11, 13, 18, 19, 20, 21 'WIOR 20160413 AGREGO CODIGO 13 LIQ. CREDITO SEG DES
            
            Dim loPers As COMDPersona.UCOMPersona
            Dim lsEstados As String
            Dim loPersCredito  As COMDColocRec.DCOMColRecCredito
            Dim lrCreditos As ADODB.Recordset
            Dim loCuentas As COMDPersona.UCOMProdPersona
            
            Set loPers = New COMDPersona.UCOMPersona
            Set loPers = frmBuscaPersona.Inicio
            If loPers Is Nothing Then
                Exit Sub
            End If
            
            ' Normal
            If lnOperacion = 10 Or lnOperacion = 13 Then 'WIOR 20160413 AGREGO lnOperacion = 12 PARA LA LIQ DE CRED CON SEG DES
                lsEstados = gColocEstVigMor & "," & gColocEstVigVenc & "," & _
                            gColocEstVigNorm & "," & gColocEstRefMor & "," & _
                            gColocEstRefVenc & "," & gColocEstRefNorm & "," & _
                            gColocEstTransferido 'FRHU 20150520 ERS022-2015: Se agrego gColocEstTransferido
            ' Judicial
            ElseIf lnOperacion = 11 Then
                lsEstados = gColocEstRecVigJud & "," & gColocEstRecVigCast
            
            ElseIf lnOperacion = 18 Or lnOperacion = 19 Or lnOperacion = 20 Or lnOperacion = 21 Then
                lsEstados = gColPEstVenci & "," & gColPEstRegis & "," & _
                            gColPEstDesem & "," & gColPEstDifer & "," & _
                            gColPEstRenov & "," & gColPEstAdjud
            
            End If
            
            If Trim(loPers.sPersCod) <> "" Then
                Set loPersCredito = New COMDColocRec.DCOMColRecCredito
                Set lrCreditos = loPersCredito.dObtieneCreditosDePersona(loPers.sPersCod, lsEstados)
                Set loPersCredito = Nothing
            End If
            
            Set loCuentas = New COMDPersona.UCOMProdPersona
            Set loCuentas = frmProdPersona.Inicio(loPers.sPersNombre, lrCreditos)
            If loCuentas.sCtaCod <> "" Then
                feOperaciones.TextMatrix(row, 5) = loCuentas.sCtaCod
                psCodigo = loCuentas.sCtaCod
                psDescripcion = loCuentas.sCtaCod
            End If
            Set loCuentas = Nothing

            
        'END RIRO *************************
        Case Else
            MsgBox "Esta Operación no esta configurado para este proceso," & Chr(13) & "Comuniquese con el Dpto. de Tecnología de la Información", vbInformation, "Aviso"
            Exit Sub
    End Select
End Sub
Private Sub feOperaciones_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    Dim Editar() As String
    Dim lnTotal As Currency
    Dim lnOperacion As String 'RIRO20140610 ERS017
    Dim oNCred As COMNCredito.NCOMCredito 'RIRO20140610 ERS017
    Dim lnMonto As Double, lnITF As Double
    
    lnOperacion = Trim(Right(feOperaciones.TextMatrix(feOperaciones.row, 2), 8))
    lnOperacion = IIf(IsNumeric(lnOperacion), lnOperacion, 0)
    
    Editar = Split(Me.feOperaciones.ColumnasAEditar, "-")
    If Editar(pnCol) = "X" Then
        MsgBox "Esta celda no es editable", vbInformation, "Aviso"
        Cancel = False
        Exit Sub
    End If
    If pnCol = 1 Then
        MsgBox "Esta celda no es editable", vbInformation, "Aviso"
        Cancel = False
        Exit Sub
    End If
    If pnCol = 3 Then
        If Not IsNumeric(feOperaciones.TextMatrix(pnRow, pnCol)) Then
            MsgBox "Ingrese un Monto válido", vbInformation, "Aviso"
            Cancel = False
            Exit Sub
        Else
            If Val(feOperaciones.TextMatrix(pnRow, pnCol)) <= 0 Then
                MsgBox "Ingrese un Monto mayor a cero", vbInformation, "Aviso"
                Cancel = False
                Exit Sub
            End If
            lnTotal = feOperaciones.SumaRow(pnCol)
            If lnTotal > CCur(txtVoucherMonto.Text) Then
                MsgBox "El Monto del Voucher es diferente a la Sumatoria del Detalle de Operaciones, verifique", vbInformation, "Aviso"
                Cancel = False
                Exit Sub
            End If
            'RIRO20140610 ERS017 ****
            If lnOperacion = 10 Or lnOperacion = 11 Then
                Set oNCred = New COMNCredito.NCOMCredito
                lnMonto = CCur((feOperaciones.TextMatrix(pnRow, pnCol)))
                lnITF = oNCred.DameMontoITF(lnMonto)
                Set oNCred = Nothing
                If lnITF > 0 Then
                    If MsgBox("El monto de Pago está afecto a un ITF de " & Format(lnITF, gsFormatoNumeroView) & Chr(13) & "¿Desea agregar el ITF en el Monto de Pago?", vbYesNo + vbInformation, "Aviso") = vbYes Then
                        feOperaciones.TextMatrix(pnRow, 3) = Format(lnMonto + lnITF, gsFormatoNumeroView)
                    End If
                End If
            End If
            'END RIRO ***************
            Dim nRow As Integer

            If Trim(Right(feOperaciones.TextMatrix(feOperaciones.row, 2), 10)) = 22 Then
                feOperaciones.TextMatrix(feOperaciones.row, 5) = 1
    
            End If
        End If
    End If
End Sub
Private Sub feOperaciones_OnCellChange(pnRow As Long, pnCol As Long)
    Dim lnTotal As Currency
    Dim i As Integer
    Dim sNroOperaciones As String 'RIRO20140407 ERS017
    Dim lnOperacion As String 'RIRO20140407 ERS017
    
    On Error GoTo ErrOnCellChangue
    If pnCol = 3 Then
        For i = 1 To feOperaciones.Rows - 1
            lnTotal = lnTotal + CCur(feOperaciones.TextMatrix(i, 3))
        Next
        txtOpeTotal.Text = Format(lnTotal, gsFormatoNumeroView)
    End If
    If pnCol = 7 Then
        feOperaciones.TextMatrix(pnRow, pnCol) = UCase(Replace(Trim(feOperaciones.TextMatrix(pnRow, pnCol)), "'", ""))
    End If
    
    'RIRO20140407 ERS017 *****************************************************************************
    'lnOperacion = Trim(Right(feOperaciones.TextMatrix(feOperaciones.row, 2), 8))
    'lnOperacion = IIf(IsNumeric(lnOperacion), lnOperacion, 0)
    'If pnCol = 5 And (lnOperacion = 8 Or lnOperacion = 9) Then
    '    sNroOperaciones = feOperaciones.TextMatrix(pnRow, pnCol)
    '    feOperaciones.TextMatrix(pnRow, pnCol) = IIf(IsNumeric(sNroOperaciones), sNroOperaciones, 0)
    'End If
    'If lnOperacion = 8 Or lnOperacion = 9 Then
    '    feOperaciones.EncabezadosAlineacion = "C-L-L-R-C-R-C-L-C"
    'Else
    '    feOperaciones.EncabezadosAlineacion = "C-L-L-R-C-L-C-L-C"
    'End If
    'END RIRO ****************************************************************************************
    
    'Guardamos el valor del combo en un temporal y si en caso cambio de operacion seteamos el detalle
    If feOperaciones.TextMatrix(pnRow, 6) <> feOperaciones.TextMatrix(pnRow, 2) Then
       feOperaciones.TextMatrix(pnRow, 5) = ""
       feOperaciones.TextMatrix(pnRow, 6) = feOperaciones.TextMatrix(pnRow, 2)
    End If
    Exit Sub
ErrOnCellChangue:
    MsgBox err.Description, vbCritical, "Aviso"
End Sub
Private Sub feOperaciones_RowColChange()
    Dim rs As New ADODB.Recordset
    Dim oConst As New COMDConstantes.DCOMConstantes
    Dim lnOperacion As String 'RIRO20140407 ERS017
    
    If feOperaciones.lbEditarFlex Then
        If feOperaciones.Col = 2 Then
            Set rs = oConst.RecuperaConstantes(10001)
            feOperaciones.CargaCombo rs
        End If
        'RIRO20140407 ERS017 *********************************************************
            'lnOperacion = Trim(Right(feOperaciones.TextMatrix(feOperaciones.row, 2), 8))
            'lnOperacion = IIf(IsNumeric(lnOperacion), lnOperacion, 0)
            'If feOperaciones.Col = 5 And (lnOperacion = 8 Or lnOperacion = 9) Then
            '    feOperaciones.ToolTipText = "Ingresar el numero de operaciones en lote"
            '    feOperaciones.ListaControles = "0-0-3-0-0-0-0-0-0"
            'Else
            '    feOperaciones.ToolTipText = ""
            '    feOperaciones.ListaControles = "0-0-3-0-0-1-0-0-0"
            'End If
        'END RIRO ********************************************************************
        feOperaciones.row = fnFilaActual 'Mantiene la posición de la fila activa
    End If
    Set rs = Nothing
    Set oConst = Nothing
End Sub
Private Sub cmdAgregar_Click()
    If Not validaDatosVoucher Then Exit Sub
    If feOperaciones.TextMatrix(1, 0) <> "" Then
        If Not validaDatosOperacion Then
            Exit Sub
        End If
    End If
    feOperaciones.lbEditarFlex = True
    feOperaciones.AdicionaFila
    fnFilaActual = feOperaciones.row
    
    InhabilitaDatosVoucher
    Call EstableceVoucherNroEnFlex(feOperaciones.row)
    feOperaciones.TextMatrix(feOperaciones.row, 3) = "0.00"
    feOperaciones.TextMatrix(feOperaciones.row, 4) = "PENDIENTE"
    If feOperaciones.Visible And feOperaciones.Enabled Then
        feOperaciones.SetFocus
    End If
    SendKeys "{Enter}"
    'Bloqueamos controles
    cmdAgregar.Visible = False
    cmdEditar.Visible = False
    cmdQuitar.Visible = False
    cmdAceptar.Visible = True
    cmdCancelar.Visible = True
End Sub
Private Sub cmdEditar_Click()
    If feOperaciones.TextMatrix(1, 0) = "" Then Exit Sub
    If feOperaciones.TextMatrix(feOperaciones.row, 4) = "REALIZADO" Then
        MsgBox "Este registro no podrá ser modificado porque ya realizaron operaciones con este Voucher", vbCritical, "Aviso"
        Exit Sub
    End If
    fnFilaActual = feOperaciones.row
    feOperaciones.Col = 3 'Para refrescar la columna 2 (combo)
    feOperaciones.lbEditarFlex = True
    If feOperaciones.Visible And feOperaciones.Enabled Then
        feOperaciones.SetFocus
    End If
    'Bloqueamos controles
    cmdAgregar.Visible = False
    cmdEditar.Visible = False
    cmdQuitar.Visible = False
    cmdAceptar.Visible = True
    cmdCancelar.Visible = True
End Sub
Private Sub cmdQuitar_Click()
    Dim i As Integer
    If feOperaciones.TextMatrix(1, 0) <> "" Then
        If MsgBox("Se va a quitar el último registro ingresado" & Chr(13) & "¿Desea continuar?", vbYesNo + vbInformation, "Aviso") = vbNo Then
            Exit Sub
        End If
        feOperaciones.EliminaFila feOperaciones.Rows - 1
        SetMatrizOperaciones 'Guardamos el Flex en la Matriz
        InhabilitaDatosVoucher
    End If
End Sub
Private Sub cmdAceptar_Click()
    Dim i As Integer
    If feOperaciones.TextMatrix(1, 0) <> "" Then
        If Not validaDatosOperacion Then
            Exit Sub
        End If
    End If
    SetMatrizOperaciones 'Guardamos el Flex en la Matriz
    InhabilitaDatosVoucher
    feOperaciones.lbEditarFlex = False
    fnFilaActual = -1
    'Desbloqueamos controles
    cmdAgregar.Visible = True
    cmdEditar.Visible = True
    cmdQuitar.Visible = True
    cmdAceptar.Visible = False
    cmdCancelar.Visible = False
    If Not fbNuevo Then
        cmdAgregar.Visible = False
        cmdQuitar.Visible = False
    End If
End Sub
Private Sub cmdCancelar_Click()
    SetFlexOperaciones
    InhabilitaDatosVoucher
    feOperaciones.lbEditarFlex = False
    fnFilaActual = -1
    'Desbloqueamos controles
    cmdAgregar.Visible = True
    cmdEditar.Visible = True
    cmdQuitar.Visible = True
    cmdAceptar.Visible = False
    cmdCancelar.Visible = False
    If Not fbNuevo Then
        cmdAgregar.Visible = False
        cmdQuitar.Visible = False
    End If
End Sub
Private Sub cmdGrabar_Click()
    Dim oNCOMCaptaGenerales As New NCOMCaptaGenerales
    Dim oNCOMContFunciones As New COMNContabilidad.NCOMContFunciones
    Dim oNContFunciones As New clases.NContFunciones
    Dim lsMovNro As String, lsBcoCtaCont As String, lsBcoSubCtaCont As String
    Dim lbExito As Boolean
    Dim MatOperacion As Variant
    Dim i As Integer
    Dim lnMontoDet As Double
    Dim lnMoneda As Integer
    Dim lsIFTpo As String, lsPersCod As String, lsCtaIFCod As String, lsVoucherNro As String
    Dim ldFechaReg As Date
    
    If Not ValidaDatosGrabar Then Exit Sub
    
    lnMoneda = CInt(Trim(Right(cboMoneda.Text, 2)))
    lsIFTpo = Left(txtCtaIFCod.Text, 2)
    lsPersCod = Mid(txtCtaIFCod.Text, 4, 13)
    lsCtaIFCod = Mid(txtCtaIFCod.Text, 18, 10)
    lsVoucherNro = Trim(txtVoucherNro.Text)
    ldFechaReg = CDate(txtFechaMov.Text)
    
    lsBcoCtaCont = "11" & lnMoneda & IIf(Mid(txtCtaIFCod.Text, 4, 13) = "1090100822183", "2", "3") & "01"
    lsBcoSubCtaCont = oNContFunciones.GetFiltroObjetos(1, lsBcoCtaCont, txtCtaIFCod.Text, False)
    
    If lsBcoSubCtaCont = "" Then
        MsgBox "Esta cuenta contable " & lsBcoCtaCont & " no esta registrado en CtaIFFiltro, comunicarse con TI", vbInformation, "Aviso"
        Exit Sub
    End If
    If oNContFunciones.verificarUltimoNivelCta(lsBcoCtaCont & lsBcoSubCtaCont) = False Then
       MsgBox "La Cuenta Contable " & lsBcoCtaCont + lsBcoSubCtaCont & " no es de Ultimo Nivel, comunicarse con Contabilidad", vbInformation, "Aviso"
       Exit Sub
    End If
    'Valida Total detalle cuadre con el monto voucher
    For i = 1 To UBound(fMatOperaciones, 2)
        lnMontoDet = lnMontoDet + fMatOperaciones(3, i)
    Next
    If CCur(lnMontoDet) <> CCur(txtVoucherMonto.Text) Then 'RIRO INC1405140004 - Se reemplazó CDbl por CCur
        MsgBox "El Monto del Voucher es diferente a la Sumatoria del Detalle de Operaciones, verifique", vbInformation, "Aviso"
        Exit Sub
    End If
    
    If fbNuevo Then
        If oNCOMCaptaGenerales.ExisteVoucherDeposito(lsIFTpo, lsPersCod, lsCtaIFCod, lsVoucherNro) Then
            MsgBox "Anteriormente ya se ha registrado un Voucher con este mismo Número, verifique", vbInformation, "Aviso"
            Exit Sub
        End If
    End If
    
    If MsgBox("¿Esta seguro de Guardar los datos del Voucher?", vbYesNo + vbInformation, "Aviso") = vbNo Then
        Exit Sub
    End If
    If fbNuevo Then
        If CDate(txtVoucherFecha.Text) < ldFechaReg Then
            frmCapRegVouDepPen.iniciarListado CDate(txtVoucherFecha.Text), Left(txtCtaIFCod.Text, 2), Mid(txtCtaIFCod.Text, 4, 13), Mid(txtCtaIFCod.Text, 18, 10), CCur(txtVoucherMonto.Text), CStr(lnMoneda), Me
            If fnMovNroPen = 0 Then
                MsgBox "Debe relacionar con una Pendiente el Voucher.", vbInformation, "Aviso"
                Exit Sub
            End If
        End If
        lbExito = oNCOMCaptaGenerales.InsertarVoucherDeposito(ldFechaReg, Right(gsCodAge, 2), gsCodUser, gsOpeCod, _
                                                                TxtTitPersCod.Text, lsIFTpo, lsPersCod, lsCtaIFCod, _
                                                                lsBcoCtaCont, lsBcoSubCtaCont, _
                                                                CStr(lnMoneda), lsVoucherNro, CCur(txtVoucherMonto.Text), CDate(txtVoucherFecha.Text), _
                                                                IIf(chkConfirmar.value = 1, True, False), fnMovNroPen, fMatOperaciones, Nothing, 0)
    Else
        lbExito = oNCOMCaptaGenerales.EditarVoucherDeposito(fnId, ldFechaReg, Right(gsCodAge, 2), gsCodUser, gsOpeCod, _
                                                                TxtTitPersCod.Text, lsIFTpo, lsPersCod, lsCtaIFCod, _
                                                                lsBcoCtaCont, lsBcoSubCtaCont, _
                                                                CStr(lnMoneda), lsVoucherNro, CCur(txtVoucherMonto.Text), CDate(txtVoucherFecha.Text), _
                                                                IIf(chkConfirmar.value = 1, True, False), fnMovNroPen, fMatOperaciones)
    End If
    If lbExito Then
        MsgBox "Se realizó correctamente la operación", vbInformation, "Aviso"
        fbNuevo = False 'Seteo la variable para poder salir del formulario una vez grabada la Operacion en el Form_Unload()
        Unload Me
    Else
        MsgBox "Hubo un error al registrar el Voucher, " & Chr(13) & "si el error persiste comuniquese con el Dpto. de TI", vbCritical, "Aviso"
    End If
    fnMovNroPen = 0
    
    Set oNCOMCaptaGenerales = Nothing
    Set oNCOMContFunciones = Nothing
    Set oNCOMCaptaGenerales = Nothing
End Sub
Private Sub cmdLimpiar_Click()
    
    LimpiarControles
End Sub
Private Sub cmdsalir_Click()
    Unload Me
End Sub
Private Sub EstableceVoucherNroEnFlex(ByVal row As Long)
    feOperaciones.TextMatrix(row, 1) = Trim(txtVoucherNro.Text) & "-" & feOperaciones.TextMatrix(row, 0)
End Sub
Private Sub CargaMoneda()
    cboMoneda.Clear
    cboMoneda.AddItem ("SOLES" & space(150) & "1")
    cboMoneda.AddItem ("DOLARES" & space(150) & "2")
End Sub
Private Function ValidaDatosGrabar() As Boolean
    Dim lsFecha As String
    ValidaDatosGrabar = True
    lsFecha = ValidaFecha(txtFechaMov.Text)
    If Len(lsFecha) > 0 Then
        MsgBox lsFecha, vbInformation, "Aviso"
        ValidaDatosGrabar = False
        If txtFechaMov.Visible And txtFechaMov.Enabled Then
            txtFechaMov.SetFocus
        End If
        Exit Function
    End If
    If Not validaDatosVoucher Then
        ValidaDatosGrabar = False
        Exit Function
    End If
    If Not validaDatosOperacion Then
        ValidaDatosGrabar = False
        Exit Function
    End If
    'Valide Acepte o Cancele los cambios en el Detalle de Operacion
    If cmdAceptar.Visible Or cmdCancelar.Visible Then
        MsgBox "Ud. debe aceptar o cancelar los cambios en el Detalle de Operaciones", vbInformation, "Aviso"
        ValidaDatosGrabar = False
        If cmdAceptar.Visible And cmdAceptar.Enabled Then
            cmdAceptar.SetFocus
        ElseIf cmdCancelar.Visible And cmdCancelar.Enabled Then
            cmdCancelar.SetFocus
        End If
        Exit Function
    End If
End Function
Private Function validaDatosVoucher() As Boolean
    Dim lsFecha As String
    validaDatosVoucher = True
    If Len(Trim(TxtTitPersCod.Text)) <> 13 Then
        MsgBox "Ud. debe de especificar el Titular del Voucher", vbInformation, "Aviso"
        validaDatosVoucher = False
        If TxtTitPersCod.Visible And TxtTitPersCod.Enabled Then
            TxtTitPersCod.SetFocus
        End If
        Exit Function
    End If
    If cboMoneda.ListIndex = -1 Then
        MsgBox "Ud. debe de seleccionar la Moneda del Depósito", vbInformation, "Aviso"
        validaDatosVoucher = False
        If cboMoneda.Visible And cboMoneda.Enabled Then
            cboMoneda.SetFocus
        End If
        Exit Function
    End If
    If Not IsNumeric(txtVoucherMonto.Text) Then
        MsgBox "Ud. debe de especificar el Monto del Depósito", vbInformation, "Aviso"
        validaDatosVoucher = False
        If txtVoucherMonto.Visible And txtVoucherMonto.Enabled Then
            txtVoucherMonto.SetFocus
        End If
        Exit Function
    Else
        If CCur(txtVoucherMonto.Text) <= 0 Then
            MsgBox "Ud. debe de especificar el Monto del Depósito", vbInformation, "Aviso"
            validaDatosVoucher = False
            If txtVoucherMonto.Visible And txtVoucherMonto.Enabled Then
                txtVoucherMonto.SetFocus
            End If
            Exit Function
        End If
    End If
    If Len(Trim(txtCtaIFCod.Text)) = 0 Then
        MsgBox "Ud. debe de especificar el Destino del Depósito", vbInformation, "Aviso"
        validaDatosVoucher = False
        If txtCtaIFCod.Visible And txtCtaIFCod.Enabled Then
            txtCtaIFCod.SetFocus
        End If
        Exit Function
    End If
    lsFecha = ValidaFecha(txtVoucherFecha.Text)
    If Len(lsFecha) > 0 Then
        MsgBox lsFecha, vbInformation, "Aviso"
        validaDatosVoucher = False
        If txtVoucherFecha.Visible And txtVoucherFecha.Enabled Then
            txtVoucherFecha.SetFocus
        End If
        Exit Function
    Else
        If CDate(txtVoucherFecha.Text) > gdFecSis Then
            MsgBox "La Fecha del Voucher no puede ser posterior a la Fecha del Sistema", vbInformation, "Aviso"
            validaDatosVoucher = False
            If txtVoucherFecha.Visible And txtVoucherFecha.Enabled Then
                txtVoucherFecha.SetFocus
            End If
            Exit Function
        End If
    End If
    If Len(Trim(txtVoucherNro.Text)) = 0 Then
        MsgBox "Ud. debe de especificar el Número del Voucher de Depósito", vbInformation, "Aviso"
        validaDatosVoucher = False
        If txtVoucherNro.Visible And txtVoucherNro.Enabled Then
            txtVoucherNro.SetFocus
        End If
        Exit Function
    End If
End Function
Private Function validaDatosOperacion() As Boolean
    Dim i As Integer, J As Integer
    Dim sMensaje As String 'RIRO20140610 ERS017
    Dim nOperacion As Integer 'RIRO20140610 ERS017
    validaDatosOperacion = True
    'Valida No este vacio el Flex
    If feOperaciones.TextMatrix(1, 0) = "" Then
        MsgBox "Ud. debe de especificar las Operaciones", vbInformation, "Aviso"
        If feOperaciones.Visible And feOperaciones.Enabled Then
            feOperaciones.SetFocus
        End If
        validaDatosOperacion = False
        Exit Function
    End If
    'Valida No este vacio los datos
    For i = 1 To feOperaciones.Rows - 1
        For J = 1 To feOperaciones.Cols - 1
            If feOperaciones.ColWidth(J) <> 0 Then
                If Len(Trim(feOperaciones.TextMatrix(i, J))) = 0 Then
                    MsgBox "El campo " & UCase(feOperaciones.TextMatrix(0, J)) & " está vacio", vbInformation, "Aviso"
                    If feOperaciones.Visible And feOperaciones.Enabled Then
                        feOperaciones.SetFocus
                    End If
                    feOperaciones.TopRow = i
                    feOperaciones.row = i
                    feOperaciones.Col = J
                    validaDatosOperacion = False
                    Exit Function
                End If
                If J = 3 Then
                    If CCur(feOperaciones.TextMatrix(i, J)) <= 0 Then
                        MsgBox "Ingrese un Monto mayor a cero", vbInformation, "Aviso"
                        If feOperaciones.Visible And feOperaciones.Enabled Then
                            feOperaciones.SetFocus
                        End If
                        feOperaciones.TopRow = i
                        feOperaciones.row = i
                        feOperaciones.Col = J
                        validaDatosOperacion = False
                        Exit Function
                    End If
                End If
                Dim valor2 As Integer
                valor2 = Trim(Right(feOperaciones.TextMatrix(i, 2), 10))
                If J = 5 And (valor2 = 7 Or valor2 = 8 Or valor2 = 9 Or valor2 = 14 Or valor2 = 15 Or valor2 = 16 Or valor2 = 17 Or valor2 = 22 Or valor2 = 23 Or valor2 = 24) Then
                    If CCur(feOperaciones.TextMatrix(i, J)) <= 0 Then
                        MsgBox "Ingrese un Monto mayor a cero", vbInformation, "Aviso"
                        If feOperaciones.Visible And feOperaciones.Enabled Then
                            feOperaciones.SetFocus
                        End If
                        feOperaciones.TopRow = i
                        feOperaciones.row = i
                        feOperaciones.Col = J
                        validaDatosOperacion = False
                        Exit Function
                    End If
                End If
            End If
        Next
    Next
    'Valida sea mayor a cero el campo detalle cuando el Tipo de Operación es CTS-Depósito en Lote
    For i = 1 To feOperaciones.Rows - 1
        nOperacion = CInt(Trim(Right(feOperaciones.TextMatrix(i, 2), 2)))
        'If CInt(Trim(Right(feOperaciones.TextMatrix(i, 2), 2))) = 7 Then
        If nOperacion = 7 Or nOperacion = 8 Or nOperacion = 9 Then
            If Val(Trim(feOperaciones.TextMatrix(i, 5))) <= 0 Then
                If nOperacion = 7 Then
                    sMensaje = "Para las Operaciones de CTS-Depósito en Lote el Detalle debe ser mayor a cero"
                ElseIf nOperacion = 8 Then
                    sMensaje = "Para las Operaciones de Ahorro-Apertura en Lote el Detalle debe ser mayor a cero"
                ElseIf nOperacion = 9 Then
                    sMensaje = "Para las Operaciones de Ahorro-Depósito de Haberes en Lote el Detalle debe ser mayor a cero"
                End If
                'MsgBox "Para las Operaciones de CTS-Depósito en Lote el Detalle debe ser mayor a cero", vbInformation, "Aviso"
                MsgBox sMensaje, vbInformation, "Aviso"
                If feOperaciones.Visible And feOperaciones.Enabled Then
                    feOperaciones.SetFocus
                End If
                feOperaciones.TopRow = i
                feOperaciones.row = i
                feOperaciones.Col = 4
                validaDatosOperacion = False
                Exit Function
            End If
        End If
    Next
End Function
Private Function EsHaberes(ByVal sCta As String) As Boolean
    Dim ssql As String
    Dim cCap As COMDCaptaGenerales.COMDCaptAutorizacion
    Set cCap = New COMDCaptaGenerales.COMDCaptAutorizacion
        EsHaberes = cCap.EsHaberes(sCta)
    Set cCap = Nothing
End Function
Private Sub SetMatrizOperaciones()
    Dim i As Integer
    Set fMatOperaciones = Nothing
    ReDim fMatOperaciones(6, 0)
    If feOperaciones.TextMatrix(1, 0) <> "" Then
        For i = 1 To feOperaciones.Rows - 1
            ReDim Preserve fMatOperaciones(6, i)
            fMatOperaciones(1, i) = feOperaciones.TextMatrix(i, 1) 'Nro Voucher
            fMatOperaciones(2, i) = feOperaciones.TextMatrix(i, 2) 'Tipo Operación
            fMatOperaciones(3, i) = feOperaciones.TextMatrix(i, 3) 'Monto Operación
            fMatOperaciones(4, i) = feOperaciones.TextMatrix(i, 4) 'Estado Operación
            fMatOperaciones(5, i) = feOperaciones.TextMatrix(i, 5) 'Detalle Operación
            fMatOperaciones(6, i) = feOperaciones.TextMatrix(i, 7) 'Glosa
        Next
    End If
End Sub
Private Sub SetFlexOperaciones()
    Dim i As Integer
    Dim lnMontoOperaciones As Currency
    Call LimpiaFlex(feOperaciones)
    For i = 1 To UBound(fMatOperaciones, 2)
        feOperaciones.AdicionaFila
        feOperaciones.TextMatrix(i, 1) = fMatOperaciones(1, i) 'Nro Voucher
        feOperaciones.TextMatrix(i, 2) = fMatOperaciones(2, i) 'Tipo Operación
        feOperaciones.TextMatrix(i, 3) = fMatOperaciones(3, i) 'Monto Operación
        feOperaciones.TextMatrix(i, 4) = fMatOperaciones(4, i) 'Estado Operación
        feOperaciones.TextMatrix(i, 5) = fMatOperaciones(5, i) 'Detalle Operación
        feOperaciones.TextMatrix(i, 6) = fMatOperaciones(2, i) 'Tipo Operación Temporal
        feOperaciones.TextMatrix(i, 7) = fMatOperaciones(6, i) 'Glosa
        lnMontoOperaciones = lnMontoOperaciones + fMatOperaciones(3, i)
    Next
    txtOpeTotal.Text = Format(lnMontoOperaciones, gsFormatoNumeroView)
End Sub
Private Sub BloqueaControles(ByVal pbBloquea As Boolean)
    TxtTitPersCod.Enabled = pbBloquea
    cboMoneda.Enabled = pbBloquea
    txtCtaIFCod.Enabled = pbBloquea
    txtVoucherFecha.Enabled = pbBloquea
    txtVoucherNro.Enabled = pbBloquea
    txtVoucherMonto.Enabled = pbBloquea
    chkConfirmar.Enabled = pbBloquea
End Sub
Private Sub InhabilitaDatosVoucher()
    'Si esta vacio habilitar
    If feOperaciones.TextMatrix(1, 0) = "" Then
        txtVoucherMonto.Locked = False
        txtVoucherNro.Locked = False
        cboMoneda.Locked = False
    Else
        txtVoucherMonto.Locked = True
        txtVoucherNro.Locked = True
        cboMoneda.Locked = True
    End If
End Sub


