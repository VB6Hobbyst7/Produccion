VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCheque 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RECEPCIÓN DE CHEQUES"
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10260
   Icon            =   "frmCheque.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   10260
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   345
      Left            =   160
      TabIndex        =   44
      Top             =   6550
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   345
      Left            =   9060
      TabIndex        =   43
      Top             =   6550
      Width           =   1050
   End
   Begin VB.CommandButton cmdLimpiar 
      Caption         =   "&Limpiar"
      Height          =   345
      Left            =   6960
      TabIndex        =   42
      Top             =   6550
      Width           =   1050
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   345
      Left            =   8010
      TabIndex        =   41
      Top             =   6550
      Width           =   1050
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6375
      Left            =   60
      TabIndex        =   19
      Top             =   60
      Width           =   10140
      _ExtentX        =   17886
      _ExtentY        =   11245
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "&Registro de Cheque"
      TabPicture(0)   =   "frmCheque.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraVoucher"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "chkVoucher"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      Begin VB.Frame Frame3 
         Caption         =   "Datos de Operación"
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
         Height          =   3405
         Left            =   60
         TabIndex        =   29
         Top             =   2880
         Width           =   9975
         Begin VB.CommandButton cmdEditar 
            Caption         =   "&Editar"
            Height          =   300
            Left            =   1865
            TabIndex        =   14
            Top             =   3020
            Width           =   885
         End
         Begin VB.CommandButton cmdAgregar 
            Caption         =   "&Agregar"
            Height          =   300
            Left            =   120
            TabIndex        =   13
            Top             =   3020
            Width           =   885
         End
         Begin VB.CommandButton cmdCancelar 
            Caption         =   "&Cancelar"
            Height          =   300
            Left            =   5085
            TabIndex        =   17
            Top             =   3020
            Visible         =   0   'False
            Width           =   885
         End
         Begin VB.CommandButton cmdAceptar 
            Caption         =   "&Aceptar"
            Height          =   300
            Left            =   4200
            TabIndex        =   16
            Top             =   3020
            Visible         =   0   'False
            Width           =   885
         End
         Begin VB.CommandButton cmdQuitar 
            Caption         =   "&Quitar"
            Height          =   300
            Left            =   980
            TabIndex        =   15
            Top             =   3020
            Width           =   885
         End
         Begin SICMACT.FlexEdit feOperaciones 
            Height          =   2730
            Left            =   120
            TabIndex        =   30
            Top             =   240
            Width           =   9780
            _extentx        =   17251
            _extenty        =   4815
            cols0           =   9
            highlight       =   2
            allowuserresizing=   3
            encabezadosnombres=   "N°-Código-Operación-Monto-Estado-Detalle de la Operación-OperacionTmp-Glosa-Aux"
            encabezadosanchos=   "0-1200-3000-1200-1200-2800-0-2500-0"
            font            =   "frmCheque.frx":0326
            font            =   "frmCheque.frx":034E
            font            =   "frmCheque.frx":0376
            font            =   "frmCheque.frx":039E
            font            =   "frmCheque.frx":03C6
            fontfixed       =   "frmCheque.frx":03EE
            lbultimainstancia=   -1
            tipobusqueda    =   6
            columnasaeditar =   "X-X-2-3-X-5-6-7-X"
            textstylefixed  =   4
            listacontroles  =   "0-0-3-0-0-1-0-0-0"
            encabezadosalineacion=   "C-L-L-R-C-L-C-L-C"
            formatosedit    =   "0-0-0-2-0-0-0-0-0"
            cantentero      =   12
            textarray0      =   "N°"
            lbflexduplicados=   0
            lbformatocol    =   -1
            lbpuntero       =   -1
            lbbuscaduplicadotext=   -1
            rowheight0      =   300
         End
         Begin VB.Label lblOpeTotal 
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
            Left            =   8760
            TabIndex        =   32
            Top             =   3000
            Width           =   1095
         End
         Begin VB.Label Label5 
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
            Left            =   8280
            TabIndex        =   31
            Top             =   3030
            Width           =   495
         End
      End
      Begin VB.CheckBox chkVoucher 
         Caption         =   "Voucher de Depósito de cheque"
         Enabled         =   0   'False
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   2190
         Width           =   2655
      End
      Begin VB.Frame fraVoucher 
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
         Height          =   690
         Left            =   80
         TabIndex        =   22
         Top             =   2190
         Width           =   9975
         Begin MSMask.MaskEdBox txtVoucherFecDep 
            Height          =   285
            Left            =   1320
            TabIndex        =   12
            Top             =   270
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
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
         Begin SICMACT.TxtBuscar txtVoucherCtaIFCod 
            Height          =   285
            Left            =   4320
            TabIndex        =   18
            Top             =   270
            Width           =   1620
            _extentx        =   2858
            _extenty        =   503
            appearance      =   0
            appearance      =   0
            font            =   "frmCheque.frx":0414
            appearance      =   0
            stitulo         =   ""
         End
         Begin VB.Label lblVoucherCtaIFDesc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   6000
            TabIndex        =   28
            Top             =   270
            Width           =   3855
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
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   280
            Width           =   1215
         End
         Begin VB.Label Label10 
            Caption         =   "Cuenta Depositada:"
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
            Left            =   2760
            TabIndex        =   23
            Top             =   280
            Width           =   1455
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Datos Generales"
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
         Height          =   1740
         Left            =   80
         TabIndex        =   20
         Top             =   360
         Width           =   9975
         Begin VB.CheckBox chkEnvioCamara 
            Caption         =   "Envío a Cámara"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   8160
            TabIndex        =   46
            Top             =   500
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.CheckBox chkMismoTitular 
            Caption         =   "Mismo Tit."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6840
            TabIndex        =   45
            Top             =   500
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox txtNroCheque4 
            Height          =   300
            Left            =   6000
            MaxLength       =   2
            TabIndex        =   5
            Top             =   480
            Width           =   405
         End
         Begin VB.TextBox txtNroCheque3 
            Height          =   300
            Left            =   3480
            MaxLength       =   3
            TabIndex        =   3
            Top             =   480
            Width           =   525
         End
         Begin VB.TextBox txtNroCheque2 
            Height          =   300
            Left            =   2880
            MaxLength       =   3
            TabIndex        =   2
            Top             =   480
            Width           =   525
         End
         Begin VB.TextBox txtNroCheque1 
            Height          =   300
            Left            =   2520
            MaxLength       =   3
            TabIndex        =   1
            Top             =   480
            Width           =   285
         End
         Begin VB.TextBox txtNroChequeCtaIF 
            Height          =   315
            Left            =   4080
            MaxLength       =   10
            TabIndex        =   4
            Top             =   480
            Width           =   1845
         End
         Begin VB.TextBox txtNroCheque 
            Height          =   300
            Left            =   960
            MaxLength       =   8
            TabIndex        =   0
            Top             =   480
            Width           =   1485
         End
         Begin VB.TextBox txtImporte 
            Alignment       =   1  'Right Justify
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
            Left            =   1800
            MaxLength       =   15
            TabIndex        =   8
            Text            =   "0.00"
            Top             =   1320
            Width           =   1580
         End
         Begin VB.CheckBox chkChequeGerencia 
            Caption         =   "Cheque de Gerencia"
            Height          =   195
            Left            =   120
            TabIndex        =   6
            Top             =   960
            Width           =   1815
         End
         Begin VB.ComboBox cboMoneda 
            Height          =   315
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   1320
            Width           =   735
         End
         Begin SICMACT.TxtBuscar txtGiradorCod 
            Height          =   285
            Left            =   4320
            TabIndex        =   9
            Top             =   960
            Width           =   1620
            _extentx        =   2858
            _extenty        =   503
            appearance      =   0
            appearance      =   0
            font            =   "frmCheque.frx":0438
            appearance      =   0
            tipobusqueda    =   3
            stitulo         =   ""
         End
         Begin SICMACT.TxtBuscar txtContactoCod 
            Height          =   285
            Left            =   4320
            TabIndex        =   10
            Top             =   1320
            Width           =   1620
            _extentx        =   2858
            _extenty        =   503
            appearance      =   0
            appearance      =   0
            font            =   "frmCheque.frx":045C
            appearance      =   0
            tipobusqueda    =   3
            stitulo         =   ""
         End
         Begin VB.Label lblBancoNombre 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   555
            Left            =   7200
            TabIndex        =   39
            Top             =   200
            Width           =   2400
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Bco"
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
            Height          =   210
            Left            =   3000
            TabIndex        =   38
            Top             =   240
            Width           =   300
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Age "
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
            Height          =   210
            Left            =   3600
            TabIndex        =   37
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta N° "
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
            Height          =   210
            Left            =   4320
            TabIndex        =   36
            Top             =   240
            Width           =   840
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Nº Cheque"
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
            Height          =   210
            Left            =   960
            TabIndex        =   35
            Top             =   285
            Width           =   855
         End
         Begin VB.Label lblContactoNombre 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   6000
            TabIndex        =   34
            Top             =   1320
            Width           =   3855
         End
         Begin VB.Label Label4 
            Caption         =   "Beneficiario"
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
            Left            =   3420
            TabIndex        =   33
            Top             =   1320
            Width           =   855
         End
         Begin VB.Label Label6 
            Caption         =   "Girador:"
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
            Left            =   3720
            TabIndex        =   27
            Top             =   975
            Width           =   615
         End
         Begin VB.Label lblGiradorNombre 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   6000
            TabIndex        =   26
            Top             =   960
            Width           =   3855
         End
         Begin VB.Label Label1 
            Caption         =   "N° Cheque:"
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
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label3 
            Caption         =   "Importe:"
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
            TabIndex        =   21
            Top             =   1320
            Width           =   615
         End
      End
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Height          =   480
      Left            =   60
      TabIndex        =   40
      Top             =   6480
      Width           =   10140
   End
End
Attribute VB_Name = "frmCheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************
'** Nombre : frmCheque
'** Descripción : Para registro de Cheques creado segun TI-ERS126-2013
'** Creación : EJVG, 20131213 11:00:00 AM
'**********************************************************************
Option Explicit

Dim fbNuevo As Boolean
Dim fMatOperaciones As Variant
Dim fnFilaActual As Integer
Dim fsTpoProgramaAho As String, fsTpoProgramaDPF As String, fsTpoProgramaCTS As String
Dim fsIFICod As String, fsIFITpo As String
Dim fnId As Long 'ID Cabecera
Dim oCCE As COMNCajaGeneral.NCOMCCE 'PASI20160803CCE
Dim fnDocTpoBenef As Integer 'PASI20160803 CCE
Dim fsDocNroBenef As String 'PASI20160803 CCE
Dim fbEsChequeCCE As Boolean 'PASI20161212 CCE
Dim fbEsChequexAjusteCCE As Boolean 'PASI20161212 CCE
Dim fnMontoAnterior As Currency 'PASI20161212 CCE
Private Sub Form_Load()
    cargarControles
    LimpiarControles
    CargarVariables
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If fbNuevo Then
        If MsgBox("¿Esta seguro de salir del Registro de Cheque?", vbYesNo + vbInformation, "Aviso") = vbNo Then
            Cancel = 1
            Exit Sub
        End If
    End If
End Sub
Public Sub Registrar()
    fbNuevo = True
    Caption = "RECEPCIÓN DE CHEQUES"
    feOperaciones.EncabezadosAnchos = "0-1200-3000-1200-0-2800-0-2500-0"
    feOperaciones.ColumnasAEditar = "X-X-2-3-X-5-6-7-X"
    Me.chkMismoTitular.Visible = True 'PASI20160912 CCE
    Me.chkEnvioCamara.Visible = True 'PASI20160912 CCE
    Show 1
End Sub
Public Sub Editar(ByVal pnId As Long)
    fbNuevo = False
    Caption = "MANTENIMIENTO DE CHEQUES"
    feOperaciones.EncabezadosAnchos = "0-1200-3000-1200-1200-2800-0-2500-0"
    'feOperaciones.ColumnasAEditar = "X-X-2-X-X-5-6-7-X" /**Comentado PASI20161212 CCE**/
    cmdLimpiar.Visible = False
    cmdImprimir.Visible = True
    'HabilitaxEditar False /**Comentado PASI20161212 CCE**/
    fnId = pnId
    If Not CargaDatos(fnId) Then
        MsgBox "No se pudo cargar los datos del Cheque, si el problema persiste comuniquese con el Dpto. de TI", vbInformation, "Aviso"
        Unload Me
    End If
    'PASI2016122 CCE************************
    If Not fbNuevo And fbEsChequeCCE And oCCE.CCE_EsChequeEnviado(fnId) Then
        If MsgBox("El cheque ya ha sido enviado a la CCE, se procederá a realizar el ajuste. " & Chr(10) & "¿Está seguro de continuar?", vbYesNo + vbInformation, "Aviso") = vbNo Then Exit Sub
        Me.Caption = "AJUSTE DE CHEQUE"
        txtGiradorCod.Enabled = False
        txtContactoCod.Enabled = False
        feOperaciones.ColumnasAEditar = "X-X-X-3-X-5-6-7-X"
        fbEsChequexAjusteCCE = True
        fnMontoAnterior = txtImporte.Text
    Else
        feOperaciones.ColumnasAEditar = "X-X-2-X-X-5-6-7-X"
    End If
    HabilitaxEditar False
    'PASI END*******************************
    Show 1
End Sub
Private Sub cboMoneda_Click()
    Dim obj As New clases.DOperacion
    Dim lnMoneda As Integer
    Dim lsColor As String
    Dim bHabilita As Boolean
    
    On Error GoTo ErrCboMoneda
    Screen.MousePointer = 11
    lsColor = &H80000005
    txtVoucherCtaIFCod.Text = ""
    lblVoucherCtaIFDesc.Caption = ""
    If cboMoneda.ListIndex <> -1 Then
        lnMoneda = CInt(Trim(Right(cboMoneda.Text, 2)))
        txtVoucherCtaIFCod.psRaiz = "Cuentas de Instituciones Financieras"
        bHabilita = txtVoucherCtaIFCod.Enabled
        txtVoucherCtaIFCod.rs = obj.listarCuentasEntidadesFinacieras("_1_[12]" & CStr(lnMoneda) & "%", CStr(lnMoneda))
        txtVoucherCtaIFCod.Enabled = bHabilita
        If lnMoneda = 1 Then
            lsColor = &H80000005
        Else
            lsColor = &HC0FFC0
        End If
    End If
    txtImporte.BackColor = lsColor
    txtVoucherCtaIFCod.BackColor = lsColor
    lblVoucherCtaIFDesc.BackColor = lsColor
    Set obj = Nothing
    
    Screen.MousePointer = 0
    Exit Sub
ErrCboMoneda:
    Screen.MousePointer = 0
    MsgBox err.Description, vbCritical, "Aviso"
End Sub
Private Sub cboMoneda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtImporte.Visible And txtImporte.Enabled Then txtImporte.SetFocus
    End If
End Sub
Private Sub chkChequeGerencia_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cboMoneda.Visible And cboMoneda.Enabled Then cboMoneda.SetFocus
    End If
End Sub
Private Sub chkVoucher_KeyPress(KeyAscii As Integer)
    If chkVoucher.value = 1 Then
        If txtVoucherFecDep.Visible And txtVoucherFecDep.Enabled Then txtVoucherFecDep.SetFocus
    Else
        If cmdAgregar.Visible And cmdAgregar.Enabled Then cmdAgregar.SetFocus
    End If
End Sub
Private Sub cmdAgregar_Click()
    If Not validaDatosCheque Then Exit Sub
    If feOperaciones.TextMatrix(1, 0) <> "" Then
        If Not validaDatosOperacion Then
            Exit Sub
        End If
    End If
    If feOperaciones.Rows - 1 > 25 Then 'Limitamos a 25 el Nro. de Operaciones para que no descuadre el PDF
        MsgBox "Se ha llegado al máximo de operaciones a realizar", vbInformation, "Aviso"
        Exit Sub
    End If
    
    feOperaciones.lbEditarFlex = True
    feOperaciones.AdicionaFila
    fnFilaActual = feOperaciones.row
    
    HabilitaDatosCheque
    Call EstableceNroCorrelativoEnFlex(feOperaciones.row)
    feOperaciones.TextMatrix(feOperaciones.row, 3) = "0.00"
    feOperaciones.TextMatrix(feOperaciones.row, 4) = "PENDIENTE"
    If feOperaciones.Visible And feOperaciones.Enabled Then feOperaciones.SetFocus
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
        MsgBox "Este registro no podrá ser modificado porque ya realizaron operaciones con este Cheque", vbCritical, "Aviso"
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
        HabilitaDatosCheque
    End If
End Sub
Private Sub cmdAceptar_Click()
    If feOperaciones.TextMatrix(1, 0) <> "" Then
        If Not validaDatosOperacion Then
            Exit Sub
        End If
    End If
    SetMatrizOperaciones 'Guardamos el Flex en la Matriz
    HabilitaDatosCheque
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
    HabilitaDatosCheque
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
Private Sub cmdImprimir_Click()
    ImprimeConstanciaPDF fnId
End Sub
Private Sub cmdLimpiar_Click()
    LimpiarControles
End Sub
Private Sub cmdGrabar_Click()
    Dim oSis As NConstSistemas
    Dim oFun As NContFunciones
    Dim oImpre As NContImprimir
    Dim oPrevio As clsprevio
    Dim oForm As frmChequeListaPendiente
    Dim oDR As COMNCajaGeneral.NCOMDocRec
    Dim ldFechaValoriza As Date
    Dim lbGerencia As Boolean
    Dim lbExito As Boolean
    Dim lnIDCheque As Long
    Dim ldDepFechaReg As Date
    Dim lsDepIFTpo As String, lsDepPersCod As String, lsDepCtaIF As String
    Dim lnTotalVerifica As Currency
    Dim lnMovNroPend As Long
    Dim lnMoneda As Moneda
    Dim lnImporte As Currency
    Dim lsCtaContRegistroD As String, lsCtaContRegistroH As String, lsCtaContDepositoD As String, lsCtaContDepositoH As String
    Dim lMatMovNro() As String
    Dim i As Integer
    Dim lsCadImpre As String
    Dim lsBcoCtaCont As String, lsBcoSubCtaCont As String

    On Error GoTo ErrGrabar
    If Not validaDatosCheque Then Exit Sub
    If Not validaDatosOperacion Then Exit Sub
    'Valide Acepte o Cancele los cambios en el Detalle de Operacion
    If cmdAceptar.Visible Or cmdCancelar.Visible Then
        MsgBox "Ud. debe aceptar o cancelar los cambios en el Detalle de Operaciones", vbInformation, "Aviso"
        If cmdAceptar.Visible And cmdAceptar.Enabled Then
            If cmdAceptar.Visible And cmdAceptar.Enabled Then cmdAceptar.SetFocus
        ElseIf cmdCancelar.Visible And cmdCancelar.Enabled Then
            If cmdCancelar.Visible And cmdCancelar.Enabled Then cmdCancelar.SetFocus
        End If
        Exit Sub
    End If
    lnTotalVerifica = feOperaciones.SumaRow(3)
    If lnTotalVerifica <> CCur(txtImporte.Text) Then
        MsgBox "El monto del cheque es diferente a la sumatoria del detalle de operaciones. Verifique.", vbInformation, "Aviso"
        Exit Sub
    End If
    If UBound(fMatOperaciones, 2) = 0 Then
        MsgBox "No existen datos de las operaciones ingresadas. Verifique.", vbCritical, "Aviso"
        Exit Sub
    End If

    lnMoneda = CInt(Trim(Right(cboMoneda.Text, 2)))
    lnImporte = CCur(txtImporte.Text)
    ldFechaValoriza = FechaValorizacion()
    lbGerencia = IIf(Me.chkChequeGerencia.value = 1, True, False)
    ldDepFechaReg = CDate(IIf(txtVoucherFecDep.Text = "__/__/____", "01/01/1900", txtVoucherFecDep.Text))
    lsDepIFTpo = Mid(txtVoucherCtaIFCod.Text, 1, 2)
    lsDepPersCod = Mid(txtVoucherCtaIFCod.Text, 4, 13)
    lsDepCtaIF = Mid(txtVoucherCtaIFCod.Text, 18, Len(txtVoucherCtaIFCod.Text))
    If Not fbNuevo And fnId = 0 Then
        MsgBox "Ud. debe seleccionar primero el cheque a realizar el mantenimiento.", vbCritical, "Aviso"
        Exit Sub
    End If
    
    If fbNuevo Then
        If chkVoucher.value = 1 Then
            If ldDepFechaReg < gdFecSis Then
                MsgBox "La fecha de Depósito del Voucher es anterior, es necesario que se relacione con una pendiente registrada el día " & Format(ldDepFechaReg, gsFormatoFechaView), vbInformation, "Aviso"
                Set oForm = New frmChequeListaPendiente
                lnMovNroPend = oForm.Inicio(ldDepFechaReg, lsDepIFTpo, lsDepPersCod, lsDepCtaIF, lnMoneda, lnImporte)
                Set oForm = Nothing
                If lnMovNroPend = 0 Then Exit Sub
            End If
        End If
        Set oSis = New NConstSistemas
        If chkEnvioCamara.value = 0 Then 'PASI20160913 CCE
            lsCtaContRegistroD = oSis.LeeConstSistema(466)
            lsCtaContRegistroD = Replace(Replace(lsCtaContRegistroD, "M", lnMoneda), "AG", Right(gsCodAge, 2))
            lsCtaContRegistroH = oSis.LeeConstSistema(467)
            lsCtaContRegistroH = Replace(Replace(lsCtaContRegistroH, "M", lnMoneda), "AG", Right(gsCodAge, 2))
        Else
            lsCtaContRegistroD = Replace(Replace(oSis.LeeConstSistema(557), "M", lnMoneda), "AG", Right(gsCodAge, 2))
            lsCtaContRegistroH = Replace(Replace(oSis.LeeConstSistema(558), "M", lnMoneda), "AG", Right(gsCodAge, 2))
        End If
        
        Set oSis = Nothing
        Set oFun = New NContFunciones
        If Not oFun.verificarUltimoNivelCta(lsCtaContRegistroD) Then
            MsgBox "La Cuenta Contable " & lsCtaContRegistroD & " no es de Ultimo Nivel, comunicarse con el Dpto. de Contabilidad", vbInformation, "Aviso"
            Set oFun = Nothing
            Exit Sub
        End If
        If Not oFun.verificarUltimoNivelCta(lsCtaContRegistroH) Then
            MsgBox "La Cuenta Contable " & lsCtaContRegistroH & " no es de Ultimo Nivel, comunicarse con el Dpto. de Contabilidad", vbInformation, "Aviso"
            Set oFun = Nothing
            Exit Sub
        End If
        If chkVoucher.value = 1 Then
            Set oSis = New NConstSistemas
            lsBcoCtaCont = "11" & CStr(lnMoneda) & IIf(lsDepPersCod = "1090100822183", "2", "3") & lsDepIFTpo
            lsBcoSubCtaCont = oFun.GetFiltroObjetos(1, lsBcoCtaCont, txtVoucherCtaIFCod.Text, False)
            If Len(lsBcoSubCtaCont) = 0 Then
                MsgBox "Esta cuenta contable " & lsBcoCtaCont & " no esta registrado en CtaIFFiltro, comunicarse con el Dpto. de TI", vbInformation, "Aviso"
                Exit Sub
            End If
            lsCtaContDepositoD = lsBcoCtaCont & lsBcoSubCtaCont
            lsCtaContDepositoH = oSis.LeeConstSistema(468)
            lsCtaContDepositoH = Replace(Replace(lsCtaContDepositoH, "M", lnMoneda), "AG", Right(gsCodAge, 2))
            Set oSis = Nothing
            If Not oFun.verificarUltimoNivelCta(lsCtaContDepositoD) Then
               MsgBox "La Cuenta Contable " & lsCtaContDepositoD & " no es de Ultimo Nivel, comunicarse con el Dpto. de Contabilidad", vbInformation, "Aviso"
               Set oFun = Nothing
               Exit Sub
            End If
            If Not oFun.verificarUltimoNivelCta(lsCtaContDepositoH) Then
               MsgBox "La Cuenta Contable " & lsCtaContDepositoH & " no es de Ultimo Nivel, comunicarse con el Dpto. de Contabilidad", vbInformation, "Aviso"
               Set oFun = Nothing
               Exit Sub
            End If
        End If
        Set oFun = Nothing
    Else 'PASI20161212 CCE
        Set oSis = New NConstSistemas
        If chkEnvioCamara.value = 1 Then
            lsCtaContRegistroD = Replace(Replace(oSis.LeeConstSistema(557), "M", lnMoneda), "AG", Right(gsCodAge, 2))
            lsCtaContRegistroH = Replace(Replace(oSis.LeeConstSistema(558), "M", lnMoneda), "AG", Right(gsCodAge, 2))
        End If
    End If

    If MsgBox("¿Esta seguro de Guardar los datos del Cheque?", vbYesNo + vbInformation, "Aviso") = vbNo Then Exit Sub
    Screen.MousePointer = 11
    cmdGrabar.Enabled = False
    Set oDR = New COMNCajaGeneral.NCOMDocRec
    If fbNuevo Then
        lnIDCheque = oDR.RegistrarCheque(gsCodArea, Right(gsCodAge, 2), gsCodUser, txtNroCheque.Text, txtNroCheque1.Text, txtNroCheque2.Text, txtNroCheque3.Text, txtNroChequeCtaIF.Text, txtNroCheque4.Text, _
                            fsIFITpo, fsIFICod, txtGiradorCod.Text, txtContactoCod.Text, lnImporte, gdFecSis, ldFechaValoriza, lnMoneda, lbGerencia, fMatOperaciones, _
                            ldDepFechaReg, lsDepIFTpo, lsDepPersCod, lsDepCtaIF, lnMovNroPend, lMatMovNro, lsCtaContRegistroD, lsCtaContRegistroH, lsCtaContDepositoD, lsCtaContDepositoH, lblContactoNombre.Caption, fnDocTpoBenef, fsDocNroBenef, lblBancoNombre.Caption, chkMismoTitular.value, chkEnvioCamara.value)
    Else
        'lnIDCheque = oDR.EditarCheque(fnId, txtGiradorCod.Text, txtContactoCod.Text, fMatOperaciones) /**Comentado PASI20161212 CCE**/
        lnIDCheque = oDR.EditarCheque(fnId, txtGiradorCod.Text, txtContactoCod.Text, fMatOperaciones, fbEsChequexAjusteCCE, CCur(txtImporte.Text), fnMontoAnterior, gdFecSis, Right(gsCodAge, 2), gsCodUser, txtNroCheque.Text, CInt(Trim(Right(cboMoneda.Text, 2))), lsCtaContRegistroD, lsCtaContRegistroH) 'PASI20161212 CCE
    End If
    Screen.MousePointer = 0
    If lnIDCheque > 0 Then
        MsgBox "Se ha registrado satisfactoriamente la Operación", vbInformation, "Aviso"
        ImprimeConstanciaPDF lnIDCheque
           'INICIO JHCU ENCUESTA 16-10-2019
        
         If fbNuevo And fnId <> 0 Then
            Encuestas gsCodUser, gsCodAge, "ERS0292019", "900031"
         End If
        'FIN
                If fbNuevo Then
            'Set oImpre = New NContImprimir
            'For i = 1 To UBound(lMatMovNro)
            '    lsCadImpre = lsCadImpre & oImpre.ImprimeAsientoContable(lMatMovNro(i), gnLinPage, gnColPage, "RECEPCIÓN DE CHEQUE", , "179") & oImpresora.gPrnSaltoPagina
            'Next
            'Set oImpre = Nothing
            'Set oPrevio = New clsprevio
            'oPrevio.Show lsCadImpre, "RECEPCIÓN DE CHEQUE", False, gnLinPage
            'Set oPrevio = Nothing
            If MsgBox("¿Desea registrar la recepción de otro cheque?", vbYesNo + vbInformation, "Aviso") = vbNo Then
                fbNuevo = False
                Unload Me
                Exit Sub
            Else
                cmdLimpiar_Click
            End If
        Else
            Unload Me
            Exit Sub
        End If
    Else
        MsgBox "Hubo un error al realizar la operación con el Cheque, " & Chr(13) & "si el problema persiste comuniquese con el Dpto. de TI", vbCritical, "Aviso"
    End If
    Set oDR = Nothing
    cmdGrabar.Enabled = True
    Exit Sub
ErrGrabar:
    cmdGrabar.Enabled = True
    Screen.MousePointer = 0
    MsgBox err.Description, vbCritical, "Aviso"
End Sub
Private Sub cmdsalir_Click()
    Unload Me
End Sub
Private Sub LimpiarControles()
    chkChequeGerencia.value = 0
    txtNroCheque.Text = ""
    txtNroCheque1.Text = ""
    txtNroCheque2.Text = ""
    txtNroCheque3.Text = ""
    txtNroChequeCtaIF.Text = ""
    txtNroCheque4.Text = ""
    lblBancoNombre.Caption = ""
    cboMoneda.ListIndex = -1
    txtImporte.Text = "0.00"
    txtGiradorCod.Text = ""
    lblGiradorNombre.Caption = ""
    txtContactoCod.Text = ""
    lblContactoNombre.Caption = ""
    chkVoucher.value = 0
    chkVoucher_Click
    txtVoucherFecDep.Text = "__/__/____"
    txtVoucherCtaIFCod.Text = ""
    lblVoucherCtaIFDesc.Caption = ""
    'Limpiar Operaciones ***
    FormateaFlex feOperaciones
    SetMatrizOperaciones
    '***********************
    txtImporte.Locked = False
    txtNroCheque.Locked = False
    cboMoneda.Locked = False
    fnFilaActual = -1
    cmdCancelar_Click
    'PASI20161122 CCE *****
    fnDocTpoBenef = 0
    fsDocNroBenef = ""
    'PASI END*****
End Sub
Private Sub cargarControles()
    cboMoneda.Clear
    '''cboMoneda.AddItem "S/." & Space(100) & "1" 'marg ers044-2016
    cboMoneda.AddItem gcPEN_SIMBOLO & space(100) & "1" 'marg ers044-2016
    cboMoneda.AddItem "US$" & space(100) & "2"
End Sub
Private Sub CargarVariables()
    Dim oConsSis As New COMDConstSistema.NCOMConstSistema
    fsTpoProgramaAho = oConsSis.LeeConstSistema(454)
    fsTpoProgramaDPF = oConsSis.LeeConstSistema(455)
    fsTpoProgramaCTS = oConsSis.LeeConstSistema(456)
    Set oCCE = New COMNCajaGeneral.NCOMCCE
    Set oConsSis = Nothing
End Sub
Private Sub SetMatrizOperaciones()
    Dim i As Integer
    Set fMatOperaciones = Nothing
    ReDim fMatOperaciones(6, 0)
    If feOperaciones.TextMatrix(1, 0) <> "" Then
        For i = 1 To feOperaciones.Rows - 1
            ReDim Preserve fMatOperaciones(6, i)
            fMatOperaciones(1, i) = feOperaciones.TextMatrix(i, 1) 'Nro Cheque
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
    FormateaFlex feOperaciones
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
    lblOpeTotal = Format(lnMontoOperaciones, gsFormatoNumeroView)
End Sub

Private Sub txtContactoCod_EmiteDatos()
    Dim rs As ADODB.Recordset 'PASI20161223 CCE
    lblContactoNombre.Caption = ""
    If txtContactoCod.Text <> "" Then
        lblContactoNombre.Caption = txtContactoCod.psDescripcion
        If chkVoucher.Visible And chkVoucher.Enabled Then chkVoucher.SetFocus
        'PASI20161122 CCE **********
        Set rs = oCCE.CCE_ObtieneDocumentoPersona(txtContactoCod.psCodigoPersona)
        If Not (rs.EOF And rs.BOF) Then
            fnDocTpoBenef = Right(rs!cPersIDTpoDesc, 2)
            fsDocNroBenef = rs!cPersIDnro
        End If
        'PASI END**********
    End If
End Sub
Private Sub txtContactoCod_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Len(Trim(txtContactoCod.Text)) = 13 Then
        If chkVoucher.Visible And chkVoucher.Enabled Then chkVoucher.SetFocus
    End If
End Sub
Private Sub txtContactoCod_LostFocus()
    If Len(txtContactoCod.Text) <> 13 Then
        txtContactoCod.Text = ""
        lblContactoNombre.Caption = ""
    End If
End Sub
Private Sub txtGiradorCod_EmiteDatos()
    Dim rs As ADODB.Recordset
    lblGiradorNombre.Caption = ""
    If txtGiradorCod.Text = gsCodPersUser Then
        MsgBox "No se puede registrar un Cheque de si mismo", vbInformation, "Aviso"
        txtGiradorCod.Text = ""
        Exit Sub
    End If
    If txtGiradorCod.Text <> "" Then
        lblGiradorNombre.Caption = txtGiradorCod.psDescripcion
        If txtContactoCod.Visible And txtContactoCod.Enabled Then txtContactoCod.SetFocus
        'PASI20161122 CCE *****
'        Set rs = oCCE.CCE_ObtieneDocumentoPersona(txtGiradorCod.psCodigoPersona)
'        If Not (rs.EOF And rs.BOF) Then
'            fnDocTpoBenef = Right(rs!cPersIDTpoDesc, 2)
'            fsDocNroBenef = rs!cPersIDnro
'        End If
        'PASI END*****
    End If
End Sub
Private Sub chkVoucher_Click()
    If chkVoucher.value = 1 Then
        fraVoucher.Enabled = True
        If fbNuevo Then
            txtVoucherFecDep.Text = Format(gdFecSis, gsFormatoFechaView)
        End If
        If txtVoucherFecDep.Visible And txtVoucherFecDep.Enabled Then txtVoucherFecDep.SetFocus
    Else
        fraVoucher.Enabled = False
        If fbNuevo Then
            txtVoucherFecDep.Text = "__/__/____"
            txtVoucherCtaIFCod.Text = ""
            lblVoucherCtaIFDesc.Caption = ""
        End If
    End If
End Sub
Private Sub txtGiradorCod_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Len(Trim(txtGiradorCod.Text)) = 13 Then
        If txtContactoCod.Visible And txtContactoCod.Enabled Then txtContactoCod.SetFocus
    End If
End Sub
Private Sub txtGiradorCod_LostFocus()
    If Len(txtGiradorCod.Text) <> 13 Then
        txtGiradorCod.Text = ""
        lblGiradorNombre.Caption = ""
    End If
End Sub
Private Sub txtImporte_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtImporte, KeyAscii, 15)
    If KeyAscii <> 13 Then Exit Sub
    If txtGiradorCod.Visible And txtGiradorCod.Enabled Then txtGiradorCod.SetFocus
    If Not txtGiradorCod.Enabled Then cmdGrabar.SetFocus 'PASI20161212
End Sub
Private Sub txtImporte_LostFocus()
    txtImporte.Text = Format(txtImporte.Text, gsFormatoNumeroView)
    If Not IsNumeric(txtImporte.Text) Then
        MsgBox "Ingrese un monto válido", vbInformation, "Aviso"
        If txtImporte.Visible And txtImporte.Enabled Then txtImporte.SetFocus
        Exit Sub
    Else
        If CCur(txtImporte.Text) <= 0 Then
            MsgBox "Ingrese un monto mayor a cero", vbInformation, "Aviso"
            If txtImporte.Visible And txtImporte.Enabled Then txtImporte.SetFocus
            Exit Sub
        End If
    End If
End Sub
Private Sub txtNroCheque_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii, True)
    If KeyAscii = 13 Then
        If Len(txtNroCheque.Text) > 0 Then
            'If txtNroCheque1.Visible And txtNroCheque1.Enabled Then txtNroCheque1.SetFocus '***Comentado PASI20161207 CCE***/
            'PASI20161207 CCE***********
            If txtNroCheque1.Visible And txtNroCheque1.Enabled Then
                txtNroCheque1.Text = oCCE.ObtieneDigitoChequeoCheque(txtNroCheque.Text)
                txtNroCheque1.SetFocus
            End If
            'PASI END*******************
        End If
    End If
End Sub
Private Sub txtNroCheque_LostFocus()
    txtNroCheque.Text = Trim(txtNroCheque.Text)
End Sub
Private Sub txtNroCheque1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(txtNroCheque1.Text) > 0 Then
            If txtNroCheque2.Visible And txtNroCheque2.Enabled Then txtNroCheque2.SetFocus
        Else
            MsgBox "Dígito no Válido, Verifique y presione Enter", vbInformation, "Aviso"
        End If
    ElseIf KeyAscii = vbKeyBack Then
        If Len(txtNroCheque1.Text) = 0 Then
            If txtNroCheque.Visible And txtNroCheque.Enabled Then txtNroCheque.SetFocus
        End If
    End If
End Sub
Private Sub txtNroCheque1_LostFocus()
    txtNroCheque1.Text = Trim(txtNroCheque1.Text)
End Sub
Private Sub txtNroCheque2_GotFocus()
    fEnfoque txtNroCheque2
End Sub
Private Sub txtNroCheque2_KeyPress(KeyAscii As Integer)
    Dim oCCE As COMNCajaGeneral.NCOMCCE 'PASI20160802 CCE
    KeyAscii = NumerosEnteros(KeyAscii)
    Dim oCajG As New COMNCajaGeneral.NCOMCajaGeneral
    Dim rsC As New ADODB.Recordset
    
    lblBancoNombre.Caption = ""
    fsIFICod = ""
    fsIFITpo = ""
    Set oCCE = New COMNCajaGeneral.NCOMCCE
    
    If KeyAscii = 13 Then
        
        If Len(txtNroCheque2.Text) = 3 Then
            'Modificado PASI20160802 CCE *************
            'Set rsC = oCajG.GetBancosCod(txtNroCheque2.Text)
            Set rsC = oCCE.CCE_ObtieneIFIxCodBCR(txtNroCheque2.Text)
            '***********************************
            If Not rsC.EOF Then
                'Modificado PASI20160802 CCE *************
                'lblBancoNombre.Caption = Trim(rsC!cNomBanco)
                'fsIFICod = Mid(rsC!cPersCod, 4, Len(rsC!cPersCod))
                'fsIFITpo = Left(rsC!cPersCod, 2)
                
                lblBancoNombre.Caption = Trim(rsC!cPersNombre)
                fsIFICod = Mid(rsC!cPersCodIfi, 4, Len(rsC!cPersCodIfi))
                fsIFITpo = Left(rsC!cPersCodIfi, 2)
                '********************************
                
                If Len(fsIFICod) = 0 Then
                    MsgBox "Código de Banco no Válido, Verifique y presione Enter, si el problema persiste comuniquese con el Dpto. de Contabilidad para el mantenimiento del mismo", vbInformation, "Aviso"
                    If txtNroCheque2.Visible And txtNroCheque2.Enabled Then txtNroCheque2.SetFocus
                Else
                    If txtNroCheque3.Visible And txtNroCheque3.Enabled Then txtNroCheque3.SetFocus
                End If
                'PASI20161229 CCE
                If Not oCCE.CCE_EsEntidadGiradoraCheque(txtNroCheque2.Text) And chkEnvioCamara.value = 1 Then
                    MsgBox "El banco ingresado no puede participar como" & Chr(13) & " entidad giradora de cheques. Verifique.", vbInformation, "¡Aviso!"
                    txtNroCheque2.Text = ""
                    lblBancoNombre.Caption = ""
                    txtNroCheque2.SetFocus
                    Exit Sub
                End If
            Else
                MsgBox "El código de banco ingresado no existe en la BD. Verifique.", vbInformation, "Aviso"
                If txtNroCheque2.Visible And txtNroCheque2.Enabled Then txtNroCheque2.SetFocus
            End If
        Else
            MsgBox "Debe indicar un Código de Banco y presione Enter", vbInformation, "Aviso"
            If txtNroCheque2.Visible And txtNroCheque2.Enabled Then txtNroCheque2.SetFocus
        End If
    ElseIf KeyAscii = vbKeyBack Then
        If Len(txtNroCheque2.Text) = 0 Then
            If txtNroCheque1.Visible And txtNroCheque1.Enabled Then txtNroCheque1.SetFocus
        End If
    End If
    Set rsC = Nothing
    Set oCajG = Nothing
End Sub
Private Sub txtNroCheque2_LostFocus()
    txtNroCheque2.Text = Trim(txtNroCheque2.Text)
End Sub
Private Sub txtNroCheque3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(txtNroCheque3.Text) > 0 Then
            If txtNroChequeCtaIF.Visible And txtNroChequeCtaIF.Enabled Then txtNroChequeCtaIF.SetFocus
        Else
            MsgBox "Agencia no Válida, Verifique y presione Enter", vbInformation, "Aviso"
        End If
    ElseIf KeyAscii = vbKeyBack Then
        If Len(txtNroCheque3.Text) = 0 Then
            If txtNroCheque2.Visible And txtNroCheque2.Enabled Then txtNroCheque2.SetFocus
        End If
    End If
End Sub
Private Sub txtNroCheque3_LostFocus()
    txtNroCheque3.Text = Trim(txtNroCheque3.Text)
End Sub
Private Sub txtNroCheque4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If chkChequeGerencia.Visible And chkChequeGerencia.Enabled Then chkChequeGerencia.SetFocus
    ElseIf KeyAscii = vbKeyBack Then
        If Len(txtNroCheque4.Text) = 0 Then
            If txtNroChequeCtaIF.Visible And txtNroChequeCtaIF.Enabled Then txtNroChequeCtaIF.SetFocus
        End If
    End If
End Sub
Private Sub txtNroCheque4_LostFocus()
    txtNroCheque4.Text = Trim(txtNroCheque4.Text)
End Sub
Private Sub txtNroChequeCtaIF_KeyPress(KeyAscii As Integer)
    Dim oDR As COMNCajaGeneral.NCOMDocRec
    
    If KeyAscii = 13 Then
        If Len(txtNroChequeCtaIF.Text) > 0 Then
            Set oDR = New COMNCajaGeneral.NCOMDocRec
            'PASI20140619
            'If oDR.ObtieneExisteDocCheque(txtNroCheque.Text, fsIFICod, fsIFITpo, txtNroChequeCtaIF.Text) Then /**Comentado PASI20161212 CCE**/
            If oDR.ObtieneExisteDocCheque(txtNroCheque.Text, fsIFICod, fsIFITpo, txtNroChequeCtaIF.Text) And fbNuevo Then 'PASI20161212 CCE
                MsgBox "El Nro. de Cheque ya fue registrado. Ingrese otro numero. ", vbInformation, "Aviso"
                LimpiarControles
                txtNroCheque.SetFocus
                Exit Sub
            Else ' end PASI
                'If txtNroCheque4.Visible And txtNroCheque4.Enabled Then txtNroCheque4.SetFocus '***Comentado PASI20161207 CCE***/
                If txtNroCheque4.Visible And txtNroCheque4.Enabled Then
                    txtNroCheque4.Text = oCCE.ObtieneDigitoChequeoCheque(txtNroCheque2.Text + txtNroCheque3) + oCCE.ObtieneDigitoChequeoCheque(txtNroChequeCtaIF.Text)
                    txtNroCheque4.SetFocus
                End If
            End If
            
        Else
            MsgBox "Debe ingresar Nro. de Cuenta, Verifique y presione Enter", vbInformation, "Aviso"
        End If
    ElseIf KeyAscii = vbKeyBack Then
        If Len(txtNroChequeCtaIF.Text) = 0 Then
            If txtNroCheque3.Visible And txtNroCheque3.Enabled Then txtNroCheque3.SetFocus
        End If
    End If
End Sub
Private Sub txtNroChequeCtaIF_LostFocus()
    txtNroChequeCtaIF.Text = Trim(txtNroChequeCtaIF.Text)
End Sub
Private Sub txtVoucherCtaIFCod_EmiteDatos()
    Dim obj As New clases.DOperacion
    lblVoucherCtaIFDesc.Caption = ""
    If txtVoucherCtaIFCod.Text <> "" Then
        lblVoucherCtaIFDesc.Caption = obj.recuperaTipoCuentaEntidadFinaciera(Mid(txtVoucherCtaIFCod.Text, 18, 10)) & " " & txtVoucherCtaIFCod.psDescripcion
    End If
    Set obj = Nothing
End Sub
Private Sub txtVoucherCtaIFCod_GotFocus()
    If cboMoneda.ListIndex = -1 Then
        MsgBox "Ud. debe de seleccionar primero la moneda", vbInformation, "Aviso"
        If cboMoneda.Visible And cboMoneda.Enabled Then cboMoneda.SetFocus
        Exit Sub
    End If
End Sub
Private Sub txtVoucherCtaIFCod_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cmdAgregar.Visible And cmdAgregar.Enabled Then cmdAgregar.SetFocus
    End If
End Sub
Private Sub txtVoucherCtaIFCod_LostFocus()
    If txtVoucherCtaIFCod.Text = "" Then
        lblVoucherCtaIFDesc.Caption = ""
    End If
End Sub
Private Sub txtVoucherFecDep_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtVoucherCtaIFCod.Visible And txtVoucherCtaIFCod.Enabled Then txtVoucherCtaIFCod.SetFocus
    End If
End Sub
Private Function validaDatosCheque() As Boolean
    Dim obj As COMNCajaGeneral.NCOMDocRec
    Dim lsFecha As String
    
    validaDatosCheque = True
    
    If Len(fsIFICod) = 0 Then
        MsgBox "Ud. debe indicar la Institución Financiera", vbInformation, "Aviso"
        If txtNroCheque2.Visible And txtNroCheque2.Enabled Then txtNroCheque2.SetFocus
        validaDatosCheque = False
        Exit Function
    End If
    If Len(txtNroCheque.Text) = 0 Then
        MsgBox "Ud. debe especificar el Nro. de Cheque", vbInformation, "Aviso"
        validaDatosCheque = False
        If txtNroCheque.Visible And txtNroCheque.Enabled Then txtNroCheque.SetFocus
        Exit Function
    Else
        If Len(txtNroCheque1.Text) = 0 Then
            MsgBox "Ud. debe especificar el Nro. de Cheque", vbInformation, "Aviso"
            validaDatosCheque = False
            If txtNroCheque1.Visible And txtNroCheque1.Enabled Then txtNroCheque1.SetFocus
            Exit Function
        End If
        If Len(txtNroCheque2.Text) = 0 Then
            MsgBox "Ud. debe especificar el Nro. de Cheque", vbInformation, "Aviso"
            validaDatosCheque = False
            If txtNroCheque2.Visible And txtNroCheque2.Enabled Then txtNroCheque2.SetFocus
            Exit Function
        End If
        If Len(txtNroCheque3.Text) = 0 Then
            MsgBox "Ud. debe especificar el Nro. de Cheque", vbInformation, "Aviso"
            validaDatosCheque = False
            If txtNroCheque3.Visible And txtNroCheque3.Enabled Then txtNroCheque3.SetFocus
            Exit Function
        End If
        If Len(txtNroChequeCtaIF.Text) = 0 Then
            MsgBox "Ud. debe especificar el Nro. de Cheque", vbInformation, "Aviso"
            validaDatosCheque = False
            If txtNroChequeCtaIF.Visible And txtNroChequeCtaIF.Enabled Then txtNroChequeCtaIF.SetFocus
            Exit Function
        End If
        If Len(txtNroCheque4.Text) = 0 Then
            MsgBox "Ud. debe especificar el Nro. de Cheque", vbInformation, "Aviso"
            validaDatosCheque = False
            If txtNroCheque4.Visible And txtNroCheque4.Enabled Then txtNroCheque4.SetFocus
            Exit Function
        End If
        'PASI20161207 CCE
        If Not txtNroCheque1.Text = oCCE.ObtieneDigitoChequeoCheque(txtNroCheque.Text) Then
            MsgBox "Ud. debe Verificar el Nro. de Cheque", vbInformation, "Aviso"
            validaDatosCheque = False
            txtNroCheque.SetFocus
            Exit Function
        End If
        If Not oCCE.CCE_EsEntidadGiradoraCheque(txtNroCheque2.Text) And chkEnvioCamara.value = 1 Then
            MsgBox "El banco ingresado no puede participar como" & Chr(13) & " entidad giradora de cheques. Verifique.", vbInformation, "¡Aviso!"
            txtNroCheque2.Text = ""
            lblBancoNombre.Caption = ""
            validaDatosCheque = False
            txtNroCheque2.SetFocus
            Exit Function
        End If
        If Not txtNroCheque4.Text = oCCE.ObtieneDigitoChequeoCheque(txtNroCheque2.Text + txtNroCheque3) + oCCE.ObtieneDigitoChequeoCheque(txtNroChequeCtaIF.Text) Then
            MsgBox "Ud. debe Verificar el Nro. de Cuenta", vbInformation, "Aviso"
            validaDatosCheque = False
            txtNroChequeCtaIF.SetFocus
            Exit Function
        End If
        If Not fbNuevo And fbEsChequeCCE And oCCE.CCE_EsChequeEnviado(fnId) Then
            If CCur(nVal(txtImporte.Text)) = fnMontoAnterior Then
                MsgBox "No se ha realizado ningún cambio en el importe del cheque. Verifique.", vbInformation, "¡Aviso!"
                validaDatosCheque = False
                txtImporte.SetFocus
                Exit Function
            End If
        End If
        'PASI END********************
        Set obj = New COMNCajaGeneral.NCOMDocRec
        If fbNuevo Then
            If obj.ExisteCheque(TpoDocCheque, txtNroCheque.Text, fsIFICod, fsIFITpo, txtNroChequeCtaIF.Text, txtNroCheque1.Text) Then
                MsgBox "Cheque ya se encuentra registrado, verifique.. ", vbExclamation, "Aviso"
                validaDatosCheque = False
                If txtNroCheque.Visible And txtNroCheque.Enabled Then txtNroCheque.SetFocus
                Set obj = Nothing
                Exit Function
            End If
        End If
    End If
    If cboMoneda.ListIndex = -1 Then
        MsgBox "Ud. debe de seleccionar la Moneda del Cheque", vbInformation, "Aviso"
        validaDatosCheque = False
        If cboMoneda.Visible And cboMoneda.Enabled Then cboMoneda.SetFocus
        Exit Function
    End If
    If Not IsNumeric(txtImporte.Text) Then
        MsgBox "Ud. debe de especificar el Monto del Cheque", vbInformation, "Aviso"
        validaDatosCheque = False
        If txtImporte.Visible And txtImporte.Enabled Then txtImporte.SetFocus
        Exit Function
    Else
        If CCur(txtImporte.Text) <= 0 Then
            MsgBox "El Monto del Cheque debe ser mayor a cero", vbInformation, "Aviso"
            validaDatosCheque = False
            If txtImporte.Visible And txtImporte.Enabled Then txtImporte.SetFocus
            Exit Function
        End If
    End If
    If Len(Trim(txtGiradorCod.Text)) = 0 Then
        MsgBox "Ud. debe seleccionar el Girador del Cheque", vbInformation, "Aviso"
        validaDatosCheque = False
        If txtGiradorCod.Visible And txtGiradorCod.Enabled Then txtGiradorCod.SetFocus
        Exit Function
    End If
    If Len(Trim(txtContactoCod.Text)) = 0 Then
        MsgBox "Ud. debe seleccionar la persona a contactar", vbInformation, "Aviso"
        validaDatosCheque = False
        If txtContactoCod.Visible And txtContactoCod.Enabled Then txtContactoCod.SetFocus
        Exit Function
    End If
    If chkVoucher.value = 1 Then
        lsFecha = ValidaFecha(txtVoucherFecDep.Text)
        If Len(lsFecha) > 0 Then
            MsgBox lsFecha, vbInformation, "Aviso"
            validaDatosCheque = False
            If txtVoucherFecDep.Visible And txtVoucherFecDep.Enabled Then txtVoucherFecDep.SetFocus
            Exit Function
        Else
            If DateDiff("D", gdFecSis, CDate(txtVoucherFecDep.Text)) > 0 Then
                MsgBox "La Fecha de Deposito no puede ser mayor a la Fecha del Sistema", vbInformation, "Aviso"
                validaDatosCheque = False
                If txtVoucherFecDep.Visible And txtVoucherFecDep.Enabled Then txtVoucherFecDep.SetFocus
                Exit Function
            End If
        End If
        If Len(Trim(txtVoucherCtaIFCod.Text)) = 0 Then
            MsgBox "Ud. debe seleccionar la Cuenta a la que se realizó el deposito", vbInformation, "Aviso"
            validaDatosCheque = False
            If txtVoucherCtaIFCod.Visible And txtVoucherCtaIFCod.Enabled Then txtVoucherCtaIFCod.SetFocus
            Exit Function
        End If
    End If
    Set obj = Nothing
End Function
Private Function validaDatosOperacion() As Boolean
    Dim i As Integer, J As Integer
    Dim lnValor As Integer
    validaDatosOperacion = True
    'Valida No este vacio el Flex
    If feOperaciones.TextMatrix(1, 0) = "" Then
        MsgBox "Ud. debe de especificar las Operaciones", vbInformation, "Aviso"
        validaDatosOperacion = False
        If feOperaciones.Visible And feOperaciones.Enabled Then feOperaciones.SetFocus
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
                        If feOperaciones.Visible And feOperaciones.Enabled Then feOperaciones.SetFocus
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
    'Valida sea mayor a cero el campo detalle cuando el Tipo de Operación es Depósito en Lote
    For i = 1 To feOperaciones.Rows - 1
        lnValor = CInt(Trim(Right(feOperaciones.TextMatrix(i, 2), 2)))
        If lnValor = 5 Or lnValor = 7 Then
            If Val(Trim(feOperaciones.TextMatrix(i, 5))) <= 0 Then
                MsgBox "Para las Operaciones en Lote el Detalle debe ser mayor a cero", vbInformation, "Aviso"
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
Private Sub HabilitaDatosCheque()
    Dim oCCEL As New COMNCajaGeneral.NCOMCCE 'PASI20161213 CCE
    Dim lbHabilita As Boolean
    'Si esta vacio habilitar
    If feOperaciones.TextMatrix(1, 0) = "" Then
        lbHabilita = False
    Else
        lbHabilita = True
    End If
    txtImporte.Locked = IIf(Not fbNuevo And fbEsChequeCCE And oCCEL.CCE_EsChequeEnviado(fnId), False, lbHabilita) 'PASI20161213 CCE
    txtNroCheque.Locked = lbHabilita
    txtNroCheque1.Locked = lbHabilita
    txtNroCheque2.Locked = lbHabilita
    txtNroCheque3.Locked = lbHabilita
    txtNroChequeCtaIF.Locked = lbHabilita
    txtNroCheque4.Locked = lbHabilita
    cboMoneda.Locked = lbHabilita
End Sub
Private Sub EstableceNroCorrelativoEnFlex(ByVal row As Long)
    feOperaciones.TextMatrix(row, 1) = Trim(txtNroCheque.Text) & "-" & feOperaciones.TextMatrix(row, 0)
End Sub
Private Sub feOperaciones_OnCellChange(pnRow As Long, pnCol As Long)
    Dim lnTotal As Currency
    Dim i As Integer
    On Error GoTo ErrOnCellChangue
    If pnCol = 3 Then
        For i = 1 To feOperaciones.Rows - 1
            lnTotal = lnTotal + CCur(feOperaciones.TextMatrix(i, 3))
        Next
        lblOpeTotal.Caption = Format(lnTotal, gsFormatoNumeroView)
    End If
    If pnCol = 7 Then
        feOperaciones.TextMatrix(pnRow, pnCol) = UCase(Replace(Trim(feOperaciones.TextMatrix(pnRow, pnCol)), "'", ""))
    End If
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
    If feOperaciones.lbEditarFlex Then
        If feOperaciones.Col = 2 Then
            Set rs = oConst.RecuperaConstantes(10034, , "cConsDescripcion")
            feOperaciones.CargaCombo rs
        End If
        feOperaciones.row = fnFilaActual 'Mantiene la posición de la fila activa
    End If
    Set rs = Nothing
    Set oConst = Nothing
End Sub
'*OJO: Si se modifica el [feOperaciones_OnClickTxtBuscar] procedimiento tambien modificar el procedimiento [cmdDetalle_Click] del formulario [frmChequeOpePendiente]
'**|
'**v
Private Sub feOperaciones_OnClickTxtBuscar(psCodigo As String, psDescripcion As String)
    Dim row As Integer
    Dim lnOperacion As TipoOperacionCheque
    Dim oPersona As COMDPersona.UCOMPersona
    Dim frmBP As frmBuscaPersona
    Dim clsCap As COMNCaptaGenerales.NCOMCaptaGenerales
    Dim clsCuenta As UCapCuenta
    Dim rsPers As ADODB.Recordset
    Dim frmMntCap As frmCapMantenimientoCtas
    Dim frmCredSel As frmChequeDetCredSel
    Dim lnProducto As Producto
    Dim lsTpoPrograma As String
    Dim lsNroCuenta As String
    Dim lsNroCuentaTmp As String
    Dim lnMoneda As Moneda
    
    On Error GoTo ErrOnClickTxtBuscar
    row = feOperaciones.row
    lnMoneda = CInt(Trim(Right(cboMoneda.Text, 2)))
    
    If Trim(feOperaciones.TextMatrix(row, 2)) = "" Then
        MsgBox "Ud. debe seleccionar primero la Operación a realizar", vbInformation, "Aviso"
        feOperaciones.TopRow = row
        feOperaciones.row = row
        feOperaciones.Col = 2
        Exit Sub
    End If
    lnOperacion = CInt(Trim(Right(feOperaciones.TextMatrix(row, 2), 8)))
    Select Case lnOperacion '*** Constante 10034
        Case DPF_Apertura, AHO_Apertura, CTS_Apertura  'Apertura
            Dim frm As New frmChequeDetApert
            feOperaciones.TextMatrix(row, 5) = frm.Inicio(feOperaciones.TextMatrix(row, 5))
            psCodigo = feOperaciones.TextMatrix(row, 5)
            psDescripcion = feOperaciones.TextMatrix(row, 5)
            Set frm = Nothing
        Case DPF_AumentoCapital, AHO_Deposito, CTS_Deposito 'Depósitos y Aumento de Capital
            Set oPersona = New COMDPersona.UCOMPersona
            Set frmBP = New frmBuscaPersona
            Set oPersona = frmBP.Inicio
            lsNroCuenta = feOperaciones.TextMatrix(row, 5)
            
            If Not oPersona Is Nothing Then
                If oPersona.sPersCod <> "" Then
                    If oPersona.sPersCod <> gsCodPersUser Then
                        Set clsCap = New COMNCaptaGenerales.NCOMCaptaGenerales
                        Set clsCuenta = New UCapCuenta
                        Set rsPers = New ADODB.Recordset
                        Set frmMntCap = New frmCapMantenimientoCtas
                                   
                        Select Case lnOperacion
                            Case DPF_AumentoCapital
                                lnProducto = gCapPlazoFijo
                                lsTpoPrograma = fsTpoProgramaDPF
                            Case AHO_Deposito
                                lnProducto = gCapAhorros
                                lsTpoPrograma = fsTpoProgramaAho
                            Case CTS_Deposito
                                lnProducto = gCapCTS
                                lsTpoPrograma = fsTpoProgramaCTS
                        End Select
                        
                        Set rsPers = clsCap.GetCuentasPersona(oPersona.sPersCod, lnProducto, True, , lnMoneda, , , lsTpoPrograma)
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
                    Else
                        MsgBox "No se puede registrar un Cheque de si mismo", vbInformation, "Aviso"
                    End If
                End If
            End If

            feOperaciones.TextMatrix(row, 5) = lsNroCuenta
            psCodigo = lsNroCuenta
            psDescripcion = lsNroCuenta
        Case CRED_Pago, CRED_LiqCreditosSegDes
            Set oPersona = New COMDPersona.UCOMPersona
            Set frmBP = New frmBuscaPersona
            Set frmCredSel = New frmChequeDetCredSel
            Set oPersona = frmBP.Inicio
            lsNroCuenta = feOperaciones.TextMatrix(row, 5)
            
            If Not oPersona Is Nothing Then
                If oPersona.sPersCod <> "" Then
                    If oPersona.sPersCod <> gsCodPersUser Then
                        lsNroCuentaTmp = frmCredSel.Inicio(oPersona.sPersCod, lnMoneda)
                        If lsNroCuentaTmp <> "" Then
                            lsNroCuenta = lsNroCuentaTmp
                        End If
                    Else
                        MsgBox "No se puede registrar un Cheque de si mismo", vbInformation, "Aviso"
                    End If
                End If
            End If
            
            feOperaciones.TextMatrix(row, 5) = lsNroCuenta
            psCodigo = lsNroCuenta
            psDescripcion = lsNroCuenta
        Case CTS_DepositoLote, CRED_PagoLote, AHO_DepositoLote, AHO_AperturaLote, AHO_DepositoHaberesLote, DPF_AperturaLote 'Lote
            Dim frmLote As New frmChequeDetLote
            feOperaciones.TextMatrix(row, 5) = frmLote.Inicio(Val(feOperaciones.TextMatrix(row, 5)))
            psCodigo = feOperaciones.TextMatrix(row, 5)
            psDescripcion = feOperaciones.TextMatrix(row, 5)
            Set frmLote = Nothing
        Case Else
            MsgBox "Esta Operación no esta configurado para este proceso," & Chr(13) & "Comuniquese con el Dpto. de TI", vbInformation, "Aviso"
            Exit Sub
    End Select
        
    Set frmCredSel = Nothing
    Set clsCap = Nothing
    Set clsCuenta = Nothing
    Set rsPers = Nothing
    Set frmMntCap = Nothing
    Set oPersona = Nothing
    Set frmBP = Nothing
    Exit Sub
ErrOnClickTxtBuscar:
    MsgBox err.Description, vbCritical, "Aviso"
End Sub
Private Sub feOperaciones_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    Dim Editar() As String
    Dim lnTotal As Currency
    Dim oNCred As COMNCredito.NCOMCredito
    Dim lnMonto As Currency, lnITF As Currency
    Dim lnValor As TipoOperacionCheque
    
    Editar = Split(Me.feOperaciones.ColumnasAEditar, "-")
    If Editar(pnCol) = "X" Or pnCol = 1 Then
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
            'Agregar ITF en Pago Crédito ***
            lnValor = Val(Trim(Right(feOperaciones.TextMatrix(pnRow, 2), 2)))
            If lnValor = 0 Then
                MsgBox "Ud. debe seleccionar primero la Operación a realizar", vbInformation, "Aviso"
                Cancel = False
                Exit Sub
            ElseIf lnValor = CRED_Pago Then
                Set oNCred = New COMNCredito.NCOMCredito
                lnMonto = CCur((feOperaciones.TextMatrix(pnRow, 3)))
                lnITF = oNCred.DameMontoITF(lnMonto)
                Set oNCred = Nothing
                If lnITF > 0 Then
                    If MsgBox("El monto de Pago está afecto a un ITF de " & Format(lnITF, gsFormatoNumeroView) & Chr(13) & "¿Desea agregar el ITF en el Monto de Pago?", vbYesNo + vbInformation, "Aviso") = vbYes Then
                        feOperaciones.TextMatrix(pnRow, 3) = Format(lnMonto + lnITF, gsFormatoNumeroView)
                    End If
                End If
            End If
            '***
            lnTotal = feOperaciones.SumaRow(pnCol)
            If lnTotal > CCur(txtImporte.Text) Then
                MsgBox "El Monto del Cheque es diferente a la Sumatoria del Detalle de Operaciones, verifique", vbInformation, "Aviso"
                Cancel = False
                Exit Sub
            End If
        End If
    End If
End Sub
Private Function FechaValorizacion() As Date
    Dim oCap As New COMNCaptaGenerales.NCOMCaptaDefinicion
    Dim dFecha As Date
    Dim lnFeriado As Integer, lnMoneda As Moneda, lnDiaValoriza As Integer
    Dim lnPlazaChq As ChequePlaza

    lnPlazaChq = 0
    lnMoneda = CInt(Trim(Right(Trim(cboMoneda.Text), 1)))
    lnDiaValoriza = GetDiasMinValorizacion(lnMoneda, lnPlazaChq)
    dFecha = DateAdd("d", CDate(gdFecSis), lnDiaValoriza)

    lnFeriado = oCap.ObtenerFeriado(gdFecSis, dFecha)
    dFecha = DateAdd("d", CDate(dFecha), lnFeriado)
    
    If Weekday(dFecha, vbMonday) = 6 Then
        dFecha = DateAdd("d", 2, dFecha)
    ElseIf Weekday(dFecha, vbMonday) = 7 Then
        dFecha = DateAdd("d", 1, dFecha)
        dFecha = CDate(dFecha) + 1
    Else
        'VERIFICA SI EL DIA Q SE VALORIZA ES FERIADO--AVMM--30-10-2006
        lnFeriado = oCap.ObtenerFeriado(dFecha, dFecha)
        dFecha = DateAdd("d", CDate(dFecha), lnFeriado)
        dFecha = CDate(dFecha) + 1
    End If
    'VERIFICA SI EL DIA Q SE VALORIZA ES FERIADO--AVMM--30-10-2006
    lnFeriado = oCap.ObtenerFeriado(dFecha, dFecha)
    dFecha = DateAdd("d", CDate(dFecha), lnFeriado)
    
    FechaValorizacion = dFecha
    Set oCap = Nothing
End Function
Private Function GetDiasMinValorizacion(Optional nmoneda As Moneda = gMonedaNacional, Optional nPlaza As ChequePlaza = gChqPlazaLocal, Optional nTipoCheque As ChequeTipo = gChqTpoSimple) As Integer
    Dim oCap As New COMNCaptaGenerales.NCOMCaptaDefinicion
    GetDiasMinValorizacion = oCap.GetDiasMinValorizacion(nmoneda, nPlaza, nTipoCheque)
    Set oCap = Nothing
End Function
Private Sub HabilitaxEditar(ByVal pbHabilita As Boolean)
    txtNroCheque.Locked = Not pbHabilita
    txtNroCheque1.Locked = Not pbHabilita
    txtNroCheque2.Locked = Not pbHabilita
    txtNroCheque3.Locked = Not pbHabilita
    txtNroChequeCtaIF.Locked = Not pbHabilita
    txtNroCheque4.Locked = Not pbHabilita
    cboMoneda.Locked = Not pbHabilita
    txtImporte.Locked = IIf(Not fbNuevo And fbEsChequeCCE And oCCE.CCE_EsChequeEnviado(fnId), False, Not pbHabilita)
    'Girador si esta disponible
    chkChequeGerencia.Enabled = pbHabilita
    Me.chkVoucher.Enabled = pbHabilita
    txtVoucherFecDep.Enabled = pbHabilita
    txtVoucherCtaIFCod.Enabled = pbHabilita
    'txtVoucherCtaIFCod.EnabledText = pbHabilita
    chkMismoTitular.Visible = True 'PASI20161211 CCE
    chkEnvioCamara.Visible = True 'PASI20161211 CCE
End Sub
Private Function CargaDatos(ByVal pnId As Long) As Boolean
    Dim obj As New COMNCajaGeneral.NCOMDocRec
    Dim rsCab As New ADODB.Recordset
    Dim rsDet As New ADODB.Recordset
    Dim rsCCE As New ADODB.Recordset 'PASI20161211 CCE
    Dim i As Integer
    
    On Error GoTo ErrCargaDatos
    Set rsCab = obj.ChequexID(pnId)
    If Not RSVacio(rsCab) Then
        txtNroCheque.Text = rsCab!cNroDoc
        txtNroCheque1.Text = rsCab!cNroCheque1
        txtNroCheque2.Text = rsCab!cNroCheque2
        txtNroCheque2_KeyPress 13
        txtNroCheque3.Text = rsCab!cNroCheque3
        txtNroChequeCtaIF.Text = rsCab!cIFCta
        txtNroCheque4.Text = rsCab!cNroCheque4
        'PASI20161212 CCE********************
        If rsCab!nMismoTitCCE = -1 Then
            chkMismoTitular.Enabled = False
            chkEnvioCamara.Enabled = False
            chkEnvioCamara.value = 0
        Else
            chkMismoTitular.value = rsCab!nMismoTitCCE
            chkEnvioCamara.value = 1
            fbEsChequeCCE = True
        End If
        chkMismoTitular.Enabled = False
        chkEnvioCamara.Enabled = False
        'PASI END ***************************
        chkChequeGerencia.value = IIf(rsCab!nConfCaja = 1, 1, 0)
        cboMoneda.ListIndex = IndiceListaCombo(cboMoneda, rsCab!nmoneda)
        txtImporte.Text = Format(rsCab!nMonto, gsFormatoNumeroView)
        txtGiradorCod.Text = rsCab!cPersCodGirador
        lblGiradorNombre.Caption = PstaNombre(rsCab!cPersNombreGirador)
        txtContactoCod.Text = rsCab!cPersCodContacto
        lblContactoNombre.Caption = PstaNombre(rsCab!cPersNombreContacto)
        chkVoucher.value = 0
        If DateDiff("d", rsCab!dDepFecha, CDate("1900-01-01")) <> 0 Then
            chkVoucher.value = 1
            txtVoucherFecDep.Text = Format(rsCab!dDepFecha, gsFormatoFechaView)
            txtVoucherCtaIFCod.Text = rsCab!cDepIFTpo & "." & rsCab!cDepPersCod & "." & rsCab!cDepIFCta
            txtVoucherCtaIFCod_EmiteDatos
        End If
        'chkVoucher_Click
        Set rsDet = obj.ChequeDetxID(pnId)
        Set fMatOperaciones = Nothing
        ReDim fMatOperaciones(6, 0)
        For i = 1 To rsDet.RecordCount
            ReDim Preserve fMatOperaciones(6, i)
            fMatOperaciones(1, i) = rsDet!cNroDoc 'Nro Cheque
            fMatOperaciones(2, i) = rsDet!cTipoOperacion & space(75) & CStr(rsDet!nTipoOperacion) 'Tipo Operación
            fMatOperaciones(3, i) = Format(rsDet!nMonto, gsFormatoNumeroView) 'Monto Operación
            fMatOperaciones(4, i) = rsDet!cEstado 'Estado Operación
            fMatOperaciones(5, i) = rsDet!lsDetalle 'Detalle Operación
            fMatOperaciones(6, i) = rsDet!lsGlosa 'Glosa
            rsDet.MoveNext
        Next
        SetFlexOperaciones 'Mostramos en Pantalla
        CargaDatos = True
    Else
        CargaDatos = False
    End If
    Set rsDet = Nothing
    Set rsCab = Nothing
    Set obj = Nothing
    Exit Function
ErrCargaDatos:
    CargaDatos = False
End Function
Public Sub ImprimeConstanciaPDF(ByVal pnId As Long)
    Dim oPDF As New cPDF
    Dim obj As New COMNCajaGeneral.NCOMDocRec
    Dim R As New ADODB.Recordset
    Dim lsLetras As String
    Dim lnMonto As Currency
    Dim lnlinea As Integer
    
    Set R = obj.ChequexImpresion(pnId)
    If R.EOF Then Exit Sub
    
    oPDF.Author = gsCodUser
    oPDF.Creator = "SICMACT - Negocio"
    oPDF.Producer = gsNomCmac
    oPDF.Subject = "CONSTANCIA DE OPERACIONES CON CHEQUE N° " & R!cNroCheque
    oPDF.Title = oPDF.Subject
    
    If Not oPDF.PDFCreate(App.Path & "\Spooler\Constancia_Operaciones_Cheque_" & R!cNroCheque & "_" & Format(gdFecSis, "YYYYMMDD") & "_" & Format(Time, "hhmmss") & ".pdf") Then
        Exit Sub
    End If

    oPDF.Fonts.Add "F1", "Arial", TrueType, Normal, WinAnsiEncoding
    oPDF.Fonts.Add "F2", "Arial", TrueType, Bold, WinAnsiEncoding
    oPDF.Fonts.Add "F3", "Arial", TrueType, Bold, WinAnsiEncoding
    oPDF.LoadImageFromFile App.Path & "\logo_cmacmaynas1.bmp", "Logo"
    oPDF.NewPage A4_Vertical
    
    oPDF.WImage 70, 450, 35, 105, "Logo"
    oPDF.WTextBox 53, 40, 15, 500, "SOLICITUD DE OPERACIONES CON CHEQUE", "F2", 14, hCenter
    oPDF.WTextBox 69, 40, 2, 515, "", "F2", 10, hCenter, vBottom, vbRed, 1, vbRed, True
    
    oPDF.WTextBox 73, 60, 10, 70, "Fecha", "F1", 7, hLeft, , , 1
    oPDF.WTextBox 73, 130, 10, 120, "Agencia", "F1", 7, hLeft, , , 1
    oPDF.WTextBox 73, 250, 10, 200, "Asesor", "F1", 7, hLeft, , , 1
    oPDF.WTextBox 83, 60, 10, 70, Format(R!dRegistro, gsFormatoFechaView), "F1", 7, hCenter, , vbBlue, 1
    oPDF.WTextBox 83, 130, 10, 120, R!cAgeDescripcion, "F1", 7, hCenter, , vbBlue, 1
    oPDF.WTextBox 83, 250, 10, 200, PstaNombre(R!cPersNombreUser, True), "F1", 7, hCenter, , vbBlue, 1
    
    oPDF.WTextBox 98, 40, 10, 515, "1. DATOS DEL CHEQUE", "F2", 8, hLeft, , vbWhite, 1, vbRed, True
    oPDF.WTextBox 109, 60, 10, 190, "N° de Cheque", "F1", 7, hLeft, , , 1
    oPDF.WTextBox 109, 250, 10, 305, "Institución Financiera", "F1", 7, hLeft, , , 1
    oPDF.WTextBox 119, 60, 10, 190, R!cNroCheque, "F1", 7, hCenter, , vbBlue, 1
    oPDF.WTextBox 119, 250, 10, 305, Left(R!cPersNombreIFi, 70), "F1", 7, hLeft, , vbBlue, 1
    oPDF.WTextBox 129, 60, 10, 345, "Importe (monto en números - letras):", "F1", 7, hLeft, , , 1
    oPDF.WTextBox 129, 405, 10, 75, "Moneda", "F1", 7, hLeft, , , 1
    oPDF.WTextBox 129, 480, 10, 75, "Fecha de Emisión", "F1", 7, hLeft, , , 1
    lnMonto = R!nMonto
    lsLetras = UCase(NumLet(lnMonto) & IIf(R!nmoneda = 2, "", " y " & IIf(InStr(1, lnMonto, ".") = 0, "00", Left(Mid(lnMonto, InStr(1, lnMonto, ".") + 1, 2) & "00", 2)) & "/100"))
    oPDF.WTextBox 139, 60, 10, 345, Format(lnMonto, gsFormatoNumeroView) & " - " & lsLetras, "F1", 7, hCenter, , vbBlue, 1
    oPDF.WTextBox 139, 405, 10, 75, R!cMoneda, "F1", 7, hCenter, , vbBlue, 1
    oPDF.WTextBox 139, 480, 10, 75, Format(R!dRegistro, gsFormatoFechaView), "F1", 7, hCenter, , vbBlue, 1
    oPDF.WTextBox 149, 60, 10, 247, "Girador", "F1", 7, hLeft, , , 1
    oPDF.WTextBox 159, 60, 10, 495, Left(PstaNombre(R!cPersNombreGirador, True), 135), "F1", 7, hLeft, , vbBlue, 1
    
    'oPDF.WTextBox 174, 40, 10, 515, "2. DATOS DEL SOLICITANTE", "F2", 8, hLeft, , vbWhite, 1, vbRed, True /**Comentado PASI20161223 **/
    oPDF.WTextBox 174, 40, 10, 515, "2. DATOS DEL BENEFICIARIO", "F2", 8, hLeft, , vbWhite, 1, vbRed, True 'PASI20161223 CCE
    oPDF.WTextBox 185, 60, 10, 190, "R.U.C / D.N.I:", "F1", 7, hLeft, , , 1
    oPDF.WTextBox 185, 250, 10, 305, "R. Social / Nombre", "F1", 7, hLeft, , , 1
    oPDF.WTextBox 195, 60, 10, 190, R!cDOIContacto, "F1", 7, hCenter, , vbBlue, 1
    oPDF.WTextBox 195, 250, 10, 305, PstaNombre(R!cPersNombreContacto, True), "F1", 7, hCenter, , vbBlue, 1
    oPDF.WTextBox 205, 60, 10, 190, "Fecha de Nacimiento/Constitución", "F1", 7, hLeft, , , 1
    oPDF.WTextBox 205, 250, 10, 230, "Dirección", "F1", 7, hLeft, , , 1
    oPDF.WTextBox 205, 480, 10, 75, "Distrito", "F1", 7, hLeft, , , 1
    oPDF.WTextBox 215, 60, 10, 190, Format(R!dNacimientoContacto, gsFormatoFechaView), "F1", 7, hCenter, , vbBlue, 1
    oPDF.WTextBox 215, 250, 10, 230, R!cDomicilioContacto, "F1", 7, hCenter, , vbBlue, 1
    oPDF.WTextBox 215, 480, 10, 75, R!cDistritoContacto, "F1", 7, hLeft, , vbBlue, 1
    oPDF.WTextBox 225, 60, 10, 190, "Telefono", "F1", 7, hLeft, , , 1
    oPDF.WTextBox 225, 250, 10, 115, "Celular", "F1", 7, hLeft, , , 1
    oPDF.WTextBox 225, 365, 10, 190, "Correo Electrónico", "F1", 7, hLeft, , , 1
    oPDF.WTextBox 235, 60, 10, 190, R!cTelefonoContacto, "F1", 7, hCenter, , vbBlue, 1
    oPDF.WTextBox 235, 250, 10, 115, R!cCelularContacto, "F1", 7, hCenter, , vbBlue, 1
    oPDF.WTextBox 235, 365, 10, 190, R!cEmailContacto, "F1", 7, hLeft, , vbBlue, 1

    oPDF.WTextBox 250, 40, 10, 515, "3. INFORMACION DE LA OPERACIÓN AUTORIZADA", "F2", 8, hLeft, , vbWhite, 1, vbRed, True
    oPDF.WTextBox 261, 60, 10, 350, "Se autoriza, luego de valorizado el importe total Cheque, realizar las siguientes operaciones:", "F1", 7, hLeft

    oPDF.WTextBox 271, 60, 10, 88, "Cod. Ope", "F1", 7, hCenter, , , 1, RGB(216, 216, 216), True
    oPDF.WTextBox 271, 148, 10, 160, "Operación", "F1", 7, hCenter, , , 1, RGB(216, 216, 216), True
    oPDF.WTextBox 271, 308, 10, 72, "Cta. Destino", "F1", 7, hCenter, , , 1, RGB(216, 216, 216), True
    oPDF.WTextBox 271, 380, 10, 126, "Concepto", "F1", 7, hCenter, , , 1, RGB(216, 216, 216), True
    oPDF.WTextBox 271, 506, 10, 50, "Saldo", "F1", 7, hCenter, , , 1, RGB(216, 216, 216), True

    lnlinea = 281
    lnMonto = 0
    'El detalle de la operación no debe superar 25 sino se descuadrará
    oPDF.WTextBox 281, 60, 240, 88, "", "F1", 7, hCenter, , , 1, RGB(216, 216, 216)
    oPDF.WTextBox 281, 148, 240, 160, "", "F1", 7, hCenter, , , 1, RGB(216, 216, 216)
    oPDF.WTextBox 281, 308, 240, 72, "", "F1", 7, hCenter, , , 1, RGB(216, 216, 216)
    oPDF.WTextBox 281, 380, 240, 126, "", "F1", 7, hCenter, , , 1, RGB(216, 216, 216)
    oPDF.WTextBox 281, 506, 240, 50, "", "F1", 7, hCenter, , , 1, RGB(216, 216, 216)
    Do While Not R.EOF
        oPDF.WTextBox lnlinea, 60, 10, 88, R!cNroDoc, "F1", 7, hCenter
        oPDF.WTextBox lnlinea, 148, 10, 160, Left(R!cTipoOperacion, 38), "F1", 7, hLeft
        oPDF.WTextBox lnlinea, 308, 10, 72, R!cCtaCodDet, "F1", 7, hCenter
        oPDF.WTextBox lnlinea, 380, 10, 126, Left(R!cMovDesc, 28), "F1", 7, hLeft
        oPDF.WTextBox lnlinea, 506, 10, 50, Format(R!nMontoDet, gsFormatoNumeroView), "F1", 7, hRight
        lnMonto = lnMonto + R!nMontoDet
        R.MoveNext
        lnlinea = lnlinea + 10
    Loop

    oPDF.WTextBox 521, 60, 10, 450, "MONTO TOTAL AUTORIZADO", "F1", 7, hCenter, , , 1, RGB(216, 216, 216), True
    oPDF.WTextBox 521, 506, 10, 50, Format(lnMonto, gsFormatoNumeroView), "F1", 7, hRight, , , 1, RGB(216, 216, 216), True
    oPDF.WTextBox 536, 60, 10, 450, "*La operaciones mencionadas deben efectuarse el mismo dia de valorizado el cheque, bajo responsabilidad de CAJA MAYNAS. S.A", "F1", 7, hLeft

    oPDF.WTextBox 555, 40, 10, 515, "4. DECLARACIONES Y FIRMAS", "F2", 8, hLeft, , vbWhite, 1, vbRed, True
    oPDF.WTextBox 566, 60, 10, 495, "Suscribo la presente: (i) Declarando la veracidad y certeza de los datos consignados en la presente solicitud; (ii) Aceptando los términos y condiciones", "F1", 7, hjustify
    oPDF.WTextBox 576, 60, 10, 495, "de las operaciones solicitadas; (iii) Declarando haber recibido previamente de CAJA MAYNAS S.A.  toda la información sobre las características y", "F1", 7, hjustify
    oPDF.WTextBox 586, 60, 10, 495, "tarifas de la operación; y (iv) Autorizando a CAJA MAYNAS S.A. a confirmar los datos consignados en la presente solicitud.", "F1", 7, hjustify
    oPDF.WTextBox 606, 60, 10, 505, "Las informaciones y declaraciones proporcionadas en la presente solicitud, tienen carácter de Declaración Jurada, conforme al Art. 179 de la Ley Nº 26072.", "F1", 7, hjustify

    oPDF.WTextBox 676, 120, 10, 170, "__________________________________________", "F1", 7, hLeft
    oPDF.WTextBox 686, 120, 10, 170, "Firma y sello del representante autorizado", "F1", 7, hCenter
    oPDF.WTextBox 676, 320, 10, 170, "__________________________________________", "F1", 7, hLeft
    oPDF.WTextBox 686, 320, 10, 170, "Firma y sello del representante autorizado", "F1", 7, hCenter
    oPDF.WTextBox 736, 120, 10, 170, "______________________________", "F1", 7, hCenter
    oPDF.WTextBox 746, 120, 10, 170, "VºBº ASESOR AL CLIENTE", "F1", 7, hCenter
    oPDF.WTextBox 736, 320, 10, 170, "____________________________________", "F1", 7, hCenter
    oPDF.WTextBox 746, 320, 10, 170, "VºBº SUPERVISOR DE OPERACIONES", "F1", 7, hCenter

    oPDF.PDFClose
    oPDF.Show
    
    Set R = Nothing
    Set oPDF = Nothing
    Set obj = Nothing
End Sub

