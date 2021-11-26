VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmPagosAdelantados 
   Caption         =   "Pagos Adelantados"
   ClientHeight    =   5070
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11085
   Icon            =   "frmPagosAdelantados.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   11085
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTabPagoAdelantado 
      Height          =   5055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11115
      _ExtentX        =   19606
      _ExtentY        =   8916
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Buscar"
      TabPicture(0)   =   "frmPagosAdelantados.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FEBuscar"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdGenerar"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdExaminar"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdBuscar"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "FEGenerar"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Pago Adelantado"
      TabPicture(1)   =   "frmPagosAdelantados.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label8"
      Tab(1).Control(1)=   "Label1"
      Tab(1).Control(2)=   "Label2"
      Tab(1).Control(3)=   "Label3"
      Tab(1).Control(4)=   "Label4"
      Tab(1).Control(5)=   "Label5"
      Tab(1).Control(6)=   "Label6"
      Tab(1).Control(7)=   "txtFecVen"
      Tab(1).Control(8)=   "txtFecIni"
      Tab(1).Control(9)=   "txtBuscarCtaCon"
      Tab(1).Control(10)=   "FrAgencia"
      Tab(1).Control(11)=   "cmdEliminarPagoAdelantado"
      Tab(1).Control(12)=   "txtImporte"
      Tab(1).Control(13)=   "cboRubro"
      Tab(1).Control(14)=   "txtDesPagAde"
      Tab(1).Control(15)=   "txtNroMes"
      Tab(1).Control(16)=   "cmdNuevoPagoAdelantado"
      Tab(1).Control(17)=   "cmdEditarPagoAdelantado"
      Tab(1).Control(18)=   "txtDesCtaCon"
      Tab(1).Control(19)=   "cmdAceptarPagoAdelantado"
      Tab(1).Control(20)=   "cmdCancelarPagoAdelantado"
      Tab(1).Control(21)=   "chkTodasAgencias"
      Tab(1).ControlCount=   22
      TabCaption(2)   =   "Amortizaciones"
      TabPicture(2)   =   "frmPagosAdelantados.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FEAmortizacion"
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame1 
         Caption         =   "Criterios para Buscar"
         Height          =   1455
         Left            =   2760
         TabIndex        =   34
         Top             =   360
         Width           =   6975
         Begin VB.CheckBox chkCancelado 
            Caption         =   "Considerar Pagos Adelantados con Saldo 0"
            Height          =   375
            Left            =   2640
            TabIndex        =   40
            Top             =   360
            Width           =   3375
         End
         Begin VB.ComboBox cboRubro2 
            Height          =   315
            Left            =   3120
            Style           =   2  'Dropdown List
            TabIndex        =   38
            Top             =   840
            Width           =   1815
         End
         Begin VB.Frame FrMoneda 
            Caption         =   "Moneda"
            Height          =   1095
            Left            =   120
            TabIndex        =   35
            Top             =   240
            Width           =   1935
            Begin VB.OptionButton OptMoneda 
               Caption         =   "Extranjera"
               Height          =   375
               Index           =   1
               Left            =   120
               TabIndex        =   37
               Top             =   600
               Width           =   1095
            End
            Begin VB.OptionButton OptMoneda 
               Caption         =   "Nacional"
               Height          =   375
               Index           =   0
               Left            =   120
               TabIndex        =   36
               Top             =   240
               Value           =   -1  'True
               Width           =   1095
            End
         End
         Begin VB.Label Label10 
            Caption         =   "Rubro"
            Height          =   255
            Left            =   2640
            TabIndex        =   39
            Top             =   960
            Width           =   495
         End
      End
      Begin VB.Frame FEGenerar 
         Caption         =   "Criterios para Generar"
         Height          =   1455
         Left            =   240
         TabIndex        =   29
         Top             =   360
         Width           =   2415
         Begin VB.ComboBox cboMes 
            Height          =   315
            ItemData        =   "frmPagosAdelantados.frx":035E
            Left            =   480
            List            =   "frmPagosAdelantados.frx":0360
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Top             =   840
            Width           =   1815
         End
         Begin MSMask.MaskEdBox txtAnio 
            Height          =   375
            Left            =   480
            TabIndex        =   30
            Top             =   360
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   661
            _Version        =   393216
            MaxLength       =   4
            Mask            =   "####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label9 
            Caption         =   "Mes"
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   960
            Width           =   375
         End
         Begin VB.Label Label7 
            Caption         =   "Año"
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   480
            Width           =   375
         End
      End
      Begin VB.CheckBox chkTodasAgencias 
         Caption         =   "&Todas las Agencias"
         Enabled         =   0   'False
         Height          =   255
         Left            =   -73560
         TabIndex        =   28
         Top             =   2280
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CommandButton cmdCancelarPagoAdelantado 
         Caption         =   "Ca&ncelar"
         Enabled         =   0   'False
         Height          =   350
         Left            =   -66360
         TabIndex        =   16
         Top             =   2160
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdAceptarPagoAdelantado 
         Caption         =   "A&ceptar"
         Enabled         =   0   'False
         Height          =   350
         Left            =   -66360
         TabIndex        =   15
         Top             =   1800
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtDesCtaCon 
         Enabled         =   0   'False
         Height          =   350
         Left            =   -71880
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   780
         Width           =   5415
      End
      Begin VB.CommandButton cmdEditarPagoAdelantado 
         Caption         =   "E&ditar"
         Enabled         =   0   'False
         Height          =   350
         Left            =   -66360
         TabIndex        =   13
         Top             =   1080
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdNuevoPagoAdelantado 
         Caption         =   "&Nuevo"
         Enabled         =   0   'False
         Height          =   350
         Left            =   -66360
         TabIndex        =   12
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtNroMes 
         Enabled         =   0   'False
         Height          =   350
         Left            =   -70920
         MaxLength       =   9
         TabIndex        =   11
         Top             =   1500
         Width           =   1215
      End
      Begin VB.TextBox txtDesPagAde 
         Enabled         =   0   'False
         Height          =   345
         Left            =   -73560
         MaxLength       =   300
         TabIndex        =   10
         Top             =   1140
         Width           =   7095
      End
      Begin VB.ComboBox cboRubro 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -73560
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1860
         Width           =   1695
      End
      Begin VB.TextBox txtImporte 
         Enabled         =   0   'False
         Height          =   345
         Left            =   -68160
         TabIndex        =   7
         Top             =   1860
         Width           =   1695
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
         Enabled         =   0   'False
         Height          =   350
         Left            =   9840
         TabIndex        =   6
         Top             =   1080
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdExaminar 
         Caption         =   "&Examinar"
         Enabled         =   0   'False
         Height          =   350
         Left            =   9840
         TabIndex        =   5
         Top             =   1860
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdGenerar 
         Caption         =   "&Generar"
         Enabled         =   0   'False
         Height          =   350
         Left            =   9840
         TabIndex        =   4
         Top             =   4380
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdEliminarPagoAdelantado 
         Caption         =   "E&liminar"
         Enabled         =   0   'False
         Height          =   350
         Left            =   -66360
         TabIndex        =   3
         Top             =   1440
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Frame FrAgencia 
         Caption         =   "Agencia"
         Height          =   2415
         Left            =   -73560
         TabIndex        =   1
         Top             =   2520
         Visible         =   0   'False
         Width           =   7095
         Begin MSComctlLib.ListView lvAgencia 
            Height          =   2100
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Visible         =   0   'False
            Width           =   6795
            _ExtentX        =   11986
            _ExtentY        =   3704
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            Enabled         =   0   'False
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Cuenta"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Descripción"
               Object.Width           =   6174
            EndProperty
         End
      End
      Begin Sicmact.FlexEdit FEBuscar 
         Height          =   2895
         Left            =   120
         TabIndex        =   9
         Top             =   1860
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   5106
         Cols0           =   9
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "Nro-Pago Adelantado-Fecha de Vencimiento-Importe-Saldo-IdPagAdeCab-nNumMes-cCtaContCod-cAgeCod"
         EncabezadosAnchos=   "500-4800-1700-1200-1200-3-3-1200-3"
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
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-C-R-R-R-R-L-C"
         FormatosEdit    =   "0-0-0-2-2-3-3-0-0"
         TextArray0      =   "Nro"
         SelectionMode   =   1
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   495
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin Sicmact.FlexEdit FEAmortizacion 
         Height          =   3975
         Left            =   -74880
         TabIndex        =   17
         Top             =   780
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   7011
         Cols0           =   8
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "Nro-Importe-Amo. Acu.-Mes Amo.-Amo. Mes-Tot. Amo.-Saldo-Mov"
         EncabezadosAnchos=   "500-1200-1200-1200-1200-1200-1200-2500"
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
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-R-R-C-R-R-R-L"
         FormatosEdit    =   "0-2-2-0-2-2-2-0"
         TextArray0      =   "Nro"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   495
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin Sicmact.TxtBuscar txtBuscarCtaCon 
         Height          =   345
         Left            =   -73560
         TabIndex        =   18
         Top             =   780
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   609
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
         Enabled         =   0   'False
         Enabled         =   0   'False
         TipoBusqueda    =   5
         sTitulo         =   ""
         EnabledText     =   0   'False
      End
      Begin MSMask.MaskEdBox txtFecIni 
         Height          =   345
         Left            =   -73560
         TabIndex        =   19
         Top             =   1500
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   609
         _Version        =   393216
         Enabled         =   0   'False
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
      Begin MSMask.MaskEdBox txtFecVen 
         Height          =   345
         Left            =   -68160
         TabIndex        =   20
         Top             =   1500
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   609
         _Version        =   393216
         Enabled         =   0   'False
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
      Begin VB.Label Label6 
         Caption         =   "Importe"
         Height          =   255
         Left            =   -68880
         TabIndex        =   27
         Top             =   1980
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Nro Meses"
         Height          =   255
         Left            =   -71760
         TabIndex        =   26
         Top             =   1620
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Vencimiento"
         Height          =   255
         Left            =   -69600
         TabIndex        =   25
         Top             =   1620
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha Inicio"
         Height          =   255
         Left            =   -74880
         TabIndex        =   24
         Top             =   1620
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Cuenta Contable"
         Height          =   255
         Left            =   -74880
         TabIndex        =   23
         Top             =   900
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Descripción"
         Height          =   255
         Left            =   -74880
         TabIndex        =   22
         Top             =   1260
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Rubro"
         Height          =   255
         Left            =   -74880
         TabIndex        =   21
         Top             =   1980
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmPagosAdelantados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************************
'***Nombre:         frmPagosAdelantados
'***Descripción:    Formulario que permite el registro del
'                   pago adelantado.
'***Creación:       ELRO el 20111109 según Acta 323-2011/TI-D
'************************************************************
Option Explicit

Private Enum CajasTextosPagosAdelantados
gValorDefectoCajasTextosPagos = 0
gTxtDesPagAde = 1
gTxtBuscarCtaCon = 2
gTxtDesCtaCon = 3
gTxtFecIni = 4
gTxtNroMes = 5
gTxtFecVen = 6
gTxtImporte = 7
End Enum

Private Enum BotonesPagosAdelantados
gValorDefectoBotonesPagos = 0
gcmdNuevoPagoAdelantado = 1
gcmdEditarPagoAdelantado = 2
gcmdAceptarPagoAdelantado = 3
gcmdCancelarPagoAdelantado = 4
gcmdEliminarPagoAdelantado = 5
End Enum

Private Enum ComboPagosAdelantados
gValorDefectoComboPagosAdelantados = 0
gcboMes = 1
gcboRubro = 2
gcboRubro2 = 3
End Enum

Private Enum BotonesBuscarPagosAdelantados
gValorDefectoBotonesBuscarPagosAdelantados = 0
gcmdBuscar = 1
gcmdExaminar = 2
gcmdGenerar = 3
End Enum

Private Enum Accion
gValorDefectoAccion = 0
gNuevoRegistro = 1
gEditarRegistro = 2
gEliminarRegistro = 3
End Enum

Private fnIdPagAdeCab As Integer
Private fnAccionPagoAdelantado As Accion
Private fnAccionAmortizacion As Accion
Private fsCtaConHaber As String
Private fsCtaConDebe  As String
Private fsOpeCod  As String
Private ldFecCie As Date
Private oPista As COMManejador.Pista


Private Sub habilitarCajasTextosPagosAdelantados(ByVal pbValor As Boolean, Optional ByVal pnControl As CajasTextosPagosAdelantados = gValorDefectoCajasTextosPagos)
    If pnControl = gValorDefectoCajasTextosPagos Then
        txtDesPagAde.Enabled = pbValor
        txtBuscarCtaCon.Enabled = pbValor
        txtDesCtaCon.Enabled = pbValor
        txtFecIni.Enabled = pbValor
        txtNroMes.Enabled = pbValor
        txtFecVen.Enabled = pbValor
        txtImporte.Enabled = pbValor
    ElseIf pnControl = gTxtDesPagAde Then
        txtDesPagAde.Enabled = pbValor
    ElseIf pnControl = gTxtBuscarCtaCon Then
        txtBuscarCtaCon.Enabled = pbValor
    ElseIf pnControl = gTxtDesCtaCon Then
        txtDesCtaCon.Enabled = pbValor
    ElseIf pnControl = gTxtFecIni Then
        txtFecIni.Enabled = pbValor
    ElseIf pnControl = gTxtNroMes Then
        txtNroMes.Enabled = pbValor
    ElseIf pnControl = gTxtFecVen Then
        txtFecVen.Enabled = pbValor
    ElseIf pnControl = gTxtImporte Then
        txtImporte.Enabled = pbValor
    End If

End Sub

Private Sub habilitarBotonesPagosAdelantados(ByVal pbValor As Boolean, Optional ByVal pnControl As BotonesPagosAdelantados = gValorDefectoBotonesPagos)
    If pnControl = gValorDefectoBotonesPagos Then
        cmdNuevoPagoAdelantado.Enabled = pbValor
        cmdEditarPagoAdelantado.Enabled = pbValor
        cmdAceptarPagoAdelantado.Enabled = pbValor
        cmdCancelarPagoAdelantado.Enabled = pbValor
        cmdEliminarPagoAdelantado.Enabled = pbValor
            
    ElseIf pnControl = gcmdNuevoPagoAdelantado Then
        cmdNuevoPagoAdelantado.Enabled = pbValor
    ElseIf pnControl = gcmdEditarPagoAdelantado Then
        cmdEditarPagoAdelantado.Enabled = pbValor
    ElseIf pnControl = gcmdAceptarPagoAdelantado Then
        cmdAceptarPagoAdelantado.Enabled = pbValor
    ElseIf pnControl = gcmdCancelarPagoAdelantado Then
        cmdCancelarPagoAdelantado.Enabled = pbValor
    ElseIf pnControl = gcmdEliminarPagoAdelantado Then
        cmdEliminarPagoAdelantado.Enabled = pbValor
    End If
End Sub
Private Sub visualizarBotonesPagosAdelantados(ByVal pbValor As Boolean, Optional ByVal pnControl As BotonesPagosAdelantados = gValorDefectoBotonesPagos)
    If pnControl = gValorDefectoBotonesPagos Then
        cmdNuevoPagoAdelantado.Visible = pbValor
        cmdEditarPagoAdelantado.Visible = pbValor
        cmdAceptarPagoAdelantado.Visible = pbValor
        cmdCancelarPagoAdelantado.Visible = pbValor
        cmdEliminarPagoAdelantado.Visible = pbValor
         
    ElseIf pnControl = gcmdNuevoPagoAdelantado Then
        cmdNuevoPagoAdelantado.Visible = pbValor
    ElseIf pnControl = gcmdEditarPagoAdelantado Then
        cmdEditarPagoAdelantado.Visible = pbValor
    ElseIf pnControl = gcmdAceptarPagoAdelantado Then
        cmdAceptarPagoAdelantado.Visible = pbValor
    ElseIf pnControl = gcmdCancelarPagoAdelantado Then
        cmdCancelarPagoAdelantado.Visible = pbValor
    ElseIf pnControl = gcmdEliminarPagoAdelantado Then
        cmdEliminarPagoAdelantado.Visible = pbValor
    End If
End Sub

Private Sub habilitarFEAmortizacion(ByVal pbValor As Boolean)
FEAmortizacion.lbEditarFlex = pbValor
End Sub

Private Sub habilitarFEBuscar(ByVal pbValor As Boolean)
FEBuscar.lbEditarFlex = pbValor
End Sub

Private Sub limpiarCajasTextosPagosAdelantados()
    txtDesPagAde = ""
    txtBuscarCtaCon = ""
    txtDesCtaCon = ""
    txtNroMes = "1"
    txtFecVen = "__/__/____"
    txtImporte = "0.00"
End Sub

Function validarCajasTextosPagosAdelantados() As Boolean
    validarCajasTextosPagosAdelantados = True
    
    If txtDesPagAde = "" Then
        MsgBox "No ingreso la Descripción del Pago Adelantado", vbInformation, "Aviso"
        validarCajasTextosPagosAdelantados = False
        txtDesPagAde.SetFocus
        Exit Function
    End If
    
    If txtBuscarCtaCon = "" Then
        MsgBox "No ingreso la cuenta contable del Pago Adelantado", vbInformation, "Aviso"
        validarCajasTextosPagosAdelantados = False
        txtBuscarCtaCon.SetFocus
        Exit Function
    End If
    
    If txtNroMes = "" Then
        MsgBox "No ingreso el Tiempo o Nro Meses del Pago Adelantado", vbInformation, "Aviso"
        validarCajasTextosPagosAdelantados = False
        txtNroMes.SetFocus
        Exit Function
    End If

    If Trim(txtImporte) = "" Then
        MsgBox "No ingreso el Importe del Pago Adelantado", vbInformation, "Aviso"
        validarCajasTextosPagosAdelantados = False
        txtImporte.SetFocus
        Exit Function
    End If
    
    If Trim(Right(cboRubro.Text, 3)) = "" Then
        MsgBox "No selecciono el Rubro del Pago Adelantado", vbInformation, "Aviso"
        validarCajasTextosPagosAdelantados = False
        cboRubro.SetFocus
        Exit Function
    End If
      
    If Trim(txtImporte) <> "" Then
        If Not IsNumeric(txtImporte) Then
            MsgBox "Debe ingresar un número", vbInformation, "Aviso"
            validarCajasTextosPagosAdelantados = False
            txtImporte.SetFocus
            Exit Function
        Else
            If CDec(txtImporte) = 0# Then
                MsgBox "El importe debe ser mayor a 0", vbInformation, "Aviso"
                validarCajasTextosPagosAdelantados = False
                txtImporte.SetFocus
            Exit Function
            End If
        End If
        
    End If
    
    If txtFecIni.Text = "__/__/____" Then
        MsgBox "Debe ingresar la fecha de inicio", vbInformation, "Aviso"
        validarCajasTextosPagosAdelantados = False
        txtFecIni.SetFocus
        Exit Function
    End If
    
    If Trim(txtNroMes) = "" Then
        MsgBox "El número de mes no debe estar vacío.", vbInformation, "Titulo"
        validarCajasTextosPagosAdelantados = False
        txtNroMes.SetFocus
        Exit Function
    End If
        
    If Not IsNumeric(txtNroMes) Then
        MsgBox "Ingrese Número.", vbInformation, "Aviso"
        validarCajasTextosPagosAdelantados = False
        txtNroMes.SetFocus
        Exit Function
    End If
        
    If CInt(txtNroMes) <= 0 Then
        MsgBox "El número de mes no debe ser menor o igual cero.", vbInformation, "Aviso"
        validarCajasTextosPagosAdelantados = False
        txtNroMes.SetFocus
        Exit Function
    End If
    
End Function

Private Sub CargarCtaCont()
Dim oDCtaCont As New DCtaCont
txtBuscarCtaCon.rs = oDCtaCont.CargaCtaCont("", "CtaCont")
txtBuscarCtaCon.TipoBusqueda = BuscaGrid
txtBuscarCtaCon.lbUltimaInstancia = False
Set oDCtaCont = Nothing
End Sub

Private Sub CargarRubros()
    Dim oDDocumento As DDocumento
    Set oDDocumento = New DDocumento
    Dim rsRubro As ADODB.Recordset
    Set rsRubro = New ADODB.Recordset
    
    Set rsRubro = oDDocumento.listarRubrosPagosAdelantados
    
    cboRubro.Clear
    While Not rsRubro.EOF
        Me.cboRubro.AddItem rsRubro.Fields(1) & Space(50) & rsRubro.Fields(0)
        rsRubro.MoveNext
    Wend
    
    cboRubro.ListIndex = -1
    Set rsRubro = Nothing
    Set oDDocumento = Nothing
End Sub

Private Sub CargarRubros2()
    Dim oDDocumento As DDocumento
    Set oDDocumento = New DDocumento
    Dim rsRubro2 As ADODB.Recordset
    Set rsRubro2 = New ADODB.Recordset
    
    Set rsRubro2 = oDDocumento.listarRubrosPagosAdelantados
    
    cboRubro2.Clear
    While Not rsRubro2.EOF
        Me.cboRubro2.AddItem rsRubro2.Fields(1) & Space(50) & rsRubro2.Fields(0)
        rsRubro2.MoveNext
    Wend
    
    cboRubro2.ListIndex = -1
    Set rsRubro2 = Nothing
    Set oDDocumento = Nothing
End Sub

Private Sub cargarMeses()
    Dim oDGeneral As DGeneral
    Set oDGeneral = New DGeneral
    Dim rsMeses As ADODB.Recordset
    Set rsMeses = New ADODB.Recordset
    
    Set rsMeses = oDGeneral.GetConstante(1010)
    
    cboMes.Clear
    While Not rsMeses.EOF
        cboMes.AddItem rsMeses.Fields(0) & Space(50) & rsMeses.Fields(1)
        rsMeses.MoveNext
    Wend
    
    Set rsMeses = Nothing
    Set oDGeneral = Nothing
    
End Sub

Private Sub habilitarComboPagosAdelantados(ByVal pbValor As Boolean, Optional ByVal pnControl As ComboPagosAdelantados = gValorDefectoComboPagosAdelantados)
    If pnControl = gValorDefectoComboPagosAdelantados Then
        cboMes.Enabled = pbValor
        cboRubro.Enabled = pbValor
        cboRubro2.Enabled = pbValor
        
    ElseIf pnControl = gcboMes Then
        cboMes.Enabled = pbValor
    ElseIf pnControl = gcboRubro Then
        cboRubro.Enabled = pbValor
    ElseIf pnControl = gcboRubro2 Then
        cboRubro2.Enabled = pbValor
    End If
End Sub

Private Sub visualizarComboPagosAdelantados(ByVal pbValor As Boolean, Optional ByVal pnControl As ComboPagosAdelantados = gValorDefectoComboPagosAdelantados)
    If pnControl = gValorDefectoComboPagosAdelantados Then
        cboMes.Visible = pbValor
        cboRubro.Visible = pbValor
        cboRubro2.Visible = pbValor
        
    ElseIf pnControl = gcboMes Then
        cboMes.Visible = pbValor
    ElseIf pnControl = gcboRubro Then
        cboRubro.Visible = pbValor
    ElseIf pnControl = gcboRubro2 Then
        cboRubro2.Visible = pbValor
    End If
End Sub

Private Sub habilitarBotonesBuscarPagosAdelantados(ByVal pbValor As Boolean, Optional ByVal pnControl As BotonesBuscarPagosAdelantados = gValorDefectoBotonesBuscarPagosAdelantados)
    If pnControl = gValorDefectoBotonesBuscarPagosAdelantados Then
        cmdBuscar.Enabled = pbValor
        cmdExaminar.Enabled = pbValor
        cmdGenerar.Enabled = pbValor
        
    ElseIf pnControl = gcmdBuscar Then
        cmdBuscar.Enabled = pbValor
    ElseIf pnControl = gcmdExaminar Then
        cmdExaminar.Enabled = pbValor
    ElseIf pnControl = gcmdGenerar Then
        cmdGenerar.Enabled = pbValor
    End If

End Sub

Private Sub visualizarBotonesBuscarPagosAdelantados(ByVal pbValor As Boolean, Optional ByVal pnControl As BotonesBuscarPagosAdelantados = gValorDefectoBotonesBuscarPagosAdelantados)
    If pnControl = gValorDefectoBotonesBuscarPagosAdelantados Then
        cmdBuscar.Visible = pbValor
        cmdExaminar.Visible = pbValor
        cmdGenerar.Visible = pbValor
        
    ElseIf pnControl = gcmdBuscar Then
        cmdBuscar.Visible = pbValor
    ElseIf pnControl = gcmdExaminar Then
        cmdExaminar.Visible = pbValor
    ElseIf pnControl = gcmdGenerar Then
        cmdGenerar.Visible = pbValor
    End If

End Sub

Private Sub cboRubro_Click()
 
    Dim oDDocumento As DDocumento
    Set oDDocumento = New DDocumento
    Dim rsRubroCtaConDebe As ADODB.Recordset
    Set rsRubroCtaConDebe = New ADODB.Recordset
    Dim oForm As New frmRubrosPagosAdelantados
  
    If Trim(Right(cboRubro.Text, 3)) = "" Then
        MsgBox "Elija un rubro...", vbInformation, "Aviso"
        If cboRubro.Enabled Then
            cboRubro.SetFocus
        Else
            cmdNuevoPagoAdelantado.SetFocus
        End If
        Exit Sub
    End If
    
    
    Set rsRubroCtaConDebe = oDDocumento.recuperarRubroPagosAdelantados(False, CInt(Trim(Right(cboRubro.Text, 3))))

    If Not rsRubroCtaConDebe.BOF And Not rsRubroCtaConDebe.EOF Then
        Dim nPosicion As Integer
        Dim sCadena As String
        
        sCadena = rsRubroCtaConDebe!cForCtaCon
        nPosicion = InStr(1, sCadena, "AG")
        
        If nPosicion = 0 Then
            FrAgencia.Visible = False
            lvAgencia.Visible = False
            chkTodasAgencias.Visible = False
        Else
            FrAgencia.Visible = True
            lvAgencia.Visible = True
            chkTodasAgencias.Visible = True
            If cmdAceptarPagoAdelantado.Visible Then
                lvAgencia.Enabled = True
                chkTodasAgencias.Enabled = True
            Else
                lvAgencia.Enabled = False
                chkTodasAgencias.Enabled = False
            End If
            
        End If
    Else
        MsgBox "Debe ingregsar la fórmula del Rubro " & Trim(Left(cboRubro.Text, 20))
        oForm.Show 1
        Call CargarRubros2
        Call CargarRubros
    End If
    
    Set rsRubroCtaConDebe = Nothing
    Set oDDocumento = Nothing

End Sub

Private Sub CargarAgencias()
    Dim oDDocumento As DDocumento
    Set oDDocumento = New DDocumento
    Dim rsAgencia As ADODB.Recordset
    Set rsAgencia = New ADODB.Recordset
    Dim lvItem As ListItem
    
    Set rsAgencia = oDDocumento.listarAgenciasPagosAdelantados
    
    lvAgencia.ListItems.Clear
    Do While Not rsAgencia.EOF
        Set lvItem = lvAgencia.ListItems.Add
        lvItem.Text = rsAgencia.Fields(0)
        lvItem.SubItems(1) = rsAgencia.Fields(1)
        lvItem.Checked = False
        rsAgencia.MoveNext
    Loop

    Set rsAgencia = Nothing
    Set oDDocumento = Nothing
End Sub

Private Function optenerCtaConDebe(ByVal psMoneda As String) As String
    Dim oDDocumento As DDocumento
    Set oDDocumento = New DDocumento
    Dim rsRubroCtaConDebe As ADODB.Recordset
    Set rsRubroCtaConDebe = New ADODB.Recordset
  
    optenerCtaConDebe = ""

    Set rsRubroCtaConDebe = oDDocumento.recuperarRubroPagosAdelantados(False, CInt(Trim(Right(cboRubro2.Text, 3))))
    
    If Not rsRubroCtaConDebe.BOF And Not rsRubroCtaConDebe.EOF Then
        Dim nPosicion, nPosicion2 As Integer
        Dim sCadena, sCadena2 As String
        
        sCadena = rsRubroCtaConDebe!cForCtaCon
        nPosicion = InStr(1, sCadena, "M")
        nPosicion2 = InStr(1, sCadena, "AG")
        
        If nPosicion > 0 Then
            If nPosicion2 > 0 Then
                sCadena2 = Left(sCadena, nPosicion - 1) & psMoneda & Mid(sCadena, nPosicion + 1, nPosicion2 - nPosicion - 1)
            Else
                sCadena2 = Left(sCadena, nPosicion - 1) & psMoneda & Mid(sCadena, nPosicion + 1)
            End If
            
        Else
            If nPosicion2 > 0 Then
                sCadena2 = Left(sCadena, nPosicion2 - 1)
            Else
                sCadena2 = sCadena
            End If
            
        End If
        
        optenerCtaConDebe = sCadena2
        
        Set rsRubroCtaConDebe = Nothing
        Set oDDocumento = Nothing
    Else
         MsgBox "Para el Rubro " & Trim(Left(cboRubro2.Text, 20)) & "no existe fórmula cta. cont. asignada, avise a TI", vbCritical, "Aviso"
        Exit Function
    End If

End Function

Private Function ValidarDatosBusqueda() As Boolean

    ValidarDatosBusqueda = True
    
    If Trim(Right(cboRubro2.Text, 3)) = "" Then
        MsgBox "Elija un rubro...", vbInformation, "Aviso"
        ValidarDatosBusqueda = False
        If cboRubro2.Enabled Then
            cboRubro2.SetFocus
        End If
        Exit Function
    End If
        
End Function

Private Function ValidarDatosGeneracion() As Boolean

ValidarDatosGeneracion = True

    If Trim(txtAnio) = "" Then
        MsgBox "Ingrese el año...", vbInformation, "Aviso"
        ValidarDatosGeneracion = False
        txtAnio.SetFocus
        Exit Function
    End If
    
    If Trim(Right(cboMes.Text, 3)) = "" Then
        MsgBox "Elija el mes...", vbInformation, "Aviso"
        ValidarDatosGeneracion = False
        cboMes.SetFocus
        Exit Function
    End If
    
End Function

Private Sub cargarAmoritzacionesPagoAdenatado()
    Dim oDDocumento As DDocumento
    Set oDDocumento = New DDocumento
    Dim rsAmortizaciones As ADODB.Recordset
    Set rsAmortizaciones = New ADODB.Recordset
    Dim i As Integer

    i = 1
    
    Set rsAmortizaciones = oDDocumento.listarAmortizacionesPagoAdelantado(fnIdPagAdeCab)
    
    Call LimpiaFlex(FEAmortizacion)

    If Not rsAmortizaciones.BOF And Not rsAmortizaciones.EOF Then
    
        Call habilitarFEAmortizacion(True)
    
        Do While Not rsAmortizaciones.EOF
        
            FEAmortizacion.AdicionaFila
            FEAmortizacion.TextMatrix(i, 1) = rsAmortizaciones!nImporte
            FEAmortizacion.TextMatrix(i, 2) = rsAmortizaciones!nAmoAcu
            FEAmortizacion.TextMatrix(i, 3) = rsAmortizaciones!cMES
            FEAmortizacion.TextMatrix(i, 4) = rsAmortizaciones!nAmortizacion
            FEAmortizacion.TextMatrix(i, 5) = rsAmortizaciones!nTotAmt
            FEAmortizacion.TextMatrix(i, 6) = rsAmortizaciones!nSaldo
            FEAmortizacion.TextMatrix(i, 7) = rsAmortizaciones!cMov
            i = i + 1
            rsAmortizaciones.MoveNext
        
        Loop
    
    End If
    
End Sub



Private Sub cboRubro_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyN And Shift = 2 Then
        Dim oForm As New frmRubrosPagosAdelantados
        oForm.Show 1
        Call CargarRubros2
        Call CargarRubros
    End If
End Sub




Private Sub chkCancelado_Click()
Call cmdbuscar_Click
End Sub

Private Sub chkTodasAgencias_Click()

Dim k As Integer

'If lvAgencia.Visible Then
'    lvAgencia.SetFocus
'End If

If chkTodasAgencias Then
    For k = 1 To CInt(lvAgencia.ListItems.Count)
        DoEvents
        lvAgencia.ListItems(k).Selected = True
        lvAgencia.ListItems(k).Checked = True
        lvAgencia.SelectedItem.EnsureVisible
    Next k
 Else
    For k = 1 To CInt(lvAgencia.ListItems.Count)
            DoEvents
            lvAgencia.ListItems(k).Selected = True
            lvAgencia.ListItems(k).Checked = False
            lvAgencia.SelectedItem.EnsureVisible
    Next k
 End If
 
End Sub

Private Sub cmdAceptarPagoAdelantado_Click()
    Dim k, l As Integer
    Dim sCodAge As String
   

    If validarCajasTextosPagosAdelantados = False Then
        Exit Sub
    End If
    
    l = CInt(lvAgencia.ListItems.Count)
    
    If lvAgencia.Visible Then
        For k = 1 To l
            If lvAgencia.ListItems(k).Checked Then
                If Trim(sCodAge) = "" Then
                    sCodAge = lvAgencia.ListItems(k).Text
                Else
                    sCodAge = sCodAge & "," & Trim(lvAgencia.ListItems(k).Text)
                End If
            End If
            If k = l Then Exit For
         Next k

    
        If Trim(sCodAge) = "" Then
            MsgBox "Debe seleccione al menos una Agencia", vbInformation, "Aviso"
            Exit Sub
        End If
    End If
    
    If verificarUltimoNivelCta(txtBuscarCtaCon) = False Then
        MsgBox "La cuenta contable " & txtBuscarCtaCon & " no es de ultimo nivel", vbInformation, "Aviso"
        txtBuscarCtaCon.SetFocus
        Exit Sub
    End If
    
    Dim lsMovNro As String
    Dim bConfirmar As Boolean
    Dim oDDocumento As DDocumento
    Set oDDocumento = New DDocumento
    Dim oNContFunciones As NContFunciones
    Set oNContFunciones = New NContFunciones
    
    lsMovNro = oNContFunciones.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    
    If fnAccionPagoAdelantado = gNuevoRegistro Then
        bConfirmar = oDDocumento.registrarPagoAdelantadoCab(txtDesPagAde, _
                                                            txtBuscarCtaCon, _
                                                            CInt(Trim(Right(cboRubro.Text, 3))), _
                                                            CInt(txtNroMes), _
                                                            Format(txtFecIni, "mm/dd/yyyy"), _
                                                            Format(txtFecVen, "mm/dd/yyyy"), _
                                                            CCur(txtImporte), _
                                                            sCodAge, _
                                                            lsMovNro)
                        
         If bConfirmar Then
            MsgBox "Se registraron correctamente los datos del Pago Adelantado", vbInformation, "Aviso"
            Call habilitarCajasTextosPagosAdelantados(False)
            Call visualizarBotonesPagosAdelantados(False, gcmdAceptarPagoAdelantado)
            Call visualizarBotonesPagosAdelantados(False, gcmdCancelarPagoAdelantado)
            Call visualizarBotonesPagosAdelantados(True, gcmdNuevoPagoAdelantado)
            Call habilitarComboPagosAdelantados(False, gcboRubro)
            
            If chkTodasAgencias.Visible Then
                chkTodasAgencias.Enabled = False
            End If
            
            If lvAgencia.Visible Then
                lvAgencia.Enabled = False
            End If
            

         Else
            MsgBox "No se registraron los datos del Pago Adelantado", vbInformation, "Aviso"
            Exit Sub
         End If
         
         
         
    End If
    
    If fnAccionPagoAdelantado = gEditarRegistro Then
       bConfirmar = oDDocumento.actualizarPagoAdelantadoCab(fnIdPagAdeCab, _
                                                            txtDesPagAde, _
                                                            txtBuscarCtaCon, _
                                                            Trim(Right(cboRubro.Text, 3)), _
                                                            CInt(txtNroMes), _
                                                            Format(txtFecIni, "mm/dd/yyyy"), _
                                                            Format(txtFecVen, "mm/dd/yyyy"), _
                                                            CCur(txtImporte), _
                                                            sCodAge, _
                                                            lsMovNro)
        If bConfirmar Then
            MsgBox "Se actualizaron correctamente los datos del Pago Adelantado", vbInformation, "Aviso"
            Call habilitarCajasTextosPagosAdelantados(False)
            Call visualizarBotonesPagosAdelantados(False, gcmdAceptarPagoAdelantado)
            Call visualizarBotonesPagosAdelantados(False, gcmdCancelarPagoAdelantado)
            Call visualizarBotonesPagosAdelantados(True, gcmdEditarPagoAdelantado)
            Call visualizarBotonesPagosAdelantados(True, gcmdEliminarPagoAdelantado)
            Call habilitarComboPagosAdelantados(False, gcboRubro)
            
            Call cmdbuscar_Click
            
            chkCancelado.Enabled = True
            Call habilitarComboPagosAdelantados(True, gcboRubro2)
            Call habilitarBotonesBuscarPagosAdelantados(True, gcmdBuscar)
            Call habilitarBotonesBuscarPagosAdelantados(True, gcmdExaminar)
            Call habilitarBotonesBuscarPagosAdelantados(True, gcmdGenerar)
            
            If chkTodasAgencias.Visible Then
                chkTodasAgencias.Enabled = False
            End If
            
            If lvAgencia.Visible Then
                lvAgencia.Enabled = False
            End If

        Else
            MsgBox "No se actualizaron los datos del Pago Adelantado", vbInformation, "Aviso"
            Exit Sub
        End If
        
    End If
    
    k = 0
    sCodAge = ""
    lsMovNro = ""
    bConfirmar = False
    Set oDDocumento = Nothing
    Set oNContFunciones = Nothing
End Sub


Private Sub cmdbuscar_Click()

If ValidarDatosBusqueda = False Then
    Exit Sub
End If

Dim oDDocumento As DDocumento
Set oDDocumento = New DDocumento
Dim rsLista As ADODB.Recordset
Set rsLista = New ADODB.Recordset
Dim i As Integer

i = 1

Set rsLista = oDDocumento.buscarPagosAdelantados(CInt(Trim(Right(cboRubro2.Text, 3))), _
                                                 IIf(OptMoneda.Item(0).value, "1", "2"), _
                                                 IIf(chkCancelado, True, False))

Call LimpiaFlex(FEBuscar)

If Not rsLista.BOF And Not rsLista.EOF Then


    Call habilitarFEBuscar(True)
    
    Do While Not rsLista.EOF
        
        FEBuscar.AdicionaFila
        FEBuscar.TextMatrix(i, 1) = rsLista!cDesPagAde
        FEBuscar.TextMatrix(i, 2) = rsLista!dFecVen
        FEBuscar.TextMatrix(i, 3) = rsLista!nImporte
        FEBuscar.TextMatrix(i, 4) = rsLista!nSaldo
        FEBuscar.TextMatrix(i, 5) = rsLista!IdPagAdeCab
        FEBuscar.TextMatrix(i, 6) = rsLista!nNumMes
        FEBuscar.TextMatrix(i, 7) = rsLista!cCtaContCod
        i = i + 1
        rsLista.MoveNext
    Loop
    
    Call habilitarFEBuscar(False)
    Call habilitarBotonesBuscarPagosAdelantados(True, gcmdExaminar)
    If chkCancelado Then
        Call habilitarBotonesBuscarPagosAdelantados(False, gcmdGenerar)
    Else
        Call habilitarBotonesBuscarPagosAdelantados(True, gcmdGenerar)
    End If
    Call visualizarBotonesBuscarPagosAdelantados(True, gcmdExaminar)
    Call visualizarBotonesBuscarPagosAdelantados(True, gcmdGenerar)
Else

    MsgBox "No existe registro de Pagos Adelantados en el Rubro " & Left(Me.cboRubro2.Text, 20), vbInformation, "Aviso"
    
End If

Set rsLista = Nothing
Set oDDocumento = Nothing

End Sub

Private Sub cmdCancelarPagoAdelantado_Click()
Dim i As Integer

If fnAccionPagoAdelantado = gNuevoRegistro Then
    Call limpiarCajasTextosPagosAdelantados
    Call habilitarCajasTextosPagosAdelantados(False)
    Call visualizarBotonesPagosAdelantados(False, gcmdAceptarPagoAdelantado)
    Call visualizarBotonesPagosAdelantados(False, gcmdCancelarPagoAdelantado)
    Call visualizarBotonesPagosAdelantados(True, gcmdNuevoPagoAdelantado)
    cboRubro.Enabled = False
    cboRubro.ListIndex = -1
    
    If chkTodasAgencias.Visible Then
        chkTodasAgencias.Enabled = False
        chkTodasAgencias.value = False
        chkTodasAgencias.Visible = False
    End If
    
    If lvAgencia.Visible Then
        lvAgencia.Enabled = False
        lvAgencia.Visible = False
    End If
    
    For i = 1 To CInt(lvAgencia.ListItems.Count)
        lvAgencia.ListItems(i).Checked = False
    Next i
    
End If

If fnAccionPagoAdelantado = gEditarRegistro Then

    Call habilitarCajasTextosPagosAdelantados(False)
    Call visualizarBotonesPagosAdelantados(False, gcmdAceptarPagoAdelantado)
    Call visualizarBotonesPagosAdelantados(False, gcmdCancelarPagoAdelantado)
    Call visualizarBotonesPagosAdelantados(False, gcmdEditarPagoAdelantado)
    Call visualizarBotonesPagosAdelantados(False, gcmdEliminarPagoAdelantado)
    Call visualizarBotonesPagosAdelantados(True, gcmdNuevoPagoAdelantado)
    cboRubro.Enabled = False
    
    chkCancelado.Enabled = True
    Call habilitarComboPagosAdelantados(True, gcboRubro2)
    Call habilitarBotonesBuscarPagosAdelantados(True, gcmdBuscar)
    Call habilitarBotonesBuscarPagosAdelantados(True, gcmdExaminar)
    Call habilitarBotonesBuscarPagosAdelantados(True, gcmdGenerar)
     
    SSTabPagoAdelantado.Tab = 0
    
    If chkTodasAgencias.Visible Then
        chkTodasAgencias.Enabled = False
    End If
    
    If lvAgencia.Visible Then
        lvAgencia.Enabled = False
    End If
    
End If

End Sub

Private Sub cmdEditarPagoAdelantado_Click()

    If CCur(FEBuscar.TextMatrix(FEBuscar.Row, 4)) = 0 Then
        MsgBox "Este Pago Adelantado tiene Saldo 0, no se puede editar"
        Exit Sub
    End If
    
    Call habilitarCajasTextosPagosAdelantados(True)
    Call habilitarBotonesPagosAdelantados(True)
    Call habilitarCajasTextosPagosAdelantados(False, gTxtFecVen)
    If CCur(FEBuscar.TextMatrix(FEBuscar.Row, 3)) > CCur(FEBuscar.TextMatrix(FEBuscar.Row, 4)) Then
        Call habilitarCajasTextosPagosAdelantados(False, gTxtNroMes)
        Call habilitarCajasTextosPagosAdelantados(False, gTxtImporte)
        Call habilitarComboPagosAdelantados(False, gcboRubro)
        If chkTodasAgencias.Visible Then
            chkTodasAgencias.Enabled = False
        End If
        
        If lvAgencia.Visible Then
            lvAgencia.Enabled = False
        End If
    Else
        Call habilitarComboPagosAdelantados(True, gcboRubro)
        If chkTodasAgencias.Visible Then
            chkTodasAgencias.Enabled = True
        End If
        
        If lvAgencia.Visible Then
            lvAgencia.Enabled = True
        End If
    End If
    Call visualizarBotonesPagosAdelantados(True, gcmdAceptarPagoAdelantado)
    Call visualizarBotonesPagosAdelantados(True, gcmdCancelarPagoAdelantado)
    Call visualizarBotonesPagosAdelantados(False, gcmdEditarPagoAdelantado)
    Call visualizarBotonesPagosAdelantados(False, gcmdEliminarPagoAdelantado)
    
    

  
    
End Sub

Private Sub cmdEliminarPagoAdelantado_Click()
    Dim oDDocumento As DDocumento
    Set oDDocumento = New DDocumento
    Dim bConfirmar As Boolean
    
    If CCur(FEBuscar.TextMatrix(FEBuscar.Row, 3)) > CCur(FEBuscar.TextMatrix(FEBuscar.Row, 4)) Then
        MsgBox "Este Pago Adelantado tiene amortización(es), no se puede eliminar", vbInformation, "Aviso"
        Exit Sub
    End If
    
    bConfirmar = False
    
    If MsgBox("¿Esta seguro que desea eliminar ?", vbYesNo, "Aviso") = vbYes Then
            bConfirmar = oDDocumento.eliminarPagoAdelantadoCab(fnIdPagAdeCab)
            If bConfirmar Then
            Call limpiarCajasTextosPagosAdelantados
            End If
    End If
    
    Call habilitarCajasTextosPagosAdelantados(False)
    Call visualizarBotonesPagosAdelantados(False, gcmdAceptarPagoAdelantado)
    Call visualizarBotonesPagosAdelantados(False, gcmdCancelarPagoAdelantado)
    Call visualizarBotonesPagosAdelantados(False, gcmdEditarPagoAdelantado)
    Call visualizarBotonesPagosAdelantados(False, gcmdEliminarPagoAdelantado)
    Call visualizarBotonesPagosAdelantados(True, gcmdNuevoPagoAdelantado)
    cboRubro.Enabled = False
    
    chkCancelado.Enabled = True
    Call habilitarComboPagosAdelantados(True, gcboRubro2)
    Call habilitarBotonesBuscarPagosAdelantados(True, gcmdBuscar)
    Call habilitarBotonesBuscarPagosAdelantados(True, gcmdExaminar)
    Call habilitarBotonesBuscarPagosAdelantados(True, gcmdGenerar)
        
    SSTabPagoAdelantado.Tab = 0
    Call cmdbuscar_Click
End Sub

Private Sub cmdexaminar_Click()
Dim oDDocumento As DDocumento
Set oDDocumento = New DDocumento
Dim rsRecuperar As ADODB.Recordset
Set rsRecuperar = New ADODB.Recordset
Dim i, j, k, l As Integer
 Dim larrCodAge() As String

fnAccionPagoAdelantado = gEditarRegistro
chkCancelado.Enabled = False
Call habilitarComboPagosAdelantados(False, gcboRubro2)
Call habilitarBotonesBuscarPagosAdelantados(False, gcmdBuscar)
Call habilitarBotonesBuscarPagosAdelantados(False, gcmdGenerar)
Call CargarRubros
Call CargarAgencias

fnIdPagAdeCab = CInt(FEBuscar.TextMatrix(FEBuscar.Row, 5))
Set rsRecuperar = oDDocumento.recuperarPagoAdelantadoCab(fnIdPagAdeCab)

If Not rsRecuperar.BOF And Not rsRecuperar.EOF Then
    txtBuscarCtaCon = rsRecuperar!cCtaContCod
    Call txtBuscarCtaCon_EmiteDatos
    txtDesPagAde = CStr(rsRecuperar!cDesPagAde)
    txtFecIni = CStr(rsRecuperar!dFecIni)
    txtNroMes = CStr(rsRecuperar!nNumMes)
    txtFecVen = CStr(rsRecuperar!dFecVen)
    txtImporte = CStr(rsRecuperar!nImporte)
    
    For i = 0 To cboRubro.ListCount - 1
        cboRubro.ListIndex = i
        If CInt(Trim(Right(cboRubro.Text, 3))) = rsRecuperar!nRubro Then
            cboRubro.ListIndex = i
            Exit For
        End If
    Next i
    
    If Trim(rsRecuperar!cAgecod) <> "" Then
        
        larrCodAge = Split(rsRecuperar!cAgecod, ",")
      
        For k = 1 To CInt(lvAgencia.ListItems.Count)
        
            For l = 0 To CInt(UBound(larrCodAge))
            
                If Trim(lvAgencia.ListItems(k).Text) = larrCodAge(l) Then
                    lvAgencia.ListItems(k).Checked = True
                    Exit For
                End If
               
            Next l
        Next k

    End If
    
    Call cargarAmoritzacionesPagoAdenatado
    SSTabPagoAdelantado.Tab = 1
    Call visualizarBotonesPagosAdelantados(False, gcmdNuevoPagoAdelantado)
    Call habilitarBotonesPagosAdelantados(True, gcmdEditarPagoAdelantado)
    Call visualizarBotonesPagosAdelantados(True, gcmdEditarPagoAdelantado)
    Call habilitarBotonesPagosAdelantados(True, gcmdEliminarPagoAdelantado)
    Call visualizarBotonesPagosAdelantados(True, gcmdEliminarPagoAdelantado)
    Call habilitarBotonesPagosAdelantados(True, gcmdCancelarPagoAdelantado)
    Call visualizarBotonesPagosAdelantados(True, gcmdCancelarPagoAdelantado)
Else
    MsgBox "No se encontro los datos del Pago Adelantado", vbInformation, "Aviso"
End If

End Sub

Private Sub cmdGenerar_Click()
    
    If ValidarDatosBusqueda = False Then
        Exit Sub
    End If
    
    If ValidarDatosGeneracion = False Then
        Exit Sub
    End If
    
    If ldFecCie >= gdFecSis Then
        MsgBox "Esta operacion no se pude realizar en un mes cerrado", vbInformation, "Aviso"
        Exit Sub
    End If
    
    txtAnio.Enabled = False
    chkCancelado.Enabled = False
    OptMoneda.Item(0).Enabled = False
    OptMoneda.Item(1).Enabled = False
    Call habilitarComboPagosAdelantados(False, gcboMes)
    Call habilitarComboPagosAdelantados(False, gcboRubro2)
    Call habilitarBotonesBuscarPagosAdelantados(False, gcmdBuscar)
    Call habilitarBotonesBuscarPagosAdelantados(False, gcmdExaminar)
    Call habilitarBotonesBuscarPagosAdelantados(False, gcmdGenerar)
    Call visualizarBotonesPagosAdelantados(False, gcmdNuevoPagoAdelantado)
    
    Dim oDDocumento As DDocumento
    Set oDDocumento = New DDocumento
    Dim oNContFunciones As NContFunciones
    Set oNContFunciones = New NContFunciones
    Dim lsMovNro, lsDescripcionPagoAdelantado, sCadImp As String
    Dim lnConfirmarAsiento As Long
    Dim rsListaGenerarAsiento As ADODB.Recordset
    Set rsListaGenerarAsiento = New ADODB.Recordset
    Dim oPrevioFinan As PrevioFinan.clsPrevioFinan
    Dim lsError As String

    Set rsListaGenerarAsiento = oDDocumento.buscarPagosAdelantados(CInt(Trim(Right(cboRubro2.Text, 3))), _
                                                                   IIf(OptMoneda.Item(0).value, "1", "2"))
    
    lsMovNro = oNContFunciones.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    
    lsDescripcionPagoAdelantado = "Amortizaciones de Pagos Adelantados"
    
    fsCtaConDebe = optenerCtaConDebe(IIf(OptMoneda.Item(0).value, "1", "2"))
    
    lnConfirmarAsiento = oDDocumento.generarAsientoAmortizacionPagoAdelantado(lsMovNro, _
                                                                              fsOpeCod, _
                                                                              lsDescripcionPagoAdelantado, _
                                                                              fsCtaConDebe, _
                                                                              Trim(Right(cboMes.Text, 3)), _
                                                                              txtAnio, _
                                                                              IIf(OptMoneda.Item(0).value, 1, 2), _
                                                                              rsListaGenerarAsiento, lsError)
    

    If lnConfirmarAsiento > 0 Then
        oPista.InsertarPista fsOpeCod, lsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", lsDescripcionPagoAdelantado
        ImprimeAsientoContable lsMovNro, , , , , , , , , , , , , , 9
        sCadImp = oDDocumento.imprimirAmortizacionesPagosAdelantados(lsMovNro, gsNomCmac, CStr(gdFecSis), IIf(OptMoneda.Item(0).value, "SOLES", "DÓLARES"))
        
        If Len(sCadImp) = 0 Then
          MsgBox "No se encontraron datos de las amortizaciones", vbInformation, "Aviso"
        Else
          Set oPrevioFinan = New PrevioFinan.clsPrevioFinan
          PrevioFinan.Show sCadImp, "LISTA DE AMORTIZACIONES DE PAGOS ADELANTADOS", True
          Set oPrevioFinan = Nothing
        End If
            
        lsMovNro = ""
        fsCtaConDebe = ""
        lnConfirmarAsiento = 0
        lsError = ""
        sCadImp = ""
        Set rsListaGenerarAsiento = Nothing
        Set oDDocumento = Nothing
        Set oNContFunciones = Nothing
        Call cmdbuscar_Click
      
    Else
        If lnConfirmarAsiento = -1 Or lnConfirmarAsiento = -2 Then
            MsgBox "La cuenta " & lsError & " no existe", vbCritical, "Error"
        ElseIf lnConfirmarAsiento = -3 Then
             MsgBox lsError, vbInformation, "Aviso"
        ElseIf lnConfirmarAsiento = 0 Then
            MsgBox "No se registraron los datos."
        End If
        
        lsMovNro = ""
        fsCtaConDebe = ""
        lnConfirmarAsiento = 0
        lsError = ""
        sCadImp = ""
        Set rsListaGenerarAsiento = Nothing
        Set oDDocumento = Nothing
        Set oNContFunciones = Nothing

        
    End If
    
    txtAnio.Enabled = True
    chkCancelado.Enabled = True
    OptMoneda.Item(0).Enabled = True
    OptMoneda.Item(1).Enabled = True
    Call habilitarComboPagosAdelantados(True, gcboMes)
    Call habilitarComboPagosAdelantados(True, gcboRubro2)
    Call habilitarBotonesBuscarPagosAdelantados(True, gcmdBuscar)
    Call habilitarBotonesBuscarPagosAdelantados(True, gcmdExaminar)
    Call habilitarBotonesBuscarPagosAdelantados(True, gcmdGenerar)
    Call visualizarBotonesPagosAdelantados(True, gcmdNuevoPagoAdelantado)
    
    Set rsListaGenerarAsiento = Nothing
    Set oNContFunciones = Nothing
    Set oDDocumento = Nothing
    
End Sub

Private Sub cmdNuevoPagoAdelantado_Click()
Dim i As Integer

fnAccionPagoAdelantado = gNuevoRegistro
Call limpiarCajasTextosPagosAdelantados
Call habilitarCajasTextosPagosAdelantados(True)
Call habilitarBotonesPagosAdelantados(True)
Call habilitarCajasTextosPagosAdelantados(False, gTxtFecVen)
Call visualizarBotonesPagosAdelantados(True, gcmdAceptarPagoAdelantado)
Call visualizarBotonesPagosAdelantados(True, gcmdCancelarPagoAdelantado)
Call visualizarBotonesPagosAdelantados(False, gcmdNuevoPagoAdelantado)
If txtBuscarCtaCon.Enabled Then
    txtBuscarCtaCon.SetFocus
End If
txtFecIni = CStr(gdFecSis)
Call CargarRubros
Call habilitarComboPagosAdelantados(True, gcboRubro)
Call CargarAgencias

If chkTodasAgencias.Visible Then
    chkTodasAgencias.Enabled = False
    chkTodasAgencias.value = False
    chkTodasAgencias.Visible = False
End If

If lvAgencia.Visible Then
    lvAgencia.Enabled = False
    lvAgencia.Visible = False
End If

For i = 1 To CInt(lvAgencia.ListItems.Count)
    lvAgencia.ListItems(i).Checked = False
Next i

FrAgencia.Visible = False

End Sub

Private Sub Form_Load()
Dim i As Integer
Dim oNConstSistemas As New NConstSistemas
Set oPista = New COMManejador.Pista

fsOpeCod = "300461"
ldFecCie = CDate(oNConstSistemas.LeeConstSistema(gConstSistCierreMensualCont))
Call limpiarCajasTextosPagosAdelantados
Call habilitarBotonesPagosAdelantados(True, gcmdNuevoPagoAdelantado)
Call CargarCtaCont
Call CargarRubros2
Call cargarMeses
SSTabPagoAdelantado.Tab = 0
txtAnio = CStr(Year(gdFecSis))

For i = 0 To cboMes.ListCount - 1
    cboMes.ListIndex = i
    If Trim(Right(cboMes.Text, 3)) = CStr(Month(gdFecSis)) Then
        cboMes.ListIndex = i
        Exit For
    End If
Next i

Call habilitarBotonesBuscarPagosAdelantados(True, gcmdBuscar)
Call visualizarBotonesBuscarPagosAdelantados(True, gcmdBuscar)
End Sub



Private Sub OptMoneda_Click(Index As Integer)
Select Case Index
    Case 0
    OptMoneda.Item(0).value = True
    OptMoneda.Item(1).value = False
    Case 1
    OptMoneda.Item(0).value = False
    OptMoneda.Item(1).value = True
    
End Select
Call cmdbuscar_Click
End Sub



Private Sub txtBuscarCtaCon_EmiteDatos()
txtDesCtaCon = txtBuscarCtaCon.psDescripcion
If txtDesCtaCon <> "" And txtDesCtaCon.Enabled Then
   txtDesPagAde.SetFocus
End If
End Sub

Private Sub txtDesPagAde_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        txtFecIni.SetFocus
End If
End Sub

Private Sub txtFecIni_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(Trim(ValidaFecha(txtFecIni.Text))) <> 0 Then
            MsgBox ValidaFecha(txtFecIni.Text), vbInformation, "Aviso"
             txtFecIni.SetFocus
             Exit Sub
        End If
        If txtNroMes.Enabled = True Then
            txtNroMes.SetFocus
        Else
            If IsNumeric(txtNroMes) Then
                txtFecVen = DateAdd("M", CDbl(txtNroMes), CDate(txtFecIni))
            End If
            Me.cmdAceptarPagoAdelantado.SetFocus
        End If
    End If
End Sub

Private Sub txtImporte_GotFocus()
    If IsNumeric(txtNroMes) Then
     txtFecVen = DateAdd("M", CDbl(txtNroMes), CDate(txtFecIni))
    End If
End Sub

Private Sub txtImporte_KeyPress(KeyAscii As Integer)
    Call NumerosDecimales(txtImporte, KeyAscii)
    If KeyAscii = 13 Then
        cmdAceptarPagoAdelantado.SetFocus
    End If
End Sub

Private Sub txtNroMes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtNroMes) = "" Then
            MsgBox "El número de mes no debe estar vacío.", vbInformation, "Titulo"
            txtNroMes.SetFocus
            Exit Sub
        End If
        
        If Not IsNumeric(txtNroMes) Then
            MsgBox "Ingrese Número.", vbInformation, "Aviso"
            txtNroMes.SetFocus
            Exit Sub
        End If
        
        If CInt(txtNroMes) <= 0 Then
            MsgBox "El número de mes no debe ser menor o igual cero.", vbInformation, "Aviso"
            txtNroMes.SetFocus
            Exit Sub
        End If
        
        txtFecVen = DateAdd("M", CDbl(txtNroMes), CDate(txtFecIni))
        cboRubro.SetFocus
    End If
End Sub

Private Sub cboRubro_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
    txtImporte.SetFocus
 End If
End Sub
