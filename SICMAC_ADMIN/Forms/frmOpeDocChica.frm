VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmOpeDocChica 
   Caption         =   "Caja Chica: Registro de Documentos"
   ClientHeight    =   7440
   ClientLeft      =   660
   ClientTop       =   1065
   ClientWidth     =   10245
   Icon            =   "frmOpeDocChica.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   10245
   Begin VB.Frame Frame4 
      Caption         =   "Movimiento"
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
      Height          =   660
      Left            =   5985
      TabIndex        =   17
      Top             =   0
      Width           =   4035
      Begin MSMask.MaskEdBox txtFecha 
         Height          =   315
         Left            =   2685
         TabIndex        =   1
         Top             =   210
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label txtMovNro 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   690
         TabIndex        =   39
         Top             =   780
         Visible         =   0   'False
         Width           =   2805
      End
      Begin VB.Label txtOpeCod 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1005
         TabIndex        =   38
         Top             =   210
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha "
         Height          =   240
         Left            =   2040
         TabIndex        =   19
         Top             =   247
         Width           =   555
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Operacion :"
         Height          =   195
         Left            =   150
         TabIndex        =   18
         Top             =   270
         Width           =   825
      End
   End
   Begin VB.Frame fraDoc 
      Caption         =   "Documento "
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
      Height          =   1620
      Left            =   5985
      TabIndex        =   21
      Top             =   675
      Width           =   4035
      Begin VB.ComboBox cboDocDestino 
         Height          =   315
         ItemData        =   "frmOpeDocChica.frx":030A
         Left            =   780
         List            =   "frmOpeDocChica.frx":031A
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1095
         Width           =   2985
      End
      Begin VB.TextBox txtDocSerie 
         Height          =   315
         Left            =   510
         MaxLength       =   3
         TabIndex        =   5
         Top             =   660
         Width           =   540
      End
      Begin VB.ComboBox cboDoc 
         Height          =   315
         Left            =   510
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   263
         Width           =   3360
      End
      Begin VB.TextBox txtDocNro 
         Height          =   315
         Left            =   1050
         MaxLength       =   8
         TabIndex        =   6
         Top             =   660
         Width           =   1080
      End
      Begin MSMask.MaskEdBox txtDocFecha 
         Height          =   315
         Left            =   2730
         TabIndex        =   7
         Top             =   675
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label6 
         Caption         =   "Destino"
         Height          =   255
         Left            =   105
         TabIndex        =   25
         Top             =   1140
         Width           =   645
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo"
         Height          =   240
         Left            =   120
         TabIndex        =   24
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha"
         Height          =   240
         Left            =   2220
         TabIndex        =   23
         Top             =   690
         Width           =   480
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nro."
         Height          =   195
         Left            =   150
         TabIndex        =   22
         Top             =   720
         Width           =   300
      End
   End
   Begin VB.Frame FrameTipCambio 
      Caption         =   "Tipo de Cambio"
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
      Height          =   600
      Left            =   7380
      TabIndex        =   47
      Top             =   5280
      Visible         =   0   'False
      Width           =   2610
      Begin VB.TextBox txtTipFijo 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   315
         Left            =   405
         TabIndex        =   49
         Top             =   210
         Width           =   735
      End
      Begin VB.TextBox txtTipVariable 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   315
         Left            =   1770
         TabIndex        =   48
         Top             =   195
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "Fijo"
         Height          =   255
         Left            =   105
         TabIndex        =   51
         Top             =   255
         Width           =   345
      End
      Begin VB.Label Label9 
         Caption         =   "Banco"
         Height          =   255
         Left            =   1215
         TabIndex        =   50
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame fraProvis 
      Caption         =   "Provisión de ..."
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
      Height          =   660
      Left            =   90
      TabIndex        =   32
      Top             =   6705
      Width           =   2505
      Begin VB.ComboBox cboProvis 
         Height          =   315
         Left            =   135
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   225
         Width           =   2265
      End
   End
   Begin VB.Frame fraServicio 
      Height          =   4410
      Left            =   105
      TabIndex        =   27
      Top             =   2280
      Width           =   9945
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "A&gregar"
         Height          =   330
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   3165
         Width           =   1110
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   330
         Left            =   1230
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   3165
         Width           =   1110
      End
      Begin VB.CommandButton cmdValVenta 
         Caption         =   "Ajuste &Valor Venta"
         Height          =   330
         Left            =   120
         TabIndex        =   41
         ToolTipText     =   "Calcula el Valor Venta de Subtotales"
         Top             =   3525
         Width           =   2220
      End
      Begin VB.CommandButton cmdAjuste 
         Caption         =   "A&juste Manual   >>>"
         Height          =   330
         Left            =   120
         TabIndex        =   40
         ToolTipText     =   "Adiciona Importe de Ajuste de Documento"
         Top             =   3885
         Width           =   2220
      End
      Begin VB.PictureBox fgImp 
         Height          =   1035
         Left            =   3630
         ScaleHeight     =   975
         ScaleWidth      =   3525
         TabIndex        =   11
         Top             =   3195
         Width           =   3585
      End
      Begin VB.TextBox txtSTotal 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   285
         Left            =   8520
         TabIndex        =   34
         Top             =   3705
         Width           =   1185
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   285
         Left            =   8520
         TabIndex        =   33
         Top             =   4035
         Width           =   1185
      End
      Begin VB.PictureBox fgObj 
         Height          =   1245
         Left            =   195
         ScaleHeight     =   1185
         ScaleWidth      =   9525
         TabIndex        =   10
         Top             =   1710
         Width           =   9585
      End
      Begin VB.PictureBox fgDetalle 
         Height          =   1470
         Left            =   195
         ScaleHeight     =   1410
         ScaleWidth      =   9525
         TabIndex        =   9
         Top             =   195
         Width           =   9585
      End
      Begin VB.Label lblSTot 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "SubTotal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   7470
         TabIndex        =   37
         Top             =   3750
         Width           =   885
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   7470
         TabIndex        =   36
         Top             =   4065
         Width           =   915
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Retenciones"
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
         Height          =   195
         Left            =   2490
         TabIndex        =   31
         Top             =   3885
         Width           =   1095
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "y/o"
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
         Height          =   195
         Left            =   2895
         TabIndex        =   30
         Top             =   3555
         Width           =   315
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Impuestos"
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
         Height          =   195
         Left            =   2580
         TabIndex        =   29
         Top             =   3270
         Width           =   885
      End
      Begin VB.Label Label7 
         Caption         =   "Servicios"
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
         Height          =   285
         Left            =   150
         TabIndex        =   28
         Top             =   0
         Width           =   795
      End
      Begin VB.Shape ShapeIGV 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H8000000C&
         Height          =   345
         Left            =   7350
         Top             =   3675
         Width           =   2385
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H8000000C&
         Height          =   345
         Left            =   7350
         Top             =   4005
         Width           =   2385
      End
      Begin VB.Label Label13 
         BorderStyle     =   1  'Fixed Single
         Height          =   1020
         Left            =   2400
         TabIndex        =   35
         Top             =   3180
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   390
      Left            =   7575
      TabIndex        =   13
      Top             =   6870
      Width           =   1230
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      CausesValidation=   0   'False
      Height          =   390
      Left            =   8835
      TabIndex        =   14
      Top             =   6870
      Width           =   1230
   End
   Begin VB.Frame Frame3 
      Caption         =   "Glosa"
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
      Height          =   915
      Left            =   120
      TabIndex        =   16
      Top             =   1365
      Width           =   5835
      Begin VB.TextBox txtMovDesc 
         Height          =   630
         Left            =   90
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   195
         Width           =   5640
      End
   End
   Begin VB.Frame frameDestino 
      Caption         =   "Proveedor"
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
      Height          =   690
      Left            =   120
      TabIndex        =   20
      Top             =   675
      Width           =   5835
      Begin VB.PictureBox txtBuscarProv 
         Height          =   360
         Left            =   135
         ScaleHeight     =   300
         ScaleWidth      =   1665
         TabIndex        =   2
         Top             =   210
         Width           =   1725
      End
      Begin VB.Label lblProvNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1875
         TabIndex        =   15
         Top             =   225
         Width           =   3705
      End
   End
   Begin VB.Frame fraAjuste 
      Height          =   660
      Left            =   2595
      TabIndex        =   26
      Top             =   6705
      Visible         =   0   'False
      Width           =   4035
      Begin VB.TextBox txtAjuste 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   135
         TabIndex        =   46
         Top             =   210
         Width           =   1410
      End
      Begin VB.CommandButton cmdCancelar 
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
         Height          =   360
         Left            =   2745
         TabIndex        =   45
         Top             =   195
         Width           =   1170
      End
      Begin VB.CommandButton cmdAplicar 
         Caption         =   "&Aplicar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1590
         TabIndex        =   44
         Top             =   195
         Width           =   1170
      End
   End
   Begin VB.Frame frameCaja 
      Caption         =   "CAJA CHICA"
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
      Height          =   660
      Left            =   120
      TabIndex        =   52
      Top             =   0
      Visible         =   0   'False
      Width           =   5835
      Begin VB.TextBox lblCajaChicaDesc 
         Appearance      =   0  'Flat
         ForeColor       =   &H00004080&
         Height          =   330
         Left            =   1230
         Locked          =   -1  'True
         TabIndex        =   61
         Top             =   217
         Width           =   3885
      End
      Begin VB.PictureBox txtBuscarAreaCH 
         Height          =   345
         Left            =   135
         ScaleHeight     =   285
         ScaleWidth      =   1035
         TabIndex        =   0
         Top             =   210
         Width           =   1095
      End
      Begin VB.Label lblNroProc 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00004080&
         Height          =   330
         Left            =   5175
         TabIndex        =   60
         Top             =   210
         Width           =   570
      End
   End
   Begin VB.Frame fraArendir 
      Caption         =   "Recibo de A rendir"
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
      Height          =   660
      Left            =   120
      TabIndex        =   53
      Top             =   0
      Width           =   5835
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Saldo :"
         Height          =   195
         Left            =   3870
         TabIndex        =   59
         Top             =   285
         Width           =   495
      End
      Begin VB.Label lblSaldoArendir 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   4470
         TabIndex        =   58
         Top             =   225
         Width           =   1185
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Nro"
         Height          =   195
         Left            =   75
         TabIndex        =   57
         Top             =   285
         Width           =   255
      End
      Begin VB.Label lblArendirNro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   405
         TabIndex        =   56
         Top             =   225
         Width           =   1440
      End
      Begin VB.Label lblFechaArendir 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2685
         TabIndex        =   55
         Top             =   225
         Width           =   1065
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Fecha :"
         Height          =   195
         Left            =   2055
         TabIndex        =   54
         Top             =   285
         Width           =   540
      End
   End
End
Attribute VB_Name = "frmOpeDocChica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lMN As Boolean, sMoney As String
Dim cCtaCodPend As String
Dim lTransActiva As Boolean
Dim nTasaIGV As Currency, nTasaImp As Currency
Dim nVariaIGV As Currency

Dim aCtaCambio(1, 2) As String
Dim sObjRendir As String
Dim sConcepCod As String
Dim nConcepNiv As Integer
Dim sCtaViaticos As String
Dim sCtaViaticosDesc As String

'''********************************************************************
Dim vsDocAbrev As String
Dim vsDocNro As String
Dim vsProveedor As String
Dim vdFechaDoc As Date
Dim vnImporteDoc As Currency
Dim vsMovNroSust As String
Dim vsMovDesc As String

Dim oContFunct As NContFunciones
Dim oNArendir As NARendir
Dim oOpe As DOperacion
Dim oCtaCont As DCtaCont
Dim lnArendirFase As ARendirFases
Dim lnTipoArendir As ArendirTipo
Dim lbCajaChica As Boolean
Dim lnSaldoArendir As Currency
Dim lsCtaCodPend As String
Dim lbNewProv As Boolean
Dim lbTieneIGV As Boolean
Dim lsCtaCodProvis As String
Dim lsMovNroAten As String

Dim OK As Boolean
Dim lSalir As Boolean
Dim lsAgeCod As String
Dim lsAgeDes As String
Dim lsAreaCod As String
Dim lsAreaDes As String
Dim lsPersCod As String
Dim lsPersNom As String
Dim lsFechaARendir As String
Dim lsMovNroSol As String
Dim lsNroArendir As String
Dim lsAreaCh As String
Dim lsAgeCh As String
Dim lnNroProc As Integer
Public Sub Inicio(ByVal pnArendirFase As ARendirFases, ByVal pnTipoArendir As ArendirTipo, _
                    ByVal psAreaCod As String, ByVal psAreaDes As String, _
                    ByVal psAgeCod As String, ByVal psAgeDes As String, _
                    ByVal psPersCod As String, ByVal psPersNom As String, _
                    ByVal psFechaARendir As String, ByVal psMovNroAten As String, _
                    ByVal psMovNroSol As String, ByVal psNroArendir As String, _
                    Optional pbCajaChica As Boolean = False, _
                    Optional pnSaldoARendir As Currency = 0, _
                    Optional psAreaCh As String = "", Optional psAgeCh As String = "", _
                    Optional pnNroProc As Integer = 0)
                    
                    
lsAreaCh = psAreaCh
lsAgeCh = psAgeCh
lnNroProc = pnNroProc
                   
lsNroArendir = psNroArendir
lsAgeCod = psAgeCod
lsAgeDes = psAgeDes
lsAreaCod = psAreaCod
lsAreaDes = psAreaDes
lsMovNroAten = psMovNroAten
lsPersCod = psPersCod
lsPersNom = psPersNom
lnArendirFase = pnArendirFase
lnTipoArendir = pnTipoArendir
lbCajaChica = pbCajaChica
lsFechaARendir = psFechaARendir
lsMovNroSol = psMovNroSol
If pnArendirFase = ArendirRendicion Or pnArendirFase = ArendirSustentacion Then
   lnSaldoArendir = pnSaldoARendir
End If
Me.Show 1
End Sub
Public Sub InicioEgresoDirecto()
lbCajaChica = True
Me.Show 1
End Sub

Private Sub FormatoImpuesto()
'fgImp.ColWidth(0) = 250
'fgImp.ColWidth(1) = 750
'fgImp.ColWidth(2) = 550
'fgImp.ColWidth(3) = 0    'CtaContCod
'fgImp.ColWidth(4) = 0    'CtaContDes
'fgImp.ColWidth(5) = 0    'D/H
'fgImp.ColWidth(6) = 1200
'fgImp.ColWidth(7) = 0 'Destino 0/1
'fgImp.ColWidth(8) = 0 'Obligatorio, Opcional 1/2
'fgImp.ColWidth(9) = 0 'Total Impuesto no Gravado
'fgImp.TextMatrix(0, 1) = ""
'fgImp.TextMatrix(0, 2) = "Tasa"
'fgImp.TextMatrix(0, 6) = "Monto"
End Sub
'Private Sub RefrescaFgObj(nItem As Integer)
'Dim K  As Integer
'For K = 1 To fgObj.Rows - 1
'    If Len(fgObj.TextMatrix(K, 1)) Then
'       If fgObj.TextMatrix(K, 0) = nItem Then
'          fgObj.RowHeight(K) = 285
'       Else
'          fgObj.RowHeight(K) = 0
'       End If
'    End If
'Next
'End Sub
'
'Private Sub CalculaTotal(Optional lCalcImpuestos As Boolean = True)
'Dim N As Integer, m As Integer
'Dim nSTot As Currency
'Dim nITot As Currency, nImp As Currency
'Dim nTot  As Currency
'nSTot = 0: nTot = 0
'If fgImp.TextMatrix(1, 1) = "" Then
'   lCalcImpuestos = False
'End If
'For m = 1 To fgImp.Rows - 1
'   nITot = 0
'   For N = 1 To fgDetalle.Rows - 1
'      If fgImp.TextMatrix(m, 2) = "." Then
'         If lCalcImpuestos Then
'            nImp = Round(Val(Format(fgDetalle.TextMatrix(N, 3), gsFormatoNumeroDato)) * Val(Format(fgImp.TextMatrix(m, 4), gsFormatoNumeroDato)) / 100, 2)
'            fgDetalle.TextMatrix(N, m + 6) = Format(nImp, gsFormatoNumeroView)
'         Else
'            nImp = fgDetalle.TextMatrix(N, m + 6)
'         End If
'         nITot = nITot + nImp
'      Else
'         If lCalcImpuestos Then fgDetalle.TextMatrix(N, m + 6) = "0.00"
'      End If
'   Next
'   fgImp.TextMatrix(m, 5) = Format(nITot, gsFormatoNumeroView)
'   nTot = nTot + nITot * IIf(fgImp.TextMatrix(m, 8) = "D", 1, -1)
'Next
'For N = 1 To fgDetalle.Rows - 1
'   nSTot = nSTot + Val(Format(fgDetalle.TextMatrix(N, 3), gsFormatoNumeroDato))
'Next
'txtSTotal = Format(nSTot, gsFormatoNumeroView)
'txtTotal = Format(nSTot + nTot, gsFormatoNumeroView)
'End Sub
'
'Private Sub CalculaImpuesto()
'Dim nTot As Currency
'nTot = Val(Format(txtSTotal, gsFormatoNumeroDato))
'If fgImp.TextMatrix(fgImp.Row, 0) = "." Then
'   fgImp.TextMatrix(fgImp.Row, 6) = Format(Round(nTot * Val(Format(fgImp.TextMatrix(fgImp.Row, 2), gsFormatoNumeroDato)) / 100, 2), gsFormatoNumeroView)
'End If
'End Sub
'
'Private Function ValidaCajaChica() As Boolean
'Dim oCajaChica As NCajaChica
'Dim lnSaldo As Currency
'Set oCajaChica = New NCajaChica
'Dim lnTope As Currency
'ValidaCajaChica = True
'lnSaldo = oCajaChica.GetDatosCajaChica(Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2))
'If lnSaldo = 0 Then
'    MsgBox "Caja Chica Sin Saldo. Es necesario Solicitar Autorización o Desembolso", vbInformation, "Aviso"
'    ValidaCajaChica = False
'    Exit Function
'End If
'lnTope = oCajaChica.GetDatosCajaChica(Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2), MontoTope)
'If CCur(txtTotal) > lnTope Then
'    MsgBox "El Importe para solicitar a Caja Chica no puede ser mayor a " & Format(lnTope, gsFormatoNumeroView) & ". " & oImpresora.gPrnSaltoLinea  & "En Caso Contrario Solicite A rendir Cuenta con Caja General", vbInformation, "Aviso"
'    ValidaCajaChica = False
'    Exit Function
'End If
'
'If oCajaChica.VerificaTopeCajaChica(Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2)) = True Then
'    MsgBox "No puede realizar Egreso porque el saldo de esta Caja Chica es menor que el permitido." & oImpresora.gPrnSaltoLinea _
'            & "Por favor es necesario que realice Rendición", vbInformation, "Aviso"
'    ValidaCajaChica = False
'    If MsgBox(" ¿ Desea continuar con solicitud de Egreso de Caja Chica ? ", vbQuestion + vbYesNo, "Aviso") = vbYes Then
'         ValidaCajaChica = True
'    End If
'    Exit Function
'End If
'If lnSaldo < CCur(txtTotal) Then
'    MsgBox "Egreso no puede ser mayor que " & Format(lnSaldo, gsFormatoNumeroView), vbInformation, "Aviso"
'    ValidaCajaChica = False
'    Exit Function
'End If
'End Function
'
''Private Sub FormatoDetalle()
''fgDetalle.TextMatrix(0, 0) = "#"
''fgDetalle.TextMatrix(0, 1) = "Código"
''fgDetalle.TextMatrix(0, 2) = "Descripción"
''fgDetalle.TextMatrix(0, 3) = "Monto"
''fgDetalle.TextMatrix(0, 4) = "D/H"
''fgDetalle.TextMatrix(0, 5) = "Gravado"
''fgDetalle.TextMatrix(0, 6) = ""
''fgDetalle.ColWidth(0) = 380
''fgDetalle.ColWidth(1) = 1300
''fgDetalle.ColWidth(2) = 3470
''fgDetalle.ColWidth(3) = 1200
''fgDetalle.ColWidth(4) = 0
''fgDetalle.ColWidth(5) = 0
''fgDetalle.ColWidth(6) = 0
''
''fgDetalle.ColAlignment(0) = 4
''fgDetalle.ColAlignment(1) = 1
''fgDetalle.ColAlignment(3) = 7
''fgDetalle.ColAlignmentFixed(0) = 4
''fgDetalle.ColAlignmentFixed(3) = 7
''fgDetalle.Row = 1
''fgDetalle.Col = 1
''fgDetalle.RowHeight(-1) = 285
''End Sub
'Private Sub cboDoc_KeyPress(KeyAscii As Integer)
'Dim N As Integer
'Dim lvItem As ListItem
'If KeyAscii = 13 Then
'   cboDoc_Click
'   If txtDocSerie.Enabled Then
'      txtDocSerie.SetFocus
'   Else
'      txtDocFecha.SetFocus
'   End If
'End If
'End Sub
'Private Sub cboDoc_Click()
'Dim rs As ADODB.Recordset
'Dim oDoc As DDocumento
'
'Set oDoc = New DDocumento
'
'Dim nRow As Integer
'   lbTieneIGV = False
'   fgDetalle.Cols = 7
'   fgImp.Clear
'   fgImp.FormaCabecera
'   fgImp.Rows = 2
'   Set rs = New ADODB.Recordset
'   Set rs = oDoc.CargaDocImpuesto(Mid(cboDoc.Text, 1, 2))
'   Do While Not rs.EOF
'      'Primero adicionamos Columna de Impuesto
'        fgDetalle.Cols = fgDetalle.Cols + 1
'        fgDetalle.ColWidth(fgDetalle.Cols - 1) = 1000
'        fgDetalle.TextMatrix(0, fgDetalle.Cols - 1) = rs!cImpAbrev
'       'Adicionamos los impuestos en el grid de impuestos
'        fgImp.AdicionaFila
'        fgImp.Col = 0
'        nRow = fgImp.Row
'        fgImp.TextMatrix(nRow, 1) = "."
'        If rs!cDocImpOpc = "1" Then
'            'activamos el check de impuesto enviandole el valor "1"
'            If cboDocDestino.ListIndex <> 3 Then fgImp.TextMatrix(nRow, 2) = "1"
'        End If
'        fgImp.TextMatrix(nRow, 3) = rs!cImpAbrev
'        fgImp.TextMatrix(nRow, 4) = Format(rs!nImpTasa, gsFormatoNumeroView)
'        fgImp.TextMatrix(nRow, 5) = Format(0, gsFormatoNumeroView)
'        fgImp.TextMatrix(nRow, 6) = rs!cCtaContCod
'        fgImp.TextMatrix(nRow, 7) = rs!cCtaContDesc
'        fgImp.TextMatrix(nRow, 8) = rs!cDocImpDH
'        fgImp.TextMatrix(nRow, 9) = rs!cImpDestino
'        fgImp.TextMatrix(nRow, 10) = rs!cDocImpOpc
'        If rs!cCtaContCod = gcCtaIGV Then
'            lbTieneIGV = True
'            nTasaIGV = rs!nImpTasa
'        End If
'        rs.MoveNext
'   Loop
'   fgImp.Col = 1
'   If lbTieneIGV = False Then
'      cboDocDestino.ListIndex = -1
'      cboDocDestino.Enabled = False
'   Else
'      cboDocDestino.Enabled = True
'      cboDocDestino.ListIndex = 2
'   End If
'   VerReciboEgreso
'   CalculaTotal
'
'Set oDoc = Nothing
'End Sub
'Private Sub VerReciboEgreso()
'Dim lsReciboEgreso As String
'   If Mid(cboDoc.Text, 1, 2) = TpoDocRecEgreso Then
'      lsReciboEgreso = oContFunct.GeneraDocNro(TpoDocRecEgreso, Mid(gsOpeCod, 3, 1), gsCodUser)
'      txtDocSerie.MaxLength = 4
'      txtDocSerie = Mid(lsReciboEgreso, 1, 4)
'      txtDocSerie.Enabled = False
'      txtDocNro = Mid(lsReciboEgreso, 6, 20)
'      txtDocNro.Enabled = False
'   Else
'      txtDocSerie.MaxLength = 3
'      txtDocSerie = ""
'      txtDocSerie.Enabled = True
'      txtDocNro = ""
'      txtDocNro.Enabled = True
'   End If
'End Sub
'
'Private Sub cboDoc_Validate(Cancel As Boolean)
'If cboDoc = "" Then
'    Cancel = True
'End If
'End Sub
'
'Private Sub cboDocDestino_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'   cboDocDestino_Click
'   If fgImp.Enabled Then
'      fgImp.SetFocus
'   Else
'      fgDetalle.SetFocus
'   End If
'End If
'End Sub
'Private Sub cboDocDestino_Click()
'Dim I As Integer
'For I = 1 To fgImp.Rows - 1
'     'si el destino del impuesto es puede ser gravado y si es obligatorio
'    If fgImp.TextMatrix(I, 9) = "1" And fgImp.TextMatrix(I, 10) = "1" Then
'        If cboDocDestino.ListIndex = 3 Then
'           fgImp.TextMatrix(I, 2) = ""
'        Else
'           fgImp.TextMatrix(I, 2) = "1"
'        End If
'        CalculaTotal
'    Else
'        'fgImp.TextMatrix(i, 2) = ""
'    End If
'Next
'If cboDocDestino.ListIndex <> 1 Then
'    fgDetalle.ColWidth(5) = 0
'    For I = 1 To fgDetalle.Rows - 1
'        'limpia las posibles gravaciones manuales realizadas
'        If fgDetalle.TextMatrix(I, 5) = "." Then
'           fgDetalle.TextMatrix(I, 5) = ""
'       End If
'    Next
'Else
'    fgDetalle.ColWidth(5) = 500
'End If
'End Sub
'
'
'Private Sub cmdAceptar_Click()
'Dim N As Integer 'Contador
'Dim m As Integer
'Dim nItem As Integer, nCol  As Integer
'Dim sTexto As String
'Dim nImporte As Currency
'Dim nImpPend As Currency
'Dim lsMovUltAct As String
'Dim lsMovNro As String
'Dim ldFecha As Date
'Dim lsDocNro As String
'Dim lsReciboEgreso As String
'Dim lsDocAbrev As String
'
'On Error GoTo ErrAceptar
'
'If Not CamposOk Then Exit Sub
'Dim oDoc As DDocumento
'Set oDoc = New DDocumento
'If oDoc.GetValidaDocProv(txtBuscarProv.Text, Left(cboDoc, 2), txtDocSerie & "-" & txtDocNro) = True Then
'    MsgBox "Documento ya se encuentra registrado", vbInformation, "Aviso"
'    Exit Sub
'End If
'Set oDoc = Nothing
'If lbCajaChica Then
'   If ValidaCajaChica = False Then
'      Exit Sub
'   End If
'End If
'
'If MsgBox(" ¿ Seguro de grabar Operación ? ", vbYesNo + vbQuestion, "Aviso de Confirmación") = vbNo Then Exit Sub
'lsMovNro = oContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
'lsMovUltAct = GeneraMovNroActualiza(gdFecSis, gsCodUser, gsCodCMAC, gsCodAge)
'nImpPend = 0
'lsCtaCodProvis = Trim(Right(cboProvis.Text, 20))
'lsDocNro = txtDocSerie + "-" + txtDocNro
'lsDocAbrev = Trim(Right(cboDoc, 3))
'    If lbCajaChica = False Then
'        If lnArendirFase = ArendirSustentacion Then
'            oNArendir.GrabaSustentacionArendir gsFormatoFecha, lbNewProv, txtBuscarProv.Text, gLogProvEstadoActivado, lsMovUltAct, lsMovNro, _
'                                gsOpeCod, Trim(txtMovDesc), CCur(txtTotal), cboDocDestino.ListIndex, _
'                                lsCtaCodProvis, lsCtaCodPend, lnTipoArendir, lnSaldoArendir, lsDocNro, _
'                                Left(cboDoc, 2), CDate(txtDocFecha), CCur(txtTotal), lsMovNroAten, lsMovNroSol, lsAgeCod, lsAreaCod, fgDetalle.GetRsNew, fgObj.GetRsNew, _
'                                fgImp.GetRsNew, lsAreaCh, lsAgeCh, lnNroProc
'
'            lbNewProv = True
'            Ok = True
'        End If
'    Else
'        Dim oCH As NCajaChica
'        Set oCH = New NCajaChica
'        oCH.GrabaSolEgresoDirecto gsFormatoFecha, lbNewProv, txtBuscarProv.Text, gLogProvEstadoActivado, lsMovUltAct, lsMovNro, _
'                                    gsOpeCod, Trim(txtMovDesc), cboDocDestino.ListIndex, _
'                                    lsCtaCodProvis, lsDocNro, Left(cboDoc, 2), CDate(txtDocFecha), CCur(txtTotal), _
'                                    Mid(Me.txtBuscarAreaCH, 4, 2), Mid(Me.txtBuscarAreaCH, 1, 3), Val(lblNroProc), _
'                                    fgDetalle.GetRsNew, fgObj.GetRsNew, _
'                                    fgImp.GetRsNew
'
'        lbNewProv = True
'        Ok = True
'        Set oCH = Nothing
'    End If
'    ldFecha = CDate(Mid(lsMovNro, 7, 2) & "/" & Mid(lsMovNro, 5, 2) & "/" & Mid(lsMovNro, 1, 4))
'    If Left(cboDoc, 2) = Format(TpoDocRecEgreso, "00") Then
'        Dim oContImp As NContImprimir
'        Set oContImp = New NContImprimir
'        lsReciboEgreso = oContImp.ImprimeReciboEgresos(gnColPage, lsMovNro, txtMovDesc, ldFecha, gsNomCmac, gsOpeCod, _
'                        lbCajaChica, txtBuscarAreaCH & "-" & lblNroProc & " " & lblCajaChicaDesc, lnArendirFase, lsNroArendir, lsDocNro, CDate(txtDocFecha), _
'                        txtBuscarProv.Text, lblProvNombre, "", CCur(txtTotal))
'        EnviaPrevio lsReciboEgreso, Me.Caption, Int(gnLinPage / 2), False
'        Set oContImp = Nothing
'    End If
'    If lbCajaChica = False Or lnTipoArendir = gArendirTipoViaticos Then
'        vsDocAbrev = lsDocAbrev
'        vsDocNro = lsDocNro
'        vsProveedor = lblProvNombre
'
'        vdFechaDoc = CDate(txtDocFecha)
'        vnImporteDoc = CCur(txtTotal)
'        vsMovNroSust = lsMovNro
'        vsMovDesc = txtMovDesc
'        Ok = True
'        lSalir = True
'        Unload Me
'    Else
'       If MsgBox(" ¿ Desea registrar Nuevo Documento de Proveedor ? ", vbQuestion + vbYesNo, "Confirmación") = vbYes Then
'          txtMovNro = oContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
'          VerReciboEgreso
'          fgDetalle.Clear
'          fgDetalle.FormaCabecera
'          fgDetalle.Rows = 2
'
'          fgObj.Clear
'          fgObj.FormaCabecera
'          fgObj.Rows = 2
'
'          If cmdAjuste.Caption = "Desasignar A&juste   >>>" Then
'             cmdAjuste_Click
'          End If
'          txtMovDesc = ""
'          txtBuscarProv = ""
'          lblProvNombre = ""
'          'cboDoc.ListIndex = -1
'          CalculaTotal
'          txtBuscarProv.SetFocus
'          lbNewProv = False
'          Ok = False
'       Else
'           Unload Me
'       End If
'    End If
'Exit Sub
'ErrAceptar:
'   MsgBox Err.Description, vbInformation, "Aviso en Actualización"
'End Sub
'
'Private Function CamposOk() As Boolean
'CamposOk = False
'If frameCaja.Visible And txtBuscarAreaCH.Text = "" Then
'   MsgBox "No se definiò Caja Chica a solicitar", vbInformation, "Aviso"
'   txtBuscarAreaCH.SetFocus
'   Exit Function
'End If
'If Trim(txtBuscarProv) = "" Then
'   MsgBox "Falta especificar datos de Proveedor"
'   txtBuscarProv.SetFocus
'   Exit Function
'End If
'If txtMovDesc = "" Then
'   MsgBox "Falta especificar Motivo de Provisión", vbInformation, "Aviso"
'   txtMovDesc.SetFocus
'   Exit Function
'End If
'If cboDoc.ListIndex < 0 Then
'   MsgBox "No se especifico Tipo de Comprobante de Provisión", vbInformation, "Aviso"
'   cboDoc.SetFocus
'   Exit Function
'End If
'If txtDocSerie = "" Then
'   MsgBox "Falta especificar Serie de Comprobante", vbInformation, "Aviso"
'   txtDocSerie.SetFocus
'   Exit Function
'End If
'If txtDocNro = "" Then
'   MsgBox "Falta especificar Número de Comprobante", vbInformation, "Aviso"
'   txtDocNro.SetFocus
'   Exit Function
'End If
'If ValFecha(txtDocFecha) = False Then
'   txtDocFecha.SetFocus
'   Exit Function
'End If
'If Val(Format(txtTotal, gsFormatoNumeroDato)) = 0 Then
'   MsgBox "Falta indicar Importe de Documento", vbInformation, "Aviso"
'   fgDetalle.SetFocus
'   Exit Function
'End If
'CamposOk = True
'End Function
'
'Private Sub cmdAgregar_Click()
'Dim lnFila As Integer
'If Me.fgDetalle.TextMatrix(1, 0) = "" Then
'    fgDetalle.AdicionaFila , Val(fgDetalle.TextMatrix(fgDetalle.Row, 0)) + 1
'    lnFila = fgDetalle.Row
'    fgDetalle.TextMatrix(lnFila, 6) = fgDetalle.TextMatrix(lnFila, 0)
'    fgDetalle.SetFocus
'    SendKeys "{ENTER}"
'Else
'    If Val(Format(fgDetalle.TextMatrix(fgDetalle.Rows - 1, 3), gsFormatoNumeroDato)) <> 0 And _
'        Len(Format(fgDetalle.TextMatrix(fgDetalle.Rows - 1, 1), gsFormatoNumeroDato)) > 0 Then
'        fgDetalle.AdicionaFila , fgDetalle.TextMatrix(fgDetalle.Row, 0) + 1
'        lnFila = fgDetalle.Row
'        fgDetalle.TextMatrix(lnFila, 6) = fgDetalle.TextMatrix(lnFila, 0)
'    Else
'        If fgDetalle.Enabled Then
'           fgDetalle.SetFocus
'        End If
'    End If
'    fgDetalle.SetFocus
'    SendKeys "{ENTER}"
'End If
'End Sub
'
'Private Sub cmdAjuste_Click()
'If Val(txtTotal) = 0 Then
'   MsgBox "Primero ingresar Concepto de Gastos de Documento", vbInformation, "Aviso"
'   Exit Sub
'End If
'If cmdAjuste.Caption = "Asignar A&juste   >>>" Then
'   fraAjuste.Visible = True
'   fraServicio.Enabled = False
'   fraDoc.Enabled = False
'   txtAjuste.SetFocus
'Else
'    fgImp.EliminaFila fgImp.Rows - 1
'    cmdAjuste.Caption = "Asignar A&juste   >>>"
'    fgDetalle.Cols = fgDetalle.Cols - 1
'    fgDetalle.SetFocus
'End If
'End Sub
'Private Sub DesactivaAjuste()
'fraAjuste.Visible = False
'fraServicio.Enabled = True
'fgDetalle.Enabled = True
'fraDoc.Enabled = True
'fgDetalle.SetFocus
'End Sub
'
'Private Sub cmdAplicar_Click()
'Dim nRow As Integer
'Dim nTot As Currency
'Dim nImp As Currency
'fgImp.AdicionaFila
'nRow = fgImp.Row
'fgImp.Col = 2
'fgImp.TextMatrix(nRow, 2) = "1"
''Set fgImp.CellPicture = picCuadroSi.Picture
''fgImp.Col = 1
''Falta adicionar Check
'fgImp.TextMatrix(nRow, 0) = "."
'fgImp.TextMatrix(nRow, 3) = "AJUSTE"
'fgImp.TextMatrix(nRow, 4) = "AJUSTE"
'fgImp.TextMatrix(nRow, 5) = Format(txtAjuste, "#,#0.00")
'fgImp.TextMatrix(nRow, 6) = "AJUSTE"
'fgImp.TextMatrix(nRow, 7) = "AJUSTE"
'fgImp.TextMatrix(nRow, 8) = "D"
'fgImp.TextMatrix(nRow, 9) = "0"
'fgImp.TextMatrix(nRow, 10) = "2"
''Distribución del Ajuste entre Cuentas de Gasto
'fgDetalle.Cols = fgDetalle.Cols + 1
'fgDetalle.TextMatrix(0, fgDetalle.Cols - 1) = "AJUSTE"
'nTot = 0
'For nRow = 1 To fgDetalle.Rows - 1
'   nImp = Round(Val(txtAjuste) * Val(fgDetalle.TextMatrix(nRow, 3)) / Val(txtSTotal), 2)
'   nTot = nTot + nImp
'   fgDetalle.TextMatrix(nRow, fgDetalle.Cols - 1) = Format(nImp, gsFormatoNumeroView)
'Next
'If nTot <> Val(txtAjuste) Then
'   fgDetalle.TextMatrix(1, fgDetalle.Cols - 1) = Val(fgDetalle.TextMatrix(1, fgDetalle.Cols - 1)) + (Val(txtAjuste) - nTot)
'End If
'cmdAjuste.Caption = "Desasignar A&juste   >>>"
'CalculaTotal False
'DesactivaAjuste
'End Sub
'
'Private Sub cmdCancelar_Click()
'cmdAjuste.Caption = "Asignar A&juste   >>>"
'DesactivaAjuste
'End Sub
'
'Private Sub cmdEliminar_Click()
'
'If fgDetalle.TextMatrix(fgDetalle.Row, 0) <> "" Then
'   EliminaCuenta fgDetalle.TextMatrix(fgDetalle.Row, 1), fgDetalle.TextMatrix(fgDetalle.Row, 0)
'   CalculaTotal
'   If fgDetalle.Enabled Then
'      fgDetalle.SetFocus
'   End If
'End If
'End Sub
'Private Sub EliminaCuenta(sCod As String, nItem As Integer)
'If fgDetalle.TextMatrix(1, 0) <> "" Then
'    EliminaFgObj Val(fgDetalle.TextMatrix(fgDetalle.Row, 0))
'    fgDetalle.EliminaFila fgDetalle.Row, False
'End If
'If Len(fgDetalle.TextMatrix(1, 1)) > 0 Then
'   RefrescaFgObj Val(fgDetalle.TextMatrix(fgDetalle.Row, 0))
'End If
'End Sub
'Private Sub EliminaFgObj(nItem As Integer)
'Dim K  As Integer, m As Integer
'K = 1
'Do While K < fgObj.Rows
'   If Len(fgObj.TextMatrix(K, 1)) > 0 Then
'      If Val(fgObj.TextMatrix(K, 0)) = nItem Then
'         fgObj.EliminaFila K, False
'      Else
'         K = K + 1
'      End If
'   Else
'      K = K + 1
'   End If
'Loop
'End Sub
'Public Sub AsignaCtaObj(ByVal psCtaContCod As String)
'Dim sql As String
'Dim rs As ADODB.Recordset
'Dim rs1 As ADODB.Recordset
'Dim lsRaiz As String
'Dim oDescObj As ClassDescObjeto
'Dim UP As UPersona
'Dim lsFiltro As String
'Dim oRHAreas As DActualizaDatosArea
'Dim oCtaCont As DCtaCont
'Dim oCtaIf As NCajaCtaIF
'Dim oEfect As Defectivo
'
'Set oEfect = New Defectivo
'Set oCtaIf = New NCajaCtaIF
'Set oRHAreas = New DActualizaDatosArea
'Set oDescObj = New ClassDescObjeto
'Set oCtaCont = New DCtaCont
'Set rs = New ADODB.Recordset
'Set rs1 = New ADODB.Recordset
'EliminaFgObj Val(fgDetalle.TextMatrix(fgDetalle.Row, 0))
'Set rs1 = oCtaCont.CargaCtaObj(psCtaContCod, , True)
'If Not rs1.EOF And Not rs1.BOF Then
'    Do While Not rs1.EOF
'        lsRaiz = ""
'        lsFiltro = ""
'        Select Case Val(rs1!cObjetoCod)
'            Case ObjCMACAgencias
'                Set rs = oRHAreas.GetAgencias(rs1!cCtaObjFiltro)
'            Case ObjCMACAgenciaArea
'                lsRaiz = "Unidades Organizacionales"
'                Set rs = oRHAreas.GetAgenciasAreas(rs1!cCtaObjFiltro)
'            Case ObjCMACArea
'                Set rs = oRHAreas.GetAreas(rs1!cCtaObjFiltro)
'            Case ObjEntidadesFinancieras
'                lsRaiz = "Cuentas de Entidades Financieras"
'                Set rs = oCtaIf.GetCtasInstFinancieras(rs1!cCtaObjFiltro, psCtaContCod)
'            Case ObjDescomEfectivo
'                Set rs = oEfect.GetBilletajes(rs1!cCtaObjFiltro)
'            Case ObjPersona
'                Set rs = Nothing
'            Case Else
'                lsRaiz = "Varios"
'                Set rs = GetObjetos(Val(rs1!cObjetoCod))
'        End Select
'        If Not rs Is Nothing Then
'            If rs.State = adStateOpen Then
'                If Not rs.EOF And Not rs.BOF Then
'                    If rs.RecordCount > 1 Then
'                        oDescObj.Show rs, "", lsRaiz
'                        If oDescObj.lbOk Then
'                            lsFiltro = oContFunct.GetFiltroObjetos(Val(rs1!cObjetoCod), psCtaContCod, oDescObj.gsSelecCod, False)
'                            AdicionaObj psCtaContCod, fgDetalle.TextMatrix(fgDetalle.Row, 0), rs1!cCtaObjOrden, oDescObj.gsSelecCod, _
'                                        oDescObj.gsSelecDesc, lsFiltro, rs1!cObjetoCod
'                        Else
'                            fgDetalle.EliminaFila fgDetalle.Row, False
'                            Exit Do
'                        End If
'                    Else
'                        AdicionaObj psCtaContCod, fgDetalle.TextMatrix(fgDetalle.Row, 0), rs1!cCtaObjOrden, rs1!cObjetoCod, _
'                                        rs1!cObjetoDesc, lsFiltro, rs1!cObjetoCod
'                    End If
'                End If
'            End If
'        Else
'            If Val(rs1!cObjetoCod) = ObjPersona Then
'                Set UP = frmBuscaPersona.Inicio
'                If Not UP Is Nothing Then
'                    AdicionaObj psCtaContCod, fgDetalle.TextMatrix(fgDetalle.Row, 0), rs1!cCtaObjOrden, _
'                                    UP.sPersCod, UP.sPersNombre, _
'                                    lsFiltro, rs1!cObjetoCod
'                End If
'            End If
'        End If
'        rs1.MoveNext
'    Loop
'End If
'rs1.Close
'Set rs1 = Nothing
'Set oDescObj = Nothing
'Set UP = Nothing
'Set oCtaCont = Nothing
'Set oCtaIf = Nothing
'Set oEfect = Nothing
'End Sub
'Private Sub AdicionaObj(sCodCta As String, nFila As Integer, _
'                        psOrden As String, psObjetoCod As String, psObjDescripcion As String, _
'                        psSubCta As String, psObjPadre As String)
'Dim nItem As Integer
'    fgObj.AdicionaFila
'    nItem = fgObj.Row
'    fgObj.TextMatrix(nItem, 0) = nFila
'    fgObj.TextMatrix(nItem, 1) = psOrden
'    fgObj.TextMatrix(nItem, 2) = psObjetoCod
'    fgObj.TextMatrix(nItem, 3) = psObjDescripcion
'    fgObj.TextMatrix(nItem, 4) = sCodCta
'    fgObj.TextMatrix(nItem, 5) = psSubCta
'    fgObj.TextMatrix(nItem, 6) = psObjPadre
'    fgObj.TextMatrix(nItem, 7) = nFila
'    'fgDetalle.TextMatrix(fgDetalle.Row, 6) = psObjetoCod
'
'End Sub
'
'
'Private Sub cmdSalir_Click()
'Ok = False
'Unload Me
'End Sub
'Private Sub cmdValVenta_Click()
'Dim N As Integer
'For N = 1 To fgDetalle.Rows - 1
'    fgDetalle.TextMatrix(N, 3) = Format(Round(Val(Format(fgDetalle.TextMatrix(N, 3), gsFormatoNumeroDato)) / (1 + (nTasaIGV / 100)), 2), gsFormatoNumeroView)
'Next
'CalculaTotal
'nVariaIGV = 0
'End Sub
''Private Sub fgDetalle_KeyPress(KeyAscii As Integer)
''If fgDetalle.TextMatrix(fgDetalle.Row, 1) = "" Then
''   Exit Sub
''End If
''If fgDetalle.Col = 1 Then
''   If KeyAscii = 13 Then EnfocaTexto txtCta, IIf(KeyAscii = 13, 0, KeyAscii), fgDetalle
''   If KeyAscii = 32 Then
''      If cboDocDestino.ListIndex = 1 Then
''         If fgDetalle.TextMatrix(fgDetalle.Row, 5) = "X" Then
''            mnuGravado.Checked = True
''         Else
' '           mnuGravado.Checked = False
' '        End If
''         mnuGravado_Click
''      End If
''   End If
''End If
''If fgDetalle.Col = 3 Or fgDetalle.Col > 6 Then
''   If fgDetalle.Col > 6 And fgDetalle.Text = "" Then
''      Exit Sub
''   End If
''   If InStr("-0123456789.", Chr(KeyAscii)) > 0 Then
''      EnfocaTexto txtCant, KeyAscii, fgDetalle
''   Else
''      If KeyAscii = 13 Then EnfocaTexto txtCant, 0, fgDetalle
''   End If
''End If
''End Sub
''Private Sub fgDetalle_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
''If Button = 2 Then
''   If cboDocDestino.ListIndex = 1 And fgDetalle.Col = 1 Then
''      mnu_.Visible = True
''      mnuGravado.Visible = True
''      If fgDetalle.TextMatrix(fgDetalle.Row, 5) = "X" Then
''         mnuGravado.Checked = True
''      Else
''         mnuGravado.Checked = False
''      End If
''   Else
''      mnu_.Visible = False
''      mnuGravado.Visible = False
''   End If
''   PopupMenu mnuObj
''End If
'
'Private Sub fgDetalle_OnCellChange(pnRow As Long, pnCol As Long)
'If fgDetalle.TextMatrix(1, 0) <> "" Then
'    CalculaTotal
'End If
'End Sub
'
'Private Sub fgDetalle_OnEnterTextBuscar(psDataCod As String, pnRow As Long, pnCol As Long, pbEsDuplicado As Boolean)
'If psDataCod <> "" Then
'    fgDetalle.TextMatrix(pnRow, 4) = oOpe.GetOpeCtaDebeHaber(gsOpeCod, psDataCod)
'    AsignaCtaObj psDataCod
'End If
'End Sub
'
'Private Sub fgDetalle_RowColChange()
'If Len(fgDetalle.TextMatrix(1, 1)) > 0 Then
'   If fgDetalle.TextMatrix(fgDetalle.Row, 0) <> "" Then
'        RefrescaFgObj Val(fgDetalle.TextMatrix(fgDetalle.Row, 0))
'   End If
'End If
'End Sub
'
'Private Sub fgDetalle_Validate(Cancel As Boolean)
'If fgDetalle.TextMatrix(1, 0) <> "" Then
'    CalculaTotal
'End If
'End Sub
'
'Private Sub fgImp_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
'If fgDetalle.TextMatrix(1, 0) <> "" Then
'    CalculaTotal
'End If
'End Sub
'
'Private Sub fgImp_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
'If fgImp.TextMatrix(fgImp.Row, 10) <> "2" Then
'    Cancel = False
'End If
'
'End Sub
'
'Private Sub Form_Activate()
'If lSalir Then
'   Unload Me
'End If
'End Sub
'
'Private Sub Form_Load()
'Dim N As Integer, nSaldo As Currency, nCant As Currency
'Dim sCtaCod As String
'Dim rs As ADODB.Recordset
'
'Set rs = New ADODB.Recordset
'Set oContFunct = New NContFunciones
'Set oNArendir = New NARendir
'Set oOpe = New DOperacion
'Set oCtaCont = New DCtaCont
'CentraForm Me
'Me.Caption = gsOpeDesc
'lSalir = False
'Ok = False
'
'CambiaTamañoCombo cboDoc, 300
'CambiaTamañoCombo cboDocDestino, 250
'
'
''Defino el Nro de Movimiento
'txtOpeCod = gsOpeCod
'txtMovNro = oContFunct.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
'txtFecha = Format(gdFecSis, gsFormatoFechaView)
'lblFechaArendir = lsFechaARendir
'fgDetalle.TipoBusqueda = BuscaArbol
'fgDetalle.psRaiz = "Cuentas Contables"
'fgDetalle.rsTextBuscar = oOpe.EmiteOpeCtasNivel(gsOpeCod, "D")
'
'If Mid(gsOpeCod, 3, 1) = gMonedaExtranjera Then    'Identificación de Tipo de Moneda
'   lMN = False
'   sMoney = gcME
'   gsSimbolo = gcME
'   If gnTipCambio = 0 Then
'      If Not GetTipCambio(gdFecSis) Then
'         lSalir = True
'         Exit Sub
'      End If
'   End If
'   FrameTipCambio.Visible = True
'   txtTipFijo = gnTipCambio
'   txtTipVariable = gnTipCambioV
'Else
'   lMN = True
'   sMoney = gcMN
'   gsSimbolo = gcMN
'End If
'
'Set rs = oOpe.CargaOpeDoc(gsOpeCod, Digitado)
'
'Do While Not rs.EOF
'   cboDoc.AddItem rs!nDocTpo & " " & Mid(rs!cDocDesc & Space(100), 1, 100) & Mid(rs!cDocAbrev & "   ", 1, 3)
'   rs.MoveNext
'Loop
'
'Set rs = oOpe.CargaOpeCta(gsOpeCod, "H", "0")
'If Not rs.EOF And Not rs.EOF Then
'    Do While Not rs.EOF
'        cboProvis.AddItem rs!cCtaContDesc & Space(100) & rs!cCtaContCod
'        rs.MoveNext
'    Loop
'    cboProvis.ListIndex = cboProvis.ListCount - 1
'Else
'    MsgBox "No se definieron Cuentas de Provisión para Operación. Por favor Consultar con Sistemas", vbInformation, "Aviso"
'    lSalir = True
'    Exit Sub
'End If
'If lnArendirFase = ArendirSustentacion Or lnArendirFase = ArendirRendicion Then 'Para a rendir de Caja General
'   fraProvis.Visible = False
'   lsCtaCodPend = oOpe.EmiteOpeCta(gsOpeCod, "H", "1")
'   If lsCtaCodPend = "" Then
'      MsgBox "No se definió Cuenta de Pendiente para Arendir", vbInformation, "Aviso"
'      lSalir = True
'      Exit Sub
'   End If
'End If
'If lbCajaChica Then   'Cuando es Caja Chica
'    txtFecha.Enabled = False
'    fraArendir.Visible = False
'    frameCaja.Visible = True
'    fraProvis.Visible = False
'    txtBuscarAreaCH.psRaiz = "Cajas Chicas"
'    txtBuscarAreaCH.rs = oNArendir.EmiteCajasChicas
'Else
'    fraArendir.Visible = True
'    lblArendirNro = lsNroArendir
'    lblSaldoArendir = Format(lnSaldoArendir, "#,#0.00")
'End If
'txtFecha.Enabled = lbCajaChica
'lbNewProv = False
'End Sub
'Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'If Not Ok And lSalir = False Then
'   If MsgBox(" ¿ Seguro de salir sin grabar Operación ? ", vbQuestion + vbYesNo) = vbNo Then
'      Cancel = 1
'      Exit Sub
'   End If
'End If
'End Sub
'
'Private Sub Form_Unload(Cancel As Integer)
'Set oContFunct = Nothing
'Set oNArendir = Nothing
'End Sub
'
'
'
'Private Sub txtAjuste_KeyPress(KeyAscii As Integer)
'KeyAscii = NumerosDecimales(txtAjuste, KeyAscii, 16, 2)
'If KeyAscii = 13 Then
'   txtAjuste = Format(txtAjuste, gsFormatoNumeroView)
'   cmdAplicar.SetFocus
'End If
'End Sub
'
'Private Sub txtAjuste_Validate(Cancel As Boolean)
'txtAjuste = Format(txtAjuste, gsFormatoNumeroView)
'End Sub
'Private Sub txtBuscarAreaCH_EmiteDatos()
'Dim oCajaCH As NCajaChica
'Dim lnSaldo As Currency
'Set oCajaCH = New NCajaChica
'If txtBuscarAreaCH.Text = "" Then Exit Sub
'lblCajaChicaDesc = txtBuscarAreaCH.psDescripcion
'lblNroProc = oCajaCH.GetDatosCajaChica(Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2), NroCajaChica)
'lnSaldo = oCajaCH.GetDatosCajaChica(Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2), SaldoActual)
'If lnSaldo = 0 Then
'    MsgBox "Caja Chica no posee Saldo", vbInformation, "Aviso"
'    lblCajaChicaDesc = ""
'    txtBuscarAreaCH = ""
'    lblNroProc = ""
'    Exit Sub
'End If
'If oCajaCH.VerificaTopeCajaChica(Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2)) Then
'    If MsgBox("Caja chica ha sobrepasado el limite establecido. debe realizar la Rendición Respectiva " & vbCrLf & "Desea Continuar ??", vbYesNo + vbQuestion, "Aviso") = vbNo Then
'        lblCajaChicaDesc = ""
'        txtBuscarAreaCH = ""
'        lblNroProc = ""
'        Exit Sub
'    End If
'End If
'If lblCajaChicaDesc <> "" Then
'    If txtFecha.Enabled Then
'        txtFecha.SetFocus
'    End If
'End If
'Set oCajaCH = Nothing
'End Sub
'Private Sub txtBuscarAreaCH_Validate(Cancel As Boolean)
'If txtBuscarAreaCH = "" Then Cancel = True
'
'End Sub
'
'Private Sub txtBuscarProv_EmiteDatos()
'Dim oProv As DLogProveedor
'Set oProv = New DLogProveedor
'lbNewProv = False
'lblProvNombre.Caption = txtBuscarProv.psDescripcion
'If lblProvNombre.Caption <> "" Then
'    lbNewProv = Not oProv.IsExisProveedor(txtBuscarProv.Text)
'    txtMovDesc.SetFocus
'End If
'Set oProv = Nothing
'End Sub
'
'Private Sub txtBuscarProv_Validate(Cancel As Boolean)
'If (txtBuscarProv = "" And txtBuscarProv.psDescripcion = "") Then
'   Cancel = True
'End If
'End Sub
'Private Sub txtDocFecha_KeyPress(KeyAscii As Integer)
'Dim nTipFijo As Currency
'If KeyAscii = 13 Then
'  ' If IsDate(txtDocFecha.Text) Then
'   If ValidaFecha(txtDocFecha.Text) = "" Then
'      If (lnArendirFase = ArendirRendicion Or lnArendirFase = ArendirSustentacion) Or lnTipoArendir = gArendirTipoViaticos Then
'         If CDate(txtDocFecha) < lsFechaARendir Then
'            MsgBox "Fecha no puede ser menor a fecha de A rendir", vbInformation, "Aviso"
'            Exit Sub
'         End If
'      End If
'      If CDate(txtDocFecha) - 3 > gdFecSis Then
'         MsgBox "Fecha no puede ser mayor a fecha Actual", vbInformation, "Aviso"
'         Exit Sub
'      End If
'      If Not lMN Then
'         nTipFijo = gnTipCambio
'         GetTipCambio CDate(txtDocFecha)
'         txtTipVariable = Format(gnTipCambioV, "###,###,##0.000")
'         gnTipCambio = nTipFijo
'      End If
'      If cboDocDestino.Enabled Then
'         cboDocDestino.SetFocus
'      Else
'         cmdAgregar.SetFocus
'      End If
'   Else
'      MsgBox "Fecha no válida...!", vbInformation, "Aviso"
'      txtDocFecha.SelStart = 0
'      txtDocFecha.SelLength = Len(txtDocFecha.Text)
'   End If
'End If
'End Sub
'
'Private Sub txtDocNro_GotFocus()
'fEnfoque txtDocNro
'End Sub
'
'Private Sub txtDocNro_Validate(Cancel As Boolean)
'Dim oDoc As DDocumento
'Set oDoc = New DDocumento
'If oDoc.GetValidaDocProv(txtBuscarProv.Text, Mid(cboDoc.Text, 1, 2), txtDocSerie & "-" & txtDocNro) Then
'    MsgBox "Documento ya ha sido Ingresado", vbInformation, "Aviso"
'    Cancel = True
'End If
'Set oDoc = Nothing
'End Sub
'
'Private Sub txtFecha_GotFocus()
'txtFecha.SelStart = 0
'txtFecha.SelLength = Len(txtFecha.Text)
'End Sub
'Private Sub txtFecha_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'   If ValidaFecha(txtFecha.Text) <> "" Then
'      MsgBox "Fecha no válida...!", vbInformation, "Error"
'      Exit Sub
'   End If
'   txtMovNro = Format(txtFecha.Text & " " & Time, "yyyymmddhhmmss") & Right(gsCodAge, 2) & "0000" & gsCodUser
'   txtBuscarProv.SetFocus
'End If
'End Sub
'Private Sub txtFecha_LostFocus()
'   If ValidaFecha(txtFecha.Text) <> "" Then
'      MsgBox "Fecha no válida...!", vbInformation, "Error"
'      Exit Sub
'   End If
'   txtMovNro = Format(txtFecha.Text & " " & Time, "yyyymmddhhmmss") & Right(gsCodAge, 2) & "0000" & gsCodUser
'End Sub
'Private Sub txtMovDesc_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'   KeyAscii = 0
'   cboDoc.SetFocus
'End If
'End Sub
'Private Sub txtDocSerie_KeyPress(KeyAscii As Integer)
'KeyAscii = NumerosEnteros(KeyAscii)
'If KeyAscii = 13 Then
'   txtDocSerie = Right(String(3, "0") & txtDocSerie, 3)
'   txtDocNro.SetFocus
'End If
'End Sub
'Private Sub txtDocNro_KeyPress(KeyAscii As Integer)
'KeyAscii = NumerosEnteros(KeyAscii)
'Dim oDoc As DDocumento
'Set oDoc = New DDocumento
'If KeyAscii = 13 Then
'   txtDocNro = Right(String(8, "0") & txtDocNro, 8)
'   If oDoc.GetValidaDocProv(txtBuscarProv.Text, Mid(cboDoc.Text, 1, 2), txtDocSerie & "-" & txtDocNro) Then
'        MsgBox "Documento ya ha sido Ingresado", vbInformation, "Aviso"
'        txtDocNro.SetFocus
'        Exit Sub
'   End If
'   txtDocFecha.SetFocus
'End If
'Set oDoc = Nothing
'End Sub
'Private Sub txtTipVariable_KeyPress(KeyAscii As Integer)
'KeyAscii = NumerosDecimales(txtTipVariable, KeyAscii, 7, 2)
'If KeyAscii = 13 Then
'    If cmdAceptar.Enabled Then
'        cmdAceptar.SetFocus
'    Else
'         cmdAplicar.SetFocus
'    End If
'End If
'End Sub
'Public Property Get lOk() As Boolean
'lOk = Ok
'End Property
'Public Property Let lOk(ByVal vNewValue As Boolean)
'Ok = vNewValue
'End Property
'Public Property Get DocAbrev() As String
'DocAbrev = vsDocAbrev
'End Property
'Public Property Let DocAbrev(ByVal vNewValue As String)
'vsDocAbrev = vNewValue
'End Property
'Public Property Get DocNro() As String
'DocNro = vsDocNro
'End Property
'Public Property Let DocNro(ByVal vNewValue As String)
'vsDocNro = vNewValue
'End Property
'Public Property Get Proveedor() As String
' Proveedor = vsProveedor
'End Property
'Public Property Let Proveedor(ByVal vNewValue As String)
'vsProveedor = vNewValue
'End Property
'Public Property Get FechaDoc() As Date
'FechaDoc = vdFechaDoc
'End Property
'Public Property Let FechaDoc(ByVal vNewValue As Date)
'vdFechaDoc = vNewValue
'End Property
'Public Property Get ImporteDoc() As Currency
'ImporteDoc = vnImporteDoc
'End Property
'Public Property Let ImporteDoc(ByVal vNewValue As Currency)
'vnImporteDoc = vNewValue
'End Property
'Public Property Get MovNroSust() As String
'MovNroSust = vsMovNroSust
'End Property
'Public Property Let MovNroSust(ByVal vNewValue As String)
'vsMovNroSust = vNewValue
'End Property
'Public Property Get MovDesc() As String
'MovDesc = vsMovDesc
'End Property
'Public Property Let MovDesc(ByVal vNewValue As String)
'vsMovDesc = vNewValue
'End Property
'
'

