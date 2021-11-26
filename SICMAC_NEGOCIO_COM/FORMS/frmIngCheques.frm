VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmIngCheques 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registro de Cheques"
   ClientHeight    =   6765
   ClientLeft      =   1875
   ClientTop       =   1590
   ClientWidth     =   8280
   Icon            =   "frmIngCheques.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   8280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraIngCheque 
      Caption         =   "Datos Generales"
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
      Height          =   1890
      Left            =   90
      TabIndex        =   30
      Top             =   105
      Width           =   8115
      Begin VB.TextBox txtIngChqNumCheque 
         Height          =   300
         Left            =   120
         MaxLength       =   15
         TabIndex        =   0
         Top             =   360
         Width           =   1485
      End
      Begin VB.TextBox txtIngChqCtaIF 
         Height          =   315
         Left            =   3240
         MaxLength       =   20
         TabIndex        =   4
         Top             =   360
         Width           =   1845
      End
      Begin VB.TextBox txtIngChqNumCheque1 
         Height          =   300
         Left            =   1680
         MaxLength       =   3
         TabIndex        =   1
         Top             =   360
         Width           =   285
      End
      Begin VB.TextBox txtIngChqNumCheque2 
         Height          =   300
         Left            =   2040
         MaxLength       =   3
         TabIndex        =   2
         Top             =   360
         Width           =   525
      End
      Begin VB.TextBox txtIngChqNumCheque3 
         Height          =   300
         Left            =   2640
         MaxLength       =   3
         TabIndex        =   3
         Top             =   360
         Width           =   525
      End
      Begin VB.TextBox txtIngChqNumCheque4 
         Height          =   300
         Left            =   5160
         MaxLength       =   3
         TabIndex        =   5
         Top             =   360
         Width           =   405
      End
      Begin VB.ComboBox cboMoneda 
         Height          =   315
         Left            =   7110
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1125
         Width           =   885
      End
      Begin VB.ComboBox cboIngChqPlaza 
         Height          =   315
         ItemData        =   "frmIngCheques.frx":030A
         Left            =   840
         List            =   "frmIngCheques.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1065
         Width           =   1380
      End
      Begin VB.CheckBox chkConfirmar 
         Caption         =   "Por Confimar en Caja Gen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   3960
      End
      Begin MSMask.MaskEdBox txtIngChqFechaReg 
         Height          =   315
         Left            =   3090
         TabIndex        =   8
         Top             =   1095
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtIngChqFechaVal 
         Height          =   315
         Left            =   5250
         TabIndex        =   9
         Top             =   1110
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin SICMACT.TxtBuscar TxtBCodPers 
         Height          =   285
         Left            =   840
         TabIndex        =   11
         Top             =   1440
         Width           =   1980
         _ExtentX        =   3281
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
         TipoBusqueda    =   3
         sTitulo         =   ""
      End
      Begin VB.Label lblbanco 
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
         Height          =   675
         Left            =   5760
         TabIndex        =   42
         Top             =   240
         Width           =   2160
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
         Left            =   120
         TabIndex        =   41
         Top             =   165
         Width           =   855
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
         Left            =   3480
         TabIndex        =   40
         Top             =   120
         Width           =   840
      End
      Begin VB.Label Label5 
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
         Left            =   2760
         TabIndex        =   39
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label6 
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
         Left            =   2160
         TabIndex        =   38
         Top             =   120
         Width           =   300
      End
      Begin VB.Label lblIngChqDescIF1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   3090
         TabIndex        =   12
         Top             =   1455
         Width           =   4635
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
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
         Height          =   210
         Left            =   120
         TabIndex        =   37
         Top             =   1560
         Width           =   585
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Moneda:"
         Height          =   195
         Left            =   6435
         TabIndex        =   36
         Top             =   1155
         Width           =   630
      End
      Begin VB.Label lblEstado 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   225
         Left            =   5895
         TabIndex        =   35
         Top             =   660
         Width           =   2040
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Valorización:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4305
         TabIndex        =   33
         Top             =   1155
         Width           =   945
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Plaza:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   105
         TabIndex        =   32
         Top             =   1110
         Width           =   435
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Registro:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2445
         TabIndex        =   31
         Top             =   1125
         Width           =   645
      End
   End
   Begin VB.Frame fraEspecifica 
      Height          =   2760
      Left            =   105
      TabIndex        =   26
      Top             =   2010
      Width           =   8115
      Begin VB.Frame Frame2 
         Caption         =   "Motivo de Recepción"
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
         Height          =   1080
         Left            =   90
         TabIndex        =   27
         Top             =   135
         Width           =   7860
         Begin SICMACT.TxtBuscar txtBuscarProd 
            Height          =   330
            Left            =   1170
            TabIndex        =   13
            Top             =   285
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   582
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            sTitulo         =   ""
         End
         Begin SICMACT.TxtBuscar txtBuscarAreaAgencia 
            Height          =   330
            Left            =   1170
            TabIndex        =   15
            Top             =   630
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   582
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            sTitulo         =   ""
         End
         Begin VB.Label Label2 
            Caption         =   "Producto"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   105
            TabIndex        =   29
            Top             =   300
            Width           =   645
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Area/Agencia:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   105
            TabIndex        =   28
            Top             =   690
            Width           =   1050
         End
         Begin VB.Label lblProdDesc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   2595
            TabIndex        =   14
            Top             =   285
            Width           =   5190
         End
         Begin VB.Label lblAreaAgeDesc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   2595
            TabIndex        =   16
            Top             =   645
            Width           =   5175
         End
      End
      Begin SICMACT.FlexEdit fgObjMotivo 
         Height          =   990
         Left            =   255
         TabIndex        =   17
         Top             =   1680
         Width           =   7470
         _ExtentX        =   13176
         _ExtentY        =   1746
         Cols0           =   5
         HighLight       =   2
         AllowUserResizing=   3
         EncabezadosNombres=   "-Objeto-Descripción-SubCta-cObjetoCod"
         EncabezadosAnchos=   "350-1600-4000-1000-0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-X"
         TextStyleFixed  =   3
         ListaControles  =   "0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-C-C"
         FormatosEdit    =   "0-0-0-0-0"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   345
         RowHeight0      =   285
      End
      Begin SICMACT.TxtBuscar txtBuscarCtaHaber 
         Height          =   330
         Left            =   1245
         TabIndex        =   18
         Top             =   1290
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   582
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         sTitulo         =   ""
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Motivo :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   330
         TabIndex        =   34
         Top             =   1320
         Width           =   645
      End
      Begin VB.Label lblCtaHaber 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2655
         TabIndex        =   19
         Top             =   1305
         Width           =   5160
      End
   End
   Begin VB.Frame fraGlosa 
      Caption         =   "Glosa :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   24
      Top             =   4800
      Width           =   8115
      Begin SICMACT.EditMoney txtMonto 
         Height          =   330
         Left            =   5760
         TabIndex        =   21
         Top             =   1035
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   582
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtMovDesc 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   180
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   20
         Top             =   210
         Width           =   7710
      End
      Begin VB.Label Label1 
         Caption         =   "Monto  : "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   210
         Left            =   4980
         TabIndex        =   25
         Top             =   1065
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   360
      Left            =   6960
      TabIndex        =   23
      Top             =   6360
      Width           =   1290
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   5685
      TabIndex        =   22
      Top             =   6360
      Width           =   1290
   End
End
Attribute VB_Name = "frmIngCheques"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oCtasIF As COMNCajaGeneral.NCOMCajaCtaIF 'NCajaCtaIF
Dim oContFunct As COMNContabilidad.NCOMContFunciones 'NContFunciones
Dim oGen As COMDConstSistema.DCOMGeneral 'DGeneral
Dim oOpe As COMDConstSistema.DCOMOperacion 'DOperacion
Dim nProducto As COMDConstantes.Producto
Dim oPersona As UPersona_Cli   ' COMDPersona.DCOMPersona

Dim lbOk As Boolean
Dim lbRegistroOperaciones As Boolean
Dim lsPersCodIf  As String
Dim lsNroCtaIf As String
Dim lsNroChq As String
Dim lnPlazaChq As COMDConstantes.ChequePlaza 'ChequePlaza
Dim ldFechaRegChq As Date
Dim ldFechaValChq As Date
Dim lsConfCheque As String
Dim lsGlosa As String
Dim lnMonto As Currency
Dim lnImporte As Double

Dim lsCtaContChq As String
Dim lbMuestra As Boolean
Dim lbNegocio As Boolean
Dim lsOpeCod As String
Dim lsCtaContHaber As String
Dim lnDiasValoriza As Long
Dim lbRegCheque As Boolean
Dim lsMovRef As String
Dim lsNombreIF As String

Dim lnDiaValoriza As Integer
'variables para Arendir
Dim lsMovNroAtenc As String
Dim lsMovNroSol As String
Dim lnTipoArendir As COMDConstantes.ArendirTipo
Dim lbArendir As Boolean
Dim lnMoneda  As COMDConstantes.Moneda
Dim lnOrdenProd As String
Dim lnOrdenAgencia As String
Dim lbApertura As Boolean

Dim lbCreaSubCta As Boolean
Dim lsSubCtaIFCod As String
Dim lsSubCtaIFDesc As String
Dim lsPersCodAper As String
Dim lsIFTpoAper As String
Dim lsCtaIFCod As String
Dim lsCtaIFDesc As String
Dim ldCtaIFAper As Date
Dim lsCtaIFVenc As String
Dim lnCtaIFPlazo As Integer
Dim lnPeriodo As Integer
Dim lnInteres  As Currency
Dim lnTpoDocAper As TpoDoc
Dim lsNroDocAper As String
Dim ldFechaDocApera As Date
Dim lsDocumentoAper As String

Dim rsObj As ADODB.Recordset
Dim lbSoloIngreso  As Boolean
Dim lsProductoCod As String
Dim lsAreaAgeCod As String
Dim lsCtaMotivo As String

Dim cCodPerBco As String 'MADM 20110628

Private Function GetDiasMinValorizacion(Optional nmoneda As Moneda = gMonedaNacional, _
    Optional nPlaza As ChequePlaza = gChqPlazaLocal, Optional nTipoCheque As ChequeTipo = gChqTpoSimple) As Integer
Dim oCap As COMNCaptaGenerales.NCOMCaptaDefinicion 'nCapDefinicion
Set oCap = New COMNCaptaGenerales.NCOMCaptaDefinicion
GetDiasMinValorizacion = oCap.GetDiasMinValorizacion(nmoneda, nPlaza, nTipoCheque)
Set oCap = Nothing
End Function

Private Sub cboIngChqPlaza_Click()
Dim dFecha As Date
Dim lnFeriado As Integer
Dim oCap As COMNCaptaGenerales.NCOMCaptaDefinicion

If cboIngChqPlaza.Text <> "" And cboMoneda.Text <> "" Then
    lnPlazaChq = CLng(Right(Trim(cboIngChqPlaza.Text), 1))
    lnMoneda = CLng(Right(Trim(cboMoneda.Text), 1))
    lnDiaValoriza = GetDiasMinValorizacion(lnMoneda, lnPlazaChq)
    dFecha = DateAdd("d", CDate(txtIngChqFechaReg.Text), lnDiaValoriza)
    
    Set oCap = New COMNCaptaGenerales.NCOMCaptaDefinicion
        lnFeriado = oCap.ObtenerFeriado(txtIngChqFechaReg.Text, dFecha)
    Set oCap = Nothing
    dFecha = DateAdd("d", CDate(dFecha), lnFeriado)
    
    If Weekday(dFecha, vbMonday) = 6 Then
        dFecha = DateAdd("d", 2, dFecha)
    ElseIf Weekday(dFecha, vbMonday) = 7 Then
        dFecha = DateAdd("d", 1, dFecha)
        dFecha = CDate(dFecha) + 1
    Else
        'VERIFICA SI EL DIA Q SE VALORIZA ES FERIADO--AVMM--30-10-2006
        Set oCap = New COMNCaptaGenerales.NCOMCaptaDefinicion
            lnFeriado = oCap.ObtenerFeriado(dFecha, dFecha)
        Set oCap = Nothing
        dFecha = DateAdd("d", CDate(dFecha), lnFeriado)
        dFecha = CDate(dFecha) + 1
    End If
    'VERIFICA SI EL DIA Q SE VALORIZA ES FERIADO--AVMM--30-10-2006
    Set oCap = New COMNCaptaGenerales.NCOMCaptaDefinicion
        lnFeriado = oCap.ObtenerFeriado(dFecha, dFecha)
    Set oCap = Nothing
    dFecha = DateAdd("d", CDate(dFecha), lnFeriado)
    
    txtIngChqFechaVal.Text = Format$(dFecha, "dd/mm/yyyy")
    
End If
End Sub

Private Sub cboMoneda_Click()
Dim dFecha As Date
Dim lnFeriado As Integer
Dim oCap As COMNCaptaGenerales.NCOMCaptaDefinicion

If cboIngChqPlaza.Text <> "" And cboMoneda.Text <> "" Then
    lnMoneda = CLng(Right(Trim(cboMoneda.Text), 1))
    lnPlazaChq = CLng(Right(Trim(cboIngChqPlaza.Text), 1))
    lnDiaValoriza = GetDiasMinValorizacion(lnMoneda, lnPlazaChq)
    dFecha = DateAdd("d", CDate(txtIngChqFechaReg.Text), lnDiaValoriza)
    
    Set oCap = New COMNCaptaGenerales.NCOMCaptaDefinicion
        lnFeriado = oCap.ObtenerFeriado(txtIngChqFechaReg.Text, dFecha)
    Set oCap = Nothing
    dFecha = DateAdd("d", CDate(dFecha), lnFeriado)
    
    If Weekday(dFecha, vbMonday) = 6 Then
        dFecha = DateAdd("d", 2, dFecha)
    ElseIf Weekday(dFecha, vbMonday) = 7 Then
        dFecha = DateAdd("d", 1, dFecha)
        dFecha = CDate(dFecha) + 1
    Else
        'VERIFICA SI EL DIA Q SE VALORIZA ES FERIADO--AVMM--30-10-2006
        Set oCap = New COMNCaptaGenerales.NCOMCaptaDefinicion
            lnFeriado = oCap.ObtenerFeriado(dFecha, dFecha)
        Set oCap = Nothing
        dFecha = DateAdd("d", CDate(dFecha), lnFeriado)
        dFecha = CDate(dFecha) + 1
    End If
    'VERIFICA SI EL DIA Q SE VALORIZA ES FERIADO--AVMM--30-10-2006
    Set oCap = New COMNCaptaGenerales.NCOMCaptaDefinicion
        lnFeriado = oCap.ObtenerFeriado(dFecha, dFecha)
    Set oCap = Nothing
    dFecha = DateAdd("d", CDate(dFecha), lnFeriado)

    txtIngChqFechaVal.Text = Format$(dFecha, "dd/mm/yyyy")
End If
End Sub

Private Sub cboMoneda_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxtBCodPers.SetFocus
End If
End Sub

Private Sub chkConfirmar_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cboIngChqPlaza.SetFocus
End If
End Sub
Private Sub cmdAceptar_Click()
Dim oDocRec As COMNCajaGeneral.NCOMDocRec
Dim oCaja As COMNCajaGeneral.NCOMCajaGeneral 'nCajaGeneral
Dim ValtxtIngChqNumCheque As String
Dim ValTottxtIngChqNumCheque As String
Dim lsMovNro As String
Dim oArendir As COMNCajaGeneral.NCOMARendir 'NARendir
Dim rs As ADODB.Recordset
Dim lnConfCaja As COMDConstantes.CGEstadoConfCheque 'CGEstadoConfCheque
Dim lsCtaDebe As String

Dim oCons As COMDConstantes.DCOMConstantes
Dim lnFeriado As Integer
ValtxtIngChqNumCheque = ""
ValTottxtIngChqNumCheque = ""
Set oContFunct = New COMNContabilidad.NCOMContFunciones
Set oCaja = New COMNCajaGeneral.NCOMCajaGeneral
Set rs = New ADODB.Recordset

If lbMuestra = False Then
    If ValidaInterfaz = False Then Exit Sub
    ValtxtIngChqNumCheque = txtIngChqNumCheque & "-" & txtIngChqNumCheque1
    ValTottxtIngChqNumCheque = txtIngChqNumCheque & "-" & txtIngChqNumCheque1 & "-" & txtIngChqNumCheque2 & "-" & txtIngChqNumCheque3 & "-" & txtIngChqCtaIF & "-" & txtIngChqNumCheque4
    If lbSoloIngreso = False Then
        Set oDocRec = New COMNCajaGeneral.NCOMDocRec
        Set oArendir = New COMNCajaGeneral.NCOMARendir
        If MsgBox("Desea Realizar el Registro del Cheque??", vbYesNo + vbInformation, "Aviso") = vbYes Then
            lsMovNro = oContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
            lnDiaValoriza = GetDiasMinValorizacion(lnMoneda, lnPlazaChq)
            If lbNegocio = True Then
                Set oCons = New COMDConstantes.DCOMConstantes
                    lnFeriado = oCons.ObtenerFeridos(Me.txtIngChqFechaReg.Text, Me.txtIngChqFechaVal.Text)
                    Me.txtIngChqFechaVal.Text = CDate(Me.txtIngChqFechaVal.Text)
                Set oCons = Nothing
                
                If chkConfirmar.value Then
                    lnConfCaja = ChqCGNoConfirmado
                Else
                    lnConfCaja = ChqCGSinConfirmacion
                End If
             
                'MADM 20110628 Mid(txtIngChqBuscaIF, 4, 13) x Trim(cCodPerBco)
                oDocRec.RegistroChequesNegocio lsMovNro, Trim(gChqOpeRegistro), txtMovDesc, Mid(Trim(cCodPerBco), 1, 2), ValtxtIngChqNumCheque, _
                                                 Mid(Trim(cCodPerBco), 4, 13), cboIngChqPlaza.ListIndex, txtIngChqCtaIF, _
                                                 CDbl(txtMonto.Text), CDate(txtIngChqFechaReg), CDate(txtIngChqFechaVal), gsFormatoFecha, Right(cboMoneda, 1), , , lnConfCaja, gsCodArea, gsCodAge, nProducto, TxtBCodPers.Text, ValTottxtIngChqNumCheque
                        
                
            Else
                If lbArendir = True Then
                   'MADM 20110628 Mid(txtIngChqBuscaIF, 4, 13) x Trim(cCodPerBco)
                    oArendir.GrabaRendicionIngresoCheque lnTipoArendir, gsFormatoFecha, lsMovNro, lsOpeCod, txtMovDesc, _
                                lsCtaContHaber, lsCtaContChq, txtBuscarProd.Text, Mid(txtBuscarAreaAgencia.Text, 1, 3), Mid(txtBuscarAreaAgencia.Text, 4, 2), _
                                CDbl(txtMonto.Text), lsMovNroAtenc, lsMovNroSol, Trim(ValtxtIngChqNumCheque), _
                                Mid(Trim(cCodPerBco), 4, 13), Mid(Trim(cCodPerBco), 1, 2), cboIngChqPlaza.ListIndex, txtIngChqCtaIF, CDate(txtIngChqFechaReg), _
                                CDate(txtIngChqFechaVal), Right(cboMoneda, 1), gChqEstEnValorizacion, gCGEstadosChqRecibido, ChqCGSinConfirmacion, Mid(txtBuscarAreaAgencia, 1, 3), Mid(txtBuscarAreaAgencia, 4, 2), TxtBCodPers.Text, ValTottxtIngChqNumCheque
                                         
                Else
                    If fgObjMotivo.TextMatrix(1, 0) <> "" Then
                        Set rs = fgObjMotivo.GetRsNew
                    End If
                    
                    lsCtaDebe = oContFunct.GetFiltroObjetos(ObjProductosCMACT, lsCtaContChq, txtBuscarProd, False)
                    lsCtaDebe = lsCtaDebe + oContFunct.GetFiltroObjetos(ObjCMACAgencias, lsCtaContChq, txtBuscarAreaAgencia, False)
                    lsCtaDebe = lsCtaContChq + lsCtaDebe
                    
                    If lbApertura = False Then
                        'MADM 20110628 Mid(txtIngChqBuscaIF, 4, 13) x Trim(cCodPerBco)
                        oDocRec.RegistroChequesContab lsMovNro, lsOpeCod, txtMovDesc, val(lsMovRef), lsCtaDebe, txtBuscarProd, _
                                    Mid(txtBuscarAreaAgencia.Text, 1, 3), Mid(txtBuscarAreaAgencia.Text, 4, 2), txtBuscarCtaHaber.Text, _
                                    rs, ValtxtIngChqNumCheque, Mid(Trim(cCodPerBco), 4, 13), Mid(Trim(cCodPerBco), 1, 2), cboIngChqPlaza.ListIndex, _
                                    Me.txtIngChqCtaIF, CDbl(txtMonto.value), CDate(txtIngChqFechaReg), CDate(txtIngChqFechaVal), _
                                    gsFormatoFecha, Right(cboMoneda, 1), gChqEstValorizado, gCGEstadosChqRecibido, ChqCGSinConfirmacion, Mid(txtBuscarAreaAgencia, 1, 3), Mid(txtBuscarAreaAgencia, 4, 2), TxtBCodPers.Text, ValTottxtIngChqNumCheque
                    Else
                        
                        'MADM 20110628 Mid(txtIngChqBuscaIF, 4, 13) x Trim(cCodPerBco)
                        oCaja.GrabaAperturaRegCheque lsMovNro, lsOpeCod, txtMovDesc, txtBuscarCtaHaber.Text, txtBuscarProd, _
                                    Mid(txtBuscarAreaAgencia.Text, 1, 3), Mid(txtBuscarAreaAgencia.Text, 4, 2), lsCtaDebe, _
                                    rs, ValtxtIngChqNumCheque, Mid(Trim(cCodPerBco), 4, 13), Mid(Trim(cCodPerBco), 1, 2), cboIngChqPlaza.ListIndex, _
                                    txtIngChqCtaIF, CDbl(txtMonto.value), CDate(txtIngChqFechaReg), CDate(txtIngChqFechaVal), _
                                    Right(cboMoneda, 1), _
                                    lbCreaSubCta, lsSubCtaIFCod, lsSubCtaIFDesc, lsPersCodAper, _
                                    lsIFTpoAper, lsCtaIFCod, lsCtaIFDesc, ldCtaIFAper, _
                                    lsCtaIFVenc, lnCtaIFPlazo, lnPeriodo, lnInteres, lnTpoDocAper, lsNroDocAper, _
                                    ldFechaDocApera, _
                                    gChqEstValorizado, gCGEstadosChqRecibido, ChqCGSinConfirmacion, Mid(txtBuscarAreaAgencia, 1, 3), Mid(txtBuscarAreaAgencia, 4, 2), TxtBCodPers.Text, ValTottxtIngChqNumCheque
                        Set oCaja = Nothing
                        If lsDocumentoAper <> "" Then
                                EnviaPrevio lsDocumentoAper & Chr(12) & lsDocumentoAper, "Carta Apertura", gnLinPage, False
                        End If
                    End If
                End If
                ImprimeAsientoContable lsMovNro, "", "", ""
            End If
            EmiteDatos
            Set oDocRec = Nothing
            Set oArendir = Nothing
            lbOk = True
            If MsgBox("¿Desea registrar otro cheque?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
                txtMonto.Text = "0.00"
                txtIngChqCtaIF = ""
                txtIngChqFechaVal = gdFecSis
                txtIngChqNumCheque = ""
                'MADM 20110628
                txtIngChqNumCheque1 = ""
                txtIngChqNumCheque2 = ""
                txtIngChqNumCheque3 = ""
                txtIngChqNumCheque4 = ""
                ValtxtIngChqNumCheque = ""
                lblbanco = ""
                cCodPerBco = ""
'                txtIngChqBuscaIF = ""
'                lblIngChqDescIF = ""
                'END MADM
                txtMovDesc = ""
                cboIngChqPlaza.ListIndex = 0
                cboIngChqPlaza_Click
                txtBuscarAreaAgencia = ""
'                txtIngChqBuscaIF.SetFocus
                If lbRegistroOperaciones Then
                    txtBuscarProd.Text = ""
                    lblProdDesc = ""
                End If
            Else
                Unload Me
            End If
        End If
    Else
        EmiteDatos
        lbOk = True
        Unload Me
    End If
Else
    EmiteDatos
    lbOk = True
    Unload Me
End If
Set oContFunct = Nothing
End Sub

Private Sub cmdCancelar_Click()
lbOk = False
Unload Me
End Sub

Private Sub Form_Activate()
    Dim nEditaFV As Integer

    Set oGen = New COMDConstSistema.DCOMGeneral
    nEditaFV = oGen.LeeConstSistema(108)
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
    If nEditaFV = 1 Then
        txtIngChqFechaVal.Enabled = True
    Else
        txtIngChqFechaVal.Enabled = False
    End If
    Set oGen = Nothing
End Sub

Private Sub Form_Load()

Set oCtasIF = New COMNCajaGeneral.NCOMCajaCtaIF
Set oContFunct = New COMNContabilidad.NCOMContFunciones 'NContFunciones
Set oGen = New COMDConstSistema.DCOMGeneral

Dim oCajitf As COMNCajaGeneral.NCOMCajaCtaIF  'DGeneral
Set oCajitf = New COMNCajaGeneral.NCOMCajaCtaIF

Dim oCajG As COMNCajaGeneral.NCOMCajaGeneral
Set oCajG = New COMNCajaGeneral.NCOMCajaGeneral

Dim oPersonas As COMDPersona.DCOMPersonas
Set oPersonas = New COMDPersona.DCOMPersonas

Set oOpe = New COMDConstSistema.DCOMOperacion 'DOperacion

Dim nEditaFV As Integer

'datos para Ing de cheques
CambiaTamañoCombo cboIngChqPlaza, 150
CambiaTamañoCombo cboMoneda, 100
txtIngChqFechaReg = gdFecSis
txtIngChqFechaVal = gdFecSis
Me.Caption = " Registro de Cheques - " & gsOpeDesc
 If lbApertura Then
    lsCtaContChq = oCajitf.EmiteOpeCta(lsOpeCod, "H", "1")
Else
    lsCtaContChq = oCajitf.EmiteOpeCta(lsOpeCod, "D")
End If
txtMonto.Text = Format$(lnMonto, "#,##0.00")
txtMonto.Enabled = Not lbMuestra
If lnMonto > 0 Then
    txtMonto.Enabled = True
End If

''Comentado x MADM 20100628
''txtIngChqBuscaIF.psRaiz = "BANCOS"
'txtIngChqBuscaIF.rs = oCtasIF.CargaCtasIF(0, "_" & gTpoIFBanco & "%", MuestraInstituciones)
''txtIngChqBuscaIF.rs = oCajG.GetOpeObj(lsOpeCod, "1")

FraIngCheque.Enabled = Not lbMuestra
txtMovDesc.Locked = lbMuestra
txtMovDesc = lsGlosa

CargaCombo cboMoneda, oGen.GetConstante(gMoneda)
CargaCombo cboIngChqPlaza, oGen.GetConstante(gChequePlaza)

txtBuscarProd.rs = oCajG.GetOpeObj(lsOpeCod, lnOrdenProd)
If txtBuscarProd.rs.State = adStateOpen Then
    If txtBuscarProd.rs.RecordCount = 1 Then
        txtBuscarProd.Text = txtBuscarProd.rs(0)
        lblProdDesc = txtBuscarProd.psDescripcion
        txtBuscarProd.Enabled = False
    End If
End If

Set oCtasIF = Nothing

txtBuscarAreaAgencia.rs = oCajG.GetOpeObj(lsOpeCod, lnOrdenAgencia)
If txtBuscarAreaAgencia.rs.State = adStateOpen Then
    If txtBuscarAreaAgencia.rs.RecordCount = 1 Then
        txtBuscarAreaAgencia.Text = txtBuscarAreaAgencia.rs(0)
        lblAreaAgeDesc = txtBuscarAreaAgencia.psDescripcion
        txtBuscarAreaAgencia.Enabled = False
    End If
End If
If lsCtaContHaber <> "" Then
    txtBuscarCtaHaber = lsCtaContHaber
    lblCtaHaber = oGen.CuentaNombre(lsCtaContHaber)
    txtBuscarCtaHaber.Enabled = False
Else
    txtBuscarCtaHaber.psRaiz = "Cuentas Contables"
    txtBuscarCtaHaber.rs = oOpe.CargaOpeCta(lsOpeCod, "H", "0")
    Set oOpe = Nothing
End If
txtIngChqFechaReg.Enabled = False
If lbRegistroOperaciones Then
    txtBuscarAreaAgencia.Visible = False
    lblAreaAgeDesc.Visible = False
    Label10.Visible = False
    fraEspecifica.Height = Frame2.Height + 210
    fraGlosa.Top = fraEspecifica.Top + fraEspecifica.Height + 50
    cmdAceptar.Top = fraGlosa.Top + fraGlosa.Height + 50 '6000
    cmdCancelar.Top = fraGlosa.Top + fraGlosa.Height + 50 '6000
    fraEspecifica.Visible = True
    Me.Height = cmdCancelar.Top + cmdCancelar.Height + 500
    Me.chkConfirmar.Visible = True
    
    Set oGen = New COMDConstSistema.DCOMGeneral
    nEditaFV = oGen.LeeConstSistema(108)
    
    If nEditaFV = 1 Then
        txtIngChqFechaVal.Enabled = True
    Else
        txtIngChqFechaVal.Enabled = False
    End If
    
    'MAVM 20100629 BAS II ***
    'txtBuscarProd.rs = oGen.GetConstanteArbol(gProducto, 10)
    'ALPA 20100707 BAS II *****
    txtBuscarProd.rs = oGen.GetConstanteArbolCreditoYAhorro(gTpoProducto, 10, gProducto, "23")
    '**************************
    Set oGen = Nothing
    If txtBuscarProd.rs.State = adStateOpen Then
        If txtBuscarProd.rs.RecordCount = 1 Then
            txtBuscarProd.Text = txtBuscarProd.rs(0)
            lblProdDesc = txtBuscarProd.psDescripcion
            txtBuscarProd.Enabled = False
        End If
    End If
ElseIf lbNegocio = False Then
    fraGlosa.Top = fraEspecifica.Top + fraEspecifica.Height + 50
    cmdAceptar.Top = fraGlosa.Top + fraGlosa.Height + 50 '6000
    cmdCancelar.Top = fraGlosa.Top + fraGlosa.Height + 50 '6000
    Me.fraEspecifica.Visible = True
    Me.Height = cmdCancelar.Top + cmdCancelar.Height + 500
    Me.chkConfirmar.Visible = False
Else
    fraGlosa.Top = FraIngCheque.Top + FraIngCheque.Height + 100
    cmdAceptar.Top = fraGlosa.Top + fraGlosa.Height + 50 '6000
    cmdCancelar.Top = fraGlosa.Top + fraGlosa.Height + 50 '6000
    fraEspecifica.Visible = False
    Me.Height = cmdCancelar.Top + cmdCancelar.Height + 500
    
    txtIngChqFechaVal = DateAdd("d", lnDiasValoriza, CDate(txtIngChqFechaReg))
End If
AsignaDatos
Set oCajitf = Nothing
Set oCajG = Nothing
End Sub

Public Sub InicioMuestra(ByVal psPersCodIF As String, ByVal psNroChq As String, _
                        ByVal pbNegocio As Boolean, ByVal psOpeCod As String, ByVal pnMoneda As Moneda)

lbMuestra = True
lbArendir = False
lsPersCodIf = psPersCodIF
lbApertura = False
lsNroChq = psNroChq
lsOpeCod = psOpeCod
lbNegocio = pbNegocio
cboMoneda.ListIndex = pnMoneda - 1
cboMoneda.Enabled = False
Me.Show 1
End Sub

Public Sub Inicio(ByVal pbNegocio As Boolean, ByVal psOpeCod As String, _
        ByVal pbRegCheque As Boolean, ByVal pnMonto As Currency, ByVal pnMoneda As Moneda, _
        Optional ByVal psMovref As String = "", Optional pnOrdenProd As Integer = 1, _
        Optional pnOrdenAgencia As Integer = 2, Optional pbSoloIngresa As Boolean = False, _
        Optional psGlosa As String = "", Optional pbRegistroOperaciones As Boolean = False, _
        Optional nProd As Producto)

nProducto = nProd
lbRegistroOperaciones = pbRegistroOperaciones
lbMuestra = False
lbArendir = False
lbNegocio = pbNegocio
lnMoneda = pnMoneda
lsOpeCod = psOpeCod
lbRegCheque = pbRegCheque
lsMovRef = psMovref
lnMonto = pnMonto
lnOrdenProd = pnOrdenProd
lnOrdenAgencia = pnOrdenAgencia
lbApertura = False
lbSoloIngreso = pbSoloIngresa
txtMonto.Text = "0.00"
txtIngChqCtaIF = ""
txtIngChqFechaVal = gdFecSis
txtIngChqNumCheque = ""

'MADM 20110628
txtIngChqNumCheque1 = ""
txtIngChqNumCheque2 = ""
txtIngChqNumCheque3 = ""
txtIngChqNumCheque4 = ""
'txtIngChqBuscaIF = ""
'lblIngChqDescIF = ""

txtMovDesc = ""
lsGlosa = psGlosa
cboIngChqPlaza.ListIndex = 0
If Not lbRegistroOperaciones Then
    cboMoneda.ListIndex = pnMoneda - 1
    cboMoneda.Enabled = False
Else
    cboMoneda.ListIndex = 0
End If
txtBuscarAreaAgencia = ""
Me.Show 1
End Sub

Public Sub InicioArendir(ByVal psOpeCod As String, ByVal pnMonto As Currency, ByVal pnTipoArendir As ArendirTipo, _
                        ByVal psMovNroAtenc As String, ByVal psMovNroSol As String, ByVal psCtaContPendiente As String, ByVal psGlosa As String, ByVal pnMoneda As Moneda, _
                        Optional pnOrdenProd As Integer = 1, Optional pnOrdenAgencia As Integer = 2)
lbMuestra = False
lbArendir = True
lsOpeCod = psOpeCod
lnMoneda = pnMoneda
lsMovNroAtenc = psMovNroAtenc
lsMovNroSol = psMovNroSol
lnTipoArendir = pnTipoArendir
lsCtaContHaber = psCtaContPendiente
lsGlosa = psGlosa
lnMonto = pnMonto
lnOrdenProd = pnOrdenProd
lnOrdenAgencia = pnOrdenAgencia
lbApertura = False
Me.Show 1
End Sub

Public Sub InicioAperturas(ByVal psOpeCod As String, ByVal pnMonto As Currency, _
                          ByVal psCtaHaber As String, ByVal psGlosa As String, ByVal pnMoneda As Moneda, _
                          ByVal pbCreaSubCta As Boolean, ByVal psSubCtaIFCod As String, ByVal psSubCtaIFDesc As String, _
                          ByVal psPersCod As String, ByVal psIFTpo As String, _
                          ByVal psCtaIFCod As String, ByVal psCtaIFDesc As String, _
                          ByVal pdCtaIFAper As Date, ByVal psCtaIFVenc As String, _
                          ByVal pnCtaIFPlazo As Integer, ByVal pnPeriodo As Integer, ByVal pnInteres As Currency, _
                          ByVal pnTpoDocAper As TpoDoc, ByVal psNroDocAper As String, _
                          ByVal pdFechaDocApera As String, ByVal psDocumentoAper As String, _
                          Optional pnOrdenProd As Integer = 1, Optional pnOrdenAgencia As Integer = 2)


lbCreaSubCta = pbCreaSubCta
lsSubCtaIFCod = psSubCtaIFCod
lsSubCtaIFDesc = psSubCtaIFDesc
lsPersCodAper = psPersCod
lsIFTpoAper = psIFTpo
lsCtaIFCod = psCtaIFCod
lsCtaIFDesc = psCtaIFDesc
ldCtaIFAper = pdCtaIFAper
lsCtaIFVenc = psCtaIFVenc
lnCtaIFPlazo = pnCtaIFPlazo
lnTpoDocAper = pnTpoDocAper
lsNroDocAper = psNroDocAper
ldFechaDocApera = pdFechaDocApera
lnTpoDocAper = pnTpoDocAper
lsNroDocAper = psNroDocAper
ldFechaDocApera = pdFechaDocApera
lsDocumentoAper = psDocumentoAper
lnPeriodo = pnPeriodo
lnInteres = pnInteres

lbMuestra = False
lsOpeCod = psOpeCod
lnMoneda = pnMoneda
lbApertura = True
lsCtaContHaber = psCtaHaber
lsGlosa = psGlosa
lnMonto = pnMonto
lnOrdenProd = pnOrdenProd
lnOrdenAgencia = pnOrdenAgencia

Me.Show 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set oCtasIF = Nothing
    Set oContFunct = Nothing
    Set oGen = Nothing
    Set oOpe = Nothing
End Sub
Private Function ValidaInterfaz() As Boolean
Dim oDoc As COMDCajaGeneral.DCOMDocumento

Set oDoc = New COMDCajaGeneral.DCOMDocumento

ValidaInterfaz = True
'If Len(Trim(txtIngChqBuscaIF)) = 0 Or Len(Trim(lblIngChqDescIF)) = 0 Then
'    MsgBox "Institución Financiera no válida", vbInformation, "Aviso"
'    ValidaInterfaz = False
'    txtIngChqBuscaIF.SetFocus
'    Exit Function
'End If
'MADM 20110628
If Trim(lblbanco) = "" Then
    MsgBox "Valide la Institución Financiera, Haga Enter en Bco !!", vbInformation, "Aviso"
    ValidaInterfaz = False
    txtIngChqNumCheque2.SetFocus
    Exit Function
End If
'END MADM
If Len(Trim(txtIngChqCtaIF)) = 0 Then
    MsgBox "Nro de Cuenta de Institución Financiera no válida", vbInformation, "Aviso"
    ValidaInterfaz = False
    txtIngChqCtaIF.SetFocus
    Exit Function
End If
If Len(Trim(txtIngChqNumCheque)) = 0 And Len(Trim(txtIngChqNumCheque1)) = 0 Then
    MsgBox "Nro de Cheque no Ingresado o no es válido", vbInformation, "Aviso"
    ValidaInterfaz = False
    txtIngChqNumCheque.SetFocus
    Exit Function
End If

'MADM 20110224
 If Len(Trim(TxtBCodPers.Text)) = 0 Then
    MsgBox "Debe Completar los datos del Girador", vbInformation, "Aviso"
    ValidaInterfaz = False
    TxtBCodPers.SetFocus
    Exit Function
 End If
'END

'MADM 20110628 - txtIngChqNumCheque & "-" & txtIngChqNumCheque1 - Mid(txtIngChqBuscaIF, 4, 13)
If oDoc.VerificaDoc(TpoDocCheque, txtIngChqNumCheque & "-" & txtIngChqNumCheque1, Mid(Trim(cCodPerBco), 4, 13)) Then
    MsgBox "Documento ya se encuentra registrado ", vbInformation, "Aviso"
    ValidaInterfaz = False
    txtIngChqNumCheque.SetFocus
    Exit Function
End If

If oDoc.VerificaCheque(TpoDocCheque, txtIngChqNumCheque & "-" & txtIngChqNumCheque1, Mid(Trim(cCodPerBco), 4, 13), Mid(cCodPerBco, 1, 2)) Then
    MsgBox "Cheque ya se encuentra registrado ", vbInformation, "Aviso"
    ValidaInterfaz = False
    txtIngChqNumCheque.SetFocus
    Exit Function
End If
'END MADM

If cboIngChqPlaza = "" Then
    MsgBox "Plaza de Cheque no válido", vbInformation, "Aviso"
    ValidaInterfaz = False
    cboIngChqPlaza.SetFocus
    Exit Function
End If
If ValFecha(txtIngChqFechaReg) = False Then
    ValidaInterfaz = False
    Exit Function
End If
If ValFecha(txtIngChqFechaVal) = False Then
    ValidaInterfaz = False
    Exit Function
End If
If CDate(txtIngChqFechaVal) < CDate(txtIngChqFechaReg) Then
    MsgBox "Fecha de Valorizacion no puede ser menor a la de registro", vbInformation, "Aviso"
    txtIngChqFechaVal.SetFocus
    ValidaInterfaz = False
    Exit Function
End If
If CDate(txtIngChqFechaReg) > CDate(txtIngChqFechaVal) Then
    MsgBox "Fecha de Registro no puede ser mayor a la de valorización", vbInformation, "Aviso"
    txtIngChqFechaReg.SetFocus
    ValidaInterfaz = False
    Exit Function
End If

If CDate(txtIngChqFechaReg) > CDate(txtIngChqFechaVal) Then
    MsgBox "Fecha de Registro no puede ser mayor a la de valorización", vbInformation, "Aviso"
    txtIngChqFechaReg.SetFocus
    ValidaInterfaz = False
    Exit Function
End If
If lbNegocio Then
    If DateDiff("d", CDate(txtIngChqFechaReg), CDate(txtIngChqFechaVal)) < lnDiasValoriza Then
        MsgBox "Fecha no válida. Día(s) mínimo(s) de Valorización : [" & lnDiasValoriza & "] ", vbInformation, "Aviso"
        ValidaInterfaz = False
        txtIngChqFechaVal.SetFocus
        Exit Function
    End If
Else
    If txtBuscarProd.Text = "" Or lblProdDesc = "" Then
        MsgBox "Producto a que pertenece el documento no ingresado", vbInformation, "Aviso"
        ValidaInterfaz = False
        If txtBuscarProd.Enabled Then txtBuscarProd.SetFocus
        
        Exit Function
    End If
    If (txtBuscarAreaAgencia = "" Or lblAreaAgeDesc = "") And Not lbRegistroOperaciones Then
        MsgBox "Area/Agencia no ingresada ", vbInformation, "Aviso"
        ValidaInterfaz = False
        If txtBuscarAreaAgencia.Enabled Then txtBuscarAreaAgencia.SetFocus
        Exit Function
    End If
    If (txtBuscarCtaHaber.Text = "" Or lblCtaHaber = "") And Not lbRegistroOperaciones Then
        MsgBox "Cuenta de Haber no Ingresado", vbInformation, "Aviso"
        ValidaInterfaz = False
        If txtBuscarCtaHaber.Enabled Then txtBuscarCtaHaber.SetFocus
        Exit Function
    End If
End If
If Len(Trim(cboMoneda)) = 0 Then
    MsgBox "Moneda de Documento no Seleccionada", vbInformation, "Aviso"
    ValidaInterfaz = False
    cboMoneda.SetFocus
    Exit Function
End If

If Len(Trim(Me.txtMovDesc)) = 0 Then
    MsgBox "Glosa o Descripcion de operación no ingresada", vbInformation, "Aviso"
    ValidaInterfaz = False
    txtMovDesc.SetFocus
    Exit Function
End If
If txtMonto.value = 0 Then
    MsgBox "Monto de Registro no válido", vbInformation, "Aviso"
    ValidaInterfaz = False
    txtMonto.SetFocus
    Exit Function
End If

'ARCV 26-10-2006
If nProducto = 0 Then
    MsgBox "Debe indicar un motivo", vbInformation, "Aviso"
    ValidaInterfaz = False
End If
'------------
Set oDoc = Nothing
End Function

Private Sub TxtBCodPers_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If fraEspecifica.Visible Then
        If txtBuscarProd.Enabled Then
            txtBuscarProd.SetFocus
        Else
            If txtBuscarCtaHaber.Enabled Then
                txtBuscarCtaHaber.SetFocus
            Else
                txtMovDesc.SetFocus
            End If
        End If
    Else
        txtMovDesc.SetFocus
    End If
End If
End Sub

Private Sub txtIngChqCtaIF_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Len(txtIngChqCtaIF.Text) > 0 Then
        txtIngChqNumCheque4.SetFocus
    Else
        MsgBox "Debe ingresar Nro. de Cuenta, Verifique y presione Enter", vbInformation, "Aviso"
    End If
End If
End Sub

Private Sub txtIngChqNumCheque1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Len(txtIngChqNumCheque1.Text) > 0 Then
        txtIngChqNumCheque2.SetFocus
    Else
        MsgBox "Dígito no Válido, Verifique y presione Enter", vbInformation, "Aviso"
    End If
End If
End Sub

Private Sub txtIngChqNumCheque2_KeyPress(KeyAscii As Integer)
Dim oCajG As COMNCajaGeneral.NCOMCajaGeneral
Dim rsC As ADODB.Recordset

Set rsC = New ADODB.Recordset
Set oCajG = New COMNCajaGeneral.NCOMCajaGeneral

If KeyAscii = 13 Then
    If txtIngChqNumCheque2.Text <> "" And Len(txtIngChqNumCheque2.Text) = 3 Then
        Set rsC = oCajG.GetBancosCod(Trim(txtIngChqNumCheque2.Text))
        lblbanco = ""
        If Not rsC.EOF And Not rsC.BOF Then
          lblbanco = Trim(rsC!cNomBanco)
          cCodPerBco = rsC!cPersCod
              If cCodPerBco = "" Then
                  MsgBox "Código de Banco no Válido, Verifique y presione Enter", vbInformation, "Aviso"
                  txtIngChqNumCheque2.SetFocus
              Else
                  txtIngChqNumCheque3.SetFocus
              End If
        End If
        rsC.Close
        Set rsC = Nothing
    Else
        MsgBox "Debe indicar un Código de Banco y presione Enter", vbInformation, "Aviso"
        txtIngChqNumCheque2.SetFocus
    End If
    Set oCajG = Nothing
End If
End Sub

Private Sub TxtBCodPers_EmiteDatos()
 If Trim(TxtBCodPers.Text) = "" Then
        Exit Sub
    End If

If Cargar_Datos_Persona(Trim(TxtBCodPers.Text)) = False Then
        MsgBox "No se pudo encontrar los datos de la Persona," & Chr(10) & " Verifique que la Persona exista", vbInformation, "Aviso"
        Exit Sub
End If
    
End Sub

Function Cargar_Datos_Persona(pcPersCod As String) As Boolean
    
    Set oPersona = Nothing
    Set oPersona = New UPersona_Cli ' COMDPersona.DCOMPersona
    oPersona.sCodAge = gsCodAge
    
    Cargar_Datos_Persona = True
    
    Call oPersona.RecuperaPersona(pcPersCod, , gsCodUser)
    
    If oPersona.PersCodigo = "" Then
        Cargar_Datos_Persona = False
        Exit Function
    Else
        lblIngChqDescIF1 = oPersona.NombreCompleto
    End If
End Function

Private Sub txtBuscarAreaAgencia_EmiteDatos()
lblAreaAgeDesc = txtBuscarAreaAgencia.psDescripcion
If txtBuscarCtaHaber.Visible And txtBuscarCtaHaber.Enabled Then
    txtBuscarCtaHaber.SetFocus
End If
End Sub

Private Sub txtBuscarAreaAgencia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If txtBuscarCtaHaber.Enabled Then
        txtBuscarCtaHaber.SetFocus
    Else
        txtMovDesc.SetFocus
    End If
End If
End Sub

Private Sub txtBuscarCtaHaber_EmiteDatos()
Set oContFunct = New COMNContabilidad.NCOMContFunciones
lblCtaHaber = oContFunct.EmiteCtaContDesc(txtBuscarCtaHaber)
If txtBuscarCtaHaber.Text <> "" Then
    AsignaCtaObj txtBuscarCtaHaber.Text
    If txtMovDesc.Visible And txtMovDesc.Enabled Then
        txtMovDesc.SetFocus
    End If
End If
Set oContFunct = Nothing
End Sub

Private Sub txtBuscarCtaHaber_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtMovDesc.SetFocus
End If
End Sub
Private Sub txtBuscarProd_EmiteDatos()
If txtBuscarProd.Text <> "" Then
    lblProdDesc = txtBuscarProd.psDescripcion
    nProducto = CLng(txtBuscarProd.Text)
    If txtBuscarAreaAgencia.Visible And txtBuscarAreaAgencia.Enabled Then
        txtBuscarAreaAgencia.SetFocus
    ElseIf txtMovDesc.Visible And txtMovDesc.Enabled Then
        txtMovDesc.SetFocus
    End If
End If
End Sub

Private Sub txtBuscarProd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If txtBuscarAreaAgencia.Visible And txtBuscarAreaAgencia.Enabled Then
        txtBuscarAreaAgencia.SetFocus
    ElseIf txtMovDesc.Visible And txtMovDesc.Enabled Then
        txtMovDesc.SetFocus
    End If
End If
End Sub

'COMENTADO X MADM 20110628
'Private Sub txtIngChqBuscaIF_EmiteDatos()
'lblIngChqDescIF = txtIngChqBuscaIF.psDescripcion
'If txtIngChqBuscaIF.psDescripcion <> "" Then
'    txtIngChqCtaIF.SetFocus
' End If
'End Sub

'Private Sub txtIngChqCtaIF_KeyPress(KeyAscii As Integer)
'KeyAscii = NumerosEnteros(KeyAscii)
'If KeyAscii = 13 Then
'    txtIngChqNumCheque.SetFocus
'End If
'End Sub

Private Sub cboIngChqPlaza_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If txtIngChqFechaReg.Enabled Then
        txtIngChqFechaReg.SetFocus
    ElseIf txtIngChqFechaVal.Enabled Then
        txtIngChqFechaVal.SetFocus
    ElseIf cboMoneda.Enabled Then
        cboMoneda.SetFocus
    Else
        txtMovDesc.SetFocus
    End If
End If
End Sub
Private Sub txtIngChqFechaReg_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If txtIngChqFechaVal.Enabled Then
        txtIngChqFechaVal.SetFocus
    End If
End If
End Sub

Private Sub txtIngChqFechaVal_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If cboMoneda.Enabled Then
        cboMoneda.SetFocus
    Else
        txtMovDesc.SetFocus
    End If
    
End If
End Sub
Private Sub txtIngChqNumCheque_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
If KeyAscii = 13 Then
    txtIngChqNumCheque1.SetFocus
'    If chkConfirmar.Visible Then
'        chkConfirmar.SetFocus
'    Else
'        cboIngChqPlaza.SetFocus
'    End If
End If
End Sub
Public Property Get OK() As Boolean
OK = lbOk
End Property
Public Property Let OK(ByVal vNewValue As Boolean)
lbOk = vNewValue
End Property
Public Property Get PersCodIF() As String
PersCodIF = lsPersCodIf
End Property

Public Property Let PersCodIF(ByVal vNewValue As String)
lsPersCodIf = vNewValue
End Property

Public Property Get NroCtaIf() As String
NroCtaIf = lsNroCtaIf
End Property
Public Property Let NroCtaIf(ByVal vNewValue As String)
lsNroCtaIf = vNewValue
End Property
Public Property Get NroChq() As String
NroChq = lsNroChq
End Property
Public Property Let NroChq(ByVal vNewValue As String)
lsNroChq = vNewValue
End Property
Public Property Get PlazaChq() As String
PlazaChq = lnPlazaChq
End Property
Public Property Let PlazaChq(ByVal vNewValue As String)
lnPlazaChq = vNewValue
End Property
Public Property Get FechaRegChq() As Date
FechaRegChq = ldFechaRegChq
End Property
Public Property Let FechaRegChq(ByVal vNewValue As Date)
ldFechaRegChq = vNewValue
End Property
Public Property Get FechaValChq() As Date
FechaValChq = ldFechaValChq
End Property
Public Property Let FechaValChq(ByVal vNewValue As Date)
FechaValChq = vNewValue
End Property
Public Property Get ConfCheque() As String
ConfCheque = lsConfCheque
End Property
Public Property Let ConfCheque(ByVal vNewValue As String)
lsConfCheque = vNewValue
End Property
Public Property Get CtaContChq() As String
CtaContChq = lsCtaContChq
End Property
Public Property Let CtaContChq(ByVal vNewValue As String)
lsCtaContChq = vNewValue
End Property
Public Property Get Glosa() As String
Glosa = lsGlosa
End Property
Public Property Let Glosa(ByVal vNewValue As String)
lsGlosa = vNewValue
End Property
Public Property Get NombreIF() As String
NombreIF = lsNombreIF
End Property
Public Property Let NombreIF(ByVal vNewValue As String)
NombreIF = vNewValue
End Property

Private Sub txtIngChqNumCheque3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Len(txtIngChqNumCheque3.Text) > 0 Then
        txtIngChqCtaIF.SetFocus
    Else
        MsgBox "Agencia no Válida, Verifique y presione Enter", vbInformation, "Aviso"
    End If
End If
End Sub

Private Sub txtIngChqNumCheque4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If chkConfirmar.Visible Then
             chkConfirmar.SetFocus
        Else
             cboIngChqPlaza.SetFocus
        End If
    End If
End Sub

Private Sub txtMonto_GotFocus()
txtMonto.MarcaTexto
End Sub

Private Sub txtMonto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdAceptar.SetFocus
End If
End Sub

Private Sub txtMovDesc_KeyPress(KeyAscii As Integer)
KeyAscii = fgIntfMayusculas(KeyAscii)
If KeyAscii = 13 Then
    KeyAscii = 0
    If Not txtMonto.Enabled Then
        txtMonto.Enabled = True
        txtMonto.SetFocus
    Else
        txtMonto.SetFocus
    End If
End If
End Sub
Private Sub EmiteDatos()
'lsPersCodIf = txtIngChqBuscaIF
lsNroCtaIf = txtIngChqCtaIF
lsNroChq = txtIngChqNumCheque
lnPlazaChq = cboIngChqPlaza.ListIndex
ldFechaRegChq = CDate(txtIngChqFechaReg)
ldFechaValChq = CDate(txtIngChqFechaVal)
lsConfCheque = chkConfirmar.value
lsGlosa = Trim(txtMovDesc)
'lsNombreIF = lblIngChqDescIF
lnImporte = txtMonto.value
Set rsObj = fgObjMotivo.GetRsNew
lsProductoCod = Trim(txtBuscarProd)
lsAreaAgeCod = Trim(txtBuscarAreaAgencia)
lsCtaMotivo = Trim(txtBuscarCtaHaber)
lnMoneda = val(Right(cboMoneda, 2))
End Sub
Private Sub AsignaDatos()
Dim oDocRec As COMNCajaGeneral.NCOMDocRec
Dim rs As New ADODB.Recordset
Set oDocRec = New COMNCajaGeneral.NCOMDocRec
If lsPersCodIf = "" Or lsNroChq = "" Then Exit Sub
Set rs = oDocRec.GetDatosCheques(lsNroChq, lsPersCodIf)
If Not rs.EOF And Not rs.BOF Then
'Comentado x MADM 20110628
'    txtIngChqBuscaIF = lsPersCodIf
'    lblIngChqDescIF = Trim(rs!IFNOMBRE)
    txtIngChqCtaIF = rs!cIFCta
    txtIngChqNumCheque = lsNroChq
    cboIngChqPlaza.ListIndex = val(rs!bPlaza)
    txtIngChqFechaReg = Mid(rs!cMovNro, 7, 2) & "/" & Mid(rs!cMovNro, 5, 2) & "/" & Left(rs!cMovNro, 4)
    txtIngChqFechaVal = rs!dValorizaRef
    chkConfirmar.value = rs!nConfCaja
    txtMovDesc = rs!cMovDesc
    lblEstado = rs!cEstado
    txtMonto.Text = Format$(rs!nMonto, "#,##0.00")
End If
rs.Close
Set rs = Nothing
Set oDocRec = Nothing
End Sub
Private Sub AsignaCtaObj(ByVal psCtaContCod As String)
Dim Sql As String
Dim rs As ADODB.Recordset
Dim rs1 As ADODB.Recordset
Dim lsRaiz As String
Dim oDescObj As ClassDescObjeto
Dim UP As COMDPersona.UCOMPersona
Dim lsFiltro As String
Dim oRHAreas As COMDConstantes.DCOMActualizaDatosArea 'DActualizaDatosArea
Dim oCtaCont As COMDContabilidad.DCOMCtaCont 'DCtaCont
Dim oCtaIf As COMNCajaGeneral.NCOMCajaCtaIF 'NCajaCtaIF
Dim oEfect As COMDCajaGeneral.DCOMEfectivo 'Defectivo

Set oEfect = New COMDCajaGeneral.DCOMEfectivo
Set oCtaIf = New COMNCajaGeneral.NCOMCajaCtaIF
Set oRHAreas = New COMDConstantes.DCOMActualizaDatosArea
Set oDescObj = New ClassDescObjeto
Set oCtaCont = New COMDContabilidad.DCOMCtaCont
Set rs = New ADODB.Recordset
Set rs1 = New ADODB.Recordset

Me.fgObjMotivo.Clear
Me.fgObjMotivo.FormaCabecera
Me.fgObjMotivo.Rows = 2
Set rs1 = oCtaCont.CargaCtaObj(psCtaContCod, , True)
If Not rs1.EOF And Not rs1.BOF Then
    Do While Not rs1.EOF
        lsRaiz = ""
        lsFiltro = ""
        Select Case val(rs1!cObjetoCod)
            Case ObjCMACAgencias
                Set rs = oRHAreas.getAgencias(rs1!cCtaObjFiltro)
            Case ObjCMACAgenciaArea
                lsRaiz = "Unidades Organizacionales"
                Set rs = oRHAreas.GetAgenciasAreas(rs1!cCtaObjFiltro)
            Case ObjCMACArea
                Set rs = oRHAreas.GetAreas(rs1!cCtaObjFiltro)
            Case ObjEntidadesFinancieras
                lsRaiz = "Cuentas de Entidades Financieras"
                'Set rs = oCtaIf.GetCtasInstFinancieras(rs1!cCtaObjFiltro, psCtaContCod)
                Set rs = oCtaIf.CargaCtasIF(Mid(psCtaContCod, 3, 1), rs1!cCtaObjFiltro)
            Case ObjDescomEfectivo
                lsRaiz = "Denominación"
                Set rs = oEfect.GetBilletajes(rs1!cCtaObjFiltro)
            Case objPersona
                Set rs = Nothing
            Case Else
                Set rs = GetObjetos(val(rs1!cObjetoCod))
        End Select
        Set oRHAreas = Nothing
        If Not rs Is Nothing Then
            If rs.State = adStateOpen Then
                If Not rs.EOF And Not rs.BOF Then
                    If rs.RecordCount > 1 Then
                        oDescObj.Show rs, "", lsRaiz
                        If oDescObj.lbOk Then
                            Set oContFunct = New COMNContabilidad.NCOMContFunciones
                                lsFiltro = oContFunct.GetFiltroObjetos(val(rs1!cObjetoCod), psCtaContCod, oDescObj.gsSelecCod, False)
                            Set oContFunct = Nothing
                            fgObjMotivo.AdicionaFila
                            fgObjMotivo.TextMatrix(fgObjMotivo.Row, 1) = oDescObj.gsSelecCod
                            fgObjMotivo.TextMatrix(fgObjMotivo.Row, 2) = oDescObj.gsSelecDesc
                            fgObjMotivo.TextMatrix(fgObjMotivo.Row, 3) = lsFiltro
                            fgObjMotivo.TextMatrix(fgObjMotivo.Row, 4) = rs1!cObjetoCod
                        Else
                            txtBuscarCtaHaber = ""
                            lblCtaHaber = ""
                            Exit Do
                        End If
                    Else
                        fgObjMotivo.AdicionaFila
                        fgObjMotivo.TextMatrix(fgObjMotivo.Row, 1) = rs1!cObjetoCod
                        fgObjMotivo.TextMatrix(fgObjMotivo.Row, 2) = rs1!cObjetoDesc
                        fgObjMotivo.TextMatrix(fgObjMotivo.Row, 3) = lsFiltro
                        fgObjMotivo.TextMatrix(fgObjMotivo.Row, 4) = rs1!cObjetoCod
                    End If
                End If
            End If
        Else
            If val(rs1!cObjetoCod) = objPersona Then
                Set UP = frmBuscaPersona.Inicio
                If Not UP Is Nothing Then
                    fgObjMotivo.AdicionaFila
                    fgObjMotivo.TextMatrix(fgObjMotivo.Row, 1) = UP.sPersCod
                    fgObjMotivo.TextMatrix(fgObjMotivo.Row, 2) = UP.sPersNombre
                    fgObjMotivo.TextMatrix(fgObjMotivo.Row, 3) = ""
                    fgObjMotivo.TextMatrix(fgObjMotivo.Row, 4) = rs1!cObjetoCod
                End If
            End If
        End If
        rs1.MoveNext
    Loop
End If
rs1.Close
Set rs1 = Nothing
Set oDescObj = Nothing
Set UP = Nothing
Set oCtaCont = Nothing
Set oCtaIf = Nothing
Set oEfect = Nothing
End Sub
Public Property Get Importe() As Double
Importe = lnImporte
End Property
Public Property Let Importe(ByVal vNewValue As Double)
lnImporte = vNewValue
End Property
Public Property Get rsObjMotivo() As ADODB.Recordset
Set rsObjMotivo = rsObj
End Property
Public Property Set rsObjMotivo(ByVal vNewValue As ADODB.Recordset)
Set rsObj = vNewValue
End Property
Public Property Get ProductoCod() As String
ProductoCod = lsProductoCod
End Property
Public Property Let ProductoCod(ByVal vNewValue As String)
lsProductoCod = vNewValue
End Property
Public Property Get AreaAgeCod() As String
AreaAgeCod = lsAreaAgeCod
End Property
Public Property Let AreaAgeCod(ByVal vNewValue As String)
lsAreaAgeCod = vNewValue
End Property
Public Property Get CtaMotivo() As String
CtaMotivo = lsCtaMotivo
End Property
Public Property Let CtaMotivo(ByVal vNewValue As String)
lsCtaMotivo = vNewValue
End Property
Public Property Get Moneda() As Moneda
Moneda = lnMoneda
End Property
Public Property Let Moneda(ByVal vNewValue As Moneda)
lnMoneda = vNewValue
End Property

Public Property Get nDiasValorizacion() As Date
nDiasValorizacion = lnDiaValoriza
End Property
