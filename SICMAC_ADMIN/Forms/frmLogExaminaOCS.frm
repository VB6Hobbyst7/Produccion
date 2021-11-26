VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmLogExaminaOCS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Titulo"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15810
   Icon            =   "frmLogExaminaOCS.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   15810
   StartUpPosition =   2  'CenterScreen
   Begin Sicmact.FlexEdit feOrden 
      Height          =   3780
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   15600
      _ExtentX        =   27517
      _ExtentY        =   6668
      Cols0           =   9
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "#-nMovNro-Número-N° OC/S Di/Pr-Fecha-Proveedor-Moneda-Importe-Observaciones"
      EncabezadosAnchos=   "400-0-1400-1500-1000-2300-1000-1200-7000"
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
      ColumnasAEditar =   "X-X-X-X-X-X-X-X-X"
      TextStyleFixed  =   3
      ListaControles  =   "0-0-0-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-C-C-C-C-L-C-R-L"
      FormatosEdit    =   "0-0-0-0-0-0-2-2-0"
      TextArray0      =   "#"
      SelectionMode   =   1
      lbUltimaInstancia=   -1  'True
      TipoBusqueda    =   7
      lbBuscaDuplicadoText=   -1  'True
      ColWidth0       =   405
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
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
      Left            =   14760
      TabIndex        =   7
      Top             =   4680
      Width           =   975
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
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
      Left            =   13800
      TabIndex        =   6
      Top             =   4680
      Width           =   975
   End
   Begin VB.Frame fraRefrescar 
      Appearance      =   0  'Flat
      Caption         =   "Busqueda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   735
      Left            =   5040
      TabIndex        =   0
      Top             =   0
      Width           =   5895
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
         Height          =   345
         Left            =   4440
         TabIndex        =   1
         Top             =   270
         Width           =   1230
      End
      Begin MSMask.MaskEdBox mskIni 
         Height          =   300
         Left            =   885
         TabIndex        =   2
         Top             =   285
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskFin 
         Height          =   300
         Left            =   3060
         TabIndex        =   3
         Top             =   285
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblFinal 
         Caption         =   "Final"
         Height          =   195
         Left            =   2355
         TabIndex        =   5
         Top             =   330
         Width           =   525
      End
      Begin VB.Label lblInicial 
         Caption         =   "Inicial"
         Height          =   195
         Left            =   210
         TabIndex        =   4
         Top             =   330
         Width           =   525
      End
   End
End
Attribute VB_Name = "frmLogExaminaOCS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fnTipo As Integer
Dim fsMatrizDatos() As String
Dim fsAreaAgeCod As String
Dim fnMoneda As Integer
Dim fsCtaContCod As String
Dim olog As DLogGeneral
Public Function Inicio(ByVal pnTipo As Integer, ByVal psAreaAgeCod As String, ByVal psCtaContCod As String) As String()
    fnTipo = pnTipo
    fsAreaAgeCod = psAreaAgeCod
    fsCtaContCod = psCtaContCod
    Me.Show 1
    Inicio = fsMatrizDatos
End Function
Private Sub cmdAceptar_Click()
    If feOrden.TextMatrix(1, 1) = "" Then Exit Sub
    fsMatrizDatos(1, 1) = feOrden.TextMatrix(feOrden.row, 1)
    fsMatrizDatos(2, 1) = feOrden.TextMatrix(feOrden.row, 2)
    fsMatrizDatos(3, 1) = feOrden.TextMatrix(feOrden.row, 5)
    Unload Me
End Sub
Private Sub cmdBuscar_Click()
    Dim nItem As Integer
    Dim nTot  As Currency

    Dim rs As ADODB.Recordset
    Dim row As Integer
    
    If fnTipo = 1 Then
        'Set rs = olog.ListaOrdenCompraxRegistroComprobante(fsAreaAgeCod, fsCtaContCod, Format(Me.mskIni.Text, "yyyyMMdd"), Format(Me.mskFin.Text, "yyyyMMdd"))
        Set rs = olog.ListaOrdenCompraxRegistroComprobante(fsAreaAgeCod, fsCtaContCod, Me.mskIni.Text, Me.mskFin.Text)
    Else
        'Set rs = olog.ListaOrdenServicioxRegistroComprobante(fsAreaAgeCod, fsCtaContCod, Format(Me.mskIni.Text, "yyyyMMdd"), Format(Me.mskFin.Text, "yyyyMMdd"))
        Set rs = olog.ListaOrdenServicioxRegistroComprobante(fsAreaAgeCod, fsCtaContCod, Me.mskIni.Text, Me.mskFin.Text)
    End If
    LimpiaFlex feOrden
    Do While Not rs.EOF
        feOrden.AdicionaFila
        row = feOrden.row
        feOrden.TextMatrix(row, 1) = rs!nMovNro
        feOrden.TextMatrix(row, 2) = rs!cDocNro
        feOrden.TextMatrix(row, 3) = rs!cDocNroODirPro
        feOrden.TextMatrix(row, 4) = Format(rs!dDocFecha, "dd/mm/yyyy")
        feOrden.TextMatrix(row, 5) = rs!cProveedorNombre
        feOrden.TextMatrix(row, 6) = rs!cMoneda
        feOrden.TextMatrix(row, 7) = Format(rs!nImporte, gsFormatoNumeroView)
        feOrden.TextMatrix(row, 8) = rs!cMovDesc
        rs.MoveNext
    Loop
    If rs.RecordCount > 0 Then
        feOrden.TabIndex = 0
        cmdAceptar.Default = True
    Else
        cmdAceptar.Default = False
    End If
    SendKeys "{Right}"
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Set olog = New DLogGeneral
    Me.Caption = "BUSQUEDA DE ORDENES DE " & IIf(fnTipo = 1, "COMPRA", "SERVICIO")
    Me.mskIni.Text = "01/" & Format(gdFecSis, "mm") & "/" & Format(gdFecSis, "yyyy")
    Me.mskFin.Text = Format(gdFecSis, gsFormatoFechaView)
    ReDim fsMatrizDatos(3, 1)
End Sub
Private Sub mskFin_GotFocus()
    mskFin.SelStart = 0
    mskFin.SelLength = 50
End Sub
Private Sub mskFin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdBuscar.SetFocus
    End If
End Sub
Private Sub mskIni_GotFocus()
    mskIni.SelStart = 0
    mskIni.SelLength = 50
End Sub
Private Sub mskIni_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
        Me.mskFin.SetFocus
    End If
End Sub
