VERSION 5.00
Begin VB.Form frmChequeListaPendiente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Pendientes"
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7065
   Icon            =   "frmChequeListaPendiente.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   7065
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRelacionar 
      Caption         =   "&Relacionar"
      Height          =   375
      Left            =   4770
      TabIndex        =   1
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5880
      TabIndex        =   2
      Top             =   2160
      Width           =   1095
   End
   Begin SICMACT.FlexEdit FEPendientes 
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   3413
      Cols0           =   4
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "Nro-Monto-Glosa-nMovNroPen"
      EncabezadosAnchos=   "500-1200-5000-0"
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-X-X-X"
      ListaControles  =   "0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-R-L-R"
      FormatosEdit    =   "0-4-0-0"
      TextArray0      =   "Nro"
      lbUltimaInstancia=   -1  'True
      ColWidth0       =   495
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
End
Attribute VB_Name = "frmChequeListaPendiente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************************************************
'** Nombre : frmChequeListadoPendiente
'** Descripción : Para seleccion de pendientes de Cheques creado segun TI-ERS126-2013
'** Creación : EJVG, 20140307 11:00:00 AM
'************************************************************************************
Option Explicit
Dim fnNroMov As Long
Dim fdFechaMov As Date
Dim fsIFTpo As String
Dim fsPersCod As String
Dim fsCtaIFCod As String
Dim fnMoneda As Moneda
Dim fnMonto As Currency

Public Function Inicio(ByVal pdFechaMov As Date, ByVal psIFTpo As String, ByVal psPersCod As String, ByVal psCtaIFCod As String, ByVal pnMoneda As Moneda, ByVal pnMonto As Currency) As Long
    fdFechaMov = pdFechaMov
    fsIFTpo = psIFTpo
    fsPersCod = psPersCod
    fsCtaIFCod = psCtaIFCod
    fnMoneda = pnMoneda
    fnMonto = pnMonto
    Show 1
    Inicio = fnNroMov
End Function
Private Sub cmdRelacionar_Click()
    If FEPendientes.TextMatrix(1, 0) = "" Then
        MsgBox "No existen pendientes para relacionar con la operación", vbInformation, "Aviso"
        If cmdRelacionar.Visible And cmdRelacionar.Enabled Then cmdRelacionar.SetFocus
        Exit Sub
    End If
    fnNroMov = CLng(FEPendientes.TextMatrix(FEPendientes.row, 3))
    Unload Me
End Sub
Private Sub Form_Load()
    CargarListaPendientes
End Sub
Private Sub cmdSalir_Click()
    fnNroMov = 0
    Unload Me
End Sub
Private Sub CargarListaPendientes()
    Dim oNCOMCaptaGenerales As New NCOMCaptaGenerales
    Dim oRS As New ADODB.Recordset
    Dim row As Long
    
    On Error GoTo ErrCargarListaPendientes
    Screen.MousePointer = 11
    FormateaFlex FEPendientes
    Set oRS = oNCOMCaptaGenerales.obtenerCajaBancosOperacionesPendientes(fdFechaMov, fsIFTpo, fsPersCod, fsCtaIFCod, fnMonto, CStr(fnMoneda))
    Do While Not oRS.EOF
        FEPendientes.AdicionaFila
        row = FEPendientes.row
        FEPendientes.TextMatrix(row, 1) = Format(oRS!nMontoOperacion, gsFormatoNumeroView)
        FEPendientes.TextMatrix(row, 2) = oRS!cMovDesc
        FEPendientes.TextMatrix(row, 3) = oRS!nMovNro
        oRS.MoveNext
    Loop
    If row > 0 Then
        cmdRelacionar.Default = True
    End If
    Set oRS = Nothing
    Set oNCOMCaptaGenerales = Nothing
    Screen.MousePointer = 0
    Exit Sub
ErrCargarListaPendientes:
    Screen.MousePointer = 0
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub

