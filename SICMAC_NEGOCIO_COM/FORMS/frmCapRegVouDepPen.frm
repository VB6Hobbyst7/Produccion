VERSION 5.00
Begin VB.Form frmCapRegVouDepPen 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Listado de Pendientes"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7305
   Icon            =   "frmCapRegVouDepPen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   7305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6120
      TabIndex        =   2
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdRelacionar 
      Caption         =   "&Relacionar"
      Height          =   375
      Left            =   5040
      TabIndex        =   1
      Top             =   2160
      Width           =   1095
   End
   Begin SICMACT.FlexEdit FEPendientes 
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7095
      _ExtentX        =   12515
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-X-X-X"
      ListaControles  =   "0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-R-C-R"
      FormatosEdit    =   "0-4-0-0"
      TextArray0      =   "Nro"
      lbUltimaInstancia=   -1  'True
      ColWidth0       =   495
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
End
Attribute VB_Name = "frmCapRegVouDepPen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************
'*** Nombre : frmCapRegVouDepPen
'*** Descripción : Formulario para mostrar las Pendientes del vouchert de Depósito.
'*** Creación : ELRO 20120530 07:40:21 PM, según OYP-RFC024-2012
'**********************************************************************************
Option Explicit

Dim frmOrigen As Form
Private fdFechaVoucher As Date
Private fsIFTpo As String
Private fsPersCod As String
Private fsCtaIFCod As String
Private fnMovImporte As Currency
Private fsMoneda As String

Public Sub iniciarListado(ByVal pdFechaVoucher As Date, _
                          ByVal psIFTpo As String, _
                          ByVal psPersCod As String, _
                          ByVal psCtaIFCod As String, _
                          ByVal pnMovImporte As Currency, _
                          ByVal psMoneda As String, _
                          ByRef pfrmOrigen As Form)
fdFechaVoucher = pdFechaVoucher
fsIFTpo = psIFTpo
fsPersCod = psPersCod
fsCtaIFCod = psCtaIFCod
fnMovImporte = pnMovImporte
fsMoneda = psMoneda
Set frmOrigen = pfrmOrigen
Me.Show 1
End Sub

Private Sub cargarListaPendientes()
    Dim oNCOMCaptaGenerales As NCOMCaptaGenerales
    Set oNCOMCaptaGenerales = New NCOMCaptaGenerales
    Dim rsPendientes As ADODB.Recordset
    Set rsPendientes = New ADODB.Recordset
    Dim i As Integer
    
    Call LimpiaFlex(FEPendientes)
    
    Set rsPendientes = oNCOMCaptaGenerales.obtenerCajaBancosOperacionesPendientes(fdFechaVoucher, _
                                                                                  fsIFTpo, _
                                                                                  fsPersCod, _
                                                                                  fsCtaIFCod, _
                                                                                  fnMovImporte, _
                                                                                  fsMoneda)

    If Not rsPendientes.BOF And Not rsPendientes.EOF Then
        i = 1
        FEPendientes.lbEditarFlex = True
        Do While Not rsPendientes.EOF
            FEPendientes.AdicionaFila
            FEPendientes.TextMatrix(i, 1) = Format(rsPendientes!nMontoOperacion, "##,###0.00")
            FEPendientes.TextMatrix(i, 2) = rsPendientes!cMovDesc
            FEPendientes.TextMatrix(i, 3) = rsPendientes!nMovNro
            i = i + 1
            rsPendientes.MoveNext
        Loop
        FEPendientes.lbEditarFlex = False
    End If
    Set rsPendientes = Nothing
    Set oNCOMCaptaGenerales = Nothing
End Sub

Private Sub cmdRelacionar_Click()
    If Trim(FEPendientes.TextMatrix(FEPendientes.Row, 3)) = "" Then Exit Sub
    frmOrigen.fnMovNroPen = CLng(FEPendientes.TextMatrix(FEPendientes.Row, 3))
    Unload Me
End Sub

Private Sub cmdSalir_Click()
    frmOrigen.fnMovNroPen = 0
    Unload Me
End Sub

Private Sub Form_Activate()
    FEPendientes.SetFocus
End Sub

Private Sub Form_Load()
    cargarListaPendientes
End Sub
