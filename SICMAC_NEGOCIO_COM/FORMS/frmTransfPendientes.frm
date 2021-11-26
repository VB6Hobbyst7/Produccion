VERSION 5.00
Begin VB.Form frmTransfpendientes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Transferencias Pendientes de Regularizar"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9450
   Icon            =   "frmTransfPendientes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   9450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   8265
      TabIndex        =   2
      Top             =   3420
      Width           =   1125
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   345
      Left            =   7095
      TabIndex        =   1
      Top             =   3420
      Width           =   1125
   End
   Begin SICMACT.FlexEdit flex 
      Height          =   3285
      Left            =   30
      TabIndex        =   0
      Top             =   105
      Width           =   9375
      _ExtentX        =   14975
      _ExtentY        =   5794
      Cols0           =   9
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "#-nMovNro-Fecha-cPersCod-Institucion-Importe_Soles-Importe_Dolares-Documento-Referencia"
      EncabezadosAnchos=   "500-0-2500-0-3000-1300-1300-1200-4000"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-X-X-X-X-X-X-X-X"
      TextStyleFixed  =   4
      ListaControles  =   "0-0-0-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-C-L-C-L-R-R-R-L"
      FormatosEdit    =   "0-0-0-0-0-2-2-3-0"
      TextArray0      =   "#"
      SelectionMode   =   1
      lbUltimaInstancia=   -1  'True
      lbPuntero       =   -1  'True
      Appearance      =   0
      ColWidth0       =   495
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
End
Attribute VB_Name = "frmTransfpendientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lnMovNro As Long
Dim lnMonto As Currency
Dim lsGlosa As String
Dim lsBanco As String
Dim lsDoc As String
Dim lnMoneda As Moneda

Public Function Ini(pnMoneda As Moneda, pnMonto As Currency, psGlosa As String, psBanco As String, psDoc As String) As Long
    lnMoneda = pnMoneda
    Me.Show 1
    pnMonto = lnMonto
    Ini = lnMovNro
    psGlosa = lsGlosa
    psBanco = lsBanco
    psDoc = lsDoc
End Function


Private Sub CmdAceptar_Click()
    If Me.flex.Row = -1 Then
        MsgBox "Debe escoger una tranferencia.", vbInformation, "Aviso"
        Exit Sub
    End If
    
    If Not IsNumeric(flex.TextMatrix(Me.flex.Row, 1)) Then
        MsgBox "Debe ingresar una tranferencia.", vbInformation, "Aviso"
        Exit Sub
    End If
    
    lnMovNro = flex.TextMatrix(Me.flex.Row, 1)
    If lnMoneda = gMonedaNacional Then
        lnMonto = flex.TextMatrix(Me.flex.Row, 5)
    Else
        lnMonto = flex.TextMatrix(Me.flex.Row, 6)
    End If
    
    lsBanco = flex.TextMatrix(Me.flex.Row, 4)
    lsGlosa = flex.TextMatrix(Me.flex.Row, 8)
    lsDoc = flex.TextMatrix(Me.flex.Row, 7)
    
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    lnMovNro = -1
    lsBanco = ""
    lsGlosa = ""
    lsDoc = ""
    Unload Me
End Sub

Private Sub Form_Activate()
    Me.cmdAceptar.SetFocus
End Sub

Private Sub Form_Load()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim oCap As COMDCaptaGenerales.DCOMCaptaMovimiento
    Set oCap = New COMDCaptaGenerales.DCOMCaptaMovimiento
        Me.Icon = LoadPicture(App.path & gsRutaIcono)
        Set rs = oCap.GetTranfPendientes(lnMoneda)
   Me.flex.rsFlex = rs
   Set oCap = Nothing
End Sub
