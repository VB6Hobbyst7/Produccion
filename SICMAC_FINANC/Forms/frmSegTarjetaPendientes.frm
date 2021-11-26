VERSION 5.00
Begin VB.Form frmSegTarjetaPendientes 
   Caption         =   "Lista de Pendientes"
   ClientHeight    =   3060
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11385
   Icon            =   "frmSegTarjetaPendientes.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   11385
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   9240
      TabIndex        =   2
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   495
      Left            =   10320
      TabIndex        =   1
      Top             =   2520
      Width           =   975
   End
   Begin Sicmact.FlexEdit fePendientes 
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   4048
      Cols0           =   12
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "#-Movimiento-Glosa-Importe-Cta Contable-Cta Cont Descripción-Moneda-Cta IF-nMovNro-AreaCod-AgeCod-cObjetoCod"
      EncabezadosAnchos=   "0-2500-1900-800-1200-2200-900-1500-0-0-0-0"
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
      ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X-X"
      ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-L-L-R-C-L-C-L-C-C-C-C"
      FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0-0-0"
      TextArray0      =   "#"
      lbUltimaInstancia=   -1  'True
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
End
Attribute VB_Name = "frmSegTarjetaPendientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Mat_Pendiente() As String
Public nLogico As Integer
Public Function Inicio(ByVal psPersCod As String, ByVal psIFTpo As String, psCtaIfCod As String, ByVal pnImporte As Currency, ByVal pdFechaDeposito As Date) As Variant
    Call CargarDatos(psPersCod, psIFTpo, psCtaIfCod, pnImporte, pdFechaDeposito)
    Me.Show 1
    Inicio = Mat_Pendiente
End Function
Private Sub CargarDatos(ByVal psPersCod As String, ByVal psIFTpo As String, psCtaIfCod As String, ByVal pnImporte As Currency, ByVal pdFechaDeposito As Date)
    Dim oNSeguro As New NSeguros
    Dim rsSeguro As ADODB.Recordset
    Dim fila As Integer
    nLogico = 0
    Set rsSeguro = oNSeguro.RecuperarDepositosPendientes(psPersCod, psIFTpo, psCtaIfCod, pnImporte, pdFechaDeposito)
    If Not rsSeguro.BOF And Not rsSeguro.EOF Then
        Do While Not rsSeguro.EOF
        fePendientes.AdicionaFila
        fila = fila + 1
        fePendientes.TextMatrix(fila, 1) = rsSeguro!cMovNro
        fePendientes.TextMatrix(fila, 2) = rsSeguro!cMovDesc
        fePendientes.TextMatrix(fila, 3) = Format(rsSeguro!nMovImporte, "#,##0.00")
        fePendientes.TextMatrix(fila, 4) = rsSeguro!cCtaContCod
        fePendientes.TextMatrix(fila, 5) = rsSeguro!cCtaContDesc
        fePendientes.TextMatrix(fila, 6) = rsSeguro!Moneda
        fePendientes.TextMatrix(fila, 7) = rsSeguro!cCtaIFDesc
        fePendientes.TextMatrix(fila, 8) = rsSeguro!nMovNro
        fePendientes.TextMatrix(fila, 9) = rsSeguro!AreaCod
        fePendientes.TextMatrix(fila, 10) = rsSeguro!AgeCod
        fePendientes.TextMatrix(fila, 11) = rsSeguro!cObjetoCod
        rsSeguro.MoveNext
        nLogico = 1
        Loop
        fePendientes.row = 1
        fePendientes.col = 1
    End If
End Sub
Private Sub cmdAceptar_Click()
    Dim fila As Integer
    ReDim Mat_Pendiente(0 To 10)
    fila = fePendientes.row
    If fePendientes.TextMatrix(fila, 8) <> "" Then
        Mat_Pendiente(0) = fePendientes.TextMatrix(fila, 8)
        Mat_Pendiente(1) = fePendientes.TextMatrix(fila, 1)
        Mat_Pendiente(2) = fePendientes.TextMatrix(fila, 2)
        Mat_Pendiente(3) = fePendientes.TextMatrix(fila, 3)
        Mat_Pendiente(4) = fePendientes.TextMatrix(fila, 4)
        Mat_Pendiente(5) = fePendientes.TextMatrix(fila, 5)
        Mat_Pendiente(6) = fePendientes.TextMatrix(fila, 6)
        Mat_Pendiente(7) = fePendientes.TextMatrix(fila, 7)
        Mat_Pendiente(8) = fePendientes.TextMatrix(fila, 9)
        Mat_Pendiente(9) = fePendientes.TextMatrix(fila, 10)
        Mat_Pendiente(10) = fePendientes.TextMatrix(fila, 11)
        nLogico = 2
    Else
        MsgBox "Tiene que elegir una fila, o nay ninguna pendinte disponible con los parametros ingresados"
        Exit Sub
    End If
    Unload Me
End Sub
Private Sub cmdSalir_Click()
    Unload Me
End Sub
