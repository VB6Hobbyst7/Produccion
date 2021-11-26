VERSION 5.00
Begin VB.Form frmCapAperturaListaChq 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7680
   Icon            =   "frmCapAperturaListaChq.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   7680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   375
      Left            =   60
      TabIndex        =   1
      Top             =   4290
      Width           =   1035
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6555
      TabIndex        =   3
      Top             =   4290
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   5415
      TabIndex        =   2
      Top             =   4290
      Width           =   1035
   End
   Begin VB.Frame fraCheque 
      Caption         =   "Cheques Válidos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   4095
      Left            =   60
      TabIndex        =   4
      Top             =   120
      Width           =   7545
      Begin SICMACT.FlexEdit grdCheque 
         Height          =   3735
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   7320
         _ExtentX        =   12912
         _ExtentY        =   6588
         Cols0           =   8
         HighLight       =   1
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Banco-Cheque-Monto-cPersCod-Valorizacion-DiasValorizacion-sIFCta"
         EncabezadosAnchos=   "350-2800-1400-1200-0-1200-0-0"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X-X"
         TextStyleFixed  =   3
         ListaControles  =   "0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-R-C-C-C-C"
         FormatosEdit    =   "0-0-0-2-0-0-0-0"
         TextArray0      =   "#"
         SelectionMode   =   1
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   345
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
End
Attribute VB_Name = "frmCapAperturaListaChq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim frmOrigen As Form
Dim nOperacion As COMDConstantes.CaptacOperacion
Dim nMoneda As Moneda
Dim nProducto As COMDConstantes.Producto
Dim bIFCuenta As Boolean

Private Sub CorrigeFormatoFecha()
Dim i As Integer
Dim sFecha As String
Dim nMonto As Double
For i = 1 To grdCheque.Rows - 1
    sFecha = grdCheque.TextMatrix(i, 5)
    If sFecha <> "" Then
        nMonto = CDbl(grdCheque.TextMatrix(i, 3))
        grdCheque.TextMatrix(i, 5) = Format(CDate(sFecha), "dd/mm/yyyy")
        grdCheque.TextMatrix(i, 3) = Format(nMonto, "#,##0.00")
    End If
Next i
End Sub

Public Sub Inicia(ByRef frmForm As Form, ByVal nOpe As CaptacOperacion, _
        ByVal nMon As Moneda, ByVal nProd As Producto, Optional ByVal bIFCta As Boolean = False)
Me.Caption = "Listado de Cheques Válidos"
Set frmOrigen = frmForm
nOperacion = nOpe
nMoneda = nMon
nProducto = nProd
bIFCuenta = bIFCta
GetCheques
Me.Show 1
End Sub

Private Sub GetCheques()
Dim clsDoc As COMNCajaGeneral.NCOMDocRec 'nDocRec
Dim rsChq As New ADODB.Recordset
Dim sNumChq As String * 15
Set clsDoc = New COMNCajaGeneral.NCOMDocRec

If nOperacion = "200252" Then
    Set rsChq = clsDoc.GetChequesCreditos(gdFecSis, nMoneda)
Else
    Set rsChq = clsDoc.GetChequesValidos(gdFecSis, nMoneda)
End If

Set clsDoc = Nothing

If Not (rsChq.EOF And rsChq.BOF) Then
    Set grdCheque.Recordset = rsChq
    CorrigeFormatoFecha
    cmdNuevo.Enabled = True
    cmdAceptar.Enabled = True
    cmdSalir.Enabled = True
    fraCheque.Enabled = True
Else
    MsgBox "No existen cheques disponible para la operación. Debe registrar el cheque", vbInformation, "Aviso"
    fraCheque.Enabled = False
    cmdNuevo.Enabled = True
    cmdAceptar.Enabled = False
    cmdSalir.Enabled = True
End If
End Sub

Private Sub CmdAceptar_Click()
Dim nFila As Long
nFila = grdCheque.Row
frmOrigen.lblNroDoc = grdCheque.TextMatrix(nFila, 2)
frmOrigen.lblNombreIF = grdCheque.TextMatrix(nFila, 1)
frmOrigen.txtMonto.Text = Format$(CDbl(grdCheque.TextMatrix(nFila, 3)), "#,##0.00")
frmOrigen.sCodIF = Right(grdCheque.TextMatrix(nFila, 4), 13)
frmOrigen.dFechaValorizacion = CDate(grdCheque.TextMatrix(nFila, 5))
frmOrigen.lnDValoriza = grdCheque.TextMatrix(nFila, 6)
If bIFCuenta Then
    frmOrigen.sIFCuenta = grdCheque.TextMatrix(nFila, 7)
End If
Unload Me
End Sub

Private Sub cmdNuevo_Click()
frmIngCheques.Inicio True, Trim(nOperacion), True, 0, nMoneda, , , , , , , nProducto
If frmIngCheques.Ok Then
    Dim nFila As Long
    grdCheque.AdicionaFila
    nFila = grdCheque.Rows - 1
    grdCheque.TextMatrix(nFila, 1) = Trim(frmIngCheques.NombreIF)
    grdCheque.TextMatrix(nFila, 2) = Trim(frmIngCheques.NroChq)
    grdCheque.TextMatrix(nFila, 3) = Format$(frmIngCheques.Importe, "#,##0.00")
    grdCheque.TextMatrix(nFila, 4) = Right(Trim(frmIngCheques.PersCodIF), 13)
    grdCheque.TextMatrix(nFila, 5) = Trim(frmIngCheques.FechaValChq)
    grdCheque.TextMatrix(nFila, 6) = Trim(frmIngCheques.nDiasValorizacion)
    grdCheque.TextMatrix(nFila, 7) = Trim(frmIngCheques.NroCtaIf)
    cmdAceptar.Enabled = True
    fraCheque.Enabled = True
End If
GetCheques
End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Icon = LoadPicture(App.path & gsRutaIcono)
End Sub

Private Sub grdCheque_DblClick()
Dim sNumChq As String, sPersCod As String
Dim nFila As Long
nFila = grdCheque.Row
sNumChq = grdCheque.TextMatrix(nFila, 2)
sPersCod = grdCheque.TextMatrix(nFila, 4)
frmIngCheques.InicioMuestra sPersCod, sNumChq, True, Trim(nOperacion), nMoneda
End Sub
