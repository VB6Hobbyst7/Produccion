VERSION 5.00
Begin VB.Form frmOpePagLetra 
   Caption         =   "Pago a Proveedores: Canje por Letras"
   ClientHeight    =   4230
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6015
   Icon            =   "frmOpePagLetra.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   6015
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   315
      Left            =   4620
      TabIndex        =   7
      Top             =   3840
      Width           =   1065
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   315
      Left            =   3510
      TabIndex        =   6
      Top             =   3840
      Width           =   1065
   End
   Begin VB.Frame Frame1 
      Caption         =   "Letras"
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
      Height          =   2475
      Left            =   90
      TabIndex        =   12
      Top             =   1320
      Width           =   5835
      Begin VB.CommandButton cmdEliminaLetra 
         Caption         =   "&Eliminar"
         Height          =   285
         Left            =   1140
         TabIndex        =   8
         Top             =   2100
         Width           =   1035
      End
      Begin VB.CommandButton cmdNuevaLetra 
         Caption         =   "&Nuevo"
         Height          =   285
         Left            =   90
         TabIndex        =   3
         Top             =   2100
         Width           =   1035
      End
      Begin Sicmact.FlexEdit fg 
         Height          =   1815
         Left            =   90
         TabIndex        =   2
         Top             =   210
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   3201
         Cols0           =   5
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Numero-Fec Giro-Fec Vencim-Monto"
         EncabezadosAnchos=   "300-1500-1200-1200-1200"
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
         ColumnasAEditar =   "X-1-2-3-4"
         TextStyleFixed  =   3
         ListaControles  =   "0-0-2-2-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-C-C-R"
         FormatosEdit    =   "0-0-0-0-2"
         CantEntero      =   12
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   300
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.CommandButton cmdAceptarLetra 
         Caption         =   "&Aceptar"
         Height          =   285
         Left            =   90
         TabIndex        =   4
         Top             =   2070
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.CommandButton cmdCancelarLetra 
         Caption         =   "&Cancelar"
         Height          =   285
         Left            =   1140
         TabIndex        =   5
         Top             =   2070
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Label lbl3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "TOTAL"
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
         Height          =   210
         Left            =   3450
         TabIndex        =   14
         Top             =   2100
         Width           =   615
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         Height          =   315
         Left            =   4290
         TabIndex        =   13
         Top             =   2055
         Width           =   1275
      End
      Begin VB.Shape ShapeS 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H8000000C&
         Height          =   345
         Left            =   3210
         Top             =   2040
         Width           =   2385
      End
   End
   Begin VB.Frame fraDoc 
      Caption         =   "Datos de Documento de Proveedor"
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
      Height          =   1125
      Left            =   90
      TabIndex        =   9
      Top             =   120
      Width           =   5835
      Begin VB.TextBox txtProv 
         Enabled         =   0   'False
         Height          =   315
         Left            =   960
         TabIndex        =   0
         Top             =   270
         Width           =   4755
      End
      Begin VB.TextBox txtMonto 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   345
         Left            =   4050
         TabIndex        =   1
         Top             =   660
         Width           =   1665
      End
      Begin VB.Label Label3 
         Caption         =   "Monto"
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   3390
         TabIndex        =   11
         Top             =   750
         Width           =   525
      End
      Begin VB.Label Label1 
         Caption         =   "Proveedor"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   150
         TabIndex        =   10
         Top             =   330
         Width           =   945
      End
   End
End
Attribute VB_Name = "frmOpePagLetra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vbOk As Boolean
Dim rsAux As ADODB.Recordset
Dim lnMonto As Currency
Dim lnFilaActualiza As Integer

Public Sub Inicio(ByVal pnMonto As Currency, _
            ByVal psPersNombre As String)
txtProv = psPersNombre
lnMonto = pnMonto
txtMonto = Format(Abs(pnMonto), "##,###0.00")
Me.Show 1
End Sub

Private Sub cmdAceptar_Click()
vbOk = True
If CCur(lblTotal) = 0 Then
    MsgBox "Ingresar Letras a Canjear", vbInformation, "Aviso!"
    txtMonto = Format(lnMonto, gsFormatoNumeroView)
    txtMonto.SetFocus
    Exit Sub
End If
If CCur(lblTotal) <> CCur(txtMonto) Then
    MsgBox "Monto Ingresado diferente a Total de Documento", vbInformation, "Aviso"
    fg.SetFocus
    Exit Sub
End If
Set rsAux = fg.GetRsNew
Unload Me
DoEvents
End Sub

Private Sub cmdAceptarLetra_Click()
If ValidaDatos(fg.Row) Then
   ActivaControles True
   lnFilaActualiza = -1
Else
   fg.SetFocus
End If
End Sub

Private Sub cmdCancelar_Click()
vbOk = False
Unload Me
DoEvents
End Sub

Private Sub cmdCancelarLetra_Click()
ActivaControles True
fg.EliminaFila fg.Row
lnFilaActualiza = -1
lblTotal = Format(fg.SumaRow(4), gsFormatoNumeroView)
End Sub

Private Sub cmdEliminaLetra_Click()
fg.EliminaFila fg.Row
End Sub

Private Sub cmdNuevaLetra_Click()
ActivaControles False
fg.AdicionaFila
lnFilaActualiza = fg.Row
fg.SetFocus
End Sub

Public Property Get lbOk() As Variant
    lbOk = vbOk
End Property
Public Property Let lbOk(ByVal vNewValue As Variant)
    vbOk = vNewValue
End Property
Public Property Get rsLetras() As ADODB.Recordset
    Set rsLetras = rsAux
End Property
Public Property Let rsLetras(ByVal vNewValue As ADODB.Recordset)
    Set rsAux = vNewValue
End Property
Public Property Get vnMonto() As Currency
vnMonto = lnMonto
End Property
Public Property Let vnMonto(ByVal vNewValue As Currency)
lnMonto = vNewValue
End Property


Private Sub ActivaControles(pbActiva As Boolean)
cmdNuevaLetra.Visible = pbActiva
cmdEliminaLetra.Visible = pbActiva
cmdAceptarLetra.Visible = Not pbActiva
cmdCancelarLetra.Visible = Not pbActiva
cmdAceptar.Enabled = pbActiva
End Sub

Private Function ValidaDatos(pnRow As Integer) As Boolean
Dim k As Integer
ValidaDatos = False
If cmdAceptarLetra.Visible Then
   If fg.TextMatrix(pnRow, 1) = "" Then    'Nro Letra
      MsgBox "Ingresar Número de Letra...", vbInformation, "Aviso"
      Exit Function
   End If
   For k = 1 To fg.Rows - 1
      If fg.TextMatrix(pnRow, 1) = fg.TextMatrix(k, 1) And pnRow <> k Then
         MsgBox "No se puede ingresar Letra Duplicada...", vbInformation, "Aviso"
         Exit Function
      End If
   Next
   If fg.TextMatrix(pnRow, 2) = "" Then    'Fecha de Giro
      MsgBox "Ingresar Fecha de Giro de Letra...", vbInformation, "Aviso"
      Exit Function
   End If
   If fg.TextMatrix(pnRow, 3) = "" Then    'Fecha de Vencimiento
      MsgBox "Ingresar Fecha de Vencimiento de Letra...", vbInformation, "Aviso"
      Exit Function
   End If
   If nVal(fg.TextMatrix(pnRow, 4)) = 0 Then    'Monto de Letra
      MsgBox "Ingresar Monto de Letra...", vbInformation, "Aviso"
      Exit Function
   End If
End If
lblTotal = Format(fg.SumaRow(4), gsFormatoNumeroView)
ValidaDatos = True
End Function

Private Sub fg_OnCellChange(pnRow As Long, pnCol As Long)
lblTotal = Format(fg.SumaRow(4), gsFormatoNumeroView)
End Sub

Private Sub fg_RowColChange()
If lnFilaActualiza <> -1 Then
   fg.Row = lnFilaActualiza
End If
End Sub

Private Sub Form_Load()
CentraForm Me
End Sub
