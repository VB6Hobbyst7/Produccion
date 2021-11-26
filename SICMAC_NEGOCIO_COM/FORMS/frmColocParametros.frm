VERSION 5.00
Begin VB.Form frmColocParametros 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6735
   Icon            =   "frmColocParametros.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   4605
      TabIndex        =   1
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   3525
      TabIndex        =   0
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5685
      TabIndex        =   2
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   2445
      TabIndex        =   4
      Top             =   4440
      Width           =   990
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "&Editar"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   4440
      Width           =   975
   End
   Begin VB.Frame fraParam 
      Height          =   4335
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   6555
      Begin Sicmact2001.FlexEdit grdParam 
         Height          =   3795
         Left            =   180
         TabIndex        =   6
         Top             =   360
         Width           =   6225
         _ExtentX        =   10980
         _ExtentY        =   6694
         Cols0           =   3
         FixedCols       =   0
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "CodParam-NomParam-Valor Param"
         EncabezadosAnchos=   "1200-2000-1200"
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
         ColumnasAEditar =   "X-1-2"
         ListaControles  =   "0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "L-L-R"
         FormatosEdit    =   "0-0-2"
         TextArray0      =   "CodParam"
         lbEditarFlex    =   -1  'True
         lbFlexDuplicados=   0   'False
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   1200
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
End
Attribute VB_Name = "frmColocParametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************
'* MANTENIMIENTO DE PARAMETROS DE COLOCACIONES
'Archivo:  frmColocParametro.frm
'LAYG   :  01/07/2001.
'Resumen:  Mantenimiento de Parametros de Colocaciones

Option Explicit
Dim fnProducto As Producto


Public Sub Inicio(Optional pbConsulta As Boolean = False, _
        Optional ByVal pbProducto As Producto = gColConsuPrendario)
If pbConsulta Then
    cmdCancelar.Visible = False
    cmdGrabar.Visible = False
    cmdEditar.Visible = False
    Me.Caption = "Colocaciones - Parámetros - Consulta"
    grdParam.lbEditarFlex = False
Else
    cmdImprimir.Visible = False
    cmdCancelar.Visible = True
    cmdGrabar.Visible = True
    cmdEditar.Visible = True
    cmdCancelar.Enabled = False
    cmdGrabar.Enabled = False
    Me.Caption = "Captaciones - Parámetros - Mantenimiento"
End If
fnProducto = pbProducto
Call CargaLista
Me.Show 1
End Sub

Private Sub cmdCancelar_Click()
    cmdEditar.Enabled = True
    cmdCancelar.Enabled = False
    cmdGrabar.Enabled = False
    cmdSalir.Enabled = True
    grdParam.SetFocus
End Sub

Private Sub cmdEditar_Click()
    cmdEditar.Enabled = False
    cmdCancelar.Enabled = True
    cmdGrabar.Enabled = True
    cmdSalir.Enabled = False
    grdParam.lbEditarFlex = True
End Sub

Private Sub cmdGrabar_Click()
If MsgBox("¿Desea grabar la información actualizada?", vbQuestion + vbYesNo, "Aviso") = vbNo Then
    Exit Sub
End If
Dim clsParam As nCapDefinicion
Dim i As Integer
Set clsParam = New nCapDefinicion
For i = 1 To grdParam.Rows - 1
    If grdParam.TextMatrix(i, 4) = "M" Then
        clsParam.ActualizaParametros grdParam.TextMatrix(i, 1), grdParam.TextMatrix(i, 2), CDbl(grdParam.TextMatrix(i, 3))
    End If
Next i
Set clsParam = Nothing
cmdCancelar.Enabled = False
cmdGrabar.Enabled = False
cmdSalir.Enabled = True
cmdEditar.Enabled = True
grdParam.SetFocus
End Sub

Private Sub cmdImprimir_Click()
Dim sCad As String
Dim Prev As Previo.clsPrevio
Dim nCarLin As Integer, i As Integer
Dim sTit1 As String, sTit2 As String
Dim sNumPag As String
Dim nLinPag As Integer, nCntPag As Integer
Dim sDesc As String * 55
Dim sVal As String * 12
sNumPag = "01"
nLinPag = 65
nCarLin = 70
sTit1 = "P A R A M E T R O S   "
sTit2 = ""
nCntPag = 0
sCad = sCad & CabeRepo("", "", nCarLin, "CAPTACIONES", sTit1, sTit2, "", sNumPag) & Chr$(10)
sCad = sCad & String(nCarLin, "=") & Chr$(10)
sCad = sCad & "DESCRIPCION" & Space(55) & "VALOR" & Chr$(10)
sCad = sCad & String(nCarLin, "=") & Chr$(10)
For i = 1 To grdParam.Rows - 1
    sDesc = Trim(grdParam.TextMatrix(i, 2))
    RSet sVal = Trim(grdParam.TextMatrix(i, 3))
    sCad = sCad & sDesc & Space(2) & sVal & Chr$(10)
Next i

Set Prev = New Previo.clsPrevio
Prev.Show sCad, "Parámetros Captaciones", True
Set Prev = Nothing
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub CargaLista()
Dim loParam As DColPCalculos
Set loParam = New DColPCalculos
    grdParam.rsFlex = loParam.dObtieneListaParametros(gColConsuPrendario)
Set loParam = Nothing
End Sub

Private Sub grdParam_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If cmdGrabar.Visible Then
        cmdEditar_Click
    Else
        cmdImprimir.SetFocus
    End If
End If
End Sub

Private Sub grdParam_OnCellChanged(pnFila As Integer, pnCol As Integer)
grdParam.TextMatrix(pnFila, 4) = "M"
End Sub
