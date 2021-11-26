VERSION 5.00
Begin VB.Form frmCredFormEvalCuotasIfis_SinConv 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cuotas Ifis"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6240
   Icon            =   "frmCredFormEvalCuotasIfis_SinConv.frx":0000
   LinkTopic       =   "Cuotas Ifis"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   6240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuitarIfis 
      Caption         =   "Quitar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1200
      TabIndex        =   6
      Top             =   2880
      Width           =   1170
   End
   Begin VB.CommandButton cmdAgregarIfis 
      Caption         =   "Agregar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   0
      TabIndex        =   5
      Top             =   2880
      Width           =   1170
   End
   Begin VB.Frame fBotones 
      Height          =   615
      Left            =   0
      TabIndex        =   2
      Top             =   3240
      Width           =   6135
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4800
         TabIndex        =   4
         Top             =   180
         Width           =   1170
      End
      Begin VB.CommandButton cmdAceptarIfis 
         Caption         =   "Aceptar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3480
         TabIndex        =   3
         Top             =   180
         Width           =   1170
      End
   End
   Begin VB.Frame Frame13 
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
      Height          =   2775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6135
      Begin SICMACT.FlexEdit feCuotaIfis 
         Height          =   2415
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   5880
         _ExtentX        =   10372
         _ExtentY        =   4260
         Cols0           =   4
         HighLight       =   1
         EncabezadosNombres=   "N°-Descripcion-Monto Cuota-Aux"
         EncabezadosAnchos=   "450-3900-1400-0"
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
         ColumnasAEditar =   "X-1-2-X"
         ListaControles  =   "0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-R-L"
         FormatosEdit    =   "0-0-2-0"
         TextArray0      =   "N°"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   3
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   450
         RowHeight0      =   300
      End
   End
   Begin SICMACT.EditMoney txtTotalIfis 
      Height          =   300
      Left            =   4920
      TabIndex        =   7
      Top             =   2880
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "0"
      Enabled         =   -1  'True
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4320
      TabIndex        =   8
      Top             =   2880
      Width           =   525
   End
End
Attribute VB_Name = "frmCredFormEvalCuotasIfis_SinConv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre      : frmCredFormEvalIfis
'** Descripción : Formulario que registra las IFIs que se asignan para los formatos de eval respectivos
'** Referencia  : ERS004-2016
'** Creación    : JOEP, 20160528 09:40:00 AM
'**********************************************************************************************
Option Explicit
Dim nTotalCuota As Currency
Dim nTotal As Currency
Dim MatOtraIFIRef As Variant
Dim i As Integer


Private Sub Form_Load()
    CentraForm Me
    txtTotalIfis.Enabled = False
End Sub

Public Sub Inicio(Optional ByVal pnMontoFLEX As Long, Optional ByRef pnTotalIFI As Currency = 0, Optional ByRef pMatIFI As Variant)
If pnMontoFLEX > 0 Then
    MatOtraIFIRef = pMatIFI
    Call CargarGridConArray
    nTotalCuota = pnTotalIFI
Else
Set pMatIFI = Nothing
Set MatOtraIFIRef = Nothing
nTotalCuota = 0
pnTotalIFI = 0
nTotal = 0
End If
Me.Show 1
'pnTotalIFI = nTotalCuota
'If MatOtraIFIRef Is Nothing Then
If nTotalCuota > 0 Then
pMatIFI = MatOtraIFIRef
pnTotalIFI = nTotalCuota
ElseIf nTotal > 0 Then
pMatIFI = MatOtraIFIRef
pnTotalIFI = nTotal
End If
End Sub

Private Sub CargarGridConArray()
    Dim i As Integer
   
    'Dim nTotal As Double
    'ReDim MatOtraIFIRef(0)
    feCuotaIfis.lbEditarFlex = True
    Call LimpiaFlex(feCuotaIfis)
    nTotal = 0
    For i = 1 To UBound(MatOtraIFIRef, 2)
        feCuotaIfis.AdicionaFila
        feCuotaIfis.TextMatrix(i, 0) = MatOtraIFIRef(1, i)
        feCuotaIfis.TextMatrix(i, 1) = MatOtraIFIRef(2, i)
        feCuotaIfis.TextMatrix(i, 2) = MatOtraIFIRef(3, i)
        nTotal = nTotal + feCuotaIfis.TextMatrix(i, 2)
    Next i
    txtTotalIfis.Text = nTotal
End Sub
Private Sub cmdAceptarIfis_Click()
    If feCuotaIfis.TextMatrix(1, 1) = "" Then
        MsgBox "Ud. primero debe registrar IFIs", vbInformation, "Aviso"
        Exit Sub
    End If

 If MsgBox("Desea Guardar IFIs??", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
    nTotalCuota = CCur(txtTotalIfis.Text)
    Dim i As Integer ' Llenado de Matriz
        ReDim MatOtraIFIRef(3, 0)
        For i = 1 To feCuotaIfis.Rows - 1
        ReDim Preserve MatOtraIFIRef(3, i)
            MatOtraIFIRef(1, i) = feCuotaIfis.TextMatrix(i, 0)
            MatOtraIFIRef(2, i) = feCuotaIfis.TextMatrix(i, 1)
            MatOtraIFIRef(3, i) = feCuotaIfis.TextMatrix(i, 2)
    Next i
    Unload Me
End Sub

Private Sub cmdAgregarIfis_Click()
    If feCuotaIfis.Rows - 1 < 25 Then
        feCuotaIfis.lbEditarFlex = True
        feCuotaIfis.AdicionaFila
        feCuotaIfis.SetFocus
        SendKeys "{Enter}"
    Else
    MsgBox "No puede agregar mas de 25 registros", vbInformation, "Aviso"
    End If
End Sub

Private Sub cmdQuitarIfis_Click()
    If MsgBox("Esta Seguro de Eliminar Registro?", vbInformation + vbYesNo, "Aviso") = vbYes Then
        feCuotaIfis.EliminaFila (feCuotaIfis.row)
        txtTotalIfis.Text = Format(SumarCampo(feCuotaIfis, 2), "#,##0.00")
    End If
End Sub

Private Sub feCuotaIfis_OnCellChange(pnRow As Long, pnCol As Long)
    If pnCol = 1 Then
        feCuotaIfis.TextMatrix(pnRow, pnCol) = UCase(feCuotaIfis.TextMatrix(pnRow, pnCol))
    End If
    txtTotalIfis.Text = Format(SumarCampo(feCuotaIfis, 2), "#,##0.00")
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub


