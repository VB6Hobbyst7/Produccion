VERSION 5.00
Begin VB.Form FrmPigClienteRemate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Piezas Pendientes de Facturar"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11580
   Icon            =   "FrmPigClienteRemate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   11580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   10410
      TabIndex        =   9
      Top             =   6060
      Width           =   975
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   9330
      TabIndex        =   8
      Top             =   6075
      Width           =   990
   End
   Begin VB.Frame Frame2 
      Caption         =   "PIezas Rematadas por Clientes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5925
      Left            =   4080
      TabIndex        =   1
      Top             =   60
      Width           =   7425
      Begin VB.TextBox txtValorVenta 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   6015
         TabIndex        =   10
         Top             =   5490
         Width           =   1290
      End
      Begin VB.TextBox txtpiezas 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   4080
         TabIndex        =   5
         Top             =   5520
         Width           =   660
      End
      Begin SICMACT.FlexEdit fepiezasrem 
         Height          =   5205
         Left            =   30
         TabIndex        =   3
         Top             =   210
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   9181
         Cols0           =   13
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "Estado-Contrato-Pieza-Material-Tipo-Observacion-PNeto-VRemate-PVenta-Tasacion-Estado-NumRemate-TipoProceso"
         EncabezadosAnchos=   "0-1700-450-1000-1000-1500-650-0-900-0-0-0-0"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "R-C-R-L-L-R-R-R-R-C-C-C-C"
         FormatosEdit    =   "3-0-0-0-0-0-2-2-2-0-0-2-2"
         TextArray0      =   "Estado"
         lbUltimaInstancia=   -1  'True
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.Label Label3 
         Caption         =   "Valor Venta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   180
         Left            =   4950
         TabIndex        =   11
         Top             =   5580
         Width           =   1005
      End
      Begin VB.Label Label1 
         Caption         =   "Piezas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   180
         Left            =   3405
         TabIndex        =   4
         Top             =   5580
         Width           =   585
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Comprador - Cliente del Remate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5925
      Left            =   15
      TabIndex        =   0
      Top             =   60
      Width           =   4020
      Begin VB.TextBox txtcancliente 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   3015
         TabIndex        =   7
         Top             =   5520
         Width           =   930
      End
      Begin SICMACT.FlexEdit feclienteremate 
         Height          =   5205
         Left            =   45
         TabIndex        =   2
         Top             =   270
         Width           =   3900
         _ExtentX        =   6879
         _ExtentY        =   9181
         Cols0           =   4
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "Item-Cliente-Pieza-Monto"
         EncabezadosAnchos=   "400-2000-510-900"
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
         EncabezadosAlineacion=   "C-L-R-R"
         FormatosEdit    =   "0-0-0-2"
         TextArray0      =   "Item"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.Label Label2 
         Caption         =   "Total CLIENTES"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   180
         Left            =   2430
         TabIndex        =   6
         Top             =   5595
         Width           =   570
      End
   End
End
Attribute VB_Name = "FrmPigClienteRemate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lnRemate As Integer

Private Sub cmdGrabar_Click()
  FrmPigFacturaRemate.Show 1
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub feclienteremate_Click()
Dim oPiezas As DPigContrato
Dim nDatosPieza As ADODB.Recordset
Dim psNombre As String
Dim i As Integer

    psNombre = feclienteremate.TextMatrix(feclienteremate.Row, 1)
    Set oPiezas = New DPigContrato
    Set nDatosPieza = oPiezas.dObtieneListadoPiezaRem(psNombre, gPigSituacionPendFacturar, gPigTipoTasacRetasac, lnRemate)
    Set oPiezas = Nothing
    
    i = 1
    fepiezasrem.Clear
    fepiezasrem.Rows = 2
    fepiezasrem.FormaCabecera
    Do While Not (nDatosPieza.EOF)
         fepiezasrem.AdicionaFila
         fepiezasrem.TextMatrix(i, 1) = nDatosPieza!cCtaCod
         fepiezasrem.TextMatrix(i, 2) = nDatosPieza!Pieza
         fepiezasrem.TextMatrix(i, 3) = nDatosPieza!Material
         fepiezasrem.TextMatrix(i, 4) = IIf(IsNull(nDatosPieza!Tipo), "", nDatosPieza!Tipo)
         If IsNull(nDatosPieza!cDescripcion) Then
            fepiezasrem.TextMatrix(i, 5) = ""
         Else
            fepiezasrem.TextMatrix(i, 5) = nDatosPieza!cDescripcion
         End If
         fepiezasrem.TextMatrix(i, 6) = nDatosPieza!npesoneto
         fepiezasrem.TextMatrix(i, 7) = Format(nDatosPieza!ValorRemate, "#####,###.00")
         fepiezasrem.TextMatrix(i, 8) = Format(nDatosPieza!ValorVenta, "#####,###.00")
         fepiezasrem.TextMatrix(i, 9) = nDatosPieza!nValorTasacion
         fepiezasrem.TextMatrix(i, 10) = nDatosPieza!nPrdEstado
         fepiezasrem.TextMatrix(i, 11) = nDatosPieza!nRemate
         fepiezasrem.TextMatrix(i, 12) = nDatosPieza!nTipoProceso
         nDatosPieza.MoveNext
         i = i + 1
     Loop
     SumaFilas
     Set nDatosPieza = Nothing
     
End Sub

Private Sub fepiezasrem_DblClick()
   
    fepiezasrem.EliminaFila (fepiezasrem.Row)
    SumaFilas
    
End Sub

Private Sub Form_Load()
    Inicia
End Sub
Public Sub Inicia()

Dim nCliente As DPigContrato
Dim nDatosCliente As ADODB.Recordset
Dim i As Integer
Dim oPigRemate As DPigContrato
Dim rs As Recordset
    
    Set oPigRemate = New DPigContrato
    Set rs = oPigRemate.dObtieneDatosRemate(oPigRemate.dObtieneMaxRemate() - 1)
    If Not (rs.EOF And rs.BOF) Then
        lnRemate = rs!nRemate
    End If
    Set rs = Nothing
    Set oPigRemate = Nothing
    
    Set nCliente = New DPigContrato
    Set nDatosCliente = nCliente.dObtieneClienteRematado(gPigSituacionPendFacturar)
    Set nCliente = Nothing
    
    feclienteremate.Clear
    feclienteremate.Rows = 2
    feclienteremate.FormaCabecera
    fepiezasrem.Clear
    fepiezasrem.Rows = 2
    fepiezasrem.FormaCabecera
    
    If Not (nDatosCliente.EOF) And Not (nDatosCliente.BOF) Then
       i = 1
       Do While Not (nDatosCliente.EOF)
        feclienteremate.AdicionaFila
        feclienteremate.TextMatrix(i, 1) = nDatosCliente!cComprador
        feclienteremate.TextMatrix(i, 2) = nDatosCliente!Piezas
        feclienteremate.TextMatrix(i, 3) = Format(nDatosCliente!ValorVenta, "#####,###.00")
        nDatosCliente.MoveNext
        i = i + 1
       Loop
       txtcancliente.Text = i - 1
    End If
    Set nDatosCliente = Nothing
    
End Sub

Private Sub SumaFilas()
Dim i As Integer
Dim lnTotalVenta As Currency
Dim lnTotalPiezas As Integer

    lnTotalVenta = 0
    lnTotalPiezas = 0
    For i = 1 To fepiezasrem.Rows - 1
        If fepiezasrem.TextMatrix(i, 8) <> "" Then
            lnTotalVenta = lnTotalVenta + CCur(IIf(fepiezasrem.TextMatrix(i, 8) = "", 0, fepiezasrem.TextMatrix(i, 8)))
            lnTotalPiezas = lnTotalPiezas + 1
        End If
        
    Next i

    txtValorVenta = Format(lnTotalVenta, "###,###,###.00")
    txtpiezas = lnTotalPiezas
    
End Sub
