VERSION 5.00
Begin VB.Form frmLogSubasta 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9330
   Icon            =   "frmLogSubasta.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   9330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Sicmact.FlexEdit FlexSerie 
      Height          =   3630
      Left            =   7230
      TabIndex        =   9
      Top             =   1335
      Width           =   2070
      _ExtentX        =   3651
      _ExtentY        =   6403
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "#-Serie"
      EncabezadosAnchos=   "300-1700"
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
      ColumnasAEditar =   "X-1"
      TextStyleFixed  =   3
      ListaControles  =   "0-1"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-L"
      FormatosEdit    =   "0-0"
      TextArray0      =   "#"
      lbEditarFlex    =   -1  'True
      TipoBusqueda    =   0
      Appearance      =   0
      ColWidth0       =   300
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   345
      Left            =   6960
      TabIndex        =   4
      Top             =   5010
      Width           =   1125
   End
   Begin VB.CommandButton cmdRechazar 
      Caption         =   "&Rechazar"
      Height          =   345
      Left            =   5745
      TabIndex        =   3
      Top             =   5010
      Width           =   1125
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   345
      Left            =   8175
      TabIndex        =   2
      Top             =   4995
      Width           =   1125
   End
   Begin Sicmact.FlexEdit FlexDetalle 
      Height          =   3270
      Left            =   30
      TabIndex        =   1
      Top             =   1335
      Width           =   7200
      _ExtentX        =   12700
      _ExtentY        =   5768
      Cols0           =   5
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "#-Codigo-Producto-Valor Adjudica-Cant Adjudica"
      EncabezadosAnchos=   "300-1200-3000-1100-1100"
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
      ColumnasAEditar =   "X-X-X-3-4"
      TextStyleFixed  =   3
      ListaControles  =   "0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-L-L-R-R"
      FormatosEdit    =   "0-0-0-2-2"
      TextArray0      =   "#"
      TipoBusqueda    =   0
      lbBuscaDuplicadoText=   -1  'True
      Appearance      =   0
      ColWidth0       =   300
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin VB.Frame fraDoc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   15
      TabIndex        =   0
      Top             =   675
      Width           =   9255
      Begin Sicmact.TxtBuscar txtDoc 
         Height          =   300
         Left            =   1140
         TabIndex        =   5
         Top             =   217
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   529
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   0
         sTitulo         =   ""
      End
      Begin VB.Label lblDoc 
         Caption         =   "Documento :"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   270
         Width           =   945
      End
      Begin VB.Label lblDocG 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   2565
         TabIndex        =   6
         Top             =   225
         Width           =   6615
      End
   End
   Begin VB.Label lblTotalG 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
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
      Height          =   270
      Left            =   5595
      TabIndex        =   11
      Top             =   4650
      Width           =   1185
   End
   Begin VB.Label lblTotal 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   4800
      TabIndex        =   10
      Top             =   4650
      Width           =   1980
   End
   Begin VB.Label lblTit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Comprobante de Adjucicación  : 2001-00001"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   645
      Left            =   15
      TabIndex        =   8
      Top             =   45
      Width           =   9255
   End
End
Attribute VB_Name = "frmLogSubasta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lsCaption As String
Dim lbRechazo As Boolean
Dim lnMovNro As Long


Public Sub Inicio(psCaption As String, pbRechazo As Boolean)
    lsCaption = psCaption
    lbRechazo = pbRechazo
    
    Me.Show 1
End Sub

Private Sub cmdAceptar_Click()
    Dim oAlmacen As DLogAlmacen
    Set oAlmacen = New DLogAlmacen
    
    oAlmacen.SetEstadoSubasta lbRechazo, lnMovNro
    
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub FlexDetalle_RowColChange()
    Dim lnContador As Integer
    Dim lnEncontrar As Integer
    Dim I As Integer
    
    If InStr(1, Me.FlexDetalle.TextMatrix(FlexDetalle.Row, 2), "[S]") <> 0 Then
        lnContador = 0
        If Me.FlexDetalle.TextMatrix(FlexDetalle.Row, 3) = "" Then Exit Sub
        
        For I = 1 To CInt(Me.FlexSerie.Rows - 1)
            If FlexSerie.TextMatrix(I, 2) = Me.FlexDetalle.TextMatrix(FlexDetalle.Row, 0) Then
                lnContador = lnContador + 1
                FlexSerie.RowHeight(I) = 285
            Else
                FlexSerie.RowHeight(I) = 0
            End If
        Next I
        
        If lnContador <> CInt(Me.FlexDetalle.TextMatrix(FlexDetalle.Row, 3)) Then
            I = 0
            While lnEncontrar < lnContador
                I = I + 1
                If FlexSerie.TextMatrix(I, 2) = Me.FlexDetalle.TextMatrix(FlexDetalle.Row, 0) Then
                    Me.FlexSerie.EliminaFila I
                    lnEncontrar = lnEncontrar + 1
                    I = I - 1
                End If
            Wend
            For I = 1 To CInt(Me.FlexDetalle.TextMatrix(FlexDetalle.Row, 3))
                If Me.FlexSerie.TextMatrix(1, 2) = "" Then
                    Me.FlexSerie.AdicionaFila
                Else
                    Me.FlexSerie.AdicionaFila , , True
                End If
                FlexSerie.TextMatrix(FlexSerie.Rows - 1, 2) = Me.FlexDetalle.TextMatrix(FlexDetalle.Row, 0)
                FlexSerie.TextMatrix(FlexSerie.Rows - 1, 0) = Me.FlexDetalle.TextMatrix(FlexDetalle.Row, 0)
            Next I
        End If
    Else
        For I = 1 To CInt(Me.FlexSerie.Rows - 1)
            FlexSerie.RowHeight(I) = 0
        Next I
    End If
End Sub

Private Sub Form_Load()
    Caption = lsCaption
    Dim oAlmacen As DLogAlmacen
    Set oAlmacen = New DLogAlmacen
    
    txtDoc.rs = oAlmacen.CargaSubasta(0)
End Sub

Private Sub txtDoc_EmiteDatos()
    Me.lblDocG.Caption = txtDoc.psDescripcion
    Dim oAlmacen As DLogAlmacen
    Set oAlmacen = New DLogAlmacen
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim I As Integer
    Dim J As Integer
    
    
    
    If txtDoc.psDescripcion <> "" Then
        lnMovNro = CLng(Mid(txtDoc.psDescripcion, InStr(1, txtDoc.psDescripcion, "[", vbTextCompare) + 1, InStr(1, txtDoc.psDescripcion, "]", vbTextCompare) - 1))
        
        Me.FlexDetalle.rsFlex = oAlmacen.CargaSubastaDetalle(lnMovNro)
        
        Set rs = oAlmacen.CargaSubastaDetalleSerie(lnMovNro)
        
        For I = 1 To Me.FlexDetalle.Rows - 1
            If InStr(1, FlexDetalle.TextMatrix(I, 2), "[S]", vbTextCompare) <> 0 Then
                rs.MoveFirst
                While Not rs.EOF
                    If rs.Fields(0) = FlexDetalle.TextMatrix(I, 1) Then
                        Me.FlexSerie.AdicionaFila
                        Me.FlexSerie.TextMatrix(FlexSerie.Rows - 1, 0) = I
                        Me.FlexSerie.TextMatrix(FlexSerie.Rows - 1, 1) = rs.Fields(1)
                        Me.FlexSerie.RowHeight(FlexSerie.Rows - 1) = 0
                    End If
                    rs.MoveNext
                Wend
            End If
        Next I
    End If

End Sub

