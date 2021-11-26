VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form frmLogAlmRecep 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Almacén : Recepción"
   ClientHeight    =   5805
   ClientLeft      =   465
   ClientTop       =   1860
   ClientWidth     =   9930
   Icon            =   "frmLogAlmRecep.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   9930
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboAlm 
      Height          =   315
      Left            =   1785
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   1470
      Width           =   1755
   End
   Begin VB.CommandButton cmdAlm 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   390
      Index           =   1
      Left            =   4590
      TabIndex        =   15
      Top             =   5235
      Width           =   1260
   End
   Begin VB.CommandButton cmdAlm 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   390
      Index           =   2
      Left            =   6435
      TabIndex        =   14
      Top             =   5235
      Width           =   1260
   End
   Begin VB.Frame fraDocumen 
      Caption         =   "Documentos "
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
      Height          =   1500
      Left            =   5070
      TabIndex        =   9
      Top             =   390
      Width           =   4665
      Begin Sicmact.FlexEdit fgeDocume 
         Height          =   1200
         Left            =   105
         TabIndex        =   10
         Top             =   210
         Width           =   4470
         _ExtentX        =   7885
         _ExtentY        =   2117
         Cols0           =   5
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "Item-TpoDoc-Tipo Documento-Número-Fecha"
         EncabezadosAnchos=   "400-0-1400-1200-1000"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-3-4"
         TextStyleFixed  =   3
         ListaControles  =   "0-0-0-0-2"
         EncabezadosAlineacion=   "C-L-L-L-C"
         FormatosEdit    =   "0-0-0-0-0"
         TextArray0      =   "Item"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         MaxLength       =   12
         Appearance      =   0
         ColWidth0       =   405
         RowHeight0      =   300
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   390
      Left            =   8280
      TabIndex        =   0
      Top             =   5235
      Width           =   1305
   End
   Begin Sicmact.TxtBuscar txtConNro 
      Height          =   285
      Left            =   1395
      TabIndex        =   1
      Top             =   120
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   503
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      Enabled         =   0   'False
      TipoBusqueda    =   2
      sTitulo         =   ""
      EnabledText     =   0   'False
   End
   Begin Sicmact.FlexEdit fgeBS 
      Height          =   2640
      Left            =   225
      TabIndex        =   7
      Top             =   1920
      Width           =   9525
      _ExtentX        =   16801
      _ExtentY        =   4657
      Cols0           =   7
      HighLight       =   1
      AllowUserResizing=   3
      EncabezadosNombres=   "Item-cBSCod-Bien/Servicio-Unidad-Precio-Solicitado-Recepción"
      EncabezadosAnchos=   "400-0-3200-800-1000-1000-1000"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-X-X-X-X-X-6"
      TextStyleFixed  =   3
      ListaControles  =   "0-0-0-0-0-0-0"
      EncabezadosAlineacion=   "C-L-L-L-R-R-R"
      FormatosEdit    =   "0-0-0-0-2-2-2"
      AvanceCeldas    =   1
      TextArray0      =   "Item"
      lbEditarFlex    =   -1  'True
      lbUltimaInstancia=   -1  'True
      lbFormatoCol    =   -1  'True
      Appearance      =   0
      ColWidth0       =   405
      RowHeight0      =   300
   End
   Begin Sicmact.FlexEdit fgeTot 
      Height          =   930
      Left            =   4425
      TabIndex        =   8
      Top             =   4245
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   1640
      Rows            =   3
      Cols0           =   3
      ScrollBars      =   0
      AllowUserResizing=   3
      EncabezadosNombres=   "Item-No utilizado-Total"
      EncabezadosAnchos=   "1200-0-1000"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-X-X"
      TextStyleFixed  =   3
      ListaControles  =   "0-0-0"
      EncabezadosAlineacion=   "C-L-R"
      FormatosEdit    =   "0-0-2"
      AvanceCeldas    =   1
      TextArray0      =   "Item"
      lbUltimaInstancia=   -1  'True
      lbFormatoCol    =   -1  'True
      Appearance      =   0
      ColWidth0       =   1200
      RowHeight0      =   300
   End
   Begin MSComCtl2.DTPicker dtpFecha 
      Height          =   315
      Left            =   8130
      TabIndex        =   11
      Top             =   75
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   56885249
      CurrentDate     =   37099
   End
   Begin Sicmact.FlexEdit fgeImpues 
      Height          =   990
      Left            =   225
      TabIndex        =   13
      Top             =   4530
      Width           =   3750
      _ExtentX        =   6615
      _ExtentY        =   1746
      Cols0           =   6
      AllowUserResizing=   3
      EncabezadosNombres=   "Item-Campo no usado-Opc-Importe-Tasa-Monto"
      EncabezadosAnchos=   "400-0-400-800-800-1000"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-X-2-X-X-X"
      TextStyleFixed  =   3
      ListaControles  =   "0-0-4-0-0-0"
      EncabezadosAlineacion=   "C-L-C-L-R-R"
      FormatosEdit    =   "0-0-0-0-2-2"
      AvanceCeldas    =   1
      TextArray0      =   "Item"
      lbUltimaInstancia=   -1  'True
      lbFormatoCol    =   -1  'True
      Appearance      =   0
      ColWidth0       =   405
      RowHeight0      =   300
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "Almacén destino"
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
      Height          =   180
      Index           =   5
      Left            =   225
      TabIndex        =   16
      Top             =   1530
      Width           =   1500
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "Fecha :"
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
      Height          =   180
      Index           =   4
      Left            =   7260
      TabIndex        =   12
      Top             =   135
      Width           =   750
   End
   Begin VB.Label lblMoneda 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   1185
      TabIndex        =   6
      Top             =   960
      Width           =   1125
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "Moneda"
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
      Height          =   180
      Index           =   2
      Left            =   240
      TabIndex        =   5
      Top             =   990
      Width           =   690
   End
   Begin VB.Label lblProveedor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   1185
      TabIndex        =   4
      Top             =   615
      Width           =   3675
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "Proveedor"
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
      Height          =   195
      Index           =   1
      Left            =   225
      TabIndex        =   3
      Top             =   645
      Width           =   900
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "Contratación"
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
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   135
      Width           =   1095
   End
End
Attribute VB_Name = "frmLogAlmRecep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAlm_Click(Index As Integer)
    If Index = 1 Then
        'CANCELAR
        If MsgBox("¿ Estás seguro de cancelar toda la operación ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
            Call Limpiar
        End If
    ElseIf Index = 2 Then
        'GRABAR
        
    Else
        MsgBox "Tipo comando no reconocido", vbInformation, " Aviso "
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub fgeImpues_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
    If fgeImpues.TextMatrix(pnRow, 2) = "." Then
        fgeImpues.TextMatrix(pnRow, 5) = Format(CCur(fgeImpues.TextMatrix(pnRow, 4)) * (fgeTot.TextMatrix(2, 2) / 100), "#0.00")
    Else
        fgeImpues.TextMatrix(pnRow, 5) = ""
    End If
    'SubTotal
    fgeTot.TextMatrix(1, 2) = Format(CCur(fgeTot.TextMatrix(2, 2)) - fgeImpues.SumaRow(5), "#0.00")
End Sub

Private Sub Form_Load()
    Dim rs As ADODB.Recordset
    Dim clsDAlm As DLogAlmacen
    
    Call CentraForm(Me)
    Call CargaProcesos
    
    Set rs = New ADODB.Recordset
    Set clsDAlm = New DLogAlmacen
    
    Set rs = clsDAlm.CargaAlmacen(ATodos)
    If rs.RecordCount > 0 Then
        CargaCombo rs, cboAlm
    End If
    
    Set clsDAlm = Nothing
    Set rs = Nothing
    
End Sub

Private Sub CargaProcesos()
    Dim rs As ADODB.Recordset
    Dim clsDAlm As DLogAlmacen
    
    Call Limpiar
    Set rs = New ADODB.Recordset
    Set clsDAlm = New DLogAlmacen
    Set rs = clsDAlm.CargaContratacion(CTodosEstado, , gLogConEstadoInicio)
    If rs.RecordCount > 0 Then
        txtConNro.EditFlex = True
        txtConNro.rs = rs
        txtConNro.EditFlex = False
        txtConNro.Enabled = True
    Else
        txtConNro.Enabled = False
        txtConNro.Text = ""
    End If
    
    Set clsDAlm = Nothing
    Set rs = Nothing
End Sub


Private Sub txtConNro_EmiteDatos()
    Dim rs As ADODB.Recordset
    'Dim clsDAdq As DLogAdquisi
    Dim clsDAlm As DLogAlmacen
    Dim clsDGnral As DLogGeneral
    Dim nTot As Currency
    Dim nCont As Integer
    Dim nConNro As Long
    
    If txtConNro.Ok = False Then
        Exit Sub
    End If
    Set clsDGnral = New DLogGeneral
    nConNro = clsDGnral.GetnMovNro(txtConNro.Text)
    Set clsDGnral = Nothing
    
    Set rs = New ADODB.Recordset
    Set clsDAlm = New DLogAlmacen
    Set rs = clsDAlm.CargaContratacion(CUnRegistro, nConNro)
    If rs.RecordCount > 0 Then
        cmdAlm(1).Enabled = True
        cmdAlm(2).Enabled = True
        fgeTot.Clear
        fgeTot.TextMatrix(1, 0) = "SUB TOTAL"
        fgeTot.TextMatrix(2, 0) = "T O T A L"
        fgeTot.TextMatrix(1, 1) = "."
        fgeTot.TextMatrix(2, 1) = "."
        
        lblProveedor.Caption = rs!cPersNombre
        lblMoneda.Caption = rs!cConsDescripcion
        'Carga Detalle
        Set rs = clsDAlm.CargaConDetalle(nConNro)
        If rs.RecordCount > 0 Then
            Set fgeBS.Recordset = rs
            'Total
            For nCont = 1 To fgeBS.Rows - 1
                nTot = nTot + (CCur(fgeBS.TextMatrix(nCont, 4)) * CCur(fgeBS.TextMatrix(nCont, 5)))
            Next
            fgeTot.TextMatrix(1, 2) = Format(nTot, "#0.00")
            fgeTot.TextMatrix(2, 2) = Format(nTot, "#0.00")
        End If
        
        'Tasas
        Set rs = clsDAlm.CargaDocImp(TpoDocFactura)
        If rs.RecordCount > 0 Then
            Set fgeImpues.Recordset = rs
            fgeImpues.lbEditarFlex = True
        End If
        
        'Documento
        Set rs = clsDAlm.CargaOpeDoc(gLogOpeAlmRegistro)
        If rs.RecordCount > 0 Then
            Set fgeDocume.Recordset = rs
            fgeDocume.lbEditarFlex = True
        End If
    End If
    
    Set clsDAlm = Nothing
    Set rs = Nothing
End Sub


Private Sub Limpiar()
    cmdAlm(1).Enabled = False
    cmdAlm(2).Enabled = False
    txtConNro.Text = ""
    dtpFecha.Value = gdFecSis
    lblProveedor.Caption = ""
    lblMoneda.Caption = ""
    dtpFecha.Value = gdFecSis
    fgeBS.Clear
    fgeBS.FormaCabecera
    fgeBS.Rows = 2
    fgeDocume.Clear
    fgeDocume.FormaCabecera
    fgeDocume.Rows = 2
    fgeImpues.Clear
    fgeImpues.FormaCabecera
    fgeImpues.Rows = 2
    fgeTot.Clear
    fgeTot.TextMatrix(1, 0) = "SUB TOTAL"
    fgeTot.TextMatrix(2, 0) = "T O T A L"
    fgeTot.TextMatrix(1, 1) = "."
    fgeTot.TextMatrix(2, 1) = "."
End Sub
