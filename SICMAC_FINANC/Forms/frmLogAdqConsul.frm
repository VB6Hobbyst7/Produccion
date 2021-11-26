VERSION 5.00
Object = "{DFDE2506-090D-11D5-BEF8-C11EAA34970C}#2.0#0"; "Spinner.ocx"
Begin VB.Form frmLogAdqConsul 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta de Adquisiciones"
   ClientHeight    =   6570
   ClientLeft      =   465
   ClientTop       =   1785
   ClientWidth     =   11115
   Icon            =   "frmLogAdqConsul.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   11115
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdAdq 
      Caption         =   "&Imprimir"
      Height          =   390
      Left            =   5415
      TabIndex        =   11
      Top             =   6060
      Width           =   1305
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   390
      Left            =   8760
      TabIndex        =   2
      Top             =   6045
      Width           =   1305
   End
   Begin VB.OptionButton optPlan 
      Caption         =   "Normal"
      Height          =   330
      Index           =   0
      Left            =   5235
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   45
      Value           =   -1  'True
      Width           =   1305
   End
   Begin VB.OptionButton optPlan 
      Caption         =   "Extemporaneo"
      Height          =   330
      Index           =   1
      Left            =   6555
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   45
      Width           =   1305
   End
   Begin Sicmact.Usuario Usuario 
      Left            =   -15
      Top             =   6135
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin Sicmact.FlexEdit fgeBS 
      Height          =   5250
      Left            =   4410
      TabIndex        =   3
      Top             =   690
      Width           =   6510
      _ExtentX        =   11483
      _ExtentY        =   9260
      Cols0           =   7
      HighLight       =   1
      AllowUserResizing=   3
      EncabezadosNombres=   "Item-Codigo-Bien/Servicio-Unidad-Cantidad-PrecioProm-SubTotal"
      EncabezadosAnchos=   "400-0-2300-650-900-900-1000"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-X-X-X-X-X-X"
      TextStyleFixed  =   3
      ListaControles  =   "0-0-0-0-0-0-0"
      EncabezadosAlineacion=   "C-L-L-L-R-R-R"
      FormatosEdit    =   "0-0-0-0-2-2-2"
      AvanceCeldas    =   1
      TextArray0      =   "Item"
      lbUltimaInstancia=   -1  'True
      lbFormatoCol    =   -1  'True
      Appearance      =   0
      ColWidth0       =   405
      RowHeight0      =   285
   End
   Begin Spinner.uSpinner spinPeriodo 
      Height          =   300
      Left            =   8910
      TabIndex        =   4
      Top             =   90
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   529
      Max             =   9999
      Min             =   2000
      MaxLength       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
   End
   Begin Sicmact.FlexEdit fgeAdq 
      Height          =   5250
      Left            =   165
      TabIndex        =   5
      Top             =   690
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   9260
      Cols0           =   5
      HighLight       =   1
      AllowUserResizing=   3
      EncabezadosNombres=   "Item-Adquisición-cEstadoCod-Estado-OK"
      EncabezadosAnchos=   "400-1900-0-1200-350"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-X-X-X-4"
      TextStyleFixed  =   3
      ListaControles  =   "0-0-0-0-4"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "L-L-L-L-C"
      FormatosEdit    =   "0-0-0-0-0"
      TextArray0      =   "Item"
      lbUltimaInstancia=   -1  'True
      Appearance      =   0
      ColWidth0       =   405
      RowHeight0      =   285
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "Detalle de planes :"
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
      Index           =   3
      Left            =   4575
      TabIndex        =   10
      Top             =   450
      Width           =   1785
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "Planes de Adquisición :"
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
      Index           =   2
      Left            =   270
      TabIndex        =   9
      Top             =   450
      Width           =   2010
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "Año :"
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
      Index           =   1
      Left            =   8340
      TabIndex        =   8
      Top             =   150
      Width           =   660
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "Area :"
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
      Left            =   345
      TabIndex        =   7
      Top             =   105
      Width           =   570
   End
   Begin VB.Label lblAreaDes 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   975
      TabIndex        =   6
      Top             =   75
      Width           =   3810
   End
End
Attribute VB_Name = "frmLogAdqConsul"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim psTpoReq As String
Dim clsDAdq As DLogAdquisi

Private Sub cmdAdq_Click()
    Dim clsNImp As NLogImpre
    Dim clsPrevio As clsPrevio
    Dim rs As ADODB.Recordset
    Dim sImpre As String
    Dim nCont As Integer, nSum As Integer
    
    'Validación de Requerimientos
    Set rs = New ADODB.Recordset
    rs.Fields.Append "cAdqNro", adVarChar, 25, adFldMayBeNull
    rs.Open
    For nCont = 1 To fgeAdq.Rows - 1
        If (fgeAdq.TextMatrix(nCont, 4)) = "." Then
            rs.AddNew "cAdqNro", fgeAdq.TextMatrix(nCont, 1)
            rs.Update
            rs.MoveNext
            nSum = nSum + 1
        End If
    Next

    If nSum = 0 Then
        Set rs = Nothing
        MsgBox "Por favor, determine que Planes se Imprimen", vbInformation, " Aviso "
        Exit Sub
    End If
    
    Set clsNImp = New NLogImpre
    sImpre = clsNImp.ImpAdquisicion(gsNomAge, gdFecSis, rs)
    Set clsNImp = Nothing
    Set rs = Nothing
    
    Set clsPrevio = New clsPrevio
    clsPrevio.Show sImpre, Me.Caption, True
    Set clsPrevio = Nothing
End Sub

Private Sub fgeAdq_OnRowChange(pnRow As Long, pnCol As Long)
    Dim rs As ADODB.Recordset
    Dim sAdqNro As String
    'Cargar información del Detalle
    sAdqNro = Trim(fgeAdq.TextMatrix(fgeAdq.Row, 1))
    If Trim(sAdqNro) <> "" Then
        Set rs = New ADODB.Recordset
        Set rs = clsDAdq.CargaAdqDetalle(AdqDetUnRegistro, sAdqNro)
        If rs.RecordCount > 0 Then
            Set fgeBS.Recordset = rs
            fgeBS.AdicionaFila
            fgeBS.BackColorRow &HC0FFFF
            fgeBS.TextMatrix(fgeBS.Row, 2) = "T O T A L "
            fgeBS.TextMatrix(fgeBS.Row, 6) = Format(fgeBS.SumaRow(6), "#,##0.00")
        End If
        Set rs = Nothing
    End If
End Sub

Private Sub Form_Load()
    Dim rs As ADODB.Recordset
    Set clsDAdq = New DLogAdquisi
    Set rs = New ADODB.Recordset
    Call CentraForm(Me)
    'Carga información de la relación usuario-area
    Usuario.Inicio gsCodUser
    If Len(Usuario.AreaCod) = 0 Then
        MsgBox "Usuario no determinado", vbInformation, "Aviso"
        Exit Sub
    End If
    lblAreaDes.Caption = Usuario.AreaNom
    
    Call optPlan_Click(0)
    'Inicia Periodo
    spinPeriodo.Valor = IIf(psTpoReq = "1", Year(gdFecSis) + 1, Year(gdFecSis))
End Sub


Private Sub cmdSalir_Click()
    Set clsDAdq = Nothing
    Unload Me
End Sub

Private Sub Limpiar()
    'Limpiar FLEX
    fgeAdq.Clear
    fgeAdq.FormaCabecera
    fgeAdq.Rows = 2
    fgeBS.Clear
    fgeBS.FormaCabecera
    fgeBS.Rows = 2
End Sub

Private Sub optPlan_Click(Index As Integer)
    psTpoReq = IIf(optPlan(0).Value = True, "1", "2")
    spinPeriodo_Change
End Sub

Private Sub spinPeriodo_Change()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    'Actualiza FLEX
    Call Limpiar
    'Carga Plan de Obtención
    Set rs = clsDAdq.CargaAdquisicion(AdqTodosGnral, psTpoReq, spinPeriodo.Valor)
    If rs.RecordCount > 0 Then
        fgeAdq.lbEditarFlex = True
        Set fgeAdq.Recordset = rs
        cmdAdq.Enabled = True
        Call fgeAdq_OnRowChange(fgeAdq.Row, fgeAdq.Col)
    Else
        cmdAdq.Enabled = False
    End If

End Sub


