VERSION 5.00
Begin VB.Form FrmCredEstRFCDIF 
   Caption         =   "Estadistica RFC - DIF"
   ClientHeight    =   4380
   ClientLeft      =   60
   ClientTop       =   165
   ClientWidth     =   8310
   LinkTopic       =   "Form1"
   ScaleHeight     =   4380
   ScaleWidth      =   8310
   StartUpPosition =   2  'CenterScreen
   Begin SICMACT.FlexEdit Flex 
      Height          =   3555
      Left            =   0
      TabIndex        =   6
      Top             =   840
      Width           =   8235
      _ExtentX        =   14526
      _ExtentY        =   6271
      Cols0           =   6
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "DESCRIPCION-AGENCIA-RFC/DIF-MONTO INICIAL-MONTO PAGADO-MONTO FINAL"
      EncabezadosAnchos=   "2500-2000-1200-1500-1500-1500"
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
      ColumnasAEditar =   "X-X-X-X-X-X"
      ListaControles  =   "0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "L-L-L-R-R-R"
      FormatosEdit    =   "0-0-0-2-2-2"
      TextArray0      =   "DESCRIPCION"
      lbUltimaInstancia=   -1  'True
      ColWidth0       =   2505
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin VB.Frame Frame1 
      Height          =   765
      Left            =   0
      TabIndex        =   0
      Top             =   60
      Width           =   8295
      Begin VB.CheckBox ChkExcel 
         Caption         =   "Excel"
         Height          =   195
         Left            =   7200
         TabIndex        =   8
         Top             =   360
         Width           =   885
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "Cancelar"
         Height          =   345
         Left            =   5820
         TabIndex        =   7
         Top             =   270
         Width           =   1275
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Buscar"
         Height          =   345
         Left            =   4440
         TabIndex        =   5
         Top             =   270
         Width           =   1275
      End
      Begin VB.ComboBox cboAno 
         Height          =   315
         ItemData        =   "FrmCredEstRFCDIF.frx":0000
         Left            =   2880
         List            =   "FrmCredEstRFCDIF.frx":001C
         TabIndex        =   4
         Top             =   270
         Width           =   1335
      End
      Begin VB.ComboBox cboMes 
         Height          =   315
         ItemData        =   "FrmCredEstRFCDIF.frx":0038
         Left            =   570
         List            =   "FrmCredEstRFCDIF.frx":0064
         TabIndex        =   2
         Top             =   270
         Width           =   1665
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Año"
         Height          =   195
         Left            =   2430
         TabIndex        =   3
         Top             =   330
         Width           =   285
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Mes"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   330
         Width           =   300
      End
   End
End
Attribute VB_Name = "FrmCredEstRFCDIF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dFechaInicial As Date
Dim dFechaPagInicial As Date
Dim dFechaPagFin As Date

Private Sub cmdBuscar_Click()
    Dim oDCredDoc As DCredDoc
    Dim oNCredDoc As NCredDoc
    Dim rs As ADODB.Recordset
    
    If cboAno.ListIndex = -1 Then
        Exit Sub
    End If

    If cboMes.ListIndex = -1 Then
        Exit Sub
    End If

    Call ArmandoFiltros


    Set oDCredDoc = New DCredDoc
    Set rs = oDCredDoc.Repo_108506(dFechaInicial, dFechaPagInicial, dFechaPagFin)
    Set oDCredDoc = Nothing
    Flex.Clear
    Flex.FormaCabecera
    Do Until rs.EOF
        Flex.AdicionaFila
        Flex.TextMatrix(Flex.Row, 0) = rs!cConsDescripcion
        Flex.TextMatrix(Flex.Row, 1) = rs!cAgeDescripcion
        Flex.TextMatrix(Flex.Row, 2) = rs!cRFA
        Flex.TextMatrix(Flex.Row, 3) = Format(rs!nMontoIni, "#0.00")
        Flex.TextMatrix(Flex.Row, 4) = Format(rs!nMontoFin, "#0.00")
        Flex.TextMatrix(Flex.Row, 5) = Format(rs!nDIF, "#0.00")
        rs.MoveNext
    Loop
    Set rs = Nothing
    
    
    If ChkExcel.value = 1 Then
        Set oNCredDoc = New NCredDoc
        Call oNCredDoc.EstadisRFCDIF(dFechaInicial, dFechaPagInicial, dFechaPagFin)
        Set oNCredDoc = Nothing
        MsgBox "Archivo de Excel generado correctamente", vbInformation, "Aviso"
    End If
End Sub

Private Sub cmdCancelar_Click()
    Call Form_Load
End Sub

Private Sub Form_Load()
    cboAno.ListIndex = 0
    cboMes.ListIndex = 0
    Flex.FormaCabecera
End Sub

Sub ArmandoFiltros()
   Dim i As Integer
   Dim nAno As Integer
   Dim nMes As Integer
    
    For i = 0 To cboMes.ListCount - 1
        cboMes.ItemData(i) = i + 1
    Next i
    
    nAno = cboAno.ItemData(cboAno.ListIndex)
    nMes = cboMes.ItemData(cboMes.ListIndex)
    
    Select Case nMes
        Case 1
                 dFechaInicial = "31/12/" & CStr(nAno - 1)
                 dFechaPagInicial = "01/01/" & CStr(nAno)
                 dFechaPagFin = "31/01/" & CStr(nAno)
        Case 2
                 dFechaInicial = "31/01/" & CStr(nAno)
                 dFechaPagInicial = "01/02/" & CStr(nAno)
                 If nAno Mod 4 = 0 Then
                    dFechaPagFin = CDate("29/02/" & CStr(nAno))
                 Else
                    dFechaPagFin = "28/02/" & CStr(nAno)
                 End If
       Case 3
                If nAno Mod 4 = 0 Then
                    dFechaInicial = "29/02/" & CStr(nAno)
                 Else
                    dFechaInicial = "28/02/" & CStr(nAno)
                 End If
                 dFechaPagInicial = "01/03/" & CStr(nAno)
                 dFechaPagFin = "31/03/" & CStr(nAno)
      Case 4
                 dFechaInicial = "31/03/" & CStr(nAno)
                 dFechaPagInicial = "01/04/" & CStr(nAno)
                 dFechaPagFin = "30/04/" & CStr(nAno)
     Case 5
                 dFechaInicial = "30/04/" & CStr(nAno)
                 dFechaPagInicial = "01/05/" & CStr(nAno)
                 dFechaPagFin = "31/05/" & CStr(nAno)
    Case 6
                 dFechaInicial = "31/05/" & CStr(nAno)
                 dFechaPagInicial = "01/06/" & CStr(nAno)
                 dFechaPagFin = "30/06/" & CStr(nAno)
                 
   Case 7
                 dFechaInicial = "30/06/" & CStr(nAno)
                 dFechaPagInicial = "01/07/" & CStr(nAno)
                 dFechaPagFin = "31/07/" & CStr(nAno)
   Case 8
                dFechaInicial = "31/07/" & CStr(nAno)
                 dFechaPagInicial = "01/08/" & CStr(nAno)
                 dFechaPagFin = "31/08/" & CStr(nAno)
   Case 9
                 dFechaInicial = "31/08/" & CStr(nAno)
                 dFechaPagInicial = "01/09/" & CStr(nAno)
                 dFechaPagFin = "30/09/" & CStr(nAno)
   Case 10
                 dFechaInicial = "30/09/" & CStr(nAno)
                 dFechaPagInicial = "01/10/" & CStr(nAno)
                 dFechaPagFin = "31/10/" & CStr(nAno)
   Case 11
                 dFechaInicial = "31/10/" & CStr(nAno)
                 dFechaPagInicial = "01/11/" & CStr(nAno)
                 dFechaPagFin = "30/11/" & CStr(nAno)
   Case 12
                 dFechaInicial = "30/11/" & CStr(nAno)
                 dFechaPagInicial = "01/12/" & CStr(nAno)
                 dFechaPagFin = "31/12/" & CStr(nAno)
    End Select
    
End Sub

