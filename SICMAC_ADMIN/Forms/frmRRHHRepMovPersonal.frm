VERSION 5.00
Begin VB.Form frmRRHHRepMovPersonal 
   ClientHeight    =   6690
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9210
   Icon            =   "frmRRHHRepMovPersonal.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   6690
   ScaleWidth      =   9210
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGenerar 
      Caption         =   "&Generar"
      Height          =   390
      Left            =   8040
      TabIndex        =   4
      Top             =   240
      Width           =   990
   End
   Begin VB.CheckBox chkTodos 
      Caption         =   "Todos"
      Height          =   315
      Left            =   3600
      TabIndex        =   3
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   390
      Left            =   8160
      TabIndex        =   2
      Top             =   6120
      Width           =   990
   End
   Begin VB.CommandButton cmdReporte 
      Caption         =   "&Exportar"
      Enabled         =   0   'False
      Height          =   390
      Left            =   6960
      TabIndex        =   1
      Top             =   6120
      Width           =   1110
   End
   Begin VB.ComboBox cmbAgencias 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   3255
   End
   Begin Sicmact.FlexEdit grdLista 
      Height          =   5145
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   8985
      _ExtentX        =   15849
      _ExtentY        =   9075
      Cols0           =   10
      HighLight       =   1
      RowSizingMode   =   1
      EncabezadosNombres=   "#-Usuario-Nombre-Estado-Agencia-Area-Cargo-Fec. Cargo-D.N.I.-Fec Ingreso"
      EncabezadosAnchos=   "350-800-4000-1000-2800-2800-4000-1200-1200-1200"
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
      ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X"
      TextStyleFixed  =   3
      ListaControles  =   "0-0-0-0-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-L-L-C-L-L-L-C-C-C"
      FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0"
      TextArray0      =   "#"
      SelectionMode   =   1
      lbUltimaInstancia=   -1  'True
      ColWidth0       =   345
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin VB.Frame Frame1 
      Caption         =   "Agencia"
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   9015
   End
End
Attribute VB_Name = "frmRRHHRepMovPersonal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet



Private Sub chkTodos_Click()
    If chkTodos.value = 1 Then
        cmbAgencias.Enabled = False
    Else
        cmbAgencias.Enabled = True
    End If
End Sub

Private Sub cmdGenerar_Click()
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim cAgeCod As String
    cAgeCod = cmbAgencias.ItemData(cmbAgencias.ListIndex)
    If cAgeCod < 10 Then
        cAgeCod = "0" & cAgeCod
    End If
    cmdReporte.Enabled = True
    If chkTodos.value = False Then
        Set grdLista.Recordset = ObtenerMovPersonal(cAgeCod)
    Else
        Set grdLista.Recordset = ObtenerMovPersonal("")
    End If
End Sub

Private Sub cmdReporte_Click()
    Dim fs As Scripting.FileSystemObject
    Dim xlsAplicacion As Excel.Application
    Dim lsArchivo As String
    Dim lsFile As String
    Dim lsNomHoja As String
    Dim xlsLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    Dim lbExisteHoja As Boolean
    Dim i As Integer: Dim IniTablas As Integer
    Dim oPersona As UPersona
    'On Error GoTo ErrorGeneraExcelFormato
    
    Set fs = New Scripting.FileSystemObject
    Set xlsAplicacion = New Excel.Application
    Set oPersona = New UPersona
    
    lsNomHoja = "Hoja1"
    lsFile = "MovimientoPersonal"
    
    lsArchivo = "\spooler\" & "MovPersonal" & "_" & gsCodUser & "_" & Format(gdFecSis, "yyyymmdd") & "_" & Format(Time, "hhmmss") & ".xls"
    If fs.FileExists(App.path & "\FormatoCarta\" & lsFile & ".xls") Then
        Set xlsLibro = xlsAplicacion.Workbooks.Open(App.path & "\FormatoCarta\" & lsFile & ".xls")
    Else
        MsgBox "No Existe Plantilla en Carpeta FormatoCarta (" & lsFile & ".xls), Consulte con el Area de TI", vbInformation, "Advertencia"
        Exit Sub
    End If

    For Each xlHoja1 In xlsLibro.Worksheets
       If xlHoja1.Name = lsNomHoja Then
            xlHoja1.Activate
         lbExisteHoja = True
        Exit For
       End If
    Next
    
    If lbExisteHoja = False Then
        Set xlHoja1 = xlsLibro.Worksheets
        xlHoja1.Name = lsNomHoja
    End If
    
    xlHoja1.Cells(IniTablas + 2, 3) = oPersona.sPersNombre
    xlHoja1.Cells(IniTablas + 3, 3) = gsNomAge
    xlHoja1.Cells(IniTablas + 4, 3) = Format(gdFecSis, "dd/mm/yyyy")
    
    IniTablas = 6
    For i = 1 To grdLista.Rows - 1
        xlHoja1.Cells(IniTablas + i, 2) = grdLista.TextMatrix(i, 1)
        xlHoja1.Cells(IniTablas + i, 3) = grdLista.TextMatrix(i, 2)
        xlHoja1.Cells(IniTablas + i, 4) = grdLista.TextMatrix(i, 3)
        xlHoja1.Cells(IniTablas + i, 5) = grdLista.TextMatrix(i, 4)
        xlHoja1.Cells(IniTablas + i, 6) = grdLista.TextMatrix(i, 5)
        xlHoja1.Cells(IniTablas + i, 7) = grdLista.TextMatrix(i, 6)
        xlHoja1.Cells(IniTablas + i, 8) = grdLista.TextMatrix(i, 7)
        xlHoja1.Cells(IniTablas + i, 9) = grdLista.TextMatrix(i, 8)
        xlHoja1.Cells(IniTablas + i, 10) = grdLista.TextMatrix(i, 9)
    Next i
    
    xlHoja1.Range(xlHoja1.Cells(6, 2), xlHoja1.Cells(i + 5, 10)).Borders.LineStyle = 1
    Dim psArchivoAGrabarC As String
    
    xlHoja1.SaveAs App.path & lsArchivo
    psArchivoAGrabarC = App.path & lsArchivo
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlHoja1 = Nothing
    cmdReporte.Enabled = False
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
     CargarCmbAgencias
     cmbAgencias.ListIndex = 0
End Sub

Public Sub Ini(psCaption As String, pForm As Form)
    Caption = psCaption
    Show 0, pForm
End Sub

Public Sub CargarCmbAgencias()
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim Conn As New DConecta
    bError = False
    sql = "SELECT cAgeCod ,cAgeDescripcion FROM Agencias A WHERE nEstado = 1"
    'Set Conn = New COMConecta.DCOMConecta
    If Not Conn.AbreConexion() Then
        bError = True
        sMsgError = "No se pudo Conectar al Servidor, Consulte con el Area de Sistemas"
        Set BuscaCliente = Nothing
        Set Conn = Nothing
        Exit Sub
    End If
    Conn.ConexionActiva.CommandTimeout = 7200
    Set rs = Conn.CargaRecordSet(sql)
    
    With rs
    Do Until .EOF
     cmbAgencias.AddItem "" & rs!cAgeDescripcion
     cmbAgencias.ItemData(cmbAgencias.NewIndex) = "" & rs!cAgeCod
       .MoveNext
    Loop
    End With
    rs.Close
    
    Conn.CierraConexion
    Set Conn = Nothing
End Sub

Public Function ObtenerMovPersonal(ByVal cAgeCod As String) As ADODB.Recordset
    Dim Conn As New DConecta
    bError = False
    'Set Conn = New DConecta
    If Not Conn.AbreConexion() Then
        bError = True
        sMsgError = "No se pudo Conectar al Servidor, Consulte con el Area de Sistemas"
        Set BuscaCliente = Nothing
        Set Conn = Nothing
        Exit Function
    End If
    Conn.ConexionActiva.CommandTimeout = 7200
    Set ObtenerMovPersonal = Conn.CargaRecordSet("stp_sel_ReporteMovimientoPersonal '" & Trim(cAgeCod) & "'")
    Conn.CierraConexion
    Set Conn = Nothing
End Function
