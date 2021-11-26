VERSION 5.00
Begin VB.Form frmRRHHRepExpedientes 
   Caption         =   "Reporte de Expediente"
   ClientHeight    =   4545
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14670
   Icon            =   "frmRRHHRepExpediente.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   14670
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkAreaAll 
      Caption         =   "Todos"
      Height          =   375
      Left            =   4200
      TabIndex        =   16
      Top             =   600
      Width           =   830
   End
   Begin VB.ComboBox cboAreas 
      Height          =   315
      ItemData        =   "frmRRHHRepExpediente.frx":030A
      Left            =   1200
      List            =   "frmRRHHRepExpediente.frx":030C
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   600
      Width           =   3015
   End
   Begin VB.CommandButton cmdGenerar 
      Caption         =   "&Generar"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   3960
      Width           =   1215
   End
   Begin VB.OptionButton opnArea 
      Caption         =   "Area"
      Height          =   200
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   1440
      TabIndex        =   6
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton cmdExportar 
      Caption         =   "&Exportar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      Height          =   4215
      Left            =   5160
      TabIndex        =   4
      Top             =   120
      Width           =   9375
      Begin Sicmact.FlexEdit grdLista 
         Height          =   3825
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   9105
         _ExtentX        =   16060
         _ExtentY        =   9287
         Cols0           =   9
         HighLight       =   1
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Usuario-Area-Agencia-Documento-Nro Documento-Desde-Hasta-Glosa"
         EncabezadosAnchos=   "350-800-2500-2500-2800-1500-1000-1000-6000"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X"
         TextStyleFixed  =   3
         ListaControles  =   "0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-C-L-L-L-C-L"
         FormatosEdit    =   "0-0-0-0-0-0-0-0-0"
         TextArray0      =   "#"
         SelectionMode   =   1
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   345
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.OptionButton opnUsuario 
      Caption         =   "Usuario"
      Height          =   200
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Filtros"
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
      Begin VB.CommandButton cmdBuscarUser 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   310
         Left            =   3840
         TabIndex        =   15
         Top             =   840
         Width           =   270
      End
      Begin VB.TextBox txtUserNombre 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1080
         TabIndex        =   14
         Top             =   840
         Width           =   3015
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Parametros"
      Height          =   2295
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   4935
      Begin VB.CheckBox chkDocAll 
         Caption         =   "Todos"
         Height          =   375
         Left            =   3960
         TabIndex        =   10
         Top             =   480
         Width           =   830
      End
      Begin VB.ComboBox cboTpoDoc 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   480
         Width           =   3615
      End
      Begin VB.Frame Frame5 
         Caption         =   "Tipo de Documento"
         Height          =   855
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   4695
      End
      Begin VB.Frame Frame4 
         Caption         =   "Periodo de Vencimiento"
         Height          =   975
         Left            =   120
         TabIndex        =   11
         Top             =   1200
         Width           =   4695
         Begin VB.ComboBox cboMes 
            Height          =   315
            Left            =   2520
            TabIndex        =   18
            Text            =   "cboMes"
            Top             =   480
            Width           =   1455
         End
         Begin VB.ComboBox cboAño 
            Height          =   315
            Left            =   120
            TabIndex        =   17
            Text            =   "cboAño"
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label Label2 
            Caption         =   "Mes :"
            Height          =   255
            Left            =   2520
            TabIndex        =   20
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Año :"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   855
         End
      End
   End
End
Attribute VB_Name = "frmRRHHRepExpedientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cPersCod As String
Dim cPersNombre As String

Public Sub CargarAreas()
    Dim Sql As String
    Dim rs As New ADODB.Recordset
    Dim Conn As New DConecta
    bError = False
    Sql = "SELECT A.cAreaCod ,A.cAreaDescripcion FROM Areas A WHERE A.nAreaEstado = 1"
    If Not Conn.AbreConexion() Then
        bError = True
        sMsgError = "No se pudo Conectar al Servidor, Consulte con el Area de Sistemas"
        Set BuscaCliente = Nothing
        Set Conn = Nothing
        Exit Sub
    End If
    Set rs = Conn.CargaRecordSet(Sql)
    
    With rs
    Do Until .EOF
     cboAreas.AddItem "" & rs!cAreaDescripcion
     cboAreas.ItemData(cboAreas.NewIndex) = "" & rs!cAreaCod
       .MoveNext
    Loop
    End With
    rs.Close
    
    Conn.CierraConexion
    Set Conn = Nothing
    cboAreas.ListIndex = 0
End Sub
Public Sub CargarTpoDoc()
    Dim Sql As String
    Dim rs As New ADODB.Recordset
    Dim Conn As New DConecta
    bError = False
    Sql = "SELECT CON.nConsValor ,CON.cConsDescripcion FROM Constante CON WHERE CON.nConsCod = '10021' AND CON.bEstado = 1"
    If Not Conn.AbreConexion() Then
        bError = True
        sMsgError = "No se pudo Conectar al Servidor, Consulte con el Area de Sistemas"
        Set BuscaCliente = Nothing
        Set Conn = Nothing
        Exit Sub
    End If
    Set rs = Conn.CargaRecordSet(Sql)
    With rs
    Do Until .EOF
     cboTpoDoc.AddItem "" & rs!cConsDescripcion
     cboTpoDoc.ItemData(cboTpoDoc.NewIndex) = "" & rs!nConsValor
       .MoveNext
    Loop
    End With
    rs.Close
    
    Conn.CierraConexion
    Set Conn = Nothing
    cboTpoDoc.ListIndex = 0
End Sub

Public Sub Ini(psCaption As String, pForm As Form)
    Caption = psCaption
    Me.Show
    Show 0, pForm
End Sub

Private Sub chkAreaAll_Click()
    If chkAreaAll.value = 1 Then
        cboAreas.Enabled = False
    Else
        cboAreas.Enabled = True
    End If
    
End Sub

Private Sub chkDocAll_Click()
    If chkDocAll.value = 1 Then
        cboTpoDoc.Enabled = False
        cboAño.Enabled = False
        cboMes.Enabled = False
    Else
        cboTpoDoc.Enabled = True
        cboAño.Enabled = True
        cboMes.Enabled = True
    End If
End Sub

Private Sub chkFecIni_Click()
    If chkFecIni.value = 1 Then
        dpkIni.Enabled = True
    Else
        dpkIni.Enabled = False
    End If
End Sub

Private Sub cmdBuscarUser_Click()
    Dim oPersona As UPersona
    Dim oRRHH As DActualizaDatosRRHH
    Dim oAcceso As UAcceso
    Set oAcceso = New UAcceso
 
    Set oRRHH = New DActualizaDatosRRHH
    Set oPersona = New UPersona
    Set oPersona = frmBuscaPersona.Inicio(True)
    If Not oPersona Is Nothing Then
        'ClearScreen
        cPersCod = oPersona.sPersCod
        cPersNombre = oPersona.sPersNombre
        txtUserNombre.Text = cPersNombre
        Me.grdLista.Clear
        Me.grdLista.Rows = 2
        Me.grdLista.FormaCabecera
    End If
End Sub

Private Sub cmdExportar_Click()
    Dim fs As Scripting.FileSystemObject
    Dim xlsAplicacion As Excel.Application
    Dim lsArchivo As String
    Dim lsFile As String
    Dim lsNomHoja As String
    Dim xlsLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    Dim lbExisteHoja As Boolean
    Dim I As Integer: Dim IniTablas As Integer
    Dim oPersona As UPersona
    'On Error GoTo ErrorGeneraExcelFormato
    
    Set fs = New Scripting.FileSystemObject
    Set xlsAplicacion = New Excel.Application
    Set oPersona = New UPersona
    
    lsNomHoja = "Hoja1"
    lsFile = "ReporteExpedienteRRHH"
    
    lsArchivo = "\spooler\" & "Prueba" & "_" & gsCodUser & "_" & Format(gdFecSis, "yyyymmdd") & "_" & Format(Time, "hhmmss") & ".xls"
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
     
    IniTablas = 3
    For I = 1 To grdLista.Rows - 1
        xlHoja1.Cells(IniTablas + I, 2) = grdLista.TextMatrix(I, 1)
        xlHoja1.Cells(IniTablas + I, 3) = grdLista.TextMatrix(I, 2)
        xlHoja1.Cells(IniTablas + I, 4) = grdLista.TextMatrix(I, 3)
        xlHoja1.Cells(IniTablas + I, 5) = grdLista.TextMatrix(I, 4)
        xlHoja1.Cells(IniTablas + I, 6) = grdLista.TextMatrix(I, 5)
        xlHoja1.Cells(IniTablas + I, 7) = grdLista.TextMatrix(I, 6)
        xlHoja1.Cells(IniTablas + I, 8) = grdLista.TextMatrix(I, 7)
        xlHoja1.Cells(IniTablas + I, 9) = grdLista.TextMatrix(I, 8)
    Next I
    
    xlHoja1.Range(xlHoja1.Cells(3, 2), xlHoja1.Cells(I + 2, 9)).Borders.LineStyle = 1
    Dim psArchivoAGrabarC As String
    
    xlHoja1.SaveAs App.path & lsArchivo
    psArchivoAGrabarC = App.path & lsArchivo
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlHoja1 = Nothing
    cmdExportar.Enabled = False
End Sub

Private Sub cmdGenerar_Click()
    Dim Sql As String
    Dim rs As New ADODB.Recordset
    Dim cAreaCod As String
    Dim dFecIni As Integer
    Dim dFecFin As Integer
    Dim nTpoDoc As Integer
    
    'dFecIni = Format(dpkIni.value, "yyyyMMdd") 'COMENTADO POR ARLO20161221
    'dFecFin = Format(dpkFin.value, "yyyyMMdd") 'COMENTADO POR ARLO20161221
    dFecIni = Str(cboAño) 'AGREGADO POR ARLO20161221
    dFecFin = cboTpoDoc.ItemData(cboMes.ListIndex) 'AGREGADO POR ARLO20161221
    cAreaCod = cboAreas.ItemData(cboAreas.ListIndex)
    cAreaCod = "0" & cAreaCod
    
    'COMENTADO POR ARLO20161221 ***
'    If chkFecIni.value = 0 Then
'        dFecIni = ""
'    End If
    'COMENTADO POR ARLO20161221 ***
    
    If chkDocAll.value = 1 Then
        nTpoDoc = 0
    Else
        nTpoDoc = cboTpoDoc.ItemData(cboTpoDoc.ListIndex)
    End If
    
    If chkAreaAll.value = 1 Then
        cAreaCod = ""
    End If
    Dim rsValida, rsValida2 As ADODB.Recordset 'lucv
    Set rsValida = CargarGrid(cAreaCod, "", nTpoDoc, dFecIni, dFecFin)
    Set rsValida2 = CargarGrid("", cPersCod, nTpoDoc, dFecIni, dFecFin)
    
    If opnArea.value = True Then
        If Not rsValida.EOF Then
            Set grdLista.Recordset = CargarGrid(cAreaCod, "", nTpoDoc, dFecIni, dFecFin)
            grdLista.Enabled = True
        Else
            MsgBox "No se encontraron Registros", vbInformation, "Aviso"
            grdLista.Enabled = False
            Me.grdLista.Clear
            Me.grdLista.Rows = 2
            Me.grdLista.FormaCabecera
        End If
    ElseIf opnUsuario.value = True Then
        If Not rsValida2.EOF Then
        Set grdLista.Recordset = CargarGrid("", cPersCod, nTpoDoc, dFecIni, dFecFin)
        grdLista.Enabled = True
        Else
            MsgBox "No se encontraron Registros", vbInformation, "Aviso"
            grdLista.Enabled = False
            Me.grdLista.Clear
            Me.grdLista.Rows = 2
            Me.grdLista.FormaCabecera
        End If
    End If
    cmdExportar.Enabled = True
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    CargarAreas
    CargarTpoDoc
    cargarAño
    CargarMes
End Sub

Public Function CargarGrid(ByVal cAreCod As String, ByVal cPersCod As String, ByVal nTpoDoc As Integer, ByVal cFecIni As String, ByVal cFecFin As String) As ADODB.Recordset
    Dim Conn As New DConecta
    Dim oConst As New DConstante
    bError = False
    'Set Conn = New DConecta
    Dim Sql As String
    Sql = "stp_sel_ReporteExpedienteRRHH '" & Trim(cAreCod) & "','" & Trim(cPersCod) & "'," & Trim(nTpoDoc) & ",'" & Trim(cFecIni) & "','" & Trim(cFecFin) & "'"
    If Not Conn.AbreConexion() Then
        bError = True
        sMsgError = "No se pudo Conectar al Servidor, Consulte con el Area de Sistemas"
        Set BuscaCliente = Nothing
        Set Conn = Nothing
        Exit Function
    End If
    Conn.ConexionActiva.CommandTimeout = 7200
    Set CargarGrid = Conn.CargaRecordSet(Sql)
    Conn.CierraConexion
    Set Conn = Nothing
End Function


Private Sub opnArea_Click()
    cmdBuscarUser.Enabled = False
    cboAreas.Enabled = True
    txtUserNombre.Enabled = False
    Me.grdLista.Clear
    Me.grdLista.Rows = 2
    Me.grdLista.FormaCabecera
End Sub

Private Sub opnUsuario_Click()
    cmdBuscarUser.Enabled = True
    cboAreas.Enabled = False
    txtUserNombre.Enabled = True
End Sub


Private Sub grdLista_DblClick()
    Dim cPath As String
    cPath = Me.grdLista.TextMatrix(Me.grdLista.row, 9)
    If cPath = "" Then
        MsgBox "No se encontro archivo PDF.", vbCritical, "Aviso"
        Exit Sub
    End If
    ShellExecute Me.hwnd, "Open", cPath, 0&, "", vbNormalFocus
End Sub

'ARLO 20161221 INICIO*********************************
Private Sub cargarAño()
    
    Dim Año As String
    Dim Año2 As Integer
    Dim AñoIncio As Integer
    
    AñoIncio = 2013
    Año = Year(gdFecSis)
    Año2 = CInt(Año)
    'With rs
    Do
    cboAño.AddItem "" & AñoIncio
    AñoIncio = AñoIncio + 1
    Loop While (AñoIncio <= Año2 + 3)
    cboAño.ListIndex = 0

End Sub

Public Sub CargarMes()
    Dim Sql As String
    Dim rs As New ADODB.Recordset
    Dim Conn As New DConecta
    bError = False
    Sql = "SELECT CON.nConsValor ,CON.cConsDescripcion FROM Constante CON WHERE CON.nConsCod = '1010' AND CON.bEstado = 1"
    If Not Conn.AbreConexion() Then
        bError = True
        sMsgError = "No se pudo Conectar al Servidor, Consulte con el Area de Sistemas"
        Set BuscaCliente = Nothing
        Set Conn = Nothing
        Exit Sub
    End If
    Set rs = Conn.CargaRecordSet(Sql)
    With rs
    Do Until .EOF
     cboMes.AddItem "" & rs!cConsDescripcion
     cboMes.ItemData(cboMes.NewIndex) = "" & rs!nConsValor
       .MoveNext
    Loop
    End With
    rs.Close
    
    Conn.CierraConexion
    Set Conn = Nothing
    cboTpoDoc.ListIndex = 0
    cboMes.ListIndex = 0
End Sub

'ARLO 20161221 FIN*********************************
