VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRRHHRegistroExpedientes 
   Caption         =   "Registro de Expediente de RR.HH."
   ClientHeight    =   7335
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12060
   Icon            =   "frmRRHHRegistroExpedientes.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   12060
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      Height          =   375
      Left            =   6600
      TabIndex        =   12
      Top             =   6840
      Width           =   1095
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "&Editar"
      Height          =   375
      Left            =   7800
      TabIndex        =   11
      Top             =   6840
      Width           =   1110
   End
   Begin VB.Frame Frame1 
      Height          =   4455
      Left            =   360
      TabIndex        =   4
      Top             =   2040
      Width           =   11295
      Begin Sicmact.FlexEdit grdLista 
         Height          =   3615
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   6376
         Cols0           =   8
         HighLight       =   2
         RowSizingMode   =   1
         VisiblePopMenu  =   -1  'True
         EncabezadosNombres=   "#-Documento-Nº Doc-Desde-Hasta-Glosa-PDF-Detalle"
         EncabezadosAnchos=   "350-2000-1650-1000-1000-3450-500-800"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X-7"
         ListaControles  =   "0-0-0-0-0-0-0-1"
         BackColor       =   16777215
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-C-L-L-L-C"
         FormatosEdit    =   "0-0-0-0-0-0-0-0"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   6
         lbPuntero       =   -1  'True
         lbOrdenaCol     =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   345
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
         CellBackColor   =   16777215
         RowHeightMin    =   1
      End
      Begin VB.Label lblCargo 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   3480
         TabIndex        =   10
         Top             =   240
         Width           =   5895
      End
      Begin VB.Label lblNroID 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   720
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Cargo Actual:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2160
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "DNI:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   390
      Left            =   10680
      TabIndex        =   2
      Top             =   6840
      Width           =   1110
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   390
      Left            =   9000
      TabIndex        =   1
      Top             =   6840
      Width           =   1110
   End
   Begin VB.CommandButton cmdExportar 
      Caption         =   "&Exportar"
      Height          =   390
      Left            =   120
      TabIndex        =   0
      Top             =   6840
      Width           =   1110
   End
   Begin Sicmact.ctrRRHHGen ctrRRHHGen 
      Height          =   1215
      Left            =   360
      TabIndex        =   9
      Top             =   720
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   2143
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   6615
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   11668
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Expediente de Usuarios"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmRRHHRegistroExpedientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'variables agregado inicio por pti1 ers029-2018
Public cUserExpediente As String
Public cPersCod As String
Public pcOpeEdit As Boolean
Public ccanFilas As Integer
'fin agregado variables

Dim cPersCodExp As String
Public pcElementos As String
Public pcOpeDet As Boolean

Public pcTpoDoc As String
Public pcNroDoc As String
Public pcPathFile As String
Public pdDesde  As String
Public pdHasta As String
Public pcGlosa As String

Private Sub cmdEditar_Click()
'*****************AGREGADO POR PTI1 ERS029-2018 18/04/2019
    If Me.ctrRRHHGen.psCodigoPersona = "" Then
        MsgBox "Debe seleccionar registro, para editar documento", vbExclamation, "Aviso"
    Else
     Dim path As String
     path = ""
        pcOpeEdit = True
        ObtenerDatosMatrix
        If pcNroDoc = "" Then
        MsgBox "No existen datos, para editar documento", vbExclamation, "Aviso"
        Else
        frmRRHHRegistroDocumento.Ini gTipoOpeRegistro, "EDICIÓN DE EXPEDIENTE", Me
        Set grdLista.Recordset = CargarExpedienteUser(cPersCod)
        grdListaAddDetalle
        End If
    End If
  '*****************FIN AGREGADO
End Sub

Private Sub CmdEliminar_Click()
'******************AGREGADO POR PTI1 ERS029-2018 18/04/2019
Dim Conn As New DConecta
Dim Sql As String
ObtenerDatosMatrix
If Me.ctrRRHHGen.psCodigoPersona = "" Or pcNroDoc = "" Then
        MsgBox "No existen datos", vbExclamation, "Aviso"
Else
 Dim msgvalue As Integer
 msgvalue = MsgBox("Ud. ¿está seguro de eliminar el expediente número " + pcNroDoc + " ?", vbQuestion + vbYesNo, "Eliminar")
 Select Case msgvalue
 Case 6
  bError = False
    Sql = "Update RHExpediente set cEstado = 0 where cPerscod =  '" & cPersCod & "' and cNumdoc=  '" & pcNroDoc & "' "
    If Not Conn.AbreConexion() Then
        bError = True
        sMsgError = "No se pudo Conectar al Servidor, Consulte con el Area de Sistemas"
        Set BuscaCliente = Nothing
        Set Conn = Nothing
        Exit Sub
    End If
    Conn.ConexionActiva.CommandTimeout = 7200
    Set ObtenerExpedientePersonal = Conn.CargaRecordSet(Sql)
    Conn.CierraConexion
    Set Conn = Nothing
    MsgBox "Expediente Eliminado con exito", vbExclamation, "Aviso"
    If ccanFilas > 1 Then
    Set grdLista.Recordset = CargarExpedienteUser(cPersCod)
    grdListaAddDetalle
    Else
    Dim i As Integer
        For i = 0 To 7
        grdLista.TextMatrix(1, i) = ""
     Next
     ccanFilas = 0
    End If
 Case 7
 'no se efectua ningún cambio si el usuario presiona NO
 End Select
End If
  '*****************FIN AGREGADO
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
    Dim i As Integer: Dim IniTablas As Integer
    Dim oPersona As UPersona
    
    
    Set fs = New Scripting.FileSystemObject
    Set xlsAplicacion = New Excel.Application
    Set oPersona = New UPersona
    
    lsNomHoja = "Hoja1"
    lsFile = "ExpedientesRRHH"
    'Se modificó la extension xls a xlsx pti1 ers029 26042018
    lsArchivo = "\spooler\" & "ExpedientesRRHH" & "_" & gsCodUser & "_" & Format(gdFecSis, "yyyymmdd") & "_" & Format(Time, "hhmmss") & ".xlsx"
    If fs.FileExists(App.path & "\FormatoCarta\" & lsFile & ".xlsx") Then
        Set xlsLibro = xlsAplicacion.Workbooks.Open(App.path & "\FormatoCarta\" & lsFile & ".xlsx")
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
    
    xlHoja1.Cells(IniTablas + 2, 3) = Me.ctrRRHHGen.psCodigoPersona
    xlHoja1.Cells(IniTablas + 3, 3) = Me.ctrRRHHGen.psNombreEmpledo
    xlHoja1.Cells(IniTablas + 4, 3) = Me.lblNroID.Caption
    xlHoja1.Cells(IniTablas + 5, 3) = Me.lblCargo.Caption
    
    IniTablas = 7
    For i = 1 To grdLista.Rows - 1
        xlHoja1.Cells(IniTablas + i, 2) = grdLista.TextMatrix(i, 1)
        xlHoja1.Cells(IniTablas + i, 3) = grdLista.TextMatrix(i, 2)
        xlHoja1.Cells(IniTablas + i, 4) = grdLista.TextMatrix(i, 3)
        xlHoja1.Cells(IniTablas + i, 5) = grdLista.TextMatrix(i, 4)
        xlHoja1.Cells(IniTablas + i, 6) = grdLista.TextMatrix(i, 5)
        xlHoja1.Cells(IniTablas + i, 7) = grdLista.TextMatrix(i, 6)
    Next i
    
    'xlHoja1.Range(xlHoja1.Cells(6, 2), xlHoja1.Cells(I + 5, 6)).Borders.LineStyle = 1 COMENTADO POR PTI1 ERS029-2018
    xlHoja1.Range(xlHoja1.Cells(7, 2), xlHoja1.Cells(i + 6, 6)).Borders.LineStyle = 1 ' AGREGADO POR PTI1 ERS029-2018
    Dim psArchivoAGrabarC As String
    
    xlHoja1.SaveAs App.path & lsArchivo
    psArchivoAGrabarC = App.path & lsArchivo
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlHoja1 = Nothing
    
End Sub

Private Sub cmdNuevo_Click()
    '**************INICIO MODIFICADO POR PTI1 ERS029-2018 18042018
    If Me.ctrRRHHGen.psCodigoPersona = "" Then
    MsgBox "Debe seleccionar a la persona, para agregar documento", vbExclamation, "Aviso"  'agregado por pti1 ers029-2018
    'ctrRRHHGen_EmiteDatos 'comentado por pti1 ers029-2018
    Else
        pcOpeEdit = False
        frmRRHHRegistroDocumento.Ini gTipoOpeRegistro, "REGISTRAR NUEVO EXPEDIENTE", Me
        CargarExpedienteUser (cPersCod)
        If ccanFilas > 0 Then     'agregado por pti1 ers029-2018
        Set grdLista.Recordset = CargarExpedienteUser(cPersCod)
        grdListaAddDetalle 'AGREGADO POR PTI1 ERS029-2108
        End If
    End If
    '***************FIN MODIFICADO
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub



Private Sub ctrRRHHGen_EmiteDatos()
    Dim oPersona As UPersona
    Dim oRRHH As DActualizaDatosRRHH
    Dim oAcceso As UAcceso
    Set oAcceso = New UAcceso
 
    Set oRRHH = New DActualizaDatosRRHH
    Set oPersona = New UPersona
    Set oPersona = frmBuscaPersona.Inicio(True)
    LimpiaFlex grdLista
   If Not oPersona Is Nothing Then
        'ClearScreen
        cPersCod = oPersona.sPersCod
        cPersCodExp = oPersona.sPersCod
        GetUserPersona (cPersCodExp)
       
        Me.ctrRRHHGen.psCodigoPersona = cUserExpediente
        Me.ctrRRHHGen.psNombreEmpledo = oPersona.sPersNombre
        Me.ctrRRHHGen.psCodigoEmpleado = oRRHH.GetCodigoEmpleado(cPersCodExp)
        Me.lblNroID = oPersona.sPersIdnroDNI
        Me.lblCargo = oRRHH.GetCargo(cPersCodExp, Format(Date, "yyyyMMdd"))
         LimpiaFlex grdLista
        CargarExpedienteUser (cPersCod)
        '******INICIO AGREGADO POR PTI1 ERS029-2108
        If ccanFilas > 0 Then
        Set grdLista.Recordset = CargarExpedienteUser(cPersCod)
        grdListaAddDetalle
        Else
        LimpiaFlex grdLista
        End If
        '******FIN AGREGADO
     
        
        
    End If
End Sub

Public Sub Ini(pnTipo As TipoOpe, psCaption As String)
    lnTipo = pnTipo
    Caption = psCaption
    Me.Show 1
    'Set grdLista.Recordset = CargarExpedienteUser("RECO")
End Sub

Public Sub GetUserPersona(cPersCod As String)
    Dim Sql As String
    Dim rs As New ADODB.Recordset
    Dim Conn As New DConecta
    bError = False
    Sql = "SELECT RH.cUser  FROM RRHH RH WHERE RH.cPersCod =  '" & cPersCod & "'"
    'Set Conn = New COMConecta.DCOMConecta
    If Not Conn.AbreConexion() Then
        bError = True
        sMsgError = "No se pudo Conectar al Servidor, Consulte con el Area de Sistemas"
        Set Conn = Nothing
        Exit Sub
    End If
    Conn.ConexionActiva.CommandTimeout = 7200
    Set rs = Conn.CargaRecordSet(Sql)
    With rs
   
     cUserExpediente = "" & rs!cuser
       .MoveNext
    
    End With
    rs.Close
    Conn.CierraConexion
    Set Conn = Nothing
End Sub

Public Function CargarExpedienteUser(cPersCod As String) As ADODB.Recordset
    'Dim Conn As COMConecta.DCOMConecta
    Dim Conn As New DConecta
    bError = False
    If Not Conn.AbreConexion() Then
        bError = True
        sMsgError = "No se pudo Conectar al Servidor, Consulte con el Area de Sistemas"
        Set BuscaCliente = Nothing
        Set Conn = Nothing
        Exit Function
    End If
    'Conn.ConexionActiva.CommandTimeout = 7200
    Set CargarExpedienteUser = Conn.CargaRecordSet("stp_sel_RHExpedienteXUser '" & Trim(cPersCod) & "'")
    ccanFilas = CargarExpedienteUser.RecordCount 'AGREGADO POR PTI1 ER029-2018
    Conn.CierraConexion
    Set Conn = Nothing
End Function

Private Sub grdLista_DblClick()
     Dim path As String
    If grdLista.col = 7 Then 'AGREGADO POR PTI1 ER029-2018
     ObtenerDatosMatrix
     If pcTpoDoc = "" Then
     MsgBox "No existen datos", vbExclamation, "Aviso"
     Else
        path = ""
        'Shell path
        'tema = Me.grdLista.TextMatrix(Me.grdLista.Row, 1)
        pcOpeDet = True
        frmRRHHRegistroDocumento.Ini 1, "DETALLE DE EXPEDIENTE", Me
        pcOpeDet = False
    End If
    End If
End Sub

Private Sub grdListaAddDetalle()
  '***** AGREGADO POR PTI1 ER029-2018*****************
    Dim i As Integer
     
    ' If ccanFilas = 0 Then
     'grdLista.Rows = 2
    ' grdLista.CantEntero = 8
    ' grdLista.CantDecimales = 2
    ' Else
      For i = 1 To ccanFilas
      grdLista.SetFocus
      grdLista.TextMatrix(i, 7) = "[...]"
      Next
    'End If
'*********FIN AGREGADO *****************************

End Sub


Public Sub ObtenerDatosMatrix()
    pcTpoDoc = Me.grdLista.TextMatrix(Me.grdLista.row, 1)
    pcNroDoc = pcElementos & Me.grdLista.TextMatrix(Me.grdLista.row, 2)
    pdDesde = pcElementos & Me.grdLista.TextMatrix(Me.grdLista.row, 3)
    pdHasta = pcElementos & Me.grdLista.TextMatrix(Me.grdLista.row, 4)
    pcGlosa = pcElementos & Me.grdLista.TextMatrix(Me.grdLista.row, 5)
    pcPathFile = pcElementos & Me.grdLista.TextMatrix(Me.grdLista.row, 6)
End Sub


