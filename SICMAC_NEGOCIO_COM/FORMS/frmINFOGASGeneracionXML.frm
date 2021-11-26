VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmINFOGASGeneracionXML 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Generación de Formato I (Archivo XML-InfoGas)"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10470
   Icon            =   "frmINFOGASGeneracionXML.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   10470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkTodos 
      Caption         =   "Todos"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   9480
      TabIndex        =   8
      Top             =   120
      Width           =   855
   End
   Begin SICMACT.FlexEdit feINFOGAS 
      Height          =   4170
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   10260
      _ExtentX        =   18098
      _ExtentY        =   7355
      Cols0           =   18
      FixedCols       =   2
      HighLight       =   2
      AllowUserResizing=   3
      EncabezadosNombres=   $"frmINFOGASGeneracionXML.frx":030A
      EncabezadosAnchos=   "350-3500-550-1000-1200-1200-1200-1200-1200-1200-1200-1200-1200-1200-1200-1200-1200-1200"
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
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-X-2-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X"
      TextStyleFixed  =   4
      ListaControles  =   "0-0-4-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
      EncabezadosAlineacion=   "C-L-C-C-L-C-C-C-C-C-R-R-C-R-R-C-C-L"
      FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0-2-2-0-2-2-0-0-0"
      AvanceCeldas    =   1
      TextArray0      =   "N°"
      lbEditarFlex    =   -1  'True
      lbUltimaInstancia=   -1  'True
      lbFormatoCol    =   -1  'True
      lbPuntero       =   -1  'True
      lbBuscaDuplicadoText=   -1  'True
      ColWidth0       =   345
      RowHeight0      =   300
      CellBackColor   =   -2147483633
   End
   Begin VB.Frame FrOpciones 
      Caption         =   "Opciones"
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
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   4560
      Width           =   10335
      Begin VB.ComboBox cboTpoArchivo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6480
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   550
         Width           =   2055
      End
      Begin VB.CommandButton cmdProcesar 
         Caption         =   "&Procesar"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   1215
      End
      Begin VB.ComboBox cboTpoCofigas 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6480
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   195
         Width           =   2055
      End
      Begin VB.CommandButton cmdExportar 
         Caption         =   "&Exportar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   9000
         TabIndex        =   6
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdBuscarPlaca 
         Caption         =   "&Buscar Placa"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1440
         TabIndex        =   5
         Top             =   360
         Width           =   1215
      End
      Begin VB.CheckBox chkExtorno 
         Caption         =   "Extorno del Envio"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2880
         TabIndex        =   2
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Archivo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4920
         TabIndex        =   4
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo de COFIGAS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4920
         TabIndex        =   3
         Top             =   240
         Width           =   1575
      End
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   4920
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmINFOGASGeneracionXML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fsTpoCOFIGAS  As String
Dim fsTpoArchivo As String
Dim fsPlacaBuscar As String

Private Sub chkTodos_Click()
    Dim i As Integer
    Dim check As String
    check = chkTodos.value
    For i = 1 To feINFOGAS.Rows - 1
        feINFOGAS.TextMatrix(i, 2) = check
    Next
End Sub
Private Sub cmdBuscarPlaca_Click()
    Dim i As Integer
    Dim lsBusqueda As String
    fsPlacaBuscar = frmBuscaPlacaINFOGAS.ObtenerPlaca(fsPlacaBuscar)
    If fsPlacaBuscar <> "" Then
        For i = 1 To feINFOGAS.Rows - 1
            If InStr(1, feINFOGAS.TextMatrix(i, 6), fsPlacaBuscar, vbTextCompare) <> 0 Then
                feINFOGAS.TopRow = i
                feINFOGAS.Row = i
                i = feINFOGAS.Rows - 1
                MsgBox "Se encontró el registro", vbInformation, "Aviso"
            End If
        Next i
    End If
End Sub
Private Sub Form_Load()
    CentraForm Me
    Call ObtenerTipoCOFIGAS
    Call ObtenerTipoArchivo
End Sub
Private Sub Inicia()

End Sub
Private Sub cmdProcesar_Click()
    If Not ValidaControles Then
        Exit Sub
    End If
    
    Dim sSql As String
    Dim oConecta As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
'    sSql = "Select ISNULL(Nombre,'') Nombre,ISNULL(TipoDoc,'') TipoDoc,ISNULL(NumDoc,'') NumDoc,ISNULL(TipoVehi,'') TipoVehi,ISNULL(Placa,'') Placa,ISNULL(taller,'') taller,ISNULL(solicitud,'') solicitud,ISNULL(codcli,'') codcli,ISNULL(porcentaje,'') porcentaje,ISNULL(cuoini,'') cuoini,ISNULL(fecha,'') fecha,ISNULL(montopresu,'') montopresu,ISNULL(montocredi,'') montocredi,ISNULL(tipocredi,'') tipocredi,ISNULL(concesionario,'') concesionario,ISNULL(nuevaplaca,'') nuevaplaca,ISNULL(estado,'') estado from GAS " _
'         & "Where Estado = " & (cboTpoArchivo.ListIndex + 1) & " And TipoVehi = " & (cboTpoCofigas.ListIndex + 1)
    sSql = "Exec stp_sel_ObtieneDatosFormatoI_Infogas '" & IIf(cboTpoArchivo.ListIndex = 0, "0", "1") & "'"
    
    fsTpoCOFIGAS = Right(cboTpoCofigas.Text, 2)
    fsTpoArchivo = Right(cboTpoArchivo.Text, 2)
    
    Set oConecta = New COMConecta.DCOMConecta
    Set rs = New ADODB.Recordset
    oConecta.AbreConexion
    Set rs = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
    
    FormatearGrillaINFOGAS
    chkTodos.value = 0
    If Not RSVacio(rs) Then
        Do While Not rs.EOF
            feINFOGAS.AdicionaFila
            feINFOGAS.TextMatrix(feINFOGAS.Row, 1) = rs!cPersNombre 'Nombre Cliente
            feINFOGAS.TextMatrix(feINFOGAS.Row, 3) = rs!cTipoDoc 'Tipo Documento
            feINFOGAS.TextMatrix(feINFOGAS.Row, 4) = rs!cPersIDnro 'Num Documento
            feINFOGAS.TextMatrix(feINFOGAS.Row, 5) = rs!nTipoVehiculo 'Tipo Vehiculo
            feINFOGAS.TextMatrix(feINFOGAS.Row, 6) = rs!cPlaca 'N° Placa
            feINFOGAS.TextMatrix(feINFOGAS.Row, 7) = rs!nTallerCod 'Id Taller
            feINFOGAS.TextMatrix(feINFOGAS.Row, 8) = rs!cCtaCod 'Id Solicitud
            feINFOGAS.TextMatrix(feINFOGAS.Row, 9) = rs!cPersCod 'Cod Cliente
            feINFOGAS.TextMatrix(feINFOGAS.Row, 10) = rs!nRecaudo 'Porcentaje
            feINFOGAS.TextMatrix(feINFOGAS.Row, 11) = rs!nCuotaInicial 'Cuota Inicial
            feINFOGAS.TextMatrix(feINFOGAS.Row, 12) = rs!dFecAprob 'Fecha
            feINFOGAS.TextMatrix(feINFOGAS.Row, 13) = rs!nMontoAprobado 'Monto Presupuesto
            feINFOGAS.TextMatrix(feINFOGAS.Row, 14) = rs!nMontoCredito 'Monto Credito
            feINFOGAS.TextMatrix(feINFOGAS.Row, 15) = rs!cNuevaPlaca 'N° Placa Nueva
            feINFOGAS.TextMatrix(feINFOGAS.Row, 16) = rs!nTipoCredito 'Tipo Credito
            feINFOGAS.TextMatrix(feINFOGAS.Row, 17) = rs!nConcesionarioCod 'Id Consecionario
            rs.MoveNext
        Loop
        HabilitarControles (True)
    Else
        MsgBox "No se encontraron datos", vbInformation, "Aviso"
        HabilitarControles (False)
    End If
End Sub
Private Sub CmdExportar_Click()
    Dim oConst As New COMDConstSistema.NCOMConstSistema
    Dim fso As Scripting.FileSystemObject
    Dim stXML As TextStream
    Dim lsArchivoXML As String, lsFilePathIni As String, sCodCMACMaynas As String
    Dim lsFileName As String
    Dim lsFilePathFin As String
    Dim i As Integer
    Dim bCheck As Boolean
    bCheck = False
    For i = 1 To feINFOGAS.Rows - 1
        If feINFOGAS.TextMatrix(i, 2) = "." Then bCheck = True
    Next
    
    If Not bCheck Then
        MsgBox "Seleccione los registros que desea exportar", vbInformation, "Aviso"
        Exit Sub
    End If
    
    If MsgBox("¿Esta seguro que desea exportar el archivo?", vbYesNo + vbInformation, "Aviso") = vbNo Then
        Exit Sub
    End If
    
    lsFileName = "CMACMAYNAS" & Format(gdFecSis, "yyyymmdd") & gsCodUser & Format(Time(), "HHMMSS") & ".xml"
    lsFilePathIni = App.path & "\Spooler\" & lsFileName
    Set fso = New Scripting.FileSystemObject
    Set stXML = fso.CreateTextFile(lsFilePathIni, True)
    sCodCMACMaynas = 924 'oConst.LeeConstSistema(gConstSistCodCMAC)
    
    stXML.WriteLine "<?xml version=""1.0"" encoding=""ISO-8859-1""?>"
    stXML.WriteLine "<Financiera Codigo=""" & sCodCMACMaynas & """>"
    Select Case fsTpoArchivo
        Case "01"
            For i = 1 To feINFOGAS.Rows - 1
                If feINFOGAS.TextMatrix(i, 2) = "." Then
                    stXML.WriteLine "<Registro Tipo=""" & fsTpoArchivo & """>"
                    stXML.WriteLine "<NombreCliente>" & feINFOGAS.TextMatrix(i, 1) & "</NombreCliente>"
                    stXML.WriteLine "<TipoDocumento>" & feINFOGAS.TextMatrix(i, 3) & "</TipoDocumento>"
                    stXML.WriteLine "<NumeroDocumento>" & feINFOGAS.TextMatrix(i, 4) & "</NumeroDocumento>"
                    stXML.WriteLine "<TipoVehiculo>" & feINFOGAS.TextMatrix(i, 5) & "</TipoVehiculo>"
                    stXML.WriteLine "<Placa>" & feINFOGAS.TextMatrix(i, 6) & "</Placa>"
                    stXML.WriteLine "<Taller>" & feINFOGAS.TextMatrix(i, 7) & "</Taller>"
                    stXML.WriteLine "<Solicitud>" & feINFOGAS.TextMatrix(i, 8) & "</Solicitud>"
                    stXML.WriteLine "<CodigoCliente>" & feINFOGAS.TextMatrix(i, 9) & "</CodigoCliente>"
                    stXML.WriteLine "<Porcentaje>" & feINFOGAS.TextMatrix(i, 10) & "</Porcentaje>"
                    stXML.WriteLine "<CuotaInicial>" & feINFOGAS.TextMatrix(i, 11) & "</CuotaInicial>"
                    stXML.WriteLine "<Fecha>" & feINFOGAS.TextMatrix(i, 12) & "</Fecha>"
                    stXML.WriteLine "<MontoPresupuesto>" & feINFOGAS.TextMatrix(i, 13) & "</MontoPresupuesto>"
                    stXML.WriteLine "<MontoCredito>" & feINFOGAS.TextMatrix(i, 14) & "</MontoCredito>"
                    stXML.WriteLine "<TipoCredito>" & feINFOGAS.TextMatrix(i, 16) & "</TipoCredito>"
                    stXML.WriteLine "<Concesionario>" & feINFOGAS.TextMatrix(i, 17) & "</Concesionario>"
                    stXML.WriteLine "</Registro>"
                End If
            Next
        Case "02"
            For i = 1 To feINFOGAS.Rows - 1
                If feINFOGAS.TextMatrix(i, 2) = "." Then
                    stXML.WriteLine "<Registro Tipo=""" & fsTpoArchivo & """>"
                    stXML.WriteLine "<NombreCliente>" & feINFOGAS.TextMatrix(i, 1) & "</NombreCliente>"
                    stXML.WriteLine "<TipoDocumento>" & feINFOGAS.TextMatrix(i, 3) & "</TipoDocumento>"
                    stXML.WriteLine "<NumeroDocumento>" & feINFOGAS.TextMatrix(i, 4) & "</NumeroDocumento>"
                    stXML.WriteLine "<TipoVehiculo>" & feINFOGAS.TextMatrix(i, 5) & "</TipoVehiculo>"
                    stXML.WriteLine "<Placa>" & feINFOGAS.TextMatrix(i, 6) & "</Placa>"
                    stXML.WriteLine "<Taller>" & feINFOGAS.TextMatrix(i, 7) & "</Taller>"
                    stXML.WriteLine "<Solicitud>" & feINFOGAS.TextMatrix(i, 8) & "</Solicitud>"
                    stXML.WriteLine "<CodigoCliente>" & feINFOGAS.TextMatrix(i, 9) & "</CodigoCliente>"
                    stXML.WriteLine "<Porcentaje>" & feINFOGAS.TextMatrix(i, 10) & "</Porcentaje>"
                    stXML.WriteLine "<CuotaInicial>" & feINFOGAS.TextMatrix(i, 11) & "</CuotaInicial>"
                    stXML.WriteLine "<Fecha>" & feINFOGAS.TextMatrix(i, 12) & "</Fecha>"
                    stXML.WriteLine "<MontoPresupuesto>" & feINFOGAS.TextMatrix(i, 13) & "</MontoPresupuesto>"
                    stXML.WriteLine "<MontoCredito>" & feINFOGAS.TextMatrix(i, 14) & "</MontoCredito>"
                    stXML.WriteLine "<PlacaNueva>" & feINFOGAS.TextMatrix(i, 15) & "</PlacaNueva>"
                    stXML.WriteLine "<TipoCredito>" & feINFOGAS.TextMatrix(i, 16) & "</TipoCredito>"
                    stXML.WriteLine "<Concesionario>" & feINFOGAS.TextMatrix(i, 17) & "</Concesionario>"
                    stXML.WriteLine "</Registro>"
                End If
            Next
        Case "03", "04"
            For i = 1 To feINFOGAS.Rows - 1
                If feINFOGAS.TextMatrix(i, 2) = "." Then
                    stXML.WriteLine "<Registro Tipo=""" & fsTpoArchivo & """>"
                    stXML.WriteLine "<Solicitud>" & feINFOGAS.TextMatrix(i, 8) & "</Solicitud>"
                    stXML.WriteLine "<CodigoCliente>" & feINFOGAS.TextMatrix(i, 9) & "</CodigoCliente>"
                    stXML.WriteLine "<Placa>" & feINFOGAS.TextMatrix(i, 6) & "</Placa>"
                    stXML.WriteLine "</Registro>"
                End If
            Next
        Case "05"
            For i = 1 To feINFOGAS.Rows - 1
                If feINFOGAS.TextMatrix(i, 2) = "." Then
                    stXML.WriteLine "<Registro Tipo=""" & fsTpoArchivo & """>"
                    stXML.WriteLine "<Solicitud>" & feINFOGAS.TextMatrix(i, 8) & "</Solicitud>"
                    stXML.WriteLine "<CodigoCliente>" & feINFOGAS.TextMatrix(i, 9) & "</CodigoCliente>"
                    stXML.WriteLine "</Registro>"
                End If
            Next
    End Select
    stXML.Write "</Financiera>"
    stXML.Close
    'MsgBox lsFile & " Archivo generado satisfactoriamente", vbInformation, "Aviso"
    'ShellExecute Me.hwnd, "open", lsFile, "", "", 4

    With dlgCommonDialog
        .DialogTitle = "Guardar"
        .CancelError = False
        .Filename = lsFileName
        .Filter = "Archivo XML (*.xml)|*.xml"
        .ShowSave
        If Len(.Filename) = 0 Then
            MsgBox "Ud. NO ha seleccionado donde guardar el archivo", vbInformation, "Aviso"
            Exit Sub
        End If
        lsFilePathFin = .Filename
    End With

    fso.CopyFile lsFilePathIni, lsFilePathFin
    MsgBox "Archivo generado satisfactoriamente en " & Chr(10) & lsFilePathFin, vbInformation, "Aviso"
    
    FormatearGrillaINFOGAS
    chkTodos.value = 0
    HabilitarControles (False)
    Set fso = Nothing
End Sub
Private Sub ObtenerTipoCOFIGAS()
    cboTpoCofigas.Clear
    cboTpoCofigas.AddItem "VEHICULO NUEVO" & Space(100) & "01"
    cboTpoCofigas.AddItem "VEHICULO USADO" & Space(100) & "02"
End Sub
Private Sub ObtenerTipoArchivo()
    cboTpoArchivo.Clear
    cboTpoArchivo.AddItem "NUEVOS" & Space(100) & "01"
    cboTpoArchivo.AddItem "MODIFICACION" & Space(100) & "02"
    cboTpoArchivo.AddItem "BLOQUEO DEL CHIP" & Space(100) & "03"
    cboTpoArchivo.AddItem "ACTIVACION DEL CHIP" & Space(100) & "04"
    cboTpoArchivo.AddItem "CANCELACION" & Space(100) & "05"
End Sub
Private Sub FormatearGrillaINFOGAS()
    feINFOGAS.Clear
    feINFOGAS.FormaCabecera
    feINFOGAS.Rows = 2
End Sub
Private Function ValidaControles() As Boolean
    If cboTpoCofigas.ListIndex = -1 Then
        MsgBox "Falta Seleccionar el tipo de COFIGAS", vbInformation, "Aviso"
        cboTpoCofigas.SetFocus
        ValidaControles = False
        Exit Function
    End If
    If cboTpoArchivo.ListIndex = -1 Then
        MsgBox "Falta Seleccionar el tipo de Archivo a Exportar", vbInformation, "Aviso"
        cboTpoArchivo.SetFocus
        ValidaControles = False
        Exit Function
    End If
    ValidaControles = True
End Function
Private Sub HabilitarControles(ByVal pbHabilita As Boolean)
    cmdExportar.Enabled = pbHabilita
    cmdBuscarPlaca.Enabled = pbHabilita
    chkTodos.Enabled = pbHabilita
End Sub
