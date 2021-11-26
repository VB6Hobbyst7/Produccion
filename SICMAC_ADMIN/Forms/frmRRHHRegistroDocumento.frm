VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRRHHRegistroDocumento 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5340
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   4680
   Icon            =   "frmRRHHRegistroDocumento.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEditar 
      Caption         =   "&Editar"
      Height          =   390
      Left            =   1920
      TabIndex        =   25
      Top             =   4680
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame3 
      Caption         =   "Vigencia"
      Height          =   975
      Left            =   360
      TabIndex        =   2
      Top             =   2280
      Width           =   3975
      Begin MSComCtl2.DTPicker dpkFechaDesde 
         Height          =   315
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         Format          =   70123521
         CurrentDate     =   37056
      End
      Begin MSComCtl2.DTPicker dpkFechaHasta 
         Height          =   315
         Left            =   2040
         TabIndex        =   24
         Top             =   600
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         Format          =   70123521
         CurrentDate     =   37056
      End
      Begin VB.Label lblHasta 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2040
         TabIndex        =   23
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label lblDesde 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   240
         TabIndex        =   22
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Fecha de Vencimiento:"
         Height          =   315
         Left            =   2040
         TabIndex        =   9
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha de Emisión:"
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   1695
      End
   End
   Begin MSComDlg.CommonDialog cdOpenFile 
      Left            =   3960
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Seleccione el Archivo que se Adjuntara"
      Filter          =   "*.pdf"
      InitDir         =   "c:\"
   End
   Begin VB.CommandButton cmdExplorar 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   3840
      TabIndex        =   14
      Top             =   1800
      Width           =   375
   End
   Begin VB.CommandButton cmdRegistrar 
      Caption         =   "&Registrar"
      Height          =   390
      Left            =   2040
      TabIndex        =   13
      Top             =   4680
      Width           =   1110
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   390
      Left            =   3240
      TabIndex        =   12
      Top             =   4680
      Width           =   1110
   End
   Begin VB.Frame Frame1 
      Caption         =   "Glosa"
      Height          =   1215
      Left            =   360
      TabIndex        =   10
      Top             =   3360
      Width           =   3975
      Begin VB.TextBox txtGlosa 
         Height          =   855
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   240
         Width           =   3735
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo de Documento"
      Height          =   1455
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Width           =   3975
      Begin VB.CommandButton cmdAbrirPDF 
         Caption         =   "Ver"
         Height          =   315
         Left            =   1800
         TabIndex        =   18
         Top             =   1080
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtNroDoc 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   720
         Width           =   2775
      End
      Begin VB.ComboBox cboTpoDoc 
         Height          =   315
         ItemData        =   "frmRRHHRegistroDocumento.frx":030A
         Left            =   1080
         List            =   "frmRRHHRegistroDocumento.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label lblNroDoc 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1080
         TabIndex        =   19
         Top             =   720
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Label lblDocumento 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   315
         Left            =   1080
         TabIndex        =   17
         Top             =   360
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Label txtFilePDF 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1080
         TabIndex        =   16
         Top             =   1080
         Width           =   2775
      End
      Begin VB.Label Label3 
         Caption         =   "PDF:"
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Número:"
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Documento:"
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   9128
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Registro de Documento"
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
   Begin VB.Label Label7 
      Caption         =   "Label7"
      Height          =   495
      Left            =   1800
      TabIndex        =   21
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      Height          =   495
      Left            =   1800
      TabIndex        =   20
      Top             =   2400
      Width           =   1215
   End
End
Attribute VB_Name = "frmRRHHRegistroDocumento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'INICIO variables agregado por PTI1 ERS029-2018 18/04/2018 ultimo cambio
 Dim dFechaAct  As String
 Dim nTpoDoc As String
 Public cNumCorrelativoExp As String
 'FIN
    
    
Private cPathFile As String
Private cCodExpediente As String
'COMENTADO POR ARLO 20161221 ************************************
'Private Sub chkHabDes_Click()
'    If chkHabDes.value = 1 Then
'        'chkHabDes.Caption = "Deshabilitar"
'        Me.dpkFechaDesde.Enabled = True
'        Me.dpkFechaHasta.Enabled = True
'        Me.dpkFechaDesde.value = Format(gdFecSis, "dd/mm/yyyy")
'        Me.dpkFechaHasta.value = Format(gdFecSis, "dd/mm/yyyy")
'    Else
'        'chkHabDes.Caption = "Habilitar"
'        Me.dpkFechaDesde.Enabled = False
'        Me.dpkFechaHasta.Enabled = False
'        Me.dpkFechaDesde.value = 0
'        Me.dpkFechaHasta.value = 0
'    End If
'End Sub
'COMENTADO POR ARLO 20161221 ************************************

'AGREGADO POR ARLO 20161221 ************************************
'Private Sub cboTpoDoc_Click()
'Set rs = CargaFechaVencimiento(cboTpoDoc.ItemData(cboTpoDoc.ListIndex))
'If (rs.RecordCount > 0) Then
'txtVencimineto.Text = DateAdd("d", rs!nPeriodoDias, dpkFechaDesde)
'Else
'txtVencimineto.Text = ""
'End If
'End Sub
'AGREGADO POR ARLO 20161221 ************************************

Private Sub cmdAbrirPDF_Click()
    ShellExecute Me.hwnd, "Open", cPathFile, 0&, "", vbNormalFocus
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdEditar_Click()
'AGREGADO POR PTI1 ERS029-2018 18/04/2018
    Dim dFechaRegistro  As String
    Dim Conn As New DConecta
    Dim Sql As String
    Dim cPathFilePDF As String
    Dim dDesde As String: Dim dHasta As String
    cPathFilePDF = txtFilePDF.Caption
    dDesde = Format(dpkFechaDesde.value, "yyyyMMdd")
    dHasta = Format(dpkFechaHasta.value, "yyyyMMdd")
        

    If txtGlosa.Text = "" Then
        MsgBox "Ingrese los datos correctos", vbCritical, "Aviso"
        Exit Sub
    End If
    
    If dDesde > dHasta Then
        MsgBox "La fecha de vencimiento no debe ser menor a la de emisión", vbCritical, "Aviso"
        Exit Sub
    End If
  
    bError = False

     
    Sql = "spt_upd_EditarExpedienteRRHH '" & frmRRHHRegistroExpedientes.cPersCod & "','" & cboTpoDoc.ItemData(cboTpoDoc.ListIndex) & "','" & Trim(txtNroDoc.Text)
    Sql = Sql & "','" & Trim(cPathFilePDF) & "','" & Trim(txtGlosa.Text) & "','" & dDesde & "','" & dHasta & "'"
   
    
    
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
    MsgBox "Expediente editado con exito", vbExclamation, "Aviso"
    Unload Me
    
    'FIN AGREGADO


End Sub

Private Sub cmdExplorar_Click()
    Dim i As Integer
    cdOpenFile.ShowOpen
    cPathFile = cdOpenFile.FileName
    txtFilePDF.Caption = cPathFile
End Sub

Private Sub cmdRegistrar_Click()
    Dim dFechaRegistro  As String
    Dim Conn As New DConecta
    Dim Sql As String
    Dim cPathFilePDF As String
    Dim dDesde As String: Dim dHasta As String
    cPathFilePDF = txtFilePDF.Caption
    'If chkHabDes.value = 1 Then 'COMENTADO POR ARLO2161221
    dDesde = Format(dpkFechaDesde.value, "yyyyMMdd") 'AGREGADO POR PTI ERS029-2018
    dHasta = Format(dpkFechaHasta.value, "yyyyMMdd") 'COMENTADO POR ARLO2161221
    'dHasta = Format(txtVencimineto.Text, "yyyyMMdd") 'AGREGADO POR ARLO2161221 COMENTADO POR  PTI20180904 ERS029-2018
  

        
'COMENTADO POR ARLO2161221
'    Else
'        dDesde = ""
'        dHasta = ""
'    End If
'COMENTADO POR ARLO2161221

    If txtGlosa.Text = "" Then 'AGREGADO POR PTI ERS029-2018
        MsgBox "Ingrese los datos correctamente", vbCritical, "Aviso"
        Exit Sub
    End If
    'AGREGADO POR PTI20180904 ERS029-2018
    If dDesde > dHasta Then
        MsgBox "la fecha de vencimiento no debe ser menor a la de emision", vbCritical, "Aviso"
        Exit Sub
    End If
    'FIN AGREGADO
    
    'dFechaRegistro = gdFecSis & Time
    bError = False
    'Set Conn = New COMConecta.DCOMConecta
    GenerarCodExpediente
    Sql = "spt_ins_RegistrarExpedienteRRHH '" & cCodExpediente & "','" & frmRRHHRegistroExpedientes.cPersCod & "','" & cboTpoDoc.ItemData(cboTpoDoc.ListIndex)
    Sql = Sql & "','" & txtNroDoc.Text & "','" & cPathFilePDF & "','" & txtGlosa.Text & "','" & dDesde & "','" & dHasta & "','" & Format(gdFecSis, "yyyyMMdd") & "'"
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
    MsgBox "Nuevo expediente registrado", vbExclamation, "Aviso"
    limpiarForm
End Sub

Private Sub Form_Load()

' INICIO MODIFICADO POR PTI1 ERS029-201810042018
    If frmRRHHRegistroExpedientes.pcOpeDet = True Then
    
           
        ObtenerElementos (frmRRHHRegistroExpedientes.pcElementos)
        frmRRHHRegistroDocumento.TabStrip1.Tabs(1).Caption = "Detalle de Documento" 'AGREGADO POR PTI1 ERS029-2018 20181004
                   
    Else
    
        If frmRRHHRegistroExpedientes.pcOpeEdit = True Then
        CargarNTpoDoc 'AGREGADO POR PTI1 ERS029-2018 20181004
        ObtenerElementosEdic (frmRRHHRegistroExpedientes.pcElementos)
        frmRRHHRegistroDocumento.TabStrip1.Tabs(1).Caption = "Edicion de Documento" 'AGREGADO POR PTI1 ERS029-2018 20181004
        
        Else
    
        CargarCmbTpoDoc
        limpiarForm
         End If
    End If
    
' FIN MODIFICADO
End Sub


Public Sub Ini(pnTipo As TipoOpe, psCaption As String, FormCerrar As Form)
    lnTipo = pnTipo
    Caption = psCaption
    Me.Show 1
End Sub

Public Sub CargarCmbTpoDoc()
    Dim Sql As String
    Dim rs As New ADODB.Recordset
    'Dim Conn As COMConecta.DCOMConecta
    Dim Conn As New DConecta
    bError = False
    
    Sql = "SELECT  CON.nConsValor,CON.cConsDescripcion FROM Constante CON WHERE CON.nConsCod = '10021'"
    Sql = Sql + "AND CON.nConsValor <> '10021' AND CON.bEstado = 1 ORDER BY CON.nConsValor  ASC"
    'Set Conn = New COMConecta.DCOMConecta
    If Not Conn.AbreConexion() Then
        bError = True
        sMsgError = "No se pudo Conectar al Servidor, Consulte con el Area de Sistemas"
        'Set BuscaCliente = Nothing
        Set Conn = Nothing
        Exit Sub
    End If
    Conn.ConexionActiva.CommandTimeout = 7200
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
End Sub






Public Sub CargarNTpoDoc()
'*******************AGREGADO POR PTI1 20181004
    Dim Sql As String
    Dim rs As New ADODB.Recordset
    'Dim Conn As COMConecta.DCOMConecta
    Dim Conn As New DConecta
    bError = False
    
    Sql = "SELECT  RHEx.NTPODOC FROM RHExpediente RHEx WHERE RHEx.cPersCod = " + frmRRHHRegistroExpedientes.cPersCod
    Sql = Sql + "AND RHEx.cNumDoc='" & Trim(frmRRHHRegistroExpedientes.pcNroDoc) & "'"

    If Not Conn.AbreConexion() Then
        bError = True
        sMsgError = "No se pudo Conectar al Servidor, Consulte con el Area de Sistemas"
        Set Conn = Nothing
        Exit Sub
    End If
   Conn.ConexionActiva.CommandTimeout = 7200
   Set rs = Conn.CargaRecordSet(Sql)
   nTpoDoc = rs!nTpoDoc
    
    rs.Close
    Conn.CierraConexion
    Set Conn = Nothing
'************************ fin agregado
   
End Sub

Public Function GetNameFile(cPath As String) As String
    Dim i As Integer
    For X = Len(cPath) - 1 To 0
        Dim cChar As String
        'cchar =
    Next
End Function

Public Sub GenerarCodExpediente()
    Dim Sql As String
    Dim rs As New ADODB.Recordset
    Dim Conn As New DConecta
    bError = False
    Sql = "SELECT TOP 1 cRHExpedienteCod  FROM RHExpediente  ORDER BY cRHExpedienteCod DESC"
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
    Do Until .EOF
     cCodExpediente = "" & rs!cRHExpedienteCod
       .MoveNext
    Loop
    End With
    rs.Close
    Conn.CierraConexion
    Set Conn = Nothing
    
    If cCodExpediente = "" Then
        cCodExpediente = "100001"
    Else
        cCodExpediente = cCodExpediente + 1
    End If
End Sub

Public Sub limpiarForm()
    
   'cboTpoDoc.ListIndex = 0 ''COMENTADO POR ARLO20161221
   'txtNroDoc.Text = "" COMENTADO POR PTI20180904 ERS029-2018
   'AGREGADO POR PTI20180904 ERS029-2018
     ObtenerNumCorrelativo (frmRRHHRegistroExpedientes.cPersCod)
     cNumCorrelativoExp = cNumCorrelativoExp + 1
     dFechaAct = Replace(Format(Date, "yyyy/mm/dd"), "/", "")
     Select Case cNumCorrelativoExp
     
     Case 0 To 9
     txtNroDoc.Text = "DOC" + dFechaAct + "000" + cNumCorrelativoExp
     Case 10 To 99
     txtNroDoc.Text = "DOC" + dFechaAct + "00" + cNumCorrelativoExp
     Case 100 To 999
     txtNroDoc.Text = "DOC" + dFechaAct + "0" + cNumCorrelativoExp
     
     End Select
     'FIN AGREGADO POR PTI20180904 ERS029-2018
     
    
    txtFilePDF.Caption = ""
'   chkHabDes.value = 1    ''COMENTADO POR ARLO20161221
    Me.dpkFechaDesde.value = Format(gdFecSis, "dd/mm/yyyy")
    Me.dpkFechaHasta.value = Format(gdFecSis, "dd/mm/yyyy")
'    Me.dpkFechaHasta.value = Format(gdFecSis, "dd/mm/yyyy") 'COMENTADO POR ARLO20161221
    txtGlosa.Text = ""
   ' txtVencimineto.Enabled = False 'AGREGADO POR ARLO20161221 COMENTADO POR PTI20180904
   
End Sub

Public Sub ObtenerElementos(pcCadena As String)
    lblDocumento.Caption = frmRRHHRegistroExpedientes.pcTpoDoc
    lblNroDoc.Caption = frmRRHHRegistroExpedientes.pcNroDoc
    cPathFile = frmRRHHRegistroExpedientes.pcPathFile
    lblDesde.Caption = frmRRHHRegistroExpedientes.pdDesde
    lblHasta.Caption = frmRRHHRegistroExpedientes.pdHasta
    txtGlosa.Text = frmRRHHRegistroExpedientes.pcGlosa
    
    lblDocumento.Visible = True
    lblNroDoc.Visible = True
    cmdAbrirPDF.Visible = True
    lblDesde.Visible = True
    lblHasta.Visible = True
    
    If cPathFile = "---" Then
        cmdAbrirPDF.Enabled = False
    End If
    cmdRegistrar.Visible = False 'Agregado por pti1 20180411
    cboTpoDoc.Visible = False
    txtNroDoc.Visible = False
    txtFilePDF.Visible = False
    cmdExplorar.Visible = False
    dpkFechaDesde.Visible = False
    cmdEditar.Visible = False
 '   txtVencimineto.Visible = False COMENTADO por pti1 20180411
  dpkFechaHasta.Visible = False

    txtGlosa.Enabled = False
 
3120
End Sub
 '******************** agregado por pti1 10042018
    Public Sub ObtenerElementosEdic(pcCadena As String)
    CargarCmbTpoDoc
    Dim i As Integer
    
    
   For i = 0 To cboTpoDoc.ListCount - 1
    If nTpoDoc = i Then
    cboTpoDoc.ListIndex = i - 1
    Exit For
    End If
    Next
    
    txtNroDoc.Text = frmRRHHRegistroExpedientes.pcNroDoc
    cPathFile = frmRRHHRegistroExpedientes.pcPathFile
    dpkFechaDesde.value = frmRRHHRegistroExpedientes.pdDesde
    dpkFechaHasta.value = frmRRHHRegistroExpedientes.pdHasta
    txtFilePDF.Caption = cPathFile
    
    txtGlosa.Text = frmRRHHRegistroExpedientes.pcGlosa
    
  
    txtFilePDF.Visible = True
    cmdExplorar.Visible = True
    dpkFechaHasta.Visible = True
    dpkFechaDesde.Visible = True
    txtGlosa.Enabled = True
    cmdEditar.Visible = True
    If cPathFile = "---" Then
        cmdAbrirPDF.Enabled = True
    End If
    
    lblDocumento.Visible = False
    lblNroDoc.Visible = False
    cmdRegistrar.Visible = False
   
    cmdAbrirPDF.Visible = False

    lblNroDoc.Visible = False
    lblDesde.Visible = False
    lblHasta.Visible = False
    
    
3120
' fin agregado
End Sub


'AGREGADO POR ARLO20161221 *****************************************
Public Function CargaFechaVencimiento(nTpoDoc As Integer) As ADODB.Recordset
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
    Set CargaFechaVencimiento = Conn.CargaRecordSet("spt_Sel_DiasVigenciaExpedienteRRHH '" & Trim(cboTpoDoc.ItemData(cboTpoDoc.ListIndex)) & "'")
    Conn.CierraConexion
    Set Conn = Nothing

End Function

Public Sub ObtenerNumCorrelativo(cPersCod As String)
   'INCIO AGREGADO POR PTI1 ERS029-2018 20181004
   Dim Sql As String
    Dim rs As New ADODB.Recordset
    Dim Conn As New DConecta
    bError = False
    Sql = "stp_sel_RHExpedienteXUserNumeroCorrelativo '" & Trim(cPersCod) & "'"
    'Set Conn = New COMConecta.DCOMConecta
    If Not Conn.AbreConexion() Then
        bError = True
        sMsgError = "No se pudo Conectar al Servidor, Consulte con el Area de Sistemas"
        Set Conn = Nothing
        Exit Sub
    End If
    Conn.ConexionActiva.CommandTimeout = 7200
    Set rs = Conn.CargaRecordSet(Sql)
   
    
    cNumCorrelativoExp = rs.GetString

    rs.Close
    Conn.CierraConexion
    Set Conn = Nothing
   
   
   'FIN AGREGADO
   
   
End Sub

