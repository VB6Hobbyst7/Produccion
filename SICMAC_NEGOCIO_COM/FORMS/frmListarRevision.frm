VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmListarRevision 
   Caption         =   "Listar Revisión"
   ClientHeight    =   6555
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9885
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmListarRevision.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   9885
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   4575
      Left            =   8160
      TabIndex        =   13
      Top             =   1800
      Width           =   1575
      Begin VB.CommandButton Command4 
         Caption         =   "&Eliminar"
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   2160
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Imprimir"
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Modificar"
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   1680
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Height          =   4575
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   7935
      Begin MSDataGridLib.DataGrid dgBuscar 
         Height          =   4215
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   7665
         _ExtentX        =   13520
         _ExtentY        =   7435
         _Version        =   393216
         AllowUpdate     =   0   'False
         ColumnHeaders   =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         RowDividerStyle =   1
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "iRevisionId"
            Caption         =   "Id"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "cFRegistro"
            Caption         =   "F. Registro"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "cPersNombre"
            Caption         =   "Nombre"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "cFCierre"
            Caption         =   "F. Cierre"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "0%"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "vCAnalista"
            Caption         =   "Analista"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            Size            =   182
            BeginProperty Column00 
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1305.071
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   3195.213
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1305.071
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1500.095
            EndProperty
         EndProperty
      End
      Begin VB.Label lblMensaje 
         Caption         =   "NO EXISTEN DATOS"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         TabIndex        =   12
         Top             =   240
         Visible         =   0   'False
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Parámetros de Búsqueda"
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9615
      Begin VB.TextBox txtTCambio 
         Height          =   285
         Left            =   8100
         TabIndex        =   10
         Top             =   720
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtCliente 
         Height          =   285
         Left            =   1320
         TabIndex        =   7
         Top             =   720
         Width           =   3615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Buscar"
         Height          =   375
         Left            =   8240
         TabIndex        =   2
         Top             =   1080
         Width           =   975
      End
      Begin MSMask.MaskEdBox mskPeriodo1Del 
         Height          =   315
         Left            =   7680
         TabIndex        =   8
         Top             =   360
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin SICMACT.TxtBuscar txtCodigo 
         Height          =   285
         Left            =   1320
         TabIndex        =   9
         Top             =   360
         Width           =   1980
         _extentx        =   3493
         _extenty        =   503
         appearance      =   1
         appearance      =   1
         font            =   "frmListarRevision.frx":030A
         appearance      =   1
         tipobusqueda    =   3
         stitulo         =   ""
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Cambio:"
         Height          =   195
         Left            =   6720
         TabIndex        =   11
         Top             =   720
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.Label Label4 
         Caption         =   "Nombre:"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Cliente:"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "F. Cierre:"
         Height          =   255
         Left            =   6720
         TabIndex        =   4
         Top             =   360
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmListarRevision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''********************************************************************
''** Nombre : frmRevisionRegistrar
''** Descripción : formulario que permitirá el registro del formato de la revision de la calificacion.
''** Creación : MAVM, 20080809 10:00:00 AM
''** Modificación:
''********************************************************************
'
''Option Explicit
'Dim objCOMNAuditoria As COMNAuditoria.NCOMRevision
'Dim FechaFinMes As Date
'Dim lsmensaje As String
'
'Public Sub BuscarDatos()
''On Error GoTo Manejador:
'    Dim rs As ADODB.Recordset
'    Dim contador As Integer
'    Set objCOMNAuditoria = New COMNAuditoria.NCOMRevision
'    lsmensaje = ""
'    Set rs = objCOMNAuditoria.DarRevisionCalificacion(txtCodigo.Text, mskPeriodo1Del, lsmensaje) ', txtTCambio.Text
'        'If objeto.ValidarError Then GoTo Manejador
'        If lsmensaje = "" Then
'            lblMensaje.Visible = False
'            dgBuscar.Visible = True
'            Set dgBuscar.DataSource = rs
'            dgBuscar.Refresh
'            Screen.MousePointer = 0
'            dgBuscar.SetFocus
'            Command2.Visible = True
'            Command3.Visible = True
'            Command4.Visible = True
'        Else
'            Set dgBuscar.DataSource = Nothing
'            dgBuscar.Refresh
'            lblMensaje.Visible = True
'            dgBuscar.Visible = False
'            Command2.Visible = False
'            Command3.Visible = False
'            Command4.Visible = False
'        End If
'        Set rs = Nothing
'        Set objCOMNAuditoria = Nothing
'        'Exit Sub
'
''Manejador:
'    'EnviarManejadorError "BuscarDatos", "frmListarRevision"
'End Sub
'
'Private Sub Command1_Click()
'    If mskPeriodo1Del.Text <> "__/__/____" Then
'        BuscarDatos
'    Else
'        MsgBox ("Por Favor Elegir la Fecha de Cierre del Mes"), vbCritical
'    End If
'End Sub
'
'Private Sub Command2_Click()
'    If lsmensaje = "" Then
'        gRevisionId = dgBuscar.Columns(0).Text
'        frmRevisionRegistrar.Show
'        'Unload Me
'    End If
'End Sub
'
'Private Sub Command3_Click()
'    If lsmensaje = "" Then
'    gRevisionId = dgBuscar.Columns(0).Text
'    frmRevisionRegistrar.ImprimeFormatoRevision
'    gRevisionId = 0
'    End If
'End Sub
'
'Private Sub Command4_Click()
'    Dim objCOMNAuditoria As COMNAuditoria.NCOMRevision
'    Set objCOMNAuditoria = New COMNAuditoria.NCOMRevision
'    If lsmensaje = "" Then
'        gRevisionId = dgBuscar.Columns(0).Text
'        objCOMNAuditoria.EliminarRevisionCalificacion gRevisionId
'        gRevisionId = 0
'        BuscarDatos
'    End If
'End Sub
'
'Private Sub Form_Load()
'    'Set objeto = New COMMANEJADOR.ManejadorError
'    CargarDatos
'    Command2.Visible = False
'    Command3.Visible = False
'    Command4.Visible = False
'End Sub
'
'Public Sub CargarDatos()
'    Dim oTipCambio As nTipoCambio
'    'Format(DateAdd("d", gdFecData, -60), "yyyymmdd")
'    'NR MAVM 20090915
'    Me.mskPeriodo1Del = gdFecData
'    FechaFinMes = gdFecData
'    Me.Icon = LoadPicture(App.Path & gsRutaIcono)
'    Set oTipCambio = New nTipoCambio
'        txtTCambio.Text = Format(oTipCambio.EmiteTipoCambio(gdFecSis, TCFijoMes), "#0.000")
'    Set oTipCambio = Nothing
'End Sub
'
'Private Sub txtCodigo_EmiteDatos()
'    If txtCodigo.Text <> "" Then
'        Call CargarDatosCliente(txtCodigo.Text)
'        Set dgBuscar.DataSource = Nothing
'        dgBuscar.Refresh
'        Command2.Visible = False
'        Command3.Visible = False
'        Command4.Visible = False
'        txtCliente.SetFocus
'    End If
'End Sub
'
'Public Sub CargarDatosCliente(ByVal CodPer As String)
'    Set objCOMNAuditoria = New COMNAuditoria.NCOMRevision
'    Dim rs2 As ADODB.Recordset
'    Dim lsmensaje As String
'    Dim lsCalificacion As String
'    Dim i As Integer
'
'    Set rs2 = objCOMNAuditoria.ObtenerDatosCliente(CodPer)
'
'    If rs2.RecordCount <> 0 Then
'        txtCliente.Text = rs2("cPersNombre")
'    Else
'        MsgBox lsmensaje, vbCritical, "Aviso"
'        txtCodigo.Text = ""
'    End If
'End Sub
