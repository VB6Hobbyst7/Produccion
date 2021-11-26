VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAuditoriaRegistrarAcuerdo 
   Caption         =   "Registrar Acuerdo de Directorio"
   ClientHeight    =   7650
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7170
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAuditoriaRegistrarAcuerdo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7650
   ScaleWidth      =   7170
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   7335
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      Begin VB.CommandButton cmdActualizar 
         Caption         =   "Actualizar"
         Height          =   350
         Left            =   2760
         TabIndex        =   20
         Top             =   6720
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmbCancelar 
         Caption         =   "Cancelar"
         Height          =   350
         Left            =   5280
         TabIndex        =   19
         Top             =   6720
         Width           =   975
      End
      Begin VB.ComboBox cmbTipoSesion 
         Height          =   315
         ItemData        =   "frmAuditoriaRegistrarAcuerdo.frx":030A
         Left            =   960
         List            =   "frmAuditoriaRegistrarAcuerdo.frx":0314
         TabIndex        =   1
         Top             =   360
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker dtFecha 
         Height          =   345
         Left            =   4920
         TabIndex        =   3
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   609
         _Version        =   393216
         Format          =   86573057
         CurrentDate     =   40235
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "Aceptar"
         Height          =   350
         Left            =   4080
         TabIndex        =   16
         Top             =   6720
         Width           =   975
      End
      Begin VB.TextBox txtNroSesion 
         Height          =   350
         Left            =   3000
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
      Begin VB.Frame Frame2 
         Caption         =   "ACUERDO"
         Height          =   5535
         Left            =   240
         TabIndex        =   11
         Top             =   960
         Width           =   6015
         Begin VB.ComboBox cmbSituacion 
            Height          =   315
            ItemData        =   "frmAuditoriaRegistrarAcuerdo.frx":0333
            Left            =   4200
            List            =   "frmAuditoriaRegistrarAcuerdo.frx":0340
            TabIndex        =   7
            Top             =   3600
            Width           =   1575
         End
         Begin VB.TextBox txtSituacionComentario 
            Height          =   1095
            Left            =   1080
            MultiLine       =   -1  'True
            TabIndex        =   8
            Top             =   4080
            Width           =   4695
         End
         Begin VB.TextBox txtDetalle 
            Height          =   1095
            Left            =   1080
            MultiLine       =   -1  'True
            TabIndex        =   6
            Top             =   2280
            Width           =   4695
         End
         Begin VB.TextBox txtAcuerdoNro 
            Height          =   350
            Left            =   4680
            MultiLine       =   -1  'True
            TabIndex        =   5
            Top             =   1680
            Width           =   1095
         End
         Begin VB.TextBox txtAsunto 
            Height          =   1095
            Left            =   1080
            MultiLine       =   -1  'True
            TabIndex        =   4
            Top             =   360
            Width           =   4695
         End
         Begin VB.Label Label6 
            Caption         =   "Situación:"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   3600
            Width           =   855
         End
         Begin VB.Label Label5 
            Caption         =   "Detalle:"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   2280
            Width           =   615
         End
         Begin VB.Label Label4 
            Caption         =   "Acuerdo Nº:"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   1680
            Width           =   1095
         End
         Begin VB.Label Label2 
            Caption         =   "Asunto:"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.Label Label8 
         Caption         =   "Nº:"
         Height          =   255
         Left            =   2640
         TabIndex        =   18
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label7 
         Caption         =   "Fecha:"
         Height          =   255
         Left            =   4320
         TabIndex        =   17
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Sesión:"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmAuditoriaRegistrarAcuerdo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Dim objCOMNAuditoria As COMNAuditoria.NCOMSeguimiento
'
'Private Sub cmdActualizar_Click()
'    Set objCOMNAuditoria = New COMNAuditoria.NCOMSeguimiento
'
'    If MsgBox("Esta Seguro de Actualizar los Datos?", vbYesNo + vbQuestion, Me.Caption) = vbYes Then
'
'        objCOMNAuditoria.ActualizarSesionDir gSesionDirId, cmbTipoSesion.Text, txtNroSesion.Text, dtFecha.value
'        objCOMNAuditoria.ActualizarAcuerdoDir gSesionDirId, gNroAcuerdo, txtAsunto.Text, txtAcuerdoNro.Text, txtDetalle.Text, cmbSituacion.Text, txtSituacionComentario.Text
'
'        MsgBox "Los Datos se Actualizaron Correctamente", vbInformation, Me.Caption
'        gSesionDirId = 0
'        txtNroSesion.Text = ""
'        Limpiar
'
'    End If
'End Sub
'
'Private Sub Form_Load()
'    If gSesionDirId <> 0 Then
'        CargarDatosModificar
'        cmdAceptar.Visible = False
'        cmdActualizar.Visible = True
'    Else
'        dtFecha.value = Date
'    End If
'End Sub
'
'Private Sub cmdAceptar_Click()
'    Set objCOMNAuditoria = New COMNAuditoria.NCOMSeguimiento
'
'    If MsgBox("Esta Seguro de Registar los Datos?", vbYesNo + vbQuestion, Me.Caption) = vbYes Then
'
'        If gSesionDirId = 0 Then
'            gSesionDirId = CInt(objCOMNAuditoria.InsertarSesionDir(cmbTipoSesion.Text, txtNroSesion.Text, dtFecha.value))
'            objCOMNAuditoria.InsertarAcuerdoDir gSesionDirId, txtAsunto.Text, txtAcuerdoNro.Text, txtDetalle.Text, cmbSituacion.Text, txtSituacionComentario.Text
'        Else
'            objCOMNAuditoria.InsertarAcuerdoDir gSesionDirId, txtAsunto.Text, txtAcuerdoNro.Text, txtDetalle.Text, cmbSituacion.Text, txtSituacionComentario.Text
'        End If
'
'        MsgBox "Los Datos se registraron correctamente", vbInformation
'
'        If MsgBox("Desea Seguir Registrando Acuerdos?", vbYesNo + vbQuestion, Me.Caption) = vbYes Then
'            Limpiar
'        Else
'            gSesionDirId = 0
'            txtNroSesion.Text = ""
'            Limpiar
'        End If
'
'    End If
'
'End Sub
'
'Private Sub cmbCancelar_Click()
'    gSesionDirId = 0
'    txtNroSesion.Text = ""
'    Limpiar
'End Sub
'
'Private Sub Limpiar()
'    txtAsunto.Text = ""
'    txtAcuerdoNro.Text = ""
'    txtDetalle.Text = ""
'    txtSituacionComentario.Text = ""
'End Sub
'
'Public Sub CargarDatosModificar()
'    Dim objCOMNAuditoria As COMNAuditoria.NCOMSeguimiento
'    Set objCOMNAuditoria = New COMNAuditoria.NCOMSeguimiento
'    Dim rs As ADODB.Recordset
'    Set rs = objCOMNAuditoria.ObtenerSesionDirXId(gSesionDirId, gNroAcuerdo)
'
'    cmbTipoSesion.Text = rs("vTipoSesion")
'    txtNroSesion.Text = rs("vNroSesion")
'    dtFecha.value = rs("vFecha")
'
'    txtAsunto.Text = rs("tAsunto")
'    txtAcuerdoNro.Text = rs("vAcuerdo")
'    txtDetalle.Text = rs("tDetalle")
'    cmbSituacion.Text = rs("vSituacion")
'    txtSituacionComentario.Text = rs("tSituacion")
'
'End Sub
'
'Private Sub Form_Unload(Cancel As Integer)
'    gSesionDirId = 0
'End Sub
