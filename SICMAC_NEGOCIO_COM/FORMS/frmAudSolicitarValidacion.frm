VERSION 5.00
Begin VB.Form frmAudSolicitarValidacion 
   Caption         =   "SOLICITAR VALIDACIÓN"
   ClientHeight    =   4080
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5325
   Icon            =   "frmAudSolicitarValidacion.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   4080
   ScaleWidth      =   5325
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRegistrar 
      Caption         =   "Registrar"
      Height          =   375
      Left            =   3120
      TabIndex        =   9
      Top             =   3480
      Width           =   855
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   4080
      TabIndex        =   8
      Top             =   3480
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Solicitar Validación"
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5055
      Begin VB.CommandButton cmdHistorial 
         Caption         =   "Ver Historial"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   3360
         Width           =   1215
      End
      Begin VB.TextBox txtConclusion 
         Height          =   855
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   2400
         Width           =   4575
      End
      Begin VB.TextBox txtComentario 
         Height          =   855
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   1200
         Width           =   4575
      End
      Begin VB.Label Label4 
         Caption         =   "Conclusión :"
         Height          =   315
         Left            =   240
         TabIndex        =   5
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Comentario :"
         Height          =   315
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label lblProcedimiento 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1440
         TabIndex        =   2
         Top             =   480
         Width           =   3375
      End
      Begin VB.Label Label1 
         Caption         =   "Procedimiento :"
         Height          =   315
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmAudSolicitarValidacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Public npProcedimientoID As Integer
'Dim slProcedimientoNombre As String
'
'Private Sub cmdCancelar_Click()
'    Unload Me
'End Sub
'
'Private Sub cmdHistorial_Click()
'    frmAudHistorialSolicitudVal.pnProcedimientoID = npProcedimientoID
'    frmAudHistorialSolicitudVal.Show 1
'End Sub
'
'Private Sub cmdRegistrar_Click()
'    Dim objCOMNAuditoria  As COMNAuditoria.NCOMRegistros
'    Set objCOMNAuditoria = New COMNAuditoria.NCOMRegistros
'
'    objCOMNAuditoria.RegistrarValidacionProcedimiento objCOMNAuditoria.ObtenerCodigoValidacionProcedimiento, npProcedimientoID, _
'                                                      txtComentario.Text, txtConclusion.Text, gdFecSis
'    MsgBox "Solicitud de Validacion Registrada", vbInformation, "Aviso"
'    Unload Me
'End Sub
'
'Private Sub Form_Load()
'    npProcedimientoID = frmAudDesarrolloProcedimiento.nProcedimientoID
'    slProcedimientoNombre = frmAudDesarrolloProcedimiento.sProcedimientoNombre
'    lblProcedimiento.Caption = slProcedimientoNombre
'End Sub
'
'Public Sub LimpiarFomr()
'    txtComentario.Text = ""
'    txtConclusion.Text = ""
'End Sub
