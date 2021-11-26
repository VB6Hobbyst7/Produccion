VERSION 5.00
Begin VB.Form frmAudVerificarProcedimiento 
   Caption         =   "AUDITORIA: VARIFICAR PROCEDIMIENTO"
   ClientHeight    =   4290
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6045
   Icon            =   "frmAudVerificarProcedimiento.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   4290
   ScaleWidth      =   6045
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Verificar Procedimiento"
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5775
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   315
         Left            =   4440
         TabIndex        =   10
         Top             =   3600
         Width           =   1095
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "Aceptar"
         Height          =   315
         Left            =   4440
         TabIndex        =   9
         Top             =   3240
         Width           =   1095
      End
      Begin VB.TextBox txtDevolvelText 
         Height          =   315
         Left            =   240
         TabIndex        =   8
         Top             =   3600
         Width           =   3735
      End
      Begin VB.OptionButton optDevolver 
         Caption         =   "Devolver"
         Height          =   255
         Left            =   1200
         TabIndex        =   7
         Top             =   3360
         Width           =   1095
      End
      Begin VB.OptionButton optAceptar 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   3360
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.Label lblConclusion 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   855
         Left            =   240
         TabIndex        =   12
         Top             =   2280
         Width           =   5295
      End
      Begin VB.Label lblComentario 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   855
         Left            =   240
         TabIndex        =   11
         Top             =   1080
         Width           =   5295
      End
      Begin VB.Label Label5 
         Caption         =   "Resolver Procedimiento:"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Conclusión:"
         Height          =   315
         Left            =   240
         TabIndex        =   4
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Comentario:"
         Height          =   315
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lblProcedimiento 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1440
         TabIndex        =   2
         Top             =   360
         Width           =   4095
      End
      Begin VB.Label Label1 
         Caption         =   "Procedimiento :"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmAudVerificarProcedimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim nProcedimientoID As Integer
'Dim sValidacionCod As String
'Private Sub cmdAceptar_Click()
'    Dim objCOMNAuditoria As COMNAuditoria.NCOMRegistros
'    Set objCOMNAuditoria = New COMNAuditoria.NCOMRegistros
'
'    If optDevolver.value = True Then
'        If txtDevolvelText.Text <> "" Then
'            If MsgBox("Esta seguro que desea devolver el procedimiento", vbOKCancel, "Aviso") = vbOk Then
'                objCOMNAuditoria.ActualizarEstadoValidacionProcedimiento nProcedimientoID, sValidacionCod, 2, txtDevolvelText.Text
'                MsgBox "El procedimiento ha sido devuelto", vbInformation, "Aviso"
'                Unload Me
'            Else
'                MsgBox "Cancelado por el usuario", vbInformation, "Aviso"
'            End If
'        Else
'            MsgBox "Ingrese texto de comentario para devolver", vbInformation, "Aviso"
'        End If
'    Else
'        If MsgBox("Esta seguro que desea aceptar el procedimiento", vbOKCancel, "Aviso") = vbOk Then
'            objCOMNAuditoria.ActualizarEstadoValidacionProcedimiento nProcedimientoID, sValidacionCod, 1, ""
'            MsgBox "El procedimiento ha sido aceptado", vbInformation, "Aviso"
'            Unload Me
'        Else
'            MsgBox "Cancelado por el usuario", vbInformation, "Aviso"
'        End If
'    End If
'End Sub
'
'Private Sub cmdCancelar_Click()
'    Unload Me
'End Sub
'
'Private Sub Form_Load()
'    Dim objCOMNAuditoria As COMNAuditoria.NCOMRegistros
'    Set objCOMNAuditoria = New COMNAuditoria.NCOMRegistros
'    nProcedimientoID = frmAudDesarrolloProcedimientoVerificar.nProcedimientoID
'    objCOMNAuditoria.ObtenerDatosValidacionProcedimiento nProcedimientoID
'    CargarDatosControles objCOMNAuditoria.ObtenerDatosValidacionProcedimiento(nProcedimientoID)
'End Sub
'
'Public Sub CargarDatosControles(ByVal DR As ADODB.Recordset)
'    lblProcedimiento.Caption = DR!cProcedimientoNombre
'    lblComentario.Caption = DR!cComentario
'    lblConclusion.Caption = DR!cConclusion
'    sValidacionCod = DR!cValidacionCod
'End Sub
