VERSION 5.00
Begin VB.Form frmAudAsignarActividad 
   Caption         =   "Asignar Actividad a Usuario"
   ClientHeight    =   5040
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmAudAsignarActividad.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   5040
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   3360
      TabIndex        =   10
      Top             =   4440
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Asignar Actividad a Usuario"
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.CommandButton cmdAsignar 
         Caption         =   "Asignar"
         Height          =   375
         Left            =   2160
         TabIndex        =   9
         Top             =   4320
         Width           =   975
      End
      Begin VB.TextBox txtObjetivoEsp 
         Height          =   1095
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   3120
         Width           =   3975
      End
      Begin VB.TextBox txtObjetivoGen 
         Height          =   975
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   1800
         Width           =   3975
      End
      Begin VB.ComboBox cboUsuarios 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   960
         Width           =   3255
      End
      Begin VB.Label Label5 
         Caption         =   "Objetivo Especifico :"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Objetivo General :"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Usuario :"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   735
      End
      Begin VB.Label lblActividad 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   960
         TabIndex        =   2
         Top             =   480
         Width           =   3255
      End
      Begin VB.Label Label1 
         Caption         =   "Actividad :"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmAudAsignarActividad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim sActividadCod As String
'Dim sArregloCodPersona() As String
'
'Private Sub cmdAsignar_Click()
'    If txtObjetivoGen.Text <> "" And txtObjetivoEsp.Text <> "" Then
'        Dim objCOMNAuditoria As COMNAuditoria.NCOMRegistros
'        Set objCOMNAuditoria = New COMNAuditoria.NCOMRegistros
'
'        objCOMNAuditoria.RegistrarAsignacionActividadUsuario sActividadCod, sArregloCodPersona(cboUsuarios.ItemData(cboUsuarios.ListIndex)), txtObjetivoGen.Text _
'        , txtObjetivoEsp.Text, 1, gdFecSis
'
'        Unload Me
'    Else
'        MsgBox "Los datos no puedes ser vacios", vbCritical, "Aviso"
'    End If
'End Sub
'
'Private Sub cmdCerrar_Click()
'    Unload Me
'End Sub
'
'Private Sub Form_Load()
'    lblActividad.Caption = frmAudRegistroActividadProgramada.sActividadDesc
'    sActividadCod = frmAudRegistroActividadProgramada.sActividadCod
'    CargarColaboradoresUAI
'End Sub
'
'Public Sub CargarColaboradoresUAI()
'    Dim objCOMNAuditoria As COMNAuditoria.NCOMRegistros
'    Set objCOMNAuditoria = New COMNAuditoria.NCOMRegistros
'
'    Dim lrDatos As ADODB.Recordset
'    Set lrDatos = New ADODB.Recordset
'    Set lrDatos = objCOMNAuditoria.ObtenerColaboradoresUAI
'
'    Call CargarComboBox(lrDatos, cboUsuarios)
'End Sub
'
'Public Sub CargarComboBox(ByVal lrDatos As ADODB.Recordset, ByVal cboControl As ComboBox)
'    Dim nContador As Integer
'    Do Until lrDatos.EOF
'     cboUsuarios.AddItem "" & lrDatos!cPersNombre
'     cboUsuarios.ItemData(cboUsuarios.NewIndex) = "" & nContador
'     'cboUsuarios.ItemData(cboUsuarios.NewIndex) = "" & lrDatos!cPersCod
'     'cArregloCodPersona(nContador) = lrDatos!cUser
'     ReDim Preserve sArregloCodPersona(nContador + 1)
'     sArregloCodPersona(nContador) = lrDatos!cPersCod
'     lrDatos.MoveNext
'     nContador = nContador + 1
'    Loop
'    Set lrDatos = Nothing
'
'    cboUsuarios.ListIndex = 0
'End Sub
