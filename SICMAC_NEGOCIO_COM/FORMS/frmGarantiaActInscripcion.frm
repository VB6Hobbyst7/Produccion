VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmGarantiaActInscripcion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Datos del Certificado de Gravamen"
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4020
   Icon            =   "frmGarantiaActInscripcion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   4020
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtNroPartidaRegistral 
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   1680
   End
   Begin VB.Frame fraDescripcion 
      Caption         =   "Datos de Actualización"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1335
      Left            =   75
      TabIndex        =   6
      Top             =   525
      Width           =   3885
      Begin VB.TextBox txtNroAsiento 
         Height          =   285
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   3
         Top             =   960
         Width           =   1560
      End
      Begin VB.TextBox txtEstado 
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   1
         TabStop         =   0   'False
         Text            =   "Inscrito"
         Top             =   240
         Width           =   1560
      End
      Begin VB.TextBox txtNroSotanos 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1680
         MaxLength       =   2
         TabIndex        =   8
         Tag             =   "txtPrincipal"
         Text            =   "0"
         Top             =   1440
         Width           =   650
      End
      Begin VB.TextBox txtAnioConstruccion 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   4080
         MaxLength       =   4
         TabIndex        =   7
         Tag             =   "txtPrincipal"
         Text            =   "2014"
         Top             =   1440
         Width           =   650
      End
      Begin MSMask.MaskEdBox txtInscripcionFecha 
         Height          =   285
         Left            =   1920
         TabIndex        =   2
         Top             =   600
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "N° de Asiento :"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   1005
         Width           =   1065
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Inscripción :"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   645
         Width           =   1575
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Nuevo Estado :"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   280
         Width           =   1110
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "N° de Sotanos :"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   1470
         Width           =   1125
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Año de Construcción :"
         Height          =   195
         Left            =   2480
         TabIndex        =   9
         Top             =   1470
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdActualizar 
      Caption         =   "&Actualizar"
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   1920
      Width           =   1000
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   2970
      TabIndex        =   5
      Top             =   1920
      Width           =   1000
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "N° de Partida Registral :"
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmGarantiaActInscripcion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********************************************************************************
'** Nombre : frmGarantiaActInscripcion
'** Descripción : Para actualización de Trámite Legal segun TI-ERS063-2014
'** Creación : EJVG, 20151021 05:00:00 PM
'*********************************************************************************
Option Explicit
Dim fsNumGarant As String
Dim fvTramiteLegal As tTramiteLegal

Dim fbOk As Boolean

Public Function Inicio(ByVal psNumGarant As String, ByRef pvTramiteLegal As tTramiteLegal) As Boolean
    fsNumGarant = psNumGarant
    fvTramiteLegal = pvTramiteLegal
    Show 1
    pvTramiteLegal = fvTramiteLegal
    Inicio = fbOk
End Function
Private Sub cmdActualizar_Click()
    Dim lsFecha As String
    
    lsFecha = ValidaFecha(txtInscripcionFecha.Text)
    If Len(lsFecha) > 0 Then
        MsgBox lsFecha, vbInformation, "Aviso"
        EnfocaControl txtInscripcionFecha
        Exit Sub
    End If
    If CDate(txtInscripcionFecha.Text) > gdFecSis Then
        MsgBox "La fecha de Inscripción no debe ser mayor a la fecha del Sistema", vbInformation, "Aviso"
        EnfocaControl txtInscripcionFecha
        Exit Sub
    End If
    
    fvTramiteLegal.nEstado = Inscrita
    fvTramiteLegal.dInscripcion = CDate(txtInscripcionFecha.Text)
    fvTramiteLegal.sNroAsiento = Trim(txtNroAsiento.Text)
    
    fbOk = True
    Unload Me
End Sub
Private Sub cmdCancelar_Click()
    fbOk = False
    Unload Me
End Sub
Private Sub Form_Load()
    fbOk = False
    
    txtNroPartidaRegistral.Text = fvTramiteLegal.sNroPartidaRegistral
End Sub
Private Sub txtInscripcionFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EnfocaControl txtNroAsiento
    End If
End Sub
Private Sub txtNroAsiento_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii)
    If KeyAscii = 13 Then
        EnfocaControl cmdActualizar
    End If
End Sub
