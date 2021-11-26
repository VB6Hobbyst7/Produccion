VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRangoFechas 
   Caption         =   "Reporte de Control de Operaciones"
   ClientHeight    =   1575
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5055
   Icon            =   "frmRangoFechas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   5055
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   3960
      TabIndex        =   4
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   1080
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Rango de Fechas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      Begin MSMask.MaskEdBox txtFecFin 
         Height          =   300
         Left            =   3360
         TabIndex        =   2
         Top             =   360
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFecIni 
         Height          =   300
         Left            =   960
         TabIndex        =   1
         Top             =   360
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Desde:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   420
         Width           =   510
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Hasta:"
         Height          =   195
         Left            =   2520
         TabIndex        =   5
         Top             =   420
         Width           =   465
      End
   End
End
Attribute VB_Name = "frmRangoFechas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()
    Dim ldFechaIni As Date
    Dim ldFechaFin As Date
    Dim lsArchivo As String

    ldFechaIni = txtFecIni.Text
    ldFechaFin = txtFecFin.Text
    
    If ldFechaIni > ldFechaFin Then
        MsgBox "Fecha final debe ser mayor", vbOKOnly, "Error"
        Exit Sub
    End If
        
    Call ListadoReporteControlOpe(ldFechaIni, ldFechaFin)
    
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    txtFecIni.Text = Format(gdFecSis, "dd/mm/yyyy")
    txtFecFin.Text = Format(gdFecSis, "dd/mm/yyyy")
End Sub

Private Sub txtFecFin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdAceptar.SetFocus
    End If
End Sub

Private Sub txtFecIni_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtFecFin.SetFocus
    End If
End Sub
