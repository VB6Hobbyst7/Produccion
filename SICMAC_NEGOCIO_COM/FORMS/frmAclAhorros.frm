VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAclAhorros 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Generar Archivo ACL"
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   4170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ProgressBar PrgBarra 
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   2640
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Frame frmReportes 
      Caption         =   "Generar Archivo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3975
      Begin VB.CheckBox chkPlazoFijo 
         Caption         =   "Depósitos a Plazo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   2055
      End
      Begin VB.CheckBox chkMovimientos 
         Caption         =   "Depósitos a Plazo Movimientos"
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
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   3255
      End
      Begin VB.CheckBox chkInactivas 
         Caption         =   "Depósitos Ahorros Inactivas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   2895
      End
      Begin VB.CheckBox chkActivas 
         Caption         =   "Depósitos Ahorros Activas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   1150
         Width           =   2895
      End
   End
   Begin VB.Label lbletiqueta1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   2280
      Width           =   3375
   End
End
Attribute VB_Name = "frmAclAhorros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Progress As clsProgressBar


Private Sub CmdAceptar_Click()
On Error GoTo ControlError
Dim loAcl As ACL
Set loAcl = New ACL

lbletiqueta1.Visible = True
If chkPlazoFijo.value = 1 Then
    lbletiqueta1.Caption = "Procesando Depósitos a Plazo"
    Call loAcl.GeneraPF(gdFecSis)
End If
If ChkMovimientos.value = 1 Then
    lbletiqueta1.Caption = "Procesando Depósitos a Plazo Movimientos"
    Call loAcl.GenerarPF_Movimiento(gdFecSis)
End If
If chkInactivas.value = 1 Then
    lbletiqueta1.Caption = "Procesando Depósitos de Ahorro Inactivas"
    Call loAcl.Generar_Inactivas(gdFecSis)
End If
If chkActivas.value = 1 Then
    lbletiqueta1.Caption = "Procesando Depósitos de Ahorro Activas"
   loAcl.Generar_Activas (gdFecSis)
End If
lbletiqueta1.Caption = "Proceso Terminado..."
MsgBox "Transferencia a Formato DBF Terminada  " & Chr(10) & "Grabado en " & App.path & "\Spooler", vbInformation, "Aviso"
chkPlazoFijo.value = 0
ChkMovimientos.value = 0
chkInactivas.value = 0
chkActivas = 0
PrgBarra.value = 0
Set loAcl = Nothing
Exit Sub
ControlError: MsgBox "El Archivo esta Siendo Usado ó no hay Conexión con el Servidor, Avisar al Area de Sistemas" & Chr(10) & Err.Description, vbInformation, "Aviso"
lbletiqueta1.Visible = False
Screen.MousePointer = 0
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Command1_Click()
frmCFTarifario.Show 1
End Sub

Private Sub Form_Load()
Set Progress = New clsProgressBar
Me.Icon = LoadPicture(App.path & gsRutaIcono)
End Sub

Private Sub loAcl_CloseProgress()
Progress.CloseForm Me
End Sub

Private Sub loAcl_Progress(pnValor As Long, pnTotal As Long)
Progress.Max = pnTotal
Progress.Progress pnValor, "Generando Reporte", "Procesando ..."
End Sub

Private Sub loAcl_ShowProgress()
Progress.ShowForm Me
End Sub

