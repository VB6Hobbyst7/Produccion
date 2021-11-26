VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRHCierreMes 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5550
   Icon            =   "frmRHCierreMes.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraCierre 
      Caption         =   "Provisiones de Mensuales de Recursos Humanos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2505
      Left            =   30
      TabIndex        =   1
      Top             =   45
      Width           =   5490
      Begin MSMask.MaskEdBox mskFecha 
         Height          =   300
         Left            =   3045
         TabIndex        =   4
         Top             =   1567
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.CommandButton cmdGenerar 
         Caption         =   "&Generar Cierre de Mes"
         Height          =   375
         Left            =   1080
         TabIndex        =   2
         Top             =   2040
         Width           =   3360
      End
      Begin VB.Label lblFechaAsientos 
         Caption         =   "Fecha de Asientos :"
         Height          =   225
         Left            =   1245
         TabIndex        =   5
         Top             =   1605
         Width           =   1620
      End
      Begin VB.Label lblCierre 
         Caption         =   $"frmRHCierreMes.frx":030A
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   1065
         Left            =   165
         TabIndex        =   3
         Top             =   270
         Width           =   5205
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   4365
      TabIndex        =   0
      Top             =   2595
      Width           =   1125
   End
End
Attribute VB_Name = "frmRHCierreMes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents oPla As NRHProcesosCierre
Attribute oPla.VB_VarHelpID = -1
Dim Progress As clsProgressBar

Private Sub Form_Load()
    Set Progress = New clsProgressBar
End Sub

Private Sub mskFecha_GotFocus()
    mskFecha.SelStart = 0
    mskFecha.SelLength = 50
End Sub

Private Sub mskFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdGenerar.SetFocus
    End If
End Sub

Private Sub oPla_CloseProgress()
    Progress.CloseForm Me
End Sub

Private Sub oPla_Progress(pnValor As Long, pnTotal As Long)
    Progress.Max = pnTotal
    Progress.Progress pnValor, "Generando Reporte"
End Sub

Private Sub cmdGenerar_Click()
    Dim oPrevio As Previo.clsPrevio
    Set oPla = New NRHProcesosCierre
    Set oPrevio = New Previo.clsPrevio
    Dim lsCadena As String
    
    If Not IsDate(Me.mskFecha.Text) Then
        MsgBox "Debe ingresar una fecha valida.", vbInformation, "Aviso"
        Me.mskFecha.SetFocus
        Exit Sub
    End If
    
    lsCadena = oPla.CierreMesRRHH(CDate(Me.mskFecha.Text), gsCodAge, gsCodUser)
    
    If lsCadena <> "" Then oPrevio.Show lsCadena, Caption, True, 66
    
    Set oPla = Nothing
    Set oPrevio = Nothing
    Me.cmdGenerar.Enabled = False
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Public Sub Ini(psCaption As String)
    Caption = psCaption
    Me.Show
End Sub

Private Sub oPla_ShowProgress()
    Progress.ShowForm Me
End Sub
