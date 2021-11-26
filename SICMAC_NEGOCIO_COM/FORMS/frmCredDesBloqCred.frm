VERSION 5.00
Begin VB.Form frmCredDesBloqCred 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Desbloqueo de Credito"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3945
   Icon            =   "frmCredDesBloqCred.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   3945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   435
      Left            =   1995
      TabIndex        =   3
      Top             =   1830
      Width           =   1350
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
      Enabled         =   0   'False
      Height          =   435
      Left            =   600
      TabIndex        =   2
      Top             =   1830
      Width           =   1350
   End
   Begin VB.CheckBox ChkBloq 
      Caption         =   "Credito Bloqueado"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   405
      Left            =   540
      TabIndex        =   1
      Top             =   945
      Width           =   3030
   End
   Begin SICMACT.ActXCodCta ActxCodCta 
      Height          =   510
      Left            =   90
      TabIndex        =   0
      Top             =   165
      Width           =   3705
      _ExtentX        =   6535
      _ExtentY        =   900
      Texto           =   "Credito  :"
      EnabledCMAC     =   -1  'True
      EnabledCta      =   -1  'True
      EnabledProd     =   -1  'True
      EnabledAge      =   -1  'True
   End
End
Attribute VB_Name = "frmCredDesBloqCred"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub HabilitaForm(ByVal pbHabilita As Boolean)
        
    ActxCodCta.Enabled = Not pbHabilita
    If Not pbHabilita Then
        ActxCodCta.NroCuenta = ""
        ActxCodCta.CMAC = gsCodCMAC
        ActxCodCta.Age = gsCodAge
        ChkBloq.value = 0
    End If
    CmdAceptar.Enabled = pbHabilita
    CmdCancelar.Enabled = pbHabilita
    ChkBloq.Enabled = pbHabilita
    
End Sub

Private Sub ActxCodCta_KeyPress(KeyAscii As Integer)
Dim nCred As COMNCredito.NCOMCredito
Dim sMsgBox As String
Dim pnBloq As Integer
    If KeyAscii = 13 Then
        Set nCred = New COMNCredito.NCOMCredito
        sMsgBox = nCred.CargaDatosBloqueoCred(ActxCodCta.NroCuenta, pnBloq)
        Set nCred = Nothing
        
        If sMsgBox <> "" Then
            MsgBox sMsgBox, vbInformation, "Desbloqueo de Credito"
            HabilitaForm False
        Else
            Me.ChkBloq.value = pnBloq
            
            HabilitaForm True
        End If
    End If

End Sub

Private Sub cmdAceptar_Click()
Dim nCred As COMNCredito.NCOMCredito

    If MsgBox("Desea Grabar la Operacion ?", vbInformation + vbYesNo, "Grabacion de Datos") = vbYes Then
    
        Set nCred = New COMNCredito.NCOMCredito
        Call nCred.ActualizaBloqueoCredito(ActxCodCta.NroCuenta, ChkBloq.value)
        Set nCred = Nothing
    
    End If
    
    HabilitaForm False
    
End Sub

Private Sub CmdCancelar_Click()
    HabilitaForm False
End Sub

Private Sub Form_Load()

    CentraForm Me
    Me.ActxCodCta.CMAC = gsCodCMAC
    Me.ActxCodCta.Age = gsCodAge
End Sub

