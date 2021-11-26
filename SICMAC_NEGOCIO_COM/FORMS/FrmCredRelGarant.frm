VERSION 5.00
Begin VB.Form FrmCredRelGarant 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Relacion de Credito Con Garantia"
   ClientHeight    =   675
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   675
   ScaleWidth      =   5130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5085
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "Aceptar"
         Height          =   315
         Left            =   3840
         TabIndex        =   2
         Top             =   210
         Width           =   1125
      End
      Begin SICMACT.ActXCodCta ActXCodCta1 
         Height          =   315
         Left            =   60
         TabIndex        =   1
         Top             =   180
         Width           =   3675
         _ExtentX        =   6482
         _ExtentY        =   556
         Texto           =   "Cuenta:"
         EnabledCMAC     =   -1  'True
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
   End
End
Attribute VB_Name = "FrmCredRelGarant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Sub InicializarCuenta()
  ActXCodCta1.CMAC = "108"
  ActXCodCta1.Age = gsCodAge
End Sub

Private Sub CmdAceptar_Click()
  If Len(ActXCodCta1.NroCuenta) <> 18 Then
    MsgBox "La cuenta de credito no tiene los digitos completos", vbInformation, "AVISO"
  Else
    If gsProyectoActual = "H" Then
        frmPersGarantiasHC.pgcCtaCod = Me.ActXCodCta1.NroCuenta
    Else
        frmPersGarantias.pgcCtaCod = Me.ActXCodCta1.NroCuenta
    End If
    Unload Me
  End If
End Sub

Private Sub Form_Load()
    InicializarCuenta
End Sub
