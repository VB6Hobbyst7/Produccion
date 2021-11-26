VERSION 5.00
Begin VB.Form frmHojaRutaAnalistaVistoInc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HOHA DE RUTA - V°B° del jefe de Agencia Pendiente"
   ClientHeight    =   2850
   ClientLeft      =   11445
   ClientTop       =   6690
   ClientWidth     =   6105
   Icon            =   "frmHojaRutaAnalistaVistoInc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   6105
   Begin VB.CommandButton cmbCerrar 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      Top             =   2280
      Width           =   1575
   End
   Begin VB.CommandButton cmbVisto 
      Caption         =   "Continuar"
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Si ya cuenta con el Visto Bueno, presione en ""Continuar"", para registrar los resultados de la visita."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   1320
      Width           =   5415
   End
   Begin VB.Label Label1 
      Caption         =   "Usted  tiene visitas de más de un dia sin completar. Para continuar se necesita un Visto Bueno por parte del Jefe de Agencia. "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   5415
   End
End
Attribute VB_Name = "frmHojaRutaAnalistaVistoInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bPasar As Boolean
Dim oDhoja As New DCOMhojaRuta
Public Function inicio()
    Me.Show 1
    inicio = bPasar
End Function
Private Sub cmbCerrar_Click()
    bPasar = False
    Unload Me
End Sub

Private Sub cmbVisto_Click()
    If Not oDhoja.tieneVistoPendiente(gsCodUser) Then
        oDhoja.recibirVisto gsCodUser
        bPasar = True
        Unload Me
    Else
        MsgBox "Aun no se aprueba el Visto Bueno"
    End If
End Sub

Private Sub Form_Load()
    bPasar = False
End Sub
