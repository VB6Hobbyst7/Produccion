VERSION 5.00
Begin VB.Form frmMyMsgBox 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4830
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   4830
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Command2"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   1560
      Width           =   1815
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Command1"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   0
      TabIndex        =   3
      Top             =   1440
      Width           =   4815
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmMyMsgBox.frx":0000
      Top             =   480
      Width           =   480
   End
   Begin VB.Label lblMessage 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label1"
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
      Left            =   840
      TabIndex        =   0
      Top             =   480
      Width           =   3615
   End
End
Attribute VB_Name = "frmMyMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'CREADO POR JHCU ENCUESTA PINPADS 06-01-2019
'MODIFICACIONES POR JHCU MEJORAS
Public valor As Integer 'Contiene el botón pulsado:
Dim lblTextoProceo As String
Dim Pinpads As New clases.clsPinpad
Dim nIntento As Integer
Dim nCorrecto As Boolean

'True -> cmdCancel (Botón 2)
'False -> cmdOk (Botón 1)

'El usuario pulsa el botón cmdOk (Botón 1)-Se establece la variable Cancel a False
Private Sub cmdOK_Click()
    cmdOk.Enabled = False
    cmdOk.Caption = "Esperando Res..."
    cmdCancel.Enabled = False
    nIntento = nIntento + 1
    nCorrecto = False
    LeertecladoPinpad
    If nCorrecto Then
        If valor <> -2 Then
            frmMensajeMostrar.Inicio ("Puntuación Registrada")
        End If
        Unload Me
    End If
    cmdOk.Enabled = True
    cmdOk.Caption = "Pedir Encuesta"
    cmdCancel.Enabled = True
End Sub
Public Function LeertecladoPinpad() As Integer
 On Error GoTo LeertecladoPinpadError
 Dim sRes As String
        Dim sEstadoPin As String
     
       ' sRes = Tarjeta.PedirTecla("Ingrese su evaluación (1-5)", gnPinPadPuerto)
        sRes = Pinpads.PedirTecla("Ingrese su evaluación (1-5)", gnPinPadPuerto)
        '-2 no hay conexión con el pinpad
        Select Case sRes
        Case "1"
            valor = sRes
            nCorrecto = True
        Case "2"
            valor = sRes
            nCorrecto = True
        Case "3"
            valor = sRes
            nCorrecto = True
        Case "4"
            valor = sRes
            nCorrecto = True
        Case "5"
            valor = sRes
            nCorrecto = True
        Case "-2"
            frmMensajeMostrar.Inicio ("Alerta (NO HAY CONEXIÓN CON EL PINPAD):se culminará el proceso de encuesta")
            valor = sRes
            nCorrecto = True
        Case Else
            If nIntento <= 3 Then
             frmMensajeMostrar.Inicio ("Alerta: Opción no válida, presione nuevamente el boton pedir encuesta")
            nCorrecto = False
            Else
            nCorrecto = True
            valor = -6
            End If
        End Select


Exit Function
LeertecladoPinpadError:
    LeertecladoPinpad = -5
End Function
'El usuario pulsa el botón cmdCancel (Botón 2)-Se establece la variable Cancel a True
Private Sub cmdCancel_Click()
    valor = -1
    Unload Me
End Sub



