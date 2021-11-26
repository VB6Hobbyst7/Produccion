VERSION 5.00
Begin VB.Form frmProgress 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " "
   ClientHeight    =   1800
   ClientLeft      =   2535
   ClientTop       =   3390
   ClientWidth     =   5940
   ControlBox      =   0   'False
   Icon            =   "frmProgress.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   Begin ProgressBar.ctlProgress Barra 
      Height          =   315
      Left            =   255
      TabIndex        =   2
      Top             =   1125
      Width           =   5400
      _ExtentX        =   9525
      _ExtentY        =   556
      Appearance      =   1
      ForeColor       =   -2147483634
      BackColor       =   -2147483637
      FillColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer tmrAnimate 
      Interval        =   100
      Left            =   4680
      Top             =   345
   End
   Begin VB.Label lblCabecera 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Loading ..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   165
      TabIndex        =   1
      Top             =   75
      Width           =   5535
   End
   Begin VB.Image imgGo 
      Height          =   240
      Index           =   0
      Left            =   360
      Picture         =   "frmProgress.frx":030A
      Top             =   465
      Width           =   240
   End
   Begin VB.Image imgGo 
      Height          =   240
      Index           =   1
      Left            =   2760
      Picture         =   "frmProgress.frx":06C0
      Top             =   465
      Width           =   240
   End
   Begin VB.Image imgMiddle 
      Height          =   480
      Left            =   2520
      Picture         =   "frmProgress.frx":0A64
      Top             =   345
      Width           =   480
   End
   Begin VB.Image imgEnd 
      Height          =   480
      Left            =   5160
      Picture         =   "frmProgress.frx":3206
      Top             =   345
      Width           =   480
   End
   Begin VB.Image imgStart 
      Height          =   480
      Left            =   120
      Picture         =   "frmProgress.frx":3510
      Top             =   345
      Width           =   480
   End
   Begin VB.Label lblSubtitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H00C00000&
      Height          =   225
      Left            =   165
      TabIndex        =   0
      Top             =   870
      Width           =   5535
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Author : Renier Barnard (renier_barnard@santam.co.za)
'
' Date    : May 2000
'
' Description :
' This code will demonstrate how to make a simple but nice
' looking progress bar. It could be more simple (Using the line command)
' but this looks better... way better. The Status bar also changes colour as it progresses.
' It will also calculate the time remaining and display it
' In addition to this , it animates some icons to keep the form "busy looking"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Private Declare Function SetWindowPos Lib "USER32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Const FLAGS = 1
Const HWND_TOPMOST = -1
Dim Aindex As Integer
Dim LastPos As Long
Dim lLastTime As Double
Dim tLastTime
Dim lnMaxValue As Variant

Public Sub Progress(value, Optional HeaderX As String = "", _
                        Optional psSubTitulo As String = "", _
                        Optional psSubCaptionBarra As String = "", _
                        Optional ColorSubTitulo As ColorConstants, _
                        Optional ForeColorBarra As ColorConstants = vbWhite)
'' This is the actual progress bar function.
DoEvents
Dim Perc
Dim bb As Integer
Dim lTime As Double
Dim lTimeDiff As Double
Dim lTimeLeft As Double
Dim lTotalTime As Double
'Me.Show

'Get a color to do it in
If ColorSubTitulo = 0 Then ColorSubTitulo = &H8000&

'Display the header , if any was returned
If HeaderX <> "" Then
    lblCabecera = HeaderX
Else
    lblCabecera = "Ejecutando Proceso...Por Favor Espera"
End If

'Now work out the percentage (0-100) of where we currently are
Perc = (value / lnMaxValue) * 100
If Perc < 0 Then Perc = 0
If Perc > 100 Then Perc = 100
Perc = Int(Perc)

'Do the time remaining calculation
If (Perc Mod 10) = 0 Or Perc = 0 Then ' Every 10 percent
        lTimeDiff = lTime - lLastTime
        lTime = Time - tLastTime
        If Perc = 0 Or Perc < 0 Then
            lTotalTime = ((100 / 1) * 2) * lTime
            lTimeLeft = (((100 / 1) * 2) * lTime) - (((100 / 100) * 2) * lTime)
        Else
            lTotalTime = ((100 / Perc) * 2) * lTime
            lTimeLeft = (((100 / Perc) * 2) * lTime) - (((100 / 100) * 2) * lTime)
        End If
        If psSubTitulo = "" Then
            lblSubtitulo = "Tiempo Transcurrido : " & Format((lTimeLeft), "hh:mm:ss")
        Else
            lblSubtitulo = psSubTitulo
        End If
        lblSubtitulo.ForeColor = ColorSubTitulo
End If
DoEvents

lblCabecera.ForeColor = ColorSubTitulo
Barra.Caption = psSubCaptionBarra
Barra.ForeColor = ForeColorBarra
Barra.value = value
DoEvents

End Sub


Public Sub Inicio(oForm As Object, Optional isModal As Boolean = False)
    'Me.Show
    'CTI6-20210503-ERS032-2019 -(Optimizar Sugerencia)
    If isModal Then
        Me.Show 1
    Else
        Me.Show
    End If
    'CTI6-20210503-ERS032-2019 -(Optimizar Sugerencia)
End Sub

Private Sub Form_Load()

Dim lResult As Long

'permite mostrar la ventana simpre como formulario Top Encima
lResult = SetWindowPos(Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)


tLastTime = Time

'Const FLAGS = 1
'Const HWND_TOPMOST = -1
Aindex = 0
LastPos = 720

'Me.Width = 5910
'Me.Height = 1545

'Sets form on always on top.
Dim Success As Integer
'Success% = SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
                                                ' Change the "0's" above to position the window.

Me.Top = Screen.Height / 2 + (Me.Height / 1.3)
Me.Left = Screen.Width / 2 - Me.Width / 2
DoEvents

End Sub


Private Sub Form_Unload(Cancel As Integer)
DoEvents
End Sub

Private Sub tmrAnimate_Timer()
'This funtion will animate a couple of icons , just to show that something is busy hapening

DoEvents
LastPos = LastPos + 1

If LastPos > 2680 And LastPos < 3250 Then
    LastPos = 3160
    Aindex = 1
Else
    If LastPos > 5360 Then
        LastPos = 720
        Aindex = 0
    Else
        
    End If
End If

If Aindex = 1 Then
    imgGo(1).Visible = True
    imgGo(0).Visible = False
Else
    imgGo(1).Visible = False
    imgGo(0).Visible = True
End If

LastPos = LastPos + 200
imgGo(Aindex).Left = LastPos
DoEvents

End Sub
Public Property Get Max() As Variant
Max = lnMaxValue
End Property
Public Property Let Max(ByVal vNewValue As Variant)
lnMaxValue = vNewValue
Me.Barra.Max = lnMaxValue
End Property
Public Property Let Color1(ByVal vNewValue As ColorConstants)
Barra.Color1 = vNewValue
End Property
Public Property Let Color2(ByVal vNewValue As ColorConstants)
Barra.Color2 = vNewValue
End Property
Public Property Let CaptionStyleBarra(ByVal vNewValue As eCaptionStyle)
Barra.CaptionStyle = vNewValue
End Property
