VERSION 5.00
Begin VB.UserControl ProgressBarra 
   BackStyle       =   0  'Transparent
   ClientHeight    =   1965
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6045
   ScaleHeight     =   1965
   ScaleWidth      =   6045
   Begin VB.Frame frmProgress 
      Height          =   1995
      Left            =   0
      TabIndex        =   0
      Top             =   -60
      Width           =   6030
      Begin Sicmact.ctlProgress Barra 
         Height          =   405
         Left            =   270
         TabIndex        =   3
         Top             =   1440
         Width           =   5445
         _ExtentX        =   9604
         _ExtentY        =   714
         Appearance      =   1
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         FillColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionStyle    =   2
         Caption         =   ""
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblSubtitulo 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   270
         TabIndex        =   2
         Top             =   1170
         Width           =   5535
      End
      Begin VB.Image imgStart 
         Height          =   480
         Left            =   240
         Picture         =   "ProgressBarra.ctx":0000
         Top             =   630
         Width           =   480
      End
      Begin VB.Image imgEnd 
         Height          =   480
         Left            =   5250
         Picture         =   "ProgressBarra.ctx":030A
         Top             =   630
         Width           =   480
      End
      Begin VB.Image imgMiddle 
         Height          =   480
         Left            =   2610
         Picture         =   "ProgressBarra.ctx":0614
         Top             =   630
         Width           =   480
      End
      Begin VB.Image imgGo 
         Height          =   240
         Index           =   1
         Left            =   2850
         Picture         =   "ProgressBarra.ctx":2DB6
         Top             =   750
         Width           =   240
      End
      Begin VB.Image imgGo 
         Height          =   240
         Index           =   0
         Left            =   450
         Picture         =   "ProgressBarra.ctx":315A
         Top             =   750
         Width           =   240
      End
      Begin VB.Label lblCabecera 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cargando..."
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
         Left            =   255
         TabIndex        =   1
         Top             =   360
         Width           =   5535
      End
   End
End
Attribute VB_Name = "ProgressBarra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Public Enum ePCaptionStyle
    ePCap_None = 0
    ePCap_CaptionOnly
    ePCap_PercentOnly
    ePCap_CaptionPercent
End Enum

Private Enum ePBorderStyle
    ePBor_None = 0
    ePBor_FixedSingle
End Enum

Private Enum ePAppearance
    ePApp_Flat = 0
    ePApp_3D
End Enum

Dim oFormPadre As Object
'variables locales para almacenar los valores de las propiedades
Private mvarMax As Variant 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarColor1 As ColorConstants 'copia local
Private mvarColor2 As ColorConstants 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarCaptionSyle As ePCaptionStyle 'copia local

Const FLAGS = 1
Const HWND_TOPMOST = -1
Dim Aindex As Integer
Dim LastPos As Long
Dim lLastTime As Double
Dim tLastTime
Dim lnMaxValue As Variant
Dim lsSubCaptionBarra As String
Dim lsSubTitulo As String

Public Property Let SubCaptionBarra(ByVal vNewValue As String)
    lsSubCaptionBarra = vNewValue
End Property
Public Property Get SubCaptionBarra() As String
    SubCaptionBarra = lsSubCaptionBarra
End Property

Public Property Let Titulo(ByVal vNewValue As String)
    lblCabecera = vNewValue
End Property
Public Property Get Titulo() As String
    Titulo = lblCabecera
End Property

Public Property Let SubTitulo(ByVal vNewValue As String)
    lsSubTitulo = vNewValue
End Property
Public Property Get SubTitulo() As String
    SubTitulo = lsSubTitulo
End Property


Public Property Let CaptionSyle(ByVal vData As ePCaptionStyle)
    mvarCaptionSyle = vData
End Property
Public Property Get CaptionSyle() As ePCaptionStyle
    CaptionSyle = mvarCaptionSyle
End Property

Public Property Let Color2(ByVal vData As ColorConstants)
    mvarColor2 = vData
End Property
Public Property Get Color2() As ColorConstants
    Color2 = mvarColor2
End Property

Public Property Let Color1(ByVal vData As ColorConstants)
    mvarColor1 = vData
End Property
Public Property Get Color1() As ColorConstants
    Color1 = mvarColor1
End Property
Public Property Let Max(ByVal vData As Variant)
    mvarMax = vData
End Property
Public Property Set Max(ByVal vData As Variant)
    Set mvarMax = vData
End Property
Public Property Get Max() As Variant
    If IsObject(mvarMax) Then
        Set Max = mvarMax
    Else
        Max = mvarMax
    End If
End Property

Public Sub ShowForm(oFormPadre As Object)
oFormPadre.Enabled = False
UserControl.Height = frmProgress.Height
UserControl.Width = frmProgress.Width + 15
If lblCabecera <> "" Then
   lblCabecera = "Ejecutando Proceso...Por Favor Espere"
End If
lblSubtitulo = lsSubTitulo
tLastTime = Time
Aindex = 0
LastPos = 720

Barra.Color1 = mvarColor1
Barra.Color2 = mvarColor2
Barra.CaptionStyle = mvarCaptionSyle

End Sub

Public Sub CloseForm(oFormPadre As Object)  'ByVal OwnerForm As Variant
oFormPadre.Enabled = True
End Sub

Public Function Progress(value)
Barra.Max = mvarMax
lnMaxValue = Barra.Max
ProgressBarra value
End Function

Private Sub ProgressBarra(value)
Dim Perc
Dim bb As Integer
Dim lTime As Double
Dim lTimeDiff As Double
Dim lTimeLeft As Double
Dim lTotalTime As Double
DoEvents
'Display the header , if any was returned

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
        If lsSubTitulo = "" Then
            lblSubtitulo = "Tiempo Transcurrido : " & Format((lTimeLeft), "hh:mm:ss")
        End If
        lblSubtitulo.ForeColor = mvarColor2
End If
DoEvents

lblCabecera.ForeColor = mvarColor2
Barra.Caption = lsSubCaptionBarra
Barra.ForeColor = mvarColor1
Barra.value = value
ProgressImagen
DoEvents
End Sub

Public Function GetTop(oFormPadre As Object) As Long
GetTop = oFormPadre.Height / 2 - UserControl.Height / 2
End Function

Public Function GetLeft(oFormPadre As Object) As Long
GetLeft = oFormPadre.Width / 2 - UserControl.Width / 2
End Function

Private Sub ProgressImagen()

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

