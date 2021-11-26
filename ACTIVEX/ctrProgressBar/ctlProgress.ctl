VERSION 5.00
Begin VB.UserControl ctlProgress 
   ClientHeight    =   300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   300
   ScaleWidth      =   4800
   Begin VB.PictureBox picProgress 
      AutoRedraw      =   -1  'True
      ClipControls    =   0   'False
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   15
      ScaleHeight     =   195
      ScaleWidth      =   4695
      TabIndex        =   0
      Top             =   15
      Width           =   4755
   End
End
Attribute VB_Name = "ctlProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum eCaptionStyle
    eCap_None = 0
    eCap_CaptionOnly
    eCap_PercentOnly
    eCap_CaptionPercent
End Enum

Public Enum eBorderStyle
    eBor_None = 0
    eBor_FixedSingle
End Enum

Public Enum eAppearance
    eApp_Flat = 0
    eApp_3D
End Enum

'Default Property Values:
Const m_def_Color1 = &HFFFFFF
Const m_def_Color2 = &HFF0000

Dim lMaxValue As Long
Dim lMinValue As Long
Dim lValue As Long
Dim sCaption As String
Dim nCaptionStyle As Integer
Dim oFillColor As OLE_COLOR
Dim StartF%
'Property Variables:
Dim m_Color1 As OLE_COLOR
Dim m_Color2 As OLE_COLOR


Public Property Let Appearance(nValue As eAppearance)
    picProgress.Appearance = nValue
    PropertyChanged
End Property
Public Property Get Appearance() As eAppearance
    Appearance = picProgress.Appearance
End Property
Public Property Let Caption(nValue As String)
    sCaption = Trim(nValue)
    PropertyChanged
End Property
Public Property Get Caption() As String
    Caption = sCaption
End Property
Public Property Let Max(nValue As Long)
    lMaxValue = nValue
    PropertyChanged
End Property
Public Property Get Max() As Long
    Max = lMaxValue
End Property
Public Property Let Min(nValue As Long)
    lMinValue = nValue
    PropertyChanged
End Property
Public Property Get Min() As Long
    Min = lMinValue
End Property
Public Property Let Enabled(nValue As Boolean)
    picProgress.Enabled = nValue
    PropertyChanged
End Property
Public Property Get Enabled() As Boolean
    Enabled = picProgress.Enabled
End Property
Public Property Let BorderStyle(nValue As eBorderStyle)
    picProgress.BorderStyle = nValue
    PropertyChanged
End Property
Public Property Get BorderStyle() As eBorderStyle
    BorderStyle = picProgress.BorderStyle
End Property
Public Property Let CaptionStyle(nValue As eCaptionStyle)
    nCaptionStyle = nValue
    PropertyChanged
End Property
Public Property Get CaptionStyle() As eCaptionStyle
    CaptionStyle = nCaptionStyle
End Property
Public Property Get CaptionFont() As Font
    Set CaptionFont = UserControl.Font
End Property
Public Property Set CaptionFont(ByVal NewFont As Font)
    Set UserControl.Font = NewFont
    SyncLabelFonts
    PropertyChanged
End Property
Private Sub SyncLabelFonts()
Dim objCtl As Object
    For Each objCtl In Controls
        Set objCtl.Font = UserControl.Font
    Next
End Sub
Public Property Let FillColor(nValue As OLE_COLOR)
    oFillColor = nValue
    PropertyChanged
End Property
Public Property Get FillColor() As OLE_COLOR
    FillColor = oFillColor
End Property
Public Property Let ForeColor(nValue As OLE_COLOR)
    picProgress.ForeColor = nValue
    PropertyChanged
End Property
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = picProgress.ForeColor
End Property
Public Property Let BackColor(nValue As OLE_COLOR)
    picProgress.BackColor = nValue
    PropertyChanged
End Property
Public Property Get BackColor() As OLE_COLOR
    BackColor = picProgress.BackColor
End Property

Public Property Let value(nValue As Long)
    lValue = nValue
    Call ChangeValue(nValue)
End Property
Public Property Get value() As Long
    value = lValue
End Property

Public Sub Refresh()
    picProgress.Refresh
End Sub

Private Sub UserControl_InitProperties()
    Max = 100
    Min = 0
    BackColor = UserControl.BackColor
    FillColor = vbBlue
    CaptionStyle = eCap_PercentOnly
    SyncLabelFonts
    m_Color1 = m_def_Color1
    m_Color2 = m_def_Color2
End Sub

Private Sub UserControl_Resize()
    picProgress.Width = UserControl.Width - 30
    picProgress.Height = UserControl.Height - 30
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    picProgress.Appearance = PropBag.ReadProperty("Appearance", picProgress.Appearance)
    picProgress.ForeColor = PropBag.ReadProperty("ForeColor", picProgress.ForeColor)
    picProgress.BackColor = PropBag.ReadProperty("BackColor", picProgress.BackColor)
    oFillColor = PropBag.ReadProperty("FillColor", oFillColor)
    BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    CaptionStyle = PropBag.ReadProperty("CaptionStyle", 3)
    Enabled = PropBag.ReadProperty("Enabled", True)
    Caption = PropBag.ReadProperty("Caption", "")
    Max = PropBag.ReadProperty("Max", 100)
    Min = PropBag.ReadProperty("Min", 0)
    Set CaptionFont = PropBag.ReadProperty("CaptionFont")
    m_Color1 = PropBag.ReadProperty("Color1", m_def_Color1)
    m_Color2 = PropBag.ReadProperty("Color2", m_def_Color2)
End Sub

Private Sub ChangeValue(nValue As Long)
Dim NewCaption As String

    If nValue > lMaxValue Then
        nValue = lMaxValue
    ElseIf nValue < lMinValue Then
        nValue = lMinValue
    End If
    
    picProgress.Cls
    If CaptionStyle <> eCap_None Then
        If CaptionStyle <> eCap_CaptionOnly Then
            If Caption = "" Or CaptionStyle = eCap_PercentOnly Then
                NewCaption = Format(Str((nValue - Min) / (Max - Min)) * 100, "0.00") + "%"
            Else
                NewCaption = Caption & " " & Format(Str((nValue - Min) / (Max - Min)) * 100, "0") + "%"
            End If
        Else
            NewCaption = Caption
        End If
    End If
    
'    picProgress.ScaleWidth = Max - Min
'    picProgress.DrawMode = 10
'
    'picProgress.Line (0, 0)-((nValue - Min), picProgress.Width), FillColor, BF
    StartF = 0
    Dim I As Integer
    For I = 1 To ((nValue - Min) / (Max - Min) * 100)
        picProgress.Line ((picProgress.Width / 100) * StartF, 0)-((picProgress.Width / 100) * (StartF + 1), picProgress.Height), Blend(Color1, Color2, I), BF    '   RGB(255 - (2.5 * StartF), 2.5 * StartF, 255)
        If StartF = 100 Then StartF = 0
        StartF = StartF + 1
    Next

    picProgress.CurrentX = (picProgress.ScaleWidth / 2 - picProgress.TextWidth(NewCaption) / 2)
    picProgress.CurrentY = (picProgress.ScaleHeight - picProgress.TextHeight(NewCaption)) / 2
    picProgress.Print NewCaption
    DoEvents

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Appearance", picProgress.Appearance)
    Call PropBag.WriteProperty("ForeColor", picProgress.ForeColor)
    Call PropBag.WriteProperty("BackColor", picProgress.BackColor)
    Call PropBag.WriteProperty("FillColor", oFillColor)
    Call PropBag.WriteProperty("Font", Font, Ambient.Font)
    Call PropBag.WriteProperty("BorderStyle", BorderStyle, 1)
    Call PropBag.WriteProperty("CaptionStyle", CaptionStyle, 3)
    Call PropBag.WriteProperty("Enabled", Enabled, True)
    Call PropBag.WriteProperty("Caption", Caption)
    Call PropBag.WriteProperty("Min", Min, 0)
    Call PropBag.WriteProperty("CaptionFont", CaptionFont)
    Call PropBag.WriteProperty("Color1", m_Color1, m_def_Color1)
    Call PropBag.WriteProperty("Color2", m_Color2, m_def_Color2)
End Sub
Public Function RGBRed(RGBCol As Long) As Integer
    RGBRed = RGBCol And &HFF
End Function

Public Function RGBGreen(RGBCol As Long) As Integer
    RGBGreen = ((RGBCol And &H100FF00) / &H100)
End Function

Public Function RGBBlue(RGBCol As Long) As Integer
    RGBBlue = (RGBCol And &HFF0000) / &H10000
End Function
Public Function Blend(Color1 As OLE_COLOR, Color2 As OLE_COLOR, Number As Integer) As OLE_COLOR
Dim r As Long, g As Long, b As Long
r = ((RGBRed(Color1) * (100 - Number)) + (RGBRed(Color2) * (Number))) / 100
g = ((RGBGreen(Color1) * (100 - Number)) + (RGBGreen(Color2) * (Number))) / 100
b = ((RGBBlue(Color1) * (100 - Number)) + (RGBBlue(Color2) * (Number))) / 100
Blend = RGB(r, g, b)
End Function
'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=10,0,0,0
Public Property Get Color1() As OLE_COLOR
    Color1 = m_Color1
End Property

Public Property Let Color1(ByVal New_Color1 As OLE_COLOR)
    m_Color1 = New_Color1
    PropertyChanged "Color1"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=10,0,0,0
Public Property Get Color2() As OLE_COLOR
    Color2 = m_Color2
End Property

Public Property Let Color2(ByVal New_Color2 As OLE_COLOR)
    m_Color2 = New_Color2
    PropertyChanged "Color2"
End Property

