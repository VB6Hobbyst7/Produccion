VERSION 5.00
Begin VB.UserControl LabelX 
   ClientHeight    =   420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   720
   LockControls    =   -1  'True
   ScaleHeight     =   420
   ScaleWidth      =   720
   ToolboxBitmap   =   "LabelX.ctx":0000
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   180
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   540
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00808080&
      X1              =   45
      X2              =   45
      Y1              =   45
      Y2              =   345
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00E0E0E0&
      X1              =   630
      X2              =   630
      Y1              =   45
      Y2              =   330
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00E0E0E0&
      X1              =   30
      X2              =   615
      Y1              =   330
      Y2              =   330
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   60
      X2              =   630
      Y1              =   45
      Y2              =   45
   End
End
Attribute VB_Name = "LabelX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Enum eColors
    eAzul = vbBlue
    eVerde = &H8000&
    eNegro = vbBlack
End Enum
Private sCaption As String
Private nResalte As eColors
Private bFondoBlanco As Boolean

Public Property Get FondoBlanco() As Boolean
    FondoBlanco = bFondoBlanco
End Property

Public Property Let FondoBlanco(ByVal vNewValue As Boolean)
    bFondoBlanco = vNewValue
    If bFondoBlanco Then
        Label1.BackColor = vbWhite
    Else
        Label1.BackColor = &H8000000F
    End If
End Property

Public Property Get Resalte() As eColors
    Resalte = nResalte
End Property

Public Property Let Resalte(ByVal vNewValue As eColors)
    nResalte = vNewValue
    Label1.ForeColor = nResalte
End Property

Public Property Get Caption() As Variant
    Caption = sCaption
End Property

Public Property Let Caption(ByVal vNewValue As Variant)
    sCaption = vNewValue
    Label1.Caption = sCaption
End Property

Private Sub UserControl_Resize()
    If UserControl.Width < 645 Then
        UserControl.Width = 645
    End If
    If UserControl.Height < 375 Then
        UserControl.Height = 375
    End If
    Line1.X1 = 60
    Line1.X2 = UserControl.Width - 100
    Line1.Y1 = 45
    Line1.Y2 = 45
    
    Line2.X1 = 30
    Line2.X2 = UserControl.Width - 100
    Line2.Y1 = UserControl.Height - 90
    Line2.Y2 = UserControl.Height - 90
    
    Line3.X1 = UserControl.Width - 90
    Line3.X2 = UserControl.Width - 90
    Line3.Y1 = 45
    Line3.Y2 = UserControl.Height - 90
    
    Line5.X1 = 45
    Line5.X2 = 45
    Line5.Y1 = 45
    Line5.Y2 = UserControl.Height - 100
    
    Label1.Top = 90
    Label1.Left = 90
    Label1.Height = UserControl.Height - 250
    Label1.Width = UserControl.Width - 240
    Label1.FontSize = CInt(Format((8 * Label1.Height) / 180, "#0"))
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Resalte = PropBag.ReadProperty("Resalte", False)
    FondoBlanco = PropBag.ReadProperty("FondoBlanco", False)
    Caption = PropBag.ReadProperty("Caption", "")
    Bold = PropBag.ReadProperty("Bold", "")
    Alignment = PropBag.ReadProperty("Alignment", "")
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "FondoBlanco", FondoBlanco, ""
    PropBag.WriteProperty "Resalte", Resalte, ""
    PropBag.WriteProperty "Caption", Caption, ""
    PropBag.WriteProperty "Bold", Bold, ""
    PropBag.WriteProperty "Alignment", Alignment, ""
End Sub

Public Property Get Alignment() As Variant
    Alignment = Label1.Alignment
End Property

Public Property Let Alignment(ByVal vNewValue As Variant)
    Label1.Alignment = vNewValue
End Property

Public Property Get Bold() As Boolean
    Bold = Label1.FontBold
End Property

Public Property Let Bold(ByVal vNewValue As Boolean)
    Label1.FontBold = vNewValue
End Property