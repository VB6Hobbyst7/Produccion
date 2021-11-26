VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl uSpinner 
   ClientHeight    =   345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   855
   ScaleHeight     =   345
   ScaleWidth      =   855
   ToolboxBitmap   =   "sPinner.ctx":0000
   Begin MSComCtl2.UpDown UpSpinner 
      Height          =   330
      Left            =   600
      TabIndex        =   0
      Top             =   15
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
      _Version        =   393216
      Max             =   100
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txtSpinner 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   15
      MaxLength       =   3
      TabIndex        =   1
      Text            =   "1"
      Top             =   15
      Width           =   615
   End
End
Attribute VB_Name = "uSpinner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim ctrControl As Control
'Default Property Values:
Const m_Ancho_UpDown = 240
'Event Declarations:
Event Change() 'MappingInfo=UpSpinner,UpSpinner,-1,Change
Attribute Change.VB_Description = "La posición actual ha cambiado"
Event UpClick() 'MappingInfo=UpSpinner,UpSpinner,-1,UpClick
Attribute UpClick.VB_Description = "Se ha hecho clic en el botón arriba del control UpDown"
Event DownClick() 'MappingInfo=UpSpinner,UpSpinner,-1,DownClick
Attribute DownClick.VB_Description = "Se ha hecho clic en el botón abajo del control UpDown"
Event KeyPress(KeyAscii As Integer) 'MappingInfo=txtSpinner,txtSpinner,-1,KeyPress
Attribute KeyPress.VB_Description = "Ocurre cuando el usuario presiona y libera una tecla ANSI."

Private Sub txtSpinner_Change()
If Int(Val(txtSpinner.Text)) <= UpSpinner.Max And Int(Val(txtSpinner.Text)) >= UpSpinner.Min Then
    UpSpinner.Value = Int(Val(txtSpinner.Text))
End If
End Sub

Private Sub txtSpinner_GotFocus()
    txtSpinner.SelStart = 0
    txtSpinner.SelLength = Len(txtSpinner)
End Sub

Private Sub txtSpinner_LostFocus()
    txtSpinner = IIf(txtSpinner = "", UpSpinner.Min, txtSpinner)
    If IsNumeric(txtSpinner) = False Then txtSpinner = UpSpinner.Min
    If Val(txtSpinner.Text) < UpSpinner.Min Then
        txtSpinner.Text = UpSpinner.Min
        UpSpinner.Value = UpSpinner.Min
        Exit Sub
    End If
    If Val(txtSpinner.Text) > UpSpinner.Max Then
        txtSpinner.Text = UpSpinner.Max
        UpSpinner.Value = UpSpinner.Max
        Exit Sub
    End If
    UpSpinner.Value = Int(Val(txtSpinner.Text))
End Sub

Private Sub UpSpinner_Change()
    RaiseEvent Change
End Sub
Private Sub UserControl_EnterFocus()
    txtSpinner.SetFocus
End Sub
'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=txtSpinner,txtSpinner,-1,MaxLength
Public Property Get MaxLength() As Long
Attribute MaxLength.VB_Description = "Devuelve o establece el número máximo de caracteres que se puede escribir en un control."
    MaxLength = txtSpinner.MaxLength
End Property
Public Property Let MaxLength(ByVal New_MaxLength As Long)
    txtSpinner.MaxLength() = New_MaxLength
    PropertyChanged "MaxLength"
End Property

'Cargar valores de propiedad desde el almacén
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UpSpinner.Max = PropBag.ReadProperty("Max", "100")
    UpSpinner.Min = PropBag.ReadProperty("Min", "0")
    UpSpinner.Value = Min
    txtSpinner.MaxLength = PropBag.ReadProperty("MaxLength", 3)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    UpSpinner.Increment = PropBag.ReadProperty("Increment", 1)
    txtSpinner.Text = PropBag.ReadProperty("Min", Min)
'    txtSpinner.ToolTipText = PropBag.ReadProperty("ToolTipText", "")
    txtSpinner.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    Set txtSpinner.Font = PropBag.ReadProperty("Font", Ambient.Font)
    txtSpinner.FontBold = PropBag.ReadProperty("FontBold", 0)
    txtSpinner.FontItalic = PropBag.ReadProperty("FontItalic", 0)
    txtSpinner.FontName = PropBag.ReadProperty("FontName", "MS Sans Serif")
    txtSpinner.FontSize = PropBag.ReadProperty("FontSize", 8)
    txtSpinner.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    txtSpinner.Locked = PropBag.ReadProperty("Locked", False)
End Sub

Private Sub UserControl_Resize()
If ScaleWidth > 100 Then
    txtSpinner.Move 0, 0, (ScaleWidth - UpSpinner.Width) + 30, ScaleHeight
    txtSpinner.Font.Size = Int(ScaleWidth / 100)
    UpSpinner.Move ScaleWidth - UpSpinner.Width, 0, m_Ancho_UpDown, ScaleHeight
End If
End Sub

'Escribir valores de propiedad en el almacén
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Max", UpSpinner.Max, "100")
    Call PropBag.WriteProperty("Min", UpSpinner.Min, "0")
    Call PropBag.WriteProperty("MaxLength", txtSpinner.MaxLength, 3)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Increment", UpSpinner.Increment, 1)
'    Call PropBag.WriteProperty("ToolTipText", txtSpinner.ToolTipText, "")
    Call PropBag.WriteProperty("BackColor", txtSpinner.BackColor, &H80000005)
    Call PropBag.WriteProperty("Font", txtSpinner.Font, Ambient.Font)
    Call PropBag.WriteProperty("FontBold", txtSpinner.FontBold, 0)
    Call PropBag.WriteProperty("FontItalic", txtSpinner.FontItalic, 0)
    Call PropBag.WriteProperty("FontName", txtSpinner.FontName, "")
    Call PropBag.WriteProperty("FontSize", txtSpinner.FontSize, 0)
    Call PropBag.WriteProperty("ForeColor", txtSpinner.ForeColor, &H80000008)
    Call PropBag.WriteProperty("Locked", txtSpinner.Locked, False)
End Sub
Private Sub UpSpinner_UpClick()
    RaiseEvent UpClick
    If Val(txtSpinner.Text) >= UpSpinner.Max Then
        txtSpinner.Text = UpSpinner.Max
    Else
        txtSpinner.Text = Int(Val(txtSpinner.Text)) + Int(UpSpinner.Increment)
    End If
End Sub
Private Sub UpSpinner_DownClick()
    RaiseEvent DownClick
    If Val(txtSpinner.Text) <= UpSpinner.Min Then
        txtSpinner.Text = UpSpinner.Min
    Else
        txtSpinner.Text = Int(Val(txtSpinner.Text)) - Int(UpSpinner.Increment)
    End If
End Sub
Private Sub txtSpinner_KeyPress(KeyAscii As Integer)
Dim lsCadeNum As String
    lsCadeNum = "-0123456789"
    RaiseEvent KeyPress(KeyAscii)
    If InStr(1, lsCadeNum, Chr(KeyAscii), vbTextCompare) = 0 Then
        KeyAscii = 0
    End If
End Sub
'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property
'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=txtSpinner,txtSpinner,-1,Text
Public Property Get Valor() As String
Attribute Valor.VB_Description = "Devuelve o establece el texto contenido en el control."
    Valor = txtSpinner.Text
End Property
Public Property Let Valor(ByVal New_Valor As String)
    txtSpinner.Text() = New_Valor
    PropertyChanged "Valor"
End Property
'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=UpSpinner,UpSpinner,-1,Increment
Public Property Get Increment() As Long
Attribute Increment.VB_Description = "Obtiene o establece la cantidad que va a cambiar la posición en cada clic"
    Increment = UpSpinner.Increment
End Property

Public Property Let Increment(ByVal New_Increment As Long)
    If New_Increment < 1 Then
        New_Increment = 1
    End If
    If New_Increment > 100 Then
        New_Increment = 100
    End If
    UpSpinner.Increment() = New_Increment
    PropertyChanged "Increment"
End Property
'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=UpSpinner,UpSpinner,-1,Max
Public Property Get Max() As String
Attribute Max.VB_Description = "Devuelve o establece el texto contenido en el control."
    Max = UpSpinner.Max
End Property
Public Property Let Max(ByVal New_Max As String)
    UpSpinner.Max() = New_Max
    PropertyChanged "Max"
    If Len(New_Max) > Len(Min) Then
        MaxLength = Len(New_Max)
    Else
        MaxLength = Len(Min)
    End If
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=UpSpinner,UpSpinner,-1,Min
Public Property Get Min() As String
Attribute Min.VB_Description = "Devuelve o establece el texto contenido en el control."
    Min = UpSpinner.Min
End Property

Public Property Let Min(ByVal New_Min As String)
    If New_Min > Max Then
        UpSpinner.Min() = Max
    Else
        UpSpinner.Min() = New_Min
    End If
    PropertyChanged "Min"
    If Len(New_Min) > Len(Max) Then
        MaxLength = Len(Min)
    Else
        MaxLength = Len(Max)
    End If
    UpSpinner.Value = Min
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=txtSpinner,txtSpinner,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Devuelve o establece el color de fondo usado para mostrar texto y gráficos en un objeto."
    BackColor = txtSpinner.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    txtSpinner.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=txtSpinner,txtSpinner,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Devuelve un objeto Font."
Attribute Font.VB_UserMemId = -512
    Set Font = txtSpinner.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set txtSpinner.Font = New_Font
    PropertyChanged "Font"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=txtSpinner,txtSpinner,-1,FontBold
Public Property Get FontBold() As Boolean
Attribute FontBold.VB_Description = "Devuelve o establece el estilo negrita de una fuente."
    FontBold = txtSpinner.FontBold
End Property

Public Property Let FontBold(ByVal New_FontBold As Boolean)
    txtSpinner.FontBold() = New_FontBold
    PropertyChanged "FontBold"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=txtSpinner,txtSpinner,-1,FontItalic
Public Property Get FontItalic() As Boolean
Attribute FontItalic.VB_Description = "Devuelve o establece el estilo cursiva de una fuente."
    FontItalic = txtSpinner.FontItalic
End Property

Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
    txtSpinner.FontItalic() = New_FontItalic
    PropertyChanged "FontItalic"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=txtSpinner,txtSpinner,-1,FontName
Public Property Get FontName() As String
Attribute FontName.VB_Description = "Especifica el nombre de la fuente que aparece en cada fila del nivel especificado."
    FontName = txtSpinner.FontName
End Property

Public Property Let FontName(ByVal New_FontName As String)
    txtSpinner.FontName() = New_FontName
    PropertyChanged "FontName"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=txtSpinner,txtSpinner,-1,FontSize
Public Property Get FontSize() As Single
Attribute FontSize.VB_Description = "Especifica el tamaño (en puntos) de la fuente que aparece en cada fila del nivel especificado."
    FontSize = txtSpinner.FontSize
End Property

Public Property Let FontSize(ByVal New_FontSize As Single)
    If New_FontSize <= 0 Then
         New_FontSize = 8
    End If
    txtSpinner.FontSize() = New_FontSize
    PropertyChanged "FontSize"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=txtSpinner,txtSpinner,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Devuelve o establece el color de primer plano usado para mostrar textos y gráficos en un objeto."
    ForeColor = txtSpinner.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    txtSpinner.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=txtSpinner,txtSpinner,-1,Locked
Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "Determina si se puede modificar un control."
    Locked = txtSpinner.Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
    txtSpinner.Locked() = New_Locked
    PropertyChanged "Locked"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=UpSpinner,UpSpinner,-1,ToolTipText
Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_Description = "Devuelve o establece el texto mostrado cuando el mouse se sitúa sobre un control."
    ToolTipText = UpSpinner.ToolTipText
End Property

