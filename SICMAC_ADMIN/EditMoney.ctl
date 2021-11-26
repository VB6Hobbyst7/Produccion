VERSION 5.00
Begin VB.UserControl EditMoney 
   ClientHeight    =   375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1695
   ScaleHeight     =   375
   ScaleWidth      =   1695
   ToolboxBitmap   =   "EditMoney.ctx":0000
   Begin VB.TextBox txtValor 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   0
      MaxLength       =   14
      TabIndex        =   0
      Top             =   0
      Width           =   1695
   End
End
Attribute VB_Name = "EditMoney"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim gbEstado As Boolean
Dim gnNumDec As Integer
Enum BorderStyle
    None = 0
    [Fixed Single] = 1
End Enum
Public Event Change()
Public Event KeyPress(KeyAscii As Integer)

Private Function TienePunto(psCadena As String) As Boolean
If InStr(1, psCadena, ".", vbTextCompare) > 0 Then
    TienePunto = True
Else
    TienePunto = False
End If
End Function

Private Function NumDecimal(psCadena As String) As Integer
Dim lnPos As Integer
lnPos = InStr(1, psCadena, ".", vbTextCompare)
If lnPos > 0 Then
    NumDecimal = Len(psCadena) - lnPos
Else
    NumDecimal = 0
End If
End Function

Private Sub txtValor_Change()
txtValor.SelStart = Len(txtValor)
gnNumDec = NumDecimal(txtValor)
If gbEstado And txtValor <> "" Then
    Select Case gnNumDec
        Case 0
                txtValor = Format(txtValor, "#,##0")
        Case 1
                txtValor = Format(txtValor, "#,##0.0")
        Case Else
                txtValor = Format(txtValor, "#,##0.00")
    End Select
End If
If txtValor = "" Then
    txtValor = 0
End If
gbEstado = False
RaiseEvent Change
End Sub

Private Sub txtValor_KeyPress(KeyAscii As Integer)
Dim lnNum As Integer
Dim Car As String * 1
RaiseEvent KeyPress(KeyAscii)
Car = Chr$(KeyAscii)
If (Car < "0" Or Car > "9") And Car <> Chr$(8) And Car <> "." Then
    Beep
    KeyAscii = 0
Else
    If TienePunto(txtValor) Then
        If (Car = "." Or gnNumDec = 2) And Car <> Chr$(8) Then
            If txtValor.SelLength = 0 Then
                Beep
                KeyAscii = 0
                gbEstado = False
            End If
        ElseIf Car = Chr$(8) And gnNumDec = 1 Then
            gbEstado = False
        Else
            gbEstado = True
        End If
    Else
        If Car = "." Then
            If txtValor.SelStart = Len(txtValor) Then
                gbEstado = False
            Else
                gbEstado = True
            End If
        Else
            gbEstado = True
        End If
    End If
End If
End Sub

Public Property Get Value() As Currency
If Mid(txtValor, 1, 1) = "," Then
    txtValor = Mid(txtValor, 2, Len(txtValor) - 1)
End If
If txtValor <> "" And txtValor <> "." Then
    Value = CCur(txtValor)
Else
    Value = 0
End If
End Property

Public Property Let Value(ByVal vNewValue As Currency)
txtValor = Trim(Str(vNewValue))
End Property

Private Sub txtValor_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then
    If txtValor <> "" Then
        Select Case gnNumDec
            Case 0
                txtValor = Format(txtValor, "#,##0")
            Case 1
                txtValor = Format(txtValor, "#,##0.0")
            Case 2
                txtValor = Format(txtValor, "#,##0.00")
        End Select
    End If
    gbEstado = False
End If
End Sub

Private Sub txtValor_LostFocus()
If Value > 0 Then
    txtValor = Format(txtValor, "#,##0.00")
End If
End Sub

Private Sub UserControl_Initialize()
gbEstado = False
txtValor = "0"
End Sub

Public Sub MarcaTexto()
    txtValor.SelStart = 0
    txtValor.SelLength = Len(txtValor.Text)
End Sub

Public Sub psSoles(pbSoles As Boolean)
    If pbSoles Then
        txtValor.BackColor = &HFFFFFF
    Else
        txtValor.BackColor = &HFF00&
    End If
End Sub

Public Sub psTipOpe(pbIngreso As Boolean)
    If pbIngreso Then
        txtValor.ForeColor = &H0&
    Else
        txtValor.ForeColor = &HFF&
    End If
End Sub

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=txtValor,txtValor,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Devuelve un objeto Font."
Attribute Font.VB_UserMemId = -512
    Set Font = txtValor.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set txtValor.Font = New_Font
    PropertyChanged "Font"
End Property

'Cargar valores de propiedad desde el almacén
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Set txtValor.Font = PropBag.ReadProperty("Font", Ambient.Font)
    txtValor.Appearance = PropBag.ReadProperty("Appearance", 1)
    txtValor.BackColor = PropBag.ReadProperty("BackColor", &HFFFFFF)
    txtValor.ForeColor = PropBag.ReadProperty("ForeColor", &H0&)
    txtValor.Text = PropBag.ReadProperty("Text", "")
    UserControl.Enabled = PropBag.ReadProperty("Enabled", Verdadero)
    txtValor.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
End Sub

Private Sub UserControl_Resize()
    txtValor.Width = UserControl.Width
    txtValor.Height = UserControl.Height
End Sub

'Escribir valores de propiedad en el almacén
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Font", txtValor.Font, Ambient.Font)
    Call PropBag.WriteProperty("Appearance", txtValor.Appearance, 1)
    Call PropBag.WriteProperty("BackColor", txtValor.BackColor, &HFFFFFF)
    Call PropBag.WriteProperty("ForeColor", txtValor.ForeColor, &H0&)
    Call PropBag.WriteProperty("Text", txtValor.Text, "")
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, Verdadero)
    Call PropBag.WriteProperty("BorderStyle", txtValor.BorderStyle, 1)
End Sub

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=txtValor,txtValor,-1,Appearance
Public Property Get Appearance() As Apariencia
Attribute Appearance.VB_Description = "Devuelve o establece si los objetos se dibujan en tiempo de ejecución con efectos 3D."
    Appearance = txtValor.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As Apariencia)
    txtValor.Appearance() = New_Appearance
    PropertyChanged "Appearance"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=txtValor,txtValor,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Devuelve o establece el color de fondo usado para mostrar texto y gráficos en un objeto."
    BackColor = txtValor.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    txtValor.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=txtValor,txtValor,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Devuelve o establece el color de primer plano usado para mostrar textos y gráficos en un objeto."
    ForeColor = txtValor.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    txtValor.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=txtValor,txtValor,-1,Text
Public Property Get Text() As String
Attribute Text.VB_UserMemId = 0
Attribute Text.VB_MemberFlags = "200"
    Text = txtValor.Text
End Property

Public Property Let Text(ByVal New_Text As String)
    txtValor.Text() = New_Text
    PropertyChanged "Text"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Devuelve o establece un valor que determina si un objeto puede responder a eventos generados por el usuario."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=txtValor,txtValor,-1,BorderStyle
Public Property Get BorderStyle() As BorderStyle
Attribute BorderStyle.VB_Description = "Devuelve o establece el estilo del borde de un objeto."
    BorderStyle = txtValor.BorderStyle
End Property
Public Property Let BorderStyle(ByVal New_BorderStyle As BorderStyle)
    txtValor.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

