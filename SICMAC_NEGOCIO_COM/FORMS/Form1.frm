VERSION 5.00
Object = "{160AE063-3670-11D5-8214-000103686C75}#4.0#0"; "PryOcxExplorer.ocx"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   5760
   ClientLeft      =   2685
   ClientTop       =   2340
   ClientWidth     =   8235
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   8235
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      Height          =   540
      Left            =   2070
      TabIndex        =   7
      Top             =   3150
      Width           =   2535
   End
   Begin SICMACT.Usuario Usuario1 
      Left            =   2640
      Top             =   2025
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.TextBox Text1 
      Height          =   525
      Left            =   2355
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   1050
      Width           =   2400
   End
   Begin PryOcxExplorer.OcxCdlgExplorer CdlgExplorer1 
      Height          =   495
      Left            =   2805
      TabIndex        =   5
      Top             =   330
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   873
      Filtro          =   ""
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Con Texto"
      Height          =   465
      Left            =   105
      TabIndex        =   4
      Top             =   2325
      Width           =   1650
   End
   Begin VB.CommandButton Command4 
      Caption         =   "ConImagen"
      Height          =   465
      Left            =   105
      TabIndex        =   3
      Top             =   1875
      Width           =   1650
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Normal"
      Height          =   465
      Left            =   105
      TabIndex        =   2
      Top             =   1425
      Width           =   1650
   End
   Begin VB.CommandButton Command2 
      Caption         =   "INI"
      Height          =   480
      Left            =   90
      TabIndex        =   1
      Top             =   765
      Width           =   1650
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Previo"
      Height          =   525
      Left            =   105
      TabIndex        =   0
      Top             =   165
      Width           =   1650
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   315
      Left            =   3765
      TabIndex        =   8
      Top             =   285
      Width           =   765
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub Draw3DBorder(TargetForm As Form, TargetControl As Control, RaisedBorder As Integer, BorderWidth As Integer)

Dim BorderOffset As Integer
Dim X1 As Integer, X2 As Integer
Dim Y1 As Integer, Y2 As Integer
Dim OriginalForeColor As Long, OriginalDrawWidth As Long
Dim UpperColor As Long, LowerColor As Long

'Define how far the 3D lines are drawn from the outer edges of the
'control. Modify to your taste.

BorderOffset = 8

'Define the four corners of the 3D box to be drawn.
X1 = TargetControl.Left - BorderOffset
Y1 = TargetControl.Top - BorderOffset
X2 = X1 + TargetControl.Width + (BorderOffset * 2)
Y2 = Y1 + TargetControl.Height + (BorderOffset * 2)

'Change the form's ForeColor and DrawWidth properties,
'so we'll save them first and restore when done.

OriginalForeColor = TargetForm.ForeColor
OriginalDrawWidth = TargetForm.DrawWidth

'If RaisedBorder is True, the white lines are drawn on the
'top and left sides.

If RaisedBorder Then
   UpperColor = QBColor(15)
   LowerColor = QBColor(8)
Else
   UpperColor = QBColor(8)
   LowerColor = QBColor(15)
End If

'Draw line on left.
TargetForm.DrawWidth = BorderWidth
TargetForm.ForeColor = UpperColor
TargetForm.Line (X1, Y2)-(X1, Y1)

'Draw line on top.
TargetForm.Line -(X2, Y1)

'Draw line on right.
TargetForm.ForeColor = LowerColor
TargetForm.Line -(X2, Y2)

'Draw line on bottom.
TargetForm.Line -(X1, Y2)

'Return the form's properties to their original state.
TargetForm.ForeColor = OriginalForeColor
TargetForm.DrawWidth = OriginalDrawWidth
End Sub
Private Sub Command1_Click()
Dim Prev As Previo.clsPrevio
Dim R As ADODB.Recordset
    'Set R = New ADODB.Recordset
    Set Prev = New clsPrevio
    Prev.Show "", "", True
End Sub

Private Sub Command2_Click()
Dim i As New clsIni.ClasIni
    i.CrearArchivoIni
End Sub

Private Sub Command3_Click()
    CdlgExplorer1.Filtro = "Todos los Archivos|*.*"
    CdlgExplorer1.TipoVentana = Normal
    CdlgExplorer1.nHwd = Me.hwnd
    CdlgExplorer1.Show
    MsgBox CdlgExplorer1.Ruta
End Sub

Private Sub Command4_Click()
    CdlgExplorer1.Filtro = "Archivos JPG|*.jpg|Archivos GIF|*.gif|Todos los Archivos|*.*"
    CdlgExplorer1.TipoVentana = Con_Imagen
    CdlgExplorer1.nHwd = Me.hwnd
    CdlgExplorer1.Show
    MsgBox CdlgExplorer1.Ruta
End Sub

Private Sub Command5_Click()
    CdlgExplorer1.Filtro = "Archivos Texto|*.txt|Archivos Word|*.doc|Todos los Archivos|*.*"
    CdlgExplorer1.TipoVentana = Con_Texto
    CdlgExplorer1.nHwd = Me.hwnd
    CdlgExplorer1.Show
    MsgBox CdlgExplorer1.Ruta
End Sub
Public Sub Inicioform2()
    Form2.Inicio
End Sub
Public Sub Inicioform3()
    Form3.Inicio
End Sub

Private Sub Command6_Click()
'Dim oPersona As UPersona
'    Screen.MousePointer = 11
'    Set oPersona = New UPersona
'        oPersona.ObtieneClientexCodigo ("1120700169307")
'    Set oPersona = Nothing
'    Screen.MousePointer = 0

Call CallByName(Me, Text1.Text, VbMethod)

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    KeyCode = KeyCode
End Sub

Private Sub Form_Load()
    Call Draw3DBorder(Form1, Command6, True, 2)
End Sub

Private Sub Form_Resize()
'Dim i
'Dim y
'Form1.Cls
'Form1.AutoRedraw = True
'Form1.DrawStyle = 6
'Form1.DrawMode = 13
'Form1.DrawWidth = 2
'Form1.ScaleMode = 3
'Form1.ScaleHeight = (256 * 2)
'For i = 0 To 255
'Form1.Line (0, y)-(Form1.Width, y + 2), RGB(255, 0, i), BF
'y = y + 2
'Next i
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = KeyCode
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii
End Sub


