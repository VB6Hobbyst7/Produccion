VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmLogTipoIncidencia 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registro de Nuevas Incidencias"
   ClientHeight    =   4830
   ClientLeft      =   1605
   ClientTop       =   3090
   ClientWidth     =   8085
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   8085
   ShowInTaskbar   =   0   'False
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Flex 
      Height          =   3615
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   6376
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      BackColorBkg    =   -2147483643
      GridColor       =   -2147483633
      FocusRect       =   0
      ScrollBars      =   2
      SelectionMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   3
   End
   Begin VB.Frame fraVis 
      Height          =   855
      Left            =   120
      TabIndex        =   5
      Top             =   3840
      Width           =   7815
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "&Agregar"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   300
         Width           =   1200
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   6480
         TabIndex        =   6
         Top             =   300
         Width           =   1200
      End
   End
   Begin VB.Frame fraReg 
      Caption         =   "Descripción"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1005
      Left            =   120
      TabIndex        =   0
      Top             =   3720
      Visible         =   0   'False
      Width           =   7845
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "Cancelar"
         Height          =   320
         Left            =   6540
         TabIndex        =   4
         Top             =   600
         Width           =   1125
      End
      Begin VB.CommandButton CmdGrabar 
         Caption         =   "Aceptar"
         Height          =   320
         Left            =   5340
         TabIndex        =   3
         Top             =   600
         Width           =   1125
      End
      Begin VB.TextBox TxtIncidencia 
         Height          =   315
         Left            =   150
         TabIndex        =   1
         Top             =   240
         Width           =   7485
      End
   End
   Begin VB.Image imgNX 
      Height          =   240
      Left            =   480
      Picture         =   "frmLogTipoIncidencia.frx":0000
      Top             =   1680
      Width           =   240
   End
   Begin VB.Image imgOK 
      Height          =   240
      Left            =   6840
      Picture         =   "frmLogTipoIncidencia.frx":0342
      Top             =   720
      Width           =   240
   End
End
Attribute VB_Name = "frmLogTipoIncidencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAgregar_Click()
fraVis.Visible = False
fraReg.Visible = True
Flex.BackColor = "&H8000000F"
Flex.BackColorBkg = "&H8000000F"
TxtIncidencia.Text = ""
TxtIncidencia.SetFocus
End Sub

Private Sub CmdCancelar_Click()
Flex.BackColor = "&H80000005"
Flex.BackColorBkg = "&H80000005"
fraVis.Visible = True
fraReg.Visible = False
Flex.SetFocus
End Sub

Private Sub cmdGrabar_Click()
Dim oConn As New DConecta
Dim rs As New ADODB.Recordset
Dim nCod As Integer

If Trim(Me.TxtIncidencia) = "" Then
    MsgBox "Ingrese la descripción de la incidencia", vbInformation, "AVISO"
    Exit Sub
End If

If oConn.AbreConexion Then
   Set rs = oConn.CargaRecordSet("select nCod=Max(nConsValor) from Constante where nConsCod=9025 and nconscod<>nconsvalor")
   If Not rs.EOF Then
      nCod = rs!nCod + 1
   Else
      nCod = 1
   End If
   oConn.CierraConexion
End If
Set rs = Nothing

If MsgBox("Esta seguro de Grabar", vbQuestion + vbYesNo, "AVISO") = vbYes Then
   If oConn.AbreConexion Then
      oConn.Ejecutar ("INSERT INTO Constante (nConsCod,nConsValor,cConsDescripcion) VALUES (9025," & nCod & ",'" & TxtIncidencia.Text & "') ")
      oConn.CierraConexion
   End If
End If
Me.TxtIncidencia = ""
CmdCancelar_Click
MarcoIncidencia
CargaTipoIncidencia
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
MarcoIncidencia
CargaTipoIncidencia
Set cmdGrabar.Picture = imgOK.Picture
Set CmdCancelar.Picture = imgNX.Picture
End Sub

Sub MarcoIncidencia()
With Flex
    .Clear
    .Rows = 2
    .HighLight = flexHighlightWithFocus
    .ColWidth(0) = 400:    .ColAlignment(0) = 4
    .ColWidth(1) = 6000
    .TextMatrix(0, 1) = "Descripcion"
End With
End Sub
Public Function intfMayusculas(intTecla As Integer) As Integer
 If Chr(intTecla) >= "a" And Chr(intTecla) <= "z" Then
    intTecla = intTecla - 32
 End If
 If intTecla = 39 Then
    intTecla = 0
 End If
 If intTecla = 209 Or intTecla = 241 Or intTecla = 8 Or intTecla = 32 Then
    intfMayusculas = Asc(UCase(Chr(intTecla)))
     Exit Function
 End If
 intfMayusculas = intTecla
End Function

Sub CargaTipoIncidencia()
Dim LV As DLogVehiculos
Dim rs As ADODB.Recordset
Dim i As Integer
Set LV = New DLogVehiculos
Set rs = LV.GetTipoIncidencia
i = 1
While Not rs.EOF
    Flex.TextMatrix(i, 0) = rs!nIncidencia
    Flex.TextMatrix(i, 1) = rs!cIncidencia
    Flex.Rows = Flex.Rows + 1
    rs.MoveNext
    i = i + 1
Wend
If Not (rs.EOF And rs.BOF) Then Flex.Rows = Flex.Rows - 1
Set rs = Nothing
Set LV = Nothing
End Sub

Private Sub TxtIncidencia_KeyPress(KeyAscii As Integer)
KeyAscii = intfMayusculas(KeyAscii)
If KeyAscii = 13 Then Me.cmdGrabar.SetFocus
End Sub
