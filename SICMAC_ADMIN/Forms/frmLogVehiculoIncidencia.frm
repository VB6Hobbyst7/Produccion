VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmLogVehiculoIncidencia 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento de Incidencias"
   ClientHeight    =   3075
   ClientLeft      =   2160
   ClientTop       =   3600
   ClientWidth     =   8055
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   8055
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraVis 
      Caption         =   "Lista de Incidencias "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2835
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   7815
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "&Agregar"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   2360
         Width           =   1200
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   6480
         TabIndex        =   5
         Top             =   2360
         Width           =   1200
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Flex 
         Height          =   1995
         Left            =   120
         TabIndex        =   8
         Top             =   300
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   3519
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
   End
   Begin VB.Frame fraReg 
      Caption         =   "Registro de Incidencia "
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
      Height          =   2745
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   7845
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "Cancelar"
         Height          =   370
         Left            =   6240
         TabIndex        =   3
         Top             =   1980
         Width           =   1125
      End
      Begin VB.CommandButton CmdGrabar 
         Caption         =   "Aceptar"
         Height          =   370
         Left            =   4980
         TabIndex        =   2
         Top             =   1980
         Width           =   1125
      End
      Begin VB.TextBox TxtIncidencia 
         Height          =   315
         Left            =   1470
         TabIndex        =   1
         Top             =   1020
         Width           =   5865
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   1080
         Width           =   1020
      End
   End
End
Attribute VB_Name = "frmLogVehiculoIncidencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAgregar_Click()
fraVis.Visible = False
fraReg.Visible = True
Flex.BackColor = "&H8000000F"
Flex.BackColorBkg = "&H8000000F"
TxtIncidencia.Text = ""
TxtIncidencia.SetFocus
End Sub

Private Sub cmdCancelar_Click()
Flex.BackColor = "&H80000005"
Flex.BackColorBkg = "&H80000005"
fraVis.Visible = True
fraReg.Visible = False
Flex.SetFocus
End Sub

Private Sub CmdGrabar_Click()
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
cmdCancelar_Click
MarcoIncidencia
CargaTipoIncidencia
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
CentraForm Me
MarcoIncidencia
CargaTipoIncidencia
End Sub

Sub MarcoIncidencia()
With Flex
    .Clear
    .Rows = 2
    .HighLight = flexHighlightWithFocus
    .ColWidth(0) = 600:    .ColAlignment(0) = 4
    .ColWidth(1) = 6000
    .TextMatrix(0, 0) = "Código"
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
If KeyAscii = 13 Then Me.CmdGrabar.SetFocus
End Sub

