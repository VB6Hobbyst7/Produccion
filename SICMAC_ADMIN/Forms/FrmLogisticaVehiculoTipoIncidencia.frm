VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmLogisticaVehiculoTipoIncidencia 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registro de Nuevas Incidencias"
   ClientHeight    =   5340
   ClientLeft      =   2760
   ClientTop       =   2175
   ClientWidth     =   6000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   360
      Left            =   4680
      TabIndex        =   5
      Top             =   4920
      Width           =   1200
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Flex 
      Height          =   4035
      Left            =   60
      TabIndex        =   2
      Top             =   780
      Width           =   5835
      _ExtentX        =   10292
      _ExtentY        =   7117
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
   Begin VB.Frame Frame16 
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
      Height          =   705
      Left            =   75
      TabIndex        =   0
      Top             =   60
      Width           =   5850
      Begin VB.CommandButton CmdCancelar 
         Height          =   380
         Left            =   5220
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   210
         Width           =   460
      End
      Begin VB.CommandButton CmdGrabar 
         Height          =   380
         Left            =   4740
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   210
         Width           =   460
      End
      Begin VB.TextBox TxtIncidencia 
         Height          =   315
         Left            =   150
         TabIndex        =   1
         Top             =   240
         Width           =   4545
      End
   End
   Begin VB.Image imgNX 
      Height          =   240
      Left            =   6840
      Picture         =   "FrmLogisticaVehiculoTipoIncidencia.frx":0000
      Top             =   1020
      Width           =   240
   End
   Begin VB.Image imgOK 
      Height          =   240
      Left            =   6840
      Picture         =   "FrmLogisticaVehiculoTipoIncidencia.frx":0342
      Top             =   720
      Width           =   240
   End
End
Attribute VB_Name = "FrmLogisticaVehiculoTipoIncidencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdCancelar_Click()
'Me.TxtIncidencia = ""
End Sub

Private Sub cmdGrabar_Click()
Dim LV As DLogVehiculo
Dim opt As Integer
If Trim(Me.TxtIncidencia) = "" Then
    MsgBox "Ingrese el tipo de incidencia", vbInformation, "AVISO"
    Exit Sub
End If

opt = MsgBox("Esta seguro de Grabar", vbQuestion + vbYesNo, "AVISO")
If opt = vbNo Then Exit Sub
Set LV = New DLogVehiculo
Call LV.InsertTipoIncidencia(Trim(Me.TxtIncidencia))
CmdCancelar_Click
MarcoIncidencia
CargaTipoIncidencia
Set LV = Nothing
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
    .ColWidth(0) = 400
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
Dim LV As DLogVehiculo
Dim Rs As ADODB.Recordset
Dim i As Integer
Set LV = New DLogVehiculo
Set Rs = LV.GetTipoIncidencia
i = 1
While Not Rs.EOF
    Flex.TextMatrix(i, 1) = Rs!cDescripcion
    Flex.Rows = Flex.Rows + 1
    Rs.MoveNext
    i = i + 1
Wend
If Not (Rs.EOF And Rs.BOF) Then Flex.Rows = Flex.Rows - 1

Set Rs = Nothing
Set LV = Nothing
End Sub


Private Sub Form_Unload(Cancel As Integer)
FrmLogisticavehiculoVarios.CargaComboTipoIncidencia
End Sub

Private Sub TxtIncidencia_KeyPress(KeyAscii As Integer)
KeyAscii = intfMayusculas(KeyAscii)
If KeyAscii = 13 Then Me.cmdGrabar.SetFocus
End Sub
