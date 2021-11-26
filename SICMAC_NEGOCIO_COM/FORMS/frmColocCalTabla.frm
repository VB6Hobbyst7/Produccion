VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmColocCalTabla 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Colocaciones : Mantenimiento de Tabla de Calificaciones"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7710
   Icon            =   "frmColocCalTabla.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   7710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraDatos 
      Height          =   1335
      Left            =   80
      TabIndex        =   5
      Top             =   0
      Width           =   7575
      Begin VB.OptionButton OptCal 
         Caption         =   "Refinanciado"
         Height          =   255
         Index           =   1
         Left            =   5880
         TabIndex        =   20
         Top             =   1000
         Width           =   1455
      End
      Begin VB.OptionButton OptCal 
         Caption         =   "Normal"
         Height          =   255
         Index           =   0
         Left            =   4800
         TabIndex        =   19
         Top             =   1000
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.ComboBox CboGarantia 
         Height          =   315
         Left            =   5640
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox TxtProvision 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3000
         TabIndex        =   8
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox TxtRangoF 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1200
         TabIndex        =   9
         Top             =   960
         Width           =   795
      End
      Begin VB.TextBox TxtRangoI 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1200
         TabIndex        =   7
         Top             =   600
         Width           =   795
      End
      Begin VB.TextBox TxtDescripcion 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   6
         Top             =   240
         Width           =   6015
      End
      Begin VB.TextBox TxtCalificacion 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3720
         TabIndex        =   11
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Garantia"
         Height          =   195
         Left            =   4800
         TabIndex        =   18
         Top             =   720
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Provision"
         Height          =   195
         Left            =   2100
         TabIndex        =   16
         Top             =   660
         Width           =   645
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Rango Final"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   1020
         Width           =   855
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Rango Inicial"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   660
         Width           =   930
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Descripcion"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   270
         Width           =   840
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Calificacion"
         Height          =   195
         Left            =   2100
         TabIndex        =   10
         Top             =   1020
         Width           =   810
      End
   End
   Begin VB.Frame FraTarifario 
      Height          =   2655
      Left            =   80
      TabIndex        =   3
      Top             =   1350
      Width           =   7575
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MshTabla 
         Height          =   2295
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   7365
         _ExtentX        =   12991
         _ExtentY        =   4048
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         SelectionMode   =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   6
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   80
      TabIndex        =   0
      Top             =   3975
      Width           =   7575
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   4560
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3120
         TabIndex        =   13
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   1680
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmColocCalTabla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim opt As Integer
Dim psServer As String

Private Sub CboGarantia_Click()
CargaDatos
End Sub

Private Sub CmdAceptar_Click()
Dim ObjColoc As COMNCredito.NCOMColocEval
Dim ban As Integer

'If opt = -1 Then Exit Sub
If VerificaText Then Exit Sub
'CargaDatos
'CmdNuevo.Caption = "&Nuevo"

'Para poder modificar los Datos
'Me.fraDatos.Enabled = False

Set ObjColoc = New COMNCredito.NCOMColocEval
'Select Case opt
'    Case 0:
'        ban = MsgBox("Esta Seguro de Guardar", vbQuestion + vbYesNo, "AVISO")
'        If ban = vbYes Then Call ObjColoc.InsertaColocCalifTabla(psServer, TxtDescripcion, Me.TxtRangoI, Me.TxtRangoF, Me.TxtCalificacion, Me.TxtProvision)
'    Case 1:
        ban = MsgBox("Esta Seguro de Modificar?", vbQuestion + vbYesNo, "AVISO")
        If ban = vbYes Then Call ObjColoc.UpdateColocCalifTabla(psServer, TxtDescripcion.Text, TxtRangoI.Text, TxtRangoF.Text, TxtCalificacion.Text, TxtProvision.Text, CInt(MshTabla.TextMatrix(MshTabla.Row, 0)))
'        CmdNuevo.Caption = "&Nuevo"
'End Select
Limpiar
CargaDatos
End Sub

Private Sub CmdModificar_Click()
Me.FraDatos.Enabled = True
TxtDescripcion.SetFocus
opt = 1
cmdCancelar.Caption = "&Cancelar"
End Sub

Private Sub cmdCancelar_Click()
    Limpiar
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
'Dim objRcd As nRcdReportes
Dim ObjRiesgo As COMDCredito.DCOMColocEval
Dim rs As ADODB.Recordset
Set ObjRiesgo = New COMDCredito.DCOMColocEval
Set rs = ObjRiesgo.TiposGaratiaRiesgos
If rs.EOF And rs.BOF Then
    MsgBox "Tipos de Datos no se encuentra en la Tabla Constante.", vbInformation, "AVISO"
    Set rs = Nothing
    Set ObjRiesgo = Nothing
'    Set objRcd = Nothing
    Exit Sub
End If
While Not rs.EOF
    Me.CboGarantia.AddItem Trim(rs!cConsDescripcion) & Space(100) & Trim(rs!nConsValor)
    rs.MoveNext
Wend

'Set objRcd = New nRcdReportes
    psServer = ObjRiesgo.GetServerConsol
'Set objRcd = Nothing
Set rs = Nothing
Set ObjRiesgo = Nothing
Me.CboGarantia.ListIndex = 0
CargaDatos
Me.Icon = LoadPicture(App.path & gsRutaIcono)
End Sub

Sub CargaDatos()
Dim ObjEvalCal As COMNCredito.NCOMColocEval
Dim rs As New ADODB.Recordset
Dim nLista As Integer
Dim sTipo As String
sTipo = IIf(Me.OptCal(0).value = True, "N", "R")
nLista = Me.CboGarantia.ListIndex + 1
Set ObjEvalCal = New COMNCredito.NCOMColocEval
  
    Set rs = ObjEvalCal.ObtieneTablaCalif(psServer)
Set ObjEvalCal = Nothing

MshTabla.Clear
MshTabla.Rows = 2
Marco

With MshTabla
    While Not rs.EOF

        .TextMatrix(.Rows - 1, 0) = Format(rs!NCalCodTab, "0000")
        .TextMatrix(.Rows - 1, 1) = rs!cCalDes
        .TextMatrix(.Rows - 1, 2) = IIf(IsNull(rs!nRangIni), "", rs!nRangIni)
        .TextMatrix(.Rows - 1, 3) = IIf(IsNull(rs!nRangFin), "", rs!nRangFin)
        .TextMatrix(.Rows - 1, 4) = IIf(IsNull(rs!cCalif), "", rs!cCalif)
        .TextMatrix(.Rows - 1, 5) = IIf(IsNull(rs!nProvision), "", Format(rs!nProvision, "#.00"))
        If Right(Format(rs!NCalCodTab, "000"), 1) = "0" Then
            .Row = .Rows - 1
            .Col = 0
            .CellFontBold = True
            .CellBackColor = vbYellow
            .Col = 1
            .CellFontBold = True
            .CellBackColor = vbYellow
        End If
        .Rows = .Rows + 1
        rs.MoveNext
    Wend
End With
If MshTabla.Rows > 2 Then MshTabla.Rows = MshTabla.Rows - 1

End Sub
Function VerificaText() As Boolean
VerificaText = False
If TxtDescripcion = "" Then
    VerificaText = True
    MsgBox "Ingrese Descripcion", vbInformation, "AVISO"
    Exit Function
End If
'If TxtRangoI < TxtRangoF Then
'    VerificaText = True
'    MsgBox "El Rango inicial no puede ser mayor", vbInformation, "AVISO"
'    Exit Function
'End If
End Function

Sub Limpiar()
TxtCalificacion = ""
TxtDescripcion = ""
TxtRangoF = ""
TxtRangoI = ""
TxtProvision = ""
End Sub
Sub Marco()
With MshTabla
    .TextMatrix(0, 0) = " Cod."
    .TextMatrix(0, 1) = " Descripcion"
    .TextMatrix(0, 2) = " Rango InI."
    .TextMatrix(0, 3) = " Rango Fin."
    .TextMatrix(0, 4) = " Calif."
    .TextMatrix(0, 5) = " Provision"
    .ColWidth(0) = 500
    .ColWidth(1) = 3500
    .ColWidth(2) = 900
    .ColWidth(3) = 900
    .ColWidth(4) = 500
    .ColWidth(5) = 700
End With

End Sub

Private Sub MshTabla_Click()
Limpiar

If MshTabla.Rows >= 2 And MshTabla.TextMatrix(MshTabla.Row, 0) <> "" Then
    TxtDescripcion = MshTabla.TextMatrix(MshTabla.Row, 1)
    TxtRangoI = MshTabla.TextMatrix(MshTabla.Row, 2)
    TxtRangoF = MshTabla.TextMatrix(MshTabla.Row, 3)
    TxtCalificacion = MshTabla.TextMatrix(MshTabla.Row, 4)
    TxtProvision = MshTabla.TextMatrix(MshTabla.Row, 5)
End If
End Sub

Private Sub OptCal_Click(Index As Integer)
CargaDatos
End Sub

Private Sub TxtCalificacion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Me.CmdAceptar.SetFocus
End Sub

Private Sub TxtDescripcion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Me.TxtRangoI.SetFocus
End Sub

Private Sub TxtProvision_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(TxtProvision, KeyAscii)
If KeyAscii = 13 Then Me.TxtCalificacion.SetFocus
End Sub

Private Sub TxtRangoF_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii, True)
If KeyAscii = 13 Then Me.TxtProvision.SetFocus
End Sub

Private Sub TxtRangoI_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii, True)
If KeyAscii = 13 Then Me.TxtRangoF.SetFocus
End Sub
