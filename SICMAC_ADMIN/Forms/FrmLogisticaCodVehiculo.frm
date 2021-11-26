VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmLogisticaCodVehiculo 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6345
   ClientLeft      =   2145
   ClientTop       =   1845
   ClientWidth     =   7605
   Icon            =   "FrmLogisticaCodVehiculo.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   7605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   675
      Left            =   60
      TabIndex        =   2
      Top             =   -60
      Width           =   7455
      Begin VB.ComboBox cboAnio 
         Height          =   315
         Left            =   2820
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Registro de Activos del año"
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
         TabIndex        =   4
         Top             =   300
         Width           =   2370
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSH 
      Height          =   5205
      Left            =   60
      TabIndex        =   0
      Top             =   645
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   9181
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      BackColorBkg    =   -2147483643
      GridColor       =   -2147483633
      ScrollBars      =   2
      SelectionMode   =   1
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   360
      Left            =   6300
      TabIndex        =   1
      Top             =   5940
      Width           =   1215
   End
End
Attribute VB_Name = "FrmLogisticaCodVehiculo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sCod As String
Dim sSerie As String
Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim rs As New ADODB.Recordset
Dim oConn As DConecta

Set oConn = New DConecta

If oConn.AbreConexion Then
   cboAnio.Clear
   Set rs = oConn.CargaRecordSet("Select distinct nAnio from bsactivofijo where  cbscod like '11203%' and dBaja is null Order by nAnio")
   If Not rs.EOF Then
      Do While Not rs.EOF
         cboAnio.AddItem rs!nAnio
         rs.MoveNext
      Loop
   End If
   cboAnio.ListIndex = cboAnio.ListCount - 1
   oConn.CierraConexion
End If
End Sub

Private Sub cboAnio_Click()
Dim LV As DLogVehiculo, i As Integer
Dim rs As New ADODB.Recordset

Set LV = New DLogVehiculo

With MSH
    .Clear
    .Rows = 2
    .RowHeight(0) = 300
    .ColWidth(0) = 300
    .ColWidth(1) = 0
    .ColWidth(2) = 860:    .ColAlignment(2) = 4
    .ColWidth(3) = 2000:   .ColAlignment(3) = 1
    .ColWidth(4) = 4000
End With

If cboAnio.ListIndex < 0 Then Exit Sub

Set rs = LV.GetVehiculoSeries(CInt(cboAnio.Text))
i = 1
While Not rs.EOF
    With MSH
        .TextMatrix(i, 0) = Format(i, "00")
        .TextMatrix(i, 1) = "" 'rs!cBSCod & " " & rs!cSerie
        .TextMatrix(i, 2) = rs!cBSCod
        .TextMatrix(i, 3) = rs!cSerie
        .TextMatrix(i, 4) = rs!cDescripcion
        .Rows = .Rows + 1
        i = 1 + i
    End With
    rs.MoveNext
Wend
If Not (rs.EOF And rs.BOF) Then MSH.Rows = MSH.Rows - 1
End Sub

Private Sub MSH_DblClick()
'If MSH.Col = 1 Or MSH.Col = 4 Then
    If (MSH.row > 0) Then
    FrmLogisticaRegVehiculo.txtBSCod = Trim(MSH.TextMatrix(MSH.row, 2))
    FrmLogisticaRegVehiculo.txtSerie = Trim(MSH.TextMatrix(MSH.row, 3))
    'FrmLogisticaRegVehiculo.LblCodCatalogo = Trim(MSH.TextMatrix(MSH.Row, 1))
    FrmLogisticaRegVehiculo.LblDescripcion.Text = Trim(MSH.TextMatrix(MSH.row, 4))
    Unload Me
    End If
'End If
End Sub

Private Sub MSH_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   MSH_DblClick
End If
End Sub
