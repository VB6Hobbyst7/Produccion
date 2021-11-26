VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmLogSelector 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4005
   ClientLeft      =   1950
   ClientTop       =   3270
   ClientWidth     =   6975
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   6975
   StartUpPosition =   2  'CenterScreen
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFlex 
      Height          =   3435
      Left            =   120
      TabIndex        =   0
      Top             =   420
      Width           =   6705
      _ExtentX        =   11827
      _ExtentY        =   6059
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      BackColorBkg    =   -2147483643
      GridColor       =   -2147483633
      FocusRect       =   0
      HighLight       =   2
      SelectionMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
   End
   Begin VB.Label lblTitulo 
      AutoSize        =   -1  'True
      Caption         =   "Seleccione"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1140
   End
End
Attribute VB_Name = "frmLogSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public vpHaySeleccion As Boolean
Public vpCodigo As String
Public vpDescripcion As String

Dim cSQL As String, cTitulo As String

Public Sub Consulta(vSQL As String, vTitulo As String)
cSQL = vSQL
cTitulo = vTitulo
Me.Show 1
End Sub

Private Sub Form_Load()
Me.vpHaySeleccion = False
Me.vpCodigo = ""
Me.vpDescripcion = ""
LBLtITULO = cTitulo
ListaDatos
End Sub

Sub ListaDatos()
Dim rs As New ADODB.Recordset
Dim oConn As DConecta, n As Integer
Dim i As Integer

Set oConn = New DConecta
MSFlex.Clear
MSFlex.RowHeight(0) = 320
MSFlex.RowHeight(1) = 8
MSFlex.ColWidth(0) = 0
MSFlex.ColWidth(1) = 6500
MSFlex.ColWidth(2) = 0
MSFlex.ColWidth(3) = 0

If oConn.AbreConexion Then
   Set rs = oConn.CargaRecordSet(cSQL)
   If Not rs.EOF Then
      i = 0
      n = rs.Fields.Count
      If n = 3 Then
         MSFlex.ColWidth(1) = 5500
         MSFlex.ColWidth(2) = 1000
      End If
      If n = 4 Then
         MSFlex.ColWidth(1) = 3500
         MSFlex.ColWidth(2) = 1500
         MSFlex.ColWidth(3) = 1500
      End If
      Do While Not rs.EOF
         i = i + 1
         InsRow MSFlex, i
         MSFlex.TextMatrix(i, 0) = rs(0)
         MSFlex.TextMatrix(i, 1) = rs(1)
         If n = 3 Then
            MSFlex.TextMatrix(i, 2) = rs(2)
         End If
         If n = 4 Then
            MSFlex.TextMatrix(i, 2) = rs(2)
            MSFlex.TextMatrix(i, 3) = rs(3)
         End If
         rs.MoveNext
      Loop
   End If
   oConn.CierraConexion
End If
End Sub

Private Sub MSFlex_DblClick()
MSFlex_KeyPress 13
End Sub

Private Sub MSFlex_KeyPress(KeyAscii As Integer)
Dim i As Integer
If KeyAscii = 13 Then
   i = MSFlex.row
   Me.vpHaySeleccion = False
   If Len(Trim(MSFlex.TextMatrix(i, 0))) > 0 Then
      Me.vpCodigo = MSFlex.TextMatrix(i, 0)
      Me.vpDescripcion = MSFlex.TextMatrix(i, 1) + " " + MSFlex.TextMatrix(i, 2) + " " + MSFlex.TextMatrix(i, 3)
      Me.vpHaySeleccion = True
      Unload Me
   Else
      MsgBox "No existe una selección válida..." + Space(10), vbInformation
   End If
End If

If KeyAscii = 27 Then
   Me.vpCodigo = ""
   Me.vpDescripcion = ""
   Me.vpHaySeleccion = False
   Unload Me
End If
End Sub
