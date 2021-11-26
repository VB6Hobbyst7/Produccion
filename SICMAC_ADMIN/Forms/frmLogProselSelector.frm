VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmLogProSelSelector 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4035
   ClientLeft      =   1845
   ClientTop       =   3150
   ClientWidth     =   7530
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   7530
   Begin VB.Frame fraSel2 
      BorderStyle     =   0  'None
      Height          =   3795
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   7275
      Begin VB.TextBox txtExp 
         Height          =   315
         Left            =   0
         TabIndex        =   6
         Top             =   240
         Width           =   7275
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSLista 
         Height          =   3135
         Left            =   0
         TabIndex        =   4
         Top             =   600
         Width           =   7275
         _ExtentX        =   12832
         _ExtentY        =   5530
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         BackColorBkg    =   -2147483643
         GridColor       =   -2147483633
         FocusRect       =   0
         HighLight       =   2
         SelectionMode   =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   4
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Expresión a buscar:"
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
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   1695
      End
   End
   Begin VB.Frame fraSel1 
      BorderStyle     =   0  'None
      Height          =   3795
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7275
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFlex 
         Height          =   3435
         Left            =   0
         TabIndex        =   1
         Top             =   300
         Width           =   7245
         _ExtentX        =   12779
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
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   1140
      End
   End
End
Attribute VB_Name = "frmLogProSelSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public vpHaySeleccion As Boolean
Public vpCodigo As String
Public vpDescripcion As String

Dim cSQL As String, cTitulo As String, bConBuscador As Boolean, cCampoBusqueda As String

Public Sub Consulta(vSQL As String, vTitulo As String, Optional ByVal vConBuscador As Boolean = False, Optional ByVal vCampoBusqueda As String = "")
cSQL = vSQL
cTitulo = vTitulo
bConBuscador = vConBuscador
cCampoBusqueda = vCampoBusqueda
Me.Show 1
End Sub

Private Sub Form_Load()
CentraForm Me
Me.vpHaySeleccion = False
Me.vpCodigo = ""
Me.vpDescripcion = ""
If Not bConBuscador Then
   fraSel1.Visible = True
   fraSel2.Visible = False
   lblTitulo = cTitulo
   ListaDatos
Else
   LimpiaFlex
   fraSel1.Visible = False
   fraSel2.Visible = True
   If Len(Trim(cCampoBusqueda)) = 0 Then
      MsgBox "Debe indicar el campo de búsqueda..." + Space(10), vbInformation
      txtExp.Enabled = False
   Else
      txtExp.TabIndex = 0
   End If
End If
End Sub

Sub ListaDatos()
Dim Rs As New ADODB.Recordset
Dim oConn As DConecta, N As Integer
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
   Set Rs = oConn.CargaRecordSet(cSQL)
   If Not Rs.EOF Then
      i = 0
      N = Rs.Fields.Count
      If N = 3 Then
         MSFlex.ColWidth(1) = 5500
         MSFlex.ColWidth(2) = 1000
      End If
      If N = 4 Then
         MSFlex.ColWidth(1) = 3500
         MSFlex.ColWidth(2) = 1500
         MSFlex.ColWidth(3) = 1500
      End If
      Do While Not Rs.EOF
         i = i + 1
         InsRow MSFlex, i
         MSFlex.TextMatrix(i, 0) = Rs(0)
         MSFlex.TextMatrix(i, 1) = Rs(1)
         If N = 3 Then
            MSFlex.TextMatrix(i, 2) = Rs(2)
         End If
         If N = 4 Then
            MSFlex.TextMatrix(i, 2) = Rs(2)
            MSFlex.TextMatrix(i, 3) = Rs(3)
         End If
         Rs.MoveNext
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
   i = MSFlex.Row
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

Private Sub txtExp_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
   If Len(Trim(txtExp)) > 0 Then
      GeneraListaExp txtExp
   End If
End If
End Sub

Sub LimpiaFlex()
MSLista.Clear
MSLista.Rows = 2
MSLista.RowHeight(1) = 8
MSLista.ColWidth(0) = 0
MSLista.ColWidth(1) = 1100
MSLista.ColWidth(2) = 5800
MSLista.ColWidth(3) = 0
End Sub

Sub GeneraListaExp(psExp As String)
Dim Rs As New ADODB.Recordset
Dim oConn As New DConecta
Dim i As Integer, sSQL As String
Dim Kini As Integer, Kfin As Integer
Dim cResto As String
Dim Kwhere As Integer
Dim Korder As Integer

LimpiaFlex

Kini = 0
Kfin = 0
Kwhere = 0
Korder = 0

Kini = InStr(UCase(cSQL), "FROM") + 4

sSQL = ""
cResto = ""
sSQL = Left(cSQL, Kini)
cResto = Right(cSQL, Len(Trim(cSQL)) - Kini + 1)

Kwhere = InStr(UCase(cSQL), "WHERE")
If Kwhere = 0 Then
   cSQL = cSQL + " WHERE " + cCampoBusqueda + " LIKE '" + txtExp + "%' "
Else
   cSQL = Left(cSQL, Kwhere - 1) + " WHERE " + cCampoBusqueda + " LIKE '" + txtExp + "%' "
End If
Korder = InStr(UCase(cSQL), "ORDER")

If oConn.AbreConexion Then
   Set Rs = oConn.CargaRecordSet(cSQL)
   If Not Rs.EOF Then
      i = 0
      Do While Not Rs.EOF
         i = i + 1
         InsRow MSLista, i
         MSLista.TextMatrix(i, 1) = Rs(0)
         MSLista.TextMatrix(i, 2) = Rs(1)
         Rs.MoveNext
      Loop
      MSLista.SetFocus
   Else
      MsgBox "No se halla la expresión..." + Space(10), vbInformation
      txtExp.SetFocus
   End If
End If
End Sub

Private Sub MSLista_KeyPress(KeyAscii As Integer)
Dim i As Integer
If KeyAscii = 13 Then
   i = MSLista.Row
   Me.vpHaySeleccion = False
   If Len(Trim(MSLista.TextMatrix(i, 1))) > 0 Then
      Me.vpCodigo = MSLista.TextMatrix(i, 1)
      Me.vpDescripcion = MSLista.TextMatrix(i, 2)
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

