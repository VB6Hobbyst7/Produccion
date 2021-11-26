VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLogBSCatalogo 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6435
   ClientLeft      =   1500
   ClientTop       =   1875
   ClientWidth     =   7815
   Icon            =   "frmLogBSCatalogo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   7815
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   20
      Top             =   -60
      Width           =   7580
      Begin VB.TextBox txtExp 
         Height          =   315
         Left            =   900
         TabIndex        =   21
         Top             =   300
         Width           =   6435
      End
      Begin VB.Label Label5 
         Caption         =   "Buscar"
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
         Left            =   180
         TabIndex        =   22
         Top             =   360
         Width           =   615
      End
   End
   Begin MSComctlLib.TreeView tvwBS 
      Height          =   2595
      Left            =   105
      TabIndex        =   0
      Top             =   720
      Width           =   7605
      _ExtentX        =   13414
      _ExtentY        =   4577
      _Version        =   393217
      Indentation     =   0
      LineStyle       =   1
      Style           =   6
      FullRowSelect   =   -1  'True
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin VB.Frame fraBienes 
      BorderStyle     =   0  'None
      Height          =   2895
      Left            =   120
      TabIndex        =   1
      Top             =   3420
      Width           =   7575
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "Nuevo"
         Height          =   375
         Left            =   6420
         TabIndex        =   5
         Top             =   0
         Width           =   1155
      End
      Begin VB.CommandButton cmdQuitar 
         Caption         =   "Quitar"
         Height          =   375
         Left            =   6420
         TabIndex        =   4
         Top             =   840
         Width           =   1155
      End
      Begin VB.CommandButton cmdModificar 
         Caption         =   "Modificar"
         Height          =   375
         Left            =   6420
         TabIndex        =   3
         Top             =   420
         Width           =   1155
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFlex 
         Height          =   2715
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   6315
         _ExtentX        =   11139
         _ExtentY        =   4789
         _Version        =   393216
         BackColor       =   16777215
         Cols            =   7
         FixedCols       =   0
         BackColorSel    =   15988975
         ForeColorSel    =   128
         BackColorBkg    =   -2147483643
         GridColor       =   -2147483633
         FocusRect       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   7
      End
   End
   Begin VB.Frame fraReg 
      Caption         =   "Registro de Bienes / Servicios "
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
      Height          =   2895
      Left            =   120
      TabIndex        =   6
      Top             =   3420
      Visible         =   0   'False
      Width           =   7575
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   6300
         TabIndex        =   19
         Top             =   2460
         Width           =   1155
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "Grabar"
         Height          =   375
         Left            =   5040
         TabIndex        =   18
         Top             =   2460
         Width           =   1215
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   315
         Left            =   1500
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   360
         Width           =   5775
      End
      Begin VB.ComboBox cboGrupo 
         Height          =   315
         Left            =   1500
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   720
         Width           =   5775
      End
      Begin VB.Frame Frame2 
         Caption         =   "Código CONSUCODE"
         ForeColor       =   &H00C00000&
         Height          =   1095
         Left            =   120
         TabIndex        =   7
         Top             =   1260
         Width           =   7335
         Begin VB.CommandButton cmdBusq 
            Caption         =   "Buscar"
            Height          =   330
            Left            =   8220
            TabIndex        =   11
            Top             =   780
            Width           =   900
         End
         Begin VB.TextBox txtCatalogo 
            Height          =   315
            Left            =   1380
            Locked          =   -1  'True
            TabIndex        =   10
            Top             =   660
            Width           =   5775
         End
         Begin VB.TextBox txtCIIU 
            Height          =   315
            Left            =   1380
            Locked          =   -1  'True
            TabIndex        =   9
            Top             =   300
            Width           =   5775
         End
         Begin VB.CommandButton cmdCiiu 
            Caption         =   "Buscar"
            Height          =   330
            Left            =   8220
            TabIndex        =   8
            Top             =   360
            Width           =   900
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Cód. Catálogo"
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   720
            Width           =   1005
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Cód. C.I.I.U. "
            Height          =   195
            Left            =   120
            TabIndex        =   12
            Top             =   360
            Width           =   915
         End
      End
      Begin VB.TextBox txtProSelBSCod 
         Height          =   315
         Left            =   1500
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   360
         Visible         =   0   'False
         Width           =   1155
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
         Left            =   180
         TabIndex        =   17
         Top             =   420
         Width           =   1020
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Grupo de Bienes"
         Height          =   195
         Left            =   180
         TabIndex        =   15
         Top             =   780
         Width           =   1185
      End
   End
   Begin VB.Menu mnuBusq 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuBusqSgte 
         Caption         =   "Buscar siguiente"
      End
   End
End
Attribute VB_Name = "frmLogBSCatalogo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sSQL As String, nFuncion As Integer
Dim nIniBusq As Integer


Private Sub cmdCancelar_Click()
fraReg.Visible = False
fraBienes.Visible = True
End Sub

Private Sub cmdGrabar_Click()
Dim oConn As New DConecta
Dim cBSCod As String
Dim cBSGrupoCod As String

cBSGrupoCod = Format(cboGrupo.ItemData(cboGrupo.ListIndex), "0000")
If MsgBox("¿ Esta seguro de grabar los datos ?" + Space(10), vbQuestion + vbYesNo, "Confirme") = vbYes Then
   Select Case nFuncion
       Case 1
       
       Case 2
            sSQL = "UPDATE LogProSelBienesServicios SET " & _
                   "       cBSDescripcion = '" & txtDescripcion.Text & "', " & _
                   "       cBSGrupoCod    = '" & cBSGrupoCod & "' " & _
                   " WHERE cProSelBSCod = '" & txtProSelBSCod.Text & "' "
   End Select
   If oConn.AbreConexion Then
      oConn.Ejecutar sSQL
   End If
   nFuncion = 0
End If
End Sub

Private Sub cmdModificar_Click()
nFuncion = 2
fraBienes.Visible = False
fraReg.Visible = True
txtProSelBSCod.Text = MSFlex.TextMatrix(MSFlex.row, 1)
txtDescripcion.Text = MSFlex.TextMatrix(MSFlex.row, 2)
End Sub

Private Sub cmdNuevo_Click()
nFuncion = 1
fraBienes.Visible = False
fraReg.Visible = True
txtProSelBSCod.Text = ""
txtDescripcion.Text = ""
End Sub

Private Sub Form_Load()
CentraForm Me
ListaBienes
End Sub

Sub ListaBienes()
Dim rs As New ADODB.Recordset
Dim n As Integer, K As Integer, cIMG As String
Dim oConn As DConecta, cProSelBSCod As String
Dim sSQL As String, cKey As String, cKeySup As String

On Error GoTo Salida

tvwBS.Nodes.Clear
tvwBS.Nodes.Add , , "K", "CMAC TRUJILLO"

sSQL = "select cProSelBSCod,cBSDescripcion from LogProSelBienesServicios  " & _
       " where bVigente=1 and len(cProSelBSCod)>=2 and len(cProSelBSCod)<=8 order by cProSelBSCod "

Set oConn = New DConecta
If oConn.AbreConexion Then
   Set rs = oConn.CargaRecordSet(sSQL)
   oConn.CierraConexion
   If Not rs.EOF Then
      Do While Not rs.EOF
         cKey = "K" + rs!cProSelBSCod
         
         Select Case Len(rs!cProSelBSCod)
             Case 2
                  cKeySup = "K"
             Case 3
                  cKeySup = "K" + Left(rs!cProSelBSCod, 2)
             Case 5
                  cKeySup = "K" + Mid(rs!cProSelBSCod, 1, Len(rs!cProSelBSCod) - 2)
             Case 8
                  cKeySup = "K" + Mid(rs!cProSelBSCod, 1, Len(rs!cProSelBSCod) - 3)
         End Select
         
         If Mid(cKey, 3, 1) <> "2" Then
            tvwBS.Nodes.Add cKeySup, tvwChild, cKey, rs!cBSDescripcion
         Else
            If Len(cKey) > 3 Then
               cKeySup = Left(cKey, 3)
               tvwBS.Nodes.Add cKeySup, tvwChild, cKey, rs!cBSDescripcion
            Else
               tvwBS.Nodes.Add cKeySup, tvwChild, cKey, rs!cBSDescripcion
            End If
         End If
         
         tvwBS.Nodes(tvwBS.Nodes.Count).Tag = rs!cProSelBSCod
         rs.MoveNext
      Loop
   End If
   oConn.CierraConexion
   '-----------------------------------------------------------------------
   cboGrupo.Clear
   sSQL = "select * from BSGrupos where len(cBSGrupoCod)=4 order by cBSGrupoCod"
   If oConn.AbreConexion Then
      Set rs = oConn.CargaRecordSet(sSQL)
      oConn.CierraConexion
      If Not rs.EOF Then
         Do While Not rs.EOF
            cboGrupo.AddItem rs!cBSGrupoDescripcion
            cboGrupo.ItemData(cboGrupo.ListCount - 1) = rs!cBSGrupoCod
            rs.MoveNext
         Loop
      End If
   End If
   '-----------------------------------------------------------------------
   
End If
Exit Sub
Salida:
End Sub



Private Sub tvwBS_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
   PopupMenu mnuBusq
End If
End Sub

Private Sub tvwBS_NodeClick(ByVal Node As MSComctlLib.Node)
If Len(Trim(Node.Tag)) = 8 Then
   GeneraListaBienes Node.Tag
Else
   LimpiaFlex
End If
End Sub

Sub GeneraListaBienes(ByVal psBSCod As String)
Dim oConn As New DConecta
Dim rs As New ADODB.Recordset
Dim i As Integer

LimpiaFlex

If Len(Trim(psBSCod)) < 8 Then Exit Sub

'sSQL = "select cProSelBSCod, cBSDescripcion  from LogProSelBienesServicios " & _
'       " where cProSelBSCod like '" & psBSCod & "%' and len(rtrim(cProSelBSCod))=10 "
       
sSQL = "select s.cProSelBSCod,s.cBSDescripcion, cGrupo=coalesce(cBSGrupoDescripcion,''), " & _
       " s.cCIIUCod,s.cCatalogoCod " & _
       " from LogProSelBienesServicios s left join BSGrupos b on s.cBSGrupoCod = b.cBSGrupoCod " & _
       " where s.cProSelBSCod like '" & psBSCod & "%' and len(rtrim(s.cProSelBSCod))=10 "


If oConn.AbreConexion Then
   Set rs = oConn.CargaRecordSet(sSQL)
   oConn.CierraConexion
   If Not rs.EOF Then
      Do While Not rs.EOF
         i = i + 1
         InsRow MSFlex, i
         MSFlex.TextMatrix(i, 1) = rs!cProSelBSCod
         MSFlex.TextMatrix(i, 2) = rs!cBSDescripcion
         MSFlex.TextMatrix(i, 3) = rs!cGrupo
         MSFlex.TextMatrix(i, 4) = rs!cCIIUCod
         MSFlex.TextMatrix(i, 5) = rs!cCatalogoCod
         rs.MoveNext
      Loop
   End If
End If
End Sub

Sub LimpiaFlex()
MSFlex.Clear
MSFlex.Rows = 2
MSFlex.RowHeight(1) = 8
MSFlex.ColWidth(0) = 0
MSFlex.ColWidth(1) = 1000
MSFlex.ColWidth(2) = 4000
MSFlex.ColWidth(3) = 0
MSFlex.ColWidth(4) = 0
End Sub

Private Sub txtExp_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   nIniBusq = 1
   BuscaExpresion
End If
End Sub

Sub BuscaExpresion()
Dim i As Integer, n As Integer
txtExp.Text = UCase(txtExp.Text)
n = tvwBS.Nodes.Count
For i = nIniBusq To n
    If InStr(tvwBS.Nodes(i).Text, txtExp.Text) > 0 Then
       tvwBS.Nodes(i).Parent.Parent.Parent.Expanded = True
       tvwBS.Nodes(i).Parent.Parent.Expanded = True
       tvwBS.Nodes(i).Parent.Expanded = True
       tvwBS.Nodes(i).Expanded = True
       tvwBS.Nodes(i).Selected = True
       tvwBS.SetFocus
       nIniBusq = i + 1
       Exit Sub
    End If
Next
MsgBox "No se halla la expresión..." + Space(10), vbInformation
End Sub

