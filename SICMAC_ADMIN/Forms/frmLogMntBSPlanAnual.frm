VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmLogMntBSPlanAnual 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6420
   ClientLeft      =   1695
   ClientTop       =   1875
   ClientWidth     =   7755
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   7755
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   380
      Left            =   6360
      TabIndex        =   9
      Top             =   5800
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Cargos de Personal "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2355
      Left            =   60
      TabIndex        =   5
      Top             =   60
      Width           =   7635
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFlex 
         Height          =   1995
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   7395
         _ExtentX        =   13044
         _ExtentY        =   3519
         _Version        =   393216
         Cols            =   7
         FixedCols       =   0
         BackColorBkg    =   -2147483643
         GridColor       =   -2147483633
         FocusRect       =   0
         HighLight       =   2
         SelectionMode   =   1
         AllowUserResizing=   1
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
         _Band(0).Cols   =   7
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Bienes y Servicios "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   3795
      Left            =   60
      TabIndex        =   1
      Top             =   2520
      Width           =   7635
      Begin VB.TextBox txtExp 
         Height          =   315
         Left            =   900
         TabIndex        =   7
         Top             =   300
         Width           =   5295
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "Buscar"
         Height          =   380
         Left            =   6300
         TabIndex        =   6
         Top             =   300
         Width           =   1215
      End
      Begin VB.CommandButton cmdQuitar 
         Caption         =   "Quitar"
         Height          =   380
         Left            =   6300
         TabIndex        =   4
         Top             =   1140
         Width           =   1215
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "Agregar"
         Height          =   380
         Left            =   6300
         TabIndex        =   3
         Top             =   720
         Width           =   1215
      End
      Begin MSComctlLib.TreeView tvwObj 
         Height          =   3015
         Left            =   120
         TabIndex        =   2
         Top             =   660
         Width           =   6090
         _ExtentX        =   10742
         _ExtentY        =   5318
         _Version        =   393217
         Indentation     =   0
         LineStyle       =   1
         Style           =   6
         FullRowSelect   =   -1  'True
         ImageList       =   "ImageList1"
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
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
         TabIndex        =   8
         Top             =   360
         Width           =   615
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   540
      Top             =   4740
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogMntBSPlanAnual.frx":0000
            Key             =   "logo"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogMntBSPlanAnual.frx":0352
            Key             =   "check0"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogMntBSPlanAnual.frx":06A4
            Key             =   "check1"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogMntBSPlanAnual.frx":09F6
            Key             =   "enlace"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuBusq 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuCopiar 
         Caption         =   "Copiar Estructura"
      End
      Begin VB.Menu mnuPegar 
         Caption         =   "Pegar Estructura"
      End
      Begin VB.Menu mnuBusqSgte 
         Caption         =   "Buscar siguiente"
      End
   End
End
Attribute VB_Name = "frmLogMntBSPlanAnual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nIniBusq As Integer
Dim sSQL As String
Dim cRHCargoOrig
Dim cRHCargoDest

Private Sub cmdAgregar_Click()
Dim cProSelBSCod As String
Dim cBSCod8 As String
Dim cBSCod5 As String
Dim cBSCod3 As String, cBSCod2 As String, cBSCod1 As String
Dim cRHCargoCod As String
Dim rs As New ADODB.Recordset
Dim rt As New ADODB.Recordset
Dim oConn As New DConecta
Dim cnt As Integer

cRHCargoCod = MSFlex.TextMatrix(MSFlex.row, 1)
frmLogProSelBSSelector.TodosConCheck False
If frmLogProSelBSSelector.vpSeleccion Then
   Set rs = Nothing
   Set rs = frmLogProSelBSSelector.gvrs
   If rs.State <> 0 Then
      If Not (rs.EOF And rs.BOF) Then
         If oConn.AbreConexion Then
            rs.MoveFirst
            Do While Not rs.EOF
               cBSCod1 = Mid(rs!cProSelBSCod, 1, 1)
               cBSCod2 = Mid(rs!cProSelBSCod, 1, 2)
               cBSCod3 = Mid(rs!cProSelBSCod, 1, 3)
               cBSCod5 = Mid(rs!cProSelBSCod, 1, 5)
               cBSCod8 = Mid(rs!cProSelBSCod, 1, 8)
               cProSelBSCod = rs!cProSelBSCod
               
               sSQL = "Select * from LogProSelBSCargos where cRHCargoCod = '" & cRHCargoCod & "' and cProSelBSCod in ('" & cBSCod1 & "','" & cBSCod2 & "','" & cBSCod3 & "','" & cBSCod5 & "','" & cBSCod8 & "','" & cProSelBSCod & "') "
               
               Set rt = oConn.CargaRecordSet(sSQL)
               If rt.EOF Then
                  sSQL = "INSERT INTO LogProSelBSCargos (cRHCargoCod,cProSelBSCod) values ('" & cRHCargoCod & "','" & cBSCod1 & "') "
                  oConn.Ejecutar sSQL

                  sSQL = "INSERT INTO LogProSelBSCargos (cRHCargoCod,cProSelBSCod) values ('" & cRHCargoCod & "','" & cBSCod2 & "') "
                  oConn.Ejecutar sSQL

                  sSQL = "INSERT INTO LogProSelBSCargos (cRHCargoCod,cProSelBSCod) values ('" & cRHCargoCod & "','" & cBSCod3 & "') "
                  oConn.Ejecutar sSQL

                  sSQL = "INSERT INTO LogProSelBSCargos (cRHCargoCod,cProSelBSCod) values ('" & cRHCargoCod & "','" & cBSCod5 & "') "
                  oConn.Ejecutar sSQL

                  sSQL = "INSERT INTO LogProSelBSCargos (cRHCargoCod,cProSelBSCod) values ('" & cRHCargoCod & "','" & cBSCod8 & "') "
                  oConn.Ejecutar sSQL
               Else
                  cnt = rt.RecordCount
                  
                  If cnt = 1 Then
                     sSQL = "INSERT INTO LogProSelBSCargos (cRHCargoCod,cProSelBSCod) values ('" & cRHCargoCod & "','" & cBSCod2 & "') "
                     oConn.Ejecutar sSQL

                     sSQL = "INSERT INTO LogProSelBSCargos (cRHCargoCod,cProSelBSCod) values ('" & cRHCargoCod & "','" & cBSCod3 & "') "
                     oConn.Ejecutar sSQL

                     sSQL = "INSERT INTO LogProSelBSCargos (cRHCargoCod,cProSelBSCod) values ('" & cRHCargoCod & "','" & cBSCod5 & "') "
                     oConn.Ejecutar sSQL

                     sSQL = "INSERT INTO LogProSelBSCargos (cRHCargoCod,cProSelBSCod) values ('" & cRHCargoCod & "','" & cBSCod8 & "') "
                     oConn.Ejecutar sSQL
                  End If
                  
                  
                  If cnt = 2 Then
                     sSQL = "INSERT INTO LogProSelBSCargos (cRHCargoCod,cProSelBSCod) values ('" & cRHCargoCod & "','" & cBSCod3 & "') "
                     oConn.Ejecutar sSQL

                     sSQL = "INSERT INTO LogProSelBSCargos (cRHCargoCod,cProSelBSCod) values ('" & cRHCargoCod & "','" & cBSCod5 & "') "
                     oConn.Ejecutar sSQL

                     sSQL = "INSERT INTO LogProSelBSCargos (cRHCargoCod,cProSelBSCod) values ('" & cRHCargoCod & "','" & cBSCod8 & "') "
                     oConn.Ejecutar sSQL
                  End If
                  
                  If cnt = 3 Then
                     sSQL = "INSERT INTO LogProSelBSCargos (cRHCargoCod,cProSelBSCod) values ('" & cRHCargoCod & "','" & cBSCod5 & "') "
                     oConn.Ejecutar sSQL

                     sSQL = "INSERT INTO LogProSelBSCargos (cRHCargoCod,cProSelBSCod) values ('" & cRHCargoCod & "','" & cBSCod8 & "') "
                     oConn.Ejecutar sSQL
                  End If
                  
                  If cnt = 4 Then
                     sSQL = "INSERT INTO LogProSelBSCargos (cRHCargoCod,cProSelBSCod) values ('" & cRHCargoCod & "','" & cBSCod8 & "') "
                     oConn.Ejecutar sSQL
                  End If
               End If
               
               sSQL = "INSERT INTO LogProSelBSCargos (cRHCargoCod,cProSelBSCod) values ('" & cRHCargoCod & "','" & cProSelBSCod & "') "
               oConn.Ejecutar sSQL
               
               rs.MoveNext
            Loop
         End If
      End If
   End If
   ListaBienesCargo cRHCargoCod
Else
   Set rs = Nothing
End If
End Sub

Private Sub cmdBuscar_Click()
txtExp_KeyPress 13
End Sub

Private Sub cmdQuitar_Click()
Dim cObjCod As String
Dim cRHCargoCod As String
Dim oConn As New DConecta

cRHCargoCod = MSFlex.TextMatrix(MSFlex.row, 1)

If tvwObj.Nodes.Count <= 1 Then Exit Sub
If MsgBox("¿ Está seguro de la estructura indicada ?" + Space(10), vbQuestion + vbYesNo, "Confirme") = vbYes Then

   cObjCod = tvwObj.Nodes(tvwObj.SelectedItem.Index).Tag
   sSQL = "DELETE FROM LogProSelBSCargos WHERE cRHCargoCod = '" & cRHCargoCod & "' and cProSelBSCod like '" & cObjCod & "%' "
   If oConn.AbreConexion Then
      oConn.Ejecutar sSQL
   End If
   'tvwObj.Nodes.Remove tvwObj.SelectedItem.Index
End If
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
CentraForm Me
ListaCargos
End Sub

Sub ListaCargos()
Dim i As Integer
Dim rs As New ADODB.Recordset
Dim oConn As New DConecta

i = 0
FormaFlexBS

sSQL = "select cRHCargoCod,cRHCargoDescripcion  from RHCargosTabla " & _
       " Where bRHCargoEstado = 1 And Len(RTrim(cRHCargoCod)) > 4 " & _
       " order by cRHCargoCod "
 
If oConn.AbreConexion Then
   Set rs = oConn.CargaRecordSet(sSQL)
   If Not rs.EOF Then
      Do While Not rs.EOF
         i = i + 1
         InsRow MSFlex, i
         MSFlex.TextMatrix(i, 1) = rs!cRHCargoCod
         MSFlex.TextMatrix(i, 2) = rs!cRHCargoDescripcion
         rs.MoveNext
      Loop
   End If
End If
End Sub

Sub FormaFlexBS()
MSFlex.Clear
MSFlex.Rows = 2
MSFlex.RowHeight(1) = 8
MSFlex.ColWidth(0) = 0
MSFlex.ColWidth(1) = 700:  MSFlex.TextMatrix(0, 1) = "Código": MSFlex.ColAlignment(1) = 4
MSFlex.ColWidth(2) = 6000: MSFlex.TextMatrix(0, 2) = "Cargo de Personal"
MSFlex.ColWidth(3) = 0
MSFlex.ColWidth(4) = 0
MSFlex.ColWidth(5) = 0
MSFlex.ColWidth(6) = 0
End Sub

Private Sub mnuBusqSgte_Click()
BuscaExpresion
End Sub

Private Sub mnuCopiar_Click()
cRHCargoOrig = MSFlex.TextMatrix(MSFlex.row, 1)
mnuCopiar.Enabled = False
End Sub

Private Sub mnuPegar_Click()
Dim oConn As New DConecta

cRHCargoDest = MSFlex.TextMatrix(MSFlex.row, 1)
mnuCopiar.Enabled = True

If Len(cRHCargoOrig) = 0 Then
   MsgBox "Falta espeficicar origen de datos..." + Space(10), vbInformation, "Aviso"
   Exit Sub
End If

If Len(cRHCargoDest) = 0 Then
   MsgBox "Falta espeficicar destino de datos..." + Space(10), vbInformation, "Aviso"
   Exit Sub
End If

If oConn.AbreConexion Then
   sSQL = "DELETE from LogProSelBSCargos where cRHCargoCod = '" & cRHCargoDest & "'"
   oConn.Ejecutar sSQL

   sSQL = "insert into LogProSelBSCargos (cRHCargoCod,cProSelBSCod) " & _
         " Select '" & cRHCargoDest & "',cProSelBSCod from LogProSelBSCargos where cRHCargoCod = '" & cRHCargoOrig & "'"
       

   oConn.Ejecutar sSQL
   MsgBox "Se ha copiado los datos correctamente..." + Space(10), vbInformation
End If
cRHCargoOrig = ""
cRHCargoDest = ""
End Sub

Private Sub MSFlex_GotFocus()
If Len(MSFlex.TextMatrix(MSFlex.row, 1)) > 0 Then
   ListaBienesCargo MSFlex.TextMatrix(MSFlex.row, 1)
End If
End Sub

Private Sub MSFlex_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
   mnuCopiar.Visible = True
   mnuPegar.Visible = True
   mnuBusqSgte.Visible = False
   PopupMenu mnuBusq
End If
End Sub

Private Sub MSFlex_RowColChange()
If Len(MSFlex.TextMatrix(MSFlex.row, 1)) > 0 Then
   ListaBienesCargo MSFlex.TextMatrix(MSFlex.row, 1)
End If
End Sub

Sub ListaBienesCargo(psRHCargoCod As String)
Dim rs As New ADODB.Recordset
Dim n As Integer, K As Integer, cIMG As String
Dim oConn As DConecta, cProSelBSCod As String
Dim sSQL As String
Dim cKey As String, cKeySup As String
Dim cColor As String

On Error GoTo Salida
Set tvwObj.ImageList = ImageList1
tvwObj.Nodes.Clear
tvwObj.Nodes.Add , , "K", "CMAC TRUJILLO" ', "logo"

'sSQL = "select cProSelBSCod,cBSDescripcion,bVigente from LogProSelBienesServicios  " & _
'       " where len(cProSelBSCod)>=2 order by cProSelBSCod "

sSQL = "Select x.cProSelBSCod,y.cBSDescripcion from LogProSelBSCargos x  " & _
       " inner join LogProSelBienesServicios y on x.cProSelBSCod = y.cProSelBSCod " & _
       " where x.cRHCargoCod = '" & psRHCargoCod & "' and len(x.cProSelBSCod)>=2 order by x.cProSelBSCod"

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
             Case 8
                  cKeySup = "K" + Mid(rs!cProSelBSCod, 1, Len(rs!cProSelBSCod) - 3)
             Case 5, 10
                  cKeySup = "K" + Mid(rs!cProSelBSCod, 1, Len(rs!cProSelBSCod) - 2)
         End Select
         'If Rs!bVigente = 0 Then
         'Else
         'End If
         If Mid(cKey, 3, 1) <> "2" Then
            If Len(rs!cProSelBSCod) = 10 Then
               tvwObj.Nodes.Add cKeySup, tvwChild, cKey, UCase(rs!cBSDescripcion) ', "check1"
               tvwObj.Nodes(tvwObj.Nodes.Count).ForeColor = "&H00C00000"
            Else
               tvwObj.Nodes.Add cKeySup, tvwChild, cKey, rs!cBSDescripcion ', "enlace"
            End If
         Else
            If Len(cKey) > 3 Then
               cKeySup = Left(cKey, 3)
               tvwObj.Nodes.Add cKeySup, tvwChild, cKey, rs!cBSDescripcion ', "check0"
            Else
               tvwObj.Nodes.Add cKeySup, tvwChild, cKey, rs!cBSDescripcion ', "enlace"
            End If
         End If
         tvwObj.Nodes(tvwObj.Nodes.Count).Tag = rs!cProSelBSCod
         'If Len(rs!cProSelBSCod) = 10 Then
         '   tvwObj.Nodes(tvwObj.Nodes.Count).ForeColor = "&H00C00000"
         'End If
         rs.MoveNext
      Loop
      tvwObj.Nodes(1).Expanded = True
   End If
End If
Exit Sub
Salida:
End Sub

Private Sub tvwObj_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
   mnuCopiar.Visible = False
   mnuPegar.Visible = False
   mnuBusqSgte.Visible = True
   PopupMenu mnuBusq
End If
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

If Len(Trim(txtExp.Text)) = 0 Then Exit Sub

n = tvwObj.Nodes.Count
For i = nIniBusq To n
    If InStr(UCase(tvwObj.Nodes(i).Text), txtExp.Text) > 0 Then
       tvwObj.Nodes(i).Parent.Parent.Parent.Expanded = True
       tvwObj.Nodes(i).Parent.Parent.Expanded = True
       tvwObj.Nodes(i).Parent.Expanded = True
       tvwObj.Nodes(i).Expanded = True
       tvwObj.Nodes(i).Selected = True
       tvwObj.SetFocus
       nIniBusq = i + 1
       Exit Sub
    End If
Next
MsgBox "No se halla la expresión..." + Space(10), vbInformation
End Sub

'Private Sub tvwObj_NodeCheck(ByVal Node As MSComctlLib.Node)
'Dim i As Integer, n As Integer, X As Boolean
'Dim oConn As New DConecta
'Dim nValor As Integer
'Dim cBSCod As String
'Dim sSQL As String
'
'X = Node.Checked
'n = Node.Children
'i = Node.Index
'
'If X Then
'   nValor = 1
'Else
'   nValor = 0
'End If
'
'If oConn.AbreConexion Then
'   If n > 0 Then
'      For i = Node.Index + 1 To Node.Index + n
'          tvwObj.Nodes(i).Checked = X
'          If Len(tvwObj.Nodes(i).Tag) >= 10 Then
'             cBSCod = tvwObj.Nodes(i).Tag
'             sSQL = "UPDATE LogProSelBienesServicios SET bVigente = " & nValor & " WHERE cProSelBSCod = '" & cBSCod & "'"
'             oConn.ConexionActiva.Execute sSQL
'          End If
'      Next
'   End If
'End If
'End Sub


Private Sub tvwObj_DblClick()
Dim i As Integer, n As Integer
Dim oConn As New DConecta
Dim cBSCod As String
Dim sSQL As String
Dim nCheck As Integer

i = tvwObj.SelectedItem.Index
If Len(tvwObj.Nodes(i).Tag) >= 10 Then
   n = InStr(tvwObj.Nodes(i).Tag, "*")
   cBSCod = Replace(tvwObj.Nodes(i).Tag, "*", "")
   If n > 0 Then
      tvwObj.Nodes(i).Image = "check0"
      'tvwObj.Nodes(i).ForeColor = "&H80000008"
      tvwObj.Nodes(i).Tag = Mid(tvwObj.Nodes(i).Tag, 1, n - 1)
      nCheck = 0
   Else
      tvwObj.Nodes(i).Image = "check1"
      'tvwObj.Nodes(i).ForeColor = "&H000000C0"
      tvwObj.Nodes(i).Tag = tvwObj.Nodes(i).Tag + "*"
      nCheck = 1
  End If
  
  sSQL = "update LogProSelBienesServicios set bVigente = " & nCheck & " where cProSelBSCod='" & cBSCod & "'"
  If oConn.AbreConexion Then
     oConn.ConexionActiva.Execute sSQL
  End If
  oConn.CierraConexion
End If
End Sub

