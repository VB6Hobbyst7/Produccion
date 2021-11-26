VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmLogProSelBSSelector 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Selector de Bienes y Servicios"
   ClientHeight    =   6315
   ClientLeft      =   2535
   ClientTop       =   1890
   ClientWidth     =   6390
   Icon            =   "frmLogProSelBSSelector.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   6390
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   -60
      Width           =   6120
      Begin VB.TextBox txtExp 
         Height          =   315
         Left            =   900
         TabIndex        =   4
         Top             =   300
         Width           =   4935
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
         TabIndex        =   5
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   380
      Left            =   5040
      TabIndex        =   2
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Seleccionar"
      Height          =   380
      Left            =   3360
      TabIndex        =   1
      Top             =   5880
      Width           =   1575
   End
   Begin MSComctlLib.TreeView tvwBS 
      Height          =   5115
      Left            =   120
      TabIndex        =   0
      Top             =   660
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   9022
      _Version        =   393217
      Indentation     =   0
      LineStyle       =   1
      Style           =   7
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList imgLista 
      Left            =   180
      Top             =   5040
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
            Picture         =   "frmLogProSelBSSelector.frx":08CA
            Key             =   "logo"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogProSelBSSelector.frx":0C1C
            Key             =   "check0"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogProSelBSSelector.frx":0F6E
            Key             =   "check1"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogProSelBSSelector.frx":12C0
            Key             =   "enlace"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuBusqueda 
      Caption         =   "Busqueda"
      Visible         =   0   'False
      Begin VB.Menu mnuBuscar 
         Caption         =   "&Buscar siguiente"
      End
   End
End
Attribute VB_Name = "frmLogProSelBSSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public vpSeleccion As Boolean
Public gvrs As ADODB.Recordset

Dim nIniBusq As Integer
Dim sSQL As String, ValCheck As Boolean
Dim cKey As String, cKeySup As String
Dim XFlag As Boolean, XTodosCheck As Boolean, XProceso As Boolean, XSelMultiple As Boolean

Public Sub TodosConCheck(Optional ByVal vSi As Boolean = False, Optional ByVal vProceso As Boolean = True, Optional ByVal vSelMultiple As Boolean = True)
XTodosCheck = vSi
XProceso = vProceso
XSelMultiple = vSelMultiple
Me.Show 1
End Sub

Private Sub Form_Load()
CentraForm Me
'---- AnimaForm Me, -1, -1, aload, 0, 3, 1, 25
tvwBS.ImageList = imgLista
vpSeleccion = False
If XTodosCheck Then
   tvwBS.CheckBoxes = True
   Set tvwBS.ImageList = Nothing
   CargaBienesConCheck
Else
   CargaBienesSinCheck
End If
If Not XSelMultiple Then
    tvwBS.CheckBoxes = False
    Exit Sub
End If
End Sub

Sub CargaBienesConCheck()
Dim Rs As New ADODB.Recordset
Dim oConn As DConecta

cKey = "K1"
tvwBS.Nodes.Clear
tvwBS.Nodes.Add , , cKey, "CAJA TRUJILLO"
tvwBS.Nodes(1).Tag = "K1"
cKeySup = "K1"

Set oConn = New DConecta
If oConn.AbreConexion Then
    If XProceso Then
        sSQL = "select cProSelBSCod,cBSDescripcion from LogProSelBienesServicios where len(cProSelBSCod)=2 and bVigente=1 "
    Else
        sSQL = "select cProSelBSCod,cBSDescripcion from BienesServicios where len(cProSelBSCod)=2 and bVigente=1 "
    End If
   Set Rs = oConn.CargaRecordSet(sSQL)
   oConn.CierraConexion
   If Not Rs.EOF Then
      Do While Not Rs.EOF
         cKey = "K" + Rs!cProSelBSCod
         tvwBS.Nodes.Add cKeySup, tvwChild, cKey, Rs!cBSDescripcion
         tvwBS.Nodes(tvwBS.Nodes.Count).Tag = Rs!cProSelBSCod
         Rs.MoveNext
      Loop
      tvwBS.Nodes(1).Expanded = True
   End If
End If
End Sub

Sub CargaBienesSinCheck()
Dim Rs As New ADODB.Recordset
Dim oConn As DConecta

cKey = "K1"
tvwBS.Nodes.Clear
tvwBS.Nodes.Add , , cKey, "CAJA TRUJILLO", "logo"
tvwBS.Nodes(1).Tag = "K1"
cKeySup = "K1"

Set oConn = New DConecta
If oConn.AbreConexion Then
    If XProceso Then
        sSQL = "select cProSelBSCod,cBSDescripcion from LogProSelBienesServicios where len(cProSelBSCod)=2 and bVigente=1 "
    Else
        sSQL = "select cProSelBSCod,cBSDescripcion from BienesServicios where len(cProSelBSCod)=2 and bVigente=1 "
    End If
   Set Rs = oConn.CargaRecordSet(sSQL)
   oConn.CierraConexion
   If Not Rs.EOF Then
      Do While Not Rs.EOF
         cKey = "K" + Rs!cProSelBSCod
         tvwBS.Nodes.Add cKeySup, tvwChild, cKey, Rs!cBSDescripcion, "enlace"
         tvwBS.Nodes(tvwBS.Nodes.Count).Tag = Rs!cProSelBSCod
         Rs.MoveNext
      Loop
      tvwBS.Nodes(1).Expanded = True
   End If
End If
End Sub

Private Sub tvwBS_DblClick()
Dim Rs As New ADODB.Recordset
Dim n As Integer, k As Integer, cIMG As String
Dim oConn As DConecta, cProSelBSCod As String


   cProSelBSCod = tvwBS.Nodes(tvwBS.SelectedItem.Index).Tag
   If InStr(cProSelBSCod, "*") > 0 Then Exit Sub

   n = Len(cProSelBSCod)
   Select Case n
    Case 2
         If cProSelBSCod = "11" Then
            sSQL = "select * from LogProSelBienesServicios where len(cProSelBSCod)=3 and bVigente=1 "
         Else
            sSQL = "select * from LogProSelBienesServicios where cProSelBSCod like '" & cProSelBSCod & "%' and len(cProSelBSCod)=10 and bVigente=1 "
         End If

    Case 3
         sSQL = "select * from LogProSelBienesServicios where len(cProSelBSCod)=5 and cProSelBSCod like '" & cProSelBSCod & "%' and bVigente=1 "
    Case 5
         sSQL = "select * from LogProSelBienesServicios where len(cProSelBSCod)=8 and cProSelBSCod like '" & cProSelBSCod & "%' and bVigente=1 "
    Case 8
         sSQL = "select * from LogProSelBienesServicios where len(cProSelBSCod)=10 and cProSelBSCod like '" & cProSelBSCod & "%' and bVigente=1 "
    Case Else
         Exit Sub
   End Select

   Set oConn = New DConecta
   If oConn.AbreConexion Then
      Set Rs = oConn.CargaRecordSet(sSQL)
      oConn.CierraConexion
      If Not Rs.EOF Then
         Do While Not Rs.EOF
            cKey = "K" + Rs!cProSelBSCod
            cKeySup = "K" + Mid(Rs!cProSelBSCod, 1, n)

            If Not XTodosCheck Then
               If Len(Rs!cProSelBSCod) = 10 Then
                  tvwBS.Nodes.Add cKeySup, tvwChild, cKey, Rs!cBSDescripcion, "check0"
               Else
                  tvwBS.Nodes.Add cKeySup, tvwChild, cKey, Rs!cBSDescripcion, "enlace"
               End If
            Else
               tvwBS.Nodes.Add cKeySup, tvwChild, cKey, Rs!cBSDescripcion ', "enlace"
            End If

            tvwBS.Nodes(tvwBS.Nodes.Count).Tag = Rs!cProSelBSCod
            Rs.MoveNext
         Loop
         tvwBS.Nodes(tvwBS.SelectedItem.Index).Expanded = True
         tvwBS.Nodes(tvwBS.SelectedItem.Index).Tag = tvwBS.Nodes(tvwBS.SelectedItem.Index).Tag + "*"
      End If
   End If

End Sub

Private Sub tvwBS_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cmdSelect_Click
End If
End Sub

Private Sub tvwBS_KeyDown(KeyCode As Integer, Shift As Integer)
XFlag = True
End Sub

Private Sub tvwBS_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
XFlag = False
If Button = 2 Then
   PopupMenu mnuBusqueda
End If
End Sub

Private Sub tvwBS_NodeClick(ByVal Node As MSComctlLib.Node)
Dim cCod As String, i As Integer
If Not XSelMultiple Then Exit Sub
If XTodosCheck Then Exit Sub
If XFlag Then Exit Sub
cCod = Trim(Node.Tag)
If Len(cCod) >= 10 Then
   i = InStr(cCod, "#")
   If i > 0 Then
      Node.Image = "check0"
      Node.ForeColor = "&H80000008"
      Node.Tag = Mid(Node.Tag, 1, i - 1)
   Else
      Node.Image = "check1"
      Node.ForeColor = "&H000000C0"
      Node.Tag = Node.Tag + "#"
  End If
End If
End Sub

Private Sub ExpandeNodo(ByVal NodeIndex As Integer)
Dim cCod As String, i As Integer

If XTodosCheck Then Exit Sub
If XFlag Then Exit Sub

cCod = Trim(tvwBS.Nodes(NodeIndex).Tag)
If Len(cCod) >= 10 Then
   i = InStr(cCod, "#")
   If i > 0 Then
      tvwBS.Nodes(NodeIndex).Image = "check0"
      tvwBS.Nodes(NodeIndex).ForeColor = "&H80000008"
      tvwBS.Nodes(NodeIndex).Tag = Mid(tvwBS.Nodes(NodeIndex).Tag, 1, i - 1)
   Else
      tvwBS.Nodes(NodeIndex).Image = "check1"
      tvwBS.Nodes(NodeIndex).ForeColor = "&H000000C0"
      tvwBS.Nodes(NodeIndex).Tag = tvwBS.Nodes(NodeIndex).Tag + "#"
  End If
End If
End Sub


Private Sub tvwBS_NodeCheck(ByVal Node As MSComctlLib.Node)
On Error GoTo tvwBS_NodeCheckErr
Dim papa As String, n As Integer, i As Integer

If Not XTodosCheck Then Exit Sub

tvwBS.SelectedItem = Node
tvwBS_DblClick

n = tvwBS.Nodes.Count
papa = Node.Key
For i = 1 To n
    If InStr(1, tvwBS.Nodes(i).Key, papa) > 0 Then
       tvwBS.Nodes(i).Checked = Node.Checked
    End If
Next
Exit Sub
tvwBS_NodeCheckErr:
    MsgBox Err.Number & vbCrLf & Err.Description
End Sub


Private Sub cmdSelect_Click()
Dim i As Integer, n As Integer
Dim k As Integer

Set gvrs = New ADODB.Recordset
gvrs.Fields.Append "cProSelBSCod", adVarChar, 12, adFldMayBeNull
gvrs.Fields.Append "cBSDescripcion", adVarChar, 120, adFldMayBeNull
gvrs.Open

n = tvwBS.Nodes.Count

If XTodosCheck Then
   For i = 1 To n
       If tvwBS.Nodes(i).Checked And Len(tvwBS.Nodes(i).Tag) >= 10 Then
          k = InStr(tvwBS.Nodes(i).Tag, "#")
          gvrs.AddNew
          gvrs.Fields(0) = tvwBS.Nodes(i).Tag
          gvrs.Fields(1) = Mid(tvwBS.Nodes(i).Text, 1, 120)
          gvrs.Update
       End If
   Next
Else
   For i = 1 To n
       If Len(tvwBS.Nodes(i).Tag) >= 10 And InStr(tvwBS.Nodes(i).Tag, "#") > 0 Then
          k = InStr(tvwBS.Nodes(i).Tag, "#")
          gvrs.AddNew
          gvrs.Fields(0) = Mid(tvwBS.Nodes(i).Tag, 1, k - 1)
          gvrs.Fields(1) = Mid(tvwBS.Nodes(i).Text, 1, 60)
          gvrs.Update
       End If
   Next
End If

If Not XSelMultiple Then
    gvrs.AddNew
    gvrs.Fields(0) = Mid(tvwBS.SelectedItem.Tag, 1)
    gvrs.Fields(1) = Mid(tvwBS.SelectedItem.Text, 1, 60)
    gvrs.Update
End If

If Not (gvrs.BOF And gvrs.EOF) Then
   gvrs.MoveFirst
End If

vpSeleccion = True
Unload Me
End Sub

Private Sub cmdSalir_Click()
Set gvrs = New ADODB.Recordset
gvrs.Fields.Append "cProSelBSCod", adVarChar, 12, adFldMayBeNull
gvrs.Fields.Append "cBSDescripcion", adVarChar, 60, adFldMayBeNull
gvrs.Open
gvrs.AddNew
gvrs.Update
Set gvrs = Nothing
vpSeleccion = False
Unload Me
End Sub

Private Sub txtExp_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   nIniBusq = 1
   ListaTodosLosBienes
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

Sub ListaTodosLosBienes()
Dim Rs As New ADODB.Recordset
Dim n As Integer, k As Integer, cIMG As String
Dim oConn As DConecta, cProSelBSCod As String

On Error GoTo Salida

tvwBS.Nodes.Clear
tvwBS.Nodes.Add , , "K", "CMAC TRUJILLO" ', "logo"

sSQL = "select cProSelBSCod,cBSDescripcion from LogProSelBienesServicios  " & _
       " where bVigente=1 and len(cProSelBSCod)>=2 order by cProSelBSCod "

Set oConn = New DConecta
If oConn.AbreConexion Then
   Set Rs = oConn.CargaRecordSet(sSQL)
   oConn.CierraConexion
   If Not Rs.EOF Then
      Do While Not Rs.EOF
         cKey = "K" + Rs!cProSelBSCod
         
         Select Case Len(Rs!cProSelBSCod)
             Case 2
                  cKeySup = "K"
             Case 3
                  cKeySup = "K" + Left(Rs!cProSelBSCod, 2)
             Case 8
                  cKeySup = "K" + Mid(Rs!cProSelBSCod, 1, Len(Rs!cProSelBSCod) - 3)
             Case 5, 10
                  cKeySup = "K" + Mid(Rs!cProSelBSCod, 1, Len(Rs!cProSelBSCod) - 2)
         End Select
         
         If Mid(cKey, 3, 1) <> "2" Then
            If Len(Rs!cProSelBSCod) = 10 Then
               tvwBS.Nodes.Add cKeySup, tvwChild, cKey, Rs!cBSDescripcion ', "check0"
            Else
               tvwBS.Nodes.Add cKeySup, tvwChild, cKey, Rs!cBSDescripcion ', "enlace"
            End If
         Else
            If Len(cKey) > 3 Then
               cKeySup = Left(cKey, 3)
               tvwBS.Nodes.Add cKeySup, tvwChild, cKey, Rs!cBSDescripcion ', "check0"
            Else
               tvwBS.Nodes.Add cKeySup, tvwChild, cKey, Rs!cBSDescripcion ', "enlace"
            End If
            
         End If
         
         If Len(Rs!cProSelBSCod) < 10 Then
            tvwBS.Nodes(tvwBS.Nodes.Count).Tag = Rs!cProSelBSCod + "#"
         Else
            tvwBS.Nodes(tvwBS.Nodes.Count).Tag = Rs!cProSelBSCod
         End If
         
         Rs.MoveNext
      Loop
   End If
End If
Exit Sub
Salida:
  
End Sub

Private Sub mnuBuscar_Click()
BuscaExpresion
End Sub
