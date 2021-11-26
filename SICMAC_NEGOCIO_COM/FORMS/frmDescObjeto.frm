VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmDescObjeto 
   Caption         =   "Operaciones: Selección de Objetos"
   ClientHeight    =   5160
   ClientLeft      =   2805
   ClientTop       =   2415
   ClientWidth     =   7785
   Icon            =   "frmDescObjeto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   7785
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      Top             =   4620
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   6420
      TabIndex        =   2
      Top             =   4620
      Width           =   1215
   End
   Begin MSComctlLib.TreeView tvwObjeto 
      Height          =   3915
      Left            =   90
      TabIndex        =   0
      Top             =   540
      Width           =   7635
      _ExtentX        =   13467
      _ExtentY        =   6906
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   441
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      ImageList       =   "imgList"
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   120
      Top             =   4320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDescObjeto.frx":030A
            Key             =   "cerrado"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDescObjeto.frx":065C
            Key             =   "abierto"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDescObjeto.frx":09AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDescObjeto.frx":0E00
            Key             =   "cuenta"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDescObjeto.frx":1252
            Key             =   "objeto"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblObjeto 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Objeto"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   75
      TabIndex        =   1
      Top             =   105
      Width           =   7590
   End
End
Attribute VB_Name = "frmDescObjeto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sSql  As String
Dim rs As New ADODB.Recordset
Dim sCod As String, sEstado As String, sObj
Dim nNiv As Integer
Dim nIndex As Integer
Dim nodX As Node
Dim llOk As Boolean
Dim raiz As String
Dim nObjNiv As Integer
Dim siGrabo As Boolean
Dim nNivel As Integer

Public Sub Inicio(prs As Recordset, sObjCod As String, pnNiv As Integer, Optional sRaiz As String = "")
sCod = prs(0)
sObj = sObjCod
nNiv = pnNiv
raiz = sRaiz
Set rs = prs
Me.Show 1
End Sub

Private Sub cmdAceptar_Click()
Dim k As Integer
If tvwObjeto.Nodes(nIndex).Image <> "objeto" Then
   MsgBox " Selección de Objeto es a último Nivel...! ", vbCritical, "Error de Seleccion"
   Exit Sub
End If
' Asignar Objetos
GetDatosObjeto nIndex
llOk = True
Unload Me
DoEvents
End Sub

Private Sub cmdCancelar_Click()
Unload Me
DoEvents
End Sub

Private Sub Form_Activate()
If RSVacio(rs) Then
   Unload Me
End If
End Sub

Private Sub Form_Load()
Dim sCod As String
'On Error GoTo ErrObj
llOk = False

tvwObjeto.Nodes.Clear
Set tvwObjeto.ImageList = imgList
Set nodX = tvwObjeto.Nodes.Add()
rs.MoveFirst
If raiz = "" Then
   lblObjeto.Caption = " Objeto: " & rs(1)
   nObjNiv = rs!nObjetoNiv
   sCod = Trim(rs!cBSCod)
   nodX.Key = "K" & sCod
   nodX.Text = sCod & " - " & rs!cBSDescripcion
   AsignaImgNodo nObjNiv, nNiv, nodX
   nodX.Tag = CStr(rs!nObjetoNiv)
   rs.MoveNext
Else
   lblObjeto.Caption = " Objeto: " & raiz
   sCod = Mid(rs!cObjetoCod, 1, 2) & "X"
   nObjNiv = rs!nObjetoNiv - 1
   nodX.Key = "K" & sCod
   nodX.Text = sCod & " - " & raiz
   AsignaImgNodo nObjNiv, nNiv, nodX
   nodX.Tag = "0"
End If
CargaNodo sCod, tvwObjeto, rs, nNiv, nObjNiv
nIndex = 1
tvwObjeto.Nodes(1).Expanded = True
If Len(sObj) > 0 Then
   ExpandeObj
End If
Exit Sub
ErrObj:
  MsgBox TextErr(Err.Description), vbInformation, "Aviso"
End Sub

Private Sub ExpandeObj()
Dim i As Integer
For i = 1 To tvwObjeto.Nodes.Count
    If InStr(sObj, Mid(tvwObjeto.Nodes(i).Key, 2, 21)) = 1 Then
       tvwObjeto.Nodes(i).Expanded = True
       tvwObjeto.Nodes(i).Selected = True
       nIndex = i
    End If
Next
End Sub

Private Sub tvwObjeto_Collapse(ByVal Node As MSComctlLib.Node)
Node.ForeColor = "&H80000008"
End Sub

Private Sub tvwObjeto_DblClick()
If tvwObjeto.Nodes(nIndex).Image = "objeto" Then
   cmdAceptar_Click
End If
End Sub

Private Sub tvwObjeto_Expand(ByVal Node As MSComctlLib.Node)
Node.ForeColor = "&H8000000D"
End Sub

Private Sub tvwObjeto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If tvwObjeto.Nodes(nIndex).Image = "objeto" Then
      cmdAceptar_Click
   End If
End If
End Sub

Private Sub tvwObjeto_NodeClick(ByVal Node As MSComctlLib.Node)
  nIndex = Node.Index
  'frmMdiMain.staMain.Panels(2).Text = Node.Tag
End Sub

Public Property Get lOk() As Boolean
lOk = llOk
End Property

Public Property Let lOk(ByVal vNewValue As Boolean)
llOk = lOk
End Property

Private Sub GetDatosObjeto(nIndex As Integer)
Dim n As Integer
With tvwObjeto
  ' n = UBound(gaObj, 1)
  ReDim gaObj(1, 2, Val(.Nodes(nIndex).Tag) - nObjNiv) As String
  n = nIndex
  Do While True
     If InStr(.Nodes(nIndex).Key, .Nodes(n).Key) = 1 Then
        gaObj(0, 0, Val(.Nodes(n).Tag) - nObjNiv - 1) = Mid(.Nodes(n).Key, 2, Len(.Nodes(n).Key))
        gaObj(0, 1, Val(.Nodes(n).Tag) - nObjNiv - 1) = Mid(.Nodes(n).Text, InStr(.Nodes(n).Text, "-") + 2, 255)
     End If
     If Val(.Nodes(n).Tag) = nObjNiv + 1 Then
       Exit Do
     End If
     n = n - 1
     If n = 0 Then
        Exit Do
     End If
  Loop
End With
End Sub

Private Sub AsignaImgNodo(nObjNivel As Integer, pnNivel As Integer, nodX As Node, Optional plExpand As Boolean = False)
    If nObjNivel = pnNivel Then
       nodX.Image = "objeto"
    Else
       nodX.Image = "cerrado"
       nodX.ExpandedImage = "abierto"
       If plExpand Then
          nodX.ForeColor = "&H8000000D"
          nodX.Expanded = True
       End If
    End If
End Sub

Private Sub CargaNodo(psRaiz As String, tvw As TreeView, rsVista As ADODB.Recordset, pnNivel As Integer, pnObjNiv As Integer, Optional plExpand As Boolean = False)
    Dim sCod As String, siSale As Boolean
    Dim SiInstancia As Boolean
    Dim nodX As Node
    Dim pnOk As Integer
    Dim nObjNiv As Integer
    siGrabo = True
    siSale = False
    Do While Not rsVista.EOF
       'If Len(rsVista(0)) > Len(psRaiz) Then
       If rsVista!nObjetoNiv > pnObjNiv Then
          nNivel = nNivel + 1
          AdicionaNodo rsVista(0), rsVista(1), pnNivel, rsVista!nObjetoNiv, tvw, psRaiz, 4, plExpand
          siGrabo = True
          nObjNiv = rsVista!nObjetoNiv
          sCod = rsVista(0)
          rsVista.MoveNext
          CargaNodo sCod, tvw, rsVista, pnNivel, nObjNiv, plExpand
          nNivel = nNivel - 1
          If Not siGrabo Then
    '         If Len(rsVista(0)) = Len(psRaiz) Then
             If rsVista!nObjetoNiv = pnObjNiv Then
                AdicionaNodo rsVista(0), rsVista(1), pnNivel, rsVista!nObjetoNiv, tvw, psRaiz, 1, plExpand
                siGrabo = True
                psRaiz = rsVista(0)
                rsVista.MoveNext
             End If
          End If
       Else
    '      If Len(psRaiz) = Len(rsVista(0)) Then
          If rsVista!nObjetoNiv = pnObjNiv Then
             AdicionaNodo rsVista(0), rsVista(1), pnNivel, rsVista!nObjetoNiv, tvw, psRaiz, 1, plExpand
             psRaiz = rsVista(0)
             siGrabo = True
             rsVista.MoveNext
          Else
             If rsVista!nObjetoNiv < pnObjNiv Then
                siGrabo = False
                Exit Sub
             End If
          End If
       End If
    Loop
End Sub

Private Sub AdicionaNodo(sCod As String, sDes As String, pnNivel As Integer, pnObjNiv As Integer, tvwObjeto As TreeView, psRaiz As String, nTipo As Integer, Optional plExpand As Boolean = False)
    Dim nodX As Node
    Set nodX = tvwObjeto.Nodes.Add("K" & psRaiz, nTipo)
    nodX.Key = "K" & sCod
    nodX.Text = sCod & " - " & sDes
    AsignaImgNodo pnObjNiv, pnNivel, nodX, plExpand
    nodX.Tag = CStr(pnObjNiv)
End Sub
