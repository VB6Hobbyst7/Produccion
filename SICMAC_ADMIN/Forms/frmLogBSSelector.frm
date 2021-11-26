VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmLogBSSelector 
   Caption         =   "Selección de Bienes / Servicios"
   ClientHeight    =   6360
   ClientLeft      =   5565
   ClientTop       =   1515
   ClientWidth     =   6375
   LinkTopic       =   "Form2"
   ScaleHeight     =   6360
   ScaleWidth      =   6375
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   0
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
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Seleccionar"
      Height          =   380
      Left            =   3360
      TabIndex        =   1
      Top             =   5940
      Width           =   1575
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   380
      Left            =   5040
      TabIndex        =   0
      Top             =   5940
      Width           =   1215
   End
   Begin MSComctlLib.ImageList imgLista 
      Left            =   0
      Top             =   5760
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
            Picture         =   "frmLogBSSelector.frx":0000
            Key             =   "logo"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogBSSelector.frx":0352
            Key             =   "check0"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogBSSelector.frx":06A4
            Key             =   "check1"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogBSSelector.frx":09F6
            Key             =   "enlace"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwBS 
      Height          =   5115
      Left            =   120
      TabIndex        =   2
      Top             =   720
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
   Begin VB.Menu mnuBusqueda 
      Caption         =   "Buscar"
      Visible         =   0   'False
      Begin VB.Menu mnuBuscar 
         Caption         =   "Buscar siguiente"
      End
   End
End
Attribute VB_Name = "frmLogBSSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public vpSeleccion As Boolean
Public gvrs As ADODB.Recordset

Dim nIniBusq As Integer, cRHCargoCod As String

Public Sub SeleccionBienesCargo(ByVal psCargoCod As String)
cRHCargoCod = psCargoCod
Me.Show 1
End Sub

Private Sub Form_Load()
CentraForm Me
Set tvwBS.ImageList = imgLista
ListaBienes
End Sub

Private Sub tvwBS_DblClick()
Dim i As Integer, n As Integer
Dim cBSCod As String
Dim sSQL As String
Dim nCheck As Integer

i = tvwBS.SelectedItem.Index
If Len(tvwBS.Nodes(i).Tag) >= 10 Then
   n = InStr(tvwBS.Nodes(i).Tag, "*")
   cBSCod = Replace(tvwBS.Nodes(i).Tag, "*", "")
   If n > 0 Then
      tvwBS.Nodes(i).Image = "check0"
      tvwBS.Nodes(i).Tag = Mid(tvwBS.Nodes(i).Tag, 1, n - 1)
      nCheck = 0
   Else
      tvwBS.Nodes(i).Image = "check1"
      tvwBS.Nodes(i).Tag = tvwBS.Nodes(i).Tag + "*"
      nCheck = 1
  End If
End If
End Sub

Private Sub tvwBS_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
   PopupMenu mnuBusqueda
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

Sub ListaBienes()
Dim rs As New ADODB.Recordset
Dim n As Integer, K As Integer, cIMG As String
Dim oConn As DConecta, cProSelBSCod As String
Dim sSQL As String
Dim cKey As String, cKeySup As String

On Error GoTo Salida

tvwBS.Nodes.Clear
tvwBS.Nodes.Add , , "K", "CMAC TRUJILLO", "logo"

sSQL = "select cProSelBSCod,cBSDescripcion from LogProSelBienesServicios  " & _
       " where bVigente=1 and len(cProSelBSCod)>=2 order by cProSelBSCod "

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
         
         If Mid(cKey, 3, 1) <> "2" Then
            If Len(rs!cProSelBSCod) = 10 Then
               tvwBS.Nodes.Add cKeySup, tvwChild, cKey, rs!cBSDescripcion, "check0"
            Else
               tvwBS.Nodes.Add cKeySup, tvwChild, cKey, rs!cBSDescripcion, "enlace"
            End If
         Else
            If Len(cKey) > 3 Then
               cKeySup = Left(cKey, 3)
               tvwBS.Nodes.Add cKeySup, tvwChild, cKey, rs!cBSDescripcion, "check0"
            Else
               tvwBS.Nodes.Add cKeySup, tvwChild, cKey, rs!cBSDescripcion, "enlace"
            End If
            
         End If
         
         If Len(rs!cProSelBSCod) < 10 Then
            tvwBS.Nodes(tvwBS.Nodes.Count).Tag = rs!cProSelBSCod + "#"
         Else
            tvwBS.Nodes(tvwBS.Nodes.Count).Tag = rs!cProSelBSCod
         End If
         
         rs.MoveNext
      Loop
   End If
End If
Exit Sub
Salida:
End Sub

Private Sub cmdSelect_Click()
Dim i As Integer, n As Integer
Dim K As Integer

Set gvrs = New ADODB.Recordset
gvrs.Fields.Append "cProSelBSCod", adVarChar, 12, adFldMayBeNull
gvrs.Fields.Append "cBSDescripcion", adVarChar, 120, adFldMayBeNull
gvrs.Open

n = tvwBS.Nodes.Count

   For i = 1 To n
       If Len(tvwBS.Nodes(i).Tag) >= 10 And InStr(tvwBS.Nodes(i).Tag, "*") > 0 Then
          K = InStr(tvwBS.Nodes(i).Tag, "*")
          gvrs.AddNew
          gvrs.Fields(0) = Mid(tvwBS.Nodes(i).Tag, 1, K - 1)
          gvrs.Fields(1) = Mid(tvwBS.Nodes(i).Text, 1, 60)
          gvrs.Update
       End If
   Next

If Not (gvrs.BOF And gvrs.EOF) Then
   gvrs.MoveFirst
End If

vpSeleccion = True
Unload Me
End Sub

Private Sub mnuBuscar_Click()
BuscaExpresion
End Sub
