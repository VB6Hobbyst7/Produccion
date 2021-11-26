VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDescObjeto 
   Caption         =   "Operaciones: Selección de Objetos"
   ClientHeight    =   5820
   ClientLeft      =   1815
   ClientTop       =   1650
   ClientWidth     =   8025
   Icon            =   "frmDescObjeto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   8025
   Begin VB.CommandButton cmdBuscarSig 
      Caption         =   "Buscar &Siguiente"
      Height          =   420
      Left            =   1500
      TabIndex        =   5
      Top             =   5280
      Width           =   1350
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Buscar"
      Height          =   420
      Left            =   120
      TabIndex        =   4
      Top             =   5280
      Width           =   1350
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   420
      Left            =   5130
      TabIndex        =   3
      Top             =   5280
      Width           =   1350
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   420
      Left            =   6480
      TabIndex        =   2
      Top             =   5280
      Width           =   1350
   End
   Begin MSComctlLib.TreeView tvwObjeto 
      Height          =   4650
      Left            =   90
      TabIndex        =   0
      Top             =   525
      Width           =   7770
      _ExtentX        =   13705
      _ExtentY        =   8202
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
      Left            =   150
      Top             =   4530
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
      Width           =   7770
   End
End
Attribute VB_Name = "frmDescObjeto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sSQL  As String
Dim rs As New ADODB.Recordset
Dim sCod As String, sEstado As String, sObj
'Dim nNiv As Integer
Dim nIndex As Integer
Dim nodX As Node
Dim llOk As Boolean
Dim raiz As String
Dim nObjNiv As Integer
Dim gaObj() As String
Dim lnColCod As Long
Dim lnColDesc As Long
Public psDatoCod As String
Public psDatoDesc As String
Public vbUltNiv As Boolean
Dim nBuscarPos As Integer
Dim sBuscarTexto As String

Public Sub inicio(prs As ADODB.Recordset, sObjCod As String, Optional sRaiz As String = "")
sCod = prs(0)
sObj = sObjCod
raiz = sRaiz
Set rs = prs
Me.Show 1
End Sub
Private Sub cmdAceptar_Click()
Dim K As Integer

'If tvwObjeto.Nodes(nIndex).Image <> "objeto" Then
'   MsgBox " Selección de Objeto es a último Nivel...! ", vbCritical, "Error de Seleccion"
'   Exit Sub
'End If
If vbUltNiv Then
    If tvwObjeto.Nodes(nIndex).Children > 0 Then
        MsgBox " Selección de Objeto es a último Nivel...! ", vbInformation, "Aviso de Seleccion"
        Exit Sub
    End If
End If

'Asignar Objetos
GetDatosObjeto nIndex
llOk = True
Unload Me
End Sub

Private Sub BuscarDato(ByVal nPos As Integer, ByVal psBuscarTexto As String)
Dim K As Integer
   For K = nPos + 1 To tvwObjeto.Nodes.Count
      If InStr(UCase(tvwObjeto.Nodes(K).Text), UCase(psBuscarTexto)) > 0 Then
         tvwObjeto.Nodes(K).Selected = True
         nBuscarPos = K
         Exit For
      End If
   Next
   If nPos = nBuscarPos Then
      MsgBox " ¡ Dato no encontrado ! ", vbInformation, "¡Aviso!"
   End If
   tvwObjeto.SetFocus
End Sub

Private Sub CmdBuscar_Click()
nBuscarPos = 0
If Me.tvwObjeto.Nodes.Count > 0 Then
   sBuscarTexto = InputBox("Descripción de Producto a Buscar ", "Busca de Bienes")
   BuscarDato nBuscarPos, sBuscarTexto
End If
End Sub

Private Sub cmdBuscarSig_Click()
BuscarDato nBuscarPos, sBuscarTexto
End Sub

Private Sub cmdCancelar_Click()
psDatoCod = ""
psDatoDesc = ""
Me.Hide
End Sub

Private Sub Form_Activate()
If rs.EOF And rs.BOF Then
   Unload Me
End If
End Sub
Private Sub Form_Load()
Dim sCod As String
On Error GoTo ErrObj
llOk = False
tvwObjeto.Nodes.Clear
CentraForm Me
Set tvwObjeto.ImageList = imgList
Set nodX = tvwObjeto.Nodes.Add()
rs.MoveFirst
If raiz = "" Then
   lblObjeto.Caption = " Objeto: " & rs(lnColDesc)
   nObjNiv = rs(2)
   sCod = rs(lnColCod)
   nodX.Key = "K" & sCod
   nodX.Text = sCod & " - " & rs(lnColDesc)
   AsignaImgNodo nodX
   nodX.Tag = CStr(rs(2))
   rs.MoveNext
Else
   lblObjeto.Caption = " Objeto: " & raiz
   sCod = Mid(rs(lnColCod), 1, 2) & "X"
   nObjNiv = rs(2) - 1
   nodX.Key = "K" & sCod
   nodX.Text = sCod & " - " & raiz
   AsignaImgNodo nodX
   nodX.Tag = "0"
End If
CargaNodo sCod, tvwObjeto, rs, nObjNiv, lnColCod, lnColDesc
nIndex = 1
tvwObjeto.Nodes(1).Expanded = True
If Len(sObj) > 0 Then
   ExpandeObj
End If
nBuscarPos = 1

Exit Sub
ErrObj:
   Err.Raise Err.Number, "frmDescObjeto-form-load", Err.Description
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

Private Sub Form_Unload(Cancel As Integer)
'Set oDescObj = Nothing
End Sub

Private Sub mnuBuscarIni_Click()

End Sub

Private Sub tvwObjeto_Collapse(ByVal Node As MSComctlLib.Node)
Node.ForeColor = "&H80000008"
End Sub

Private Sub tvwObjeto_DblClick()
If tvwObjeto.Nodes(nIndex).Children = 0 Then
    cmdAceptar_Click
End If
End Sub

Private Sub tvwObjeto_Expand(ByVal Node As MSComctlLib.Node)
Node.ForeColor = "&H8000000D"
End Sub

Private Sub tvwObjeto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If tvwObjeto.Nodes(nIndex).Children = 0 Then
      cmdAceptar_Click
   End If
End If
End Sub
Private Sub tvwObjeto_NodeClick(ByVal Node As MSComctlLib.Node)
    nIndex = Node.Index
End Sub

Public Property Get lOk() As Boolean
lOk = llOk
End Property

Public Property Let lOk(ByVal vNewValue As Boolean)
llOk = lOk
End Property

Private Sub GetDatosObjeto(nIndex As Integer)
Dim n As Integer

psDatoCod = Mid(tvwObjeto.Nodes(nIndex).Key, 2, Len(tvwObjeto.Nodes(nIndex).Key))
psDatoDesc = Mid(tvwObjeto.Nodes(nIndex).Text, InStr(tvwObjeto.Nodes(nIndex).Text, "-") + 2, 255)
End Sub
Public Property Get ColCod() As Long
ColCod = lnColCod
End Property

Public Property Let ColCod(ByVal vNewValue As Long)
lnColCod = vNewValue
End Property
Public Property Get ColDesc() As Long
ColDesc = lnColDesc
End Property
Public Property Let ColDesc(ByVal vNewValue As Long)
lnColDesc = vNewValue
End Property
