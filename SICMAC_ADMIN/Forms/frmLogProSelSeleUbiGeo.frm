VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmLogProSelSeleUbiGeo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Selector de Ubicacion Geografica"
   ClientHeight    =   6330
   ClientLeft      =   3240
   ClientTop       =   2025
   ClientWidth     =   5715
   Icon            =   "frmLogProSelSeleUbiGeo.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   5715
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Seleccionar"
      Default         =   -1  'True
      Height          =   380
      Left            =   2640
      TabIndex        =   2
      Top             =   5880
      Width           =   1575
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   380
      Left            =   4320
      TabIndex        =   1
      Top             =   5880
      Width           =   1215
   End
   Begin MSComctlLib.ImageList ImgLista 
      Left            =   300
      Top             =   4860
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogProSelSeleUbiGeo.frx":08CA
            Key             =   "enlace"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwObj 
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   5475
      _ExtentX        =   9657
      _ExtentY        =   10186
      _Version        =   393217
      Indentation     =   529
      LineStyle       =   1
      Style           =   7
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
End
Attribute VB_Name = "frmLogProSelSeleUbiGeo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public vpCodUbigeo As String
Public vpUbigeoDesc As String
Dim pbConsucode As Boolean
Dim cKey As String, cKeySup As String, bConsucode As Boolean
Public gvNoddo As String, gvCodigo As String

Public Sub FuenteConsucode(Optional ByVal vSi As Boolean = False)
bConsucode = vSi
Me.Show 1
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdSelect_Click()
   gvNoddo = tvwObj.SelectedItem.Text
   gvCodigo = tvwObj.SelectedItem.Tag
   Unload Me
End Sub

Private Sub Form_Load()
CentraForm Me
If bConsucode Then
   UbigeoConsucode
Else
   PrimerNivel
   Me.vpCodUbigeo = ""
   Me.vpUbigeoDesc = ""
End If
End Sub

Sub UbigeoConsucode()
Dim Rs As New ADODB.Recordset
Dim oConn As DConecta
Dim sSQL As String

Set oConn = New DConecta

tvwObj.Nodes.Clear
sSQL = "select cUbigeoCod as cCodigo, cUbigeoDescripcion as cDescripcion from LogProSelUbigeo order by cUbigeoCod "
If oConn.AbreConexion Then
   Set Rs = oConn.CargaRecordSet(sSQL)
   If Not Rs.EOF Then
      cKey = "K"
      tvwObj.Nodes.Add , , cKey, "PERU"
      Do While Not Rs.EOF
         cKey = "K" + Rs!cCodigo
         cKeySup = "K" + Left(Rs!cCodigo, Len(Rs!cCodigo) - 2)
         tvwObj.Nodes.Add cKeySup, tvwChild, cKey, Rs!cDescripcion
         tvwObj.Nodes(tvwObj.Nodes.Count).Tag = Rs!cCodigo
         Rs.MoveNext
      Loop
      tvwObj.Nodes(1).Expanded = True
   End If
End If
End Sub

Sub PrimerNivel()
Dim Rs As New ADODB.Recordset
Dim oConn As DConecta
Dim sSQL As String

Set oConn = New DConecta

tvwObj.Nodes.Clear

sSQL = "select cCodigo='4'+substring(cUbigeoCod,2,2), cDescripcion=cUbigeoDescripcion from UbicacionGeografica " & _
       " where left(cUbigeoCod,1)='1' and cUbigeoDescripcion <>'Migracion NO DEFINIDO EN MAESTRO' " & _
       " order by cUbigeoCod "

If oConn.AbreConexion Then
   Set Rs = oConn.CargaRecordSet(sSQL)
   If Not Rs.EOF Then
      cKey = "K"
      tvwObj.Nodes.Add , , cKey, "PERU"
      Do While Not Rs.EOF
         cKey = "K" + Rs!cCodigo
         cKeySup = "K"
         tvwObj.Nodes.Add cKeySup, tvwChild, cKey, Rs!cDescripcion
         tvwObj.Nodes(tvwObj.Nodes.Count).Tag = Rs!cCodigo
         Rs.MoveNext
      Loop
      tvwObj.Nodes(1).Expanded = True
   End If
End If
End Sub

Private Sub tvwObj_KeyPress(KeyAscii As Integer)
Dim xCodigo As String
Dim xRaiz As String
If KeyAscii = 13 Then
   xCodigo = tvwObj.Nodes(tvwObj.SelectedItem.Index).Tag
   If InStr(tvwObj.Nodes(tvwObj.SelectedItem.Index).Tag, "#") > 0 Then
      MsgBox "No es una selección de último nivel..." + Space(10), vbInformation
      Exit Sub
   Else
      Me.vpCodUbigeo = ""
      Me.vpUbigeoDesc = ""
      
      If bConsucode Then
         Me.vpCodUbigeo = xCodigo
         Me.vpUbigeoDesc = GetUbigeoConsucode(xCodigo)
      Else
         If Len(xCodigo) > 0 And Len(xCodigo) < 12 Then
            Select Case Len(xCodigo)
             Case 3
                  xRaiz = "1"
             Case 5
                  xRaiz = "2"
             Case 7
                  xRaiz = "3"
            End Select
            xCodigo = xRaiz + Right(xCodigo, Len(xCodigo) - 1)
            Me.vpCodUbigeo = xCodigo + String(12 - Len(xCodigo), "0")
         Else
            xRaiz = "4"
            Me.vpCodUbigeo = xCodigo
         End If
         Me.vpUbigeoDesc = UbigeoDescCompleto(xRaiz, Me.vpCodUbigeo)
      End If
      Unload Me
   End If
End If
If KeyAscii = 27 Then
   Me.vpCodUbigeo = ""
   Me.vpUbigeoDesc = ""
   Unload Me
End If
End Sub

Private Sub tvwObj_NodeClick(ByVal Node As MSComctlLib.Node)
On Error GoTo tvwObj_NodeErr
Dim Rs As New ADODB.Recordset
Dim oConn As DConecta
Dim xCodigo As String
Dim sSQL As String
Dim xCod As String
Dim cFig As String
Set oConn = New DConecta

xCodigo = Node.Tag
If InStr(xCodigo, "#") > 0 Then Exit Sub

xCod = Right(xCodigo, Len(xCodigo) - 1)
Select Case Len(xCod)
    Case 2
         sSQL = "select cCodigo='4'+substring(cUbigeoCod,2,4), cDescripcion=cUbigeoDescripcion from UbicacionGeografica " & _
         " where left(cUbigeoCod,1)='2' and substring(cUbigeoCod,2,2)='" & xCod & "' and cUbigeoDescripcion <>'Migracion NO DEFINIDO EN MAESTRO' " & _
         " order by cUbigeoCod"
    Case 4
         sSQL = "select cCodigo='4'+substring(cUbigeoCod,2,6), cDescripcion=cUbigeoDescripcion from UbicacionGeografica " & _
         " where left(cUbigeoCod,1)='3' and substring(cUbigeoCod,2,4)='" & xCod & "' and cUbigeoDescripcion <>'Migracion NO DEFINIDO EN MAESTRO' " & _
         " order by cUbigeoCod"
    Case 6
         sSQL = "select cCodigo=cUbigeoCod, cDescripcion=cUbigeoDescripcion from UbicacionGeografica " & _
         " where left(cUbigeoCod,1)='4' and substring(cUbigeoCod,2,6)='" & xCod & "' and cUbigeoDescripcion <>'Migracion NO DEFINIDO EN MAESTRO' " & _
         " order by cUbigeoDescripcion"
    Case Else
         Exit Sub
End Select

If oConn.AbreConexion Then
   Set Rs = oConn.CargaRecordSet(sSQL)
   If Not Rs.EOF Then
      Do While Not Rs.EOF
         cKey = "K" + Rs!cCodigo
         cKeySup = "K" + xCodigo
         tvwObj.Nodes.Add cKeySup, tvwChild, cKey, Rs!cDescripcion
         tvwObj.Nodes(tvwObj.Nodes.Count).Tag = Rs!cCodigo
         Rs.MoveNext
      Loop
      tvwObj.Nodes(Node.Index).Tag = tvwObj.Nodes(Node.Index).Tag + "#"
      tvwObj.Nodes(Node.Index).Expanded = True
   End If
End If
Exit Sub
tvwObj_NodeErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Aviso"
End Sub

