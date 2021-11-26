VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLogMntBSGrupos 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4815
   ClientLeft      =   1695
   ClientTop       =   2610
   ClientWidth     =   8430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   8430
   Begin VB.Frame Frame1 
      Caption         =   "Grupos de Bienes y Servicios"
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
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8175
      Begin MSComctlLib.TreeView tvwObj 
         Height          =   3375
         Left            =   180
         TabIndex        =   1
         Top             =   300
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   5953
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   529
         LineStyle       =   1
         Style           =   7
         FullRowSelect   =   -1  'True
         Appearance      =   1
      End
      Begin VB.Frame fraBoton 
         BorderStyle     =   0  'None
         Height          =   795
         Left            =   180
         TabIndex        =   2
         Top             =   3660
         Width           =   7815
         Begin VB.CommandButton cmdSalir 
            Caption         =   "Salir"
            Height          =   375
            Left            =   6600
            TabIndex        =   5
            Top             =   420
            Width           =   1215
         End
         Begin VB.CommandButton cmdQuitar 
            Caption         =   "Quitar"
            Height          =   375
            Left            =   1260
            TabIndex        =   4
            Top             =   420
            Width           =   1155
         End
         Begin VB.CommandButton cmdNuevo 
            Caption         =   "Agregar"
            Height          =   375
            Left            =   0
            TabIndex        =   3
            Top             =   420
            Width           =   1215
         End
         Begin VB.Label Label1 
            BorderStyle     =   1  'Fixed Single
            Height          =   345
            Left            =   0
            TabIndex        =   10
            Top             =   15
            Width           =   7815
         End
      End
      Begin VB.Frame fraTexto 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   180
         TabIndex        =   6
         Top             =   3720
         Visible         =   0   'False
         Width           =   7815
         Begin VB.TextBox txtNodo 
            BackColor       =   &H00EAFFFF&
            Height          =   315
            Left            =   0
            TabIndex        =   9
            Top             =   0
            Width           =   7815
         End
         Begin VB.CommandButton cmdCancelar 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   6660
            TabIndex        =   8
            Top             =   360
            Width           =   1155
         End
         Begin VB.CommandButton cmdGrabar 
            Caption         =   "Grabar"
            Height          =   375
            Left            =   5460
            TabIndex        =   7
            Top             =   360
            Width           =   1155
         End
      End
   End
   Begin VB.Menu mnuObj 
      Caption         =   "MenuGeneral"
      Visible         =   0   'False
      Begin VB.Menu mnuAgregar 
         Caption         =   "Agregar"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuitar 
         Caption         =   "Quitar"
      End
   End
End
Attribute VB_Name = "frmLogMntBSGrupos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nNivel As Integer

Private Sub Form_Load()
CentraForm Me
Me.Caption = "Mantenimiento de Grupos de Bienes y Servicios"
CargaTreeView
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub cmdCancelar_Click()
fraBoton.Visible = True
fraTexto.Visible = False
End Sub

Sub CargaTreeView()
Dim rs As New ADODB.Recordset
Dim oConn As New DConecta
Dim sSql As String, cKey As String, cKeySup As String

tvwObj.Nodes.Clear

cKey = "K"
tvwObj.Nodes.Add , , cKey, "CAJA TRUJILLO"
tvwObj.Nodes(1).Tag = ""
   
sSql = "SELECT cBSGrupoCod, cBSGrupoDescripcion from  BSGrupos order by cBSGrupoCod"

If oConn.AbreConexion Then
   Set rs = oConn.CargaRecordSet(sSql)
   oConn.CierraConexion
   If Not rs.EOF Then
      Do While Not rs.EOF
         cKey = "K" + rs!cBSGrupoCod
         cKeySup = Left(cKey, Len(cKey) - 2)
         tvwObj.Nodes.Add cKeySup, tvwChild, cKey, rs!cBSGrupoCod + " - " + rs!cBSGrupoDescripcion
         tvwObj.Nodes(tvwObj.Nodes.Count).Tag = rs!cBSGrupoCod
         rs.MoveNext
      Loop
      tvwObj.Nodes(1).Expanded = True
   End If
End If
End Sub

Private Sub tvwObj_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
   PopupMenu mnuObj
End If
End Sub

Private Sub cmdGrabar_Click()
Dim k As Integer, rs As New ADODB.Recordset
Dim oConn As New DConecta, cBSGrupo As String
Dim cBSNuevoGrupo As String, sSql As String
Dim i As Integer, cCabecera As String

k = tvwObj.SelectedItem.Index

If nNivel = 1 Then
   cBSGrupo = Left(tvwObj.Nodes(k).Tag, 2)
   For i = k To 1 Step -1
    If tvwObj.Nodes(i).Tag = cBSGrupo Then
       cCabecera = tvwObj.Nodes(i).Text
    End If
   Next
End If

If nNivel = 2 Then
   cBSGrupo = tvwObj.Nodes(k).Tag
   cCabecera = tvwObj.Nodes(k).Text
   If Len(cBSGrupo) > 2 Then
      MsgBox "No es un nivel de cabecera..." + Space(10), vbInformation
      Exit Sub
   End If
End If

If Len(Trim(txtNodo.Text)) = 0 Then
   MsgBox "Debe indicar la descripción del Grupo..." + Space(10), vbInformation
   txtNodo.SetFocus
   Exit Sub
End If

If MsgBox("¿ Agregar el nuevo elemento al Grupo " + Space(10) + vbCrLf + cCabecera & " ?" + Space(10), vbQuestion + vbYesNo, "Confirme") = vbYes Then
   If oConn.AbreConexion Then
   
      Set rs = oConn.CargaRecordSet("Select cMax = coalesce(Max(cBSGrupoCod),'00') from BSGrupos where cBSGrupoCod like '" & cBSGrupo & "%' and len(cBSGrupoCod)=4 ")
      If Not rs.EOF Then
         cBSNuevoGrupo = cBSGrupo + Format(CInt(Right(rs!cMax, 2)) + 1, "00")
      Else
         cBSNuevoGrupo = rs!cMax
      End If
   
   
      sSql = "INSERT INTO BSGrupos (cBSGrupoCod,cBSGrupoDescripcion) " & _
             "  VALUES ('" & cBSNuevoGrupo & "','" & txtNodo.Text & "') "
      oConn.Ejecutar sSql
      
      tvwObj.Nodes.Add "K" + cBSGrupo, tvwChild, "K" + cBSNuevoGrupo, cBSNuevoGrupo + " - " + txtNodo.Text
      tvwObj.Nodes(tvwObj.Nodes.Count).Tag = cBSNuevoGrupo
      
      nNivel = 0
      fraBoton.Visible = True
      fraTexto.Visible = False
  End If
End If
End Sub

Private Sub mnuAgregar_Click()
nNivel = 2
fraBoton.Visible = False
fraTexto.Visible = True
txtNodo.Text = ""
txtNodo.SetFocus
End Sub

Private Sub cmdNuevo_Click()
nNivel = 1
fraBoton.Visible = False
fraTexto.Visible = True
txtNodo.Text = ""
txtNodo.SetFocus
End Sub

Private Sub tvwObj_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then
   cmdQuitar_Click
End If
End Sub

Private Sub cmdQuitar_Click()
Dim i As Integer
Dim cBSGrupoCod As String
Dim sSql As String
Dim oConn As New DConecta

i = tvwObj.SelectedItem.Index
cBSGrupoCod = tvwObj.Nodes(i).Tag

If MsgBox("¿ Está seguro de quitar el elemento indicado ?" + Space(10), vbQuestion + vbYesNo, "Confirme operación") = vbYes Then

   If Len(cBSGrupoCod) = 4 Then
      sSql = "DELETE FROM BSGrupos WHERE cBSGrupoCod = '" & cBSGrupoCod & "'"
   End If
   
   If Len(cBSGrupoCod) = 2 Then
      sSql = "DELETE FROM BSGrupos WHERE cBSGrupoCod like '" & cBSGrupoCod & "%'"
   End If
   
   If oConn.AbreConexion Then
      oConn.Ejecutar sSql
   End If
   
   tvwObj.Nodes.Remove (i)
   MsgBox "Se ha eliminado el Item!" + Space(10), vbInformation
End If
End Sub

Private Sub txtNodo_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
   cmdGrabar.SetFocus
End If
End Sub


      'tvwObj.Nodes.Add "K" + cBSGrupo, tvwChild, "K" + cBSNuevoGrupo, cBSNuevoGrupo + " - "
      'tvwObj.Nodes(tvwObj.Nodes.Count).Tag = cBSNuevoGrupo
      'tvwObj.SetFocus
      'For i = k To tvwObj.Nodes.Count
      '    If tvwObj.Nodes(i).Tag = cBSNuevoGrupo Then
      '
      '       Exit For
      '    End If
      'Next

