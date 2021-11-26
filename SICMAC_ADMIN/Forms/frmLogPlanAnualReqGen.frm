VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLogPlanAnualReqGen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Plan Anual de Adquisiciones  y Contrataciones"
   ClientHeight    =   6150
   ClientLeft      =   105
   ClientTop       =   2130
   ClientWidth     =   11715
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   11.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   11715
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   60
      TabIndex        =   6
      Top             =   60
      Width           =   11595
      Begin VB.TextBox txtAnio 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5640
         TabIndex        =   10
         Text            =   "2005"
         Top             =   90
         Width           =   675
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Requerimientos para CMAC-T"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   7680
         TabIndex        =   9
         Top             =   120
         Width           =   3390
      End
      Begin VB.Label lblPlan 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Plan Anual de Adquisiciones y Contrataciones"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   5280
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00EAFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         Height          =   495
         Left            =   0
         Top             =   0
         Width           =   6915
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00EAFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         Height          =   495
         Left            =   6960
         Top             =   0
         Width           =   4635
      End
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2340
      Left            =   60
      TabIndex        =   8
      Top             =   480
      Width           =   11595
      Begin MSComctlLib.TreeView tvwObj 
         Height          =   1995
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   11355
         _ExtentX        =   20029
         _ExtentY        =   3519
         _Version        =   393217
         Indentation     =   353
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         ImageList       =   "imgLista"
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
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10440
      TabIndex        =   5
      Top             =   5700
      Width           =   1215
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "Agregar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   60
      TabIndex        =   2
      Top             =   5700
      Width           =   1155
   End
   Begin VB.CommandButton cmdQuitar 
      Caption         =   "Quitar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1260
      TabIndex        =   1
      Top             =   5700
      Width           =   1155
   End
   Begin VB.TextBox txtEdit 
      BackColor       =   &H00DDFFFE&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   9540
      TabIndex        =   0
      Top             =   3420
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "Grabar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9180
      TabIndex        =   4
      Top             =   5700
      Width           =   1215
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFlex 
      Height          =   2835
      Left            =   60
      TabIndex        =   3
      Top             =   2820
      Width           =   11595
      _ExtentX        =   20452
      _ExtentY        =   5001
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   18
      FixedCols       =   0
      ForeColorFixed  =   -2147483646
      BackColorBkg    =   16777215
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483633
      GridColorUnpopulated=   -2147483633
      FocusRect       =   0
      HighLight       =   2
      ScrollBars      =   2
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
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   18
   End
   Begin MSComctlLib.ImageList imgLista 
      Left            =   4080
      Top             =   4920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Image Image3 
      Height          =   240
      Left            =   3360
      Picture         =   "frmLogPlanAnualReqGen.frx":0000
      Top             =   5340
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   3060
      Picture         =   "frmLogPlanAnualReqGen.frx":0342
      Top             =   5340
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   2760
      Picture         =   "frmLogPlanAnualReqGen.frx":0684
      Top             =   5340
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "MenuReq"
      Visible         =   0   'False
      Begin VB.Menu mnuInfo 
         Caption         =   "Info del trámite "
      End
   End
End
Attribute VB_Name = "frmLogPlanAnualReqGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mMes(1 To 12) As String, sSQL As String
'Dim cRHAgeCod As String, cRHAreaCod As String, cRHCargoCod As String
Dim nEditable As Boolean, nPlanReqActual As Long

Private Sub cmdGrabar_Click()
Dim oConn As DConecta, sSQL As String
Dim cProSelBSCod As String, i As Integer, N As Integer
Dim nPlanNro As Long, Rs As New ADODB.Recordset
Dim nItem As Integer, cLogNro As String, nAnio As Integer
Dim cRHAgeCod As String, cRHAreaCod As String, cRHCargoCod As String
Dim cPersCod As String

'Persona: CAJA MUNICIPAL DE TRUJILLO ------------------
cPersCod = "1120800013498"
cRHAgeCod = ""
cRHAreaCod = ""
cRHCargoCod = ""

nItem = 0
N = MSFlex.Rows - 1
For i = 1 To N
    If Len(MSFlex.TextMatrix(i, 1)) > 0 And Len(MSFlex.TextMatrix(i, 2)) > 0 Then
       nItem = nItem + 1
    End If
Next

If nItem = 0 Then
   MsgBox "Debe seleccionar al menos un Bien / Servicio..." + Space(10), vbInformation
   Exit Sub
End If

nItem = 0
nAnio = CInt(VNumero(txtAnio.Text))
Set oConn = New DConecta
If oConn.AbreConexion Then
   If MsgBox("¿ Está seguro de grabar ?" + Space(10), vbQuestion + vbYesNo, "Confirme") = vbYes Then
   
      cLogNro = GetLogMovNro
      
      '1º Deshabilitamos los requerimientos anteriores del usuario
      'sSQL = "UPDATE LogPlanAnualReq SET nEstado = 0 WHERE cPersCod = '" & txtPersCod.Text & "'"
      'oConn.Ejecutar sSQL
      
      'sSQL = "UPDATE LogPlanAnualReqDetalle SET nEstado = 0 WHERE cPersCod = '" & txtPersCod.Text & "'"
      'oConn.Ejecutar sSQL
      
      '2º Insertamos cabecera del requerimiento actual del usuario
      sSQL = "INSERT INTO LogPlanAnualReq( nAnio, cPersCod, cRHCargoCod, cRHAreaCod, cRHAgeCod, cMovNro) " & _
             "    VALUES (" & nAnio & ",'" & cPersCod & "','" & cRHCargoCod & "','" & cRHAreaCod & "','" & cRHAgeCod & "','" & cLogNro & "') "
      oConn.Ejecutar sSQL
      
      '3º Hallamos ultima secuencia de los requerimientos
      nPlanNro = UltimaSecuenciaIdentidad("LogPlanAnualReq")
      
      '---------------------------------------------------------------------------------
      nItem = 0
      For i = 1 To N
          If Len(MSFlex.TextMatrix(i, 1)) > 0 Then
             nItem = nItem + 1
             sSQL = "INSERT INTO LogPlanAnualReqDetalle (nPlanReqNro,cPersCod,nAnio, nItem, cBSCod, " & _
                  "            nMes01, nMes02, nMes03, nMes04, nMes05, nMes06, nMes07, nMes08, nMes09, nMes10, nMes11, nMes12) " & _
                  " VALUES (" & nPlanNro & ",'" & cPersCod & "'," & nAnio & "," & nItem & ",'" & MSFlex.TextMatrix(i, 1) & "'," & _
                  "         " & VNumero(MSFlex.TextMatrix(i, 4)) & "," & VNumero(MSFlex.TextMatrix(i, 5)) & "," & VNumero(MSFlex.TextMatrix(i, 6)) & "," & VNumero(MSFlex.TextMatrix(i, 7)) & "," & VNumero(MSFlex.TextMatrix(i, 8)) & "," & VNumero(MSFlex.TextMatrix(i, 9)) & ", " & _
                  "         " & VNumero(MSFlex.TextMatrix(i, 10)) & "," & VNumero(MSFlex.TextMatrix(i, 11)) & "," & VNumero(MSFlex.TextMatrix(i, 12)) & "," & VNumero(MSFlex.TextMatrix(i, 13)) & "," & VNumero(MSFlex.TextMatrix(i, 14)) & "," & VNumero(MSFlex.TextMatrix(i, 15)) & " )"
             oConn.Ejecutar sSQL
          End If
      Next
      
      sSQL = " insert into LogPlanAnualAprobacion (nPlanReqNro,cRHCargoCodAprobacion,nNivelAprobacion) " & _
             " select " & nPlanNro & ",cRHCargoCodAprobacion,nNivelAprobacion " & _
             " from LogNivelAprobacion where cRHCargoCod = '" & cRHCargoCod & "' order by nNivelAprobacion "
      oConn.Ejecutar sSQL
      
      MsgBox "El requerimiento se ha grabado con éxito!" + Space(10), vbInformation
      cPersCod = ""
      FormaFlex
   End If
   oConn.CierraConexion
End If
End Sub

Private Sub cmdQuitar_Click()
Dim i As Integer
Dim K As Integer

i = MSFlex.row
If Len(Trim(MSFlex.TextMatrix(i, 1))) = 0 Then
   Exit Sub
End If

If MsgBox("¿ está seguro de quitar el elemento ?" + Space(10), vbQuestion + vbYesNo, "Confirme") = vbYes Then
   If MSFlex.Rows - 1 > 1 Then
      MSFlex.RemoveItem i
   Else
      'MSFlex.Clear          Quita las cabeceras
      For K = 0 To MSFlex.Cols - 1
          MSFlex.TextMatrix(i, K) = ""
      Next
      MSFlex.RowHeight(i) = 8
   End If
End If

End Sub

Private Sub Form_Load()
CentraForm Me
txtAnio.Text = Year(gdFecSis) + 1
mMes(1) = "ENE"
mMes(2) = "FEB"
mMes(3) = "MAR"
mMes(4) = "ABR"
mMes(5) = "MAY"
mMes(6) = "JUN"
mMes(7) = "JUL"
mMes(8) = "AGO"
mMes(9) = "SEP"
mMes(10) = "OCT"
mMes(11) = "NOV"
mMes(12) = "DIC"
imgLista.ListImages.Add 1, "cmact", Image1
imgLista.ListImages.Add 2, "agencia", Image2
imgLista.ListImages.Add 3, "area", Image3
nEditable = True
CargaAgencias
FormaFlex
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Sub CargaAgencias()
Dim oConn As New DConecta
Dim Rs As New ADODB.Recordset
Dim cKey As String, cKeySup  As String

tvwObj.Nodes.Clear

sSQL = "select distinct r.cAgenciaActual as cCodigo, a.cAgeDescripcion as cDescripcion " & _
       "  from rrhh r inner join Agencias a on r.cAgenciaActual = a.cAgeCod " & _
       " Union " & _
       " select distinct cAgenciaActual+cAreaCodActual as cCodigo, a.cAreaDescripcion as cDescripcion " & _
       "  from rrhh r inner join Areas a on r.cAreaCodActual = a.cAreaCod " & _
       " order by cCodigo"
       
If oConn.AbreConexion Then
   Set Rs = oConn.CargaRecordSet(sSQL)
   If Not Rs.EOF Then
      tvwObj.Nodes.Add , , "K", "CAJA TRUJILLO", "cmact"
      Do While Not Rs.EOF
         cKeySup = ""
         cKey = "K" + Rs!cCodigo
         If Len(Rs!cCodigo) > 2 Then
            cKeySup = Left(cKey, 3)
            tvwObj.Nodes.Add cKeySup, tvwChild, cKey, Rs(1), "area"
         Else
            cKeySup = "K"
            tvwObj.Nodes.Add cKeySup, tvwChild, cKey, Rs(1), "agencia"
         End If
         'tvwObj.Nodes.Add cKeySup, tvwChild, cKey, rs(1), "agencia"
         tvwObj.Nodes(tvwObj.Nodes.Count).Tag = Rs(0)
         Rs.MoveNext
      Loop
   End If
End If
End Sub

Private Sub mnuInfo_Click()
frmLogPlanAnualInfo.PlanAnual nPlanReqActual
End Sub

Private Sub MSFlex_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete And MSFlex.Col >= 4 And MSFlex.Col <= 15 And nEditable Then
   MSFlex.TextMatrix(MSFlex.row, MSFlex.Col) = ""
End If
End Sub

Private Sub MSFlex_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
   PopupMenu mnuMenu
End If
End Sub

Private Sub tvwObj_NodeCheck(ByVal Node As MSComctlLib.Node)
Dim i As Integer, N As Integer, K As Integer
Dim xValor As Boolean

i = Node.Index
Node.Expanded = True
xValor = Node.Checked
If Len(Node.Tag) <= 2 Then
   For K = i To Node.Child.LastSibling.Index
       tvwObj.Nodes(K).Checked = xValor
   Next
End If
End Sub

Private Sub txtAnio_GotFocus()
SelTexto txtAnio
End Sub

Private Sub txtanio_KeyPress(KeyAscii As Integer)
nKeyAscii = KeyAscii
KeyAscii = DigNumEnt(KeyAscii)
If nKeyAscii = 13 Then
   tvwObj.SetFocus
End If
End Sub

Sub Totaliza()
Dim i As Integer, j As Integer, N As Integer
Dim nSuma As Currency
N = MSFlex.Rows - 1
For i = 1 To N
    nSuma = 0
    For j = 1 To 12
        nSuma = nSuma + VNumero(MSFlex.TextMatrix(i, j + 3))
    Next
    MSFlex.TextMatrix(i, 16) = nSuma
Next
End Sub

Private Sub cmdAgregar_Click()
Dim i As Integer, Rs As New ADODB.Recordset

i = MSFlex.Rows - 1
If Len(Trim(MSFlex.TextMatrix(i, 1))) = 0 Then
   i = i - 1
End If

'frmBSSelector.TodosConCheck True
frmLogProSelBSSelector.TodosConCheck False
Set Rs = frmLogProSelBSSelector.gvrs
If Rs.State <> 0 Then
   If Not Rs.EOF Then
      Do While Not Rs.EOF
         If Not YaEsta(Rs!cProSelBSCod) Then
            i = i + 1
            InsRow MSFlex, i
            MSFlex.TextMatrix(i, 1) = Rs!cProSelBSCod
            MSFlex.TextMatrix(i, 2) = Rs!cBSDescripcion
            MSFlex.TextMatrix(i, 3) = GetBSUnidadLog(Rs!cProSelBSCod)
            If Left(Rs!cProSelBSCod, 2) = "12" Then
               MSFlex.row = i
               MSFlex.Col = 4
               frmLogProSelEspecificaciones.Inicio MSFlex.Left + MSFlex.CellLeft + 120, MSFlex.Top + MSFlex.CellTop + 1720, ""
               MSFlex.TextMatrix(i, 17) = frmLogProSelEspecificaciones.vpTexto
            End If
         End If
         Rs.MoveNext
      Loop
   End If
End If
End Sub

Private Sub MSFlex_DblClick()
If Left(MSFlex.TextMatrix(MSFlex.row, 1), 2) = "12" Then
   MSFlex.row = MSFlex.row
   MSFlex.Col = 4
   frmLogProSelEspecificaciones.Inicio MSFlex.Left + MSFlex.CellLeft + 120, MSFlex.Top + MSFlex.CellTop + 1720, MSFlex.TextMatrix(MSFlex.row, 17)
   MSFlex.TextMatrix(MSFlex.row, 17) = frmLogProSelEspecificaciones.vpTexto
End If
End Sub


Function YaEsta(vBSCod As String) As Boolean
Dim i As Integer, N As Integer
YaEsta = False
N = MSFlex.Rows - 1

For i = 1 To N
    If MSFlex.TextMatrix(i, 1) = vBSCod Then
       YaEsta = True
       Exit Function
    End If
Next
End Function

Sub FormaFlex()
Dim i As Integer
MSFlex.Clear
MSFlex.Rows = 2
MSFlex.RowHeight(-1) = 260
MSFlex.RowHeight(0) = 320
MSFlex.RowHeight(1) = 8
MSFlex.ColWidth(0) = 0
MSFlex.ColWidth(1) = 0:         MSFlex.TextMatrix(0, 1) = ""
MSFlex.ColWidth(2) = 2500:      MSFlex.TextMatrix(0, 2) = "Descripción"
MSFlex.ColWidth(3) = 1200:      MSFlex.TextMatrix(0, 3) = "Unidad":   MSFlex.ColAlignment(3) = 1
For i = 1 To 12
    MSFlex.TextMatrix(0, i + 3) = Space(2) + mMes(i)
    MSFlex.ColWidth(i + 3) = 550:   MSFlex.ColAlignment(i + 3) = 4
Next
MSFlex.ColWidth(16) = 900: MSFlex.TextMatrix(0, 16) = "   TOTAL"
MSFlex.ColWidth(17) = 0
End Sub

'*********************************************************************
'PROCEDIMIENTOS DEL FLEX
'*********************************************************************


'Private Sub MSFlex_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyDelete And MSFlex.Col = 3 Then
'   MSFlex.TextMatrix(MSFlex.Row, 3) = ""
'End If
'End Sub

Private Sub MSFlex_KeyPress(KeyAscii As Integer)
If MSFlex.Col >= 3 And MSFlex.Col < 16 And nEditable Then
   EditaFlex MSFlex, txtEdit, KeyAscii
End If
End Sub

Sub EditaFlex(MSFlex As Control, Edt As Control, KeyAscii As Integer)
If InStr("0123456789", Chr(KeyAscii)) Then
Select Case KeyAscii
    Case 0 To 32
         Edt = MSFlex
         Edt.SelStart = 1000
    Case Else
         Edt = Chr(KeyAscii)
         Edt.SelStart = 1
End Select
Edt.Move MSFlex.Left + MSFlex.CellLeft - 15, MSFlex.Top + MSFlex.CellTop - 15, _
         MSFlex.CellWidth, MSFlex.CellHeight
Edt.Visible = True
Edt.SetFocus
End If
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
nKeyAscii = KeyAscii
KeyAscii = DigNumEnt(KeyAscii)
If KeyAscii = Asc(vbCr) Then
   KeyAscii = 0
   txtEdit = FNumero(txtEdit)
End If
End Sub

Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
EditKeyCode MSFlex, txtEdit, KeyCode, Shift
End Sub

Sub EditKeyCode(MSFlex As Control, Edt As Control, KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 27
         Edt.Visible = False
         MSFlex.SetFocus
    Case 13
         MSFlex.SetFocus
    Case 37                     'Izquierda
         MSFlex.SetFocus
         DoEvents
         If MSFlex.Col > 1 Then
            MSFlex.Col = MSFlex.Col - 1
         End If
    Case 39                     'Derecha
         MSFlex.SetFocus
         DoEvents
         If MSFlex.Col < MSFlex.Cols - 1 Then
            MSFlex.Col = MSFlex.Col + 1
         End If
    Case 38
         MSFlex.SetFocus
         DoEvents
         If MSFlex.row > MSFlex.FixedRows + 1 Then
            MSFlex.row = MSFlex.row - 1
         End If
    Case 40
         MSFlex.SetFocus
         DoEvents
         'If MSFlex.Row < MSFlex.FixedRows - 1 Then
         If MSFlex.row < MSFlex.Rows - 1 Then
            MSFlex.row = MSFlex.row + 1
         End If
End Select
End Sub

Private Sub MSFlex_GotFocus()
If txtEdit.Visible = False Then Exit Sub
MSFlex = txtEdit
txtEdit.Visible = False
TotalFila MSFlex.row
'If MSFlex.Row < MSFlex.Rows - 1 Then
'   MSFlex.Row = MSFlex.Row + 1
'End If
End Sub

Private Sub MSFlex_LeaveCell()
If txtEdit.Visible = False Then Exit Sub
MSFlex = txtEdit
txtEdit.Visible = False
TotalFila MSFlex.row
'If MSFlex.Row < MSFlex.Rows - 1 Then
'   MSFlex.Row = MSFlex.Row + 1
'End If
End Sub

Sub TotalFila(i As Integer)
Dim j As Integer, N As Integer
Dim nSuma As Currency
nSuma = 0
For j = 1 To 12
    nSuma = nSuma + VNumero(MSFlex.TextMatrix(i, j + 3))
Next
MSFlex.TextMatrix(i, 16) = nSuma
End Sub

