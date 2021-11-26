VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmLogPlanAnualValorizacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Plan Anual de Adquisiciones y Contrataciones - Asignación de Valor Referencial "
   ClientHeight    =   5175
   ClientLeft      =   735
   ClientTop       =   2490
   ClientWidth     =   10515
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   10515
   Begin VB.TextBox txtEdit 
      BackColor       =   &H00DDFFFE&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5280
      TabIndex        =   4
      Top             =   1620
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.ComboBox cboMoneda 
      BackColor       =   &H00EAFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmLogPlanAnualValorizacion.frx":0000
      Left            =   4260
      List            =   "frmLogPlanAnualValorizacion.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1620
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00DADEDC&
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   120
      TabIndex        =   1
      Top             =   -30
      Width           =   10275
      Begin VB.CommandButton cmdConsolidar 
         Caption         =   "Consolidar &Requerimientos"
         Height          =   375
         Left            =   5880
         TabIndex        =   19
         Top             =   210
         Width           =   2115
      End
      Begin VB.CommandButton cmdEjeVal 
         Caption         =   "Asignar &Valor Referencial"
         Height          =   375
         Left            =   8040
         TabIndex        =   18
         Top             =   210
         Width           =   2115
      End
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
         Left            =   4920
         TabIndex        =   0
         Text            =   "2005"
         Top             =   225
         Width           =   675
      End
      Begin VB.Label lblTitulo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Plan Anual de Adquisiciones y Contrataciones del Año"
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
         Height          =   195
         Left            =   180
         TabIndex        =   2
         Top             =   300
         Width           =   4620
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFlex 
      Height          =   3435
      Left            =   120
      TabIndex        =   5
      Top             =   660
      Width           =   10275
      _ExtentX        =   18124
      _ExtentY        =   6059
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   10
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
      _Band(0).Cols   =   10
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00EAFFFF&
      ForeColor       =   &H80000008&
      Height          =   570
      Left            =   120
      TabIndex        =   6
      Top             =   4020
      Width           =   10275
      Begin VB.TextBox txtTotalDol 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   8460
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   180
         Width           =   1575
      End
      Begin VB.TextBox txtTotalSol 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6180
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   180
         Width           =   1575
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL   S/.                                US$"
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
         Height          =   195
         Left            =   4980
         TabIndex        =   9
         Top             =   240
         Width           =   3345
      End
   End
   Begin VB.Frame fraBot1 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   120
      TabIndex        =   10
      Top             =   4530
      Width           =   10275
      Begin VB.CommandButton cmdAgencias 
         Caption         =   "Agencias"
         Height          =   375
         Left            =   1320
         TabIndex        =   22
         Top             =   120
         Width           =   1275
      End
      Begin VB.CommandButton cmdAreas 
         Caption         =   "Areas"
         Height          =   375
         Left            =   0
         TabIndex        =   21
         Top             =   120
         Width           =   1275
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   375
         Left            =   9060
         TabIndex        =   11
         Top             =   150
         Width           =   1215
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H00DADEDC&
         ForeColor       =   &H80000008&
         Height          =   580
         Left            =   5160
         TabIndex        =   12
         Top             =   0
         Width           =   3855
         Begin VB.CommandButton cmdImprimir 
            Caption         =   "Imprimir"
            Height          =   375
            Left            =   2520
            TabIndex        =   20
            Top             =   150
            Width           =   1275
         End
         Begin VB.OptionButton opSol 
            BackColor       =   &H00DADEDC&
            Caption         =   "Soles"
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
            Height          =   195
            Left            =   240
            TabIndex        =   14
            Top             =   260
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton opDol 
            BackColor       =   &H00DADEDC&
            Caption         =   "Dólares"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00185B11&
            Height          =   195
            Left            =   1320
            TabIndex        =   13
            Top             =   260
            Width           =   1095
         End
      End
   End
   Begin VB.Frame fraBot2 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   7500
      TabIndex        =   15
      Top             =   4680
      Visible         =   0   'False
      Width           =   2895
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   1680
         TabIndex        =   17
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar"
         Height          =   375
         Left            =   420
         TabIndex        =   16
         Top             =   0
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmLogPlanAnualValorizacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sSQL As String
Dim pbEstaValorizando As Boolean
Dim nFILA As Integer
Dim bRecuperado As Boolean

Private Sub cmdAgencias_Click()
frmLogPlanAgeArea.Inicio 2, 0, CInt(txtAnio.Text)
End Sub

Private Sub cmdAreas_Click()
frmLogPlanAgeArea.Inicio 1, 0, CInt(txtAnio.Text)
End Sub

Private Sub cmdCancelar_Click()
Dim i As Integer, n As Integer

n = MSFlex.Rows - 1
For i = 1 To n
    MSFlex.TextMatrix(i, 5) = MSFlex.TextMatrix(i, 8)
    MSFlex.TextMatrix(i, 6) = MSFlex.TextMatrix(i, 9)
Next
SumaTotal
pbEstaValorizando = False
fraBot1.Visible = True
fraBot2.Visible = False
End Sub

Private Sub Form_Load()
CentraForm Me
txtAnio.Text = Year(gdFecSis) + 1
FormaFlexValor
pbEstaValorizando = False
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub cmdEjeVal_Click()
pbEstaValorizando = True
fraBot1.Visible = False
fraBot2.Visible = True
MSFlex.row = 1
MSFlex.Col = 6
MSFlex.SetFocus
End Sub

Private Sub cmdGrabar_Click()
Dim oConn As New DConecta, i As Integer, n As Integer
Dim nPrecio As Currency, nMoneda As Integer, cBSCod As String

n = MSFlex.Rows - 1

If MsgBox("¿ Está seguro de grabar los valores indicados ?" + Space(10), vbQuestion + vbYesNo, "Confirme") = vbYes Then

   If oConn.AbreConexion Then
      For i = 1 To n
          If Len(MSFlex.TextMatrix(i, 5)) > 0 And VNumero(MSFlex.TextMatrix(i, 6)) > 0 Then
             nMoneda = 0
             Select Case MSFlex.TextMatrix(i, 5)
                 Case "S/."
                      nMoneda = 1
                 Case "US$"
                      nMoneda = 2
             End Select
             nPrecio = VNumero(MSFlex.TextMatrix(i, 6))
             cBSCod = MSFlex.TextMatrix(i, 1)
             
             sSQL = "UPDATE LogPlanAnualReqDetalle SET nMoneda = " & nMoneda & ", " & _
                    "                                  nPrecioUnitario = " & nPrecio & " " & _
                    " WHERE cBSCod = '" & cBSCod & "' "
             oConn.Ejecutar sSQL
             
          End If
      Next
   End If
   pbEstaValorizando = False
   fraBot1.Visible = True
   fraBot2.Visible = False
End If
End Sub

Sub ListaRequerimientosNoAprobados(ByVal pnAnio As Integer)
Dim oConn As New DConecta, sSQL As String, rs As New ADODB.Recordset
Dim v As Double, f As Integer
Dim cArchivo As String

'sSQL = "select * from " & _
'       " (select r.nPlanReqNro,nNro=count(*) from LogPlanAnualAprobacion a " & _
'       "  inner join LogPlanAnualReq r on a.nPlanReqNro = r.nPlanReqNro " & _
'       "  Where r.nAnio = " & pnAnio & " and r.nEstado=1 group by r.nPlanReqNro) x left join " & _
'       " (select nPlanReqNro,nApro=count(*) from LogPlanAnualAprobacion where nEstadoAprobacion=1 group by nPlanReqNro) y on x.nPlanReqNro = y.nPlanReqNro " & _
'       "  Where x.nNro <> Y.nApro  "
       
sSQL = "select r.nPlanReqNro,r.cPersCod, p.cPersNombre, a.cAgeDescripcion, t.cRHCargoDescripcion " & _
"  from LogPlanAnualReq r inner join Persona p on r.cPersCod = p.cPersCod " & _
"       inner join Agencias a on r.cRHAgeCod = a.cAgeCod " & _
"       inner join RHCargosTabla t on r.cRHCargoCod = t.cRHCargoCod " & _
" where r.nPlanReqNro in  (select x.nPlanReqNro from " & _
" (select r.nPlanReqNro,nNro=count(*) from LogPlanAnualAprobacion a   inner join LogPlanAnualReq r on a.nPlanReqNro = r.nPlanReqNro " & _
"   Where r.nAnio = 2006 and r.nEstado=1 group by r.nPlanReqNro) x left join " & _
" (select nPlanReqNro,nApro=count(*) from LogPlanAnualAprobacion " & _
"   where nEstadoAprobacion=1 group by nPlanReqNro) y on x.nPlanReqNro = y.nPlanReqNro " & _
"   Where x.nNro <> Y.nApro)"

If oConn.AbreConexion Then
   Set rs = oConn.CargaRecordSet(sSQL)
   If Not rs.EOF Then
   
      f = FreeFile
      cArchivo = App.path + "\replan.txt"
      Open cArchivo For Output As #f

      Print #f, Centra("LISTA DE REQUERIMIENTOS SIN APROBAR", 70)
      Print #f, ""
   
      Do While Not rs.EOF
         Print #f, rs!cPersCod + " " + JIZQ(rs!cPersNombre, 40) + " " + JIZQ(rs!cAgeDescripcion, 25) + " " + rs!cRHCargoDescripcion
         rs.MoveNext
      Loop
      
      Close #f
      v = Shell("notepad " & cArchivo & "", vbNormalFocus)
      
   End If
End If
End Sub

Private Sub cmdConsolidar_Click()
Dim rs As New ADODB.Recordset, oConn As New DConecta
Dim YaValorizo As Boolean, nAnio As Integer, nResp As Integer
Dim k As Integer

nAnio = CInt(VNumero(txtAnio.Text))

If RequerimientosNoAprobados(nAnio) > 0 Then
   
   k = MsgBox("Existen requerimientos sin aprobar ¿ Continuar ? " + Space(10) + vbCrLf + "   Presione SI para consolidar solo requerimientos aprobados" + Space(10) + vbCrLf + _
          "   Presione NO para visualizar los requerimientos sin aprobar " + Space(10), vbQuestion + vbYesNoCancel, "Verifique requerimientos")
   If k = vbNo Then
      ListaRequerimientosNoAprobados nAnio
      Exit Sub
   End If
   If k = vbCancel Then
      Exit Sub
   End If
End If

sSQL = "select top 1 nPrecioUnitario from LogPlanAnualReqDetalle where nAnio = " & nAnio & " and nPrecioUnitario>0"
If oConn.AbreConexion Then
   Set rs = oConn.CargaRecordSet(sSQL)
   oConn.CierraConexion
End If

If Not rs.EOF Then
   YaValorizo = True
Else
   YaValorizo = False
End If

Set rs = Nothing

If YaValorizo Then
   nResp = MsgBox(Space(30) + "A T E N C I O N" + vbCrLf + "El Consolidado ya tiene valores referenciales asignados" + Space(10) + vbCrLf + vbCrLf + Space(20) + "¿ Consolidar requerimientos ?" + vbCrLf + vbCrLf + _
                  "Presione SI para consolidar requerimientos nuevamente" + Space(10) + vbCrLf + _
                  "Presione NO para recuperar el consolidado anterior   " + Space(10), vbQuestion + vbYesNoCancel + vbDefaultButton3, "Aviso")

   If nResp = vbYes Then
      ConsolidacionRequerimientos
      bRecuperado = False
   ElseIf nResp = vbNo Then
      RecuperaConsolidado
      bRecuperado = True
   ElseIf nResp = vbCancel Then
   
   End If
Else
   ConsolidacionRequerimientos
End If
End Sub

Sub ConsolidacionRequerimientos()
Dim rs As New ADODB.Recordset
Dim oConn As New DConecta, i As Integer
Dim nAnio As Integer
nAnio = CInt(VNumero(txtAnio.Text))

FormaFlexValor
txtTotalSol = ""
txtTotalDol = ""

 sSQL = "select d.cBSCod, g.cBSDescripcion, u.cUnidad , nTotal = sum(nMes01+nMes02+nMes03+nMes04+nMes05+nMes06+nMes07+nMes08+nMes09+nMes10+nMes11+nMes12) " & _
        "  from LogPlanAnualReqDetalle d " & _
        "       inner join BienesServicios g on d.cBSCod = g.cBSCod " & _
        "       inner join (select nConsValor as nBSUnidad, cConsDescripcion as cUnidad from Constante where nConsCod = 9097) u on g.nBSUnidad = u.nBSUnidad " & _
        " WHERE d.nAnio = " & nAnio & " " & _
        "  group by d.cBSCod, g.cBSDescripcion, u.cUnidad "
       
   If oConn.AbreConexion Then
      Set rs = oConn.CargaRecordSet(sSQL)
      oConn.CierraConexion
   End If
   If Not rs.EOF Then
      i = 0
      Do While Not rs.EOF
         i = i + 1
         InsRow MSFlex, i
         MSFlex.TextMatrix(i, 1) = rs!cBSCod
         MSFlex.TextMatrix(i, 2) = rs!cBSDescripcion
         MSFlex.TextMatrix(i, 3) = rs!nTotal
         MSFlex.TextMatrix(i, 4) = rs!cUnidad
         MSFlex.TextMatrix(i, 5) = "S/."
         MSFlex.TextMatrix(i, 6) = GetPrecioUnitario(nAnio - 1, 2, rs!cBSCod)
         MSFlex.TextMatrix(i, 7) = VNumero(MSFlex.TextMatrix(i, 3)) * VNumero(MSFlex.TextMatrix(i, 6))
         DoEvents
         rs.MoveNext
      Loop
      SumaTotal
    End If
End Sub

Sub RecuperaConsolidado()
Dim rs As New ADODB.Recordset
Dim oConn As New DConecta, i As Integer
Dim nAnio As Integer
nAnio = CInt(VNumero(txtAnio.Text))

FormaFlexValor
txtTotalSol = ""
txtTotalDol = ""

sSQL = "select d.cBSCod, g.cBSDescripcion, d.nMoneda, d.nPrecioUnitario," & _
       "            u.cUnidad , nTotal = sum(nMes01+nMes02+nMes03+nMes04+nMes05+nMes06+nMes07+nMes08+nMes09+nMes10+nMes11+nMes12) " & _
       "  from LogPlanAnualReqDetalle d " & _
       "       inner join BienesServicios g on d.cBSCod = g.cBSCod " & _
       "       inner join (select nConsValor as nBSUnidad, cConsDescripcion as cUnidad from Constante where nConsCod = 9097) u on g.nBSUnidad = u.nBSUnidad " & _
       " WHERE d.nAnio = " & nAnio & " and d.nEstado = 1 " & _
       "  group by d.cBSCod, g.cBSDescripcion, u.cUnidad, d.nMoneda, d.nPrecioUnitario "
      
    If oConn.AbreConexion Then
       Set rs = oConn.CargaRecordSet(sSQL)
       oConn.CierraConexion
    End If
    
    If Not rs.EOF Then
       i = 0
       Do While Not rs.EOF
          i = i + 1
          InsRow MSFlex, i
          MSFlex.TextMatrix(i, 0) = Format(i, "00")
          MSFlex.TextMatrix(i, 1) = rs!cBSCod
          MSFlex.TextMatrix(i, 2) = rs!cBSDescripcion
          MSFlex.TextMatrix(i, 3) = rs!nTotal
          MSFlex.TextMatrix(i, 4) = rs!cUnidad
          MSFlex.TextMatrix(i, 5) = IIf(rs!nMoneda = 1, "S/.", "US$")
          MSFlex.TextMatrix(i, 6) = FNumero(rs!nPrecioUnitario)
          MSFlex.TextMatrix(i, 7) = FNumero(VNumero(MSFlex.TextMatrix(i, 3)) * VNumero(MSFlex.TextMatrix(i, 6)))
          'Para restaurar en caso cancele operación
          MSFlex.TextMatrix(i, 8) = IIf(rs!nMoneda = 1, "S/.", "US$")
          MSFlex.TextMatrix(i, 9) = FNumero(rs!nPrecioUnitario)
          DoEvents
          rs.MoveNext
       Loop
       SumaTotal
    End If
End Sub

Sub FormaFlexValor()
MSFlex.Clear
MSFlex.Rows = 2
MSFlex.RowHeight(-1) = 280
MSFlex.RowHeight(0) = 320
MSFlex.RowHeight(1) = 8
MSFlex.ColWidth(0) = 350:       MSFlex.TextMatrix(0, 0) = "Nº":     MSFlex.ColAlignment(0) = 4
MSFlex.ColWidth(1) = 850:       MSFlex.TextMatrix(0, 1) = "Codigo"
MSFlex.ColWidth(2) = 4000:      MSFlex.TextMatrix(0, 2) = "Descripción"
MSFlex.ColWidth(3) = 700:       MSFlex.TextMatrix(0, 3) = "Cantidad":     MSFlex.ColAlignment(3) = 4
MSFlex.ColWidth(4) = 1500:      MSFlex.TextMatrix(0, 4) = " U. Medida":   MSFlex.ColAlignment(4) = 1
MSFlex.ColWidth(5) = 600:       MSFlex.TextMatrix(0, 5) = "Moneda":       MSFlex.ColAlignment(5) = 4
MSFlex.ColWidth(6) = 950:       MSFlex.TextMatrix(0, 6) = " Precio Unit"
MSFlex.ColWidth(7) = 1000:       MSFlex.TextMatrix(0, 7) = " Sub-Total"
MSFlex.ColWidth(8) = 0
MSFlex.ColWidth(9) = 0
End Sub

Private Sub MSFlex_DblClick()
If MSFlex.Col = 5 And pbEstaValorizando Then
   cboMoneda.Visible = True
   cboMoneda.ListIndex = 0
   cboMoneda.Move MSFlex.Left + MSFlex.CellLeft - 30, MSFlex.Top + MSFlex.CellTop - 30, MSFlex.CellWidth + 30
End If
End Sub

Private Sub MSFlex_GotFocus()
If cboMoneda.Visible Then
   MSFlex.TextMatrix(MSFlex.row, 5) = cboMoneda.Text
   cboMoneda.Visible = False
End If
If txtEdit.Visible = False Then Exit Sub
MSFlex = txtEdit
txtEdit.Visible = False
SumaTotal
End Sub

Private Sub MSFlex_LeaveCell()
If cboMoneda.Visible Then
   MSFlex.TextMatrix(MSFlex.row, 5) = cboMoneda.Text
   cboMoneda.Visible = False
End If
If txtEdit.Visible = False Then Exit Sub
MSFlex = FNumero(txtEdit)
txtEdit.Visible = False
SumaTotal
End Sub


'*********************************************************************
'PROCEDIMIENTOS DEL FLEX
'*********************************************************************

Private Sub MSFlex_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete And MSFlex.Col = 6 And pbEstaValorizando Then
   MSFlex.TextMatrix(MSFlex.row, 6) = ""
   SumaTotal
End If
End Sub

Private Sub MSFlex_KeyPress(KeyAscii As Integer)
If MSFlex.Col = 6 And pbEstaValorizando Then
   If Len(MSFlex.TextMatrix(MSFlex.row, 5)) = 0 Then
      MsgBox "Debe seleccionar un tipo de moneda..." + Space(10), vbInformation
   Else
      EditaFlex MSFlex, txtEdit, KeyAscii
   End If
End If

If MSFlex.Col = 5 And pbEstaValorizando Then
   nFILA = MSFlex.row
   cboMoneda.Move MSFlex.Left + MSFlex.CellLeft - 30, MSFlex.Top + MSFlex.CellTop - 30, MSFlex.CellWidth + 30
   cboMoneda.Visible = True
   cboMoneda.SetFocus
End If
End Sub

Sub EditaFlex(MSFlex As Control, Edt As Control, KeyAscii As Integer)
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
End Sub

Private Sub txtAnio_GotFocus()
SelTexto txtAnio
End Sub

Private Sub txtAnio_KeyPress(KeyAscii As Integer)
nKeyAscii = KeyAscii
KeyAscii = DigNumEnt(KeyAscii)
If nKeyAscii = 13 Then
   cmdConsolidar.SetFocus
End If
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
If KeyAscii = Asc(vbCr) Then
   KeyAscii = 0
   'txtEdit = FNumero(txtEdit)
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
         If MSFlex.row < MSFlex.Rows - 1 Then
            MSFlex.row = MSFlex.row + 1
         End If
End Select
End Sub

Sub SumaTotal()
Dim i As Integer, n As Integer, nSumaSol As Currency, nSumaDol As Currency

n = MSFlex.Rows - 1
nSumaSol = 0
nSumaDol = 0
For i = 1 To n
    MSFlex.TextMatrix(i, 7) = FNumero(VNumero(MSFlex.TextMatrix(i, 6)) * VNumero(MSFlex.TextMatrix(i, 3)))
    
    If VNumero(MSFlex.TextMatrix(i, 6)) > 0 Then
       Select Case MSFlex.TextMatrix(i, 5)
           Case "S/."
                nSumaSol = nSumaSol + VNumero(MSFlex.TextMatrix(i, 7))
           Case "US$"
                nSumaDol = nSumaDol + VNumero(MSFlex.TextMatrix(i, 7))
       End Select
    Else
       MSFlex.TextMatrix(i, 7) = ""
    End If

Next
txtTotalSol = FNumero(nSumaSol)
txtTotalDol = FNumero(nSumaDol)
End Sub

Private Sub cmdImprimir_Click()
Dim i As Integer, n As Integer, f As Integer, v As Variant
Dim nTotal As Currency, cMoneda As String

If opSol.value Then
   cMoneda = "S/."
End If
If opDol.value Then
   cMoneda = "US$"
End If

n = MSFlex.Rows - 1
f = FreeFile
Open App.path + "\Val01.txt" For Output As #f
nTotal = 0
Print #f, ""
Print #f, Space(18) + "PLAN ANUAL DE ADQUISICIONES Y CONTRATACIONES"
Print #f, Space(18) + "  VALORIZACION DE REQUERIMIENTOS EN " + IIf(opSol.value, "SOLES", "DOLARES")
Print #f, ""
Print #f, String(100, "=")
Print #f, "No Codigo     Descripcion" + Space(35) + "Cantidad   Precio Unitario   Sub - Total"
Print #f, String(100, "-")
For i = 1 To n
    If MSFlex.TextMatrix(i, 5) = cMoneda Then
       Print #f, Format(i, "00") + " " + MSFlex.TextMatrix(i, 1) + " " + JIZQ(MSFlex.TextMatrix(i, 2), 40, "_") + " " + JDER(MSFlex.TextMatrix(i, 3), 8) + " " + JIZQ(MSFlex.TextMatrix(i, 4), 12) + " " + IIf(VNumero(MSFlex.TextMatrix(i, 6)) > 0, JDER(FNumero(VNumero(MSFlex.TextMatrix(i, 6))), 9) + "  " + JDER(FNumero(VNumero(MSFlex.TextMatrix(i, 3)) * VNumero(MSFlex.TextMatrix(i, 6))), 12), "")
       nTotal = nTotal + VNumero(MSFlex.TextMatrix(i, 3)) * VNumero(MSFlex.TextMatrix(i, 6))
    End If
Next
Print #f, String(100, "-")
Print #f, Space(65) + "TOTAL " + Space(10) + cMoneda + JDER(FNumero(nTotal), 15)
Print #f, String(100, "=")
Close #f
v = Shell("notepad.exe " + App.path + "\val01.txt", vbNormalFocus)
If MsgBox(" ¿ Exportar a Excel ? " + Space(10), vbQuestion + vbYesNo, "Exportación de datos") = vbYes Then
   ExportaExcel
End If
End Sub

Sub ExportaExcel()
Dim i As Integer, n As Integer, k As Integer
Dim appExcel As New Excel.Application
Dim wbExcel As Excel.Workbook

Set wbExcel = appExcel.Workbooks.Add

k = 1
n = MSFlex.Rows - 1
For i = 1 To n
    'If MSFlex.TextMatrix(i, 5) =  Then
       k = k + 1
       wbExcel.Worksheets(1).Range("A" + CStr(k)).value = MSFlex.TextMatrix(i, 0)
       wbExcel.Worksheets(1).Range("B" + CStr(k)).value = MSFlex.TextMatrix(i, 1)
       wbExcel.Worksheets(1).Range("C" + CStr(k)).value = MSFlex.TextMatrix(i, 2)
       wbExcel.Worksheets(1).Range("D" + CStr(k)).value = MSFlex.TextMatrix(i, 3)
       wbExcel.Worksheets(1).Range("E" + CStr(k)).value = MSFlex.TextMatrix(i, 4)
       If VNumero(MSFlex.TextMatrix(i, 6)) > 0 Then
          wbExcel.Worksheets(1).Range("F" + CStr(k)).value = MSFlex.TextMatrix(i, 6)
          wbExcel.Worksheets(1).Range("G" + CStr(k)).value = VNumero(MSFlex.TextMatrix(i, 3)) * VNumero(MSFlex.TextMatrix(i, 6))
       End If
    'End If
Next

appExcel.Application.Visible = True
appExcel.Windows(1).Visible = True
End Sub

