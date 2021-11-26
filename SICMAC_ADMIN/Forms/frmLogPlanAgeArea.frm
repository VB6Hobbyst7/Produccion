VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmLogPlanAgeArea 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5655
   ClientLeft      =   1695
   ClientTop       =   2250
   ClientWidth     =   8130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   8130
   Begin VB.CommandButton cmdDetalle 
      Caption         =   "Agencias y Areas"
      Height          =   375
      Left            =   2340
      TabIndex        =   7
      Top             =   5220
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   6840
      TabIndex        =   5
      Top             =   5220
      Width           =   1215
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   5580
      TabIndex        =   4
      Top             =   5220
      Width           =   1215
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFlex 
      Height          =   2415
      Left            =   60
      TabIndex        =   0
      Top             =   2700
      Width           =   7995
      _ExtentX        =   14102
      _ExtentY        =   4260
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      BackColorSel    =   6956042
      ForeColorSel    =   -2147483643
      BackColorBkg    =   -2147483643
      GridColor       =   -2147483633
      FocusRect       =   0
      AllowUserResizing=   3
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
   Begin VB.CommandButton cmdConsolida 
      Caption         =   "Consolidar Requerimientos"
      Height          =   375
      Left            =   60
      TabIndex        =   3
      Top             =   5220
      Width           =   2235
   End
   Begin VB.Frame fraSelector 
      Height          =   2535
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   7995
      Begin VB.CheckBox chkTodos 
         Height          =   195
         Left            =   280
         TabIndex        =   6
         Top             =   360
         Width           =   200
      End
      Begin MSComctlLib.ListView lsvObj 
         Height          =   2150
         Left            =   180
         TabIndex        =   2
         Top             =   300
         Width           =   7635
         _ExtentX        =   13467
         _ExtentY        =   3784
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Width           =   10583
         EndProperty
      End
   End
End
Attribute VB_Name = "frmLogPlanAgeArea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nTipo As Integer, nPlanAnualNro As Integer, nAnio As Integer

Public Sub Inicio(ByVal pnTipo As Integer, ByVal pnPlanAnualNro As Integer, ByVal pnAnio As Integer)
nTipo = pnTipo
nPlanAnualNro = pnPlanAnualNro
nAnio = pnAnio
Me.Show 1
End Sub

Private Sub chkTodos_Click()
Dim i As Integer, n As Integer

n = lsvObj.ListItems.Count
If chkTodos.value = 0 Then
   For i = 1 To n
       lsvObj.ListItems(i).Checked = False
   Next
Else
   For i = 1 To n
       lsvObj.ListItems(i).Checked = True
   Next
End If
End Sub

Private Sub cmdImprimir_Click()
Dim f As String, i As Integer, n As Integer
Dim v As Variant

f = FreeFile
Open App.path + "\temp.txt" For Output As #f

Print #f, ""
Print #f, Space(12) + "CONSOLIDACION POR AGENCIA DE REQUERIMIENTOS PARA VALORIZACION AL " + CStr(Date)
Print #f, ""
Print #f, String(100, "=")
Print #f, "   Agencia" + Space(16) + "Requerimiento" + Space(52) + "Cantidad"
Print #f, String(100, "-")
n = MSFlex.Rows - 1
For i = 1 To n
    Print #f, MSFlex.TextMatrix(i, 1) + " " + JIZQ(MSFlex.TextMatrix(i, 2), 22) + " " + JIZQ(MSFlex.TextMatrix(i, 3), 60) + " " + JDER(Format(MSFlex.TextMatrix(i, 5), "###,##0"), 12)
Next
Print #f, String(100, "=")
Print #f, "Usuario: [" + gsCodUser + "]  " + CStr(Date) + " - " + CStr(Time)
Close #f
v = Shell("notepad.exe " + App.path + "\temp.txt", vbNormalFocus)
End Sub

Private Sub Form_Load()
CentraForm Me
FormaFlex
Select Case nTipo
    Case 1
         GeneraListaAreas
         Me.Caption = "Distribución de requerimientos por Areas"
         fraSelector.Caption = "Seleccione Areas"
    Case 2
         GeneraListaAgencias
         cmdDetalle.Visible = True
         Me.Caption = "Distribución de requerimientos por Agencias"
         fraSelector.Caption = "Seleccione Agencias"
End Select
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Sub GeneraListaAreas()
Dim rs As New ADODB.Recordset
Dim oConn As New DConecta
Dim i As Integer, sSQL As String

sSQL = "Select cAreaCod, cAreaDescripcion from DBCmactAux..Areas where nAreaEstado=1 order by cAreaDescripcion"

If oConn.AbreConexion Then
   Set rs = oConn.CargaRecordSet(sSQL)
   If Not rs.EOF Then
      i = 0
      Do While Not rs.EOF
         i = i + 1
         lsvObj.ListItems.Add i
         lsvObj.ListItems(i).SubItems(1) = rs!cAreaCod
         lsvObj.ListItems(i).SubItems(2) = rs!cAreaDescripcion
         rs.MoveNext
      Loop
   End If
End If
End Sub

Sub GeneraListaAgencias()
Dim rs As New ADODB.Recordset
Dim oConn As New DConecta
Dim i As Integer, sSQL As String

sSQL = "Select cAgeCod, cAgeDescripcion from DBCmactAux..Agencias where nEstado=1 order by cAgeCod"

If oConn.AbreConexion Then
   Set rs = oConn.CargaRecordSet(sSQL)
   If Not rs.EOF Then
      i = 0
      Do While Not rs.EOF
         i = i + 1
         lsvObj.ListItems.Add i
         lsvObj.ListItems(i).SubItems(1) = rs!cAgeCod
         lsvObj.ListItems(i).SubItems(2) = rs!cAgeDescripcion
         rs.MoveNext
      Loop
   End If
End If
End Sub

Private Sub cmdConsolida_Click()
Dim rs As New ADODB.Recordset, oConn As New DConecta
Dim i As Integer, n As Integer, sSQL As String
Dim cLista As String

cLista = ""
n = lsvObj.ListItems.Count
For i = 1 To n
    If lsvObj.ListItems(i).Checked Then
       cLista = cLista + "'" + lsvObj.ListItems(i).SubItems(1) + "',"
    End If
Next

If Len(Trim(cLista)) = 0 Then
   MsgBox "Debe seleccionar al menos un Item..." + Space(10), vbInformation
   Exit Sub
End If

cLista = Left(cLista, Len(cLista) - 1)

FormaFlex

If nPlanAnualNro = 0 Then
   'Consolidación simple ------ sin plan anual ------
   Select Case nTipo
     Case 1
          sSQL = "select r.cRHAreaCod as cCodigo,g.cBSDescripcion as cConcepto, a.cAreaDescripcion as cDescripcion, d.nMoneda, d.nPrecioUnitario, u.cUnidad,  " & _
          " nMonto = Sum(nMes01 + nMes02 + nMes03 + nMes04 + nMes05 + nMes06 + nMes07 + nMes08 + nMes09 + nMes10 + nMes11 + nMes12) " & _
          " from LogPlanAnualReq r inner join LogPlanAnualReqDetalle d on r.nPlanReqNro = d.nPlanReqNro " & _
          " inner join DBCmactAux..Areas a on r.cRHAreaCod = a.cAreaCod " & _
          " inner join LogProSelBienesServicios g on d.cBSCod = g.cProSelBSCod " & _
          " inner join (select nConsValor as nBSUnidad, cConsDescripcion as cUnidad from Constante where nConsCod = 9097) u on g.nBSUnidad = u.nBSUnidad " & _
          " WHERE d.nAnio = " & nAnio & "  and r.cRHAreaCod  in (" & cLista & ") " & _
          " group by r.cRHAreaCod,a.cAreaDescripcion,g.cBSDescripcion,u.cUnidad, d.nMoneda, d.nPrecioUnitario"

     Case 2
          sSQL = "select r.cRHAgeCod as cCodigo,g.cBSDescripcion as cConcepto, a.cAgeDescripcion as cDescripcion, d.nMoneda, d.nPrecioUnitario, u.cUnidad,  " & _
          " nMonto = Sum(nMes01 + nMes02 + nMes03 + nMes04 + nMes05 + nMes06 + nMes07 + nMes08 + nMes09 + nMes10 + nMes11 + nMes12) " & _
          " from LogPlanAnualReq r inner join LogPlanAnualReqDetalle d on r.nPlanReqNro = d.nPlanReqNro " & _
          " inner join DBCmactAux..Agencias a on r.cRHAgeCod = a.cAgeCod " & _
          " inner join LogProSelBienesServicios g on d.cBSCod = g.cProSelBSCod " & _
          " inner join (select nConsValor as nBSUnidad, cConsDescripcion as cUnidad from Constante where nConsCod = 9097) u on g.nBSUnidad = u.nBSUnidad " & _
          " WHERE d.nAnio = " & nAnio & "  and r.cRHAgeCod  in (" & cLista & ") " & _
          " group by r.cRHAgeCod, a.cAgeDescripcion,g.cBSDescripcion, u.cUnidad, d.nMoneda, d.nPrecioUnitario "

   End Select
   
Else
   'Consolidación simple ---- desde el plan anual ----
   Select Case nTipo
    Case 1
         sSQL = "select r.cRHAreaCod as cCodigo, a.cAreaDescripcion as cDescripcion, " & _
         "      cBSGrupoCod=coalesce(g.cBSGrupoCod,''), " & _
         "      nObjetoCod = convert(tinyint,substring(p.cProSelBSCod,2,1)), cConcepto=g.cBSGrupoDescripcion, p.nMoneda, sum(p.nPrecioUnitario*(p.nMes01+p.nMes02+p.nMes03+p.nMes04+p.nMes05+p.nMes06+p.nMes07+p.nMes08+p.nMes09+p.nMes10+p.nMes11+p.nMes12)) as nMonto " & _
         " from LogPlanAnualReq r inner join LogPlanAnualReqDetalle p on r.nPlanReqNro = p.nPlanReqNro " & _
         "      inner join DBCmactAux..Areas a on r.cRHAreaCod = a.cAreaCod " & _
         "      left join LogProSelBienesServicios b on p.cBSCod = b.cProSelBSCod " & _
         "      left join BSGrupos g on b.cBSGrupoCod = g.cBSGrupoCod " & _
         " Where p.nPlanAnualNro = " & nPlanAnualNro & " And P.nEstado = 1 " & _
         "       and r.cRHAreaCod in (" & cLista & ") " & _
         " group by r.cRHAreaCod, a.cAreaDescripcion,g.cBSGrupoCod,convert(tinyint,substring(p.cProSelBSCod,2,1)),g.cBSGrupoDescripcion,p.nMoneda " & _
         " order by r.cRHAreaCod,g.cBSGrupoCod "
    Case 2
         sSQL = "select r.cRHAgeCod as cCodigo, a.cAgeDescripcion as cDescripcion, " & _
         "      cBSGrupoCod=coalesce(g.cBSGrupoCod,''), " & _
         "      nObjetoCod = convert(tinyint,substring(p.cProSelBSCod,2,1)), cConcepto=g.cBSGrupoDescripcion, p.nMoneda, sum(p.nPrecioUnitario*(p.nMes01+p.nMes02+p.nMes03+p.nMes04+p.nMes05+p.nMes06+p.nMes07+p.nMes08+p.nMes09+p.nMes10+p.nMes11+p.nMes12)) as nMonto " & _
         " from LogPlanAnualReq r inner join LogPlanAnualReqDetalle p on r.nPlanReqNro = p.nPlanReqNro " & _
         "      inner join DBCmactAux..Agencias a on r.cRHAgeCod = a.cAgeCod " & _
         "      left join LogProSelBienesServicios b on p.cBSCod = b.cProSelBSCod " & _
         "      left join BSGrupos g on b.cBSGrupoCod = g.cBSGrupoCod " & _
         " Where P.nPlanAnualNro = " & nPlanAnualNro & " And P.nEstado = 1 " & _
         "       and r.cRHAgeCod in (" & cLista & ") " & _
         " group by r.cRHAgeCod, a.cAgeDescripcion,g.cBSGrupoCod,convert(tinyint,substring(p.cProSelBSCod,2,1)),g.cBSGrupoDescripcion,p.nMoneda " & _
         " order by r.cRHAgeCod,g.cBSGrupoCod"
  End Select
End If

If oConn.AbreConexion Then
   Set rs = oConn.CargaRecordSet(sSQL)
   If Not rs.EOF Then
      i = 0
      Do While Not rs.EOF
         i = i + 1
         InsRow MSFlex, i
         MSFlex.TextMatrix(i, 1) = rs!cCodigo
         MSFlex.TextMatrix(i, 2) = rs!cDescripcion
         MSFlex.TextMatrix(i, 3) = rs!cConcepto
         If nPlanAnualNro = 0 Then
            MSFlex.TextMatrix(i, 5) = rs!nMonto
         Else
            MSFlex.TextMatrix(i, 4) = IIf(rs!nMoneda = 1, "S/.", "US$")
            MSFlex.TextMatrix(i, 5) = FNumero(rs!nMonto)
         End If
         rs.MoveNext
      Loop
   End If
End If
End Sub

Sub FormaFlex()
MSFlex.Clear
MSFlex.Rows = 2
MSFlex.RowHeight(0) = 420
MSFlex.RowHeight(1) = 8
MSFlex.ColWidth(0) = 0
MSFlex.ColWidth(1) = 360:     MSFlex.TextMatrix(0, 1) = " ID":     MSFlex.ColAlignment(1) = 4
MSFlex.ColWidth(2) = 3300:    MSFlex.TextMatrix(0, 2) = " Descripción"
If nPlanAnualNro = 0 Then
   MSFlex.ColWidth(3) = 3000:    MSFlex.TextMatrix(0, 3) = " Grupo B/S"
   MSFlex.ColWidth(4) = 0
   MSFlex.ColWidth(5) = 1000:    MSFlex.TextMatrix(0, 5) = "  Cantidad"
Else
   MSFlex.ColWidth(3) = 2500:    MSFlex.TextMatrix(0, 3) = " Grupo B/S"
   MSFlex.ColWidth(4) = 400:     MSFlex.TextMatrix(0, 4) = "Moneda":     MSFlex.ColAlignment(4) = 4
   MSFlex.ColWidth(5) = 1100:    MSFlex.TextMatrix(0, 5) = "   Monto"
   MSFlex.ColWidth(6) = 0:       MSFlex.TextMatrix(0, 6) = "Objeto"
End If
End Sub


Private Sub cmdDetalle_Click()
Dim rs As New ADODB.Recordset, oConn As New DConecta
Dim i As Integer, n As Integer, sSQLAge As String, sSQLArea As String
Dim cLista As String, k As Integer
Dim cSintesis As String

cLista = ""
n = lsvObj.ListItems.Count
For i = 1 To n
    If lsvObj.ListItems(i).Checked Then
       cLista = cLista + "'" + lsvObj.ListItems(i).SubItems(1) + "',"
    End If
Next

If Len(Trim(cLista)) = 0 Then
   MsgBox "Debe seleccionar al menos un Item..." + Space(10), vbInformation
   Exit Sub
End If

cLista = Left(cLista, Len(cLista) - 1)

FormaFlex

sSQLAge = "select r.cRHAgeCod, '' as cRHAreaCod, a.cAgeDescripcion as cDescripcion,cBSGrupoCod=coalesce(g.cBSGrupoCod,''), " & _
"        nObjetoCod = convert(tinyint,substring(p.cBSCod,2,1)), cSintesis=g.cBSGrupoDescripcion, p.nMoneda, sum(p.nPrecioUnitario*(p.nMes01+p.nMes02+p.nMes03+p.nMes04+p.nMes05+p.nMes06+p.nMes07+p.nMes08+p.nMes09+p.nMes10+p.nMes11+p.nMes12)) as nValorEstimado, 1 as nNivel " & _
"   from LogPlanAnualReq r inner join LogPlanAnualReqDetalle p on r.nPlanReqNro = p.nPlanReqNro " & _
"        inner join DBCmactAux..Agencias a on r.cRHAgeCod = a.cAgeCod " & _
"        left join LogProSelBienesServicios b on p.cBSCod = b.cProSelBSCod " & _
"        left join BSGrupos g on b.cBSGrupoCod = g.cBSGrupoCod " & _
"  Where p.nPlanAnualNro = " & nPlanAnualNro & " And p.nEstado = 1 " & _
"       and r.cRHAgeCod in (" & cLista & ") " & _
"  group by r.cRHAgeCod, a.cAgeDescripcion,g.cBSGrupoCod,convert(tinyint,substring(p.cBSCod,2,1)),g.cBSGrupoDescripcion,p.nMoneda "

sSQLArea = "select r.cRHAgeCod,r.cRHAreaCod, a.cAreaDescripcion as cDescripcion,cBSGrupoCod=coalesce(g.cBSGrupoCod,''), " & _
"        nObjetoCod = convert(tinyint,substring(p.cBSCod,2,1)), cSintesis=g.cBSGrupoDescripcion, p.nMoneda, sum(p.nPrecioUnitario*(p.nMes01+p.nMes02+p.nMes03+p.nMes04+p.nMes05+p.nMes06+p.nMes07+p.nMes08+p.nMes09+p.nMes10+p.nMes11+p.nMes12)) as nValorEstimado, 2 as nNivel " & _
"   from LogPlanAnualReq r inner join LogPlanAnualReqDetalle p on r.nPlanReqNro = p.nPlanReqNro " & _
"        inner join DBCmactAux..Areas a on r.cRHAreaCod = a.cAreaCod " & _
"        left join LogProSelBienesServicios b on p.cBSCod = b.cProSelBSCod " & _
"        left join BSGrupos g on b.cBSGrupoCod = g.cBSGrupoCod " & _
"  Where p.nPlanAnualNro = " & nPlanAnualNro & " And p.nEstado = 1 " & _
"       and r.cRHAgeCod in (" & cLista & ") " & _
"  group by r.cRHAgeCod,r.cRHAreaCod, a.cAreaDescripcion,g.cBSGrupoCod,convert(tinyint,substring(p.cBSCod,2,1)),g.cBSGrupoDescripcion,p.nMoneda " & _
"  order by r.cRHAgeCod,g.cBSGrupoCod "
  
If oConn.AbreConexion Then
   Set rs = oConn.CargaRecordSet(sSQLAge + " UNION " + sSQLArea)
   If Not rs.EOF Then
      i = 0
      Do While Not rs.EOF
         i = i + 1
         InsRow MSFlex, i
         
         If Len(rs!cRHAreaCod) = 0 Then
            MSFlex.TextMatrix(i, 1) = rs!cRHAgeCod
            MSFlex.row = i
            For k = 1 To 5
                MSFlex.Col = k
                MSFlex.CellBackColor = "&H00EAFFFF"
                MSFlex.CellFontBold = True
            Next
         End If
         
         If IsNull(rs!cSintesis) Then
            cSintesis = ""
         Else
            cSintesis = rs!cSintesis
         End If
         
         MSFlex.TextMatrix(i, 2) = rs!cDescripcion
         MSFlex.TextMatrix(i, 3) = cSintesis
         MSFlex.TextMatrix(i, 4) = IIf(rs!nMoneda = 1, "S/.", "US$")
         MSFlex.TextMatrix(i, 5) = FNumero(rs!nValorEstimado)
         rs.MoveNext
      Loop
      MSFlex.row = 1
      MSFlex.Col = 1
   End If
End If
End Sub

Private Sub lsvObj_ItemCheck(ByVal Item As MSComctlLib.ListItem)
FormaFlex
lsvObj.ListItems(Item.Index).Selected = True
End Sub
