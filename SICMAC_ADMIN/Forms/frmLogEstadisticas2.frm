VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{DFDE2506-090D-11D5-BEF8-C11EAA34970C}#2.0#0"; "Spinner.ocx"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmLogEstadisticas 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6720
   ClientLeft      =   1155
   ClientTop       =   1560
   ClientWidth     =   8385
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   8385
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   675
      Left            =   60
      TabIndex        =   15
      Top             =   -60
      Width           =   8270
      Begin VB.ComboBox cboRep 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   240
         Width           =   6615
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Estadísticas de"
         Height          =   195
         Left            =   180
         TabIndex        =   17
         Top             =   300
         Width           =   1095
      End
   End
   Begin TabDlg.SSTab sstReg 
      Height          =   6015
      Left            =   60
      TabIndex        =   0
      Top             =   660
      Width           =   8265
      _ExtentX        =   14579
      _ExtentY        =   10610
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabHeight       =   564
      TabCaption(0)   =   "Procesamiento "
      TabPicture(0)   =   "frmLogEstadisticas.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraRep02"
      Tab(0).Control(1)=   "fraRep01"
      Tab(0).Control(2)=   "flxVis"
      Tab(0).Control(3)=   "cmdSalir"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Gráfico            "
      TabPicture(1)   =   "frmLogEstadisticas.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label7"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Graf"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "chkGrafico"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "spEscala"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      Begin Spinner.uSpinner spEscala 
         Height          =   255
         Left            =   7500
         TabIndex        =   14
         Top             =   480
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   450
         Increment       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
      End
      Begin VB.CheckBox chkGrafico 
         Caption         =   "Grafico de sectores"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   1755
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "&Salir"
         Height          =   360
         Left            =   -68160
         TabIndex        =   10
         Top             =   5580
         Width           =   1275
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxVis 
         Height          =   3735
         Left            =   -74880
         TabIndex        =   5
         Top             =   1800
         Width           =   7995
         _ExtentX        =   14102
         _ExtentY        =   6588
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         BackColorBkg    =   -2147483643
         GridColor       =   -2147483633
         ScrollBars      =   2
         _NumberOfBands  =   1
         _Band(0).Cols   =   4
      End
      Begin MSChart20Lib.MSChart Graf 
         Height          =   5115
         Left            =   60
         OleObjectBlob   =   "frmLogEstadisticas.frx":0038
         TabIndex        =   1
         Top             =   780
         Width           =   8115
      End
      Begin VB.Frame fraRep01 
         Height          =   1395
         Left            =   -74880
         TabIndex        =   2
         Top             =   360
         Width           =   7995
         Begin VB.TextBox txtFecha 
            Height          =   315
            Left            =   6540
            TabIndex        =   11
            Top             =   240
            Width           =   1215
         End
         Begin VB.ComboBox cboBase 
            Height          =   315
            ItemData        =   "frmLogEstadisticas.frx":255B
            Left            =   1800
            List            =   "frmLogEstadisticas.frx":2565
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   960
            Width           =   5955
         End
         Begin VB.ComboBox cboMoneda 
            Height          =   315
            ItemData        =   "frmLogEstadisticas.frx":25A8
            Left            =   1800
            List            =   "frmLogEstadisticas.frx":25B2
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   240
            Width           =   1335
         End
         Begin VB.ComboBox cboDoc 
            Height          =   315
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   600
            Width           =   5955
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Fecha"
            Height          =   195
            Left            =   5940
            TabIndex        =   12
            Top             =   300
            Width           =   450
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Porcentaje en base a"
            Height          =   195
            Left            =   180
            TabIndex        =   9
            Top             =   1020
            Width           =   1515
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Moneda"
            Height          =   195
            Left            =   180
            TabIndex        =   7
            Top             =   300
            Width           =   585
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Documento"
            Height          =   195
            Left            =   180
            TabIndex        =   4
            Top             =   660
            Width           =   825
         End
      End
      Begin VB.Frame fraRep02 
         Height          =   1395
         Left            =   -74880
         TabIndex        =   18
         Top             =   360
         Visible         =   0   'False
         Width           =   7995
         Begin VB.ComboBox cboMes 
            Height          =   315
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   240
            Width           =   1635
         End
         Begin VB.ComboBox Combo3 
            Height          =   315
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   600
            Width           =   6615
         End
         Begin VB.TextBox txtAnio 
            Height          =   315
            Left            =   2880
            MaxLength       =   4
            TabIndex        =   19
            Top             =   240
            Width           =   675
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Documento"
            Height          =   195
            Left            =   180
            TabIndex        =   22
            Top             =   660
            Width           =   825
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Mes y Año"
            Height          =   195
            Left            =   180
            TabIndex        =   21
            Top             =   300
            Width           =   750
         End
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Escala"
         Height          =   195
         Left            =   6900
         TabIndex        =   24
         Top             =   495
         Width           =   480
      End
   End
End
Attribute VB_Name = "frmLogEstadisticas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sSQL As String, EsPrimera As Boolean

Private Sub cboBase_Click()
If Not EsPrimera Then
   GenerarPorcentajes
End If
End Sub

Private Sub cboDoc_Click()
If Not EsPrimera Then
   GenerarPorcentajes
End If
End Sub


Private Sub cboMes_Click()
If Not EsPrimera Then
   GenerarReporte02
End If
End Sub



Private Sub txtAnio_Change()
If Len(Trim(txtAnio)) = 4 And Not EsPrimera Then
   GenerarReporte02
End If
End Sub

Sub GenerarReporte02()
Dim rs As New ADODB.Recordset, nSuma  As Currency
Dim oCon As New DConecta, i As Integer, n As Integer
Dim nMes As Integer, cAnioMes As String

   'n = flxVis.Rows - 1
   'For i = 1 To n
   '    flxVis.TextMatrix(i, 2) = 100 * flxVis.TextMatrix(i, 2)
   'Next
   
nMes = cboMes.ListIndex + 1
cAnioMes = txtAnio.Text + Format(nMes, "00")
If Len(Trim(cAnioMes)) < 6 Or nMes = 0 Then
   'MsgBox "Mes/año no válido..." + Space(10), vbInformation
   Exit Sub
End If
FormaFlex 0
sSQL = "select x.cAreaCod, a.cAreaDescripcion, count(x.cAreaCod) as Nro " & _
" from (Select MOAA.cAreaCod,nNro=coalesce(MREF.Atendidos,0) " & _
" From Mov M Inner Join MovBS MBS On M.nMovNro = MBS.nMovNro " & _
"           Inner Join MovCant MCT On MBS.nMovNro = MCT.nMovNro And MBS.nMovItem = MCT.nMovItem " & _
"           Inner Join MovObjAreaAgencia MOAA On MBS.nMovNro = MOAA.nMovNro And MBS.nMovItem = MOAA.nMovItem " & _
" Left Join (Select Sum(nMovCant) Atendidos, MRMBS.cBSCod, MR.nMovNroRef nMovNro " & _
"            From MovRef MR Inner Join MovCant MRMCT On MR.nMovNro = MRMCT.nMovNro " & _
"                 Inner Join Mov MRM On MRM.nMovNro = MR.nMovNro " & _
"                 Inner Join MovBS MRMBS On MRMCT.nMovNro = MRMBS.nMovNro And MRMCT.nMovItem = MRMBS.nMovItem " & _
"           Where MRM.nMovFlag Not IN ('2','1','3','5') And MRM.cOpeCod like '591201%' " & _
"           Group By MRMBS.cBSCod,  MR.nMovNroRef) MREF On MREF.nMovNro = M.nMovNro And MREF.cBSCod = MBS.cBSCod " & _
" where left(M.cMovNro,6)='" & cAnioMes & "') x inner join areas a on x.cAreaCod = a.cAreaCod " & _
" where x.nNro = 0 group by x.cAreaCod, a.cAreaDescripcion order by x.cAreaCod "

oCon.AbreConexion
Set rs = oCon.CargaRecordSet(sSQL)
If Not rs.EOF Then
   i = 0: nSuma = 0
   rs.MoveFirst
   Do While Not rs.EOF
      i = i + 1
      If i > 1 Then
         flxVis.Rows = i + 1
      Else
         flxVis.RowHeight(1) = 240
      End If
      flxVis.TextMatrix(i, 0) = rs!cAreaCod
      flxVis.TextMatrix(i, 1) = rs!cAreaDescripcion
      flxVis.TextMatrix(i, 2) = rs!nro
      nSuma = nSuma + rs!nro
      rs.MoveNext
   Loop
   n = flxVis.Rows - 1
   For i = 1 To n
       flxVis.TextMatrix(i, 2) = Format(100 * Val(flxVis.TextMatrix(i, 2)) / nSuma, "##0.00")
   Next
   sstReg.TabEnabled(1) = True
Else
   sstReg.TabEnabled(1) = False
End If
sstReg.Tab = 0
oCon.CierraConexion
End Sub

Private Sub cboMoneda_Click()
If Not EsPrimera Then
   GenerarPorcentajes
End If
End Sub

Private Sub cboRep_Click()
Select Case cboRep.ListIndex
    Case 0
         fraRep01.Visible = True
         fraRep02.Visible = False
    Case 1
         cboMes.Clear
         cboMes.AddItem "ENERO"
         cboMes.AddItem "FEBRERO"
         cboMes.AddItem "MARZO"
         cboMes.AddItem "ABRIL"
         cboMes.AddItem "MAYO"
         cboMes.AddItem "JUNIO"
         cboMes.AddItem "JULIO"
         cboMes.AddItem "AGOSTO"
         cboMes.AddItem "SEPTIEMBRE"
         cboMes.AddItem "OCTUBRE"
         cboMes.AddItem "NOVIEMBRE"
         cboMes.AddItem "DICIEMBRE"
         txtAnio = Year(Date)
         fraRep01.Visible = False
         fraRep02.Visible = True
End Select
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
EsPrimera = True
txtFecha = Date
spEscala.Valor = 50
cboRep.Clear
cboRep.AddItem "COMPRAS NO PRESUPUESTADAS"
cboRep.AddItem "REQUERIMIENTOS NO ATENDIDOS"
cboRep.ListIndex = 0
CargaDocumentos
cboMoneda.ListIndex = 0
cboBase.ListIndex = 0
sstReg.TabEnabled(1) = False
sstReg.Tab = 0
FormaFlex 0

If EsPrimera Then
   GenerarPorcentajes
   EsPrimera = False
End If

End Sub

Sub FormaFlex(nFilas As Integer)
flxVis.Clear
flxVis.RowHeight(0) = 300
Select Case nFilas
    Case 0
         flxVis.Rows = 2
         flxVis.RowHeight(1) = 0
    Case 1
         flxVis.Rows = 2
         flxVis.RowHeight(1) = 240
    Case Is > 1
         flxVis.Rows = nFilas - 1
         flxVis.RowHeight(1) = 240
End Select
flxVis.ColWidth(0) = 400
flxVis.ColWidth(1) = 5550
flxVis.ColWidth(2) = 1200
flxVis.ColWidth(3) = 0
End Sub

Sub CargaDocumentos()
Dim rs As New ADODB.Recordset
Dim oCon As New DConecta
    
cboDoc.Clear
sSQL = "select nDocTpo,cDocDesc from Documento where nDocTpo in (33,70) "
oCon.AbreConexion
Set rs = oCon.CargaRecordSet(sSQL)
If Not rs.EOF Then
   'CargaCombo rs, cboDoc
   Do While Not rs.EOF
      cboDoc.AddItem rs!cDocDesc
      cboDoc.ItemData(cboDoc.ListCount - 1) = rs!nDocTpo
      rs.MoveNext
   Loop
   cboDoc.ListIndex = 0
End If
oCon.CierraConexion
End Sub

Private Sub GenerarPorcentajes()
Dim rs As New ADODB.Recordset
Dim oCon As New DConecta, i As Integer
Dim nDocTpo As Integer, cMoneda As Integer, nBase As Integer

FormaFlex 0
sSQL = ""
nDocTpo = cboDoc.ItemData(cboDoc.ListIndex)
cMoneda = CStr(cboMoneda.ListIndex + 1)
nBase = cboBase.ListIndex

Select Case nBase
    Case 0
         sSQL = "select left(m.cAreaCod,3) as cAreaCod, a.cAreaDescripcion, " & _
         "       nPorc= 100*convert(money,count(d.nDocTpo))/(select count(d.nDocTpo) as Nro from MOVCOTIZAC m inner join MovDoc d on m.nMovNro = d.nMovNro " & _
         "       left outer join Areas a on left(m.cAreaCod,3) = a.cAreaCod " & _
         " where m.nPresupuestado>0 and d.nDocTpo = " & nDocTpo & " AND a.cAreaDescripcion is not NULL and m.cMovMoneda='" & cMoneda & "') " & _
         "  from MOVCOTIZAC m inner join MovDoc d on m.nMovNro = d.nMovNro " & _
         "       left outer join Areas a on left(m.cAreaCod,3) = a.cAreaCod " & _
         " where m.nPresupuestado>0 and d.nDocTpo = " & nDocTpo & " AND " & _
         "       a.cAreaDescripcion is not NULL and m.cMovMoneda='" & cMoneda & "' " & _
         " group by left(m.cAreaCod,3), a.cAreaDescripcion "


         'sSQL = "select cAreaCod,cAreaDescripcion, nPorc=100*Nro/(select count(d.nDocTpo) " & _
         '"  from MOVCOTIZAC m inner join MovDoc d on m.nMovNro = d.nMovNro " & _
         '"       left outer join Areas a on m.cAreaCod = a.cAreaCod " & _
         '" where m.nPresupuestado>0 AND len(m.cAreaCod)=3 AND d.nDocTpo = " & nDocTpo & " and m.cMovMoneda = '" & cMoneda & "' ) " & _
         '" from (select m.cAreaCod, a.cAreaDescripcion, Nro = convert(money,count(d.nDocTpo)) " & _
         '"  from MOVCOTIZAC m inner join MovDoc d on m.nMovNro = d.nMovNro " & _
         '"       left outer join Areas a on m.cAreaCod = a.cAreaCod " & _
         '" Where m.nPresupuestado > 0 And Len(m.cAreaCod) = 3 And d.nDocTpo = " & nDocTpo & " and m.cMovMoneda = '" & cMoneda & "' " & _
         '" group by m.cAreaCod, a.cAreaDescripcion) x "
    Case 1
         sSQL = "Select m.cAreaCod, a.cAreaDescripcion, nPorc=100*sum(c.nMovImporte)/(select sum(c.nMovImporte) " & _
         "  from MOVCOTIZAC m inner join MovDoc d on m.nMovNro = d.nMovNro " & _
         "       INNER JOIN MovCta c on m.nMovNro = c.nMovNro " & _
         "       left outer join Areas a on m.cAreaCod = a.cAreaCod " & _
         " where m.nPresupuestado>0 AND len(m.cAreaCod)=3 AND d.nDocTpo = " & nDocTpo & " and nMovImporte>0 and m.cMovMoneda = '" & cMoneda & "')  " & _
         "  from MOVCOTIZAC m inner join MovDoc d on m.nMovNro = d.nMovNro " & _
         "       INNER JOIN MovCta c on m.nMovNro = c.nMovNro " & _
         "       left outer join Areas a on m.cAreaCod = a.cAreaCod " & _
         " Where m.nPresupuestado > 0 And Len(m.cAreaCod) = 3 And d.nDocTpo = " & nDocTpo & " And nMovImporte > 0 and m.cMovMoneda = '" & cMoneda & "'" & _
         " Group by m.cAreaCod, a.cAreaDescripcion"
End Select

If Len(sSQL) = 0 Then Exit Sub

oCon.AbreConexion
Set rs = oCon.CargaRecordSet(sSQL)
If Not rs.EOF Then
   i = 0
   Do While Not rs.EOF
      i = i + 1
      If i > 1 Then
         flxVis.Rows = i + 1
      Else
         flxVis.RowHeight(1) = 240
      End If
      flxVis.TextMatrix(i, 0) = rs!cAreaCod
      flxVis.TextMatrix(i, 1) = rs!cAreaDescripcion
      flxVis.TextMatrix(i, 2) = Format(rs!nPorc, "##0.00")
      rs.MoveNext
   Loop
   sstReg.TabEnabled(1) = True
Else
   sstReg.TabEnabled(1) = False
End If
sstReg.Tab = 0
oCon.CierraConexion
End Sub

Private Sub spEscala_Change()
Graf.Plot.Axis(VtChAxisIdY).ValueScale.Maximum = spEscala.Valor
End Sub

Private Sub sstReg_Click(PreviousTab As Integer)

If sstReg.Tab = 1 Then
   chkGrafico.value = 0
   GraficoBarra
End If

End Sub

Private Sub chkGrafico_Click()
If chkGrafico.value = 1 Then
   GraficoPie
Else
   GraficoBarra
End If
End Sub

Sub GraficoBarra()
Dim i As Integer, n As Integer

Graf.RowCount = 0
Graf.ColumnCount = 0
If sstReg.Tab = 1 Then
   Graf.ChartType = VtChChartType2dBar
   'Graf.ChartType = VtChChartType3dBar
   Graf.Plot.Axis(VtChAxisIdY).ValueScale.Minimum = 0
   Graf.Plot.Axis(VtChAxisIdY).ValueScale.Maximum = 50
   'Graf.ShowLegend = False
   Graf.ShowLegend = True
   n = flxVis.Rows - 1
   Graf.ColumnCount = n
   Graf.RowCount = 1
   For i = 1 To n
       Graf.Column = i
       Graf.row = 1
       Graf.data = flxVis.TextMatrix(i, 2)
       'Graf.RowLabel = Mid(flxVis.TextMatrix(i, 1), 1, 10)
       Graf.ColumnLabel = Mid(flxVis.TextMatrix(i, 1), 1, 20)
   Next
End If
End Sub

Sub GraficoPie()
Graf.ChartType = VtChChartType2dPie
End Sub

Sub GraficoPie2()
Dim i As Integer, n As Integer
'Necesita la opcion series apiladas
Graf.RowCount = 0
Graf.ColumnCount = 0
'Graf.Legend.Font.Name = "Tahoma"
'Graf.Legend.Font.Size = 7
Graf.ShowLegend = True

If sstReg.Tab = 1 Then
   Graf.ChartType = VtChChartType2dPie
   Graf.Plot.Axis(VtChAxisIdY).ValueScale.Minimum = 0
   Graf.Plot.Axis(VtChAxisIdY).ValueScale.Maximum = 10
   
   n = flxVis.Rows - 1
   
   Graf.ColumnCount = 1
   Graf.RowCount = n
   
   For i = 1 To n
       Graf.row = i
       Graf.Column = 1
       Graf.data = flxVis.TextMatrix(i, 2)
       Graf.RowLabel = flxVis.TextMatrix(i, 1)
   Next
   Graf.ShowLegend = False
End If
End Sub

