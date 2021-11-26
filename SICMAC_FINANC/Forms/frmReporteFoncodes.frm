VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmReporteFoncodes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte de Creditos Convenio FONCODES"
   ClientHeight    =   6330
   ClientLeft      =   1110
   ClientTop       =   2430
   ClientWidth     =   11880
   Icon            =   "frmReporteFoncodes.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   11880
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   405
      Left            =   10695
      TabIndex        =   3
      Top             =   5790
      Width           =   1125
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   405
      Left            =   9570
      TabIndex        =   2
      Top             =   5790
      Width           =   1125
   End
   Begin Sicmact.FlexEdit FlexEdit1 
      Height          =   4800
      Left            =   195
      TabIndex        =   1
      Top             =   810
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   8467
      Cols0           =   17
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   $"frmReporteFoncodes.frx":030A
      EncabezadosAnchos=   "450-800-1800-3500-1200-2500-1200-1200-1200-1500-1500-1200-1200-1200-1200-1000-1200"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X"
      ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-L-C-L-L-L-R-R-R-R-R-R-R-R-R-R-R"
      FormatosEdit    =   "0-0-0-0-0-0-2-2-2-2-2-2-2-2-2-2-2"
      TextArray0      =   "Item"
      lbUltimaInstancia=   -1  'True
      ColWidth0       =   450
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin VB.CommandButton cmdProcesar 
      Caption         =   "&Procesar"
      Height          =   390
      Left            =   10695
      TabIndex        =   0
      Top             =   180
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker txtFechaIni 
      Height          =   330
      Left            =   840
      TabIndex        =   5
      Top             =   210
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   582
      _Version        =   393216
      Format          =   59637761
      CurrentDate     =   38478
   End
   Begin MSComCtl2.DTPicker txtFechaFin 
      Height          =   330
      Left            =   2730
      TabIndex        =   6
      Top             =   210
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   582
      _Version        =   393216
      Format          =   59637761
      CurrentDate     =   38478
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Hasta:"
      Height          =   195
      Left            =   2220
      TabIndex        =   7
      Top             =   285
      Width           =   465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Desde:"
      Height          =   195
      Left            =   255
      TabIndex        =   4
      Top             =   285
      Width           =   510
   End
End
Attribute VB_Name = "frmReporteFoncodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdImprimir_Click()
ImprimirExcell
End Sub

Private Sub cmdProcesar_Click()
CargaDatosFoncodes Me.txtFechaini, Me.txtFechafin
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
CentraForm Me
Me.txtFechaini = gdFecSis
Me.txtFechafin = gdFecSis

cmdImprimir.Enabled = False

End Sub
Sub CargaDatosFoncodes(ByVal pdFechaIni As Date, ByVal pdFechaFin As Date, Optional ByVal pbDesembolso As Boolean = False)
Dim sql As String
Dim rs As ADODB.Recordset
Dim oCon As DConecta

Set oCon = New DConecta

oCon.AbreConexion

sql = "SELECT  SUBSTRING(MD.CCTACOD,4,2) AS cCodOfi, MD.cCtaCod, " _
    & "         cPersNombre = ( SELECT P.cPersnombre " _
    & "                         FROM ProductoPersona R JOIN PERSONA P ON P.cPersCod = R.cPersCod " _
    & "                         WHERE R.cCtaCod = MD.cCtaCod AND nPrdPersRelac=20 ), " _
    & "         CASE WHEN SUBSTRING (C.cLineaCred,6,1) ='1' THEN 'CORTO PLAZO' ELSE 'LARGO PLAZO' END AS cCodLin, " _
    & "         L.cDescripcion, " _
    & "         SUM(CASE WHEN MD.COPECOD LIKE '1001%'  THEN MD.NMONTO ELSE 0  END) AS NDESEMBOLSO, " _
    & "         SUM(CASE WHEN MD.nPrdConceptoCod = 1000 AND MD.COPECOD LIKE '100[234567]%' THEN MD.nMonto ELSE 0 END) AS nCapPag, " _
    & "         SUM(CASE WHEN MD.nPrdConceptoCod IN (1100,1105) THEN MD.nMonto*0.69 ELSE 0 END) AS INTPAGCMAC, " _
    & "         SUM(CASE WHEN MD.nPrdConceptoCod IN (1100,1105) THEN MD.nMonto*0.20 ELSE 0 END) AS INTPAGFONCCAPIT, " _
    & "         SUM(CASE WHEN MD.nPrdConceptoCod IN (1100,1105) THEN MD.nMonto*0.11 ELSE 0 END) AS INTPAGFONCCAPAC, " _
    & "         SUM(CASE WHEN MD.nPrdConceptoCod IN (1100,1105) THEN MD.nMonto ELSE 0 END) AS TOTINTPAG, " _
    & "         SUM(CASE WHEN MD.nPrdConceptoCod = 1106 THEN MD.nMonto ELSE 0 END) AS INTDESAG, " _
    & "         SUM(CASE WHEN MD.nPrdConceptoCod = 1101 THEN MD.nMonto ELSE 0 END) AS MORAPAG, " _
    & "         SUM(CASE WHEN MD.nPrdConceptoCod NOT IN (1000, 1100, 1105,1101,1106) AND NOT MD.COPECOD LIKE '99%' THEN MD.nMonto ELSE 0 END) AS GASTOS, " _
    & "         SUM(CASE WHEN MD.COPECOD LIKE '99%' THEN MD.nMonto ELSE 0 END) AS NITF, " _
    & "         SUM(MD.nMonto) As NTOTALCAJA " _
    & "   FROM MOV M    " _
    & "         JOIN MOVCOL MC ON MC.NMOVNRO = M.NMOVNRO " _
    & "         JOIN MOVCOLDET MD ON MD.NMOVNRO = MC.NMOVNRO AND MC.COPECOD = MD.COPECOD AND MC.CCTACOD = MD.CCTACOD " _
    & "         JOIN COLOCACIONES C ON C.CCTACOD = MD.CCTACOD " _
    & "         JOIN COLOCLINEACREDITO L ON L.cLineacred = C.cLineaCred " _
    & "   WHERE C.CLINEACRED IN ('04991120103','04991220101','04991120102') " _
    & "         AND LEFT(M.CMOVNRO ,8) BETWEEN '" & Format(pdFechaIni, "yyyymmdd") & "' AND '" & Format(pdFechaFin, "yyyymmdd") & "' AND M.NMOVFLAG = 0 " _
    & "         AND NOT MD.COPECOD LIKE '107%' and Not MD.cOpeCod in ('107002','107003') "
sql = sql & " GROUP BY SUBSTRING(MD.CCTACOD,4,2), MD.CCTACOD, C.CLINEACRED, L.cDescripcion " _
    & "         ORDER BY SUBSTRING(MD.CCTACOD,4,2)"

Set rs = oCon.CargaRecordSet(sql)

If rs.RecordCount > 0 Then
   cmdImprimir.Enabled = True
End If

Set Me.FlexEdit1.Recordset = rs

oCon.CierraConexion
Set oCon = Nothing

End Sub
Sub ImprimirExcell()
Dim vExcelObj As Excel.Application
Dim vNHC As String
Dim lsCodCta As String


If Me.FlexEdit1.TextMatrix(0, 1) = "" Then
    Exit Sub
End If

vNHC = App.path & "\spooler\Rep_Foncodes" & Format(txtFechaini, "yyyymm") & ".XLS"

Set vExcelObj = New Excel.Application  '   = CreateObject("Excel.Application")
vExcelObj.DisplayAlerts = False

vExcelObj.Workbooks.Add
vExcelObj.Sheets("Hoja1").Select
vExcelObj.Sheets("Hoja1").Name = "CONTROL"

vExcelObj.Range("A1:IV65536").Font.Name = "Arial Narrow"
vExcelObj.Range("A1:IV65536").Font.Size = 8


vExcelObj.Range("A1").Select
vExcelObj.Range("A1").Font.Bold = True
vExcelObj.Range("A1").HorizontalAlignment = 1
vExcelObj.ActiveCell.value = UCase(Trim(gsNomCmac))

vExcelObj.Range("H1").Select
vExcelObj.Range("H1").Font.Bold = True
vExcelObj.Range("H1").HorizontalAlignment = 1
vExcelObj.ActiveCell.value = "Informacion del mes de " & Format(txtFechafin, "mm/yyyy")

vExcelObj.Range("A2").Select
vExcelObj.Range("A2").Font.Bold = True
vExcelObj.Range("A2").HorizontalAlignment = 1
vExcelObj.ActiveCell.value = gsNomAge

vExcelObj.Range("A4").Select
vExcelObj.Range("A4").Font.Bold = True
vExcelObj.Range("A4").HorizontalAlignment = 1
vExcelObj.ActiveCell.value = "LISTADO DE DESEMBOLSOS Y PAGOS CONVENIO FONCODES"

vExcelObj.Range("A6").Select
vExcelObj.Range("A6").Font.Bold = True
vExcelObj.Range("A6").ColumnWidth = 8
vExcelObj.ActiveCell.value = "Agencia"

vExcelObj.Range("B6").Select
vExcelObj.Range("B6").Font.Bold = True
vExcelObj.Range("B6").ColumnWidth = 15
vExcelObj.ActiveCell.value = "N° Credito"

vExcelObj.Range("C6").Select
vExcelObj.Range("C6").Font.Bold = True
vExcelObj.Range("C6").ColumnWidth = 30
vExcelObj.ActiveCell.value = "Cliente"

vExcelObj.Range("D6").Select
vExcelObj.Range("D6").Font.Bold = True
vExcelObj.Range("D6").ColumnWidth = 15
vExcelObj.ActiveCell.value = "Plazo"

vExcelObj.Range("E6").Select
vExcelObj.Range("E6").Font.Bold = True
vExcelObj.Range("E6").ColumnWidth = 30
vExcelObj.ActiveCell.value = "Linea Credito"

vExcelObj.Range("F6").Select
vExcelObj.Range("F6").Font.Bold = True
vExcelObj.Range("F6").ColumnWidth = 10
vExcelObj.ActiveCell.value = "Desembolso"

vExcelObj.Range("G6").Select
vExcelObj.Range("G6").Font.Bold = True
vExcelObj.Range("G6").ColumnWidth = 10
vExcelObj.ActiveCell.value = "Cap.Pagado"

vExcelObj.Range("H6").Select
vExcelObj.Range("H6").Font.Bold = True
vExcelObj.Range("H6").ColumnWidth = 10
vExcelObj.ActiveCell.value = "Int.Pag.CMAC"

vExcelObj.Range("I6").Select
vExcelObj.Range("I6").Font.Bold = True
vExcelObj.Range("I6").ColumnWidth = 10
vExcelObj.ActiveCell.value = "Int.Pag.FONC.CAPITAL"

vExcelObj.Range("J6").Select
vExcelObj.Range("J6").Font.Bold = True
vExcelObj.Range("J6").ColumnWidth = 10
vExcelObj.ActiveCell.value = "Int.Pag.FONC.CAPAC"

vExcelObj.Range("K6").Select
vExcelObj.Range("K6").Font.Bold = True
vExcelObj.Range("K6").ColumnWidth = 10
vExcelObj.ActiveCell.value = "TOTAL.INT.PAG"

vExcelObj.Range("L6").Select
vExcelObj.Range("L6").Font.Bold = True
vExcelObj.Range("L6").ColumnWidth = 10
vExcelObj.ActiveCell.value = "INT.DESAG"

vExcelObj.Range("M6").Select
vExcelObj.Range("M6").Font.Bold = True
vExcelObj.Range("M6").ColumnWidth = 10
vExcelObj.ActiveCell.value = "MORA.PAG."

vExcelObj.Range("N6").Select
vExcelObj.Range("N6").Font.Bold = True
vExcelObj.Range("N6").ColumnWidth = 10
vExcelObj.ActiveCell.value = "GASTOS"

vExcelObj.Range("O6").Select
vExcelObj.Range("O6").Font.Bold = True
vExcelObj.Range("O6").ColumnWidth = 10
vExcelObj.ActiveCell.value = "ITF"

vExcelObj.Range("P6").Select
vExcelObj.Range("P6").Font.Bold = True
vExcelObj.Range("P6").ColumnWidth = 10
vExcelObj.ActiveCell.value = "TOTAL.CAJA"

vIni = 6
vItem = vIni
lnTotalCtasCMAC = 0
For I = 1 To FlexEdit1.Rows - 1
    vItem = vItem + 1
    
    vCel = "A" + Trim(Str(vItem))
    vExcelObj.Range(vCel).Select
    vExcelObj.ActiveCell.value = "'" + FlexEdit1.TextMatrix(I, 1)
    
    vCel = "B" + Trim(Str(vItem))
    vExcelObj.Range(vCel).Select
    vExcelObj.ActiveCell.value = "'" + FlexEdit1.TextMatrix(I, 2)
    
    vCel = "C" + Trim(Str(vItem))
    vExcelObj.Range(vCel).Select
    vExcelObj.ActiveCell.value = "'" + Trim(FlexEdit1.TextMatrix(I, 3))
    
    vCel = "D" + Trim(Str(vItem))
    vExcelObj.Range(vCel).Select
    vExcelObj.ActiveCell.value = "'" + FlexEdit1.TextMatrix(I, 4)
    
    vCel = "E" + Trim(Str(vItem))
    vExcelObj.Range(vCel).Select
    vExcelObj.ActiveCell.value = "'" + FlexEdit1.TextMatrix(I, 5)
    
    vCel = "F" + Trim(Str(vItem))
    vExcelObj.Range(vCel).Select
    vExcelObj.ActiveCell.value = FlexEdit1.TextMatrix(I, 6)
    
    vCel = "G" + Trim(Str(vItem))
    vExcelObj.Range(vCel).Select
    vExcelObj.ActiveCell.value = FlexEdit1.TextMatrix(I, 7)
    
    vCel = "H" + Trim(Str(vItem))
    vExcelObj.Range(vCel).Select
    vExcelObj.ActiveCell.value = FlexEdit1.TextMatrix(I, 8)
    
    vCel = "I" + Trim(Str(vItem))
    vExcelObj.Range(vCel).Select
    vExcelObj.ActiveCell.value = FlexEdit1.TextMatrix(I, 9)
    
    vCel = "J" + Trim(Str(vItem))
    vExcelObj.Range(vCel).Select
    vExcelObj.ActiveCell.value = FlexEdit1.TextMatrix(I, 10)
    
    vCel = "K" + Trim(Str(vItem))
    vExcelObj.Range(vCel).Select
    vExcelObj.ActiveCell.value = FlexEdit1.TextMatrix(I, 11)
    
    vCel = "L" + Trim(Str(vItem))
    vExcelObj.Range(vCel).Select
    vExcelObj.ActiveCell.value = FlexEdit1.TextMatrix(I, 12)
    
    vCel = "M" + Trim(Str(vItem))
    vExcelObj.Range(vCel).Select
    vExcelObj.ActiveCell.value = FlexEdit1.TextMatrix(I, 13)
    
    vCel = "N" + Trim(Str(vItem))
    vExcelObj.Range(vCel).Select
    vExcelObj.ActiveCell.value = FlexEdit1.TextMatrix(I, 14)
    
    vCel = "O" + Trim(Str(vItem))
    vExcelObj.Range(vCel).Select
    vExcelObj.ActiveCell.value = FlexEdit1.TextMatrix(I, 15)
    
    vCel = "P" + Trim(Str(vItem))
    vExcelObj.Range(vCel).Select
    vExcelObj.ActiveCell.value = FlexEdit1.TextMatrix(I, 16)
Next

vExcelObj.Range("A6").Select
vExcelObj.Range("A6").Subtotal 1, xlSum, Array(6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16), True


'If Dir(vNHC) <> "" Then
'   If MsgBox("Archivo Ya Existe ...  Desea Reemplazarlo ??", vbQuestion + vbYesNo + vbDefaultButton1, " Mensaje del Sistema ...") = vbNo Then
'      Exit Sub
'   End If
'End If
vExcelObj.Range("A1").Select
vExcelObj.ActiveWorkbook.SaveAs (vNHC)
vExcelObj.ActiveWorkbook.Close
MsgBox "SE HA GENERADO CON ÉXITO EL ARCHIVO !!  ", vbInformation, " Mensaje del Sistema ..."
vExcelObj.Workbooks.Open (vNHC)
vExcelObj.Visible = True

Set vExcelObj = Nothing


End Sub
