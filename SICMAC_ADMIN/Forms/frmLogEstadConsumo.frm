VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLogEstadConsumo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Estadísticas de Consumo"
   ClientHeight    =   5670
   ClientLeft      =   2205
   ClientTop       =   2025
   ClientWidth     =   8250
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   8250
   Begin MSComctlLib.ProgressBar Barra 
      Height          =   315
      Left            =   1500
      TabIndex        =   20
      Top             =   3300
      Visible         =   0   'False
      Width           =   5475
      _ExtentX        =   9657
      _ExtentY        =   556
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.CommandButton cmdExpExcel 
      Caption         =   "Exportar a Excel"
      Height          =   375
      Left            =   1320
      TabIndex        =   19
      Top             =   5280
      Width           =   1575
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6960
      TabIndex        =   18
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "imprimir"
      Height          =   375
      Left            =   60
      TabIndex        =   7
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   1995
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8175
      Begin VB.ComboBox cboMesFin 
         Height          =   315
         Left            =   3900
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1560
         Width           =   1755
      End
      Begin VB.TextBox txtAnioFin 
         Height          =   315
         Left            =   5640
         MaxLength       =   4
         TabIndex        =   11
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox txtAnioIni 
         Height          =   315
         Left            =   2580
         MaxLength       =   4
         TabIndex        =   10
         Top             =   1560
         Width           =   615
      End
      Begin VB.ComboBox cboMesIni 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1560
         Width           =   1755
      End
      Begin VB.CommandButton cmdGenerar 
         Caption         =   "Generar Lista"
         Height          =   360
         Left            =   6360
         TabIndex        =   6
         Top             =   1540
         Width           =   1635
      End
      Begin VB.ComboBox cboBSF 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1200
         Width           =   7155
      End
      Begin VB.ComboBox cboBS8 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   900
         Width           =   7155
      End
      Begin VB.ComboBox cboBS5 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   600
         Width           =   7155
      End
      Begin VB.ComboBox cboBS3 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   300
         Width           =   7155
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Nivel 4"
         Height          =   195
         Left            =   180
         TabIndex        =   17
         Top             =   1260
         Width           =   495
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Nivel 3"
         Height          =   195
         Left            =   180
         TabIndex        =   16
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nivel 2"
         Height          =   195
         Left            =   180
         TabIndex        =   15
         Top             =   660
         Width           =   495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nivel 1"
         Height          =   195
         Left            =   180
         TabIndex        =   14
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         Height          =   195
         Left            =   3360
         TabIndex        =   13
         Top             =   1635
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         Height          =   195
         Left            =   180
         TabIndex        =   9
         Top             =   1635
         Width           =   465
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxObj 
      Height          =   3255
      Left            =   60
      TabIndex        =   0
      Top             =   1980
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   5741
      _Version        =   393216
      Cols            =   20
      FixedCols       =   0
      BackColorBkg    =   -2147483643
      GridColor       =   -2147483633
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
      _Band(0).Cols   =   20
   End
End
Attribute VB_Name = "frmLogEstadConsumo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sSQL As String, cCod As String
Dim mMes(1 To 12) As String
'ARLO 20170126******************
Dim objPista As COMManejador.Pista
'*******************************

Private Sub cmdExpExcel_Click()
Dim appExcel As New Excel.Application
Dim wbExcel As Excel.Workbook
Dim i As Integer, j As Integer, NC As Integer, NF As Integer, f As Integer
Dim cArchivo As String, c() As String, Celda As String, cLinea As String
Dim nFilaIni As Integer, nFilaFin As Integer, v As Double
Dim cLetra As String, nro As Integer

ReDim c(1 To 50)
cLetra = ""
NC = flxObj.Cols - 1
NF = flxObj.Rows - 1
f = 0
nro = 0
For i = 1 To NC
    nro = nro + 1
    c(i) = cLetra + Chr(nro + 64)
    If nro = 26 Then
       f = f + 1
       cLetra = Chr(f + 64)
       nro = 0
    End If
Next

Set wbExcel = appExcel.Workbooks.Add
wbExcel.Worksheets(1).Range("A1:AZ80").Font.Size = 8
NC = flxObj.Cols - 1
NF = flxObj.Rows - 1
wbExcel.Worksheets(1).Range("B1").ColumnWidth = 30

For j = 2 To NC
    If flxObj.ColWidth(j) > 9 Then
       wbExcel.Worksheets(1).Range(c(j) + "1").value = flxObj.TextMatrix(0, j)
    End If
Next

Barra.value = 0
Barra.Max = NF
Barra.Visible = True
DoEvents
  
For i = 1 To NF
    wbExcel.Worksheets(1).Range("A" + CStr(i + 2)).value = flxObj.TextMatrix(i, 1)
    For j = 2 To NC
        wbExcel.Worksheets(1).Range(c(j) + CStr(i + 2)).value = flxObj.TextMatrix(i, j)
    Next
    Barra.value = i
Next
Barra.Visible = False
appExcel.Application.Visible = True
appExcel.Windows(1).Visible = True

        'ARLO 20160126 ***
        gsOpeCod = LogPistaReporteEstadistico
        Set objPista = New COMManejador.Pista
        objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "", "Genero el Reporte Estadistico de Comsumo en Excel "
        Set objPista = Nothing
        '**************
End Sub

Private Sub cmdImprimir_Click()
Dim rs As New ADODB.Recordset
Dim i As Integer, n As Integer
Dim v As Double, f As Integer
Dim cArchivo As String
Dim k As Integer, sCad As String

f = FreeFile
cArchivo = "\RepStat.txt"
Open App.path + cArchivo For Output As #f

Print #f, Centra("ESTADISTICAS DE CONSUMO MENSUAL", 70)
Print #f, ""
Print #f, String(80, "=")
Print #f, " "
Print #f, String(80, "-")

n = flxObj.Rows - 1
For i = 1 To n
    sCad = ""
    For k = 3 To 14
        If flxObj.ColWidth(k) > 10 Then
           sCad = sCad + " " + JDER(flxObj.TextMatrix(i, k), 10)
        End If
    Next
    Print #f, flxObj.TextMatrix(i, 0) + " " + flxObj.TextMatrix(i, 1) + " " + JIZQ(flxObj.TextMatrix(i, 2), 30) + sCad
Next i
Print #f, String(80, "=")
Close #f
v = Shell("notepad " & App.path + cArchivo & "", vbNormalFocus)
        'ARLO 20160126 ***
        gsOpeCod = LogPistaReporteEstadistico
        Set objPista = New COMManejador.Pista
        objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "", "Imprimio el Reporte Estadistico de Comsumo "
        Set objPista = Nothing
        '**************
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
CentraForm Me
txtAnioIni = Year(Date)
txtAnioFin = Year(Date)
CargaCombos
End Sub

Sub CargaCombos()
Dim rs As ADODB.Recordset
Dim oCon As DConecta
Set oCon = New DConecta

Set rs = New ADODB.Recordset

cboMesIni.Clear:    cboMesFin.Clear
cboMesIni.AddItem "--------------":   cboMesFin.AddItem "--------------"

sSQL = "select nConsValor,cConsDescripcion from Constante where nConsCod = 1010"
oCon.AbreConexion
Set rs = oCon.CargaRecordSet(sSQL)
If Not rs.EOF Then
   Do While Not rs.EOF
      mMes(rs!nConsValor) = Mid(UCase(rs!cConsDescripcion), 1, 3)
      cboMesIni.AddItem UCase(rs!cConsDescripcion)
      cboMesFin.AddItem UCase(rs!cConsDescripcion)
      rs.MoveNext
   Loop
   cboMesIni.ListIndex = 0
   cboMesFin.ListIndex = 0
End If
oCon.CierraConexion

cboBS3.Clear
sSQL = "select cBSCod, cBSDescripcion " & _
       " from BienesServicios where len(rtrim(cBSCod))=3 and cBSCod like '11%'"
       
oCon.AbreConexion
Set rs = oCon.CargaRecordSet(sSQL)
If Not rs.EOF Then
   cboBS3.AddItem "--------------"
   Do While Not rs.EOF
      cboBS3.AddItem rs!cBSCod + " - " + rs!cBSDescripcion
      rs.MoveNext
   Loop
   cboBS3.ListIndex = 0
End If
oCon.CierraConexion
'sSQL = "select * from BienesServicios where len(rtrim(cBSCod))=8 and cBSCod like '11%'"
'sSQL = "select * from BienesServicios where len(rtrim(cBSCod))>8 and cBSCod like '11%'"
End Sub

Private Sub cboBS3_Click()
Dim rs As ADODB.Recordset
Dim oCon As DConecta
Set oCon = New DConecta
Set rs = New ADODB.Recordset

LimpiaFlex 0, 0, 0
cboBS5.Clear
cCod = Mid(cboBS3.Text, 1, 3)

sSQL = "select cBSCod, cBSDescripcion " & _
       " from BienesServicios where len(rtrim(cBSCod))=5 and cBSCod like '" & cCod & "%'"

oCon.AbreConexion
Set rs = oCon.CargaRecordSet(sSQL)
If Not rs.EOF Then
   cboBS5.AddItem "--------------"
   Do While Not rs.EOF
      cboBS5.AddItem rs!cBSCod + " - " + rs!cBSDescripcion
      rs.MoveNext
   Loop
   cboBS5.ListIndex = 0
End If
oCon.CierraConexion
End Sub

Private Sub cboBS5_Click()
Dim rs As ADODB.Recordset
Dim oCon As DConecta
Set oCon = New DConecta
Set rs = New ADODB.Recordset
 
LimpiaFlex 0, 0, 0
cboBS8.Clear
cCod = Mid(cboBS5.Text, 1, 5)

sSQL = "select cBSCod, cBSDescripcion " & _
       " from BienesServicios where len(rtrim(cBSCod))=8 and cBSCod like '" & cCod & "%'"

oCon.AbreConexion
Set rs = oCon.CargaRecordSet(sSQL)
If Not rs.EOF Then
   cboBS8.AddItem "--------------"
   Do While Not rs.EOF
      cboBS8.AddItem rs!cBSCod + " - " + rs!cBSDescripcion
      rs.MoveNext
   Loop
   cboBS8.ListIndex = 0
End If
oCon.CierraConexion
End Sub

Private Sub cboBS8_Click()
Dim rs As ADODB.Recordset
Dim oCon As DConecta
Set oCon = New DConecta

Set rs = New ADODB.Recordset

LimpiaFlex 0, 0, 0
cboBSF.Clear
cCod = Mid(cboBS8.Text, 1, 8)

sSQL = "select cBSCod, cBSDescripcion " & _
       " from BienesServicios where len(rtrim(cBSCod))>8 and cBSCod like '" & cCod & "%'"

oCon.AbreConexion
Set rs = oCon.CargaRecordSet(sSQL)
If Not rs.EOF Then
   cboBSF.BackColor = "&H80000005"
   cboBSF.AddItem "--------------"
   Do While Not rs.EOF
      cboBSF.AddItem rs!cBSCod + " - " + rs!cBSDescripcion
      rs.MoveNext
   Loop
   cboBSF.ListIndex = 0
End If
oCon.CierraConexion
End Sub

Private Sub cmdGenerar_Click()
Dim k As Integer

If cboMesIni.ListIndex = 0 Then
   MsgBox "Debe indicar el mes desde el cual se desea examinar..." + Space(10), vbInformation
   cboMesIni.SetFocus
   Exit Sub
End If
If cboMesFin.ListIndex = 0 Then
   MsgBox "Debe indicar el mes hasta el cual se desea examinar..." + Space(10), vbInformation
   cboMesFin.SetFocus
   Exit Sub
End If

If Len(txtAnioIni) < 0 Or txtAnioIni < 2001 Then
   MsgBox "el Año inicial indicado no es válido..." + Space(10), vbInformation
   Exit Sub
End If

If Len(txtAnioFin) < 0 Or txtAnioFin < 2001 Then
   MsgBox "el Año final indicado no es válido..." + Space(10), vbInformation
   Exit Sub
End If


cCod = ""
If cboBS3.ListIndex > 0 Then
   k = InStr(cboBS3.Text, "-")
   cCod = RTrim(Mid(cboBS3.Text, 1, k - 1))
   If cboBS5.ListIndex > 0 Then
      k = InStr(cboBS5.Text, "-")
      cCod = RTrim(Mid(cboBS5.Text, 1, k - 1))
      If cboBS8.ListIndex > 0 Then
         k = InStr(cboBS8.Text, "-")
         cCod = RTrim(Mid(cboBS8.Text, 1, k - 1))
         If cboBSF.ListIndex > 0 Then
            k = InStr(cboBSF.Text, "-")
            cCod = RTrim(Mid(cboBSF.Text, 1, k - 1))
         End If
      End If
   End If
End If

If Len(Trim(cCod)) > 0 Then
   GeneraListaBS cCod, Format(cboMesIni.ListIndex, "00"), txtAnioIni, Format(cboMesFin.ListIndex, "00"), txtAnioFin
End If
End Sub

Sub GeneraListaBS(vBSCod As String, vMesIni As String, vAnioIni As String, vMesFin As String, vAnioFin As String)
Dim oCon As DConecta
Set oCon = New DConecta

Dim i As Integer, k As Integer, nMax As Currency
Dim rs As ADODB.Recordset, nMeses As Integer
Dim nMaxCol As Integer
Dim nSuma As Currency
Dim n As Integer
Dim nCol As Integer

Dim oALmacen As DLogAlmacen
Set oALmacen = New DLogAlmacen
Dim lnStock As Double

Set rs = New ADODB.Recordset

If vAnioIni = vAnioFin And vMesIni > vMesFin Then
   MsgBox "El mes inicial no puede ser mayor al mes final..." + Space(10), vbInformation
   Exit Sub
End If

If vAnioIni = vAnioFin And vAnioFin + vMesFin >= Format(Date, "YYYYMM") Then
   MsgBox "El mes " & mMes(CInt(vMesFin)) & " no tiene saldos todavía..." + Space(10), vbInformation
   Exit Sub
End If

nMeses = CalculaNroMeses(vMesIni, vAnioIni, vMesFin, vAnioFin)

LimpiaFlex CInt(vMesIni), CInt(vAnioIni), nMeses
DoEvents
nMaxCol = flxObj.Cols - 1
oCon.AbreConexion
Set rs = oCon.Ejecutar("paGetConsumoAlmacen '" & vBSCod & "','" & vAnioIni + vMesIni & "','" & vAnioFin + vMesFin & "',0")
If Not rs.EOF Then
   i = 0
   Do While Not rs.EOF
      i = i + 1
      InsRow flxObj, i
      flxObj.TextMatrix(i, 0) = Format(i, "000")
      flxObj.TextMatrix(i, 1) = rs!cBSCod
      flxObj.TextMatrix(i, 2) = rs!cBSDescripcion
      flxObj.TextMatrix(i, 3) = rs!cUnidad
      'lnStock = oALmacen.GetStock("-1", flxObj.TextMatrix(i, 1), 0)
      If rs!nStock > 0 Then
         flxObj.TextMatrix(i, nMaxCol - 2) = Format(rs!nMonto / rs!nStock, "##,##0.00")
         flxObj.TextMatrix(i, nMaxCol - 1) = Format(rs!nStock, "##,##0.00")
      End If
      'flxObj.TextMatrix(i, nMaxCol - 2) = rs!nStock
      rs.MoveNext
   Loop
End If

Set rs = oCon.Ejecutar("paGetConsumoAlmacen '" & vBSCod & "','" & vAnioIni + vMesIni & "','" & vAnioFin + vMesFin & "',1")
If Not rs.EOF Then
   Do While Not rs.EOF
      flxObj.TextMatrix(rs!nBBSS, 3 + rs!nMMAA) = rs!nCant
      rs.MoveNext
   Loop
End If
oCon.CierraConexion

For k = 1 To flxObj.Rows - 1
    n = 0
    nMax = 0
    nSuma = 0
    For i = 4 To nMaxCol - 6
        If flxObj.ColWidth(i) > 10 And Len(flxObj.TextMatrix(k, i)) > 0 Then
           If CCur(flxObj.TextMatrix(k, i)) > nMax Then
              nMax = flxObj.TextMatrix(k, i)
           End If
           nSuma = nSuma + CCur(flxObj.TextMatrix(k, i))
        End If
        n = n + 1
    Next i
    flxObj.TextMatrix(k, nMaxCol - 5) = Format(nSuma / n, "###,##0.00")
    flxObj.TextMatrix(k, nMaxCol - 4) = nMax
    flxObj.TextMatrix(k, nMaxCol - 3) = nSuma
    'flxObj.TextMatrix(k, nmaxcol-1) = nMax
Next k
End Sub

Function GetStockBSSaldos(ByVal pnMes As Integer, ByVal pnAnio As Integer, ByVal psBSCod As String) As Long
Dim sSQL As String
Dim rs As New ADODB.Recordset
Dim oConn As DConecta
Set oConn = New DConecta

GetStockBSSaldos = 0
sSQL = "select nStock from BSSaldos where year(dSaldo)=" & pnAnio & " and " & _
       "       cBSCod = '" & psBSCod & "' " & _
       "       and month(dSaldo)=(select max(month(dSaldo)) from BSSaldos where year(dSaldo)=" & pnAnio & " and month(dSaldo)<=" & pnMes & " and cBSCod = '" & psBSCod & "')"

If oConn.AbreConexion Then
   Set rs = oConn.CargaRecordSet(sSQL)
   If Not rs.EOF Then
      GetStockBSSaldos = rs!nStock
   End If
End If
End Function

Function CalculaNroMeses(vMesIni As String, vAnioIni As String, vMesFin As String, vAnioFin As String) As Integer
Dim nMes As Integer, nAnio As Integer, nMeses As Integer
Dim Sigue As Boolean

nMes = CInt(vMesIni)
nAnio = CInt(vAnioIni)

If CInt(vMesIni) = CInt(vMesFin) And CInt(vAnioIni) = CInt(vAnioFin) Then
   CalculaNroMeses = 1
   Exit Function
End If

CalculaNroMeses = 0
Sigue = True
nMeses = 0
Do While Sigue
   nMeses = nMeses + 1
   nMes = nMes + 1
   If nMes = 13 Then
      nAnio = nAnio + 1
      nMes = 1
   End If
   If nMes = CInt(vMesFin) And nAnio = CInt(vAnioFin) Then
      nMeses = nMeses + 1
      Sigue = False
   End If
Loop
CalculaNroMeses = nMeses
End Function

Sub LimpiaFlex(nMesIni As Integer, nAnioIni As Integer, nMeses As Integer)
Dim i As Integer, k As Integer
Dim nMes As Integer, nAnio As Integer

flxObj.Clear
flxObj.Rows = 2
flxObj.Cols = nMeses + 10
flxObj.RowHeight(0) = 320
flxObj.RowHeight(1) = 10
flxObj.ColWidth(0) = 350:  flxObj.ColAlignment(0) = 4:  flxObj.TextMatrix(0, 0) = "  Nº"
flxObj.ColWidth(1) = 850:  flxObj.ColAlignment(1) = 4:  flxObj.TextMatrix(0, 1) = "  Codigo"
flxObj.ColWidth(2) = 2000: flxObj.TextMatrix(0, 2) = " Descripción "
flxObj.ColWidth(3) = 1000: flxObj.TextMatrix(0, 3) = " Unidad"
If nMeses = 0 Or nMesIni = 0 Or nAnioIni = 0 Then
   Exit Sub
End If
'--------------------------------------------------------
nMes = nMesIni
nAnio = nAnioIni
k = 3
For i = 1 To nMeses
    k = k + 1
    flxObj.TextMatrix(0, k) = mMes(nMes) + " - " + CStr(nAnio)
    nMes = nMes + 1
    If nMes > 12 Then
       nAnio = nAnio + 1
       nMes = 1
    End If
    flxObj.ColWidth(k) = 900
    flxObj.ColAlignment(k) = 4
Next
k = k + 1
flxObj.ColWidth(k) = 1000:     flxObj.TextMatrix(0, k) = "Cons. Promedio"
flxObj.ColWidth(k + 1) = 1000: flxObj.TextMatrix(0, k + 1) = "Cons. Máximo"
flxObj.ColWidth(k + 2) = 1000: flxObj.TextMatrix(0, k + 2) = " Cons. TOTAL"
flxObj.ColWidth(k + 3) = 1000: flxObj.TextMatrix(0, k + 3) = " Precio Prom"
flxObj.ColWidth(k + 4) = 1000: flxObj.TextMatrix(0, k + 4) = " Stock Actual"
End Sub

