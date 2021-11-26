VERSION 5.00
Begin VB.Form frmMuestraOpeTramNeg 
   Caption         =   "Consulta de Operaciones con cuentas de tramite de negocio"
   ClientHeight    =   7065
   ClientLeft      =   -180
   ClientTop       =   1710
   ClientWidth     =   11880
   Icon            =   "frmMuestraOpeTramNeg.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7065
   ScaleWidth      =   11880
   Begin VB.TextBox txtAnio 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   7305
      TabIndex        =   10
      Text            =   "2005"
      Top             =   120
      Width           =   840
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   150
      TabIndex        =   9
      Top             =   6540
      Width           =   1575
   End
   Begin VB.CommandButton cmdsalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   405
      Left            =   11010
      TabIndex        =   8
      Top             =   6495
      Width           =   1485
   End
   Begin VB.CommandButton cmdExcel 
      Caption         =   "Migrar Excel"
      Height          =   405
      Left            =   9480
      TabIndex        =   7
      Top             =   6495
      Width           =   1485
   End
   Begin VB.CommandButton cmdProcesar 
      Caption         =   "&Procesar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10845
      TabIndex        =   6
      Top             =   120
      Width           =   1575
   End
   Begin Sicmact.FlexEdit FlexEdit1 
      Height          =   5655
      Left            =   180
      TabIndex        =   5
      Top             =   720
      Width           =   12270
      _ExtentX        =   21643
      _ExtentY        =   9975
      Cols0           =   14
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   $"frmMuestraOpeTramNeg.frx":030A
      EncabezadosAnchos=   "450-2500-1200-1200-1000-1000-2500-4500-900-4500-2500-1200-2500-1200"
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X-X-X-X"
      ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-L-C-C-C-L-L-L-C-L-L-L-L-R"
      FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0-0-0-0-2"
      TextArray0      =   "Item"
      lbUltimaInstancia=   -1  'True
      ColWidth0       =   450
      RowHeight0      =   300
      ForeColorFixed  =   -2147483645
   End
   Begin VB.ComboBox cboMes 
      Height          =   315
      Left            =   4635
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   105
      Width           =   1905
   End
   Begin VB.TextBox txtCtaCont 
      Height          =   330
      Left            =   1950
      TabIndex        =   1
      Top             =   97
      Width           =   1545
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Año:"
      Height          =   195
      Left            =   6810
      TabIndex        =   3
      Top             =   165
      Width           =   330
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Mes:"
      Height          =   195
      Left            =   4110
      TabIndex        =   2
      Top             =   165
      Width           =   345
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cuenta Contable:"
      Height          =   195
      Left            =   585
      TabIndex        =   0
      Top             =   165
      Width           =   1230
   End
End
Attribute VB_Name = "frmMuestraOpeTramNeg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboMes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.txtAnio.SetFocus
End If
End Sub

Private Sub cmdCancelar_Click()
Me.txtCtaCont = ""
Me.FlexEdit1.Clear
Me.FlexEdit1.Rows = 2
Me.FlexEdit1.FormaCabecera
Me.cboMes.ListIndex = -1

End Sub

Private Sub cmdExcel_Click()
MigrarExcell
End Sub

Private Sub cmdProcesar_Click()
If txtCtaCont = "" Then
    MsgBox "Cuenta Contable no Ingresada", vbInformation, "aviso"
    Exit Sub
End If
If Me.cboMes = "" Then
    MsgBox "Mes no válido", vbInformation, "Aviso"
    Exit Sub
End If

CargaTramites Trim(txtCtaCont), Format(Me.cboMes.ListIndex + 1, "00"), txtAnio
 
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
CentraForm Me
Me.cboMes.AddItem "ENERO"
Me.cboMes.AddItem "FEBRERO"
Me.cboMes.AddItem "MARZO"
Me.cboMes.AddItem "ABRIL"
Me.cboMes.AddItem "MAYO"
Me.cboMes.AddItem "JUNIO"
Me.cboMes.AddItem "JULIO"
Me.cboMes.AddItem "AGOSTO"
Me.cboMes.AddItem "SETIEMBRE"
Me.cboMes.AddItem "OCTUBRE"
Me.cboMes.AddItem "NOVIEMBRE"
Me.cboMes.AddItem "DICIEMBRE"

txtAnio = Year(gdFecSis)
End Sub

Private Sub txtAnio_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
If KeyAscii = 13 Then
    Me.cmdProcesar.SetFocus
End If
End Sub

Private Sub txtCtaCont_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
If KeyAscii = 13 Then
    Me.cboMes.SetFocus
End If
End Sub
Sub CargaTramites(ByVal lsCtaContCod As String, ByVal lsMes As String, ByVal lsAnio As String)
Dim oCon As DConecta
Set oCon = New DConecta
Dim sql As String
Dim rs As ADODB.Recordset

lsCtaContCod = Left(lsCtaContCod, 2) & "M" & Mid(lsCtaContCod, 4, Len(lsCtaContCod))

sql = "     SELECT * "
sql = sql & "   From "
sql = sql & "       (SELECT  'OPERACIONES CAPTACIONES' AS [TIPO OPERACION], M.NMOVNRO AS [Nro Mov],"
sql = sql & "               dbo.GetFechaMov(M.CMOVNRO,103) AS [FECHA MOV], right(M.CMOVNRO,4) as [USUARIO] , M.COPECOD as [COD.OPE.], O.COPEDESC AS [DESC.OPERACION], M.CMOVDESC AS [DESC.MOV.], SUBSTRING(MC.CCTACOD,9,1) AS MONEDA,"
sql = sql & "               Titular = (    SELECT TOP 1 CPERSNOMBRE "
sql = sql & "                               FROM    PRODUCTOPERSONA R"
sql = sql & "                               JOIN PERSONA P ON P.CPERSCOD = R.CPERSCOD AND nPrdPersRelac=10"
sql = sql & "                               WHERE   R.CCTACOD = MC.CCTACOD ),"
sql = sql & "               MC.CCTACOD AS [Nro CUENTA] , 0 AS [CONCEPTO], '' AS [DESC.CONCEPTO],  MC.NMONTO As MONTO"
sql = sql & "           FROM    MOV M"
sql = sql & "                   JOIN MOVCAP MC ON MC.NMOVNRO = M.NMOVNRO"
sql = sql & "                   JOIN OPETPO O ON O.COPECOD = M.COPECOD"
sql = sql & "           WHERE   M.COPECOD IN (  SELECT  DISTINCT COPECOD"
sql = sql & "                                   From    OPECTANEG"
sql = sql & "                                   WHERE   CCTACONTCOD LIKE '" & lsCtaContCod & "%')"
sql = sql & "                   AND LEFT(M.CMOVNRO,6) = '" & lsAnio & lsMes & "' AND M.NMOVFLAG = 0 AND NOT MC.COPECOD LIKE '99%'"
sql = sql & "                   AND NMOVFLAG = 0"
sql = sql & "       Union"
sql = sql & "       SELECT  'OPERACIONES COLOCACIONES' AS CTIPOPE, M.NMOVNRO, dbo.GetFechaMov(M.CMOVNRO,103) AS CFECHA, right(M.CMOVNRO,4) as cUser ,"
sql = sql & "                MD.COPECOD, O.COPEDESC, M.CMOVDESC, SUBSTRING(MD.CCTACOD,9,1) AS NMONEDA,"
sql = sql & "                cTitular = (    SELECT TOP 1 CPERSNOMBRE"
sql = sql & "                               FROM    PRODUCTOPERSONA R"
sql = sql & "                               JOIN PERSONA P ON P.CPERSCOD = R.CPERSCOD AND nPrdPersRelac=20"
sql = sql & "                               WHERE   R.CCTACOD = MD.CCTACOD ),"
sql = sql & "                MD.CCTACOD ,"
sql = sql & "                PC.nPrdconceptoCod, PC.cDescripcion,"
sql = sql & "                SUM(MD.nMonto) As nMonto"
sql = sql & "        FROM    MOV M"
sql = sql & "                JOIN MOVCOLDET MD ON MD.nMovNro = M.NMOVNRO"
sql = sql & "                JOIN OPETPO O ON O.COPECOD = MD.COPECOD"
sql = sql & "                JOIN OPECTANEG OG ON OG.COPECOD = MD.COPECOD AND  CCTACONTCOD LIKE '" & lsCtaContCod & "%' AND nConcepto = MD.nPrdconceptoCod"
sql = sql & "                JOIN PRODUCTOCONCEPTO PC ON PC.nPrdconceptoCod = MD.nPrdconceptoCod"
sql = sql & "        WHERE   LEFT(M.CMOVNRO,6) = '" & lsAnio & lsMes & "' AND M.NMOVFLAG = 0"
sql = sql & "        GROUP BY M.NMOVNRO, M.CMOVNRO, MD.COPECOD,COPEDESC,CMOVDESC,PC.cDescripcion, PC.nPrdconceptoCod, CCTACOD"
sql = sql & "        Union"
sql = sql & "        SELECT     'OTRAS OPERACIONES' AS CTIPOPE,M.NMOVNRO, dbo.GetFechaMov(M.CMOVNRO,103) AS CFECHA, right(M.CMOVNRO,4) as cUser ,"
sql = sql & "                   M.COPECOD, O.COPEDESC, M.CMOVDESC, MO.NMONEDA, P.CPERSNOMBRE, '' AS cCtaCod,"
sql = sql & "                   0 as nPrdConcepto, '' as cDescConcep , MO.NMOVIMPORTE  as nMonto"
sql = sql & "        FROM       MOV M"
sql = sql & "                   JOIN MOVOPEVARIAS MO ON MO.NMOVNRO = M.NMOVNRO"
sql = sql & "                   JOIN OPECTANEG OG ON OG.COPECOD = M.COPECOD"
sql = sql & "                   JOIN OPETPO O ON O.COPECOD = OG.COPECOD"
sql = sql & "                   LEFT JOIN MOVGASTO MG ON MG.NMOVNRO = M.NMOVNRO"
sql = sql & "                   LEFT JOIN PERSONA P ON P.CPERSCOD = MG.CPERSCOD"
sql = sql & "        WHERE      OG.CCTACONTCOD LIKE '" & lsCtaContCod & "%'"
sql = sql & "                   AND LEFT(M.CMOVNRO,6) = '" & lsAnio & lsMes & "' AND M.NMOVFLAG = 0  ) AS DATA "
sql = sql & "        ORDER BY [TIPO OPERACION], [FECHA MOV]"

Set oCon = New DConecta
oCon.AbreConexion

Set rs = oCon.CargaRecordSet(sql)
FlexEdit1.rsFlex = rs
oCon.CierraConexion
Set oCon = Nothing
End Sub
Sub MigrarExcell()
Dim vExcelObj As Excel.Application
Dim vNHC As String
Dim Fila As Long
Dim Col As Long
Dim lsCol As String
Dim lnFila As Long

If Me.FlexEdit1.TextMatrix(1, 0) = "" Then Exit Sub
'lsCadena = CabeRepo(gsNomCmac, lblDescAge, "", IIf(psMoneda = "1", "SOLES", "DOLARES"), Format(pdFecha, "dd/mm/yyyy"), "Control Calidad MOVIMIENTOS AHORROS", "USUARIO :" & psCodUsu, lsInfoSicmac_Siafc, "", lnPag, 64)
'lsCadena = lsCadena + Chr(10)
                               
vNHC = App.path & "\spooler\OPTRAM" & txtCtaCont & "_" & Trim(cboMes) & txtAnio & ".XLS"

Set vExcelObj = New Excel.Application  '   = CreateObject("Excel.Application")
vExcelObj.DisplayAlerts = False

vExcelObj.Workbooks.Add
vExcelObj.Sheets("Hoja1").Select
vExcelObj.Sheets("Hoja1").Name = "OPETRAMNEG"

vExcelObj.Range("A1:IV65536").Font.Name = "Arial Narrow"
vExcelObj.Range("A1:IV65536").Font.Size = 8
vExcelObj.Columns("A:IV").Select
vExcelObj.Selection.VerticalAlignment = 3

vExcelObj.Range("A1").Select
vExcelObj.Range("A1").Font.Bold = True
vExcelObj.Range("A1").HorizontalAlignment = 1
vExcelObj.ActiveCell.value = UCase(Trim(gsNomCmac))

vExcelObj.Range("N1").Select
vExcelObj.Range("N1").Font.Bold = True
vExcelObj.Range("N1").HorizontalAlignment = 1
vExcelObj.ActiveCell.value = Trim(cboMes) & "-" & Me.txtAnio

vExcelObj.Range("A4").Select
vExcelObj.Range("A4").Font.Bold = True
vExcelObj.Range("A4").HorizontalAlignment = 1
vExcelObj.ActiveCell.value = "REPORTE DE OPERACIONES DE TRAMITE NEGOCIO  - CUENTA CONT." & txtCtaCont
Dim lnAlinCol As XlHAlign
Dim lnAnchoCol As Long
Dim vCel As String
Dim vItem As Long
Dim vIni As Long


vExcelObj.Range("A7:M7").Select
vExcelObj.Range("A7:M7").Font.Bold = True
vExcelObj.Selection.AutoFilter

vIni = 6
vItem = vIni
For Fila = 0 To Me.FlexEdit1.Rows - 1
    vItem = vItem + 1
    For Col = 0 To Me.FlexEdit1.Cols - 1
        Select Case Col
            Case Is = 0
                lsCol = "A"
                lnAnchoCol = 5
                lnAlinCol = xlHAlignCenter
            Case Is = 1
                lsCol = "B"
                lnAnchoCol = 25
                lnAlinCol = xlHAlignLeft
            Case Is = 2
                lsCol = "C"
                lnAnchoCol = 7
                lnAlinCol = xlHAlignLeft
            Case Is = 3
                lsCol = "D"
                lnAnchoCol = 9
                lnAlinCol = xlHAlignCenter
            Case Is = 4
                lsCol = "E"
                lnAnchoCol = 6
                lnAlinCol = xlHAlignCenter
            Case Is = 5
                lsCol = "F"
                lnAnchoCol = 6
                lnAlinCol = xlHAlignCenter
            Case Is = 6
                lsCol = "G"
                lnAnchoCol = 45
                lnAlinCol = xlHAlignLeft
            Case Is = 7
                lsCol = "H"
                lnAnchoCol = 120
                lnAlinCol = xlHAlignLeft
            Case Is = 8
                lsCol = "I"
                lnAnchoCol = 6
                lnAlinCol = xlHAlignCenter
            Case Is = 9
                lsCol = "J"
                lnAnchoCol = 70
                lnAlinCol = xlHAlignLeft
            Case Is = 10
                lsCol = "K"
                lnAnchoCol = 15
                lnAlinCol = xlHAlignCenter
            Case Is = 11
                lsCol = "L"
                lnAnchoCol = 15
                lnAlinCol = xlHAlignCenter
            Case Is = 12
                lsCol = "M"
                lnAnchoCol = 35
                lnAlinCol = xlHAlignLeft
            Case Is = 13
                lsCol = "N"
                lnAnchoCol = 10
                lnAlinCol = xlHAlignRight
        End Select
        vCel = lsCol + Trim(Str(vItem))
        vExcelObj.Range(vCel).Select
        If Fila = 0 Then
            vExcelObj.Range(vCel).Font.Bold = True
        End If
        vExcelObj.Range(vCel).HorizontalAlignment = lnAlinCol
        vExcelObj.Range(vCel).ColumnWidth = lnAnchoCol
        
        If Col = 10 And Fila > 0 Then
            vExcelObj.ActiveCell.value = "'" & FlexEdit1.TextMatrix(Fila, Col)
        Else
            If Col = 13 And Fila > 0 Then
                vExcelObj.ActiveCell.value = Format(FlexEdit1.TextMatrix(Fila, Col), "#,#0.00")
            Else
                If Col = 3 And Fila > 0 Then
                    vExcelObj.ActiveCell.value = "'" & Format(FlexEdit1.TextMatrix(Fila, Col), "dd/mm/yyyy")
                Else
                    vExcelObj.ActiveCell.value = Me.FlexEdit1.TextMatrix(Fila, Col)
                End If
            End If
        End If
    Next
Next
If Dir(vNHC) <> "" Then
   If MsgBox("Archivo Ya Existe ...  Desea Reemplazarlo ??", vbQuestion + vbYesNo + vbDefaultButton1, " Mensaje del Sistema ...") = vbNo Then
      Exit Sub
   End If
End If
vExcelObj.Range("A1").Select
vExcelObj.ActiveWorkbook.SaveAs (vNHC)
vExcelObj.ActiveWorkbook.Close
MsgBox "SE HA GENERADO CON ÉXITO EL ARCHIVO !!  ", vbInformation, " Mensaje del Sistema ..."
vExcelObj.Workbooks.Open (vNHC)
vExcelObj.Visible = True

Set vExcelObj = Nothing

End Sub
