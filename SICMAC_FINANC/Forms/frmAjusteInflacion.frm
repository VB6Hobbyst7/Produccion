VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAjusteInflacion 
   Caption         =   "Ajuste por Inflación"
   ClientHeight    =   5775
   ClientLeft      =   990
   ClientTop       =   2040
   ClientWidth     =   9600
   Icon            =   "frmAjusteInflacion.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   9600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   465
      Left            =   7650
      TabIndex        =   16
      Top             =   795
      Width           =   1755
      Begin VB.CommandButton cmdTodo 
         Caption         =   "&Todo"
         Height          =   255
         Left            =   60
         TabIndex        =   18
         Top             =   150
         Width           =   795
      End
      Begin VB.CommandButton cmdNada 
         Caption         =   "&Ninguno"
         Height          =   255
         Left            =   900
         TabIndex        =   17
         Top             =   150
         Width           =   795
      End
   End
   Begin Sicmact.FlexEdit fgIndice 
      Height          =   3975
      Left            =   180
      TabIndex        =   14
      Top             =   1230
      Width           =   3945
      _ExtentX        =   6959
      _ExtentY        =   7011
      Rows            =   13
      Cols0           =   3
      HighLight       =   2
      AllowUserResizing=   3
      EncabezadosNombres=   "-Mes-Factor"
      EncabezadosAnchos=   "300-2050-1500"
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
      ColumnasAEditar =   "X-X-X"
      TextStyleFixed  =   3
      ListaControles  =   "0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-L-R"
      FormatosEdit    =   "0-0-0"
      SelectionMode   =   1
      lbUltimaInstancia=   -1  'True
      ColWidth0       =   300
      RowHeight0      =   285
      ForeColorFixed  =   -2147483630
   End
   Begin VB.CommandButton cmdAsiento 
      Caption         =   "&Asiento"
      Enabled         =   0   'False
      Height          =   360
      Left            =   5610
      TabIndex        =   4
      Top             =   5340
      Width           =   1200
   End
   Begin VB.ComboBox cboDec 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      ItemData        =   "frmAjusteInflacion.frx":030A
      Left            =   5250
      List            =   "frmAjusteInflacion.frx":0320
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   390
      Width           =   735
   End
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   4230
      TabIndex        =   10
      Top             =   120
      Width           =   1935
      Begin VB.Label Label3 
         Caption         =   "Decimales"
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Top             =   330
         Width           =   795
      End
   End
   Begin VB.CommandButton cmdProcesar 
      Caption         =   "&Procesar"
      Height          =   360
      Left            =   6840
      TabIndex        =   5
      Top             =   5340
      Width           =   1200
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      CausesValidation=   0   'False
      Height          =   360
      Left            =   8070
      TabIndex        =   6
      Top             =   5340
      Width           =   1200
   End
   Begin VB.Frame Frame3 
      Caption         =   "&Periodo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   735
      Left            =   180
      TabIndex        =   7
      Top             =   120
      Width           =   3900
      Begin VB.ComboBox cboMes 
         Height          =   315
         ItemData        =   "frmAjusteInflacion.frx":0336
         Left            =   2100
         List            =   "frmAjusteInflacion.frx":035E
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   270
         Width           =   1665
      End
      Begin VB.TextBox txtAnio 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   630
         MaxLength       =   4
         TabIndex        =   1
         Top             =   270
         Width           =   855
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Mes :"
         Height          =   195
         Left            =   1680
         TabIndex        =   13
         Top             =   315
         Width           =   390
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Año :"
         Height          =   195
         Left            =   150
         TabIndex        =   12
         Top             =   330
         Width           =   375
      End
   End
   Begin MSComctlLib.ListView lvCta 
      Height          =   3825
      Left            =   4350
      TabIndex        =   3
      Top             =   1320
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   6747
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Cuenta "
         Object.Width           =   2187
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Descripción"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "CateCta"
         Object.Width           =   0
      EndProperty
   End
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
      ForeColor       =   &H8000000D&
      Height          =   4065
      Left            =   4230
      TabIndex        =   9
      Top             =   1140
      Width           =   5175
   End
   Begin VB.Label Label2 
      Caption         =   "Cuentas de Ajuste"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   4230
      TabIndex        =   15
      Top             =   960
      Width           =   2685
   End
   Begin VB.Label Label1 
      Caption         =   "Indice de Precios al por Mayor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   225
      Left            =   180
      TabIndex        =   8
      Top             =   960
      Width           =   2805
   End
End
Attribute VB_Name = "frmAjusteInflacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xlAplicacion  As Excel.Application
Dim xlLibro       As Excel.Workbook
Dim xlHoja1       As Excel.Worksheet

Dim sSql As String
Dim rs   As ADODB.Recordset
Dim N    As Integer

Dim aAsiento() As String

Private Sub MuestraIPM()
Dim nIndiceMes As Currency
Dim oIPM As DAjusteCont
If txtAnio = "" Or cboMes.ListIndex = -1 Then
   Exit Sub
End If
fgIndice.TextMatrix(1, 1) = "Dic-" & (txtAnio - 1)
For nIndiceMes = 1 To fgIndice.Rows - 1
   fgIndice.TextMatrix(nIndiceMes, 2) = ""
Next
Set oIPM = New DAjusteCont
Set rs = oIPM.CargaIPM(, " Year(dFecha) = " & txtAnio & " and Month(dFecha) = " & cboMes.ListIndex + 1)
If rs.EOF Then
   MsgBox "Aún no se define Indice del Mes requerido", vbInformation, "Aviso"
   RSClose rs
   Exit Sub
End If
nIndiceMes = rs!nValor

Set rs = oIPM.CargaIPM("12-31-" & (txtAnio - 1))
If Not rs.EOF Then
   fgIndice.TextMatrix(1, 2) = Format(Round(nIndiceMes / rs!nValor, Val(cboDec.Text)), "#,##0." & String(cboDec, "0"))
End If

Set rs = oIPM.CargaIPM(, "Year(dFecha) = " & txtAnio & " and Month(dfecha) <= " & cboMes.ListIndex + 1)
Do While Not rs.EOF
   fgIndice.TextMatrix(Month(rs!dFecha) + 1, 2) = Format(Round(nIndiceMes / rs!nValor, Val(cboDec.Text)), "#,##0." & String(cboDec, "0"))
   rs.MoveNext
Loop
RSClose rs
Set oIPM = Nothing
End Sub

Private Sub CabeceraExcel(sCtaCod As String, sCtaDes As String)
xlHoja1.Cells(1, 2) = gsNomCmac
xlHoja1.Cells(3, 2) = "CEDULA DE AJUSTE POR INFLACION PARA EL AÑO " & txtAnio
xlHoja1.Cells(4, 2) = "CUENTA : " & sCtaCod & ". " & sCtaDes
xlHoja1.Cells(6, 2) = "MES"
xlHoja1.Cells(6, 3) = "VALOR"
xlHoja1.Cells(7, 3) = "HISTORICO"
xlHoja1.Cells(6, 4) = "FACTOR DE"
xlHoja1.Cells(7, 4) = "AJUSTE"
xlHoja1.Cells(6, 5) = "VALOR"
xlHoja1.Cells(7, 5) = "AJUSTADO"
xlHoja1.Cells(6, 6) = "VARIACION"

xlHoja1.Range("A1:F7").Font.Bold = True

xlHoja1.Range("A6:F7").HorizontalAlignment = xlHAlignCenter
xlHoja1.Range(xlHoja1.Cells(6, 2), xlHoja1.Cells(7, 6)).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
xlHoja1.Range(xlHoja1.Cells(6, 2), xlHoja1.Cells(7, 6)).Borders(xlInsideVertical).LineStyle = xlContinuous
End Sub

Private Sub cboDec_Click()
MuestraIPM
End Sub

Private Sub cboMes_Click()
If cboMes.ListIndex = 11 Then
   cboDec.ListIndex = 2
Else
   cboDec.ListIndex = 3
End If
MuestraIPM
End Sub

Private Sub cmdAsiento_Click()
Dim nItem As Integer
Dim sImp  As String
Dim k     As Integer
On Error GoTo ErrAsiento

If UBound(aAsiento, 2) = 0 Then
   MsgBox "Es necesario Generar Cuadros de Ajuste para generar Asiento Contable!", vbInformation, "Aviso"
   cmdAsiento.Enabled = False
   Exit Sub
End If
If MsgBox(" ¿ Seguro que desea Guardar Asiento de Ajuste Generado ? ", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmación") = vbNo Then
   Exit Sub
End If
Dim oMov As DMov
Set oMov = New DMov
If oMov.BuscarMov(txtAnio & Format(cboMes.ListIndex + 1, "00"), "cOpeCod = '" & gsOpeCod & "' and nMovEstado = '" & gMovEstContabMovContable & "' and NOT nMovFlag IN ('" & gMovFlagDeExtorno & "','" & gMovFlagExtornado & "','" & gMovFlagEliminado & "') ") Then
   MsgBox "Asiento de Ajuste por Inflación ya fue generado", vbInformation, "Aviso"
   Exit Sub
End If

gdFecha = DateAdd("m", 1, CDate("01/" & Format(cboMes.ListIndex + 1, "00") & "/" & txtAnio)) - 1
gsMovNro = oMov.GeneraMovNro(gdFecha, Right(gsCodAge, 2), gsCodUser)
gsGlosa = "ASIENTO DE AJUSTE POR INFLACION DE INGRESOS/EGRESOS : " & cboMes.Text & " " & txtAnio

'GRABACION DE ASIENTO DE AJUSTE
oMov.BeginTrans
oMov.InsertaMov gsMovNro, gsOpeCod, gsGlosa
gnMovNro = oMov.GetnMovNro(gsMovNro)
oMov.InsertaMovCont gnMovNro, 0, 0, ""
nItem = 0
For k = 1 To UBound(aAsiento, 2)
    If Val(aAsiento(1, k)) <> 0 Then
        nItem = nItem + 1
        oMov.InsertaMovCta gnMovNro, nItem, aAsiento(0, k), aAsiento(1, k)
    End If
Next
oMov.ActualizaSaldoMovimiento gsMovNro, "+"
oMov.CommitTrans
Set oMov = Nothing

Dim oImp As New NContImprimir
sImp = oImp.ImprimeAsientoContable(gsMovNro, gnLinPage, gnColPage, "ASIENTO DE AJUSTE POR INFLACION", , , , gsNomCmac)
EnviaPrevio sImp, Me.Caption, gnLinPage, False
Set oImp = Nothing

Exit Sub
ErrAsiento:
   MsgBox TextErr(Err.Description), vbInformation, "Aviso"
   oMov.RollbackTrans
End Sub

Private Sub cmdNada_Click()
Dim N As Integer
For N = 1 To lvCta.ListItems.Count
   lvCta.ListItems(N).Checked = False
Next
cmdProcesar.SetFocus
End Sub

Private Sub cmdProcesar_Click()
Dim nItem     As Integer
Dim k         As Integer
Dim lsArchivo As String
Dim lbLibroOpen As Boolean
Dim rsCta As New ADODB.Recordset
Dim rsBal As New ADODB.Recordset
Dim sDH As String
Dim pbImprime As Boolean
Dim nResp As Integer
Dim oBarra As New clsProgressBar
On Error GoTo ErrProcesa


If MsgBox(" ¿ Seguro de Generar Hojas de Trabajo de Ajuste por Inflación ? ", vbQuestion + vbYesNo, "Confirmación") = vbNo Then
   Exit Sub
End If

Dim oBalance As New NBalanceCont
If Not oBalance.ValidaBalanceHistoricoAjuste(txtAnio, Format(cboMes.ListIndex + 1, "00")) Then
   Set oBalance = Nothing
   Exit Sub
End If

ReDim aAsiento(1, 0)
Me.Enabled = False
For nItem = 1 To lvCta.ListItems.Count
   If lvCta.ListItems(nItem).Checked Then
      sDH = lvCta.ListItems(nItem).SubItems(2)
      
      oBarra.ShowForm frmMdiMain
      oBarra.CaptionSyle = eCap_CaptionPercent
      oBarra.Max = 1
      oBarra.Progress 0, "AJUSTE POR INFLACION", "Captura de Cuentas...", "", vbBlue
      Dim oCont As New DAjusteCont
      Set rs = oCont.GetCuentasAjusteBalance(lvCta.ListItems(nItem), txtAnio, cboMes.ListIndex + 1, gsOpeCod)
      Set oCont = Nothing
      oBarra.Progress 1, "AJUSTE POR INFLACION", "Captura de Cuentas...", "", vbBlue
      
      If Not rs.EOF Then
         lsArchivo = App.path & "\Spooler\AI" & lvCta.ListItems(nItem).Text & "_" & txtAnio & Format(cboMes.ListIndex + 1, "0#") & ".xls"
         lbLibroOpen = ExcelBegin(lsArchivo, xlAplicacion, xlLibro)
         If lbLibroOpen Then
            Set xlHoja1 = xlLibro.Worksheets(1)
            oBarra.Max = rs.RecordCount
            oBarra.Progress 1, "AJUSTE POR INFLACION", "Captura de Cuentas...", "", vbBlue
            Do While Not rs.EOF
               oBarra.Progress rs.Bookmark, "AJUSTE POR INFLACION", "Generando Cuadros", "Cuenta " & rs!cCtaContCod, vbBlue
               
               Dim nFila As Integer, nFilaMax As Integer
               Dim sCtaCod As String
               
               Set rsCta = oBalance.CargaBalanceAjusteInflacion(rs!cCtaContCod, txtAnio, Format(cboMes.ListIndex + 1, "00"))
               
               nFila = 8
               nFilaMax = cboMes.ListCount + nFila
               gnImporte = 0
               If Not rsCta.EOF Then
                  Do While Not rsCta.EOF
                     ExcelAddHoja "Cta_" & rsCta!cCtaContCod, xlLibro, xlHoja1
                     CabeceraExcel rsCta!cCtaContCod, rs!cCtaContDesc
                     xlHoja1.Cells(nFila, 2) = "Dic-" & (txtAnio - 1)
                     xlHoja1.Cells(nFila, 4) = fgIndice.TextMatrix(1, 2)
                     
                     For k = 1 To cboMes.ListCount
                        xlHoja1.Cells(k + 8, 2) = cboMes.List(k - 1)
                        xlHoja1.Cells(k + 8, 4) = fgIndice.TextMatrix(k + 1, 2)
                     Next
                     sCtaCod = rsCta!cCtaContCod
                     Do While rsCta!cCtaContCod = sCtaCod
                        nFila = 8 + Val(rsCta!cBalancemes)
                        If Val(rsCta!cBalancemes) <> 0 Then
                            If sDH = "D" Then
                               xlHoja1.Cells(nFila, 3) = rsCta!nDebe - rsCta!nHaber
                            Else
                               xlHoja1.Cells(nFila, 3) = rsCta!nHaber - rsCta!nDebe
                            End If
                        Else
                               xlHoja1.Cells(nFila, 3) = rsCta!nDebe
                        End If
                        xlHoja1.Range(xlHoja1.Cells(nFila, 5), xlHoja1.Cells(nFila, 5)).Formula = "=ROUND(" + xlHoja1.Range(xlHoja1.Cells(nFila, 3), xlHoja1.Cells(nFila, 3)).Address(False, False) + "*" + xlHoja1.Range(xlHoja1.Cells(nFila, 4), xlHoja1.Cells(nFila, 4)).Address(False, False) + ",2)"
                        xlHoja1.Range(xlHoja1.Cells(nFila, 6), xlHoja1.Cells(nFila, 6)).Formula = "=" + xlHoja1.Range(xlHoja1.Cells(nFila, 5), xlHoja1.Cells(nFila, 5)).Address(False, False) + "-" + xlHoja1.Range(xlHoja1.Cells(nFila, 3), xlHoja1.Cells(nFila, 3)).Address(False, False)
                        rsCta.MoveNext
                        If rsCta.EOF Then
                           Exit Do
                        End If
                     Loop
                     xlHoja1.Range(xlHoja1.Cells(8, 2), xlHoja1.Cells(nFilaMax + 1, 6)).Borders.LineStyle = xlContinuous
                     xlHoja1.Range(xlHoja1.Cells(1, 2), xlHoja1.Cells(nFilaMax + 1, 6)).Font.Size = 8
                     xlHoja1.Range(xlHoja1.Cells(8, 3), xlHoja1.Cells(nFilaMax + 1, 6)).NumberFormat = "#,##0.00;-#,##0.00"
                     xlHoja1.Range(xlHoja1.Cells(8, 4), xlHoja1.Cells(nFilaMax + 1, 4)).NumberFormat = "#,##0." & String(cboDec.Text, "0") & ";-#,##0." & String(cboDec.Text, "0")
                     xlHoja1.Cells(nFilaMax + 1, 2) = "TOTAL"
                     xlHoja1.Range(xlHoja1.Cells(nFilaMax + 1, 2), xlHoja1.Cells(nFilaMax + 1, 2)).Font.Bold = True
                     If nFila > 7 Then
                        xlHoja1.Range(xlHoja1.Cells(nFilaMax + 1, 3), xlHoja1.Cells(nFilaMax + 1, 3)).Formula = "=SUM(" + xlHoja1.Range(xlHoja1.Cells(8, 3), xlHoja1.Cells(nFilaMax, 3)).Address(False, False) + ")"
                        xlHoja1.Range(xlHoja1.Cells(nFilaMax + 1, 5), xlHoja1.Cells(nFilaMax + 1, 5)).Formula = "=SUM(" + xlHoja1.Range(xlHoja1.Cells(8, 5), xlHoja1.Cells(nFilaMax, 5)).Address(False, False) + ")"
                        xlHoja1.Range(xlHoja1.Cells(nFilaMax + 1, 6), xlHoja1.Cells(nFilaMax + 1, 6)).Formula = "=SUM(" + xlHoja1.Range(xlHoja1.Cells(8, 6), xlHoja1.Cells(nFilaMax, 6)).Address(False, False) + ")"
                     End If
                     gnImporte = gnImporte + nVal(xlHoja1.Cells(nFilaMax + 1, 6))
                  Loop
                     
                  'ASIENTO CONTABLE
                  sCtaCod = ""
                  Dim oConA As New NContAsientos
                  sCtaCod = oConA.GetCtaEquivalente(rs!cCtaContCod, sDH, gsOpeCod)
                                  
                  gnImporte = gnImporte - oBalance.GetBalanceSaldoCuenta(rs!cCtaContCod, txtAnio, Format(cboMes.ListIndex, "00"))
                  
                  If sCtaCod <> "" Then
                     If sDH = "D" Then
                        oConA.LlenaArrayAsiento aAsiento, rs!cCtaContCod, gnImporte
                        oConA.LlenaArrayAsiento aAsiento, sCtaCod, gnImporte * -1
                     Else
                        oConA.LlenaArrayAsiento aAsiento, rs!cCtaContCod, gnImporte * -1
                        oConA.LlenaArrayAsiento aAsiento, sCtaCod, gnImporte
                     End If
                  Else
                     MsgBox "Cuenta Contable " & rs!cCtaContCod & " no tiene Equivalente en Operación", vbInformation, "Aviso"
                  End If
                  Set oConA = Nothing
               End If
               RSClose rsCta
               rs.MoveNext
            Loop
            ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1
            'CargaArchivo lsArchivo, App.path & "\SPOOLER"
         End If
      End If
      RSClose rs
      oBarra.CloseForm frmMdiMain
   End If
Next
MsgBox "Hojas de Trabajo generadas correctamente", vbInformation, "¡Aviso!"
Me.Enabled = True
Set oBalance = Nothing
Set oBarra = Nothing
cmdAsiento.Enabled = True
cmdAsiento.SetFocus
Exit Sub
ErrProcesa:
   MsgBox TextErr(Err.Description), vbInformation, "Aviso"
   
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub cmdTodo_Click()
Dim N As Integer
For N = 1 To lvCta.ListItems.Count
   lvCta.ListItems(N).Checked = True
Next
cmdProcesar.SetFocus
End Sub

Private Sub Form_Load()
Dim dFecha As String
Dim lvItem As ListItem
Dim oCons As New DConstantes
CentraForm Me
Set rs = oCons.CargaConstante(gMeses)
fgIndice.Rows = 14
fgIndice.Row = 1
fgIndice.CellBackColor = "&H00DBDBDB"
fgIndice.TextMatrix(1, 0) = 12
fgIndice.TextMatrix(1, 1) = "Dic-" & Year(gdFecSis) - 1

Do While Not rs.EOF
   N = rs!nConsValor
   fgIndice.Row = N + 1
   fgIndice.CellBackColor = "&H00DBDBDB"
   fgIndice.TextMatrix(N + 1, 0) = N
   fgIndice.TextMatrix(N + 1, 1) = rs!cConsDescripcion
   rs.MoveNext
Loop
Set oCons = Nothing

Dim oOpe  As New DOperacion
Set rs = oOpe.CargaOpeCta(gsOpeCod, , "0", True, True)
Do While Not rs.EOF
   Set lvItem = lvCta.ListItems.Add
   lvItem.Text = rs!cCtaContCod
   lvItem.SubItems(1) = rs!cCtaContDesc
   lvItem.SubItems(2) = rs!cCtaCaracter
   rs.MoveNext
Loop
RSClose rs
dFecha = DateAdd("m", 1, CDate(LeeConstanteSist(gConstSistCierreMensualCont)))
txtAnio = Year(dFecha)
cboMes.ListIndex = Month(dFecha) - 1
cboDec.ListIndex = 3
End Sub

Private Sub Form_Unload(Cancel As Integer)
CierraConexion
End Sub

Private Sub txtAnio_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
If KeyAscii = 13 Then
   If Not ValidaAnio(txtAnio) Then
      Exit Sub
   End If
   MuestraIPM
   cboMes.SetFocus
End If
End Sub

Private Sub txtAnio_Validate(Cancel As Boolean)
   If Not ValidaAnio(txtAnio) Then
      Cancel = True
   Else
      MuestraIPM
   End If
End Sub

