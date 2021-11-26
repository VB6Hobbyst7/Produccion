VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAjusteInflaHisto 
   Caption         =   "Ajuste por Inflación"
   ClientHeight    =   3990
   ClientLeft      =   2430
   ClientTop       =   2205
   ClientWidth     =   4305
   Icon            =   "frmAjusteInflaHisto.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   4305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar prgBar 
      Height          =   255
      Left            =   1530
      TabIndex        =   3
      Top             =   3690
      Visible         =   0   'False
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.StatusBar sBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   3615
      Width           =   4305
      _ExtentX        =   7594
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   9173
            MinWidth        =   9173
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdAsiento 
      Caption         =   "&Asiento"
      Enabled         =   0   'False
      Height          =   360
      Left            =   300
      TabIndex        =   10
      Top             =   2460
      Width           =   3660
   End
   Begin VB.Frame Frame4 
      Height          =   705
      Left            =   180
      TabIndex        =   4
      Top             =   960
      Width           =   3915
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
         ItemData        =   "frmAjusteInflaHisto.frx":030A
         Left            =   1440
         List            =   "frmAjusteInflaHisto.frx":0320
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   240
         Width           =   2325
      End
      Begin VB.Label Label3 
         Caption         =   "Decimales"
         Height          =   285
         Left            =   180
         TabIndex        =   5
         Top             =   300
         Width           =   795
      End
   End
   Begin VB.CommandButton cmdProcesar 
      Caption         =   "&Procesar"
      Height          =   360
      Left            =   300
      TabIndex        =   0
      Top             =   2040
      Width           =   3660
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      CausesValidation=   0   'False
      Height          =   360
      Left            =   300
      TabIndex        =   1
      Top             =   2880
      Width           =   3660
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
      TabIndex        =   2
      Top             =   120
      Width           =   3900
      Begin VB.ComboBox cboMes 
         Height          =   315
         ItemData        =   "frmAjusteInflaHisto.frx":0336
         Left            =   2100
         List            =   "frmAjusteInflaHisto.frx":035E
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   270
         Width           =   1665
      End
      Begin VB.TextBox txtAnio 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   630
         MaxLength       =   4
         TabIndex        =   6
         Top             =   270
         Width           =   855
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Mes :"
         Height          =   195
         Left            =   1680
         TabIndex        =   9
         Top             =   315
         Width           =   390
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Año :"
         Height          =   195
         Left            =   180
         TabIndex        =   8
         Top             =   330
         Width           =   375
      End
   End
   Begin RichTextLib.RichTextBox rtxt 
      Height          =   315
      Left            =   4590
      TabIndex        =   12
      Top             =   390
      Visible         =   0   'False
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   556
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"frmAjusteInflaHisto.frx":03C6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   1605
      Left            =   180
      TabIndex        =   13
      Top             =   1800
      Width           =   3915
   End
End
Attribute VB_Name = "frmAjusteInflaHisto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet

Dim sSql As String
Dim rs   As ADODB.Recordset
Dim N    As Integer
Dim aAsiento() As String

Private Sub CabeceraExcel(sCtaCod As String, sCtaDes As String)
xlHoja1.Cells(1, 2) = gsNomCmac
xlHoja1.Cells(3, 2) = "CEDULA DE AJUSTE POR INFLACION PARA EL AÑO " & txtAnio
xlHoja1.Cells(4, 2) = "CUENTA : " & sCtaCod & ". " & sCtaDes
xlHoja1.Cells(6, 2) = "FECHA"
xlHoja1.Cells(6, 3) = "DETALLE"
xlHoja1.Cells(6, 4) = "VALOR"
xlHoja1.Cells(7, 4) = "HISTORICO"
xlHoja1.Cells(6, 5) = "FACTOR DE"
xlHoja1.Cells(7, 5) = "AJUSTE"
xlHoja1.Cells(6, 6) = "VALOR"
xlHoja1.Cells(7, 6) = "AJUSTADO"
xlHoja1.Cells(6, 7) = "VARIACION"

xlHoja1.Range("A1:G7").Font.Bold = True

xlHoja1.Range("A6:G7").HorizontalAlignment = xlHAlignCenter
xlHoja1.Range("B1").ColumnWidth = 9
xlHoja1.Range("C1").ColumnWidth = 20
xlHoja1.Range(xlHoja1.Cells(6, 2), xlHoja1.Cells(7, 7)).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
xlHoja1.Range(xlHoja1.Cells(6, 2), xlHoja1.Cells(7, 7)).Borders(xlInsideVertical).LineStyle = xlContinuous

End Sub

Private Sub cboMes_Click()
If cboMes.ListIndex = 11 Then
   cboDec.ListIndex = 2
Else
   cboDec.ListIndex = 3
End If
End Sub

Private Sub cmdAsiento_Click()
Dim nItem As Integer
Dim sImp  As String
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
gsMovNro = oMov.GeneraMovNro(gdFecha, gsCodAge, gsCodUser)
gsGlosa = "ASIENTO DE AJUSTE POR INFLACION DE ACTIVOS y PATRIMONIO : " & cboMes.Text & " " & txtAnio

'GRABACION DE ASIENTO DE AJUSTE
oMov.BeginTrans
oMov.InsertaMov gsMovNro, gsOpeCod, gsGlosa
gnMovNro = oMov.GetnMovNro(gsMovNro)
oMov.InsertaMovCont gnMovNro, 0, 0, ""
For nItem = 1 To UBound(aAsiento, 2)
   oMov.InsertaMovCta gnMovNro, nItem, aAsiento(0, nItem), aAsiento(1, nItem)
Next
If gdFecha < gdFecSis Then
   oMov.ActualizaSaldoMovimiento gsMovNro, "+"
End If
oMov.CommitTrans
Set oMov = Nothing

Dim oImp As New NContImprimir
sImp = oImp.ImprimeAsientoContable(gsMovNro, gnLinPage, gnColPage, "ASIENTO DE AJUSTE POR INFLACION")
EnviaPrevio sImp, Me.Caption, gnLinPage, False
Set oImp = Nothing

Exit Sub
ErrAsiento:
   MsgBox TextErr(TextErr(Err.Description)), vbInformation, "Aviso"
End Sub

Private Sub cmdProcesar_Click()
Dim nItem     As Integer
Dim lsArchivo As String
Dim lbLibroOpen As Boolean
Dim rsCta As New ADODB.Recordset
Dim sCtaCod   As String
Dim sCtaCodEq As String
Dim sDH       As String
Dim nResp As Integer

On Error GoTo ErrProcesa
If MsgBox(" ¿ Seguro de Generar Asiento de Ajuste por Inflación ? ", vbQuestion + vbYesNo, "Confirmación") = vbNo Then
   Exit Sub
End If

ReDim aAsiento(1, 0)
Me.Enabled = False
sBar.Panels(1).Text = "Procesando..."
prgBar.Visible = True
prgBar.Min = 0
prgBar.value = 0
Dim oAjuste As New DAjusteCont
Set rs = oAjuste.CargaAjusteHistorico()
If Not rs.EOF Then
   lsArchivo = App.path & "\Spooler\AIDet" & "_" & txtAnio & Format(cboMes.ListIndex + 1, "0#") & ".xls"
   lbLibroOpen = ExcelBegin(lsArchivo, xlAplicacion, xlLibro)
   prgBar.Max = rs.RecordCount
   If lbLibroOpen Then
      Dim oBalance As New NBalanceCont
      Do While Not rs.EOF
         sCtaCod = rs!cCtaContCod
         sDH = rs!cTipoDH
         ExcelAddHoja "Cta_" & sCtaCod, xlLibro, xlHoja1
         CabeceraExcel sCtaCod, rs!cCtaContDesc
         Dim nFila As Integer
         nFila = 7
         
         Dim oCont As New NContFunciones
         Do While sCtaCod = rs!cCtaContCod
            prgBar.value = rs.Bookmark
            nFila = nFila + 1
            xlHoja1.Cells(nFila, 2) = Format(rs!dAjusteFecha, gsFormatoFechaView)
            xlHoja1.Cells(nFila, 3) = rs!cAjusteDescrip
            xlHoja1.Cells(nFila, 4) = rs!nAjusteValor3
            xlHoja1.Cells(nFila, 5) = oCont.FactorAjuste(rs!dAjusteFecha, DateAdd("m", 1, CDate("01/" & Format(cboMes.ListIndex + 1, "00") & "/" & txtAnio)) - 1, cboDec)
            xlHoja1.Range(xlHoja1.Cells(nFila, 6), xlHoja1.Cells(nFila, 6)).Formula = "=Round(" + xlHoja1.Range(xlHoja1.Cells(nFila, 4), xlHoja1.Cells(nFila, 4)).Address(False, False) + "*" + xlHoja1.Range(xlHoja1.Cells(nFila, 5), xlHoja1.Cells(nFila, 5)).Address(False, False) & ",2)"
            xlHoja1.Range(xlHoja1.Cells(nFila, 7), xlHoja1.Cells(nFila, 7)).Formula = "=" + xlHoja1.Range(xlHoja1.Cells(nFila, 6), xlHoja1.Cells(nFila, 6)).Address(False, False) + "-" + xlHoja1.Range(xlHoja1.Cells(nFila, 4), xlHoja1.Cells(nFila, 4)).Address(False, False)
            rs.MoveNext
            If rs.EOF Then
               Exit Do
            End If
         Loop
         Set oCont = Nothing
         
         xlHoja1.Range(xlHoja1.Cells(8, 2), xlHoja1.Cells(nFila + 1, 7)).Borders.LineStyle = xlContinuous
         xlHoja1.Range(xlHoja1.Cells(1, 2), xlHoja1.Cells(nFila + 1, 7)).Font.Size = 8
         xlHoja1.Range(xlHoja1.Cells(8, 4), xlHoja1.Cells(nFila + 1, 7)).NumberFormat = "#,##0.00;-#,##0.00"
         xlHoja1.Range(xlHoja1.Cells(8, 5), xlHoja1.Cells(nFila + 1, 5)).NumberFormat = "#,##0." & String(cboDec.Text, "0") & ";-#,##0." & String(cboDec.Text, "0")
         xlHoja1.Cells(nFila + 1, 2) = "TOTAL"
         xlHoja1.Range(xlHoja1.Cells(nFila + 1, 2), xlHoja1.Cells(nFila + 1, 2)).Font.Bold = True
         If nFila > 7 Then
            xlHoja1.Range(xlHoja1.Cells(nFila + 1, 4), xlHoja1.Cells(nFila + 1, 4)).Formula = "=SUM(" + xlHoja1.Range(xlHoja1.Cells(8, 4), xlHoja1.Cells(nFila, 4)).Address(False, False) + ")"
            xlHoja1.Range(xlHoja1.Cells(nFila + 1, 6), xlHoja1.Cells(nFila + 1, 6)).Formula = "=SUM(" + xlHoja1.Range(xlHoja1.Cells(8, 6), xlHoja1.Cells(nFila, 6)).Address(False, False) + ")"
            xlHoja1.Range(xlHoja1.Cells(nFila + 1, 7), xlHoja1.Cells(nFila + 1, 7)).Formula = "=SUM(" + xlHoja1.Range(xlHoja1.Cells(8, 7), xlHoja1.Cells(nFila, 7)).Address(False, False) + ")"
         End If
         
         'Asiento Contable
         gnImporte = nVal(xlHoja1.Cells(nFila + 1, 7))
         gnImporte = gnImporte - oBalance.GetBalanceSaldoCuenta(sCtaCod, txtAnio, Format(cboMes.ListIndex, "00"))
         
         Dim oConA As New NContAsientos
         sCtaCodEq = oConA.GetCtaEquivalente(sCtaCod, sDH, gsOpeCod)
         If sCtaCodEq <> "" Then
            If sDH = "D" Then
               oConA.LlenaArrayAsiento aAsiento, sCtaCod, gnImporte
               oConA.LlenaArrayAsiento aAsiento, sCtaCodEq, gnImporte * -1
            Else
               oConA.LlenaArrayAsiento aAsiento, sCtaCod, gnImporte * -1
               oConA.LlenaArrayAsiento aAsiento, sCtaCodEq, gnImporte
            End If
         Else
            MsgBox "Cuenta Contable " & sCtaCod & " no tiene Equivalente en Operación", vbInformation, "Aviso"
         End If
         Set oConA = Nothing
      Loop
      Set oBalance = Nothing
      ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1
      CargaArchivo lsArchivo, App.path & "\SPOOLER"
   End If
End If
Set oAjuste = Nothing
Me.Enabled = True
RSClose rs

cmdAsiento.Enabled = True
prgBar.Visible = False
sBar.Panels(1).Text = "Proceso Terminado"
Exit Sub
ErrProcesa:
   MsgBox TextErr(Err.Description), vbInformation, "Aviso"
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
CentraForm Me
Dim dFecha As String
Dim lvItem As ListItem
dFecha = DateAdd("m", 1, CDate(LeeConstanteSist(gConstSistCierreMensualCont)))

txtAnio = Year(dFecha)
cboMes.ListIndex = Month(dFecha) - 1
cboDec.ListIndex = 3
End Sub

Private Sub txtAnio_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
If KeyAscii = 13 Then
   If Not ValidaAnio(txtAnio) Then
      Exit Sub
   End If
   cboMes.SetFocus
End If
End Sub

Private Sub txtAnio_Validate(Cancel As Boolean)
   If Not ValidaAnio(txtAnio) Then
      Cancel = True
   Else
   End If
End Sub
