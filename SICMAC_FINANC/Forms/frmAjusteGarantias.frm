VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmAjusteGarantias 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Colocaciones: Garantías: Reclasificación"
   ClientHeight    =   4890
   ClientLeft      =   2025
   ClientTop       =   2805
   ClientWidth     =   4485
   Icon            =   "frmAjusteGarantias.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   4485
   Begin VB.CommandButton cmdGeneraGarantia 
      Caption         =   "&Reporte de Garantías"
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
      Left            =   120
      TabIndex        =   14
      Top             =   2280
      Width           =   4155
   End
   Begin VB.CommandButton cmdavales 
      Caption         =   "Resumen Garantías &Avales"
      Enabled         =   0   'False
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
      Left            =   120
      TabIndex        =   13
      Top             =   1920
      Visible         =   0   'False
      Width           =   4155
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Resumen por Tipo de Garantías"
      Enabled         =   0   'False
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
      Left            =   135
      TabIndex        =   12
      Top             =   1800
      Visible         =   0   'False
      Width           =   4155
   End
   Begin VB.CommandButton cmdGenerar 
      Caption         =   "&Generar Cuadro de Reclasificación"
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
      Left            =   120
      TabIndex        =   11
      Top             =   2760
      Width           =   4155
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
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
      Left            =   120
      TabIndex        =   10
      Top             =   3720
      Width           =   4155
   End
   Begin VB.CommandButton cmdAsiento 
      Caption         =   "Grabar &Asiento Contable"
      Enabled         =   0   'False
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
      Left            =   120
      TabIndex        =   9
      Top             =   3240
      Width           =   4155
   End
   Begin VB.Frame frmMoneda 
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
      Height          =   750
      Left            =   120
      TabIndex        =   5
      Top             =   930
      Width           =   4170
      Begin VB.TextBox txtTipCambio 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2010
         MaxLength       =   16
         TabIndex        =   7
         Top             =   270
         Width           =   1425
      End
      Begin VB.OptionButton optMoneda 
         Caption         =   "A&justado"
         Height          =   255
         Index           =   3
         Left            =   4500
         TabIndex        =   6
         Top             =   330
         Width           =   1005
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo de Cambio :"
         Height          =   315
         Left            =   615
         TabIndex        =   8
         Top             =   330
         Width           =   1320
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Periodo"
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
      Height          =   795
      Left            =   120
      TabIndex        =   0
      Top             =   75
      Width           =   4170
      Begin VB.TextBox txtAnio 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3120
         MaxLength       =   4
         TabIndex        =   2
         Top             =   300
         Width           =   855
      End
      Begin VB.ComboBox CboMes 
         Height          =   315
         ItemData        =   "frmAjusteGarantias.frx":030A
         Left            =   690
         List            =   "frmAjusteGarantias.frx":0332
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   315
         Width           =   1830
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Año :"
         Height          =   195
         Left            =   2640
         TabIndex        =   4
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Mes :"
         Height          =   195
         Left            =   165
         TabIndex        =   3
         Top             =   390
         Width           =   390
      End
   End
   Begin ComctlLib.ProgressBar pbProceso 
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   4440
      Visible         =   0   'False
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   0
   End
   Begin VB.Label lblPorcentaje 
      Alignment       =   2  'Center
      Caption         =   "0%"
      Height          =   255
      Left            =   2160
      TabIndex        =   17
      Top             =   4200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblDetalleRpt 
      Caption         =   "Dato"
      Height          =   255
      Left            =   1200
      TabIndex        =   16
      Top             =   4200
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "frmAjusteGarantias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim aAsiento() As String
Dim nCta As Integer
Dim dFecha As Date
Dim sCtaDebe  As String
Dim sCtaHaber As String
Dim WithEvents oImp As NContImprimir
Attribute oImp.VB_VarHelpID = -1
Dim oBarra As New clsProgressBar

Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet
Dim lsArchivo As String 'LUCV20161007

'ARLO20170208****
Dim objPista As COMManejador.Pista
'************

Private Sub GeneraReporteAsiento()
Dim nMes As Integer
Dim nAnio As Integer
Dim dFecha As Date
Dim sImpre As String
Dim oCont As New NContFunciones
On Error GoTo AjusteErr
nMes = CboMes.ListIndex + 1
nAnio = txtAnio
dFecha = DateAdd("m", 1, "01/" & Format(nMes, "00") & "/" & Format(txtAnio, "0000")) - 1

If Not oCont.PermiteModificarAsiento(Format(dFecha, gsFormatoMovFecha), False) Then
   MsgBox "Mes ya cerrado. Imposible generar Cuadro de Reclasificación", vbInformation, "!Aviso!"
   Set oCont = Nothing
   Exit Sub
End If
Me.Enabled = False
sImpre = oImp.ImprimeCuadroReclasificacion("G", dFecha, CInt(Mid(gsOpeCod, 3, 1)), "84", nVal(txtTipCambio), gnLinPage, sCtaDebe)
EnviaPrevio sImpre, "CUADRO DE RECLASIFICACION DE GARANTIAS", gnLinPage, False
Set oCont = Nothing
Me.Enabled = True
cmdAsiento.Enabled = True
cmdAsiento.SetFocus

            'ARLO20170208
            Set objPista = New COMManejador.Pista
            'gsOpeCod = LogPistaMantDocumento
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "", "Se ha Generado el Cuadro de Reclasificación al Cierre :" & dFecha & " Con Tipo de Cambio :" & txtTipCambio.Text
            Set objPista = Nothing
            '*******
Exit Sub
AjusteErr:
    MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
End Sub

Private Sub GeneraReporteCuadro(ByVal pnTipo As Integer)
Dim nMes As Integer
Dim nAnio As Integer
Dim dFecha As Date
Dim sImpre As String
Dim oCont As New NContFunciones
On Error GoTo GeneraReporteCuadroErr
nMes = CboMes.ListIndex + 1
nAnio = txtAnio
dFecha = DateAdd("m", 1, "01/" & Format(nMes, "00") & "/" & Format(txtAnio, "0000")) - 1

Me.Enabled = False
ImprimeCuadroGarantias dFecha, CInt(Mid(gsOpeCod, 3, 1)), nVal(txtTipCambio), gnLinPage, pnTipo
Set oCont = Nothing
Me.Enabled = True
Exit Sub
GeneraReporteCuadroErr:
    MsgBox Err.Description, vbInformation, "¡Aviso!"
End Sub

'->***** LUCV20161007, Agregó
Private Sub GeneraRptGarantia()
    Dim nMes As Integer
    Dim nAnio As Integer
    Dim dFecha As Date

    Dim xlsAplicacion As New Excel.Application
    Dim xlsLibro As Excel.Workbook
    Dim xlsHoja As Excel.Worksheet
    Dim ldFecha As Date
    
    'JOEP
        Me.Height = 5310
    'JOEP
    
On Error GoTo GeneraRptGarantiaError
    nMes = CboMes.ListIndex + 1
    nAnio = txtAnio
    dFecha = DateAdd("m", 1, "01/" & Format(nMes, "00") & "/" & Format(txtAnio, "0000")) - 1

    'Generacion
    lsArchivo = "\spooler\RptGarantias" & UCase(gsCodUser) & "_" & Format(gdFecSis, "YYYYMMDD") & "_" & Format$(Time(), "HHMMSS") & ".xlsx"
    ldFecha = dFecha
    
    Set xlsLibro = xlsAplicacion.Workbooks.Add
    'HOJA JOYAS
    lblDetalleRpt.Caption = "JOYAS" 'JOEP
    Set xlsHoja = xlsLibro.Worksheets.Add
    xlsHoja.Name = "JOYAS"
    xlsHoja.Cells.Font.Name = "Arial"
    xlsHoja.Cells.Font.Size = 9
    Call GeneraHojaGarantiaJoyasRpt(ldFecha, xlsHoja)
    
    'HOJA GARANTIAS
    lblDetalleRpt.Caption = "GARANTIAS" 'JOEP
    Set xlsHoja = xlsLibro.Worksheets.Add
    xlsHoja.Name = "GARANTIAS"
    xlsHoja.Cells.Font.Name = "Arial"
    xlsHoja.Cells.Font.Size = 9
    Call GeneraHojaGarantiaRpt(ldFecha, xlsHoja)
    
    'JOEP
        Me.Height = 4560
    'JOEP
    
    MsgBox "Se ha generado satisfactoriamente el reporte de garantias", vbInformation, "Aviso"
    xlsHoja.SaveAs App.path & lsArchivo
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
      
    'ARLO20170208
    Set objPista = New COMManejador.Pista
    'gsOpeCod = LogPistaMantDocumento
    objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "", "Se ha Generado el Reporte de Garantias al Cierre :" & dFecha & " Con Tipo de Cambio :" & txtTipCambio.Text
    Set objPista = Nothing
    '*******
    
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlsHoja = Nothing
    
    Exit Sub
    'fin generacion
GeneraRptGarantiaError:
    MsgBox Err.Description, vbInformation, "¡Aviso!"
End Sub
'<-***** LUCV20161007

Private Sub cboMes_Click()
txtTipCambio = TipoCambioCierre(nVal(txtAnio), CboMes.ListIndex + 1, False)

End Sub

Private Sub cboMes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtTipCambio = TipoCambioCierre(nVal(txtAnio), CboMes.ListIndex + 1, False)
   txtAnio.SetFocus
End If
End Sub

Private Sub cmdAsiento_Click()
Dim rs       As ADODB.Recordset
Dim nTotal   As Currency
Dim nItem    As Integer
Dim lTransActiva As Boolean
Dim nMes     As Integer, nAnio As Integer, dFecha As Date

On Error GoTo AsientoErr

If MsgBox("¿ Seguro desea Grabar Asiento de Garantías ? ", vbQuestion + vbYesNo + vbDefaultButton2, "¡Confirmación¡") = vbNo Then
   Exit Sub
End If

nMes = CboMes.ListIndex + 1
nAnio = txtAnio
dFecha = DateAdd("m", 1, "01/" & Format(nMes, "00") & "/" & Format(txtAnio, "0000")) - 1

Dim oCont As New NContFunciones
Dim oMov  As New DMov
Dim oAju  As New DAjusteCont

If Not oCont.PermiteModificarAsiento(Format(dFecha, gsFormatoMovFecha), False) Then
   MsgBox "Mes ya cerrado. Imposible generar Asiento de Reclasificación", vbInformation, "!Aviso!"
   Exit Sub
End If
gsMovNro = Format(dFecha, "yyyymmdd")
If oCont.ExisteMovimiento(gsMovNro, gsOpeCod) Then
   MsgBox "Asiento de Garantías ya generado", vbInformation, "¡Aviso!"
   Exit Sub
End If

Me.Enabled = False
Set rs = oAju.AjusteReclasificaGarantia(Format(dFecha, gsFormatoFecha), CInt(Mid(gsOpeCod, 3, 1)), "84", nVal(txtTipCambio))
If rs.EOF Then
   MsgBox "No existen diferencias entre Estadísticas y Saldos Contables ", vbInformation, "!Aviso!"
Else
   oImp_BarraShow rs.RecordCount
   oMov.BeginTrans
   gsGlosa = "Asiento de Reclasificación de Garantías al " & dFecha & " en " & IIf(Mid(gsOpeCod, 3, 1) = "1", "M.N.", "M.E.")
   gsMovNro = oMov.GeneraMovNro(dFecha, gsCodAge, gsCodUser)
   lTransActiva = True
   oMov.InsertaMov gsMovNro, gsOpeCod, gsGlosa, gMovEstContabMovContable, gMovFlagVigente
   gnMovNro = oMov.GetnMovNro(gsMovNro)
   nItem = 1

   Dim lsCtaCod As String
   Do While Not rs.EOF
      lsCtaCod = IIf(IsNull(rs!Cta1), rs!Cta2, rs!Cta1)
      gnImporte = Val(Format(rs!nSaldo, "#0.00")) - rs!nCtaSaldoImporte
      If gnImporte <> 0 Then
        
        'ALPA 20090316**********************************************************
        If gsOpeCod = gReclasiGarantME Then
            'ALPA 20090710****************************************************************************
            'oMov.InsertaMovCta gnMovNro, nItem, lsCtaCod, (gnImporte * IIf(nVal(txtTipCambio) = 0, 1, nVal(txtTipCambio))) * -1
            oMov.InsertaMovCta gnMovNro, nItem, lsCtaCod, Round(gnImporte * IIf(nVal(txtTipCambio) = 0, 1, nVal(txtTipCambio)), 2) * -1
            'oMov.InsertaMovMe gnMovNro, nItem, gnImporte * -1
            oMov.InsertaMovMe gnMovNro, nItem, Round(gnImporte, 2) * -1
            '*****************************************************************************************
        Else
            oMov.InsertaMovCta gnMovNro, nItem, lsCtaCod, gnImporte * -1
        End If
        '***********************************************************************
        
        nTotal = nTotal + gnImporte
        nItem = nItem + 1
      End If
      oImp_BarraProgress rs.Bookmark, "ASIENTO DE GARANTIAS", "", "Grabando...", vbBlue
      rs.MoveNext
   Loop
   'ALPA 20090316**********************************************************
    If gsOpeCod = gReclasiGarantME Then
        'ALPA 20090710************************************************************************************************
        'oMov.InsertaMovCta gnMovNro, nItem, "83" & Mid(gsOpeCod, 3, 1) & "1", (nTotal * IIf(nVal(txtTipCambio) = 0, 1, nVal(txtTipCambio)))
        oMov.InsertaMovCta gnMovNro, nItem, "83" & Mid(gsOpeCod, 3, 1) & "1", Round(nTotal * IIf(nVal(txtTipCambio) = 0, 1, nVal(txtTipCambio)), 2)
        'oMov.InsertaMovMe gnMovNro, nItem, nTotal
        oMov.InsertaMovMe gnMovNro, nItem, Round(nTotal, 2)
        '**************************************************************************************************************
    Else
        oMov.InsertaMovCta gnMovNro, nItem, "83" & Mid(gsOpeCod, 3, 1) & "1", nTotal
    End If
    '***********************************************************************
   
   If dFecha < gdFecSis Then
      oMov.ActualizaSaldoMovimiento gsMovNro, "+"
   End If
   oMov.CommitTrans
   lTransActiva = False
   oImp_BarraClose
End If
RSClose rs
            'ARLO20170208
            Set objPista = New COMManejador.Pista
            'gsOpeCod = LogPistaMantDocumento
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", "Se Grabo Asiento Contable de Garantias al Cierre :" & dFecha & " Con Tipo de Cambio :" & txtTipCambio.Text
            Set objPista = Nothing
            '*******
ImprimeAsientoContable gsMovNro, , , , , , , , , , , , 1, "ASIENTO DE GARANTIAS"
Set oMov = Nothing
Set oCont = Nothing
Set oAju = Nothing
Me.Enabled = True
Exit Sub
AsientoErr:
   MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
   Me.Enabled = True
   If lTransActiva Then
        oMov.RollbackTrans
   End If
End Sub


Private Sub cmdGenerar_Click()
If Not ValidaDatos Then
   Exit Sub
End If
If MsgBox("¿ Seguro que desea generar Cuadro de Reclasificación de Garantias ? ", vbQuestion + vbYesNo, "!Confirmación!") = vbNo Then
   Exit Sub
End If
GeneraReporteAsiento
End Sub

Private Sub cmdImprimir_Click()
If Not ValidaDatos Then
   Exit Sub
End If
If MsgBox("¿ Seguro que desea generar Resumen de Garantías ? ", vbQuestion + vbYesNo, "!Confirmación!") = vbNo Then
   Exit Sub
End If
GeneraReporteCuadro (1)
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub
Private Sub cmdAvales_Click()
If Not ValidaDatos Then
   Exit Sub
End If
If MsgBox("¿ Seguro que desea generar Resumen de Avales ? ", vbQuestion + vbYesNo, "!Confirmación!") = vbNo Then
   Exit Sub
End If
GeneraReporteCuadro (2)
End Sub

'->*****LUCV20161007
Private Sub cmdGeneraGarantia_Click()
If MsgBox("¿ Desea generar el reporte de garantías? ", vbQuestion + vbYesNo, "!Confirmación!") = vbNo Then
   Exit Sub
End If
Call GeneraRptGarantia
End Sub
'<-*****LUCV20161007

Private Sub Form_Load()
CentraForm Me
'JOEP
Me.Height = 4560
'JOEP
frmOperaciones.Enabled = False
CboMes.ListIndex = Month(gdFecSis) - 1
txtAnio = Year(gdFecSis)
Set oImp = New NContImprimir
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set oImp = Nothing
frmOperaciones.Enabled = True
End Sub

Private Sub txtAnio_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtTipCambio = TipoCambioCierre(nVal(txtAnio), CboMes.ListIndex + 1, False)
   txtTipCambio.SetFocus
End If
End Sub

Private Sub txtTipCambio_GotFocus()
fEnfoque txtTipCambio
End Sub

Private Sub txtTipCambio_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtTipCambio, KeyAscii, 14, 5)
If KeyAscii = 13 Then
   'cmdImprimir.SetFocus
   cmdGeneraGarantia.SetFocus
End If
End Sub

Private Function ValidaDatos() As Boolean
ValidaDatos = False
If CboMes.ListIndex = -1 Then
   MsgBox "Debe seleccionarse mes de proceso", vbInformation, "!Aviso!"
   CboMes.SetFocus
   Exit Function
End If
If Val(txtAnio) = 0 Then
   MsgBox "Debe indicar año de proceso", vbInformation, "!Aviso!"
   txtAnio.SetFocus
   Exit Function
End If
If Not ValidaAnio(txtAnio) Then
   txtAnio.SetFocus
   Exit Function
End If
If Val(txtTipCambio) = 0 Then
   MsgBox "Debe indicar Tipo de Cambio", vbInformation, "!Aviso!"
   txtTipCambio.SetFocus
   Exit Function
End If
ValidaDatos = True
End Function

Private Sub oImp_BarraClose()
oBarra.CloseForm Me
End Sub

Private Sub oImp_BarraProgress(value As Variant, psTitulo As String, psSubTitulo As String, psTituloBarra As String, ColorLetras As ColorConstants)
oBarra.Progress value, psTitulo, psSubTitulo, psTituloBarra, ColorLetras
End Sub

Private Sub oImp_BarraShow(pnMax As Variant)
oBarra.ShowForm Me
oBarra.CaptionSyle = eCap_CaptionPercent
oBarra.Max = pnMax
End Sub
Public Function ImprimeCuadroGarantias(pdFecha As Date, pnMoneda As Integer, pnTipCambio As Currency, pnLinPage As Integer, Optional ByVal pnTipo As Integer) As String ' By Capi 10122007 se agrego un parametro opcional
Dim rs As ADODB.Recordset
Dim sSql As String
Dim nLin As Integer
Dim P    As Integer
Dim sP As String
Dim sTit As String, sTitBarra As String
Dim nTotDif As Currency
Dim oAju  As New DAjusteCont

Dim lsArchivo   As String
Dim lbLibroOpen As Boolean
Dim lsTitulo    As String
Dim N           As Integer
On Error GoTo ErrImprime

sTitBarra = "RESUMEN DE GARANTIAS"
oImp_BarraShow 1
oImp_BarraProgress 0, sTitBarra, "", "Obteniendo datos", vbBlue
Set rs = oAju.CargaEstadGarantias(Format(pdFecha, gsFormatoFecha), pnMoneda, pnTipCambio, pnTipo) 'By capi 10122007 se agrego par pnTipo
oImp_BarraClose
If rs.EOF Then
   Err.Raise "50001", "ImprimeCuadroGarantias", "No existen Datos para generar Resumen de Garantías"
Else
    'By Capi 10122007 titulo de acuerdo a parametro
    If pnTipo = 1 Then
        lsTitulo = "CREDITOS POR TIPO DE GARANTIA AL " & pdFecha
        lsArchivo = App.path & "\Spooler\RGARANT_" & Year(pdFecha) & "_" & IIf(pnMoneda = 1, "MN", "ME") & ".xls"
    Else
        lsTitulo = "CREDITOS POR TIPO DE GARANTIA-SOLO AVALES AL " & pdFecha
        lsArchivo = App.path & "\Spooler\RAVALES_" & Year(pdFecha) & "_" & IIf(pnMoneda = 1, "MN", "ME") & ".xls"
    End If
    
    oImp_BarraShow 1
    oImp_BarraProgress 0, lsTitulo, "Generando Hoja Excel...", "", vbBlue
    
    
    lbLibroOpen = ExcelBegin(lsArchivo, xlAplicacion, xlLibro)
    If lbLibroOpen Then
        ExcelAddHoja Format(Month(pdFecha), "00"), xlLibro, xlHoja1, False
        oImp_BarraProgress 1, lsTitulo, "Hoja Excel generada...", "", vbBlue
        oImp_BarraShow rs.RecordCount
        '------CABECERA
        xlHoja1.Cells(1, 1) = gsNomCmac
        xlHoja1.Cells(2, 3) = lsTitulo
        xlHoja1.Range("A1:D2").Font.Bold = True
        P = rs.Fields.Count - 1
        sP = ExcelColumnaString(P + 1)
        For N = 0 To P
            xlHoja1.Cells(4, N + 1) = Replace(rs.Fields(N).Name, "_", " ")
        Next
        xlHoja1.Range("A4:" & sP & "4").BorderAround xlContinuous
        xlHoja1.Range("A4:" & sP & "4").Borders(xlInsideVertical).LineStyle = xlContinuous
        xlHoja1.Range("A4:" & sP & "4").Font.Bold = True
        xlHoja1.Range("A4:" & sP & "4").WrapText = True
        xlHoja1.Range("A4:" & sP & "4").HorizontalAlignment = xlCenter
        xlHoja1.Range("A4:" & sP & "4").VerticalAlignment = xlCenter
        xlHoja1.Range("D4:" & sP & "4").ColumnWidth = 13
        
        '------FIN CABECERA
        xlHoja1.Range("A5").CopyFromRecordset rs
'        Do While Not rs.EOF
'            oImp_BarraProgress rs.Bookmark, sTitBarra, "", "Generando Cuadro", vbBlue
'            For N = 0 To P
'                xlHoja1.Cells(rs.Bookmark + 4, N + 1) = rs.Fields(N).value
'                If N >= 3 Then
'                    xlHoja1.Range(ExcelColumnaString(N + 1) & rs.Bookmark + 4 & ":" & ExcelColumnaString(N + 1) & rs.Bookmark + 4).NumberFormat = "#,##0.00"
'                End If
'            Next
'            rs.MoveNext
'        Loop
        nLin = rs.RecordCount + 5
        xlHoja1.Range("A" & nLin & ":" & sP & nLin).BorderAround xlContinuous
        xlHoja1.Range("A" & nLin & ":" & sP & nLin).Font.Bold = True
        xlHoja1.Range("D" & nLin & ":" & sP & nLin).NumberFormat = "#,##0.00"
        xlHoja1.Range("A" & 5 & ":" & sP & nLin).Font.Size = 9
        xlHoja1.Range("A" & 5 & ":" & "C" & nLin).HorizontalAlignment = xlCenter
        For N = 3 To P
            xlHoja1.Range(ExcelColumnaString(N + 1) & nLin & ":" & ExcelColumnaString(N + 1) & nLin).Formula = "=SUM(" & ExcelColumnaString(N + 1) & "4:" & ExcelColumnaString(N + 1) & nLin - 1 & ") "
        Next
        ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1
        oImp_BarraClose
        CargaArchivo lsArchivo, App.path & "\Spooler"
        MsgBox "Archivo generado satisfactoriamente", vbInformation, "Aviso!!!"
    Else
        oImp_BarraClose
        MsgBox "Archivo no puede generarse por falta de datos", vbInformation, "Aviso!!!"
    End If
End If
Exit Function
ErrImprime:
   MsgBox TextErr(Err.Description), vbInformation, "!Aviso!"
   If lbLibroOpen Then
      ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1, False
      lbLibroOpen = False
   End If
   If Not oBarra Is Nothing Then
      oImp_BarraClose
   End If
   Set oBarra = Nothing
   MousePointer = 0
End Function

'->***** LUCV20161010
Public Sub GeneraHojaGarantiaRpt(ByVal pdFecha As Date, ByRef xlsHoja As Worksheet)
    Dim oGarantias As New DGarantia
    Dim rsGarantias As New ADODB.Recordset
    Dim lsCodPersAnt As String
    
    Dim lsNomHoja As String, lsArchivo As String, lsFormulaTotal As String
    Dim ldFechaRep, ldFechaTC As Date
    Dim lnAcumDPOvernight As Currency, lnAcumInversiones As Currency
    'Comento JOEP20200313 Motivo Desbordamiento
    'Dim lnPosActual As Integer, lnPosAnterior As Integer
    'Dim i As Integer
    'Dim iMat As Integer
    'Comento JOEP20200313 Motivo Desbordamiento
    'Add JOEP20200313
    Dim i As Long
    Dim lnPosActual As Long, lnPosAnterior As Long
    Dim iMat As Long
    'Add JOEP20200313

    ldFechaRep = pdFecha
    lnPosActual = 3
    
    xlsHoja.Cells.Font.Name = "Arial"
    xlsHoja.Cells.Font.Size = 9
    xlsHoja.Columns("A:A").ColumnWidth = 2
    xlsHoja.Cells(3, 3) = "Reporte de Garantias"
    xlsHoja.Cells(3, 3).Font.Size = 13
    xlsHoja.Range("B3", "Q3").MergeCells = True
    xlsHoja.Cells(3, 3).HorizontalAlignment = 3
    
    'Configuración del Ancho de Columnas
    xlsHoja.Columns("B:B").ColumnWidth = 10
    xlsHoja.Columns("C:C").ColumnWidth = 7
    xlsHoja.Columns("D:D").ColumnWidth = 30
    xlsHoja.Columns("E:E").ColumnWidth = 25
    xlsHoja.Columns("F:F").ColumnWidth = 25
    xlsHoja.Columns("G:G").ColumnWidth = 25
    xlsHoja.Columns("H:H").ColumnWidth = 13
    xlsHoja.Columns("I:I").ColumnWidth = 13
    xlsHoja.Columns("N:N").ColumnWidth = 18
    xlsHoja.Columns("O:O").ColumnWidth = 10
    xlsHoja.Columns("P:P").ColumnWidth = 10
    
    
    xlsHoja.Cells(2, 15) = "FECHA DE REPORTE"
    xlsHoja.Cells(2, 15).HorizontalAlignment = 3
    xlsHoja.Cells(2, 15).Font.Size = 10
    
    xlsHoja.Cells(2, 17).NumberFormat = "dd/mm/yyyy" 'EJVG20130927
    xlsHoja.Cells(2, 17) = pdFecha 'Format(pdFecha, gsFormatoFechaView)
    xlsHoja.Cells(2, 17).HorizontalAlignment = 3
    xlsHoja.Cells(2, 17).Font.Size = 10
    xlsHoja.Cells(2, 17).Interior.Color = RGB(255, 204, 153)
    
    xlsHoja.Range(xlsHoja.Cells(5, 7), xlsHoja.Cells(5, 9)).Borders.LineStyle = xlContinuous
    xlsHoja.Range(xlsHoja.Cells(5, 7), xlsHoja.Cells(5, 9)).Borders(xlEdgeTop).Weight = xlMedium 'Ancho del borde de Linea
    xlsHoja.Range(xlsHoja.Cells(5, 7), xlsHoja.Cells(5, 9)).Borders(xlEdgeBottom).Weight = xlMedium 'Ancho del borde de Linea
    xlsHoja.Range(xlsHoja.Cells(5, 7), xlsHoja.Cells(5, 9)).Borders(xlEdgeLeft).Weight = xlMedium 'Ancho del borde de Linea
    xlsHoja.Range(xlsHoja.Cells(5, 7), xlsHoja.Cells(5, 9)).Borders(xlEdgeRight).Weight = xlMedium 'Ancho del borde de Linea
    xlsHoja.Range(xlsHoja.Cells(5, 7), xlsHoja.Cells(5, 9)).Borders.ColorIndex = xlAutomatic 'Color del Borde de Linea
    
    xlsHoja.Range("B9", "Q" & lnPosActual).Borders.LineStyle = xlContinuous
    xlsHoja.Range("B9", "Q" & lnPosActual).Borders.Weight = xlThin
    xlsHoja.Range("B9", "Q" & lnPosActual).Borders.ColorIndex = xlAutomatic
    
    xlsHoja.Range("B9", "B" & lnPosActual).Borders(xlEdgeLeft).Weight = xlMedium
    xlsHoja.Range("B" & lnPosActual, "Q" & lnPosActual).Borders.Weight = xlMedium
    xlsHoja.Range("I9", "Q" & lnPosActual).Borders(xlEdgeRight).Weight = xlMedium

    'REPORTE DE GARANTIAS
    lnPosActual = lnPosActual + 1
    lnPosAnterior = lnPosActual
    xlsHoja.Cells(lnPosActual, 2) = "Garantia"
    xlsHoja.Cells(lnPosActual, 3) = "Moneda"
    xlsHoja.Cells(lnPosActual, 4) = "Titular"
    xlsHoja.Cells(lnPosActual, 5) = "Calificacion SBS"
    xlsHoja.Cells(lnPosActual, 6) = "Clasificacion"
    xlsHoja.Cells(lnPosActual, 7) = "Bien Garantía"
    xlsHoja.Cells(lnPosActual, 8) = "VRM"
    xlsHoja.Cells(lnPosActual, 9) = "Gravamen"
    xlsHoja.Cells(lnPosActual, 10) = "GravamenValor"
    xlsHoja.Cells(lnPosActual, 11) = "GarantiaDisponible"
    xlsHoja.Cells(lnPosActual, 12) = "FechaTasacion"
    xlsHoja.Cells(lnPosActual, 13) = "FechaTra"
    xlsHoja.Cells(lnPosActual, 14) = "TipoTramite"
    xlsHoja.Cells(lnPosActual, 15) = "leasing"
    xlsHoja.Cells(lnPosActual, 16) = "Aval"
    xlsHoja.Cells(lnPosActual, 17) = "CuentaContable"
    
    xlsHoja.Range(xlsHoja.Cells(lnPosActual, 2), xlsHoja.Cells(lnPosActual, 17)).Interior.Color = RGB(255, 204, 153)
    xlsHoja.Range(xlsHoja.Cells(lnPosActual, 2), xlsHoja.Cells(lnPosActual, 17)).HorizontalAlignment = xlCenter
    xlsHoja.Range(xlsHoja.Cells(lnPosActual, 2), xlsHoja.Cells(lnPosActual, 17)).Font.Bold = True
    
    Set rsGarantias = oGarantias.ObtenerGarantiasJoyasRpt(ldFechaRep, 1)
    
    'JOEP20200313
    lblDetalleRpt.Visible = True
    lblPorcentaje.Visible = True
    pbProceso.Visible = True
    pbProceso.value = 0
    pbProceso.Min = 0
    pbProceso.Max = IIf(rsGarantias.RecordCount = 0, 1, (rsGarantias.RecordCount - 1))
    'JOEP20200313
        
    For i = 0 To rsGarantias.RecordCount - 1
        lnPosActual = lnPosActual + 1
            xlsHoja.Cells(lnPosActual, 2) = rsGarantias!Garantia_Cod
            xlsHoja.Cells(lnPosActual, 3) = rsGarantias!Garantia_Moneda
            xlsHoja.Cells(lnPosActual, 4) = rsGarantias!Garantia_Titular_Nombre
            xlsHoja.Cells(lnPosActual, 5) = rsGarantias!Garantia_Clasificacion_SBS
            xlsHoja.Cells(lnPosActual, 6) = rsGarantias!Garantia_Clasificacion
            xlsHoja.Cells(lnPosActual, 7) = rsGarantias!Garantia_Bien_Garantia
            xlsHoja.Cells(lnPosActual, 8).NumberFormat = "#,##0.00"
            xlsHoja.Cells(lnPosActual, 8) = rsGarantias!Garantia_VRM
            xlsHoja.Cells(lnPosActual, 9).NumberFormat = "#,##0.00"
            xlsHoja.Cells(lnPosActual, 9) = rsGarantias!Garantia_Gravamen
            xlsHoja.Cells(lnPosActual, 10).NumberFormat = "#,##0.00"
            xlsHoja.Cells(lnPosActual, 10) = rsGarantias!Garantia_Valor
            xlsHoja.Cells(lnPosActual, 11).NumberFormat = "#,##0.00"
            xlsHoja.Cells(lnPosActual, 11) = rsGarantias!Garantia_Disponible
            'xlsHoja.Range(xlsHoja.Cells(lnPosActual, 12), xlsHoja.Cells(lnPosActual, 13)).NumberFormat = "dd/mm/yyyy"
            xlsHoja.Cells(lnPosActual, 12) = rsGarantias!Garantia_VAL_Tasacion_Fecha
            xlsHoja.Cells(lnPosActual, 13) = rsGarantias!Garantia_TRA_Fecha
            xlsHoja.Cells(lnPosActual, 14) = rsGarantias!Garantia_TRA_Tipo
            xlsHoja.Cells(lnPosActual, 15) = rsGarantias!Garantia_Leasing
            xlsHoja.Cells(lnPosActual, 16) = rsGarantias!Garantia_AVAL
            xlsHoja.Cells(lnPosActual, 17) = rsGarantias!Garantia_Cuenta_Contable
        rsGarantias.MoveNext
        pbProceso.value = i 'JOEP20200313
        lblPorcentaje = CLng((pbProceso.value * 100) / pbProceso.Max) & " %"
    Next
       
    xlsHoja.Range("B" & lnPosAnterior, "Q" & lnPosActual).Borders.Weight = xlThin
    'JOEP20200313
    pbProceso.Visible = False
    lblDetalleRpt.Visible = False
    lblPorcentaje.Visible = False
    'JOEP20200313
    '******************************
    
    Set oGarantias = Nothing
    Set rsGarantias = Nothing
End Sub

Public Sub GeneraHojaGarantiaJoyasRpt(ByVal pdFecha As Date, ByRef xlsHoja As Worksheet)
    Dim oGarantias As New DGarantia
    Dim rsGarantias As New ADODB.Recordset
    Dim lsCodPersAnt As String
    
    Dim lsNomHoja As String, lsArchivo As String, lsFormulaTotal As String
    Dim ldFechaRep, ldFechaTC As Date
    Dim lnAcumDPOvernight As Currency, lnAcumInversiones As Currency
'JOEP20200313 Desbordamiento
    'Dim lnPosActual As Integer, lnPosAnterior As Integer
    'Dim i As Integer
    'Dim iMat As Integer
'JOEP20200313 Desbordamiento
    'JOEP20200313
    Dim lnPosActual As Long, lnPosAnterior As Long
    Dim i As Long
    Dim iMat As Long
    'JOEP20200313

    ldFechaRep = pdFecha
    lnPosActual = 3
    
    xlsHoja.Cells.Font.Name = "Arial"
    xlsHoja.Cells.Font.Size = 9
    xlsHoja.Columns("A:A").ColumnWidth = 2
    xlsHoja.Cells(3, 3) = "Reporte de Garantias - Joyas"
    xlsHoja.Cells(3, 3).Font.Size = 13
    xlsHoja.Range("B3", "F3").MergeCells = True
    xlsHoja.Cells(3, 3).HorizontalAlignment = 3
    
    'Configuración del Ancho de Columnas
    xlsHoja.Columns("B:B").ColumnWidth = 18
    xlsHoja.Columns("C:C").ColumnWidth = 32
    xlsHoja.Columns("D:D").ColumnWidth = 6
    xlsHoja.Columns("E:E").ColumnWidth = 25
    xlsHoja.Columns("F:F").ColumnWidth = 15
    
    xlsHoja.Cells(2, 5) = "Fecha de Reporte"
    xlsHoja.Cells(2, 5).HorizontalAlignment = 4 'Credito/Cliente
    xlsHoja.Cells(2, 5).Font.Size = 10
    
    xlsHoja.Cells(2, 6).NumberFormat = "dd/mm/yyyy" 'EJVG20130927
    xlsHoja.Cells(2, 6) = pdFecha 'Format(pdFecha, gsFormatoFechaView)
    xlsHoja.Cells(2, 6).HorizontalAlignment = 3
    xlsHoja.Cells(2, 6).Font.Size = 10
    xlsHoja.Cells(2, 6).Interior.Color = RGB(221, 248, 255)
    
    xlsHoja.Range(xlsHoja.Cells(5, 6), xlsHoja.Cells(5, 6)).Borders.LineStyle = xlContinuous
    xlsHoja.Range(xlsHoja.Cells(5, 6), xlsHoja.Cells(5, 6)).Borders(xlEdgeTop).Weight = xlMedium 'Ancho del borde de Linea
    xlsHoja.Range(xlsHoja.Cells(5, 6), xlsHoja.Cells(5, 6)).Borders(xlEdgeBottom).Weight = xlMedium 'Ancho del borde de Linea
    xlsHoja.Range(xlsHoja.Cells(5, 6), xlsHoja.Cells(5, 6)).Borders(xlEdgeLeft).Weight = xlMedium 'Ancho del borde de Linea
    xlsHoja.Range(xlsHoja.Cells(5, 6), xlsHoja.Cells(5, 6)).Borders(xlEdgeRight).Weight = xlMedium 'Ancho del borde de Linea
    xlsHoja.Range(xlsHoja.Cells(5, 6), xlsHoja.Cells(5, 6)).Borders.ColorIndex = xlAutomatic 'Color del Borde de Linea
    
    xlsHoja.Range("B9", "F" & lnPosActual).Borders.LineStyle = xlContinuous
    xlsHoja.Range("B9", "F" & lnPosActual).Borders.Weight = xlThin
    xlsHoja.Range("B9", "F" & lnPosActual).Borders.ColorIndex = xlAutomatic
    
    'xlsHoja.Range("B9", "B" & lnPosActual).Borders(xlEdgeLeft).Weight = xlMedium
    xlsHoja.Range("B" & lnPosActual, "F" & lnPosActual).Borders.Weight = xlMedium
    'xlsHoja.Range("I9", "F" & lnPosActual).Borders(xlEdgeRight).Weight = xlMedium

    'REPORTE DE GARANTIAS
    lnPosActual = lnPosActual + 1
    lnPosAnterior = lnPosActual
    xlsHoja.Cells(lnPosActual, 2) = "Credito"
    xlsHoja.Cells(lnPosActual, 3) = "Cliente"
    xlsHoja.Cells(lnPosActual, 4) = "Moneda"
    xlsHoja.Cells(lnPosActual, 5) = "Agencia"
    xlsHoja.Cells(lnPosActual, 6) = "Monto"
    
    xlsHoja.Range(xlsHoja.Cells(lnPosActual, 2), xlsHoja.Cells(lnPosActual, 6)).Interior.Color = RGB(221, 248, 255)
    xlsHoja.Range(xlsHoja.Cells(lnPosActual, 2), xlsHoja.Cells(lnPosActual, 6)).HorizontalAlignment = xlCenter
    xlsHoja.Range(xlsHoja.Cells(lnPosActual, 2), xlsHoja.Cells(lnPosActual, 6)).Font.Bold = True
    
    Set rsGarantias = oGarantias.ObtenerGarantiasJoyasRpt(ldFechaRep, 2)
    'JOEP20200313
    lblDetalleRpt.Visible = True
    lblPorcentaje.Visible = True
    pbProceso.Visible = True
    pbProceso.Min = 0
    pbProceso.Max = IIf(rsGarantias.RecordCount = 0, 1, (rsGarantias.RecordCount - 1))
    'JOEP20200313
    
    For i = 0 To rsGarantias.RecordCount - 1
        lnPosActual = lnPosActual + 1
        xlsHoja.Cells(lnPosActual, 2) = rsGarantias!cCtaCod
        xlsHoja.Cells(lnPosActual, 3) = rsGarantias!cPersNombre
        xlsHoja.Cells(lnPosActual, 4) = rsGarantias!cMoneda
        xlsHoja.Cells(lnPosActual, 5) = rsGarantias!cAgencia
        xlsHoja.Cells(lnPosActual, 6).NumberFormat = "#,##0.00"
        xlsHoja.Cells(lnPosActual, 6) = rsGarantias!nMonto
        rsGarantias.MoveNext
        pbProceso.value = i 'JOEP20200313
        lblPorcentaje = CLng((pbProceso.value * 100) / pbProceso.Max) & " %"
    Next


    xlsHoja.Range("B" & lnPosAnterior, "F" & lnPosActual).Borders.Weight = xlThin
    'JOEP20200313
    pbProceso.Visible = False
    lblDetalleRpt.Visible = False
    lblPorcentaje.Visible = False
    'JOEP20200313
    '******************************
    
    Set oGarantias = Nothing
    Set rsGarantias = Nothing
End Sub
'<-***** LUCV20161010
