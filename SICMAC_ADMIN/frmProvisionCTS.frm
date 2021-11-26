VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmProvisionCTS 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PROVISION CTS"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10350
   Icon            =   "frmProvisionCTS.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   10350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   315
      Left            =   8685
      TabIndex        =   10
      Text            =   "0.00"
      Top             =   6030
      Width           =   1560
   End
   Begin VB.CommandButton cmdAsiento 
      Caption         =   "&Asiento"
      Height          =   360
      Left            =   5160
      TabIndex        =   9
      Top             =   6450
      Width           =   1230
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   360
      Left            =   9000
      TabIndex        =   3
      Top             =   6450
      Width           =   1230
   End
   Begin VB.CommandButton cmdProvisionar 
      Caption         =   "&Provisionar"
      Height          =   360
      Left            =   7710
      TabIndex        =   4
      Top             =   6450
      Width           =   1230
   End
   Begin VB.CommandButton CmdNuevoPeriodo 
      Caption         =   "Nuevo Periodo"
      Enabled         =   0   'False
      Height          =   360
      Left            =   6435
      TabIndex        =   2
      Top             =   6450
      Width           =   1230
   End
   Begin VB.CommandButton CmdExcel3 
      Caption         =   "<<Exp.Excel>> "
      Height          =   360
      Left            =   120
      TabIndex        =   5
      Top             =   6060
      Width           =   1230
   End
   Begin Sicmact.FlexEdit Flex 
      Height          =   4755
      Left            =   45
      TabIndex        =   1
      Top             =   1170
      Width           =   10200
      _extentx        =   17965
      _extenty        =   9366
      cols0           =   16
      highlight       =   1
      allowuserresizing=   3
      rowsizingmode   =   1
      encabezadosnombres=   "#-Codigo-Nombre--Fec Ing-Rem Ant AFP-3% AFP-Total-Gratific-1/6 Grat-Rem Indem-Mes-Total_Dep-Valida-CodPers-"
      encabezadosanchos=   "500-800-4000-250-1000-1200-1000-1100-1000-1000-1100-900-1000-1200-0-1000"
      font            =   "frmProvisionCTS.frx":08CA
      font            =   "frmProvisionCTS.frx":08F2
      font            =   "frmProvisionCTS.frx":091A
      font            =   "frmProvisionCTS.frx":0942
      font            =   "frmProvisionCTS.frx":096A
      fontfixed       =   "frmProvisionCTS.frx":0992
      backcolorcontrol=   -2147483643
      backcolorcontrol=   -2147483643
      backcolorcontrol=   -2147483643
      lbultimainstancia=   -1  'True
      tipobusqueda    =   7
      columnasaeditar =   "X-X-X-3-X-X-X-X-X-X-X-X-X-X-X-X"
      textstylefixed  =   3
      listacontroles  =   "0-0-0-4-0-0-0-0-0-0-0-0-0-0-0-0"
      encabezadosalineacion=   "C-L-L-R-C-R-R-R-R-R-R-C-R-L-C-C"
      formatosedit    =   "0-0-0-0-0-4-4-4-4-0-0-0-4-0-0-4"
      textarray0      =   "#"
      lbeditarflex    =   -1  'True
      lbbuscaduplicadotext=   -1  'True
      appearance      =   0
      colwidth0       =   495
      rowheight0      =   300
      forecolorfixed  =   -2147483630
   End
   Begin MSComctlLib.ProgressBar PrgBar 
      Height          =   240
      Left            =   90
      TabIndex        =   6
      Top             =   6495
      Visible         =   0   'False
      Width           =   4905
      _ExtentX        =   8652
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSMask.MaskEdBox txtFecha 
      Height          =   300
      Left            =   9135
      TabIndex        =   7
      Top             =   270
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label lnlTot 
      Caption         =   "Total :"
      Height          =   240
      Left            =   8085
      TabIndex        =   11
      Top             =   6075
      Width           =   585
   End
   Begin VB.Label LBLtITULO 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "XXXXXXXXXXXX XXX XXXXXX XX XXXXXXXX XXXXXXXX"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   555
      Left            =   105
      TabIndex        =   8
      Top             =   885
      Visible         =   0   'False
      Width           =   10065
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "PROVISION DE CTS :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   360
      Left            =   6045
      TabIndex        =   0
      Top             =   225
      Width           =   3060
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Height          =   840
      Left            =   45
      Top             =   30
      Width           =   2865
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   105
      Picture         =   "frmProvisionCTS.frx":09B8
      Stretch         =   -1  'True
      Top             =   75
      Width           =   2745
   End
End
Attribute VB_Name = "frmProvisionCTS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ban As Boolean

Private Sub cmdAsiento_Click()
    
    Dim oAsi As NContImprimir
    Set oAsi = New NContImprimir
    Dim lsCadena As String
    
    Dim oPrevio As Previo.clsPrevio
    Set oPrevio = New Previo.clsPrevio
    
    Dim lsMovNroEstad As String
    Dim lsMovNroContra As String
    Dim lnMontoEstad As Currency
    
    If Not IsDate(Me.txtFecha.Text) Then
        MsgBox "Debe ingresar una fecha valida.", vbInformation, "Aviso"
        Exit Sub
    End If
    
    If MsgBox("Desea Generar Asiento Contable ??? ", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
    
    lsMovNroEstad = GeneraProvisionCTS(gsRHPlanillaCTSProvEst, "21160101", CDate(Me.txtFecha.Text), 0, lnMontoEstad)
    lsMovNroContra = GeneraProvisionCTS(gsRHPlanillaCTSProvCon, "21160102", CDate(Me.txtFecha.Text), 1, lnMontoEstad)
 
    lsCadena = oAsi.ImprimeAsientoContable(lsMovNroEstad, 66, 80, , , , , False) & oImpresora.gPrnSaltoPagina & oAsi.ImprimeAsientoContable(lsMovNroContra, 66, 80, , , , , False)
    
    oPrevio.Show lsCadena, Me.Caption, True, , gImpresora
End Sub

Private Sub CmdCarga_Click()
ban = False
End Sub

Private Sub CmdExcel3_Click()
Dim rs As New ADODB.Recordset
Dim fs As Scripting.FileSystemObject
Dim lsArchivo As String

Dim lsNomHoja As String
Dim i As Integer
Dim Y As Integer
Dim sSuma As String
Dim lbExisteHoja As Boolean

Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet
Dim sCabecera As String

sCabecera = "PROVISION DE CTS DEL MES DE " & Format(Me.txtFecha, "MMMM")

If Me.Flex.TextMatrix(1, 1) = "" Then
    Exit Sub
End If

'On Error GoTo ErrorINFO4Excel
Screen.MousePointer = 11
lsArchivo = "ProvCts" & Format(txtFecha, "YYYYMM") & ".xls"

Set fs = New Scripting.FileSystemObject

Set xlAplicacion = New Excel.Application
If fs.FileExists(App.path & "\SPOOLER\" & lsArchivo) Then
    Set xlLibro = xlAplicacion.Workbooks.Open(App.path & "\SPOOLER\" & lsArchivo)
Else
    Set xlLibro = xlAplicacion.Workbooks.Add
End If

lsNomHoja = "CTS" & Format(gdFecSis, "YYYYMM")
For Each xlHoja1 In xlLibro.Worksheets
    If xlHoja1.Name = lsNomHoja Then
        xlHoja1.Activate
        lbExisteHoja = True
        Exit For
    End If
Next
If lbExisteHoja = False Then
    Set xlHoja1 = xlLibro.Worksheets.Add
    xlHoja1.Name = lsNomHoja
End If

Me.PrgBar.Visible = True
xlHoja1.Range("A1") = "CAJA TRUJILLO"
xlHoja1.Range("A1").Font.Bold = True
xlHoja1.Range("F1") = gdFecSis
xlHoja1.Range("F2") = gsCodUser
xlHoja1.Range("F2").HorizontalAlignment = xlRight
xlHoja1.Range("F1:F2").Font.Bold = True

xlHoja1.Range("B5") = "Cod Emp"
xlHoja1.Range("C5") = "Nombre"
xlHoja1.Range("D5") = "Ren_Ant_AFP"
xlHoja1.Range("E5") = "Inc 3%"
xlHoja1.Range("F5") = "Total"
xlHoja1.Range("G5") = "Grati"
xlHoja1.Range("H5") = "1/6 Grati"
xlHoja1.Range("I5") = "Remu Inden"
xlHoja1.Range("J5") = "Mes"
xlHoja1.Range("K5") = "Total Dep"

xlHoja1.Range("B4:K4").MergeCells = True
xlHoja1.Range("B4") = sCabecera
xlHoja1.Range("B4").Font.Bold = True
xlHoja1.Range("B4").HorizontalAlignment = xlCenter
xlHoja1.Range("B5:K5").HorizontalAlignment = xlCenter
xlHoja1.Range("B5:K5").Interior.ColorIndex = 35
xlHoja1.Range("B5:K5").Font.Bold = True
xlHoja1.Range("A1").ColumnWidth = 6
xlHoja1.Range("B1").ColumnWidth = 9
xlHoja1.Range("C1").ColumnWidth = 45
xlHoja1.Range("D1").ColumnWidth = 12
xlHoja1.Range("E1").ColumnWidth = 12
xlHoja1.Range("F1").ColumnWidth = 11
xlHoja1.Range("G1").ColumnWidth = 13
xlHoja1.Range("H1").ColumnWidth = 13
xlHoja1.Range("I1").ColumnWidth = 13
xlHoja1.Range("J1").ColumnWidth = 13
xlHoja1.Range("K1").ColumnWidth = 13
xlHoja1.Application.ActiveWindow.Zoom = 80

xlHoja1.Range("D6:I1000").Style = "Comma"
'xlHoja1.Range("F6:F1000").Style = "Comma"
xlHoja1.Range("K6:K1000").Style = "Comma"
Y = 6
Me.PrgBar.Min = 1
Me.PrgBar.Max = Flex.Rows - 1

For i = 1 To Flex.Rows - 1
   xlHoja1.Range("B" & Y) = Flex.TextMatrix(i, 1)
   xlHoja1.Range("C" & Y) = Flex.TextMatrix(i, 2)
   xlHoja1.Range("D" & Y) = Flex.TextMatrix(i, 5)
   xlHoja1.Range("E" & Y) = Flex.TextMatrix(i, 6)
   xlHoja1.Range("F" & Y) = Flex.TextMatrix(i, 7)
   xlHoja1.Range("G" & Y) = Flex.TextMatrix(i, 8)
   xlHoja1.Range("H" & Y) = Flex.TextMatrix(i, 9)
   xlHoja1.Range("I" & Y) = Flex.TextMatrix(i, 10)
   xlHoja1.Range("J" & Y) = Flex.TextMatrix(i, 11)
   xlHoja1.Range("K" & Y) = Flex.TextMatrix(i, 12)
   Y = Y + 1
   Me.PrgBar.value = i
Next i
    
'xlHoja1.SaveAs App.path & "\SPOOLER\" & "VacPRov"
xlHoja1.SaveAs App.path & "\SPOOLER\" & lsArchivo
'Cierra el libro de trabajo
xlLibro.Close
' Cierra Microsoft Excel con el método Quit.
xlAplicacion.Quit
'Libera los objetos.

Set xlAplicacion = Nothing
Set xlLibro = Nothing
Set xlHoja1 = Nothing
Screen.MousePointer = 0
Me.PrgBar.Visible = False
MsgBox "Se ha Generado el Archivo " & lsArchivo & " Satisfactoriamente", vbInformation, "Aviso"
'CargaArchivo lsArchivo, App.path & "\SPOOLER\"
Exit Sub
'ErrorINFO4Excel:
'    MsgBox "Error Nº [" & Str(Err.Number) & "] " & Err.Description, vbInformation, "Aviso"
'    xlLibro.Close
'    ' Cierra Microsoft Excel con el método Quit.
'    xlAplicacion.Quit
'    'Libera los objetos.
'    Set xlAplicacion = Nothing
'    Set xlLibro = Nothing
'    Set xlHoja1 = Nothing

End Sub

Private Sub CmdNuevoPeriodo_Click()
Dim RHCTS As DRHCTS
Set RHCTS = New DRHCTS

On Error GoTo CTSErr
    If Not RHCTS.VerificaNuevoPeriodo(Format(Me.txtFecha, "YYYYMM")) Then
            Call RHCTS.NuevoPeriodo(Format(Me.txtFecha, "YYYYMM"), FechaHora(gdFecSis), gsCodUser)
        ban = True
        Me.CmdNuevoPeriodo.Visible = False
        MsgBox "Se Inicio el Nuevo Periodo de Provison de CTS", vbInformation, "AVISO"
    Else
        MsgBox "Ya se Inicio el Nuevo periodo. Tenga mas cuidado!", vbInformation, "AVISO"
    End If
    Exit Sub
CTSErr:
    MsgBox Err.Description
End Sub

Private Sub cmdProvisionar_Click()
Dim FechaProv As String
Dim RHCTS As DRHCTS
Dim Mes As Integer
Dim i As Integer
Dim Resultado As Integer
Dim DiaUltimo As Integer
Dim Cantidad As Double
Dim Opt As Integer
If Not ValFecha(Me.txtFecha) Then
    Exit Sub
End If
FechaProv = Format(Me.txtFecha, "YYYYMM")
Mes = Month(Me.txtFecha)
Set RHCTS = New DRHCTS


If Mes = 5 Or Mes = 11 Then
    If RHCTS.VerificaNuevoPeriodo(FechaProv) Then
    Else
        If RHCTS.ProvisionadoCTS(FechaProv) > 0 Then
            ban = True
            Me.CmdNuevoPeriodo.Enabled = False
        End If
    
        If Not ban Then
            MsgBox "USTED INICIAR EL NUEVO PERIODO DE PROVISION ", vbInformation, "AVISO"
            Me.CmdNuevoPeriodo.Enabled = True
            Exit Sub
        End If
    End If
End If


If RHCTS.ProvisionadoCTS(FechaProv) > 0 Then
    MsgBox "Ya se Provisiono el mes de " & Format(Me.txtFecha, "MMMM") & ", Tenga mas cuidado al Provisionar", vbCritical, "AVISO"
    Set RHCTS = Nothing
    Exit Sub
End If

Opt = MsgBox("Esta seguro de Realizar la Provision del mes de " & Format(Me.txtFecha, "MMMM"), vbQuestion + vbYesNo, "AVISO")
If vbNo = Opt Then Exit Sub

Me.PrgBar.Visible = True
Me.PrgBar.Min = 1
Me.PrgBar.Max = Flex.Rows - 1

If RHCTS.VerificaProvCTSMes(Format(Me.txtFecha, "YYYYMM")) Then
    MsgBox "Ya se Provisiono  el mes de " & Format(Me.txtFecha, "YYYYMM") & ". Tenga cuidado", vbInformation, "AVISO"
    Set RHCTS = Nothing
    Exit Sub
End If

If -1 = RHCTS.AbonaCTSPasado(gsCodUser, FechaHora(gdFecSis), Format(Me.txtFecha, "YYYYMM")) Then
    MsgBox "Error de Conexion vuelva a intentarlo", vbInformation, "AVISO"
    Exit Sub
End If

For i = 1 To Flex.Rows - 1
    If Flex.TextMatrix(i, 3) = "." Then
        With Flex
            Resultado = RHCTS.ActualizaMesProvison(.TextMatrix(i, 14), .TextMatrix(i, 1), gsCodUser, FechaHora(gdFecSis), Format(Me.txtFecha, "YYYYMM"), 1, 1)
            .TextMatrix(i, 13) = IIf(Resultado = 1, "OK", "Error")
        End With
    Else
        With Flex
        DiaUltimo = CInt(Left(DateSerial(Year(Me.txtFecha), Month(Me.txtFecha) + 1, 0), 2))
        
        Cantidad = (((CInt(Left(Trim(.TextMatrix(i, 4)), 2)) - DiaUltimo) * -1) + 1) / DiaUltimo
        Resultado = RHCTS.ActualizaMesProvison(.TextMatrix(i, 14), .TextMatrix(i, 1), gsCodUser, FechaHora(gdFecSis), Format(Me.txtFecha, "YYYYMM"), 2, Cantidad, 2)
        Resultado = RHCTS.InsertaCTSTemp(.TextMatrix(i, 14), Cantidad, Format(Me.txtFecha, "YYYYMM"), FechaHora(gdFecSis))
        .TextMatrix(i, 13) = IIf(Resultado = 1, "OK", "Error")
        End With
    End If
    Me.PrgBar.value = i
Next i
'Flex.Visible = False
MsgBox "Provision de CTS realizada", vbInformation, "AVISO"
CargaMatrizCTS
Set RHCTS = Nothing
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Sub CargaMatrizCTS()
Dim rs As ADODB.Recordset
Dim FechaGrati As String
Dim FechaCTS As String
Dim Año As Integer
Dim Mes As Integer
Dim CTS As DRHCTS
Dim i As Integer
Set CTS = New DRHCTS

Dim lnMonto As Currency

lnMonto = 0

PrgBar.Min = 1
'txtFecha = gdFecSis
Mes = Month(Me.txtFecha)

Flex.Rows = 2
For i = 1 To Flex.Cols - 1
    Flex.TextMatrix(1, i) = ""
Next i


Select Case Mes
Case 1, 2, 3, 4, 5, 6:
    Año = Year(Me.txtFecha) - 1
    FechaGrati = Año & "12"
Case 7, 8, 9, 10, 11, 12:
    Año = Year(Me.txtFecha)
    FechaGrati = Año & "07"
End Select

FechaCTS = Format(Me.txtFecha, "YYYYMM")
Set rs = CTS.Carga_CTS_Mes(FechaCTS, FechaGrati, Me.txtFecha)
If Not (rs.EOF And rs.BOF) Then
    PrgBar.Visible = True
    PrgBar.Max = rs.RecordCount
End If

While Not rs.EOF
    Flex.AdicionaFila
    If Right(Format(rs!dIngreso, "DD/MM/YYYY"), 7) <> Right(Me.txtFecha, 7) Then
            Flex.TextMatrix(Flex.Rows - 1, 3) = "1"
    End If
    Flex.TextMatrix(Flex.Rows - 1, 1) = rs!cRHCod
    Flex.TextMatrix(Flex.Rows - 1, 2) = rs!cPersNombre
    Flex.TextMatrix(Flex.Rows - 1, 4) = Format(rs!dIngreso, "DD/MM/YYYY")
    Flex.TextMatrix(Flex.Rows - 1, 5) = Format(IIf(IsNull(rs!REM_ANT_AFP), 0, rs!REM_ANT_AFP), "#0.00")
    Flex.TextMatrix(Flex.Rows - 1, 6) = Format(IIf(IsNull(rs!INCRE_AFP3), 0, rs!INCRE_AFP3), "#0.00")
    Flex.TextMatrix(Flex.Rows - 1, 7) = Format(IIf(IsNull(rs!Total), 0, rs!Total), "#0.00")
    Flex.TextMatrix(Flex.Rows - 1, 8) = Format(IIf(IsNull(rs!GRATI), 0, rs!GRATI), "#0.00")
    Flex.TextMatrix(Flex.Rows - 1, 9) = Format(IIf(IsNull(rs!GRATI6), 0, rs!GRATI6), "#0.00")
    Flex.TextMatrix(Flex.Rows - 1, 10) = Format(IIf(IsNull(rs!REMUNERA_IND), 0, rs!REMUNERA_IND), "#0.00")
    Flex.TextMatrix(Flex.Rows - 1, 11) = rs!MESCTS
    Flex.TextMatrix(Flex.Rows - 1, 12) = Format(IIf(IsNull(rs!TOTALDEP), 0, rs!TOTALDEP), "#0.00")
    Flex.TextMatrix(Flex.Rows - 1, 14) = rs!cPersCod
    lnMonto = lnMonto + Format(IIf(IsNull(rs!TOTALDEP), 0, rs!TOTALDEP), "#0.00")
    PrgBar.value = rs.Bookmark
    rs.MoveNext
Wend
PrgBar.Visible = False
Me.txtTotal.Text = Format(lnMonto, "#,##0.00")

End Sub

Private Sub Form_Load()
Dim FechaP As Date
FechaP = "01" & Mid(gdFecSis, 3, 10)
Me.txtFecha = DateAdd("d", -1, FechaP)
ban = False
End Sub

Private Sub txtFecha_GotFocus()
    
    txtFecha.SelStart = 0
    txtFecha.SelLength = 50
End Sub

Private Sub txtFecha_KeyPress(KeyAscii As Integer)
Dim Mes As Integer
If KeyAscii = 13 Then
    If Not ValFecha(Me.txtFecha) Then
        Me.txtFecha.SetFocus
        Exit Sub
    End If
    Mes = Month(Me.txtFecha)
    Me.LBLtITULO.Visible = True
    Me.LBLtITULO = "COMPENSACION POR TIEMPO DE SERVICIO SEMESTRE "
    Select Case Mes
        Case 11, 12, 1, 2, 3, 4:
        Me.LBLtITULO = Me.LBLtITULO & " NOV " & Year(Me.txtFecha) - 1 & " - " & " ABR " & Year(Me.txtFecha)
        Case 5, 6, 7, 8, 9, 10:
        Me.LBLtITULO = Me.LBLtITULO & " MAY " & Year(Me.txtFecha) & " - " & " OCT " & Year(Me.txtFecha)
    End Select
    
    CargaMatrizCTS
End If
End Sub

Private Function GeneraProvisionCTS(psOpeCod As String, psCtaContCod As String, pdFecha As Date, pnTipoContrato As Integer, pnMontoAcum As Currency)
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim rsMontoAnt As New ADODB.Recordset
    Dim oCon As New DConecta
    Dim oMov As New DMov
    Dim lnMovNro As Long
    Dim lsMovNro As String
    Dim lnItem As Long
    
    Dim lnAcum As Currency
    Dim lnAcumDiff As Currency
    Dim lnUltimo  As Currency
    
    Dim lsCadena As String
    
    sql = "Select cMovNro from mov where cmovnro like '" & Format(pdFecha, gsFormatoMovFecha) & "%' And cOpeCod = '" & psOpeCod & "' And nMovflag = 0"
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sql)
    If Not rs.EOF And Not rs.BOF Then
        lsMovNro = rs!cMovNro
        MsgBox "El Asiento de Provision ya fue generado.", vbInformation
        GeneraProvisionCTS = lsMovNro
        oCon.CierraConexion
        Exit Function
    Else
        rs.Close
    End If
    
    sql = "Select dbo.getsaldocta('" & Format(pdFecha, gsFormatoFecha) & "','" & psCtaContCod & "',1)"
    Set rsMontoAnt = oCon.CargaRecordSet(sql)
    
    sql = " Select isnull(a.cage, b.cage) Age, IsNull(a.nmonto,0) MontoAct , IsNull(b.nmonto,0) MontoAnt, IsNull(a.nmonto,0) - IsNull(b.nmonto,0) Diff from dbo.RHGetProvisionCTSMes('" & Format(pdFecha, gsFormatoFecha) & "',0) a" _
        & " full outer join dbo.RHGetProvisionCTSMesAnt('" & Format(pdFecha, gsFormatoFecha) & "'," & rsMontoAnt.Fields(0) & ",0) b on a.cage = b.cage"
    Set rs = oCon.CargaRecordSet(sql)
    
    lsMovNro = oMov.GeneraMovNro(pdFecha, Right(gsCodAge, 2), gsCodUser)
    lnAcum = 0
    lnItem = 0
    lnAcumDiff = 0
    
    oMov.BeginTrans
        oMov.InsertaMov lsMovNro, psOpeCod, "Provison de Planilla de CTS. " & IIf(pnTipoContrato = 0, "ESTABLES ", "CONTRATADOS ") & Format(pdFecha, gsFormatoFechaView)
        lnMovNro = oMov.GetnMovNro(lsMovNro)
        
        While Not rs.EOF
            lnAcum = lnAcum + Round(rs!Diff, 2) + Round(rs!MontoAnt, 2)
            lnAcumDiff = lnAcumDiff + Round(rs!Diff, 2)
            lnItem = lnItem + 1
            oMov.InsertaMovCta lnMovNro, lnItem, "45110506" & rs!Age, Round(rs!Diff, 2)
            lnUltimo = rs!Diff
            rs.MoveNext
        Wend
        
        If pnTipoContrato = 1 Then
            If lnAcum + pnMontoAcum <> CCur(Me.txtTotal.Text) Then
                oMov.ActualizaMovCta lnItem, lnUltimo + (CCur(Me.txtTotal.Text) - lnAcum - pnMontoAcum)
                lnAcumDiff = lnAcumDiff + (CCur(Me.txtTotal.Text) - lnAcum - pnMontoAcum)
            End If
        Else
            pnMontoAcum = lnAcum
        End If
        
        lnItem = lnItem + 1
        oMov.InsertaMovCta lnMovNro, lnItem, psCtaContCod, lnAcumDiff * -1
        
    oMov.CommitTrans
        
    GeneraProvisionCTS = lsMovNro
    oCon.CierraConexion
End Function

