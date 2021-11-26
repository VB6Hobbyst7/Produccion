VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
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
   Begin VB.CommandButton cmdConsol 
      Caption         =   "&Consol Excel"
      Height          =   345
      Left            =   1680
      TabIndex        =   14
      Top             =   6060
      Width           =   1260
   End
   Begin VB.CheckBox chkPeriodo 
      Caption         =   "Ultimo Mes del Periodo (Asiento Contable)"
      Height          =   375
      Left            =   6120
      TabIndex        =   12
      Top             =   480
      Width           =   4095
   End
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
      _ExtentX        =   17965
      _ExtentY        =   9366
      Cols0           =   16
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "#-Codigo-Nombre--F. Ingreso-Sueldo-3% AFP-Sueldo-Grati-1/6 Grati-Total-Mes-Total_Dep-Valida-CodPers-Agencia"
      EncabezadosAnchos=   "500-800-3500-250-1000-0-0-900-900-900-900-500-900-800-0-800"
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-X-X-3-X-X-X-X-X-X-X-X-X-X-X-X"
      TextStyleFixed  =   3
      ListaControles  =   "0-0-0-4-0-0-0-0-0-0-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-L-L-R-C-R-R-R-R-R-R-C-R-L-C-C"
      FormatosEdit    =   "0-0-0-0-0-4-4-4-4-0-0-0-4-0-0-4"
      TextArray0      =   "#"
      lbEditarFlex    =   -1  'True
      lbUltimaInstancia=   -1  'True
      TipoBusqueda    =   7
      lbBuscaDuplicadoText=   -1  'True
      Appearance      =   0
      ColWidth0       =   495
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
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
      Left            =   9120
      TabIndex        =   7
      Top             =   150
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "CAJA MAYNAS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   555
      Left            =   240
      TabIndex        =   13
      Top             =   240
      Width           =   3615
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
         Weight          =   900
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
      Left            =   6120
      TabIndex        =   0
      Top             =   105
      Width           =   3060
   End
End
Attribute VB_Name = "frmProvisionCTS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ban As Boolean

Private Sub chkPeriodo_Click()
Dim Co  As DConecta
Dim sql As String
Dim rs As ADODB.Recordset

Dim i As Integer
Dim lnMonto  As Double
lnMonto = 0

'MAVM 20111020
'sql = "select Count(*) Nro from RHPlanillaDetCon where cPlanillaCod ='E05' and cRRHHPeriodo like '" & Format(Me.txtFecha, "YYYYMM") & "%'"
sql = "select Count(*) Nro from RHPlanillaDetCon where cPlanillaCod ='E05' and cRRHHPeriodo like '" & Format(gdFecSis, "YYYYMM") & "%'"
'***

If chkPeriodo.value = 1 Then
    Set Co = New DConecta
    Co.AbreConexion
    Set rs = Co.CargaRecordSet(sql)
    If rs!nro = 0 Then
        MsgBox "Generara la Planilla CTS del Periodo Correspondiente", vbInformation, "AVISO"
        Co.CierraConexion
        Set rs = Nothing
        Set Co = Nothing
        Exit Sub
    Else
        lnMonto = 0
        PrgBar.Visible = True
        PrgBar.Min = 1
        PrgBar.Max = Flex.Rows - 1
        For i = 1 To Flex.Rows - 1
            'MAVM 20111020 ***
            'sql = " select"
            'sql = sql & " Isnull(sum(case when cRHConceptoCod = '130' then nMonto else 0 end),0) C1,"
            'sql = sql & " IsNull(sum(case when cRHConceptoCod = '164' then nMonto else 0 end),0) C2"
            'sql = sql & " from RHPlanillaDetCon  RP"
            'sql = sql & " where cPlanillaCod ='E05' and cRRHHPeriodo like '" & Format(Me.txtFecha, "YYYYMM") & "%' and cRHConceptoCod in ('130','164')"
            'sql = sql & " and cPersCod = '" & Trim(Flex.TextMatrix(i, 14)) & "'"
            'Set rs = Co.CargaRecordSet(sql)
            
            sql = " select"
            sql = sql & " SUM(nProvision) SumProvision, PlaCTS = ISNULL((Select nMonto From RHPlanillaDetCon Where cPlanillaCod = 'E05' and cRRHHPeriodo like '" & Format(gdFecSis, "YYYYMM") & "%'" & " And cPersCod = RH.cPersCod And cRHConceptoCod = 130), 0)"
            sql = sql & " from MovCts MC Inner Join RRHH RH on MC.cRHCod = RH.cRHCod "
            sql = sql & " where cPeriodo In ('201105','201106','201107','201108','201109') And RH.cPersCod = '" & Trim(Flex.TextMatrix(i, 14)) & "'"
            sql = sql & " group by cPersCod"
            Set rs = Co.CargaRecordSet(sql)
                        
            'If Not (rs.EOF And rs.BOF) Then
            '    Flex.TextMatrix(i, 12) = Format(rs!C1 - rs!C2, "#0.00")
            'End If
            'lnMonto = lnMonto + Flex.TextMatrix(i, 12)
            
            If Not (rs.EOF And rs.BOF) Then
                If rs!PlaCTS <> 0 Then
                    If rs!PlaCTS >= rs!SumProvision Then
                        Flex.TextMatrix(i, 12) = Format(rs!PlaCTS - rs!SumProvision, "#0.00")
                    Else
                        Flex.TextMatrix(i, 12) = Format(0, "#0.00")
                    End If
                Else
                    Flex.TextMatrix(i, 12) = Format(0, "#0.00")
                End If
            End If
            lnMonto = lnMonto + Flex.TextMatrix(i, 12)
            '***
            
            Me.PrgBar.value = i
        Next i
        Me.txtTotal = Format(lnMonto, "#,##0.00")
        PrgBar.Visible = False
    End If
    Co.CierraConexion
    Set rs = Nothing
    Set Co = Nothing
End If
End Sub

Private Sub cmdAsiento_Click()
    Dim oCon  As DConecta
    Set oCon = New DConecta
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    Dim oAsi As NContImprimir
    Set oAsi = New NContImprimir
    Dim lsCadena As String
    
    Dim oPrevio As Previo.clsPrevio
    Set oPrevio = New Previo.clsPrevio
    
    Dim lsMovNroEstad As String
    Dim lsMovNroContra As String
    Dim lnMontoEstad As Currency
    
    Dim sql As String
    Dim lbBan As Boolean
    Dim ldFechaAnt As Date
    Dim ldFechaMes As Date
    
    Dim lsMovNro As String
    
    If Not IsDate(Me.txtFecha.Text) Then
        MsgBox "Debe ingresar una fecha valida.", vbInformation, "Aviso"
        Exit Sub
    End If
    lbBan = False
    
    ldFechaAnt = DateAdd("d", -1, CDate("01/" & Format(CDate(Me.txtFecha.Text), "mm/yyyy")))
    ldFechaMes = Format(CDate(Me.txtFecha.Text), "dd/mm/yyyy")
    
    sql = "Select cMovNro  from mov where cmovnro like '" & Format(ldFechaMes, gsFormatoMovFecha) & "%' And cOpeCod = '622401' And nMovflag = 0"
    oCon.AbreConexion
    
    Set rs = oCon.CargaRecordSet(sql)
     
    If Not rs.EOF And Not rs.BOF Then
        lsMovNro = rs!cMovNro
        MsgBox "El Asiento de Provision ya fue generado.", vbInformation
        
        lsCadena = oAsi.ImprimeAsientoContable(lsMovNro, 66, 80, , , , , False)
        oPrevio.Show lsCadena, Me.Caption, True, , gImpresora
        'lbBan = True
        oCon.CierraConexion
        Exit Sub
    Else
        rs.Close
    End If
    
    sql = "Select cMovNro  from mov where cmovnro like '" & Format(ldFechaMes, gsFormatoMovFecha) & "%' And cOpeCod = '622402' And nMovflag = 0"
    oCon.AbreConexion
    
    Set rs = oCon.CargaRecordSet(sql)
     
    If Not rs.EOF And Not rs.BOF Then
        lsMovNro = rs!cMovNro
        MsgBox "El Asiento de Provision ya fue generado.", vbInformation
        
        lsCadena = oAsi.ImprimeAsientoContable(lsMovNro, 66, 80, , , , , False)
        oPrevio.Show lsCadena, Me.Caption, True, , gImpresora
        'lbBan = True
        oCon.CierraConexion
        Exit Sub
    Else
        rs.Close
    End If
    
    'MAVM 20110711 ***
    sql = "Select cPeriodo from MovCTS where cPeriodo like '" & Mid(Format(ldFechaMes, gsFormatoMovFecha), 1, 6) & "'" & " And Not cRHCod = 'XXXXXX'"
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sql)
     
    If rs.RecordCount = 0 Then
        MsgBox "USTED DEBE REALIZAR LA PROVISION ", vbInformation, "AVISO"
        oCon.CierraConexion
        Exit Sub
    Else
        rs.Close
    End If
    '***
    
    'If lbBan Then
    '    oCon.CierraConexion
    '    Exit Sub
    'End If
    
    If MsgBox("Desea Generar Asiento Contable ??? ", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
    
    lsMovNroEstad = GeneraProvisionCTS(gsRHPlanillaCTSProvEst, CDate(Me.txtFecha.Text))
    lsCadena = oAsi.ImprimeAsientoContable(lsMovNroEstad, 66, 80, , , , , False)
    
    oPrevio.Show lsCadena, Me.Caption, True, , gImpresora
End Sub

Private Sub CmdCarga_Click()
ban = False
End Sub

Private Sub cmdConsol_Click()
Dim RHC As DRHCTS
Dim rs As New ADODB.Recordset
Dim fs As Scripting.FileSystemObject
Dim lsArchivo As String

Dim lsNomHoja As String
Dim i, Y As Integer
Dim lbExisteHoja As Boolean

Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet

Dim dProv05Vert, dProv06Vert, dProv07Vert, dProv08Vert, dProv09Vert, dProv10Vert, AA As Double
Dim dProv11Vert, dProv12Vert, dProv01Vert, dProv02Vert, dProv03Vert, dProv04Vert, BB As Double
Dim dProvHorizSum, dProvHorizSumTot, CC As Double

Set RHC = New DRHCTS
Set rs = RHC.CargarProvCTSConsol(Mid(Format(Me.txtFecha.Text, gsFormatoMovFecha), 1, 6))

Screen.MousePointer = 11
lsArchivo = "ProvCTSConsol" & Format(Now, "yyyymm") & "_" & Format(Time(), "HHMMSS") & ".xls"

Set fs = New Scripting.FileSystemObject

Set xlAplicacion = New Excel.Application
If fs.FileExists(App.path & "\Spooler\" & lsArchivo) Then
    Set xlLibro = xlAplicacion.Workbooks.Open(App.path & "\Spooler\" & lsArchivo)
Else
    Set xlLibro = xlAplicacion.Workbooks.Add
End If

lsNomHoja = Format(gdFecSis, "YYYYMM")
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
xlHoja1.Range("B1") = "CAJA MAYNAS"
xlHoja1.Range("B1:C1").MergeCells = True
xlHoja1.Range("B1").Font.Bold = True
xlHoja1.Range("S1") = gdFecSis
xlHoja1.Range("S2") = gsCodUser
xlHoja1.Range("S2").HorizontalAlignment = xlRight
xlHoja1.Range("H1:H2").Font.Bold = True

xlHoja1.Range("B5") = "Cod Emp"
xlHoja1.Range("C5") = "Nombre"
xlHoja1.Range("D5") = "Noviembre"
xlHoja1.Range("E5") = "Diciembre"
xlHoja1.Range("F5") = "Enero"
xlHoja1.Range("G5") = "Febrero"
xlHoja1.Range("H5") = "Marzo"
xlHoja1.Range("I5") = "Abril"
xlHoja1.Range("J5") = "Total Nov Abr"
xlHoja1.Range("K5") = "Pago"
xlHoja1.Range("L5") = "Mayo"
xlHoja1.Range("M5") = "Junio"
xlHoja1.Range("N5") = "Julio"
xlHoja1.Range("O5") = "Agosto"
xlHoja1.Range("P5") = "Setiembre"
xlHoja1.Range("Q5") = "Octubre"
xlHoja1.Range("R5") = "Total May Oct"
xlHoja1.Range("S5") = "Pago"

xlHoja1.Range("B4:S4").MergeCells = True
xlHoja1.Range("B4") = "PROVISION DE CTS " & " " & UCase(Format(txtFecha.Text, "MMMM")) & " DEL " & Format(DateAdd("m", -1, gdFecSis), "YYYY")
xlHoja1.Range("B4").Font.Bold = True
xlHoja1.Range("B4").HorizontalAlignment = xlCenter

xlHoja1.Range("B5:S5").HorizontalAlignment = xlCenter
xlHoja1.Range("B5:S5").Interior.ColorIndex = 35
xlHoja1.Range("B5:S5").Font.Bold = True

xlHoja1.Range("B1").ColumnWidth = 9
xlHoja1.Range("C1").ColumnWidth = 40
xlHoja1.Range("D1").ColumnWidth = 12
xlHoja1.Range("E1").ColumnWidth = 12
xlHoja1.Range("F1").ColumnWidth = 12
xlHoja1.Range("G1").ColumnWidth = 12
xlHoja1.Range("H1").ColumnWidth = 12
xlHoja1.Range("I1").ColumnWidth = 12
xlHoja1.Range("J1").ColumnWidth = 17
xlHoja1.Range("K1").ColumnWidth = 17
xlHoja1.Range("L1").ColumnWidth = 12
xlHoja1.Range("M1").ColumnWidth = 12
xlHoja1.Range("N1").ColumnWidth = 12
xlHoja1.Range("O1").ColumnWidth = 12
xlHoja1.Range("P1").ColumnWidth = 12
xlHoja1.Range("Q1").ColumnWidth = 12
xlHoja1.Range("R1").ColumnWidth = 17
xlHoja1.Range("S1").ColumnWidth = 17

xlHoja1.Application.ActiveWindow.Zoom = 80
Y = 6

For i = 1 To rs.RecordCount
    xlHoja1.Range("B" & Y) = rs!cRHCod
    xlHoja1.Range("C" & Y) = rs!cPersNombre
    
    If (Mid(Format(txtFecha.Text, gsFormatoMovFecha), 5, 2) = "11" Or Mid(Format(txtFecha.Text, gsFormatoMovFecha), 5, 2) = "12" Or Mid(Format(txtFecha.Text, gsFormatoMovFecha), 5, 2) = "01" Or Mid(Format(txtFecha.Text, gsFormatoMovFecha), 5, 2) = "02" Or Mid(Format(txtFecha.Text, gsFormatoMovFecha), 5, 2) = "03" Or Mid(Format(txtFecha.Text, gsFormatoMovFecha), 5, 2) = "04") Then
        xlHoja1.Range("D" & Y) = rs!Prov11
        dProv11Vert = dProv11Vert + rs!Prov11
        xlHoja1.Range("E" & Y) = rs!Prov12
        dProv12Vert = dProv12Vert + rs!Prov12
        xlHoja1.Range("F" & Y) = rs!Prov01
        dProv01Vert = dProv01Vert + rs!Prov01
        xlHoja1.Range("G" & Y) = rs!Prov02
        dProv02Vert = dProv02Vert + rs!Prov02
        xlHoja1.Range("H" & Y) = rs!Prov03
        dProv03Vert = dProv03Vert + rs!Prov03
        xlHoja1.Range("I" & Y) = rs!Prov04
        dProv04Vert = dProv04Vert + rs!Prov04
    End If
    
    If (Mid(Format(txtFecha.Text, gsFormatoMovFecha), 5, 2) = "05" Or Mid(Format(txtFecha.Text, gsFormatoMovFecha), 5, 2) = "06" Or Mid(Format(txtFecha.Text, gsFormatoMovFecha), 5, 2) = "07" Or Mid(Format(txtFecha.Text, gsFormatoMovFecha), 5, 2) = "08" Or Mid(Format(txtFecha.Text, gsFormatoMovFecha), 5, 2) = "09" Or Mid(Format(txtFecha.Text, gsFormatoMovFecha), 5, 2) = "10") Then
        xlHoja1.Range("L" & Y) = rs!Prov05
        dProv05Vert = dProv05Vert + rs!Prov05
        dProvHorizSum = dProvHorizSum + rs!Prov05
        
        xlHoja1.Range("M" & Y) = rs!Prov06
        dProv06Vert = dProv06Vert + rs!Prov06
        dProvHorizSum = dProvHorizSum + rs!Prov06
        
        xlHoja1.Range("N" & Y) = rs!Prov07
        dProv07Vert = dProv07Vert + rs!Prov07
        dProvHorizSum = dProvHorizSum + rs!Prov07
        
        xlHoja1.Range("O" & Y) = rs!Prov08
        dProv08Vert = dProv08Vert + rs!Prov08
        dProvHorizSum = dProvHorizSum + rs!Prov08
        
        xlHoja1.Range("P" & Y) = rs!Prov09
        dProv09Vert = dProv09Vert + rs!Prov09
        dProvHorizSum = dProvHorizSum + rs!Prov09
        
        xlHoja1.Range("Q" & Y) = rs!Prov10
        dProv10Vert = dProv10Vert + rs!Prov10
        dProvHorizSum = dProvHorizSum + rs!Prov10
        xlHoja1.Range("R" & Y) = Format(dProvHorizSum, "#,##0.00")
        
    End If
  
    dProvHorizSumTot = dProvHorizSumTot + dProvHorizSum
    dProvHorizSum = 0
    
    rs.MoveNext
    Y = Y + 1
Next i

xlHoja1.Cells(rs.RecordCount + 6, 2) = "TOTALES"

xlHoja1.Cells(rs.RecordCount + 6, 12) = Format(dProv05Vert, "#,##0.00")
xlHoja1.Cells(rs.RecordCount + 6, 13) = Format(dProv06Vert, "#,##0.00")
xlHoja1.Cells(rs.RecordCount + 6, 14) = Format(dProv07Vert, "#,##0.00")
xlHoja1.Cells(rs.RecordCount + 6, 15) = Format(dProv08Vert, "#,##0.00")
xlHoja1.Cells(rs.RecordCount + 6, 16) = Format(dProv09Vert, "#,##0.00")
xlHoja1.Cells(rs.RecordCount + 6, 17) = Format(dProv10Vert, "#,##0.00")

xlHoja1.Cells(rs.RecordCount + 6, 4) = Format(dProv11Vert, "#,##0.00")
xlHoja1.Cells(rs.RecordCount + 6, 5) = Format(dProv12Vert, "#,##0.00")
xlHoja1.Cells(rs.RecordCount + 6, 6) = Format(dProv01Vert, "#,##0.00")
xlHoja1.Cells(rs.RecordCount + 6, 7) = Format(dProv02Vert, "#,##0.00")
xlHoja1.Cells(rs.RecordCount + 6, 8) = Format(dProv03Vert, "#,##0.00")
xlHoja1.Cells(rs.RecordCount + 6, 9) = Format(dProv04Vert, "#,##0.00")

xlHoja1.Cells(rs.RecordCount + 6, 18) = Format(dProvHorizSumTot, "#,##0.00")

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
MsgBox "Se ha Generado el Archivo " & lsArchivo & " Satisfactoriamente en la carpeta Spooler de SICMACT ADM", vbInformation, "Aviso"

CargaArchivo lsArchivo, App.path & "\SPOOLER\"
Exit Sub
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
xlHoja1.Range("A1") = "CAJA MAYNAS"
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
Dim mes As Integer
Dim i As Integer
Dim Resultado As Integer
Dim DiaUltimo As Integer
Dim Cantidad As Double
Dim opt As Integer
If Not ValFecha(Me.txtFecha) Then
    Exit Sub
End If
FechaProv = Format(Me.txtFecha, "YYYYMM")
mes = Month(Me.txtFecha)
Set RHCTS = New DRHCTS


If mes = 5 Or mes = 11 Then
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

opt = MsgBox("Esta seguro de Realizar la Provision del mes de " & Format(Me.txtFecha, "MMMM"), vbQuestion + vbYesNo, "AVISO")
If vbNo = opt Then Exit Sub

Me.PrgBar.Visible = True
Me.PrgBar.Min = 1
Me.PrgBar.Max = Flex.Rows - 1

If RHCTS.VerificaProvCTSMes(Format(Me.txtFecha, "YYYYMM")) Then
    MsgBox "Ya se Provisiono  el mes de " & Format(Me.txtFecha, "YYYYMM") & ". Tenga cuidado", vbInformation, "AVISO"
    Set RHCTS = Nothing
    Exit Sub
End If

'If -1 = RHCTS.AbonaCTSPasado(gsCodUser, FechaHora(gdFecSis), Format(Me.txtFecha, "YYYYMM")) Then
'    MsgBox "Error de Conexion vuelva a intentarlo", vbInformation, "AVISO"
'    Exit Sub
'End If

For i = 1 To Flex.Rows - 1
    If Flex.TextMatrix(i, 3) = "." Then
        With Flex
            Resultado = RHCTS.ActualizaMesProvison(.TextMatrix(i, 14), .TextMatrix(i, 1), gsCodUser, FechaHora(gdFecSis), Format(Me.txtFecha, "YYYYMM"), 1, 1, , .TextMatrix(i, 12))
            .TextMatrix(i, 13) = IIf(Resultado = 1, "OK", "Error")
        End With
    Else
        With Flex
        DiaUltimo = CInt(Left(DateSerial(Year(Me.txtFecha), Month(Me.txtFecha) + 1, 0), 2))
        
        'Cantidad = (((CInt(Left(Trim(.TextMatrix(i, 4)), 2)) - DiaUltimo) * -1) + 1) / DiaUltimo
        Cantidad = (((CInt(Left(Trim(.TextMatrix(i, 4)), 2)) - DiaUltimo) * -1)) / DiaUltimo 'MAVM 20110711
        'Resultado = RHCTS.ActualizaMesProvison(.TextMatrix(i, 14), .TextMatrix(i, 1), gsCodUser, FechaHora(gdFecSis), Format(Me.txtFecha, "YYYYMM"), 2, Cantidad, 2, .TextMatrix(i, 12))
        Resultado = RHCTS.ActualizaMesProvison(.TextMatrix(i, 14), .TextMatrix(i, 1), gsCodUser, FechaHora(gdFecSis), Format(Me.txtFecha, "YYYYMM"), 1, Cantidad, , .TextMatrix(i, 12))
        Resultado = RHCTS.InsertaCTSTemp(.TextMatrix(i, 14), Cantidad, Format(Me.txtFecha, "YYYYMM"), FechaHora(gdFecSis))
        .TextMatrix(i, 13) = IIf(Resultado = 1, "OK", "Error")
        End With
    End If
    Me.PrgBar.value = i
Next i

'Flex.Visible = False
MsgBox "Provision de CTS realizada", vbInformation, "AVISO"

'MAVM 20111020 ***
If Me.chkPeriodo.value <> 1 Then
    CargaMatrizCTS
End If
'***
Set RHCTS = Nothing
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Sub CargaMatrizCTS()
Dim rs As ADODB.Recordset
Dim FechaGrati As String
Dim FechaCTS As String
Dim Año As Integer
Dim mes As Integer
Dim CTS As DRHCTS
Dim i As Integer
Set CTS = New DRHCTS

Dim lnMonto As Currency

lnMonto = 0

PrgBar.Min = 1
'txtFecha = gdFecSis
mes = Month(Me.txtFecha)

Flex.Rows = 2
For i = 1 To Flex.Cols - 1
    Flex.TextMatrix(1, i) = ""
Next i

Select Case mes
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
        Flex.TextMatrix(Flex.Rows - 1, 8) = Format(IIf(IsNull(rs!Grati), 0, rs!Grati), "#0.00")
        Flex.TextMatrix(Flex.Rows - 1, 9) = Format(IIf(IsNull(rs!GRATI6), 0, rs!GRATI6), "#0.00")
        Flex.TextMatrix(Flex.Rows - 1, 10) = Format(IIf(IsNull(rs!REMUNERA_IND), 0, rs!REMUNERA_IND), "#0.00")
        Flex.TextMatrix(Flex.Rows - 1, 11) = rs!MESCTS
        Flex.TextMatrix(Flex.Rows - 1, 12) = Format(IIf(IsNull(rs!TotalDep), 0, rs!TotalDep), "#0.00")
        Flex.TextMatrix(Flex.Rows - 1, 14) = rs!cPersCod
        'Flex.TextMatrix(Flex.Rows - 1, 15) = rs!cAgenciaAsig
        Flex.TextMatrix(Flex.Rows - 1, 15) = rs!cAgenciaActual 'MAVM 20110711 ***
        lnMonto = lnMonto + Format(IIf(IsNull(rs!TotalDep), 0, rs!TotalDep), "#0.00")
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
Dim mes As Integer
If KeyAscii = 13 Then
    If Not ValFecha(Me.txtFecha) Then
        Me.txtFecha.SetFocus
        Exit Sub
    End If
    mes = Month(Me.txtFecha)
    Me.LBLtITULO.Visible = True
    Me.LBLtITULO = "COMPENSACION POR TIEMPO DE SERVICIO SEMESTRE "
    Select Case mes
        Case 11, 12, 1, 2, 3, 4:
        Me.LBLtITULO = Me.LBLtITULO & " NOV " & Year(Me.txtFecha) - 1 & " - " & " ABR " & Year(Me.txtFecha)
        Case 5, 6, 7, 8, 9, 10:
        Me.LBLtITULO = Me.LBLtITULO & " MAY " & Year(Me.txtFecha) & " - " & " OCT " & Year(Me.txtFecha)
    End Select
    
    CargaMatrizCTS
End If
End Sub

Private Function GeneraProvisionCTS(psOpeCod As String, pdFecha As Date)
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim oCon As New DConecta
    Dim oMov As New DMov
    Dim lnMovNro As Long
    Dim lsMovNro As String
    Dim lnItem As Long
    
    Dim lnFechaAux As Date
    Dim lnAcumDiffCont As Currency
    Dim lnAcumDiffEst As Currency
    Dim lnDiffCont As Currency
    Dim lnDiffEst As Currency
    
    Dim lnUltimo  As Currency
    Dim lsCadena As String
    Dim lnDiffTotal As Currency
    
    Dim CTS As DRHCTS
    Set CTS = New DRHCTS
    
    Dim lnProvAntCont As Currency
    Dim lnProvAntEstab As Currency
    
    Dim lnProvActCont As Currency
    Dim lnProvActEstab As Currency
    Dim lnProvAct As Currency 'MAVM 20110711
    
    Dim FechaGrati As String
    Dim Año As Long
    
    Dim sql1 As String
    Dim opt As Integer
    
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
    
    lnFechaAux = DateAdd("d", -1, CDate("01/" & Format(Month(pdFecha), "00") & "/" & Trim(Str(Year(pdFecha)))))
        
    'sql = "Select dbo.getsaldocta('" & Format(lnFechaAux, gsFormatoFecha) & "','21160102',1)"
    'Set rsTemp = oCon.CargaRecordSet(sql)
    'lnProvAntCont = rsTemp.Fields(0)
    lnProvAntCont = 0
    
    
    sql = "Select dbo.getsaldoctaAcumulado('" & Format(lnFechaAux, gsFormatoFecha) & "','21_601%',1)"
    Set rsTemp = oCon.CargaRecordSet(sql)
    lnProvAntEstab = rsTemp.Fields(0)
    
    
    Select Case Month(lnFechaAux)
    Case 1, 2, 3, 4, 5, 6:
        Año = Year(Me.txtFecha) - 1
        FechaGrati = Año & "12"
    Case 7, 8, 9, 10, 11, 12:
        Año = Year(Me.txtFecha)
        FechaGrati = Año & "07"
    End Select
    
    If Me.chkPeriodo.value = 1 Then
        opt = MsgBox("Esta seguro de que es el ultimo mes del perido del CTS a Provisionar", vbInformation + vbYesNo, "AVISO")
        If opt = vbNo Then Exit Function
        'MAVM 20111020 ***
        'sql1 = " Select Sum(TotalDep) TotalDep, cAgenciaAsig , Cont from ("
        'sql1 = sql1 & "    Select C1 - C2 TOTALDEP, cAgenciaAsig,"
        'sql1 = sql1 & "     case when 0 = dbo.GetRHTpoContrato (cPersCod,'" & Format(lnFechaAux, gsFormatoMovFecha) & "') then 0 else 1 end Cont"
        'sql1 = sql1 & "     from ("
        'sql1 = sql1 & "         select RP.cPersCod,cAgenciaAsig,"
        'sql1 = sql1 & "         sum(case when cRHConceptoCod = '130' then nMonto else 0 end) C1,"
        'sql1 = sql1 & "         sum(case when cRHConceptoCod in ('164X','165X') then nMonto else 0 end) C2"
        'sql1 = sql1 & "         from RHPlanillaDetCon  RP"
        'sql1 = sql1 & "         Inner Join RRHH R on R.cPersCod = RP.cPersCod"
        'sql1 = sql1 & "         where cPlanillaCod ='E05' and cRRHHPeriodo like '" & Format(Me.txtFecha, "YYYYMM") & "%' and cRHConceptoCod in ('130','164','165')"
        'sql1 = sql1 & "         and  RP.cPersCod in (select cPersCod from RRHH  where nRhEstado < 700)"
        'sql1 = sql1 & "         group by RP.cPersCod,cAgenciaAsig"
        'sql1 = sql1 & "     ) ABC"
        'sql1 = sql1 & " ) Total"
        'sql1 = sql1 & " Group by cAgenciaAsig, Cont"
        'sql1 = sql1 & " Order by Cont,cAgenciaAsig"
        'Set rs = oCon.CargaRecordSet(sql1)
        
        sql1 = "Select cAgenciaActual, SUM (nProvision) as TotalDep"
        sql1 = sql1 & " From RRHH RH Inner Join MovCTS MC on RH.cRHCod = MC.cRHCod"
        sql1 = sql1 & " And MC.cPeriodo LIKE '" & Mid(Format(CDate(pdFecha), gsFormatoMovFecha), 1, 6) & "%'"
        sql1 = sql1 & " Group by cAgenciaActual"
        Set rs = oCon.CargaRecordSet(sql1)
        '***
    Else
        Set rs = CTS.Carga_CTS_Mes(Format(lnFechaAux, "YYYYMM"), FechaGrati, Me.txtFecha, True)
    End If
    lnProvActCont = 0
    lnProvActEstab = 0
    
    While Not rs.EOF
        'If rs!Cont = 5 Then
        '    lnProvActCont = lnProvActCont + rs!TotalDep
        'Else
        '    lnProvActEstab = lnProvActEstab + rs!TotalDep
        'End If
        lnProvAct = lnProvAct + Format(rs!TotalDep, "#0.00") 'MAVM 20110711
        rs.MoveNext
    Wend
    
    'lnDiffTotal = CCur(Me.txtTotal.Text) - lnProvAntEstab
    'pdFecha = lnFechaAux
    lsMovNro = oMov.GeneraMovNro(pdFecha, Right(gsCodAge, 2), gsCodUser)
    lnItem = 0

    lnDiffCont = 0
    lnAcumDiffCont = 0
    lnDiffEst = 0
    lnAcumDiffEst = 0

    
    rs.MoveFirst
    oMov.BeginTrans
        oMov.InsertaMov lsMovNro, psOpeCod, "Provisión de Planilla de CTS. " & Format(pdFecha, gsFormatoFechaView)
        lnMovNro = oMov.GetnMovNro(lsMovNro)
        
        While Not rs.EOF
            'lnDiffEst = Round(((rs!TotalDep / lnProvActEstab) * lnDiffTotal), 2)
            'lnAcumDiffEst = lnAcumDiffEst + lnDiffEst
            lnItem = lnItem + 1
            'oMov.InsertaMovCta lnMovNro, lnItem, "451105" & rs!cAgenciaAsig, lnDiffEst
            oMov.InsertaMovCta lnMovNro, lnItem, "451105" & rs!cAgenciaActual, Round(rs!TotalDep, 2)
            rs.MoveNext
        Wend
        
        lnItem = lnItem + 1
        oMov.InsertaMovCta lnMovNro, lnItem, "211601", Round(lnProvAct, 2) * -1
        'lnItem = lnItem + 1
        'oMov.InsertaMovCta lnMovNro, lnItem, "21160101", lnAcumDiffEst * -1
    oMov.CommitTrans
        
    GeneraProvisionCTS = lsMovNro
    oCon.CierraConexion
End Function

