VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmResumenGastosJudiciales 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Resumen de Gastos Judiciales por Agencia"
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10665
   Icon            =   "frmResumenGastosJudiciales.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   10665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Rango de Fechas"
      Height          =   735
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   3615
      Begin MSMask.MaskEdBox txtFechaDel 
         Height          =   345
         Left            =   600
         TabIndex        =   18
         Top             =   240
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   609
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtFechaAl 
         Height          =   345
         Left            =   2340
         TabIndex        =   19
         Top             =   240
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   609
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "AL :"
         Height          =   195
         Left            =   1920
         TabIndex        =   21
         Top             =   360
         Width           =   285
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "DEL :"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   405
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Moneda"
      Height          =   735
      Left            =   3840
      TabIndex        =   14
      Top             =   120
      Width           =   1935
      Begin VB.OptionButton optDolar 
         Caption         =   "Dolar"
         Height          =   195
         Left            =   960
         TabIndex        =   16
         Top             =   360
         Width           =   735
      End
      Begin VB.OptionButton optSoles 
         Caption         =   "Soles"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdExtorna 
      Caption         =   "Extorna Asiento"
      Height          =   345
      Left            =   2520
      TabIndex        =   12
      Top             =   4320
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.ComboBox cboTpo 
      Height          =   315
      Left            =   5895
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   480
      Width           =   3240
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "&Agregar"
      Height          =   345
      Left            =   45
      TabIndex        =   8
      Top             =   4335
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar Asiento"
      Height          =   345
      Left            =   4215
      TabIndex        =   7
      Top             =   4320
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.ComboBox cmbMes 
      Height          =   315
      ItemData        =   "frmResumenGastosJudiciales.frx":030A
      Left            =   2055
      List            =   "frmResumenGastosJudiciales.frx":030C
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   5430
      Visible         =   0   'False
      Width           =   2220
   End
   Begin MSMask.MaskEdBox mskAnio 
      Height          =   300
      Left            =   720
      TabIndex        =   4
      Top             =   5460
      Visible         =   0   'False
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   4
      Mask            =   "####"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   345
      Left            =   9375
      TabIndex        =   2
      Top             =   4335
      Width           =   960
   End
   Begin VB.CommandButton cmdDeprecia 
      Caption         =   "&Generar Cálculo"
      Height          =   345
      Left            =   7740
      TabIndex        =   1
      Top             =   4320
      Width           =   1560
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   345
      Left            =   6690
      TabIndex        =   0
      Top             =   4335
      Width           =   960
   End
   Begin MSDataGridLib.DataGrid grdEntiOpeRecipro 
      Height          =   3255
      Left            =   120
      TabIndex        =   13
      Top             =   960
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   5741
      _Version        =   393216
      BackColor       =   -2147483634
      HeadLines       =   1
      RowHeight       =   17
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label lblTipo 
      AutoSize        =   -1  'True
      Caption         =   "Agencias :"
      Height          =   195
      Left            =   5895
      TabIndex        =   11
      Top             =   180
      Width           =   750
   End
   Begin VB.OLE OleExcel 
      Appearance      =   0  'Flat
      AutoActivate    =   3  'Automatic
      Enabled         =   0   'False
      Height          =   240
      Left            =   1215
      SizeMode        =   1  'Stretch
      TabIndex        =   9
      Top             =   4380
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Label lblMes 
      Caption         =   "Mes :"
      Height          =   210
      Left            =   1560
      TabIndex        =   6
      Top             =   5475
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Label lblAnio 
      Caption         =   "Año :"
      Height          =   210
      Left            =   255
      TabIndex        =   3
      Top             =   5490
      Visible         =   0   'False
      Width           =   705
   End
End
Attribute VB_Name = "frmResumenGastosJudiciales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lsCaption As String
Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet
Dim nMes As Integer
Dim nAnio As Integer
Dim dFecha As Date
Dim rsRes As ADODB.Recordset
Dim lsCodProd As String

Public Sub Ini(psCaption As String)
    lsCaption = psCaption
    Me.Show 1
End Sub
Private Sub cmdDeprecia_Click()

    If Me.optSoles.value = False And Me.optDolar.value = False Then
        MsgBox "Seleccione una moneda", vbInformation, "Aviso"
        Me.optSoles.SetFocus
        Exit Sub
    End If

    If Me.cboTpo.Text = "" Then
        MsgBox "Indique alguna agencia o todas.", vbInformation, "Aviso"
        Me.cboTpo.SetFocus
        Exit Sub
    End If
    If Not IsDate(Me.txtFechaDel.Text) Then
        MsgBox "Debe ingresar una fecha valida.", vbInformation, "Aviso"
        txtFechaDel.SetFocus
        Exit Sub
    ElseIf Not IsDate(Me.txtFechaAl.Text) Then
        MsgBox "Debe ingresar una fecha valida.", vbInformation, "Aviso"
        txtFechaAl.SetFocus
        Exit Sub
    ElseIf CDate(Me.txtFechaAl.Text) < CDate(Me.txtFechaDel.Text) Then
        MsgBox "Debe ingresar mayor que la fecha inicial.", vbInformation, "Aviso"
        txtFechaAl.SetFocus
        Exit Sub
    End If
    
    Call llenagrid

End Sub

Private Sub cmdGrabar_Click()
    
'    Dim oMov As DMov
'    Set oMov = New DMov
'
'    Dim oDep As DOperacion
'    Set oDep = New DOperacion
'
'    Dim oConect As DConecta
'    Set oConect = New DConecta
'
'    Dim lnMovNro As Long
'    Dim lsMovNro As String
'
'    Dim lsTipo As String
'    Dim lsFecha As String
'    Dim i As Integer
'    Dim lnI As Long
'    Dim lnContador As Long
'    Dim lsCtaCont As String
'    Dim oPrevio As clsPrevioFinan
'    Dim oAsiento As NContImprimir
'    Dim nConta As Integer, lcCtaDif As String
'    Dim overi As DOperacion
'    Dim lnDebe As Double, lnHaber As Double, lnTotHaber As Double, lnTotDebe As Double
'    Dim lnDebeME As Double, lnTotDebeME As Double
'    Dim oContFunc As NContFunciones
'    Dim lnMontoPrin As Currency
'    Dim rsAgesDistrib As ADODB.Recordset
'    Dim lsSql As String, lcCtaCont As String
'    Dim rsBuscaCuenta As ADODB.Recordset
'    Dim lnItemDistri As Integer
'    Dim lnRegImporte As Currency
'    Dim lnItemPrin As Integer
'    Dim lnImpoPrin As Currency
'    Dim lnMontoPrME As Currency
'    Dim lnRegImporMETot As Currency
'    Dim lnMonto As Currency
'    Dim lsCtasInexis As String
'    Dim lnRegImporME As Currency
'    Dim lnRegImporteTot As Currency
'    Dim lnImpoPrinME As Currency
'    Dim lnMontoDebePrin As Currency
'    Set oPrevio = New clsPrevioFinan
'    Set oAsiento = New NContImprimir
'
'    Dim rs As ADODB.Recordset, rs1 As ADODB.Recordset
'
'    Dim lcAgeActual As String
'
'    Dim ldFechaDepre As Date
'    Dim ldFechaRegistro As Date
'
'    Dim lnAdjudi  As Double, lnTotProv As Double, lnValorNeto As Double
'    Dim lcCtaHAdju As String, lcCtaDTotProv As String, lcCtaDValorNeto As String
'   Dim lcUltAgeReg As String
'   Dim lnFlag As Integer
'   Dim lnItem As Integer
'   Dim lcMovNro As String
'
''If Me.FlexEdit1.Rows - 2 = 0 Then
''    MsgBox "Genere el cálculo por favor."
''    Exit Sub
''End If
'
'gsOpeCod = "300460"
'lcAgeActual = ""
'
'lnAdjudi = 0
'lnTotProv = 0
'lnValorNeto = 0
'
'lsFecha = mskAnio.Text & Trim(Right(cmbMes.Text, 2))
'
'If Len(Trim(Right(cmbMes.Text, 2))) = 1 Then
'    lsFecha = mskAnio.Text & "0" & Trim(Right(cmbMes.Text, 2))
'Else
'    lsFecha = mskAnio.Text & Trim(Right(cmbMes.Text, 2))
'End If
'
'If MsgBox("¿Desea grabar el asiento? ", vbQuestion + vbYesNo, "Atención") = vbNo Then Exit Sub
'
'lnFlag = 0
'lcUltAgeReg = ""
'For lnI = 1 To Me.FlexEdit1.Rows - 1
'
'    If lnI = 1 Then
'        lcAgeActual = Me.FlexEdit1.TextMatrix(lnI, 1)
'    End If
'
'    If Me.FlexEdit1.TextMatrix(lnI, 1) = lcAgeActual Then
'        lnAdjudi = lnAdjudi + Me.FlexEdit1.TextMatrix(lnI, 13)
'        lnTotProv = lnTotProv + Me.FlexEdit1.TextMatrix(lnI, 17)
'        lnValorNeto = lnValorNeto + Me.FlexEdit1.TextMatrix(lnI, 18)
'
'        lcAgeActual = Me.FlexEdit1.TextMatrix(lnI, 1)
'    Else
'
'        Set rs = New ADODB.Recordset
'
'        Set overi = New DOperacion
'        Set rs = overi.VerificaAsientoCont(gsOpeCod, lsFecha, lcAgeActual)
'        Set rs1 = overi.ObtieneCtasResumenCredPigno
'        Set overi = Nothing
'
'        If Not rs.EOF Then
'            If lcUltAgeReg <> lcAgeActual And lsCtasInexis <> lcAgeActual Then
'                lsCtasInexis = lsCtasInexis + lcAgeActual + "-"
'            End If
'        Else
'            If Abs(lnAdjudi + lnTotProv + lnValorNeto) > 0 Then
'
'                lnItem = 0
'
'                nMes = Val(Trim(Right(cmbMes.Text, 2)))
'                nAnio = Val(mskAnio.Text)
'                dFecha = DateAdd("m", 1, "01/" & Format(nMes, "00") & "/" & Format(nAnio, "0000")) - 1
'                Set oContFunc = New NContFunciones
'                If Not oContFunc.PermiteModificarAsiento(Format(dFecha, gsFormatoMovFecha), False) Then
'                   Set oContFunc = Nothing
'                   MsgBox "Imposible grabar el asiento en un mes cerrado.", vbInformation, "Aviso"
'                   Exit Sub
'                End If
'                ldFechaRegistro = DateAdd("d", -1, DateAdd("m", 1, CDate("01/" & Right(Me.cmbMes.Text, 2) & "/" & Me.mskAnio.Text)))
'
'                oMov.BeginTrans
'
'                    lcUltAgeReg = lcAgeActual
'
'                    'lsMovNro = oMov.GeneraMovNro(ldFechaRegistro, Right(gsCodAge, 2), gsCodUser)
'                    lsMovNro = oMov.GeneraMovNro(ldFechaRegistro, lcAgeActual, gsCodUser)
'                    oMov.InsertaMov lsMovNro, gsOpeCod, "REG. " & Trim(Mid(Me.cboTpo.Text, 1, Len(Me.cboTpo.Text) - 2))
'                    lnMovNro = oMov.GetnMovNro(lsMovNro)
'
'                    lcMovNro = lcMovNro + "'" + lsMovNro + "',"
'
'                    lnFlag = 1
'
'                    lcCtaHAdju = Replace(rs1!HAdju, "AG", lcAgeActual)
'                    lcCtaDTotProv = Replace(rs1!DTotProv, "AG", lcAgeActual)
'                    lcCtaDValorNeto = Replace(rs1!DValorNeto, "AG", lcAgeActual)
'
'                    lnItem = lnItem + 1
'                    oMov.InsertaMovCta lnMovNro, lnItem, lcCtaHAdju, lnAdjudi * -1
'                    lnItem = lnItem + 1
'                    oMov.InsertaMovCta lnMovNro, lnItem, lcCtaDTotProv, lnTotProv
'                    lnItem = lnItem + 1
'                    oMov.InsertaMovCta lnMovNro, lnItem, lcCtaDValorNeto, lnValorNeto
'
'                oMov.CommitTrans
'            End If
'        End If
'
'        lnAdjudi = 0
'        lnTotProv = 0
'        lnValorNeto = 0
'
'        If Me.FlexEdit1.TextMatrix(lnI, 1) <> "" Then
'            lnAdjudi = lnAdjudi + Me.FlexEdit1.TextMatrix(lnI, 13)
'            lnTotProv = lnTotProv + Me.FlexEdit1.TextMatrix(lnI, 17)
'            lnValorNeto = lnValorNeto + Me.FlexEdit1.TextMatrix(lnI, 18)
'
'            lcAgeActual = Me.FlexEdit1.TextMatrix(lnI, 1)
'        End If
'    End If
'Next
'
'    If lsCtasInexis <> "" Then
'        MsgBox "Las siguientes agencias ya tienen asientos generados en este mes y año.: " + Chr(10) + lsCtasInexis, vbOKOnly, "Aviso"
'    End If
'
'    If lnFlag = 1 Then
'
'        lcMovNro = Left(lcMovNro, Len(lcMovNro) - 1) + IIf(Right(lcMovNro, 1) = ",", "", "")
'        oPrevio.Show oAsiento.ImprimeAsientoContResVtaPigno(lcMovNro, 60, 80, "A S I E N T O  C O N T A B L E - " + Trim(Left(Me.cboTpo.Text, 30)) + " - " + CStr(dFecha)), "", True
'    End If

End Sub

Private Sub cmdImprimir_Click()
        
'    Dim rsE As ADODB.Recordset
'    Set rsE = New ADODB.Recordset
'    Dim lsArchivoN As String
'    Dim lbLibroOpen As Boolean
        
'    Dim lsImpre As String
'    If rsRes.EOF And rsRes.BOF Then
'       MsgBox "No Existen datos para Imprimir", vbInformation, "Aviso"
'       Exit Sub
'    End If
'    Set oImp = New NContImprimir
'
'       lsImpre = oImp.ImprimeObjetos(gnLinPage)
'    Set oImp = Nothing
'    EnviaPrevio lsImpre, "Reporte", gnLinPage, False

    Dim rsE As ADODB.Recordset
    Set rsE = New ADODB.Recordset
    Dim lsArchivoN As String
    Dim lbLibroOpen As Boolean

    If rsRes.EOF And rsRes.BOF Then
       MsgBox "Debe generar el calculo primero.", vbInformation, "Aviso"
       Exit Sub
    End If

     MousePointer = 0
    lsArchivoN = App.path & "\Spooler\" & Format(gdFecSis, "yyyymmdd") & ".xls"
    OleExcel.Class = "ExcelWorkSheet"
    lbLibroOpen = ExcelBegin(lsArchivoN, xlAplicacion, xlLibro)
    If lbLibroOpen Then
       Set xlHoja1 = xlLibro.Worksheets(1)
       ExcelAddHoja Format(gdFecSis, "yyyymmdd"), xlLibro, xlHoja1
       Call GeneraReporte(rsRes.DataSource)
       OleExcel.Class = "ExcelWorkSheet"
       ExcelEnd lsArchivoN, xlAplicacion, xlLibro, xlHoja1
       OleExcel.SourceDoc = lsArchivoN
       OleExcel.Verb = 1
       OleExcel.Action = 1
       OleExcel.DoVerb -1
    End If
    MousePointer = 0
    
    
   'Call exportar_Datagrid(Me.grdEntiOpeRecipro, grdEntiOpeRecipro.ApproxCount)
    

    
'    If Me.FlexEdit1.TextMatrix(1, 1) = "" Then
'        MsgBox "Debe Depreciar antes de imprimir.", vbInformation, "Aviso"
'        Me.cmdDeprecia.SetFocus
'        Exit Sub
'    End If
'
'    lsArchivoN = App.path & "\Spooler\" & Format(gdFecSis, "yyyymmdd") & ".xls"
'    OleExcel.Class = "ExcelWorkSheet"
'    lbLibroOpen = ExcelBegin(lsArchivoN, xlAplicacion, xlLibro)
'    If lbLibroOpen Then
'       Set xlHoja1 = xlLibro.Worksheets(1)
'       ExcelAddHoja Format(gdFecSis, "yyyymmdd"), xlLibro, xlHoja1
'       'Call GeneraReporte(Me.FlexEdit1.GetRsNew)
'
'       OleExcel.Class = "ExcelWorkSheet"
'       ExcelEnd lsArchivoN, xlAplicacion, xlLibro, xlHoja1
'       OleExcel.SourceDoc = lsArchivoN
'       OleExcel.Verb = 1
'       OleExcel.Action = 1
'       OleExcel.DoVerb -1
'    End If
'    MousePointer = 0
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()

    Dim overi As DOperacion
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    Dim rsPr As ADODB.Recordset
    Set rsPr = New ADODB.Recordset

    
    Dim oGen As DGeneral
    Set oGen = New DGeneral
    
    Set rs = oGen.GetConstante(1010)
    Me.cmbMes.Clear
    While Not rs.EOF
        cmbMes.AddItem rs.Fields(0) & Space(50) & rs.Fields(1)
        rs.MoveNext
    Wend
    
    
    Set overi = New DOperacion
    Set rs = overi.ObtieneAgencias()
    Set rsPr = overi.ObtieneProductos()
    Set overi = Nothing

    Me.cboTpo.Clear
    While Not rs.EOF
        cboTpo.AddItem rs.Fields(1) & Space(50) & rs.Fields(0)
        rs.MoveNext
    Wend

    While Not rsPr.EOF
        lsCodProd = lsCodProd & rsPr!cValor & ","
        rsPr.MoveNext
    Wend

    lsCodProd = Left(lsCodProd, Len(lsCodProd) - 1)
lsCodProd = lsCodProd
      
    Me.mskAnio.Text = Format(gdFecSis, "yyyy")
     

End Sub

Private Sub GeneraReporte(prRs As ADODB.Recordset)
    Dim i As Integer
    Dim K As Integer
    Dim j As Integer
    Dim nFila As Integer
    Dim nIni  As Integer
    Dim lNegativo As Boolean
    Dim sConec As String
    Dim lsSuma As String
    Dim sTipoGara As String
    Dim sTipoCred As String
    Dim lnAcum As Variant 'Currency
    Dim lnSer As Currency
    
    'i = -1
    
    prRs.MoveFirst
    
    xlHoja1.Cells(3, 1) = "RESUMEN DE GASTOS JUDICIALES POR AGENCIA DEL " + Me.txtFechaDel + " AL " + Me.txtFechaAl
    
    i = 4
    i = i + 1
    
    Dim lmMontoCols() As Double
    
    For j = 0 To prRs.Fields.Count - 1
        xlHoja1.Cells(i + 1, j + 1) = IIf(j = 0, prRs.Fields.Item(j).Name + "/Agencias", "'" + Right(prRs.Fields.Item(j).Name, 2))
        xlHoja1.Range(xlHoja1.Cells(i + 1, j + 1), xlHoja1.Cells(i + 1, j + 1)).Borders.LineStyle = xlContinuous
        xlHoja1.Range(xlHoja1.Cells(i + 1, j + 1), xlHoja1.Cells(i + 1, j + 1)).Interior.Color = RGB(213, 240, 228)
       ReDim lmMontoCols(j)
    Next j
    
    While Not prRs.EOF
        If Len(prRs.Fields(0)) > 0 And IsNumeric(prRs.Fields(0)) Then
            i = i + 1
            For j = 0 To prRs.Fields.Count - 1
                xlHoja1.Cells(i + 1, j + 1) = prRs.Fields(j)
                xlHoja1.Range(xlHoja1.Cells(i + 1, j + 1), xlHoja1.Cells(i + 1, j + 1)).Borders.LineStyle = xlContinuous
    
                If j > 0 Then
                    lmMontoCols(j) = lmMontoCols(j) + prRs.Fields(j)
                End If
            Next j
        End If
        prRs.MoveNext
    Wend
    
    i = i + 1
    xlHoja1.Cells(i + 1, 1) = "Totales"
        xlHoja1.Range(xlHoja1.Cells(i + 1, 1), xlHoja1.Cells(i + 1, 1)).Borders.LineStyle = xlContinuous
    For j = 1 To prRs.Fields.Count - 1
        xlHoja1.Cells(i + 1, j + 1) = lmMontoCols(j)
        xlHoja1.Range(xlHoja1.Cells(i + 1, j + 1), xlHoja1.Cells(i + 1, j + 1)).Borders.LineStyle = xlContinuous
    Next j

    prRs.MoveFirst
    While Not prRs.EOF
        If Len(prRs.Fields(0)) > 0 And Not IsNumeric(prRs.Fields(0)) Then
            i = i + 1
            For j = 0 To prRs.Fields.Count - 1
                xlHoja1.Cells(i + 1, j + 1) = prRs.Fields(j)
                xlHoja1.Range(xlHoja1.Cells(i + 1, j + 1), xlHoja1.Cells(i + 1, j + 1)).Borders.LineStyle = xlContinuous
            Next j
        End If
        prRs.MoveNext
    Wend

    xlHoja1.Range("3:1").Font.Bold = True
    xlHoja1.Columns.AutoFit
    xlHoja1.Columns("A:A").ColumnWidth = 17

End Sub

Private Sub GeneraReporteSaldoHistorico(prRs As ADODB.Recordset)
    Dim i As Integer
    Dim K As Integer
    Dim j As Integer
    Dim nFila As Integer
    Dim nIni  As Integer
    Dim lNegativo As Boolean
    Dim sConec As String
    Dim lsSuma As String
    Dim sTipoGara As String
    Dim sTipoCred As String
    Dim lnAcum As Currency
    
    i = 1
    
    With xlHoja1.Range("A1:N1")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .MergeCells = False
    End With
    xlHoja1.Range("A1:N1").Merge
    xlHoja1.Range("A1:N1").FormulaR1C1 = " REPORTE DE SALDOS HISTORICOS "
    
    With xlHoja1.Range("A2:G2")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .MergeCells = False
    End With
    xlHoja1.Range("A2:G2").Merge
    xlHoja1.Range("A2:G2").FormulaR1C1 = " COSTO DEL ACTIVO FIJO "
    
    With xlHoja1.Range("H2:N2")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .MergeCells = False
    End With
    xlHoja1.Range("H2:N2").Merge
    xlHoja1.Range("H2:N2").FormulaR1C1 = " DEPRECIACION DE ACTIVOS FIJOS "
    
    prRs.MoveFirst
    While Not prRs.EOF
        i = i + 1
        
        If i = 2 Then
            xlHoja1.Cells(i + 1, 1) = "COD CTIVO"
            xlHoja1.Cells(i + 1, 2) = "COD. PATRIM."
            xlHoja1.Cells(i + 1, 3) = "F. ADQUIS"
            xlHoja1.Cells(i + 1, 4) = "SALDO AÑO ANT."
            xlHoja1.Cells(i + 1, 5) = "COMPRAS AÑO"
            xlHoja1.Cells(i + 1, 6) = "RETIROS AÑO"
            xlHoja1.Cells(i + 1, 7) = "SALDO ACTUAL"
            xlHoja1.Cells(i + 1, 8) = "DEP ACUM EJER ANT"
            xlHoja1.Cells(i + 1, 9) = "DEP AL MES ANT"
            xlHoja1.Cells(i + 1, 10) = "DEP DEP MES"
            xlHoja1.Cells(i + 1, 11) = "TOT DEP DEL EJER"
            xlHoja1.Cells(i + 1, 12) = "DEP ACUM DE RET"
            xlHoja1.Cells(i + 1, 13) = "DEP ACUM TOTAL"
            xlHoja1.Cells(i + 1, 14) = "VALOR EN LIBROS"
            
            i = i + 1
            
            xlHoja1.Cells(i + 1, 1) = prRs!Codigo & "-" & prRs!Serie
            xlHoja1.Cells(i + 1, 2) = Format(i - 2, "00000")
            xlHoja1.Cells(i + 1, 3) = "'" & Format(prRs!F_Compra, gsFormatoFechaView)
            xlHoja1.Cells(i + 1, 4) = Format(IIf(Year(prRs!F_Compra) <> Me.mskAnio.Text, prRs!Valor, 0), "#,##0.00")
            xlHoja1.Cells(i + 1, 5) = Format(IIf(Year(prRs!F_Compra) = Me.mskAnio.Text, prRs!Valor, 0), "#,##0.00")
            
            If prRs!Baja = "False" Then
                xlHoja1.Cells(i + 1, 6) = Format(0, "#,##0.00")
            Else
                If Year(prRs!F_Baja) = Me.mskAnio Then
                    xlHoja1.Cells(i + 1, 6) = Format(prRs!Valor, "#,##0.00")
                Else
                    xlHoja1.Cells(i + 1, 6) = Format(0, "#,##0.00")
                End If
            End If
            
            xlHoja1.Cells(i + 1, 7) = Format(xlHoja1.Cells(i + 1, 4) + xlHoja1.Cells(i + 1, 5) - xlHoja1.Cells(i + 1, 6), "#,##0.00")
            
            xlHoja1.Cells(i + 1, 8) = Format(IIf(Year(prRs!F_Compra) <> Me.mskAnio.Text, prRs!Dep_H_Ejer_Ant, 0), "#,##0.00")
            xlHoja1.Cells(i + 1, 9) = IIf(prRs!Baja = "False", Format(prRs!Dep_H_Mes_Ant, "#,##0.00"), Format(0, "#,##0.00"))
            xlHoja1.Cells(i + 1, 10) = Format(prRs!Dep_H_Mes, "#,##0.00")
            xlHoja1.Cells(i + 1, 11) = Format(xlHoja1.Cells(i + 1, 9) + xlHoja1.Cells(i + 1, 10), "#,##0.00")
            If prRs!Baja = "False" Then
                xlHoja1.Cells(i + 1, 12) = Format(0, "#,##0.00")
            Else
                If Year(prRs!F_Baja) = Me.mskAnio Then
                    xlHoja1.Cells(i + 1, 12) = Format(prRs!Dep_H_Mes_Ant, "#,##0.00")
                Else
                    xlHoja1.Cells(i + 1, 12) = Format(0, "#,##0.00")
                End If
            End If
            xlHoja1.Cells(i + 1, 13) = Format(xlHoja1.Cells(i + 1, 8) + xlHoja1.Cells(i + 1, 11) - xlHoja1.Cells(i + 1, 12), "#,##0.00")
            xlHoja1.Cells(i + 1, 14) = Format(xlHoja1.Cells(i + 1, 13) - xlHoja1.Cells(i + 1, 7), "#,##0.00")
        Else
            xlHoja1.Cells(i + 1, 1) = prRs!Codigo & "-" & prRs!Serie
            xlHoja1.Cells(i + 1, 2) = Format(i - 2, "00000")
            xlHoja1.Cells(i + 1, 3) = "'" & Format(prRs!F_Compra, gsFormatoFechaView)
            xlHoja1.Cells(i + 1, 4) = Format(IIf(Year(prRs!F_Compra) <> Me.mskAnio.Text, prRs!Valor, 0), "#,##0.00")
            xlHoja1.Cells(i + 1, 5) = Format(IIf(Year(prRs!F_Compra) = Me.mskAnio.Text, prRs!Valor, 0), "#,##0.00")
            
            If prRs!Baja = "False" Then
                xlHoja1.Cells(i + 1, 6) = Format(0, "#,##0.00")
            Else
                If Year(prRs!F_Baja) = Me.mskAnio Then
                    xlHoja1.Cells(i + 1, 6) = Format(prRs!Valor, "#,##0.00")
                Else
                    xlHoja1.Cells(i + 1, 6) = Format(0, "#,##0.00")
                End If
            End If
            
            xlHoja1.Cells(i + 1, 7) = Format(xlHoja1.Cells(i + 1, 4) + xlHoja1.Cells(i + 1, 5) - xlHoja1.Cells(i + 1, 6), "#,##0.00")
            
            xlHoja1.Cells(i + 1, 8) = Format(IIf(Year(prRs!F_Compra) <> Me.mskAnio.Text, prRs!Dep_H_Ejer_Ant, 0), "#,##0.00")
            xlHoja1.Cells(i + 1, 9) = IIf(prRs!Baja = "False", Format(prRs!Dep_H_Mes_Ant, "#,##0.00"), Format(0, "#,##0.00"))
            xlHoja1.Cells(i + 1, 10) = Format(prRs!Dep_H_Mes, "#,##0.00")
            xlHoja1.Cells(i + 1, 11) = Format(xlHoja1.Cells(i + 1, 9) + xlHoja1.Cells(i + 1, 10), "#,##0.00")
            If prRs!Baja = "False" Then
                xlHoja1.Cells(i + 1, 12) = Format(0, "#,##0.00")
            Else
                If Year(prRs!F_Baja) = Me.mskAnio Then
                    xlHoja1.Cells(i + 1, 12) = Format(prRs!Dep_H_Mes_Ant, "#,##0.00")
                Else
                    xlHoja1.Cells(i + 1, 12) = Format(0, "#,##0.00")
                End If
            End If
            xlHoja1.Cells(i + 1, 13) = Format(xlHoja1.Cells(i + 1, 8) + xlHoja1.Cells(i + 1, 11) - xlHoja1.Cells(i + 1, 12), "#,##0.00")
            xlHoja1.Cells(i + 1, 14) = Format(xlHoja1.Cells(i + 1, 13) - xlHoja1.Cells(i + 1, 7), "#,##0.00")
        End If
        prRs.MoveNext
    Wend
        
    xlHoja1.Cells.Select
    xlHoja1.Columns.AutoFit
    
    xlHoja1.Range("A1:N3").Select
    xlHoja1.Range("A1:N3").Font.Bold = True

    With xlHoja1.Range("A2:G2").Interior
        .ColorIndex = 36
        .Pattern = xlSolid
    End With
    With xlHoja1.Range("H2:N2").Interior
        .ColorIndex = 35
        .Pattern = xlSolid
    End With
    
    xlHoja1.Range("A2:N" & Trim(Str(i + 1))).Select
    xlHoja1.Range("A2:N" & Trim(Str(i + 1))).Borders(xlDiagonalDown).LineStyle = xlNone
    xlHoja1.Range("A2:N" & Trim(Str(i + 1))).Borders(xlDiagonalUp).LineStyle = xlNone
    
    With xlHoja1.Range("A2:N" & Trim(Str(i + 1))).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlHoja1.Range("A2:N" & Trim(Str(i + 1))).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlHoja1.Range("A2:N" & Trim(Str(i + 1))).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlHoja1.Range("A2:N" & Trim(Str(i + 1))).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlHoja1.Range("A2:N" & Trim(Str(i + 1))).Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlHoja1.Range("A2:N" & Trim(Str(i + 1))).Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
End Sub

Private Sub GeneraReporteSaldoAjustado(prRs As ADODB.Recordset)
    Dim i As Integer
    Dim K As Integer
    Dim j As Integer
    Dim nFila As Integer
    Dim nIni  As Integer
    Dim lNegativo As Boolean
    Dim sConec As String
    Dim lsSuma As String
    Dim sTipoGara As String
    Dim sTipoCred As String
    Dim lnAcum As Currency
    
    i = 1
    
    With xlHoja1.Range("A1:S1")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .MergeCells = False
    End With
    xlHoja1.Range("A1:T1").Merge
    xlHoja1.Range("A1:T1").FormulaR1C1 = " REPORTE DE SALDOS AJUSTADOS "
    
    With xlHoja1.Range("A2:L2")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .MergeCells = False
    End With
    xlHoja1.Range("A2:L2").Merge
    xlHoja1.Range("A2:L2").FormulaR1C1 = " COSTO DEL AJUSTADO DE ACTIVO FIJO "
    
    With xlHoja1.Range("M2:T2")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .MergeCells = False
    End With
    xlHoja1.Range("M2:T2").Merge
    xlHoja1.Range("M2:T2").FormulaR1C1 = " DEPRECIACION AJUSTADA DE ACTIVOS FIJOS "
    
    prRs.MoveFirst
    While Not prRs.EOF
        i = i + 1
        
        If i = 2 Then
            xlHoja1.Cells(i + 1, 1) = "COD CTIVO"
            xlHoja1.Cells(i + 1, 2) = "COD. PATRIM."
            xlHoja1.Cells(i + 1, 3) = "F. ADQUIS"
            xlHoja1.Cells(i + 1, 4) = "SALDO ACT HIST"
            xlHoja1.Cells(i + 1, 5) = "SALDO AJUS AÑO ANT"
            xlHoja1.Cells(i + 1, 6) = "COM DEL AÑO"
            xlHoja1.Cells(i + 1, 7) = "RET DEL AÑO"
            xlHoja1.Cells(i + 1, 8) = "FAC DE AJUS"
            xlHoja1.Cells(i + 1, 9) = "REEXP VAL AJUS ANT"
            xlHoja1.Cells(i + 1, 10) = "COM AJUS"
            xlHoja1.Cells(i + 1, 11) = "RET AJUS"
            xlHoja1.Cells(i + 1, 12) = "VAL ACT AJUS"
            xlHoja1.Cells(i + 1, 13) = "DEP ACUM AÑO ANT"
            xlHoja1.Cells(i + 1, 14) = "REEXP DEP AJUS AÑO ANT"
            xlHoja1.Cells(i + 1, 15) = "DEP AJUS EJER MES ANT"
            xlHoja1.Cells(i + 1, 16) = "DEP AJUST DEL MES"
            xlHoja1.Cells(i + 1, 17) = "TOT DEP DEL EJER AJSU"
            xlHoja1.Cells(i + 1, 18) = "DEP AJUS ACUM RET"
            xlHoja1.Cells(i + 1, 19) = "DEP AJUS ACUM TOTAL"
            xlHoja1.Cells(i + 1, 20) = "VALOR EN LIBROS AJUS"
            
            i = i + 1

            xlHoja1.Cells(i + 1, 1) = prRs!Codigo & "-" & prRs!Serie
            xlHoja1.Cells(i + 1, 2) = Format(i - 2, "00000")
            xlHoja1.Cells(i + 1, 3) = "'" & Format(prRs!F_Compra, gsFormatoFechaView)
            xlHoja1.Cells(i + 1, 4) = Format(prRs!Valor, "#,##0.00")
            xlHoja1.Cells(i + 1, 5) = IIf(Year(prRs!F_Compra) = Me.mskAnio, 0, Format(prRs!Valor_Ajustado, "#,##0.00"))
            xlHoja1.Cells(i + 1, 6) = Format(IIf(Year(prRs!F_Compra) = Me.mskAnio, prRs!Valor, 0), "#,##0.00")
            
            If prRs!Baja = "False" Or Not IsNumeric(prRs!F_Baja) Then
                xlHoja1.Cells(i + 1, 7) = Format(0, "#,##0.00")
            Else
                If Year(prRs!F_Baja) = Me.mskAnio Then
                    xlHoja1.Cells(i + 1, 7) = Format(prRs!Valor, "#,##0.00")
                Else
                    xlHoja1.Cells(i + 1, 7) = Format(0, "#,##0.00")
                End If
            End If
            
            xlHoja1.Cells(i + 1, 8) = Format(prRs!F_Ajuste, "#,##0.000")
            xlHoja1.Cells(i + 1, 9) = Format(IIf(Year(prRs!F_Compra) = Me.mskAnio, 0, prRs!Valor_Ajustado) * prRs!F_Ajuste, "#,##0.00")
            xlHoja1.Cells(i + 1, 10) = Format(IIf(Year(prRs!F_Compra) = Year(gdFecSis), prRs!Valor, 0) * prRs!F_Ajuste, "#,##0.00")
            
            If prRs!Baja = "False" Or Not IsNumeric(prRs!F_Baja) Then
                xlHoja1.Cells(i + 1, 11) = Format(0, "#,##0.00")
            Else
                If Year(prRs!F_Baja) = Me.mskAnio Then
                    xlHoja1.Cells(i + 1, 11) = Format(prRs!Valor * prRs!F_Ajuste, "#,##0.00")
                Else
                    xlHoja1.Cells(i + 1, 11) = Format(0, "#,##0.00")
                End If
            End If
            
            xlHoja1.Cells(i + 1, 12) = Format(xlHoja1.Cells(i + 1, 9) + xlHoja1.Cells(i + 1, 10) - xlHoja1.Cells(i + 1, 11), "#,##0.00")
            
            xlHoja1.Cells(i + 1, 13) = Format(IIf(Year(prRs!F_Compra) >= Me.mskAnio, 0, prRs!Valor_Ajustado), "#,##0.00")
            xlHoja1.Cells(i + 1, 14) = Format(IIf(Year(prRs!F_Compra) >= Me.mskAnio, 0, prRs!Valor_Ajustado) * prRs!F_Ajuste, "#,##0.00")
            xlHoja1.Cells(i + 1, 15) = IIf(prRs!Baja = "False", Format(prRs!Dep_A_Mes_Ant, "#,##0.00"), Format(0, "#,##0.00"))
            xlHoja1.Cells(i + 1, 16) = Format(prRs!Dep_A_Mes, "#,##0.00")
            xlHoja1.Cells(i + 1, 17) = Format(xlHoja1.Cells(i + 1, 15) + xlHoja1.Cells(i + 1, 16), "#,##0.00")
            xlHoja1.Cells(i + 1, 18) = IIf(prRs!Baja <> "False", Format(prRs!Dep_A_Mes_Ant, "#,##0.00"), Format(0, "#,##0.00"))
            xlHoja1.Cells(i + 1, 19) = Format(xlHoja1.Cells(i + 1, 14) + xlHoja1.Cells(i + 1, 17) - xlHoja1.Cells(i + 1, 18), "#,##0.00")
            xlHoja1.Cells(i + 1, 20) = Format(xlHoja1.Cells(i + 1, 12) - xlHoja1.Cells(i + 1, 19), "#,##0.00")
        Else
            xlHoja1.Cells(i + 1, 1) = prRs!Codigo & "-" & prRs!Serie
            xlHoja1.Cells(i + 1, 2) = Format(i - 2, "00000")
            xlHoja1.Cells(i + 1, 3) = "'" & Format(prRs!F_Compra, gsFormatoFechaView)
            xlHoja1.Cells(i + 1, 4) = Format(prRs!Valor, "#,##0.00")
            xlHoja1.Cells(i + 1, 5) = IIf(Year(prRs!F_Compra) = Me.mskAnio, 0, Format(prRs!Valor_Ajustado, "#,##0.00"))
            xlHoja1.Cells(i + 1, 6) = Format(IIf(Year(prRs!F_Compra) = Year(gdFecSis), prRs!Valor, 0), "#,##0.00")
            
            If prRs!Baja = "False" Or Not IsNumeric(prRs!F_Baja) Then
                xlHoja1.Cells(i + 1, 7) = Format(0, "#,##0.00")
            Else
                If Year(prRs!F_Baja) = Me.mskAnio Then
                    xlHoja1.Cells(i + 1, 7) = Format(prRs!Valor, "#,##0.00")
                Else
                    xlHoja1.Cells(i + 1, 7) = Format(0, "#,##0.00")
                End If
            End If
            
            xlHoja1.Cells(i + 1, 8) = Format(prRs!F_Ajuste, "#,##0.000")
            xlHoja1.Cells(i + 1, 9) = Format(IIf(Year(prRs!F_Compra) = Me.mskAnio, 0, prRs!Valor_Ajustado) * prRs!F_Ajuste, "#,##0.00")
            xlHoja1.Cells(i + 1, 10) = Format(IIf(Year(prRs!F_Compra) = Year(gdFecSis), prRs!Valor, 0) * prRs!F_Ajuste, "#,##0.00")
            
            If prRs!Baja = "False" Or Not IsNumeric(prRs!F_Baja) Then
                xlHoja1.Cells(i + 1, 11) = Format(0, "#,##0.00")
            Else
                If Year(prRs!F_Baja) = Me.mskAnio Then
                    xlHoja1.Cells(i + 1, 11) = Format(prRs!Valor * prRs!F_Ajuste, "#,##0.00")
                Else
                    xlHoja1.Cells(i + 1, 11) = Format(0, "#,##0.00")
                End If
            End If
            
            xlHoja1.Cells(i + 1, 12) = Format(xlHoja1.Cells(i + 1, 9) + xlHoja1.Cells(i + 1, 10) - xlHoja1.Cells(i + 1, 11), "#,##0.00")
            
            xlHoja1.Cells(i + 1, 13) = Format(IIf(Year(prRs!F_Compra) >= Me.mskAnio, 0, prRs!Valor_Ajustado), "#,##0.00")
            xlHoja1.Cells(i + 1, 14) = Format(IIf(Year(prRs!F_Compra) >= Me.mskAnio, 0, prRs!Valor_Ajustado) * prRs!F_Ajuste, "#,##0.00")
            xlHoja1.Cells(i + 1, 15) = IIf(prRs!Baja = "False", Format(prRs!Dep_A_Mes_Ant, "#,##0.00"), Format(0, "#,##0.00"))
            xlHoja1.Cells(i + 1, 16) = Format(prRs!Dep_A_Mes, "#,##0.00")
            xlHoja1.Cells(i + 1, 17) = Format(xlHoja1.Cells(i + 1, 15) + xlHoja1.Cells(i + 1, 16), "#,##0.00")
            xlHoja1.Cells(i + 1, 18) = IIf(prRs!Baja <> "False", Format(prRs!Dep_A_Mes_Ant, "#,##0.00"), Format(0, "#,##0.00"))
            xlHoja1.Cells(i + 1, 19) = Format(xlHoja1.Cells(i + 1, 14) + xlHoja1.Cells(i + 1, 17) - xlHoja1.Cells(i + 1, 18), "#,##0.00")
            xlHoja1.Cells(i + 1, 20) = Format(-xlHoja1.Cells(i + 1, 12) + xlHoja1.Cells(i + 1, 19), "#,##0.00")

        End If
        prRs.MoveNext
    Wend
        
    xlHoja1.Select
    xlHoja1.Columns.AutoFit
    
    xlHoja1.Range("A1:T3").Select
    xlHoja1.Range("A1:T3").Font.Bold = True

    With xlHoja1.Range("A2:L2").Interior
        .ColorIndex = 36
        .Pattern = xlSolid
    End With
    With xlHoja1.Range("M2:T2").Interior
        .ColorIndex = 35
        .Pattern = xlSolid
    End With
    
    xlHoja1.Range("A2:T" & Trim(Str(i + 1))).Select
    xlHoja1.Range("A2:T" & Trim(Str(i + 1))).Borders(xlDiagonalDown).LineStyle = xlNone
    xlHoja1.Range("A2:T" & Trim(Str(i + 1))).Borders(xlDiagonalUp).LineStyle = xlNone
    
    With xlHoja1.Range("A2:T" & Trim(Str(i + 1))).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlHoja1.Range("A2:T" & Trim(Str(i + 1))).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlHoja1.Range("A2:T" & Trim(Str(i + 1))).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlHoja1.Range("A2:T" & Trim(Str(i + 1))).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlHoja1.Range("A2:T" & Trim(Str(i + 1))).Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlHoja1.Range("A2:T" & Trim(Str(i + 1))).Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
End Sub

Private Sub llenagrid()
    Dim oDepo As DAgencia
    Set oDepo = New DAgencia
    Dim ldFecha As Date, i As Integer
    Dim rs1 As ADODB.Recordset
    Dim rs9 As ADODB.Recordset
    
    Set rs9 = New ADODB.Recordset
    
    'FlexEdit1.Clear
'    ldFecha = CDate("01/" & Format(Trim(Right(Me.cmbMes.Text, 5)), "00") & "/" & Me.mskAnio.Text)
    
    MousePointer = 11
    Set rsRes = oDepo.ObtieneResumenGastosJudiciales(Format(Trim(Right(Me.cboTpo.Text, 3)), "00"), lsCodProd, Format(Me.txtFechaDel.Text, "yyyymmdd"), Format(Me.txtFechaAl.Text, "yyyymmdd"), IIf(Me.optSoles.value, "1", "2"))
    Set Me.grdEntiOpeRecipro.DataSource = rsRes
    MousePointer = 0
    
    'Me.grdEntiOpeRecipro.
    
    'FlexEdit1.rsFlex = oDepo.ObtieneResumenCredPigno(Format(ldFecha, "yyyymmdd"), Format(Trim(Right(Me.cboTpo.Text, 3)), "00"))
    
End Sub

Private Sub txtFechaAl_GotFocus()
fEnfoque txtFechaAl
End Sub

Private Sub txtFechaAl_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If ValidaFecha(txtFechaAl.Text) <> "" Then
      MsgBox "Fecha no válida...!", vbInformation, "Error"
      txtFechaAl.SetFocus
   End If
End If
End Sub

Private Sub txtFechaAl_Validate(Cancel As Boolean)
    If ValidaFecha(txtFechaAl.Text) <> "" Then
       MsgBox "Fecha no válida...!", vbInformation, "Error"
       Cancel = True
    End If
End Sub

Private Sub txtFechaDel_GotFocus()
fEnfoque txtFechaDel
End Sub

Private Sub txtFechaDel_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If ValidaFecha(txtFechaDel.Text) <> "" Then
      MsgBox "Fecha no válida...!", vbInformation, "Error"
      txtFechaDel.SetFocus
   End If
   txtFechaAl.SetFocus
End If

End Sub

Private Sub txtFechaDel_Validate(Cancel As Boolean)
If ValidaFecha(txtFechaDel.Text) <> "" Then
   MsgBox "Fecha no válida...!", vbInformation, "Error"
   Cancel = True
End If
End Sub
