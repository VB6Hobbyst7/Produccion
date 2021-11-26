VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmRHReportesC 
   Caption         =   "Reportes"
   ClientHeight    =   9030
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12660
   Icon            =   "frmRHReportesC.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9030
   ScaleWidth      =   12660
   Begin VB.Frame fremuneracion 
      Height          =   855
      Left            =   9720
      TabIndex        =   38
      Top             =   8280
      Width           =   1935
      Begin VB.OptionButton optLiquid 
         Alignment       =   1  'Right Justify
         Caption         =   "Liquidados"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   480
         Width           =   1575
      End
      Begin VB.OptionButton potActivos 
         Alignment       =   1  'Right Justify
         Caption         =   "Activos"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   120
         Value           =   -1  'True
         Width           =   1575
      End
   End
   Begin VB.Frame fraAgencias 
      Appearance      =   0  'Flat
      Caption         =   "Agencias"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   580
      Left            =   240
      TabIndex        =   34
      Top             =   8400
      Width           =   8715
      Begin VB.CheckBox chkTodos 
         Appearance      =   0  'Flat
         Caption         =   "&Todos"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   150
         TabIndex        =   36
         Top             =   240
         Width           =   930
      End
      Begin Sicmact.TxtBuscar TxtAgencia 
         Height          =   285
         Left            =   1065
         TabIndex        =   35
         Top             =   210
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   0
         sTitulo         =   ""
      End
      Begin VB.Label lblAgencia 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2460
         TabIndex        =   37
         Top             =   195
         Width           =   6165
      End
   End
   Begin VB.Frame fperpres 
      Caption         =   "Periodo"
      Height          =   615
      Left            =   600
      TabIndex        =   29
      Top             =   8640
      Visible         =   0   'False
      Width           =   4575
      Begin VB.ComboBox cmbmes2 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   240
         Width           =   1575
      End
      Begin VB.ComboBox cmbano2 
         Height          =   315
         Left            =   480
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Mes"
         Height          =   195
         Left            =   1800
         TabIndex        =   33
         Top             =   240
         Width           =   300
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Año"
         Height          =   195
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   285
      End
   End
   Begin VB.CommandButton cmdgenerar 
      Caption         =   "Ver"
      Height          =   375
      Left            =   10440
      TabIndex        =   28
      Top             =   7920
      Width           =   1335
   End
   Begin VB.Frame fcts 
      Caption         =   "CTS"
      Height          =   560
      Left            =   240
      TabIndex        =   25
      Top             =   7800
      Width           =   8715
      Begin Sicmact.TxtBuscar txtPlaCTS 
         Height          =   285
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   503
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TipoBusqueda    =   2
         sTitulo         =   ""
      End
      Begin VB.Label lblctsdescripcion 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2040
         TabIndex        =   27
         Top             =   240
         Width           =   6540
      End
   End
   Begin MSComDlg.CommonDialog cdprint 
      Left            =   240
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fperiodo 
      Caption         =   "Periodo"
      Height          =   615
      Left            =   840
      TabIndex        =   18
      Top             =   600
      Width           =   5295
      Begin VB.CommandButton cmdProceso 
         Caption         =   "Procesar"
         Height          =   375
         Left            =   3840
         TabIndex        =   24
         Top             =   210
         Width           =   1335
      End
      Begin VB.ComboBox cmbano 
         Height          =   315
         Left            =   480
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   240
         Width           =   1215
      End
      Begin VB.ComboBox cmbmes 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Año"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   285
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Mes"
         Height          =   195
         Left            =   1800
         TabIndex        =   21
         Top             =   240
         Width           =   300
      End
   End
   Begin VB.Frame fafp 
      Caption         =   "AFP"
      Height          =   615
      Left            =   6360
      TabIndex        =   13
      Top             =   600
      Width           =   4335
      Begin VB.CommandButton cmdMostrar 
         Caption         =   "Ver"
         Height          =   375
         Left            =   2880
         TabIndex        =   23
         Top             =   210
         Width           =   1335
      End
      Begin VB.ComboBox cmbPag 
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   240
         Width           =   855
      End
      Begin VB.ComboBox cmbafp 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Pag 
         AutoSize        =   -1  'True
         Caption         =   "Pag"
         Height          =   195
         Left            =   1560
         TabIndex        =   16
         Top             =   240
         Width           =   285
      End
      Begin VB.Label Label2 
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   615
      End
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   5775
      Left            =   120
      TabIndex        =   12
      Top             =   1680
      Width           =   11655
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
   Begin VB.ComboBox cmbReportes 
      Height          =   315
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   120
      Width           =   6375
   End
   Begin VB.CommandButton cmdProcesar 
      Caption         =   "Procesar"
      Height          =   375
      Left            =   7320
      TabIndex        =   8
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdsalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   10440
      TabIndex        =   6
      Top             =   7440
      Width           =   1335
   End
   Begin VB.CommandButton cmdExportar 
      Caption         =   "Exportar"
      Height          =   375
      Left            =   8520
      TabIndex        =   5
      Top             =   7440
      Width           =   1695
   End
   Begin VB.Frame ffechas 
      Caption         =   "Fecha"
      Height          =   615
      Left            =   840
      TabIndex        =   0
      Top             =   600
      Width           =   5535
      Begin MSMask.MaskEdBox mskFecIni 
         Height          =   300
         Left            =   1080
         TabIndex        =   1
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskFecFin 
         Height          =   300
         Left            =   3600
         TabIndex        =   2
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label lblFecFin 
         Caption         =   "Fecha Fin:"
         Height          =   255
         Left            =   2640
         TabIndex        =   4
         Top             =   270
         Width           =   855
      End
      Begin VB.Label lblFecINi 
         Caption         =   "Fecha Ini:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   270
         Width           =   735
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFLista 
      Height          =   5775
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   10186
      _Version        =   393216
      FixedCols       =   0
      BackColorBkg    =   16777215
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Reporte"
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   570
   End
   Begin VB.OLE OleExcel 
      Class           =   "Excel.Sheet.8"
      Height          =   255
      Left            =   3600
      TabIndex        =   10
      Top             =   7320
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "frmRHReportesC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oReporte As DRHReportes
Dim rs As New ADODB.Recordset
Dim oPla As DActualizaDatosConPlanilla




Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet
Dim Progress As clsProgressBar



Public GSConnRpt As CRAXDRT.Application
Dim crReport As CRAXDRT.Report
Dim crReport1 As CRAXDRT.Report







Private Sub chkTodos_Click()
If Me.chkTodos.value = 1 Then
    Me.TxtAgencia.Text = ""
    Me.lblAgencia.Caption = ""
    Me.TxtAgencia.Enabled = False
    Else
    Me.TxtAgencia.Enabled = True
End If

End Sub

Private Sub cmbReportes_Click()

Select Case Val(Right(cmbReportes.Text, 2))
Case 1
        ffechas.Visible = False
        MSHFLista.Visible = True
        Me.CRViewer1.Visible = False
        fafp.Visible = False
        fperiodo.Visible = False
        cmdProcesar.Visible = True
        cmdExportar.Visible = True
        fcts.Visible = False
        cmdgenerar.Visible = False
        fperpres.Visible = False
        fraAgencias.Visible = False
        fremuneracion.Visible = False
Case 2
        ffechas.Visible = False
        MSHFLista.Visible = True
        Me.CRViewer1.Visible = False
        fafp.Visible = False
        fperiodo.Visible = False
        cmdProcesar.Visible = True
        cmdExportar.Visible = True
        fcts.Visible = False
        cmdgenerar.Visible = False
        fperpres.Visible = False
        fraAgencias.Visible = False
        fremuneracion.Visible = False
Case 3
        ffechas.Visible = True
        MSHFLista.Visible = True
        Me.CRViewer1.Visible = False
        fafp.Visible = False
        fperiodo.Visible = False
        cmdProcesar.Visible = True
        cmdExportar.Visible = True
        fcts.Visible = False
        cmdgenerar.Visible = False
        fperpres.Visible = False
        fraAgencias.Visible = False
        fremuneracion.Visible = False
Case 4
        MSHFLista.Visible = False
        Me.CRViewer1.Visible = True
        fafp.Visible = True
        fperiodo.Visible = True
        ffechas.Visible = False
        cmdProcesar.Visible = False
        cmdExportar.Visible = False
        fcts.Visible = False
        cmdgenerar.Visible = False
        fperpres.Visible = False
        fraAgencias.Visible = False
        fremuneracion.Visible = False
Case 5
        txtPlaCTS.rs = oReporte.GetPlaSemstralCTS("S")
        MSHFLista.Visible = False
        Me.CRViewer1.Visible = True
        fafp.Visible = False
        fperiodo.Visible = False
        ffechas.Visible = False
        cmdProcesar.Visible = False
        cmdExportar.Visible = False
        fcts.Visible = True
        fcts.Top = 480
        fcts.Left = 840
        cmdgenerar.Top = 650
        cmdgenerar.Left = 10000
        cmdgenerar.Visible = True
        fperpres.Visible = False
        fraAgencias.Visible = False
        fremuneracion.Visible = False
Case 6
        txtPlaCTS.rs = oReporte.GetPlaSemstralCTS("M")
        MSHFLista.Visible = False
        Me.CRViewer1.Visible = True
        fafp.Visible = False
        fperiodo.Visible = False
        ffechas.Visible = False
        cmdProcesar.Visible = False
        cmdExportar.Visible = False
        fcts.Visible = True
        cmdgenerar.Top = 650
        cmdgenerar.Left = 10000
        cmdgenerar.Visible = True
        fperpres.Visible = False
        fraAgencias.Visible = True
        fcts.Top = 450
        fcts.Left = 840
        fraAgencias.Left = 840
        fraAgencias.Top = 1000
        fremuneracion.Visible = False
Case 7
        txtPlaCTS.rs = oReporte.GetProvGrati()
        MSHFLista.Visible = False
        Me.CRViewer1.Visible = True
        fafp.Visible = False
        fperiodo.Visible = False
        ffechas.Visible = False
        cmdProcesar.Visible = False
        cmdExportar.Visible = False
        fcts.Visible = True
        cmdgenerar.Top = 650
        cmdgenerar.Left = 10000
        cmdgenerar.Visible = True
        fperpres.Visible = False
        fraAgencias.Visible = True
        fcts.Top = 450
        fcts.Left = 840
        fraAgencias.Left = 840
        fraAgencias.Top = 1000
        fremuneracion.Visible = False
        
Case 8
        ffechas.Visible = False
        MSHFLista.Visible = False
        Me.CRViewer1.Visible = True
        fafp.Visible = False
        fperiodo.Visible = True
        cmdProcesar.Visible = True
        cmdExportar.Visible = True
        fcts.Visible = False
        cmdgenerar.Visible = False
        fperpres.Visible = False
        fraAgencias.Visible = False
        fremuneracion.Visible = True
        fremuneracion.Left = 6200
        fremuneracion.Top = 400
        cmdProceso.Visible = False
        cmdProcesar.Visible = False
        cmdgenerar.Top = 650
        cmdgenerar.Left = 10000
        cmdgenerar.Visible = True

Case 9
        txtPlaCTS.rs = oReporte.GetCalculoEPs()
        MSHFLista.Visible = False
        Me.CRViewer1.Visible = True
        fafp.Visible = False
        fperiodo.Visible = False
        ffechas.Visible = False
        cmdProcesar.Visible = False
        cmdExportar.Visible = False
        fcts.Visible = True
        cmdgenerar.Top = 650
        cmdgenerar.Left = 10000
        cmdgenerar.Visible = True
        fperpres.Visible = False
        fraAgencias.Visible = True
        fcts.Top = 450
        fcts.Left = 840
        fraAgencias.Left = 840
        fraAgencias.Top = 1000
        fremuneracion.Visible = False
 Case 10
        ffechas.Visible = False
        MSHFLista.Visible = False
        Me.CRViewer1.Visible = True
        fafp.Visible = False
        fperiodo.Visible = True
        cmdProcesar.Visible = True
        cmdExportar.Visible = True
        fcts.Visible = False
        cmdgenerar.Visible = False
        fperpres.Visible = False
        fraAgencias.Visible = False
        fremuneracion.Visible = False
        fremuneracion.Left = 6200
        fremuneracion.Top = 400
        cmdProceso.Visible = False
        cmdProcesar.Visible = False
        cmdgenerar.Top = 650
        cmdgenerar.Left = 10000
        cmdgenerar.Visible = True
        
        
Case 50
       fperpres.Top = 600
       fperpres.Left = 840
       fperpres.Visible = True
       MSHFLista.Visible = True
       Me.CRViewer1.Visible = False
       fafp.Visible = False
       fperiodo.Visible = False
       ffechas.Visible = False
       cmdProcesar.Visible = False
       cmdExportar.Visible = False
       fcts.Visible = False
       cmdgenerar.Visible = False
       cmdProcesar.Visible = True
       cmdExportar.Visible = True
       fraAgencias.Visible = False
       fremuneracion.Visible = False
       
Case 51
       fperpres.Top = 600
       fperpres.Left = 840
       fperpres.Visible = True
       MSHFLista.Visible = True
       Me.CRViewer1.Visible = False
       fafp.Visible = False
       fperiodo.Visible = False
       ffechas.Visible = False
       cmdProcesar.Visible = False
       cmdExportar.Visible = False
       fcts.Visible = False
       cmdgenerar.Visible = False
       cmdProcesar.Visible = True
       cmdExportar.Visible = True
       fraAgencias.Visible = False
       fremuneracion.Visible = False
End Select



End Sub

Private Sub cmdExportar_Click()
Dim lsArchivoN As String
    Dim lbLibroOpen As Boolean
    If Me.MSHFLista.TextMatrix(1, 1) = "" Then
        MsgBox "No existen datos.", vbInformation, "Aviso"
        Exit Sub
    End If
    lsArchivoN = App.path & "\Spooler\" & Format(CDate(Date), "yyyy") & Format(Time, "hhmmss") & ".xls"
    OleExcel.Class = "ExcelWorkSheet"
    lbLibroOpen = ExcelBegin(lsArchivoN, xlAplicacion, xlLibro)
    If lbLibroOpen Then
       Set xlHoja1 = xlLibro.Worksheets(1)
       ExcelAddHoja Format(gdFecSis, "yyyymmdd"), xlLibro, xlHoja1
       GeneraReporteRH MSHFLista, xlHoja1
       OleExcel.Class = "ExcelWorkSheet"
       ExcelEnd lsArchivoN, xlAplicacion, xlLibro, xlHoja1
       OleExcel.SourceDoc = lsArchivoN
       OleExcel.Verb = 1
       OleExcel.Action = 1
       OleExcel.DoVerb -1
       OleExcel.Appearance = 0
       OleExcel.Width = 500
    End If
    MousePointer = 0
End Sub

Sub generar_reporte()
Dim sCodafp As String
sCodafp = Right(cmbafp.Text, 2)
Set crReport = Nothing
Set crReport1 = Nothing
                
Select Case Right(cmbReportes.Text, 2)
Case 4
                Select Case Val(cmbPag.Text)
                    Case 1
                            'PAGOSAFP
                             Set crReport = GSConnRpt.OpenReport(App.path & "\DataReport\PAGOSAFP.rpt")
                             crReport.FormulaFields(5).Text = "'" & sCodafp & "'"
                    Case 2
                            'PAGOSAFP_PARTE_2
                             Set crReport = GSConnRpt.OpenReport(App.path & "\DataReport\PAGOSAFP_PARTE_2.rpt")
                             crReport.FormulaFields(5).Text = "'" & sCodafp & "'"
                             'For i = 1 To crReport.FormulaFields.Count
                                          '        If crReport.FormulaFields(i).Name = "{@wtipopago}" Then
                                          '             MsgBox crReport.FormulaFields(i).Name + Str(i)
                                          '              Exit Sub
                                          '          End If
                            'Next
                    Case 3
                            'PAGOSAFP_PARTE_3
                            Set crReport = GSConnRpt.OpenReport(App.path & "\DataReport\PAGOSAFP_PARTE_3.rpt")
                            crReport.FormulaFields(5).Text = "'" & sCodafp & "'"
                    End Select

Case 5
                            If txtPlaCTS.Text = "" Then
                                MsgBox "debe Seleccionar un codigo de planilla"
                                Exit Sub
                            End If
                            
                            Set crReport = GSConnRpt.OpenReport(App.path & "\DataReport\CARTACTSRH.rpt")
                            'crReport.ParameterFields(1).AddCurrentValue ("20050511")
                            crReport.ParameterFields(1).AddCurrentValue (txtPlaCTS.Text)
                            'crReport.ParameterFields(1).AddDefaultValue ("20050511")
                            'Print crReport.ParameterFields(1).value = "'" & "20050511" & "'"

                            
                                                        
                           ' crReport.ParameterFields("{?@Periodo}") = "" & "20050511" & " ; " & "(?@Periodo)" & " ;true "
                            
                            'crReport.ParameterFields(1).AddCurrentValue ("200505011") ' ' = "?@Periodo;" & sCodafp & ";True"
                            
                            
                            'Set crReport.ParameterFields(1).value = "(@Periodo);" & sCodafp & ";false"
                            '.ParameterFields(2) = "pUser;" & gsCodUser & ";True"
                            
                           ' With CR
                           '     .Connect = oConec1.GetStringConnection
                           '     .Connect = "dsn=" & oConec1.servername & ";UID=SA;DSQ=" & oConec1.DatabaseName & ";pwd=cmacica"
                           '     .WindowControls = True
                           '     .WindowState = crptMaximized
                           '
                           '     .ParameterFields(0) = "pMes;" & MonthName(Month(Me.TxtFecIniA02), False) & ";True"
                           '     .ParameterFields(1) = "pAno;" & Year(Me.TxtFecIniA02) & ";True"
                           '     .ParameterFields(2) = "pUser;" & gsCodUser & ";True"
                           '     .ParameterFields(3) = "@dFecha;" & Format(Me.TxtFecIniA02, "MM/dd/yyyy") & ";True"
                           '     .ReportFileName = App.path & "\Rpts\ListaClientesGravament.rpt"
                           '     .Destination = crptToWindow
                           '     '.PrintFileType = crptExcel50
                           '     '.PrintFileName = "C:\ListaClientes.xls"
                           '     .WindowState = crptNormal
                           '     .Action = 1
                           ' End With
                            
                            
                            
                            
                            
                            
                            'For i = 1 To crReport.ParameterFields.Count
                            '                      If crReport.ParameterFields(i).Name = "{@wtipopago}" Then
                            '                           MsgBox crReport.FormulaFields(i).Name + Str(i)
                            '                            Exit Sub
                            '                        End If
                            'Next
                            
Case 6
                          If txtPlaCTS.Text = "" Then
                                MsgBox "debe Seleccionar un codigo de planilla"
                                Exit Sub
                          End If
                          
                          Set crReport = GSConnRpt.OpenReport(App.path & "\DataReport\RH_CTSMENSUAL.rpt")
                           
                          crReport.ParameterFields(1).AddCurrentValue (txtPlaCTS.Text)
                          If chkTodos.value = 1 Then
                           crReport.ParameterFields(2).AddCurrentValue ("T")
                          Else
                          'TxtAgencia.Text
                          crReport.ParameterFields(2).AddCurrentValue (TxtAgencia.Text)
                          End If
                          
Case 7
                          If txtPlaCTS.Text = "" Then
                                MsgBox "debe Seleccionar un codigo de planilla"
                                Exit Sub
                          End If
                          
                          Set crReport = GSConnRpt.OpenReport(App.path & "\DataReport\RH_GRATMENSUAL.rpt")
                           
                          crReport.ParameterFields(1).AddCurrentValue (txtPlaCTS.Text)
                          If chkTodos.value = 1 Then
                           crReport.ParameterFields(2).AddCurrentValue ("T")
                           crReport.FormulaFields(1).Text = "'" & "Todas " & "'"
                          Else
                          'TxtAgencia.Text
                          crReport.ParameterFields(2).AddCurrentValue (TxtAgencia.Text)
                          crReport.FormulaFields(1).Text = "'" & lblAgencia.Caption & "'"
                          End If

Case 8
                          
                          Set crReport = GSConnRpt.OpenReport(App.path & "\DataReport\RH_REMUNERACIONMENSUAL_AG.rpt")
                          crReport.ParameterFields(1).AddCurrentValue (cmbano.Text + Right(cmbmes.Text, 2))
                          crReport.ParameterFields(2).AddCurrentValue (IIf(potActivos.value = True, "A", "L"))
                          
                         
Case 10
                          Set crReport = GSConnRpt.OpenReport(App.path & "\DataReport\RH_TPOCONTRATONIVELREM.rpt")
                          crReport.ParameterFields(1).AddCurrentValue (cmbano.Text + Right(cmbmes.Text, 2))
                          
                         
                         
Case 9
                          If txtPlaCTS.Text = "" Then
                                MsgBox "Debe Seleccionar un Mes de Calculo EPS"
                                Exit Sub
                          End If
                          
                          If TxtAgencia.Text = "" And chkTodos.value = 0 Then
                                MsgBox "Debe Seleccionar una Agencia o Todas ", vbInformation, "Seleccione Agencia"
                                Exit Sub
                          End If
                          
                          Set crReport = GSConnRpt.OpenReport(App.path & "\DataReport\RH_FACT_EPS.rpt")
                           
                          crReport.ParameterFields(1).AddCurrentValue (txtPlaCTS.Text)
                          If chkTodos.value = 1 Then
                           crReport.ParameterFields(2).AddCurrentValue ("T")
                           crReport.FormulaFields(1).Text = "'" & "Todas " & "'"
                          Else
                          'TxtAgencia.Text
                          crReport.ParameterFields(2).AddCurrentValue (TxtAgencia.Text)
                          crReport.FormulaFields(1).Text = "'" & lblAgencia.Caption & "'"
                          End If
                          

End Select
'GSConnRpt.LogOnServer "p2sodbc.dll", "SICMACI", ObtenerBaseDatos, ObtenerUsuarioDatos, ObtenerPassword
GSConnRpt.LogOnServer "p2sodbc.dll", "SICMACI", ObtenerBaseDatos, "USERSICMACCONS", "sicmacicons"
CRViewer1.ReportSource = crReport
CRViewer1.ViewReport



End Sub



Private Sub cmdGenerar_Click()
If Right(cmbReportes.Text, 1) = 5 Then
   Dim Numeros As New clsNumeros
   Dim oconcepto  As DRHConcepto
   Set oconcepto = New DRHConcepto
   'oconcepto.ActualizaLetrasMontoCTS (txtPlaCTS.Text)
End If
generar_reporte
End Sub

Private Sub cmdMostrar_Click()
generar_reporte
End Sub


Private Sub cmdProcesar_Click()
If cmbReportes.Text = "" Then Exit Sub

Select Case Val(Right(cmbReportes.Text, 2))
Case 1
        
        MSHFLista.Clear
        MSHFLista.Rows = 2
        Set rs = oReporte.GetRHlistaTrabDep
        MSHFLista.ColWidth(0) = 1300
        MSHFLista.ColWidth(1) = 3000
        MSHFLista.ColWidth(2) = 2000
        MSHFLista.ColWidth(3) = 2500
        MSHFLista.ColWidth(4) = 2500
        MSHFLista.ColWidth(5) = 2000 'fecha
        MSHFLista.ColWidth(6) = 2000 'fecha
        Set MSHFLista.DataSource = rs
Case 2
        
        MSHFLista.Clear
        MSHFLista.Rows = 2
        Set rs = oReporte.SP_RHlistaAnalistasAg
        MSHFLista.ColWidth(0) = 1300
        MSHFLista.ColWidth(1) = 3000
        MSHFLista.ColWidth(2) = 2000
        MSHFLista.ColWidth(3) = 2500
        MSHFLista.ColWidth(4) = 2500
        MSHFLista.ColWidth(5) = 1 'fecha
        MSHFLista.ColWidth(6) = 2000 'fecha
        Set MSHFLista.DataSource = rs
        
Case 3
        
        MSHFLista.Clear
        MSHFLista.Rows = 2
        MSHFLista.Cols = 11
        MSHFLista.ColWidth(0) = 500
        MSHFLista.ColWidth(1) = 980
        MSHFLista.ColWidth(2) = 2000
        MSHFLista.ColWidth(3) = 980
        MSHFLista.ColWidth(4) = 980
        MSHFLista.ColWidth(5) = 980
        MSHFLista.ColWidth(6) = 980
        MSHFLista.ColWidth(7) = 980
        MSHFLista.ColWidth(8) = 980
        MSHFLista.ColWidth(9) = 980
        MSHFLista.ColWidth(10) = 980
        ReporteQuinta gdFecSis, Me.mskFecIni.Text, Me.mskFecFin.Text, gsEmpresa
        'lsCad = oRepEvento.Rep5taRRHH(gdFecSis, Me.mskFecIni.Text, Me.mskFecFin.Text, gsEmpresa)
        
         
        If MSHFLista.TextMatrix(MSHFLista.Rows - 1, 2) <> "" Then
        MSHFLista.Rows = MSHFLista.Rows + 1
        MSHFLista.TextMatrix(MSHFLista.Rows - 1, 0) = "Total"
        For J = 3 To Me.MSHFLista.Cols - 1
            lnAcumulador = 0
            If Left(MSHFLista.TextMatrix(0, J), 2) <> "U_" And Left(MSHFLista.TextMatrix(0, J), 1) <> "_" Then
                For i = 1 To Me.MSHFLista.Rows - 2
                    If MSHFLista.TextMatrix(i, J) <> "" Then
                        lnAcumulador = lnAcumulador + CCur(MSHFLista.TextMatrix(i, J))
                        
                        If MSHFLista.TextMatrix(i, 8) < 0 Then
                            MSHFLista.Row = i
                            MSHFLista.Col = J
                            MSHFLista.CellBackColor = RGB(300, 150, 150)
                        End If
                        
                    End If
                Next i
                MSHFLista.TextMatrix(MSHFLista.Rows - 1, J) = Format(lnAcumulador, "#,##.00")
                MSHFLista.Row = MSHFLista.Rows - 1
                MSHFLista.Col = J
                MSHFLista.CellBackColor = &HA0C000
                'FlexPrePla.CellFontBold = True
                lnAcumulador = lnAcumulador + CCur(MSHFLista.TextMatrix(i, J))
                
            End If
        Next J
        End If
        MSHFLista.TextMatrix(MSHFLista.Rows - 1, 1) = MSHFLista.Rows - 2
Case 4
Case 50
        Set rs = oPla.GetRHListaMontoCargo(Trim(cmbano2.Text) + Format(Right(cmbmes2.Text, 2), "00"))
        Set MSHFLista.Recordset = rs
        
        MSHFLista.ColWidth(0) = 1200
        MSHFLista.ColWidth(1) = 1200
        MSHFLista.ColWidth(2) = 1200
        MSHFLista.ColWidth(3) = 1200
        MSHFLista.ColWidth(4) = 1200
        MSHFLista.ColWidth(5) = 1200
        MSHFLista.ColWidth(6) = 1200
        MSHFLista.ColWidth(7) = 1200
Case 51
       
        Set rs = oPla.GetRHListaMontoCargoDet(Trim(cmbano2.Text) + Format(Right(cmbmes2.Text, 2), "00"))
        Set MSHFLista.Recordset = rs
                MSHFLista.ColWidth(0) = 1200
                MSHFLista.ColWidth(1) = 1200
                MSHFLista.ColWidth(2) = 1200
                MSHFLista.ColWidth(3) = 1200
                MSHFLista.ColWidth(4) = 1200
                MSHFLista.ColWidth(5) = 1200
                MSHFLista.ColWidth(6) = 1200
                MSHFLista.ColWidth(7) = 1200
        
        If MSHFLista.TextMatrix(MSHFLista.Rows - 1, 2) <> "" Then
        MSHFLista.Rows = MSHFLista.Rows + 1
        MSHFLista.TextMatrix(MSHFLista.Rows - 1, 0) = "Total"
        For J = 3 To Me.MSHFLista.Cols - 1
            lnAcumulador = 0
            If Left(MSHFLista.TextMatrix(0, J), 2) <> "U_" And Left(MSHFLista.TextMatrix(0, J), 1) <> "_" Then
                For i = 1 To Me.MSHFLista.Rows - 2
                    If MSHFLista.TextMatrix(i, J) <> "" Then
                        lnAcumulador = lnAcumulador + CCur(MSHFLista.TextMatrix(i, J))
                    End If
                Next i
                MSHFLista.TextMatrix(MSHFLista.Rows - 1, J) = Format(lnAcumulador, "#,##.00")
                MSHFLista.Row = MSHFLista.Rows - 1
                MSHFLista.Col = J
                MSHFLista.CellBackColor = &HA0C000
                'FlexPrePla.CellFontBold = True
                lnAcumulador = lnAcumulador + CCur(MSHFLista.TextMatrix(i, J))
                
            End If
        Next J
        End If
        
        MSHFLista.TextMatrix(MSHFLista.Rows - 1, 1) = MSHFLista.Rows - 2
        
End Select

End Sub

Private Sub cmdProceso_Click()
Dim sAno As String
Dim sMes As String
sAno = cmbano.Text
sMes = Right(cmbmes.Text, 2)
Dim i As Long
oRepEvento_ShowProgress


For i = 1 To 50
oRepEvento_Progress i, 100
Next
oReporte.GeneraReporteAFP sAno + sMes
For i = 51 To 100
oRepEvento_Progress i, 100
Next

oRepEvento_CloseProgress
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub



Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
Dim oCon As DConstantes

Set oReporte = New DRHReportes
Set rs = New ADODB.Recordset

Set oPla = New DActualizaDatosConPlanilla

Set oCon = New DConstantes
 Me.TxtAgencia.rs = oCon.GetAgencias(, , True)

Set rs = oReporte.GetRHReportes
CargaCombo rs, cmbReportes
cmbReportes.ListIndex = 0
Set Progress = New clsProgressBar


mskFecIni.Text = Format(gdFecSis, gsFormatoFechaView)
mskFecFin.Text = Format(gdFecSis, gsFormatoFechaView)
Me.Width = 11985
Me.Height = 8250

cmbafp.AddItem "Integra                      IN"
cmbafp.AddItem "Union Vida                   UV"
cmbafp.AddItem "Profuturo                    PR"
cmbafp.AddItem "Horizonte                    HO"
cmbafp.ListIndex = 0

cmbPag.AddItem "1"
cmbPag.AddItem "2"
cmbPag.AddItem "3"

cmbafp.ListIndex = 0


cmbano.AddItem "2005"
cmbano.AddItem "2006"
cmbano.ListIndex = 0

cmbmes.AddItem "ENERO" & Space(20) & "01"
cmbmes.AddItem "FEBRERO" & Space(20) & "02"
cmbmes.AddItem "MARZO" & Space(20) & "03"
cmbmes.AddItem "ABRIL" & Space(20) & "04"
cmbmes.AddItem "MAYO" & Space(20) & "05"
cmbmes.AddItem "JUNIO" & Space(20) & "06"
cmbmes.AddItem "JULIO" & Space(20) & "07"
cmbmes.AddItem "AGOSTO" & Space(20) & "08"
cmbmes.AddItem "SETIEMBRE" & Space(20) & "09"
cmbmes.AddItem "OCTUBRE" & Space(20) & "10"
cmbmes.AddItem "NOVIEMBRE" & Space(20) & "11"
cmbmes.AddItem "DICIEMBRE" & Space(20) & "12"
cmbmes.ListIndex = 0
cmbPag.ListIndex = 0



cmbano2.AddItem "2005"
cmbano2.AddItem "2006"
cmbano2.ListIndex = 0


cmbmes2.AddItem "ENERO" & Space(20) & "01"
cmbmes2.AddItem "FEBRERO" & Space(20) & "02"
cmbmes2.AddItem "MARZO" & Space(20) & "03"
cmbmes2.AddItem "ABRIL" & Space(20) & "04"
cmbmes2.AddItem "MAYO" & Space(20) & "05"
cmbmes2.AddItem "JUNIO" & Space(20) & "06"
cmbmes2.AddItem "JULIO" & Space(20) & "07"
cmbmes2.AddItem "AGOSTO" & Space(20) & "08"
cmbmes2.AddItem "SETIEMBRE" & Space(20) & "09"
cmbmes2.AddItem "OCTUBRE" & Space(20) & "10"
cmbmes2.AddItem "NOVIEMBRE" & Space(20) & "11"
cmbmes2.AddItem "DICIEMBRE" & Space(20) & "12"

cmbmes2.ListIndex = 0



'********** Coneccion para los Reportes *****************
Set GSConnRpt = New CRAXDRT.Application
'GSConnRpt.LogOnServer "p2sodbc.dll", "SICMACI", ObtenerBaseDatos, ObtenerUsuarioDatos, ObtenerPassword
GSConnRpt.LogOnServer "p2sodbc.dll", "SICMACI", ObtenerBaseDatos, "USERSICMACCONS", "sicmacicons"


End Sub

Private Sub GeneraReporteRH(pflex As MSHFlexGrid, pxlHoja1 As Excel.Worksheet, Optional pnColFiltroVacia As Integer = 0)
    Dim i As Integer
    Dim K As Integer
    Dim J As Integer
    Dim nFila As Integer
    Dim nIni  As Integer
    Dim lNegativo As Boolean
    Dim sConec As String
    Dim lsSuma As String
    Dim sTipoGara As String
    Dim sTipoCred As String
    Dim lnAcum As Currency
    
    For i = 0 To pflex.Rows - 1
        If pnColFiltroVacia = 0 Then
            For J = 0 To pflex.Cols - 1
                pxlHoja1.Cells(i + 1, J + 1) = pflex.TextMatrix(i, J)
            Next J
        Else
            If pflex.TextMatrix(i, pnColFiltroVacia) <> "" Then
                For J = 0 To pflex.Cols - 1
                    pxlHoja1.Cells(i + 1, J + 1) = pflex.TextMatrix(i, J)
                Next J
            End If
        End If
    Next i
    
End Sub

Sub ReporteQuinta(pgdFecSis As Date, psFecIni As String, psFecFin As String, psEmpresa As String)
 Dim sqlE As String
    Dim rsE As ADODB.Recordset
    Set rsE = New ADODB.Recordset
    Dim lsCadena As String
    Dim lnMargen As Integer
    Dim lnPagina As Integer
    Dim lnItem As Long
    Dim lsCadAux1 As String
    Dim lsCadAux4 As String
    
    Dim lsProyeccion As String
    Dim lsIngAcumulado As String
    Dim lsImpProyeccion As String
    Dim lsImpAcumulado As String
    Dim lsImpuesto As String
    
    Dim lsCodigo As String * 10
    Dim lsNombre As String * 35
    Dim lsVProy As String * 18
    Dim lsVIngMes As String * 18
    Dim lsVIngAcum As String * 18
    Dim lsVValUIT As String * 18
    Dim lsVIngAfecto As String * 18
    Dim lsVImpuesto As String * 18
    Dim lsVRetencion As String * 18
    Dim lsVImpuestoMes As String * 18
    
    Dim oRep As DRHReportes
    Set oRep = New DRHReportes
    Dim oInterprete As DInterprete
    Set oInterprete = New DInterprete
    
    Dim lsCadUIT7 As String
    Dim lsCadUIT27 As String
    Dim lsCadUIT54 As String
    Dim lsCadPorHasta27 As String
    Dim lsCadPorHasta54 As String
    Dim lsCadPorMas54 As String
    
    Dim lnCorr As Long
    
    Set rsE = oRep.Rep5taRRHH
    
    'Item
    'Código
    'Apellidos y Nombres
    'Ingreso Mes
    'Ing.Acumulado
    'Ing.Anu.Proy
    'Val.UIT
    'Ing.Afecto
    'Impuesto
    'Impu.Rete
    'Impu.Mes
    
    MSHFLista.Cols = 11
    MSHFLista.TextMatrix(0, 0) = "Item"
    MSHFLista.TextMatrix(0, 1) = "Código"
    MSHFLista.TextMatrix(0, 2) = "Apellidos y Nombres"
    MSHFLista.TextMatrix(0, 3) = "Ing.Mes"
    MSHFLista.TextMatrix(0, 4) = "Ing.Acumul"
    MSHFLista.TextMatrix(0, 5) = "Ing.Anu.Proy"
    MSHFLista.TextMatrix(0, 6) = "Val.UIT"
    MSHFLista.TextMatrix(0, 7) = "Ing.Afecto"
    MSHFLista.TextMatrix(0, 8) = "Impuesto"
    MSHFLista.TextMatrix(0, 9) = "Impu.Rete"
    MSHFLista.TextMatrix(0, 10) = "Impu.Mes"
    
    
    
    lsCadena = ""
    If Not (rsE.EOF And rsE.BOF) Then
        lsCadena = lsCadena & Space(lnMargen) & CentrarCadena(psEmpresa, 180) & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
        lsCadena = lsCadena & Space(lnMargen) & CentrarCadena("DETALLE DE RETENCIONES-" & Format(pgdFecSis, gsFormatoFechaView), 180) & oImpresora.gPrnSaltoLinea
        oInterprete.Interprete_InI
        lsCadena = lsCadena & Space(lnMargen) & Encabezado("Item;4; ;2;Código;7; ;3;Apellidos y Nombres;23; ;12;Ingreso mes;16; ;2;Ing. Acumulado;17; ;1;Ing.Anu.Proy;15; ;3;Val.UIT;15; ;3;Ing.Afecto;15; ;3;Impuesto;12; ;6;Impu.Rete;15; ;3;Impu.Mes;15; ;3;", lnItem)
        
        lsCadUIT7 = ExprANum(oInterprete.FunEvalua("V_UIT_7", "", CDate(psFecIni), CDate(psFecFin), False, "VVVV", ""))
        lsCadUIT27 = ExprANum(oInterprete.FunEvalua("V_UIT_27", "", CDate(psFecIni), CDate(psFecFin), False, "VVVV", ""))
        lsCadUIT54 = ExprANum(oInterprete.FunEvalua("V_UIT_54", "", CDate(psFecIni), CDate(psFecFin), False, "VVVV", ""))
        lsCadPorHasta27 = oInterprete.FunEvalua("V_POR_IMP_5TA", "", CDate(psFecIni), CDate(psFecFin), False, "VVVV", "")
        lsCadPorHasta54 = oInterprete.FunEvalua("V_POR_5TA_H54", "", CDate(psFecIni), CDate(psFecFin), False, "VVVV", "")
        lsCadPorMas54 = oInterprete.FunEvalua("V_POR_5TA_M54", "", CDate(psFecIni), CDate(psFecFin), False, "VVVV", "")
            
        RSet lsVValUIT = Format(lsCadUIT7, "#.##0.00")
            
            
         oRepEvento_ShowProgress
         lnCorr = 0
        
        While Not rsE.EOF
            lnCorr = lnCorr + 1
            MSHFLista.Rows = MSHFLista.Rows + 1
            lsCodigo = rsE!cRhCod
            lsNombre = PstaNombre(rsE!cPersNombre, False)
            oInterprete.Reinicia
            
            lsCadAux1 = ExprANum(oInterprete.FunEvalua("I_REM_NO_AFEC", rsE!cPersCod, CDate(psFecIni), CDate(psFecFin), False, "", ""))
            lsCadAux4 = Format(oInterprete.GetImp5taEmpRRHH(rsE!cPersCod, CCur(lsCadAux1), CDate(psFecIni), CDate(psFecFin), lsCadUIT7, lsCadUIT27, lsCadUIT54, lsCadPorHasta27, lsCadPorHasta54, lsCadPorMas54, lsCadAux1, lsIngAcumulado, lsProyeccion, lsImpProyeccion, lsImpAcumulado, "", ""), "#0.0000")
            
            RSet lsVIngMes = Format(lsCadAux1, "#.##0.00")
            RSet lsVProy = Format(lsProyeccion, "#.##0.00")
            RSet lsVIngAcum = Format(lsIngAcumulado, "#.##0.00")
            
            If CCur(lsProyeccion) - CCur(lsCadUIT7) < 0 Then
                RSet lsVIngAfecto = Format(0, "#.##0.00")
            Else
                RSet lsVIngAfecto = Format(CCur(lsProyeccion) - CCur(lsCadUIT7), "#.##0.00")
            End If
            lsVImpuesto = FillNum(Format(lsImpProyeccion, "#.##0.00"), 18, " ")
            lsVRetencion = FillNum(Format(lsImpAcumulado, "#.##0.00"), 18, " ")
            lsVImpuestoMes = FillNum(Format(lsCadAux4, "#.##0.00"), 18, " ")
            
            lsCadena = lsCadena & Space(lnMargen) & Format(lnCorr, "0000") & "  " & lsCodigo & lsNombre & lsVIngMes & lsVIngAcum & lsVProy & lsVValUIT & lsVIngAfecto & lsVImpuesto & lsVRetencion & lsVImpuestoMes & oImpresora.gPrnSaltoLinea
            
            MSHFLista.TextMatrix(lnCorr, 0) = Format(lnCorr, "0000")
            MSHFLista.TextMatrix(lnCorr, 1) = lsCodigo
            MSHFLista.TextMatrix(lnCorr, 2) = lsNombre
            MSHFLista.TextMatrix(lnCorr, 3) = Val(lsVIngMes)
            MSHFLista.TextMatrix(lnCorr, 4) = Val(lsVIngAcum)
            MSHFLista.TextMatrix(lnCorr, 5) = Val(lsVProy)
            MSHFLista.TextMatrix(lnCorr, 6) = Val(lsVValUIT)
            MSHFLista.TextMatrix(lnCorr, 7) = Val(lsVIngAfecto)
            MSHFLista.TextMatrix(lnCorr, 8) = Val(lsVImpuesto)
            MSHFLista.TextMatrix(lnCorr, 9) = Val(lsVRetencion)
            MSHFLista.TextMatrix(lnCorr, 10) = Val(lsVImpuestoMes)
            
             oRepEvento_Progress rsE.Bookmark, rsE.RecordCount
            rsE.MoveNext
        Wend
        
        oRepEvento_CloseProgress
    End If
    MSHFLista.Rows = MSHFLista.Rows - 1
    rsE.Close
    Set rsE = Nothing
    Set oRep = Nothing
    Set oInterprete = Nothing
    
    'Rep5taRRHH = lsCadena
End Sub

Private Sub oRepEvento_CloseProgress()
    Progress.CloseForm Me
End Sub

Private Sub oRepEvento_Progress(pnValor As Long, pnTotal As Long)
    Progress.Max = pnTotal
    Progress.Progress pnValor, "Generando Reporte"
End Sub

Private Sub oRepEvento_ShowProgress()
    Progress.ShowForm Me
End Sub



Private Sub mskFecIni_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    mskFecFin.SetFocus
End If
End Sub


Private Sub txtAgencia_EmiteDatos()
 Me.lblAgencia.Caption = TxtAgencia.psDescripcion
End Sub

Private Sub txtPlaCTS_EmiteDatos()
lblctsdescripcion.Caption = txtPlaCTS.psDescripcion
End Sub
