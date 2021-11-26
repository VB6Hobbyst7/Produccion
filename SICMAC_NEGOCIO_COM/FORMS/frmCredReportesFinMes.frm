VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmCredReportesFinMes 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9225
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   9225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin RichTextLib.RichTextBox RTF 
      Height          =   420
      Left            =   8340
      TabIndex        =   10
      Top             =   4305
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   741
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmCredReportesFinMes.frx":0000
   End
   Begin VB.Frame frmPlanCuentas 
      Caption         =   "Plan de Cuentas"
      Height          =   825
      Left            =   5280
      TabIndex        =   7
      Top             =   4320
      Width           =   1635
      Begin VB.OptionButton optPlanActual 
         Caption         =   "Plan Contable"
         Enabled         =   0   'False
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1395
      End
      Begin VB.OptionButton optPlanNuevo 
         Caption         =   "Manual Nuevo"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   540
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.Frame FraCambio 
      Height          =   735
      Left            =   5640
      TabIndex        =   4
      Top             =   5400
      Width           =   3255
      Begin VB.TextBox TxtCambio 
         Height          =   285
         Left            =   1320
         TabIndex        =   5
         Top             =   300
         Width           =   1095
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Cambio"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   885
      End
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
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
      Height          =   435
      Left            =   5760
      TabIndex        =   3
      Top             =   6480
      Width           =   1245
   End
   Begin VB.CommandButton cmdSalir 
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
      Height          =   435
      Left            =   7080
      TabIndex        =   2
      Top             =   6480
      Width           =   1245
   End
   Begin VB.Frame FrameOperaciones 
      Caption         =   "Lista de Reportes"
      Height          =   6975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5055
      Begin MSComctlLib.TreeView tvwReporte 
         Height          =   6555
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   11562
         _Version        =   393217
         HideSelection   =   0   'False
         LabelEdit       =   1
         Style           =   7
         ImageList       =   "imglstFiguras"
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.ImageList imglstFiguras 
         Left            =   240
         Top             =   4320
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCredReportesFinMes.frx":0082
               Key             =   "Padre"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCredReportesFinMes.frx":039C
               Key             =   "Hijo"
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmCredReportesFinMes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Co As New nCredRepoFinMes
Dim P As Previo.clsPrevio
Dim Conecta As DConecta

Dim TipoCambio As Currency
Dim fnRepoSelec As Long
Dim lsRtfImpG As String




Private Sub cmdImprimir_Click()

Dim CredRepoMEs As nCredRepoFinMes
Set CredRepoMEs = New nCredRepoFinMes
Dim FMes As Date
Dim Fechaini As String
Select Case fnRepoSelec

Case 108701:
         lsRtfImpG = ""
         lsRtfImpG = Co.nRepo108701_CarteraColocacionesxMoneda(IIf(optPlanActual.value = True, "1", "2"), "dbCmactConsolidada..")
         If lsRtfImpG = "" Then
            MsgBox "No Existen Datos para este Reporte", vbInformation, "AVISO,"
         Else
            P.Show lsRtfImpG, "Situacion de Cartera Colocaciones x Moneda", True
         End If
Case 108702:
        lsRtfImpG = ""
        
        lsRtfImpG = Co.nRepo108702_ImpCarteraCredConsolidada(IIf(optPlanActual.value = True, "1", "2"), TxtCambio, "dbCmactConsolidada..")
        If lsRtfImpG = "" Then
           MsgBox "No Existen Datos para este Reporte", vbInformation, "AVISO,"
        Else
           P.Show lsRtfImpG, "Situacion de Cartera Colocaciones x Moneda", True
        End If
Case 108703:
        lsRtfImpG = ""
        lsRtfImpG = Co.nRepo108703_ImpRepCarteraProd_Venc(IIf(optPlanActual.value = True, "1", "2"), Val(TxtCambio), "dbCmactConsolidada..")
        If lsRtfImpG = "" Then
           MsgBox "No Existen Datos para este Reporte", vbInformation, "AVISO,"
        Else
           P.Show lsRtfImpG, "Resumen por Producto y Vencimiento  (A-2.1)", True
        End If
Case 108704: '  Reporte por  Producto  Y Agencia (A-2.3)
         lsRtfImpG = ""
         lsRtfImpG = Co.nRepo108704_ImpRepCarteraAgencia_Prod(IIf(optPlanActual.value = True, "1", "2"), Val(TxtCambio), "C", "dbCmactConsolidada..")
         If lsRtfImpG = "" Then
           MsgBox "No Existen Datos para este Reporte", vbInformation, "AVISO,"
         Else
           P.Show lsRtfImpG, "Reporte por  Producto  Y Agencia (A-2.3)", True
         End If

Case 108705: 'Reporte para Reclasificacion de Cartera (A-4)
        lsRtfImpG = ""
        lsRtfImpG = Co.nRepo108705_ImpCarteraReclasificacion(IIf(optPlanActual.value = True, "1", "2"), "dbCmactConsolidada..")
         If lsRtfImpG = "" Then
           MsgBox "No Existen Datos para este Reporte", vbInformation, "AVISO,"
         Else
           P.Show lsRtfImpG, "Reporte para Reclasificacion de Cartera (A-4)", True
         End If

Case 108706: ' Reporte de Intereses Devengados Vigentes (A-5)
         lsRtfImpG = ""
         lsRtfImpG = Co.nREpo108706_ImpRepDevengados_Vigentes("dbCmactConsolidada..", IIf(optPlanActual.value = True, "1", "2"), Val(TxtCambio))
         If lsRtfImpG = "" Then
           MsgBox "No Existen Datos para este Reporte", vbInformation, "AVISO,"
         Else
           P.Show lsRtfImpG, "Reporte de Intereses Devengados Vigentes (A-5)", True
         End If
Case 108707: ' Reporte de Intereses Devengados Vencidos y Cobranza Judicial (A-6)
         lsRtfImpG = ""
         lsRtfImpG = Co.nRepo108707_ImpRepDevengados_Vencidos("dbCmactConsolidada..", IIf(optPlanActual.value = True, "1", "2"), Val(TxtCambio))
         If lsRtfImpG = "" Then
           MsgBox "No Existen Datos para este Reporte", vbInformation, "AVISO,"
         Else
           P.Show lsRtfImpG, "Reporte de Intereses Devengados Vencidos y Cobranza Judicial (A-6)", True
         End If
Case 108708:  ' Resumen de Garantias  (A-7)
        lsRtfImpG = ""
        lsRtfImpG = Co.nRepo108708_ImpRepResumenGarantias("dbCmactConsolidada..", IIf(optPlanActual.value = True, "1", "2"), Val(TxtCambio))
         If lsRtfImpG = "" Then
           MsgBox "No Existen Datos para este Reporte", vbInformation, "AVISO,"
         Else
           P.Show lsRtfImpG, "Reporte de Intereses Devengados Vencidos y Cobranza Judicial (A-6)", True
         End If
Case 108709:  ' Cartera de Alto Riesgo  (A-8)

Case 108710: '  Colocaciones x Sectores Economicos  (A-9)
          lsRtfImpG = ""
          lsRtfImpG = Co.nRepo108710_ImpRepColocxSectEcon("dbCmactConsolidada..", IIf(optPlanActual.value = True, "1", "2"), Val(TxtCambio))
          If lsRtfImpG = "" Then
            MsgBox "No Existen Datos para este Reporte", vbInformation, "AVISO,"
          Else
            P.Show lsRtfImpG, "Reporte de Intereses Devengados Vencidos y Cobranza Judicial (A-6)", True
          End If

Case 108711: ' Reporte de Intereses de Créditos (A-4)
            lsRtfImpG = ""
            lsRtfImpG = Co.nRepo108711_ImpCarteraReclasificacion("dbCmactConsolidada..", IIf(optPlanActual.value = True, "1", "2"), "nIntDev")
            If lsRtfImpG = "" Then
                 MsgBox "No Existen Datos para este Reporte", vbInformation, "AVISO,"
            Else
                P.Show lsRtfImpG, "Reporte de Intereses de Créditos (A-4)", True
            End If
Case 108801:
            Fechaini = "01" & Mid(CStr(Co.GEtFechaCierreMes), 3, 10)
            lsRtfImpG = Co.nRepo108801_("dbCmactConsolidada..", Fechaini, Co.GEtFechaCierreMes, Val(TxtCambio), gsInstCmac)
            P.Show lsRtfImpG, "INFO1", True
            
Case 108802:
            If Not (IsNumeric(TxtCambio)) Then
                MsgBox "Ingrese Correctamente el Tipo de Cambio", vbInformation, "AVISO"
            Else
                lsRtfImpG = Co.nRepo108802_("dbCmactConsolidada..", Val(TxtCambio))
                P.Show lsRtfImpG, "INFO2", True
            End If
            
Case 108803:
            If Not (IsNumeric(TxtCambio)) Then
                MsgBox "Ingrese Correctamente el Tipo de Cambio", vbInformation, "AVISO"
            Else
                lsRtfImpG = Co.nRepo108803_("dbCmactConsolidada..", Val(TxtCambio), gsInstCmac)
                P.Show lsRtfImpG, "INFO3", True
            End If
            
Case 108804:
            If Not (IsNumeric(TxtCambio)) Then
                MsgBox "Ingrese Correctamente el Tipo de Cambio", vbInformation, "AVISO"
            Else
                lsRtfImpG = Co.nRepo108804_("dbCmactConsolidada..", Val(TxtCambio), gsInstCmac)
                P.Show lsRtfImpG, "INFO4", True
            End If
            
Case 108806:
            If Not (IsNumeric(TxtCambio)) Then
                MsgBox "Ingrese Correctamente el Tipo de Cambio", vbInformation, "AVISO"
            Else
                lsRtfImpG = Co.nRepo108806_("dbCmactConsolidada..", Val(TxtCambio))
                P.Show lsRtfImpG, "INFO6", True
            End If
            
Case 108808:
            If Not (IsNumeric(TxtCambio)) Then
                MsgBox "Ingrese Correctamente el Tipo de Cambio", vbInformation, "AVISO"
            Else
                Fechaini = "01" & Mid(CStr(Co.GEtFechaCierreMes), 3, 10)
                Call nRepo108808_("dbCmactConsolidada..", Fechaini, Co.GEtFechaCierreMes, Val(TxtCambio))
            End If
            'P.Show lsRtfImpG, "Informe Colocaciones para BCR", True
            
End Select
End Sub


Public Sub nRepo108808_(ByVal psServConsol As String, ByVal pdFechaDesde As Date, ByVal pdFechaHasta As Date, _
ByVal pnTipoCambio As Double)
Dim Co As nCredRepoFinMes
Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet
Dim xlHojaP As Excel.Worksheet
Dim lnFil As Integer, lnCol As Integer

Dim Sql As String
Dim rs As New ADODB.Recordset

Dim fs As Scripting.FileSystemObject

Dim Total As Double
Dim Tabula As Integer

Dim lsCond1(11) As String, lsCond2(11) As String
Dim Det As Integer


Dim lnNroCreFMS As Currency, lnNroCreFMD As Currency
Dim lnMonCreFMS As Currency, lnMonCreFMD As Currency
Dim lnNroCreOtorgS As Currency, lnNroCreOtorgD As Currency
Dim lnMonCreOtorgS As Currency, lnMonCreOtorgD As Currency
Dim lnNroCreCancelS As Currency, lnNroCreCancelD As Currency
Dim lnMonCreCancelS As Currency, lnMonCreCancelD As Currency
Dim lnNroCredS As Currency, lnNroCredD As Currency
Dim lnMonCredS As Currency, lnMonCredD As Currency

Dim Titulo As String
Dim lsCreditosVigentes As String
Dim lsPignoraticio As String
Dim lsVig As String
'Dim Tabula As Integer

Set Co = New nCredRepoFinMes
lsCreditosVigentes = gColocEstVigNorm & "," & gColocEstVigMor & "," & gColocEstVigVenc & "," & gColocEstRefNorm & "," & gColocEstRefMor & "," & gColocEstRefVenc
lsPignoraticio = gColPEstDesem & "," & gColPEstVenci & "," & gColPEstPRema & "," & gColPEstRenov
lsVig = gColocEstRecVigJud & "," & gColocEstRecVigCast & "," & gColocEstRecCanJud & "," & gColocEstRecCanCast
'On Error GoTo ErrorExcel
Screen.MousePointer = 11

Total = 4 * 25
'Me.barra.Max = Total
'rtf.Text = ""
Tabula = 20
ReDim Lineas(20)
Set fs = New Scripting.FileSystemObject

Set xlAplicacion = New Excel.Application
If fs.FileExists(App.path & "\SPOOLER\INFORME_COLOC_BCR.xls") Then
    Set xlLibro = xlAplicacion.Workbooks.Open(App.path & "\SPOOLER\INFO4.xls")
Else
    Set xlLibro = xlAplicacion.Workbooks.Add
End If
Set xlHoja1 = xlLibro.Worksheets.Add

'--************************** CLIENTES NUEVOS Y CONOCIDOS *****************************
'EncabezadoInfo4Excel
Titulo = " C R E D I T O   E M P R E S A R I A L"
xlAplicacion.Range("A1:E7").Font.Bold = True
xlAplicacion.Range("A1:E7").Font.Size = 9
xlAplicacion.Range("A5:P15").Font.Size = 8
xlAplicacion.Range("A4:E20").Font.Size = 8
xlAplicacion.Range("A7:E7").HorizontalAlignment = xlHAlignCenter
xlAplicacion.Range("A11:E11").HorizontalAlignment = xlHAlignCenter
xlAplicacion.Range("A11:E12").Font.Bold = True
xlHoja1.Cells(1, 3) = "R E P O R T E   I N F O 4"
xlHoja1.Cells(2, 2) = gsNomCmac
xlHoja1.Range("B2:E3").MergeCells = True
xlHoja1.Cells(3, 2) = "INFORMACION AL " & Format(gdFecSis, "dd/mm/yyyy")
xlHoja1.Cells(4, 2) = "T.C.F. :" & Format(pnTipoCambio, "#,#0.000")
xlHoja1.Cells(5, 3) = Titulo

'---------------------------------------
For Det = 1 To 11
    Select Case Det
        Case 1
            lsCond1(Det) = " AND Substring(C.cCtaCod,6,3) in('101','201') "
            lsCond2(Det) = " AND Substring(C.cLineacred,1,2) in('01') "
        Case 2
            lsCond1(Det) = " AND Substring(C.cCtaCod,6,3) in('101','201') "
            lsCond2(Det) = " AND Substring(C.cLineacred,1,2) not in('01') "
        Case 3
            lsCond1(Det) = " AND Substring(C.cCtaCod,6,3) in('301') "
            lsCond2(Det) = " AND Substring(C.cLineacred,1,2) in('01') "
        Case 4
            lsCond1(Det) = " AND Substring(C.cCtaCod,6,3) in('302','303') "
            lsCond2(Det) = " AND Substring(C.cLineacred,1,2) in('01') "
        Case 5
            lsCond1(Det) = " AND Substring(C.cCtaCod,6,3) in('304') "
            lsCond2(Det) = " AND Substring(C.cLineacred,1,2) in('01') "
        Case 6
            lsCond1(Det) = " AND Substring(C.cCtaCod,6,3) in('401','423') "
            lsCond2(Det) = " AND Substring(C.cLineacred,1,2) in('01') "
        Case 7
            lsCond1(Det) = " AND Substring(C.cCtaCod,6,3) in('301') "
            lsCond2(Det) = " AND Substring(C.cLineacred,1,2) in ('03','05','06','07','08') "
        Case 8
            lsCond1(Det) = " AND Substring(C.cCtaCod,6,3) in('302','303') "
            lsCond2(Det) = " AND Substring(C.cLineacred,1,2) in ('03','05','06','07','08') "
        Case 9
            lsCond1(Det) = " AND Substring(C.cCtaCod,6,3) in('304') "
            lsCond2(Det) = " AND Substring(C.cLineacred,1,2) in ('03','05','06','07','08') "
        Case 10
            lsCond1(Det) = " AND Substring(C.cCtaCod,6,3) in('401','423') "
            lsCond2(Det) = " AND Substring(C.cLineacred,1,2) in ('03','05','06','07','08') "
        Case 11
            lsCond1(Det) = " AND Substring(C.cCtaCod,6,3) in('305') "
            lsCond2(Det) = "  "
    End Select

    Sql = "SELECT Count( CASE WHEN SUBSTRING(C.cCtaCod,9,1)='1' THEN C.cCtaCod End ) NumOtorgS , " _
        & " Count( CASE WHEN SUBSTRING(C.cCtaCod,9,1)='2' THEN C.cCtaCod End ) NumOtorgD , " _
        & " Isnull(Sum ( CASE  WHEN SUBSTRING(C.cCtaCod,9,1)='1' THEN C.NMONTODESEMB End ),  0 ) SKOtorgS,  " _
        & " Isnull(Sum ( CASE  WHEN SUBSTRING(C.cCtaCod,9,1)='2' THEN C.NMONTODESEMB*" & pnTipoCambio & "  End ),  0 ) SKOtorgD  " _
        & " From " & psServConsol & "CreditoConsol  C " _
        & " WHERE C.nPrdEstado in (" & lsCreditosVigentes & "," & gColocEstCancelado & "," & lsVig & ") " _
        & "  AND  C.DFECVIG BETWEEN '" & Format(pdFechaDesde, "mm/dd/yyyy") & "' AND '" & Format(pdFechaHasta, "mm/dd/yyyy") & " 23:59' " _
        & lsCond1(Det) & lsCond2(Det)
    
    'rs.Open SQL, dbCmactCentral, adOpenStatic, adLockReadOnly, adCmdText
    Set rs = Co.GetQuery(Sql)
    lnNroCreOtorgS = rs!NumOtorgS
    lnNroCreOtorgD = rs!NumOtorgD
    lnMonCreOtorgS = rs!SKOtorgS
    lnMonCreOtorgS = rs!SKOtorgS
    
    rs.Close

    Sql = "SELECT Count( CASE WHEN SUBSTRING(C.cCtaCod,9,1)='1' THEN C.cCtaCod End ) NumFinMesS , " _
        & " Count( CASE WHEN SUBSTRING(C.cCtaCod,9,1)='2' THEN C.cCtaCod End ) NumFinMesD , " _
        & " Isnull(SUM( CASE WHEN substring(C.cCtaCod,9,1) = '1' THEN (C.nSaldoCap) End ), 0 ) SKFinMesS , " _
        & " Isnull(SUM( CASE WHEN substring(C.cCtaCod,9,1) = '2' THEN (C.nSaldoCap * " & pnTipoCambio & ") End ), 0 ) SKFinMesD " _
        & " From " & psServConsol & "CreditoSaldoConsol C " _
        & " JOIN " & psServConsol & "CreditoConsol CC on C.cCtaCod = CC.cCtaCod " _
        & " WHERE C.nPrdEstado in (" & lsCreditosVigentes & ") And Datediff(d,dFecha,'" & Format(pdFechaDesde, "mm/dd/yyyy") & "') = 0 " _
        & "  " _
        & lsCond1(Det) & Replace(lsCond2(Det), "C", "CC")

    'rs.Open SQL, dbCmactCentral, adOpenStatic, adLockReadOnly, adCmdText
    Set rs = Co.GetQuery(Sql)
    lnNroCreFMS = rs!NumFinMesS
    lnNroCreFMD = rs!NumFinMesD
    lnMonCreFMS = rs!SKFinMesS
    lnMonCreFMD = rs!SKFinMesD
    
    rs.Close

    Sql = "SELECT Count( CASE WHEN SUBSTRING(C.cCtaCod,9,1)='1' THEN C.cCtaCod End ) NumCredS ,  " _
        & " Count( CASE WHEN SUBSTRING(C.cCtaCod,9,1)='2' THEN C.cCtaCod End ) NumCredD , " _
        & " Isnull(SUM( CASE WHEN substring(C.cCtaCod,9,1) = '1' THEN (C.nSaldoCap) End ), 0 ) SKCredS, " _
        & " Isnull(SUM( CASE WHEN substring(C.cCtaCod,9,1) = '2' THEN (C.nSaldoCap * " & pnTipoCambio & " ) End ), 0 ) SKCredD " _
        & " From " & psServConsol & "CreditoSaldoConsol C " _
        & " JOIN " & psServConsol & "CreditoConsol CC on C.cCtaCod = CC.cCtaCod " _
        & " WHERE C.nPrdEstado in (" & lsCreditosVigentes & ") And Datediff(d,dFecha,'" & Format(pdFechaHasta, "mm/dd/yyyy") & "')=0" _
        & "  " _
        & lsCond1(Det) & Replace(lsCond2(Det), "C", "CC")
    
    'rs.Open SQL, dbCmactCentral, adOpenStatic, adLockReadOnly, adCmdText
    Set rs = Co.GetQuery(Sql)
    lnNroCredS = rs!NumCredS
    lnNroCredD = rs!NumCredD
    lnMonCredS = rs!SKCredS
    lnMonCredD = rs!SKCredD
    
    rs.Close

    lnNroCreCancelS = lnNroCreFMS + lnNroCreOtorgS - lnNroCredS
    lnNroCreCancelD = lnNroCreFMD + lnNroCreOtorgD - lnNroCredD
    lnMonCreCancelS = lnMonCreFMS + lnMonCreOtorgS - lnMonCredS
    lnMonCreCancelD = lnMonCreFMD + lnMonCreOtorgD - lnMonCredD
    
    If Det = 1 Or Det = 3 Or Det = 7 Or Det = 10 Or Det = 11 Then
        lnFil = lnFil + 3
        lnCol = 1
        
        xlHoja1.Cells(lnFil, lnCol) = "Nro Cred. Vigentes " & Format(pdFechaDesde, "dd/mm/yyyy")
        xlHoja1.Cells(lnFil + 1, lnCol) = "Nro Cred. Otorgados   " & Format(pdFechaHasta, "dd/mm/yyyy")
        xlHoja1.Cells(lnFil + 2, lnCol) = "Nro Cred. Cancelados  " & Format(pdFechaHasta, "dd/mm/yyyy")
        xlHoja1.Cells(lnFil + 3, lnCol) = "Nro Cred. Vigentes    " & Format(pdFechaHasta, "dd/mm/yyyy")
        xlHoja1.Cells(lnFil + 4, lnCol) = "Saldo Cred. Vigentes  " & Format(pdFechaDesde, "dd/mm/yyyy")
        xlHoja1.Cells(lnFil + 5, lnCol) = "Monto Cred. Otorgados " & Format(pdFechaHasta, "dd/mm/yyyy")
        xlHoja1.Cells(lnFil + 6, lnCol) = "Monto Cred. Cancelados" & Format(pdFechaHasta, "dd/mm/yyyy")
        xlHoja1.Cells(lnFil + 7, lnCol) = "Saldo Cred. Vigentes  " & Format(pdFechaHasta, "dd/mm/yyyy")
    End If
    
    xlHoja1.Cells(lnFil, lnCol + 1) = lnNroCreFMS
    xlHoja1.Cells(lnFil, lnCol + 2) = lnNroCreFMD
    xlHoja1.Cells(lnFil + 1, lnCol + 1) = lnNroCreOtorgS
    xlHoja1.Cells(lnFil + 1, lnCol + 2) = lnNroCreOtorgD
    xlHoja1.Cells(lnFil + 2, lnCol + 1) = lnNroCreCancelS
    xlHoja1.Cells(lnFil + 2, lnCol + 2) = lnNroCreCancelD
    xlHoja1.Cells(lnFil + 3, lnCol + 1) = lnNroCredS
    xlHoja1.Cells(lnFil + 3, lnCol + 2) = lnNroCredD
    xlHoja1.Cells(lnFil + 4, lnCol + 1) = lnMonCreFMS
    xlHoja1.Cells(lnFil + 4, lnCol + 2) = lnMonCreFMD
    xlHoja1.Cells(lnFil + 5, lnCol + 1) = lnMonCreOtorgS
    xlHoja1.Cells(lnFil + 5, lnCol + 2) = lnMonCreOtorgD
    xlHoja1.Cells(lnFil + 6, lnCol + 1) = lnMonCreCancelS
    xlHoja1.Cells(lnFil + 6, lnCol + 2) = lnMonCreCancelD
    xlHoja1.Cells(lnFil + 7, lnCol + 1) = lnMonCredS
    xlHoja1.Cells(lnFil + 7, lnCol + 2) = lnMonCredD

Next Det

xlHoja1.SaveAs App.path & "\SPOOLER\INFO4.xls"
'Cierra el libro de trabajo
xlLibro.Close
' Cierra Microsoft Excel con el método Quit.
xlAplicacion.Quit
'Libera los objetos.
Set xlAplicacion = Nothing
Set xlLibro = Nothing
Set xlHoja1 = Nothing
Screen.MousePointer = 0
MsgBox "Se ha Generado el Archivo INFO4.XLS Satisfactoriamente", vbInformation, "Aviso"
Exit Sub

ErrorExcel:
    MsgBox "Error Nº [" & Str(Err.Number) & "] " & Err.Description, vbInformation, "Aviso"
    xlLibro.Close
    ' Cierra Microsoft Excel con el método Quit.
    xlAplicacion.Quit
    'Libera los objetos.
    Set xlAplicacion = Nothing
    Set xlLibro = Nothing
    Set xlHoja1 = Nothing

End Sub


Private Sub cmdsalir_Click()
Unload Me
End Sub


Private Sub Form_Load()
TipoCambio = Co.TipoDiaCambio
Me.TxtCambio = TipoCambio
Set P = New Previo.clsPrevio
Set Co = New nCredRepoFinMes

CargaMenu
'-----------
Set Conecta = New DConecta
End Sub

Private Sub CargaMenu()
Dim clsGen As DGeneral
Dim rsUsu As Recordset
Dim sOperacion As String
Dim sOpeCod As String
Dim sOpePadre As String
Dim sOpeHijo As String
Dim sOpeHijito As String
Dim nodOpe As Node
Dim lsTipREP As String
lsTipREP = "1088"
Set clsGen = New DGeneral
Set rsUsu = clsGen.GetOperacionesUsuario(gsCodUser, lsTipREP, MatOperac, NroRegOpe)
Set clsGen = Nothing
Do While Not rsUsu.EOF
    sOpeCod = rsUsu("cOpeCod")
    sOperacion = sOpeCod & " - " & UCase(rsUsu("cOpeDesc"))
    Select Case rsUsu("nOpeNiv")
        Case "1"
            sOpePadre = "P" & sOpeCod
         Set nodOpe = tvwReporte.Nodes.Add(, , sOpePadre, sOperacion, "Padre")
            nodOpe.Tag = sOpeCod
        Case "2"
            sOpeHijo = "H" & sOpeCod
            Set nodOpe = tvwReporte.Nodes.Add(sOpePadre, tvwChild, sOpeHijo, sOperacion, "Hijo")
            nodOpe.Tag = sOpeCod
    End Select
    rsUsu.MoveNext
Loop
rsUsu.Close
Set rsUsu = Nothing
End Sub
Function ImpCarteraCredxMoneda(pPlanCuentas As String) As Boolean
Dim SQL9 As String
Dim SQL9B As String
Dim SQL9C As String
Dim rsM As New ADODB.Recordset
Dim rsPla As New ADODB.Recordset
Dim rsFF As New ADODB.Recordset
Dim rsAg As New ADODB.Recordset
Dim Reg9 As New ADODB.Recordset
Dim lsDescM As String
Dim lsDescPr As String
Dim lsDescP As String
Dim lsDescF As String
Dim lsDescAG As String
Dim i As Integer
Dim liLineas As Integer
Dim liDatos  As Long

Dim lnTCred As Single  ' *** General
Dim lnTSaldoCap As Double
Dim lnTNumVig As Single
Dim lnTCapVig As Double
Dim lnTNumVen1 As Integer
Dim lnTCapVen1  As Double
Dim lnTNumVen2 As Integer
Dim lnTCapVen2 As Double
Dim lnTNumJud As Integer
Dim lnTCapJud As Double

Dim lnTCredM As Single  ' *** Moneda
Dim lnTSaldoCapM As Double
Dim lnTNumVigM As Single
Dim lnTCapVigM As Double
Dim lnTNumVen1M As Integer
Dim lnTCapVen1M  As Double
Dim lnTNumVen2M As Integer
Dim lnTCapVen2M As Double
Dim lnTNumJudM As Integer
Dim lnTCapJudM As Double

Dim lnTCredF As Single  ' *** Fuente Financiamiento
Dim lnTSaldoCapF As Double
Dim lnTNumVigF As Single
Dim lnTCapVigF As Double
Dim lnTNumVen1F As Integer
Dim lnTCapVen1F  As Double
Dim lnTNumVen2F As Integer
Dim lnTCapVen2F As Double
Dim lnTNumJudF As Integer
Dim lnTCapJudF As Double

Dim lnTCredAG As Single  ' *** Agencia
Dim lnTSaldoCapAG As Double
Dim lnTNumVigAG As Single
Dim lnTCapVigAG As Double
Dim lnTNumVen1AG As Integer
Dim lnTCapVen1AG  As Double
Dim lnTNumVen2AG As Integer
Dim lnTCapVen2AG As Double
Dim lnTNumJudAG As Integer
Dim lnTCapJudAG As Double

Dim lnTCredPr As Single  ' *** Producto
Dim lnTSaldoCapPr As Double
Dim lnTNumVigPr As Single
Dim lnTCapVigPr As Double
Dim lnTNumVen1Pr As Integer
Dim lnTCapVen1Pr  As Double
Dim lnTNumVen2Pr As Integer
Dim lnTCapVen2Pr As Double
Dim lnTNumJudPr As Integer
Dim lnTCapJudPr As Double

Dim lbSiCambio As Boolean
Dim Cadena As String

Dim lnCapVencido As Currency  'Para los consumo

Dim Prod(7) As String
Dim lnRangoIni As Integer
Dim lnRangoFin As Integer
Dim lsRTfImp As String
Dim cpag As Integer

Dim lnCredAdmin As Integer ' Bandera Para Cred Admin

' Para Cada Producto
Prod(0) = " '101' "  ' Comerciales
Prod(1) = " '201' "  ' Pyme
Prod(2) = " '301','302','303','304'"  ' Personales
Prod(3) = " '102','202' "  ' Agricola
Prod(4) = " '305'" ' Pignoraticio
Prod(5) = " '320'" ' Administrativos
Prod(6) = " '401'" ' Hipotecaja
Prod(7) = " '423'" ' Mivivienda

lbSiCambio = False
lsRTfImp = ""
cpag = 1

Cadena = ""
'Cadena = ValorMoneda & ValorNorRefPar & ValorProducto
 
'===== Verificar si existen datos para el reporte                                                'CreditoConsol.nEstado='F'
SQL9 = "SELECT COUNT(cCodCta) as NumCred from dbCmactConsolidada..CreditoConsol where nEstado in ('2020','2021','2022','2030','2031','2032') "

Conecta.AbreConexion
Set Reg9 = Conecta.CargaRecordSet(SQL9)
Conecta.CierraConexion
'Reg9.Open SQL9, dbCmactCentral, adOpenStatic, adLockReadOnly, adCmdText

If Reg9.EOF And Reg9.BOF Then
    liDatos = 0
Else
    liDatos = Reg9!NumCred
End If
Reg9.Close
Set Reg9 = Nothing

liDatos = liDatos + Round(liDatos / 10, 0)

If liDatos > 0 Then  ' General
    '======== Tipo de Moneda =========
    Set rsM = Co.ObtieneTablaCod(1011) ' Moneda
   '========= Fuente Financiamiento =========
    Set rsFF = Co.ObtieneLineaCredito 'Fuente Financiamiento
   
   '======= Para Cada Agencia  =========
    SQL9 = " select cAgeDescripcion cNombTab, cValor='112'+cAgeCod from DBCmactAux..Agencias"
           '" WHERE cCodTab like '47%' AND cValor like '112%' ORDER BY cValor "
    'rsAg.Open SQL9, dbCmactCentral, adOpenStatic, adLockReadOnly, adCmdText
    Conecta.AbreConexion
    Set rsAg = Conecta.CargaRecordSet(SQL9)
    Conecta.CierraConexion
    
    ImpCarteraCredxMoneda = True
    'PrgBarra.Min = 0
    'PrgBarra.Max = liDatos
    'StBarra.Panels(1).Text = "Procesando Reporte, Por Favor Espere..."
    If rsM.EOF And rsM.BOF Then
    Else
        Do While Not rsM.EOF
            '=====LIMPIAR LOS DATOS SUBTOTALES POR MONEDA ======
            lnTCredM = 0:   lnTSaldoCapM = 0
            lnTNumVigM = 0: lnTCapVigM = 0
            lnTNumVen1M = 0: lnTCapVen1M = 0
            lnTNumVen2M = 0: lnTCapVen2M = 0
            lnTNumJudM = 0: lnTCapJudM = 0
            lsDescM = Trim(rsM!cNomTab)
            '============ CREDITOS VIGENTES PARA UNA DETERMINADA MONEDA =========
            SQL9 = "SELECT Count(cCodCta) AS NumCred FROM CreditoConsol " & _
                "WHERE CreditoConsol.nEstado in  ('2020','2021','2022','2030','2031','2032') and " & _
                "substring(cCodCta,9,1 )= '" & Trim(rsM!cValor) & "'"
            
            'Reg9.Open SQL9, dbCmactCentral, adOpenStatic, adLockReadOnly, adCmdText
            Conecta.AbreConexion
            Set Reg9 = Conecta.CargaRecordSet(SQL9)
            Conecta.CierraConexion
            
            If Reg9.EOF And Reg9.BOF Then
                liDatos = 0
            Else
                liDatos = Reg9!NumCred
            End If
            Reg9.Close
            Set Reg9 = Nothing
            
            If liDatos > 0 Then  ' De la Moneda
                lsRTfImp = lsRTfImp & Co.CabCarteraCredxMoneda(cpag, pPlanCuentas)
                liLineas = 5
                If lbSiCambio Then
                    lbSiCambio = False
                End If
                lsRTfImp = lsRTfImp & Chr(27) & Chr(69)
                lsRTfImp = lsRTfImp & " MONEDA      : " & lsDescM & Chr(27) & Chr(70) & Chr(10)
                liLineas = liLineas + 1
                If rsFF.EOF And rsFF.BOF Then 'Fuente Financ
                Else
                    rsFF.MoveFirst
                    Do While Not rsFF.EOF 'Fuente Financ
                       lnTCredF = 0: lnTSaldoCapF = 0
                       lnTNumVigF = 0: lnTCapVigF = 0
                       lnTNumVen1F = 0: lnTCapVen1F = 0
                       lnTNumVen2F = 0: lnTCapVen2F = 0
                       lnTNumJudF = 0: lnTCapJudF = 0
                       lsDescF = Mid(rsFF!cNomTab, 1, 25)
                       '=========== CREDITOS VIGENTES PARA UNA FUENTE FINANCIAMIENTO =============
                       SQL9 = " SELECT count(cCodCta) as NumCred from dbcmactconsolidada..CreditoConsol " & _
                              " WHERE CreditoConsol.nEstado in ('2020','2021','2022','2030','2031','2032')  " & _
                              " AND substring(cCodCta,9,1 )= '" & Trim(rsM!cValor) & "'" & _
                              " AND substring(cCodLinCred, 9,1) IN ('" & Trim(rsFF!cValor) & "')"
                       SQL9B = "SELECT Count(cCodCta) as NumCred FROM dbcmactconsolidada..CreditoConsol "
                       SQL9B = SQL9B & "WHERE nEstado =" & gColocEstRecVigJud & " AND cCondCre =" & gColocEstJudicial & " AND nSaldoCap > 0 "
                       SQL9B = SQL9B & "AND substring(cCodLinCred,4,1 )= '" & Trim(rsM!cValor) & "'"
                       SQL9B = SQL9B & "AND substring(cCodLinCred,6,1) IN ('" & Trim(rsFF!cValor) & "')"


                        If Co.HayDatos(SQL9) + Co.HayDatos(SQL9B) > 0 Then  ' De la Fuente Financ
                            
                          ' ****************************
                          ' ****  Verifica Si Cambio pagina
                          If liLineas >= 55 Then 'Para el control de salto de página
                             lsRTfImp = lsRTfImp & Chr(12)
                             cpag = cpag + 1
                             lsRTfImp = lsRTfImp & Co.CabCarteraCredxMoneda(cpag, pPlanCuentas)
                             liLineas = 5
                          End If
                            
                          lsRTfImp = lsRTfImp & Chr(27) & Chr(69)
                          lsRTfImp = lsRTfImp & " * FUENTE FINANCIAMIENTO   : " & lsDescF & Chr(27) & Chr(70) & Chr(10)
                          liLineas = liLineas + 1
                            '======= Para Cada Agencia  =========
                          rsAg.MoveFirst
                          Do While Not rsAg.EOF
                               
                            SQL9 = "SELECT count(cCodCta) as NumCred FROM dbcmactconsolidada..CreditoConsol " & _
                                    " where CreditoConsol.nEstado in ('2020','2021','2022','2030','2031','2032') " & _
                                    " AND substring(cCodLinCred,4,1 )= '" & Trim(rsM!cValor) & "'" & _
                                    " AND substring(cCodLinCred,6,1) = '" & Trim(rsFF!cValor) & "'" & _
                                    " AND substring(cCodCta,4,2) = '" & Right(Trim(rsAg!cValor), 2) & "'"
                            If Trim(rsFF!cValor) = "1" Then
                               SQL9B = "SELECT count(cCodCta) as NumCred FROM dbcmactconsolidada..CreditoConsol " & _
                                       "WHERE nEstado IN (gColPEstDesem,gColPEstVenci,gColPEstPRema,gColPEstRenov) " & _
                                       "AND substring(cCodCta,9,1 )= '" & Trim(rsM!cValor) & "'" & _
                                       "AND substring(cCodCta,4,2) = '" & Right(Trim(rsAg!cValor), 2) & "'"
                            Else
                               SQL9B = "SELECT count(cCodCta) as NumCred FROM dbcmactconsolidada..CreditoConsol "
                               SQL9B = SQL9B & "WHERE nESTADO = 'Y' "
                               '             -----------------> ver
                            End If
                            SQL9C = "SELECT Count(cCodCta) as NumCred FROM dbcmactconsolidada..CreditoConsol "
                            SQL9C = SQL9C & "WHERE cEstado =" & gColocEstRecVigJud & " AND cCondCre =" & gColocEstJudicial & " AND nSaldoCap > 0 "
                            SQL9C = SQL9C & "AND substring(cCodLinCred,4,1 )= '" & Trim(rsM!cValor) & "'"
                            SQL9C = SQL9C & "AND substring(cCodLinCred,6,1) = '" & Trim(rsFF!cValor) & "'"
                            SQL9C = SQL9C & "AND substring(cCodCta,4,2) = '" & Right(Trim(rsAg!cValor), 2) & "'"
                            '******* Administrativos
                            If Trim(rsFF!cValor) = "1" And Right(Trim(rsAg!cValor), 2) = "07" And i = "5" Then
                               lnCredAdmin = 1
                            Else
                               lnCredAdmin = 0
                            End If
                            
                            If Co.HayDatos(SQL9) + Co.HayDatos(SQL9B) + Co.HayDatos(SQL9C) + lnCredAdmin > 0 Then    ' De la Agencia

                               lnTCredAG = 0: lnTSaldoCapAG = 0
                               lnTNumVigAG = 0: lnTCapVigAG = 0
                               lnTNumVen1AG = 0: lnTCapVen1AG = 0
                               lnTNumVen2AG = 0: lnTCapVen2AG = 0
                               lnTNumJudAG = 0: lnTCapJudAG = 0
                               lsDescAG = Mid(rsAg!cNomTab, 1, 20)
                               ' ****************************
                               ' ****  Verifica Si Cambio pagina
                               If liLineas >= 55 Then 'Para el control de salto de página
                                  lsRTfImp = lsRTfImp & Chr(12)
                                  cpag = cpag + 1
                                  lsRTfImp = lsRTfImp & Co.CabCarteraCredxMoneda(cpag, pPlanCuentas)
                                  liLineas = 5
                               End If
                               lsRTfImp = lsRTfImp & Chr(27) & Chr(69)
                               lsRTfImp = lsRTfImp & " ** " & Mid(lsDescAG, 1, 25) & Chr(27) & Chr(70) & Chr(10)
                               liLineas = liLineas + 1
                               
                               For i = 0 To 7  'Para Cada Producto
                                  lnTCredPr = 0:   lnTSaldoCapPr = 0
                                  lnTNumVigPr = 0: lnTCapVigPr = 0
                                  lnTNumVen1Pr = 0: lnTCapVen1Pr = 0
                                  lnTNumVen2Pr = 0: lnTCapVen2Pr = 0
                                  lnTNumJudPr = 0: lnTCapJudPr = 0
                                  Select Case i
                                     Case 0
                                        lsDescPr = " Comercial    "
                                     Case 1
                                        lsDescPr = " Pyme         "
                                     Case 2
                                        lsDescPr = " Personal     "
                                     Case 3
                                        lsDescPr = " Agricola     "
                                     Case 4
                                        lsDescPr = " Pignoraticio "
                                     Case 5
                                        lsDescPr = " Administrat  "
                                     Case 6
                                        lsDescPr = " Hipotecaja   "
                                     Case 7
                                        lsDescPr = " MiVivienda   "
                                  End Select
                                  
                                  '=========== CREDITOS DEL PRODUCTO =============
                                    SQL9 = "SELECT count(cCodCta) as NumCred FROM dbcmactconsolidada..CreditoConsol " & _
                                        " where CreditoConsol.nEstado in ('2020','2021','2022','2030','2031','2032') " & _
                                        " AND substring(cCodLinCred,4,1 )= '" & Trim(rsM!cValor) & "'" & _
                                        " AND substring(cCodLinCred,6,1) = '" & Trim(rsFF!cValor) & "'" & _
                                        " AND substring(cCodCta,4,2) = '" & Right(Trim(rsAg!cValor), 2) & "'" & _
                                        " AND substring(cCodCta,6,3) in ( " & Prod(i) & ")"
                                     
                                    If Trim(rsFF!cValor) = "1" And Trim(rsM!cValor) = "1" Then  ' Si es RR PP
                                       SQL9B = "SELECT count(cCodCta) as NumCred FROM dbcmactconsolidada..CreditoConsol " & _
                                            " WHERE cEstado IN ('2101','2104','2106','2107') " & _
                                            " AND substring(cCodCta,9,1 )= '" & Trim(rsM!cValor) & "'" & _
                                            " AND substring(cCodCta,4,2) = '" & Right(Trim(rsAg!cValor), 2) & "'" & _
                                            " AND substring(cCodCta,6,3) in ( " & Prod(i) & ")"
                                    Else
                                       SQL9B = "SELECT count(cCodCta) as NumCred FROM CreditoConsol " & _
                                               "WHERE nESTADO = 'Y' "
                                    End If
                                    ' Judicial
                                    SQL9C = "SELECT COUNT(cCodCta) AS NumCred " & _
                                           " FROM dbcmactconsolidada..CreditoConsol " & _
                                           " WHERE nEstado =" & gColocEstRecVigJud & " AND cCondCre =" & gColocEstJudicial & "  AND nSaldoCap > 0" & _
                                           " AND substring(cCodLinCred,4,1 )= '" & Trim(rsM!cValor) & "'" & _
                                           " AND substring(cCodLinCred,6,1) = '" & Trim(rsFF!cValor) & "'" & _
                                           " AND substring(cCodCta,4,2) = '" & Right(Trim(rsAg!cValor), 2) & "'" & _
                                           " AND substring(cCodCta,6,3) in (" & Prod(i) & ")"
                                    
                                    '******* Administrativos
                                    If Trim(rsFF!cValor) = "1" And Right(Trim(rsAg!cValor), 2) = "07" And i = "5" Then
                                       lnCredAdmin = 1
                                    Else
                                       lnCredAdmin = 0
                                    End If
                                    
                                    If Co.HayDatos(SQL9) + Co.HayDatos(SQL9B) + Co.HayDatos(SQL9C) + lnCredAdmin > 0 Then ' Para Producto
                                       '=================================================
                                       '===== Total Cartera  Producto de Agencia  ======
                                       '  Creditos ------------
                                       SQL9 = "SELECT COUNT(cCodCta) AS Numero, SUM(nSaldoCap) AS Capital " & _
                                              " FROM dbcmactconsolidada..CreditoConsol " & _
                                              " WHERE cEstado in ('2020','2021','2022','2030','2031','2032') " & _
                                              " AND substring(cCodLinCred,4,1 )= '" & Trim(rsM!cValor) & "'" & _
                                              " AND substring(cCodLinCred,6,1) = '" & Trim(rsFF!cValor) & "'" & _
                                              " AND substring(cCodCta,4,2) = '" & Right(Trim(rsAg!cValor), 2) & "'" & _
                                              " AND substring(cCodCta,6,3) in ( " & Prod(i) & ")"
                                              
                                       'Reg9.Open SQL9, dbCmactCentral, adOpenStatic, adLockReadOnly, adCmdText
                                       Conecta.AbreConexion
                                       Set Reg9 = Conecta.CargaRecordSet(SQL9)
                                       Conecta.CierraConexion
                                       
                                       lnTCredPr = Reg9!Numero
                                       lnTSaldoCapPr = IIf(IsNull(Reg9!CAPITAL), 0, Reg9!CAPITAL)
                                       Reg9.Close
                                       Set Reg9 = Nothing
                                                
                                       ' Pignoraticio  ------------
                                       If Trim(rsFF!cValor) = "1" And Trim(rsM!cValor) = "1" Then  ' Si es RR PP
                                          SQL9 = "SELECT count(cCodCta) as Numero, Sum(nSaldoCap) AS Capital " & _
                                              " FROM dbcmactconsolidada..CreditoConsol WHERE cEstado IN ('2101','2104','2106','2107') " & _
                                              " AND substring(cCodCta,9,1 )= '" & Trim(rsM!cValor) & "'" & _
                                              " AND substring(cCodCta,4,2) = '" & Right(Trim(rsAg!cValor), 2) & "'" & _
                                              " AND substring(cCodCta,6,3) in ( " & Prod(i) & ")"
                                          
                                          Conecta.AbreConexion
                                          Set Reg9 = Conecta.CargaRecordSet(SQL9)
                                          Conecta.CierraConexion
                                          
                                          'Reg9.Open SQL9, dbCmactCentral, adOpenStatic, adLockReadOnly, adCmdText
                                          lnTCredPr = lnTCredPr + Reg9!Numero
                                          lnTSaldoCapPr = lnTSaldoCapPr + IIf(IsNull(Reg9!CAPITAL), 0, Reg9!CAPITAL)
                                          Reg9.Close
                                          Set Reg9 = Nothing
                                       End If
                                       ' Judicial -------------
                                       SQL9 = "SELECT COUNT(cCodCta) AS Numero, " & _
                                              " SUM(nSaldoCap) AS Capital " & _
                                              " FROM dbcmactconsolidada..CreditoConsol " & _
                                              " WHERE cEstado =" & gColocEstRecVigJud & " AND cCondCre =" & gColocEstJudicial & " AND nSaldoCap > 0" & _
                                              " AND substring(cCodLinCred,4,1 )= '" & Trim(rsM!cValor) & "'" & _
                                              " AND substring(cCodLinCred,6,1) = '" & Trim(rsFF!cValor) & "'" & _
                                              " AND substring(cCodCta,4,2) = '" & Right(Trim(rsAg!cValor), 2) & "'" & _
                                              " AND substring(cCodCta,6,3) in ( " & Prod(i) & ")"
                                       
                                       'Reg9.Open SQL9, dbCmactCentral, adOpenStatic, adLockReadOnly, adCmdText
                                         Conecta.AbreConexion
                                         Set Reg9 = Conecta.CargaRecordSet(SQL9)
                                         Conecta.CierraConexion
                                       
                                       lnTCredPr = lnTCredPr + Reg9!Numero
                                       lnTSaldoCapPr = lnTSaldoCapPr + IIf(IsNull(Reg9!CAPITAL), 0, Reg9!CAPITAL)
                                       Reg9.Close
                                       Set Reg9 = Nothing
                                       '=================================================
                                       '====== Cartera Vigente  Producto de Agencia  ===
                                       Select Case i
                                       
                                           Case 0  ' Comercial  ******************
                                             If pPlanCuentas = "1" Then
                                                lnRangoIni = -9999   ' Comerciales
                                                lnRangoFin = 15
                                             Else
                                                lnRangoIni = -9999   ' Comerciales
                                                lnRangoFin = 15
                                             End If

                                             SQL9 = "SELECT COUNT(cCodCta) AS Numero, " & _
                                                   " SUM(nSaldoCap) AS Capital " & _
                                                   " FROM dbcmactconsolidada..CreditoConsol " & _
                                                   " WHERE CreditoConsol.nEstado in ('2020','2021','2022','2030','2031','2032')" & _
                                                   " AND substring(cCodLinCred,4,1 )= '" & Trim(rsM!cValor) & "'" & _
                                                   " AND substring(cCodLinCred,6,1) = '" & Trim(rsFF!cValor) & "'" & _
                                                   " AND substring(cCodCta,4,2) = '" & Right(Trim(rsAg!cValor), 2) & "'" & _
                                                   " AND substring(cCodCta,6,3) in (" & Prod(i) & ") " & _
                                                   " And CreditoConsol.ndiasAtraso >= " & Val(lnRangoIni) & _
                                                   " And CreditoConsol.ndiasatraso <= " & Val(lnRangoFin)
                                                   
                                             'Reg9.Open SQL9, dbCmactCentral, adOpenStatic, adLockReadOnly, adCmdText
                                             Conecta.AbreConexion
                                             Set Reg9 = Conecta.CargaRecordSet(SQL9)
                                             Conecta.CierraConexion
                                             lnTNumVigPr = Reg9!Numero
                                             lnTCapVigPr = IIf(IsNull(Reg9!CAPITAL), 0, Reg9!CAPITAL)
                                             Reg9.Close
                                             Set Reg9 = Nothing
                                             
                                             '********
                                       
                                           Case 1  ' Pyme  ******************
                                             If pPlanCuentas = "1" Then
                                                 lnRangoIni = -9999   ' MicroEmpresa
                                                 lnRangoFin = 30
                                             Else
                                                 lnRangoIni = -9999   ' MicroEmpresa
                                                 lnRangoFin = 30
                                             End If
                                             
                                             SQL9 = "SELECT COUNT(dbcmactconsolidada..CreditoConsol.cCodCta) AS Numero, " & _
                                                   " SUM(dbcmactconsolidada..CreditoConsol.nSaldoCap) AS Capital " & _
                                                   " FROM dbcmactconsolidada..CreditoConsol " & _
                                                   " WHERE dbcmactconsolidada..CreditoConsol.nEstado in ('2020','2021','2022','2030','2031','2032') " & _
                                                   " AND substring(cCodLinCred,4,1 )= '" & Trim(rsM!cValor) & "'" & _
                                                   " AND substring(cCodLinCred,6,1) = '" & Trim(rsFF!cValor) & "'" & _
                                                   " AND substring(cCodCta,4,2) = '" & Right(Trim(rsAg!cValor), 2) & "'" & _
                                                   " AND substring(cCodCta,6,3) in (" & Prod(i) & ")" & _
                                                   " And dbcmactconsolidada..CreditoConsol.ndiasAtraso >= " & Val(lnRangoIni) & _
                                                   " And dbcmactconsolidada..CreditoConsol.ndiasatraso <= " & Val(lnRangoFin)
                                                   
                                             'Reg9.Open SQL9, dbCmactCentral, adOpenStatic, adLockReadOnly, adCmdText
                                             Conecta.AbreConexion
                                             Set Reg9 = Conecta.CargaRecordSet(SQL9)
                                             Conecta.CierraConexion
                                             
                                             lnTNumVigPr = Reg9!Numero
                                             lnTCapVigPr = IIf(IsNull(Reg9!CAPITAL), 0, Reg9!CAPITAL)
                                             Reg9.Close
                                             Set Reg9 = Nothing
                                                   
                                          Case 2, 6, 7 ' Consumo  / Hipotecaja / MiVivienda ***************
                                          
                                             If pPlanCuentas = "1" Then
                                                 lnRangoIni = -9999
                                                 lnRangoFin = 30
                                             Else
                                                 lnRangoIni = -9999
                                                 lnRangoFin = 30
                                             End If
                                          
                                             SQL9 = "SELECT COUNT(dbcmactconsolidada..CreditoConsol.cCodCta) AS Numero, " & _
                                                   " SUM(dbcmactconsolidada..CreditoConsol.nSaldoCap) AS Capital " & _
                                                   " FROM dbcmactconsolidada..CreditoConsol " & _
                                                   " WHERE CreditoConsol.nEstado in ('2020','2021','2022','2030','2031','2032') " & _
                                                   " AND substring(cCodLinCred,4,1 )= '" & Trim(rsM!cValor) & "'" & _
                                                   " AND substring(cCodLinCred,6,1) = '" & Trim(rsFF!cValor) & "'" & _
                                                   " AND substring(cCodCta,1,2) = '" & Right(Trim(rsAg!cValor), 2) & "'" & _
                                                   " AND substring(cCodCta,3,3) in ( " & Prod(i) & ")" & _
                                                   " And dbcmactconsolidada..CreditoConsol.ndiasAtraso >= " & Val(lnRangoIni) & _
                                                   " And dbcmactconsolidada..CreditoConsol.ndiasatraso <= " & Val(lnRangoFin)
                                                   
                                             'Reg9.Open SQL9, dbCmactCentral, adOpenStatic, adLockReadOnly, adCmdText
                                             Conecta.AbreConexion
                                             Set Reg9 = Conecta.CargaRecordSet(SQL9)
                                             Conecta.CierraConexion
                                             
                                             lnTNumVigPr = Reg9!Numero
                                             lnTCapVigPr = IIf(IsNull(Reg9!CAPITAL), 0, Reg9!CAPITAL)
                                             Reg9.Close
                                             Set Reg9 = Nothing
                                                   
                                          Case 3  ' Agricola  *************************
                                             
                                             If pPlanCuentas = "1" Then
                                                 lnRangoIni = -9999
                                                 lnRangoFin = 30
                                             Else
                                                 lnRangoIni = -9999
                                                 lnRangoFin = 30
                                             End If
                                             SQL9 = "SELECT COUNT(CreditoConsol.cCodCta) AS Numero, " & _
                                                   " SUM(CreditoConsol.nSaldoCap) AS Capital " & _
                                                   " FROM dbcmactconsolidada..CreditoConsol " & _
                                                   " WHERE dbcmactconsolidada..CreditoConsol.nEstado in ('2020','2021','2022','2030','2031','2032') " & _
                                                   " AND substring(cCodLinCred,4,1 )= '" & Trim(rsM!cValor) & "'" & _
                                                   " AND substring(cCodLinCred,6,1) = '" & Trim(rsFF!cValor) & "'" & _
                                                   " AND substring(cCodCta,4,2) = '" & Right(Trim(rsAg!cValor), 2) & "'" & _
                                                   " AND substring(cCodCta,6,3) in ( " & Prod(i) & ")" & _
                                                   " And dbcmactconsolidada..CreditoConsol.ndiasAtraso >= " & Val(lnRangoIni) & _
                                                   " And dbcmactconsolidada..CreditoConsol.ndiasatraso <= " & Val(lnRangoFin)
                                            
                                            'Reg9.Open SQL9, dbCmactCentral, adOpenStatic, adLockReadOnly, adCmdText
                                            
                                            lnTNumVigPr = Reg9!Numero
                                            lnTCapVigPr = IIf(IsNull(Reg9!CAPITAL), 0, Reg9!CAPITAL)
                                            Reg9.Close
                                            Set Reg9 = Nothing
                                                   
                                          Case 4  ' Prendario  *************************
                                             If Trim(rsFF!cValor) = "1" Then  ' Si es RR PP
                                                lnRangoIni = -999
                                                lnRangoFin = 30
                                                SQL9 = "SELECT COUNT(cCodCta) AS Numero, " & _
                                                    " SUM(nSaldoCap) AS Capital " & _
                                                   " FROM dbcmactconsolidada..CreditoConsol " & _
                                                   " WHERE nEstado in ('2101','2104','2106','2107') " & _
                                                   " AND substring(cCodCta,4,2) = '" & Right(Trim(rsAg!cValor), 2) & "'" & _
                                                   " AND substring(cCodCta,6,3) in ( " & Prod(i) & ")" & _
                                                   " And DATEDIFF(dd ,dFecVenc,'" & Format(Co.GEtFechaCierreMes, "mm/dd/yyyy") & "') >= " & Val(lnRangoIni) & _
                                                   " And DATEDIFF(dd ,dFecVenc,'" & Format(Co.GEtFechaCierreMes, "mm/dd/yyyy") & "') <= " & Val(lnRangoFin)
                                                   
                                                 'Reg9.Open SQL9, dbCmactCentral, adOpenStatic, adLockReadOnly, adCmdText
                                                 Conecta.AbreConexion
                                                 Set Reg9 = Conecta.CargaRecordSet(SQL9)
                                                 Conecta.CierraConexion
                                                 
                                                 lnTNumVigPr = Reg9!Numero
                                                 lnTCapVigPr = IIf(IsNull(Reg9!CAPITAL), 0, Reg9!CAPITAL)
                                                 Reg9.Close
                                                 Set Reg9 = Nothing
                                             End If
                                          
                                          Case 5  ' Administrativos  *************************
                                             '*****  Cred Administrativos
                                             
                                             If pPlanCuentas = "1" Then
                                                 lnRangoIni = -9999
                                                 lnRangoFin = 30
                                             Else
                                                 lnRangoIni = -9999
                                                 lnRangoFin = 30
                                             End If
                                          
                                             SQL9 = "SELECT COUNT(dbcmactconsolidada..CreditoConsol.cCodCta) AS Numero, " & _
                                                   " SUM(dbcmactconsolidada..CreditoConsol.nSaldoCap) AS Capital " & _
                                                   " FROM dbcmactconsolidada..CreditoConsol " & _
                                                   " WHERE dbcmactconsolidada..CreditoConsol.nEstado in ('2020','2021','2022','2030','2031','2032') " & _
                                                   " AND substring(cCodLinCred,4,1 )= '" & Trim(rsM!cValor) & "'" & _
                                                   " AND substring(cCodLinCred,6,1) = '" & Trim(rsFF!cValor) & "'" & _
                                                   " AND substring(cCodCta,4,2) = '" & Right(Trim(rsAg!cValor), 2) & "'" & _
                                                   " AND substring(cCodCta,6,3) in ( " & Prod(i) & ")" & _
                                                   " And dbcmactconsolidada..CreditoConsol.ndiasAtraso >= " & Val(lnRangoIni) & _
                                                   " And dbcmactconsolidada..CreditoConsol.ndiasatraso <= " & Val(lnRangoFin)
                                                   
                                             'Reg9.Open SQL9, dbCmactCentral, adOpenStatic, adLockReadOnly, adCmdText
                                             Conecta.AbreConexion
                                             Set Reg9 = Conecta.CargaRecordSet(SQL9)
                                             Conecta.CierraConexion
                                             
                                             lnTNumVigPr = Reg9!Numero
                                             lnTCapVigPr = IIf(IsNull(Reg9!CAPITAL), 0, Reg9!CAPITAL)
                                             Reg9.Close
                                             Set Reg9 = Nothing
                                             
                                           Case Else
                                               lnTNumVigPr = 0
                                               lnTCapVigPr = 0
                                           
                                       End Select
                                       '=================================================
                                       '====== Cartera Vencida 1 Producto de Agencia  ===
                                       Select Case i
                                           
                                           Case 0  ' Comerciales  ******************
                                           
                                             If pPlanCuentas = "1" Then
                                                 lnRangoIni = 16   ' Comerciales
                                                 lnRangoFin = 120
                                             Else
                                                 lnRangoIni = 16   ' Comerciales
                                                 lnRangoFin = 9999
                                             End If
                                           
                                             SQL9 = "SELECT COUNT(dbcmactconsolidada..CreditoConsol.cCodCta) AS Numero, " & _
                                                   " SUM(dbcmactconsolidada..CreditoConsol.nSaldoCap) AS Capital " & _
                                                   " FROM dbcmactconsolidada..CreditoConsol " & _
                                                   " WHERE dbcmactconsolidada..CreditoConsol.nEstado in ('2020','2021','2022','2030','2031','2032') " & _
                                                   " AND substring(cCodLinCred,4,1 )= '" & Trim(rsM!cValor) & "'" & _
                                                   " AND substring(cCodLinCred,6,1) = '" & Trim(rsFF!cValor) & "'" & _
                                                   " AND substring(cCodCta,4,2) = '" & Right(Trim(rsAg!cValor), 2) & "'" & _
                                                   " AND substring(cCodCta,6,3) in ( " & Prod(i) & ")" & _
                                                   " And dbcmactconsolidada..CreditoConsol.ndiasAtraso >= " & Val(lnRangoIni) & _
                                                   " And dbcmactconsolidada..CreditoConsol.ndiasatraso <= " & Val(lnRangoFin)
                                                   
                                             'Reg9.Open SQL9, dbCmactCentral, adOpenStatic, adLockReadOnly, adCmdText
                                             Conecta.AbreConexion
                                             Set Reg9 = Conecta.CargaRecordSet(SQL9)
                                             Conecta.CierraConexion
                                             
                                             lnTNumVen1Pr = Reg9!Numero
                                             lnTCapVen1Pr = IIf(IsNull(Reg9!CAPITAL), 0, Reg9!CAPITAL)
                                             Reg9.Close
                                             Set Reg9 = Nothing
                                           
                                           Case 1  ' Pyme  ******************
                                           
                                             If pPlanCuentas = "1" Then
                                                 lnRangoIni = 31   ' MicroEmpresa
                                                 lnRangoFin = 120
                                             Else
                                                 lnRangoIni = 31   ' MicroEmpresa
                                                 lnRangoFin = 9999
                                             End If
                                             
                                             SQL9 = "SELECT COUNT(dbcmactconsolidada..CreditoConsol.cCodCta) AS Numero, " & _
                                                   " SUM(dbcmactconsolidada..CreditoConsol.nSaldoCap) AS Capital " & _
                                                   " FROM dbcmactconsolidada..CreditoConsol " & _
                                                   " WHERE dbcmactconsolidada..CreditoConsol.nEstado ('2020','2021','2022','2030','2031','2031') " & _
                                                   " AND substring(cCodLinCred,4,1 )= '" & Trim(rsM!cValor) & "'" & _
                                                   " AND substring(cCodLinCred,6,1) = '" & Trim(rsFF!cValor) & "'" & _
                                                   " AND substring(cCodCta,4,2) = '" & Right(Trim(rsAg!cValor), 2) & "'" & _
                                                   " AND substring(cCodCta,6,3) in ( " & Prod(i) & ")" & _
                                                   " And ndiasAtraso >= " & Val(lnRangoIni) & _
                                                   " And ndiasatraso <= " & Val(lnRangoFin)
                                                   
                                             'Reg9.Open SQL9, dbCmactCentral, adOpenStatic, adLockReadOnly, adCmdText
                                             lnTNumVen1Pr = Reg9!Numero
                                             lnTCapVen1Pr = IIf(IsNull(Reg9!CAPITAL), 0, Reg9!CAPITAL)
                                             Reg9.Close
                                             Set Reg9 = Nothing
                                                   
                                          Case 2, 6, 7   ' Consumo / Hipotecario / MiVivienda *******************  OJO
                                             
                                             If pPlanCuentas = "1" Then
                                                 lnRangoIni = 31   ' Capital Vencido OJO
                                                 lnRangoFin = 90
                                             Else
                                                 lnRangoIni = 31   ' Capital Vencido OJO
                                                 lnRangoFin = 90
                                             End If
                                             
                                             SQL9 = "SELECT COUNT(cCodCta) AS Numero, " & _
                                                   " SUM(nSaldoCap) AS Capital, " & _
                                                   " SUM(nCapVencido) AS CapVencido " & _
                                                   " FROM dbcmactconsolidada..CreditoConsol " & _
                                                   " WHERE dbcmactconsolidada..CreditoConsol.nEstado in ('2020','2021','2022','2030','2031','2032') " & _
                                                   " AND substring(cCodLinCred,4,1 )= '" & Trim(rsM!cValor) & "'" & _
                                                   " AND substring(cCodLinCred,6,1) = '" & Trim(rsFF!cValor) & "'" & _
                                                   " AND substring(cCodCta,4,2) = '" & Right(Trim(rsAg!cValor), 2) & "'" & _
                                                   " AND substring(cCodCta,6,3) in ( " & Prod(i) & ")" & _
                                                   " And ndiasAtraso >= " & Val(lnRangoIni) & _
                                                   " And ndiasatraso <= " & Val(lnRangoFin)
                                                   
                                             'Reg9.Open SQL9, dbCmactCentral, adOpenStatic, adLockReadOnly, adCmdText
                                             Conecta.AbreConexion
                                             Set Reg9 = Conecta.CargaRecordSet(SQL9)
                                             Conecta.CierraConexion
                                             
                                             lnTNumVen1Pr = Reg9!Numero
                                             lnTCapVen1Pr = IIf(IsNull(Reg9!CapVencido), 0, Reg9!CapVencido)
                                             ' Sumo la Dif (Sald Cap - Cap Venc) *****
                                             lnTCapVigPr = lnTCapVigPr + IIf(IsNull(Reg9!CAPITAL - Reg9!CapVencido), 0, Reg9!CAPITAL - Reg9!CapVencido)
                                             Reg9.Close
                                             Set Reg9 = Nothing
                                             
                                             If pPlanCuentas = "1" Then
                                                 lnRangoIni = 91  ' Capital Vencido OJO
                                                 lnRangoFin = 120
                                         
                                                 SQL9 = "SELECT COUNT(cCodCta) AS Numero, " & _
                                                      " SUM(nSaldoCap) AS Capital " & _
                                                      " FROM dbcmactconsolidada..CreditoConsol " & _
                                                      " WHERE nEstado in ('2020','2021','2022','2030','2031','2032') " & _
                                                      " AND substring(cCodLinCred,4,1 )= '" & Trim(rsM!cValor) & "'" & _
                                                      " AND substring(cCodLinCred,6,1) = '" & Trim(rsFF!cValor) & "'" & _
                                                      " AND substring(cCodCta,4,2) = '" & Right(Trim(rsAg!cValor), 2) & "'" & _
                                                      " AND substring(cCodCta,6,3) in ( " & Prod(i) & ")" & _
                                                      " And ndiasAtraso >= " & Val(lnRangoIni) & _
                                                      " And ndiasatraso <= " & Val(lnRangoFin)
                                                      
                                                 'Reg9.Open SQL9, dbCmactCentral, adOpenStatic, adLockReadOnly, adCmdText
                                                 Conecta.AbreConexion
                                                 Set Reg9 = Conecta.CargaRecordSet(SQL9)
                                                 Conecta.CierraConexion
                                                 
                                                 lnTNumVen1Pr = lnTNumVen1Pr + Reg9!Numero
                                                 lnTCapVen1Pr = lnTCapVen1Pr + IIf(IsNull(Reg9!CAPITAL), 0, Reg9!CAPITAL)
                                                 Reg9.Close
                                                 Set Reg9 = Nothing
                                             
                                             End If
                                             
                                             
                                                   
                                          Case 3  ' Agricola  *************************
                                             
                                             If pPlanCuentas = "1" Then
                                                 lnRangoIni = 31
                                                 lnRangoFin = 120
                                             Else
                                                 lnRangoIni = 31
                                                 lnRangoFin = 9999
                                             End If
                                             
                                             SQL9 = "SELECT COUNT(dbcmactconsolidada..CreditoConsol.cCodCta) AS Numero, " & _
                                                   " SUM(dbcmactconsolidada..CreditoConsol.nSaldoCap) AS Capital " & _
                                                   " FROM dbcmactconsolidada..CreditoConsol " & _
                                                   " WHERE CreditoConsol.nEstado in ('2020','2021','2022','2030','2031','2032') " & _
                                                   " AND substring(cCodLinCred,4,1 )= '" & Trim(rsM!cValor) & "'" & _
                                                   " AND substring(cCodLinCred,6,1) = '" & Trim(rsFF!cValor) & "'" & _
                                                   " AND substring(cCodCta,4,2) = '" & Right(Trim(rsAg!cValor), 2) & "'" & _
                                                   " AND substring(cCodCta,6,3) in ( " & Prod(i) & ")" & _
                                                   " And ndiasAtraso >= " & Val(lnRangoIni) & _
                                                   " And ndiasatraso <= " & Val(lnRangoFin)
                                             'Reg9.Open SQL9, dbCmactCentral, adOpenStatic, adLockReadOnly, adCmdText
                                             Conecta.AbreConexion
                                             Set Reg9 = Conecta.CargaRecordSet(SQL9)
                                             Conecta.CierraConexion
                                             
                                             lnTNumVen1Pr = Reg9!Numero
                                             lnTCapVen1Pr = IIf(IsNull(Reg9!CAPITAL), 0, Reg9!CAPITAL)
                                             Reg9.Close
                                             Set Reg9 = Nothing
                                                   
                                          Case 4  ' Prendario  *************************
                                              If Trim(rsFF!cValor) = "1" Then  ' Si es RR PP
                                                lnRangoIni = 31
                                                lnRangoFin = 999
                                                SQL9 = "SELECT COUNT(cCodCta) AS Numero, " & _
                                                    " SUM(nSaldoCap) AS Capital " & _
                                                   " FROM dbcmactconsolidada..CreditoConsol " & _
                                                   " WHERE cEstado in ('2101','2104','2106','2107') " & _
                                                   " AND substring(cCodCta,4,2) = '" & Right(Trim(rsAg!cValor), 2) & "'" & _
                                                   " AND substring(cCodCta,6,3) in ( " & Prod(i) & ")" & _
                                                   " And DATEDIFF(dd ,dFecVenc,'" & Format(Co.GEtFechaCierreMes, "mm/dd/yyyy") & "') >= " & Val(lnRangoIni) & _
                                                   " And DATEDIFF(dd ,dFecVenc,'" & Format(Co.GEtFechaCierreMes, "mm/dd/yyyy") & "') <= " & Val(lnRangoFin)
                                                   
                                                 'Reg9.Open SQL9, dbCmactCentral, adOpenStatic, adLockReadOnly, adCmdText
                                                 Conecta.AbreConexion
                                                 Set Reg9 = Conecta.CargaRecordSet(SQL9)
                                                 Conecta.CierraConexion
                                                 
                                                 lnTNumVen1Pr = Reg9!Numero
                                                 lnTCapVen1Pr = IIf(IsNull(Reg9!CAPITAL), 0, Reg9!CAPITAL)
                                                 Reg9.Close
                                                 Set Reg9 = Nothing
                                             End If
                                           
                                           
                                          Case 5  '  Administrativos   *************************  OJO
                                             
                                             If pPlanCuentas = "1" Then
                                                 lnRangoIni = 31   ' Capital Vencido OJO
                                                 lnRangoFin = 90
                                             Else
                                                 lnRangoIni = 31   ' Capital Vencido OJO
                                                 lnRangoFin = 90
                                             End If
                                             
                                             SQL9 = "SELECT COUNT(cCodCta) AS Numero, " & _
                                                   " SUM(nSaldoCap) AS Capital, " & _
                                                   " SUM(nCapVencido) AS CapVencido " & _
                                                   " FROM dbcmactconsolidada..CreditoConsol " & _
                                                   " WHERE dbcmactconsolidada..CreditoConsol.nEstado in ('2020','2021','2022','2030','2031','2032') " & _
                                                   " AND substring(cCodLinCred,4,1 )= '" & Trim(rsM!cValor) & "'" & _
                                                   " AND substring(cCodLinCred,6,1) = '" & Trim(rsFF!cValor) & "'" & _
                                                   " AND substring(cCodCta,4,2) = '" & Right(Trim(rsAg!cValor), 2) & "'" & _
                                                   " AND substring(cCodCta,6,3) in ( " & Prod(i) & ")" & _
                                                   " And ndiasAtraso >= " & Val(lnRangoIni) & _
                                                   " And ndiasatraso <= " & Val(lnRangoFin)
                                                   
                                             'Reg9.Open SQL9, dbCmactCentral, adOpenStatic, adLockReadOnly, adCmdText
                                             Conecta.AbreConexion
                                             Set Reg9 = Conecta.CargaRecordSet(SQL9)
                                             Conecta.CierraConexion
                                             
                                             lnTNumVen1Pr = Reg9!Numero
                                             lnTCapVen1Pr = IIf(IsNull(Reg9!CapVencido), 0, Reg9!CapVencido)
                                             lnTCapVigPr = lnTCapVigPr + IIf(IsNull(Reg9!CAPITAL - Reg9!CapVencido), 0, Reg9!CAPITAL - Reg9!CapVencido)
                                             Reg9.Close
                                             Set Reg9 = Nothing
                                             
                                             If pPlanCuentas = "1" Then
                                                 lnRangoIni = 91  ' Capital Vencido OJO
                                                 lnRangoFin = 120
                                         
                                                 SQL9 = "SELECT COUNT(cCodCta) AS Numero, " & _
                                                      " SUM(nSaldoCap) AS Capital " & _
                                                      " FROM dbcmactconsolidada..CreditoConsol " & _
                                                      " WHERE nEstado in ('2020','2021','2022','2030','2031','2032') " & _
                                                      " AND substring(cCodLinCred,4,1 )= '" & Trim(rsM!cValor) & "'" & _
                                                      " AND substring(cCodLinCred,6,1) = '" & Trim(rsFF!cValor) & "'" & _
                                                      " AND substring(cCodCta,4,2) = '" & Right(Trim(rsAg!cValor), 2) & "'" & _
                                                      " AND substring(cCodCta,6,3) in ( " & Prod(i) & ")" & _
                                                      " And ndiasAtraso >= " & Val(lnRangoIni) & _
                                                      " And ndiasatraso <= " & Val(lnRangoFin)
                                                      
                                                 'Reg9.Open SQL9, dbCmactCentral, adOpenStatic, adLockReadOnly, adCmdText
                                                 Conecta.AbreConexion
                                                 Set Reg9 = Conecta.CargaRecordSet(SQL9)
                                                 Conecta.CierraConexion
                                                 
                                                 lnTNumVen1Pr = lnTNumVen1Pr + Reg9!Numero
                                                 lnTCapVen1Pr = lnTCapVen1Pr + IIf(IsNull(Reg9!CAPITAL), 0, Reg9!CAPITAL)
                                                 Reg9.Close
                                                 Set Reg9 = Nothing
                                             
                                             End If
                                           
                                           
                                          Case Else
                                               lnTNumVen1Pr = 0
                                               lnTCapVen1Pr = 0
                                          
                                       End Select
                                       
                                       '=================================================
                                       '====== Cartera Vencida 2 Producto de Agencia  ===
                                       Select Case i
                                                   
                                           Case 0  ' Comerciales  ******************
                                           
                                             If pPlanCuentas = "1" Then  'Solo Plan Actual
                                                 lnRangoIni = 121   ' Comerciales
                                                 lnRangoFin = 999
                                           
                                                SQL9 = "SELECT COUNT(CreditoConsol.cCodCta) AS Numero, " & _
                                                      " SUM(dbcmactconsolidada..CreditoConsol.nSaldoCap) AS Capital " & _
                                                      " FROM dbcmactconsolidada..CreditoConsol " & _
                                                      " WHERE dbcmactconsolidada..CreditoConsol.nEstado in ('2020','2021','2022','2030','2031','2032') " & _
                                                      " AND substring(cCodLinCred,4,1 )= '" & Trim(rsM!cValor) & "'" & _
                                                      " AND substring(cCodLinCred,6,1) = '" & Trim(rsFF!cValor) & "'" & _
                                                      " AND substring(cCodCta,4,2) = '" & Right(Trim(rsAg!cValor), 2) & "'" & _
                                                      " AND substring(cCodCta,6,3) in ( " & Prod(i) & ")" & _
                                                      " And ndiasAtraso >= " & Val(lnRangoIni) & _
                                                      " And ndiasatraso <= " & Val(lnRangoFin)
                                                      
                                                'Reg9.Open SQL9, dbCmactCentral, adOpenStatic, adLockReadOnly, adCmdText
                                                
                                                Conecta.AbreConexion
                                                Set Reg9 = Conecta.CargaRecordSet(SQL9)
                                                Conecta.CierraConexion
                                                
                                                lnTNumVen2Pr = Reg9!Numero
                                                lnTCapVen2Pr = IIf(IsNull(Reg9!CAPITAL), 0, Reg9!CAPITAL)
                                                Reg9.Close
                                                Set Reg9 = Nothing
                                             End If
                                           
                                           Case 1  ' Pyme  ******************
                                           
                                             If pPlanCuentas = "1" Then  'Solo Plan Actual
                                                lnRangoIni = 121   ' MicroEmpresa
                                                lnRangoFin = 999
                                                
                                                SQL9 = "SELECT COUNT(dbcmactconsolidada..CreditoConsol.cCodCta) AS Numero, " & _
                                                      " SUM(dbcmactconsolidada..CreditoConsol.nSaldoCap) AS Capital " & _
                                                      " FROM dbcmactconsolidada..CreditoConsol " & _
                                                      " WHERE dbcmactconsolidada..CreditoConsol.nEstado in ('2020','2021','2022','2030','2031','2032') " & _
                                                      " AND substring(cCodLinCred,4,1 )= '" & Trim(rsM!cValor) & "'" & _
                                                      " AND substring(cCodLinCred,6,1) = '" & Trim(rsFF!cValor) & "'" & _
                                                      " AND substring(cCodCta,4,2) = '" & Right(Trim(rsAg!cValor), 2) & "'" & _
                                                      " AND substring(cCodCta,6,3) in ( " & Prod(i) & ")" & _
                                                      " And ndiasAtraso >= " & Val(lnRangoIni) & _
                                                      " And ndiasatraso <= " & Val(lnRangoFin)
                                                      
                                                'Reg9.Open SQL9, dbCmactCentral, adOpenStatic, adLockReadOnly, adCmdText
                                                Conecta.AbreConexion
                                                Set Reg9 = Conecta.CargaRecordSet(SQL9)
                                                Conecta.CierraConexion
                                                
                                                lnTNumVen2Pr = Reg9!Numero
                                                lnTCapVen2Pr = IIf(IsNull(Reg9!CAPITAL), 0, Reg9!CAPITAL)
                                                Reg9.Close
                                                Set Reg9 = Nothing
                                             End If
                                                   
                                          Case 2, 6, 7   ' Consumo / Hipotecaja / MiVivienda **************  OJO
                                          
                                             If pPlanCuentas = "1" Then
                                                 lnRangoIni = 121  ' Capital Vencido OJO
                                                 lnRangoFin = 999
                                             Else
                                                 lnRangoIni = 91   ' Capital Vencido OJO
                                                 lnRangoFin = 9999
                                             End If
                                          
                                             SQL9 = "SELECT COUNT(cCodCta) AS Numero, " & _
                                                   " SUM(nSaldoCap) AS Capital " & _
                                                   " FROM dbcmactconsolidada..CreditoConsol " & _
                                                   " WHERE nEstado in ('2020','2021','2022','2030','2031','2032') " & _
                                                   " AND substring(cCodLinCred,4,1 )= '" & Trim(rsM!cValor) & "'" & _
                                                   " AND substring(cCodLinCred,6,1) = '" & Trim(rsFF!cValor) & "'" & _
                                                   " AND substring(cCodCta,4,2) = '" & Right(Trim(rsAg!cValor), 2) & "'" & _
                                                   " AND substring(cCodCta,6,3) in ( " & Prod(i) & ")" & _
                                                   " And ndiasAtraso >= " & Val(lnRangoIni) & _
                                                   " And ndiasatraso <= " & Val(lnRangoFin)
                                                   
                                             'Reg9.Open SQL9, dbCmactCentral, adOpenStatic, adLockReadOnly, adCmdText
                                             
                                             Conecta.AbreConexion
                                             Set Reg9 = Conecta.CargaRecordSet(SQL9)
                                             Conecta.CierraConexion
                                             
                                             lnTNumVen2Pr = Reg9!Numero
                                             lnTCapVen2Pr = IIf(IsNull(Reg9!CAPITAL), 0, Reg9!CAPITAL)
                                             Reg9.Close
                                             Set Reg9 = Nothing
                                          
                                          Case 3  ' Agricola  *************************
                                             
                                             If pPlanCuentas = "1" Then
                                                 lnRangoIni = 121
                                                 lnRangoFin = 999
                                                 
                                                 SQL9 = "SELECT COUNT(dbcmactconsolidada..CreditoConsol.cCodCta) AS Numero, " & _
                                                       " SUM(dbcmactconsolidada..CreditoConsol.nSaldoCap) AS Capital " & _
                                                       " FROM dbcmactconsolidada..CreditoConsol " & _
                                                       " WHERE dbcmactconsolidada..CreditoConsol.nEstado ('2020','2021','2022','2030','2031','2032') " & _
                                                       " AND substring(cCodLinCred,4,1 )= '" & Trim(rsM!cValor) & "'" & _
                                                       " AND substring(cCodLinCred,6,1) = '" & Trim(rsFF!cValor) & "'" & _
                                                       " AND substring(cCodCta,2,2) = '" & Right(Trim(rsAg!cValor), 2) & "'" & _
                                                       " AND substring(cCodCta,4,3) in ( " & Prod(i) & ")" & _
                                                       " And ndiasAtraso >= " & Val(lnRangoIni) & _
                                                       " And ndiasatraso <= " & Val(lnRangoFin)
                                                 'Reg9.Open SQL9, dbCmactCentral, adOpenStatic, adLockReadOnly, adCmdText
                                                 Conecta.AbreConexion
                                                 Set Reg9 = Conecta.CargaRecordSet(SQL9)
                                                 Conecta.CierraConexion
                                                 
                                                 lnTNumVen2Pr = Reg9!Numero
                                                 lnTCapVen2Pr = IIf(IsNull(Reg9!CAPITAL), 0, Reg9!CAPITAL)
                                                 Reg9.Close
                                                 Set Reg9 = Nothing
                                             End If
                                          
                                          
                                          Case 5  ' Administrativo *************************  OJO
                                          
                                             If pPlanCuentas = "1" Then
                                                 lnRangoIni = 121  ' Capital Vencido OJO
                                                 lnRangoFin = 9999
                                             Else
                                                 lnRangoIni = 91   ' Capital Vencido OJO
                                                 lnRangoFin = 9999
                                             End If
                                          
                                             SQL9 = "SELECT COUNT(cCodCta) AS Numero, " & _
                                                   " SUM(nSaldoCap) AS Capital " & _
                                                   " FROM dbcmactconsolidada..CreditoConsol " & _
                                                   " WHERE nEstado in ('2020','2021','2023','2030','2031','2032') " & _
                                                   " AND substring(cCodLinCred,4,1 )= '" & Trim(rsM!cValor) & "'" & _
                                                   " AND substring(cCodLinCred,6,1) = '" & Trim(rsFF!cValor) & "'" & _
                                                   " AND substring(cCodCta,4,2) = '" & Right(Trim(rsAg!cValor), 2) & "'" & _
                                                   " AND substring(cCodCta,6,3) in ( " & Prod(i) & ")" & _
                                                   " And ndiasAtraso >= " & Val(lnRangoIni) & _
                                                   " And ndiasatraso <= " & Val(lnRangoFin)
                                                   
                                             'Reg9.Open SQL9, dbCmactCentral, adOpenStatic, adLockReadOnly, adCmdText
                                             Conecta.AbreConexion
                                             Set Reg9 = Conecta.CargaRecordSet(SQL9)
                                             Conecta.CierraConexion
                                             
                                             lnTNumVen2Pr = Reg9!Numero
                                             lnTCapVen2Pr = IIf(IsNull(Reg9!CAPITAL), 0, Reg9!CAPITAL)
                                             Reg9.Close
                                             Set Reg9 = Nothing
                                          
                                          Case Else
                                              lnTNumVen2Pr = 0
                                              lnTCapVen2Pr = 0
                                               
                                       End Select
                                       
                                       '=================================================
                                       '====== En Asesoria Legal Producto de Agencia  ===
                                       Select Case i
                                           Case 0 ' Comercial ******************
                                             'lnRangoIni = -999
                                             'lnRangoFin = 999
                                             SQL9 = "SELECT COUNT(cCodCta) AS Numero, " & _
                                                   " SUM(nSaldoCap) AS Capital " & _
                                                   " FROM dbcmactconsolidada..CreditoConsol " & _
                                                   " WHERE nEstado = '2201' AND cCondCre =" & gColocEstJudicial & " AND nSaldoCap > 0" & _
                                                   " AND substring(cCodLinCred,4,1 )= '" & Trim(rsM!cValor) & "'" & _
                                                   " AND substring(cCodLinCred,6,1) = '" & Trim(rsFF!cValor) & "'" & _
                                                   " AND substring(cCodCta,4,2) = '" & Right(Trim(rsAg!cValor), 2) & "'" & _
                                                   " AND substring(cCodCta,6,3) in ( " & Prod(i) & ")"
                                                   
                                             'Reg9.Open SQL9, dbCmactCentral, adOpenStatic, adLockReadOnly, adCmdText
                                             
                                             Conecta.AbreConexion
                                             Set Reg9 = Conecta.CargaRecordSet(SQL9)
                                             Conecta.CierraConexion
                                             
                                             lnTNumJudPr = Reg9!Numero
                                             lnTCapJudPr = IIf(IsNull(Reg9!CAPITAL), 0, Reg9!CAPITAL)
                                             Reg9.Close
                                             Set Reg9 = Nothing
                                             '********
                                           Case 1 ' Pyme  ******************
                                             'lnRangoIni = -999
                                             'lnRangoFin = 999
                                             SQL9 = "SELECT COUNT(cCodCta) AS Numero, " & _
                                                   " SUM(nSaldoCap) AS Capital " & _
                                                   " FROM dbcmactconsolidada..CreditoConsol " & _
                                                   " WHERE nEstado = '2201' AND cCondCre =" & gColocEstJudicial & " AND nSaldoCap > 0" & _
                                                   " AND substring(cCodLinCred,4,1 )= '" & Trim(rsM!cValor) & "'" & _
                                                   " AND substring(cCodLinCred,6,1) = '" & Trim(rsFF!cValor) & "'" & _
                                                   " AND substring(cCodCta,1,2) = '" & Right(Trim(rsAg!cValor), 2) & "'" & _
                                                   " AND substring(cCodCta,3,3) in ( " & Prod(i) & ")"
                                                   
                                             'Reg9.Open SQL9, dbCmactCentral, adOpenStatic, adLockReadOnly, adCmdText
                                             Conecta.AbreConexion
                                             Set Reg9 = Conecta.CargaRecordSet(SQL9)
                                             Conecta.CierraConexion
                                             
                                             lnTNumJudPr = Reg9!Numero
                                             lnTCapJudPr = IIf(IsNull(Reg9!CAPITAL), 0, Reg9!CAPITAL)
                                             Reg9.Close
                                             Set Reg9 = Nothing
                                             '********
                                           Case 2, 6, 7   ' Consumo / Hipotecario / MiVivienda *************
                                           '  lnRangoIni = -999
                                           '  lnRangoFin = 30
                                             SQL9 = "SELECT COUNT(cCodCta) AS Numero, " & _
                                                   " SUM(nSaldoCap) AS Capital " & _
                                                   " FROM dbcmactconsolidada..CreditoConsol " & _
                                                   " WHERE nEstado = '2201' AND cCondCre =" & gColocEstJudicial & " AND nSaldoCap > 0" & _
                                                   " AND substring(cCodLinCred,4,1 )= '" & Trim(rsM!cValor) & "'" & _
                                                   " AND substring(cCodLinCred,6,1) = '" & Trim(rsFF!cValor) & "'" & _
                                                   " AND substring(cCodCta,4,2) = '" & Right(Trim(rsAg!cValor), 2) & "'" & _
                                                   " AND substring(cCodCta,6,3) in ( " & Prod(i) & ")"
                                                   
                                             'Reg9.Open SQL9, dbCmactCentral, adOpenStatic, adLockReadOnly, adCmdText
                                             Conecta.AbreConexion
                                             Set Reg9 = Conecta.CargaRecordSet(SQL9)
                                             Conecta.CierraConexion
                                             
                                             lnTNumJudPr = Reg9!Numero
                                             lnTCapJudPr = IIf(IsNull(Reg9!CAPITAL), 0, Reg9!CAPITAL)
                                             Reg9.Close
                                             Set Reg9 = Nothing
                                                   
                                          Case 3  ' Agricola  *************************
                                             'lnRangoIni = -999
                                             'lnRangoFin = 30
                                             SQL9 = "SELECT COUNT(cCodCta) AS Numero, " & _
                                                   " SUM(nSaldoCap) AS Capital " & _
                                                   " FROM dbcmactconsolidada..CreditoConsol " & _
                                                   " WHERE nEstado = '2201' AND cCondCre =" & gColocEstJudicial & " AND nSaldoCap > 0" & _
                                                   " AND substring(cCodLinCred,4,1 )= '" & Trim(rsM!cValor) & "'" & _
                                                   " AND substring(cCodLinCred,6,1) = '" & Trim(rsFF!cValor) & "'" & _
                                                   " AND substring(cCodCta,4,2) = '" & Right(Trim(rsAg!cValor), 2) & "'" & _
                                                   " AND substring(cCodCta,6,3) in ( " & Prod(i) & ")"
                                                   '" And CreditoConsol.ndiasAtraso >= " & Val(lnRangoIni) & _
                                                   '" And CreditoConsol.ndiasatraso <= " & Val(lnRangoFin)
                                            'Reg9.Open SQL9, dbCmactCentral, adOpenStatic, adLockReadOnly, adCmdText
                                             Conecta.AbreConexion
                                             Set Reg9 = Conecta.CargaRecordSet(SQL9)
                                             Conecta.CierraConexion
                                            
                                            lnTNumJudPr = Reg9!Numero
                                            lnTCapJudPr = IIf(IsNull(Reg9!CAPITAL), 0, Reg9!CAPITAL)
                                            Reg9.Close
                                            Set Reg9 = Nothing
                                       
                                          Case Else
                                              lnTNumJudPr = 0
                                              lnTCapJudPr = 0
                                               
                                       End Select
                                       
                                       ' ***************************************
                                       ' ***** Asigna Valores a la Agencia ******
                                       lnTCredAG = lnTCredAG + lnTCredPr
                                       lnTSaldoCapAG = lnTSaldoCapAG + lnTSaldoCapPr
                                       lnTNumVigAG = lnTNumVigAG + lnTNumVigPr
                                       lnTCapVigAG = lnTCapVigAG + lnTCapVigPr
                                       lnTNumVen1AG = lnTNumVen1AG + lnTNumVen1Pr
                                       lnTCapVen1AG = lnTCapVen1AG + lnTCapVen1Pr
                                       lnTNumVen2AG = lnTNumVen2AG + lnTNumVen2Pr
                                       lnTCapVen2AG = lnTCapVen2AG + lnTCapVen2Pr
                                       lnTNumJudAG = lnTNumJudAG + lnTNumJudPr
                                       lnTCapJudAG = lnTCapJudAG + lnTCapJudPr
                                       ' *****  Imprime
                                       lsRTfImp = lsRTfImp & " " & lsDescPr
                                       lsRTfImp = lsRTfImp & ImpreFormat(lnTCredPr, 10, 0, True)
                                       lsRTfImp = lsRTfImp & ImpreFormat(lnTSaldoCapPr, 13, 2, True) & " | "
                                       lsRTfImp = lsRTfImp & ImpreFormat(lnTNumVigPr, 8, 0, True)
                                       lsRTfImp = lsRTfImp & ImpreFormat(lnTCapVigPr, 13, 2, True) & " | "
                                       lsRTfImp = lsRTfImp & ImpreFormat(lnTNumVen1Pr, 7, 0, True)
                                       lsRTfImp = lsRTfImp & ImpreFormat(lnTCapVen1Pr, 12, 2, True) & " | "
                                       lsRTfImp = lsRTfImp & ImpreFormat(lnTNumVen2Pr, 7, 0, True)
                                       lsRTfImp = lsRTfImp & ImpreFormat(lnTCapVen2Pr, 12, 2, True) & " | "
                                       lsRTfImp = lsRTfImp & ImpreFormat(lnTNumJudPr, 7, 0, True)
                                       lsRTfImp = lsRTfImp & ImpreFormat(lnTCapJudPr, 12, 2, True) & " | "
                                       lsRTfImp = lsRTfImp & ImpreFormat(Round(((lnTSaldoCapPr - lnTCapVigPr) / lnTSaldoCapPr) * 100, 2), 3, 2) & "%"
                                       lsRTfImp = lsRTfImp & Chr(10)
                                       liLineas = liLineas + 1

                                    End If  ' Fin de Si encuentra Datos
                                        
                                        
                               Next i   ' Para el Sgte Producto
                               
                               
                               ' **********************
                               ' ************ Asigna Valores a Fuente Financ
                               lnTCredF = lnTCredF + lnTCredAG
                               lnTSaldoCapF = lnTSaldoCapF + lnTSaldoCapAG
                               lnTNumVigF = lnTNumVigF + lnTNumVigAG
                               lnTCapVigF = lnTCapVigF + lnTCapVigAG
                               lnTNumVen1F = lnTNumVen1F + lnTNumVen1AG
                               lnTCapVen1F = lnTCapVen1F + lnTCapVen1AG
                               lnTNumVen2F = lnTNumVen2F + lnTNumVen2AG
                               lnTCapVen2F = lnTCapVen2F + lnTCapVen2AG
                               lnTNumJudF = lnTNumJudF + lnTNumJudAG
                               lnTCapJudF = lnTCapJudF + lnTCapJudAG
                               ' *****  Imprime
                               lsRTfImp = lsRTfImp & String(155, "-") & Chr(10)
                               lsRTfImp = lsRTfImp & Space(15)
                               lsRTfImp = lsRTfImp & ImpreFormat(lnTCredAG, 10, 0, True)
                               lsRTfImp = lsRTfImp & ImpreFormat(lnTSaldoCapAG, 13, 2, True) & " | "
                               lsRTfImp = lsRTfImp & ImpreFormat(lnTNumVigAG, 8, 0, True)
                               lsRTfImp = lsRTfImp & ImpreFormat(lnTCapVigAG, 13, 2, True) & " | "
                               lsRTfImp = lsRTfImp & ImpreFormat(lnTNumVen1AG, 7, 0, True)
                               lsRTfImp = lsRTfImp & ImpreFormat(lnTCapVen1AG, 12, 2, True) & " | "
                               lsRTfImp = lsRTfImp & ImpreFormat(lnTNumVen2AG, 7, 0, True)
                               lsRTfImp = lsRTfImp & ImpreFormat(lnTCapVen2AG, 12, 2, True) & " | "
                               lsRTfImp = lsRTfImp & ImpreFormat(lnTNumJudAG, 7, 0, True)
                               lsRTfImp = lsRTfImp & ImpreFormat(lnTCapJudAG, 12, 2, True) & " | "
                               lsRTfImp = lsRTfImp & ImpreFormat(Round(((lnTSaldoCapAG - lnTCapVigAG) / lnTSaldoCapAG) * 100, 2), 3, 2) & "%"
                               lsRTfImp = lsRTfImp & Chr(10)
                               lsRTfImp = lsRTfImp & String(155, "-") & Chr(10)
                               'lsRTfImp = lsRTfImp & Chr(10)
                               liLineas = liLineas + 4
                               ' *** Sgte agencia
                             End If  ' Si hay datos
                            rsAg.MoveNext
                            

                            Loop  ' de Agencia
                            
                            ' **********************
                            ' ************ Asigna Valores a Moneda
                            lnTCredM = lnTCredM + lnTCredF
                            lnTSaldoCapM = lnTSaldoCapM + lnTSaldoCapF
                            lnTNumVigM = lnTNumVigM + lnTNumVigF
                            lnTCapVigM = lnTCapVigM + lnTCapVigF
                            lnTNumVen1M = lnTNumVen1M + lnTNumVen1F
                            lnTCapVen1M = lnTCapVen1M + lnTCapVen1F
                            lnTNumVen2M = lnTNumVen2M + lnTNumVen2F
                            lnTCapVen2M = lnTCapVen2M + lnTCapVen2F
                            lnTNumJudM = lnTNumJudM + lnTNumJudF
                            lnTCapJudM = lnTCapJudM + lnTCapJudF
                            ' *****  Imprime
                            'lsRTfImp = lsRTfImp & Chr(10)
                            lsRTfImp = lsRTfImp & String(155, "=") & Chr(10)
                            lsRTfImp = lsRTfImp & Mid(lsDescF, 1, 15)
                            lsRTfImp = lsRTfImp & ImpreFormat(lnTCredF, 10, 0, True)
                            lsRTfImp = lsRTfImp & ImpreFormat(lnTSaldoCapF, 13, 2, True) & " | "
                            lsRTfImp = lsRTfImp & ImpreFormat(lnTNumVigF, 8, 0, True)
                            lsRTfImp = lsRTfImp & ImpreFormat(lnTCapVigF, 13, 2, True) & " | "
                            lsRTfImp = lsRTfImp & ImpreFormat(lnTNumVen1F, 7, 0, True)
                            lsRTfImp = lsRTfImp & ImpreFormat(lnTCapVen1F, 12, 2, True) & " | "
                            lsRTfImp = lsRTfImp & ImpreFormat(lnTNumVen2F, 7, 0, True)
                            lsRTfImp = lsRTfImp & ImpreFormat(lnTCapVen2F, 12, 2, True) & " | "
                            lsRTfImp = lsRTfImp & ImpreFormat(lnTNumJudF, 7, 0, True)
                            lsRTfImp = lsRTfImp & ImpreFormat(lnTCapJudF, 12, 2, True) & " | "
                            lsRTfImp = lsRTfImp & ImpreFormat(Round(((lnTSaldoCapF - lnTCapVigF) / lnTSaldoCapF) * 100, 2), 3, 2) & "%"
                            lsRTfImp = lsRTfImp & Chr(10)
                            lsRTfImp = lsRTfImp & String(155, "=") & Chr(10)
                            lsRTfImp = lsRTfImp & Chr(10)
                            liLineas = liLineas + 5
                               
                               
                        End If  ' Fin de Fuente financiamiento
                        
                        rsFF.MoveNext
                          ' Arturo
                    Loop  ' De la Fuente Financiamiento
                    
                            
                    ' **********************
                    ' ************ Asigna Valores a Moneda
                    lnTCred = lnTCred + lnTCredM
                    lnTSaldoCap = lnTSaldoCap + lnTSaldoCapM
                    lnTNumVig = lnTNumVig + lnTNumVigM
                    lnTCapVig = lnTCapVig + lnTCapVigM
                    lnTNumVen1 = lnTNumVen1 + lnTNumVen1M
                    lnTCapVen1 = lnTCapVen1 + lnTCapVen1M
                    lnTNumVen2 = lnTNumVen2 + lnTNumVen2M
                    lnTCapVen2 = lnTCapVen2 + lnTCapVen2M
                    lnTNumJud = lnTNumJud + lnTNumJudM
                    lnTCapJud = lnTCapJud + lnTCapJudM
                    ' *****  Imprime
                    lsRTfImp = lsRTfImp & Chr(10)
                    lsRTfImp = lsRTfImp & String(155, "=") & Chr(10)
                    lsRTfImp = lsRTfImp & Chr(27) & Chr(69)
                    lsRTfImp = lsRTfImp & " TOTAL Moneda  "
                    lsRTfImp = lsRTfImp & ImpreFormat(lnTCredM, 10, 0, True)
                    lsRTfImp = lsRTfImp & ImpreFormat(lnTSaldoCapM, 13, 2, True) & " | "
                    lsRTfImp = lsRTfImp & ImpreFormat(lnTNumVigM, 8, 0, True)
                    lsRTfImp = lsRTfImp & ImpreFormat(lnTCapVigM, 13, 2, True) & " | "
                    lsRTfImp = lsRTfImp & ImpreFormat(lnTNumVen1M, 7, 0, True)
                    lsRTfImp = lsRTfImp & ImpreFormat(lnTCapVen1M, 12, 2, True) & " | "
                    lsRTfImp = lsRTfImp & ImpreFormat(lnTNumVen2M, 7, 0, True)
                    lsRTfImp = lsRTfImp & ImpreFormat(lnTCapVen2M, 12, 2, True) & " | "
                    lsRTfImp = lsRTfImp & ImpreFormat(lnTNumJudM, 7, 0, True)
                    lsRTfImp = lsRTfImp & ImpreFormat(lnTCapJudM, 12, 2, True) & " | "
                    lsRTfImp = lsRTfImp & ImpreFormat(Round(((lnTSaldoCapM - lnTCapVigM) / lnTSaldoCapM) * 100, 2), 3, 2) & "%"
                    lsRTfImp = lsRTfImp & Chr(27) & Chr(70)
                    lsRTfImp = lsRTfImp & Chr(10)
                    lsRTfImp = lsRTfImp & String(155, "=") & Chr(10)
                    lsRTfImp = lsRTfImp & Chr(10)
                    liLineas = liLineas + 5
                End If
             
            End If '---liDatos - Moneda
            lbSiCambio = True
            rsM.MoveNext
            If Not rsM.EOF Then
                lsRTfImp = lsRTfImp & Chr(12)
                cpag = cpag + 1
                liLineas = 0
            End If
        Loop ' ---- Moneda
    End If '--- Moneda
    rsM.Close
    Set rsM = Nothing
    
Else
    ImpCarteraCredxMoneda = False
End If '--- liDatos - General
lsRtfImpG = lsRTfImp ' Asigna el valor
End Function


Private Sub tvwReporte_Click()
Dim NodRep  As Node
Dim lsDesc As String

Set NodRep = tvwReporte.SelectedItem

If NodRep Is Nothing Then
   Exit Sub
End If

lsDesc = Mid(NodRep.Text, 8, Len(NodRep.Text) - 7)
fnRepoSelec = CLng(NodRep.Tag)
cmdImprimir.Enabled = True
Select Case fnRepoSelec
    Case 108701:
    Case 108702:
    Case 108703:
    Case 108704:
    Case 108705:
    Case 108706:
    Case 108707:
    Case 108708:
    Case 108709:
    Case 108710:
    Case 108711:
    Case 108801:
    Case 108802:
    Case 108803:
    Case 108804:
    Case 108806:
    Case 108808:
Case Else
    cmdImprimir.Enabled = False

End Select

Set NodRep = Nothing
End Sub
