VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmAnxVarCambYBackTesting 
   Caption         =   "VAR Cambiario y BackTesting"
   ClientHeight    =   5055
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   6735
   Begin VB.Frame fraFechaCentral 
      Caption         =   "Fecha VarCambiario"
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
      Height          =   745
      Left            =   240
      TabIndex        =   11
      Top             =   240
      Width           =   2010
      Begin MSMask.MaskEdBox txtFecCentral 
         Height          =   360
         Left            =   135
         TabIndex        =   1
         Top             =   270
         Width           =   1745
         _ExtentX        =   3069
         _ExtentY        =   635
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "Buscar"
      Height          =   335
      Left            =   3840
      TabIndex        =   10
      Top             =   1680
      Width           =   810
   End
   Begin VB.CheckBox chkRangFec 
      Caption         =   "Rango de Fechas"
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
      Height          =   375
      Left            =   4320
      TabIndex        =   9
      Top             =   90
      Width           =   1840
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   345
      Left            =   5350
      TabIndex        =   3
      Top             =   4560
      Width           =   1155
   End
   Begin VB.Frame fraTipoCamb 
      Caption         =   "Tipo de Cambio SBS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   3015
      Left            =   240
      TabIndex        =   0
      Top             =   1440
      Width           =   6255
      Begin Sicmact.FlexEdit flxTipoCambioSBS 
         Height          =   2055
         Left            =   240
         TabIndex        =   13
         Top             =   720
         Width           =   4620
         _ExtentX        =   8149
         _ExtentY        =   3625
         Cols0           =   5
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "Nro-Fecha-TC Compra-TC Venta-tc"
         EncabezadosAnchos=   "500-1300-1200-1200-3"
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-1-2-3-X"
         ListaControles  =   "0-2-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-R-R"
         FormatosEdit    =   "0-0-0-0-3"
         CantEntero      =   10
         TextArray0      =   "Nro"
         lbUltimaInstancia=   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   495
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.TextBox txtAnio 
         Height          =   315
         Left            =   2760
         TabIndex        =   18
         Top             =   255
         Width           =   615
      End
      Begin VB.ComboBox cboMes 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   260
         Width           =   1335
      End
      Begin VB.CommandButton cmdGuardar 
         Caption         =   "Guardar"
         Height          =   375
         Left            =   5110
         TabIndex        =   16
         Top             =   1560
         Width           =   900
      End
      Begin VB.CheckBox chkDatoBal 
         Caption         =   "Dato de Balance"
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
         Height          =   495
         Left            =   5050
         TabIndex        =   14
         Top             =   2160
         Width           =   1095
      End
      Begin VB.CommandButton cmdQuitar 
         Caption         =   "Quitar"
         Height          =   375
         Left            =   5110
         TabIndex        =   15
         Top             =   1080
         Width           =   900
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "Agregar"
         Height          =   375
         Left            =   5110
         TabIndex        =   12
         Top             =   600
         Width           =   900
      End
   End
   Begin MSMask.MaskEdBox txtFecFin 
      Height          =   300
      Left            =   5160
      TabIndex        =   6
      Top             =   870
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Frame frmRangFec 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4080
      TabIndex        =   4
      Top             =   120
      Width           =   2415
      Begin MSMask.MaskEdBox txtFecIni 
         Height          =   300
         Left            =   1080
         TabIndex        =   5
         Top             =   330
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblHasta 
         Caption         =   "Hasta :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   770
         Width           =   735
      End
      Begin VB.Label lblDesde 
         Caption         =   "Desde :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdGenerar 
      Caption         =   "Generar"
      Height          =   345
      Left            =   240
      TabIndex        =   2
      Top             =   4560
      Width           =   1155
   End
End
Attribute VB_Name = "frmAnxVarCambYBackTesting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************************
'*** Nombre : frmAnxVarCambYBackTesting
'*** Descripción : Formulario para generar el Var Cambiario y BackTesting
'*** Creación : por NAGL el 20170717
'********************************************************************************

Dim nItem As Integer
Dim DAnxRies As New DAnexoRiesgos
Dim fnFilSel As Integer
Dim rs As New ADODB.Recordset
Dim oGen As New DGeneral
Dim pnTipo As Integer
Dim valorCeldaFec As String, valorCelda1 As String, valorCelda2 As String

Public Sub Inicio()
Dim pdFecha As Date
    Me.txtFecCentral = gdFecSis - 1
    pdFecha = Me.txtFecCentral
    flxTipoCambioSBS.lbEditarFlex = False
    txtFecIni.Enabled = False
    txtFecFin.Enabled = False
    pnTipo = 2
    Set rs = oGen.GetConstante(1010)
    While Not rs.EOF
       cboMes.AddItem rs.Fields(0) & space(50) & rs.Fields(1)
       rs.MoveNext
    Wend
    Call CargaTipoCambioSBS(pdFecha, pdFecha, pnTipo)
    CentraForm Me
    Me.Show 1
End Sub

Public Function CalculaFechaPeriodSel() As Date
Dim pdFecha As Date
Dim pdFechaFin As Date
Dim psMes As String
Dim psAnioNew As String

    If CInt(cboMes.ListIndex) + 1 < 10 Then
        psMes = "0" & CStr(CInt(cboMes.ListIndex) + 1)
    Else
        psMes = CStr(CInt(cboMes.ListIndex) + 1)
    End If
    psAnioNew = txtAnio
    pdFecha = "01" & "/" & psMes & "/" & psAnioNew
    CalculaFechaPeriodSel = pdFecha
    
End Function '*****NAGL 20170719

Private Sub chkRangFec_Click()
Dim Carg As Integer
    If chkRangFec.value = 1 Then
       pnTipo = 0
       Carg = 1
       txtFecIni.Enabled = True
       txtFecFin.Enabled = True
       txtFecIni.Text = gdFecSis - 1
       txtFecFin.Text = gdFecSis - 1
       txtFecCentral.Text = "__/__/____"
       txtFecCentral.Enabled = False
       flxTipoCambioSBS.Enabled = False
       cmdAgregar.Enabled = False
       cmdQuitar.Enabled = False
       cmdGuardar.Enabled = False
       cmdBuscar.Enabled = False
       Call CargaTipoCambioSBS(txtFecIni, txtFecFin, pnTipo, Carg)
       txtFecIni.SetFocus
    Else
       pnTipo = 2
       txtFecIni.Text = "__/__/____"
       txtFecFin.Text = "__/__/____"
       txtFecIni.Enabled = False
       txtFecFin.Enabled = False
       txtFecCentral.Text = gdFecSis - 1
       txtFecCentral.Enabled = True
       cboMes.Enabled = True
       txtAnio.Enabled = True
       cmdAgregar.Enabled = True
       cmdQuitar.Enabled = True
       cmdGuardar.Enabled = True
       cmdBuscar.Enabled = True
       flxTipoCambioSBS.Enabled = True
       Call CargaTipoCambioSBS(txtFecCentral, txtFecCentral, pnTipo)
       txtFecCentral.SetFocus
    End If
End Sub

Private Sub cmdGenerar_Click()
Dim pdFechaIni As Date
Dim pdFechaFin As Date
Dim pnOptRangFec As Integer, pnOptBal As Integer

pnOptRangFec = chkRangFec.value
pnOptBal = chkDatoBal.value
        If (pnOptRangFec = 1) Then
            If ValFecha(txtFecIni) And ValFecha(txtFecFin) Then
                pnTipo = 0
                pdFechaIni = txtFecIni.Text
                pdFechaFin = txtFecFin.Text
                If ValRegTipoCambioSBS(pdFechaIni, pdFechaFin, pnTipo) Then 'Verifica si existen datos TC en el Rango de Fechas Especificado
                        Call GenerarAnexoVarCambBackTest(pdFechaIni, pdFechaFin, pnOptBal)
                End If
            End If
        Else
            If ValFecha(txtFecCentral) Then
                pnTipo = 1
                pdFechaFin = txtFecCentral.Text
                pdFechaIni = DateAdd("D", 1, DateAdd("D", -Day(pdFechaFin), pdFechaFin))
                If ValRegTipoCambioSBS(pdFechaIni, pdFechaFin, pnTipo) Then 'Verifica si existen datos TC a la fecha Ingresada
                    'If MsgBox("Está seguro de que guardó los Datos Ingresados?", vbInformation + vbYesNo, "Atención") = vbNo Then
                        'txtFecCentral.SetFocus
                        'Exit Sub
                    'Else
                        Call GenerarAnexoVarCambBackTest(pdFechaIni, pdFechaFin, pnOptBal)
                    'End If
                End If
            End If
        End If
End Sub

Private Sub GenerarAnexoVarCambBackTest(pdFechaIni As Date, pdFechaFin As Date, pnOptBal As Integer)
Dim fs As Scripting.FileSystemObject
Dim lbExisteHoja As Boolean
Dim lsArchivo1 As String
Dim lsNomHoja  As String
Dim lsArchivo As String
Dim xlsAplicacion As Excel.Application
Dim xlsLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet
Dim TituloProgress As String
Dim MensajeProgress As String
Dim oBarra As clsProgressBar
Dim nprogress As Integer

Dim rsNeg As New ADODB.Recordset
Dim lilineas As Long
Dim nCorrelativo As Long
Dim liInicio As Long
Dim lsCadena() As String
Dim TCC As String, TCV As String, TCProm As String, TCPosGl As String, PSGlb As String
Dim DesvEst As String, IntConf As String, PlazLiquid As String, Rend As String, Var As String, PatrimEfec As String, GanPer As String
Dim CantVarSBS As Long, CantVarNeg As Long
ReDim lsCadena(2)

On Error GoTo GeneraExcelErr

    Set oBarra = New clsProgressBar
    Unload Me
    oBarra.ShowForm frmReportes
    oBarra.Max = 100
    nprogress = 0
    oBarra.Progress nprogress, "Anexo: VAR Cambiario y Análisis para BackTesting", "GENERANDO EL ARCHIVO", "", vbBlue
    TituloProgress = "Anexo: VAR Cambiario y Análisis para BackTesting"
    MensajeProgress = "GENERANDO EL ARCHIVO"

    Set fs = New Scripting.FileSystemObject
    Set xlsAplicacion = New Excel.Application
    lsArchivo = "AnxVarCambyBackTest"
    'Primera Hoja ******************************************************
    'CON RESPECTO AL ESCENARIO TIPO CAMBIO SBS
    lsNomHoja = "VAR_SBS"
    '*******************************************************************
    lsArchivo1 = "\spooler\ANEXO_VarCambYBackTest_" & gsCodUser & "_" & Format(pdFecha, "YYYYMMDD") & "_" & Format$(Time(), "HHMMSS") & ".xlsx"
    
    If fs.FileExists(App.path & "\FormatoCarta\" & lsArchivo & ".xlsx") Then
        Set xlsLibro = xlsAplicacion.Workbooks.Open(App.path & "\FormatoCarta\" & lsArchivo & ".xlsx")
    Else
        MsgBox "No Existe Plantilla en Carpeta FormatoCarta, Consulte con el Area de  TI", vbInformation, "Advertencia"
        Exit Sub
    End If
    For Each xlHoja1 In xlsLibro.Worksheets
       If xlHoja1.Name = lsNomHoja Then
            xlHoja1.Activate
         lbExisteHoja = True
        Exit For
       End If
    Next
    If lbExisteHoja = False Then
        Set xlHoja1 = xlsLibro.Worksheets
        xlHoja1.Name = lsNomHoja
    End If
    
    lilineas = 5
    liInicio = lilineas
    nCorrelativo = 1
    oBarra.Progress 10, TituloProgress, MensajeProgress, "", vbBlue
    nprogress = 10
    Set rs = DAnxRies.DevueleReporteVarCambYTipoCambioSBS(pdFechaIni, pdFechaFin, "VarSBS", pnOptBal)
    CantVarSBS = rs.RecordCount
    If Not (rs.EOF And rs.BOF) Then
        Do While Not rs.EOF
        
                xlHoja1.Cells(lilineas, 2) = Format(nCorrelativo, "#,##0") 'Correlativo
                xlHoja1.Cells(lilineas, 3) = Format(rs!dFecha, "mm/dd/yyyy") 'Fecha
                
                xlHoja1.Cells(lilineas, 4) = Format(rs!TCCompra, "#,##0.000") 'TCCompra
                TCC = xlHoja1.Range(xlHoja1.Cells(lilineas, 4), xlHoja1.Cells(lilineas, 4)).Address(False, False)
                
                xlHoja1.Cells(lilineas, 5) = Format(rs!TCVenta, "#,##0.000") 'TCVenta
                TCV = xlHoja1.Range(xlHoja1.Cells(lilineas, 5), xlHoja1.Cells(lilineas, 5)).Address(False, False)
                
                xlHoja1.Cells(lilineas, 6) = Format(rs!PGlobal, gsFormatoNumeroView) 'PosicionGlobal
                PSGlb = xlHoja1.Range(xlHoja1.Cells(lilineas, 6), xlHoja1.Cells(lilineas, 6)).Address(False, False)
                
                xlHoja1.Cells(lilineas, 7).Formula = "=" & "If" & "(" & TCC & "=" & "0" & "," & "0" & "," & TCC & "*" & PSGlb & ")"
                xlHoja1.Cells(lilineas, 7).NumberFormat = "#,###0.00" 'PosicionGlobalTC
                TCPosGl = xlHoja1.Range(xlHoja1.Cells(lilineas, 7), xlHoja1.Cells(lilineas, 7)).Address(False, False)
                
                xlHoja1.Cells(lilineas, 8) = Format(rs!PatrimEfectivo, "#,##0.00") 'PatrimonioEfectivo
                PatrimEfec = xlHoja1.Range(xlHoja1.Cells(lilineas, 8), xlHoja1.Cells(lilineas, 8)).Address(False, False)
                
                lsCadena(1) = xlHoja1.Range(xlHoja1.Cells(lilineas, 4), xlHoja1.Cells(lilineas, 5)).Address(False, False)
                xlHoja1.Cells(lilineas, 9).Formula = "=" & "Average" & "(" & lsCadena(1) & ")" 'TipCambProm
                xlHoja1.Cells(lilineas, 9).NumberFormat = "#,###0.000"
                TCProm = xlHoja1.Range(xlHoja1.Cells(lilineas, 9), xlHoja1.Cells(lilineas, 9)).Address(False, False)
                
                If lilineas >= 15 Then 'Rendimiento
                    xlHoja1.Cells(lilineas, 10).Formula = "=" & "If" & "(" & xlHoja1.Range(xlHoja1.Cells(lilineas, 9), xlHoja1.Cells(lilineas, 9)).Address(False, False) & "<>" & "0" & "," & "LN" & "(" & xlHoja1.Range(xlHoja1.Cells(lilineas, 9), xlHoja1.Cells(lilineas, 9)).Address(False, False) & "/" & xlHoja1.Range(xlHoja1.Cells(lilineas - 10, 9), xlHoja1.Cells(lilineas - 10, 9)).Address(False, False) & ")" & "," & """""" & ")"
                    xlHoja1.Cells(lilineas, 10).NumberFormat = "#,##0.0000%"
                    'xlHoja1.Cells(lilineas, 9) = Application.WorksheetFunction.ImAbs
                Else
                    xlHoja1.Cells(lilineas, 10) = Format(rs!Rendimiento, "#,##0.0000%")
                End If
                    Rend = xlHoja1.Range(xlHoja1.Cells(lilineas, 10), xlHoja1.Cells(lilineas, 10)).Address(False, False)
                    xlHoja1.Range(xlHoja1.Cells(lilineas, 10), xlHoja1.Cells(lilineas, 10)).Interior.ColorIndex = 43
                
                If lilineas >= 256 Then 'Desviación Estandar
                    xlHoja1.Cells(lilineas, 11).Formula = "=" & "If" & "(" & TCProm & "<>" & "0" & "," & "StDev" & "(" & xlHoja1.Range(xlHoja1.Cells(lilineas - 251, 10), xlHoja1.Cells(lilineas, 10)).Address(False, False) & ")" & "," & "0" & ")"
                    xlHoja1.Cells(lilineas, 11).NumberFormat = "#,##0.0000%"
                Else
                    xlHoja1.Cells(lilineas, 11) = Format(rs!DesvEstand, "#,##0.0000%")
                End If
                    DesvEst = xlHoja1.Range(xlHoja1.Cells(lilineas, 11), xlHoja1.Cells(lilineas, 11)).Address(False, False)
                
                xlHoja1.Cells(lilineas, 12).Formula = "=" & "NormSInv" & "(" & xlHoja1.Range(xlHoja1.Cells(4, 12), xlHoja1.Cells(4, 12)).Address(False, False) & ")"
                xlHoja1.Cells(lilineas, 12).NumberFormat = "#,###0.00" 'Intervalo de Confianza
                IntConf = xlHoja1.Range(xlHoja1.Cells(lilineas, 12), xlHoja1.Cells(lilineas, 12)).Address(False, False)
                
                xlHoja1.Cells(lilineas, 13).Formula = "=" & "Sqrt" & "(" & xlHoja1.Range(xlHoja1.Cells(4, 13), xlHoja1.Cells(4, 13)).Address(False, False) & ")"
                xlHoja1.Cells(lilineas, 13).NumberFormat = "#,###0.0000" 'Plazo de Liquid
                PlazLiquid = xlHoja1.Range(xlHoja1.Cells(lilineas, 13), xlHoja1.Cells(lilineas, 13)).Address(False, False)
                
                If DesvEst <> "" Then 'VAR
                    xlHoja1.Cells(lilineas, 14).Formula = "=" & "If" & "(" & "IsError" & "(" & "Abs" & "(" & TCPosGl & ")" & "*" & IntConf & "*" & PlazLiquid & "*" & DesvEst & ")" & "," & "0" & "," & "(" & "Abs" & "(" & TCPosGl & ")" & "*" & IntConf & "*" & PlazLiquid & "*" & DesvEst & ")" & ")"
                    xlHoja1.Cells(lilineas, 14).NumberFormat = "#,###0.00"
                    Var = xlHoja1.Range(xlHoja1.Cells(lilineas, 14), xlHoja1.Cells(lilineas, 14)).Address(False, False)
                    
                    xlHoja1.Cells(lilineas, 18).Formula = "=" & "(" & Var & "/" & PatrimEfec & ")" & "*" & "100"
                    xlHoja1.Cells(lilineas, 18).NumberFormat = "#,###0.00"
                    xlHoja1.Cells(lilineas, 19) = Format(rs!ReqCapital, "#,##0.00")
                    xlHoja1.Cells(lilineas, 20).Formula = "=" & "(" & xlHoja1.Range(xlHoja1.Cells(lilineas, 19), xlHoja1.Cells(lilineas, 19)).Address(False, False) & "/" & PatrimEfec & ")" & "*" & "100"
                    xlHoja1.Cells(lilineas, 20).NumberFormat = "#,###0.00"
                    xlHoja1.Range(xlHoja1.Cells(liInicio, 18), xlHoja1.Cells(lilineas, 20)).HorizontalAlignment = xlCenter
                    
                    xlHoja1.Cells(lilineas, 21).Formula = "=" & Var '10 dayVaR(+)
                    xlHoja1.Cells(lilineas, 21).NumberFormat = "#,###0"
                    xlHoja1.Cells(lilineas, 22).Formula = "=" & "-" & Var '10 dayVaR(-)
                    xlHoja1.Cells(lilineas, 22).NumberFormat = "#,###0"
                End If
                
                'LimPosicionLarge
                xlHoja1.Cells(lilineas, 15).Formula = "=" & "If" & "(" & TCPosGl & ">" & "1" & "," & TCPosGl & "," & """""" & ")"
                xlHoja1.Cells(lilineas, 15).NumberFormat = "#,###0.00"
                
                'LimPosicionSmall
                xlHoja1.Cells(lilineas, 16).Formula = "=" & "If" & "(" & TCPosGl & "<" & "1" & "," & TCPosGl & "," & """""" & ")"
                xlHoja1.Cells(lilineas, 16).NumberFormat = "#,###0.00"
                xlHoja1.Range(xlHoja1.Cells(lilineas, 16), xlHoja1.Cells(lilineas, 16)).Font.Color = vbRed
                
                If Rend <> "" Then 'Gan o Perd Efect
                    xlHoja1.Cells(lilineas, 17).Formula = "=" & "If" & "(" & "IsError" & "(" & TCPosGl & "*" & Rend & ")" & "," & """""" & "," & "(" & TCPosGl & "*" & Rend & ")" & ")"
                    xlHoja1.Cells(lilineas, 17).NumberFormat = "#,###0.00"
                    GanPer = xlHoja1.Range(xlHoja1.Cells(lilineas, 17), xlHoja1.Cells(lilineas, 17)).Address(False, False)
                End If
                
                If Rend <> "" And Var <> "" Then 'Para la Excepción
                    xlHoja1.Cells(lilineas, 23).Formula = "=" & "If" & "(" & "ABS" & "(" & GanPer & ")" & ">" & Var & "," & """Yes""" & "," & """No""" & ")"
                    xlHoja1.Cells(lilineas, 23).NumberFormat = "#,###0.00"
                    xlHoja1.Cells(lilineas, 23).Font.Bold = True
                    xlHoja1.Range(xlHoja1.Cells(liInicio, 23), xlHoja1.Cells(lilineas, 23)).HorizontalAlignment = xlCenter
                End If
                
                If rs!FinMes = "SI" Then
                    xlHoja1.Range(xlHoja1.Cells(lilineas, 2), xlHoja1.Cells(lilineas, 23)).Interior.ColorIndex = 44
                End If
                
                If lilineas = CantVarSBS + 4 Then
                   xlHoja1.Cells(8, 25) = "Al" & " " & ArmaFecha(rs!dFecha)
                   xlHoja1.Cells(12, 26).Formula = "=" & "+" & TCPosGl & "/" & "1000" & "-" & "1"
                   xlHoja1.Cells(12, 26).NumberFormat = "#,###0"
                   xlHoja1.Cells(12, 27).Formula = "=" & "+" & DesvEst
                   xlHoja1.Cells(12, 27).NumberFormat = "#,###0.000000000"
                   xlHoja1.Cells(13, 31).Formula = "=" & "+" & Var & "/" & "1000"
                   xlHoja1.Cells(13, 31).NumberFormat = "#,###0.00"
                   xlHoja1.Cells(14, 31).Formula = "=" & "+" & xlHoja1.Range(xlHoja1.Cells(13, 31), xlHoja1.Cells(13, 31)).Address(False, False) & "*" & "3"
                   xlHoja1.Cells(14, 31).NumberFormat = "#,###0.00"
                   xlHoja1.Cells(15, 31).Formula = "=" & "+" & xlHoja1.Range(xlHoja1.Cells(14, 31), xlHoja1.Cells(14, 31)).Address(False, False) & "/" & PatrimEfec & "*" & "1000"
                   xlHoja1.Cells(15, 31).NumberFormat = "#,###0.00%"
                   xlHoja1.Cells(12, 28).Formula = "=" & "+" & xlHoja1.Range(xlHoja1.Cells(13, 31), xlHoja1.Cells(13, 31)).Address(False, False)
                   xlHoja1.Cells(12, 28).NumberFormat = "#,###0.00"
                End If
                
                If Len(CStr(lilineas / IIf(Round((CantVarSBS / 10)) < 1, 1, Round((CantVarSBS / 10))))) = Len(Replace(CStr(lilineas / IIf(Round((CantVarSBS / 10)) < 1, 1, Round((CantVarSBS / 10)))), ".", "")) And nprogress < 45 Then
                    oBarra.Progress nprogress + 5, TituloProgress, MensajeProgress, "", vbBlue
                    nprogress = nprogress + 5
                End If
                
                lilineas = lilineas + 1
                nCorrelativo = nCorrelativo + 1
                rs.MoveNext
        Loop
        xlHoja1.Range(xlHoja1.Cells(liInicio, 2), xlHoja1.Cells(lilineas, 10)).HorizontalAlignment = xlCenter
        xlHoja1.Range(xlHoja1.Cells(liInicio, 12), xlHoja1.Cells(lilineas, 13)).HorizontalAlignment = xlCenter
        xlHoja1.Range(xlHoja1.Cells(liInicio, 2), xlHoja1.Cells(lilineas, 10)).Font.Size = 9
        xlHoja1.Range(xlHoja1.Cells(liInicio, 2), xlHoja1.Cells(lilineas, 10)).Font.Name = "Arial"
        'xlHoja1.Range(xlHoja1.Cells(liInicio, 2), xlHoja1.Cells(lilineas, 10)).Interior.ColorIndex = 2
        xlHoja1.Range(xlHoja1.Cells(liInicio, 7), xlHoja1.Cells(lilineas, 7)).Font.Color = vbRed
    End If
   
   'CON RESPECTO AL ESCENARIO TIPO DE CAMBIO DEL NEGOCIO
    lsNomHoja = "VAR_NEG"
    For Each xlHoja1 In xlsLibro.Worksheets
       If xlHoja1.Name = lsNomHoja Then
            xlHoja1.Activate
         lbExisteHoja = True
        Exit For
       End If
    Next
    If lbExisteHoja = False Then
        Set xlHoja1 = xlsLibro.Worksheets
        xlHoja1.Name = lsNomHoja
    End If
    Set rs = Nothing
    lilineas = 5
    liInicio = lilineas
    nCorrelativo = 1
    Set rs = DAnxRies.DevueleReporteVarCambYTipoCambioSBS(pdFechaIni, pdFechaFin, "VarNeg", pnOptBal)
    CantVarNeg = rs.RecordCount
    If Not (rs.EOF And rs.BOF) Then
        Do While Not rs.EOF
        
                xlHoja1.Cells(lilineas, 2) = Format(nCorrelativo, "#,##0") 'Correlativo
                xlHoja1.Cells(lilineas, 3) = Format(rs!dFecha, "mm/dd/yyyy") 'Fecha
                
                xlHoja1.Cells(lilineas, 4) = Format(rs!TCCompra, "#,##0.000") 'TCCompra
                TCC = xlHoja1.Range(xlHoja1.Cells(lilineas, 4), xlHoja1.Cells(lilineas, 4)).Address(False, False)
                
                xlHoja1.Cells(lilineas, 5) = Format(rs!TCVenta, "#,##0.000") 'TCVenta
                TCV = xlHoja1.Range(xlHoja1.Cells(lilineas, 5), xlHoja1.Cells(lilineas, 5)).Address(False, False)
                
                xlHoja1.Cells(lilineas, 6) = Format(rs!PGlobal, gsFormatoNumeroView) 'PosicionGlobal
                PSGlb = xlHoja1.Range(xlHoja1.Cells(lilineas, 6), xlHoja1.Cells(lilineas, 6)).Address(False, False)
                
                xlHoja1.Cells(lilineas, 7).Formula = "=" & "If" & "(" & TCC & "=" & "0" & "," & "0" & "," & TCC & "*" & PSGlb & ")"
                xlHoja1.Cells(lilineas, 7).NumberFormat = "#,###0.00" 'PosicionGlobalTC
                TCPosGl = xlHoja1.Range(xlHoja1.Cells(lilineas, 7), xlHoja1.Cells(lilineas, 7)).Address(False, False)
                
                xlHoja1.Cells(lilineas, 8) = Format(rs!PatrimEfectivo, "#,##0.00") 'PatrimonioEfectivo
                PatrimEfec = xlHoja1.Range(xlHoja1.Cells(lilineas, 8), xlHoja1.Cells(lilineas, 8)).Address(False, False)
                
                lsCadena(1) = xlHoja1.Range(xlHoja1.Cells(lilineas, 4), xlHoja1.Cells(lilineas, 5)).Address(False, False)
                xlHoja1.Cells(lilineas, 9).Formula = "=" & "Average" & "(" & lsCadena(1) & ")" 'TipCambProm
                xlHoja1.Cells(lilineas, 9).NumberFormat = "#,###0.000"
                TCProm = xlHoja1.Range(xlHoja1.Cells(lilineas, 9), xlHoja1.Cells(lilineas, 9)).Address(False, False)
                
                If lilineas >= 15 Then 'Rendimiento
                    xlHoja1.Cells(lilineas, 10).Formula = "=" & "If" & "(" & xlHoja1.Range(xlHoja1.Cells(lilineas, 9), xlHoja1.Cells(lilineas, 9)).Address(False, False) & "<>" & "0" & "," & "LN" & "(" & xlHoja1.Range(xlHoja1.Cells(lilineas, 9), xlHoja1.Cells(lilineas, 9)).Address(False, False) & "/" & xlHoja1.Range(xlHoja1.Cells(lilineas - 10, 9), xlHoja1.Cells(lilineas - 10, 9)).Address(False, False) & ")" & "," & """""" & ")"
                    xlHoja1.Cells(lilineas, 10).NumberFormat = "#,##0.0000%"
                    'xlHoja1.Cells(lilineas, 9) = Application.WorksheetFunction.ImAbs
                Else
                    xlHoja1.Cells(lilineas, 10) = Format(rs!Rendimiento, "#,##0.0000%")
                End If
                    Rend = xlHoja1.Range(xlHoja1.Cells(lilineas, 10), xlHoja1.Cells(lilineas, 10)).Address(False, False)
                    xlHoja1.Range(xlHoja1.Cells(lilineas, 10), xlHoja1.Cells(lilineas, 10)).Interior.ColorIndex = 43
                
                If lilineas >= 256 Then 'Desviación Estandar
                    xlHoja1.Cells(lilineas, 11).Formula = "=" & "If" & "(" & TCProm & "<>" & "0" & "," & "StDev" & "(" & xlHoja1.Range(xlHoja1.Cells(lilineas - 251, 10), xlHoja1.Cells(lilineas, 10)).Address(False, False) & ")" & "," & "0" & ")"
                    xlHoja1.Cells(lilineas, 11).NumberFormat = "#,##0.0000%"
                Else
                    xlHoja1.Cells(lilineas, 11) = Format(rs!DesvEstand, "#,##0.0000%")
                End If
                    DesvEst = xlHoja1.Range(xlHoja1.Cells(lilineas, 11), xlHoja1.Cells(lilineas, 11)).Address(False, False)
                
                xlHoja1.Cells(lilineas, 12).Formula = "=" & "NormSInv" & "(" & xlHoja1.Range(xlHoja1.Cells(4, 12), xlHoja1.Cells(4, 12)).Address(False, False) & ")"
                xlHoja1.Cells(lilineas, 12).NumberFormat = "#,###0.00" 'Intervalo de Confianza
                IntConf = xlHoja1.Range(xlHoja1.Cells(lilineas, 12), xlHoja1.Cells(lilineas, 12)).Address(False, False)
                
                xlHoja1.Cells(lilineas, 13).Formula = "=" & "Sqrt" & "(" & xlHoja1.Range(xlHoja1.Cells(4, 13), xlHoja1.Cells(4, 13)).Address(False, False) & ")"
                xlHoja1.Cells(lilineas, 13).NumberFormat = "#,###0.0000" 'Plazo de Liquid
                PlazLiquid = xlHoja1.Range(xlHoja1.Cells(lilineas, 13), xlHoja1.Cells(lilineas, 13)).Address(False, False)
                
                If DesvEst <> "" Then 'VAR
                    xlHoja1.Cells(lilineas, 14).Formula = "=" & "If" & "(" & "IsError" & "(" & "Abs" & "(" & TCPosGl & ")" & "*" & IntConf & "*" & PlazLiquid & "*" & DesvEst & ")" & "," & "0" & "," & "(" & "Abs" & "(" & TCPosGl & ")" & "*" & IntConf & "*" & PlazLiquid & "*" & DesvEst & ")" & ")"
                    xlHoja1.Cells(lilineas, 14).NumberFormat = "#,###0.00"
                    Var = xlHoja1.Range(xlHoja1.Cells(lilineas, 14), xlHoja1.Cells(lilineas, 14)).Address(False, False)
                    
                    xlHoja1.Cells(lilineas, 18).Formula = "=" & "(" & Var & "/" & PatrimEfec & ")" & "*" & "100"
                    xlHoja1.Cells(lilineas, 18).NumberFormat = "#,###0.00"
                    xlHoja1.Cells(lilineas, 19) = Format(rs!ReqCapital, "#,##0.00")
                    xlHoja1.Cells(lilineas, 20).Formula = "=" & "(" & xlHoja1.Range(xlHoja1.Cells(lilineas, 19), xlHoja1.Cells(lilineas, 19)).Address(False, False) & "/" & PatrimEfec & ")" & "*" & "100"
                    xlHoja1.Cells(lilineas, 20).NumberFormat = "#,###0.00"
                    xlHoja1.Range(xlHoja1.Cells(liInicio, 18), xlHoja1.Cells(lilineas, 20)).HorizontalAlignment = xlCenter
                    
                    xlHoja1.Cells(lilineas, 21).Formula = "=" & Var '10 dayVaR(+)
                    xlHoja1.Cells(lilineas, 21).NumberFormat = "#,###0"
                    xlHoja1.Cells(lilineas, 22).Formula = "=" & "-" & Var '10 dayVaR(-)
                    xlHoja1.Cells(lilineas, 22).NumberFormat = "#,###0"
                End If
                
                'LimPosicionLarge
                xlHoja1.Cells(lilineas, 15).Formula = "=" & "If" & "(" & TCPosGl & ">" & "1" & "," & TCPosGl & "," & """""" & ")"
                xlHoja1.Cells(lilineas, 15).NumberFormat = "#,###0.00"
                
                'LimPosicionSmall
                xlHoja1.Cells(lilineas, 16).Formula = "=" & "If" & "(" & TCPosGl & "<" & "1" & "," & TCPosGl & "," & """""" & ")"
                xlHoja1.Cells(lilineas, 16).NumberFormat = "#,###0.00"
                xlHoja1.Range(xlHoja1.Cells(lilineas, 16), xlHoja1.Cells(lilineas, 16)).Font.Color = vbRed
                
                If Rend <> "" Then 'Gan o Perd Efect
                    xlHoja1.Cells(lilineas, 17).Formula = "=" & "If" & "(" & "IsError" & "(" & TCPosGl & "*" & Rend & ")" & "," & """""" & "," & "(" & TCPosGl & "*" & Rend & ")" & ")"
                    xlHoja1.Cells(lilineas, 17).NumberFormat = "#,###0.00"
                    GanPer = xlHoja1.Range(xlHoja1.Cells(lilineas, 17), xlHoja1.Cells(lilineas, 17)).Address(False, False)
                End If
                
                If Rend <> "" And Var <> "" Then 'Para la Excepción
                    xlHoja1.Cells(lilineas, 23).Formula = "=" & "If" & "(" & "ABS" & "(" & GanPer & ")" & ">" & Var & "," & """Yes""" & "," & """No""" & ")"
                    xlHoja1.Cells(lilineas, 23).NumberFormat = "#,###0.00"
                    xlHoja1.Cells(lilineas, 23).Font.Bold = True
                    xlHoja1.Range(xlHoja1.Cells(liInicio, 23), xlHoja1.Cells(lilineas, 23)).HorizontalAlignment = xlCenter
                End If
                
                If rs!FinMes = "SI" Then
                    xlHoja1.Range(xlHoja1.Cells(lilineas, 2), xlHoja1.Cells(lilineas, 23)).Interior.ColorIndex = 44
                End If
                
                If lilineas = CantVarNeg + 4 Then
                   xlHoja1.Cells(8, 25) = "Al" & " " & ArmaFecha(rs!dFecha)
                   xlHoja1.Cells(12, 26).Formula = "=" & "+" & TCPosGl & "/" & "1000" & "-" & "1"
                   xlHoja1.Cells(12, 26).NumberFormat = "#,###0"
                   xlHoja1.Cells(12, 27).Formula = "=" & "+" & DesvEst
                   xlHoja1.Cells(12, 27).NumberFormat = "#,###0.000000000"
                   xlHoja1.Cells(13, 31).Formula = "=" & "+" & Var & "/" & "1000"
                   xlHoja1.Cells(13, 31).NumberFormat = "#,###0.00"
                   xlHoja1.Cells(14, 31).Formula = "=" & "+" & xlHoja1.Range(xlHoja1.Cells(13, 31), xlHoja1.Cells(13, 31)).Address(False, False) & "*" & "3"
                   xlHoja1.Cells(14, 31).NumberFormat = "#,###0.00"
                   xlHoja1.Cells(15, 31).Formula = "=" & "+" & xlHoja1.Range(xlHoja1.Cells(14, 31), xlHoja1.Cells(14, 31)).Address(False, False) & "/" & PatrimEfec & "*" & "1000"
                   xlHoja1.Cells(15, 31).NumberFormat = "#,###0.00%"
                   xlHoja1.Cells(12, 28).Formula = "=" & "+" & xlHoja1.Range(xlHoja1.Cells(13, 31), xlHoja1.Cells(13, 31)).Address(False, False)
                   xlHoja1.Cells(12, 28).NumberFormat = "#,###0.00"
                End If
                
                 If Len(CStr(lilineas / IIf(Round((CantVarNeg / 10)) < 1, 1, Round((CantVarNeg / 10))))) = Len(Replace(CStr(lilineas / IIf(Round((CantVarNeg / 10)) < 1, 1, Round((CantVarNeg / 10)))), ".", "")) And nprogress < 95 Then
                    oBarra.Progress nprogress + 5, TituloProgress, MensajeProgress, "", vbBlue
                    nprogress = nprogress + 5
                End If
                
                lilineas = lilineas + 1
                nCorrelativo = nCorrelativo + 1
                rs.MoveNext
        Loop
        xlHoja1.Range(xlHoja1.Cells(liInicio, 2), xlHoja1.Cells(lilineas, 10)).HorizontalAlignment = xlCenter
        xlHoja1.Range(xlHoja1.Cells(liInicio, 12), xlHoja1.Cells(lilineas, 13)).HorizontalAlignment = xlCenter
        xlHoja1.Range(xlHoja1.Cells(liInicio, 2), xlHoja1.Cells(lilineas, 10)).Font.Size = 9
        xlHoja1.Range(xlHoja1.Cells(liInicio, 2), xlHoja1.Cells(lilineas, 10)).Font.Name = "Arial"
        'xlHoja1.Range(xlHoja1.Cells(liInicio, 2), xlHoja1.Cells(lilineas, 10)).Interior.ColorIndex = 2
        xlHoja1.Range(xlHoja1.Cells(liInicio, 7), xlHoja1.Cells(lilineas, 7)).Font.Color = vbRed
    End If
   
    oBarra.Progress 100, "Anexo: Modelo de Valor en Riesgo y Análisis para BackTesting", "Generación Terminada", "", vbBlue
    oBarra.CloseForm frmReportes
    Set oBarra = Nothing

    xlHoja1.SaveAs App.path & lsArchivo1
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlHoja1 = Nothing

Exit Sub
GeneraExcelErr:
    MsgBox Err.Description, vbInformation, "Aviso"
    Exit Sub
End Sub

Private Function LlenarRsflxTipoCambioSBSDet(ByVal feControl As FlexEdit) As ADODB.Recordset

 Dim rsVD As New ADODB.Recordset
 Dim nIndex As Integer
  If feControl.Rows >= 2 Then
        If feControl.TextMatrix(nIndex, 1) = "" Then
            Exit Function
        End If
            rsVD.CursorType = adOpenStatic
            rsVD.Fields.Append "pdFecCambSBS", adDate, adFldIsNullable
            rsVD.Fields.Append "TCCompra", adDouble, adFldIsNullable
            rsVD.Fields.Append "TCCVenta", adDouble, adFldIsNullable
            rsVD.Open
            
        For nIndex = 1 To feControl.Rows - 1
            rsVD.AddNew
            rsVD.Fields("pdFecCambSBS") = feControl.TextMatrix(nIndex, 1)
            rsVD.Fields("TCCompra") = feControl.TextMatrix(nIndex, 2)
            rsVD.Fields("TCCVenta") = feControl.TextMatrix(nIndex, 3)
            rsVD.Update
            rsVD.MoveFirst
        Next
    End If
    Set LlenarRsflxTipoCambioSBSDet = rsVD
End Function

Private Sub MuestraTipoCambioSBS(pdFecha As Date, pdFechaFin As Date, pnTipo As Integer)
Set rs = DAnxRies.DevuelveTipoCambioSBS(pdFecha, pdFechaFin, pnTipo)
flxTipoCambioSBS.Clear
FormateaFlex flxTipoCambioSBS
    If Not (rs.EOF And rs.BOF) Then
        For X = 1 To rs.RecordCount
            flxTipoCambioSBS.AdicionaFila
            flxTipoCambioSBS.TextMatrix(X, 1) = Format(rs!dFecCambSBS, "dd/mm/yyyy")
            flxTipoCambioSBS.TextMatrix(X, 2) = Format(rs!TCCompra, "#,##0.000")
            flxTipoCambioSBS.TextMatrix(X, 3) = Format(rs!TCVenta, "#,##0.000")
            rs.MoveNext
        Next
    End If
End Sub

Private Function ValRegTipoCambioSBS(pdFechaIni As Date, pdFechaFin As Date, pnTipo As Integer) As Boolean
Dim DAnxVal As New DAnexoRiesgos
Dim psRegistro As String
 psRegistro = DAnxVal.ObtieneValRegTipoCambioSBS(pdFechaIni, pdFechaFin, pnTipo)
    If psRegistro = "NO" Then
        If pnTipo = 0 Then
                If MsgBox("No existen Datos en el Rango Ingresado, Desea continuar?", vbInformation + vbYesNo, "Atención") = vbNo Then
                    txtFecIni.SetFocus
                    Exit Function
                End If
        ElseIf pnTipo = 1 Then
                MsgBox "No existen Datos en el Periodo Seleccionado", vbOKOnly + vbInformation, "Atención"
                flxTipoCambioSBS.Clear
                FormateaFlex flxTipoCambioSBS
                cboMes.SetFocus
                Exit Function
        Else
                MsgBox "No existen Datos en el Mes de la Fecha Ingresada", vbOKOnly + vbInformation, "Atención"
                txtFecCentral.SetFocus
                Exit Function
        End If
    End If
ValRegTipoCambioSBS = True
End Function

Private Sub cmdBuscar_Click()
Dim pdFecha As Date
pnTipo = 1
pdFecha = CalculaFechaPeriodSel
Call CargaTipoCambioSBS(pdFecha, pdFecha, pnTipo)
cmdGenerar.Enabled = True
chkRangFec.Enabled = True
txtFecCentral.Enabled = True
End Sub

Private Sub CargaTipoCambioSBS(pdFecha As Date, pdFechaFin As Date, pnTipo As Integer, Optional Carg As Integer = 0)
    If pnTipo = 0 Then 'Rango de Fechas
            cboMes.ListIndex = CInt(Month(pdFechaFin)) - 1
            txtAnio.Text = CInt(Year(pdFechaFin))
            cboMes.Enabled = False
            txtAnio.Enabled = False
            If Carg = 0 Then 'Pendiente
                If ValRegTipoCambioSBS(pdFecha, pdFechaFin, pnTipo) Then
                   Call MuestraTipoCambioSBS(pdFecha, pdFechaFin, pnTipo)
                End If
            Else
                flxTipoCambioSBS.Clear
                FormateaFlex flxTipoCambioSBS
            End If
            
    ElseIf pnTipo = 1 Then 'Para Cargar Datos con el Mes Correspondiente.
            If ValRegTipoCambioSBS(pdFecha, pdFechaFin, pnTipo) Then
               Call MuestraTipoCambioSBS(pdFecha, pdFechaFin, pnTipo)
            End If
        
    Else  'Para cargar desde el primer dia hasta la fecha Actual (Mes) (pnTipo = 2 -> Para q no tome en cuenta el cboMes)
            cboMes.ListIndex = CInt(Month(pdFecha)) - 1
            txtAnio.Text = CInt(Year(pdFecha))
            Call MuestraTipoCambioSBS(pdFecha, pdFechaFin, pnTipo)
    End If
End Sub

Private Sub cmdGuardar_Click()
    Dim lsMovNro As String
    Dim oCont As New NContFunciones
    If flxTipoCambioSBS.TextMatrix(1, 1) = "" Then
        Exit Sub
    End If
    lsMovNro = oCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    If flxTipoCambioSBS.TextMatrix(flxTipoCambioSBS.row, 1) = "" Or flxTipoCambioSBS.TextMatrix(flxTipoCambioSBS.row, 2) = "" Or flxTipoCambioSBS.TextMatrix(flxTipoCambioSBS.row, 3) = "" Then
        nItem = flxTipoCambioSBS.row
        flxTipoCambioSBS.EliminaFila nItem
    End If
    If MsgBox("¿Desea Guardar los Tipos de Cambio Ingresados..", vbInformation + vbYesNo, "Atención") = vbYes Then
         If ValidarDatos = False Then Exit Sub
         Call DAnxRies.ControlTipoCambioSBS(lsMovNro, LlenarRsflxTipoCambioSBSDet(flxTipoCambioSBS))
         MsgBox "Los datos se guardaron satisfactoriamente.", vbOKOnly + vbInformation, "Atención"
         cmdGenerar.Enabled = True
         chkRangFec.Enabled = True
         txtFecCentral.Enabled = True
    End If
End Sub

Private Sub cmdAgregar_Click()
   'fnFilSel = -1
   cmdGenerar.Enabled = False
   chkRangFec.Enabled = False
   txtFecCentral.Enabled = False
   flxTipoCambioSBS.lbEditarFlex = True
   
   If (flxTipoCambioSBS.TextMatrix(CInt(flxTipoCambioSBS.Rows) - 1, 1) = "" Or flxTipoCambioSBS.TextMatrix(CInt(flxTipoCambioSBS.Rows) - 1, 2) = "" Or flxTipoCambioSBS.TextMatrix(CInt(flxTipoCambioSBS.Rows) - 1, 3) = "") Then
       nItem = CInt(flxTipoCambioSBS.Rows) - 1
       flxTipoCambioSBS.EliminaFila nItem
   End If
   If (flxTipoCambioSBS.TextMatrix(flxTipoCambioSBS.row, 1) <> "" And flxTipoCambioSBS.TextMatrix(flxTipoCambioSBS.row, 2) <> "" And flxTipoCambioSBS.TextMatrix(flxTipoCambioSBS.row, 3) <> "") Then
        If ValidarDatos = False Then Exit Sub
        flxTipoCambioSBS.AdicionaFila
        fnFilSel = flxTipoCambioSBS.Rows - 1
    ElseIf CInt(flxTipoCambioSBS.row) = 1 Then
        flxTipoCambioSBS.AdicionaFila
        fnFilSel = flxTipoCambioSBS.Rows - 1
    Else
        MsgBox "Para continuar, se debe llenar todos los campos obligatoriamente", vbInformation, "Atención"
    End If
   flxTipoCambioSBS.SetFocus
End Sub
Private Function ValidarDatos(Optional CambCell As String = "") As Boolean
Dim i As Integer
Dim Cant As Integer
Cant = 0
    For i = 1 To CInt(flxTipoCambioSBS.Rows) - 1
            If CInt(Month(Trim(flxTipoCambioSBS.TextMatrix(i, 1)))) <> cboMes.ListIndex + 1 Or CInt(Year(Trim(flxTipoCambioSBS.TextMatrix(i, 1)))) <> txtAnio Then
               Cant = Cant + 1
                If Cant >= 2 Then
                    nItem = flxTipoCambioSBS.row
                    MsgBox "El Registro de Tipos de Cambio SBS no corresponde al periodo seleccionado, actualice la Búsqueda !!", vbInformation, "Aviso"
                    flxTipoCambioSBS.EliminaFila nItem
                    cboMes.SetFocus
                    ValidarDatos = False
                    Exit Function
                End If
            End If
    Next i
    
    If (CInt(Month(Trim(flxTipoCambioSBS.TextMatrix(flxTipoCambioSBS.row, 1)))) <> cboMes.ListIndex + 1) Or Year(Trim(flxTipoCambioSBS.TextMatrix(flxTipoCambioSBS.row, 1))) <> txtAnio Then
         nItem = CInt(flxTipoCambioSBS.Rows) - 1
         MsgBox "La fecha " & Trim(flxTipoCambioSBS.TextMatrix(flxTipoCambioSBS.row, 1)) & ", no pertenece al periodo seleccionado !!", vbInformation, "Aviso"
         ValidarDatos = False
         flxTipoCambioSBS.TextMatrix(flxTipoCambioSBS.row, 1) = ""
         flxTipoCambioSBS.SetFocus
         Exit Function
    End If
        
    For i = 1 To CInt(flxTipoCambioSBS.Rows) - 1 'Para Controlar Fechas Repetidas
            If Trim(flxTipoCambioSBS.TextMatrix(i, 1)) = CDate(Trim(flxTipoCambioSBS.TextMatrix(flxTipoCambioSBS.row, 1))) Then
                If i <> flxTipoCambioSBS.row Then
                    If Trim(flxTipoCambioSBS.TextMatrix(i, 1)) = CDate(Trim(flxTipoCambioSBS.TextMatrix(flxTipoCambioSBS.row, 1))) Then
                        nItem = flxTipoCambioSBS.row
                        MsgBox "La fecha " & Trim(flxTipoCambioSBS.TextMatrix(flxTipoCambioSBS.row, 1)) & ", ya ha sido ingresado !!", vbInformation, "Aviso"
                        ValidarDatos = False
                        flxTipoCambioSBS.EliminaFila nItem
                        flxTipoCambioSBS.SetFocus
                        Exit Function
                    End If
                End If
            End If
    Next i
ValidarDatos = True
End Function

Private Sub flxTipoCambioSBS_EnterCell()
    If flxTipoCambioSBS.col = 2 Then
        valorCelda1 = flxTipoCambioSBS.TextMatrix(flxTipoCambioSBS.row, flxTipoCambioSBS.col)
    ElseIf flxTipoCambioSBS.col = 3 Then
        valorCelda2 = flxTipoCambioSBS.TextMatrix(flxTipoCambioSBS.row, flxTipoCambioSBS.col)
    End If
End Sub

Private Sub flxTipoCambioSBS_OnCellChange(pnRow As Long, pnCol As Long)
    Dim ValorNew1 As Currency, ValorNew2 As Currency
    Dim ValorAnterior1 As Currency, ValorAnterior2 As Currency
    Dim ValorFechaAnt As Date
    Dim ValorFechaNew As Date

        If (pnCol = 1) Then
            If ValidarDatos("SC") = False Then Exit Sub 'SC : Salto entre celda
        ElseIf (pnCol = 2) Then
            If (valorCelda1 <> "") Then
                ValorAnterior1 = CDbl(valorCelda1)
            End If
        Else
            If (valorCelda2 <> "") Then
                ValorAnterior2 = CDbl(valorCelda2)
            End If
        End If

    If pnCol = 2 Then
             If (IsNumeric(flxTipoCambioSBS.TextMatrix(pnRow, pnCol)) And Len(flxTipoCambioSBS.TextMatrix(pnRow, pnCol)) < 10) Then
                ValorNew1 = CDbl(flxTipoCambioSBS.TextMatrix(pnRow, pnCol))
                    If (ValorNew1 < 0) Then
                        MsgBox "No se puede asignar un valor Negativo", vbInformation, "Aviso"
                        flxTipoCambioSBS.TextMatrix(pnRow, pnCol) = IIf(ValorAnterior1 = 0, Format(0, "###,##0.00"), ValorAnterior1)
                        Exit Sub
                    Else
                        flxTipoCambioSBS.TextMatrix(pnRow, pnCol) = ValorNew1
                    End If
              Else
                        flxTipoCambioSBS.TextMatrix(pnRow, pnCol) = IIf(ValorAnterior1 = 0, Format(ValorAnterior1, "###,##0.00"), ValorAnterior1)
              End If

    ElseIf pnCol = 3 Then
             If (IsNumeric(flxTipoCambioSBS.TextMatrix(pnRow, pnCol)) And Len(flxTipoCambioSBS.TextMatrix(pnRow, pnCol)) < 10) Then
                ValorNew2 = CDbl(flxTipoCambioSBS.TextMatrix(pnRow, pnCol))
                    If (ValorNew2 < 0) Then
                        MsgBox "No se puede asignar un valor Negativo", vbInformation, "Aviso"
                        flxTipoCambioSBS.TextMatrix(pnRow, pnCol) = IIf(ValorAnterior2 = 0, Format(ValorAnterior2, "###,##0.00"), ValorAnterior2)
                        Exit Sub
                    Else
                        flxTipoCambioSBS.TextMatrix(pnRow, pnCol) = ValorNew2
                    End If
              Else
                        flxTipoCambioSBS.TextMatrix(pnRow, pnCol) = IIf(ValorAnterior2 = 0, Format(ValorAnterior2, "###,##0.00"), ValorAnterior2)
              End If
    End If


End Sub

Private Sub cmdQuitar_Click()
Dim psReg As String
    pnTipo = 3
    If flxTipoCambioSBS.TextMatrix(flxTipoCambioSBS.row, 1) = "" Or flxTipoCambioSBS.TextMatrix(flxTipoCambioSBS.row, 2) = "" Or flxTipoCambioSBS.TextMatrix(flxTipoCambioSBS.row, 3) = "" Then
        nItem = flxTipoCambioSBS.row
        flxTipoCambioSBS.EliminaFila nItem
        Exit Sub
    End If
    nItem = flxTipoCambioSBS.row
    If MsgBox("¿Esta seguro que desea quitar el Tipo de Cambio seleccionado?", vbYesNo + vbQuestion, "Atención") = vbNo Then Exit Sub
    psReg = DAnxRies.ObtieneValRegTipoCambioSBS(Trim(flxTipoCambioSBS.TextMatrix(nItem, 1)), Trim(flxTipoCambioSBS.TextMatrix(nItem, 1)), pnTipo)
    If psReg = "SI" Then
        If DAnxRies.PermiteDeleteModificarTpoCambSBS(Format(CDate(flxTipoCambioSBS.TextMatrix(nItem, 1)), gsFormatoFechaView)) = True Then
            If MsgBox("El Tipo de Cambio pertenece a un mes ya Cerrado, Desea Continuar...", vbInformation + vbYesNo, "Atención") = vbYes Then
                DAnxRies.EliminaTipoCambioSBS (Format(CDate(flxTipoCambioSBS.TextMatrix(nItem, 1)), gsFormatoFechaView))
                flxTipoCambioSBS.EliminaFila nItem
            End If
        Else
                DAnxRies.EliminaTipoCambioSBS (Format(CDate(flxTipoCambioSBS.TextMatrix(nItem, 1)), gsFormatoFechaView))
                flxTipoCambioSBS.EliminaFila nItem
        End If
    Else
         flxTipoCambioSBS.EliminaFila nItem
    End If
End Sub

Private Function ValidaDatos() As Boolean

   If flxTipoCambioSBS.col = 1 Then
        If (flxTipoCambioSBS.TextMatrix(flxTipoCambioSBS.row, 1)) = "" Then
            MsgBox "Falta ingresar la Fecha correspondiente", vbInformation, "Aviso"
            flxTipoCambioSBS.SetFocus
            Exit Function
        End If
        If Not IsDate(flxTipoCambioSBS.TextMatrix(flxTipoCambioSBS.row, 1)) Then
            MsgBox "La fecha Ingresada es incorrecta", vbInformation, "Aviso"
            flxTipoCambioSBS.SetFocus
        End If
    End If

    If flxTipoCambioSBS.col = 2 Or flxTipoCambioSBS.col = 3 Then
        If ValorNew = "" Then
            MsgBox "Falta ingresar el Tipo de Cambio de Compra", vbInformation, "Aviso"
            flxTipoCambioSBS.SetFocus
            Exit Function
        End If

        If Not IsNumeric(ValorNew) Or ValorNew < 0 Then
            MsgBox "El Tipo de Cambio Ingresado es Incorrecto", vbInformation, "Aviso"
            flxTipoCambioSBS.SetFocus
        End If
    End If

    ValidaDatos = True
End Function

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub txtFecCentral_GotFocus()
    fEnfoque txtFecCentral
End Sub

Private Sub txtFecCentral_KeyPress(keyAscii As Integer)
pnTipo = 2
    If keyAscii = 13 Then
        If ValFecha(txtFecCentral) Then
           Call CargaTipoCambioSBS(txtFecCentral, txtFecCentral, pnTipo)
           cmdGenerar.SetFocus
        End If
    End If
End Sub

Private Sub txtFecCentral_LostFocus()
    pnTipo = 2
If ValFecha(txtFecCentral) Then
    Call CargaTipoCambioSBS(txtFecCentral, txtFecCentral, pnTipo)
End If
End Sub

Private Sub txtFecIni_GotFocus()
    fEnfoque txtFecIni
End Sub

Private Sub txtFecIni_KeyPress(keyAscii As Integer)
  If keyAscii = 13 Then
        If ValFecha(txtFecIni) Then
            txtFecFin.SetFocus
        End If
  End If
End Sub

Private Sub txtFecFin_GotFocus()
   fEnfoque txtFecFin
End Sub

Private Sub txtFecFin_KeyPress(keyAscii As Integer)
   If keyAscii = 13 Then
        If ValFecha(txtFecFin) Then
            cmdGenerar.SetFocus
        End If
   End If
End Sub

Private Sub txtAnio_GotFocus()
    fEnfoque txtAnio
End Sub

