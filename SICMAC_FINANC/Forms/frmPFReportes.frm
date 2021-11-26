VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmPFReportes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte: Resumen de Plazo Fijo"
   ClientHeight    =   2805
   ClientLeft      =   3255
   ClientTop       =   3135
   ClientWidth     =   8175
   Icon            =   "frmPFReportes.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   8175
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FraCuentaDesde 
      Caption         =   "Cuenta Desde"
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
      Height          =   1035
      Left            =   90
      TabIndex        =   8
      Top             =   840
      Width           =   7965
      Begin VB.CheckBox chktodos 
         Caption         =   "&Todos"
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   7020
         TabIndex        =   13
         Top             =   -30
         Value           =   1  'Checked
         Width           =   765
      End
      Begin Sicmact.TxtBuscar txtCtaIFDesde 
         Height          =   315
         Left            =   1155
         TabIndex        =   9
         Top             =   255
         Width           =   2820
         _ExtentX        =   4974
         _ExtentY        =   556
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
         sTitulo         =   ""
         ForeColor       =   -2147483635
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Institución :"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   180
         TabIndex        =   14
         Top             =   330
         Width           =   810
      End
      Begin VB.Label lblDescCtaDesde 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H8000000D&
         Height          =   300
         Left            =   1155
         TabIndex        =   12
         Top             =   637
         Width           =   6630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Cuenta :"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   135
         TabIndex        =   11
         Top             =   690
         Width           =   960
      End
      Begin VB.Label lblDescIFDesde 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H8000000D&
         Height          =   300
         Left            =   3990
         TabIndex        =   10
         Top             =   255
         Width           =   3795
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   735
      Left            =   90
      TabIndex        =   6
      Top             =   45
      Width           =   4785
      Begin MSMask.MaskEdBox txtfecha 
         Height          =   330
         Left            =   915
         TabIndex        =   0
         Top             =   240
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   582
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Fecha  :"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   255
         TabIndex        =   7
         Top             =   285
         Width           =   585
      End
   End
   Begin MSComctlLib.ProgressBar Barra 
      Height          =   180
      Left            =   2550
      TabIndex        =   5
      Top             =   2610
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   318
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   360
      Left            =   4005
      TabIndex        =   2
      Top             =   2040
      Width           =   1440
   End
   Begin VB.CommandButton cmdGenerar 
      Caption         =   "&Generar"
      Height          =   360
      Left            =   2565
      TabIndex        =   1
      Top             =   2040
      Width           =   1440
   End
   Begin MSComctlLib.StatusBar Estado 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      TabIndex        =   3
      Top             =   2565
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   423
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   9878
            MinWidth        =   9878
         EndProperty
      EndProperty
   End
   Begin VB.OLE OleExcel 
      Height          =   465
      Left            =   195
      SizeMode        =   1  'Stretch
      TabIndex        =   4
      Top             =   3645
      Visible         =   0   'False
      Width           =   270
   End
End
Attribute VB_Name = "frmPFReportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lsReporteGeneral() As String
Dim lsCtaContDebe() As String
Dim lsCtaContHaber() As String
Dim lsObjetos() As String

Dim lbExcel As Boolean
Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet
Dim lsArchivo As String

Dim n As Integer

Dim lbLoad As Boolean
Dim oCtaIf As NCajaCtaIF
Dim oOpe As DOperacion
Dim oBarra As clsProgressBar
Dim objPista As COMManejador.Pista 'ARLO20170217

Private Sub chkBonos_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.txtFecha.SetFocus
End If
End Sub

Private Sub chkTodos_Click()
    If Me.chkTodos.value = 1 Then
        Me.txtCtaIFDesde.Enabled = False
    Else
        Me.txtCtaIFDesde.Enabled = True
        Me.txtCtaIFDesde.SetFocus
    End If
End Sub

Private Sub chkTodos_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Me.txtCtaIFDesde.Enabled = True Then
        Me.txtCtaIFDesde.SetFocus
    Else
        CmdGenerar.SetFocus
    End If
End If
End Sub

Private Sub cmdGenerar_Click()
On Error GoTo ErrorGenerar

If chkTodos.value = 0 Then
    If txtCtaIFDesde = "" Then
        MsgBox "Institución Financiera no Válida", vbInformation, "Aviso"
        Exit Sub
    End If
End If

If ValFecha(txtFecha) = False Then
    Exit Sub
End If

lbExcel = False
ReDim lsReporteGeneral(10, 0)
n = 0
GeneraReporteGeneral
If lsArchivo <> "" Then
   CargaArchivo lsArchivo, App.path & "\SPOOLER\"
End If
Exit Sub
ErrorGenerar:
    MsgBox "Error N°[" & Err.Number & "] " & Err.Description, vbInformation, "Aviso"
    If lbExcel = True Then
       ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1, False
    End If
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub
Private Sub Form_Activate()
If lbLoad = False Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
CentraForm Me
lbLoad = True
txtFecha = gdFecSis

Set oOpe = New DOperacion
Set oBarra = New clsProgressBar
Set oCtaIf = New NCajaCtaIF

CentraForm Me
txtCtaIFDesde.rs = oOpe.GetRsOpeObj(gsOpeCod, "1")
Me.Caption = gsOpeDesc
txtFecha = gdFecSis
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set oOpe = Nothing
Set oBarra = Nothing
End Sub

Private Sub GeneraReporteGeneral()
Dim fs As New Scripting.FileSystemObject
Dim lbExisteHoja As Boolean
Dim lnFila As Integer
Dim i As Integer
Dim lsTotal As String
Dim Y1 As Currency, Y2 As Currency
Dim prs As ADODB.Recordset

    Select Case gsOpeCod
        Case OpeCGRepRepBancosResumenPFMN, OpeCGRepRepBancosResumenPFME
            lsArchivo = App.path & "\SPOOLER\RepPFResBanc_" & IIf(Mid(gsOpeCod, 3, 1) = "1", "MN", "ME") & ".XLSX"

        Case OpeCGRepRepCMACSResumenPFMN, OpeCGRepRepCMACSResumenPFME
            lsArchivo = App.path & "\SPOOLER\RepPFResCmac_" & IIf(Mid(gsOpeCod, 3, 1) = "1", "MN", "ME") & ".XLSX"
    End Select
    lbExcel = ExcelBegin(lsArchivo, xlAplicacion, xlLibro, True)
    If Not lbExcel Then
        Exit Sub
    End If
    ExcelAddHoja "R_" & Format(txtFecha, "ddmmyyyy"), xlLibro, xlHoja1
    
    'edpyme -
    'xlHoja1.PageSetup.Zoom = 80
    'xlHoja1.PageSetup.Orientation = xlLandscape
    'xlHoja1.PageSetup.CenterHorizontally = True
    xlAplicacion.Range("A1:R100").Font.Size = 9
    
    xlHoja1.Range("A1").ColumnWidth = 20
    xlHoja1.Range("B1").ColumnWidth = 7
    xlHoja1.Range("C1").ColumnWidth = 27
    xlHoja1.Range("D1").ColumnWidth = 10
    xlHoja1.Range("E1").ColumnWidth = 10
    xlHoja1.Range("F1").ColumnWidth = 12
    xlHoja1.Range("G1").ColumnWidth = 7
    xlHoja1.Range("H1").ColumnWidth = 12
    xlHoja1.Range("I1").ColumnWidth = 12
    
    lnFila = 4
    xlHoja1.Cells(lnFila, 1) = gsNomCmac
    xlAplicacion.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 1)).Font.Bold = True
    xlHoja1.Cells(lnFila, 8) = "Area de Caja General"
    xlAplicacion.Range(xlHoja1.Cells(lnFila, 8), xlHoja1.Cells(lnFila, 9)).Font.Bold = True
    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 1) = "Datos al :" & Format(txtFecha, "dd mmmm yyyy")
    xlAplicacion.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 1)).Font.Bold = True
    xlHoja1.Cells(lnFila, 8) = "Fecha      :" & Format(gdFecSis, "dd mmmm yyyy")
    xlAplicacion.Range(xlHoja1.Cells(lnFila, 8), xlHoja1.Cells(lnFila, 9)).Font.Bold = True
    
    lnFila = lnFila + 2
    xlHoja1.Cells(lnFila, 3) = "REPORTE CONSOLIDADO CUENTAS A PLAZO EN " & IIf(Mid(gsOpeCod, 3, 1) = "1", "MONEDA NACIONAL", "MONEDA EXTRANJERA")
    xlAplicacion.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 3)).Font.Bold = True
    
    lnFila = lnFila + 2
    Y1 = lnFila
    xlAplicacion.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 10)).Font.Bold = True
    xlAplicacion.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 10)).HorizontalAlignment = xlCenter
    xlHoja1.Cells(lnFila, 1) = "ENTIDAD FINANCIERA"
    xlHoja1.Cells(lnFila, 2) = "PLAZO"
    xlHoja1.Cells(lnFila, 3) = "N° CUENTA"
    xlHoja1.Cells(lnFila, 4) = "FECHA DE"
    xlHoja1.Cells(lnFila, 5) = "FECHA DE"
    xlHoja1.Cells(lnFila, 6) = "CAP. INI."
    xlHoja1.Cells(lnFila, 7) = "T.E.A"
    xlHoja1.Cells(lnFila, 8) = "INTERES"
    xlHoja1.Cells(lnFila, 9) = "CAPITAL ACT."
    xlHoja1.Cells(lnFila, 10) = "TIPO CUENTA"
    lnFila = lnFila + 1
    xlAplicacion.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 10)).Font.Bold = True
    xlAplicacion.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 10)).HorizontalAlignment = xlCenter
    xlHoja1.Cells(lnFila, 1) = ""
    xlHoja1.Cells(lnFila, 2) = ""
    xlHoja1.Cells(lnFila, 3) = ""
    xlHoja1.Cells(lnFila, 4) = "APERTURA"
    xlHoja1.Cells(lnFila, 5) = "VCTO."
    xlHoja1.Cells(lnFila, 6) = IIf(Mid(gsOpeCod, 3, 1) = "2", "US$", "S/.")
    xlHoja1.Cells(lnFila, 7) = ""
    xlHoja1.Cells(lnFila, 8) = "DEVENGADOS"
    xlHoja1.Cells(lnFila, 9) = IIf(Mid(gsOpeCod, 3, 1) = "2", "US$", "S/.")
    Y2 = lnFila
    ExcelCuadro xlHoja1, 1, Y1, 10, Y2
    lsTotal = ""
    Y1 = lnFila + 1
    Set prs = CargaDatosReporte(Me.txtFecha, Mid(gsOpeCod, 3, 1))
    If Not prs.EOF Then
        Do While Not prs.EOF
            lnFila = lnFila + 1
            xlHoja1.Cells(lnFila, 1) = prs!cPersNombre
            xlHoja1.Cells(lnFila, 2) = prs!nCtaIFPlazo
            xlAplicacion.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 2)).HorizontalAlignment = xlCenter
            xlHoja1.Cells(lnFila, 3) = prs!cCtaIFDesc
            xlHoja1.Cells(lnFila, 4) = "'" & Format(prs!dCtaIFAper, gsFormatoFechaView)
            xlAplicacion.Range(xlHoja1.Cells(lnFila, 4), xlHoja1.Cells(lnFila, 4)).HorizontalAlignment = xlCenter
            xlHoja1.Cells(lnFila, 5) = "'" & Format(prs!dCtaIFVenc, gsFormatoFechaView)
            xlAplicacion.Range(xlHoja1.Cells(lnFila, 5), xlHoja1.Cells(lnFila, 5)).HorizontalAlignment = xlCenter
            xlHoja1.Cells(lnFila, 6) = prs!nCapitalIni
            xlHoja1.Cells(lnFila, 7) = prs!nCtaIFIntValor
            xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 7)).NumberFormat = "#,##0.00;-#,##0.00"
            xlHoja1.Cells(lnFila, 8) = prs!nInteres
            xlHoja1.Cells(lnFila, 9) = prs!nSaldo
            xlHoja1.Range(xlHoja1.Cells(lnFila, 8), xlHoja1.Cells(lnFila, 9)).NumberFormat = "#,##0.00;-#,##0.00"
            lsTotal = lsTotal + xlHoja1.Range(xlHoja1.Cells(lnFila, 9), xlHoja1.Cells(lnFila, 9)).Address(False, False) & "+"
            xlHoja1.Cells(lnFila, 10) = prs!cTpoCuenta
            prs.MoveNext
        Loop
        Y2 = lnFila
        ExcelCuadro xlHoja1, 1, Y1, 10, Y2
        lnFila = lnFila + 1
        Y1 = lnFila
        xlAplicacion.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 9)).Font.Bold = True
        xlHoja1.Cells(lnFila, 8) = "TOTAL : "
        xlHoja1.Range(xlHoja1.Cells(lnFila, 9), xlHoja1.Cells(lnFila, 9)).Formula = "=Sum(" & Mid(lsTotal, 1, Len(Trim(lsTotal)) - 1) & ")"
        xlHoja1.Range(xlHoja1.Cells(lnFila, 9), xlHoja1.Cells(lnFila, 9)).NumberFormat = "#,##0.00;-#,##0.00"
        Y2 = lnFila
        ExcelCuadro xlHoja1, 8, Y1, 9, Y2
    End If
    ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1, True
    Me.Estado.Panels(1).Text = "Reporte Generado con Exito"
    lbExcel = False
                'ARLO20170217
                Set objPista = New COMManejador.Pista
                objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "", Me.Caption & " Se Genero Reporte "
                Set objPista = Nothing
                '****
End Sub

Private Function CargaDatosReporte(pdFecha As Date, psMoneda As String) As ADODB.Recordset
Dim n As Integer
Dim sql As String
Dim rs As New ADODB.Recordset
Dim lsFiltro As String
Dim lsFiltroObjeto As String
Dim lsCtaHaber As String
Dim i As Integer
Dim oCon As DConecta
Dim lsCtaCapitalSdo As String
Set oCon = New DConecta

Set rs = oOpe.CargaOpeObj(gsOpeCod, adLockReadOnly)
If rs.EOF Then
    RSClose rs
    MsgBox "No se asignó definición de Institución Financiera para reporte", vbInformation, "¡Aviso!"
    Exit Function
End If
lsFiltro = rs!cOpeObjFiltro

Set rs = oOpe.CargaOpeCtaUltimoNivel(gsOpeCod, "D")
lsCtaCapitalSdo = RSMuestraLista(rs)

RSClose rs
sql = ""
sql = sql & " SELECT ci.cIFTpo, ci.cPersCod, ci.cCtaIFCod, p.cPersNombre, ci.cCtaIFDesc, ci.nCtaIFPlazo, dCtaIFaper, dCtaIFVenc, ISNULL(cii.nCtaIFIntValor,0) nCtaIFIntValor, "
sql = sql & "  ci.nInteres, CASE WHEN ci.cCtaIFCod LIKE '__1%' THEN ini.nSaldo ELSE ini.nSaldoME END nCapitalIni, CASE WHEN ci.cCtaIFCod LIKE '__1%' THEN cis.nSaldo ELSE cis.nSaldoME END nSaldo,"
sql = sql & "  CASE WHEN LEFT(ci.cCtaIFCod,2) = '01' THEN 'CTA CTE'"
sql = sql & "      WHEN LEFT(ci.cCtaIFCod,2) = '02' THEN 'CTA AHORROS'"
sql = sql & "      WHEN LEFT(ci.cCtaIFCod,2) = '03' THEN 'PLAZO FIJO'"
sql = sql & "      WHEN LEFT(ci.cCtaIFCod,2) = '05' THEN 'ADEUDADOS' END cTpoCuenta"
sql = sql & " FROM ctaif ci join persona p on p.cPersCod = ci.cPersCod"
sql = sql & "   LEFT JOIN CtaIFInteres cii ON cii.cIFTpo = ci.cIFTpo and cii.cPersCod = ci.cPersCod and cii.cCtaIFCod = ci.cCtaIFCod"
sql = sql & "   JOIN (SELECT  CS.cPersCod, CS.cIFTpo, CS.cCtaIFCod, SUM(CS.nSaldo) nSaldo, SUM(CS.nSaldoME) nSaldoME FROM CTAIFSALDO CS "
sql = sql & "         WHERE   CS.cCtaContCod in (" & lsCtaCapitalSdo & ") and CS.dCtaIFSaldo = ( SELECT  MAX(dCtaIFSaldo) "
sql = sql & "                         FROM    CTAIFSALDO CS1 Where CS1.cCtaContCod = CS.cCtaContCod And CS1.cPersCod = CS.cPersCod And CS1.cIFTpo = CS.cIFTpo And CS1.cCtaIFCod = CS.cCtaIFCod "
sql = sql & "                             AND CS1.dCtaIFSaldo<='" & Format(pdFecha, gsFormatoFecha) & "') "
sql = sql & "         GROUP BY CS.cPersCod, CS.cIFTpo, CS.cCtaIFCod "
sql = sql & "        ) cis ON cis.cIFTpo = ci.cIFTpo and cis.cPersCod = ci.cPersCod and cis.cCtaIFCod = ci.cCtaIFCod"
sql = sql & "  JOIN (Select cIFTpo, cPersCod, cCtaIFCod, SUM(nSaldo) nSaldo, SUM(nSaldoME) nSaldoME"
sql = sql & "        FROM CtaIFSaldo ini1"
sql = sql & "        where ini1.cCtaContCod in (" & lsCtaCapitalSdo & ") "
sql = sql & "        and ini1.dCtaIFSaldo = (SELECT Min(dCtaIFSaldo) FROM CtaIFSaldo cis1"
sql = sql & "                 WHERE cis1.cIFTpo = ini1.cIFTpo and cis1.cPersCod = ini1.cPersCod and cis1.cCtaIFCod = ini1.cCtaIFCod and cis1.dCtaIFSaldo <= '" & Format(pdFecha, gsFormatoFecha) & "')"
sql = sql & "        GROUP BY cIFTpo, cPersCod, cCtaIFCod"
sql = sql & "       ) ini ON  ini.cIFTpo = ci.cIFTpo and ini.cPersCod = ci.cPersCod and ini.cCtaIFCod = ci.cCtaIFCod"
sql = sql & " WHERE ci.cIFTpo+ci.cCtaIFCod like '" & lsFiltro & "%'"
sql = sql & "   and cii.dCtaIFIntRegistro = (SELECT Max(dCtaIFIntRegistro) FROM CtaIFInteres cii1"
sql = sql & "                  WHERE cii1.cIFTpo = cii.cIFTpo and cii1.cPersCod = cii.cPersCod and cii1.cCtaIFCod = cii.cCtaIFCod and cii1.dCtaIFIntRegistro <= '" & Format(pdFecha, gsFormatoFecha) & "')"
sql = sql & "   and ((ci.cCtaIFCod like '__1%' and cis.nSaldo <> 0) or (ci.cCtaIFCod like '__2%' and cis.nSaldoME <> 0) )"
sql = sql & "   and ci.cCtaIFEstado  = " & gEstadoCtaIFActiva
sql = sql & " ORDER BY ci.cIFTpo, ci.cPersCod, ci.cCtaIFCod"

oCon.AbreConexion
Set CargaDatosReporte = oCon.CargaRecordSet(sql)
oCon.CierraConexion
Set oCon = Nothing
End Function

Private Sub txtCtaIFDesde_EmiteDatos()
lblDescCtaDesde = oCtaIf.EmiteTipoCuentaIF(Mid(txtCtaIFDesde, 18, 10)) + " " + txtCtaIFDesde.psDescripcion
lblDescIFDesde = oCtaIf.NombreIF(Mid(txtCtaIFDesde, 4, 13))
CmdGenerar.SetFocus
End Sub

Private Sub txtFecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.chkTodos.SetFocus
End If
End Sub

