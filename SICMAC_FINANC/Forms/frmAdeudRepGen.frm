VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmAdeudRepGen 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "  Adeudados: Reporte General de Pagarés"
   ClientHeight    =   2745
   ClientLeft      =   2430
   ClientTop       =   2790
   ClientWidth     =   4995
   Icon            =   "frmAdeudRepGen.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   4995
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   90
      TabIndex        =   11
      Top             =   45
      Width           =   4785
      Begin MSMask.MaskEdBox txtfecha 
         Height          =   330
         Left            =   840
         TabIndex        =   0
         Top             =   255
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   582
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   390
         Left            =   4245
         Picture         =   "frmAdeudRepGen.frx":08CA
         Stretch         =   -1  'True
         Top             =   150
         Width           =   465
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Fecha  :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   105
         TabIndex        =   12
         Top             =   225
         Width           =   720
      End
   End
   Begin MSComctlLib.ProgressBar Barra 
      Height          =   165
      Left            =   2580
      TabIndex        =   10
      Top             =   2550
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   360
      Left            =   3735
      TabIndex        =   5
      Top             =   2070
      Width           =   1185
   End
   Begin VB.CommandButton cmdGenerar 
      Caption         =   "&Generar"
      Height          =   360
      Left            =   2445
      TabIndex        =   4
      Top             =   2085
      Width           =   1185
   End
   Begin MSComctlLib.StatusBar Estado 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      TabIndex        =   9
      Top             =   2505
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   423
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraopciones 
      Caption         =   "Banco"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1125
      Left            =   105
      TabIndex        =   6
      Top             =   840
      Width           =   4815
      Begin Sicmact.TxtBuscar txtCodObjeto 
         Height          =   345
         Left            =   1185
         TabIndex        =   1
         Top             =   180
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   609
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
         Enabled         =   0   'False
         Enabled         =   0   'False
         sTitulo         =   ""
         EnabledText     =   0   'False
      End
      Begin VB.CheckBox chktodos 
         Caption         =   "&Todos"
         Height          =   270
         Left            =   3810
         TabIndex        =   2
         Top             =   240
         Value           =   1  'Checked
         Width           =   765
      End
      Begin VB.Label lblObjDesc 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   1170
         TabIndex        =   3
         Top             =   570
         Width           =   3450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Descripcion :"
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
         Height          =   195
         Left            =   105
         TabIndex        =   8
         Top             =   630
         Width           =   1140
      End
      Begin VB.Label Label1 
         Caption         =   "Objeto :"
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
         Height          =   255
         Left            =   105
         TabIndex        =   7
         Top             =   270
         Width           =   675
      End
   End
End
Attribute VB_Name = "frmAdeudRepGen"
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
Dim lbBancos As Boolean
Dim lbCortoPlazo As Boolean
Dim lbLoad As Boolean
Dim dbCmact As DConecta
Dim objPista As COMManejador.Pista 'ARLO20170217

Private Sub chkTodos_Click()
    If Me.chkTodos.value = 1 Then
        Me.txtCodObjeto.Enabled = False
        Me.lblObjDesc = ""
        Me.txtCodObjeto = ""
    Else
        Me.txtCodObjeto.Enabled = True
    End If
End Sub
Public Sub Inicio(Optional pbCortoPlazo As Boolean = False)
    lbCortoPlazo = pbCortoPlazo
    Me.Show 1
End Sub
Private Sub chkTodos_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       CmdGenerar.SetFocus
    End If
End Sub

Private Sub cmdGenerar_Click()
  On Error GoTo ErrorGenerar

    If chkTodos.value = 0 Then
        If txtCodObjeto = "" Then
            MsgBox "No se selecciono Cuenta de Adeudado", vbInformation, "Aviso"
            Exit Sub
        End If
    End If
    If ValFecha(txtFecha) = False Then
        Exit Sub
    End If

    lbExcel = False
    ReDim lsReporteGeneral(22, 0)
    n = 0
    GeneraReporte
    If UBound(lsReporteGeneral, 2) = 0 Then
        MsgBox "No se posee Información para Procesar el reporte", vbInformation, "Aviso"
        Exit Sub
    End If
    If lsArchivo <> "" Then
        ExcelEnd App.path & "\SPOOLER\" & lsArchivo, xlAplicacion, xlLibro, xlHoja1, True
        MsgBox "Archivo Generado Satisfactoriamente", vbInformation, "Aviso"
        CargaArchivo lsArchivo, App.path & "\SPOOLER"
        lbExcel = False
                'ARLO20170217
                Set objPista = New COMManejador.Pista
                objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "", Me.Caption & " Se Genero Reporte "
                Set objPista = Nothing
                '****
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
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim n As Long, m As Long

    Me.Caption = "  " & UCase(gsOpeDesc)
    CentraForm Me
    lbLoad = True
    Set dbCmact = New DConecta
    dbCmact.AbreConexion
    txtFecha = gdFecSis
    ReDim lsObjetos(4, 0)
    n = 0
    sql = "Select * from OpeObj where cOpeCod ='" & gsOpeCod & "' and cOpeObjOrden = '0'"
    Set rs = dbCmact.CargaRecordSet(sql)
    If Not RSVacio(rs) Then
        Do While Not rs.EOF
            n = n + 1
            ReDim Preserve lsObjetos(4, n)
            lsObjetos(1, n) = Trim(rs!cObjetoCod)
            lsObjetos(2, n) = Trim(rs!nOpeObjNiv)
            lsObjetos(3, n) = Trim(rs!cOpeObjFiltro)
            lsObjetos(4, n) = Trim(rs!cOpeObjOrden)
            rs.MoveNext
        Loop
    Else
        RSClose rs
        MsgBox "No se han Definido Objetos para Reporte", vbInformation, "Aviso"
        lbLoad = False
        Exit Sub
    End If
    RSClose rs
    
    Dim oIF As New DCajaCtasIF
    Me.txtCodObjeto.rs = oIF.CargaCtasIF(Mid(gsOpeCod, 3, 1), lsObjetos(3, 1), MuestraInstituciones)
    Set oIF = Nothing
    
End Sub

Private Function DatosReporteGeneral(lsBanco As String) As Boolean
    Dim rs As New ADODB.Recordset
    Dim lsFiltro As String
    Dim lsFiltro1 As String
    Dim lnImporteActual As Currency
    Dim lnCapitalIni As Currency
    Dim lnDevengados As Currency
    Dim lnTotal As Integer, j As Integer
    Dim lnIndiceVac As Double
    Dim ntempo As Currency
    
    lsFiltro = ""
    DatosReporteGeneral = False
    Me.Barra.value = 0
    Me.Estado.Panels(1).Text = ""

    Dim oIF As New NCajaAdeudados
    Dim oDAdeud As DCaja_Adeudados
    Set oDAdeud = New DCaja_Adeudados
    lnIndiceVac = oDAdeud.CargaIndiceVAC(txtFecha)
    Set oDAdeud = Nothing
    
    'Set rs = oIF.CargaDatosGeneralesCtaIF(lsBanco, lsObjetos(3, 1), , lnIndiceVac, txtfecha, gsOpeCod)
    If lsBanco <> "" Then lsBanco = " AND CIF.cPersCod = '" & Mid(lsBanco, 4, 50) & "'"
    
    Set rs = oIF.GetReporteDatosAdeudadoPago(txtFecha.Text, Mid(gsOpeCod, 3, 1), lsBanco, , False, True)
    
    lnTotal = rs.RecordCount
    j = 0
    If Not RSVacio(rs) Then
        Do While Not rs.EOF
        
            DatosReporteGeneral = True
            j = j + 1
            n = n + 1
            ReDim Preserve lsReporteGeneral(22, n)
            lsReporteGeneral(1, n) = Trim(rs!cPersNombre)
            lsReporteGeneral(2, n) = Trim(rs!cCtaIFDesc)
            lsReporteGeneral(3, n) = rs!dCtaIFAper
            If Val(Left(rs!cCtaIfCod, 2)) = gTpoCtaIFCtaAdeud Then
                    If Not IsNull(rs!dVencimiento) Then
                        lsReporteGeneral(4, n) = rs!dVencimiento
                        'lsReporteGeneral(11, N) = Format(rs!dVencimiento - CDate(txtfecha), "#,#0")
                    End If
            Else
                    If Not IsNull(rs!dCtaIFVenc) Then
                       lsReporteGeneral(4, n) = rs!dCtaIFVenc
                       'lsReporteGeneral(11, N) = Format(rs!dCtaIFVenc - CDate(txtfecha), "#,#0")
                    End If
            End If
            lsReporteGeneral(5, n) = Format(rs!nMontoPrestado, "#,#0.00")
            lsReporteGeneral(11, n) = Format(rs!nSumInteres, "#,#0.00")
            lsReporteGeneral(6, n) = Format(rs!nCtaIFIntValor, "#,#0.00")
            lsReporteGeneral(7, n) = rs!nNroCuotas
            lsReporteGeneral(8, n) = rs!cCtaIFPlazo
            lsReporteGeneral(9, n) = Format((rs!nSaldoCap), "#,#0.00")
            lsReporteGeneral(13, n) = Format((rs!nCapitalPagado), "#,#0.00")
            lsReporteGeneral(14, n) = Format((rs!nInteresPagado), "#,#0.00")
            lsReporteGeneral(15, n) = Format((rs!nSumInteresPen), "#,#0.00")
            lsReporteGeneral(10, n) = Trim(rs!cCodLinCred)
            lsReporteGeneral(12, n) = Trim(rs!cDesLinCred)
'            If Mid(rs!cCtaIfCod, 3, 1) = "1" And rs!cMonedaPago = "2" Then
'               'SaldoCortoPlazo
'               lsReporteGeneral(13, N) = Format((rs!nSaldoCap - rs!nSaldoCapLP) * lnIndiceVac, "#,#0.00")
'               'nSaldoCapLP
'               lsReporteGeneral(14, N) = Format((rs!nSaldoCapLP * lnIndiceVac), "#,#0.00")
'            Else
'               'SaldoCortoPlazo
'               lsReporteGeneral(13, N) = Format(rs!nSaldoCap - rs!nSaldoCapLP, "#,#0.00")
'               'nSaldoCapLP
'               lsReporteGeneral(14, N) = Format(rs!nSaldoCapLP, "#,#0.00")
'            End If
            
'            If lbCortoPlazo Then
'                If Mid(gsOpeCod, 3, 1) = "1" And rs!cMonedaPago = "2" Then
'                        lnImporteActual = Format((rs!nSaldoCap - rs!nSaldoCapLP) * lnIndiceVac, "#,#0.00")
'                Else
'                    lnImporteActual = Format(rs!nSaldoCap - rs!nSaldoCapLP, "#,#0.00")
'                End If
'            Else
'                If Mid(gsOpeCod, 3, 1) = "1" And rs!cMonedaPago = "2" Then
'                    lnImporteActual = Format(rs!nSaldoCap * lnIndiceVac, "#,#0.00")
'                Else
'                    lnImporteActual = rs!nSaldoCap
'                End If
'            End If
'            If lnImporteActual > 0 Or lbCortoPlazo Then
'                lnDevengados = 0
'                lnCapitalIni = 0
'                N = N + 1
'                lnCapitalIni = IIf(Mid(gsOpeCod, 3, 1) = "1", rs!nSaldoCap, rs!nSaldoCap)
'                ReDim Preserve lsReporteGeneral(22, N)
'
'                lsReporteGeneral(1, N) = Trim(rs!cPersNombre)
'
'                If Mid(rs!cCtaIfCod, 3, 1) = "1" And rs!cMonedaPago = "2" Then
'                    'SaldoCortoPlazo
'                    lsReporteGeneral(13, N) = Format((rs!nSaldoCap - rs!nSaldoCapLP) * lnIndiceVac, "#,#0.00")
'                    'nSaldoCapLP
'                    lsReporteGeneral(14, N) = Format((rs!nSaldoCapLP * lnIndiceVac), "#,#0.00")
'                Else
'                    'SaldoCortoPlazo
'                    lsReporteGeneral(13, N) = Format(rs!nSaldoCap - rs!nSaldoCapLP, "#,#0.00")
'                    'nSaldoCapLP
'                    lsReporteGeneral(14, N) = Format(rs!nSaldoCapLP, "#,#0.00")
'                End If
'
'                'FechaVencimiento
'                lsReporteGeneral(15, N) = Format(rs!dVencimiento, gsFormatoFecha)
'
'
'                ntempo = DateDiff("d", CDate(txtfecha.Text), rs!dVencimientoFinal)
'                If ntempo <= 360 Then
'                    lsReporteGeneral(2, N) = Trim(Str(ntempo)) & "d"
'                Else
'                    lsReporteGeneral(2, N) = Round(ntempo / 360, 2) & "a"
'                End If
'
'                lsReporteGeneral(3, N) = Trim(rs!cCtaIFDesc)
'
'                lsReporteGeneral(4, N) = rs!dCtaIFAper
'                If Val(Left(rs!cCtaIfCod, 2)) = gTpoCtaIFCtaAdeud Then
'                    If Not IsNull(rs!dVencimiento) Then
'                        lsReporteGeneral(5, N) = rs!dVencimiento
'                        lsReporteGeneral(11, N) = Format(rs!dVencimiento - CDate(txtfecha), "#,#0")
'                    End If
'                Else
'                    If Not IsNull(rs!dCtaIFVenc) Then
'                       lsReporteGeneral(5, N) = rs!dCtaIFVenc
'                       lsReporteGeneral(11, N) = Format(rs!dCtaIFVenc - CDate(txtfecha), "#,#0")
'                    End If
'                End If
'                If gsCodCMAC = "102" Then
'                    lsReporteGeneral(6, N) = Format(lnImporteActual - oIF.GetAdeudadosSaldoCap(rs!cPersCod, rs!cCtaIfCod, rs!ciftpo, , , , , gCGTipoCuotCalIFNoConcesional), "#,#0.00")
'                Else
'                    lsReporteGeneral(6, N) = Format(rs!nMontoPrestado, "#,#0.00")
'                End If
'                lsReporteGeneral(7, N) = Format(rs!nCtaIFIntValor, "#,#0.00")
'                lsReporteGeneral(8, N) = Format(rs!nInteresProvisionadoReal, "#,#0.00")
'
'                lsReporteGeneral(9, N) = Format(lnImporteActual, "#,#0.00")
'
'                lsReporteGeneral(10, N) = Trim(rs!cCodLinCred)
'                lsReporteGeneral(12, N) = Trim(rs!cDesLinCred)
'                If lbCortoPlazo Then
'                    Dim lnInteresProvisionMes  As Currency
'                    lnInteresProvisionMes = 0
'                    If rs!dCuotaUltModSaldos = Me.txtfecha Then
'                        lnInteresProvisionMes = Format(rs!nInteresProvisionadoReal, "#,#0.00")
'                    End If
'
'                    If Mid(rs!cCtaIfCod, 3, 1) = "1" And rs!cMonedaPago = "2" Then
'                        lsReporteGeneral(6, N) = Format(lnIndiceVac, "#,#0.00###")
'                        lsReporteGeneral(5, N) = Round(lnImporteActual, 2)
'                        'lsReporteGeneral(9, N) = Format(lnImporteActual / rs!nVac, "#,#0.00")
'                        lsReporteGeneral(9, N) = Format(lnImporteActual / lnImporteActual, "#,#0.00")
'
'                    Else
'                        lsReporteGeneral(6, N) = Format(0, "#,#0.00")
'                        'lsReporteGeneral(5, N) = rs!nSaldoCap
'                        lsReporteGeneral(5, N) = lnImporteActual
'                    End If
'
'                    lsReporteGeneral(8, N) = Format(rs!nInteresProvisionadoReal, "#,#0.00")
'
'                    lsReporteGeneral(7, N) = Format(lnInteresProvisionMes, "#,#0.00")
'                End If
'            End If
            Me.Barra.value = Int(j / lnTotal * 100)
            Me.Estado.Panels(1).Text = "Avance :" & Format(j / lnTotal * 100, "#0.00") & "%"
            DoEvents
            rs.MoveNext
        Loop
    End If
    rs.Close
    Set rs = Nothing
End Function
Private Sub GeneraReporteGeneral()
    Dim fs As New Scripting.FileSystemObject
    Dim lbExisteHoja As Boolean
    Dim lnFila As Integer
    Dim i As Integer
    Dim lsCodLinea As String
    Dim lsDesLinea As String
    Dim sTotCap As String
    Dim sTotInt As String
    Dim sTotSdo As String
    Dim sTotSdoCorto As String
    Dim sTotSdoLargo As String
    Dim sTotGCap As String
    Dim sTotGInt As String
    Dim sTotGSdo As String
    Dim sTotGSdoCorto As String
    Dim sTotGSdoLargo As String
    Dim sTotGSdoK, sTotGSdoKPM, sTotGSdoIPM, sTotSIPP, sTotGIPP As String
    Dim sTotSdoKPM, sTotSdoIPM, sTotIPP As String
    Dim Y1 As Integer, Y2 As Integer
    Dim nSaldoCap As Currency
    Dim nSaldoInt As Currency
    lsCodLinea = ""
    gsPersNombre = ""
    lsArchivo = "RepGenAdeud_" & IIf(Mid(gsOpeCod, 3, 1) = "1", "MN", "ME") & Format(FechaHora(gdFecSis), gsFormatoMovFechaHora) & ".XLS"
    ExcelBegin lsArchivo, xlAplicacion, xlLibro, True
    lbExcel = True
    lbExisteHoja = False
    Me.Estado.Panels(1).Text = "Generando Reporte ..."
    If lbCortoPlazo Then
        ExcelAddHoja "CORTO_PLAZO", xlLibro, xlHoja1
    Else
        ExcelAddHoja "ADEUDADOS", xlLibro, xlHoja1
    End If
    xlHoja1.PageSetup.Zoom = 80
    xlHoja1.PageSetup.Orientation = xlLandscape
    xlHoja1.PageSetup.CenterHorizontally = True
    xlAplicacion.Range("A1:R100").Font.Size = 9
    
    xlHoja1.Range("A1").ColumnWidth = 20
    xlHoja1.Range("B1").ColumnWidth = 7
    xlHoja1.Range("C1").ColumnWidth = 27
    xlHoja1.Range("D1").ColumnWidth = 10
    xlHoja1.Range("E1").ColumnWidth = 10
    xlHoja1.Range("F1").ColumnWidth = 7
    xlHoja1.Range("G1").ColumnWidth = 12
    xlHoja1.Range("H1").ColumnWidth = 9
    xlHoja1.Range("I1:M1").ColumnWidth = 12

    lnFila = 2
    xlHoja1.Cells(lnFila, 1) = gsNomCmac
    xlHoja1.Range("A" & lnFila & ":C" & lnFila).MergeCells = True
    xlHoja1.Cells(lnFila, 12) = "Departamento de Finanzas"
    xlHoja1.Range("L" & lnFila & ":M" & lnFila).MergeCells = True
      
    xlHoja1.Range("A" & lnFila & ":O" & lnFila).Font.Bold = True
    xlHoja1.Range("A" & lnFila & ":O" & lnFila).HorizontalAlignment = xlLeft
    
    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 1) = "Datos al"
    xlHoja1.Cells(lnFila, 2) = Format(txtFecha, "dd mmmm yyyy")
    xlHoja1.Cells(lnFila, 12) = "Fecha"
    xlHoja1.Cells(lnFila, 13) = Format(gdFecSis, "dd mmmm yyyy")
    
    xlHoja1.Range("A" & lnFila & ":O" & lnFila).Font.Bold = True
    xlHoja1.Range("A" & lnFila & ":O" & lnFila).HorizontalAlignment = xlLeft
    
    lnFila = lnFila + 2
    xlHoja1.Cells(lnFila, 3) = "REPORTE CONSOLIDADO DE PAGARES DE ADEUDADOS"
    xlHoja1.Range("A" & lnFila & ":O" & lnFila).MergeCells = True
    xlHoja1.Range("A" & lnFila & ":O" & lnFila).Font.Bold = True
    xlHoja1.Range("A" & lnFila & ":O" & lnFila).Font.Underline = True
    xlHoja1.Range("A" & lnFila & ":O" & lnFila).HorizontalAlignment = xlCenter
    
    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 3) = IIf(Mid(gsOpeCod, 3, 1) = "1", "MONEDA NACIONAL", "MONEDA EXTRANJERA")
    xlHoja1.Range("A" & lnFila & ":O" & lnFila).MergeCells = True
    xlHoja1.Range("A" & lnFila & ":O" & lnFila).Font.Bold = True
    xlHoja1.Range("A" & lnFila & ":O" & lnFila).HorizontalAlignment = xlCenter
    
    lnFila = lnFila + 2
    Y1 = lnFila
'ALPA 20110812
'    xlHoja1.Cells(lnFila, 1) = "ENTIDAD FINANCIERA"
'    xlHoja1.Cells(lnFila, 2) = "PLAZO"
'    xlHoja1.Cells(lnFila, 3) = "N° PAGARE"
'    xlHoja1.Cells(lnFila, 4) = "FECHA DE"
'    xlHoja1.Cells(lnFila, 5) = "FECHA DE"
'    xlHoja1.Cells(lnFila, 6) = "VCTO "
'    If gsCodCMAC = "102" Then
'        xlHoja1.Cells(lnFila, 7) = "SDO.CAP.INT."
'    Else
'        xlHoja1.Cells(lnFila, 7) = "CAP. INI."
'    End If
'    xlHoja1.Cells(lnFila, 8) = "TASA"
'    xlHoja1.Cells(lnFila, 9) = "INTERES"
'    xlHoja1.Cells(lnFila, 10) = "CAP. ACTUAL"
'    xlHoja1.Cells(lnFila, 11) = "CAP. CORTO"
'    xlHoja1.Cells(lnFila, 12) = "CAP. LARGO"
'    xlHoja1.Cells(lnFila, 13) = "FECHA VENC"

'    xlHoja1.Cells(lnFila, 1) = ""
'    xlHoja1.Cells(lnFila, 2) = ""
'    xlHoja1.Cells(lnFila, 3) = ""
'    xlHoja1.Cells(lnFila, 4) = "APERTURA"
'    xlHoja1.Cells(lnFila, 5) = "VCTO."
'    xlHoja1.Cells(lnFila, 6) = "DIAS"
'    xlHoja1.Cells(lnFila, 7) = IIf(Mid(gsOpeCod, 3, 1) = "2", "US$", "S/.")
'    xlHoja1.Cells(lnFila, 8) = "%"
'    xlHoja1.Cells(lnFila, 9) = "PROVISION"
'    xlHoja1.Cells(lnFila, 10) = IIf(Mid(gsOpeCod, 3, 1) = "2", "US$", "S/.")
'    xlHoja1.Cells(lnFila, 11) = "PLAZO " & IIf(Mid(gsOpeCod, 3, 1) = "2", "US$", "S/.")
'    xlHoja1.Cells(lnFila, 12) = "PLAZO " & IIf(Mid(gsOpeCod, 3, 1) = "2", "US$", "S/.")
'    xlHoja1.Cells(lnFila, 13) = "ADEUDO"
'    lnFila = lnFila + 1
    
    xlHoja1.Cells(lnFila, 1) = "MONEDA NACIONAL"
    xlHoja1.Cells(lnFila, 2) = "CODIFICACION"
    xlHoja1.Cells(lnFila, 3) = "VIGENCIA"
    xlHoja1.Cells(lnFila, 5) = "TOTAL DESEMBOLSO"
    xlHoja1.Cells(lnFila, 8) = "Nº"
    xlHoja1.Cells(lnFila, 9) = "FRECUENCIA"
    xlHoja1.Cells(lnFila, 10) = "SALDO DE"
    xlHoja1.Cells(lnFila, 11) = "K PAGADO"
    xlHoja1.Cells(lnFila, 12) = "INT.PAGADO"
    xlHoja1.Cells(lnFila, 13) = "INTERERES"
    xlHoja1.Cells(lnFila, 14) = "DEP. EN GARANTIA"
    xlHoja1.Range("C" & lnFila & ":D" & lnFila).MergeCells = True
    xlHoja1.Range("E" & lnFila & ":G" & lnFila).MergeCells = True
    xlHoja1.Range("N" & lnFila & ":O" & lnFila).MergeCells = True
    lnFila = lnFila + 1

    xlHoja1.Cells(lnFila, 1) = "ENTIDAD ACREEDORA"
    xlHoja1.Cells(lnFila, 2) = "N° PAGARE"
    xlHoja1.Cells(lnFila, 3) = "DESEMBOLSO"
    xlHoja1.Cells(lnFila, 4) = "VENCE"
    xlHoja1.Cells(lnFila, 5) = "PRINCIPAL"
    xlHoja1.Cells(lnFila, 6) = "INTERES"
    xlHoja1.Cells(lnFila, 7) = "TEA"
    xlHoja1.Cells(lnFila, 8) = "CUOTAS"
    xlHoja1.Cells(lnFila, 9) = "DE PAGO"
    xlHoja1.Cells(lnFila, 10) = "CAPITAL"
    xlHoja1.Cells(lnFila, 11) = "EN EL MES"
    xlHoja1.Cells(lnFila, 12) = "EN EL MES"
    xlHoja1.Cells(lnFila, 13) = "POR PAGAR"
    xlHoja1.Cells(lnFila, 14) = "IMPORTE"
    xlHoja1.Cells(lnFila, 15) = "TEA"
    Y2 = lnFila
    CuadroExcel 1, Y1, 15, Y2
    CuadroExcel 1, Y1, 15, Y1
    
    '
    xlHoja1.Range(xlHoja1.Cells(lnFila - 1, 1), xlHoja1.Cells(lnFila, 15)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(lnFila - 1, 1), xlHoja1.Cells(lnFila, 15)).HorizontalAlignment = xlCenter
    xlHoja1.Range("A" & lnFila - 1 & ":O" & lnFila).Interior.ColorIndex = 36
    xlHoja1.Range("A" & lnFila - 1 & ":O" & lnFila).Font.ColorIndex = 11
    '
    
    Y1 = lnFila + 1
    lnFila = lnFila + 1
    sTotGCap = "": sTotGInt = "": sTotGSdo = ""
    sTotCap = "": sTotInt = "": sTotSdo = ""
    sTotSdoKPM = "": sTotSdoIPM = "": sTotIPP = ""
    sTotGSdoKPM = "": sTotGSdoIPM = "": sTotGIPP = ""
    For i = 1 To UBound(lsReporteGeneral, 2)
        If lsCodLinea <> lsReporteGeneral(10, i) Or gsPersNombre <> lsReporteGeneral(1, i) Then
            If i > 1 Then
                Y2 = lnFila
                CuadroExcel 1, Y1, 15, Y2
                lnFila = lnFila + 1
                xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 15)).Font.Bold = True
                xlHoja1.Range(xlHoja1.Cells(lnFila, 5), xlHoja1.Cells(lnFila, 12)).NumberFormat = "#,##0.00;-#,##0.00"
                If lsCodLinea <> "" Then
                    xlHoja1.Cells(lnFila, 1) = "TOTAL LINEA " & lsDesLinea
                Else
                    xlHoja1.Cells(lnFila, 1) = "TOTAL " & gsPersNombre
                End If
                xlHoja1.Range(xlHoja1.Cells(lnFila, 5), xlHoja1.Cells(lnFila, 5)).Formula = "=SUM(E" & Y1 + 1 & ":E" & lnFila - 1 & ")"
                xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Formula = "=SUM(F" & Y1 + 1 & ":F" & lnFila - 1 & ")"
                'xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).Formula = "=SUM(G" & Y1 + 1 & ":G" & lnFila - 1 & ")"
                xlHoja1.Range(xlHoja1.Cells(lnFila, 10), xlHoja1.Cells(lnFila, 10)).Formula = "=SUM(J" & Y1 + 1 & ":J" & lnFila - 1 & ")"
                xlHoja1.Range(xlHoja1.Cells(lnFila, 11), xlHoja1.Cells(lnFila, 11)).Formula = "=SUM(K" & Y1 + 1 & ":K" & lnFila - 1 & ")"
                xlHoja1.Range(xlHoja1.Cells(lnFila, 12), xlHoja1.Cells(lnFila, 12)).Formula = "=SUM(L" & Y1 + 1 & ":L" & lnFila - 1 & ")"
                xlHoja1.Range(xlHoja1.Cells(lnFila, 13), xlHoja1.Cells(lnFila, 13)).Formula = "=SUM(M" & Y1 + 1 & ":M" & lnFila - 1 & ")"

                'Total General
                sTotGCap = sTotGCap & "+" & xlHoja1.Range(xlHoja1.Cells(lnFila, 5), xlHoja1.Cells(lnFila, 5)).Address
                sTotGInt = sTotGInt & "+" & xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Address
                sTotGSdoK = sTotGSdoK & "+" & xlHoja1.Range(xlHoja1.Cells(lnFila, 10), xlHoja1.Cells(lnFila, 10)).Address 'sTotGSdoCorto
                sTotGSdoKPM = sTotGSdoKPM & "+" & xlHoja1.Range(xlHoja1.Cells(lnFila, 11), xlHoja1.Cells(lnFila, 11)).Address 'sTotGSdoCorto
                sTotGSdoIPM = sTotGSdoIPM & "+" & xlHoja1.Range(xlHoja1.Cells(lnFila, 12), xlHoja1.Cells(lnFila, 12)).Address 'sTotGSdoLargo
                sTotGIPP = sTotGIPP & "+" & xlHoja1.Range(xlHoja1.Cells(lnFila, 13), xlHoja1.Cells(lnFila, 13)).Address 'sTotGSdoLargo
                
'                sTotCap = sTotGCap
'                sTotInt = sTotGInt
'                sTotSdo = sTotGSdoK
'                sTotSdoKPM = sTotGSdoKPM
'                sTotSdoIPM = sTotGSdoIPM
'                sTotIPP = sTotGIPP

                If lsCodLinea <> "" Then
                    If gsPersNombre <> "" And gsPersNombre <> lsReporteGeneral(1, i) Then
                        lnFila = lnFila + 2
                        xlHoja1.Cells(lnFila, 1) = "TOTAL " & gsPersNombre
                        xlHoja1.Range(xlHoja1.Cells(lnFila, 5), xlHoja1.Cells(lnFila, 5)).Formula = "=" & sTotGCap
                        xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Formula = "=" & sTotGInt
                        xlHoja1.Range(xlHoja1.Cells(lnFila, 10), xlHoja1.Cells(lnFila, 10)).Formula = "=" & sTotGSdoK
                        xlHoja1.Range(xlHoja1.Cells(lnFila, 11), xlHoja1.Cells(lnFila, 11)).Formula = "=" & sTotGSdoKPM
                        xlHoja1.Range(xlHoja1.Cells(lnFila, 12), xlHoja1.Cells(lnFila, 12)).Formula = "=" & sTotGSdoIPM
                        xlHoja1.Range(xlHoja1.Cells(lnFila, 13), xlHoja1.Cells(lnFila, 13)).Formula = "=" & sTotGIPP
                        xlHoja1.Range(xlHoja1.Cells(lnFila, 5), xlHoja1.Cells(lnFila, 13)).NumberFormat = "#,##0.00;-#,##0.00"
                        xlHoja1.Range(xlHoja1.Cells(lnFila, 5), xlHoja1.Cells(lnFila, 13)).Font.Bold = True
                        sTotCap = "": sTotInt = "": sTotSdo = ""
                        sTotSdoKPM = "": sTotSdoIPM = "": sTotIPP = ""
                        sTotCap = ""
                        sTotInt = ""
                        sTotSdo = ""
                        sTotSdoKPM = ""
                        sTotSdoIPM = ""
                        sTotIPP = ""
                    Else
                        sTotCap = sTotCap & "+" & xlHoja1.Range(xlHoja1.Cells(lnFila, 5), xlHoja1.Cells(lnFila, 5)).Address
                        sTotInt = sTotInt & "+" & xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Address
                        sTotSdo = sTotSdo & "+" & xlHoja1.Range(xlHoja1.Cells(lnFila, 10), xlHoja1.Cells(lnFila, 10)).Address
                        sTotSdoKPM = sTotSdoKPM & "+" & xlHoja1.Range(xlHoja1.Cells(lnFila, 11), xlHoja1.Cells(lnFila, 11)).Address
                        sTotSdoIPM = sTotSdoIPM & "+" & xlHoja1.Range(xlHoja1.Cells(lnFila, 12), xlHoja1.Cells(lnFila, 12)).Address
                        sTotIPP = sTotIPP & "+" & xlHoja1.Range(xlHoja1.Cells(lnFila, 13), xlHoja1.Cells(lnFila, 13)).Address
                    End If
                End If
                Y1 = lnFila
                lnFila = lnFila + 1
            End If
            lnFila = lnFila + 1
            If lsReporteGeneral(10, i) <> "" Then
                xlHoja1.Cells(lnFila, 1) = "LINEA DE CREDITO : " & lsReporteGeneral(12, i)
            End If
            lsCodLinea = lsReporteGeneral(10, i)
            lsDesLinea = lsReporteGeneral(12, i)
            gsPersNombre = lsReporteGeneral(1, i)
        End If
        lnFila = lnFila + 1
        xlHoja1.Cells(lnFila, 1) = lsReporteGeneral(1, i)
        xlHoja1.Cells(lnFila, 2) = "'" & lsReporteGeneral(2, i)
        xlAplicacion.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 2)).HorizontalAlignment = xlLeft
        xlHoja1.Cells(lnFila, 3) = "'" & lsReporteGeneral(3, i)
        xlHoja1.Cells(lnFila, 4) = "'" & lsReporteGeneral(4, i)
        xlHoja1.Cells(lnFila, 5) = lsReporteGeneral(5, i)
        xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 4)).NumberFormat = "dd/mm/yyyy"
        xlHoja1.Cells(lnFila, 6) = lsReporteGeneral(11, i)
        xlHoja1.Cells(lnFila, 7) = lsReporteGeneral(6, i)
        xlHoja1.Cells(lnFila, 8) = Format(lsReporteGeneral(7, i), "#,#0.00")
        xlHoja1.Cells(lnFila, 9) = lsReporteGeneral(8, i)
        xlHoja1.Cells(lnFila, 10) = lsReporteGeneral(9, i)
        xlHoja1.Cells(lnFila, 11) = lsReporteGeneral(13, i)
        xlHoja1.Cells(lnFila, 12) = lsReporteGeneral(14, i)
        xlHoja1.Cells(lnFila, 13) = lsReporteGeneral(15, i)
        xlHoja1.Cells(lnFila, 14) = ""
        xlHoja1.Cells(lnFila, 15) = ""
        nSaldoCap = nSaldoCap + IIf(lsReporteGeneral(9, i) = "", 0, lsReporteGeneral(9, i))
        nSaldoInt = nSaldoInt + IIf(lsReporteGeneral(13, i) = "", 0, lsReporteGeneral(13, i))
        xlAplicacion.Range(xlHoja1.Cells(lnFila, 4), xlHoja1.Cells(lnFila, 4)).HorizontalAlignment = xlCenter
        xlAplicacion.Range(xlHoja1.Cells(lnFila, 5), xlHoja1.Cells(lnFila, 5)).HorizontalAlignment = xlCenter
        xlAplicacion.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).HorizontalAlignment = xlCenter
        xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 12)).NumberFormat = "#,##0.00;-#,##0.00"
    Next
    
    Y2 = lnFila
    CuadroExcel 1, Y1, 15, Y2
    lnFila = lnFila + 1
    xlAplicacion.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 15)).Font.Bold = True
    If lsCodLinea <> "" Then
        xlHoja1.Cells(lnFila, 1) = "TOTAL LINEA " & lsDesLinea
    Else
        xlHoja1.Cells(lnFila, 1) = "TOTAL"
    End If
    xlHoja1.Range(xlHoja1.Cells(lnFila, 5), xlHoja1.Cells(lnFila, 5)).Formula = "=SUM(E" & Y1 + 1 & ":E" & lnFila - 1 & ")"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Formula = "=SUM(F" & Y1 + 1 & ":F" & lnFila - 1 & ")"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 10), xlHoja1.Cells(lnFila, 10)).Formula = "=SUM(J" & Y1 + 1 & ":J" & lnFila - 1 & ")"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 11), xlHoja1.Cells(lnFila, 11)).Formula = "=SUM(K" & Y1 + 1 & ":K" & lnFila - 1 & ")"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 12), xlHoja1.Cells(lnFila, 12)).Formula = "=SUM(L" & Y1 + 1 & ":L" & lnFila - 1 & ")"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 13), xlHoja1.Cells(lnFila, 13)).Formula = "=SUM(M" & Y1 + 1 & ":M" & lnFila - 1 & ")"
    xlHoja1.Range(xlHoja1.Cells(lnFila, 5), xlHoja1.Cells(lnFila, 13)).NumberFormat = "#,##0.00;-#,##0.00"
    
    'Total General
    sTotGCap = sTotGCap & "+" & xlHoja1.Range(xlHoja1.Cells(lnFila, 5), xlHoja1.Cells(lnFila, 5)).Address
    sTotGInt = sTotGInt & "+" & xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Address
    sTotGSdoK = sTotGSdoK & "+" & xlHoja1.Range(xlHoja1.Cells(lnFila, 10), xlHoja1.Cells(lnFila, 10)).Address
    sTotGSdoKPM = sTotGSdoKPM & "+" & xlHoja1.Range(xlHoja1.Cells(lnFila, 11), xlHoja1.Cells(lnFila, 11)).Address
    sTotGSdoIPM = sTotGSdoIPM & "+" & xlHoja1.Range(xlHoja1.Cells(lnFila, 12), xlHoja1.Cells(lnFila, 12)).Address
    sTotGIPP = sTotGIPP & "+" & xlHoja1.Range(xlHoja1.Cells(lnFila, 13), xlHoja1.Cells(lnFila, 13)).Address
    
    'Total IF
    sTotCap = sTotCap & "+" & xlHoja1.Range(xlHoja1.Cells(lnFila, 5), xlHoja1.Cells(lnFila, 5)).Address
    sTotInt = sTotInt & "+" & xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Address
    sTotSdo = sTotSdo & "+" & xlHoja1.Range(xlHoja1.Cells(lnFila, 10), xlHoja1.Cells(lnFila, 10)).Address
    sTotSdoKPM = sTotSdoKPM & "+" & xlHoja1.Range(xlHoja1.Cells(lnFila, 11), xlHoja1.Cells(lnFila, 11)).Address
    sTotSdoIPM = sTotSdoIPM & "+" & xlHoja1.Range(xlHoja1.Cells(lnFila, 12), xlHoja1.Cells(lnFila, 12)).Address
    sTotIPP = sTotIPP & "+" & xlHoja1.Range(xlHoja1.Cells(lnFila, 13), xlHoja1.Cells(lnFila, 13)).Address
    Y1 = lnFila
    
    lnFila = lnFila + 2
    If lsCodLinea <> "" Then
        xlHoja1.Cells(lnFila, 1) = "TOTAL " & gsPersNombre
        xlHoja1.Range(xlHoja1.Cells(lnFila, 5), xlHoja1.Cells(lnFila, 5)).Formula = "=" & sTotCap
        xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Formula = "=" & sTotInt
        xlHoja1.Range(xlHoja1.Cells(lnFila, 10), xlHoja1.Cells(lnFila, 10)).Formula = "=" & sTotSdo
        xlHoja1.Range(xlHoja1.Cells(lnFila, 11), xlHoja1.Cells(lnFila, 11)).Formula = "=" & sTotSdoKPM
        xlHoja1.Range(xlHoja1.Cells(lnFila, 12), xlHoja1.Cells(lnFila, 12)).Formula = "=" & sTotSdoIPM
        xlHoja1.Range(xlHoja1.Cells(lnFila, 13), xlHoja1.Cells(lnFila, 13)).Formula = "=" & sTotIPP
        xlHoja1.Range(xlHoja1.Cells(lnFila, 5), xlHoja1.Cells(lnFila, 13)).NumberFormat = "#,##0.00;-#,##0.00"
        xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 13)).Font.Bold = True
    End If
    sTotCap = "=": sTotInt = "=": sTotSdo = ""
    sTotSdoKPM = "": sTotSdoIPM = "": sTotIPP = ""
    Y2 = lnFila
    CuadroExcel 1, Y1, 15, Y2
    Y1 = lnFila
    If sTotGCap <> "" Then
        lnFila = lnFila + 2
        xlHoja1.Cells(lnFila, 1) = "TOTAL ADEUDADOS "
        xlHoja1.Range(xlHoja1.Cells(lnFila, 5), xlHoja1.Cells(lnFila, 5)).Formula = "=" & sTotGCap
        xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Formula = "=" & sTotGInt
        xlHoja1.Range(xlHoja1.Cells(lnFila, 10), xlHoja1.Cells(lnFila, 10)).Formula = "=" & sTotGSdoK
        xlHoja1.Range(xlHoja1.Cells(lnFila, 11), xlHoja1.Cells(lnFila, 11)).Formula = "=" & sTotGSdoKPM
        xlHoja1.Range(xlHoja1.Cells(lnFila, 12), xlHoja1.Cells(lnFila, 12)).Formula = "=" & sTotGSdoIPM
        xlHoja1.Range(xlHoja1.Cells(lnFila, 13), xlHoja1.Cells(lnFila, 13)).Formula = "=" & sTotGIPP
        xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 12)).NumberFormat = "#,##0.00;-#,##0.00"
        xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 12)).Font.Bold = True
        Y2 = lnFila
        CuadroExcel 1, Y1, 15, Y2
        
        xlHoja1.Range("A" & lnFila & ":M" & lnFila).Interior.ColorIndex = 36
        xlHoja1.Range("A" & lnFila & ":M" & lnFila).Font.ColorIndex = 53

        xlHoja1.Cells(lnFila + 5, 1) = "Resumen de Capital"
        xlHoja1.Cells(lnFila + 6, 1) = "Resumen de Interes"
        xlHoja1.Cells(lnFila + 5, 2) = nSaldoCap
        xlHoja1.Cells(lnFila + 6, 2) = nSaldoInt
        CuadroExcel 1, lnFila + 5, 2, lnFila + 6
        
        
    End If
    
    '
    xlHoja1.Cells.Select
    xlHoja1.Cells.Font.Name = "Arial"
    xlHoja1.Cells.Font.Size = 8
    xlHoja1.Cells.EntireColumn.AutoFit
    '
    
    Me.Estado.Panels(1).Text = "Reporte Generado con Exito"
End Sub
Private Sub CuadroExcel(X1 As Integer, Y1 As Integer, X2 As Integer, Y2 As Integer, Optional lbLineasVert As Boolean = False)
    Dim i, j As Integer

    For i = X1 To X2
        xlHoja1.Range(xlHoja1.Cells(Y1, i), xlHoja1.Cells(Y1, i)).Borders(xlEdgeTop).LineStyle = xlContinuous
        xlHoja1.Range(xlHoja1.Cells(Y2, i), xlHoja1.Cells(Y2, i)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    Next i
    If lbLineasVert = False Then
        For i = X1 To X2
            For j = Y1 To Y2
                xlHoja1.Range(xlHoja1.Cells(j, i), xlHoja1.Cells(j, i)).Borders(xlEdgeLeft).LineStyle = xlContinuous
            Next j
        Next i
    End If
    If lbLineasVert Then
        For j = Y1 To Y2
            xlHoja1.Range(xlHoja1.Cells(j, X1), xlHoja1.Cells(j, X1)).Borders(xlEdgeRight).LineStyle = xlContinuous
        Next j
    End If

    For j = Y1 To Y2
        xlHoja1.Range(xlHoja1.Cells(j, X2), xlHoja1.Cells(j, X2)).Borders(xlEdgeRight).LineStyle = xlContinuous
    Next j
End Sub
Private Function InteresDevengado(lsObjetoCta As String) As Currency
    Dim sql As String
    Dim rs As New ADODB.Recordset

    sql = " SELECT  MAX(MC.CMOVNRO) as MaxFecha ,ISNULL(SUM(ISNULL(ME.nMovMEImporte,0)),0) AS TotalME, ISNULL(SUM(ISNULL(MC.NMOVIMPORTE,0)),0) as TotalMN " _
       & " FROM    MOVCTA MC JOIN MOVOBJ MO ON MO.CMOVNRO=MC.CMOVNRO AND MO.CMOVITEM=MC.CMOVITEM " _
       & "         LEFT JOIN MOVME ME   ON ME.CMOVNRO = MC.CMOVNRO AND ME.CMOVITEM = MC.CMOVITEM " _
       & "         JOIN MOV M           ON M.CMOVNRO = MC.CMOVNRO JOIN OPECTA OC ON OC.CCTACONTCOD = SUBSTRING(MC.CCTACONTCOD,1,LEN(OC.CCTACONTCOD))" _
       & " WHERE   OC.COPECOD ='" & gsOpeCod & "' AND OC.COPECTADH='H' AND OC.cOpeCtaTpo='1'  " _
       & "         AND COBJETOCOD='" & lsObjetoCta & "' " _
       & "         AND MC.nMovImporte>0 AND M.CMOVFLAG NOT IN ('X','E','N') AND SUBSTRING(MC.cMovNro,1,8)<='" & Format(txtFecha, "yyyymmdd") & "'" _
       & " UNION " _
       & " SELECT  MAX(MC.CMOVNRO) as MaxFecha ,ISNULL(SUM(ISNULL(ME.nMovMEImporte,0)),0) AS TotalME, ISNULL(SUM(ISNULL(MC.NMOVIMPORTE,0)),0) as TotalMN " _
       & " FROM    MOVCTA MC JOIN MOVOBJ MO ON MO.CMOVNRO=MC.CMOVNRO AND MO.CMOVITEM=MC.CMOVITEM " _
       & "         LEFT JOIN MOVME ME   ON ME.CMOVNRO = MC.CMOVNRO AND ME.CMOVITEM = MC.CMOVITEM " _
       & "         JOIN MOV M           ON M.CMOVNRO = MC.CMOVNRO  JOIN OPECTA OC ON OC.CCTACONTCOD = SUBSTRING(MC.CCTACONTCOD,1,LEN(OC.CCTACONTCOD)) " _
       & " WHERE   OC.COPECOD ='" & gsOpeCod & "' AND OC.COPECTADH='H' AND OC.cOpeCtaTpo='2' " _
       & "         AND COBJETOCOD='" & lsObjetoCta & "' " _
       & "         AND MC.nMovImporte>0 AND M.CMOVFLAG NOT IN ('X','E','N') AND SUBSTRING(MC.cMovNro,1,8)<='" & Format(txtFecha, "yyyymmdd") & "'"
    InteresDevengado = 0
    rs.CursorLocation = adUseClient
    rs.Open sql, dbCmact, adOpenStatic, adLockBatchOptimistic, adCmdText
    rs.ActiveConnection = Nothing
    If Not RSVacio(rs) Then
        Do While Not rs.EOF
            InteresDevengado = InteresDevengado + IIf(Mid(gsOpeCod, 3, 1) = "1", rs!TotalMN, rs!TotalME)
            rs.MoveNext
        Loop
    End If
    rs.Close
    Set rs = Nothing
End Function

Private Sub Form_Unload(Cancel As Integer)
    CierraConexion
End Sub
Private Sub GeneraReporte()
    Dim n As Integer
    Dim sql As String
    Dim rs As ADODB.Recordset
    Dim lsFiltro As String
    Dim lsFiltroObjeto As String
    Dim lsCtaHaber As String


    Call DatosReporteGeneral(Trim(txtCodObjeto))

    If UBound(lsReporteGeneral, 2) > 0 Then
        If lbCortoPlazo Then
            GeneraReporteCorto
        Else
            GeneraReporteGeneral
        End If
    Else
        lsArchivo = ""
    End If
End Sub

Private Sub txtCodObjeto_EmiteDatos()
If txtCodObjeto <> "" Then
    lblObjDesc = txtCodObjeto.psDescripcion
    CmdGenerar.SetFocus
End If
End Sub

Private Sub txtFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.chkTodos.SetFocus
    End If
End Sub

Private Sub GeneraReporteCorto()
    Dim fs As New Scripting.FileSystemObject
    Dim lbExisteHoja As Boolean
    Dim lnFila As Integer
    Dim i As Integer
    Dim lsCodLinea As String
    Dim lsDesLinea As String
    Dim sTotCCal  As String
    Dim sTotCap   As String
    Dim sTotInt   As String
    Dim sTotIMes  As String
    Dim sTotSdo   As String
    Dim sTotGCCal As String
    Dim sTotGCap  As String
    Dim sTotGInt  As String
    Dim sTotGIMes As String
    Dim sTotGSdo  As String
    
    Dim Y1 As Integer, Y2 As Integer
    lsCodLinea = ""
    gsPersNombre = ""
    
    lsArchivo = "RepCortoAdeud_" & IIf(Mid(gsOpeCod, 3, 1) = "1", "MN", "ME") & Format(FechaHora(gdFecSis), gsFormatoMovFechaHora) & ".XLS"
 
    ExcelBegin App.path & "\Spooler\" & lsArchivo, xlAplicacion, xlLibro, True
    lbExcel = True
    lbExisteHoja = False
    Me.Estado.Panels(1).Text = "Generando Reporte ..."
    ExcelAddHoja "CORTO_PLAZO", xlLibro, xlHoja1
    xlHoja1.PageSetup.Zoom = 80
    xlHoja1.PageSetup.Orientation = xlLandscape
    xlHoja1.PageSetup.CenterHorizontally = True
    xlAplicacion.Range("A1:R100").Font.Size = 9
    
    xlHoja1.Range("A1").ColumnWidth = 20
    xlHoja1.Range("B1").ColumnWidth = 27
    xlHoja1.Range("C1").ColumnWidth = 14
    xlHoja1.Range("D1").ColumnWidth = 12
    xlHoja1.Range("E1").ColumnWidth = 14
    xlHoja1.Range("F1").ColumnWidth = 12
    xlHoja1.Range("G1").ColumnWidth = 14
    xlHoja1.Range("H1:K1").ColumnWidth = 12

    lnFila = 4
    xlHoja1.Cells(lnFila, 1) = gsNomCmac
    xlHoja1.Range("A" & lnFila & ":C" & lnFila).MergeCells = True
    xlHoja1.Cells(lnFila, 7) = "Departamento de Finanzas"
    xlHoja1.Range("G" & lnFila & ":H" & lnFila).MergeCells = True
    
    xlHoja1.Range("A" & lnFila & ":H" & lnFila).Font.Bold = True
    xlHoja1.Range("A" & lnFila & ":H" & lnFila).HorizontalAlignment = xlLeft
    
    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 1) = "Datos al"
    xlHoja1.Cells(lnFila, 2) = Format(txtFecha, "dd mmmm yyyy")
    xlHoja1.Cells(lnFila, 7) = "Fecha"
    xlHoja1.Cells(lnFila, 8) = Format(gdFecSis, "dd mmmm yyyy")
    
    xlHoja1.Range("A" & lnFila & ":H" & lnFila).Font.Bold = True
    xlHoja1.Range("A" & lnFila & ":H" & lnFila).HorizontalAlignment = xlLeft

    lnFila = lnFila + 2
    xlHoja1.Cells(lnFila, 3) = "REPORTE ADEUDADOS CORTO PLAZO"
    xlHoja1.Range("A" & lnFila & ":H" & lnFila).MergeCells = True
    xlHoja1.Range("A" & lnFila & ":H" & lnFila).Font.Bold = True
    xlHoja1.Range("A" & lnFila & ":H" & lnFila).Font.Underline = True
    xlHoja1.Range("A" & lnFila & ":H" & lnFila).HorizontalAlignment = xlCenter
    
    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 3) = IIf(Mid(gsOpeCod, 3, 1) = "1", "MONEDA NACIONAL", "MONEDA EXTRANJERA")
    xlHoja1.Range("A" & lnFila & ":H" & lnFila).MergeCells = True
    xlHoja1.Range("A" & lnFila & ":H" & lnFila).Font.Bold = True
    xlHoja1.Range("A" & lnFila & ":H" & lnFila).HorizontalAlignment = xlCenter

    lnFila = lnFila + 2
    Y1 = lnFila
'    xlAplicacion.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 11)).Font.Bold = True
'    xlAplicacion.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 11)).HorizontalAlignment = xlCenter
    xlHoja1.Cells(lnFila, 1) = "ENTIDAD FINANCIERA"
    xlHoja1.Cells(lnFila, 2) = "N° PAGARE"
    xlHoja1.Cells(lnFila, 3) = "CAPITAL"
    xlHoja1.Cells(lnFila, 4) = "VAC"
    xlHoja1.Cells(lnFila, 5) = "CAPITAL "
    xlHoja1.Cells(lnFila, 6) = "INTERES"
    xlHoja1.Cells(lnFila, 7) = "INTERES"
    xlHoja1.Cells(lnFila, 8) = "TOTAL "
    lnFila = lnFila + 1
'    xlAplicacion.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 11)).Font.Bold = True
'    xlAplicacion.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 11)).HorizontalAlignment = xlCenter
    xlHoja1.Cells(lnFila, 1) = ""
    xlHoja1.Cells(lnFila, 3) = "CALENDARIO"
    xlHoja1.Cells(lnFila, 5) = IIf(Mid(gsOpeCod, 3, 1) = "2", "US$", "S/.")
    xlHoja1.Cells(lnFila, 6) = "ACUMULADO"
    xlHoja1.Cells(lnFila, 7) = "MES"
    xlHoja1.Cells(lnFila, 8) = IIf(Mid(gsOpeCod, 3, 1) = "2", "US$", "S/.")
    Y2 = lnFila
    CuadroExcel 1, Y1, 8, Y2
    
    '
    xlHoja1.Range(xlHoja1.Cells(lnFila - 1, 1), xlHoja1.Cells(lnFila, 8)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(lnFila - 1, 1), xlHoja1.Cells(lnFila, 8)).HorizontalAlignment = xlCenter
    xlHoja1.Range("A" & lnFila - 1 & ":H" & lnFila).Interior.ColorIndex = 36
    xlHoja1.Range("A" & lnFila - 1 & ":H" & lnFila).Font.ColorIndex = 11
    '
    
    Y1 = lnFila + 1
    lnFila = lnFila + 1
    sTotGCCal = "": sTotGCap = "": sTotGInt = "": sTotGIMes = "": sTotGSdo = ""
    sTotCCal = "": sTotCap = "": sTotInt = "": sTotIMes = "": sTotSdo = ""
    For i = 1 To UBound(lsReporteGeneral, 2)
        If lsCodLinea <> lsReporteGeneral(10, i) Or gsPersNombre <> lsReporteGeneral(1, i) Then
            If i > 1 Then
                Y2 = lnFila
                CuadroExcel 1, Y1, 8, Y2
                lnFila = lnFila + 1
                xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 11)).Font.Bold = True
                xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 8)).NumberFormat = "#,##0.00;-#,##0.00"
                If lsCodLinea <> "" Then
                    xlHoja1.Cells(lnFila, 1) = "TOTAL LINEA " & lsDesLinea
                Else
                    xlHoja1.Cells(lnFila, 1) = "TOTAL " & gsPersNombre
                End If
                xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 8)).Formula = "=SUM(C" & Y1 + 1 & ":C" & lnFila - 1 & ")"
                xlHoja1.Cells(lnFila, 4) = ""
                
                'Total General
                sTotGCap = sTotGCap & "+" & Replace(xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 3)).AddressLocal, "$", "")
                
                If lsCodLinea <> "" Then
                    If gsPersNombre <> "" And gsPersNombre <> lsReporteGeneral(1, i) Then
                        lnFila = lnFila + 2
                        xlHoja1.Cells(lnFila, 1) = "TOTAL " & gsPersNombre
                        xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 8)).Formula = "=" & sTotCap
                        xlHoja1.Cells(lnFila, 4) = ""
                        sTotCap = ""
                    Else
                        sTotCap = sTotCap & "+" & Replace(xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 3)).AddressLocal, "$", "")
                    End If
                End If
                Y1 = lnFila
                lnFila = lnFila + 1
            End If
            lnFila = lnFila + 1
            If lsReporteGeneral(10, i) <> "" Then
                xlHoja1.Cells(lnFila, 1) = "LINEA DE CREDITO : " & lsReporteGeneral(12, i)
            End If
            lsCodLinea = lsReporteGeneral(10, i)
            lsDesLinea = lsReporteGeneral(12, i)
            gsPersNombre = lsReporteGeneral(1, i)
        End If
        lnFila = lnFila + 1
        xlHoja1.Cells(lnFila, 1) = lsReporteGeneral(1, i)
        xlHoja1.Cells(lnFila, 2) = lsReporteGeneral(3, i)
        xlAplicacion.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 2)).HorizontalAlignment = xlCenter
        xlHoja1.Cells(lnFila, 3) = lsReporteGeneral(9, i)
        xlHoja1.Cells(lnFila, 4) = lsReporteGeneral(6, i)
        xlHoja1.Cells(lnFila, 5) = lsReporteGeneral(5, i)
        xlHoja1.Cells(lnFila, 6) = lsReporteGeneral(8, i)
        xlHoja1.Cells(lnFila, 7) = lsReporteGeneral(7, i)
        
        xlHoja1.Range(xlHoja1.Cells(lnFila, 8), xlHoja1.Cells(lnFila, 8)).Formula = "=SUM(" & xlHoja1.Range(xlHoja1.Cells(lnFila, 5), xlHoja1.Cells(lnFila, 5)).Address & ":" & xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).Address & ") "
        xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 8)).NumberFormat = "#,##0.00;-#,##0.00"
        xlHoja1.Range(xlHoja1.Cells(lnFila, 4), xlHoja1.Cells(lnFila, 4)).NumberFormat = "#,##0.00####;-#,##0.00####"
    Next
    
    Y2 = lnFila
    CuadroExcel 1, Y1, 8, Y2
    lnFila = lnFila + 1
    xlAplicacion.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 11)).Font.Bold = True
    If lsCodLinea <> "" Then
        xlHoja1.Cells(lnFila, 1) = "TOTAL LINEA " & lsDesLinea
    Else
        xlHoja1.Cells(lnFila, 1) = "TOTAL"
    End If
    xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 8)).Formula = "=SUM(C" & Y1 + 1 & ":C" & lnFila - 1 & ")"
    xlHoja1.Cells(lnFila, 4) = ""
    xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 8)).NumberFormat = "#,##0.00;-#,##0.00"
    'Total General
    sTotGCap = sTotGCap & "+" & Replace(xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 3)).AddressLocal, "$", "")
    'Total IF
    sTotCap = sTotCap & "+" & Replace(xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 3)).AddressLocal, "$", "")
    Y1 = lnFila
    
    lnFila = lnFila + 2
    If lsCodLinea <> "" Then
        xlHoja1.Cells(lnFila, 1) = "TOTAL " & gsPersNombre
        xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 8)).Formula = "=" & sTotCap
        xlHoja1.Cells(lnFila, 4) = ""
        xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 8)).NumberFormat = "#,##0.00;-#,##0.00"
        xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 10)).Font.Bold = True
    End If
    sTotCap = "=": sTotInt = "=": sTotSdo = ""
    Y2 = lnFila
    CuadroExcel 1, Y1, 8, Y2
    Y1 = lnFila
    If sTotGCap <> "" Then
        lnFila = lnFila + 2
        xlHoja1.Cells(lnFila, 1) = "TOTAL ADEUDADOS "
        xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 8)).Formula = "=" & sTotGCap
        xlHoja1.Cells(lnFila, 4) = ""
        xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 8)).NumberFormat = "#,##0.00;-#,##0.00"
        xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 10)).Font.Bold = True
        Y2 = lnFila
        CuadroExcel 1, Y1, 8, Y2
        
        xlHoja1.Range("A" & lnFila & ":H" & lnFila).Interior.ColorIndex = 36
        xlHoja1.Range("A" & lnFila & ":H" & lnFila).Font.ColorIndex = 53
        
    End If
    
    '
    xlHoja1.Cells.Select
    xlHoja1.Cells.Font.Name = "Arial"
    xlHoja1.Cells.Font.Size = 8
    xlHoja1.Cells.EntireColumn.AutoFit
    '
     
      
End Sub


