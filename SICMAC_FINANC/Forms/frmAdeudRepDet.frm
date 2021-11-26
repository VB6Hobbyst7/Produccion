VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.OCX"
Begin VB.Form frmAdeudRepDet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "  ADEUDADOS: REPORTE DETALLADO POR PAGARES"
   ClientHeight    =   3285
   ClientLeft      =   2265
   ClientTop       =   3195
   ClientWidth     =   8115
   Icon            =   "frmAdeudRepDet.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   8115
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkCancelados 
      BackColor       =   &H8000000A&
      Caption         =   "Mostrar Cancelados"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   6165
      TabIndex        =   18
      Top             =   45
      Width           =   1860
   End
   Begin VB.Frame FracuentaHasta 
      Caption         =   "Datos Generales"
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
      Height          =   1065
      Left            =   90
      TabIndex        =   11
      Top             =   1380
      Width           =   7965
      Begin Sicmact.TxtBuscar txtCodObjetoFin 
         Height          =   360
         Left            =   1125
         TabIndex        =   16
         Top             =   255
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   635
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
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta IF:"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   135
         TabIndex        =   15
         Top             =   322
         Width           =   720
      End
      Begin VB.Label lblDescBancoFin 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000080&
         Height          =   300
         Left            =   4890
         TabIndex        =   2
         Top             =   277
         Width           =   2895
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Cuenta :"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   135
         TabIndex        =   14
         Top             =   690
         Width           =   945
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Institución :"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   4020
         TabIndex        =   13
         Top             =   322
         Width           =   810
      End
      Begin VB.Label lblNumCuentaFin 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000080&
         Height          =   300
         Left            =   1140
         TabIndex        =   12
         Top             =   645
         Width           =   6645
      End
   End
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
      TabIndex        =   7
      Top             =   300
      Width           =   7965
      Begin Sicmact.TxtBuscar txtCodObjetoIni 
         Height          =   360
         Left            =   1140
         TabIndex        =   17
         Top             =   225
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   635
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
      End
      Begin VB.Label lblNumCuentaIni 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H8000000D&
         Height          =   300
         Left            =   1155
         TabIndex        =   1
         Top             =   637
         Width           =   6630
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Institución :"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   4035
         TabIndex        =   10
         Top             =   315
         Width           =   810
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Cuenta :"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   135
         TabIndex        =   9
         Top             =   690
         Width           =   960
      End
      Begin VB.Label lblDescBancoIni 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H8000000D&
         Height          =   300
         Left            =   4890
         TabIndex        =   0
         Top             =   262
         Width           =   2895
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta IF:"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   135
         TabIndex        =   8
         Top             =   315
         Width           =   735
      End
   End
   Begin MSComctlLib.ProgressBar Barra 
      Height          =   210
      Left            =   3105
      TabIndex        =   6
      Top             =   3060
      Width           =   4860
      _ExtentX        =   8573
      _ExtentY        =   370
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   345
      Left            =   6705
      TabIndex        =   4
      Top             =   2535
      Width           =   1305
   End
   Begin VB.CommandButton cmdGenerar 
      Caption         =   "&Generar"
      Height          =   345
      Left            =   5355
      TabIndex        =   3
      Top             =   2535
      Width           =   1305
   End
   Begin MSComctlLib.StatusBar Estado 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   5
      Top             =   2970
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   8820
            MinWidth        =   8820
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmAdeudRepDet"
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
Dim lsCadenaPrint As String
Dim lbAdeudados As Boolean
Dim lbLoad As Boolean
Dim lsImpre As String
Dim dbCmact As DConecta
Dim objPista As COMManejador.Pista 'ARLO20170217

Public Sub Inicio(pbAdeudados As Boolean)
    lbAdeudados = pbAdeudados
    Me.Show 1
End Sub


Private Sub chkCancelados_Click()
    Dim oIF As New DCajaCtasIF
    If Me.chkCancelados.value = 1 Then
        Me.txtCodObjetoIni.rs = oIF.CargaCtasIF(Mid(gsOpeCod, 3, 1), lsObjetos(3, 1), MuestraCuentas, , , False)
        Me.txtCodObjetoFin.rs = oIF.CargaCtasIF(Mid(gsOpeCod, 3, 1), lsObjetos(3, 1), MuestraCuentas, , , False)
        Set oIF = Nothing
    Else
        Me.txtCodObjetoIni.rs = oIF.CargaCtasIF(Mid(gsOpeCod, 3, 1), lsObjetos(3, 1), MuestraCuentas)
        Me.txtCodObjetoFin.rs = oIF.CargaCtasIF(Mid(gsOpeCod, 3, 1), lsObjetos(3, 1), MuestraCuentas)
        Set oIF = Nothing
    End If

End Sub

Private Sub cmdGenerar_Click()
    On Error GoTo ErrorGenerar
    Dim lbDatos1 As Boolean
    Dim lbDatos2 As Boolean
    
    lbExcel = False
    ReDim lsReporteGeneral(9, 0)
    n = 0
    If lbAdeudados Then
        If DatosReporteGeneral1(Me.txtCodObjetoIni, Me.txtCodObjetoFin, lsCtaContDebe(1, 1), lsCtaContHaber(1, 1)) = False Then
            MsgBox "No se han encontrado Datos para procesar Reporte", vbInformation, "Aviso"
            Exit Sub
        End If
    Else
        If DatosReporteGeneralPF(Me.txtCodObjetoIni, Me.txtCodObjetoFin, lsCtaContDebe(1, 1), lsCtaContHaber(1, 1), Format(gdFecSis, gsFormatoFecha)) = False Then
            MsgBox "No se han encontrado Datos para procesar Reporte", vbInformation, "Aviso"
            Exit Sub
        End If
    End If
    EnviaPrevio lsImpre, gsOpeDesc, gnLinPage, True
                'ARLO20170217
                Set objPista = New COMManejador.Pista
                objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "", Me.Caption & " Se Genero Reporte "
                Set objPista = Nothing
                '****
    Exit Sub
ErrorGenerar:
    MsgBox "Error N°[" & Err.Number & "] " & Err.Description, vbInformation, "Aviso"
  
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
    Set dbCmact = New DConecta
    dbCmact.AbreConexion
    CentraForm Me
    Me.Caption = gsOpeDesc
    ReDim lsCtaContDebe(2, 0)
    ReDim lsCtaContHaber(2, 0)

    lbLoad = True
    sql = "SELECT * FROM OpeCta WHERE cOpeCod ='" & gsOpeCod & "' ORDER BY cOpeCtaOrden"
    Set rs = dbCmact.CargaRecordSet(sql)
    If Not RSVacio(rs) Then
        Do While Not rs.EOF
            If Trim(rs!cOpeCtaDH) = "D" Then
                n = n + 1
                ReDim Preserve lsCtaContDebe(2, n)
                lsCtaContDebe(1, n) = Trim(rs!cCtaContCod)
                lsCtaContDebe(2, n) = Trim(rs!cOpeCtaOrden)
            Else
                m = m + 1
                ReDim Preserve lsCtaContHaber(2, m)
                lsCtaContHaber(1, m) = Trim(rs!cCtaContCod)
                lsCtaContHaber(2, m) = Trim(rs!cOpeCtaOrden)
            End If
            rs.MoveNext
        Loop
    Else
        RSClose rs
        MsgBox "No se han definido Cuentas Contables", vbInformation, "Aviso"
        lbLoad = False
        Exit Sub
    End If
    RSClose rs
    If n = 0 Then
        MsgBox "No se han definido Cuentas de Adeudados en el Debe"
        lbLoad = False
        Exit Sub
    End If
    If m = 0 Then
        MsgBox "No se han definido Cuentas de Intereses en el Haber"
        lbLoad = False
        Exit Sub
    End If
    ReDim lsObjetos(4, 0)
    n = 0
    sql = "Select * from OpeObj where cOpeCod ='" & gsOpeCod & "'"
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
    
    Select Case gsOpeCod
        Case OpeCGAdeudRepDetalleMN, OpeCGADeudRepDetalleME
            Me.chkCancelados.Visible = True
        Case Else
            Me.chkCancelados.Visible = False
    End Select

    Dim oIF As New DCajaCtasIF
    Me.txtCodObjetoIni.rs = oIF.CargaCtasIF(Mid(gsOpeCod, 3, 1), lsObjetos(3, 1), MuestraCuentas)
    Me.txtCodObjetoFin.rs = oIF.CargaCtasIF(Mid(gsOpeCod, 3, 1), lsObjetos(3, 1), MuestraCuentas)
    Set oIF = Nothing

End Sub
Private Function Encabezado(lsBanco As String, lnInteres As Currency, lsCuenta As String, lsApertura As String, lsVencimiento As String, lnPlazo As Long, lnGracia As Long, lnPlazoInt As Long, Optional lbGeneral As Boolean = True) As Integer
    If lbGeneral Then
        lsCadenaPrint = lsCadenaPrint & oImpresora.gPrnCondensadaON & oImpresora.gPrnBoldON + ImpreFormat(gsNomCmac, 100) & ImpreFormat("Fecha :" & gdFecSis & " - Area Caja General ", 50) & oImpresora.gPrnSaltoLinea
        lsCadenaPrint = lsCadenaPrint & oImpresora.gPrnSaltoLinea
        lsCadenaPrint = lsCadenaPrint & CentrarCadena("REPORTE DETALLADO DE PAGARES DE ADEUDADOS", 150) + oImpresora.gPrnSaltoLinea
        lsCadenaPrint = lsCadenaPrint & CentrarCadena("=========================================", 150) + PrnSet("B-") + oImpresora.gPrnSaltoLinea
        lsCadenaPrint = lsCadenaPrint & CentrarCadena(IIf(Mid(gsOpeCod, 3, 1) = "1", "MONEDA NACIONAL", "MONEDA EXTRANJERA"), 150) + oImpresora.gPrnSaltoLinea
    End If
    lsCadenaPrint = lsCadenaPrint & oImpresora.gPrnSaltoLinea
    lsCadenaPrint = lsCadenaPrint + PrnSet("B+") + ImpreFormat("Banco          :" & Trim(lsBanco), 100) & ImpreFormat("Pagare      :" & Trim(lsCuenta), 50) + PrnSet("B-") + oImpresora.gPrnSaltoLinea
    lsCadenaPrint = lsCadenaPrint & ImpreFormat("Tasa Interes   :" & Format(lnInteres, "#,#0.00") & "% " & IIf(lnPlazoInt < 360, "Mensual", "Anual"), 100) & ImpreFormat("N° Cuotas   :" & Format(lnPlazo, "#,#0") & " ", 50) & oImpresora.gPrnSaltoLinea
    lsCadenaPrint = lsCadenaPrint & ImpreFormat("Apertura       :" & Format(lsApertura, "dd/mm/yyyy"), 100) & ImpreFormat("Fecha Venc. :" & Format(lsVencimiento, "dd/mm/yyyy"), 50) & oImpresora.gPrnSaltoLinea
    lsCadenaPrint = lsCadenaPrint & ImpreFormat("Vencimiento a  :" & Format(lnGracia, "#,#0") & " Dias ", 100) & oImpresora.gPrnSaltoLinea
    lsCadenaPrint = lsCadenaPrint + PrnSet("B+") + PrnSet("Esp", 6) & String(155, "-") & oImpresora.gPrnSaltoLinea
    lsCadenaPrint = lsCadenaPrint & ImpreFormat(" ", 122) & ImpreFormat("INTERES", 22) & oImpresora.gPrnSaltoLinea
    lsCadenaPrint = lsCadenaPrint & ImpreFormat("FECHA", 12) & ImpreFormat("MOVIMIENTO", 60) & ImpreFormat("DESEMBOLSO", 13) & ImpreFormat("AMORTIZACION", 13) & ImpreFormat(String(30, "-"), 35) & ImpreFormat("SALDO", 15) & oImpresora.gPrnSaltoLinea
    lsCadenaPrint = lsCadenaPrint & ImpreFormat(" ", 115) & ImpreFormat("Pagado.", 12) & ImpreFormat("Provisionado", 12) & oImpresora.gPrnSaltoLinea
    lsCadenaPrint = lsCadenaPrint & String(155, "-") & PrnSet("EspN") + PrnSet("B-") & oImpresora.gPrnSaltoLinea
    Encabezado = 15
End Function

Private Function EncabezadoPF(lsBanco As String, lnInteres As Currency, lsCuenta As String, lsApertura As String, lsVencimiento As String, lnPlazo As Long, Optional lbGeneral As Boolean = True) As Integer
If lbGeneral Then
    lsCadenaPrint = lsCadenaPrint + PrnSet("B+") + ImpreFormat(gsNomCmac, 80) & ImpreFormat("Fecha :" & gdFecSis & " - Area Caja General ", 50) & oImpresora.gPrnSaltoLinea
    lsCadenaPrint = lsCadenaPrint & oImpresora.gPrnSaltoLinea
    lsCadenaPrint = lsCadenaPrint & CentrarCadena("REPORTE DETALLADO DE CUENTAS PLAZO FIJO", 120) + oImpresora.gPrnSaltoLinea
    lsCadenaPrint = lsCadenaPrint & CentrarCadena("=======================================", 120) + PrnSet("B-") + oImpresora.gPrnSaltoLinea
End If
lsCadenaPrint = lsCadenaPrint & oImpresora.gPrnSaltoLinea
lsCadenaPrint = lsCadenaPrint + PrnSet("B+") + ImpreFormat("Institucion Financiera :" & Trim(lsBanco), 100) & ImpreFormat("Cuenta :" & Trim(lsCuenta), 50) + PrnSet("B-") + oImpresora.gPrnSaltoLinea
lsCadenaPrint = lsCadenaPrint & ImpreFormat("T.E.A.     :" & Format(lnInteres, "#,#0.00") & "%", 100) & ImpreFormat("Plazo  :" & Format(lnPlazo, "#,#0"), 50) & oImpresora.gPrnSaltoLinea
lsCadenaPrint = lsCadenaPrint & ImpreFormat("Apertura   :" & Format(lsApertura, "dd/mm/yyyy"), 100) & ImpreFormat("Vencimiento :" & Format(lsVencimiento, "dd/mm/yyyy"), 50) & oImpresora.gPrnSaltoLinea
lsCadenaPrint = lsCadenaPrint & PrnSet("Esp", 7) & String(150 - 22, "-") & oImpresora.gPrnSaltoLinea

'lsCadenaPrint = lsCadenaPrint & ImpreFormat(" ", 112) & ImpreFormat("HABER", 22) & oImpresora.gPrnSaltoLinea
'LsCadenaPrint = lsCadenaPrint & ImpreFormat("FECHA", 12) & ImpreFormat("MOVIMIENTO", 70) & ImpreFormat("DEBE", 13) & ImpreFormat(String(30, "-"), 35) & ImpreFormat("TOTAL", 15) & oImpresora.gPrnSaltoLinea
'lsCadenaPrint = lsCadenaPrint & ImpreFormat(" ", 100) & ImpreFormat("Int. Acum.", 15) & ImpreFormat("Int.Calc.", 15) & oImpresora.gPrnSaltoLinea

lsCadenaPrint = lsCadenaPrint & "                                                                                                HABER                        " & oImpresora.gPrnSaltoLinea
lsCadenaPrint = lsCadenaPrint & "  FECHA       MOVIMIENTO                                                  DEBE    ------------------------------        TOTAL" & oImpresora.gPrnSaltoLinea
lsCadenaPrint = lsCadenaPrint & "                                                                                      Int. Acum.      Int. Calc.             " & oImpresora.gPrnSaltoLinea

lsCadenaPrint = lsCadenaPrint & String(150 - 22, "-") & PrnSet("EspN") & oImpresora.gPrnSaltoLinea
EncabezadoPF = 13
End Function

Private Sub Form_Unload(Cancel As Integer)
    CierraConexion
End Sub

Private Function DatosReporteGeneral1(lsObjetoDesde As String, lsObjetoHasta As String, lsCtaContEntidad As String, lsCtaIntCalc As String) As Boolean
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim rs1 As New ADODB.Recordset
    Dim lsFiltro As String
    Dim lnCapital As Currency
    Dim lnCapitalInicial As Currency
    Dim lnAmortiza As Currency
    Dim lnDevengados As Currency
    Dim lnTotal As Integer, j As Integer
    Dim lnDesembolso As Currency, lnHaber As Currency
    Dim lnInteres As Currency
    Dim lnInteresProv As Currency
    Dim lnInteresDev As Currency
    Dim lnLineas As Long
    Dim lsFecha As String
    Dim lsMovimiento As String
    Dim lsDesMov As String, lsDato As String

    Dim lnTotalAmortizacion As Currency
    Dim lnTotalIntPagado As Currency
    Dim lnTotalIntProv As Currency
    Dim lsCtaIntProv  As String

    Dim oSdo As NCajaCtaIF
    Set oSdo = New NCajaCtaIF
    lsFiltro = ""
    DatosReporteGeneral1 = False
    lnLineas = 0
    Barra.value = 0
    Estado.Panels(1).Text = ""
    
    Dim oIF As New NCajaAdeudados
    Set rs = oIF.CargaDatosGeneralesCtaIF(lsObjetoDesde, lsObjetos(3, 1), lsObjetoHasta, , gdFecSis, gsOpeCod)
    lnTotal = rs.RecordCount
    j = 0
    lsCadenaPrint = ""
    lsCtaIntProv = lsCtaContHaber(1, 2)
    If Not RSVacio(rs) Then
        Do While Not rs.EOF
            j = j + 1
            sql = " SELECT  M.CMOVNRO, " _
               & "     ABS(SUM(CASE WHEN oc.cOpeCtaOrden = 0 THEN " & IIf(Mid(gsOpeCod, 3, 1) = "2", " ISNULL(ME.nMovMEImporte,0) ", " MC.nMovImporte ") & " ELSE 0 END)) AS nAmortiza, " _
               & "     ABS(SUM(CASE WHEN oc.cOpeCtaOrden = 1 THEN " & IIf(Mid(gsOpeCod, 3, 1) = "2", " ISNULL(ME.nMovMEImporte,0) ", " MC.nMovImporte ") & " ELSE 0 END)) AS nIntPago, " _
               & "     ABS(SUM(CASE WHEN oc.cOpeCtaOrden = 2 THEN " & IIf(Mid(gsOpeCod, 3, 1) = "2", " ISNULL(ME.nMovMEImporte,0) ", " MC.nMovImporte ") & " ELSE 0 END)) AS nIntProv, " _
               & "     M.cMovDesc as Movimiento, CASE WHEN cia.cMonedaPago = '2' and SubString(cia.cCtaIFCod, 3, 1) = '1' THEN i.nIndiceVac ELSE 1 END nFactor " _
               & " FROM    MOV M JOIN MOVCTA MC   ON MC.nMOVNRO= M.nMOVNRO " _
               & "          LEFT JOIN MOVME ME    ON ME.nMOVNRO=MC.nMOVNRO AND MC.nMOVITEM=ME.nMOVITEM " _
               & "               JOIN MOVOBJIF MO ON MO.nMOVNRO=MC.nMOVNRO AND MO.nMOVITEM=MC.nMOVITEM " _
               & "          LEFT JOIN CtaIFAdeudados cia ON cia.cIFTpo = mo.cIFTpo and cia.cPersCod = mo.cPersCod and cia.cCtaIFCod = mo.cCtaIFCod " _
               & "               JOIN OpeCta OC   ON MC.cCtaContCod LIKE OC.cCtaContCod + '%' " _
               & "          LEFT JOIN IndiceVac i ON i.dIndiceVac = Convert(datetime,LEFT(m.cMovNro,8) ) " _
               & " WHERE  MO.cIFTpo = '" & rs!cIFTpo & "' and MO.cPersCod = '" & Trim(rs!cPersCod) & "' and MO.cCtaIFCod = '" & rs!cCtaIfCod & "' " _
               & "    AND M.nMOVFLAG NOT IN(1, 2, 3, 5) and m.nMovEstado = " & gMovEstContabMovContable & " " _
               & " AND SUBSTRING(M.CMOVNRO,1,8)> '" & Format(rs!dCtaIFAper, "yyyymmdd") & "' and oc.cOpeCod = '" & gsOpeCod & "' "
               
            sql = sql & " And M.nMovNro Not IN(682714, 683330, 692474, 718469, 718476)" ' Arrastran datos antiguos

               
            sql = sql & " GROUP BY M.CMOVNRO, M.cMovDesc, CASE WHEN cia.cMonedaPago = '2' and SubString(cia.cCtaIFCod, 3, 1) = '1' THEN i.nIndiceVac ELSE 1 END " _
               & " ORDER BY M.CMOVNRO "
     
                lnCapitalInicial = IIf(Mid(gsOpeCod, 3, 1) = "1", rs!nMontoPrestado, rs!nMontoPrestado)
                lnCapital = 0
                DatosReporteGeneral1 = True
                If lbAdeudados Then
                    lnLineas = lnLineas + Encabezado(Trim(rs!cPersNombre), rs!nTasaInteres, rs!cCtaIFDesc, rs!dCtaIFAper, IIf(IsNull(rs!dCtaIFVenc), "__/__/____", rs!dCtaIFVenc), rs!nCtaIFCuotas, rs!nPeriodoGracia, rs!nCtaIFIntPeriodo, IIf(lnLineas = 0, True, False))
                Else
                    lnLineas = lnLineas + EncabezadoPF(Trim(rs!cPersNombre), rs!nTasaInteres, rs!cCtaIFDesc, rs!dCtaIFAper, IIf(IsNull(rs!dCtaIFVenc), "__/__/____", rs!dCtaIFVenc), rs!nPlazo, IIf(lnLineas = 0, True, False))
                End If
                lnDesembolso = lnCapitalInicial
                lnAmortiza = 0
                lnInteresDev = 0
                lnInteresProv = 0
                lnTotalAmortizacion = 0
                lnTotalIntPagado = 0
                lnTotalIntProv = 0

                lnCapital = lnCapitalInicial
     
                lsCadenaPrint = lsCadenaPrint & ImpreFormat(Format(rs!dCtaIFAper, "dd/mm/yyyy"), 12) & ImpreFormat("Desembolso de Pagaré", 58) & ImpreFormat(lnDesembolso, 15, , True) & ImpreFormat(lnAmortiza, 15, , True) & ImpreFormat(lnInteresDev, 12, , True) & ImpreFormat(lnInteresProv, 12, , True) & ImpreFormat(lnCapital, 15, , True) & oImpresora.gPrnSaltoLinea
                lnLineas = lnLineas + 1
                lnDesembolso = 0
 
                lnInteresDev = 0:             lnInteresProv = 0:             lnAmortiza = 0
                lnTotalAmortizacion = 0:         lnTotalIntPagado = 0:     lnTotalIntProv = 0
            
            Set rs1 = dbCmact.CargaRecordSet(sql)
            If Not RSVacio(rs1) Then
     
                Do While Not rs1.EOF
                    lsDato = Trim(rs1!cMovNro)
                    lnTotalAmortizacion = lnTotalAmortizacion + rs1!nAmortiza
                    If rs1!nAmortiza <> 0 Then
                        lnTotalIntPagado = lnTotalIntPagado + rs1!nIntPago
                    End If
                    lnTotalIntProv = lnTotalIntProv + rs1!nIntProv
                    lsFecha = Mid(rs1!cMovNro, 7, 2) & "/" & Mid(rs1!cMovNro, 5, 2) & "/" & Mid(rs1!cMovNro, 1, 4)
                    lsDesMov = Trim(rs1!Movimiento)
                    If rs1!nFactor > 1 Then
                       lnCapital = oSdo.GetSaldoCtaIfcalendario(rs!cPersCod, rs!cIFTpo, rs!cCtaIfCod, CDate(lsFecha))
                       lnCapital = Round(lnCapital * rs1!nFactor, 2)
                    Else
                       lnCapital = lnCapital - rs1!nAmortiza
                    End If
                    lsCadenaPrint = lsCadenaPrint & ImpreFormat(lsFecha, 12) & ImpreFormat(lsDesMov, 58) & ImpreFormat(lnDesembolso, 15, , True) & ImpreFormat(rs1!nAmortiza, 15, , True) & ImpreFormat(Abs(IIf(rs1!nAmortiza <> 0, rs1!nIntPago, 0)), 12, , True) & ImpreFormat(Abs(rs1!nIntProv), 12, , True) & ImpreFormat(lnCapital, 15, , True) & oImpresora.gPrnSaltoLinea
                    lnLineas = lnLineas + 1
                    If lnLineas > gnLinPage - 4 Then
                        lnLineas = 0
                        lsCadenaPrint = lsCadenaPrint & oImpresora.gPrnSaltoPagina
                        lnLineas = lnLineas + Encabezado(Trim(rs!cPersNombre), rs!nTasaInteres, rs!cCtaIFDesc, rs!dCtaIFAper, IIf(IsNull(rs!dCtaIFVenc), "__/__/____", rs!dCtaIFVenc), rs!nCtaIFCuotas, rs!nPeriodoGracia, rs!nCtaIFIntPeriodo, IIf(lnLineas = 0, True, False))
                    End If
             
                    rs1.MoveNext
                Loop
     
                If lnLineas > 62 Then
                    lnLineas = 0
                    lsCadenaPrint = lsCadenaPrint & oImpresora.gPrnSaltoPagina
                    lnLineas = lnLineas + Encabezado(Trim(rs!EntidadFinan), rs!Interes, rs!Cuenta, rs!dCtaIFAper, rs!FechVenc, rs!Plazo, rs!PerGracia, rs!PlazoInt, IIf(lnLineas = 0, True, False))
                End If
                lsCadenaPrint = lsCadenaPrint & String(160, "=") & oImpresora.gPrnSaltoLinea
                lnLineas = lnLineas + 1
                lsCadenaPrint = lsCadenaPrint + PrnSet("B+") + ImpreFormat("", 74) & ImpreFormat("TOTAL  :", 16, , True) & ImpreFormat(lnTotalAmortizacion, 15, , True) & ImpreFormat(Abs(lnTotalIntPagado), 12, , True) & ImpreFormat(Abs(lnTotalIntProv), 12, , True) & ImpreFormat("", 15, , True) + PrnSet("B-") + oImpresora.gPrnSaltoLinea
                lnLineas = lnLineas + 1
            End If
            RSClose rs1
            Me.Barra.value = Int(j / lnTotal * 100)
            Me.Estado.Panels(1).Text = "Generando Reporte :" & Format(j / lnTotal * 100, "#0.00") & "%. Espere... "
            DoEvents
            rs.MoveNext
            If lnLineas > gnLinPage - 4 Then
                lnLineas = 0
                lsCadenaPrint = lsCadenaPrint & oImpresora.gPrnSaltoPagina
            End If
        Loop
        lsCadenaPrint = lsCadenaPrint & oImpresora.gPrnCondensadaOFF
    End If
    RSClose rs
    lsImpre = lsCadenaPrint
End Function

Private Function DatosReporteGeneralPF(lsObjetoDesde As String, lsObjetoHasta As String, lsCtaContEntidad As String, lsCtaIntCalc As String, lsCtaIntFecha As String) As Boolean
Dim sql As String
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim lsFiltro As String
Dim lnCapital As Currency
Dim lnCapitalInicial As Currency
Dim lnDevengados As Currency
Dim lnTotal As Integer, j As Integer
Dim lnDebe As Currency, lnHaber As Currency
Dim lnInteres As Currency
Dim lnInteresAcum As Currency
Dim lnInteresCalc As Currency
Dim lnLineas As Long
Dim lsFecha As String

Dim lsCtaInt As String
Dim lsCtaIng As String

lsFiltro = ""
DatosReporteGeneralPF = False
lnLineas = 0
Barra.value = 0
Estado.Panels(1).Text = ""

Dim oIF As New NCajaAdeudados
Set rs = oIF.CargaDatosGeneralesCtaIFPlazoFijo(lsObjetoDesde, lsObjetos(3, 1), gsOpeCod, lsObjetoHasta)
lnTotal = rs.RecordCount
j = 0
lsCadenaPrint = ""
If Not RSVacio(rs) Then
    Do While Not rs.EOF
        DatosReporteGeneralPF = True
        j = j + 1
        lnCapitalInicial = IIf(Mid(gsOpeCod, 3, 1) = "1", rs!nMontoIni, rs!nMontoIniME)
        lnLineas = lnLineas + EncabezadoPF(Trim(rs!cPersNombre), rs!nTasaInteres, rs!cCtaIFDesc, rs!dCtaIFAper, rs!dCtaIFVenc, IIf(IsNull(rs!nCtaIFIntPeriodo), 0, rs!nCtaIFIntPeriodo), IIf(lnLineas = 0, True, False))
       
        'Elimina por ser duplicado
        'lnCapital = lnCapitalInicial
        'lsCadenaPrint = lsCadenaPrint & ImpreFormat(Format(rs!dCtaIFAper, "dd/mm/yyyy"), 10) & ImpreFormat("Apertura de Cuenta", 50) & ImpreFormat(lnCapitalInicial, 13, , True) & ImpreFormat(0, 13, , True) & ImpreFormat(0, 13, , True) & ImpreFormat(lnCapital, 13, , True) & oImpresora.gPrnSaltoLinea
        
        lnLineas = lnLineas + 1

'1 - Interes Devengado
'2 - Interes Clase 5
            
        sql = " SELECT  M.CMOVNRO, " _
           & "     SUM(CASE WHEN oc.cOpeCtaOrden = 0 THEN " & IIf(Mid(gsOpeCod, 3, 1) = gMonedaExtranjera, "ISNULL(ME.nMovMEImporte,0)", "ISNULL(MC.nMovImporte,0)") & " ELSE 0 END) AS nCapitaliza, " _
           & "     SUM(CASE WHEN oc.cOpeCtaOrden = 1 THEN " & IIf(Mid(gsOpeCod, 3, 1) = gMonedaExtranjera, "ISNULL(ME.nMovMEImporte,0)", "ISNULL(MC.nMovImporte,0)") & " ELSE 0 END) AS nIntPago, " _
           & "     SUM(CASE WHEN oc.cOpeCtaOrden = 2 THEN " & IIf(Mid(gsOpeCod, 3, 1) = gMonedaExtranjera, "ISNULL(ME.nMovMEImporte,0)", "ISNULL(MC.nMovImporte,0)") & " ELSE 0 END) AS nIntProv, " _
           & "     M.cMovDesc as Movimiento " _
           & " FROM    MOV M JOIN MOVCTA MC   ON MC.nMOVNRO= M.nMOVNRO " _
           & IIf(Mid(gsOpeCod, 3, 1) = gMonedaExtranjera, " JOIN MOVME ME    ON ME.nMOVNRO=MC.nMOVNRO AND MC.nMOVITEM=ME.nMOVITEM ", "") _
           & "               JOIN MOVOBJIF MO ON MO.nMOVNRO=MC.nMOVNRO AND MO.nMOVITEM=MC.nMOVITEM " _
           & "               JOIN OpeCta OC   ON MC.cCtaContCod LIKE OC.cCtaContCod + '%' " _
           & " WHERE  MO.cIFTpo = '" & rs!cIFTpo & "' and MO.cPersCod = '" & Trim(rs!cPersCod) & "' and MO.cCTAIFCod = '" & rs!cCtaIfCod & "' " _
           & "    AND M.nMOVFLAG NOT IN(1, 2, 3, 5) and m.nMovEstado = " & gMovEstContabMovContable & " " _
           & " AND SUBSTRING(M.CMOVNRO,1,8)>='" & Format(rs!dCtaIFAper, "yyyymmdd") & "' and oc.cOpeCod = '" & gsOpeCod & "' " & " AND M.NMOVNRO not in(621231, 668578)  "
           
        sql = sql & " GROUP BY M.CMOVNRO, M.cMovDesc " _
           & " ORDER BY M.CMOVNRO "
        Set rs1 = dbCmact.CargaRecordSet(sql)
        If Not RSVacio(rs1) Then
            Do While Not rs1.EOF
                lnInteres = 0
                If rs1!nIntPago > 0 Then
                    lnInteresCalc = rs1!nIntPago
                    lnInteresAcum = 0
                Else
                    lnInteresCalc = rs1!nIntProv
                    lnInteresAcum = rs1!nIntPago
                End If
                If rs1!nCapitaliza <> 0 Then
                    lnInteres = rs1!nCapitaliza
                    lnCapital = lnCapital + lnInteres
                End If
                lsFecha = GetFechaMov(rs1!cMovNro, True)
                lsCadenaPrint = lsCadenaPrint & ImpreFormat(lsFecha, 10) & ImpreFormat(rs1!Movimiento, 50) & ImpreFormat(lnInteres, 13, , True) & ImpreFormat(lnInteresAcum, 13, , True) & ImpreFormat(lnInteresCalc, 13, , True) & ImpreFormat(lnCapital, 13, , True) & oImpresora.gPrnSaltoLinea
                lnLineas = lnLineas + 1
                If lnLineas > gnLinPage - 4 Then
                    lnLineas = 0
                    lsCadenaPrint = lsCadenaPrint & oImpresora.gPrnSaltoPagina
                    lnLineas = lnLineas + EncabezadoPF(Trim(rs!cPersNombre), rs!nTasaInteres, rs!cCtaIFDesc, rs!dCtaIFAper, rs!dCtaIFVenc, IIf(IsNull(rs!nCtaIFIntPeriodo), 0, rs!nCtaIFIntPeriodo), IIf(lnLineas = 0, True, False))
                End If
                rs1.MoveNext
            Loop
        End If
        rs1.Close
        Set rs1 = Nothing
        lsCadenaPrint = lsCadenaPrint & String(150 - 22, "=") & oImpresora.gPrnSaltoLinea
        lnLineas = lnLineas + 1
        Me.Barra.value = Int(j / lnTotal * 100)
        Me.Estado.Panels(1).Text = "Generando Reporte :" & Format(j / lnTotal * 100, "#0.00") & "%. Espere... "
        DoEvents
        rs.MoveNext
        If lnLineas > 62 Then
            lnLineas = 0
            lsCadenaPrint = lsCadenaPrint & oImpresora.gPrnSaltoPagina
        End If
    Loop
End If
RSClose rs
lsImpre = lsCadenaPrint
End Function

Private Sub txtCodObjetoFin_EmiteDatos()
Dim oIF As New NCajaCtaIF
If txtCodObjetoFin <> "" Then
    Me.lblDescBancoFin = oIF.GetIFDesc(Mid(txtCodObjetoFin, 4, 13), True, Mid(txtCodObjetoFin, 18, 5))
    Me.lblNumCuentaFin = txtCodObjetoFin.psDescripcion
    Me.CmdGenerar.SetFocus
End If
Set oIF = Nothing
End Sub

Private Sub txtCodObjetoIni_EmiteDatos()
Dim oIF As New NCajaCtaIF
If txtCodObjetoIni <> "" Then
    Me.lblDescBancoIni = oIF.GetIFDesc(Mid(txtCodObjetoIni, 4, 13), True, Mid(txtCodObjetoIni, 18, 5))
    Me.lblNumCuentaIni = txtCodObjetoIni.psDescripcion
    Me.txtCodObjetoFin.SetFocus
End If
Set oIF = Nothing
End Sub
