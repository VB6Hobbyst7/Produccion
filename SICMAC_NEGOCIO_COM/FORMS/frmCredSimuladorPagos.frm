VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCredSimuladorPagos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Simulador de Pagos"
   ClientHeight    =   5385
   ClientLeft      =   4980
   ClientTop       =   3585
   ClientWidth     =   4830
   Icon            =   "frmCredSimuladorPagos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   60
      TabIndex        =   23
      Top             =   4740
      Width           =   4680
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   345
         Left            =   75
         TabIndex        =   25
         Top             =   195
         Width           =   1155
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   345
         Left            =   3450
         TabIndex        =   24
         Top             =   180
         Width           =   1155
      End
   End
   Begin VB.Frame Frame3 
      Height          =   2505
      Left            =   90
      TabIndex        =   9
      Top             =   2235
      Width           =   4665
      Begin VB.CommandButton CmdGastos 
         Caption         =   "Gastos"
         Height          =   345
         Left            =   3540
         TabIndex        =   30
         Top             =   2100
         Width           =   1005
      End
      Begin VB.CommandButton CmdAplicar 
         Caption         =   "&Aplicar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2760
         TabIndex        =   13
         Top             =   225
         Width           =   1290
      End
      Begin MSMask.MaskEdBox TxtFechaPago 
         Height          =   300
         Left            =   1440
         TabIndex        =   11
         Top             =   240
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblDeuda 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   315
         Left            =   1605
         TabIndex        =   34
         Top             =   2025
         Width           =   1215
      End
      Begin VB.Label Label12 
         Caption         =   "Deuda a la Fecha :"
         Height          =   195
         Left            =   180
         TabIndex        =   33
         Top             =   2040
         Width           =   1590
      End
      Begin VB.Label Label13 
         Caption         =   "ITF :"
         Height          =   240
         Left            =   2940
         TabIndex        =   32
         Top             =   930
         Width           =   705
      End
      Begin VB.Label lblITF 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   315
         Left            =   3705
         TabIndex        =   31
         Top             =   900
         Width           =   855
      End
      Begin VB.Label LblMetLiq 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   315
         Left            =   3705
         TabIndex        =   29
         Top             =   1275
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "Met. Liq. :"
         Height          =   240
         Left            =   2940
         TabIndex        =   28
         Top             =   1305
         Width           =   705
      End
      Begin VB.Label Label11 
         Caption         =   "Mora :"
         Height          =   240
         Left            =   2940
         TabIndex        =   27
         Top             =   1665
         Width           =   540
      End
      Begin VB.Label LblMora 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   315
         Left            =   3705
         TabIndex        =   26
         Top             =   1635
         Width           =   855
      End
      Begin VB.Label LblGastos 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   315
         Left            =   1440
         TabIndex        =   18
         Top             =   1628
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Gastos               :"
         Height          =   240
         Left            =   165
         TabIndex        =   17
         Top             =   1665
         Width           =   1245
      End
      Begin VB.Label LblInteres 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   315
         Left            =   1440
         TabIndex        =   16
         Top             =   1268
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Interes               :"
         Height          =   240
         Left            =   165
         TabIndex        =   15
         Top             =   1305
         Width           =   1245
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000E&
         X1              =   90
         X2              =   4575
         Y1              =   735
         Y2              =   735
      End
      Begin VB.Line Line1 
         X1              =   90
         X2              =   4575
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label LblMontoPago 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   315
         Left            =   1440
         TabIndex        =   14
         Top             =   915
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Monto a Pagar  :"
         Height          =   240
         Left            =   165
         TabIndex        =   12
         Top             =   945
         Width           =   1245
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha de Pago :"
         Height          =   240
         Left            =   165
         TabIndex        =   10
         Top             =   255
         Width           =   1245
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos"
      Height          =   1725
      Left            =   90
      TabIndex        =   0
      Top             =   495
      Width           =   4665
      Begin VB.Label LblFechaPend 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   270
         Left            =   3390
         TabIndex        =   22
         Top             =   1305
         Width           =   1125
      End
      Begin VB.Label Label10 
         Caption         =   "Cuota Pendiente :"
         Height          =   240
         Left            =   2055
         TabIndex        =   21
         Top             =   1320
         Width           =   1305
      End
      Begin VB.Label LblCuotaPend 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
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
         Height          =   285
         Left            =   1440
         TabIndex        =   20
         Top             =   1290
         Width           =   390
      End
      Begin VB.Label Label7 
         Caption         =   "Cuota Pendiente :"
         Height          =   240
         Left            =   120
         TabIndex        =   19
         Top             =   1320
         Width           =   1305
      End
      Begin VB.Label LblSaldo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3375
         TabIndex        =   8
         Top             =   900
         Width           =   1095
      End
      Begin VB.Label Saldo 
         AutoSize        =   -1  'True
         Caption         =   "Saldo :"
         Height          =   195
         Left            =   2070
         TabIndex        =   7
         Top             =   900
         Width           =   495
      End
      Begin VB.Label LblPrestamo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   945
         TabIndex        =   6
         Top             =   885
         Width           =   900
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Prestamo :"
         Height          =   195
         Left            =   105
         TabIndex        =   5
         Top             =   885
         Width           =   750
      End
      Begin VB.Label LblAnalista 
         Height          =   195
         Left            =   855
         TabIndex        =   4
         Top             =   540
         Width           =   3645
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Analista :"
         Height          =   195
         Left            =   90
         TabIndex        =   3
         Top             =   540
         Width           =   645
      End
      Begin VB.Label lblTitular 
         Height          =   195
         Left            =   705
         TabIndex        =   2
         Top             =   285
         Width           =   3795
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Titular :"
         Height          =   195
         Left            =   90
         TabIndex        =   1
         Top             =   285
         Width           =   525
      End
   End
   Begin SICMACT.ActXCodCta ActxCta 
      Height          =   435
      Left            =   75
      TabIndex        =   35
      Top             =   75
      Width           =   3660
      _ExtentX        =   6456
      _ExtentY        =   767
      Texto           =   "Credito :"
      EnabledCMAC     =   -1  'True
      EnabledCta      =   -1  'True
      EnabledProd     =   -1  'True
      EnabledAge      =   -1  'True
   End
End
Attribute VB_Name = "frmCredSimuladorPagos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nMontoGasto As Double
Dim nMontoIntCompVencCierre As Double
Dim psCtaCod As String
Dim vnTasa As Double
Dim MatCalendParalelo As Variant
Dim MatCalend As Variant
Dim oNegCredito As COMNCredito.NCOMCredito
Dim nTasaMorat As Double
Dim MatGastosCargados(100, 2) As String
Dim nNumGasCarg As Integer
Dim lsTipoProducto As String

Private Sub ActxCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CargaDatos ActxCta.NroCuenta
    End If
End Sub

'Private Function MontoTotalGastosGenerado(ByVal MatGastos As Variant, ByVal pnNumGastosCancel As Integer, _
'    Optional ByVal psTipoGastoProc As Variant = "") As Double
'Dim i As Integer
'    MontoTotalGastosGenerado = 0
'    For i = 0 To pnNumGastosCancel - 1
'        If MatGastos(i, 4) = psTipoGastoProc(0) Or MatGastos(i, 4) = psTipoGastoProc(1) Or MatGastos(i, 4) = psTipoGastoProc(2) Then
'            MontoTotalGastosGenerado = MontoTotalGastosGenerado + CDbl(MatGastos(i, 3))
'        End If
'    Next i
'End Function

Private Sub cmdAplicar_Click()
'Dim nDias As Integer
'Dim nDiasAtraso As Integer
'Dim I As Integer
'Dim j As Integer
'Dim pdFechaIni As Date
'Dim pdFechaFin As Date
'Dim oGasto As COMNCredito.NCOMGasto
'Dim nNumgastos As Integer
'Dim MatGastos As Variant
Dim sCad As String
'Dim oDataGastos As COMDCredito.DCOMGasto
'Dim R, RFiltro, RFiltroGar As ADODB.Recordset
'Dim oNCred As COMNCredito.NCOMCredito
'Dim MatCalendTempo As Variant
'Dim dCal As COMDCredito.DCOMCalendario
'Dim odGastos1 As COMDCredito.DCOMGasto
'Dim cInstitucion As String

'    nNumGasCarg = 0

Dim nMontoPago As Double
Dim nInteres As Double
Dim nGastos As Double
Dim nMora As Double
Dim nIntAFecha As Double

'CUSCO
Dim nitf As Double

    sCad = ValidaFecha(TxtFechaPago.Text)
    If sCad <> "" Then
        MsgBox sCad, vbInformation, "Aviso"
        TxtFechaPago.SetFocus
        Exit Sub
    End If
    
    Screen.MousePointer = 11
    LblMontoPago.Caption = "0.00"
    LblMora.Caption = "0.00"
    LblInteres.Caption = "0.00"
    LblGastos.Caption = "0.00"
    'CUSCO
    lblITF.Caption = "0.00"
    '*****
    psCtaCod = ActxCta.NroCuenta
    nMontoGasto = 0
        
    Set oNegCredito = New COMNCredito.NCOMCredito
    Call oNegCredito.AplicarSimuladorPagos(psCtaCod, gdFecSis, CDate(TxtFechaPago.Text), CDbl(LblPrestamo.Caption), _
                            vnTasa, LblMetLiq.Caption, nTasaMorat, MatCalend, MatGastosCargados, nNumGasCarg, nMontoGasto, _
                            nMontoIntCompVencCierre, nMontoPago, nInteres, nGastos, nMora, nitf, nIntAFecha, lsTipoProducto)
    Set oNegCredito = Nothing
    
    LblInteres.Caption = nInteres
    LblGastos.Caption = nGastos
    LblMora.Caption = nMora
    LblMontoPago.Caption = nMontoPago
    'CUSCO
    lblITF.Caption = nitf
    '*****
    lblDeuda.Caption = Format(CDbl(LblSaldo.Caption) + nIntAFecha + nGastos + nMora, "#0.00")
    
    Screen.MousePointer = 0
'    pdFechaIni = gdFecSis
'    pdFechaFin = CDate(TxtFechaPago.Text)
'    nDias = DateDiff("d", pdFechaIni, pdFechaFin)
'    nNumGasCarg = 0
'    Set dCal = New COMDCredito.DCOMCalendario
'    Set R = dCal.RecuperaCalendarioGastosPendientesAFecha(ActxCta.NroCuenta, gColocCalendAplCuota, CDate(TxtFechaPago.Text))
'    Do While Not R.EOF
'        MatGastosCargados(nNumGasCarg, 0) = Trim(R!cDescripcion)
'        MatGastosCargados(nNumGasCarg, 1) = Format(R!nMonto, "#0.00")
'        nNumGasCarg = nNumGasCarg + 1
'        R.MoveNext
'    Loop
'    R.Close
'
'    Set oDataGastos = New COMDCredito.DCOMGasto
'    Set R = oDataGastos.RecuperaGastosAplicablesCuotas(CInt(Mid(ActxCta.NroCuenta, 9, 1)), "CD", , True)
'    If Mid(ActxCta.NroCuenta, 6, 3) <> "301" Then
'            Set RFiltro = oDataGastos.RecuperaFiltroAplicadoCuenta("P", True)
'            Set RFiltroGar = oDataGastos.RecuperaFiltroAplicadoCuenta("G", True)
'    Else
'            'verifico la institucion
'            Set odGastos1 = New COMDCredito.DCOMGasto
'            cInstitucion = odGastos1.ObtenerCodInstitucionByCredito(ActxCta.NroCuenta)
'            Set odGastos1 = Nothing
'            Set RFiltro = oDataGastos.RecuperaFiltroAplicadoCuenta("P", True, cInstitucion)
'            Set RFiltroGar = oDataGastos.RecuperaFiltroAplicadoCuenta("G", True, cInstitucion)
'    End If
'    'Set RFiltro = oDataGastos.RecuperaFiltroAplicadoCuenta("P")
'    'Set RFiltroGar = oDataGastos.RecuperaFiltroAplicadoCuenta("G")
'    Set oDataGastos = Nothing
'
'    nMontoIntCompVencCierre = 0
'    Set oNegCredito = New COMNCredito.NCOMCredito
'    Set oGasto = New COMNCredito.NCOMGasto
'
'    Set oNCred = New COMNCredito.NCOMCredito
'
'    MatCalendTempo = MatCalend
'    MatGastos = oGasto.GeneraCalendarioGastos(MatCalendTempo, Array(0, 0), nNumgastos, pdFechaIni, psCtaCod, 1, "PA", , , CDbl(LblPrestamo.Caption), , , True, R, RFiltro, RFiltroGar, nDiasAtraso)
'    For j = 0 To nNumgastos - 1
'            MatGastosCargados(nNumGasCarg, 0) = MatGastos(j, 2)
'            MatGastosCargados(nNumGasCarg, 1) = MatGastos(j, 3)
'            nNumGasCarg = nNumGasCarg + 1
'    Next j
'    nMontoGasto = nMontoGasto + MontoTotalGastosGenerado(MatGastos, nNumgastos, Array("PA", "", ""))
'    MatCalendTempo = MatCalend
'    For I = 1 To nDias
'        nDiasAtraso = (pdFechaIni - CDate(MatCalend(0, 0)))
'        nDiasAtraso = IIf(nDiasAtraso < 0, 0, nDiasAtraso)
'        MatGastos = oGasto.GeneraCalendarioGastos(MatCalendTempo, Array(0, 0), nNumgastos, pdFechaIni, psCtaCod, 1, "CD", , , CDbl(LblPrestamo.Caption), , , True, R, RFiltro, RFiltroGar, nDiasAtraso + 1)
'
'        For j = 0 To nNumgastos - 1
'            MatGastosCargados(nNumGasCarg, 0) = MatGastos(j, 2)
'            MatGastosCargados(nNumGasCarg, 1) = MatGastos(j, 3)
'            nNumGasCarg = nNumGasCarg + 1
'        Next j
'
'        nMontoGasto = nMontoGasto + MontoTotalGastosGenerado(MatGastos, nNumgastos, Array("CD", "", ""))
'        nMontoIntCompVencCierre = oNegCredito.MatrizIntCompVencidoCierre(MatCalendTempo, pdFechaIni, vnTasa, nMontoIntCompVencCierre)
'        pdFechaIni = pdFechaIni + 1
'    Next I
'    Set oGasto = Nothing
'
'    If LblMetLiq.Caption = "GMiC" Or LblMetLiq.Caption = "GMYC" Then
'        LblMontoPago.Caption = Format(oNegCredito.MatrizCapitalVencido(MatCalend, CDate(TxtFechaPago.Text)) + oNegCredito.MatrizInteresTotalesAFecha(ActxCta.NroCuenta, MatCalend, CDate(TxtFechaPago.Text)) + nMontoGasto + oNegCredito.MatrizGastosVencidos(MatCalend, CDate(Me.TxtFechaPago.Text)), "#0.00")
'    Else
'        'LblMontoPago.Caption = Format(oNegCredito.MatrizMontoAPagar(MatCalend, pdFechaFin) + nMontoIntCompVencCierre + nMontoGasto + oNegCredito.MatrizGastosVencidos(MatCalend, CDate(Me.TxtFechaPago.Text)), "#0.00")
'        LblMontoPago.Caption = Format(oNegCredito.MatrizMontoAPagar(MatCalend, pdFechaFin) + nMontoIntCompVencCierre + nMontoGasto, "#0.00")
'    End If
'
'    If LblMetLiq.Caption = "GMiC" Or LblMetLiq.Caption = "GMYC" Then
'        LblInteres.Caption = Format(oNegCredito.MatrizInteresTotalesAFecha(ActxCta.NroCuenta, MatCalend, CDate(TxtFechaPago.Text)), "#0.00")
'    Else
'        If pdFechaFin >= CDate(MatCalend(0, 0)) Then
'            LblInteres.Caption = Format(oNegCredito.MatrizIntCompVencido(MatCalend, pdFechaFin) + oNegCredito.MatrizIntCompVencidoCalendCierre(MatCalend, pdFechaFin), "#0.00")
'        Else
'            LblInteres.Caption = Format(CDbl(MatCalend(0, 4)) + CDbl(MatCalend(0, 5)), "#0.00")
'        End If
'    End If
'    LblInteres.Caption = Format(CDbl(LblInteres.Caption) + nMontoIntCompVencCierre, "#0.00")
'
'    LblGastos.Caption = Format(nMontoGasto + oNegCredito.MatrizGastosVencidos(MatCalend, CDate(Me.TxtFechaPago.Text)), "#0.00")
'    LblMora.Caption = Format(oNegCredito.MatrizInteresMorFecha(ActxCta.NroCuenta, MatCalend) + oNegCredito.MatrizCalculoMoraSimuladorPagos(MatCalend, gdFecSis, pdFechaFin, nTasaMorat), "#0.00")
'    LblMontoPago.Caption = Format(CDbl(LblMontoPago.Caption) + oNegCredito.MatrizCalculoMoraSimuladorPagos(MatCalend, gdFecSis, pdFechaFin, nTasaMorat), "#0.00")
'    LblMontoPago.Caption = Format(CDbl(LblMontoPago.Caption), "#0.00")
'
'    Set oNegCredito = Nothing
'    Set oNCred = Nothing
'    Screen.MousePointer = 0
End Sub

Private Sub cmdCancelar_Click()
    Call LimpiaPantalla
End Sub

Private Sub CmdGastos_Click()
    frmCredSimPagosGastos.Mostrar MatGastosCargados, nNumGasCarg
    
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 And ActxCta.Enabled = True Then 'F12
        Dim bRetSinTarjeta As Boolean
        Dim sCuenta As String
        sCuenta = frmValTarCodAnt.Inicia(0, bRetSinTarjeta)
        If sCuenta <> "" Then
            ActxCta.NroCuenta = sCuenta
            ActxCta.SetFocusCuenta
        End If
    End If
End Sub

Private Sub Form_Load()
    CentraForm Me
    ActxCta.CMAC = gsCodCMAC
    ActxCta.Age = gsCodAge
    nNumGasCarg = 0
End Sub

Private Sub HabilitaControles(ByVal pbHab As Boolean)
    ActxCta.Enabled = Not pbHab
    CmdAplicar.Enabled = pbHab
    TxtFechaPago.Enabled = pbHab
End Sub

Private Sub LimpiaPantalla()
    ActxCta.NroCuenta = ""
    ActxCta.CMAC = gsCodCMAC
    ActxCta.Age = gsCodAge
    lblTitular.Caption = ""
    LblAnalista.Caption = ""
    LblPrestamo.Caption = "0.00"
    LblSaldo.Caption = "0.00"
    TxtFechaPago.Text = "__/__/____"
    LblMontoPago.Caption = "0.00"
    LblInteres.Caption = "0.00"
    LblGastos.Caption = "0.00"
    LblFechaPend.Caption = ""
    LblCuotaPend.Caption = ""
    HabilitaControles False
    
End Sub

Private Sub CargaDatos(ByVal psCtaCod As String)
'Dim oCredito As COMDCredito.DCOMCredito
Dim R2 As ADODB.Recordset

'    Set oCredito = New COMDCredito.DCOMCredito
'    Set R2 = oCredito.RecuperaDatosComunes(psCtaCod, False)

    Set oNegCredito = New COMNCredito.NCOMCredito
    Call oNegCredito.CargarDatosSimuladorPagos(psCtaCod, R2, MatCalend, MatCalendParalelo)
    Set oNegCredito = Nothing
    
    If R2.RecordCount > 0 Then
        lblTitular.Caption = PstaNombre(R2!cTitular)
        LblAnalista.Caption = PstaNombre(R2!cAnalista)
        LblSaldo.Caption = Format(R2!nSaldo, "#0.00")
        LblPrestamo.Caption = Format(R2!nMontoCol, "#0.00")
        nTasaMorat = IIf(IsNull(R2!nTasaMora), 0, R2!nTasaMora)
        lsTipoProducto = R2!cTpoProdCod
        vnTasa = R2!nTasaInteres
        HabilitaControles True
        LblMetLiq.Caption = R2!cMetLiquidacion
        'Set oNegCredito = New COMNCredito.NCOMCredito
        'If CInt(Mid(psCtaCod, 6, 3)) = gColHipoMiVivienda Then
        '        MatCalend = oNegCredito.RecuperaMatrizCalendarioPendiente(psCtaCod)
        '        MatCalendParalelo = oNegCredito.RecuperaMatrizCalendarioPendiente(psCtaCod, True)
        '        MatCalend = UnirMatricesMiViviendaAmortizacion(MatCalend, MatCalendParalelo)
        'Else
        '        MatCalend = oNegCredito.RecuperaMatrizCalendarioPendiente(psCtaCod)
        'End If
        'Set oNegCredito = Nothing
        
        LblCuotaPend.Caption = MatCalend(0, 1)
        LblFechaPend.Caption = MatCalend(0, 0)
    Else
        MsgBox "Credito No Existe, o No Esta Vigente, o No es un Credito Hipotecario", vbInformation, "Aviso"
        LimpiaPantalla
    End If
    R2.Close
    
    'Set oCredito = Nothing
        
End Sub

