VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmOrdPagSolicitud 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9030
   Icon            =   "frmOrdPagSolicitud.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   9030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   105
      TabIndex        =   5
      Top             =   4605
      Width           =   990
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   7860
      TabIndex        =   4
      Top             =   4620
      Width           =   1035
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   6720
      TabIndex        =   3
      Top             =   4620
      Width           =   1035
   End
   Begin VB.Frame fraSolicitud 
      Caption         =   "Solicitud"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2115
      Left            =   5280
      TabIndex        =   10
      Top             =   2400
      Width           =   3615
      Begin VB.ComboBox cboNumOP 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   1830
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   990
         Width           =   1305
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "# Ordenes de Pago :"
         Height          =   195
         Left            =   270
         TabIndex        =   23
         Top             =   1050
         Width           =   1485
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Monto Descuento :"
         Height          =   195
         Left            =   330
         TabIndex        =   22
         Top             =   405
         Width           =   1365
      End
      Begin VB.Label lblDescuento 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   1860
         TabIndex        =   21
         Top             =   315
         Width           =   1035
      End
      Begin VB.Line Line1 
         X1              =   180
         X2              =   3420
         Y1              =   825
         Y2              =   825
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Del :"
         Height          =   195
         Left            =   240
         TabIndex        =   20
         Top             =   1590
         Width           =   330
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Al :"
         Height          =   195
         Left            =   1875
         TabIndex        =   19
         Top             =   1590
         Width           =   225
      End
      Begin VB.Label lblInicio 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   630
         TabIndex        =   18
         Top             =   1500
         Width           =   1155
      End
      Begin VB.Label lblFin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   2175
         TabIndex        =   17
         Top             =   1500
         Width           =   1155
      End
   End
   Begin VB.Frame fraHistoria 
      Caption         =   "Historia"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2115
      Left            =   60
      TabIndex        =   9
      Top             =   2400
      Width           =   5175
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdHistoria 
         Height          =   1755
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   3096
         _Version        =   393216
         TextStyleFixed  =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame fraCuenta 
      Caption         =   "Datos Cuenta"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2295
      Left            =   60
      TabIndex        =   6
      Top             =   60
      Width           =   8895
      Begin VB.Frame fraDetCuenta 
         Height          =   1515
         Left            =   5280
         TabIndex        =   7
         Top             =   660
         Width           =   3495
         Begin VB.Label lblFirmas 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   1140
            TabIndex        =   16
            Top             =   1080
            Width           =   2175
         End
         Begin VB.Label lblTipoCuenta 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   1140
            TabIndex        =   15
            Top             =   660
            Width           =   2175
         End
         Begin VB.Label lblApertura 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   1140
            TabIndex        =   14
            Top             =   240
            Width           =   2175
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "# Firmas :"
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   1170
            Width           =   690
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Cuenta :"
            Height          =   195
            Left            =   120
            TabIndex        =   12
            Top             =   750
            Width           =   960
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Apertura :"
            Height          =   195
            Left            =   120
            TabIndex        =   11
            Top             =   330
            Width           =   690
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdCliente 
         Height          =   1455
         Left            =   180
         TabIndex        =   1
         Top             =   720
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   2566
         _Version        =   393216
         TextStyleFixed  =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin SICMACT.ActXCodCta txtCuenta 
         Height          =   435
         Left            =   180
         TabIndex        =   0
         Top             =   240
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   767
      End
      Begin VB.Label lblDatosCuenta 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   435
         Left            =   3360
         TabIndex        =   8
         Top             =   240
         Width           =   5235
      End
   End
End
Attribute VB_Name = "frmOrdPagSolicitud"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nMontoDescuento As Double
Dim nSaldoMinimo As Double

Private Sub CargaNumTalonario(ByVal sMoneda As String)
    Dim rsOP As ADODB.Recordset
    VSQL = "Select nNumOP, nCosto FROM OrdPagTarifa Where cMoneda = '" & sMoneda & "' " _
        & "Order by nNumOP"
    Set rsOP = New ADODB.Recordset
    rsOP.CursorLocation = adUseClient
    rsOP.Open VSQL, dbCmactN, adOpenStatic, adLockReadOnly, adCmdText
    Set rsOP.ActiveConnection = Nothing
    Do While Not rsOP.EOF
        cboNumOP.AddItem rsOP("nNumOP") & Space(100) & rsOP("nCosto")
        rsOP.MoveNext
    Loop
    cboNumOP.ListIndex = 0
End Sub

Private Function GeneraDescuento(ByVal sCuenta As String) As Boolean
    Dim rsCta As ADODB.Recordset
    Dim nSaldo As Double, nIntGanado As Double, nTasa As Double
    Dim nSaldoContable As Double, nDescuentoTotal As Double
    Dim nDiasTranscurridos As Long
    Set rsCta = New ADODB.Recordset
    rsCta.CursorLocation = adUseClient
    AbreConexion
    VSQL = "SELECT nSaldDispAC,nSaldCntAC, nInteres, nTasaIntAC, dUltMovAC FROM AhorroC " _
        & "WHERE cCodCta = '" & sCuenta & "'"
    rsCta.Open VSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    Set rsCta.ActiveConnection = Nothing
    
    nTasa = rsCta("nTasaIntAC")
    nSaldo = rsCta("nSaldDispAC")
    nSaldoContable = rsCta("nSaldCntAC")
    nDiasTranscurridos = DateDiff("d", rsCta("dUltMovAC"), gdFecSis) - 1
    rsCta.Close
    Set rsCta = Nothing
    nDescuentoTotal = CDbl(lblDescuento)
    If nSaldo - nDescuentoTotal >= nSaldoMinimo Then
        nIntGanado = Round((nTasa / 36000) * nSaldo * nDiasTranscurridos, 2)
        dbCmact.BeginTrans
        VSQL = "UPDATE AhorroC " _
            & "SET nSaldAntAC = nSaldDispAC, " _
            & "nSaldDispAC = nSaldDispAC - " & nDescuentoTotal & ", " _
            & "nSaldCntAC = nSaldCntAC - " & nDescuentoTotal & ", " _
            & "nInteres = nInteres + " & nIntGanado & ", " _
            & "dUltActAC = '" & FechaHora(gdFecSis) & "', " _
            & "dUltMovAC = '" & FechaHora(DateAdd("d", -1, gdFecSis)) & "', cCodUsu = '" & gsCodUser & "' " _
            & "Where cCodCta = '" & sCuenta & "'"
        dbCmact.Execute VSQL
        VSQL = "INSERT TranDiaria (dFecTran, cCodUsu, cCodOpe, cCodCta, cNumDoc, nMonTran, nSaldCnt, cCodUsuRem, cCodAge, cFlag, nTipCambio)  VALUES('" & FechaHora(gdFecSis) & "','" & gsCodUser & "','" _
            & gsACImpOP & "','" & sCuenta & "', NULL," _
            & nDescuentoTotal & "," & nSaldoContable & ",NULL, '" & gsCodAge & "', NULL, 0)"
        dbCmact.Execute VSQL
        dbCmact.CommitTrans
        GeneraDescuento = True
    Else
        GeneraDescuento = False
    End If
    CierraConexion
End Function

Private Sub AgregaHistoria(ByVal rsHist As ADODB.Recordset)
Dim nFila As Long
Do While Not rsHist.EOF
    If grdHistoria.TextMatrix(1, 0) <> "" Then grdHistoria.Rows = grdHistoria.Rows + 1
    nFila = grdHistoria.Rows - 1
    grdHistoria.TextMatrix(nFila, 0) = nFila
    grdHistoria.TextMatrix(nFila, 1) = rsHist("nInicio")
    grdHistoria.TextMatrix(nFila, 2) = rsHist("nFin")
    grdHistoria.TextMatrix(nFila, 3) = Trim(rsHist("cEstado"))
    grdHistoria.TextMatrix(nFila, 4) = Format$(rsHist("dFecha"), "dd/mm/yyyy")
    rsHist.MoveNext
Loop
End Sub

Private Function GetMaxOrdPagEmitida(ByVal sCuenta As String) As Long
Dim rsNum As ADODB.Recordset
Set rsNum = New ADODB.Recordset
rsNum.CursorLocation = adUseClient
VSQL = "Select ISNULL(MAX(nFin),0) nNum From OrdPagEmision Where cCodCta LIKE '" & Mid(sCuenta, 1, 6) & "%' " _
    & "And cEstado NOT IN ('5')"
rsNum.Open VSQL, dbCmactN, adOpenStatic, adLockReadOnly, adCmdText
Set rsNum.ActiveConnection = Nothing
If Not (rsNum.EOF And rsNum.BOF) Then
    GetMaxOrdPagEmitida = rsNum("nNum")
Else
    GetMaxOrdPagEmitida = 0
End If
rsNum.Close
Set rsNum = Nothing
End Function

Private Sub SetupGrid()
grdCliente.Clear
grdCliente.Rows = 2
grdCliente.Cols = 3
grdCliente.ColAlignmentFixed(0) = 4
grdCliente.ColAlignmentFixed(1) = 4
grdCliente.ColAlignmentFixed(2) = 4
grdCliente.ColAlignment(0) = 4
grdCliente.ColAlignment(1) = 1
grdCliente.ColAlignment(2) = 4
grdCliente.TextMatrix(0, 0) = "#"
grdCliente.TextMatrix(0, 1) = "Nombre"
grdCliente.TextMatrix(0, 2) = "RE"
grdCliente.ColWidth(0) = 350
grdCliente.ColWidth(1) = 3500
grdCliente.ColWidth(2) = 500

grdHistoria.Clear
grdHistoria.Rows = 2
grdHistoria.Cols = 5
grdHistoria.ColAlignmentFixed(0) = 4
grdHistoria.ColAlignmentFixed(1) = 4
grdHistoria.ColAlignmentFixed(2) = 4
grdHistoria.ColAlignmentFixed(3) = 4
grdHistoria.ColAlignmentFixed(4) = 4
grdHistoria.ColAlignment(0) = 4
grdHistoria.ColAlignment(1) = 4
grdHistoria.ColAlignment(2) = 4
grdHistoria.ColAlignment(3) = 4
grdHistoria.ColAlignment(4) = 4
grdHistoria.TextMatrix(0, 0) = "#"
grdHistoria.TextMatrix(0, 1) = "Inicio"
grdHistoria.TextMatrix(0, 2) = "Fin"
grdHistoria.TextMatrix(0, 3) = "Estado"
grdHistoria.TextMatrix(0, 4) = "Fecha"
grdHistoria.ColWidth(0) = 350
grdHistoria.ColWidth(1) = 1000
grdHistoria.ColWidth(2) = 1000
grdHistoria.ColWidth(3) = 1200
grdHistoria.ColWidth(4) = 1200
End Sub

Private Sub ObtieneDatosCuenta(ByVal sCuenta As String)
Dim rsCta As ADODB.Recordset

Dim nFila As Long
VSQL = "Select P.cNomPers, TC.cNomTab cTipoCuenta, PC.cRelaCta, A.nNumFirmAC, A.cCodCta, A.dAperAC " _
    & "From " & gcCentralPers & "Persona P INNER JOIN PersCuenta PC INNER JOIN AhorroC A " _
    & "ON PC.cCodCta = A.cCodCta ON P.cCodPers = PC.cCodPers INNER JOIN " & gcCentralCom _
    & "TablaCod TC ON A.cTipCtaAC = RTRIM(TC.cValor) Where A.cEstCtaAC NOT IN ('C','U') " _
    & "And A.cOrdPag = 'S' And A.cCodCta = '" & sCuenta & "' And TC.cCodTab LIKE '16__'"

AbreConexion
Set rsCta = New ADODB.Recordset
rsCta.CursorLocation = adUseClient
rsCta.Open VSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsCta.ActiveConnection = Nothing

If Not (rsCta.EOF And rsCta.BOF) Then
    lblDatosCuenta = "CUENTA CON ORDEN DE PAGO" & Chr$(13)
    If Mid(sCuenta, 6, 1) = "1" Then
        lblDescuento.BackColor = &H80000005
        lblDatosCuenta = lblDatosCuenta & "MONEDA NACIONAL"
        nSaldoMinimo = GeCapParametro("23102")
    Else
        lblDescuento.BackColor = &HC0FFC0
        lblDatosCuenta = lblDatosCuenta & "MONEDA EXTRANJERA"
        nSaldoMinimo = GeCapParametro("23103")
    End If
    
    lblApertura = Format$(rsCta("dAperAC"), "dd-mmm-yyyy")
    lblTipoCuenta = Trim(rsCta("cTipoCuenta"))
    lblFirmas.Caption = rsCta("nNumFirmAC")
    Do While Not rsCta.EOF
        If grdCliente.TextMatrix(1, 0) <> "" Then grdCliente.Rows = grdCliente.Rows + 1
        nFila = grdCliente.Rows - 1
        grdCliente.TextMatrix(nFila, 0) = nFila
        grdCliente.TextMatrix(nFila, 1) = PstaNombre(Trim(rsCta("cNomPers")), True)
        grdCliente.TextMatrix(nFila, 2) = UCase(Trim(rsCta("cRelaCta")))
        rsCta.MoveNext
    Loop
    Dim rsHist As ADODB.Recordset
    Set rsHist = New ADODB.Recordset
    rsHist.CursorLocation = adUseClient
    VSQL = "Select nNumIni nInicio, nNumFin nFin, dRegOp dFecha, cEstado = 'ENTREGADO' From " _
        & "OPEmiti Where cCodCta = '" & sCuenta & "' Order by dFecha"
    
    rsHist.Open VSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    Set rsHist.ActiveConnection = Nothing
    If Not (rsHist.EOF And rsHist.BOF) Then
        AgregaHistoria rsHist
    End If
    rsHist.Close
    If AbreConeccion(gsAgenciaCentralOP & "2321000019", False) Then
        Dim nMaxNumOP As Long
        nMaxNumOP = GetMaxOrdPagEmitida(sCuenta)
        VSQL = "Select nInicio, nFin, dFecha, cEstado = RTRIM(TC.cNomTab) From " _
            & "OrdPagEmision O INNER JOIN " & gcCentralCom & "TablaCod TC ON O.cEstado = TC.cValor " _
            & "Where O.cCodCta = '" & sCuenta & "' And TC.cCodTab LIKE 'P5%' Order by dFecha"
        
        rsHist.Open VSQL, dbCmactN, adOpenStatic, adLockReadOnly, adCmdText
        Set rsHist.ActiveConnection = Nothing
        If Not (rsHist.EOF And rsHist.BOF) Then
            AgregaHistoria rsHist
        End If
        If nMaxNumOP = 0 Then
            lblInicio = Mid(sCuenta, 6, 1) & Format$(1, "0000000")
        Else
            lblInicio = Trim(nMaxNumOP + 1)
        End If
        CargaNumTalonario Mid(sCuenta, 6, 1)
        fraCuenta.Enabled = False
        fraHistoria.Enabled = True
        fraSolicitud.Enabled = True
        cmdGrabar.Enabled = True
        cmdCancelar.Enabled = True
        cboNumOP.SetFocus
    Else
        MsgBox "No es posible conectar con la Agencia Central. Avise al Area de Sistemas", vbExclamation, "Error"
        cmdCancelar_Click
    End If
    CierraConeccion
Else
    MsgBox "Cuenta Cancelada, Anulada, o Sin Orden de Pago.", vbInformation, "Aviso"
    cmdCancelar_Click
End If
CierraConexion
End Sub

Private Sub cboNumOP_Click()
Dim nFin As Long
Dim nNumTal As Integer
nMontoDescuento = CDbl(Trim(Right(cboNumOP, 10)))
lblDescuento = Format$(nMontoDescuento, "#,##0.00")
nNumTal = CInt(Trim(Left(cboNumOP, 10)))
nFin = CLng(lblInicio) - 1
nFin = nFin + nNumTal
lblFin = Trim(nFin)
End Sub

Private Sub cmdCancelar_Click()
cmdGrabar.Enabled = False
fraCuenta.Enabled = True
fraHistoria.Enabled = False
fraSolicitud.Enabled = False
txtCuenta.psAge = Right(gsCodAge, 2)
txtCuenta.psProd = gsCodProAC
txtCuenta.pbEnabledAge = False
txtCuenta.pbEnabledProd = False
txtCuenta.psCuenta = ""
txtCuenta.SetFocusCuenta
cmdCancelar.Enabled = False
lblInicio = ""
lblFin = ""
lblDescuento.BackColor = &H80000005
lblDatosCuenta = ""
lblApertura = ""
lblTipoCuenta = ""
lblFirmas = ""
SetupGrid
nMontoDescuento = 0
lblDescuento = "0.00"
cboNumOP.Clear
End Sub

Private Sub cmdGrabar_Click()
If lblInicio = "" Or lblFin = "" Then
    MsgBox "Error en la generación de los números a emitir", vbExclamation, "Error"
    Unload Me
End If
If Len(lblInicio) <> 8 Or Len(lblFin) <> 8 Then
    MsgBox "Error en la generación de los números a emitir", vbExclamation, "Error"
    Unload Me
End If
If MsgBox("Desea grabar la solicitud de emisión de la Orden de Pago", vbQuestion + vbYesNo, "Aviso") = vbYes Then
    Dim sCuenta As String
    Dim bTrans As Boolean
    sCuenta = txtCuenta.NroCuenta
    If GeneraDescuento(sCuenta) Then
        AbreConexion
        If AbreConeccion(gsAgenciaCentralOP & "2321000019", False) Then
            Dim sInicio As String, sFin As String, sFecha As String
            dbCmactN.BeginTrans
            bTrans = True
            On Error GoTo ErrGraba
            sInicio = Trim(lblInicio)
            sFin = Trim(lblFin)
            sFecha = FechaHora(gdFecSis)
            VSQL = "Insert OrdPagEmision (cCodCta,nInicio,nFin,dFecha,cEstado,cCodUsu) " _
                & "Values ('" & sCuenta & "'," & sInicio & "," & sFin & ",'" & sFecha & " ','1','" & gsCodUser & "')"
            dbCmactN.Execute VSQL
            VSQL = "Insert OrdPagEstado (cCodCta,nInicio,dFecha,cEstado,cCodUsu) " _
                & "Values ('" & sCuenta & "'," & sInicio & ",'" & sFecha & " ','1','" & gsCodUser & "')"
            dbCmactN.Execute VSQL
            dbCmactN.CommitTrans
            bTrans = False
            CierraConeccion
        Else
            MsgBox "No es posible conectar con la Agencia Central. Avise al Area de Sistemas", vbExclamation, "Error"
            cmdCancelar_Click
        End If
        CierraConexion
    Else
        MsgBox "La cuenta no posee el saldo suficiente para realizar el cargo correspondiente.", vbInformation, "Aviso"
    End If
    cmdCancelar_Click
End If
Exit Sub
ErrGraba:
    If bTrans Then dbCmactN.RollbackTrans
    CierraConeccion
    MsgBox Err.Description, vbExclamation, "Error"
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Caption = "Orden Pago - Solicitud Emisión"
cmdGrabar.Enabled = False
cmdCancelar.Enabled = False
fraHistoria.Enabled = False
fraSolicitud.Enabled = False
txtCuenta.psAge = Right(gsCodAge, 2)
txtCuenta.psProd = gsCodProAC
txtCuenta.pbEnabledAge = False
txtCuenta.pbEnabledProd = False
lblInicio = ""
lblFin = ""
nMontoDescuento = 0
lblDescuento = "0.00"
lblDatosCuenta = ""
lblApertura = ""
lblTipoCuenta = ""
lblFirmas = ""
SetupGrid
cboNumOP.Clear
End Sub

Private Sub txtCuenta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Dim sCuenta As String
    sCuenta = txtCuenta.NroCuenta
    ObtieneDatosCuenta sCuenta
End If
End Sub

