VERSION 5.00
Begin VB.Form frmCredReprogCredConvenio 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ampliación de Plazos para Créditos Convenio"
   ClientHeight    =   7470
   ClientLeft      =   8835
   ClientTop       =   7680
   ClientWidth     =   10935
   Icon            =   "frmCredReprogCredConvenio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   10935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   360
      Left            =   8520
      TabIndex        =   32
      Top             =   6720
      Width           =   1050
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      Height          =   360
      Left            =   9720
      TabIndex        =   31
      Top             =   6720
      Width           =   1050
   End
   Begin VB.Frame Frame4 
      Caption         =   " Glosa "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   735
      Left            =   120
      TabIndex        =   28
      Top             =   6600
      Width           =   7095
      Begin VB.TextBox txtGlosa 
         Height          =   315
         Left            =   120
         TabIndex        =   30
         Top             =   285
         Width           =   5655
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "Grabar"
         Height          =   360
         Left            =   5880
         TabIndex        =   29
         Top             =   240
         Width           =   1050
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   " Calendario "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2835
      Left            =   120
      TabIndex        =   26
      Top             =   3720
      Width           =   10695
      Begin SICMACT.FlexEdit FECalend 
         Height          =   2460
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   10485
         _ExtentX        =   18494
         _ExtentY        =   4339
         Cols0           =   13
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "-Fecha-Nro-Monto-Capital-Int. Comp-Int. Mor-Int. Reprog-Int Gracia-Gasto-Saldo-Estado-nCapPag"
         EncabezadosAnchos=   "400-1000-400-1000-1000-1000-1000-1000-1000-1000-1200-0-0"
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
         ColumnasAEditar =   "X-1-X-3-X-X-X-X-X-X-X-X-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-2-0-0-0-0-0-0-0-0-0-0-0"
         BackColorControl=   65535
         BackColorControl=   65535
         BackColorControl=   65535
         EncabezadosAlineacion=   "C-C-C-C-C-C-C-C-C-C-C-C-C"
         FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0-0-0-0"
         lbUltimaInstancia=   -1  'True
         lbPuntero       =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
         ForeColorFixed  =   -2147483635
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " Ampliación de Crédito "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   855
      Left            =   120
      TabIndex        =   21
      Top             =   2760
      Width           =   10695
      Begin VB.CheckBox chkSoloInt 
         Caption         =   "Excluir Capital en la 1era Cuota"
         Height          =   195
         Left            =   5520
         TabIndex        =   35
         Top             =   420
         Width           =   2655
      End
      Begin VB.TextBox txtDias 
         Enabled         =   0   'False
         Height          =   315
         Left            =   4200
         TabIndex        =   34
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox txtNuevoPlazo 
         Height          =   315
         Left            =   1440
         TabIndex        =   33
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton cmdCalcular 
         Caption         =   "Calcular"
         Height          =   360
         Left            =   8280
         TabIndex        =   24
         Top             =   330
         Width           =   1050
      End
      Begin VB.CheckBox chkReprogramar 
         Caption         =   "Reprogramar dias :"
         Height          =   255
         Left            =   2520
         TabIndex        =   23
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "días"
         Height          =   195
         Left            =   5040
         TabIndex        =   25
         Top             =   405
         Width           =   315
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Nuevas Cuotas :"
         Height          =   195
         Left            =   240
         TabIndex        =   22
         Top             =   405
         Width           =   1185
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Datos de Crédito "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2055
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   10695
      Begin VB.Label lblTasaInteres 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   9360
         TabIndex        =   20
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Tasa Interes:"
         Height          =   195
         Left            =   8280
         TabIndex        =   19
         Top             =   1485
         Width           =   930
      End
      Begin VB.Label lblCuotas 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   6480
         TabIndex        =   18
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Cuotas: "
         Height          =   195
         Left            =   5880
         TabIndex        =   17
         Top             =   1485
         Width           =   585
      End
      Begin VB.Label lblSaldo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   3720
         TabIndex        =   16
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Saldo:"
         Height          =   195
         Left            =   3120
         TabIndex        =   15
         Top             =   1485
         Width           =   450
      End
      Begin VB.Label lblPrestamo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   960
         TabIndex        =   14
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Préstamo: "
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   1485
         Width           =   750
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   10560
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Label lblProducto 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   8160
         TabIndex        =   12
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Producto :"
         Height          =   195
         Left            =   7320
         TabIndex        =   11
         Top             =   765
         Width           =   735
      End
      Begin VB.Label lblAnalista 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   8160
         TabIndex        =   10
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Analista :"
         Height          =   195
         Left            =   7320
         TabIndex        =   9
         Top             =   405
         Width           =   645
      End
      Begin VB.Label lblFechaUltCuota 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   6000
         TabIndex        =   8
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Ult Cuota :"
         Height          =   195
         Left            =   4680
         TabIndex        =   7
         Top             =   765
         Width           =   1245
      End
      Begin VB.Label lblTipoCredito 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1200
         TabIndex        =   6
         Top             =   720
         Width           =   3375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Crédito :"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   765
         Width           =   945
      End
      Begin VB.Label lblTitular 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1200
         TabIndex        =   4
         Top             =   360
         Width           =   5895
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Titular :"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   405
         Width           =   525
      End
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "..."
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   120
      Width           =   375
   End
   Begin SICMACT.ActXCodCta ActxCta 
      Height          =   435
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3660
      _ExtentX        =   6456
      _ExtentY        =   767
      Texto           =   "Crédito:"
      EnabledCMAC     =   -1  'True
      EnabledCta      =   -1  'True
      EnabledProd     =   -1  'True
      EnabledAge      =   -1  'True
   End
End
Attribute VB_Name = "frmCredReprogCredConvenio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'** Nombre : frmCredReprogCredConvenio
'** Descripción : Formulario para reprogramar créditos convenio según TI-ERS049-2014
'** Creación : JUEZ-WIOR, 20140426 10:00:00 AM
'*****************************************************************************************************

Option Explicit

Private oDCred As COMDCredito.DCOMCredito
Private oDCalend As COMDCredito.DCOMCalendario
Private oNGasto As COMNCredito.NCOMGasto
Private oNCred As COMNCredito.NCOMCredito
 
Private MatGracia As Variant
Private MatCalend As Variant
Private MatCalend_2 As Variant
Private MatDesemb As Variant
Private MatGastos As Variant
Private nNumGastos As Integer
Private MatCalendIC As Variant
Private R As ADODB.Recordset

Private fnPlazo As Integer
Private fnTipoCuota As Integer
Private fnTipoGracia As Integer
Private fsFecVenc As String
Private fnPeriodoGracia As Integer
Private fnFechaFija As Integer
Private fnFechaFija2 As Integer
Private fbProxMes As Boolean
Private fnTasaGracia As Double
Private fsTpoCred As String
Private fsTpoProdCred As String
Private fbQuincenal As Boolean
Private fnNroCalen As Integer

Private fnTCuota As Integer
Private fnTPeriodo As Integer
Private fnDiasReprog As Integer
Private fnMontoApr As Double
Private fnSaldoGracia As Double
Private fdFechaPago As Date
Private fdUltCuotaPag As Date
Private fnValosDiasReprog As Long
'WIOR 20150523 ***
Private fdFechaCredOtor As Date
Private fdFechaCalif As Date
Private fsTipoProd As String
Private fsCodCalif As String
'WIOR FIN ********
Private fnIntPag As Double 'WIOR 20150527

Private Sub ActxCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If ValidaCredito(ActxCta.NroCuenta) Then
            CargarDatos ActxCta.NroCuenta
        Else
            LimpiarCredito
        End If
    End If
End Sub

Private Sub chkReprogramar_Click()
    txtDias.Text = ""
    If chkReprogramar.value Then
        txtDias.Enabled = True
    Else
        txtDias.Enabled = False
    End If
    LimpiaFlex FECalend
End Sub

Private Sub chkSoloInt_Click()
LimpiaFlex FECalend
End Sub

Private Sub cmdBuscar_Click()
Dim oPers As COMDPersona.UCOMPersona
    Limpiar
    Set oPers = frmBuscaPersona.Inicio()
    If Not oPers Is Nothing Then
        Call FrmVerCredito.Inicio(oPers.sPersCod, , , True, ActxCta)
        ActxCta.SetFocusCuenta
    End If
    Set oPers = Nothing
End Sub

Private Sub cmdCalcular_Click()
    If Trim(txtNuevoPlazo.Text) = "" Or Trim(txtNuevoPlazo.Text) = "0" Then
        MsgBox "Debe ingresar la nueva cantidad de cuotas para el calendario", vbInformation, "Aviso"
        txtNuevoPlazo.SetFocus
        Exit Sub
    End If
    If chkReprogramar.value = 1 And (Trim(txtDias.Text) = "" Or Trim(txtDias.Text) = "0") Then
        MsgBox "Debe ingresar los dias de reprogramación", vbInformation, "Aviso"
        txtNuevoPlazo.SetFocus
        Exit Sub
    End If
    
Dim bErrorCargaCalendario As Boolean
Dim rsCalend As ADODB.Recordset
Dim i As Integer
'WIOR 20150527 ***
Dim oCredito As COMNCredito.NCOMCredito
'WIOR FIN ********

    txtDias.Text = IIf(Trim(txtDias.Text) = "", 0, txtDias.Text)
    fnDiasReprog = CInt(txtDias.Text)
    fnMontoApr = CDbl(lblSaldo.Caption)
    
    Call LimpiaFlex(FECalend)
    MatCalend = Array(0)
    
    Set oDCalend = New COMDCredito.DCOMCalendario
    Set rsCalend = oDCalend.RecuperaCalendarioPagos(ActxCta.NroCuenta)
    Set oDCalend = Nothing

    ReDim MatGracia(rsCalend.RecordCount)
    fnSaldoGracia = 0
    
    Do While Not rsCalend.EOF
        MatGracia(rsCalend.Bookmark - 1) = Format(rsCalend!nIntGracia, "#0.00")
        fnSaldoGracia = fnSaldoGracia + CDbl(IIf(IsNull(rsCalend!nIntGracia), 0, rsCalend!nIntGracia)) - CDbl(IIf(IsNull(rsCalend!nIntGraciaPag), 0, rsCalend!nIntGraciaPag))
        rsCalend.MoveNext
        
    Loop
    
    
    Select Case fnTipoCuota
        Case Trim(str(gColocCalendCodPFCF)), Trim(str(gColocCalendCodPFCFPG))  'Periodo Fijo Cuota Fija
            fnTCuota = 1
            fnTPeriodo = 1
        Case Trim(str(gColocCalendCodPFCC)), Trim(str(gColocCalendCodPFCCPG))   'Periodo Fijo - Cuota Creciente
            fnTCuota = 2
            fnTPeriodo = 1
        Case Trim(str(gColocCalendCodPFCD)), Trim(str(gColocCalendCodPFCDPG))  'Periodo Fijo - Cuota Decreciente"
            fnTCuota = 3
            fnTPeriodo = 1
        Case Trim(str(gColocCalendCodFFCF)), Trim(str(gColocCalendCodFFCFPG))   'Fecha Fija - Cuota Fija
            fnTCuota = 1
            fnTPeriodo = 2
        Case Trim(str(gColocCalendCodFFCC)), Trim(str(gColocCalendCodFFCCPG))   'Fecha Fija - Cuota Creciente
            fnTCuota = 2
            fnTPeriodo = 2
        Case Trim(str(gColocCalendCodFFCD)), Trim(str(gColocCalendCodFFCDPG))       'Fecha Fija - Cuota Decreciente
            fnTCuota = 3
            fnTPeriodo = 2
        Case Trim(str(gColocCalendCodCL))
            fnTCuota = 5
            fnTPeriodo = 2
    End Select
    
ReDim MatDesemb(1, 2)
MatDesemb(0, 0) = fdUltCuotaPag
MatDesemb(0, 1) = Format(fnMontoApr, "#0.00")

fdFechaPago = DateAdd("d", fnDiasReprog, CDate(fsFecVenc))
fnPeriodoGracia = 0
fnPeriodoGracia = DateDiff("d", CDate(MatDesemb(0, 0)), fdFechaPago) - fnPlazo

If fnPeriodoGracia < 0 Or fnDiasReprog = 0 Then
    fnPeriodoGracia = 0
End If

fbProxMes = IIf(DateDiff("m", CDate(MatDesemb(0, 0)), fdFechaPago) > 0, True, False)

fnFechaFija = IIf(fnFechaFija > 0, Day(fdFechaPago), 0)

fnTasaGracia = IIf(fnSaldoGracia > 0, 6, -1)

MatGastos = GenerarMatrices(True, MatCalendIC, bErrorCargaCalendario)

If Not bErrorCargaCalendario Then
    MsgBox "Ha ocurrido un error en mostrar el calendario", vbInformation, "Aviso"
    Limpiar
    Exit Sub
End If

MatCalend = MatCalendIC

Call MostrarGasto(MatCalend)

'WIOR 20150527 ***
Set oCredito = New COMNCredito.NCOMCredito

If CDbl(MatCalend(0, 1)) = 0 Then
    MatCalend(1, 4) = oCredito.MontoIntPerDias(CDbl(lblTasaInteres.Caption), DateDiff("D", fdUltCuotaPag, MatCalend(1, 0)), fnMontoApr) - fnIntPag
Else
    MatCalend(0, 4) = oCredito.MontoIntPerDias(CDbl(lblTasaInteres.Caption), DateDiff("D", fdUltCuotaPag, MatCalend(0, 0)), fnMontoApr) - fnIntPag
End If
Set oCredito = Nothing
'WIOR ************

For i = 0 To UBound(MatCalend) - 1
    FECalend.AdicionaFila
    FECalend.TextMatrix(i + 1, 1) = Format(CDate(MatCalend(i, 0)), "dd/mm/yyyy")
    FECalend.TextMatrix(i + 1, 2) = Trim(str(MatCalend(i, 1)))
    
    FECalend.TextMatrix(i + 1, 4) = Format(MatCalend(i, 3), "#0.00") 'Capital
    FECalend.TextMatrix(i + 1, 5) = Format(MatCalend(i, 4), "#0.00") 'IntComp
    FECalend.TextMatrix(i + 1, 6) = "0.00" 'IntMora
    FECalend.TextMatrix(i + 1, 7) = "0.00" 'IntReprog
    FECalend.TextMatrix(i + 1, 8) = Format(MatCalend(i, 5), "#0.00") 'IntGracia
    'Gasto
    If Trim(str(MatCalend(i, 1))) <> "0" Then
        If IsNumeric(MatCalend(i, 8)) Then
            FECalend.TextMatrix(i + 1, 9) = Format(CDbl(MatCalend(i, 6) + CDbl(MatCalend(i, 8))), "#0.00")
        Else
            FECalend.TextMatrix(i + 1, 9) = Format(CDbl(MatCalend(i, 6)), "#0.00")
        End If
    Else
        FECalend.TextMatrix(i + 1, 9) = "0.00"
    End If

    
    If CInt(Trim(str(MatCalend(i, 1)))) > 0 Then
        'Monto
        FECalend.TextMatrix(i + 1, 3) = Format(CDbl(FECalend.TextMatrix(i + 1, 4)) + CDbl(FECalend.TextMatrix(i + 1, 5)) + CDbl(FECalend.TextMatrix(i + 1, 6)) + CDbl(FECalend.TextMatrix(i + 1, 7)) + CDbl(FECalend.TextMatrix(i + 1, 8)) + CDbl(FECalend.TextMatrix(i + 1, 9)), "#0.00")
        
        fnMontoApr = fnMontoApr - MatCalend(i, 3)
    Else
        'Monto
        FECalend.TextMatrix(i + 1, 3) = Format(MatCalend(i, 2), "#0.00")
    End If
    fnMontoApr = CDbl(Format(fnMontoApr, "#0.00"))
    FECalend.TextMatrix(i + 1, 10) = Format(fnMontoApr, "#0.00")
    Call FECalend.ForeColorRow(vbBlack)
Next i
FECalend.TopRow = 1 'WIOR 20150527
fnMontoApr = CDbl(lblSaldo.Caption)
End Sub

Private Sub MostrarGasto(ByRef pMatCalend As Variant)
Dim nTotalGasto As Double
Dim nTotalGastoSeg As Double
Dim i, j As Integer

If IsArray(MatGastos) Then
    For j = 0 To UBound(pMatCalend) - 1
        nTotalGasto = 0
        nTotalGastoSeg = 0
        For i = 0 To UBound(MatGastos) - 1
            If Trim(Right(MatGastos(i, 0), 2)) = "1" And _
               (Trim(MatGastos(i, 1)) = Trim(pMatCalend(j, 1)) And Trim(Right(MatGastos(i, 2), 4)) = "1217") Then
                nTotalGastoSeg = nTotalGastoSeg + CDbl(MatGastos(i, 3))
                pMatCalend(j, 6) = Format(nTotalGastoSeg, "#0.00")
            Else
                If Trim(MatGastos(i, 1)) = "*" Or (Trim(Right(MatGastos(i, 0), 2)) = "1" And Trim(MatGastos(i, 1)) = Trim(pMatCalend(j, 1)) And Trim(Right(MatGastos(i, 2), 4)) <> "1217") Then
                    nTotalGasto = nTotalGasto + CDbl(MatGastos(i, 3))
                    pMatCalend(j, 8) = Format(nTotalGasto, "#0.00")
                End If
            End If
        
        Next i
    Next j
End If
End Sub

Private Sub cmdCancelar_Click()
    Limpiar
    ActxCta.SetFocusCuenta
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub cmdGrabar_Click()
Dim bError As Boolean
Dim objPista As COMManejador.Pista
On Error GoTo ErrorReprogCredConvenio

bError = False
    If ValidaDatos Then
        If MsgBox("Estas seguro de guardar los datos?", vbInformation + vbYesNo, "Aviso") = vbYes Then
        
            If fnTasaGracia = 6 Then
                Call RegularCalendario
            End If
            
            MatGastos = GenerarMatrices(False, MatCalendIC)
            
            Set oNCred = New COMNCredito.NCOMCredito
            Call oNCred.ReprogramarCredConvenio(Trim(ActxCta.NroCuenta), CDate(gdFecSis), fnNroCalen + 1, MatDesemb, MatCalend, MatGastos, nNumGastos, fnTipoCuota, bError)
            
            If Not bError Then
                Set objPista = New COMManejador.Pista
                'objPista.InsertarPista "100922", GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, gInsertar, Trim(txtGlosa.Text), Trim(ActxCta.NroCuenta), gCodigoCuenta
                objPista.InsertarPista "100922", GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, gInsertar, Trim(txtGlosa.Text) & " - NroCuotas: " & txtNuevoPlazo.Text & IIf(Me.chkReprogramar.value = 1, ", Dias Reporgramados:" & txtDias.Text, "") & IIf(chkSoloInt.value = 1, ", Capital excluido en la 1era Cuota", ""), Trim(ActxCta.NroCuenta), gCodigoCuenta
                'WIOR 20150527 IIf(chkSoloInt.value = 1, ", Capital excluido en la 1era Cuota" , "")
                Set objPista = Nothing
                MsgBox "Datos guardados Satisfactoriamente", vbInformation, "Aviso"
                Limpiar
            End If
            
        End If
    End If
    
Exit Sub
ErrorReprogCredConvenio:
    MsgBox err.Description, vbInformation, "Aviso"
End Sub

Private Sub Form_Load()
Call CargaVariables 'WIOR 20150523
Call Limpiar
End Sub

Private Sub LimpiarCredito()
    ActxCta.CMAC = "109"
    ActxCta.Age = gsCodAge
    ActxCta.Prod = gColProConsumoPerDesPla
    ActxCta.Cuenta = ""
    ActxCta.EnabledCMAC = False
End Sub

Private Function ValidaCredito(ByVal psCtaCod As String) As Boolean
    
    ValidaCredito = False
    
    Set oDCred = New COMDCredito.DCOMCredito
    Set R = oDCred.VerificaAmpliacionPlazoCreditoConvenio(psCtaCod)
    Set oDCred = Nothing

    If Not R.BOF And Not R.EOF Then
        If InStr(1, fsTipoProd, Trim(R!cTpoProdCod)) = 0 Then 'If R!cTpoProdCod <> gColProConsumoPerDesPla And R!cTpoProdCod <> "512" Then'WIOR 20150523
            MsgBox "El crédito debe ser necesariamente un crédito Convenio con Descuento por Planilla o Negocio por Convenio", vbInformation, "Aviso"
            Exit Function
        End If
        
        If CDate(R!dVigencia) > fdFechaCredOtor Then  'If Not CDate(R!dVigencia) < CDate("16/01/2014") Then'WIOR 20150523
            'MsgBox "El crédito debe tener la fecha de desembolso anterior al 16 de Enero del 2014", vbInformation, "Aviso"'WIOR 20150523
            MsgBox "La fecha de desembolso del crédito debe ser hasta el " & Format(fdFechaCredOtor, "dd/mm/yyyy") & ".", vbInformation, "Aviso" 'WIOR 20150523
            Exit Function
        End If
        
        If InStr(1, fsCodCalif, Trim(R!cCalGen)) = 0 Or Trim(R!cCalGen) = "" Then 'If R!cCalGen <> "0" Then'WIOR 20150523
            If (Trim(R!cCalGen) <> "") Or (Trim(R!cCalGen) = "" And CDate(R!dVigencia) <= fdFechaCalif) Then 'WIOR 20150523
                'MsgBox "El crédito debe tener calificación 100% normal al 31 de Diciembre del 2013", vbInformation, "Aviso"'WIOR 20150523
                MsgBox "El crédito debe tener Calificación 100% Normal(Sin Alineamiento) al " & Format(fdFechaCalif, "dd/mm/yyyy") & ".", vbInformation, "Aviso" 'WIOR 20150523
                Exit Function
            End If 'WIOR 20150523
        End If
        
        'WIOR 20150523 ***
        If CDbl(R!nMora) > 0 Then
            MsgBox "Debe cancelar los interes moratorios(" & IIf(Mid(psCtaCod, 9, 1) = "1", "S/. ", "$ ") & Format(R!nMora, "#0.00") & ") antes de realizar la ampliación de de plazo del crédito.", vbInformation, "Aviso"
            Exit Function
        End If
        'WIOR FIN ********
        
        If R!cAmpliado <> "" Then
            MsgBox "El crédito ya fue ampliado anteriormente por esta opción.", vbInformation, "Aviso"
            Exit Function
        End If
        ValidaCredito = True
    Else
        MsgBox "No se pudo encontrar crédito", vbInformation, "Aviso"
    End If
End Function

Private Sub CargarDatos(ByVal psCtaCod As String)

    Set oDCred = New COMDCredito.DCOMCredito
    Set R = oDCred.RecuperaCreditoConvenioAmpliacionPlazo(psCtaCod)
    Set oDCred = Nothing
    
    If Not R.BOF And Not R.EOF Then
        lblTitular.Caption = R!cPersNombre
        lblAnalista.Caption = R!cAnalista
        lblTipoCredito.Caption = R!cTipoCredDesc
        lblFechaUltCuota.Caption = R!dFecUltCuota
        lblProducto.Caption = R!cTipoProdDesc
        lblPrestamo.Caption = Format(R!nMontoCol, "#,##0.00") & " "
        lblSaldo.Caption = Format(R!nSaldo, "#,##0.00") & " "
        lblCuotas.Caption = R!nCuotas
        lblTasaInteres.Caption = Format(R!nTasaInteres, "#0.0000")
        
        fsFecVenc = Format(R!dFecVenc, "dd/mm/yyyy")
        fnPlazo = R!nPlazo
        fnTipoCuota = R!nTipoCuota
        fnTipoGracia = R!nTipoGracia
        fnPeriodoGracia = R!nPeriodoGracia
        fnFechaFija = R!nPeriodoFechaFija
        fnFechaFija2 = R!nPeriodoFechaFija2
        fbProxMes = R!nProxMes
        fnTasaGracia = R!nTasaGracia
        fsTpoCred = R!cTpoCredCod
        fsTpoProdCred = R!cTpoProdCod
        fbQuincenal = IIf(R!nColocCalendCod = "81", True, False)
        fnNroCalen = CInt(R!nNroCalen)
        ActxCta.Enabled = False
        cmdBuscar.Enabled = False
        fdUltCuotaPag = CDate(R!dUltCuotaPag)
        fnIntPag = CDate(R!nIntPag) 'WIOR 20150527
        Call CargarCalendario(psCtaCod)
    Else
        MsgBox "No se encuentra el crédito o el tipo de la institución relacionada al crédito no es admitido para esta opción", vbInformation, "Aviso"
    End If
End Sub

Private Sub Limpiar()
ActxCta.Enabled = True
cmdBuscar.Enabled = True
LimpiarCredito
lblTitular.Caption = ""
lblAnalista.Caption = ""
lblTipoCredito.Caption = ""
lblFechaUltCuota.Caption = ""
lblProducto.Caption = ""
lblPrestamo.Caption = ""
lblSaldo.Caption = ""
lblCuotas.Caption = ""
lblTasaInteres.Caption = ""
txtNuevoPlazo.Text = ""
txtDias.Text = ""
chkReprogramar.value = 0
txtGlosa.Text = ""

Set oDCred = Nothing
Set oDCalend = Nothing
Set oNGasto = Nothing
Set oNCred = Nothing
 
Set MatGracia = Nothing
Set MatCalend = Nothing
Set MatCalend_2 = Nothing
Set MatDesemb = Nothing
Set MatGastos = Nothing

nNumGastos = 0
Set MatCalendIC = Nothing
Set R = Nothing

fnPlazo = 0
fnTipoCuota = 0
fnTipoGracia = 0
fsFecVenc = "01/01/1900"
fnPeriodoGracia = 0
fnFechaFija = 0
fnFechaFija2 = 0
fbProxMes = False
fnTasaGracia = 0
fsTpoCred = ""
fsTpoProdCred = ""
fbQuincenal = False
fnNroCalen = 0

fnTCuota = 0
fnTPeriodo = 0
fnDiasReprog = 0
fnMontoApr = 0
fnSaldoGracia = 0
fdFechaPago = "01/01/1900"
fdUltCuotaPag = "01/01/1900"
LimpiaFlex FECalend
fnIntPag = 0 'WIOR 20150527
End Sub

Function SoloNumeros(ByVal KeyAscii As Integer) As Integer
    'permite que solo sean ingresados los numeros, el ENTER y el RETROCESO
    If InStr("0123456789", Chr(KeyAscii)) = 0 Then
        SoloNumeros = 0
    Else
        SoloNumeros = KeyAscii
    End If
    ' teclas especiales permitidas
    If KeyAscii = 8 Then SoloNumeros = KeyAscii ' borrado atras
    If KeyAscii = 13 Then SoloNumeros = KeyAscii 'Enter
End Function

Private Sub txtDias_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNumeros(KeyAscii)
    LimpiaFlex FECalend
    If KeyAscii = 13 Then
        cmdCalcular.SetFocus
    End If
End Sub

Private Sub txtDias_KeyUp(KeyCode As Integer, Shift As Integer)
    If CInt(IIf(txtDias.Text = "", 0, txtDias.Text)) > fnValosDiasReprog Then
        MsgBox "El plazo no debe ser superior a " & fnValosDiasReprog, vbInformation, "Aviso"
        txtDias.Text = fnValosDiasReprog
    End If
End Sub

Private Sub txtNuevoPlazo_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNumeros(KeyAscii)
    LimpiaFlex FECalend
    If KeyAscii = 13 Then
        If txtDias.Visible And txtDias.Enabled Then
            txtDias.SetFocus
        End If
    End If
End Sub

Private Sub txtNuevoPlazo_KeyUp(KeyCode As Integer, Shift As Integer)
    If IIf(txtNuevoPlazo.Text = "", 0, txtNuevoPlazo.Text) > 84 Then
        MsgBox "El plazo no debe ser superior a 84", vbInformation, "Aviso"
        txtNuevoPlazo.Text = 84
    End If
End Sub

Private Function ValidaDatos() As Boolean
    ValidaDatos = False
    
    If ValidaFlexVacio Then
        MsgBox "No se generó el calendario", vbInformation, "Aviso"
        Exit Function
    End If
    
    If Trim(txtGlosa.Text) = "" Then
        MsgBox "Ingrese la glosa para continuar", vbInformation, "Aviso"
        Exit Function
    End If
    
    If Trim(txtNuevoPlazo.Text) = "" Then
        MsgBox "Ingrese la nueva cantidad de cuotas para continuar", vbInformation, "Aviso"
        Exit Function
    End If
    
    ValidaDatos = True
End Function

Public Function ValidaFlexVacio() As Boolean
    Dim i As Integer
    For i = 1 To FECalend.Rows - 1
        If FECalend.TextMatrix(i, 1) = "" Then
            ValidaFlexVacio = True
            Exit Function
        End If
    Next i
End Function

Private Sub CargarCalendario(ByVal psCtaCod As String)
Dim nCuoPag As Integer, nCuoNoPag As Integer
Dim lnCapital As Double
Dim lnIntComp As Double
Dim lnIntGra As Double
Dim nMontoApr As Double
Dim lnSaldoNew As Double
Dim lnIntMora As Double 'WIOR 20150528
    Set oDCalend = New COMDCredito.DCOMCalendario
    Set R = oDCalend.RecuperaCalendarioPagos(psCtaCod)
    Set oDCalend = Nothing
    
    Set oDCred = New COMDCredito.DCOMCredito
    nMontoApr = oDCred.SaldoPactadoCredito(psCtaCod)
    Set oDCred = Nothing
    
    Do While Not R.EOF
        nCuoPag = nCuoPag + 1
        nCuoNoPag = nCuoNoPag + 1
        If R!nColocCalendEstado = gColocCalendEstadoPagado Then
            lnCapital = R!nCapital
            lnIntComp = R!nIntComp
            lnIntGra = R!nIntGracia
            lnIntMora = R!nIntMor  'WIOR 20150528
        Else
            lnCapital = R!nCapital - R!nCapitalPag
            lnIntComp = R!nIntComp - R!nIntCompPag
            lnIntGra = R!nIntGracia - R!nIntGraciaPag
            lnIntMora = R!nIntMor - R!nIntMorPag 'WIOR 20150528
        End If
        
        FECalend.AdicionaFila
        FECalend.TextMatrix(R.Bookmark, 1) = Format(R!dVenc, "dd/mm/yyyy")
        FECalend.TextMatrix(R.Bookmark, 2) = Trim(str(R!nCuota))
        FECalend.TextMatrix(R.Bookmark, 3) = Format(IIf(IsNull(lnCapital), 0, lnCapital) + _
                                        IIf(IsNull(lnIntComp), 0, lnIntComp) + _
                                        IIf(IsNull(lnIntGra), 0, lnIntGra) + _
                                        IIf(IsNull(R!nIntMor), 0, R!nIntMor) + _
                                        IIf(IsNull(R!nIntReprog), 0, R!nIntReprog) + _
                                        IIf(IsNull(R!nGasto), 0, R!nGasto), "#0.00")
        FECalend.TextMatrix(R.Bookmark, 4) = Format(IIf(IsNull(lnCapital), 0, lnCapital), "#0.00")
        FECalend.TextMatrix(R.Bookmark, 5) = Format(IIf(IsNull(lnIntComp), 0, lnIntComp), "#0.00")
        FECalend.TextMatrix(R.Bookmark, 6) = Format(IIf(IsNull(lnIntMora), 0, lnIntMora), "#0.00") 'WIOR 20150528 CAMBIO R!nIntMor POR lnIntMora
        FECalend.TextMatrix(R.Bookmark, 7) = Format(IIf(IsNull(R!nIntReprog), 0, R!nIntReprog), "#0.00")
        FECalend.TextMatrix(R.Bookmark, 8) = Format(IIf(IsNull(lnIntGra), 0, lnIntGra), "#0.00")
        FECalend.TextMatrix(R.Bookmark, 9) = Format(IIf(IsNull(R!nGasto), 0, R!nGasto), "#0.00")
        nMontoApr = nMontoApr - IIf(IsNull(R!nCapital), 0, R!nCapital)
        nMontoApr = CDbl(Format(nMontoApr, "#0.00"))
        FECalend.TextMatrix(R.Bookmark, 10) = Format(nMontoApr, "#0.00")
        FECalend.TextMatrix(R.Bookmark, 11) = Trim(str(R!nColocCalendEstado))
        FECalend.TextMatrix(R.Bookmark, 12) = Format(IIf(IsNull(R!nCapitalPag), 0, R!nCapitalPag), "#0.00")
        
        lnSaldoNew = lnSaldoNew + IIf(IsNull(R!nCapital), 0, R!nCapital) - IIf(IsNull(R!nCapitalPag), 0, R!nCapitalPag)
        If R!nColocCalendEstado = gColocCalendEstadoPagado Then
            FECalend.row = R.Bookmark
            Call FECalend.ForeColorRow(vbRed)
            nCuoNoPag = nCuoNoPag - 1
        End If
        R.MoveNext
    Loop
    FECalend.TopRow = 1 'WIOR 20150527
    R.Close
    Set R = Nothing
End Sub

Private Sub RegularCalendario()
Dim Y As Integer
Dim MatCalendTemp() As String
Dim i As Integer
ReDim MatCalendTemp(UBound(MatCalend) - 1, 13)
For i = 0 To UBound(MatCalend) - 1
    For Y = 0 To 13
        MatCalendTemp(i, Y) = MatCalend(i + 1, Y)
    Next Y
Next i
Erase MatCalend
ReDim MatCalend(UBound(MatCalendTemp), 13)

For i = 0 To UBound(MatCalendTemp)
    For Y = 0 To 13
        MatCalend(i, Y) = MatCalendTemp(i, Y)
    Next Y
Next i
Erase MatCalendTemp

End Sub

Private Function GenerarMatrices(ByVal pbMostrar As Boolean, ByRef pbMatCalendIC As Variant, Optional ByRef pbErrorCargaCalendario As Boolean = False) As Variant
Set oNGasto = New COMNCredito.NCOMGasto

'WIOR 20140825 (Se adecuo ya que en la funcion GeneraCalendarioGastos_NEW no da para más parametros )*************************
Dim ReprogConvenio As Variant
ReDim ReprogConvenio(4, 0)
ReprogConvenio(0, 0) = "X"
'WIOR FIN *****************************

GenerarMatrices = oNGasto.GeneraCalendarioGastos_NEW(fnMontoApr, CDbl(lblTasaInteres.Caption), CInt(txtNuevoPlazo.Text), _
                        fnPlazo, CDate(MatDesemb(0, 0)), fnTCuota, _
                        fnTPeriodo, fnTasaGracia, fnPeriodoGracia, _
                        CDbl(fnTasaGracia), fnFechaFija, _
                        fbProxMes, MatGracia, 0, 0, MatCalend_2, _
                        MatDesemb, nNumGastos, gdFecSis, _
                        ActxCta.NroCuenta, 1, "DE", "F", _
                        CInt(txtNuevoPlazo.Text), fnMontoApr, , , , , , , , pbMostrar, _
                        2, True, 0, MatDesemb, fbQuincenal, pbErrorCargaCalendario, _
                        0, True, , _
                        gnITFMontoMin, gnITFPorcent, gbITFAplica, 0, Mid(ActxCta.NroCuenta, 4, 2), fsTpoProdCred, fsTpoCred, , , fnSaldoGracia, pbMatCalendIC, , chkSoloInt.value, , , ReprogConvenio) 'WIOR 20140825 CAMBIO True POR MatCalendSegDes
                        'WIOR 20150527 AGREGO chkSoloInt.value

Set oNGasto = Nothing
End Function

'WIOR 20150523 ***
Private Sub CargaVariables()
Dim oPar As COMDCredito.DCOMParametro
Dim oConsSist As COMDConstSistema.NCOMConstSistema

Set oPar = New COMDCredito.DCOMParametro
fnValosDiasReprog = oPar.RecuperaValorParametro(3216)
Set oPar = Nothing

Set oConsSist = New COMDConstSistema.NCOMConstSistema
fdFechaCredOtor = CDate(oConsSist.LeeConstSistema(489))
fdFechaCalif = CDate(oConsSist.LeeConstSistema(490))
fsCodCalif = oConsSist.LeeConstSistema(491)
fsTipoProd = oConsSist.LeeConstSistema(492)

Set oConsSist = Nothing
End Sub
'WIOR FIN ********
