VERSION 5.00
Begin VB.Form FrmColRecVistoRecup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Visto de Recuperaciones"
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7575
   Icon            =   "FrmColRecVistoRecup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   7575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraCredito 
      Caption         =   "Credito"
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
      Height          =   3735
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   7335
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "&Buscar..."
         Height          =   360
         Left            =   6120
         TabIndex        =   7
         Top             =   255
         Width           =   1080
      End
      Begin SICMACT.ActXCodCta AXCodCta 
         Height          =   465
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   3615
         _extentx        =   6376
         _extenty        =   820
         texto           =   "Crédito"
         enabledcta      =   -1  'True
         enabledprod     =   -1  'True
         enabledage      =   -1  'True
      End
      Begin VB.Label lblMetLiquid 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   6180
         TabIndex        =   37
         Top             =   1200
         Width           =   1065
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Met.Liquid."
         Height          =   195
         Index           =   4
         Left            =   5340
         TabIndex        =   36
         Top             =   1200
         Width           =   780
      End
      Begin VB.Label lblEstudioJur 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   900
         TabIndex        =   35
         Top             =   1560
         Width           =   4305
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Abogado"
         Height          =   195
         Index           =   5
         Left            =   90
         TabIndex        =   34
         Top             =   1560
         Width           =   645
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Comision"
         Height          =   195
         Index           =   6
         Left            =   5340
         TabIndex        =   33
         Top             =   1560
         Width           =   630
      End
      Begin VB.Label lblTipoCobranza 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3780
         TabIndex        =   32
         Top             =   1200
         Width           =   1395
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Cobranza"
         Height          =   195
         Index           =   3
         Left            =   2610
         TabIndex        =   31
         Top             =   1260
         Width           =   1035
      End
      Begin VB.Label lblCondicion 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   900
         TabIndex        =   30
         Top             =   1200
         Width           =   1485
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Condicion"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   29
         Top             =   1200
         Width           =   705
      End
      Begin VB.Label lblCliente 
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
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   900
         TabIndex        =   28
         Top             =   810
         Width           =   4305
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   27
         Top             =   840
         Width           =   480
      End
      Begin VB.Label lblComision 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   6180
         TabIndex        =   26
         Top             =   1560
         Width           =   1050
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Demanda"
         Height          =   195
         Index           =   1
         Left            =   5340
         TabIndex        =   25
         Top             =   840
         Width           =   690
      End
      Begin VB.Label lblDemanda 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   6180
         TabIndex        =   24
         Top             =   810
         Width           =   1080
      End
      Begin VB.Label lblTotalAct 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   6060
         TabIndex        =   23
         Top             =   3240
         Width           =   1155
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Total"
         Height          =   195
         Index           =   4
         Left            =   4920
         TabIndex        =   22
         Top             =   3300
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Saldo Capital"
         Height          =   195
         Index           =   0
         Left            =   4920
         TabIndex        =   21
         Top             =   2130
         Width           =   930
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Saldo Interes"
         Height          =   195
         Index           =   1
         Left            =   4920
         TabIndex        =   20
         Top             =   2400
         Width           =   930
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Saldo Gastos"
         Height          =   195
         Index           =   2
         Left            =   4920
         TabIndex        =   19
         Top             =   2700
         Width           =   945
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Saldo Mora"
         Height          =   195
         Index           =   3
         Left            =   4920
         TabIndex        =   18
         Top             =   3000
         Width           =   810
      End
      Begin VB.Label lblSaldAct 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   6060
         TabIndex        =   17
         Top             =   2940
         Width           =   1155
      End
      Begin VB.Label lblSaldAct 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   6060
         TabIndex        =   16
         Top             =   2640
         Width           =   1155
      End
      Begin VB.Label lblSaldAct 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   6060
         TabIndex        =   15
         Top             =   2340
         Width           =   1155
      End
      Begin VB.Label lblSaldAct 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   6060
         TabIndex        =   14
         Top             =   2040
         Width           =   1155
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Ing. Recup."
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   13
         Top             =   2100
         Width           =   840
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Ultima Tran"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   12
         Top             =   2580
         Width           =   810
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tasa Int"
         Height          =   195
         Index           =   7
         Left            =   180
         TabIndex        =   11
         Top             =   2940
         Width           =   585
      End
      Begin VB.Label lblIngresoRecup 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1020
         TabIndex        =   10
         Top             =   2100
         Width           =   1410
      End
      Begin VB.Label lblUltimaTran 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1020
         TabIndex        =   9
         Top             =   2520
         Width           =   1410
      End
      Begin VB.Label lblTasaInt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1020
         TabIndex        =   8
         Top             =   2880
         Width           =   1410
      End
   End
   Begin VB.Frame FraComandos 
      Height          =   675
      Left            =   120
      TabIndex        =   1
      Top             =   4560
      Width           =   7335
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar"
         Enabled         =   0   'False
         Height          =   360
         Left            =   3660
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdsalir 
         Caption         =   "&Salir"
         Height          =   360
         Left            =   6060
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   360
         Left            =   4800
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame FraComen 
      Caption         =   "Visto de Recupeciones"
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
      ForeColor       =   &H8000000D&
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   3780
      Width           =   7335
      Begin VB.CheckBox ChkVisto 
         Caption         =   "Cancelación"
         Height          =   255
         Index           =   2
         Left            =   5040
         TabIndex        =   40
         Top             =   360
         Width           =   1695
      End
      Begin VB.CheckBox ChkVisto 
         Caption         =   "Negociación"
         Height          =   255
         Index           =   1
         Left            =   2760
         TabIndex        =   39
         Top             =   360
         Width           =   1575
      End
      Begin VB.CheckBox ChkVisto 
         Caption         =   "Metodo de Liquidación"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   38
         Top             =   360
         Width           =   1935
      End
   End
End
Attribute VB_Name = "FrmColRecVistoRecup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************
'* VISTO - DE CREDITOS EN RECUPERACIONES
'Archivo:  frmColRecVistoRecup.frm
'AVMM   :  12/12/2006
'******************************************************

Option Explicit
Dim fnSaldoCap As Currency, fnSaldoIntComp As Currency, fnSaldoIntMorat As Currency, fnSaldoGasto As Currency
Dim fnNewSaldoCap As Currency, fnNewSaldoIntComp As Currency
Dim fnNewIntCompGen As Currency
Dim fnNroCalend As Integer

Dim fnPorcComision As Double
Dim fnComisionAbog As Currency, fnIntCompGenerado As Currency
Dim fsFecUltPago As String
Dim fnTasaInt As Double
Dim fsCondicion As String, fsDemanda As String
Dim fsCancSKMayorCero As String

Private Sub HabilitaControles(ByVal pbCmdGrabar As Boolean, ByVal pbCmdCancelar As Boolean, _
            ByVal pbCmdSalir As Boolean)
    cmdGrabar.Enabled = pbCmdGrabar
    cmdCancelar.Enabled = pbCmdCancelar
    cmdSalir.Enabled = pbCmdSalir
End Sub

Private Sub Limpiar()
Dim lnI As Integer
    Me.AxCodCta.Enabled = True
    Me.AxCodCta.NroCuenta = fgIniciaAxCuentaRecuperaciones
    Me.lblCliente.Caption = ""
    Me.lblDemanda.Caption = ""
    Me.lblCondicion.Caption = ""
    Me.lblTipoCobranza.Caption = ""
    Me.lblMetLiquid.Caption = ""
    Me.lblEstudioJur.Caption = ""
    Me.lblComision.Caption = ""
    Me.lblIngresoRecup.Caption = ""
    Me.lblUltimaTran.Caption = ""
    Me.lblTasaInt.Caption = ""
    
    For lnI = 0 To 3
        Me.lblSaldAct(lnI) = 0
    Next
    Me.lblTotalAct = 0
    ChkVisto(0).value = 0
    ChkVisto(1).value = 0
    ChkVisto(2).value = 0
    Me.FraComen.Enabled = False
End Sub

Private Sub AXCodCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then BuscaCredito (AxCodCta.NroCuenta)
    FraComen.Enabled = True
End Sub

Private Sub BuscaCredito(ByVal psCtaCod As String)
Dim lbOk As Boolean
Dim lrDatCredito As ADODB.Recordset
Dim lrDatGastos As New ADODB.Recordset
Dim loValCred As COMDColocRec.DCOMColRecCredito
Dim lnDiasUltTrans As Integer
Dim lnIntCompGenCal As Double

'On Error GoTo ControlError

Dim lsmensaje As String
    'Carga Datos
    Set loValCred = New COMDColocRec.DCOMColRecCredito
        Set lrDatCredito = loValCred.dObtieneDatosCancelaCredRecup(psCtaCod, lsmensaje)
        If Trim(lsmensaje) <> "" Then
            MsgBox lsmensaje, vbInformation, "Aviso"
            Exit Sub
        End If
    Set loValCred = Nothing
    
    If lrDatCredito Is Nothing Then   ' Hubo un Error
        MsgBox "No se Encontro el Credito ", vbInformation, "Aviso"
        Limpiar
        Set lrDatCredito = Nothing
        Exit Sub
    End If
        
        'Muestra Datos
        Me.lblCliente.Caption = PstaNombre(Trim(lrDatCredito!cPersNombre))
        Me.lblDemanda.Caption = IIf(lrDatCredito!nDemanda = gColRecDemandaSi, "S", "N")
        Me.lblCondicion = fgCondicionColRecupDesc(lrDatCredito!nPrdEstado)
        Me.lblTipoCobranza = IIf(lrDatCredito!nTipCJ = gColRecTipCobJudicial, "Judicial", "ExtraJudicial")
        Me.lblMetLiquid = lrDatCredito!cMetLiquid
        Me.lblEstudioJur.Caption = lrDatCredito!cPersNombreAbog
        Me.lblComision = lrDatCredito!nValorCom
        
        fnSaldoCap = lrDatCredito!nSaldo
        fnSaldoIntComp = lrDatCredito!nSaldoIntComp
        fnSaldoIntMorat = lrDatCredito!nSaldoIntMor
        fnSaldoGasto = lrDatCredito!nSaldoGasto
        fsFecUltPago = CDate(fgFechaHoraGrab(lrDatCredito!cUltimaActualizacion))
        fnNroCalend = lrDatCredito!nNroCalen
        fnTasaInt = lrDatCredito!nTasaInteres
        lnDiasUltTrans = CDate(Format(gdFecSis, "dd/mm/yyyy")) - CDate(Format(fsFecUltPago, "dd/mm/yyyy"))
        fnIntCompGenerado = lrDatCredito!nIntCompGen
        
        
        Me.lblIngresoRecup.Caption = Format(lrDatCredito!dIngRecup, "dd/mm/yyyy")
        Me.lblUltimaTran.Caption = Format(fsFecUltPago, "dd/mm/yyyy")
        Me.lblTasaInt.Caption = lrDatCredito!nTasaInteres

        
        Me.lblSaldAct(0) = Format(fnSaldoCap, "#,##0.00")
        Me.lblSaldAct(1) = Format(fnSaldoIntComp, "#,##0.00")
        Me.lblSaldAct(2) = Format(fnSaldoIntMorat, "#,##0.00")
        Me.lblSaldAct(3) = Format(fnSaldoGasto, "#,##0.00")
        Me.lblTotalAct = Format(fnSaldoCap + fnSaldoIntComp + fnSaldoIntMorat + fnSaldoGasto, "#,##0.00")
        ChkVisto(0).value = IIf(IsNull(lrDatCredito!nVisMetodo), 0, lrDatCredito!nVisMetodo)
        ChkVisto(1).value = IIf(IsNull(lrDatCredito!nVisNegociacion), 0, lrDatCredito!nVisNegociacion)
        ChkVisto(2).value = IIf(IsNull(lrDatCredito!nVisCancelacion), 0, lrDatCredito!nVisCancelacion)
        '***
    Set lrDatCredito = Nothing
        
    'Calcula el Int Comp Generado
    
    cmdGrabar.Enabled = True
    cmdGrabar.SetFocus
        
    AxCodCta.Enabled = False
    Call HabilitaControles(True, True, True)
Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox "Error: " & Err.Number & " " & Err.Description & vbCr & _
        "Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

Private Sub cmdBuscar_Click()
Dim loPers As COMDPersona.UCOMPersona
Dim lsPersCod As String, lsPersNombre As String
Dim lsEstados As String
Dim loPersCredito  As COMDColocRec.DCOMColRecCredito
Dim lrCreditos As New ADODB.Recordset
Dim loCuentas As COMDPersona.UCOMProdPersona

On Error GoTo ControlError

Set loPers = New COMDPersona.UCOMPersona
    Set loPers = frmBuscaPersona.Inicio
    If Not loPers Is Nothing Then
        lsPersCod = loPers.sPersCod
        lsPersNombre = loPers.sPersNombre
    Else
        Exit Sub
    End If
Set loPers = Nothing

' Selecciona Estados
lsEstados = gColocEstRecVigJud & "," & gColocEstRecVigCast

If Trim(lsPersCod) <> "" Then
    Set loPersCredito = New COMDColocRec.DCOMColRecCredito
        Set lrCreditos = loPersCredito.dObtieneCreditosDePersona(lsPersCod, lsEstados)
    Set loPersCredito = Nothing
End If

Set loCuentas = New COMDPersona.UCOMProdPersona
    Set loCuentas = frmProdPersona.Inicio(lsPersNombre, lrCreditos)
    If loCuentas.sCtaCod <> "" Then
        AxCodCta.NroCuenta = Mid(loCuentas.sCtaCod, 1, 18)
        AxCodCta.SetFocusCuenta
    End If
Set loCuentas = Nothing

Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

Private Sub cmdCancelar_Click()
    Limpiar
    Call HabilitaControles(False, True, True)
    AxCodCta.SetFocusAge
    FraComen.Enabled = False
End Sub

Private Sub cmdGrabar_Click()

'On Error GoTo ControlError
Dim loContFunct As COMNContabilidad.NCOMContFunciones
Dim loGrabar As COMNColocRec.NCOMColRecCredito
Dim loImprime As COMNColocRec.NCOMColRecImpre
Dim loPrevio As previo.clsPrevio

Dim lsMovNro As String
Dim lsFechaHoraGrab As String
Dim lsCadImprimir As String
Dim lsNombreCliente As String
Dim lsOpeCod As String

lsNombreCliente = Mid(Me.lblCliente.Caption, 1, 30)

If MsgBox(" Desea Grabar el Visto de Credito en Recuperaciones ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
    cmdGrabar.Enabled = False
        Set loGrabar = New COMNColocRec.NCOMColRecCredito
            Call loGrabar.bVistoRecuperaciones(AxCodCta.NroCuenta, ChkVisto(0).value, ChkVisto(1).value, ChkVisto(2).value)
        Set loGrabar = Nothing
        
        Limpiar
        
        AxCodCta.Enabled = True
        AxCodCta.SetFocus
        
Else
    MsgBox " Grabación cancelada ", vbInformation, " Aviso "
End If
Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Limpiar
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
    Me.AxCodCta.NroCuenta = fgIniciaAxCuentaRecuperaciones
End Sub


