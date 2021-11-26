VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmColRecCancelacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Colocaciones - Recuperaciones : Cancelaciones  de Credito"
   ClientHeight    =   5685
   ClientLeft      =   2430
   ClientTop       =   615
   ClientWidth     =   7515
   Icon            =   "frmColRecCancelacion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame FraComen 
      Caption         =   "Comentario"
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
      Height          =   1095
      Left            =   120
      TabIndex        =   39
      Top             =   3840
      Width           =   7335
      Begin VB.TextBox TxtComentario 
         Height          =   855
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   40
         Top             =   160
         Width           =   7095
      End
   End
   Begin VB.Frame FraComandos 
      Height          =   675
      Left            =   120
      TabIndex        =   19
      Top             =   4920
      Width           =   7335
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   360
         Left            =   4800
         TabIndex        =   22
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdsalir 
         Caption         =   "&Salir"
         Height          =   360
         Left            =   6060
         TabIndex        =   21
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar"
         Enabled         =   0   'False
         Height          =   360
         Left            =   3600
         TabIndex        =   20
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.TextBox Txttemp 
      Height          =   405
      Left            =   8430
      TabIndex        =   2
      Tag             =   "txtcodigo"
      Top             =   375
      Visible         =   0   'False
      Width           =   810
   End
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
      TabIndex        =   1
      Top             =   60
      Width           =   7335
      Begin SICMACT.ActXCodCta AXCodCta 
         Height          =   465
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   820
         Texto           =   "Crédito"
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "&Buscar..."
         Height          =   360
         Left            =   6120
         TabIndex        =   0
         Top             =   255
         Width           =   1080
      End
      Begin VB.Label lblTasaInt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1020
         TabIndex        =   38
         Top             =   2880
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
         TabIndex        =   37
         Top             =   2520
         Width           =   1410
      End
      Begin VB.Label lblIngresoRecup 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1020
         TabIndex        =   36
         Top             =   2100
         Width           =   1410
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tasa Int"
         Height          =   195
         Index           =   7
         Left            =   180
         TabIndex        =   35
         Top             =   2940
         Width           =   585
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Ultima Tran"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   34
         Top             =   2580
         Width           =   810
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Ing. Recup."
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   33
         Top             =   2100
         Width           =   840
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
         TabIndex        =   32
         Top             =   2040
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
         TabIndex        =   31
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
         Index           =   2
         Left            =   6060
         TabIndex        =   30
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
         Index           =   3
         Left            =   6060
         TabIndex        =   29
         Top             =   2940
         Width           =   1155
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Saldo Mora"
         Height          =   195
         Index           =   3
         Left            =   4920
         TabIndex        =   28
         Top             =   2680
         Width           =   810
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Saldo Gastos"
         Height          =   195
         Index           =   2
         Left            =   4920
         TabIndex        =   27
         Top             =   3000
         Width           =   945
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Saldo Interes"
         Height          =   195
         Index           =   1
         Left            =   4920
         TabIndex        =   26
         Top             =   2400
         Width           =   930
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Saldo Capital"
         Height          =   195
         Index           =   0
         Left            =   4920
         TabIndex        =   25
         Top             =   2130
         Width           =   930
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Total"
         Height          =   195
         Index           =   4
         Left            =   4920
         TabIndex        =   24
         Top             =   3300
         Width           =   720
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
      Begin VB.Label lblDemanda 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   6180
         TabIndex        =   18
         Top             =   810
         Width           =   1080
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Demanda"
         Height          =   195
         Index           =   1
         Left            =   5340
         TabIndex        =   17
         Top             =   840
         Width           =   690
      End
      Begin VB.Label lblComision 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   6180
         TabIndex        =   16
         Top             =   1560
         Width           =   1050
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   15
         Top             =   840
         Width           =   480
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
         TabIndex        =   14
         Top             =   810
         Width           =   4305
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Condicion"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   13
         Top             =   1200
         Width           =   705
      End
      Begin VB.Label lblCondicion 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   900
         TabIndex        =   12
         Top             =   1200
         Width           =   1485
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Cobranza"
         Height          =   195
         Index           =   3
         Left            =   2610
         TabIndex        =   11
         Top             =   1260
         Width           =   1035
      End
      Begin VB.Label lblTipoCobranza 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3780
         TabIndex        =   10
         Top             =   1200
         Width           =   1395
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Comision"
         Height          =   195
         Index           =   6
         Left            =   5340
         TabIndex        =   9
         Top             =   1560
         Width           =   630
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Abogado"
         Height          =   195
         Index           =   5
         Left            =   90
         TabIndex        =   8
         Top             =   1560
         Width           =   645
      End
      Begin VB.Label lblEstudioJur 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   900
         TabIndex        =   7
         Top             =   1560
         Width           =   4305
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Met.Liquid."
         Height          =   195
         Index           =   4
         Left            =   5340
         TabIndex        =   6
         Top             =   1200
         Width           =   780
      End
      Begin VB.Label lblMetLiquid 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   6180
         TabIndex        =   5
         Top             =   1200
         Width           =   1065
      End
   End
   Begin RichTextLib.RichTextBox rtf 
      Height          =   300
      Left            =   8385
      TabIndex        =   3
      Top             =   885
      Visible         =   0   'False
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   529
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmColRecCancelacion.frx":030A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmColRecCancelacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************
'* RECUPERACIONES - CANCELACION DE CREDITOS EN RECUPERACIONES
'Archivo:  frmColRecCancelacion.frm
'LAYG   :  15/08/2002.
'Resumen:  Nos permite registrar la Cancelacion de un Credito en Recuperaciones
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
    cmdsalir.Enabled = pbCmdSalir
End Sub

Private Sub Limpiar()
Dim lnI As Integer
    Me.AXCodCta.Enabled = True
    Me.AXCodCta.NroCuenta = fgIniciaAxCuentaRecuperaciones
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
    Me.TxtComentario.Text = ""
    Me.FraComen.Enabled = False
End Sub

Private Sub AXCodCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then BuscaCredito (AXCodCta.NroCuenta)
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
        ' Asigna Valores a las Variables
        fnSaldoCap = lrDatCredito!nSaldo
        fnSaldoIntComp = lrDatCredito!nSaldoIntComp
        fnSaldoIntMorat = lrDatCredito!nSaldoIntMor
        fnSaldoGasto = lrDatCredito!nSaldoGasto
        fsFecUltPago = CDate(fgFechaHoraGrab(lrDatCredito!cUltimaActualizacion))
        fnNroCalend = lrDatCredito!nNroCalen
        fnTasaInt = lrDatCredito!nTasaInteres
        lnDiasUltTrans = CDate(Format(gdFecSis, "dd/mm/yyyy")) - CDate(Format(fsFecUltPago, "dd/mm/yyyy"))
        fnIntCompGenerado = lrDatCredito!nIntCompGen
        
        'Muestra Datos
        Me.lblCliente.Caption = PstaNombre(Trim(lrDatCredito!cPersNombre))
        Me.lblDemanda.Caption = IIf(lrDatCredito!nDemanda = gColRecDemandaSi, "S", "N")
        Me.lblCondicion = fgCondicionColRecupDesc(lrDatCredito!nPrdEstado)
        Me.lblTipoCobranza = IIf(lrDatCredito!nTipCJ = gColRecTipCobJudicial, "Judicial", "ExtraJudicial")
        Me.lblMetLiquid = lrDatCredito!cMetLiquid
        Me.lblEstudioJur.Caption = lrDatCredito!cPersNombreAbog
        Me.lblComision = lrDatCredito!nValorCom
        fnPorcComision = lrDatCredito!nValorCom
        fsCondicion = IIf(lrDatCredito!nPrdEstado = gColocEstRecVigJud, "J", "A")
        fsDemanda = IIf(lrDatCredito!nDemanda = gColRecDemandaSi, "S", "N")
        
        Me.lblIngresoRecup.Caption = Format(lrDatCredito!dIngRecup, "dd/mm/yyyy")
        Me.lblUltimaTran.Caption = Format(fsFecUltPago, "dd/mm/yyyy")
        Me.lblTasaInt.Caption = lrDatCredito!nTasaInteres

        
        Me.lblSaldAct(0) = Format(fnSaldoCap, "#,##0.00")
        Me.lblSaldAct(1) = Format(fnSaldoIntComp, "#,##0.00")
        Me.lblSaldAct(2) = Format(fnSaldoIntMorat, "#,##0.00")
        Me.lblSaldAct(3) = Format(fnSaldoGasto, "#,##0.00")
        Me.lblTotalAct = Format(fnSaldoCap + fnSaldoIntComp + fnSaldoIntMorat + fnSaldoGasto, "#,##0.00")
        '***
    Set lrDatCredito = Nothing
        
    'Calcula el Int Comp Generado
    Dim loCalcula As COMNColocRec.NCOMColRecCalculos
    Set loCalcula = New COMNColocRec.NCOMColRecCalculos
        lnIntCompGenCal = loCalcula.nCalculaIntCompGenerado(lnDiasUltTrans, fnTasaInt, fnSaldoCap)
    Set loCalcula = Nothing
    fnNewSaldoIntComp = fnSaldoIntComp + lnIntCompGenCal
    fnNewIntCompGen = fnIntCompGenerado + lnIntCompGenCal
    
    cmdGrabar.Enabled = True
    cmdGrabar.SetFocus
        
    AXCodCta.Enabled = False
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
        AXCodCta.NroCuenta = Mid(loCuentas.sCtaCod, 1, 18)
        AXCodCta.SetFocusCuenta
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
    AXCodCta.SetFocusAge
    FraComen.Enabled = False
End Sub

Private Sub cmdGrabar_Click()

'On Error GoTo ControlError
Dim loContFunct As COMNContabilidad.NCOMContFunciones
Dim loGrabar As COMNColocRec.NCOMColRecCredito
Dim loImprime As COMNColocRec.NCOMColRecImpre
Dim loPrevio As previo.clsprevio

Dim lsMovNro As String
Dim lsFechaHoraGrab As String
Dim lsCadImprimir As String
Dim lsNombreCliente As String
Dim lsOpeCod As String

lsNombreCliente = Mid(Me.lblCliente.Caption, 1, 30)
lsOpeCod = "130600"
If ValidaOperacion = False Then
    Exit Sub
End If

'********* VERIFICAR VISTO AVMM - 13-12-2006 **********************
'Dim loVisto As COMDColocRec.DCOMColRecCredito
'Set loVisto = New COMDColocRec.DCOMColRecCredito
'    '3=Cancelación
'    If loVisto.bVerificarVisto(AXCodCta.NroCuenta, 3) = False Then
'        MsgBox "No existe Visto para realizar Negociación"
'        Exit Sub
'    End If
'Set loVisto = Nothing
'********************************************************************


If MsgBox(" Desea Grabar Cancelacion de Credito en Recuperaciones ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
    cmdGrabar.Enabled = False
        
        'Genera el Mov Nro
        Set loContFunct = New COMNContabilidad.NCOMContFunciones
            lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        Set loContFunct = Nothing
        
        lsFechaHoraGrab = fgFechaHoraGrab(lsMovNro)
        Set loGrabar = New COMNColocRec.NCOMColRecCredito
            'Grabar Cancelacion de Credito en Recuperaciones
            Call loGrabar.nCancelacionCreditoRecup(AXCodCta.NroCuenta, lsFechaHoraGrab, lsOpeCod, _
                 lsMovNro, IIf(fsCondicion = "J", gColocEstRecCanJud, gColocEstRecCanCast), fnSaldoCap, _
                 fnNewSaldoIntComp, fnNewIntCompGen, fnNroCalend, False, TxtComentario.Text)
        Set loGrabar = Nothing

        'Impresión
'        Set loImprime = New COMNColocRec.NCOMColRecImpre
'           lsCadImprimir = loImprime.nPrintReciboPagoCredRecup(gsNomAge, lsFechaHoraGrab, AXCodCta.NroCuenta, _
'                lsNombreCliente, CCur(Me.AXMontoPago.Text), gsCodUser, lsmensaje)
'        If Trim(lsmensaje) <> "" Then
'             MsgBox lsmensaje, vbInformation, "Aviso"
'             Exit Sub
'        End If
'        Set loImprime = Nothing
'        Set loPrevio = New Previo.clsPrevio
'            loPrevio.Show lsCadImprimir, "Recuperaciones - Cancelacion de Credito", True
'        Set loPrevio = Nothing
        
        Limpiar
        
        AXCodCta.Enabled = True
        AXCodCta.SetFocus
        
Else
    MsgBox " Grabación cancelada ", vbInformation, " Aviso "
End If
Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Limpiar
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
    Me.AXCodCta.NroCuenta = fgIniciaAxCuentaRecuperaciones
End Sub

Private Function ValidaOperacion() As Boolean
'--- Valida la grabacion

ValidaOperacion = True
'Verifica si el Saldo de Capital es igual a Cero
'MADM 20110804
If fnSaldoCap > 0 Then
    'If fsCancSKMayorCero = "N" Then
        MsgBox "Credito tiene Capital Vigente,  NO puede cancelarlo. ", vbInformation, "Aviso"
        ValidaOperacion = False
    'End If
End If
'END MADM
End Function

Private Sub fCargaParametro()
Dim loParam As COMDConstSistema.NCOMConstSistema
Set loParam = New COMDConstSistema.NCOMConstSistema
    fsCancSKMayorCero = loParam.LeeConstSistema(154)
Set loParam = Nothing

End Sub

Private Sub TxtComentario_KeyPress(KeyAscii As Integer)
     KeyAscii = fgIntfMayusculas(KeyAscii)
     If KeyAscii = 13 Then
        cmdGrabar.SetFocus
     End If
End Sub
