VERSION 5.00
Begin VB.Form frmColPCancelacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Crédito Pignoraticio -  Cancelación de Crédito"
   ClientHeight    =   7485
   ClientLeft      =   1200
   ClientTop       =   2445
   ClientWidth     =   8025
   Icon            =   "frmColPCancelacion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7485
   ScaleWidth      =   8025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   380
      Left            =   5820
      TabIndex        =   1
      Top             =   7000
      Width           =   975
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   380
      Left            =   6900
      TabIndex        =   2
      Top             =   7000
      Width           =   975
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   380
      Left            =   4740
      TabIndex        =   0
      Top             =   7000
      Width           =   990
   End
   Begin VB.Frame fraContenedor 
      Height          =   6210
      Index           =   0
      Left            =   135
      TabIndex        =   3
      Top             =   75
      Width           =   7785
      Begin VB.TextBox txtCampRete 
         Height          =   285
         Left            =   5640
         TabIndex        =   40
         Text            =   "0.00"
         Top             =   240
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton cmdBuscar 
         Height          =   345
         Left            =   7110
         Picture         =   "frmColPCancelacion.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Buscar ..."
         Top             =   180
         Width           =   420
      End
      Begin VB.Frame fraContenedor 
         Height          =   1800
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   4320
         Width           =   7395
         Begin VB.TextBox txtPgoAdelInt 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Height          =   315
            Left            =   3525
            TabIndex        =   32
            Top             =   1320
            Width           =   1215
         End
         Begin VB.TextBox txtCostoRemate 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Height          =   315
            Left            =   3525
            TabIndex        =   30
            Top             =   960
            Width           =   1215
         End
         Begin VB.TextBox txtInteres 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Height          =   315
            Left            =   885
            TabIndex        =   23
            Top             =   960
            Width           =   1215
         End
         Begin VB.TextBox txtMora 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Height          =   315
            Left            =   885
            TabIndex        =   22
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox txtCapital 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Height          =   315
            Left            =   885
            TabIndex        =   21
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox txtIntVen 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Height          =   315
            Left            =   885
            TabIndex        =   20
            Top             =   1320
            Width           =   1215
         End
         Begin VB.TextBox txtCostoCus 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Height          =   315
            Left            =   3525
            TabIndex        =   19
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox txtCostoNoti 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Height          =   315
            Left            =   3525
            TabIndex        =   18
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox txtDiasAtraso 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Height          =   315
            Left            =   6255
            TabIndex        =   4
            Top             =   240
            Width           =   1050
         End
         Begin VB.Label lblInteres 
            AutoSize        =   -1  'True
            Caption         =   "Amortización Int.:"
            Height          =   195
            Index           =   1
            Left            =   2160
            TabIndex        =   33
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label lblCostoRemate 
            AutoSize        =   -1  'True
            Caption         =   "Costo Remate"
            Height          =   195
            Index           =   0
            Left            =   2160
            TabIndex        =   31
            Top             =   1080
            Width           =   1005
         End
         Begin VB.Label lblInteres 
            AutoSize        =   -1  'True
            Caption         =   "Interés"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   29
            Top             =   1080
            Width           =   480
         End
         Begin VB.Label lblMora 
            AutoSize        =   -1  'True
            Caption         =   "Mora"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   28
            Top             =   720
            Width           =   360
         End
         Begin VB.Label lblCapital 
            AutoSize        =   -1  'True
            Caption         =   "Capital"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   27
            Top             =   360
            Width           =   480
         End
         Begin VB.Label lblIntVen 
            AutoSize        =   -1  'True
            Caption         =   "Int. Vcdo."
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   26
            Top             =   1440
            Width           =   690
         End
         Begin VB.Label lblCostoCus 
            AutoSize        =   -1  'True
            Caption         =   "Costo Custodia"
            Height          =   195
            Index           =   4
            Left            =   2160
            TabIndex        =   25
            Top             =   360
            Width           =   1065
         End
         Begin VB.Label lblCostoNoti 
            AutoSize        =   -1  'True
            Caption         =   "Costo Notificación"
            Height          =   195
            Index           =   5
            Left            =   2160
            TabIndex        =   24
            Top             =   720
            Width           =   1290
         End
         Begin VB.Label LblITF 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
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
            Height          =   315
            Left            =   6270
            TabIndex        =   17
            Top             =   960
            Width           =   1005
         End
         Begin VB.Label LblMontoPagar 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
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
            Height          =   315
            Left            =   6270
            TabIndex        =   16
            Top             =   1320
            Width           =   1005
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "ITF:"
            Height          =   195
            Left            =   5850
            TabIndex        =   15
            Top             =   1080
            Width           =   285
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Mono Pagar:"
            Height          =   195
            Left            =   5205
            TabIndex        =   14
            Top             =   1440
            Width           =   915
         End
         Begin VB.Label lblTotalDeuda 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
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
            Height          =   315
            Left            =   6270
            TabIndex        =   9
            Top             =   600
            Width           =   1005
         End
         Begin VB.Label lblMoneda 
            Height          =   255
            Left            =   240
            TabIndex        =   8
            Top             =   240
            Width           =   360
         End
         Begin VB.Label lblEtiqueta 
            AutoSize        =   -1  'True
            Caption         =   "Monto  Deuda:"
            Height          =   195
            Index           =   6
            Left            =   5040
            TabIndex        =   7
            Top             =   720
            Width           =   1065
         End
         Begin VB.Label lblEtiqueta 
            AutoSize        =   -1  'True
            Caption         =   "Dias de atraso:"
            Height          =   195
            Index           =   7
            Left            =   5040
            TabIndex        =   6
            Top             =   360
            Width           =   1065
         End
      End
      Begin SICMACT.ActXCodCta AXCodCta 
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   3615
         _extentx        =   6376
         _extenty        =   661
         texto           =   "Crédito"
         enabledcta      =   -1
         enabledprod     =   -1
         enabledage      =   -1
      End
      Begin SICMACT.ActXColPDesCon AXDesCon 
         Height          =   3735
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   6588
      End
      Begin VB.Label lblCampRetenPrend 
         Caption         =   "Campaña y Retencion"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   3960
         TabIndex        =   41
         Top             =   240
         Visible         =   0   'False
         Width           =   3135
      End
   End
   Begin VB.Frame fraFormaPago 
      Height          =   600
      Left            =   135
      TabIndex        =   39
      Top             =   6300
      Width           =   7785
      Begin VB.ComboBox CmbForPag 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   200
         Width           =   1785
      End
      Begin SICMACT.ActXCodCta txtCuentaCargo 
         Height          =   375
         Left            =   3120
         TabIndex        =   38
         Top             =   200
         Visible         =   0   'False
         Width           =   3630
         _extentx        =   6403
         _extenty        =   661
         texto           =   "Cuenta N°:"
         enabledcta      =   -1
         enabledage      =   -1
      End
      Begin VB.Label lblNroDocumento 
         AutoSize        =   -1  'True
         Caption         =   "Nº Documento"
         Height          =   195
         Left            =   3105
         TabIndex        =   37
         Top             =   250
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.Label LblNumDoc 
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   4300
         TabIndex        =   36
         Top             =   200
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.Label lblFormaPago 
         AutoSize        =   -1  'True
         Caption         =   "Forma Pago"
         Height          =   195
         Left            =   180
         TabIndex        =   34
         Top             =   250
         Width           =   855
      End
   End
   Begin VB.Label lblMensaje 
      Caption         =   "<F12> Número Contrato Antiguo"
      Height          =   165
      Left            =   180
      TabIndex        =   10
      Top             =   7000
      Width           =   2655
   End
End
Attribute VB_Name = "frmColPCancelacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************
'* CANCELACION DE CONTRATO PIGNORATICIO
'Archivo:  frmColPCancelacion.frm
'LAYG   :  10/07/2001.
'Resumen:  Nos permite cancelar la deuda total de credito pignoraticio
Option Explicit

Option Compare Text

Dim fnVarOpeCod As Long
Dim fsVarOpeDesc As String
Dim fsVarPersCodCMAC As String
Dim fsVarNombreCMAC As String

Dim fnTasaPreparacionRemate As Double
Dim fnTasaCustodiaVencida As Double
Dim fnTasaImpuesto As Double
Dim fnVarDiasCambCart As Integer

Dim fnVarTasaInteres As Double

Dim vTasaInteres As Double ' peac 20070820

Dim vNroContrato As String
Dim vFecContrato As Date
Dim vFecVencimiento As Date
Dim vdiasAtraso As Double
Dim vDiasAtrasoReal As Double
Dim vDeuda As Double
Dim vInteresVencido As Double
Dim vInteresMoratorio As Double
Dim vImpuesto As Double
Dim fnVarGastoCorrespondencia As Double
Dim vCostoCustodiaMoratorio As Double
Dim vOroNeto As Double
Dim vValorTasacion As Double
Dim vSaldoCapital As Double
Dim vNewSaldoCapital As Double
Dim vPlazo As Integer
Dim vEstado As String
Dim vCostoPreparacionRemate As Double
Dim vPrestamo As Double
Dim vTasaInteresVencido As Double
Dim vTasaInteresMoratorio As Double
Dim vCostoCustMoratorio As Double
Dim v14k As Double
Dim v16k As Double
Dim v18k As Double
Dim v21k As Double

Dim fnVarCostoNotificacion As Currency 'PEAC 20070926

Dim gnNotifiAdju As Integer  ' peac 20080515
Dim gnNotifiCob As Integer  ' PEAC 20080715

Dim gcCredAntiguo As String  ' peac 20070923
Dim fnVarEstUltProcRem As Integer 'DAOR 20070714
Dim fsColocLineaCredPig As String, vFecEstado As Date ' PEAC 20070813
Dim vDiasAdel As Integer, vInteresAdel As Double, vMontoCol As Double ' PEAC 20070813
Dim nRedondeoITF As Double ' BRGO 20110914
Dim fnDiaFer As Integer 'RECO20141226 ERS170-2014
Dim fnPgoAdelInt As Double '*** Peac 20161116
Dim nPagoIntVenMor As Double '*** PEAC 20170331
Dim fnIntPendSaldo As Double  '= lrValida!nPagosPendIntSaldo
Private nMontoVoucher As Currency 'CTI4 ERS0112020
Dim nMovNroRVD As Long, nMovNroRVDPend As Long 'CTI4 ERS0112020
Dim sNumTarj As String 'CTI4 ERS0112020
Dim loVistoElectronico As frmVistoElectronico 'CTI4 ERS0112020
Dim nRespuesta As Integer 'CTI4 ERS0112020

'Dim ventana As Integer 'MADM 20090928
Public Sub Inicio(ByVal pnOpeCod As Long, ByVal psOpeDesc As String, _
        ByVal psPersCodCMAC As String, ByVal psNomCmac As String)

    fnVarOpeCod = pnOpeCod
    fsVarOpeDesc = psOpeDesc
    fsVarPersCodCMAC = psPersCodCMAC
    fsVarNombreCMAC = psNomCmac
    
    Select Case fnVarOpeCod
        Case gColPOpeCancelacEFE
            'txtDocumento.Visible = false
        Case gColPOpeCancelacCHQ
            'txtDocumento.Visible = True
    '    Case Else
    '        txtDocumento.Visible = False
    End Select
    CargaParametros
    Limpiar
    Me.Show 1
End Sub

'Inicializa las variables del formulario
Private Sub Limpiar()
    Me.AXCodCta.NroCuenta = fgIniciaAxCuentaPignoraticio
    AXDesCon.Limpiar
    txtDiasAtraso.Text = ""
    lblTotalDeuda.Caption = "0.00"
    Me.LblMontoPagar = "0.00"
    Me.LblITF = "0.00"
    nRedondeoITF = 0
    fnDiaFer = 0
    
    Me.txtCapital.Text = "": Me.txtCostoCus.Text = ""
    Me.txtMora.Text = "": Me.txtCostoNoti.Text = ""
    Me.txtInteres.Text = "": Me.txtCostoRemate.Text = ""
    Me.txtIntVen.Text = "": Me.txtPgoAdelInt.Text = ""
    CmbForPag.ListIndex = -1 'CTI4 ERS0112020
    txtCuentaCargo.NroCuenta = "" 'CTI4 ERS0112020
    LblNumDoc.Caption = "" 'CTI4 ERS0112020
    cmdGrabar.Enabled = False 'CTI4 ERS0112020
    sNumTarj = "" 'CTI4 ERS0112020
    'JOEP20210921 campana prendario
    lblCampRetenPrend.Visible = False
    lblCampRetenPrend.Caption = ""
    txtCampRete.Text = 0#
    'JOEP20210921 campana prendario
End Sub

'Busca el contrato ingresado
Private Sub BuscaContrato(ByVal psNroContrato As String)
Dim lbok As Boolean
Dim lrValida As ADODB.Recordset
Dim loValContrato As COMNColoCPig.NCOMColPValida 'nColPValida
Dim loCalculos As COMNColoCPig.NCOMColPCalculos 'NColPCalculos
Dim loPigFunc As COMDColocPig.DCOMColPFunciones 'dColPFunciones
Dim lnDeuda As Currency, lnMinimoPagar As Currency
Dim lnDiasAtraso  As Integer
Dim lsFecVenTemp As String
Dim lsmensaje As String
'----- MADM 20091120 ---------------------
Dim lafirma As frmPersonaFirma
Dim ClsPersona As COMDPersona.DCOMPersonas
Dim Rf As ADODB.Recordset
Dim loParam As COMDColocPig.DCOMColPCalculos
Dim lbNrRenovacion As Integer 'JOEP20210923 campna prendario
Set loParam = New COMDColocPig.DCOMColPCalculos
'----- END MADM --------------------------

'On Error GoTo ControlError
    'gITF.fgITFParamAsume (Mid(psNroContrato, 4, 2))
    'Valida Contrato
    Set lrValida = New ADODB.Recordset
    Set loValContrato = New COMNColoCPig.NCOMColPValida
        Set lrValida = loValContrato.nValidaCancelacionCredPignoraticio(psNroContrato, gdFecSis, 0, lsmensaje)
        If Trim(lsmensaje) <> "" Then
             MsgBox lsmensaje, vbInformation, "Aviso"
             Exit Sub
        End If
        
    Set loValContrato = Nothing
    
    If lrValida Is Nothing Then ' Hubo un Error
        Limpiar
        Set lrValida = Nothing
        Exit Sub
    End If
    
    ' Asigna Valores a las Variables
    vValorTasacion = Format(lrValida!nTasacion, "#0.00")
    nPagoIntVenMor = lrValida!nPagosIntVenMora
    
    vTasaInteresVencido = lrValida!nTasaIntVenc
    vTasaInteresMoratorio = lrValida!nTasaIntMora

    vEstado = lrValida!nPrdEstado
    
    gcCredAntiguo = lrValida!cCredB 'PEAC 20070925
    
    gnNotifiAdju = lrValida!nCodNotifiAdj 'PEAC 20080515
    gnNotifiCob = lrValida!nCodNotifiCob 'PEAC 20080715
    
    vFecEstado = lrValida!dPrdEstado ' PEAC 20070813
    vSaldoCapital = lrValida!nMontoCol ' PEAC 20070813
        
    'vCostoCustMoratorio = Format(RegCredPrend!nCostCusto, "#0.00")
    vSaldoCapital = Format(lrValida!nSaldo, "#0.00")
    vTasaInteres = lrValida!nTasaInteres
    'vTasaImpuesto = Format(RegCredPrend!nTasaImpu, "#0.00")
    vFecVencimiento = Format(lrValida!dVenc, "dd/mm/yyyy")
    'Me.txtFecVencimiento.Text = Format(lrValida!dFecVenc, "dd/mm/yyyy")
    fnVarEstUltProcRem = lrValida!nEstUltProcRem 'DAOR 20070714
    
    fnIntPendSaldo = lrValida!nPagosPendIntSaldo '*** PEAC 20170926
    fnPgoAdelInt = lrValida!nPagosPendIntPagados '*** PEAC 20161116
    lbNrRenovacion = lrValida!nNroRenov 'JOEP20210923 Campana Prendario
    'Muestra Datos
    If fgMuestraCredPig_AXDesCon(psNroContrato, Me.AXDesCon, False) Then

    End If
    
    'JOEP20210921 Campana Prendario
    Call CampPrendVerificaCampanas(psNroContrato, gdFecSis, 3, lbNrRenovacion)
    If lrValida!nAplicaCence = 0 Then
        AXDesCon.TasaEfectivaMensual = vTasaInteres
    End If
   'JOEP20210921 Campana Prendario
   
    ' Fecha de Vencimiento es feriado - OJO
    lsFecVenTemp = vFecVencimiento
    Set loPigFunc = New COMDColocPig.DCOMColPFunciones
    
    If loPigFunc.dVerSiFeriado(lsFecVenTemp, lsmensaje) = True Then
        If Trim(lsmensaje) <> "" Then
            MsgBox lsmensaje, vbInformation, "Aviso"
            Exit Sub
        End If
        Do While True
            lsFecVenTemp = DateAdd("d", 1, lsFecVenTemp)
            fnDiaFer = fnDiaFer + 1 'RECO20141226 ERS170-2014
            If Not loPigFunc.dVerSiFeriado(lsFecVenTemp, lsmensaje) = True Then
                If Trim(lsmensaje) <> "" Then
                    MsgBox lsmensaje, vbInformation, "Aviso"
                    Exit Sub
                End If
                Exit Do
            End If
        Loop
        If lsFecVenTemp = gdFecSis Then
            vFecVencimiento = lsFecVenTemp
        Else
            fnDiaFer = 0 'RECO20141226 ERS170-2014
        End If
    End If
    Set loPigFunc = Nothing
    
    'lnDiasAtraso = DateDiff("d", Format(lrValida!dVenc, "dd/mm/yyyy"), Format(gdFecSis, "dd/mm/yyyy")) RIRO 20200406 COMENTADO
    lnDiasAtraso = lrValida!nDiasAtraso 'RIRO 20200406
    'vDiasAtrasoReal = vDiasAtraso
    Me.txtDiasAtraso = Val(lnDiasAtraso)
    
    Me.txtPgoAdelInt.Text = Format(fnPgoAdelInt, "#0.00") '*** PEAC 20161117
    
'    'vDiasAtrasoReal = vDiasAtraso
'    Set loCalculos = New NColPCalculos
'        lnDeuda = loCalculos.nCalculaDeudaPignoraticio(lrValida!nSaldo, lrValida!dVenc, lrValida!nTasacion, _
'                 fnTasaInteresVencido, fnTasaCustodiaVencida, fnTasaImpuesto, lrValida!nPrdEstado, fnTasaPreparacionRemate, gdFecSis)
'    Set loCalculos = Nothing
        
   
    'madm 20100120
    If Me.AXCodCta.Age <> "" Then
        'RECO20140623 ERS081-2014*******************************************
'        Select Case CInt(Me.AXCodCta.Age)
'            Case 1
'               fnVarCostoNotificacion = loParam.dObtieneColocParametro(3103)
'            Case 2
'               fnVarCostoNotificacion = loParam.dObtieneColocParametro(3104)
'            Case 3
'               fnVarCostoNotificacion = loParam.dObtieneColocParametro(3105)
'            Case 4
'               fnVarCostoNotificacion = loParam.dObtieneColocParametro(3106)
'            Case 5
'               fnVarCostoNotificacion = loParam.dObtieneColocParametro(3107)
'            Case 6
'               fnVarCostoNotificacion = loParam.dObtieneColocParametro(3108)
'            Case 7
'               fnVarCostoNotificacion = loParam.dObtieneColocParametro(3109)
'            Case 9
'               fnVarCostoNotificacion = loParam.dObtieneColocParametro(3111)
'            Case 10
'               fnVarCostoNotificacion = loParam.dObtieneColocParametro(3112)
'            Case 12
'               fnVarCostoNotificacion = loParam.dObtieneColocParametro(3113)
'            Case 13
'               fnVarCostoNotificacion = loParam.dObtieneColocParametro(3114)
'            Case 24
'               fnVarCostoNotificacion = loParam.dObtieneColocParametro(3115)
'            Case 25
'               fnVarCostoNotificacion = loParam.dObtieneColocParametro(3116)
'            Case 31
'             fnVarCostoNotificacion = loParam.dObtieneColocParametro(3117)
'        End Select
        'RECO20140722 ERS114-2014******************************************
        If AXCodCta.Age = "33" Then
            Dim oColPNotif As New COMDColocPig.DCOMColPActualizaBD
            Dim drNotif As ADODB.Recordset
            
            Set oColPNotif = New COMDColocPig.DCOMColPActualizaBD
            Set drNotif = New ADODB.Recordset
            
            Set drNotif = oColPNotif.DevuelveValorNotificacionCarNotMinka(AXCodCta.NroCuenta)
            If Not (drNotif.EOF And drNotif.BOF) Then
                fnVarCostoNotificacion = drNotif!nValor
            Else
                fnVarCostoNotificacion = loParam.dObtieneParamPignoCostoNotif("COSTO NOTIFIC", Me.AXCodCta.Age)
            End If
        Else
            fnVarCostoNotificacion = loParam.dObtieneParamPignoCostoNotif("COSTO NOTIFIC", Me.AXCodCta.Age)
        End If
        'RECO FIN**********************************************************
        'fnVarCostoNotificacion = loParam.dObtieneParamPignoCostoNotif("COSTO NOTIFIC", Me.AXCodCta.Age)
        'RECO FIN***********************************************************
   End If
    Set loParam = Nothing
    
    '------------------------------------------------------------------------
     ' ** Calcula deuda Total del Crédito
    fgCalculaDeuda lnDiasAtraso 'RIRO 20200406 Se añaadió lnDiasAtraso
    ' Muestro los Resultados
    'lblTotalDeuda.Caption = Format(CDbl(vDeuda), "#0.00")
    'LblMontoPagar.Caption = Format(CDbl(vDeuda) / (1 - gITF.gnITFPorcent), "#0.00")
    '*****************
    
    '******ITF*******
    lblTotalDeuda.Caption = Format(CDbl(vDeuda), "#0.00")
    If gITF.gbITFAplica Then
        If Not gITF.gbITFAsumidocreditos Then
            'Me.TxtITF = Format(gITF.fgITFCalculaImpuestoIncluido(TxtMontoTotal), "#0.00")
            'txtMontoPagar = Format(CDbl(Me.TxtMontoTotal) - CDbl(Me.TxtITF), "#0.00")
            'TxtMontoTotal.Text = Format(gITF.fgITFCalculaImpuestoIncluido(txtMontoPagar.Text) + Val(txtMontoPagar), "#0.00")
            'Me.TxtITF = gITF.fgITFCalculaImpuestoIncluido(txtMontoPagar.Text)
            LblITF.Caption = Format(gITF.fgITFCalculaImpuesto(lblTotalDeuda.Caption), "#0.00")
            '*** BRGO 20110908 ************************************************
            nRedondeoITF = fgDiferenciaRedondeoITF(CCur(Me.LblITF.Caption))
            If nRedondeoITF > 0 Then
               Me.LblITF.Caption = Format(CCur(Me.LblITF.Caption) - nRedondeoITF, "#,##0.00")
            End If
            '*** END BRGO
            LblMontoPagar = Format(CDbl(Me.lblTotalDeuda.Caption) + CDbl(LblITF.Caption), "#0.00")
        Else
            LblITF = Format(gITF.fgITFCalculaImpuesto(LblMontoPagar.Caption), "#0.00")
            '*** BRGO 20110908 ************************************************
            nRedondeoITF = fgDiferenciaRedondeoITF(CCur(Me.LblITF.Caption))
            If nRedondeoITF > 0 Then
               Me.LblITF.Caption = Format(CCur(Me.LblITF.Caption) - nRedondeoITF, "#,##0.00")
            End If
            '*** END BRGO
            LblMontoPagar = Format(CDbl(Me.lblTotalDeuda.Caption), "#0.00")
        End If
    Else
            LblITF = Format(0, "#0.00")
            LblMontoPagar = Format(CDbl(Me.lblTotalDeuda.Caption), "#0.00")
    End If
    '****************

'*** PEAC 20080701
txtCapital.Text = Format(vSaldoCapital, "#0.00") 'Format(LblMontoPagar.Caption - fnVarCostoNotificacion - vTasaInteresMoratorio - vInteresAdel - vTasaInteresVencido - LblItf.Caption, "#0.00")
txtMora.Text = Format(vInteresMoratorio, "#0.00")
txtInteres.Text = Format(vInteresAdel, "#0.00")
txtIntVen.Text = Format(vInteresVencido, "#0.00")
txtCostoCus.Text = Format(vCostoCustodiaMoratorio, "#0.00")
txtCostoNoti.Text = Format(fnVarCostoNotificacion, "#0.00")
txtCostoRemate.Text = Format(vCostoPreparacionRemate, "#0.00")
'*** FIN PEAC 20080701
    
    
    Set lrValida = Nothing
        
    cmdGrabar.Enabled = True
    cmdGrabar.SetFocus
    
        '************ firma madm 20091120 ----------------------------------------
         Set lafirma = New frmPersonaFirma
         Set ClsPersona = New COMDPersona.DCOMPersonas
        
         Set Rf = ClsPersona.BuscaCliente(gColPigFunciones.vcodper, BusquedaCodigo)
         
         If Not Rf.BOF And Not Rf.EOF Then
            If Rf!nPersPersoneria = 1 Then
           Call frmPersonaFirma.Inicio(Trim(gColPigFunciones.vcodper), Mid(gColPigFunciones.vcodper, 4, 2), False, False) 'MOD BY JATO 20210324
            End If
         End If
         Set Rf = Nothing
        '************ firma madm -------------------------------------------------

    AXCodCta.Enabled = False
    CmbForPag.Enabled = True 'CTI4 ERS0112020
    CmbForPag.ListIndex = IndiceListaCombo(CmbForPag, 1) 'CTI4 ERS0112020
    
Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox "Error: " & err.Number & " " & err.Description & vbCr & _
        "Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub


Private Sub fgCalculaDeuda(Optional ByVal nDiasAtr As Integer = -1)
'RIRO 20200406 ADD nDiasAtr

Dim loCalculos As COMNColoCPig.NCOMColPCalculos 'NColPCalculos
Dim lsmensaje As String
'vDiasAtraso = DateDiff("d", vFecVencimiento, gdFecSis)
'vdiasAtraso = DateDiff("d", Format(vFecVencimiento, "dd/mm/yyyy"), Format(gdFecSis, "dd/mm/yyyy")) RIRO 20200401 Comentado
vdiasAtraso = IIf(nDiasAtr < 0, DateDiff("d", Format(vFecVencimiento, "dd/mm/yyyy"), Format(gdFecSis, "dd/mm/yyyy")), nDiasAtr) 'RIRO 20200401
vDiasAtrasoReal = vdiasAtraso
If vdiasAtraso <= 0 Then

        'PEAC 20070813
        If gcCredAntiguo = "A" Then
            vInteresAdel = Round(0, 2)
        Else
            Set loCalculos = New COMNColoCPig.NCOMColPCalculos
                vDiasAdel = DateDiff("d", Format(vFecEstado, "dd/mm/yyyy"), Format(gdFecSis, "dd/mm/yyyy"))
                vDiasAdel = vDiasAdel - fnDiaFer 'RECO20141226 ERS170-2014
                '*** PEAC 20080806 ***************************
                'vInteresAdel = loCalculos.nCalculaInteresAdelantado(vSaldoCapital, vTasaInteres, vDiasAdel)
                 vInteresAdel = loCalculos.nCalculaInteresAlVencimiento(vSaldoCapital, vTasaInteres, vDiasAdel)
                '*** FIN PEAC ********************************
                
                '*** PEAC 20161221
                'vInteresAdel = Round(vInteresAdel, 2)
                vInteresAdel = vInteresAdel + fnIntPendSaldo 'nUltIntAPagar
                'vInteresAdel = Round(IIf((vInteresAdel - fnPgoAdelInt) < 0, vInteresAdel, vInteresAdel - fnPgoAdelInt), 2) '*** PEAC 20161221
                '*** FIN PEAC
    
            Set loCalculos = Nothing
        End If
    
    'end peac
    vdiasAtraso = 0
    vInteresVencido = 0
    vInteresMoratorio = 0
    vCostoCustodiaMoratorio = 0
    vImpuesto = 0
    fnVarGastoCorrespondencia = 0
Else
    Set loCalculos = New COMNColoCPig.NCOMColPCalculos
        'PEAC 20070813
        
        If gcCredAntiguo = "A" Then
            vInteresAdel = Round(vInteresAdel, 2)
        Else
            
            '*** PEAC 20170906 - mejora el calculo de dias de mora
            If Format(vFecEstado, "yyyymmdd") >= Format(vFecVencimiento, "yyyymmdd") Then
                vDiasAdel = DateDiff("d", Format(vFecEstado, "dd/mm/yyyy"), Format(gdFecSis, "dd/mm/yyyy"))
            Else
                
                'vDiasAdel = 30 'PEAC 20200808 - se igualo la lógica al proceso de renovacion
                'vDiasAdel = DateDiff("d", Format(vFecEstado, "dd/mm/yyyy"), Format(vFecVencimiento, "dd/mm/yyyy")) 'APRI
                
                'PEAC 20200808 - se igualo la lógica al proceso de renovacion
                vDiasAdel = DateDiff("d", Format(vFecEstado, "dd/mm/yyyy"), Format(vFecVencimiento, "dd/mm/yyyy"))
                
            End If
            
'            If Format(vFecEstado, "dd/mm/yyyy") > Format(vFecVencimiento, "dd/mm/yyyy") Then
'                vDiasAdel = 30
'            Else
'                vDiasAdel = DateDiff("d", Format(vFecEstado, "dd/mm/yyyy"), Format(vFecVencimiento, "dd/mm/yyyy"))
'            End If
            '*** FIN PEAC

            '*** PEAC 20080806 *********************************
            'vInteresAdel = loCalculos.nCalculaInteresAdelantado(vSaldoCapital, vTasaInteres, vDiasAdel)
             vInteresAdel = loCalculos.nCalculaInteresAlVencimiento(vSaldoCapital, vTasaInteres, vDiasAdel)
            '*** FIN PEAC **************************************
            
            '*** PEAC 20161221
            'vInteresAdel = Round(vInteresAdel, 2)
            'vInteresAdel = Round(vInteresAdel - fnPgoAdelInt, 2) '*** PEAC 20161221
            
            vInteresAdel = vInteresAdel + fnIntPendSaldo
            vInteresAdel = Round(IIf((vInteresAdel - fnPgoAdelInt) < 0, vInteresAdel, vInteresAdel - fnPgoAdelInt), 2) '*** PEAC 20161221
            
            '*** FIN PEAC
            
        End If
    
        'If nPagoIntVenMor > 0 Then
        If Format(vFecEstado, "dd/mm/yyyy") = Format(gdFecSis, "dd/mm/yyyy") Then 'APRI
            vInteresVencido = Round(0, 2)
            vInteresMoratorio = Round(0, 2)
        Else
        
            If Format(vFecEstado, "yyyymmdd") >= Format(vFecVencimiento, "yyyymmdd") Then
                vdiasAtraso = vDiasAdel
            End If
        
            vInteresVencido = loCalculos.nCalculaInteresMoratorio(vSaldoCapital, vTasaInteresVencido, vdiasAtraso, vInteresAdel)
            vInteresVencido = Round(vInteresVencido, 2)
        
            vInteresMoratorio = loCalculos.nCalculaInteresMoratorio(vSaldoCapital, vTasaInteresMoratorio, vdiasAtraso)
            vInteresMoratorio = Round(vInteresMoratorio, 2)
        End If
        
        vCostoCustodiaMoratorio = loCalculos.nCalculaCostoCustodiaMoratorio(vValorTasacion, fnTasaCustodiaVencida, vdiasAtraso)
        vCostoCustodiaMoratorio = Round(vCostoCustodiaMoratorio, 2)
        
        vImpuesto = (vInteresVencido + vCostoCustodiaMoratorio + vInteresMoratorio) * fnTasaImpuesto
        vImpuesto = Round(vImpuesto, 2)
        fnVarGastoCorrespondencia = loCalculos.nCalculaGastosCorrespondencia(AXCodCta.NroCuenta, lsmensaje)
        
        If Trim(lsmensaje) <> "" Then
             MsgBox lsmensaje, vbInformation, "Aviso"
             Exit Sub
        End If
    Set loCalculos = Nothing
End If
vCostoPreparacionRemate = 0

'*** PEAC 20080515
'Modificado por DAOR 20070714, Se esta para remate y además estado en el ultimo remate fue no vendido=2
If vEstado = gColPEstPRema And fnVarEstUltProcRem = 2 Then  ' Si esta en via de Remate
    vCostoPreparacionRemate = fnTasaPreparacionRemate * vValorTasacion
    vCostoPreparacionRemate = Round(vCostoPreparacionRemate, 2)
End If

'**PEAC 20070926 *********************************************
'If vEstado <> gColPEstPRema Then  ' Si no esta en via de Remate
'    fnVarCostoNotificacion = 0
'End If

'*** PEAC 20080515
'If gnNotifiAdju <> 1 And gnNotifiCob <> 0 Then
'    fnVarCostoNotificacion = 0
'End If

If gnNotifiAdju = 1 Then
    If gnNotifiCob = 1 Then
        fnVarCostoNotificacion = 0
    End If
Else
    fnVarCostoNotificacion = 0
End If


 vDeuda = vSaldoCapital + vInteresAdel + vInteresVencido + vCostoCustodiaMoratorio + vImpuesto + vCostoPreparacionRemate + fnVarGastoCorrespondencia + vInteresMoratorio + fnVarCostoNotificacion

'vDeuda = vSaldoCapital + vInteresAdel + vInteresVencido + vCostoCustodiaMoratorio + vImpuesto + vCostoPreparacionRemate + fnVarGastoCorrespondencia + vInteresMoratorio
End Sub

Private Sub AXCodCta_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then Call BuscaContrato(Me.AXCodCta.NroCuenta)
End Sub

Private Sub cmdBuscar_Click()
Dim loPers As COMDPersona.UCOMPersona 'UPersona
Dim lsPersCod As String, lsPersNombre As String
Dim lsEstados As String
Dim loPersContrato As COMDColocPig.DCOMColPContrato
Dim lrContratos As ADODB.Recordset
Dim loCuentas As COMDPersona.UCOMProdPersona

On Error GoTo ControlError

Set loPers = New COMDPersona.UCOMPersona
    Set loPers = frmBuscaPersona.Inicio
    If loPers Is Nothing Then Exit Sub

'--------- PEAC 20120216
    Limpiar
    cmdGrabar.Enabled = False
    AXCodCta.Enabled = True
    AXCodCta.SetFocus
'--------- FIN PEAC
    
    lsPersCod = loPers.sPersCod
    lsPersNombre = loPers.sPersNombre
Set loPers = Nothing

' Selecciona Estados
lsEstados = gColPEstDesem & "," & gColPEstVenci & "," & gColPEstPRema & "," & gColPEstRenov

If Trim(lsPersCod) <> "" Then
    Set loPersContrato = New COMDColocPig.DCOMColPContrato
        Set lrContratos = loPersContrato.dObtieneCredPigDePersona(lsPersCod, lsEstados, Mid(gsCodAge, 4, 2))
    Set loPersContrato = Nothing
End If

Set loCuentas = New COMDPersona.UCOMProdPersona
    'Set loCuentas = frmProdPersona.Inicio(lsPersNombre, lrContratos) RIRO 20130724 SEGUN ERS101-2013
    Set loCuentas = frmCuentasPersona.Inicio(lsPersNombre, lrContratos) 'RIRO 20130724 SEGUN ERS101-2013
    If loCuentas.sCtaCod <> "" Then
        AXCodCta.NroCuenta = Mid(loCuentas.sCtaCod, 1, 18)
        AXCodCta.SetFocusCuenta
    End If
Set loCuentas = Nothing
'ventana = 1
Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & err.Number & " " & err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "

End Sub

'Cancela el proceso actual e inicializa uno nuevo
Private Sub cmdCancelar_Click()
Limpiar
cmdGrabar.Enabled = False
AXCodCta.Enabled = True
CmbForPag.Enabled = False 'CTI4 ERS0112020
AXCodCta.SetFocus
End Sub

Private Sub cmdGrabar_Click()
'WIOR 20130301 ******************
Dim PersonaPago() As Variant
Dim nI As Integer
Dim fnCondicion As Integer
Dim regPersonaRealizaPago As Boolean
nI = 0
gnMovNro = 0
'WIOR ***************************
'On Error GoTo ControlError
Dim loContFunct As COMNContabilidad.NCOMContFunciones
Dim loGrabarCan As COMNColoCPig.NCOMColPContrato 'NColPContrato
Dim loImprime As COMNColoCPig.NCOMColPImpre 'NColPImpre
Dim loPrevio As previo.clsprevio
Dim loMov As COMDMov.DCOMMov

Dim lsMovNro As String
Dim lsFechaHoraGrab As String
Dim lsCuenta As String

Dim lnSaldoCap As Currency, lnInteresComp As Currency, lnImpuesto As Currency
Dim lnCostoTasacion As Currency, lnCostoCustodia As Currency
Dim lnMontoTransaccion As Currency

Dim lsCadImprimir As String
Dim lsBoletaCargo  As String 'CTI4 ERS0112020
Dim MatDatosAho(14) As String 'CTI4 ERS0112020
Dim lsNombreClienteCargoCta As String 'CTI4 ERS0112020
Dim lsNombreCliente As String


Dim loLavDinero As frmMovLavDinero
Set loLavDinero = New frmMovLavDinero
Dim objCmPr As COMDColocPig.DCOMColPContrato 'JOEP20210921 Campana Prendario

lnMontoTransaccion = CCur(Me.lblTotalDeuda)
lsNombreCliente = AXDesCon.listaClientes.ListItems(1).ListSubItems.Item(1)
'WIOR 20121009 Clientes Observados *************************************
Dim oDPersona As COMDPersona.DCOMPersona
Dim rsPersonaCred As ADODB.Recordset
Dim rsPersona As ADODB.Recordset
Dim Cont As Integer
Set oDPersona = New COMDPersona.DCOMPersona

If Not ValidaFormaPago Then Exit Sub 'CTI4 ERS0112020

Set rsPersonaCred = oDPersona.ObtenerPersCuentaRelac(Trim(AXCodCta.NroCuenta), gColRelPersTitular)

If rsPersonaCred.RecordCount > 0 Then
    If Not (rsPersonaCred.EOF And rsPersonaCred.BOF) Then
        For Cont = 0 To rsPersonaCred.RecordCount - 1
            Call VerSiClienteActualizoAutorizoSusDatos(Trim(rsPersonaCred!cperscod), fnVarOpeCod) 'FRHU ERS077-2015 20151204 OBSERVACION
            'WIOR 20130301 **************
            ReDim Preserve PersonaPago(Cont, 1)
            PersonaPago(Cont, 0) = Trim(rsPersonaCred!cperscod)
            PersonaPago(Cont, 1) = Trim(rsPersonaCred!cPersNombre)
            'WIOR **********************
            Set rsPersona = oDPersona.ObtenerUltimaVisita(Trim(rsPersonaCred!cperscod))
            If rsPersona.RecordCount > 0 Then
                If Not (rsPersona.EOF And rsPersona.BOF) Then
                    If Trim(rsPersona!sUsual) = "3" Then
                    MsgBox PstaNombre(Trim(rsPersonaCred!cPersNombre), True) & "." & Chr(10) & "CLIENTE OBSERVADO: " & Trim(rsPersona!cVisObserva), vbInformation, "Aviso"
                        Call frmPersona.Inicio(Trim(rsPersonaCred!cperscod), PersonaActualiza)
                    End If
                End If
            End If
            'Call VerSiClienteActualizoAutorizoSusDatos(Trim(rsPersonaCred!cPersCod), fnVarOpeCod) 'FRHU ERS077-2015 20151204
            Set rsPersona = Nothing
            rsPersonaCred.MoveNext
        Next Cont
    End If
End If
'WIOR FIN ***************************************************************

'*** AMDO20130705 TI-ERS063-2013
            Dim oDPersonaAct As COMDPersona.DCOMPersona
            Set oDPersonaAct = New COMDPersona.DCOMPersona
                            If oDPersonaAct.VerificaExisteSolicitudDatos(gColPigFunciones.vcodper) Then
                                MsgBox Trim("SE SOLICITA DATOS DEL CLIENTE: " & lsNombreCliente) & "." & Chr(10), vbInformation, "Aviso"
                                Call frmActInfContacto.Inicio(gColPigFunciones.vcodper)
                            End If
'***END AMDO

If MsgBox(" Grabar Cancelacion de Contrato Pignoraticio ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
    cmdGrabar.Enabled = False
        
        Dim clsExo As New COMNCaptaServicios.NCOMCaptaServicios 'NCapServicios
        If Not clsExo.EsCuentaExoneradaLavadoDinero(AXCodCta.NroCuenta) Then
          Dim sPersLavDinero As String
          Dim nMontoLavDinero As Double, nTC As Double
          Dim clsLav As New COMNCaptaGenerales.NCOMCaptaDefinicion, nmoneda As Integer, nMonto As Double
    
            nMonto = CDbl(lblTotalDeuda.Caption)
            Set clsExo = Nothing
            sPersLavDinero = ""
            nMontoLavDinero = clsLav.GetCapParametro(gMonOpeLavDineroME)
            Set clsLav = Nothing
            nmoneda = gMonedaNacional
            If nmoneda = gMonedaNacional Then
                Dim clsTC As COMDConstSistema.NCOMTipoCambio
                Set clsTC = New COMDConstSistema.NCOMTipoCambio
                nTC = clsTC.EmiteTipoCambio(gdFecSis, TCFijoDia)
                Set clsTC = Nothing
            Else
                nTC = 1
            End If
            If nMonto >= Round(nMontoLavDinero * nTC, 2) Then
                'By Capi 1402208
                    Call IniciaLavDinero(loLavDinero)
                    sPersLavDinero = loLavDinero.Inicia(, , , , False, True, nMonto, AXCodCta.NroCuenta, Mid(Me.Caption, 15), True, "", , , , , nmoneda)
                    If loLavDinero.OrdPersLavDinero = "" Then Exit Sub
                'End
                'sPersLavDinero = IniciaLavDinero()
                'If sPersLavDinero = "" Then Exit Sub
            End If
         Else
            Set clsExo = Nothing
         End If
        'WIOR 20130301 ***SEGUN TI-ERS005-2013 ************************************************************
        If loLavDinero.OrdPersLavDinero = "Exit" Then
            Dim oPersonaSPR As UPersona_Cli
            Dim oPersonaU As COMDPersona.UCOMPersona
            Dim nTipoConBN As Integer
            Dim sConPersona As String
            Dim pbClienteReforzado As Boolean
            Dim rsAgeParam As Recordset
            Dim objCred As COMNCredito.NCOMCredito
            Dim lnMonto As Double, lnTC As Double
            Dim ObjTc As COMDConstSistema.NCOMTipoCambio
            
            
            Set oPersonaU = New COMDPersona.UCOMPersona
            Set oPersonaSPR = New UPersona_Cli
            
            regPersonaRealizaPago = False
            pbClienteReforzado = False
            fnCondicion = 0
            
            For nI = 0 To UBound(PersonaPago)
                oPersonaSPR.RecuperaPersona Trim(PersonaPago(nI, 0))
                                    
                If oPersonaSPR.Personeria = 1 Then
                    If oPersonaSPR.Nacionalidad <> "04028" Then
                        sConPersona = "Extranjera"
                        fnCondicion = 1
                        pbClienteReforzado = True
                        Exit For
                    ElseIf oPersonaSPR.Residencia <> 1 Then
                        sConPersona = "No Residente"
                        fnCondicion = 2
                        pbClienteReforzado = True
                        Exit For
                    ElseIf oPersonaSPR.RPeps = 1 Then
                        sConPersona = "PEPS"
                        fnCondicion = 4
                        pbClienteReforzado = True
                        Exit For
                    ElseIf oPersonaU.ValidaEnListaNegativaCondicion(IIf(Trim(oPersonaSPR.ObtenerDNI) = "", oPersonaSPR.ObtenerNumeroDoc(0), oPersonaSPR.ObtenerDNI), oPersonaSPR.ObtenerRUC, nTipoConBN, oPersonaSPR.NombreCompleto) Then
                        If nTipoConBN = 1 Or nTipoConBN = 3 Then
                            sConPersona = IIf(nTipoConBN = 1, "Negativa", "PEPS")
                            fnCondicion = IIf(nTipoConBN = 1, 3, 4)
                            pbClienteReforzado = True
                            Exit For
                        End If
                    End If
                Else
                    If oPersonaU.ValidaEnListaNegativaCondicion(oPersonaSPR.ObtenerDNI, oPersonaSPR.ObtenerRUC, nTipoConBN, oPersonaSPR.NombreCompleto) Then
                        If nTipoConBN = 1 Or nTipoConBN = 3 Then
                            sConPersona = IIf(nTipoConBN = 1, "Negativa", "PEPS")
                            fnCondicion = IIf(nTipoConBN = 1, 3, 4)
                            pbClienteReforzado = True
                            Exit For
                        End If
                    End If
                End If
            Next nI
            
            If pbClienteReforzado Then
                MsgBox "El Cliente: " & Trim(PersonaPago(nI, 1)) & " es un Cliente de Procedimiento Reforzado (Persona " & sConPersona & ")", vbInformation, "Aviso"
                frmPersRealizaOpeGeneral.Inicia fsVarOpeDesc & " (Persona " & sConPersona & ")", fnVarOpeCod
                regPersonaRealizaPago = frmPersRealizaOpeGeneral.PersRegistrar
                
                If Not regPersonaRealizaPago Then
                    MsgBox "Se va a proceder a Anular la Operación", vbInformation, "Aviso"
                    cmdGrabar.Enabled = True
                    Exit Sub
                End If
            Else
                fnCondicion = 0
                lnMonto = nMonto
                pbClienteReforzado = False
                
                Set ObjTc = New COMDConstSistema.NCOMTipoCambio
                lnTC = ObjTc.EmiteTipoCambio(gdFecSis, TCFijoDia)
                Set ObjTc = Nothing
            
            
                Set objCred = New COMNCredito.NCOMCredito
                Set rsAgeParam = objCred.obtieneCredPagoCuotasAgeParam(gsCodAge)
                Set objCred = Nothing
                
                If Mid(AXCodCta.NroCuenta, 9, 1) = 2 Then
                    lnMonto = Round(lnMonto * lnTC, 2)
                End If
            
                If Not (rsAgeParam.EOF And rsAgeParam.BOF) Then
                    If lnMonto >= rsAgeParam!nMontoMin And lnMonto <= rsAgeParam!nMontoMax Then
                        frmPersRealizaOpeGeneral.Inicia fsVarOpeDesc, fnVarOpeCod
                        regPersonaRealizaPago = frmPersRealizaOpeGeneral.PersRegistrar
                        If Not regPersonaRealizaPago Then
                            MsgBox "Se va a proceder a Anular la Operación", vbInformation, "Aviso"
                            cmdGrabar.Enabled = True
                            Exit Sub
                        End If
                    End If
                End If
                
            End If
        End If
        'WIOR FIN ***************************************************************
        
        'Genera el Mov Nro
        Set loContFunct = New COMNContabilidad.NCOMContFunciones
            lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        Set loContFunct = Nothing
        
        Set loGrabarCan = New COMNColoCPig.NCOMColPContrato 'CTI4 ERS0112020
        
        'CTI4 ERS0112020
        Select Case CInt(Trim(Right(CmbForPag.Text, 10)))
            Case gColocTipoPagoEfectivo
                fnVarOpeCod = gColPOpeCancelacEFE
            Case gColocTipoPagoVoucher
                fnVarOpeCod = gColPOpeCancelVoucher
            Case gColocTipoPagoCargoCta
                fnVarOpeCod = gColPOpeCancelCargoCta
        End Select
        If CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoCargoCta Then lsNombreClienteCargoCta = PstaNombre(loGrabarCan.ObtieneNombreTitularCargoCta(txtCuentaCargo.NroCuenta))
        'end CTI4

        lsFechaHoraGrab = fgFechaHoraGrab(lsMovNro)
        Set loGrabarCan = New COMNColoCPig.NCOMColPContrato
            'Grabar Cancelacion Pignoraticio
            Call loGrabarCan.nCancelacionCredPignoraticio(AXCodCta.NroCuenta, vNewSaldoCapital, lsFechaHoraGrab, _
                 lsMovNro, lnMontoTransaccion, vSaldoCapital, vInteresVencido, _
                  vCostoCustodiaMoratorio, vCostoPreparacionRemate, vImpuesto, _
                  vDiasAtrasoReal, fnVarDiasCambCart, vValorTasacion, fnVarOpeCod, fsVarOpeDesc, fsVarPersCodCMAC, _
                  gITF.gbITFAplica, gITF.gbITFAsumidocreditos, CCur(Me.LblITF), False, Val(fnVarGastoCorrespondencia), vInteresMoratorio, fnVarCostoNotificacion, _
                  vInteresAdel, gnMovNro, , , CInt(Trim(Right(CmbForPag.Text, 10))), nMovNroRVD, nMovNroRVDPend, txtCuentaCargo.NroCuenta, MatDatosAho) 'WIOR 20130301 agrego gnMovNro '***peac 20071220 "fnVarCostoNotificacion"
                  '***PEAC 20080122 "vInteresAdel"
            
        Set loGrabarCan = Nothing
        '*** BRGO 20110906 ***************************
        If gITF.gbITFAplica Then
            Set loMov = New COMDMov.DCOMMov
            Call loMov.InsertaMovRedondeoITF(lsMovNro, 1, CCur(Me.LblITF) + nRedondeoITF, CCur(Me.LblITF))
            Set loMov = Nothing
        End If
        '*** BRGO
        
         'JOEP20210921 campana prendario
    If lblCampRetenPrend.Visible = True Then
        Set objCmPr = New COMDColocPig.DCOMColPContrato
        Call objCmPr.CampPrenRegCampCred(Trim(AXCodCta.NroCuenta), txtCampRete.Text, "Cancelacion", AXDesCon.TasaEfectivaMensual, 0, 0, 0, lsMovNro, 5, 4)
        Set objCmPr = Nothing
    End If
    'JOEP20210921 campana prendario
        
        'ADD JHCU 09-07-2020 REVERSIÓN PIGNORATICIO
         Set loGrabarCan = New COMNColoCPig.NCOMColPContrato
         Call loGrabarCan.ReversionReprogramacion(lsMovNro, AXCodCta.NroCuenta)
         Set loGrabarCan = Nothing
        'FIN JHCU 09-07-2020
        'Impresión
        Set loImprime = New COMNColoCPig.NCOMColPImpre
'            lsCadImprimir = loImprime.nPrintReciboCancelacion(gsNomAge, lsFechaHoraGrab, AXCodCta.NroCuenta, lsNombreCliente, _
'                Format(AXDesCon.FechaPrestamo, "mm/dd/yyyy"), vdiasAtraso, CCur(AXDesCon.prestamo), vSaldoCapital, _
'                vInteresAdel, vInteresVencido, vImpuesto, vCostoCustodiaMoratorio, vCostoPreparacionRemate, _
'                lnMontoTransaccion, vNewSaldoCapital, vTasaInteres, gsCodUser, fnVarCostoNotificacion, fsVarNombreCMAC, "", Val(fnVarGastoCorrespondencia), _
'                CDbl(LblItf.Caption), gImpresora, vInteresMoratorio, gbImpTMU)

            lsCadImprimir = loImprime.nPrintReciboCancelacion(gsNomAge, lsFechaHoraGrab, AXCodCta.NroCuenta, lsNombreCliente, _
                Format(AXDesCon.FechaPrestamo, "dd/MM/yyyy"), vdiasAtraso, CCur(AXDesCon.prestamo), vSaldoCapital, _
                vInteresAdel, vInteresVencido, vImpuesto, vCostoCustodiaMoratorio, vCostoPreparacionRemate, _
                lnMontoTransaccion, vNewSaldoCapital, vTasaInteres, gsCodUser, fnVarCostoNotificacion, fsVarNombreCMAC, "", Val(fnVarGastoCorrespondencia), _
                CDbl(LblITF.Caption), gImpresora, vInteresMoratorio, gbImpTMU)
  
        'CTI4 ERS0112020
        If CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoCargoCta Then
            lsBoletaCargo = loImprime.ImprimeBoletaAhorro("RETIRO AHORROS", "CARGO A CUENTA POR CANC. PIGNO.", "", CStr(lnMontoTransaccion + Me.LblITF), lsNombreClienteCargoCta, txtCuentaCargo.NroCuenta, "", CDbl(MatDatosAho(10)), CDbl(MatDatosAho(3)), "", 1, CDbl(MatDatosAho(11)), , , , , , , , , , gdFecSis, gsNomAge, gsCodUser, sLpt, gsCodCMAC, 0, , , gbImpTMU)
        End If
        'END CTI4

        Set loImprime = Nothing
        Set loPrevio = New previo.clsprevio
             loPrevio.PrintSpool sLpt, lsCadImprimir, False, 22
            Do While True
                If MsgBox("Reimprimir Recibo de Cancelación ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
                                                                                        
                    loPrevio.PrintSpool sLpt, lsCadImprimir, False, 22
                Else
                    Set loPrevio = Nothing
                    Exit Do
                End If
            Loop
        Set loLavDinero = Nothing
        
        'CTI4 ERS0112020
        If Trim(lsBoletaCargo) <> "" Then
            Set loPrevio = New previo.clsprevio
            loPrevio.PrintSpool sLpt, lsBoletaCargo, False, 22
            
            Do While True
                If MsgBox("Desea reimprimir boleta del cargo a cuenta?", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
                    Set loPrevio = New previo.clsprevio
                        loPrevio.PrintSpool sLpt, lsBoletaCargo, False
                    Set loPrevio = Nothing
                Else
                    Exit Do
                End If
            Loop
        End If
        'END CTI4 ERS0112020
        'WIOR 20130301 ************************************************************
        If regPersonaRealizaPago And gnMovNro > 0 Then
            frmPersRealizaOpeGeneral.InsertaPersonasOperacion gnMovNro, Trim(AXCodCta.NroCuenta), fnCondicion
            regPersonaRealizaPago = False
        End If
        'WIOR FIN *****************************************************************
        'CTI4 ERS0112020
        If CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoCargoCta Then
            Dim oMovOperacion As COMDMov.DCOMMov
            Dim nMovNroOperacion As Long
            Dim rsCli As New ADODB.Recordset
            Dim clsCli As New COMNCaptaGenerales.NCOMCaptaGenerales
            Dim oSolicitud As New COMDCaptaGenerales.DCOMCaptaGenerales
            Set oMovOperacion = New COMDMov.DCOMMov
            nMovNroOperacion = oMovOperacion.GetnMovNro(lsMovNro)

            loVistoElectronico.RegistraVistoElectronico nMovNroOperacion, , gsCodUser, nMovNroOperacion

            If nRespuesta = 2 Then
                Set rsCli = clsCli.GetPersonaCuenta(txtCuentaCargo.NroCuenta, gCapRelPersTitular)
                oSolicitud.ActualizarCapAutSinTarjetaVisto_nMovNro gsCodUser, gsCodAge, txtCuentaCargo.NroCuenta, rsCli!cperscod, nMovNroOperacion, CStr(gAhoCargoCtaCancelaPigno)
            End If
            Set oMovOperacion = Nothing
            nRespuesta = 0
        End If
        'CTI4 end
        'INICIO JHCU ENCUESTA 16-10-2019
         Encuestas gsCodUser, gsCodAge, "ERS0292019", fnVarOpeCod
        'FIN
                Limpiar
        
        
        AXCodCta.Enabled = True
        AXCodCta.SetFocus
        
        
Else
    MsgBox " Grabación cancelada ", vbInformation, " Aviso "
End If
Exit Sub


ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & err.Number & " " & err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub
Private Sub IniciaLavDinero(poLavDinero As frmMovLavDinero)
Dim i As Long
Dim nRelacion As CaptacRelacPersona
Dim sPersCod As String, sNombre As String, sDocId As String, sDireccion As String
Dim nMonto As Double, nPersoneria As Integer
Dim sCuenta As String
'For i = 1 To grdCliente.Rows - 1
    'nRelacion = CLng(Trim(Right(grdCliente.TextMatrix(i, 3), 4)))
    nPersoneria = gPersonaNat
    If nPersoneria = gPersonaNat Then
        'If nRelacion = gCapRelPersApoderado Or nRelacion = gCapRelPersTitular Then
            poLavDinero.TitPersLavDinero = AXDesCon.listaClientes.ListItems(1).Text
            poLavDinero.TitPersLavDineroNom = AXDesCon.listaClientes.ListItems(1).SubItems(1)
            poLavDinero.TitPersLavDineroDir = AXDesCon.listaClientes.ListItems(1).SubItems(2)
            poLavDinero.TitPersLavDineroDoc = AXDesCon.listaClientes.ListItems(1).SubItems(7)
         '   Exit For
       ' End If
    Else
        'If nRelacion = gCapRelPersTitular Then
            poLavDinero.TitPersLavDinero = AXDesCon.listaClientes.ListItems(1).Text
            poLavDinero.TitPersLavDineroNom = AXDesCon.listaClientes.ListItems(1).SubItems(1)
            poLavDinero.TitPersLavDineroDir = AXDesCon.listaClientes.ListItems(1).SubItems(2)
            poLavDinero.TitPersLavDineroDoc = AXDesCon.listaClientes.ListItems(1).SubItems(9)
          '  Exit For
        'End If
    End If
'Next i
nMonto = CDbl((lblTotalDeuda.Caption))
sCuenta = AXCodCta.NroCuenta
'If sPersCodCMAC <> "" Then
'    IniciaLavDinero = frmMovLavDinero.Inicia(sPersCod, sNombre, sDireccion, sDocId, False, True, nmonto, sCuenta, sOperacion, , sTipoCuenta)
'Else
    'IniciaLavDinero = frmMovLavDinero.Inicia(sPersCod, sNombre, sDireccion, sDocId, False, True, nMonto, sCuenta, CStr(fnVarOpeCod), , gMonedaNacional)
'End If
End Sub


'Finaliza el formulario actual
Private Sub cmdsalir_Click()
    Unload Me
End Sub

'Inicializa el formulario actual
Private Sub Form_Load()
'ventana = 0
    Call CargaControles 'CTI4 ERS0112020
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 And AXCodCta.Enabled = True Then 'F12
        Dim sCuenta As String
        sCuenta = frmValTarCodAnt.Inicia(gColConsuPrendario, False)
        If sCuenta <> "" Then
            AXCodCta.NroCuenta = sCuenta
            AXCodCta.SetFocusCuenta
        End If
   ElseIf KeyCode = 13 And AXCodCta.EnabledCta And AXCodCta.Age <> "" And Trim(AXCodCta.Cuenta) = "" Then
                AXCodCta.SetFocusCuenta
                 Exit Sub
 
    End If
End Sub

Private Sub CargaParametros()
Dim loParam As COMDColocPig.DCOMColPCalculos
Set loParam = New COMDColocPig.DCOMColPCalculos
    fnTasaCustodiaVencida = loParam.dObtieneColocParametro(gConsColPTasaCustodiaVencida)
    fnTasaPreparacionRemate = loParam.dObtieneColocParametro(gConsColPTasaPreparaRemate)
    fnTasaImpuesto = loParam.dObtieneColocParametro(gConsColPTasaImpuesto)
    fnVarDiasCambCart = loParam.dObtieneColocParametro(gConsColPDiasCambioCartera)
    
    'fnVarCostoNotificacion = loParam.dObtieneColocParametro(3103) 'ARCV 14-03-2007
    'madm 20091204 ---------------------------------------------------------
    If Me.AXCodCta.Age <> "" Then
        Select Case CInt(Me.AXCodCta.Age)
            Case 1
               fnVarCostoNotificacion = loParam.dObtieneColocParametro(3103)
            Case 2
               fnVarCostoNotificacion = loParam.dObtieneColocParametro(3104)
            Case 3
               fnVarCostoNotificacion = loParam.dObtieneColocParametro(3105)
            Case 4
               fnVarCostoNotificacion = loParam.dObtieneColocParametro(3106)
            Case 5
               fnVarCostoNotificacion = loParam.dObtieneColocParametro(3107)
            Case 6
               fnVarCostoNotificacion = loParam.dObtieneColocParametro(3108)
            Case 7
               fnVarCostoNotificacion = loParam.dObtieneColocParametro(3109)
            Case 9
               fnVarCostoNotificacion = loParam.dObtieneColocParametro(3111)
            Case 10
               fnVarCostoNotificacion = loParam.dObtieneColocParametro(3112)
            Case 12
               fnVarCostoNotificacion = loParam.dObtieneColocParametro(3113)
            Case 13
               fnVarCostoNotificacion = loParam.dObtieneColocParametro(3114)
            Case 24
               fnVarCostoNotificacion = loParam.dObtieneColocParametro(3115)
            Case 25
               fnVarCostoNotificacion = loParam.dObtieneColocParametro(3116)
            Case 31
               fnVarCostoNotificacion = loParam.dObtieneColocParametro(3117)
        End Select
   End If
    '************** end madm ************************
    
    
Set loParam = Nothing
End Sub
'CTI4 ERS0112020 *****************
Private Sub CmbForPag_Click()
    EstadoFormaPago IIf(CmbForPag.ListIndex = -1, -1, CInt(Trim(Right(IIf(CmbForPag.Text = "", "-1", CmbForPag.Text), 10))))
    If CmbForPag.ListIndex <> -1 Then
        If CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoVoucher Then
     
            Dim oformVou As New frmCapRegVouDepBus
            Dim lnTipMot As Integer
            Dim sGlosa As String
            Dim sIF As String
            Dim sVaucher As String
            Dim sPersCod As String
            Dim sNombre As String
            Dim sDireccion As String
            Dim sDocumento As String
                        
            lnTipMot = 19 ' Cancelacion Credito Pignoraticio
            oformVou.iniciarFormularioDeposito CInt(Mid(AXCodCta.NroCuenta, 9, 1)), lnTipMot, sGlosa, sIF, sVaucher, nMontoVoucher, sPersCod, nMovNroRVD, nMovNroRVDPend, sNombre, sDireccion, sDocumento, AXCodCta.NroCuenta
            If Len(sVaucher) = 0 Then Exit Sub
            LblNumDoc.Caption = Trim(Mid(sVaucher, 1, Len(sVaucher) - 10))
            cmdGrabar.Enabled = True
        ElseIf CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoCargoCta Then
            Dim sCuenta As String
            
            sCuenta = frmATMCargaCuentas.RecuperaCuenta(CStr(gAhoCargoCtaCancelaPigno), sNumTarj, 232, False)
            If Val(Mid(sCuenta, 6, 3)) <> "232" And sCuenta <> "" Then
                MsgBox "Esta operación no le corresponde a este producto.", vbOKOnly + vbInformation, App.Title
                Exit Sub
            End If
                                  
                                             
                                        
                   
            If Len(sCuenta) = 18 Then
                If CInt(Mid(sCuenta, 9, 1)) <> CInt(Mid(AXCodCta.NroCuenta, 9, 1)) Then
                    MsgBox "La cuenta de ahorro no tiene el mismo tipo de moneda que la cuenta a amortizar.", vbOKOnly + vbInformation, App.Title
                End If
            End If
            If Len(sCuenta) = 0 Then txtCuentaCargo.SetFocusAge: Exit Sub
            txtCuentaCargo.NroCuenta = sCuenta
            txtCuentaCargo.Enabled = False
            AsignaValorITF
            cmdGrabar.Enabled = True
            cmdGrabar.SetFocus
        End If
    End If
End Sub
Private Sub EstadoFormaPago(ByVal nFormaPago As Integer)
    LblNumDoc.Caption = ""
    txtCuentaCargo.NroCuenta = ""
    Select Case nFormaPago
        Case -1
            LblNumDoc.Visible = False
            lblNroDocumento.Visible = False
            txtCuentaCargo.Visible = False
            cmdGrabar.Enabled = True
        Case gColocTipoPagoEfectivo
            txtCuentaCargo.Visible = False
            LblNumDoc.Visible = False
            lblNroDocumento.Visible = False
            lblNroDocumento.Visible = False
            cmdGrabar.Enabled = True
        Case gColocTipoPagoCargoCta
            LblNumDoc.Visible = False
            lblNroDocumento.Visible = False
            txtCuentaCargo.Visible = True
            txtCuentaCargo.Enabled = True
            txtCuentaCargo.CMAC = gsCodCMAC
            txtCuentaCargo.Prod = Trim(Str(gCapAhorros))
            cmdGrabar.Enabled = False
        Case gColocTipoPagoVoucher
            LblNumDoc.Visible = True
            lblNroDocumento.Visible = True
            txtCuentaCargo.Visible = False
            cmdGrabar.Enabled = False
    End Select
End Sub
Private Function ValidaFormaPago() As Boolean
Dim clsCap As New COMNCaptaGenerales.NCOMCaptaMovimiento
    ValidaFormaPago = False
    If CmbForPag.ListIndex = -1 Then
        MsgBox "No se ha seleccionado la forma de pago. Verifique.", vbInformation, "¡Aviso!"
        CmbForPag.SetFocus
        Exit Function
    End If
    If CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoVoucher And Len(Trim(LblNumDoc.Caption)) = 0 Then
        MsgBox "No se ha seleccionado el voucher correctamente. Verifique.", vbInformation, "¡Aviso!"
        CmbForPag.SetFocus
        Exit Function
    End If
    If CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoVoucher And Len(Trim(LblNumDoc.Caption)) > 0 _
        And CCur(LblMontoPagar.Caption) <> CCur(nMontoVoucher) Then
        MsgBox "El Monto de Transacción debe ser igual al Monto Total. Verifique.", vbInformation, "¡Aviso!"
        CmbForPag.SetFocus
        Exit Function
    End If
    
    If CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoCargoCta And Len(txtCuentaCargo.NroCuenta) <> 18 Then
        MsgBox "Favor de ingresar la cuenta de ahorros a debitar", vbInformation, "¡Aviso!"
        CmbForPag.SetFocus
        Exit Function
    End If
        
    If CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoCargoCta Then
        If Not clsCap.ValidaSaldoCuenta(txtCuentaCargo.NroCuenta, CDbl(LblMontoPagar.Caption)) Then
            MsgBox "Cuenta a debitar NO posee saldo suficiente o NO está ACTIVA", vbInformation, "¡Aviso!"
            Exit Function
        End If
    End If
    
    ValidaFormaPago = True
End Function
Private Sub txtCuentaCargo_KeyPress(KeyAscii As Integer)
    Dim sMsg As String
    If KeyAscii = 13 Then sMsg = ValidaCuentaACargo(txtCuentaCargo.NroCuenta)
    If Len(sMsg) > 0 Then
        MsgBox sMsg, vbInformation, "¡Aviso!"
        txtCuentaCargo.SetFocus
        Exit Sub
    End If
    If Len(txtCuentaCargo.NroCuenta) = 18 Then
        If CInt(Mid(txtCuentaCargo.NroCuenta, 9, 1)) <> CInt(Mid(AXCodCta.NroCuenta, 9, 1)) Then
            MsgBox "La cuenta de ahorro no tiene el mismo tipo de moneda que la cuenta a amortizar.", vbOKOnly + vbInformation, App.Title
        End If
    End If
    ObtieneDatosCuenta txtCuentaCargo.NroCuenta
End Sub
Private Function ValidaCuentaACargo(ByVal psCuenta As String) As String
    Dim clsCap As COMNCaptaGenerales.NCOMCaptaMovimiento
    Dim sMsg As String
    Set clsCap = New COMNCaptaGenerales.NCOMCaptaMovimiento
    sMsg = clsCap.ValidaCuentaOperacion(psCuenta)
    ValidaCuentaACargo = sMsg
End Function
Private Sub ObtieneDatosCuenta(ByVal psCuenta As String)
    Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales
    Dim rsCta As ADODB.Recordset
    Dim rsV As ADODB.Recordset
    Dim lnTpoPrograma As Integer
    Dim lsTieneTarj As String
    Dim lbVistoVal As Boolean
   
    Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
    Set rsV = New ADODB.Recordset
    Set rsCta = New ADODB.Recordset
    Set rsCta = clsMant.GetDatosCuenta(psCuenta)
    If Not (rsCta.EOF And rsCta.BOF) Then
        lnTpoPrograma = IIf(IsNull(rsCta("nTpoPrograma")), 0, rsCta("nTpoPrograma"))
        If sNumTarj = "" Then
            If Mid(psCuenta, 6, 3) = "232" And lnTpoPrograma <> 1 Then
                Set rsV = clsMant.ValidaTarjetizacion(psCuenta, lsTieneTarj)
                If rsV.RecordCount > 0 Then
                    Dim tipoCta As Integer
                    tipoCta = rsCta("nPrdCtaTpo")
                    If tipoCta = 0 Or tipoCta = 2 Then
                        Dim rsCli As New ADODB.Recordset
                        Dim clsCli As New COMNCaptaGenerales.NCOMCaptaGenerales
                        Dim oSolicitud As New COMDCaptaGenerales.DCOMCaptaGenerales
                        Dim bExitoSol As Integer
                        Set rsCli = clsCli.GetPersonaCuenta(psCuenta, gCapRelPersTitular)
                        nRespuesta = oSolicitud.SolicitarVistoAtencionSinTarjeta(gsCodUser, gsCodAge, psCuenta, rsCli!cperscod, CStr(gAhoCargoCtaCancelaPigno))
                    
                        If nRespuesta = 1 Then '1:Tiene Visto de atencion sin tarjeta pendiente de autorizar
                             MsgBox "La solicitud de atención sin tarjeta fue enviada. " & vbNewLine & "Comuníquese con el Coordinador o Jefe de Operaciones para la aprobación o rechazo de la misma", vbInformation, "Aviso"
                             Exit Sub
                        End If
                        If nRespuesta = 2 Then '2:Tiene visto aceptado
                            MsgBox "La solicitud de atención sin tarjeta fue Aprobada, proceda con la atención", vbInformation, "Aviso"
                        End If
                        If nRespuesta = 3 Then '3:Tiene visto rechazado
                           If MsgBox("La solicitud de atención sin tarjeta fue RECHAZADA. ¿Desea realizar una nueva solicitud?", vbInformation + vbYesNo, "Aviso") = vbYes Then
                                Call oSolicitud.RegistrarVistoDeUsuario(gsCodUser, gsCodAge, psCuenta, rsCli!cperscod, CStr(gAhoCargoCtaCancelaPigno))
                                Exit Sub
                            Else
                                Exit Sub
                            End If
                        End If
                        If nRespuesta = 4 Then '4:Se permite registrar la solicitud
                            Dim mensaje As String
                            If lsTieneTarj = "SI" Then
                                mensaje = "El Cliente posee tarjeta. Para continuar deberá registrar el Motivo de Autorización y comunicar al Coordinador o Jefe de Operaciones para su Aprobación. ¿Desea Continuar?"
                            Else
                                mensaje = "El Cliente NO posee tarjeta activa. Para continuar deberá registrar el Motivo de Autorización y comunicar al Coordinador o Jefe de Operaciones para su Aprobación. ¿Desea Continuar?"
                            End If
                        
                            If MsgBox(mensaje, vbInformation + vbYesNo, "Aviso") = vbYes Then
                        
                                bExitoSol = frmCapMotivoAutorizacion.Inicio(gsCodUser, gsCodAge, psCuenta, rsCli!cperscod, CStr(gAhoCargoCtaCancelaPigno))
                                If bExitoSol > 0 Then
                                    MsgBox "La solicitud de atención sin tarjeta fue enviada. " & vbNewLine & "Comuníquese con el Coordinador o Jefe de Operaciones para la aprobación o rechazo de la misma", vbInformation, "Aviso"
                                End If
                                Exit Sub
                            Else
                                Exit Sub
                            End If
                        End If
                        lbVistoVal = loVistoElectronico.Inicio(5, gAhoCargoCtaCancelaPigno)
                        If Not lbVistoVal Then
                            MsgBox "Visto Incorrecto por favor comunicar al supervisor de operaciones", vbInformation, "Mensaje del Sistema"
                            Exit Sub
                        End If
                    End If
                ElseIf lsTieneTarj = "NO" And rsV.RecordCount > 0 Then
                    If MsgBox("El Cliente debe solicitar su tarjeta para realizar las operaciones, si desea continuar con la operacion? ", vbInformation + vbYesNo, "Mensaje del Sistema") = vbYes Then 'add by marg ers 065-2017
                        lbVistoVal = loVistoElectronico.Inicio(5, gAhoCargoCtaCancelaPigno)
                        If Not lbVistoVal Then
                            MsgBox "Visto Incorrecto por favor comunicar al supervisor de operaciones", vbInformation, "Mensaje del Sistema" 'ADD BY MARG ERS065-2017
                            Exit Sub
                        End If
                    Else
                        Exit Sub
                    End If
                End If
            End If
        End If
        txtCuentaCargo.Enabled = False
        AsignaValorITF
        cmdGrabar.Enabled = True
        cmdGrabar.SetFocus
    End If
End Sub
Private Sub AsignaValorITF()
Dim loValida As COMNColoCPig.NCOMColPValida
Dim bEsMismoTitular As Boolean
    Set loValida = New COMNColoCPig.NCOMColPValida
    
    bEsMismoTitular = loValida.EsMismoTitulardeCuentaPignoYAhorro(txtCuentaCargo.NroCuenta, AXCodCta.NroCuenta)
    
    If gITF.gbITFAplica And Not bEsMismoTitular Then
        If Not gITF.gbITFAsumidocreditos Then
            LblITF.Caption = Format(gITF.fgITFCalculaImpuesto(lblTotalDeuda.Caption), "#0.00")
            nRedondeoITF = fgDiferenciaRedondeoITF(CCur(Me.LblITF.Caption))
            If nRedondeoITF > 0 Then
               Me.LblITF.Caption = Format(CCur(Me.LblITF.Caption) - nRedondeoITF, "#,##0.00")
            End If
            LblMontoPagar = Format(CDbl(Me.lblTotalDeuda.Caption) + CDbl(LblITF.Caption), "#0.00")
        Else
            LblITF = Format(gITF.fgITFCalculaImpuesto(LblMontoPagar.Caption), "#0.00")
            nRedondeoITF = fgDiferenciaRedondeoITF(CCur(Me.LblITF.Caption))
            If nRedondeoITF > 0 Then
               Me.LblITF.Caption = Format(CCur(Me.LblITF.Caption) - nRedondeoITF, "#,##0.00")
            End If
            LblMontoPagar = Format(CDbl(Me.lblTotalDeuda.Caption), "#0.00")
        End If
    Else
            LblITF = Format(0, "#0.00")
            LblMontoPagar = Format(CDbl(Me.lblTotalDeuda.Caption), "#0.00")
    End If
End Sub
Private Sub CargaControles()
Dim oCons As COMDConstantes.DCOMConstantes
Dim R As ADODB.Recordset

    On Error GoTo ERRORCargaControles
    Set oCons = New COMDConstantes.DCOMConstantes
    Set R = oCons.RecuperaConstantes(gColocTipoPago, , , 4)
    Set oCons = Nothing
    Call Llenar_Combo_con_Recordset(R, CmbForPag)
    Set loVistoElectronico = New frmVistoElectronico
    Exit Sub
ERRORCargaControles:
        MsgBox err.Description, vbCritical, "Aviso"
End Sub
'END CTI4


'JOEP20210916 campana prendario
Private Sub CampPrendVerificaCampanas(ByVal pcCtaCod As String, ByVal pdFechaSis As String, ByVal pnModulo As Integer, ByVal pnNrRenov As Integer)
    Dim oCampPrend As COMDColocPig.DCOMColPContrato
    Dim rsCampPrend As ADODB.Recordset
    Set oCampPrend = New COMDColocPig.DCOMColPContrato
    
    Set rsCampPrend = oCampPrend.CampPrendarioDesbCampa(pcCtaCod, pdFechaSis, pnModulo, pnNrRenov)
    If Not (rsCampPrend.BOF And rsCampPrend.EOF) Then
        lblCampRetenPrend.Caption = rsCampPrend!cResultado
        lblCampRetenPrend.Visible = True
        txtCampRete.Text = rsCampPrend!nCampana
    Else
        lblCampRetenPrend.Visible = False
        lblCampRetenPrend.Caption = ""
        txtCampRete.Text = 0#
    End If
    Set oCampPrend = Nothing
    RSClose oCampPrend
End Sub
'JOEP20210916 campana prendario
