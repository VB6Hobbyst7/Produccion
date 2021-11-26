VERSION 5.00
Begin VB.Form frmColPAmortizacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Crédito Pignoraticio - Amortización de Crédito"
   ClientHeight    =   7725
   ClientLeft      =   1935
   ClientTop       =   2385
   ClientWidth     =   7995
   Icon            =   "frmColPAmortizacion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7725
   ScaleWidth      =   7995
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   5640
      TabIndex        =   4
      Top             =   7215
      Width           =   975
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   360
      Left            =   4440
      TabIndex        =   3
      Top             =   7230
      Width           =   1035
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   6795
      TabIndex        =   5
      Top             =   7185
      Width           =   975
   End
   Begin VB.Frame fraContenedor 
      Height          =   6990
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   60
      Width           =   7785
      Begin VB.CommandButton cmdBuscar 
         Height          =   345
         Left            =   7110
         Picture         =   "frmColPAmortizacion.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Buscar ..."
         Top             =   270
         Width           =   420
      End
      Begin VB.Frame fraContenedor 
         Height          =   1425
         Index           =   5
         Left            =   120
         TabIndex        =   12
         Top             =   5400
         Width           =   7425
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
            Left            =   3600
            TabIndex        =   34
            Top             =   240
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
            Left            =   1320
            TabIndex        =   33
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
            Left            =   1320
            TabIndex        =   32
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
            Left            =   1320
            TabIndex        =   31
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox TxtITF 
            Alignment       =   1  'Right Justify
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
            Height          =   285
            Left            =   6000
            TabIndex        =   28
            Top             =   585
            Width           =   1215
         End
         Begin VB.TextBox TxtMontoTotal 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   6000
            Locked          =   -1  'True
            TabIndex        =   27
            Top             =   930
            Width           =   1215
         End
         Begin VB.ComboBox cboPlazoNuevo 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmColPAmortizacion.frx":040C
            Left            =   1440
            List            =   "frmColPAmortizacion.frx":0413
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   -120
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox txtMontoPagar 
            Alignment       =   1  'Right Justify
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
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   6000
            MaxLength       =   9
            TabIndex        =   2
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label lblCostoNoti 
            AutoSize        =   -1  'True
            Caption         =   "Costo Notif."
            Height          =   195
            Index           =   3
            Left            =   2640
            TabIndex        =   38
            Top             =   240
            Width           =   825
         End
         Begin VB.Label lblCostoCus 
            AutoSize        =   -1  'True
            Caption         =   "Costo Custodia"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   37
            Top             =   960
            Width           =   1065
         End
         Begin VB.Label lblInteres 
            AutoSize        =   -1  'True
            Caption         =   "Interes"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   36
            Top             =   600
            Width           =   480
         End
         Begin VB.Label lblCapital 
            AutoSize        =   -1  'True
            Caption         =   "Capital"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   35
            Top             =   240
            Width           =   480
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "ITF"
            Height          =   195
            Left            =   5010
            TabIndex        =   30
            Top             =   675
            Width           =   240
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Monto Pagar"
            Height          =   195
            Left            =   5010
            TabIndex        =   29
            Top             =   990
            Width           =   915
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Plazo Nuevo "
            Height          =   255
            Index           =   15
            Left            =   240
            TabIndex        =   14
            Top             =   0
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Monto "
            Height          =   255
            Index           =   12
            Left            =   5010
            TabIndex        =   13
            Top             =   240
            Width           =   600
         End
      End
      Begin VB.Frame fraContenedor 
         Height          =   1005
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   4320
         Width           =   7485
         Begin VB.TextBox txtSaldoCapitalNuevo 
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
            Height          =   285
            Left            =   6360
            TabIndex        =   24
            Top             =   600
            Width           =   960
         End
         Begin VB.TextBox txtPlazoActual 
            Alignment       =   2  'Center
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
            Left            =   6720
            MaxLength       =   2
            TabIndex        =   22
            Top             =   240
            Width           =   555
         End
         Begin VB.TextBox txtDiasAtraso 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1440
            TabIndex        =   19
            Top             =   240
            Width           =   690
         End
         Begin VB.TextBox txtNroRenovacion 
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
            Left            =   1440
            TabIndex        =   18
            Top             =   600
            Width           =   690
         End
         Begin VB.TextBox txtMontoMinimoPagar 
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
            Left            =   3720
            TabIndex        =   8
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox txtTotalDeuda 
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
            Left            =   3720
            TabIndex        =   7
            Top             =   240
            Width           =   1245
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Nuevo Saldo"
            Height          =   255
            Index           =   14
            Left            =   5280
            TabIndex        =   25
            Top             =   615
            Width           =   960
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Plazo Actual"
            Height          =   255
            Index           =   11
            Left            =   5280
            TabIndex        =   23
            Top             =   285
            Width           =   960
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Dias Atraso"
            Height          =   210
            Index           =   7
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Width           =   960
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Nro Renovación"
            Height          =   330
            Index           =   8
            Left            =   120
            TabIndex        =   20
            Top             =   600
            Width           =   1200
         End
         Begin VB.Label lblMoneda 
            Height          =   255
            Left            =   2070
            TabIndex        =   15
            Top             =   210
            Width           =   255
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Mínimo a Pagar"
            Height          =   255
            Index           =   10
            Left            =   2520
            TabIndex        =   11
            Top             =   600
            Width           =   1170
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Total Deuda"
            Height          =   195
            Index           =   9
            Left            =   2520
            TabIndex        =   10
            Top             =   240
            Width           =   1185
         End
      End
      Begin SICMACT.ActXCodCta AXCodCta 
         Height          =   375
         Left            =   240
         TabIndex        =   0
         Top             =   240
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   661
         Texto           =   "Crédito"
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
      Begin SICMACT.ActXColPDesCon AXDesCon 
         Height          =   3495
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   6165
      End
   End
   Begin VB.Label lblMensaje 
      Caption         =   "<F12> Número Contrato Antiguo"
      Height          =   285
      Left            =   240
      TabIndex        =   16
      Top             =   7215
      Width           =   2655
   End
End
Attribute VB_Name = "frmColPAmortizacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************
'* AMORTIZACION  DE CONTRATO PIGNORATICIO
'Archivo:  frmColPAmortizacion.frm
'LAYG   :  15/06/2003.
'Resumen:  Nos permite amortizar un contrato afectando directamente al capital

Option Explicit

Dim fnVarOpeCod As Long
Dim fsVarOpeDesc As String
Dim fsVarPersCodCMAC As String
Dim fsVarNombreCMAC As String

Dim fnVarTasaPreparacionRemate As Double
Dim fnVarTasaImpuesto As Double
Dim fnVarTasaCustodia As Double
Dim fnVarTasaCustodiaVencida As Double

Dim fnVarTasaInteres As Double
Dim fnVarTasaInteresVencido As Double

Dim fnVarDiasCambCart As Double

Dim fnVarSaldoCap As Currency
Dim fnVarValorTasacion As Currency
Dim fnVarPlazo As Integer
Dim fdVarFecVencimiento As Date
Dim fnVarEstado As ColocEstado
Dim fnVarNroRenovacion As Integer
Dim fnVarFechaRenovacion As Date

Dim fnVarNewSaldoCap As Currency
Dim fnVarNewPlazo As Integer
Dim fsVarNewFecVencimiento As String
Dim fnVarCapitalPagado As Currency   ' Capital a Pagar
Dim fnVarFactor As Double
Dim fnVarInteresVencido As Currency
Dim fnVarInteres As Currency
Dim fnVarCostoCustodia As Currency
Dim fnVarCostoCustodiaVencida As Currency
Dim fnVarImpuesto As Currency
Dim fnVarCostoPreparacionRemate As Double

Dim fnVarDiasAtraso As Double
Dim vDiasAtrasoReal As Double
Dim vSumaCostoCustodia As Double
Dim fnVarDeuda As Currency

Dim fnVarMontoMinimo As Currency
Dim fnVarMontoAPagar As Currency

Dim fnVarCostoNotificacion As Currency '*** PEAC 20080515

Dim fsColocLineaCredPig As String ' PEAC 20070813
Dim vFecEstado As Date ' PEAC 20070813
Dim vDiasAdel As Integer, vInteresAdel As Double, vMontoCol As Double ' PEAC 20070813
Dim gcCredAntiguo As String  ' peac 20070923
Dim gnNotifiAdju As Integer  ' peac 20080515
Dim nRedondeoITF As Double  'BRGO 20110906


Public Sub Inicio(ByVal pnOpeCod As Long, ByVal psOpeDesc As String, _
        ByVal psPersCodCMAC As String, ByVal psNomCmac As String)

    fnVarOpeCod = pnOpeCod
    fsVarOpeDesc = psOpeDesc
    fsVarPersCodCMAC = psPersCodCMAC
    fsVarNombreCMAC = psNomCmac
    
    Select Case fnVarOpeCod
        Case gColPOpeAmortizEFE
            'txtDocumento.Visible = false
        Case gColPOpeAmortizCHQ
            'txtDocumento.Visible = True
    '    Case Else
    '        txtDocumento.Visible = False
    End Select
    CargaParametros
    Limpiar
    Me.Show 1

End Sub

Private Sub Limpiar()
    Me.AXCodCta.NroCuenta = fgIniciaAxCuentaPignoraticio
    AXDesCon.Limpiar
    txtDiasAtraso.Text = ""
    txtNroRenovacion.Text = ""
    txtTotalDeuda.Text = ""
    txtMontoMinimoPagar.Text = Format(0, "#0.00")
    TxtITF.Text = Format(0, "#0.00")
    TxtMontoTotal = Format(0, "#0.00")
    txtPlazoActual.Text = ""
    txtSaldoCapitalNuevo.Text = Format(0, "#0.00")
    cboPlazoNuevo.ListIndex = 0
    txtMontoPagar.Text = Format(0, "#0.00")
    fnVarCapitalPagado = 0
    fnVarNewSaldoCap = 0
    txtSaldoCapitalNuevo.Text = Format(0, "#0.00")
    nRedondeoITF = 0
End Sub

Private Sub AXCodCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call BuscaContrato(AXCodCta.NroCuenta)
End Sub

'Busca el contrato ingresado
Private Sub BuscaContrato(ByVal psNroContrato As String)
Dim lbok As Boolean
Dim lrValida As ADODB.Recordset
Dim loValContrato As COMNColoCPig.NCOMColPValida
Dim lnDeuda As Currency, lnMinimoPagar As Currency
Dim lnDiasAtraso  As Integer
Dim lsmensaje As String
'----- MADM 20091120 ---------------------
Dim lafirma As frmPersonaFirma
Dim ClsPersona As COMDPersona.DCOMPersonas
Dim Rf As ADODB.Recordset
'----- MADM ------------------------------
'On Error GoTo ControlError

    'Valida Contrato
    Set lrValida = New ADODB.Recordset
    Set loValContrato = New COMNColoCPig.NCOMColPValida
      '  Set lrValida = loValContrato.nValidaAmortizacionCredPignoraticio(psNroContrato, gdFecSis, 0)
    Set lrValida = loValContrato.nValidaAmortizacionCredPignoraticio(psNroContrato, gdFecSis, 0, gsCodUser, lsmensaje)
    If Trim(lsmensaje) <> "" Then
         MsgBox lsmensaje, vbInformation, "Aviso"
         Exit Sub
    End If
    Set loValContrato = Nothing
    
    
    If (lrValida Is Nothing) Then       ' Hubo un Error
        Limpiar
        Set lrValida = Nothing
        Exit Sub
    End If
    ' Asigna Valores a Variables del Form
          
Dim loParam As COMDColocPig.DCOMColPCalculos
Set loParam = New COMDColocPig.DCOMColPCalculos
        
    'madm 20091204 --------------------------------------------------------
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
        'end madm ----------------------------------------------------------
    'fnVarCostoNotificacion = loParam.dObtieneColocParametro(3103) '*** PEAC 20080515

Set loParam = Nothing
          
    fnVarPlazo = lrValida!nPlazo
    fnVarSaldoCap = Format(lrValida!nSaldo, "#0.00")
    fnVarValorTasacion = lrValida!nTasacion
    fnVarTasaInteresVencido = lrValida!nTasaIntVenc
    fnVarEstado = lrValida!nPrdEstado
    fnVarTasaInteres = lrValida!nTasaInteres
    fdVarFecVencimiento = Format(lrValida!dVenc, "dd/mm/yyyy")
    fnVarFechaRenovacion = lrValida!dRenovacion
    gcCredAntiguo = lrValida!cCredB 'PEAC 20071106
    
    vFecEstado = Format(lrValida!dPrdEstado, "dd/mm/yyyy") ' PEAC 20070813
    fnVarSaldoCap = lrValida!nMontoCol ' PEAC 20070813
    gnNotifiAdju = lrValida!nCodNotifiAdj 'PEAC 20080515
    
    fnVarNroRenovacion = lrValida!nNroRenov
    fnVarNewPlazo = lrValida!nPlazo
    'Muestra Datos
    If fgMuestraCredPig_AXDesCon(psNroContrato, Me.AXDesCon, False) Then
    
    End If

    ' Fecha de Vencimiento es feriado - OJO
    
    lnDiasAtraso = DateDiff("d", Format(lrValida!dVenc, "dd/mm/yyyy"), Format(gdFecSis, "dd/mm/yyyy"))
    'vDiasAtrasoReal = vDiasAtraso
    Me.txtDiasAtraso = val(lnDiasAtraso)
    txtNroRenovacion.Text = val(lrValida!nNroRenov) + 1
    txtPlazoActual.Text = val(lrValida!nPlazo)
    
    
    
    If lnDiasAtraso > 0 Then
        MsgBox "Contrato se encuentra Vencido, No es posible Amortizar el Crédito", vbInformation, "Aviso"
        Exit Sub
    End If
    
    'vDiasAtrasoReal = vDiasAtraso
'    Set loCalculos = New NColPCalculos
'        lnDeuda = loCalculos.nCalculaDeudaPignoraticio(fnVarSaldoCap, fdVarFecVencimiento, fnVarValorTasacion, _
'                fnVarTasaInteresVencido, fnVarTasaCustodiaVencida, fnVarTasaImpuesto, fnVarEstado, _
'                fnVarTasaPreparacionRemate, gdFecSis)
'
'        lnMinimoPagar = loCalculos.nCalculaMinimoPagar(fnVarSaldoCap, fnVarTasaInteres, fnVarPlazo, fnVarTasaCustodia, _
'                fdVarFecVencimiento, fnVarValorTasacion, fnVarTasaInteresVencido, fnVarTasaCustodiaVencida, _
'                fnVarTasaImpuesto, fnVarEstado, fnVarTasaPreparacionRemate, gdFecSis)
'    Set loCalculos = Nothing
        
    ' Calcula el Monto Total de la Deuda
    fgCalculaDeuda
   
   fgCalculaMinimoPagar
        
    ' Muestra datos
    
    
    'txtTotalDeuda.Text = Format(fnVarDeuda, "#0.00")
    'fnVarMontoMinimo = 0
    txtMontoMinimoPagar.Text = Format(fnVarMontoMinimo, "#0.00")
  '  fnVarMontoMinimo = 0
    'txtMontoPagar.Text = Format(fnVarMontoMinimo, "#0.00")
     TxtMontoTotal.Text = Format(fnVarMontoMinimo, "#0.00")
    
    
    
    
    If gITF.gbITFAplica Then
        If Not gITF.gbITFAsumidocreditos Then
'            Me.TxtITF = Format(gITF.fgITFCalculaImpuestoIncluido(TxtMontoTotal), "#0.00")
'            txtMontoPagar = Format(CDbl(Me.TxtMontoTotal) - CDbl(Me.TxtITF), "#0.00")
'            TxtMontoTotal.Text = Format(gITF.fgITFCalculaImpuestoIncluido(txtMontoPagar.Text) + Val(txtMontoPagar), "#0.00")
            'Me.TxtITF = gITF.fgITFCalculaImpuestoIncluido(txtMontoPagar.Text)
             
'             TxtMontoTotal.Text = Format(gITF.fgITFCalculaImpuestoIncluido(txtMontoPagar.Text), "#0.00")
'              TxtITF.Text = CCur(TxtMontoTotal.Text) - CCur(txtMontoPagar.Text)

             
             
           TxtITF = gITF.fgITFCalculaImpuesto(txtMontoPagar)
           txtMontoPagar = CCur(TxtMontoTotal.Text) + CCur(TxtITF.Text)
                        
           Dim Aux As String
           If InStr(1, CStr(TxtITF), ".", vbTextCompare) > 0 Then
                Aux = CDbl(CStr(Int(TxtITF)) & "." & Mid(CStr(TxtITF), InStr(1, CStr(TxtITF), ".", vbTextCompare) + 1, 2))
           Else
                Aux = CDbl(CStr(Int(TxtITF)))
           End If
            
            TxtITF.Text = Format(Aux, "#0.00")
            '*** BRGO 20110908 ************************************************
            nRedondeoITF = fgDiferenciaRedondeoITF(CCur(Me.TxtITF.Text))
            If nRedondeoITF > 0 Then
               Me.TxtITF.Text = Format(CCur(Me.TxtITF.Text) - nRedondeoITF, "#,##0.00")
            End If
            '*** END BRGO
            '****ULTIMAS ACTUALIZACIONES
           ' Me.TxtITF = Format(gITF.fgITFCalculaImpuesto(txtMontoPagar.Text), "#0.00")
           ' TxtMontoTotal = Format(CDbl(Me.txtMontoPagar) + CDbl(val(TxtITF.Text)), "#0.00")
        Else
            Me.TxtITF = Format(gITF.fgITFCalculaImpuesto(txtMontoPagar.Text), "#0.00")
            '*** BRGO 20110908 ************************************************
            nRedondeoITF = fgDiferenciaRedondeoITF(CCur(Me.TxtITF.Text))
            If nRedondeoITF > 0 Then
               Me.TxtITF.Text = Format(CCur(Me.TxtITF.Text) - nRedondeoITF, "#,##0.00")
            End If
            '*** END BRGO
            TxtMontoTotal = Format(CDbl(Me.txtMontoPagar), "#0.00")
        End If
    Else
            Me.TxtITF = Format(0, "#0.00")
            TxtMontoTotal = Format(Me.txtMontoPagar, "#0.00")
    End If
    
    
        
    txtSaldoCapitalNuevo.Text = "0.00"
'    cboPlazoNuevo.Text = lrValida!nPlazo
'    cboPlazoNuevo.Enabled = True
'    cboPlazoNuevo.SetFocus
    
'*** PEAC 20080528
If CCur(AXDesCon.SaldoCapital) > 0 Then
    fnVarDeuda = CCur(AXDesCon.SaldoCapital) + TxtITF.Text + fnVarInteres + fnVarInteresVencido + fnVarCostoCustodiaVencida + fnVarImpuesto + fnVarCostoPreparacionRemate + fnVarCostoNotificacion
    txtTotalDeuda.Text = Format(fnVarDeuda, "#0.00")
    
    txtCapital.Text = 0 'Format(txtMontoPagar.Text - fnVarCostoNotificacion - fnVarInteres - TxtITF.Text, "#0.00")
    txtInteres.Text = Format(fnVarInteres, "#0.00")
    txtCostoCus.Text = Format(fnVarCostoCustodiaVencida, "#0.00")
    txtCostoNoti.Text = Format(fnVarCostoNotificacion, "#0.00")
     
    txtSaldoCapitalNuevo.Text = Format(CCur(AXDesCon.SaldoCapital) - fnVarCapitalPagado, "0#.00")
End If

    Set lrValida = Nothing
    
    AXCodCta.Enabled = False
    cmdGrabar.Enabled = True
    txtMontoPagar.Enabled = True
    Me.txtMontoPagar.SetFocus
     '   TxtMontoTotal.SetFocus

        'madm 20091120 --------------------------------------------------------
         Set lafirma = New frmPersonaFirma
         Set ClsPersona = New COMDPersona.DCOMPersonas
        
         Set Rf = ClsPersona.BuscaCliente(gColPigFunciones.vcodper, BusquedaCodigo)
         If Not Rf.BOF And Not Rf.EOF Then
            If Rf!nPersPersoneria = 1 Then
            Call frmPersonaFirma.Inicio(Trim(gColPigFunciones.vcodper), Mid(gColPigFunciones.vcodper, 4, 2), False, True)
            End If
         End If
         Set Rf = Nothing
        'firma madm -----------------------------------------------------------



Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox "Error: " & Err.Number & " " & Err.Description & vbCr & _
        "Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

Private Sub cmdBuscar_Click()

Dim loPers As COMDPersona.UCOMPersona
Dim lsPersCod As String, lsPersNombre As String
Dim lsEstados As String
Dim loPersContrato As COMDColocPig.DCOMColPContrato
Dim lrContratos As New ADODB.Recordset
Dim loCuentas As COMDPersona.UCOMProdPersona

On Error GoTo ControlError

Set loPers = New COMDPersona.UCOMPersona
    Set loPers = frmBuscaPersona.Inicio
    If loPers Is Nothing Then Exit Sub
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
    Set loCuentas = frmCuentasPersona.Inicio(lsPersNombre, lrContratos) ' RIRO 20130724 SEGUN ERS101-2013
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

'Cancela el proceso actual y permite inicializar ls variables para otro proceso
Private Sub cmdCancelar_Click()
    Limpiar
    cmdGrabar.Enabled = False
    cboPlazoNuevo.Enabled = False
    txtMontoPagar.Enabled = False
    AXCodCta.Enabled = True
    AXCodCta.SetFocusCuenta
End Sub

'Actualiza los cambios en la basede datos
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
Dim loGrabarRen As COMNColoCPig.NCOMColPContrato 'COMNColoCPig.NCOMColPContrato
Dim loImprime As COMNColoCPig.NCOMColPImpre
Dim loPrevio As previo.clsprevio

Dim lsMovNro As String
Dim lsFechaHoraGrab As String
Dim lsFechaVenc As String
Dim lnMontoTransaccion As Currency
Dim lsCadImprimir As String
Dim lsNombreCliente As String

Dim loLavDinero As frmMovLavDinero
Set loLavDinero = New frmMovLavDinero
Dim loMov As COMDMov.DCOMMov 'BRGO 20110908
Dim lsFecVenImp  As String 'RECO20160412

If Not ValidaAlGrabar Then Exit Sub '*** PEAC 20160728

'*** PEAC 20071206 - cuando amortiza se correra la fecha de venc a 30 dias desde el dia que paga ********
'lsFechaVenc = Format$(fdVarFecVencimiento + fnVarNewPlazo, "mm/dd/yyyy")
'lsFechaVenc = Format$(fdVarFecVencimiento, "mm/dd/yyyy")
lsFechaVenc = Format$(gdFecSis + fnVarNewPlazo, "mm/dd/yyyy")
lsFecVenImp = Format$(gdFecSis + fnVarNewPlazo, "dd/MM/yyyy") 'RECO20160412
'*********************************************************************************************************

'lnMontoTransaccion = CCur(Me.txtMontoPagar.Text)

lsNombreCliente = AXDesCon.listaClientes.ListItems(1).ListSubItems.iTem(1)

'    If CCur(TxtMontoTotal.Text) < CCur(txtMontoMinimoPagar.Text) Then
'             MsgBox "El monto a pagar debe ser igual o mayor al monto mínimo", vbOKOnly + vbExclamation, "AVISO"
'             TxtMontoTotal.Text = txtMontoMinimoPagar.Text
'              If gITF.gbITFAplica Then
'                      If Not gITF.gbITFAsumidocreditos Then
'
'                             txtMontoPagar = gITF.fgITFCalculaImpuestoNOIncluido(TxtMontoTotal.Text)
'                             TxtITF = CCur(txtMontoPagar.Text) - CCur(TxtMontoTotal.Text)
'
'                                Dim aux As String
'                                If InStr(1, CStr(TxtITF), ".", vbTextCompare) > 0 Then
'                                     aux = CDbl(CStr(Int(TxtITF)) & "." & Mid(CStr(TxtITF), InStr(1, CStr(TxtITF), ".", vbTextCompare) + 1, 2))
'                                Else
'                                     aux = CDbl(CStr(Int(TxtITF)))
'                                End If
'
'                                 TxtITF.Text = Format(aux, "#0.00")
'
'                        Else
'                            Me.TxtITF = Format(gITF.fgITFCalculaImpuesto(txtMontoPagar.Text), "#0.00")
'                            TxtMontoTotal = Format(CDbl(Me.txtMontoPagar), "#0.00")
'                        End If
'                Else
'                        Me.TxtITF = Format(0, "#0.00")
'                        TxtMontoTotal = Format(Me.txtMontoPagar, "#0.00")
'                End If
'
'
'             Exit Sub
'    End If


'If CDbl(txtMontoPagar.Text) <= 0 Then ' peac 20071106
If CDbl(txtMontoPagar.Text) <= 0 And gcCredAntiguo <> "A" Then ' peac 20071106
   MsgBox "El monto a amortizar debe ser mayor a cero.", vbOKOnly + vbInformation, "AVISO"
   Exit Sub
End If

'If CDbl(txtMontoMinimoPagar.Text) > CDbl(txtMontoPagar.Text) Then
'    MsgBox "El valor del monto debe ser mayor o igual al monto mínimo a pagar"
'    Exit Sub
'End If

'*** PEAC 20100125
If CDbl(txtMontoPagar) < CDbl(txtMontoMinimoPagar) Then 'Monto Pagar < MinimoPagar
   MsgBox " Monto a Pagar debe ser Mayor o igual que el Mínimo.", , " Aviso "
   txtMontoPagar.SetFocus
   cmdGrabar.Enabled = False
   Exit Sub
End If
'*** FIN PEAC
'WIOR 20121009**********************************************************
Dim oDPersona As COMDPersona.DCOMPersona
Dim rsPersonaCred As ADODB.Recordset
Dim rsPersona As ADODB.Recordset
Dim Cont As Integer
Set oDPersona = New COMDPersona.DCOMPersona

Set rsPersonaCred = oDPersona.ObtenerPersCuentaRelac(Trim(AXCodCta.NroCuenta), gColRelPersTitular)

If rsPersonaCred.RecordCount > 0 Then
    If Not (rsPersonaCred.EOF And rsPersonaCred.BOF) Then
        For Cont = 0 To rsPersonaCred.RecordCount - 1
            'WIOR 20130301 **************
            ReDim Preserve PersonaPago(Cont, 1)
            PersonaPago(Cont, 0) = Trim(rsPersonaCred!cPersCod)
            PersonaPago(Cont, 1) = Trim(rsPersonaCred!cPersNombre)
            'WIOR **********************
            Set rsPersona = oDPersona.ObtenerUltimaVisita(Trim(rsPersonaCred!cPersCod))
            If rsPersona.RecordCount > 0 Then
                If Not (rsPersona.EOF And rsPersona.BOF) Then
                    If Trim(rsPersona!sUsual) = "3" Then
                    MsgBox PstaNombre(Trim(rsPersonaCred!cPersNombre), True) & "." & Chr(10) & "CLIENTE OBSERVADO: " & Trim(rsPersona!cVisObserva), vbInformation, "Aviso"
                        Call frmPersona.Inicio(Trim(rsPersonaCred!cPersCod), PersonaActualiza)
                    End If
                   
                End If
            End If
            Call VerSiClienteActualizoAutorizoSusDatos(Trim(rsPersonaCred!cPersCod), fnVarOpeCod) 'FRHU ERS077-2015 20151204
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

If MsgBox(" Grabar Amortización de Contrato Pignoraticio ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
    cmdGrabar.Enabled = False
        
        If fnVarCapitalPagado <= 0 Then
            fnVarCapitalPagado = 0
            lnMontoTransaccion = CCur(Me.txtMontoPagar.Text) + CCur(TxtITF.Text)
        Else
            lnMontoTransaccion = CCur(Me.txtMontoPagar.Text)
            'lnMontoTransaccion = fnVarCapitalPagado '***PEAC 20080118
        End If
        
        ' por peac 20070814
        fnVarNewSaldoCap = Format(CCur(AXDesCon.SaldoCapital) - fnVarCapitalPagado, "0#.00")
        'txtSaldoCapitalNuevo.Text = fnVarNewSaldoCap
      
        Dim clsExo As New COMNCaptaServicios.NCOMCaptaServicios
        If Not clsExo.EsCuentaExoneradaLavadoDinero(AXCodCta.NroCuenta) Then
          Dim sPersLavDinero As String
          Dim nMontoLavDinero As Double, nTC As Double
          Dim clsLav As New COMNCaptaGenerales.NCOMCaptaDefinicion, nMoneda As Integer, nMonto As Double
    
            'nMonto = CDbl(TxtMontoTotal.Text)
            nMonto = CCur(Me.txtMontoPagar.Text) '***PEAC 20080118
            
            Set clsExo = Nothing
            sPersLavDinero = ""
            nMontoLavDinero = clsLav.GetCapParametro(gMonOpeLavDineroME)
            Set clsLav = Nothing
            nMoneda = gMonedaNacional
            If nMoneda = gMonedaNacional Then
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
                sPersLavDinero = loLavDinero.Inicia(, , , , False, True, nMonto, AXCodCta.NroCuenta, Mid(Me.Caption, 15), True, "", , , , , nMoneda)
                If loLavDinero.OrdPersLavDinero = "" Then Exit Sub
            'End

               ' sPersLavDinero = IniciaLavDinero()
               ' If sPersLavDinero = "" Then Exit Sub
            End If
         Else
            Set clsExo = Nothing
         End If
         Set clsExo = Nothing
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
        
        lsFechaHoraGrab = fgFechaHoraGrab(lsMovNro)
        Set loGrabarRen = New COMNColoCPig.NCOMColPContrato
            'Grabar Amoritizacion Pignoraticio
                          
            Call loGrabarRen.nAmortizacionCredPignoraticio(AXCodCta.NroCuenta, fnVarNewSaldoCap, lsFechaHoraGrab, _
                 lsMovNro, lsFechaVenc, fnVarNewPlazo, lnMontoTransaccion - CCur(val(Me.TxtITF.Text)), fnVarCapitalPagado, fnVarInteresVencido, _
                 fnVarCostoCustodiaVencida, fnVarCostoPreparacionRemate, fnVarInteres, fnVarImpuesto, fnVarCostoCustodia, _
                 fnVarDiasAtraso, fnVarDiasCambCart, fnVarValorTasacion, fnVarOpeCod, _
                 fsVarOpeDesc, fsVarPersCodCMAC, gITF.gbITFAplica, gITF.gbITFAsumidocreditos, CCur(val(Me.TxtITF.Text)), False, fnVarCostoNotificacion, gnMovNro) 'WIOR 20130301 agrego gnMovNro
                 
        Set loGrabarRen = Nothing
        '*** BRGO 20110906 ***************************
        If gITF.gbITFAplica Then
           Set loMov = New COMDMov.DCOMMov
           Call loMov.InsertaMovRedondeoITF(lsMovNro, 1, CCur(Me.TxtITF) + nRedondeoITF, CCur(Me.TxtITF))
           Set loMov = Nothing
        End If
        '*** BRGO
'-- fnVarInteresVencido + fnVarInteres  EL INTERES ES 0 PORQUE SOLO SE MUEVE EL SALDO DE CAPITAL
        Set loImprime = New COMNColoCPig.NCOMColPImpre
'            lsCadImprimir = loImprime.nPrintReciboAmortizacion(gsNomAge, lsFechaHoraGrab, AXCodCta.NroCuenta, lsNombreCliente, _
'                Format(AXDesCon.FechaPrestamo, "mm/dd/yyyy"), fnVarDiasAtraso, CCur(AXDesCon.SaldoCapital), fnVarCapitalPagado, _
'                fnVarInteres, fnVarImpuesto, fnVarCostoCustodiaVencida + fnVarCostoCustodia, _
'                fnVarCostoPreparacionRemate, lnMontoTransaccion, fnVarNewSaldoCap, fnVarTasaInteres, _
'                Val(Me.txtNroRenovacion.Text), Format(lsFechaVenc, "mm/dd/yyyy"), gsCodUser, fnVarNewPlazo, _
'                fsVarNombreCMAC, " ", CDbl(Val(TxtITF.Text)), gImpresora, gbImpTMU, fnVarCostoNotificacion)

            lsCadImprimir = loImprime.nPrintReciboAmortizacion(gsNomAge, lsFechaHoraGrab, AXCodCta.NroCuenta, lsNombreCliente, _
                Format(AXDesCon.FechaPrestamo, "dd/MM/yyyy"), fnVarDiasAtraso, CCur(AXDesCon.SaldoCapital), fnVarCapitalPagado, _
                fnVarInteres, fnVarImpuesto, fnVarCostoCustodiaVencida + fnVarCostoCustodia, _
                fnVarCostoPreparacionRemate, lnMontoTransaccion, fnVarNewSaldoCap, fnVarTasaInteres, _
                val(Me.txtNroRenovacion.Text), Format(lsFecVenImp, "dd/MM/yyyy"), gsCodUser, fnVarNewPlazo, _
                fsVarNombreCMAC, " ", CDbl(val(TxtITF.Text)), gImpresora, gbImpTMU, fnVarCostoNotificacion)

        Set loImprime = Nothing
        Set loPrevio = New previo.clsprevio
            loPrevio.PrintSpool sLpt, lsCadImprimir, False, 22
            Do While True
                If MsgBox("Reimprimir Recibo de Amortización ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
                    loPrevio.PrintSpool sLpt, lsCadImprimir, False, 22
                    
                Else
                    Set loPrevio = Nothing
                    Exit Do
                End If
            Loop
            Set loPrevio = Nothing
            Set loLavDinero = Nothing
        'WIOR 20130301 ************************************************************
        If regPersonaRealizaPago And gnMovNro > 0 Then
            frmPersRealizaOpeGeneral.InsertaPersonasOperacion gnMovNro, Trim(AXCodCta.NroCuenta), fnCondicion
            regPersonaRealizaPago = False
        End If
        'WIOR FIN *****************************************************************
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
nMonto = CDbl(TxtMontoTotal.Text)
sCuenta = AXCodCta.NroCuenta
'If sPersCodCMAC <> "" Then
'    IniciaLavDinero = frmMovLavDinero.Inicia(sPersCod, sNombre, sDireccion, sDocId, False, True, nmonto, sCuenta, sOperacion, , sTipoCuenta)
'Else
    'IniciaLavDinero = frmMovLavDinero.Inicia(sPersCod, sNombre, sDireccion, sDocId, False, True, nMonto, sCuenta, CStr(fnVarOpeCod), , gMonedaNacional)
'End If
End Sub

'Termina el formulario actual
Private Sub cmdSalir_Click()
    Unload Me
End Sub

'Procedimiento que calcula la deuda del cliente
Private Sub fgCalculaDeuda()
Dim loCalculos As COMNColoCPig.NCOMColPCalculos
fnVarDiasAtraso = DateDiff("d", fdVarFecVencimiento, gdFecSis)
If fnVarDiasAtraso <= 0 Then
    fnVarDiasAtraso = 0
    
    'PEAC 20070813
    'end peac
    
    'PEAC 20070813
    If gcCredAntiguo = "A" Then
        fnVarInteres = Round(0, 2)
    Else
        Set loCalculos = New COMNColoCPig.NCOMColPCalculos
        vDiasAdel = DateDiff("d", vFecEstado, Format(gdFecSis, "dd/mm/yyyy"))
        '*** PEAC 20080806 *************************************
        'fnVarInteres = loCalculos.nCalculaInteresAdelantado(CCur(AXDesCon.SaldoCapital), fnVarTasaInteres, vDiasAdel) 'fnVarSaldoCap
         fnVarInteres = loCalculos.nCalculaInteresAlVencimiento(CCur(AXDesCon.SaldoCapital), fnVarTasaInteres, vDiasAdel)
        '*** FIN PEAC ******************************************
        fnVarInteres = Round(fnVarInteres, 2)
        Set loCalculos = Nothing
    End If
    
    fnVarInteresVencido = 0
    fnVarCostoCustodia = 0
    fnVarImpuesto = 0
Else
    Set loCalculos = New COMNColoCPig.NCOMColPCalculos
        fnVarInteresVencido = loCalculos.nCalculaInteresMoratorio(fnVarSaldoCap, fnVarTasaInteresVencido, fnVarDiasAtraso)
        fnVarInteresVencido = Round(fnVarInteresVencido, 2)
        fnVarCostoCustodiaVencida = loCalculos.nCalculaCostoCustodiaMoratorio(fnVarValorTasacion, fnVarTasaCustodiaVencida, fnVarDiasAtraso)
        fnVarCostoCustodiaVencida = Round(fnVarCostoCustodiaVencida, 2)
        fnVarImpuesto = (fnVarInteresVencido + fnVarCostoCustodiaVencida) * fnVarTasaImpuesto
        fnVarImpuesto = Round(fnVarImpuesto, 2)
    Set loCalculos = Nothing
End If
fnVarCostoPreparacionRemate = 0

'*** PEAC 20080515
If fnVarEstado = gColPEstPRema Then    ' Si esta en via de Remate
    fnVarCostoPreparacionRemate = fnVarTasaPreparacionRemate * fnVarValorTasacion
    fnVarCostoPreparacionRemate = Round(fnVarCostoPreparacionRemate, 2)
End If

'*** PEAC 20080515
'If gnNotifiAdju = 0 Then
    fnVarCostoNotificacion = 0
'End If

'fnVarDeuda = fnVarSaldoCap + fnVarInteresVencido + fnVarCostoCustodiaVencida + fnVarImpuesto + fnVarCostoPreparacionRemate + fnVarCostoNotificacion





End Sub

Private Sub fgCalculaMinimoPagar()
Dim loCalculos As COMNColoCPig.NCOMColPCalculos

' **************************************
' ** Calculo del Monto Minimo a Pagar **
' **************************************
'    Dim NumeroDias As Single
Set loCalculos = New COMNColoCPig.NCOMColPCalculos
    'fnVarInteres = 0
    fnVarFactor = loCalculos.nCalculaFactorRenovacion(fnVarTasaInteres, fnVarNewPlazo)
    'Ubicacion corte
    fnVarCostoCustodia = loCalculos.nCalculaCostoCustodia(fnVarValorTasacion, fnVarTasaCustodia, fnVarNewPlazo)
    fnVarCostoCustodia = Round(fnVarCostoCustodia, 2)
    
'    fnVarInteres = fnVarSaldoCap * fnVarFactor
'    fnVarInteres = Round(fnVarInteres, 2)
    
    fnVarImpuesto = (fnVarInteresVencido + fnVarInteres + fnVarCostoCustodia + fnVarCostoCustodiaVencida) * fnVarTasaImpuesto
    fnVarImpuesto = Round(fnVarImpuesto, 2)
    
    
    If fnVarFechaRenovacion = gdFecSis Then
        fnVarMontoMinimo = Round(0.01, 2)
    Else
        If gcCredAntiguo = "A" Then
            fnVarMontoMinimo = Round(0.01, 2)
        Else
            fnVarMontoMinimo = fnVarInteresVencido + fnVarCostoCustodiaVencida + fnVarInteres + fnVarCostoCustodia + fnVarImpuesto + fnVarCostoPreparacionRemate + fnVarCostoNotificacion
            fnVarMontoMinimo = Round(fnVarMontoMinimo, 2)
        End If
    End If
    
Set loCalculos = Nothing
End Sub
'Valida el campo cboplazonuevo
Private Sub cboPlazoNuevo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    'txtMontoPagar.SetFocus
   ' TxtMontoTotal.SetFocus
End If
End Sub
Private Sub cboPlazoNuevo_Click()
    fnVarNewPlazo = val(cboPlazoNuevo.Text)
    fgCalculaMinimoPagar
    txtMontoMinimoPagar.Text = Format(fnVarMontoMinimo, "#0.00")
    txtMontoPagar.Text = Format(fnVarMontoMinimo, "#0.00")
    fnVarCapitalPagado = 0
    txtSaldoCapitalNuevo.Text = Format(fnVarSaldoCap - fnVarCapitalPagado, "#0.00")
End Sub

Private Sub Form_Load()
Me.Icon = LoadPicture(App.path & gsRutaIcono)
'ventana = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 And AXCodCta.Enabled = True Then 'F12
        Dim sCuenta As String
        'sCuenta = frmValTarCodAnt.Inicia(gColConsuPrendario, False)
        sCuenta = frmValTarCodAnt.Inicia(gColProConsumoPrendario, False)
        If sCuenta <> "" Then
            AXCodCta.NroCuenta = sCuenta
            AXCodCta.SetFocusCuenta
        End If
    ElseIf KeyCode = 13 And Trim(AXCodCta.EnabledCta) And AXCodCta.Age <> "" And Trim(AXCodCta.Cuenta) = "" Then
                AXCodCta.SetFocusCuenta
                 Exit Sub
    End If
End Sub


Private Sub txtMontoPagar_Change()

fnVarInteres = fnVarInteres

If IsNumeric(txtMontoPagar.Text) Then
    'fnVarCapitalPagado = txtMontoPagar.Text 'peac 20070820
    fnVarCapitalPagado = txtMontoPagar.Text - fnVarInteres - TxtITF.Text ' peac 20070820
          
          
     If gITF.gbITFAplica Then
        If Not gITF.gbITFAsumidocreditos Then
'            Me.TxtITF = Format(gITF.fgITFCalculaImpuestoIncluido(TxtMontoTotal), "#0.00")
'            txtMontoPagar = Format(CDbl(Me.TxtMontoTotal) - CDbl(Me.TxtITF), "#0.00")
'            TxtMontoTotal.Text = Format(gITF.fgITFCalculaImpuestoIncluido(txtMontoPagar.Text) + Val(txtMontoPagar), "#0.00")
'            Me.TxtITF = gITF.fgITFCalculaImpuestoIncluido(txtMontoPagar.Text)
            TxtMontoTotal.Text = "0.00"
           
           '*** PEAC 20080828 ***********************************************************
'            If TxtMontoTotal.Text <> txtMontoMinimoPagar.Text Then
'              TxtItf.Text = Format(gITF.fgITFCalculaImpuesto(txtMontoPagar.Text))
'              TxtMontoTotal.Text = CCur(txtMontoPagar.Text) + CCur(TxtItf.Text)
'            End If

            If txtMontoPagar.Text <> txtMontoMinimoPagar.Text Then
              TxtITF.Text = Format(gITF.fgITFCalculaImpuesto(txtMontoPagar.Text))
              '*** BRGO 20110908 ************************************************
              nRedondeoITF = fgDiferenciaRedondeoITF(CCur(Me.TxtITF.Text))
              If nRedondeoITF > 0 Then
                  Me.TxtITF.Text = Format(CCur(Me.TxtITF.Text) - nRedondeoITF, "#,##0.00")
              End If
              '*** END BRGO
                  TxtMontoTotal.Text = CCur(txtMontoPagar.Text) + CCur(TxtITF.Text)
              End If
            '*** FIN PEAC ***********************************************************

'
           Dim Aux As String
           If InStr(1, CStr(TxtITF), ".", vbTextCompare) > 0 Then
            Aux = CDbl(CStr(Int(TxtITF)) & "." & Mid(CStr(TxtITF), InStr(1, CStr(TxtITF), ".", vbTextCompare) + 1, 2))
           Else
            Aux = CDbl(CStr(Int(TxtITF)))
           End If
            TxtITF.Text = Format(TxtITF.Text, "#0.00")
            
            
            
        '--
'            Me.TxtITF = gITF.fgITFCalculaImpuesto(txtMontoPagar.Text)
'            TxtMontoTotal = Format(CDbl(Me.txtMontoPagar) + CDbl(val(TxtITF.Text)), "#0.00")
        Else
            Me.TxtITF = gITF.fgITFCalculaImpuesto(txtMontoPagar.Text)
            '*** BRGO 20110908 ************************************************
            nRedondeoITF = fgDiferenciaRedondeoITF(CCur(Me.TxtITF.Text))
            If nRedondeoITF > 0 Then
               Me.TxtITF.Text = Format(CCur(Me.TxtITF.Text) - nRedondeoITF, "#,##0.00")
            End If
            '*** END BRGO
            TxtMontoTotal = Format(CDbl(Me.txtMontoPagar), "#0.00")
        End If
    Else
            Me.TxtITF = Format(0, "#0.00")
            TxtMontoTotal = Format(Me.txtMontoPagar, "#0.00")
    End If
    
'    If CCur(TxtMontoTotal.Text) < CCur(txtMontoMinimoPagar.Text) Then
'             MsgBox "El monto a pagar debe ser igual o mayor al monto mínimo", vbOKOnly + vbExclamation, "AVISO"
'             TxtMontoTotal.Text = txtMontoMinimoPagar.Text
'             TxtMontoTotal_KeyPress 13
'             Exit Sub
'    End If
    
    Me.TxtITF = Me.TxtITF
    
    fnVarNewSaldoCap = Format(fnVarSaldoCap - fnVarCapitalPagado, "0#.00")
    txtSaldoCapitalNuevo.Text = fnVarNewSaldoCap
    
    
'*** PEAC 20080528
If CCur(AXDesCon.SaldoCapital) > 0 Then
    fnVarDeuda = CCur(AXDesCon.SaldoCapital) + TxtITF.Text + fnVarInteres + fnVarInteresVencido + fnVarCostoCustodiaVencida + fnVarImpuesto + fnVarCostoPreparacionRemate + fnVarCostoNotificacion
    txtTotalDeuda.Text = Format(fnVarDeuda, "#0.00")
    
    txtCapital.Text = Format(txtMontoPagar.Text - fnVarCostoNotificacion - fnVarInteres - TxtITF.Text, "#0.00")
    txtInteres.Text = Format(fnVarInteres, "#0.00")
    txtCostoCus.Text = Format(fnVarCostoCustodiaVencida, "#0.00")
    txtCostoNoti.Text = Format(fnVarCostoNotificacion, "#0.00")
        
    txtSaldoCapitalNuevo.Text = Format(CCur(AXDesCon.SaldoCapital) - fnVarCapitalPagado, "0#.00")
End If

End If
End Sub

'Valida el campo txtmontopagar
Private Sub txtMontoPagar_GotFocus()
    fEnfoque txtMontoPagar
End Sub
Private Function ValidaAlGrabar() As Boolean
    ValidaAlGrabar = False
    
    If CDbl(txtMontoPagar) >= CDbl(txtTotalDeuda) Then  'Monto Pagar >= Total Deuda
       MsgBox " Cliente debe Cancelar su contrato ", , " Aviso "
       cmdGrabar.Enabled = False
       Exit Function
    End If
    
    If CDbl(txtMontoPagar) < CDbl(txtMontoMinimoPagar) Then 'Monto Pagar < MinimoPagar
       MsgBox " Monto a Pagar debe ser Mayor o igual que el Mínimo.", , " Aviso "
       txtMontoPagar.SetFocus
       cmdGrabar.Enabled = False
       Exit Function
    End If
    
    If val(txtMontoPagar) <= 0 Then
       cmdGrabar.Enabled = False
       Exit Function
    End If
    
    ValidaAlGrabar = True

End Function


Private Sub txtMontoPagar_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtMontoPagar, KeyAscii)

If KeyAscii = 13 Then
    fnVarInteres = fnVarInteres ' peac 20070820
    
    If Not ValidaAlGrabar Then Exit Sub
    
'////////////////////////////////////////////////////////////////////////////
'    If CDbl(txtMontoPagar) >= CDbl(txtTotalDeuda) Then  'Monto Pagar >= Total Deuda
'       MsgBox " Cliente debe Cancelar su contrato ", , " Aviso "
'       'txtMontoPagar.SetFocus
'       'TxtMontoTotal.SetFocus
'       cmdGrabar.Enabled = False
'       Exit Sub
'    End If
'
'    '*** PEAC 20100125
'    If CDbl(txtMontoPagar) < CDbl(txtMontoMinimoPagar) Then 'Monto Pagar < MinimoPagar
'       MsgBox " Monto a Pagar debe ser Mayor o igual que el Mínimo.", , " Aviso "
'       txtMontoPagar.SetFocus
'       cmdGrabar.Enabled = False
'       Exit Sub
'    End If
'    '*** FIN PEAC
'
'    'If Val(txtMontoPagar) < Val(txtMontoMinimoPagar) Then 'Monto Pagar < MinimoPagar
'    '   MsgBox " Monto a Pagar debe ser Mayor que el Mìnimo", , " Aviso "
'    '   txtMontoPagar.SetFocus
'    '   cmdGrabar.Enabled = False
'    '   Exit Sub
'    'End If
'    If val(txtMontoPagar) <= 0 Then
'       'txtMontoPagar.SetFocus
'      ' TxtMontoTotal.SetFocus
'       cmdGrabar.Enabled = False
'       Exit Sub
'    End If
'    'Distribuye los importes a las diferentes rubros (Todo va ha capital)
'    'fnVarCapitalPagado = 0
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

    vSumaCostoCustodia = fnVarCostoCustodia + fnVarCostoCustodiaVencida
    If CDbl(txtMontoPagar.Text) > CDbl(txtMontoMinimoPagar) Then ' Monto Pagar = Minimo Pagar
        'fnVarCapitalPagado = (Val(txtMontoPagar) - fnVarFactor * fnVarSaldoCap - vSumaCostoCustodia - fnVarInteresVencido - _
            fnVarCostoPreparacionRemate - fnVarTasaImpuesto * fnVarFactor * fnVarSaldoCap - fnVarTasaImpuesto * vSumaCostoCustodia - _
            fnVarTasaImpuesto * fnVarInteresVencido) / (1 - fnVarFactor - fnVarTasaImpuesto * fnVarFactor)
        fnVarInteres = fnVarInteres
         
        ' se agrego fnVarInteres - CDbl(TxtITF) peac 20070820
        fnVarCapitalPagado = (CDbl(txtMontoPagar) - fnVarInteres - CDbl(TxtITF) - vSumaCostoCustodia - (fnVarTasaImpuesto * vSumaCostoCustodia))
        fnVarCapitalPagado = Round(fnVarCapitalPagado, 2)
        'fnVarInteres = fnVarFactor * (fnVarSaldoCap - fnVarCapitalPagado)
        
    'PEAC 20070813
    Dim loCalculos As COMNColoCPig.NCOMColPCalculos
    Set loCalculos = New COMNColoCPig.NCOMColPCalculos
        vDiasAdel = DateDiff("d", vFecEstado, Format(gdFecSis, "dd/mm/yyyy"))
        
        If gcCredAntiguo = "A" Then
            fnVarInteres = Round(0, 2)
        Else
            '*** PEAC 20080806 ************************************
            'fnVarInteres = loCalculos.nCalculaInteresAdelantado(CCur(AXDesCon.SaldoCapital), fnVarTasaInteres, vDiasAdel)
             fnVarInteres = loCalculos.nCalculaInteresAlVencimiento(CCur(AXDesCon.SaldoCapital), fnVarTasaInteres, vDiasAdel)
            '*** FIN PEAC *****************************************
            fnVarInteres = Round(fnVarInteres, 2)
        End If
        
    Set loCalculos = Nothing
'        fnVarInteres = 0
'        fnVarInteres = Round(fnVarInteres, 2)

        fnVarImpuesto = fnVarTasaImpuesto * (fnVarInteres + vSumaCostoCustodia)
        fnVarImpuesto = Round(fnVarImpuesto, 2)
'       'PARA REDONDEO
        'fnVarInteres = Val(txtMontoPagar.Text) - fnVarCapitalPagado - fnVarInteresVencido - vSumaCostoCustodia - fnVarImpuesto - fnVarCostoPreparacionRemate
        'fnVarInteres = Round(fnVarInteres, 2)
    End If
    
    'fnVarNewSaldoCap = Format(fnVarSaldoCap - fnVarCapitalPagado, "#0.00")
    'txtSaldoCapitalNuevo.Text = fnVarNewSaldoCap
        
'*** PEAC 20080528
If CCur(AXDesCon.SaldoCapital) > 0 Then
    fnVarDeuda = CCur(AXDesCon.SaldoCapital) + TxtITF.Text + fnVarInteres + fnVarInteresVencido + fnVarCostoCustodiaVencida + fnVarImpuesto + fnVarCostoPreparacionRemate + fnVarCostoNotificacion
    txtTotalDeuda.Text = Format(fnVarDeuda, "#0.00")
    
    txtCapital.Text = Format(txtMontoPagar.Text - fnVarCostoNotificacion - fnVarInteres - TxtITF.Text, "#0.00")
    txtInteres.Text = Format(fnVarInteres, "#0.00")
    txtCostoCus.Text = Format(fnVarCostoCustodiaVencida, "#0.00")
    txtCostoNoti.Text = Format(fnVarCostoNotificacion, "#0.00")
        
    txtSaldoCapitalNuevo.Text = Format(CCur(AXDesCon.SaldoCapital) - fnVarCapitalPagado, "0#.00")
End If

    cmdGrabar.Enabled = True
    cmdGrabar.SetFocus
End If
End Sub


Private Sub CargaParametros()
Dim loParam As COMDColocPig.DCOMColPCalculos
Set loParam = New COMDColocPig.DCOMColPCalculos
    fnVarTasaPreparacionRemate = loParam.dObtieneColocParametro(gConsColPTasaPreparaRemate)
    fnVarTasaImpuesto = loParam.dObtieneColocParametro(gConsColPTasaImpuesto)
    fnVarTasaCustodiaVencida = loParam.dObtieneColocParametro(gConsColPTasaCustodiaVencida)
    fnVarTasaCustodia = loParam.dObtieneColocParametro(gConsColPTasaCustodia)
    fnVarDiasCambCart = loParam.dObtieneColocParametro(gConsColPDiasCambioCartera)
    
   'madm 20091204 -------------------------------------------------------------------
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

    'fnVarCostoNotificacion = loParam.dObtieneColocParametro(3103) '*** PEAC 20080515
    'end madm --------------------------------------------------------------
    
Set loParam = Nothing
End Sub

Private Sub txtMontoPagar_LostFocus()
    
     If Trim(txtMontoPagar.Text) = "" Then
        txtMontoPagar.Text = "0.00"
    End If
    txtMontoPagar.Text = Format(txtMontoPagar.Text, "#0.00")
    

    
End Sub

Private Sub TxtMontoTotal_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxtMontoTotal.Text = Format(TxtMontoTotal.Text, "#0.00")
    txtMontoPagar_Change
 End If
End Sub

