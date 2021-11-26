VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCFAutRenovacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Autorización de Pago Por Renovacion CF"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7560
   Icon            =   "frmCFAutRenovacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   7560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Aval"
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
      Height          =   675
      Left            =   120
      TabIndex        =   35
      Top             =   1880
      Width           =   7410
      Begin VB.Label lblNomAvalado 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2280
         TabIndex        =   38
         Tag             =   "txtnombre"
         Top             =   240
         Width           =   5010
      End
      Begin VB.Label lblCodAvalado 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1080
         TabIndex        =   37
         Tag             =   "txtcodigo"
         Top             =   240
         Width           =   1185
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Aval"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   36
         Top             =   240
         Width           =   315
      End
   End
   Begin VB.Frame FraCredito 
      Caption         =   "Carta Fianza"
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
      Height          =   1575
      Left            =   120
      TabIndex        =   12
      Top             =   2640
      Width           =   7395
      Begin VB.Label lblMontoApr 
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
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5760
         TabIndex        =   44
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Monto Inicial:"
         Height          =   195
         Left            =   4680
         TabIndex        =   43
         Top             =   600
         Width           =   945
      End
      Begin VB.Label lblMoneda 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5760
         TabIndex        =   22
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblModalidad 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   525
         Left            =   1020
         TabIndex        =   21
         Top             =   540
         Width           =   3420
      End
      Begin VB.Label lblTipoCF 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1020
         TabIndex        =   20
         Top             =   240
         Width           =   3420
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Modalidad:"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   780
      End
      Begin VB.Label Label7 
         Caption         =   "Moneda:"
         Height          =   255
         Left            =   4980
         TabIndex        =   18
         Top             =   315
         Width           =   735
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Tipo:"
         Height          =   195
         Left            =   480
         TabIndex        =   17
         Top             =   240
         Width           =   360
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Vencimiento:"
         Height          =   195
         Left            =   4680
         TabIndex        =   16
         Top             =   960
         Width           =   915
      End
      Begin VB.Label lblFecVencCF 
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
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5760
         TabIndex        =   15
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label lblAnalista 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1020
         TabIndex        =   14
         Top             =   1080
         Width           =   3420
      End
      Begin VB.Label Label9 
         Caption         =   "Analista:"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   1140
         Width           =   720
      End
   End
   Begin VB.Frame fraDatos 
      Caption         =   "Datos de Renovación"
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
      Height          =   1305
      Left            =   120
      TabIndex        =   11
      Top             =   4320
      Width           =   7380
      Begin VB.TextBox TxtMonApr 
         Height          =   285
         Left            =   1080
         TabIndex        =   41
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox TxtPeriodo 
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
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   3480
         MaxLength       =   15
         TabIndex        =   26
         Top             =   240
         Width           =   660
      End
      Begin MSMask.MaskEdBox txtFecVencNueva 
         Height          =   315
         Left            =   6000
         TabIndex        =   27
         Top             =   720
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtFecEmiNue 
         Height          =   315
         Left            =   6000
         TabIndex        =   32
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label6 
         Caption         =   "Monto:"
         Height          =   255
         Left            =   360
         TabIndex        =   42
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblcomision 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1080
         TabIndex        =   40
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Comisión:"
         Height          =   255
         Left            =   240
         TabIndex        =   39
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nueva Emision:"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   2
         Left            =   4680
         TabIndex        =   31
         Top             =   240
         Width           =   1110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Periodo:"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   1
         Left            =   2760
         TabIndex        =   30
         Top             =   240
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nuevo Vencimiento:"
         ForeColor       =   &H80000001&
         Height          =   375
         Index           =   0
         Left            =   4440
         TabIndex        =   29
         Top             =   720
         Width           =   1440
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Acreedor"
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
      Height          =   675
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   7410
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Acreedor "
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   690
      End
      Begin VB.Label lblCodAcreedor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1080
         TabIndex        =   9
         Tag             =   "txtcodigo"
         Top             =   240
         Width           =   1185
      End
      Begin VB.Label lblNomAcreedor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2280
         TabIndex        =   8
         Tag             =   "txtnombre"
         Top             =   240
         Width           =   5010
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Afianzado"
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
      Height          =   645
      Left            =   120
      TabIndex        =   3
      Top             =   540
      Width           =   7425
      Begin VB.Label lblNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2280
         TabIndex        =   6
         Tag             =   "txtnombre"
         Top             =   210
         Width           =   5025
      End
      Begin VB.Label lblCodigo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1080
         TabIndex        =   5
         Tag             =   "txtcodigo"
         Top             =   210
         Width           =   1185
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Afianzado"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   705
      End
   End
   Begin VB.Frame Frame4 
      Height          =   675
      Left            =   120
      TabIndex        =   0
      Top             =   5760
      Width           =   7380
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   1680
         TabIndex        =   24
         Top             =   180
         Width           =   1155
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   5940
         TabIndex        =   2
         Top             =   195
         Width           =   1155
      End
      Begin VB.CommandButton cmdAutorizar 
         Caption         =   "&Autorizar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   480
         TabIndex        =   1
         Top             =   180
         Width           =   1155
      End
   End
   Begin SICMACT.ActXCodCta ActXCodCta 
      Height          =   390
      Left            =   120
      TabIndex        =   23
      Top             =   0
      Width           =   3645
      _ExtentX        =   6429
      _ExtentY        =   688
      Texto           =   "Cta Fianza"
      EnabledCMAC     =   -1  'True
      EnabledCta      =   -1  'True
      EnabledProd     =   -1  'True
      EnabledAge      =   -1  'True
   End
   Begin VB.Label lblPoliza 
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
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   6000
      TabIndex        =   34
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Num Folio :"
      Height          =   195
      Index           =   1
      Left            =   6000
      TabIndex        =   33
      Top             =   0
      Width           =   795
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Renovación: "
      Height          =   195
      Index           =   0
      Left            =   4440
      TabIndex        =   28
      Top             =   0
      Width           =   960
   End
   Begin VB.Label LblRenovacion 
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
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   4440
      TabIndex        =   25
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "frmCFAutRenovacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'*  APLICACION : Carta Fianza
'*  ARCHIVO : frmCFAuthRenovacion
'*  CREACION: 12/03/2013     - WIOR
'*************************************************************************
'*  RESUMEN: Autorización de Pago Por Renovacion CF
'***************************************************************************

Option Explicit
Dim vCodCta As String
Dim fbComisionTrimestral As Boolean
Dim loCFCalculo As COMNCartaFianza.NCOMCartaFianzaCalculos
Dim lcCons As COMDConstSistema.DCOMConstSistema
Dim loParam As COMDColocPig.DCOMColPCalculos
Dim fpComision As Double
Dim lr As New ADODB.Recordset
Dim sCodCta As String
Dim lnModalidad As Integer
Dim objPista As COMManejador.Pista
Dim nPeriodoMax As Integer 'JOEP20181222 CP
Dim nPeriodoMin As Integer 'JOEP20181222 CP

Private Sub cmdAutorizar_Click()
  Dim oDCOMCredito As COMDCredito.DCOMCredito 'LUCV20171212
  If ValidaDatos Then
    Dim oCF As COMDCartaFianza.DCOMCartaFianza
    Dim sMovNro As String
    Set oCF = New COMDCartaFianza.DCOMCartaFianza
    
    sMovNro = GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    
    If MsgBox("Estas seguro de Autorizar la renovacion de la Carta Fianza?", vbInformation + vbYesNo, "Aviso") = vbYes Then
        Call oCF.InsActAutRenovacionCF(1, _
                                       Trim(ActXCodCta.NroCuenta), _
                                       CDbl(TxtMonApr.Text), _
                                       CDbl(lblcomision.Caption), _
                                       CDbl(TxtPeriodo.Text), _
                                       Format(TxtFecEmiNue.Text, "dd/mm/yyyy"), _
                                       Format(Me.txtFecVencNueva.Text, "dd/mm/yyyy"), _
                                       True, _
                                       CInt(LblRenovacion.Caption), _
                                       sMovNro, , 0, , 0)
        
        '***** LUCV20171212, Según observacion SBS *****
        If CDbl(TxtPeriodo.Text) > 360 Then
            Set oDCOMCredito = New COMDCredito.DCOMCredito
            Call oDCOMCredito.RegistraAutorizacionesRequeridas(Format(gdFecSis & " " & GetHoraServer(), "yyyy/mm/dd hh:mm:ss"), gsCodUser, gsCodAge, ActXCodCta.NroCuenta)
            Set oDCOMCredito = Nothing
        End If
        '***** Fin LUCV20171212 *****
        
        MsgBox "Datos Guardados Satisfactoriamiente", vbInformation, "Aviso"
        LimpiarControles
        ActXCodCta.SetFocus
    End If
 End If
End Sub

Private Sub cmdCancelar_Click()
    LimpiarControles
    ActXCodCta.SetFocus
End Sub

Sub CargaDatosR(ByVal psCodCta As String)
Dim oCF As COMDCartaFianza.DCOMCartaFianza
Dim loConstante As COMDConstantes.DCOMConstantes
Dim R As New ADODB.Recordset
Dim rsVerif As New ADODB.Recordset
Dim dFecha As Date

Dim lbTienePermiso As Boolean


dFecha = gdFecSis
ActXCodCta.Enabled = False

    
        Set oCF = New COMDCartaFianza.DCOMCartaFianza
        Set R = oCF.RecuperaCartaFianzaRenovacion(psCodCta, dFecha)
        Set oCF = Nothing
        
        If Not (R.BOF And R.EOF) Then
        
            Set oCF = New COMDCartaFianza.DCOMCartaFianza
            Set rsVerif = oCF.ObternerAutRenovacion(psCodCta, "0,1", IIf(IsNull(R!nRenovacion), "0", R!nRenovacion))
            Set oCF = Nothing
    
            If (rsVerif.BOF And rsVerif.EOF) Then
            Call CP_CargaDatosAuto 'JOEP20181222 CP
            lblCodigo.Caption = R!cPersCod
            lblNombre.Caption = PstaNombre(R!cPersNombre)
        
            lblCodAcreedor.Caption = R!cPersAcreedor
            lblNomAcreedor.Caption = PstaNombre(R!cPersNomAcre)
            lnModalidad = R!nModalidad
            'JOEP20181222 CP
            If R!nModalidad = 13 Then
                lblModalidad.Caption = R!OtrsModalidades
            Else
                lblModalidad.Caption = R!sModalidad
            End If
            'JOEP20181222 CP
            'lblModalidad.Caption = R!sModalidad'comento JOEP20181222 CP
    
            lblCodAvalado.Caption = IIf(IsNull(R!cPersAvalado), "", R!cPersAvalado)
            If R!cAvalNombre <> "" Then
                lblNomAvalado.Caption = IIf(IsNull(PstaNombre(R!cAvalNombre)), "", PstaNombre(R!cAvalNombre))
            End If
    
            lblTipoCF = IIf(IsNull(R!cConsDescripcion), "", R!cConsDescripcion)
    
            If Mid(Trim(psCodCta), 9, 1) = "1" Then
                lblMoneda = "Soles"
            ElseIf Mid(Trim(psCodCta), 9, 1) = "2" Then
                lblMoneda = "Dolares"
            End If
            lblAnalista.Caption = PstaNombre(IIf(IsNull(R!cAnalista), "", R!cAnalista))
            TxtMonApr.Text = IIf(IsNull(R!nMontoApr), "", Format(R!nMontoApr, "#0.00"))
            lblMontoApr.Caption = IIf(IsNull(R!nMontoApr), "", Format(R!nMontoApr, "#0.00"))
            
            lblFecVencCF.Caption = IIf(IsNull(R!dVencimiento), "", Format(R!dVencimiento, "dd/mm/yyyy"))
            LblRenovacion.Caption = IIf(IsNull(R!nRenovacion), "0", R!nRenovacion)
            fraDatos.Enabled = True
            cmdAutorizar.Enabled = True
    
            TxtFecEmiNue.Text = CDate(lblFecVencCF.Caption) + 1
            lblPoliza.Caption = IIf(IsNull(R!nPoliza), "0", R!nPoliza)
    
        Else
            MsgBox "Carta Fianza ya fue Autorizada", vbInformation, "Aviso"
            Exit Sub
        End If
    Else
        MsgBox "La Fecha de Vencimiento puede ser Menor o Mayor a la fecha que se desea Renovar", vbInformation, "Aviso"
        Exit Sub
    End If
  
Exit Sub

ErrorCargaDat:
    MsgBox "Error Nº [" & str(Err.Number) & "] " & Err.Description, vbCritical, "Error del Sistema"
    Exit Sub
End Sub

Private Sub ActXCodCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CargaDatosR (ActXCodCta.NroCuenta)
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Icon = LoadPicture(App.Path & gsRutaIcono)
    LimpiarControles
    gsOpeCod = gCredRenovacionCF
End Sub

Sub LimpiarControles()
   ActXCodCta.Enabled = True
   ActXCodCta.NroCuenta = fgIniciaAxCuentaCF
   lblCodigo.Caption = ""
   lblNombre.Caption = ""
   lblCodAcreedor.Caption = ""
   lblNomAcreedor.Caption = ""
   lblCodAvalado.Caption = ""
   lblNomAvalado.Caption = ""
   lblTipoCF.Caption = ""
   lblMoneda.Caption = ""
   TxtMonApr.Text = ""
   lblModalidad.Caption = ""
   lblAnalista.Caption = ""
   lblFecVencCF.Caption = ""
   LblRenovacion.Caption = ""
   lblPoliza.Caption = ""
   lblcomision.Caption = ""
   txtFecVencNueva.Text = "__/__/____"
   TxtPeriodo.Text = " "
   TxtFecEmiNue.Text = "__/__/____"
   fraDatos.Enabled = False
   cmdAutorizar.Enabled = False
   TxtMonApr.Enabled = True
End Sub

Private Sub txtFecVencNueva_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdAutorizar.SetFocus
    End If
End Sub

Private Sub TxtPeriodo_Change()
    If IsNumeric(TxtPeriodo) Then
        txtFecVencNueva.Text = CDate(lblFecVencCF.Caption) + CInt(TxtPeriodo.Text) + 1
    End If
    
    Set loParam = New COMDColocPig.DCOMColPCalculos
    fpComision = loParam.dObtieneColocParametro(4001)
    Set loParam = Nothing
    sCodCta = (ActXCodCta.NroCuenta)

    Set lcCons = New COMDConstSistema.DCOMConstSistema
    Set lr = lcCons.ObtenerVarSistema()
        fbComisionTrimestral = IIf(lr!nConsSisValor = 2, True, False)
    Set lr = Nothing
    Set lcCons = Nothing
    
    If val(TxtMonApr.Text) > 0 Then '*** PEAC 20090930
        '**Inicio,Capi Octubre2007, mostrar comisión de carta fianza
        '**Las siguientes lineas de código fueron obtenidos de la pantalla de sugerenia
        Set loCFCalculo = New COMNCartaFianza.NCOMCartaFianzaCalculos
        If fbComisionTrimestral = False Then ' Caja Trujillo
            lblcomision = Format(loCFCalculo.nCalculaComisionCF(val(TxtMonApr.Text), DateDiff("d", CDate(TxtFecEmiNue), CDate(txtFecVencNueva)), fpComision, Mid(sCodCta, 9, 1)), "#,##0.00")
        Else  ' Caja Metropolitana
            lblcomision = Format(loCFCalculo.nCalculaComisionTrimestralCF(val(TxtMonApr.Text), DateDiff("d", CDate(TxtFecEmiNue), CDate(txtFecVencNueva)), lnModalidad, Mid(Trim(sCodCta), 9, 1), ActXCodCta.NroCuenta, 6), "#,###0.00")
        End If
        Set loCFCalculo = Nothing
    End If
End Sub

Private Sub TxtPeriodo_KeyPress(KeyAscii As Integer)
KeyAscii = SoloNumeros(KeyAscii) 'JOEP20181222 CP
    If KeyAscii = 13 Then
        'JOEP20181222 CP
        If TxtPeriodo.Text < nPeriodoMin Then
            MsgBox "El periodo mínimo es " & nPeriodoMin & " días", vbInformation, "Aviso"
            TxtPeriodo.Text = nPeriodoMin
            TxtPeriodo.SetFocus
        End If
        'JOEP20181222 CP
    If IsNumeric(TxtPeriodo) Then
        txtFecVencNueva.Text = CDate(lblFecVencCF.Caption) + CInt(TxtPeriodo.Text) + 1
        cmdAutorizar.SetFocus
    ElseIf Len(TxtPeriodo) = 0 Then
        txtFecVencNueva.Text = "__/__/____"
        cmdAutorizar.SetFocus
    ElseIf Not IsNumeric(TxtPeriodo) Then
        MsgBox "Solo Ingrese Valores Numéricos", vbInformation, "Aviso"
        TxtPeriodo.Text = ""
    End If
        
    End If
End Sub

Private Sub TxtPeriodo_LostFocus()

    If IsNumeric(TxtPeriodo) Then
        txtFecVencNueva.Text = CDate(lblFecVencCF.Caption) + CInt(TxtPeriodo.Text) + 1
        cmdAutorizar.SetFocus
    ElseIf Len(TxtPeriodo) = 0 Then
        txtFecVencNueva.Text = "__/__/____"
        cmdAutorizar.SetFocus
    ElseIf Not IsNumeric(TxtPeriodo) Then
        MsgBox "Solo Ingrese Valores Numericos", vbInformation, "Aviso"
        TxtPeriodo.Text = ""
    End If

End Sub

Private Function ValidaDatos() As Boolean
    If Trim(TxtMonApr.Text) = "" Or Trim(TxtMonApr.Text) = "0" Then
        MsgBox "Ingrese el Monto de la Carta Fianza.", vbInformation, "AVISO"
        ValidaDatos = False
        Exit Function
    End If
    
    If Trim(TxtPeriodo.Text) = "" Or Trim(TxtPeriodo.Text) = "0" Then
        MsgBox "Ingrese el Periodo de la Carta Fianza.", vbInformation, "AVISO"
        ValidaDatos = False
        Exit Function
    End If
    
    If txtFecVencNueva = "__/__/____" Then
        MsgBox "Falta calcular la nueva fecha de vencimiento.", vbInformation, "AVISO"
        ValidaDatos = False
        Exit Function
    End If
    
    If CDbl(TxtPeriodo.Text) > 360 Then
        'MsgBox "Renovación Necesita Autorización de la Gerencia Mancomunada.", vbInformation, "AVISO" 'LUCV20171212, Comentó según Observacion SBS
        MsgBox "El periodo ingresado implica que la Renovación necesite una Autorización.", vbInformation, "AVISO"
        If MsgBox("Desea Continuar con el proceso?", vbInformation + vbYesNo, "Aviso") = vbNo Then
            ValidaDatos = False
            Exit Function
        End If
    End If
    
'JOEP20181222 CP
    If TxtPeriodo.Text < nPeriodoMin Then
        MsgBox "El periodo minimo es " & nPeriodoMin & " dias", vbInformation, "Aviso"
        TxtPeriodo.Text = nPeriodoMin
        TxtPeriodo.SetFocus
        ValidaDatos = False
        Exit Function
    End If
'JOEP20181222 CP
ValidaDatos = True
End Function

'JOEP20181218 CP
Private Sub CP_CargaDatosAuto()
Dim oDCred As COMDCredito.DCOMCredito
Dim rsDefaut As ADODB.Recordset
Set oDCred = New COMDCredito.DCOMCredito

Set rsDefaut = oDCred.CatalogoProDefaut(514, 7000)

If Not (rsDefaut.BOF And rsDefaut.EOF) Then
    nPeriodoMin = rsDefaut!MinPlazo
    nPeriodoMax = rsDefaut!MaxPlazo
End If

End Sub
'JOEP20181218 CP
