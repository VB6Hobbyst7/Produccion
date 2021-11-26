VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCFAprobacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Carta Fianza - Aprobación"
   ClientHeight    =   8025
   ClientLeft      =   2370
   ClientTop       =   405
   ClientWidth     =   7275
   Icon            =   "frmCFAprobacion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   7275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Avalado "
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
      TabIndex        =   40
      Top             =   2280
      Width           =   7065
      Begin VB.Label lblNomAvalado 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2460
         TabIndex        =   43
         Tag             =   "txtnombre"
         Top             =   240
         Width           =   4365
      End
      Begin VB.Label lblCodAvalado 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1140
         TabIndex        =   42
         Tag             =   "txtcodigo"
         Top             =   240
         Width           =   1350
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Avalado"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   240
         TabIndex        =   41
         Top             =   270
         Width           =   585
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Aprobación"
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
      Height          =   1935
      Left            =   100
      TabIndex        =   26
      Top             =   5400
      Width           =   7095
      Begin VB.Frame frModOtrs 
         Height          =   615
         Left            =   50
         TabIndex        =   49
         Top             =   1280
         Width           =   4575
         Begin VB.TextBox txtModOtrs 
            Height          =   285
            Left            =   1080
            TabIndex        =   51
            Top             =   240
            Width           =   3375
         End
         Begin VB.Label lblModOtrs 
            Caption         =   "Modalidad Otros"
            Height          =   375
            Left            =   120
            TabIndex        =   50
            Top             =   165
            Width           =   855
         End
      End
      Begin VB.TextBox TxtPeriodoApr 
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
         Left            =   5520
         MaxLength       =   15
         TabIndex        =   48
         Top             =   600
         Width           =   660
      End
      Begin VB.TextBox txtMontoApr 
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
         Left            =   1125
         MaxLength       =   10
         TabIndex        =   29
         Top             =   240
         Width           =   1335
      End
      Begin VB.ComboBox cboApoderado 
         Height          =   315
         Left            =   1125
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   960
         Width           =   3255
      End
      Begin VB.ComboBox CboModalidad 
         Height          =   315
         Left            =   1125
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   600
         Width           =   3255
      End
      Begin MSMask.MaskEdBox TxtFecVenApr 
         Height          =   285
         Left            =   5520
         TabIndex        =   30
         Top             =   960
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtfechaAsigApr 
         Height          =   315
         Left            =   5520
         TabIndex        =   47
         Top             =   240
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label6 
         Caption         =   "Periodo"
         Height          =   255
         Left            =   4780
         TabIndex        =   46
         Top             =   650
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "Asig. Apr"
         Height          =   255
         Left            =   4690
         TabIndex        =   45
         Top             =   260
         Width           =   735
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Monto Apr. "
         Height          =   195
         Index           =   4
         Left            =   170
         TabIndex        =   34
         Top             =   300
         Width           =   825
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Venc. Apr."
         Height          =   195
         Index           =   7
         Left            =   4680
         TabIndex        =   33
         Top             =   980
         Width           =   750
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Apoderado"
         Height          =   195
         Index           =   6
         Left            =   170
         TabIndex        =   32
         Top             =   1000
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Modalidad"
         Height          =   195
         Index           =   5
         Left            =   170
         TabIndex        =   31
         Top             =   680
         Width           =   735
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
      Height          =   555
      Left            =   120
      TabIndex        =   21
      Top             =   1620
      Width           =   7050
      Begin VB.Label lblNomAcreedor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2400
         TabIndex        =   24
         Tag             =   "txtnombre"
         Top             =   180
         Width           =   4470
      End
      Begin VB.Label lblCodAcreedor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1140
         TabIndex        =   23
         Tag             =   "txtcodigo"
         Top             =   180
         Width           =   1305
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Acreedor "
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   690
      End
   End
   Begin VB.CommandButton cmdExaminar 
      Caption         =   "E&xaminar..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5820
      TabIndex        =   20
      ToolTipText     =   "Buscar Credito"
      Top             =   120
      Width           =   1215
   End
   Begin VB.Frame frmCFEmision 
      Height          =   690
      Left            =   100
      TabIndex        =   14
      Top             =   7320
      Width           =   7125
      Begin VB.CommandButton cmdGenerarPDF 
         Caption         =   "Vista Previa"
         Enabled         =   0   'False
         Height          =   390
         Left            =   4200
         TabIndex        =   44
         Top             =   195
         Width           =   1215
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   390
         Left            =   5520
         TabIndex        =   19
         Top             =   195
         Width           =   1185
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar"
         Enabled         =   0   'False
         Height          =   390
         Left            =   240
         TabIndex        =   18
         ToolTipText     =   "Grabar Datos de Aprobacion de Credito"
         Top             =   195
         Width           =   1185
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   390
         Left            =   1620
         TabIndex        =   17
         ToolTipText     =   "Ir al Menu Principal"
         Top             =   195
         Width           =   1185
      End
   End
   Begin VB.Frame fraDatos 
      Caption         =   "Carta Fianza"
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
      Height          =   2295
      Left            =   120
      TabIndex        =   8
      Top             =   3060
      Width           =   7065
      Begin VB.TextBox txtFinalidad 
         Height          =   855
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   39
         Text            =   "frmCFAprobacion.frx":030A
         Top             =   1320
         Width           =   6735
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Comisión "
         Height          =   195
         Index           =   9
         Left            =   4680
         TabIndex        =   38
         Top             =   1020
         Width           =   675
      End
      Begin VB.Label lblComision 
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
         Height          =   300
         Left            =   5640
         TabIndex        =   37
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Finalidad"
         Height          =   195
         Index           =   8
         Left            =   180
         TabIndex        =   35
         Top             =   1140
         Width           =   630
      End
      Begin VB.Label lblAnalista 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1080
         TabIndex        =   16
         Top             =   600
         Width           =   3195
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Analista"
         Height          =   195
         Index           =   1
         Left            =   195
         TabIndex        =   15
         Top             =   600
         Width           =   555
      End
      Begin VB.Label lblFecVencSug 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   5640
         TabIndex        =   0
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblMontoSug 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Left            =   5640
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Vencimiento"
         Height          =   195
         Index           =   3
         Left            =   4680
         TabIndex        =   12
         Top             =   600
         Width           =   870
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Monto"
         Height          =   195
         Index           =   2
         Left            =   4680
         TabIndex        =   11
         Top             =   240
         Width           =   450
      End
      Begin VB.Label lblTipoCF 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1080
         TabIndex        =   10
         Top             =   240
         Width           =   3180
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tipo"
         Height          =   195
         Index           =   0
         Left            =   195
         TabIndex        =   9
         Top             =   240
         Width           =   315
      End
   End
   Begin VB.Frame fraCliente 
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
      Height          =   1080
      Left            =   180
      TabIndex        =   1
      Top             =   540
      Width           =   6990
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Activ. Eco."
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   600
         Width           =   780
      End
      Begin VB.Label lblActiv 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1140
         TabIndex        =   7
         Top             =   600
         Width           =   5700
      End
      Begin VB.Label lblRazSoc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1140
         TabIndex        =   6
         Top             =   540
         Visible         =   0   'False
         Width           =   5700
      End
      Begin VB.Label lblNomcli 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2460
         TabIndex        =   5
         Top             =   180
         Width           =   4365
      End
      Begin VB.Label lblCodcli 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1140
         TabIndex        =   4
         Top             =   180
         Width           =   1350
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fuente Ing."
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   540
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   225
         Width           =   480
      End
   End
   Begin SICMACT.ActXCodCta ActXCodCta 
      Height          =   390
      Left            =   180
      TabIndex        =   36
      Top             =   120
      Width           =   3645
      _ExtentX        =   6429
      _ExtentY        =   688
      Texto           =   "Cta Fianza"
      EnabledCMAC     =   -1  'True
      EnabledCta      =   -1  'True
      EnabledProd     =   -1  'True
      EnabledAge      =   -1  'True
   End
End
Attribute VB_Name = "frmCFAprobacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'*  APLICACION : Carta Fianza
'*  ARCHIVO : frmCFAprobacion
'*  CREACION: 10/09/2002     - LAYG
'*************************************************************************
'*  RESUMEN: PERMITE APROBAR LA CARTA FIANZA
'*************************************************************************
Option Explicit
Dim vCodCta As String
Dim fpComision As Double
Dim fbComisionTrimestral  As Boolean
Dim objPista As COMManejador.Pista
Dim fvGravamen() As tGarantiaGravamen 'EJVG20150715
Dim nPeriodoMax As Integer 'JOEP20181218 CP
Dim nPeriodoMin As Integer 'JOEP20181218 CP

'*  VALIDACION DE DATOS DEL FORMULARIO ANTES DE GRABAR
Function ValidaDatos() As Boolean

Dim ldVencApr As Date 'FRHU20131126
Dim ldVigeApr As Date 'FRHU20131126
Dim lnPeriodoFecha As Integer 'FRHU20131126
Dim reg As New ADODB.Recordset
Dim lsSQL As String
Dim MonGarant As Currency
Dim loCFValida As COMNCartaFianza.NCOMCartaFianzaValida 'EJVG20150713
Dim loGen As COMDConstSistema.DCOMGeneral 'EJVG20150713
Dim lnTipoCambioFijo As Double, lnValorGarantGrav As Double 'EJVG20150713
Dim lsmensaje As String 'EJVG20150713
    
    'If VerificaMontoAprobCredito(gsCodUser, Mid(ActXCodCta.NroCuenta, 3, 3), CCur(txtMontoApr.Text), Mid(ActXCodCta.NroCuenta, 6, 1), "N") = False Then
    '    MsgBox "No tiene el Nivel de Autorizacion para Aprobar este Credito", vbInformation, "Aviso"
    '    ValidaDatos = False
    '    Exit Function
    'End If
 
    'valida fecha Vencimiento
    If ValidaFecha(TxtFecVenApr.Text) <> "" Then
        MsgBox "No se registro fecha de Vencimiento", vbInformation, "Aviso"
        ValidaDatos = False
        TxtFecVenApr.SetFocus
        Exit Function
    End If
    'verificando Combo
    If cboApoderado.ListIndex = -1 Then
        MsgBox "El Apoderado No ha Sido Seleccionado", vbInformation, "Aviso"
        ValidaDatos = False
        cboApoderado.SetFocus
        Exit Function
    End If
    'FRHU 20140106 RQ13778
    If Not IsNumeric(TxtPeriodoApr) Then
        MsgBox "El periodo tiene que ser numerico", vbInformation, "Aviso"
        TxtPeriodoApr.Text = ""
        ValidaDatos = False
        Exit Function
    End If
    ldVencApr = Format(TxtFecVenApr.Text, "dd/mm/yyyy")
    ldVigeApr = Format(Me.txtfechaAsigApr.Text, "dd/mm/yyyy")
    lnPeriodoFecha = DateDiff("d", ldVigeApr, ldVencApr)
    If lnPeriodoFecha <> Me.TxtPeriodoApr.Text Then
        MsgBox "Presionar Enter en el Periodo para que el numero de dias en el periodo sean iguales a las fechas", vbInformation, "Aviso"
        TxtPeriodoApr.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    'FRHU 20140106 RQ13778
    
    'If TieneGarantias(vCodCta) = False Then
    '    Exit Sub
    'End If
    
    'VERIFICAR QUE MONTO DE GARANTIAS SEA MAYOR QUE MONTO APROBADO
    'MonGarant = VerGravGarant(ActXCodCta.Text)
    'If Val(Format(txtMontoApr.Text, "#0.00")) > MonGarant Then
    '    If MsgBox("Monto de garantias : " & Format(MonGarant, "#0.00") & "es Menor " & _
    '        "Que Monto Aprobado Desea Continuar ", vbInformation + vbYesNo, "Aviso") = 7 Then
    '        txtMontoApr.SetFocus
    '        Exit Function
    '    End If
    'End If
    
    'EJVG20150713 *** Igual que sugerencia
    Set loGen = New COMDConstSistema.DCOMGeneral
        lnTipoCambioFijo = loGen.EmiteTipoCambio(gdFecSis, TCFijoDia)
    Set loGen = Nothing
    Set loCFValida = New COMNCartaFianza.NCOMCartaFianzaValida
        lnValorGarantGrav = loCFValida.nCFGarantiasGravada(vCodCta, lnTipoCambioFijo, lsmensaje)
        If Trim(lsmensaje) <> "" Then
            MsgBox lsmensaje, vbInformation, "Aviso"
            Exit Function
        End If
    Set loCFValida = Nothing
    
    If lnValorGarantGrav = 0 Then
        MsgBox "El crédito no cuenta con Garantías relacionadas", vbInformation, "Aviso"
        Exit Function
    ElseIf CDbl(Format(txtMontoApr.Text, "#0.00")) > lnValorGarantGrav Then
        'VERIFICA QUE MONTO DE GARANTIAS SEA MAYOR QUE MONTO SUGERIDO
        If MsgBox("Monto de Garantias : " & Format(lnValorGarantGrav, "#0.00") & " es Menor " & _
                  "Que Monto Aprobado. Desea Continuar ", vbInformation + vbYesNo, "Aviso") = vbNo Then
            Exit Function
        End If
    End If
    'END EJVG *******

    'joep20181227 CP
    If Not CP_ValidaMsgApr(1) Then
        ValidaDatos = False
        Exit Function
    End If
    If Not CP_ValidaMsgApr(3) Then
        ValidaDatos = False
        Exit Function
    End If
    'joep20181227 CP
    ValidaDatos = True
End Function


'****************************************************************
'*  LIMPIA LOS DATOS DE LA PANTALLA PARA UNA NUEVA APROBACION
'****************************************************************
Sub LimpiaDatos()
    ActXCodCta.Enabled = True
    ActXCodCta.NroCuenta = fgIniciaAxCuentaCF
    lblNomcli.Caption = ""
    lblCodcli.Caption = ""
    lblRazSoc.Caption = ""
    lblActiv.Caption = ""
    lblCodAcreedor.Caption = ""
    lblNomAcreedor.Caption = ""
    lblCodAvalado.Caption = ""
    lblNomAvalado.Caption = ""
    lblTipoCF.Caption = ""
    lblMontoSug.Caption = ""
    lblFecVencSug.Caption = ""
    txtFinalidad.Text = ""
    txtMontoApr.Text = ""
    TxtFecVenApr.Text = "__/__/____"
    lblAnalista.Caption = ""
    cboApoderado.ListIndex = -1
    CboModalidad.ListIndex = -1
    cmdGrabar.Enabled = False
    fraDatos.Enabled = False
    lblComision.Caption = ""
    cmdGenerarPDF.Enabled = False 'WIOR 20120613
    TxtPeriodoApr = "" 'FRHU 20131126
    txtfechaAsigApr = "__/__/____" 'FRHU 20131126
End Sub

Private Sub CargaApoderados()
Dim R As New ADODB.Recordset
Dim sSql As String
Dim lcPers As COMDPersona.DCOMPersonas
Dim oGen As COMDConstSistema.DCOMGeneral
Dim sApoderados As String
Dim rs As New ADODB.Recordset
On Error GoTo ERRORCargaApoderado

    Set oGen = New COMDConstSistema.DCOMGeneral
        sApoderados = oGen.LeeConstSistema(gConstSistRHCargoCodApoderados)
    Set oGen = Nothing
    
    Set lcPers = New COMDPersona.DCOMPersonas
      Set rs = lcPers.ObtenerApoderados(sApoderados)
    Set lcPers = Nothing
    
    cboApoderado.Clear
    Do While Not rs.EOF
        cboApoderado.AddItem PstaNombre(rs!cPersNombre) & Space(100) & rs!cPersCod
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    Exit Sub
ERRORCargaApoderado:
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub

'PROCEDIMIENTO QUE CARGA LOS DATOS QUE SE REQUIEREN PARA EL FORMULARIO
Private Sub CargaDatosApr(ByVal psCta As String)

Dim oCF As COMDCartaFianza.DCOMCartaFianza
Dim R As ADODB.Recordset
'Dim loConstante As COMDConstantes.DCOMConstantes
Dim loCFCalculo As COMNCartaFianza.NCOMCartaFianzaCalculos
Dim loCFValida As COMNCartaFianza.NCOMCartaFianzaValida
Dim lbTienePermiso As Boolean

txtMontoApr.Text = "0.00"
TxtFecVenApr.Text = "__/__/____"
ActXCodCta.Enabled = False
    
    Set oCF = New COMDCartaFianza.DCOMCartaFianza
    Set R = oCF.RecuperaCartaFianzaAprobacion(psCta)
    Set oCF = Nothing
    If Not R.BOF And Not R.EOF Then
        lblCodcli.Caption = R!cPersCod
        lblNomcli.Caption = PstaNombre(R!cPersNombre)
    
        lblCodAcreedor.Caption = R!cPersAcreedor
        lblNomAcreedor.Caption = PstaNombre(R!cPersNomAcre)
        'MADM 20111020
        lblCodAvalado.Caption = IIf(IsNull(R!cAvalCod), "", R!cAvalCod)
         If R!cAvalNombre <> "" Then
            lblNomAvalado.Caption = PstaNombre(R!cAvalNombre)
        End If
        'END MADM
        'MAVM 20100605 BAS II
        'If Mid(Trim(psCta), 9, 1) = "1" Then
            lblTipoCF = IIf(IsNull(R!cConsDescripcion), "", R!cConsDescripcion) 'IIf(Mid(Trim(psCta), 9, 1) = "1", "COMERCIALES - SOLES", "COMERCIALES - DOLARES")
        'ElseIf Mid(Trim(psCta), 9, 1) = "2" Then
            'lblTipoCF = IIf(Mid(Trim(psCta), 9, 1) = "1", "MICROEMPRESA - SOLES", "MICROEMPRESA - DOLARES")
        'End If
        lblAnalista.Caption = IIf(IsNull(R!cAnalista), "", R!cAnalista)
        lblMontoSug.Caption = IIf(IsNull(R!nMontoSug), "", Format(R!nMontoSug, "#0.00"))
        lblFecVencSug.Caption = IIf(IsNull(R!dVencSug), "", Format(R!dVencSug, "dd/mm/yyyy"))
        txtFinalidad.Text = IIf(IsNull(R!cfinalidad), "", R!cfinalidad)
        
        '**Inicio,DAOR 20070110, mostrar comisión de carta fianza
        '**Las siguientes lineas de código fueron obtenidos de la pantalla de sugerenia
        Set loCFCalculo = New COMNCartaFianza.NCOMCartaFianzaCalculos
            If fbComisionTrimestral = False Then ' Caja Trujillo
                lblComision = Format(loCFCalculo.nCalculaComisionCF(R!nMontoSug, DateDiff("d", gdFecSis, R!dVencSug), fpComision, Mid(psCta, 9, 1)), "#,##0.00")
            Else  ' Caja Metropolitana
                lblComision = Format(loCFCalculo.nCalculaComisionTrimestralCF(R!nMontoSug, DateDiff("d", gdFecSis, R!dVencSug), R!nModalidad, Mid(Trim(psCta), 9, 1), psCta, 6), "#,###0.00")
            End If
        Set loCFCalculo = Nothing
        '**Fin************************************************************
        
        txtMontoApr.Text = IIf(IsNull(R!nMontoSug), "", Format(R!nMontoSug, "#0.00"))
        TxtFecVenApr.Text = IIf(IsNull(R!dVencSug), "", Format(R!dVencSug, "dd/mm/yyyy"))
        txtfechaAsigApr.Text = IIf(IsNull(R!dAsignacion), "", Format(R!dAsignacion, "dd/mm/yyyy")) 'FRHU20131126
        
         Call CP_CargaDatosApr 'JOEP20181218 CP
         
        TxtPeriodoApr.Text = CDate(IIf(IsNull(R!dVencSug), 0, Format(R!dVencSug, "dd/mm/yyyy"))) - CDate(IIf(IsNull(R!dVencSug), 0, Format(R!dAsignacion, "dd/mm/yyyy"))) 'FRHU20131126
        
        'lblRazSoc.Caption = IIf(IsNull(R!cRazSocDescrip), "", R!cRazSocDescrip) 'LUCV20160919, Comentó. Según ERS004-2016
        lblActiv.Caption = IIf(IsNull(R!cCIIUdescripcion), "", R!cCIIUdescripcion)
    
        '** Modalidad
        CboModalidad.ListIndex = IndiceListaCombo(CboModalidad, R!nModalidad)
        
        'JOEP20181220 CP
        TxtPeriodoApr.Text = R!nPeriodo
        If R!nModalidad = 13 Then
            frModOtrs.Visible = True
            Frame4.Height = 1935
            Frame4.Width = 7095
            Frame4.Left = 100
            frmCFEmision.Left = 100
            frmCFEmision.Top = 7320
            Me.Width = 7365
            Me.Height = 8445
            txtModOtrs.Text = R!OtrsModalidades
        Else
            frModOtrs.Visible = False
            Frame4.Height = 1350
            frmCFEmision.Left = 100
            frmCFEmision.Top = 6800
            Me.Width = 7365
            Me.Height = 7995
            txtModOtrs.Text = ""
        End If
    'JOEP20181220 CP
        
        CargaApoderados
'-------------- Comentado por AVMM  28-07-2006 -------------------
'------------------------------------------------------------------
'        Set loCFValida = New COMNCartaFianza.NCOMCartaFianzaValida
'            lbTienePermiso = loCFValida.nCFPermisoAprobacion(gsCodPersUser, Mid(psCta, 6, 3))
'        Set loCFValida = Nothing
'        If lbTienePermiso = False Then
'            MsgBox "NO TIENE EL PERMISO NECESARIO PARA APROBAR ESTE TIPO DE PRODUCTO "
'            fraDatos.Enabled = False
'            cmdGrabar.Enabled = False
'            Exit Sub
'        End If
        fraDatos.Enabled = True
        cmdGrabar.Enabled = True
        cmdGenerarPDF.Enabled = True 'WIOR 20120613
    End If

End Sub

Private Sub ActXCodCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(Trim(ActXCodCta.NroCuenta)) > 0 Then
            Call CargaDatosApr(ActXCodCta.NroCuenta)
        Else
            Call LimpiaDatos
        End If
    End If
End Sub

Private Sub cboApoderado_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdGrabar.SetFocus
End If
End Sub

Private Sub cboModalidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cboApoderado.SetFocus
    End If
End Sub

Private Sub cmdCancelar_Click()
    Call LimpiaDatos
End Sub

Private Sub cmdExaminar_Click()
Dim lsCta As String
    'MAVM 20100605 BAS II
    lsCta = frmCFPersEstado.Inicio(Array(gColocEstSug, gColocEstSug), "Aprobacion de Carta Fianza", Array(gColCFComercial, gColCFPYME, gColCFTpoProducto))
    If Len(Trim(lsCta)) > 0 Then
        ActXCodCta.NroCuenta = lsCta
        Call CargaDatosApr(lsCta)
    Else
        Call LimpiaDatos
    End If
End Sub

Private Sub cmdGrabar_Click()
Dim loNCartaFianza As COMNCartaFianza.NCOMCartaFianza
Dim loContFunct As COMNContabilidad.NCOMContFunciones
Dim lsMovNro As String
Dim lsFechaHoraGrab As String

Dim lnMontoApr As Currency
Dim ldVencApr As Date
Dim ldVigeApr As Date 'FRHU20131126

vCodCta = ActXCodCta.NroCuenta
lnMontoApr = Format(txtMontoApr.Text, "#0.00")
ldVencApr = Format(TxtFecVenApr.Text, "dd/mm/yyyy")
ldVigeApr = Format(Me.txtfechaAsigApr.Text, "dd/mm/yyyy") 'FRHU20131126

If ValidaDatos = False Then
    Exit Sub
End If

If Not RecalcularCoberturaGarantias(vCodCta, False, "514", "CARTA FIAMZA", CCur(txtMontoApr.Text), fvGravamen) Then Exit Sub   'EJVG20150715
If MsgBox("Desea Guardar Aprobación de Carta Fianza", vbInformation + vbYesNo, "Aprobacion Carta Fianza") = vbYes Then
    'EJVG20150715 ***
    If Not IsArray(fvGravamen) Then
        ReDim fvGravamen(0)
    End If
    'END EJVG *******
    'Genera el Mov Nro
    Set loContFunct = New COMNContabilidad.NCOMContFunciones
        lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Set loContFunct = Nothing
    lsFechaHoraGrab = fgFechaHoraGrab(lsMovNro)
    
    Set loNCartaFianza = New COMNCartaFianza.NCOMCartaFianza
        'FRHU20131126 - Se agrego el parametro: ldVigeApr
        Call loNCartaFianza.nCFAprobacion(vCodCta, lsFechaHoraGrab, ldVencApr, lnMontoApr, Right(Trim(cboApoderado), 13), Trim(txtFinalidad), Trim(lblCodAvalado.Caption), ldVigeApr, , Trim(txtModOtrs.Text)) 'JOEP20181222 CP Trim(txtModOtrs.Text)
    Set loNCartaFianza = Nothing
    
    'MAVM 20100621
    objPista.InsertarPista gsOpeCod, lsMovNro, gsCodPersUser, GetMaquinaUsuario, gInsertar, "Aprobacion de CF", vCodCta, gCodigoCuenta
            
    cmdGrabar.Enabled = False
    cmdGenerarPDF.Enabled = False 'WIOR 20120613
    LimpiaDatos
End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub ActxCodCta_keypressEnter()
'    vCodCta = ActXCodCta.NroCuenta
'    If Len(gsCodCred) > 0 Then
'        Call CargaDatosApr
'        ActXCodCta.Enabled = False
'        txtMontoApr.Enabled = True
'    Else
'        Call LimpiaDatos
'    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 And ActXCodCta.Enabled = True Then 'F12
        Dim sCuenta As String
        sCuenta = frmValTarCodAnt.inicia(gColCFComercial, False)
        If sCuenta <> "" Then
            ActXCodCta.NroCuenta = sCuenta
            ActXCodCta.SetFocusCuenta
        End If
    End If
End Sub

Private Sub Form_Load()
    Call LimpiaDatos
    Call CargaParametros 'DAOR 20070110
    
    Call CP_CargaComboxApr(49000) 'JOEP20181218 CP
    
    'Call CargaComboConstante(gColCFModalidad, CboModalidad)'Comento JOEP20181218 CP
    
    CboModalidad.ListIndex = 0
    Call CargaApoderados
    Me.Icon = LoadPicture(App.Path & gsRutaIcono)
    'MAVM 20100621 ***
    Set objPista = New COMManejador.Pista
    gsOpeCod = gCredAprobacionCF
    '***
'joep20181221 CP
    frModOtrs.Visible = False
    Frame4.Height = 1350
    frmCFEmision.Left = 100
    frmCFEmision.Top = 6800
    Me.Width = 7365
    Me.Height = 7995
    txtModOtrs.Text = ""
    CboModalidad.Enabled = False
'joep20181221 CP

End Sub

'Carga los Parametros
Private Sub CargaParametros()
Dim loParam As COMDColocPig.DCOMColPCalculos 'DColPCalculos
Dim lcCons As COMDConstSistema.DCOMConstSistema

Dim lr As New ADODB.Recordset

Set loParam = New COMDColocPig.DCOMColPCalculos
    fpComision = loParam.dObtieneColocParametro(4001)
Set loParam = Nothing

Set lcCons = New COMDConstSistema.DCOMConstSistema
    Set lr = lcCons.ObtenerVarSistema()
        fbComisionTrimestral = IIf(lr!nConsSisValor = 2, True, False)
    Set lr = Nothing
Set lcCons = Nothing
End Sub

Private Sub TxtFecVenApr_KeyPress(KeyAscii As Integer)
If IsDate(TxtFecVenApr.Text) Then
    If CDate(Format(TxtFecVenApr.Text, "dd/mm/yyyy")) < CDate(Format(gdFecSis, "dd/mm/yyyy")) Then
        MsgBox "Fecha de Vencimiento no puede ser anterior a la fecha actual", vbInformation, "Aviso"
        TxtFecVenApr.SetFocus
        Exit Sub
    Else
        If CboModalidad.Enabled = True Then 'JOEP20190227 CP
            CboModalidad.SetFocus
        End If
    End If
Else
    TxtFecVenApr.SetFocus
    Exit Sub
End If
End Sub

Private Sub txtMontoApr_GotFocus()
    txtMontoApr.SelStart = 0
    txtMontoApr.SelLength = Len(txtMontoApr.Text)
End Sub

Private Sub txtMontoApr_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtMontoApr, KeyAscii)
    If KeyAscii = 13 Then
        If TxtFecVenApr.Enabled Then
            TxtFecVenApr.SetFocus
        Else
            TxtFecVenApr.SetFocus
        End If
    End If
End Sub

Private Sub txtMontoApr_LostFocus()
    If Trim(txtMontoApr.Text) = "" Then
        txtMontoApr.Text = "0.00"
    Else
        txtMontoApr.Text = Format(txtMontoApr.Text, "#,#0.00")
    End If
End Sub
'WIOR 20120613*************************************************************
Private Sub cmdGenerarPDF_Click()
On Error GoTo ErrorGenerarPdf
vCodCta = ActXCodCta.NroCuenta
If ValidaDatos = False Then
    Exit Sub
End If
Call ImprimirPDF(vCodCta, IIf(Me.lblCodAvalado.Caption = "", False, True), "1", 1)
MsgBox "Archivo Previo Generado Satisfacoriamente.", vbInformation, "Aviso"
Exit Sub
ErrorGenerarPdf:
    MsgBox Err.Description, vbInformation, "Aviso"
End Sub

Private Sub ImprimirPDF(ByVal psCodCta As String, ByVal pbAvalado As Boolean, ByVal psNumFolio As String, ByVal nTipo As Integer)
    On Error GoTo ErrorImprimirPDF
    Dim loRs As COMNCartaFianza.NCOMCartaFianzaValida
    Dim lrDataCR As ADODB.Recordset
    Dim dfechaini As Date
    
    Dim nPoliza As Long
    Dim nCFPoliza As Long
    Set loRs = New COMNCartaFianza.NCOMCartaFianzaValida
    nCFPoliza = psNumFolio

    
    Dim oDoc  As cPDF
    Dim sCadena As String
    Dim sParrafo1 As String
    Dim sParrafo2 As String
    Dim sParrafo3 As String
    Dim sParrafo4 As String
    Dim nTamano As Integer
    Dim nValidar As Double
    Dim nTop As Integer
    
    Dim sFechaActual As String
    Dim sSenores As String
    Dim sAval As String
    Dim sSolicitante As String
    Dim sMonto As String
    Dim sModalidad As String
    Dim sFinalidad As String
    
    Dim dfechafin As Date
    Dim dfechainipdf As Date 'FRHU20131126
    
    Dim sVencimiento As String
    Dim sVigenciapdf As String 'FRHU20131126
    
    Dim sDireccion As String
    Dim lnPosicion As Integer
    
    Set oDoc = New cPDF
    
    oDoc.Author = gsCodUser
    oDoc.Creator = "SICMACT - Negocio"
    oDoc.Producer = "Caja Municipal de Ahorros y Crédito de Maynas S.A."
    oDoc.Subject = "Carta Fianza Nº " & psCodCta
    oDoc.Title = "Carta Fianza Nº " & psCodCta

    If Not oDoc.PDFCreate(App.Path & "\Spooler\" & IIf(nTipo = 1, "Previo", "") & "Aprobacion" & psCodCta & "_" & Format(gdFecSis, "YYYYMMDD") & "_" & Format(Time, "hhmmss") & ".pdf") Then
        Exit Sub
    End If
    
    oDoc.Fonts.Add "F1", "Times New Roman", TrueType, , WinAnsiEncoding 'FRHU20131126 DE BOLD A NORMAL
    oDoc.Fonts.Add "F2", "Times New Roman", TrueType, Bold, WinAnsiEncoding 'FRHU20131126
    oDoc.NewPage A4_Vertical
    
    sFechaActual = Format(gdFecSis, "dd") & " de " & Format(gdFecSis, "mmmm") & " del " & Format(gdFecSis, "yyyy")
    sSenores = PstaNombre(lblNomAcreedor, True)
    If pbAvalado Then
        sAval = PstaNombre(lblNomAvalado.Caption, True)
    End If
    sSolicitante = PstaNombre(lblNomcli, True)
    sMonto = IIf(Mid(psCodCta, 9, 1) = "1", "S/ ", "$ ") & Format(txtMontoApr.Text, "#,###0.00") & " " & "(" & UCase(NumLet(txtMontoApr.Text)) & IIf(Mid(psCodCta, 9, 1) = "2", "", " Y " & IIf(InStr(1, txtMontoApr.Text, ".") = 0, "00", Mid(txtMontoApr.Text, InStr(1, txtMontoApr.Text, ".") + 1, 2)) & "/100 ") & IIf(Mid(psCodCta, 9, 1) = "1", "SOLES)", " US DOLARES)") 'EAAS20181128 SE CAMBIO DE NUEVO SOLES A SOLES
    'sModalidad = Trim(Mid(cboModalidad.Text, 1, Len(cboModalidad.Text) - 3))'Comento JOEP20181221 CP
    sModalidad = IIf(Trim(Right(CboModalidad.Text, 9)) = 13, UCase(Trim(txtModOtrs.Text)), Trim(Mid(CboModalidad.Text, 1, Len(CboModalidad.Text) - 3)))    'JOEP20181221 CP
    sFinalidad = Trim(txtFinalidad.Text)
    
    'FRHU20131126
    dfechainipdf = CDate(Me.txtfechaAsigApr.Text)
    sVigenciapdf = Format(dfechainipdf, "dd") & " de " & Format(dfechainipdf, "mmmm") & " del " & Format(dfechainipdf, "yyyy")
    'FIN FRHU20131126
    
    dfechafin = CDate(TxtFecVenApr.Text)
    sVencimiento = Format(dfechafin, "dd") & " de " & Format(dfechafin, "mmmm") & " del " & Format(dfechafin, "yyyy")
    sDireccion = loRs.Get_Agencia_CF(psCodCta)
    lnPosicion = InStr(sDireccion, "(")
    sDireccion = Left(sDireccion, lnPosicion - 2)
    
    oDoc.WTextBox 70, 50, 10, 450, Left(psCodCta, 3) & "-" & Mid(psCodCta, 4, 2) & "-" & Mid(psCodCta, 6, 3) & "-" & Mid(psCodCta, 9, 10), "F1", 12, hRight
    oDoc.WTextBox 120, 50, 10, 450, "CARTA FIANZA N° " & Format(nCFPoliza, "0000000"), "F1", 12, hCenter
    oDoc.WTextBox 170, 50, 10, 450, sFechaActual, "F2", 12, hRight 'FRHU20131126 F1 a F2
    oDoc.WTextBox 220, 50, 10, 450, "Señores:", "F1", 12, hLeft
    oDoc.WTextBox 232, 50, 10, 450, sSenores, "F2", 12, hLeft 'FRHU20131126 F1 a F2
    oDoc.WTextBox 260, 50, 10, 450, "Ciudad.-", "F1", 12, hLeft
    oDoc.WTextBox 280, 50, 10, 450, "Muy Señores Nuestros:", "F1", 12, hLeft
    
    sAval = " garantizando a " & sAval
    nTop = 270
    
    sParrafo1 = "A solicitud de " & sSolicitante & ", otorgamos por el presente " & _
                "documento una fianza solidaria, irrevocable, incondicional, de " & _
                "ejecución inmediata, con renuncia expresa al beneficio de " & _
                "excusión e indivisible, a favor de ustedes" & IIf(pbAvalado = True, sAval, "") & _
                ", hasta por la suma de " & sMonto & ", a fin de garantizar " & _
                "la Carta Fianza por " & sModalidad & ", objeto del proceso: " & sFinalidad & "."
    nTamano = Len(sParrafo1)
    nValidar = nTamano / 75
    nTamano = CInt(IIf(CInt(nValidar) > nValidar, CInt(nValidar), CInt(nValidar) + 1))
 
    oDoc.WTextBox nTop, 0, nTamano * 20, 580, String(20, "-") & " " & sParrafo1, "F1", 11, hjustify, , , , , , 50

    
    oDoc.WTextBox nTop + 50, 50, 10, 75, "", "F1", 10, hjustify, , vbWhite, 1, vbWhite, True
    nTop = nTop + (nTamano * 10) + 12
    
    sParrafo2 = "Dejamos claramente establecido que la presente " & String(1, vbTab) & "Carta " & String(1, vbTab) & "Fianza no " & _
                "podrá ser usada " & String(1, vbTab) & "para operaciones comprendidas en la prohibición " & _
                "indicada en el inciso ''5'' del Articulo 217 de la " & String(1, vbTab) & "Ley  26702, Ley " & _
                "General del " & String(1, vbTab) & "Sistema " & String(1, vbTab) & "Financiero y del Sistema de Seguros y Orgánica " & _
                "de la Superintendencia de --- Banca y Seguros."
    nTamano = Len(sParrafo2)
    nValidar = nTamano / 80
    nTamano = CInt(IIf(CInt(nValidar) > nValidar, CInt(nValidar), CInt(nValidar) + 1))
    oDoc.WTextBox nTop, 0, nTamano * 10, 580, String(20, "-") & " " & sParrafo2, "F1", 11, hjustify, , , , , , 50
    oDoc.WTextBox nTop + 50, 50, 10, 75, "", "F1", 10, hjustify, , vbWhite, 1, vbWhite, True
    oDoc.WTextBox nTop + 75, 520, 10, 20, "", "F1", 10, hjustify, , vbWhite, 1, vbWhite, True
    
    nTop = nTop + (nTamano * 12) + 12
    'JGPA20190614 Cambio razón social según Memorandum Nº 1037-2019-GM-DI/CMACM
    sParrafo3 = "Por efecto de este compromiso la CAJA MUNICIPAL DE AHORRO Y CRÉDITO MAYNAS S.A. " & _
                        "asume con su fiado las responsabilidades en que éste llegara a " & _
                        "incurrir siempre que el " & String(1, vbTab) & "monto de las  mismas  no " & String(1, vbTab) & "exceda por ningún " & _
                        "motivo de la suma antes mencionada y que estén estrictamente " & _
                        "vinculadas al cumplimiento de lo arriba indicado."
    nTamano = Len(sParrafo3)
    nValidar = nTamano / 80
    nTamano = CInt(IIf(CInt(nValidar) > nValidar, CInt(nValidar), CInt(nValidar) + 1))
    oDoc.WTextBox nTop, 0, nTamano * 10, 580, String(20, "-") & " " & sParrafo3, "F1", 11, hjustify, , , , , , 50
    oDoc.WTextBox nTop + 50, 50, 10, 75, "", "F1", 10, hjustify, , vbWhite, 1, vbWhite, True
    nTop = nTop + (nTamano * 12) + 12
    
    'FRHU20131126
    sParrafo4 = "La presente garantía rige a partir del " & sVigenciapdf & " y vencerá " & _
                        "el " & sVencimiento & ". Cualquier  reclamo en virtud de esta " & _
                        "garantía deberá ceñirse estrictamente a lo estipulado por " & _
                        "el Art. 1898 del Código Civil y deberá ser formulado por vía " & _
                        "notarial y en nuestra oficina ubicada en " & sDireccion & "."
    nTamano = Len(sParrafo4)
    nValidar = nTamano / 80
    nTamano = CInt(IIf(CInt(nValidar) > nValidar, CInt(nValidar), CInt(nValidar) + 1))
    oDoc.WTextBox nTop, 0, nTamano * 10, 580, String(20, "-") & " " & sParrafo4, "F1", 11, hjustify, , , , , , 50
    oDoc.WTextBox nTop + 50, 50, 10, 75, "", "F1", 10, hjustify, , vbWhite, 1, vbWhite, True
    nTop = nTop + (nTamano * 13) + 80

    oDoc.WTextBox nTop, 50, 10, 450, "Atentamente,", "F1", 12, hCenter, vMiddle, , , , False
    oDoc.WTextBox nTop + 12, 50, 10, 450, "CAJA MUNICIPAL DE AHORRO Y CRÉDITO MAYNAS S.A.", "F1", 12, hCenter, vMiddle, , , , False 'JGPA20190614 Cambio razón social según Memorandum Nº 1037-2019-GM-DI/CMACM

    oDoc.PDFClose
    oDoc.Show
    Exit Sub
ErrorImprimirPDF:
    MsgBox Err.Description, vbInformation, "Aviso"
End Sub
'WIOR FIN ******************************************************************
'***** FRHU 20131126
Private Sub txtfechaAsigApr_GotFocus()
    fEnfoque txtfechaAsigApr
End Sub
Private Sub txtfechaAsigApr_KeyPress(KeyAscii As Integer)
    Dim dfechaini As Date
    Dim dfechafin As Date
    Dim dfechafin3Meses As Date 'WIOR 20140130
    dfechaini = (gdFecSis - CInt(Mid(gdFecSis, 1, 2))) + 1
    'dfechafin = DateAdd("m", 1, gdFecSis) - CInt(Mid(gdFecSis, 1, 2))'WIOR 20140130 COMENTO
    'WIOR 20140130 ***************************
    dfechafin = obtenerFechaFinMes(Month(dfechaini), Year(dfechaini))
    dfechafin3Meses = DateAdd("m", 4, dfechaini) - 1
    'WIOR FIN ********************************
    If KeyAscii = 13 Then
        If IsDate(Me.txtfechaAsigApr.Text) Then
            If Me.txtfechaAsigApr <= gdFecSis Then
                If CDate(Format(txtfechaAsigApr.Text, "dd/mm/yyyy")) >= CDate(Format(dfechaini, "dd/mm/yyyy")) And CDate(Format(txtfechaAsigApr.Text, "dd/mm/yyyy")) <= CDate(Format(dfechafin, "dd/mm/yyyy")) Then
                    Call CP_CargaDatosApr 'JOEP20181218 CP
                    TxtPeriodoApr.SetFocus
                Else
                    MsgBox "Si la fecha de vigencia es anterior a la fecha actual del sistema; debe estar dentro del mismo mes", vbInformation, "Aviso"
                    Me.txtfechaAsigApr.SetFocus
                    Exit Sub
                End If
            Else
                'If CDate(Format(Me.txtfechaAsigApr, "dd/mm/yyyy")) <= CDate(Format(DateAdd("m", 3, dfechafin), "dd/mm/yyyy")) Then'WIOR 20140130 COMENTO
                If CDate(Format(Me.txtfechaAsigApr, "dd/mm/yyyy")) <= CDate(Format(dfechafin3Meses, "dd/mm/yyyy")) Then 'WIOR 20140130
                    Me.TxtPeriodoApr.SetFocus
                Else
                    MsgBox "Si la fecha de vigencia es posterior a la fecha actual del sistema, debera estar dentro de los siguientes 3 meses inclusive", vbInformation, "Aviso"
                    Me.txtfechaAsigApr.SetFocus
                    Exit Sub
                End If
            End If
        Else
            MsgBox "Escribe un Fecha Correcta", vbInformation, "Aviso"
            Me.txtfechaAsigApr.SetFocus
            Exit Sub
        End If
    End If
End Sub
Private Sub txtfechaAsigApr_LostFocus()
    Dim dfechaini As Date
    Dim dfechafin As Date
    Dim dfechafin3Meses As Date 'WIOR 20140130
    dfechaini = (gdFecSis - CInt(Mid(gdFecSis, 1, 2))) + 1
    'dfechafin = DateAdd("m", 1, gdFecSis) - CInt(Mid(gdFecSis, 1, 2))'WIOR 20140130 COMENTO
    'WIOR 20140130 ***************************
    dfechafin = obtenerFechaFinMes(Month(dfechaini), Year(dfechaini))
    dfechafin3Meses = DateAdd("m", 4, dfechaini) - 1
    'WIOR FIN ********************************
    'If IsDate(Me.txtfechaAsigApr.Text) Then'WIOR 20140130 COMENTO
    If IsDate(Me.txtfechaAsigApr.Text) Then
        If Me.txtfechaAsigApr <= gdFecSis Then
            If CDate(Format(txtfechaAsigApr.Text, "dd/mm/yyyy")) >= CDate(Format(dfechaini, "dd/mm/yyyy")) And CDate(Format(txtfechaAsigApr.Text, "dd/mm/yyyy")) <= CDate(Format(dfechafin, "dd/mm/yyyy")) Then
                TxtPeriodoApr.SetFocus
            Else
                MsgBox "Si la fecha de vigencia es anterior a la fecha actual del sistema; debe estar dentro del mismo mes", vbInformation, "Aviso"
                Me.txtfechaAsigApr.SetFocus
                Exit Sub
            End If
        Else
            'If CDate(Format(Me.txtfechaAsigApr, "dd/mm/yyyy")) <= CDate(Format(DateAdd("m", 3, dfechafin), "dd/mm/yyyy")) Then'WIOR 20140130 COMENTO
            If CDate(Format(Me.txtfechaAsigApr, "dd/mm/yyyy")) <= CDate(Format(dfechafin3Meses, "dd/mm/yyyy")) Then 'WIOR 20140130
                Me.TxtPeriodoApr.SetFocus
            Else
                MsgBox "Si la fecha de vigencia es posterior a la fecha actual del sistema, debera estar dentro de los siguientes 3 meses inclusive", vbInformation, "Aviso"
                Me.txtfechaAsigApr.SetFocus
                Exit Sub
            End If
        End If
    Else
        MsgBox "Escribe un Fecha Correcta", vbInformation, "Aviso"
        Me.txtfechaAsigApr.SetFocus
        Exit Sub
    End If
    'Else'WIOR 20140130 COMENTO
    '    MsgBox "Escribe un Fecha Correcta", vbInformation, "Aviso"
    '    Me.txtfechaAsigApr.SetFocus
    '    Exit Sub
    'End If
End Sub
Private Sub TxtPeriodoApr_Change()
    If IsNumeric(TxtPeriodoApr) Then
    If Not CP_ValidaMsgApr(2) Then Exit Sub
        TxtFecVenApr.Text = CDate(txtfechaAsigApr.Text) + CInt(TxtPeriodoApr.Text)
    End If
End Sub
Private Sub TxtPeriodoApr_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If IsNumeric(TxtPeriodoApr) Then
        If Not CP_ValidaMsgApr(1) Then Exit Sub
            TxtFecVenApr.Text = CDate(txtfechaAsigApr.Text) + CInt(TxtPeriodoApr.Text)
            If CboModalidad.Enabled = True Then 'JOEP20190227 CP
                CboModalidad.SetFocus
            End If
        Else
            MsgBox "El periodo tiene que ser numerico", vbInformation, "Aviso"
            TxtPeriodoApr.Text = ""
        End If
    End If
End Sub
'***** FIN FRHU 20131126

'JOEP20181221 CP
Private Sub CP_CargaComboxApr(ByVal nParCod As Long)
Dim objCatalogoLlenaCombox As COMDCredito.DCOMCredito
Dim rsCatalogoCombox As ADODB.Recordset
Set objCatalogoLlenaCombox = New COMDCredito.DCOMCredito
Set rsCatalogoCombox = objCatalogoLlenaCombox.getCatalogoCombo("514", nParCod)

If Not (rsCatalogoCombox.BOF And rsCatalogoCombox.EOF) Then
    If nParCod = 49000 Then
        Call Llenar_Combo_con_Recordset(rsCatalogoCombox, CboModalidad)
        Call CambiaTamañoCombo(CboModalidad, 300)
    End If
End If

End Sub

Private Sub CP_CargaDatosApr()
Dim oDCred As COMDCredito.DCOMCredito
Dim rsDefaut As ADODB.Recordset
Set oDCred = New COMDCredito.DCOMCredito

Set rsDefaut = oDCred.CatalogoProDefaut(514, 7000)

If Not (rsDefaut.BOF And rsDefaut.EOF) Then
    'TxtPeriodoApr.Text = rsDefaut!MinPlazo
    nPeriodoMin = rsDefaut!MinPlazo
    nPeriodoMax = rsDefaut!MaxPlazo
End If

End Sub

Private Sub cboModalidad_Click()
If Trim(Right(CboModalidad.Text, 9)) = "13" Then
    frModOtrs.Visible = True
    Frame4.Height = 1935
    Frame4.Width = 7095
    Frame4.Left = 100
    frmCFEmision.Left = 100
    frmCFEmision.Top = 7320
    Me.Width = 7365
    Me.Height = 8445
Else
    frModOtrs.Visible = False
    Frame4.Height = 1350
    frmCFEmision.Left = 100
    frmCFEmision.Top = 6800
    Me.Width = 7365
    Me.Height = 7995
    txtModOtrs.Text = ""
End If
End Sub
Private Function CP_ValidaMsgApr(ByVal nTpOp As Integer) As Boolean
CP_ValidaMsgApr = True
Select Case nTpOp
    Case 1
        If CInt(TxtPeriodoApr.Text) < nPeriodoMin Then
            MsgBox "El Periodo minimo es " & nPeriodoMin & " dias", vbInformation, "Aviso"
            TxtPeriodoApr.Text = nPeriodoMin
            CP_ValidaMsgApr = False
            Exit Function
        End If
        If CInt(TxtPeriodoApr.Text) > nPeriodoMax Then
            MsgBox "El Periodo maximo es " & nPeriodoMax & " dias", vbInformation, "Aviso"
            TxtPeriodoApr.Text = nPeriodoMax
            CP_ValidaMsgApr = False
            Exit Function
        End If
    Case 2
        If (txtfechaAsigApr.Text = "" Or txtfechaAsigApr.Text = "__/__/____") Then
            MsgBox "Ingrese la Fecha de Asignacion", vbInformation, "Aviso"
            txtfechaAsigApr.SetFocus
            TxtPeriodoApr.Text = 0
            CP_ValidaMsgApr = False
            Exit Function
        End If
    Case 3
        If frModOtrs.Visible = True And txtModOtrs.Text = "" Then
            MsgBox "Registre Otras Modalidades", vbInformation, "Aviso"
            txtModOtrs.SetFocus
            CP_ValidaMsgApr = False
            Exit Function
        End If
End Select
End Function
'JOEP20181221 CP
