VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmColPRemateRegVenta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Crédito Pignoraticio -  Registrar Venta en Remate"
   ClientHeight    =   7620
   ClientLeft      =   855
   ClientTop       =   1230
   ClientWidth     =   8070
   Icon            =   "frmColPRemateRegVenta.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   8070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   5640
      TabIndex        =   3
      Top             =   7080
      Width           =   975
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4440
      TabIndex        =   2
      Top             =   7080
      Width           =   975
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6840
      TabIndex        =   4
      Top             =   7080
      Width           =   975
   End
   Begin VB.Frame fraContenedor 
      Height          =   6900
      Index           =   0
      Left            =   135
      TabIndex        =   5
      Top             =   75
      Width           =   7740
      Begin VB.Frame fraContenedor 
         Caption         =   "Adjudicatario"
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
         Height          =   930
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   4200
         Width           =   7350
         Begin VB.CommandButton cmdBuscar 
            Caption         =   "&Buscar..."
            Enabled         =   0   'False
            Height          =   300
            Left            =   2520
            TabIndex        =   31
            Top             =   120
            Width           =   930
         End
         Begin VB.TextBox txtNatAdj 
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
            Left            =   6045
            TabIndex        =   10
            Tag             =   "txtDocumento"
            Top             =   180
            Width           =   1080
         End
         Begin VB.TextBox txtTriAdj 
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
            Left            =   6030
            TabIndex        =   9
            Tag             =   "txtTributario"
            Top             =   510
            Width           =   1080
         End
         Begin VB.TextBox txtCodAdj 
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
            Height          =   330
            Left            =   825
            TabIndex        =   8
            Tag             =   "txtcodigo"
            Top             =   165
            Width           =   1215
         End
         Begin VB.TextBox txtNomAdj 
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
            Left            =   825
            TabIndex        =   7
            Tag             =   "txtnombre"
            Top             =   510
            Width           =   3780
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Doc.Tributario : "
            Height          =   255
            Index           =   3
            Left            =   4890
            TabIndex        =   14
            Top             =   555
            Width           =   1110
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Doc.Natural : "
            Height          =   255
            Index           =   2
            Left            =   4875
            TabIndex        =   13
            Top             =   210
            Width           =   1110
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Código :"
            Height          =   225
            Index           =   8
            Left            =   120
            TabIndex        =   12
            Top             =   225
            Width           =   615
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Nombre :"
            Height          =   225
            Index           =   7
            Left            =   135
            TabIndex        =   11
            Top             =   555
            Width           =   735
         End
      End
      Begin MSMask.MaskEdBox txtNroDocumento 
         Height          =   330
         Left            =   6000
         TabIndex        =   1
         Top             =   6510
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
         _Version        =   393216
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   9
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "###-#####"
         PromptChar      =   "_"
      End
      Begin VB.Frame fraContenedor 
         Height          =   1185
         Index           =   6
         Left            =   120
         TabIndex        =   15
         Top             =   5160
         Width           =   7365
         Begin VB.TextBox txtSubTotalBoleta 
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
            Left            =   3360
            TabIndex        =   34
            Top             =   840
            Width           =   1080
         End
         Begin VB.TextBox txtITF 
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
            Left            =   825
            TabIndex        =   32
            Top             =   840
            Width           =   1080
         End
         Begin VB.TextBox txtPreVentaNeto 
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
            Height          =   270
            Left            =   6075
            TabIndex        =   27
            Top             =   450
            Width           =   1125
         End
         Begin VB.TextBox txtSobrante 
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
            Left            =   3360
            TabIndex        =   24
            Top             =   495
            Width           =   1125
         End
         Begin VB.TextBox txtDeuda 
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
            Left            =   825
            TabIndex        =   18
            Top             =   150
            Width           =   1080
         End
         Begin VB.TextBox txtPreBaseVenta 
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
            Left            =   3375
            TabIndex        =   17
            Top             =   150
            Width           =   1125
         End
         Begin VB.TextBox txtPreVentaBruta 
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
            Left            =   6075
            TabIndex        =   0
            Top             =   150
            Width           =   1125
         End
         Begin VB.TextBox txtComision 
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
            Left            =   825
            TabIndex        =   16
            Top             =   495
            Width           =   1080
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "SubTotal :"
            Height          =   225
            Index           =   1
            Left            =   2100
            TabIndex        =   35
            Top             =   750
            Width           =   765
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "ITF :"
            Height          =   225
            Index           =   0
            Left            =   75
            TabIndex        =   33
            Top             =   870
            Width           =   765
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Precio Venta Neto :"
            Height          =   225
            Index           =   11
            Left            =   4620
            TabIndex        =   26
            Top             =   495
            Width           =   1455
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Sobrante :"
            Height          =   225
            Index           =   9
            Left            =   2100
            TabIndex        =   25
            Top             =   450
            Width           =   735
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Precio Base Venta :"
            Height          =   225
            Index           =   6
            Left            =   2100
            TabIndex        =   22
            Top             =   195
            Width           =   1410
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Comisión :"
            Height          =   225
            Index           =   5
            Left            =   75
            TabIndex        =   21
            Top             =   480
            Width           =   765
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Precio Venta Bruta :"
            Height          =   225
            Index           =   14
            Left            =   4605
            TabIndex        =   20
            Top             =   210
            Width           =   1485
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Deuda :"
            Height          =   225
            Index           =   13
            Left            =   75
            TabIndex        =   19
            Top             =   195
            Width           =   675
         End
      End
      Begin SICMACT.ActXCodCta AXCodCta 
         Height          =   375
         Left            =   180
         TabIndex        =   29
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
         TabIndex        =   30
         Top             =   600
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   6165
      End
      Begin MSMask.MaskEdBox txtNroBoleta 
         Height          =   330
         Left            =   1680
         TabIndex        =   36
         Top             =   6405
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
         _Version        =   393216
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   9
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "###-#####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Boleta Venta :"
         Height          =   225
         Index           =   4
         Left            =   360
         TabIndex        =   37
         Top             =   6480
         Width           =   1185
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Poliza Adjudicación :"
         Height          =   225
         Index           =   10
         Left            =   4200
         TabIndex        =   23
         Top             =   6555
         Width           =   1545
      End
   End
   Begin VB.Label lblMensaje 
      Caption         =   "<F12> Número Contrato Antiguo"
      Height          =   285
      Left            =   240
      TabIndex        =   28
      Top             =   7080
      Width           =   2655
   End
End
Attribute VB_Name = "frmColPRemateRegVenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************
'* REGISTRO DE VENTA DE REMATE.
'Archivo:  frmColPRemateRegVenta.frm
'LAYG   :  15/07/2001.
'Resumen:  Nos permite registrar una venta de contrato en remate
Option Explicit

Dim fsVarNroRemate As String

'Dim pDifeDiasRema As Integer
Dim fnVarTasaRemateComision As Double
Dim fnVarTasaImpuesto As Double
Dim fnVarTasaPreparacionRemate As Double
Dim fnVarTasaCustoriaVencida As Double
Dim fnVarTasaIGV As Double

Dim vNroContrato As String

Dim fnVarCapital As Currency
Dim fnVarDeuda As Currency
Dim fdFecVencimiento As Date
Dim fnVarDiasAtraso As Integer
Dim fnVarInteresVencido As Currency
Dim fnVarCostoCustodiaMoratorio As Currency
Dim fnVarImpuesto As Currency
Dim fnVarCostoPreparacionRemate As Currency
Dim fnVarValorTasacion As Currency
Dim fnVarTasaInteresVencido As Double
Dim fnVarEstado As Integer
Dim fnVarComisionRemate As Currency
Dim fnVarSobranteRemate As Currency
Dim fnVarIGVventa As Currency
Dim fnVarSubTotalBoleta As Currency
Dim nRedondeoITF As Double 'BRGO 20110914

'Inicializa el formulario
Public Sub Inicio(Optional ByVal psNroProceso As String = "")
Dim lsVerificaRemate As String
Dim lsmensaje As String
Dim lsAge As String
    If psNroProceso = "" Then
        'Obtiene el ultimo Remate y que este Inicializado.
        Dim loDat As COMNColoCPig.NCOMColPRecGar
        Set loDat = New COMNColoCPig.NCOMColPRecGar
            lsAge = gsCodCMAC & gsCodAge
            lsVerificaRemate = loDat.nObtieneNroUltimoProceso("R", lsAge, lsmensaje)
        Set loDat = Nothing
        If lsVerificaRemate <> "" Then
            fsVarNroRemate = lsVerificaRemate
        Else
             MsgBox "Remate no inicializado", vbInformation, "Aviso"
             Exit Sub
        End If
        'lsUltRemate = loDatos.nObtieneNroUltimoProceso("R", fsRemateCadaAgencia, lsmensaje)
    Else
        fsVarNroRemate = psNroProceso
    End If
    CargaParametros
    Limpiar
    Me.Show 1
End Sub

Private Sub Limpiar()
    Me.AXCodCta.NroCuenta = fgIniciaAxCuentaPignoraticio
    AXDesCon.Limpiar
    txtCodAdj = ""
    txtNomAdj = ""
    txtNatAdj = ""
    txtTriAdj = ""
    txtDeuda.Text = Format(0, "#0.00")
    txtPreBaseVenta.Text = Format(0, "#0.00")
    txtPreVentaBruta.Text = Format(0, "#0.00")
    txtPreVentaNeto.Text = Format(0, "#0.00")
    txtComision.Text = Format(0, "#0.00")
    txtSobrante.Text = Format(0, "#0.00")
End Sub

'Busca el contrato ingresado
Private Sub BuscaContrato(ByVal psNroContrato As String)

Dim lbok As Boolean
Dim lrValida As ADODB.Recordset
Dim loValContrato As COMNColoCPig.NCOMColPValida
Dim lsmensaje As String
'On Error GoTo ControlError

    'Valida Contrato
    Set lrValida = New ADODB.Recordset
    Set loValContrato = New COMNColoCPig.NCOMColPValida
        Set lrValida = loValContrato.nValidaRegVentaRemateCredPignoraticio(psNroContrato, "R", fsVarNroRemate, lsmensaje)
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
    fnVarValorTasacion = Format(lrValida!nTasacion, "#0.00")
    fnVarTasaInteresVencido = lrValida!nTasaIntVenc
    fnVarEstado = lrValida!nPrdEstado
    fnVarCapital = Format(lrValida!nSaldo, "#0.00")
    fdFecVencimiento = Format(lrValida!dVenc, "mm/dd/yyyy")

    'Muestra Datos
    lbok = fgMuestraCredPig_AXDesCon(psNroContrato, Me.AXDesCon, False)
    
    Call fgCalculaDeuda
    txtDeuda = Format(fnVarDeuda, "#0.00")
    txtPreBaseVenta = Format(lrValida!nRemSubBaseVta, "#0.00")
    txtPreVentaBruta = Format(lrValida!nRemSubBaseVta, "#0.00")
    Set lrValida = Nothing
        
    cmdGrabar.Enabled = True
   ' cmdGrabar.SetFocus
        
    AXCodCta.Enabled = False
    cmdBuscar.Enabled = True
    cmdBuscar.SetFocus

Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox "Error: " & Err.Number & " " & Err.Description & vbCr & _
        "Avise al Area de Sistemas ", vbInformation, " Aviso "


End Sub

Private Sub fgCalculaDeuda()
Dim loCalculos As COMNColoCPig.NCOMColPCalculos
fnVarDiasAtraso = DateDiff("d", Format(fdFecVencimiento, "mm/dd/yyyy"), Format(gdFecSis, "mm/dd/yyyy"))

If fnVarDiasAtraso <= 0 Then
    fnVarDiasAtraso = 0
    fnVarInteresVencido = 0
    fnVarCostoCustodiaMoratorio = 0
    fnVarImpuesto = 0
Else
    Set loCalculos = New COMNColoCPig.NCOMColPCalculos
        fnVarInteresVencido = loCalculos.nCalculaInteresMoratorio(fnVarCapital, fnVarTasaInteresVencido, fnVarDiasAtraso)
        fnVarInteresVencido = Round(fnVarInteresVencido, 2)
        fnVarCostoCustodiaMoratorio = loCalculos.nCalculaCostoCustodiaMoratorio(fnVarValorTasacion, fnVarTasaCustoriaVencida, fnVarDiasAtraso)
        fnVarCostoCustodiaMoratorio = Round(fnVarCostoCustodiaMoratorio, 2)
        fnVarImpuesto = (fnVarInteresVencido + fnVarCostoCustodiaMoratorio) * fnVarTasaImpuesto
        fnVarImpuesto = Round(fnVarImpuesto, 2)
    Set loCalculos = Nothing
End If
fnVarCostoPreparacionRemate = 0
If fnVarEstado = gColPEstPRema Then    ' Si esta en via de Remate
    fnVarCostoPreparacionRemate = fnVarTasaPreparacionRemate * fnVarValorTasacion
    fnVarCostoPreparacionRemate = Round(fnVarCostoPreparacionRemate, 2)
End If
fnVarDeuda = fnVarCapital + fnVarInteresVencido + fnVarCostoCustodiaMoratorio + fnVarImpuesto + fnVarCostoPreparacionRemate
End Sub

Private Sub AXCodCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call BuscaContrato(AXCodCta.NroCuenta)
End Sub


'Busca el Adjudicatario
Private Sub cmdBuscar_Click()
Dim loPers As COMDPersona.UCOMPersona
Dim lsPersCod As String
Dim liFil As Integer
Dim ls As String
On Error GoTo ControlError

Set loPers = New COMDPersona.UCOMPersona
Set loPers = frmBuscaPersona.Inicio

If Not loPers Is Nothing Then
    lsPersCod = loPers.sPersCod
    Me.txtCodAdj = loPers.sPersCod
    txtNomAdj = PstaNombre(loPers.sPersNombre, False)
    
    txtPreVentaBruta.Enabled = True
    txtPreVentaBruta.SetFocus

End If

Set loPers = Nothing
Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

'Cancela el proceso actual e inicializa uno nuevo
Private Sub cmdCancelar_Click()
    Limpiar
    txtNroDocumento.Enabled = False
    cmdGrabar.Enabled = False
    'cmdBuscar.Enabled = False
    txtPreVentaBruta.Enabled = False
    AXCodCta.Enabled = True
    AXCodCta.SetFocusCuenta
End Sub


Private Sub cmdGrabar_Click()
                                                                                                                                   
'On Error GoTo ControlError
Dim loContFunct As COMNContabilidad.NCOMContFunciones
Dim loGrabarVta As COMNColoCPig.NCOMColPContrato
Dim oMov As COMDMov.DCOMMov 'BRGO 20110914
Set oMov = New COMDMov.DCOMMov 'BRGO 20110914

Dim loPrevio As previo.clsprevio
Dim loRecImp As COMNColoCPig.NCOMColPRecGar

Dim lsMovNro As String
Dim lsFechaHoraGrab As String
Dim lsFechaVenc As String

Dim lsCadImp As String
 
'asigna valores a variables
fnVarComisionRemate = CCur(txtComision.Text)
fnVarSobranteRemate = CCur(txtSobrante.Text)
fnVarIGVventa = Val(txtITF.Text)
fnVarSubTotalBoleta = Val(txtSubTotalBoleta.Text)


If MsgBox(" Grabar Venta de Joyas en Remate ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
    cmdGrabar.Enabled = False
        
        
        'Genera el Mov Nro
        Set loContFunct = New COMNContabilidad.NCOMContFunciones
            lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        Set loContFunct = Nothing
        
        lsFechaHoraGrab = fgFechaHoraGrab(lsMovNro)
        Set loGrabarVta = New COMNColoCPig.NCOMColPContrato
            'Grabar Venta de Remate
            Call loGrabarVta.nRemateVentaCredPignoraticio(AXCodCta.NroCuenta, fsVarNroRemate, lsFechaHoraGrab, _
                 lsMovNro, Val(Me.txtPreVentaBruta.Text), fnVarCapital, fnVarInteresVencido, _
                  fnVarCostoCustodiaMoratorio, fnVarCostoPreparacionRemate, fnVarComisionRemate, fnVarImpuesto, _
                  fnVarSobranteRemate, fnVarValorTasacion, Val(Me.AXDesCon.Oro14), _
                  Val(Me.AXDesCon.Oro16), Val(Me.AXDesCon.Oro18), Val(Me.AXDesCon.Oro21), _
                  Trim(txtCodAdj.Text), Trim(txtNroDocumento.Text), Trim(txtNroBoleta.Text), _
                  fnVarIGVventa, fnVarSubTotalBoleta, False)
        Set loGrabarVta = Nothing
        If CCur(txtITF.Text) > 0 Then
            Call oMov.InsertaMovRedondeoITF(lsMovNro, 1, CCur(txtITF.Text) + nRedondeoITF, CCur(txtITF.Text)) 'BRGO 20110914
        End If
        Set oMov = Nothing
        'Impresión
        If MsgBox(" Desea Imprimir POLIZA ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
            ImprimirPolizaNueva
            Do While True
                If MsgBox("Desea Reimprimir POLIZA ?", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
                    ImprimirPolizaNueva
                Else
                    Exit Do
                End If
            Loop
        End If
        If MsgBox(" Desea imprimir RECIBO ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
                Set loRecImp = New COMNColoCPig.NCOMColPRecGar
                    lsCadImp = loRecImp.ImprimirComision(vNroContrato, txtNomAdj.Text, CCur(txtComision.Text), CCur(Me.txtSubTotalBoleta.Text), CCur(txtPreVentaNeto.Text), CCur(txtITF.Text), gdFecSis, gsCodUser, gsNomAge, gImpresora)
                Set loRecImp = Nothing
                Set loPrevio = New previo.clsprevio
                    loPrevio.PrintSpool sLpt, Chr(10) & Chr(10) & Chr(10) & lsCadImp, False
                    loPrevio.PrintSpool sLpt, Chr(10) & Chr(10) & Chr(10) & Chr(10), False
                Set loPrevio = Nothing
            Do While True
                If MsgBox("Desea Reimprimir RECIBO ?", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
                
                Set loPrevio = New previo.clsprevio
                    loPrevio.PrintSpool sLpt, Chr(10) & Chr(10) & Chr(10) & lsCadImp, False
                    loPrevio.PrintSpool sLpt, Chr(10) & Chr(10) & Chr(10) & Chr(10) & lsCadImp, False
                Set loPrevio = Nothing
                Else
                    Exit Do
                End If
            Loop
        End If
        txtNroDocumento = Val(txtNroDocumento) + 1
        txtNroDocumento = String(8 - Len(txtNroDocumento), "0") & Trim(Str(txtNroDocumento))
        Limpiar
        AXCodCta.Enabled = True
        AXCodCta.SetFocus
End If
Exit Sub


ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "

End Sub

'Finaliza el formulario
Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Me.Icon = LoadPicture(App.path & gsRutaIcono)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 And AXCodCta.Enabled = True Then 'F12
        Dim sCuenta As String
        sCuenta = frmValTarCodAnt.Inicia(gColConsuPrendario, False)
        If sCuenta <> "" Then
            AXCodCta.NroCuenta = sCuenta
            AXCodCta.SetFocusCuenta
        End If
    End If
End Sub



Private Sub txtNroBoleta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
'    Set loValid = New COMNColoCPig.NCOMColPValida
'        lbExiste = loValid.nDocumentoEmitido(23, txtNroDocumento.Text, "'" & geColPVtaRemate & "'")
'    Set loValid = Nothing
'    If lbExiste = True Then
'        MsgBox "Número de Poliza duplicada" & vbCr & "Ingrese un número diferente", vbInformation, " Aviso "
'    Else
        txtNroDocumento.Enabled = True
        txtNroDocumento.SetFocus
'    End If
End If
End Sub

'Valida el campo txtnrodocumento
Private Sub txtNroDocumento_KeyPress(KeyAscii As Integer)
Dim loValid As COMNColoCPig.NCOMColPValida
Dim lbExiste As Boolean
If KeyAscii = 13 And Len(Trim(txtNroDocumento)) = 8 Then
    Set loValid = New COMNColoCPig.NCOMColPValida
        lbExiste = loValid.nDocumentoEmitido(23, txtNroDocumento.Text, "'" & geColPVtaRemate & "'")
    Set loValid = Nothing
    If lbExiste = True Then
        MsgBox "Número de Poliza duplicada" & vbCr & "Ingrese un número diferente", vbInformation, " Aviso "
    Else
        cmdGrabar.Enabled = True
        cmdGrabar.SetFocus
    End If
End If
End Sub

'Valida el campo txtpreventabruta
Private Sub txtPreVentaBruta_GotFocus()
    fEnfoque txtPreVentaBruta
End Sub
Private Sub txtPreVentaBruta_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtPreVentaBruta, KeyAscii, 10, 2)
If KeyAscii = 13 Then
    If (Val(txtPreVentaBruta.Text) + 5) < Val(txtPreBaseVenta.Text) Then
        MsgBox " Precio Venta debe ser mayor a Precio Base ", vbInformation, " Aviso "
        txtPreVentaBruta.SetFocus
    Else
        txtNroBoleta.Enabled = True
        txtNroBoleta.SetFocus
    End If
End If
End Sub
Private Sub txtPreVentaBruta_LostFocus()
    txtPreVentaBruta = Format(txtPreVentaBruta, "#0.00")
    VeriVenBru
End Sub

'Procedimiento de impresión de la poliza
'Private Sub ImprimirPoliza()
'Dim lstTmpJoyas As ListItem
'
'
'Dim lnItem As Integer
'    ImpreBegChe False, 39
''    'Adjudicatario
'    Print #ArcSal, " ": Print #ArcSal, " ": Print #ArcSal, " ": Print #ArcSal, " "
'    Print #ArcSal, " ": Print #ArcSal, " ": Print #ArcSal, " ": Print #ArcSal, " "
'    Print #ArcSal, Tab(22); txtNomAdj
'    Print #ArcSal, Tab(22); txtNatAdj & " " & txtTriAdj;
'    Print #ArcSal, " ": Print #ArcSal, " "
'    'Cliente que joyas son rematadas
'    Print #ArcSal, Tab(22); AXDesCon.listaClientes.ListItems(1).ListSubItems.Item(1); 'vNombre;
'    Print #ArcSal, Tab(22); AXDesCon.listaClientes.ListItems(1).ListSubItems.Item(7) & "  " & AXDesCon.listaClientes.ListItems(1).ListSubItems.Item(9); 'txtNatural & " " & txtTributario;
'    Print #ArcSal, Tab(30); fsVarNroRemate & Space(18) & Format(gdFecSis, "dd/mm/yyyy")
'    Print #ArcSal, " ": Print #ArcSal, " "
'    For lnItem = 1 To Me.AXDesCon.listaJoyasDet.ListItems.Count
'        Print #ArcSal, Tab(10); Me.AXDesCon.listaJoyasDet.ListItems(lnItem).ListSubItems.Item(1); Space(2) & Me.AXDesCon.listaJoyasDet.ListItems(lnItem).ListSubItems.Item(4); Me.AXDesCon.listaJoyasDet.ListItems(lnItem).ListSubItems.Item(3) & " gr";
'        Print #ArcSal, Tab(10);
'    Next
'    For lnItem = Me.AXDesCon.listaJoyasDet.ListItems.Count To 15
'        Print #ArcSal, " "
'    Next
'
'    'Print #ArcSal, " ": Print #ArcSal, " "
''    Print #ArcSal, Tab(8); "01     LOTE          JOYAS CONTRATO N°  " & vNroContrato & Space(25) & txtPreVentaNeto
''    Print #ArcSal, " "
''    If (AXDesCon.Oro14) > 0 Then
''        Print #ArcSal, Tab(30); Format(AXDesCon.Oro14, "#0.00") & " grs de 14Kl."
''    Else
''        Print #ArcSal, " "
''    End If
''    If (AXDesCon.Oro16) > 0 Then
''        Print #ArcSal, Tab(30); Format(AXDesCon.Oro16, "#0.00") & " grs de 16Kl."
''    Else
''        Print #ArcSal, ""
''    End If
''    If (AXDesCon.Oro18) > 0 Then
''        Print #ArcSal, Tab(30); Format(AXDesCon.Oro18, "#0.00") & " grs de 18Kl."
''    Else
''        Print #ArcSal, ""
''    End If
''    If (AXDesCon.Oro21) > 0 Then
''        Print #ArcSal, Tab(30); Format(AXDesCon.Oro21, "#0.00") & " grs de 21Kl."
''    Else
''        Print #ArcSal, ""
''    End If
''    Print #ArcSal, " ": Print #ArcSal, " ": Print #ArcSal, " ": Print #ArcSal, " ": Print #ArcSal, " "
'    Print #ArcSal, Tab(85); Format(txtSubTotalBoleta.Text, "#0.00")
'    Print #ArcSal, Tab(85); Format(txtITF.Text, "#0.00")
'    Print #ArcSal, Tab(85); Format(txtPreVentaNeto, "#0.00")
'
'    'Print #ArcSal, Tab(15); UCase(NumLet(Str(Int(CCur(txtPreVentaNeto.Text))))) & " Y " & Left(Str(CCur(txtPreVentaNeto.Text) - Int(CCur(txtPreVentaNeto.Text))) * 100, 2) & "/100 NUEVOS SOLES"
'    Print #ArcSal, " "
'
'    ImpreEnd
'End Sub

Private Sub ImprimirPolizaNueva()
    Dim lstTmpJoyas As ListItem
    Dim lnItem As Integer, i As Integer, conta As Integer
    Dim rs As New ADODB.Recordset
    Dim rs1 As New ADODB.Recordset
    Dim loImp As COMNColoCPig.NCOMColPImpre
    Dim lsCadImp As String
    With rs
        'Crear RecordSet
        .Fields.Append "cCliente", adVarChar, 150
        .Fields.Append "cTelefono", adVarChar, 50
        .Fields.Append "cZona", adVarChar, 100
        .Fields.Append "cNroDNI", adVarChar, 10
        .Fields.Append "cNroRUC", adVarChar, 15
        .Open
        'Llenar Recordset
        For lnItem = 1 To Me.AXDesCon.listaClientes.ListItems.Count
            .AddNew
            .Fields("cCliente") = AXDesCon.listaClientes.ListItems(lnItem).ListSubItems.iTem(1)
            .Fields("cTelefono") = AXDesCon.listaClientes.ListItems(lnItem).ListSubItems.iTem(3)
            .Fields("cZona") = AXDesCon.listaClientes.ListItems(lnItem).ListSubItems.iTem(4)
            .Fields("cNroDNI") = AXDesCon.listaClientes.ListItems(lnItem).ListSubItems.iTem(7)
            .Fields("cNroRUC") = AXDesCon.listaClientes.ListItems(lnItem).ListSubItems.iTem(9)
        Next lnItem
    End With
    
    With rs1
        'Llenar Recordset
        .Fields.Append "cPieza", adVarChar, 50
        .Fields.Append "cDescripcion", adVarChar, 150
        .Open
        For lnItem = 1 To Me.AXDesCon.listaJoyasDet.ListItems.Count
            .AddNew
            .Fields("cPieza") = Trim(AXDesCon.listaJoyasDet.ListItems(lnItem).ListSubItems.iTem(1))
            .Fields("cDescripcion") = AXDesCon.listaJoyasDet.ListItems(lnItem).ListSubItems.iTem(4)
        Next
    End With
    
    Set loImp = New COMNColoCPig.NCOMColPImpre
         lsCadImp = loImp.ImprimePolizaNueva(txtNomAdj, txtNatAdj, txtTriAdj, fsVarNroRemate, gdFecSis, txtSubTotalBoleta, txtITF, txtPreVentaNeto, rs, rs1, gImpresora)
    Set loImp = Nothing
    ImpreBegChe True, 22
    Print #ArcSal, lsCadImp
    Print #ArcSal, " "
    Close ArcSal
   ' ImpreEnd
End Sub


'Procedimiento de verificación de la venta bruta
Private Sub VeriVenBru()
    If (Val(txtPreVentaBruta.Text) + 5) < Val(txtPreBaseVenta.Text) Then
        MsgBox " Precio Venta debe ser mayor a Precio Base ", vbInformation, " Aviso "
        txtPreVentaBruta.SetFocus
    Else
        txtComision = Format(Round((CCur(txtPreVentaBruta.Text) * fnVarTasaRemateComision), 2), "#0.00")
        txtPreVentaNeto = Format(Round(CCur(txtPreVentaBruta.Text) - CCur(txtComision.Text), 2), "#0.00")
        'txtITF.Text = Format(Round(CCur(txtPreVentaNeto.Text) - (CCur(txtPreVentaNeto.Text) / (1 + fnVarTasaIGV)), 2), "#0.00")
        txtITF.Text = CCur(txtPreVentaNeto.Text) - Format(Round(CCur(txtPreVentaNeto.Text) / (1.0008), 2), "#0.00")
        '*** BRGO 20110908 ************************************************
        nRedondeoITF = fgDiferenciaRedondeoITF(CCur(Me.txtITF.Text))
        If nRedondeoITF > 0 Then
            Me.txtITF.Text = Format(CCur(Me.txtITF.Text) - nRedondeoITF, "#,##0.00")
        End If
        '*** END BRGO
        txtSubTotalBoleta.Text = Format(Round(CCur(txtPreVentaNeto.Text) - CCur(txtITF.Text), 2), "#0.00")
        'fnVarIGVventa = Format(Round(CCur(txtPreVentaNeto.Text) - (CCur(txtPreVentaNeto.Text) / (1 + fnVarTasaIGV)), 2), "#0.00")
        'txtSobrante.Text = Format(Round(CCur(txtPreVentaBruta.Text) - CCur(txtITF.Text) - CCur(txtDeuda.Text), 2), "#0.00")
        'txtSobrante.Text = Format(Round(CCur(txtPreVentaBruta.Text) - CCur(txtDeuda.Text), 2), "#0.00")
         txtSobrante.Text = Format(Round(CCur(txtPreVentaNeto.Text) / (1.0008), 2), "#0.00") - CCur(txtDeuda.Text)
    End If
End Sub

Private Sub CargaParametros()
Dim loParam As COMDColocPig.DCOMColPCalculos
Set loParam = New COMDColocPig.DCOMColPCalculos
    fnVarTasaRemateComision = loParam.dObtieneColocParametro(gConsColPTasaComisionRemate)
    fnVarTasaImpuesto = loParam.dObtieneColocParametro(gConsColPTasaImpuesto)
    fnVarTasaPreparacionRemate = loParam.dObtieneColocParametro(gConsColPTasaPreparaRemate)
    fnVarTasaCustoriaVencida = loParam.dObtieneColocParametro(gConsColPTasaCustodiaVencida)
    fnVarTasaIGV = loParam.dObtieneColocParametro(gConsColPTasaIGV)
    'pDifeDiasRema = loParam.dObtieneColocParametro(gConsColPCostoDuplicadoContrato)
Set loParam = Nothing
End Sub

