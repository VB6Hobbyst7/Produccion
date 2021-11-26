VERSION 5.00
Begin VB.Form frmColPRescateJoyas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Crédito Pignoraticio - Rescate de Joyas"
   ClientHeight    =   6204
   ClientLeft      =   1296
   ClientTop       =   1548
   ClientWidth     =   8004
   Icon            =   "frmColPRescateJoyas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6204
   ScaleWidth      =   8004
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   5640
      TabIndex        =   1
      Top             =   5745
      Width           =   975
   End
   Begin VB.Frame fraContenedor 
      Height          =   5475
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   7815
      Begin VB.CommandButton cmdBuscar 
         Height          =   345
         Left            =   7110
         Picture         =   "frmColPRescateJoyas.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Buscar ..."
         Top             =   270
         Width           =   420
      End
      Begin VB.Frame fraContenedor 
         Enabled         =   0   'False
         Height          =   1005
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   4320
         Width           =   7425
         Begin VB.Label lblCostoCustodia 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   1800
            TabIndex        =   15
            Top             =   600
            Width           =   1245
         End
         Begin VB.Label lblFecPago 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   1800
            TabIndex        =   14
            Top             =   240
            Width           =   1245
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Costo Custodia"
            Height          =   225
            Index           =   18
            Left            =   360
            TabIndex        =   13
            Top             =   600
            Width           =   1380
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Dias Custodia"
            Height          =   225
            Index           =   19
            Left            =   3840
            TabIndex        =   12
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label lblDiasTranscurridos 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   5160
            TabIndex        =   11
            Top             =   600
            Width           =   735
         End
         Begin VB.Label lblNroDuplic 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   5160
            TabIndex        =   7
            Top             =   240
            Width           =   765
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Nro.Duplicado"
            Height          =   255
            Index           =   12
            Left            =   3840
            TabIndex        =   6
            Top             =   240
            Width           =   1275
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Fec.Cancelación"
            Height          =   225
            Index           =   16
            Left            =   345
            TabIndex        =   5
            Top             =   240
            Width           =   1275
         End
      End
      Begin SICMACT.ActXCodCta AXCodCta 
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   3615
         _ExtentX        =   6371
         _ExtentY        =   656
         Texto           =   "Crédito"
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
      Begin SICMACT.ActXColPDesCon AXDesCon 
         Height          =   3615
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   7575
         _ExtentX        =   13356
         _ExtentY        =   6371
      End
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4560
      TabIndex        =   0
      Top             =   5745
      Width           =   975
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6840
      TabIndex        =   2
      Top             =   5745
      Width           =   975
   End
   Begin VB.Label lblMensaje 
      Caption         =   "<F12> Número Contrato Antiguo"
      Height          =   285
      Left            =   240
      TabIndex        =   8
      Top             =   5760
      Visible         =   0   'False
      Width           =   2655
   End
End
Attribute VB_Name = "frmColPRescateJoyas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'RESCATE DE JOYAS
'Archivo                : frmColPRescateJoyas.frm
'El Proceso de Rescate de Joyas nos permite registrar la devolución de joyas
'Fecha : 10/05/2001

Option Explicit
Dim fnMaxDiasCustodiaDiferida As Double
Dim fnTasaIGV As Double
Dim fnPorcentajeCustodiaDiferida As Double

Dim vCostoCustodiaExtra As Double
Dim vSaldoCustodiaExtra As Double

Dim fnVarOpeCod As Long
Dim fsVarOpeDesc As String
Dim objPista As COMManejador.Pista


Public Sub Inicio(ByVal pnOpeCod As Long, ByVal psOpeDesc As String)

    fnVarOpeCod = pnOpeCod
    fsVarOpeDesc = psOpeDesc
    
    Select Case fnVarOpeCod
        Case gColPOpeDevJoyas
            Me.Caption = "Credito Pignoraticio - Rescate de Joya "
        Case gColPOpeDevJoyasNoDesemb
            Me.Caption = "Credito Pignoraticio - Rescate de Joya No Desembolsada"
    End Select
    Limpiar
    CargaParametros
    AXCodCta.NroCuenta = fgIniciaAxCuentaPignoraticio
    Me.Show 1

End Sub
'Permite inicializar las variables
Private Sub Limpiar()
    Me.AXCodCta.NroCuenta = fgIniciaAxCuentaPignoraticio
    AXDesCon.Limpiar
    Me.lblCostoCustodia.Caption = Format(0, "#0.00")
    Me.lblFecPago.Caption = "  /  /    "
    Me.lblNroDuplic.Caption = ""
    Me.lblDiasTranscurridos.Caption = ""
End Sub

'Permite buscar el contrato ingresado
Public Sub BuscaContrato(ByVal psNroContrato As String)

Dim lbOk As Boolean
Dim lrValida As ADODB.Recordset
Dim loValContrato As COMNColoCPig.NCOMColPValida
Dim loCalculos As COMNColoCPig.NCOMColPCalculos
Dim lnCustodiaDiferida  As Currency
Dim lsmensaje As String
'On Error GoTo ControlError

    'Valida Contrato
    Set lrValida = New ADODB.Recordset
        Set loValContrato = New COMNColoCPig.NCOMColPValida
          '  Set lrValida = loValContrato.nValidaRescateCredPignoraticio(psNroContrato, gsCodAge, gdFecSis, fnVarOpeCod)
            Set lrValida = loValContrato.nValidaRescateCredPignoraticio(psNroContrato, gsCodAge, gdFecSis, fnVarOpeCod, gsCodUser, lsmensaje)
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
        '** Muestra Datos
        lbOk = fgMuestraCredPig_AXDesCon(psNroContrato, Me.AXDesCon, True)
        'Muestra otros datos
        
        Set loCalculos = New COMNColoCPig.NCOMColPCalculos
            lnCustodiaDiferida = loCalculos.nCalculaCostoCustodiaDiferida(lrValida!nTasacion, IIf(IsNull(lrValida!nDiasTranscurridos), 0, lrValida!nDiasTranscurridos), lrValida!nPorcentajeCustodia, lrValida!nTasaIGV)
        Set loCalculos = Nothing
        
        'Me.lblCostoCustodia = Format(lnCustodiaDiferida - lrValida!nCustodiaDiferida, "#0.00")
        Me.lblCostoCustodia = Format(lnCustodiaDiferida - lrValida!nCustodiaPag, "#0.00") '*** PEAC 20170929
        
        Me.lblFecPago = Format(lrValida!dCancelado, "dd/mm/yyyy")
        Me.lblDiasTranscurridos = IIf(IsNull(lrValida!nDiasTranscurridos), 0, lrValida!nDiasTranscurridos)
        Me.lblNroDuplic = lrValida!nNroDuplic
    Set lrValida = Nothing
        
    cmdGrabar.Enabled = True
    cmdGrabar.SetFocus
        
    AXCodCta.Enabled = False
    'cmdBuscar.Enabled = False
Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "

End Sub

Private Sub AXCodCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then BuscaContrato (AXCodCta.NroCuenta)
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
    lsPersCod = loPers.sPerscod
    lsPersNombre = loPers.sPersNombre
Set loPers = Nothing

' Selecciona Estados
lsEstados = gColPEstDifer

If Trim(lsPersCod) <> "" Then
    Set loPersContrato = New COMDColocPig.DCOMColPContrato
        Set lrContratos = loPersContrato.dObtieneCredPigDePersona(lsPersCod, lsEstados, Mid(gsCodAge, 4, 2))
    Set loPersContrato = Nothing
End If

Set loCuentas = New COMDPersona.UCOMProdPersona
    Set loCuentas = frmProdPersona.Inicio(lsPersNombre, lrContratos)
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

'Cancela un proceso y reinicializa las variables par iniciar otro
Private Sub cmdCancelar_Click()
    Limpiar
    cmdGrabar.Enabled = False
    AXCodCta.Enabled = True
    'AXCodCta.SetFocusCuenta
    AXCodCta.SetFocusProd
End Sub

Private Sub CmdGrabar_Click()

'On Error GoTo ControlError
Dim loContFunct As COMNContabilidad.NCOMContFunciones
Dim loGrabarResc As COMNColoCPig.NCOMColPContrato
Dim oImp As New COMNColoCPig.NCOMColPImpre
Dim loColPRes As COMDColocPig.DCOMColPContrato
Dim loPrevio As previo.clsprevio

Dim lsMovNro As String
Dim lsFechaHoraGrab As String
Dim lsCuenta As String
Dim lsCadImp As String

Dim lnSaldoCap As Currency, lnInteresComp As Currency, lnImpuesto As Currency
Dim lnCostoTasacion As Currency, lnCostoCustodia As Currency
Dim lnMontoEntregar As Currency

' Verificar que el Contrato no este Rescatado

Set loColPRes = New COMDColocPig.DCOMColPContrato
    If loColPRes.VerificarCreditoxRescatar(AXCodCta.NroCuenta) = True Then
        MsgBox "Este cretido ya esta Rescatado", vbInformation, "Aviso"
        cmdGrabar.Enabled = False
        Exit Sub
    End If
Set loColPRes = Nothing
If MsgBox(" Grabar Rescate de Contrato Pignoraticio ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
    cmdGrabar.Enabled = False
        
        'Genera el Mov Nro
        Set loContFunct = New COMNContabilidad.NCOMContFunciones
            lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        Set loContFunct = Nothing
        
        lsFechaHoraGrab = fgFechaHoraGrab(lsMovNro)
        Set loGrabarResc = New COMNColoCPig.NCOMColPContrato
            'Grabar Cancelacion Pignoraticio
            Call loGrabarResc.nRescataJoyaCredPignoraticio(AXCodCta.NroCuenta, lsFechaHoraGrab, _
                 lsMovNro, val(Me.AXDesCon.Oro14), val(Me.AXDesCon.Oro16), _
                 val(Me.AXDesCon.Oro18), val(Me.AXDesCon.Oro21), val(Me.lblDiasTranscurridos.Caption), val(Me.AXDesCon.ValTasa), gColPOpeDevJoyas, False)
            
            ''*** PEAC 20090126
            objPista.InsertarPista gsOpeCod, lsMovNro, gsCodPersUser, GetMaquinaUsuario, gModificar, , AXCodCta.NroCuenta, gCodigoCuenta
                                  
        Set loGrabarResc = Nothing

        'Impresión
        If MsgBox("Desea realizar impresiones ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
            Set loPrevio = New previo.clsprevio
            lsCadImp = oImp.ImprimirRescate(AXCodCta.NroCuenta, AXDesCon.listaClientes.ListItems(1).SubItems(1), AXDesCon.prestamo, AXDesCon.SaldoCapital, gdFecSis, gsCodUser, gImpresora)
            loPrevio.PrintSpool sLpt, lsCadImp, False, 22
            Do While True
                If MsgBox("Desea reimprimir ?", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
                     loPrevio.PrintSpool sLpt, lsCadImp, False, 22
                Else
                    Exit Do
                    Set loPrevio = Nothing
                End If
            Loop
            Set loPrevio = Nothing
        End If
        Limpiar
        Set oImp = Nothing
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

'Permite salir del formulario actual
Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub CargaParametros()
Dim loParam As COMDColocPig.DCOMColPCalculos
Set loParam = New COMDColocPig.DCOMColPCalculos
    fnMaxDiasCustodiaDiferida = loParam.dObtieneColocParametro(gConsColPMaxDiasCustodiaDiferida)
    fnTasaIGV = loParam.dObtieneColocParametro(gConsColPTasaIGV)
    fnPorcentajeCustodiaDiferida = loParam.dObtieneColocParametro(gConsColPPorcentajeCustodiaDiferida)
Set loParam = Nothing
End Sub

Private Sub Form_Load()
'ARCV 05-07-2007
fnVarOpeCod = gColPOpeDevJoyas
 '--------------
 Me.Icon = LoadPicture(App.path & gsRutaIcono)
 Limpiar
 
 Set objPista = New COMManejador.Pista
 gsOpeCod = gPigRescatarJoyas
 
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim loParam As COMDColocPig.DCOMColPContrato
Set loParam = New COMDColocPig.DCOMColPContrato
Dim rstemp As New ADODB.Recordset
Dim nEstado As Integer

    If KeyCode = vbKeyF12 And AXCodCta.Enabled = True Then 'F12
        Dim sCuenta As String
        sCuenta = frmValTarCodAnt.Inicia(gColConsuPrendario, False)
                                   
        If sCuenta <> "" Then
          Set rstemp = loParam.dObtieneDatosCreditoPignoraticio(sCuenta)
            If rstemp!nPrdEstado <> gColPEstDifer Then
                MsgBox "Este contrato no esta diferido.", vbOKOnly + vbInformation, App.Title
            Else
            
                AXCodCta.NroCuenta = sCuenta
                AXCodCta.SetFocusCuenta
            End If
        End If
    End If
    Set rstemp = Nothing
    Set loParam = Nothing
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objPista = Nothing
End Sub
