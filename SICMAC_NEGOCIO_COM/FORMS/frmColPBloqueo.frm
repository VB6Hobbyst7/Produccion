VERSION 5.00
Begin VB.Form frmColPBloqueo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Crédito Pignoraticio - Bloque/Desbloqueo de Contrato"
   ClientHeight    =   6225
   ClientLeft      =   2190
   ClientTop       =   915
   ClientWidth     =   8100
   Icon            =   "frmColPBloqueo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   8100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4320
      TabIndex        =   5
      Top             =   5775
      Width           =   975
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6525
      TabIndex        =   2
      Top             =   5775
      Width           =   975
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   5400
      TabIndex        =   1
      Top             =   5775
      Width           =   975
   End
   Begin VB.Frame fraContenedor 
      Height          =   5640
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7815
      Begin VB.CommandButton cmdBuscar 
         Height          =   345
         Left            =   7200
         Picture         =   "frmColPBloqueo.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Buscar ..."
         Top             =   270
         Width           =   420
      End
      Begin SICMACT.ActXColPDesCon AXColPDesCon 
         Height          =   3495
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   6165
      End
      Begin SICMACT.ActXCodCta AXCodCta 
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   661
         Texto           =   "Crédito"
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
      Begin VB.Frame fraBloqueo 
         Caption         =   "Bloqueo/Desbloqueo "
         Enabled         =   0   'False
         Height          =   960
         Left            =   120
         TabIndex        =   4
         Top             =   4440
         Width           =   7560
         Begin VB.Frame fraMotivo 
            Caption         =   "Motivo "
            Enabled         =   0   'False
            Height          =   930
            Left            =   2040
            TabIndex        =   7
            Top             =   0
            Width           =   5505
            Begin VB.TextBox txtDescripcion 
               Height          =   675
               Left            =   1800
               MaxLength       =   255
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   10
               Top             =   150
               Width           =   3600
            End
            Begin VB.OptionButton optMotivo 
               Caption         =   "Mandato Judicial"
               Height          =   195
               Index           =   1
               Left            =   195
               TabIndex        =   9
               Top             =   510
               Width           =   1530
            End
            Begin VB.OptionButton optMotivo 
               Caption         =   "Administrativo"
               Height          =   195
               Index           =   0
               Left            =   195
               TabIndex        =   8
               Top             =   240
               Value           =   -1  'True
               Width           =   1320
            End
         End
         Begin VB.CheckBox chkBloqueo 
            Caption         =   "Bloquear / Desbloqueo    Contrato"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   570
            Left            =   240
            TabIndex        =   6
            Top             =   240
            Width           =   1530
         End
      End
   End
   Begin VB.Label lblMensaje 
      Caption         =   "<F12> Número Contrato Antiguo"
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   5835
      Visible         =   0   'False
      Width           =   2655
   End
End
Attribute VB_Name = "frmColPBloqueo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************
'* REGISTRO DE CONTRATO.
'Archivo:  frmColPBloqueo.frm
'LAYG   :  05/05/2001.
'Resumen:  Nos permite Bloquear/Desbloquear un contrato pignoraticio

Option Explicit
Dim fbBlqIni As Boolean
Dim fsMovNroBloqueo As String
Dim fbNewBlq As String
Dim objPista As COMManejador.Pista

'Permite inicializarlas variables del formulario
Private Sub Limpiar()
    Me.AXCodCta.NroCuenta = fgIniciaAxCuentaPignoraticio
    AXColPDesCon.Limpiar
    txtDescripcion.Text = ""
    ChkBloqueo.value = 0
    optMotivo.iTem(0).value = True
End Sub

'Permite buscar el contrato ingresado
Private Sub BuscaContrato(ByVal psNroContrato As String)

Dim lbOk As Boolean
Dim lrValida As ADODB.Recordset
Dim loValContrato As COMNColoCPig.NCOMColPValida
Dim lsmensaje As String
'On Error GoTo ControlError
    
    'Valida Contrato
    Set lrValida = New ADODB.Recordset
    Set loValContrato = New COMNColoCPig.NCOMColPValida
        Set lrValida = loValContrato.nValidaBloqueoCredPignoraticio(psNroContrato, lsmensaje)
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
    If lrValida.BOF And lrValida.EOF Then
        Limpiar
        Set lrValida = Nothing
        Exit Sub
    End If
    lbOk = fgMuestraCredPig_AXDesCon(psNroContrato, Me.AXColPDesCon, False)
    fbBlqIni = IIf(lrValida!cBloqueo = "S", True, False)
    If lrValida!cBloqueo = "S" Then
        fsMovNroBloqueo = lrValida!cMovNroBloqueo
        ChkBloqueo.value = 1
        txtDescripcion.Text = lrValida!cComentario
        If lrValida!nBlqTpo = "08" Then
            optMotivo.iTem(0).value = True
            optMotivo.iTem(1).value = False
        Else
            optMotivo.iTem(0).value = False
            optMotivo.iTem(1).value = True
        End If
    Else
        ChkBloqueo.value = 0
    End If
    fraBloqueo.Enabled = True
    Me.fraMotivo.Enabled = True
    
    ChkBloqueo.SetFocus
    
    AXCodCta.Enabled = False
    cmdbuscar.Enabled = False

Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

Private Sub AXCodCta_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then BuscaContrato (AXCodCta.NroCuenta)
End Sub


Private Sub chkBloqueo_Click()
If fbBlqIni = True Then
    cmdgrabar.Enabled = IIf(ChkBloqueo.value = 1, False, True)
    fbNewBlq = True
Else
    cmdgrabar.Enabled = IIf(ChkBloqueo.value = 1, True, False)
    fraMotivo.Enabled = IIf(ChkBloqueo.value = 1, True, False)
    fbNewBlq = False
End If
End Sub

'Permite buscar un cliente por su nombre y/o documento
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
lsEstados = gColPEstRegis & "," & gColPEstDesem

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

'Permite cancelar un proceso e inicializar los campos para otro proceso
Private Sub cmdCancelar_Click()
    Limpiar
    cmdgrabar.Enabled = False
    cmdbuscar.Enabled = True
    fraBloqueo.Enabled = False
    AXCodCta.Enabled = True
    'MAVM BAS II
    'AXCodCta.SetFocusCuenta
    AXCodCta.SetFocusProd
End Sub

Private Sub CmdGrabar_Click()

Dim loContFunct As COMNContabilidad.NCOMContFunciones
Dim loGrabarLote As COMNColoCPig.NCOMColPContrato
Dim loImpBlo As COMNColoCPig.NCOMColPImpre
Dim loPrevio As previo.clsprevio

Dim lsMovNro As String
Dim lsFechaHoraGrab As String
Dim lsCuenta As String
Dim lsLote As String
Dim lsCadImp As String

'On Error GoTo ControlError

lsCuenta = AXCodCta.NroCuenta

    If MsgBox(" Grabar Bloqueo/Desbloqueo ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
        cmdgrabar.Enabled = False
        
        'Genera el Mov Nro
        Set loContFunct = New COMNContabilidad.NCOMContFunciones
            lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        Set loContFunct = Nothing
        
        lsFechaHoraGrab = fgFechaHoraGrab(lsMovNro)
        Set loGrabarLote = New COMNColoCPig.NCOMColPContrato
            'Grabar Bloqueo / DesBloqueo
            Call loGrabarLote.nBloqueo_DesBloqueoCredPignoraticio(lsCuenta, lsFechaHoraGrab, lsMovNro, fbBlqIni, Me.txtDescripcion.Text, fsMovNroBloqueo, False)
            
        ''*** PEAC 20090126
        objPista.InsertarPista gsOpeCod, lsMovNro, gsCodPersUser, GetMaquinaUsuario, gEliminar, "Bloquear credito pignoraticio", lsCuenta, gCodigoCuenta
                        
        Set loGrabarLote = Nothing

'       Set loImpBlo = New COMNColoCPig.NCOMColPImpre
'          lsCadImp = loImpBlo.ImpBloqueo(IIf(fbNewBlq, True, False), vNroContrato, AXDesCon.Plazo, lstCliente.ListItems(1).ListSubItems.Item(1))
'             lsCadImp = loImpBlo.ImpBloqueo(IIf(fbNewBlq, True, False), lsCuenta, AXDesCon.Plazo, lstCliente.ListItems(1).ListSubItems.iTem(1), gsCodUser)
'        Set loImpBlo = Nothing
'
'        Set loPrevio = New Previo.clsPrevio
'
'        If MsgBox("Desea realizar de Comprobante de Bloqueo/DesBloqueo ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
'            loPrevio.PrintSpool sLpt, lsCadImprimir, False
'            Do While True
'                If MsgBox("Desea reimprimir ?", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
'                    loPrevio.PrintSpool sLpt, lsCadImprimir, False
'                Else
'                    Exit Do
'                End If
'            Loop
'        End If
        
        Set loPrevio = Nothing
        Limpiar
        'fraListado.Visible = False
        fraBloqueo.Enabled = False
        cmdbuscar.Enabled = True
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

'Carga formulario de busqueda de contrato antiguo
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
    If KeyCode = vbKeyF12 And AXCodCta.Enabled = True Then 'F12
        Dim sCuenta As String
        sCuenta = frmValTarCodAnt.Inicia(gColConsuPrendario, False)
        If sCuenta <> "" Then
            AXCodCta.NroCuenta = sCuenta
            AXCodCta.SetFocusCuenta
            SendKeys "{Enter}"
        End If
    End If
End Sub

'Permite inicializar el formulario actual
Private Sub Form_Load()
    Limpiar
    
    Set objPista = New COMManejador.Pista
    gsOpeCod = gPigBloquearContrato
    
End Sub

'Permite mostrar los valores del list en los campos del número de contrato
Private Sub cboContratos_Click()
'    AXCodCta.Cuenta = Mid(Trim(cboContratos.Text), 1, 18)
'    cmdGrabar.Enabled = False
'    If Trim(AXCodCta.Cuenta) <> "" Then
'        If Right(Trim(gsCodAge), 2) = Mid(AXCodCta.Cuenta, 1, 2) Then
'            'Call BuscaContrato(dbCmact)
'        Else
'            'If AbreConeccion(vNroContrato, , True) Then
'            '    Call BuscaContrato(dbCmactN)
'            'Else
'            '    Limpiar
'            '    AXCodCta.SetFocus
'            'End If
'        End If
'        cboContratos.SetFocus
'    Else
'        MsgBox " Contrato no encontrado ", vbInformation, " Aviso "
'    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set objPista = Nothing
End Sub

Private Sub TxtDescripcion_KeyPress(KeyAscii As Integer)
    KeyAscii = fgIntfMayusculas(KeyAscii)
End Sub

