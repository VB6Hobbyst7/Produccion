VERSION 5.00
Begin VB.Form frmColPAnularPrestamoPig 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Crédito Pignoraticio - Anulación de Contrato"
   ClientHeight    =   6090
   ClientLeft      =   1275
   ClientTop       =   1770
   ClientWidth     =   7965
   Icon            =   "frmColPAnularPrestamoPig.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   7965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraContenedor 
      Height          =   5580
      Index           =   0
      Left            =   105
      TabIndex        =   3
      Top             =   30
      Width           =   7800
      Begin VB.TextBox txtMotivo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C16A0B&
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   720
         MaxLength       =   280
         TabIndex        =   7
         ToolTipText     =   "Ingrese su Usuario"
         Top             =   5040
         Width           =   6855
      End
      Begin VB.CommandButton cmdBuscar 
         Height          =   345
         Left            =   7110
         Picture         =   "frmColPAnularPrestamoPig.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Buscar ..."
         Top             =   270
         Width           =   420
      End
      Begin SICMACT.ActXColPDesCon AXDesCon 
         Height          =   3615
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   7575
         _extentx        =   13361
         _extenty        =   6376
      End
      Begin SICMACT.ActXCodCta AXCodCta 
         Height          =   375
         Left            =   180
         TabIndex        =   4
         Top             =   240
         Width           =   3615
         _extentx        =   6376
         _extenty        =   661
         texto           =   "Credito"
         enabledcta      =   -1  'True
      End
      Begin VB.Label Label1 
         Caption         =   "N° Holograma:"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   4560
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Tasador:"
         Height          =   255
         Left            =   3120
         TabIndex        =   11
         Top             =   4560
         Width           =   735
      End
      Begin VB.Label lblHolograma 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1440
         TabIndex        =   10
         Top             =   4560
         Width           =   1455
      End
      Begin VB.Label lblTasador 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   3960
         TabIndex        =   9
         Top             =   4560
         Width           =   975
      End
      Begin VB.Label lblComentario 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Motivo:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   5040
         Width           =   525
      End
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4320
      TabIndex        =   0
      Top             =   5640
      Width           =   1005
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6720
      TabIndex        =   2
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   5520
      TabIndex        =   1
      Top             =   5640
      Width           =   975
   End
End
Attribute VB_Name = "frmColPAnularPrestamoPig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************
'* MANTENIMIENTO DE CONTRATO PIGNORATICIO
'Archivo:  frmColPAnularPrestamoPig.frm
'LAYG   :  01/05/2001.
'Resumen:  Nos permite anular un Credito Pignoraticio
'          (solo si es digitado el mismo dia, y esta en estado "Registrado")
Option Explicit
Dim objPista As COMManejador.Pista

Private Sub limpiar()
    Me.AXCodCta.NroCuenta = fgIniciaAxCuentaPignoraticio
    AXDesCon.limpiar
    Me.txtMotivo.Text = ""
End Sub

Private Sub BuscaContrato(ByVal psNroContrato As String)
Dim lbOk As Boolean
Dim lrValida As ADODB.Recordset
Dim loValContrato As COMNColoCPig.NCOMColPValida
Dim lsmensaje As String
Dim rs As ADODB.Recordset 'APRI20180705 ERS063-2017
'On Error GoTo ControlError

    'Valida Contrato
    Set lrValida = New ADODB.Recordset
    Set loValContrato = New COMNColoCPig.NCOMColPValida
        'Set lrValida = loValContrato.nValidaAnulacionCredPignoraticio(psNroContrato, gdFecSis)
    Set lrValida = loValContrato.nValidaAnulacionCredPignoraticio(psNroContrato, gdFecSis, gsCodUser, lsmensaje)
    If Trim(lsmensaje) <> "" Then
         MsgBox lsmensaje, vbInformation, "Aviso"
         Exit Sub
    End If
    Set loValContrato = Nothing
    
    If lrValida Is Nothing Then ' Hubo un Error
        limpiar
        Set lrValida = Nothing
        Exit Sub
    End If

    lbOk = fgMuestraCredPig_AXDesCon(psNroContrato, Me.AXDesCon, False)
    'APRI20180705 ERS063-2017
    Set rs = ObtieneTasadorHolograma(psNroContrato)
    lblHolograma.Caption = rs!nHolograma
    lblTasador.Caption = rs!cUser
    Set rs = Nothing
    'END APRI
    AXCodCta.Enabled = False
    cmdBuscar.Enabled = False
    Me.cmdGrabar.Enabled = True
    cmdGrabar.SetFocus

Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

Private Sub AXCodCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then BuscaContrato (AXCodCta.NroCuenta)
End Sub

'Valida el campo AXDesCon.DescLote
Private Sub AXDesCon_KeyPressDesLot(KeyAscii As Integer)
KeyAscii = fgIntfMayusculas(KeyAscii)
KeyAscii = fgIntfLineas(AXDesCon.DescLote, KeyAscii, 14)
If Len(Trim(fgEliminaEnters(AXDesCon.DescLote))) > 0 Then
    cmdGrabar.Enabled = True
Else
    cmdGrabar.Enabled = False
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
lsEstados = gColPEstRegis '& "," & gColPEstDesem

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
    limpiar
    AXDesCon.EnabledDescLot = False
    cmdGrabar.Enabled = False
    cmdBuscar.Enabled = True
    AXCodCta.Enabled = True
    AXCodCta.SetFocusCuenta
End Sub

Private Sub CmdGrabar_Click()
'On Error GoTo ControlError
Dim loContFunct As COMNContabilidad.NCOMContFunciones
Dim loGrabar As COMNColoCPig.NCOMColPContrato
Dim loMovAnterior As COMDColocPig.DCOMColPFunciones
Dim lnMovNroAnt As Long
Dim lsMovNro As String
Dim lsFechaHoraGrab As String
Dim lsCuenta As String
Dim lsLote As String

lsCuenta = AXCodCta.NroCuenta
lsLote = fgEliminaEnters(AXDesCon.DescLote) & vbCr

'*** PEAC 20161220
If Trim(Me.txtMotivo.Text) = "" Then
    MsgBox "Por favor ingrese el motivo de la anulación del crédito Pignoraticio.", vbExclamation + vbOKOnly, "Atención"
    Me.txtMotivo.SetFocus
    Exit Sub
End If
'*** FIN PEAC

'If Len(Trim(fgEliminaEnters(AXDesCon.DescLote))) > 0 Then
    If MsgBox(" Grabar Anulación de Credito Pignoraticio ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
        AXDesCon.EnabledDescLot = False
        cmdGrabar.Enabled = False
        'Obtiene el Mov Nro Anterior
        Set loMovAnterior = New COMDColocPig.DCOMColPFunciones
            lnMovNroAnt = loMovAnterior.dObtieneMovNroAnterior(lsCuenta, geColPRegContrato)
        Set loMovAnterior = Nothing
            
        'Genera el Mov Nro
        Set loContFunct = New COMNContabilidad.NCOMContFunciones
            lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        Set loContFunct = Nothing
        
        lsFechaHoraGrab = fgFechaHoraGrab(lsMovNro)
        Set loGrabar = New COMNColoCPig.NCOMColPContrato
            'Grabar la Modificacion
            Call loGrabar.nAnulaCredPignoraticio(lsCuenta, gColPEstAnula, lsFechaHoraGrab, lsMovNro, lnMovNroAnt, True, Trim(Me.txtMotivo.Text))
        
            ''*** PEAC 20090126
            objPista.InsertarPista gsOpeCod, lsMovNro, gsCodPersUser, GetMaquinaUsuario, gEliminar, "Anula el credito pignoraticio", lsCuenta, gCodigoCuenta
            
        Set loGrabar = Nothing
        
        limpiar
        cmdBuscar.Enabled = True
        'fraListado.Visible = False
        AXCodCta.Enabled = True
        Me.AXCodCta.EnabledProd = True 'ARLO20190304
        AXCodCta.SetFocus
    Else
        MsgBox " Grabación cancelada ", vbInformation, " Aviso "
    End If
'Else
'    MsgBox " Falta información " & vbCr & " No se puede Grabar Contrato ", vbInformation, " Aviso "
'End If
Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

'Permite salir del formulario actual
Private Sub cmdsalir_Click()
Unload Me
End Sub


'Permite inicializar el formulario actual
Private Sub Form_Load()
    Me.Icon = LoadPicture(App.Path & gsRutaIcono)
    limpiar
    Me.AXCodCta.NroCuenta = fgIniciaAxCuentaPignoraticio
    Me.AXCodCta.EnabledProd = True 'ARLO20190304
    
    Set objPista = New COMManejador.Pista
    gsOpeCod = gPigAnularContrato
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 And AXCodCta.Enabled = True Then 'F12
        Dim sCuenta As String
        sCuenta = frmValTarCodAnt.inicia(gColConsuPrendario, False)
        If sCuenta <> "" Then
            AXCodCta.NroCuenta = sCuenta
            AXCodCta.SetFocusCuenta
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objPista = Nothing
End Sub
