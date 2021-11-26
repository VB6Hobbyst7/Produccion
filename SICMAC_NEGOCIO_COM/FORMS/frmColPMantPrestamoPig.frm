VERSION 5.00
Begin VB.Form frmColPMantPrestamoPig 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Crédito Pignoraticio - Consulta"
   ClientHeight    =   5520
   ClientLeft      =   1275
   ClientTop       =   1770
   ClientWidth     =   7965
   Icon            =   "frmColPMantPrestamoPig.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   7965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
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
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   5040
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Frame fraContenedor 
      Height          =   4980
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   7800
      Begin SICMACT.ActXColPDesCon AXDesCon 
         Height          =   3615
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   7575
         _extentx        =   13361
         _extenty        =   6588
      End
      Begin SICMACT.ActXCodCta AXCodCta 
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   3615
         _extentx        =   6376
         _extenty        =   661
         texto           =   "Crédito"
         enabledcta      =   -1  'True
         enabledprod     =   -1  'True
      End
      Begin VB.CommandButton cmdBuscar 
         Height          =   345
         Left            =   6930
         Picture         =   "frmColPMantPrestamoPig.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Buscar ..."
         Top             =   270
         Width           =   420
      End
      Begin VB.Label lblTasador 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   3960
         TabIndex        =   11
         Top             =   4440
         Width           =   975
      End
      Begin VB.Label lblHolograma 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1440
         TabIndex        =   10
         Top             =   4440
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Tasador:"
         Height          =   255
         Left            =   3120
         TabIndex        =   9
         Top             =   4440
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "N° Holograma:"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   4440
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4320
      TabIndex        =   1
      Top             =   5040
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6720
      TabIndex        =   3
      Top             =   5040
      Width           =   975
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   5520
      TabIndex        =   2
      Top             =   5040
      Width           =   975
   End
End
Attribute VB_Name = "frmColPMantPrestamoPig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************
'* MANTENIMIENTO DE CONTRATO PIGNORATICIO
'Archivo:  frmColPMantPrestamoPig.frm
'LAYG   :  01/05/2001.
'Resumen:  Nos permite modificar la descripcion del lote
Option Explicit

Dim objPista As COMManejador.Pista

'Permite inicializarlas variables del formulario
Private Sub limpiar()
    Me.AXCodCta.NroCuenta = fgIniciaAxCuentaPignoraticio
    AXDesCon.limpiar
End Sub

'Permite buscar el contrato ingresado
Public Sub BuscaContrato(ByVal psNroContrato As String, Optional ByVal nTipoBusqueda As Integer = 0, Optional ByVal psPersCod As String)
Dim loValContrato As COMNColoCPig.NCOMColPValida
Dim lrValida As ADODB.Recordset
Dim lbOk As Boolean
Dim lsmensaje As String
Dim rs As ADODB.Recordset 'APRI20180705 ERS063-2017
'On Error GoTo ControlError

    'Valida Contrato
    Set lrValida = New ADODB.Recordset
    Set loValContrato = New COMNColoCPig.NCOMColPValida
        Set lrValida = loValContrato.nValidaAnulacionCredPignoraticio(psNroContrato, gdFecSis, , lsmensaje, nTipoBusqueda)
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
    'Muestra Datos
    If nTipoBusqueda = 0 Then
        lbOk = fgMuestraCredPig_AXDesCon(psNroContrato, Me.AXDesCon, True)
    Else
        limpiar
        lbOk = fgMuestraCredPig_AXDesCon(psNroContrato, Me.AXDesCon, True, 1)
        AXCodCta.NroCuenta = psNroContrato
        cmdGrabar.Visible = False
        cmdCancelar.Visible = False
        '** Juez 20120529 ****************
        If gsCodArea = "001" Then
            cmdImprimir.Visible = True
            cmdImprimir.Enabled = True
        End If
        '** End Juez *********************
    End If
    If lbOk = False Then
        limpiar
        AXCodCta.SetFocusCuenta
        Exit Sub
    End If
    AXCodCta.Enabled = False
    cmdBuscar.Enabled = False
    If nTipoBusqueda = 0 Then
        AXDesCon.SetFocusDesLot
    End If
    'APRI20180705 ERS063-2017
    Set rs = ObtieneTasadorHolograma(psNroContrato)
    lblHolograma.Caption = rs!nHolograma
    lblTasador.Caption = rs!cUser
    Set rs = Nothing
    'END APRI
    
    Set lrValida = Nothing
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
lsEstados = gColPEstRegis ' & "," & gColPEstDesem

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
    'AXCodCta.SetFocusCuenta
    'MAVM 20100609 BAS II
    AXCodCta.SetFocusProd
End Sub

'Permite la impresión del duplicado del contrato deseado así como la actulaización
' en la base de datos de los campos respectivos
Private Sub CmdGrabar_Click()
'On Error GoTo ControlError
Dim loContFunct As COMNContabilidad.NCOMContFunciones
Dim loGrabarLote As COMNColoCPig.NCOMColPContrato
Dim loImprime As COMNColoCPig.NCOMColPImpre
Dim loPrevio As previo.clsprevio

Dim lsMovNro As String
Dim lsFechaHoraGrab As String
Dim lsCuenta As String
Dim lsLote As String
Dim lsCadImprimir As String

Dim lsmensaje As String

lsCuenta = AXCodCta.NroCuenta
lsLote = fgEliminaEnters(AXDesCon.DescLote) & vbCr

If Len(Trim(fgEliminaEnters(AXDesCon.DescLote))) > 0 Then
    If MsgBox(" Grabar Cambios de la descripción ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
        AXDesCon.EnabledDescLot = False
        cmdGrabar.Enabled = False
        
        'Genera el Mov Nro
        Set loContFunct = New COMNContabilidad.NCOMContFunciones
            lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        Set loContFunct = Nothing
        
        lsFechaHoraGrab = fgFechaHoraGrab(lsMovNro)
        Set loGrabarLote = New COMNColoCPig.NCOMColPContrato
            'Grabar la Modificacion
            Call loGrabarLote.nModificaLoteCredPignoraticio(lsCuenta, lsFechaHoraGrab, lsLote, lsMovNro, True)
        Set loGrabarLote = Nothing
        
        '' *** PEAC 20090126
        objPista.InsertarPista gsOpeCod, lsMovNro, gsCodPersUser, GetMaquinaUsuario, gModificar, , lsCuenta, gCodigoCuenta

        
        ' Imprimir
        If MsgBox("Imprimir Contrato Pignoraticio ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
            Set loImprime = New COMNColoCPig.NCOMColPImpre
                lsCadImprimir = loImprime.nPrintContratoPignoraticio(AXCodCta.NroCuenta, True, , , , , , , , , , _
                                        , , , , , , , , , , gsCodUser, , lsmensaje, gImpresora)
            If Trim(lsmensaje) <> "" Then
                 MsgBox lsmensaje, vbInformation, "Aviso"
                 Exit Sub
            End If
            
            Set loImprime = Nothing
            Set loPrevio = New previo.clsprevio
                loPrevio.PrintSpool sLpt, lsCadImprimir, False
                Do While True
                    If MsgBox("Reimprimir Contrato Pignoraticio ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
                        loPrevio.PrintSpool sLpt, lsCadImprimir, False
                    Else
                        Set loPrevio = Nothing
                        Exit Do
                    End If
                Loop
        End If
       
        limpiar
        cmdBuscar.Enabled = True
        
        AXCodCta.Enabled = True
        AXCodCta.SetFocus
    Else
        MsgBox " Grabación cancelada ", vbInformation, " Aviso "
    End If
Else
    MsgBox " Falta información " & vbCr & " No se puede Grabar Contrato ", vbInformation, " Aviso "
End If
Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

'** Juez 20120529 ****************************************
Private Sub cmdImprimir_Click()
    Dim loRep As COMNColoCPig.NCOMColPRepo
    
    Dim lsCadImp As String
    Dim loPrevio As previo.clsprevio
    Set loRep = New COMNColoCPig.NCOMColPRepo
    loRep.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis
    lsCadImp = loRep.ImprimeConsultaCreditoPignoraticio(AXCodCta.NroCuenta, gImpresora, gdFecSis)
    Set loRep = Nothing
    
    If Len(Trim(lsCadImp)) > 0 Then
        Set loPrevio = New previo.clsprevio
        loPrevio.Show lsCadImp, "Consulta de Créditos Pignoraticios", True
        Set loPrevio = Nothing
    Else
        MsgBox "No Existen Datos para el reporte", vbInformation, "Aviso"
    End If
End Sub
'** End Juez *********************************************

'Permite salir del formulario actual
Private Sub cmdsalir_Click()
    Unload Me
End Sub

'Permite inicializar el formulario actual
Private Sub Form_Load()
    Me.Icon = LoadPicture(App.Path & gsRutaIcono)
    limpiar
    Me.AXCodCta.NroCuenta = fgIniciaAxCuentaPignoraticio
    
    Set objPista = New COMManejador.Pista
    gsOpeCod = gPigConsultarContro
    
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
    cmdGrabar.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objPista = Nothing
End Sub
