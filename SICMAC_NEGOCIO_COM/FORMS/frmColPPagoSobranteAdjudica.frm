VERSION 5.00
Begin VB.Form frmColPPagoSobranteAdjudicacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Crédito Pignoraticio -  Pago de Sobrantes de Adjudicado"
   ClientHeight    =   6000
   ClientLeft      =   885
   ClientTop       =   1845
   ClientWidth     =   7995
   Icon            =   "frmColPPagoSobranteAdjudica.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   7995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtNroAdjudicacion 
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
      Left            =   1680
      TabIndex        =   10
      Top             =   4560
      Width           =   975
   End
   Begin VB.Frame fraContenedor 
      Height          =   5370
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   7725
      Begin VB.Frame fraContenedor 
         Height          =   660
         Index           =   5
         Left            =   120
         TabIndex        =   4
         Top             =   4200
         Width           =   7425
         Begin VB.TextBox txtTasacion 
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
            Left            =   6150
            TabIndex        =   15
            Top             =   240
            Width           =   1035
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
            Left            =   3510
            TabIndex        =   13
            Top             =   240
            Width           =   1035
         End
         Begin VB.Label lblEtiqueta 
            AutoSize        =   -1  'True
            Caption         =   "Valor Tasación :"
            Height          =   195
            Index           =   7
            Left            =   4920
            TabIndex        =   16
            Top             =   270
            Width           =   1155
         End
         Begin VB.Label lblEtiqueta 
            AutoSize        =   -1  'True
            Caption         =   "Deuda :"
            Height          =   195
            Index           =   2
            Left            =   2520
            TabIndex        =   14
            Top             =   270
            Width           =   570
         End
         Begin VB.Label lblEtiqueta 
            AutoSize        =   -1  'True
            Caption         =   "Nro Adjudicación"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   1215
         End
      End
      Begin SICMACT.ActXCodCta AXCodCta 
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   661
         Texto           =   "Crédito"
         EnabledCta      =   -1  'True
         EnabledAge      =   -1  'True
      End
      Begin SICMACT.ActXColPDesCon AXDesCon 
         Height          =   3615
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   6376
      End
      Begin VB.Label lblSobrante 
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
         Left            =   6240
         TabIndex        =   6
         Top             =   4920
         Width           =   1035
      End
      Begin VB.Label lblEtiqueta 
         AutoSize        =   -1  'True
         Caption         =   "Sobrante :"
         Height          =   195
         Index           =   9
         Left            =   5385
         TabIndex        =   5
         Top             =   4995
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6300
      TabIndex        =   2
      Top             =   5565
      Width           =   975
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3495
      TabIndex        =   0
      Top             =   5565
      Width           =   975
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4905
      TabIndex        =   1
      Top             =   5565
      Width           =   975
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "Nro Remate "
      Height          =   225
      Index           =   0
      Left            =   360
      TabIndex        =   11
      Top             =   5070
      Width           =   975
   End
   Begin VB.Label lblMensaje 
      AutoSize        =   -1  'True
      Caption         =   "<F12> Número Contrato Antiguo"
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   5610
      Width           =   2280
   End
End
Attribute VB_Name = "frmColPPagoSobranteAdjudicacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************
'* PAGO DE SOBRANTES.
'Archivo:  frmColPPagoSobranteRemate.frm
'LAYG   :  20/07/2001.
'ICA - 07/01/2005 - LAYG
'Resumen:  Nos permite pagar los sobrantes de los contratos que fueron rematados,
'          a sus respectivos dueños previa identificación.
Option Explicit

Dim pCtaAhoSob As String
Dim RegCredPrend As New ADODB.Recordset
Dim RegJoyas As New ADODB.Recordset
Dim RegDetRem As New ADODB.Recordset

Dim fsRemateCadaAgencia As String
Dim fsAdjudicaCadaAgencia As String

Dim fsVarNroRemate As String
Dim fsVarNroAdjudicacion As String ''*** PEAC 20090313

Dim objPista As COMManejador.Pista ''*** PEAC 20090313

'Permite inicializar las variables
Sub Limpiar()
    Me.AXCodCta.NroCuenta = fgIniciaAxCuentaPignoraticio
    AXDesCon.Limpiar
    
'    txtNroRemate.Text = Format(0, "#0.00")
    '*** PEAC 20090313
    Me.txtNroAdjudicacion.Text = Format(0, "#0.00")
    Me.txtDeuda.Text = Format(0, "#0.00")
    Me.txtTasacion.Text = Format(0, "#0.00")
    
'    txtPreBaseVenta.Text = Format(0, "#0.00")
'    txtPreVentaBruta.Text = Format(0, "#0.00")
'    txtPreVentaNeto.Text = Format(0, "#0.00")
'    txtComision.Text = Format(0, "#0.00")
    lblSobrante.Caption = Format(0, "#0.00")
 
End Sub


'Busca el contrato ingresado
Private Sub BuscaContrato(ByVal psNroContrato As String)
Dim lbOk As Boolean
Dim lrValida As ADODB.Recordset
Dim loValContrato As COMNColoCPig.NCOMColPValida
Dim lsmensaje As String

On Error GoTo ControlError

    'Valida Contrato
    Set lrValida = New ADODB.Recordset
    Set loValContrato = New COMNColoCPig.NCOMColPValida
        Set lrValida = loValContrato.nValidaPagoSobranteAdjudicadoCredPignoraticio(psNroContrato, "A", lsmensaje)
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
    'Muestra Datos
    lbOk = fgMuestraCredPig_AXDesCon(psNroContrato, Me.AXDesCon, False)


    'txtNroRemate.Text = lrValida!cNroProceso
    '*** PEAC 20090313
    Me.txtNroAdjudicacion.Text = lrValida!cNroProceso
    
    'fsVarNroRemate = lrValida!cNroProceso
    fsVarNroAdjudicacion = lrValida!cNroProceso
    
'    txtPreBaseVenta = Format(lrValida!nRemSubBaseVta, "#0.00")
'    txtPreVentaBruta.Text = Format(lrValida!nMontoProceso, "#0.00")
'    txtComision = Format(lrValida!nComision, "#0.00")
'    txtPreVentaNeto.Text = Format(lrValida!nMontoProceso - lrValida!nComision, "#0.00")
    
    '*** PEAC 20090313
    Me.txtDeuda.Text = Format(lrValida!nDeuda, "#0.00")
    Me.txtTasacion.Text = Format(lrValida!nTasacion, "#0.00")
    
    'lblSobrante.Caption = Format(lrValida!nSobrante, "#0.00")
    '*** PEAC 20090313
    lblSobrante.Caption = Format(lrValida!nDevolver, "#0.00")

    Set lrValida = Nothing
        
    AXCodCta.Enabled = False
    cmdGrabar.Enabled = True
    cmdGrabar.SetFocus

Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox "Error: " & Err.Number & " " & Err.Description & vbCr & _
        "Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

Private Sub AXCodCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call BuscaContrato(AXCodCta.NroCuenta)
End Sub

'Permite cancelar un proceso iniciado
Private Sub cmdCancelar_Click()
    Limpiar
    cmdGrabar.Enabled = False
    AXCodCta.Enabled = True
    AXCodCta.SetFocus
End Sub

'Permite grabar la información del pago de sobrante
Private Sub cmdGrabar_Click()

Dim loContFunct As COMNContabilidad.NCOMContFunciones
Dim loGrabarVta As COMNColoCPig.NCOMColPContrato

Dim lsMovNro As String
Dim lsFechaHoraGrab As String

'Dim loConstSis As NConstSistemas
'Dim PObjConec As DConstante
'Dim lrCtaAho As ADODB.Recordset
Dim lsCtaSobranteRemate As String
Dim lsCtaSobranteAdjudicado As String

Dim loImprime As COMNColoCPig.NCOMColPImpre
Dim loPrevio As previo.clsPrevio
Dim lsNombreCliente As String
Dim lsCadImprimir As String
Dim lsmensaje As String

On Error GoTo ControlError

    lsNombreCliente = AXDesCon.listaClientes.ListItems(1).ListSubItems.iTem(1)
    ' Obtiene la cuenta de Ahorros ****
    Dim loDatRem As COMNColoCPig.NCOMColPRecGar
    Set loDatRem = New COMNColoCPig.NCOMColPRecGar
        lsCtaSobranteAdjudicado = loDatRem.nObtieneCtaSobranteAdjudicado(fsAdjudicaCadaAgencia, lsmensaje)
        If Trim(lsmensaje) <> "" Then
             MsgBox lsmensaje, vbInformation, "Aviso"
             Exit Sub
        End If
    Set loDatRem = Nothing
    
    If lsCtaSobranteAdjudicado = "" Then
         MsgBox "No se encuentra configurada la Cta de Ahorros de Sobrante", vbInformation, "Aviso"
         Exit Sub
    End If

    If MsgBox("Está seguro de retirar Sobrante de Adjudicado ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
        
        cmdGrabar.Enabled = False
       
        'Genera el Mov Nro
        Set loContFunct = New COMNContabilidad.NCOMContFunciones
            lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        Set loContFunct = Nothing
        
        lsFechaHoraGrab = fgFechaHoraGrab(lsMovNro)
        Set loGrabarVta = New COMNColoCPig.NCOMColPContrato
'            'Grabar Pago Sobrante Remate
'            Call loGrabarVta.nPagoSobranteRemate(AXCodCta.NroCuenta, fsVarNroRemate, lsFechaHoraGrab, _
'                 lsMovNro, CCur(lblSobrante.Caption), lsCtaSobranteRemate, sLpt, False)
                 
            '*** PEAC 20090313
            'Grabar Pago Sobrante Adjudicado
            Call loGrabarVta.nPagoSobranteAdjudicado(AXCodCta.NroCuenta, fsVarNroAdjudicacion, lsFechaHoraGrab, _
                 lsMovNro, CCur(lblSobrante.Caption), lsCtaSobranteAdjudicado, sLpt, False)
        
            ''*** PEAC 20090313
            objPista.InsertarPista gsOpeCod, lsMovNro, gsCodPersUser, GetMaquinaUsuario, gInsertar, "Registrar Pago sobrante adjudicado", AXCodCta.NroCuenta, gCodigoCuenta
                    
        Set loGrabarVta = Nothing

        'Impresión
        'If MsgBox(" Imprimir Comprobante de Sobrante de Remate ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
        If MsgBox(" Imprimir Comprobante de Sobrante de Adjudicado ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
            Set loImprime = New COMNColoCPig.NCOMColPImpre
                lsCadImprimir = loImprime.nPrintReciboPagoSobrante(gsNomAge, lsFechaHoraGrab, AXCodCta.NroCuenta, lsNombreCliente, _
                       CCur(lblSobrante.Caption), gsCodUser, "", gImpresora)
            Set loImprime = Nothing
            Set loPrevio = New previo.clsPrevio
                loPrevio.PrintSpool sLpt, lsCadImprimir, False, 22
                Do While True
                    'If MsgBox("Reimprimir Comprobante de Sobrante de Remate ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
                    If MsgBox("Reimprimir Comprobante de Sobrante de Adjudicado ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
                        loPrevio.PrintSpool sLpt, lsCadImprimir, False, 22
                    Else
                        Set loPrevio = Nothing
                        Exit Do
                    End If
                Loop
           Set loPrevio = Nothing
        End If
    End If
        '************************************
Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
        
End Sub

'Permite salir del formulario actual
Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 And AXCodCta.Enabled = True Then 'F12
        Dim sCuenta As String
        sCuenta = frmValTarCodAnt.Inicia(gColConsuPrendario, False)
        If sCuenta <> "" Then
            AXCodCta.NroCuenta = sCuenta
            AXCodCta.SetFocusCuenta
        End If
   ElseIf KeyCode = 13 And AXCodCta.EnabledCta And AXCodCta.Age <> "" And Trim(AXCodCta.cuenta) = "" Then
        AXCodCta.SetFocusCuenta
        Exit Sub
   End If
End Sub

'Permite inicializar el formulario actual
Private Sub Form_Load()
    CargaParametros
    Limpiar
    
   Set objPista = New COMManejador.Pista
   gsOpeCod = gPigPagoSobranteAdjudicado
    
End Sub

Private Sub CargaParametros()
'Dim loParam As DColPCalculos
Dim loConstSis As COMDConstSistema.NCOMConstSistema
Dim lnProcesoCadaAgencia As Integer

Set loConstSis = New COMDConstSistema.NCOMConstSistema
    lnProcesoCadaAgencia = loConstSis.LeeConstSistema(121)
    If lnProcesoCadaAgencia = 1 Then  ' En cada agencia
        fsAdjudicaCadaAgencia = gsCodCMAC & gsCodAge
    Else
        fsAdjudicaCadaAgencia = gsCodCMAC & "00"
    End If
Set loConstSis = Nothing

End Sub
