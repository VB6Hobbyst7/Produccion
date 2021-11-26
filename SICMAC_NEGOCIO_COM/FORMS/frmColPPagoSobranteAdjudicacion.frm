VERSION 5.00
Begin VB.Form frmColPPagoSobranteAdjudicacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Crédito Pignoraticio -  Pago de Sobrantes de Remate"
   ClientHeight    =   6150
   ClientLeft      =   885
   ClientTop       =   1845
   ClientWidth     =   7995
   Icon            =   "frmColPPagoSobranteAdjudicacion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   7995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraContenedor 
      Height          =   5490
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   7725
      Begin SICMACT.ActXColPDesCon AXDesCon 
         Height          =   3495
         Left            =   120
         TabIndex        =   19
         Top             =   720
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   6165
      End
      Begin SICMACT.ActXCodCta AXCodCta 
         Height          =   495
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   873
         Texto           =   "Crédito"
         EnabledCMAC     =   -1  'True
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
      Begin VB.Frame fraContenedor 
         Height          =   780
         Index           =   5
         Left            =   135
         TabIndex        =   4
         Top             =   4200
         Width           =   7425
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
            Left            =   3555
            TabIndex        =   9
            Top             =   480
            Width           =   975
         End
         Begin VB.TextBox txtPreVentaBruta 
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
            Left            =   6090
            TabIndex        =   8
            Top             =   150
            Width           =   1035
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
            Left            =   3555
            TabIndex        =   7
            Top             =   150
            Width           =   975
         End
         Begin VB.TextBox txtNroRemate 
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
            Left            =   1050
            TabIndex        =   6
            Top             =   165
            Width           =   975
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
            Left            =   6090
            TabIndex        =   5
            Top             =   450
            Width           =   1035
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Nro Remate "
            Height          =   225
            Index           =   3
            Left            =   90
            TabIndex        =   14
            Top             =   195
            Width           =   975
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Precio Venta Bruta :"
            Height          =   225
            Index           =   4
            Left            =   4620
            TabIndex        =   13
            Top             =   180
            Width           =   1485
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Comisión :"
            Height          =   225
            Index           =   5
            Left            =   2115
            TabIndex        =   12
            Top             =   480
            Width           =   765
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Precio Base Venta :"
            Height          =   225
            Index           =   6
            Left            =   2130
            TabIndex        =   11
            Top             =   180
            Width           =   1410
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Precio Venta Neto :"
            Height          =   225
            Index           =   11
            Left            =   4620
            TabIndex        =   10
            Top             =   465
            Width           =   1455
         End
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
         TabIndex        =   16
         Top             =   5040
         Width           =   1035
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Sobrante :"
         Height          =   225
         Index           =   9
         Left            =   4785
         TabIndex        =   15
         Top             =   5115
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6300
      TabIndex        =   2
      Top             =   5685
      Width           =   975
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3495
      TabIndex        =   0
      Top             =   5685
      Width           =   975
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4905
      TabIndex        =   1
      Top             =   5685
      Width           =   975
   End
   Begin VB.Label lblMensaje 
      Caption         =   "<F12> Número Contrato Antiguo"
      Height          =   285
      Left            =   240
      TabIndex        =   17
      Top             =   5730
      Width           =   2655
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
Dim fsVarNroRemate As String

'Permite inicializar las variables
Sub Limpiar()
    Me.AXCodCta.NroCuenta = fgIniciaAxCuentaPignoraticio
    AXDesCon.Limpiar
    txtNroRemate.Text = Format(0, "#0.00")
    txtPreBaseVenta.Text = Format(0, "#0.00")
    txtPreVentaBruta.Text = Format(0, "#0.00")
    txtPreVentaNeto.Text = Format(0, "#0.00")
    txtComision.Text = Format(0, "#0.00")
    lblSobrante.Caption = Format(0, "#0.00")
 
End Sub


'Busca el contrato ingresado
Private Sub BuscaContrato(ByVal psNroContrato As String)
Dim lbOk As Boolean
Dim lrValida As ADODB.Recordset
Dim loValContrato As COMNColoCPig.NCOMColPValida
Dim lsMensaje As String

On Error GoTo ControlError

    'Valida Contrato
    Set lrValida = New ADODB.Recordset
    Set loValContrato = New COMNColoCPig.NCOMColPValida
        Set lrValida = loValContrato.nValidaPagoSobranteRemateCredPignoraticio(psNroContrato, "R", lsMensaje)
        If Trim(lsMensaje) <> "" Then
            MsgBox lsMensaje, vbInformation, "Aviso"
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


    txtNroRemate.Text = lrValida!cNroProceso
    fsVarNroRemate = lrValida!cNroProceso
    
    txtPreBaseVenta = Format(lrValida!nRemSubBaseVta, "#0.00")
    txtPreVentaBruta.Text = Format(lrValida!nMontoProceso, "#0.00")
    txtComision = Format(lrValida!nComision, "#0.00")
    txtPreVentaNeto.Text = Format(lrValida!nMontoProceso - lrValida!nComision, "#0.00")
    lblSobrante.Caption = Format(lrValida!nSobrante, "#0.00")

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

Dim loImprime As COMNColoCPig.NCOMColPImpre
Dim loPrevio As previo.clsprevio
Dim lsNombreCliente As String
Dim lsCadImprimir As String
Dim lsMensaje As String

On Error GoTo ControlError

    lsNombreCliente = AXDesCon.listaClientes.ListItems(1).ListSubItems.iTem(1)
    ' Obtiene la cuenta de Ahorros ****
    Dim loDatRem As COMNColoCPig.NCOMColPRecGar
    Set loDatRem = New COMNColoCPig.NCOMColPRecGar
        lsCtaSobranteRemate = loDatRem.nObtieneCtaSobranteRemate(fsRemateCadaAgencia, lsMensaje)
        If Trim(lsMensaje) <> "" Then
             MsgBox lsMensaje, vbInformation, "Aviso"
             Exit Sub
        End If
    Set loDatRem = Nothing
    
'    If fsRemateCadaAgencia = gsCodCMAC & "00" Then  ' Remate Centralizado
'        Set loConstSis = New NConstSistemas
'            lsCtaSobranteRemate = loConstSis.LeeConstSistema(61)  ' Cuenta SobranteRemate
'        Set loConstSis = Nothing
'    Else ' Remate en Cada Agencia
'        Set PObjConec = New DConstante
'        Set lrCtaAho = PObjConec.RecuperaConstantes(3207, , "C.nConsValor")
'        If lrCtaAho.BOF And lrCtaAho.EOF Then
'           MsgBox "No se han configurado las Ctas de Ahorro de Sobrantes", vbInformation, "Aviso"
'           Exit Sub
'        End If
'        Do While Not lrCtaAho.EOF
'           If Val(lrCtaAho!nConsValor) = Val(Right(fsRemateCadaAgencia, 2)) Then
'               lsCtaSobranteRemate = Trim(lrCtaAho!cConsDescripcion)
'               Exit Do
'           End If
'           lrCtaAho.MoveNext
'        Loop
'        Set lrCtaAho = Nothing
'    End If
    
    If lsCtaSobranteRemate = "" Then
         MsgBox "No se encuentra configurada la Cta de Ahorros de Sobrante", vbInformation, "Aviso"
         Exit Sub
    End If

    If MsgBox("Está seguro de retirar Sobrante de Remate ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
        
        cmdGrabar.Enabled = False
       
        'Genera el Mov Nro
        Set loContFunct = New COMNContabilidad.NCOMContFunciones
            lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        Set loContFunct = Nothing
        
        lsFechaHoraGrab = fgFechaHoraGrab(lsMovNro)
        Set loGrabarVta = New COMNColoCPig.NCOMColPContrato
            'Grabar Pago Sobrante Remate
            Call loGrabarVta.nPagoSobranteRemate(AXCodCta.NroCuenta, fsVarNroRemate, lsFechaHoraGrab, _
                 lsMovNro, CCur(lblSobrante.Caption), lsCtaSobranteRemate, sLpt, False)
        Set loGrabarVta = Nothing

        'Impresión
        If MsgBox(" Imprimir Comprobante de Sobrante de Remate ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
            Set loImprime = New COMNColoCPig.NCOMColPImpre
                lsCadImprimir = loImprime.nPrintReciboPagoSobrante(gsNomAge, lsFechaHoraGrab, AXCodCta.NroCuenta, lsNombreCliente, _
                       CCur(lblSobrante.Caption), gsCodUser, "", gImpresora)
            Set loImprime = Nothing
            Set loPrevio = New previo.clsprevio
                loPrevio.PrintSpool sLpt, lsCadImprimir, False, 22
                Do While True
                    If MsgBox("Reimprimir Comprobante de Sobrante de Remate ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
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
Private Sub cmdSalir_Click()
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
End Sub

Private Sub CargaParametros()
'Dim loParam As DColPCalculos
Dim loConstSis As COMDConstSistema.NCOMConstSistema
Dim lnProcesoCadaAgencia As Integer

'Set loParam = New DColPCalculos
'    pTasaRemateComision = loParam.dObtieneColocParametro(gConsColPTasaComisionRemate)
'Set loParam = Nothing

Set loConstSis = New COMDConstSistema.NCOMConstSistema
    lnProcesoCadaAgencia = loConstSis.LeeConstSistema(121)  ' gConstSistPigRemateCadaAg
    If lnProcesoCadaAgencia = 1 Then  ' En cada agencia
        fsRemateCadaAgencia = gsCodCMAC & gsCodAge
    Else
        fsRemateCadaAgencia = gsCodCMAC & "00"
    End If
Set loConstSis = Nothing

End Sub

