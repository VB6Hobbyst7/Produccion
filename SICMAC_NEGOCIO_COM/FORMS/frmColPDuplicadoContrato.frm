VERSION 5.00
Begin VB.Form frmColPDuplicadoContrato 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Crédito Pignoraticio - Duplicado de Contrato u Hoja Resumen"
   ClientHeight    =   6390
   ClientLeft      =   555
   ClientTop       =   1800
   ClientWidth     =   7995
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   7995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraContenedor 
      Height          =   5745
      Index           =   0
      Left            =   75
      TabIndex        =   4
      Top             =   60
      Width           =   7830
      Begin VB.CommandButton cmdBuscar 
         Height          =   345
         Left            =   7200
         Picture         =   "frmColPDuplicadoContrato.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Buscar ..."
         Top             =   270
         Width           =   420
      End
      Begin VB.Frame fraContenedor 
         Enabled         =   0   'False
         Height          =   915
         Index           =   2
         Left            =   240
         TabIndex        =   9
         Top             =   4320
         Width           =   7395
         Begin VB.Label lblEtiqueta 
            Caption         =   "Nro.Duplicado :"
            Height          =   255
            Index           =   19
            Left            =   4680
            TabIndex        =   19
            Top             =   240
            Width           =   1275
         End
         Begin VB.Label lblNroDuplic 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
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
            Height          =   285
            Left            =   6120
            TabIndex        =   18
            Top             =   240
            Width           =   840
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Impuesto "
            Height          =   255
            Index           =   17
            Left            =   2430
            TabIndex        =   17
            Top             =   540
            Width           =   855
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Interes"
            Height          =   255
            Index           =   16
            Left            =   180
            TabIndex        =   16
            Top             =   540
            Width           =   795
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Cost. Custod."
            Height          =   255
            Index           =   18
            Left            =   2430
            TabIndex        =   15
            Top             =   180
            Width           =   990
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Cost. Tasac."
            Height          =   255
            Index           =   2
            Left            =   180
            TabIndex        =   14
            Top             =   210
            Width           =   1005
         End
         Begin VB.Label lblCostoCustodia 
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
            Height          =   270
            Left            =   3420
            TabIndex        =   13
            Top             =   180
            Width           =   1035
         End
         Begin VB.Label lblCostoTasacion 
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
            Height          =   270
            Left            =   1260
            TabIndex        =   12
            Top             =   180
            Width           =   1035
         End
         Begin VB.Label lblInteres 
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
            Height          =   270
            Left            =   1260
            TabIndex        =   11
            Top             =   480
            Width           =   1035
         End
         Begin VB.Label lblImpuesto 
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
            Height          =   270
            Left            =   3420
            TabIndex        =   10
            Top             =   480
            Width           =   1035
         End
      End
      Begin SICMACT.ActXCodCta AXCodCta 
         Height          =   375
         Left            =   180
         TabIndex        =   0
         Top             =   270
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
         TabIndex        =   8
         Top             =   720
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   6165
      End
      Begin VB.Label LblITF 
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
         Height          =   270
         Left            =   3870
         TabIndex        =   22
         Top             =   5340
         Width           =   1035
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "I.T.F."
         Height          =   255
         Index           =   0
         Left            =   3345
         TabIndex        =   21
         Top             =   5355
         Width           =   645
      End
      Begin VB.Label lblCostoDuplicado 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
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
         Height          =   315
         Left            =   1890
         TabIndex        =   6
         Top             =   5295
         Width           =   1290
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Costo Duplicado :  S/."
         Height          =   255
         Index           =   11
         Left            =   240
         TabIndex        =   5
         Top             =   5355
         Width           =   1680
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   5655
      TabIndex        =   2
      Top             =   5910
      Width           =   975
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6855
      TabIndex        =   3
      Top             =   5910
      Width           =   975
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   345
      Left            =   4455
      TabIndex        =   1
      Top             =   5910
      Width           =   1005
   End
   Begin VB.Label lblMensaje 
      Caption         =   "<F12> Número Contrato Antiguo"
      Height          =   285
      Left            =   255
      TabIndex        =   7
      Top             =   5910
      Width           =   2655
   End
End
Attribute VB_Name = "frmColPDuplicadoContrato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************
'* DUPLICADO DE CONTRATO.
'Archivo:  frmColPDuplicadoContrato.frm
'LAYG   :  10/07/2001.
'Resumen:  Permite reimprimir el Contrato pignoraticio

Option Explicit
Dim pCostoDuplicado As Double
Dim RegCredPrend As New ADODB.Recordset
Dim RegPerCta As New ADODB.Recordset
Dim vNroContrato As String
Dim vNetoARecibir As Double
Dim fnTasaInteresAdelantado As Double
Dim nRedondeoITF As Double


'Permite inicializarlas variables del formulario
Private Sub Limpiar()
    Me.AXCodCta.NroCuenta = fgIniciaAxCuentaPignoraticio
    AXDesCon.Limpiar
    'lblCostoDuplicado.Caption = Format(pCostoDuplicado, "#0.00")
    'LblITF.Caption = fgITFCalculaImpuestoNOIncluido(CDbl(lblCostoDuplicado.Caption)) - CDbl(lblCostoDuplicado.Caption)
    'lblCostoDuplicado.Caption = fgITFCalculaImpuestoNOIncluido(CDbl(lblCostoDuplicado.Caption))
    
    lblNroDuplic.Caption = ""
    Me.lblCostoTasacion = "0.00"
    Me.lblCostoCustodia = "0.00"
    Me.lblNroDuplic = ""
    Me.lblInteres = "0.00"
    Me.lblImpuesto = "0.00"
    Me.lblCostoDuplicado = "0.00"
    Me.LblITF.Caption = "0.00"
    nRedondeoITF = 0
End Sub

'Permite buscar el contrato ingresado
Private Sub BuscaContrato(ByVal psNroContrato As String)

Dim lbok As Boolean
Dim lrValida As ADODB.Recordset
Dim loValContrato As COMNColoCPig.NCOMColPValida
Dim lsmensaje As String
'On Error GoTo ControlError

    'Valida Contrato
    Set lrValida = New ADODB.Recordset
    Set loValContrato = New COMNColoCPig.NCOMColPValida
        Set lrValida = loValContrato.nValidaDuplicadoContratoCredPignoraticio(psNroContrato, lsmensaje)
        If Trim(lsmensaje) <> "" Then
            If Trim(lsmensaje) = "Contrato se encuentra" Then lsmensaje = "Contrato se encuentra en estado NO vigente"
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
    lbok = fgMuestraCredPig_AXDesCon(psNroContrato, Me.AXDesCon, False)

    fnTasaInteresAdelantado = lrValida!nTasaInteres
    Me.lblInteres = Format(lrValida!nInteres, "#0.00")
    Me.lblImpuesto = Format(lrValida!nImpuesto, "#0.00")
    Me.lblCostoTasacion = Format(lrValida!nTasacion, "#0.00")
    Me.lblCostoCustodia = Format(lrValida!nCustodia, "#0.00")
    
    Me.lblNroDuplic = Format(lrValida!nNroDuplic + 1, "#0")
    
    Set lrValida = Nothing
    
    Me.lblCostoDuplicado = Format(pCostoDuplicado, "#0.00")
    LblITF.Caption = Format(fgITFCalculaImpuestoNOIncluido(CDbl(lblCostoDuplicado.Caption)) - CDbl(lblCostoDuplicado.Caption), "#0.00")
    '*** BRGO 20110908 ************************************************
    nRedondeoITF = fgDiferenciaRedondeoITF(CCur(Me.LblITF.Caption))
    If nRedondeoITF > 0 Then
       Me.LblITF.Caption = Format(CCur(Me.LblITF.Caption) - nRedondeoITF, "#,##0.00")
       'Me.lblTotal.Caption = Format(CCur(Me.lblITF.Caption) + txtMonto.value, "#,##0.00")
    End If
    '*** END BRGO
    lblCostoDuplicado.Caption = Format(fgITFCalculaImpuestoNOIncluido(CDbl(lblCostoDuplicado.Caption)), "#0.00")
    AXCodCta.Enabled = False
    cmdImprimir.Enabled = True
   ' cmdImprimir.SetFocus

Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & err.Number & " " & err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "

End Sub

Private Sub AXCodCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call BuscaContrato(AXCodCta.NroCuenta)
End Sub

Private Sub cmdBuscar_Click()
Dim loPers As COMDPersona.UCOMPersona
Dim lsPersCod As String, lsPersNombre As String
Dim lsEstados As String
Dim loPersContrato As COMDColocPig.DCOMColPContrato
Dim lrContratos As ADODB.Recordset
Dim loCuentas As COMDPersona.UCOMProdPersona

On Error GoTo ControlError

Set loPers = New COMDPersona.UCOMPersona
    Set loPers = frmBuscaPersona.Inicio
    If loPers Is Nothing Then Exit Sub
    lsPersCod = loPers.sPersCod
    lsPersNombre = loPers.sPersNombre
Set loPers = Nothing

' Selecciona Estados
lsEstados = gColPEstRegis & "," & gColPEstDesem & "," & gColPEstVenci & "," & gColPEstPRema & "," & _
            gColPEstRenov & "," & gColPEstDifer & "," & gColPEstCance

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
    MsgBox " Error: " & err.Number & " " & err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "

End Sub

'Permite cancelar un proceso e inicializar los campos para otro proceso
Private Sub cmdCancelar_Click()
    Limpiar
    cmdImprimir.Enabled = False
    AXCodCta.Enabled = True
    AXCodCta.SetFocusCuenta
End Sub

Private Sub cmdImprimir_Click()

'On Error GoTo ControlError
Dim loContFunct As COMNContabilidad.NCOMContFunciones
Dim loGrabarDup As COMNColoCPig.NCOMColPContrato
Dim loImprime As COMNColoCPig.NCOMColPImpre
Dim loPrevio As previo.clsprevio
Dim loRegPig As COMDColocPig.DCOMColPActualizaBD
Dim oMov As COMDMov.DCOMMov

Dim lsMovNro As String
Dim lsFechaHoraGrab As String

Dim lnNumDuplicado As Integer
Dim lnMontoTransaccion As Currency
Dim lsCadImprimir As String
Dim lsNombreCliente As String
Dim lsmensaje As String
'Dim lsLote As String
Dim lrPersonas As ADODB.Recordset
'RIRO20210923 campana prendario
Dim nCampana As Integer
'RIRO20210923 campana prendario

lnNumDuplicado = Val(Me.lblNroDuplic.Caption)
lnMontoTransaccion = CCur(Me.lblCostoDuplicado.Caption) - CDbl(LblITF.Caption)


lsNombreCliente = AXDesCon.listaClientes.ListItems(1).ListSubItems.Item(1)
Set lrPersonas = fgGetCodigoPersonaListaRsNew(Me.AXDesCon.listaClientes)
'lsLote = fgEliminaEnters(Me.AXDesCon.DescLote) & vbCr
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
            Set rsPersona = oDPersona.ObtenerUltimaVisita(Trim(rsPersonaCred!cperscod))
            If rsPersona.RecordCount > 0 Then
                If Not (rsPersona.EOF And rsPersona.BOF) Then
                    If Trim(rsPersona!sUsual) = "3" Then
                    MsgBox PstaNombre(Trim(rsPersonaCred!cPersNombre), True) & "." & Chr(10) & "CLIENTE OBSERVADO: " & Trim(rsPersona!cVisObserva), vbInformation, "Aviso"
                        Call frmPersona.Inicio(Trim(rsPersonaCred!cperscod), PersonaActualiza)
                    End If
                End If
            End If
            Call VerSiClienteActualizoAutorizoSusDatos(Trim(rsPersonaCred!cperscod), gColPOpeImpDuplicado) 'FRHU ERS077-2015 20151204
            Set rsPersona = Nothing
            rsPersonaCred.MoveNext
        Next Cont
    End If
End If
'WIOR FIN ***************************************************************
If MsgBox(" Grabar Duplicado de Contrato Pignoraticio ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
    cmdImprimir.Enabled = False

    'Genera el Mov Nro
    Set loContFunct = New COMNContabilidad.NCOMContFunciones
        lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Set loContFunct = Nothing
    
    lsFechaHoraGrab = fgFechaHoraGrab(lsMovNro)
    Set loGrabarDup = New COMNColoCPig.NCOMColPContrato
        'Grabar Duplicado de Contrato Pignoraticio
        Call loGrabarDup.nDuplicadoContratoCredPignoraticio(AXCodCta.NroCuenta, lnNumDuplicado - 1, lsFechaHoraGrab, _
              lsMovNro, lnMontoTransaccion, False, CDbl(LblITF.Caption))
    Set loGrabarDup = Nothing
    '*** BRGO 20110915 *****************************************
    If CDbl(LblITF.Caption) > 0 Then
        Set oMov = New COMDMov.DCOMMov
        Call oMov.InsertaMovRedondeoITF(lsMovNro, 1, CDbl(LblITF.Caption) + nRedondeoITF, CDbl(LblITF.Caption))
        Set oMov = Nothing
    End If
    '*** END BRGO ***************************************************
    
'*** PEAC 20161220

    ' *** Impresion del voucher del costo por duplicado
    If MsgBox(" Imprimir Recibo de Duplicado de Contrato ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
        Set loImprime = New COMNColoCPig.NCOMColPImpre
            lsCadImprimir = loImprime.nPrintReciboDuplicadoContrato(gsNomAge, lsFechaHoraGrab, AXCodCta.NroCuenta, lsNombreCliente, _
                   lnMontoTransaccion, lnNumDuplicado, fnTasaInteresAdelantado, gsCodUser, "", CDbl(LblITF.Caption), gImpresora)
        Set loImprime = Nothing
        Set loPrevio = New previo.clsprevio
            loPrevio.PrintSpool sLpt, lsCadImprimir, False, 22
            Do While True
                If MsgBox("Reimprimir Recibo de Duplicado de Contrato ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
                    loPrevio.PrintSpool sLpt, lsCadImprimir, False, 22
                Else
                    Set loPrevio = Nothing
                    Exit Do
                End If
            Loop
    End If


    If MsgBox("Imprimir Contrato Pignoraticio ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then

        Set loImprime = New COMNColoCPig.NCOMColPImpre
        Dim rsPig As ADODB.Recordset
        Dim rsPigJoyas As ADODB.Recordset
        Dim rsPigPers As ADODB.Recordset
        Dim rsPigCostos As ADODB.Recordset
        Dim rsPigDet As ADODB.Recordset
        Dim rsPigTasas As ADODB.Recordset
        Dim rsPigCosNot As ADODB.Recordset

        'Comento RIRO20210923 campana prendario
        'Call loImprime.RecuperaDatosHojaResumenPigno(AXCodCta.NroCuenta, rsPig, rsPigJoyas, rsPigPers, rsPigCostos, rsPigDet, rsPigTasas, rsPigCosNot)
        'Call CargaHojaResumenPignoPDF(rsPig, rsPigJoyas, rsPigPers, rsPigCostos, rsPigDet, rsPigTasas, rsPigCosNot)
        'Comento RIRO20210923 campana prendario
        
        'RIRO20210923 campana prendario
        nCampana = 0
        Call loImprime.RecuperaDatosHojaResumenPigno(AXCodCta.NroCuenta, rsPig, rsPigJoyas, rsPigPers, rsPigCostos, rsPigDet, rsPigTasas, rsPigCosNot, nCampana)
        Call CargaHojaResumenPignoPDF(rsPig, rsPigJoyas, rsPigPers, rsPigCostos, rsPigDet, rsPigTasas, rsPigCosNot, nCampana)
        'RIRO20210923 campana prendario
        
        Set loImprime = Nothing
        
'    ' *** Impresion
'    If MsgBox(" Imprimir Recibo de Duplicado de Contrato ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
'        Set loImprime = New COMNColoCPig.NCOMColPImpre
'            lsCadImprimir = loImprime.nPrintReciboDuplicadoContrato(gsNomAge, lsFechaHoraGrab, AXCodCta.NroCuenta, lsNombreCliente, _
'                   lnMontoTransaccion, lnNumDuplicado, fnTasaInteresAdelantado, gsCodUser, "", CDbl(LblITF.Caption), gImpresora)
'        Set loImprime = Nothing
'        Set loPrevio = New previo.clsprevio
'            loPrevio.PrintSpool sLpt, lsCadImprimir, False, 22
'            Do While True
'                If MsgBox("Reimprimir Recibo de Duplicado de Contrato ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
'                    loPrevio.PrintSpool sLpt, lsCadImprimir, False, 22
'                Else
'                    Set loPrevio = Nothing
'                    Exit Do
'                End If
'            Loop
'    End If
'
'    If MsgBox("Imprimir Duplicado de Contrato Pignoraticio ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
'        Set loImprime = New COMNColoCPig.NCOMColPImpre
'
''        lsCadImprimir = loRegImp.nPrintContratoPignoraticioDet(lsContrato, True, lrPersonas, fnTasaInteresAdelantado, _
''                lnMontoPrestamo, lsFechaHoraGrab, Format(lsFechaVenc, "mm/dd/yyyy"), lnPlazo, lnOroBruto, lnOroNeto, lnValTasacion, _
''                lnPiezas, lsLote, ln14k, ln16k, ln18k, ln21k, lnIntAdelantado, lnCostoTasac, lnCostoCustodia, lnImpuesto, gsCodUser)
'
'            lsCadImprimir = loImprime.nPrintContratoPignoraticioDet(AXCodCta.NroCuenta, True, , , , , , , , , , _
'                                    , , , , , , Format(lblInteres.Caption, "#0.00"), , , , gsCodUser, lnNumDuplicado, lsmensaje, gImpresora)
'            If Trim(lsmensaje) <> "" Then
'                MsgBox lsmensaje, vbInformation, "Aviso"
'                Exit Sub
'            End If
'
'        Set loImprime = Nothing
'        Set loPrevio = New previo.clsprevio
'            loPrevio.PrintSpool sLpt, lsCadImprimir, False
'
'            Do While True
'                If MsgBox("Reimprimir Contrato Pignoraticio ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
'                    loPrevio.PrintSpool sLpt, Chr(10) & Chr(10) & Chr(10) & Chr(10) & lsCadImprimir, False
'                Else
'                    Set loPrevio = Nothing
'                    Exit Do
'                End If
'            Loop
'        Set loPrevio = Nothing
'    End If
'*** FIN PEAC

'        Limpiar
'        AXCodCta.Enabled = True
'        AXCodCta.SetFocus
    'INICIO JHCU ENCUESTA 16-10-2019
    Encuestas gsCodUser, gsCodAge, "ERS0292019", gColPOpeImpDuplicado
    'FIN
    Else
        MsgBox " Impresión cancelada ", vbInformation, " Aviso "
    End If
    Limpiar
    AXCodCta.Enabled = True
    AXCodCta.SetFocus
    
Else
    MsgBox " Grabación cancelada ", vbInformation, " Aviso "
End If

Exit Sub


ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & err.Number & " " & err.Description & vbCr & _
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
            SendKeys "{Enter}"
        End If
    End If
End Sub

'Permite inicializar el formulario actual
Private Sub Form_Load()
    Me.Icon = LoadPicture(App.Path & gsRutaIcono)
    CargaParametros
    Limpiar
End Sub


Private Sub CargaParametros()
Dim loParam As COMDColocPig.DCOMColPCalculos
Set loParam = New COMDColocPig.DCOMColPCalculos
    pCostoDuplicado = loParam.dObtieneColocParametro(gConsColPCostoDuplicadoContrato)
Set loParam = Nothing
End Sub

