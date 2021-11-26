VERSION 5.00
Begin VB.Form frmColPCustodiaDiferida 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Crédito Pignoraticio - Cobrar Custodia Diferida"
   ClientHeight    =   6780
   ClientLeft      =   390
   ClientTop       =   2145
   ClientWidth     =   8085
   Icon            =   "frmColPCustodiaDiferida.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   8085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   5685
      TabIndex        =   1
      Top             =   6270
      Width           =   975
   End
   Begin VB.Frame fraContenedor 
      Height          =   6180
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   7815
      Begin VB.CommandButton cmdBuscar 
         Height          =   345
         Left            =   7110
         Picture         =   "frmColPCustodiaDiferida.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Buscar ..."
         Top             =   270
         Width           =   420
      End
      Begin VB.Frame fraContenedor 
         Height          =   1575
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   4440
         Width           =   7455
         Begin SICMACT.EditMoney EditMoney1 
            Height          =   255
            Left            =   5820
            TabIndex        =   19
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "0"
         End
         Begin VB.TextBox txtFecPago 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            Left            =   1440
            TabIndex        =   15
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox txtNroBoletaVentaSerie 
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
            Height          =   300
            Left            =   5835
            MaxLength       =   3
            TabIndex        =   12
            Top             =   615
            Width           =   480
         End
         Begin VB.TextBox txtNroBoletaVenta 
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
            Height          =   300
            Left            =   6315
            MaxLength       =   9
            TabIndex        =   11
            Top             =   615
            Width           =   1050
         End
         Begin VB.TextBox txtCostoCustodiaExtra 
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
            Height          =   300
            Left            =   1440
            MaxLength       =   7
            TabIndex        =   6
            Top             =   600
            Width           =   1050
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Total a Pagar :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   1
            Left            =   4395
            TabIndex        =   24
            Top             =   1125
            Width           =   1245
         End
         Begin VB.Label LblTotalAPagar 
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
            Left            =   5805
            TabIndex        =   23
            Top             =   1110
            Width           =   930
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "I.T.F."
            Height          =   225
            Index           =   0
            Left            =   2520
            TabIndex        =   22
            Top             =   1125
            Width           =   630
         End
         Begin VB.Label LblItf 
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
            Left            =   3210
            TabIndex        =   21
            Top             =   1065
            Width           =   930
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Fecha Pago : "
            Height          =   225
            Index           =   7
            Left            =   120
            TabIndex        =   18
            Top             =   285
            Width           =   1035
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Dias Transcurridos : "
            Height          =   225
            Index           =   3
            Left            =   2640
            TabIndex        =   17
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label lblDiasTranscurridos 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   300
            Left            =   4080
            TabIndex        =   16
            Top             =   240
            Width           =   1200
         End
         Begin VB.Label lblSaldoCostoCustodia 
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
            Left            =   1425
            TabIndex        =   9
            Top             =   1065
            Width           =   930
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Saldo a pagar :"
            Height          =   225
            Index           =   13
            Left            =   150
            TabIndex        =   8
            Top             =   1110
            Width           =   1110
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Costo Custodia :"
            Height          =   225
            Index           =   6
            Left            =   135
            TabIndex        =   7
            Top             =   600
            Width           =   1305
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Nro Boleta :"
            Height          =   225
            Index           =   4
            Left            =   4905
            TabIndex        =   5
            Top             =   645
            Width           =   900
         End
      End
      Begin SICMACT.ActXCodCta AXCodCta 
         Height          =   375
         Left            =   120
         TabIndex        =   13
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
         Height          =   3615
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   6376
      End
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4485
      TabIndex        =   0
      Top             =   6270
      Width           =   975
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6885
      TabIndex        =   2
      Top             =   6270
      Width           =   975
   End
   Begin VB.Label lblMensaje 
      Caption         =   "<F12> Número Contrato Antiguo"
      Height          =   285
      Left            =   165
      TabIndex        =   10
      Top             =   6270
      Visible         =   0   'False
      Width           =   2655
   End
End
Attribute VB_Name = "frmColPCustodiaDiferida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************
'* COBRO DE CUSTODIA DIFERIDA
'Archivo:  frmColPCustodiaDiferida.frm
'LAYG   :  10/07/2001.
'Resumen:  Permite cobrar el Costo de Custodia
'          sobre las joyas cuando no se viene a recoger dentro del plazo limite.
Option Explicit

Dim pMaxDiasCustodiaDiferida As Double
Dim pTasaIGV As Double
Dim pPorcentajeCustodiaDiferida As Double
Dim sSql As String
Dim vNroContrato As String * 12
Dim vPlazo As Integer
Dim vTasaInteres As Double
Dim vCostoCustodiaExtra As Currency
Dim vImpuesto As Currency
Dim nRedondeoITF As Double

'Inicializa las variables  a utilizar en el formulario
Sub Limpiar()
    Me.AXCodCta.NroCuenta = fgIniciaAxCuentaPignoraticio
    AXDesCon.Limpiar
    txtCostoCustodiaExtra.Text = Format(0, "#0.00")
    lblSaldoCostoCustodia.Caption = Format(0, "#0.00")
    txtFecPago.Text = ""
    txtNroBoletaVentaSerie.Text = ""
    txtNroBoletaVenta.Text = ""
    vImpuesto = 0
    lblDiasTranscurridos.Caption = ""
    Me.LblItf.Caption = "0.00"
    Me.LblTotalAPagar.Caption = "0.00"
    nRedondeoITF = 0
 End Sub

'Busca el contrato ingresado
Private Sub BuscaContrato(ByVal psNroContrato As String)
Dim lbOk As Boolean
Dim lrValida As ADODB.Recordset
Dim loValContrato As COMNColoCPig.NCOMColPValida


Dim lsmensaje  As String
'On Error GoTo ControlError

    'Valida Contrato
    Set lrValida = New ADODB.Recordset
    Set loValContrato = New COMNColoCPig.NCOMColPValida
       ' Set lrValida = loValContrato.nValidaCustodiaDiferidaCredPignoraticio(psNroContrato, gdFecSis, 0)
        Set lrValida = loValContrato.nValidaCustodiaDiferidaCredPignoraticio(psNroContrato, gdFecSis, 0, gsCodUser, lsmensaje)
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
    If Not (lrValida.EOF And lrValida.BOF) Then
        'txtFecPago = Format(lrValida!dCancelado, "dd/mm/yyyy hh:mm")
        'MAVM 20120111 ***
        If txtFecPago = "" Then
            txtFecPago = Format(lrValida!dCancelado, "dd/mm/yyyy hh:mm")
        End If
        '***
        
        vCostoCustodiaExtra = lrValida!nCustodiaPag
    
        'lblDiasTranscurridos = DateDiff("d", txtFecPago, gdFecSis)
        'MAVM 20120111 ***
        If txtFecPago <> "" Then
            lblDiasTranscurridos = DateDiff("d", txtFecPago, gdFecSis)
        Else
            txtFecPago.Enabled = True
        End If
        '***
        
        'Calcula el Costo de Custodia Extra
        If val(lblDiasTranscurridos) > pMaxDiasCustodiaDiferida Then ' Dias Transcurridos Mayor al Plazo
            
            txtCostoCustodiaExtra.Text = Round(CalculaCostoCustodiaDiferida(val(lrValida!nTasacion), val(lblDiasTranscurridos), pPorcentajeCustodiaDiferida, pTasaIGV), 2)
            lblSaldoCostoCustodia.Caption = val(txtCostoCustodiaExtra) - vCostoCustodiaExtra
            vImpuesto = val(lblSaldoCostoCustodia.Caption) - (val(lblSaldoCostoCustodia.Caption) / (1 + pTasaIGV))
            vImpuesto = Round(vImpuesto, 2)
            If val(lblSaldoCostoCustodia.Caption) > 0 Then
                txtNroBoletaVentaSerie.Enabled = True
                txtNroBoletaVenta.Enabled = True
                txtNroBoletaVentaSerie.SetFocus
            Else
                MsgBox "Ya se ha cancelado el Costo de Custodia ", vbInformation + vbOKOnly, " Aviso "
                txtNroBoletaVentaSerie.Enabled = False
                txtNroBoletaVenta.Enabled = False
            End If
        Else
            MsgBox "Contrato no ha generado Costo de Custodia", vbInformation + vbOKOnly, " Aviso "
            txtCostoCustodiaExtra.Text = Format(0, "#0.00")
            txtNroBoletaVentaSerie.Enabled = False
            txtNroBoletaVenta.Enabled = False
        End If
        
            
        Set lrValida = Nothing
        
        LblItf.Caption = Abs(Format(fgITFCalculaImpuestoNOIncluido(CDbl(lblSaldoCostoCustodia.Caption)) - CDbl(lblSaldoCostoCustodia.Caption), "#0.00"))
        '*** BRGO 20110908 ************************************************
        nRedondeoITF = fgDiferenciaRedondeoITF(CCur(Me.LblItf.Caption))
        If nRedondeoITF > 0 Then
           Me.LblItf.Caption = Format(CCur(Me.LblItf.Caption) - nRedondeoITF, "#,##0.00")
        End If
        '*** END BRGO
        LblTotalAPagar.Caption = Abs(Format(CDbl(lblSaldoCostoCustodia.Caption) + CDbl(LblItf.Caption), "#0.00"))
        
        cmdGrabar.Enabled = True
        cmdGrabar.SetFocus
            
        AXCodCta.Enabled = False
        'cmdBuscar.Enabled = False
    End If
Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox "Error: " & Err.Number & " " & Err.Description & vbCr & _
        "Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

Private Sub AXCodCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then BuscaContrato (AXCodCta.NroCuenta)
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

'Cancela la operación actual y limpia los campos para iniciar otra operación
Private Sub cmdCancelar_Click()
    Limpiar
    txtNroBoletaVentaSerie.Enabled = False
    txtNroBoletaVenta.Enabled = False
    cmdBuscar.Enabled = True
    AXCodCta.Enabled = True
    AXCodCta.SetFocus
End Sub

'Permite actualizar los cambios en la base de datos
Private Sub cmdGrabar_Click()

'On Error GoTo ControlError
Dim loContFunct As COMNContabilidad.NCOMContFunciones
Dim loGrabarCustod As COMNColoCPig.NCOMColPContrato
Dim loImprime As COMNColoCPig.NCOMColPImpre
Dim loMov As COMDMov.DCOMMov
Dim loPrevio As previo.clsprevio
Dim lnMovNro As Long

Dim lsMovNro As String
Dim lsFechaHoraGrab As String
Dim lsCuenta As String
Dim lsCadImprimir As String

Dim lnSaldoCap As Currency, lnInteresComp As Currency, lnImpuesto As Currency
Dim lnCostoTasacion As Currency, lnCostoCustodia As Currency
Dim lnMontoEntregar As Currency

If MsgBox(" Grabar Costo Custodia de Contrato Pignoraticio ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
    cmdGrabar.Enabled = False
        
        'Genera el Mov Nro
        Set loContFunct = New COMNContabilidad.NCOMContFunciones
            lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        Set loContFunct = Nothing
        
        lsFechaHoraGrab = fgFechaHoraGrab(lsMovNro)
        Set loGrabarCustod = New COMNColoCPig.NCOMColPContrato
            'Grabar Costo Custodia Pignoraticio
            Call loGrabarCustod.nCustodiaDiferidaCredPignoraticio(AXCodCta.NroCuenta, lsFechaHoraGrab, _
                 lsMovNro, CCur(Me.lblSaldoCostoCustodia.Caption), CCur(Me.lblSaldoCostoCustodia.Caption) - vImpuesto, vImpuesto, _
                  3, txtNroBoletaVentaSerie & "-" & txtNroBoletaVenta, False, CDbl(LblItf.Caption))
        Set loGrabarCustod = Nothing
        '*** BRGO 20110916 ******************************
        If CCur(LblItf.Caption) > 0 Then
            Set loMov = New COMDMov.DCOMMov
            Call loMov.InsertaMovRedondeoITF(lsMovNro, 1, CCur(LblItf.Caption) + nRedondeoITF, CCur(LblItf.Caption))
            Set loMov = Nothing
        End If
        '*** END BRGO

        If MsgBox(" Imprimir Recibo de Cobro de Custodia ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
        Set loImprime = New COMNColoCPig.NCOMColPImpre
            lsCadImprimir = loImprime.nPrintReciboCobroCustodia(gsNomAge, lsFechaHoraGrab, AXCodCta.NroCuenta, AXDesCon.listaClientes.ListItems(1).ListSubItems.iTem(1), _
                   CCur(Me.lblSaldoCostoCustodia.Caption), 0, 0, gsCodUser, "", CDbl(LblItf.Caption), gImpresora)
        Set loImprime = Nothing
        Set loPrevio = New previo.clsprevio
            loPrevio.PrintSpool sLpt, lsCadImprimir, False, 22
            Do While True
                If MsgBox("Reimprimir Recibo de Cobro de Custodia ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
                    loPrevio.PrintSpool sLpt, lsCadImprimir, False, 22
                Else
                    Set loPrevio = Nothing
                    Exit Do
                End If
            Loop
            Set loPrevio = Nothing
        End If

        'Impresión
'        If MsgBox("Desea realizar impresión de recibo ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
'            ImprimirRecibo
'            Do While True
'                If MsgBox("Desea reimprimir Recibo ?", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
'                    ImprimirRecibo
'                Else
'                    Exit Do
'                End If
'            Loop
'        End If
'
'        If MsgBox(" Desea Imprimir BOLETA DE VENTA ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
'            ImprimirBoletaVenta
'            Do While True
'                If MsgBox("Desea Reimprimir BOLETA DE VENTA ?", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
'                    ImprimirBoletaVenta
'                Else
'                    Exit Do
'                End If
'            Loop
'        End If
        
        Limpiar
        
        cmdBuscar.Enabled = True
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
Private Sub cmdSalir_Click()
    Unload Me
End Sub

'Inicializa el formulario actual
Private Sub Form_Load()
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
    CargaParametros
    Limpiar
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

'Valida el campo txtcostocustodiaextra
Private Sub txtCostoCustodiaExtra_GotFocus()
    fEnfoque txtCostoCustodiaExtra
End Sub
Private Sub txtCostoCustodiaExtra_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtCostoCustodiaExtra, KeyAscii)
If KeyAscii = 13 Then
    'txtCostoCustodiaExtra.Enabled = False
    txtNroBoletaVenta.Enabled = True
    txtNroBoletaVenta.SetFocus
End If
End Sub

'MAVM 20120111 ***
Private Sub txtFecPago_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    BuscaContrato ((AXCodCta.NroCuenta))
End If
End Sub
'***

'Valida el campo txtnroBoletaVenta
Private Sub txtNroBoletaVenta_KeyPress(KeyAscii As Integer)
If KeyAscii <> vbKeySubtract Then
    KeyAscii = NumerosEnteros(KeyAscii)
End If
If KeyAscii = vbKeyBack And Len(txtNroBoletaVenta) <= 1 Then
   txtNroBoletaVentaSerie.SetFocus
End If
If KeyAscii = 13 Then
    If Len(Mid(txtNroBoletaVenta, 1, 3)) > 0 Then
        cmdGrabar.Enabled = True
        cmdGrabar.SetFocus
    Else
        MsgBox " Ingrese Número de Boleta de Venta", vbInformation, " Aviso "
    End If
End If
End Sub

Private Sub CargaParametros()
Dim loParam As COMDColocPig.DCOMColPCalculos
Set loParam = New COMDColocPig.DCOMColPCalculos
    pMaxDiasCustodiaDiferida = loParam.dObtieneColocParametro(gConsColPMaxDiasCustodiaDiferida)
    pTasaIGV = loParam.dObtieneColocParametro(gConsColPTasaIGV)
    pPorcentajeCustodiaDiferida = loParam.dObtieneColocParametro(gConsColPPorcentajeCustodiaDiferida)
Set loParam = Nothing
End Sub

Private Sub txtNroBoletaVentaSerie_KeyPress(KeyAscii As Integer)
If KeyAscii <> vbKeySubtract Then
    KeyAscii = NumerosEnteros(KeyAscii)
    If Len(txtNroBoletaVentaSerie) = 3 Then
        txtNroBoletaVenta.Enabled = True
        txtNroBoletaVenta.SetFocus
    End If
End If
If KeyAscii = 13 Then
    If Len(Mid(txtNroBoletaVentaSerie, 1, 3)) > 0 Then
        txtNroBoletaVenta.Enabled = True
        txtNroBoletaVenta.SetFocus
    Else
        MsgBox " Ingrese Nro Serie de Boleta de Venta", vbInformation, " Aviso "
    End If
End If
End Sub

