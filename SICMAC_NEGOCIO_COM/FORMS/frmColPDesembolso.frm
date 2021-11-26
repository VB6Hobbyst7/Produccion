VERSION 5.00
Begin VB.Form frmColPDesembolso 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Desembolso Pignoraticio"
   ClientHeight    =   7065
   ClientLeft      =   1440
   ClientTop       =   1980
   ClientWidth     =   7935
   HelpContextID   =   210
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCtasAhorros 
      Caption         =   "Cuentas Ahorros"
      Height          =   375
      Left            =   720
      TabIndex        =   22
      Top             =   6600
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   5415
      TabIndex        =   1
      Top             =   6585
      Width           =   1095
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6615
      TabIndex        =   2
      Top             =   6585
      Width           =   1095
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4215
      TabIndex        =   0
      Top             =   6585
      Width           =   1095
   End
   Begin VB.Frame fraContenedor 
      Height          =   6420
      Index           =   1
      Left            =   105
      TabIndex        =   3
      Top             =   90
      Width           =   7725
      Begin VB.CommandButton cmdBuscar 
         Height          =   345
         Left            =   7110
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Buscar ..."
         Top             =   180
         Width           =   420
      End
      Begin VB.Frame fraContenedor 
         Enabled         =   0   'False
         Height          =   2055
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   4200
         Width           =   7480
         Begin VB.TextBox txtImpuesto 
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
            Height          =   255
            Left            =   3930
            TabIndex        =   8
            Top             =   480
            Width           =   1170
         End
         Begin VB.TextBox txtInteres 
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
            Left            =   3945
            TabIndex        =   7
            Top             =   165
            Width           =   1170
         End
         Begin VB.TextBox txtCostoCustodia 
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
            Height          =   255
            Left            =   1515
            TabIndex        =   6
            Top             =   465
            Width           =   1215
         End
         Begin VB.TextBox txtCostoTasacion 
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
            Left            =   1515
            TabIndex        =   5
            Top             =   165
            Width           =   1215
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Total a Pagar :"
            Height          =   255
            Index           =   1
            Left            =   5775
            TabIndex        =   21
            Top             =   1335
            Width           =   1170
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "ITF"
            Height          =   255
            Index           =   0
            Left            =   5760
            TabIndex        =   20
            Top             =   750
            Width           =   1170
         End
         Begin VB.Label LblTotalPagar 
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
            Left            =   5760
            TabIndex        =   19
            Top             =   1575
            Width           =   1395
         End
         Begin VB.Label LblITF 
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
            Left            =   5760
            TabIndex        =   18
            Top             =   990
            Width           =   1395
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Neto a Recibir :"
            Height          =   255
            Index           =   7
            Left            =   5760
            TabIndex        =   16
            Top             =   180
            Width           =   1170
         End
         Begin VB.Label lblNetoRecibir 
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
            Left            =   5760
            TabIndex        =   15
            Top             =   420
            Width           =   1395
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Impuesto :"
            Height          =   255
            Index           =   3
            Left            =   3075
            TabIndex        =   12
            Top             =   435
            Width           =   855
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Interes :"
            Height          =   255
            Index           =   2
            Left            =   3075
            TabIndex        =   11
            Top             =   165
            Width           =   780
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Costo Custodia :"
            Height          =   240
            Index           =   4
            Left            =   180
            TabIndex        =   10
            Top             =   450
            Width           =   1260
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Costo Tasación :"
            Height          =   255
            Index           =   5
            Left            =   165
            TabIndex        =   9
            Top             =   195
            Width           =   1245
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
         Height          =   3495
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   6165
      End
   End
End
Attribute VB_Name = "frmColPDesembolso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************
'* DESEMBOLSO DE CREDITO PIGNORATICIO
'Archivo:  frmColPDesembolso.frm
'LAYG   :  01/06/2001.
'Resumen:  Nos permite registrar el desembolso del Credito Pignoraticio

Option Explicit
Dim loItf1 As Double, loItf2 As Double
Dim vFecVenc As Date 'peac 20070820
Dim lnOpeCod As Long 'By capi 05032009 Acta 022-2009
Dim lsPersCod As String 'By capi 06032009

'Inicializa las variables del formulario
Private Sub Limpiar()
    Me.AXCodCta.NroCuenta = fgIniciaAxCuentaPignoraticio
    AXDesCon.Limpiar
    txtCostoTasacion.Text = ""
    txtCostoCustodia.Text = ""
    txtInteres.Text = ""
    txtImpuesto.Text = ""
    lblNetoRecibir.Caption = ""
End Sub

'Busca el contrato ingresado
Private Sub BuscaContrato(ByVal psNroContrato As String)
Dim lbOk As Boolean
Dim lrValida As ADODB.Recordset
Dim loValContrato As COMNColoCPig.NCOMColPValida
Dim loValMontoPrestamo As Double
Dim loValOtrosCostos As Double
Dim lsmensaje As String



'On Error GoTo ControlError
    gITF.fgITFParamAsume (Mid(psNroContrato, 4, 2)), Mid(psNroContrato, 6, 3)
    'Valida Contrato
    Set lrValida = New ADODB.Recordset
    Set loValContrato = New COMNColoCPig.NCOMColPValida
        Set lrValida = loValContrato.nValidaDesembolsoCredPignoraticio(psNroContrato, gdFecSis, 0, lsmensaje)
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

    Me.txtInteres = Format(lrValida!nInteres, "#0.00")
    Me.txtImpuesto = Format(lrValida!nImpuesto, "#0.00")
    Me.txtCostoTasacion = Format(lrValida!nTasacion, "#0.00")
    Me.txtCostoCustodia = Format(lrValida!nCustodia, "#0.00")
    'PEAC 20070813
    vFecVenc = Format(lrValida!dVenc, "dd/mm/yyyy")

    Set lrValida = Nothing
    
    loValMontoPrestamo = CCur(AXDesCon.SaldoCapital)
    loValOtrosCostos = CCur(Me.txtCostoTasacion) + CCur(txtCostoCustodia.Text) + CCur(txtInteres.Text) + CCur(txtImpuesto.Text)
        
    '*** PEAC 20071206 - en el neto a recibir no estara el interes adelantado ********************
    'Me.lblNetoRecibir.Caption = Format(CCur(AXDesCon.SaldoCapital) - (CCur(Me.txtCostoTasacion) + CCur(txtCostoCustodia.Text) + CCur(txtInteres.Text) + CCur(txtImpuesto.Text)), "#0.00")
    Me.lblNetoRecibir.Caption = Format(CCur(AXDesCon.SaldoCapital) - (CCur(Me.txtCostoTasacion) + CCur(txtCostoCustodia.Text) + CCur(txtImpuesto.Text)), "#0.00")
    '*********************************************************************************************
    
    ' **************  ITF ***************
    If gITF.gbITFAplica Then
        If Not gITF.gbITFAsumidocreditos Then
            loItf1 = Format(gITF.fgITFCalculaImpuesto(Val(Me.lblNetoRecibir.Caption)), "#0.00")
            'loItf2 = Format(gITF.fgITFCalculaImpuesto(loValOtrosCostos), "#0.00")
            'Me.LblITF = Format(gITF.fgITFCalculaImpuesto(CDbl(Me.lblNetoRecibir)), "#0.00")
            'Me.LblITF = Format(loItf1 + loItf2, "#0.00")
            Me.LblITF = Format(loItf1, "#0.00")
            Me.LblTotalPagar = Format(CDbl(Me.lblNetoRecibir) - CDbl(Me.LblITF), "#0.00")
        Else
            loItf1 = Format(gITF.fgITFCalculaImpuesto(loValMontoPrestamo), "#0.00")
            loItf2 = Format(gITF.fgITFCalculaImpuesto(loValOtrosCostos), "#0.00")
            'Me.LblITF = Format(gITF.fgITFCalculaImpuesto(CDbl(Me.lblNetoRecibir)), "#0.00")
            Me.LblITF = Format(loItf1 + loItf2, "#0.00")
            Me.LblTotalPagar = Me.lblNetoRecibir
        End If
    Else
        Me.LblITF = Format(0, "#0.00")
        Me.LblTotalPagar = Me.lblNetoRecibir
    End If
    ' **************  ITF ***************
    
    'By capi 05032009 Acta 022-2009
    
    cmdGrabar.Enabled = True
    cmdGrabar.SetFocus
    
'ALPA 20090420 Comentado para compilacion sin cambios dejado por CAPI
'    If lnOpeCod = gColPOpeDesembolsoEFE Then
'        cmdGrabar.Enabled = True
'        cmdGrabar.SetFocus
'    ElseIf lnOpeCod = gColPOpeDesembolsoAboCta Then
'        cmdCtasAhorros.Enabled = True
'    End If
'End by
    
            
    
        
    AXCodCta.Enabled = False
    'cmdBuscar.Enabled = False

Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox "Error: " & Err.Number & " " & Err.Description & vbCr & _
        "Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

Private Sub AXCodCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then BuscaContrato (AXCodCta.NroCuenta)
'ALPA 20090420 Comentado para compilacion sin cambios dejado por CAPI
'By capi 16032009
'    If (AXCodCta.NroCuenta) <> "" Then
'        cmdCtasAhorros.Visible = True
'        cmdCtasAhorros.SetFocus
'    End If
    '
End Sub

Private Sub cmdBuscar_Click()
Dim loPers As COMDPersona.UCOMPersona

'By capi 06032009
'Dim lsPersCod As String, lsPersNombre As String
Dim lsPersNombre As String

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
lsEstados = gColPEstRegis

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

Private Sub cmdCancelar_Click()
    Limpiar
    cmdGrabar.Enabled = False
    AXCodCta.Enabled = True
    AXCodCta.SetFocusCuenta
End Sub
'ALPA 20090420 Comentado para compilacion sin cambios dejado por CAPI
'Private Sub cmdCtasAhorros_Click()
'    frmCredDesembAbonoCta.DesembolsoPigAbonoCuenta Str(gColPOpeDesembolsoAboCta), lsPersCod, AXCodCta.NroCuenta
'End Sub

Private Sub cmdGrabar_Click()

'On Error GoTo ControlError
Dim loContFunct As COMNContabilidad.NCOMContFunciones
Dim loGrabarDesem As COMNColoCPig.NCOMColPContrato
Dim loColImp As COMNColoCPig.NCOMColPImpre
Dim lscadimp As String
Dim opt As Integer
Dim OptBt2 As Integer
Dim lsMovNro As String
Dim lsFechaHoraGrab As String
Dim lsFechaHoraPrend As String
Dim lsCuenta As String
Dim nFicSal As Integer

Dim lnSaldoCap As Currency, lnInteresComp As Currency, lnImpuesto As Currency
Dim lnCostoTasacion As Currency, lnCostoCustodia As Currency
Dim lnMontoEntregar As Currency
Set loColImp = New COMNColoCPig.NCOMColPImpre
lsCuenta = AXCodCta.NroCuenta
lnSaldoCap = Me.AXDesCon.SaldoCapital
lnInteresComp = CCur(Me.txtInteres.Text)
lnImpuesto = CCur(Me.txtImpuesto.Text)
lnCostoCustodia = CCur(Me.txtCostoCustodia.Text)
lnCostoTasacion = CCur(Me.txtCostoTasacion.Text)
lnMontoEntregar = CCur(Me.LblTotalPagar.Caption)

If MsgBox(" Grabar Desembolso de Contrato Pignoraticio ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
    cmdGrabar.Enabled = False
        
        'Genera el Mov Nro
        Set loContFunct = New COMNContabilidad.NCOMContFunciones
            lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        Set loContFunct = Nothing
        
        lsFechaHoraGrab = fgFechaHoraGrab(lsMovNro)
        lsFechaHoraPrend = fgFechaHoraPrend(lsMovNro)
        
        Set loGrabarDesem = New COMNColoCPig.NCOMColPContrato
            'Grabar Desembolso Pignoraticio
        Dim clsExo As New COMNCaptaServicios.NCOMCaptaServicios
        If Not clsExo.EsCuentaExoneradaLavadoDinero(lsCuenta) Then
          Dim sPersLavDinero As String
          Dim nMontoLavDinero As Double, nTC As Double
          Dim clsLav As New COMNCaptaGenerales.NCOMCaptaDefinicion, nMoneda As Integer, nMonto As Double
    
            nMonto = CDbl(LblTotalPagar.Caption)
            Set clsExo = Nothing
            sPersLavDinero = ""
            nMontoLavDinero = clsLav.GetCapParametro(gMonOpeLavDineroME)
            Set clsLav = Nothing
            nMoneda = gMonedaNacional
            If nMoneda = gMonedaNacional Then
                'Modificar
                Dim clsTC As COMDConstSistema.NCOMTipoCambio
                Set clsTC = New COMDConstSistema.NCOMTipoCambio
                nTC = clsTC.EmiteTipoCambio(gdFecSis, TCFijoDia)
                Set clsTC = Nothing
            Else
                nTC = 1
            End If
            If nMonto >= Round(nMontoLavDinero * nTC, 2) Then
                'By Capi 18022008 no aplica a desembolsos segun manifestado por riesgos
                'sPersLavDinero = IniciaLavDinero()
                 sPersLavDinero = ""
                'If sPersLavDinero = "" Then Exit Sub
            End If
         Else
            Set clsExo = Nothing
         End If
                        
            '*** PEAC 20071206 - no se graba el interes solo se muestra ***********************************
            'Call loGrabarDesem.nDesembolsoCredPignoraticio(lsCuenta, lnSaldoCap, lsFechaHoraGrab, _
                 lsMovNro, lnMontoEntregar, lnInteresComp, lnImpuesto, lnCostoTasacion, lnCostoCustodia, gITF.gbITFAplica, gITF.gbITFAsumidocreditos, loItf1, loItf2, CCur(lblNetoRecibir.Caption), False)
                 
            Call loGrabarDesem.nDesembolsoCredPignoraticio(lsCuenta, lnSaldoCap, lsFechaHoraGrab, _
                 lsMovNro, lnMontoEntregar, 0, lnImpuesto, lnCostoTasacion, lnCostoCustodia, gITF.gbITFAplica, gITF.gbITFAsumidocreditos, loItf1, loItf2, CCur(lblNetoRecibir.Caption), False)
                 
                 
        
        
            lscadimp = loColImp.nPrintReciboDesembolso(vFecVenc, lsCuenta, lnSaldoCap, lsFechaHoraPrend, _
                       lnMontoEntregar, lnInteresComp, gsNomAge, gsCodUser, CDbl(LblITF.Caption), gImpresora)
           
'-------------------- comentado por CMAC - CUSCO----------------------
            Do
             OptBt2 = MsgBox("Desea Imprimir la Boleta", vbInformation + vbYesNo, "Aviso")
             If vbYes = OptBt2 Then
             nFicSal = FreeFile
             Open sLpt For Output As nFicSal
                Print #nFicSal, Chr$(27) & Chr$(50);   'espaciamiento lineas 1/6 pulg.
                    Print #nFicSal, Chr$(27) & Chr$(67) & Chr$(22);  'Longitud de página a 22 líneas'
                    Print #nFicSal, Chr$(27) & Chr$(77);   'Tamaño 10 cpi
                    Print #nFicSal, Chr$(27) + Chr$(107) + Chr$(0);     'Tipo de Letra Sans Serif
                    Print #nFicSal, Chr$(27) + Chr$(72) ' desactiva negrita
                    Print #nFicSal, lscadimp & Chr$(12)
                    Print #nFicSal, ""
                    Close #nFicSal
             End If
            Loop Until OptBt2 = vbNo
'----------------------------------------------------------------------
        Set loGrabarDesem = Nothing
        Set loColImp = Nothing
        Limpiar
        Me.lblNetoRecibir = "0.00"
        Me.LblITF = "0.00"
        Me.LblTotalPagar = "0.00"

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

Private Function IniciaLavDinero() As String
Dim i As Long
Dim nRelacion As CaptacRelacPersona
Dim sPersCod As String, sNombre As String, sDocId As String, sDireccion As String
Dim nMonto As Double, nPersoneria As Integer
Dim sCuenta As String
'For i = 1 To grdCliente.Rows - 1
    'nRelacion = CLng(Trim(Right(grdCliente.TextMatrix(i, 3), 4)))
    nPersoneria = gPersonaNat
    If nPersoneria = gPersonaNat Then
        'If nRelacion = gCapRelPersApoderado Or nRelacion = gCapRelPersTitular Then
            sPersCod = AXDesCon.listaClientes.ListItems(1).Text
            sNombre = AXDesCon.listaClientes.ListItems(1).SubItems(1)
            sDireccion = AXDesCon.listaClientes.ListItems(1).SubItems(2)
            sDocId = AXDesCon.listaClientes.ListItems(1).SubItems(7)
         '   Exit For
       ' End If
    Else
        'If nRelacion = gCapRelPersTitular Then
            sPersCod = AXDesCon.listaClientes.ListItems(1).Text
            sNombre = AXDesCon.listaClientes.ListItems(1).SubItems(1)
            sDireccion = AXDesCon.listaClientes.ListItems(1).SubItems(2)
            sDocId = AXDesCon.listaClientes.ListItems(1).SubItems(9)
          '  Exit For
        'End If
    End If
'Next i
nMonto = CDbl(LblTotalPagar.Caption)
sCuenta = AXCodCta.NroCuenta
'If sPersCodCMAC <> "" Then
'    IniciaLavDinero = frmMovLavDinero.Inicia(sPersCod, sNombre, sDireccion, sDocId, False, True, nmonto, sCuenta, sOperacion, , sTipoCuenta)
'Else
    IniciaLavDinero = frmMovLavDinero.Inicia(sPersCod, sNombre, sDireccion, sDocId, False, True, nMonto, sCuenta, gColPOpeDesembolsoEFE, , gMonedaNacional)
'End If
End Function

'Finaliza el formulario actual
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
            SendKeys "{Enter}"
        End If
    ElseIf KeyCode = 13 And Trim(AXCodCta.EnabledCta) And AXCodCta.Age <> "" And Trim(AXCodCta.Cuenta) = "" Then
                AXCodCta.SetFocusCuenta
                 Exit Sub
        
    End If
End Sub

'Inicializa el formulario actual
Private Sub Form_Load()
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
    Limpiar
End Sub


'ALPA 20090420 Comentado para compilacion sin cambios dejado por CAPI
'By capi 05032009 Acta 022-2009
'Public Sub Inicio(ByVal pnOpeCod As Long)
'
'    Select Case pnOpeCod
'        Case gColPOpeDesembolsoEFE
'            frmColPDesembolso.Caption = "Credito Pignoraticio - Desembolso en Efectivo"
'        Case gColPOpeDesembolsoAboCta
'            cmdCtasAhorros.Visible = True
'            frmColPDesembolso.Caption = "Credito Pignoraticio - Desembolso con Abono en Cuenta"
'    End Select
'    Limpiar
'    Me.Show 1
'End Sub

