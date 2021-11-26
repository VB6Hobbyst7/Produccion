VERSION 5.00
Begin VB.Form FrmColPDuplicadoContratoSinCosto 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Crédito Pignoraticio - Duplicado de Contrato"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8100
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   8100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Enabled         =   0   'False
      Height          =   345
      Left            =   4365
      TabIndex        =   19
      Top             =   5580
      Width           =   1005
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6765
      TabIndex        =   18
      Top             =   5580
      Width           =   975
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   5565
      TabIndex        =   17
      Top             =   5580
      Width           =   975
   End
   Begin VB.Frame fraContenedor 
      Height          =   5415
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7830
      Begin VB.Frame fraContenedor 
         Enabled         =   0   'False
         Height          =   915
         Index           =   2
         Left            =   240
         TabIndex        =   2
         Top             =   4320
         Width           =   7395
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
            TabIndex        =   12
            Top             =   480
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
            TabIndex        =   10
            Top             =   180
            Width           =   1035
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
            TabIndex        =   9
            Top             =   180
            Width           =   1035
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Cost. Tasac."
            Height          =   255
            Index           =   2
            Left            =   180
            TabIndex        =   8
            Top             =   210
            Width           =   1005
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Cost. Custod."
            Height          =   255
            Index           =   18
            Left            =   2430
            TabIndex        =   7
            Top             =   180
            Width           =   990
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Interes"
            Height          =   255
            Index           =   16
            Left            =   180
            TabIndex        =   6
            Top             =   540
            Width           =   795
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Impuesto "
            Height          =   255
            Index           =   17
            Left            =   2430
            TabIndex        =   5
            Top             =   540
            Width           =   855
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
            TabIndex        =   4
            Top             =   240
            Width           =   840
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Nro.Duplicado :"
            Height          =   255
            Index           =   19
            Left            =   4680
            TabIndex        =   3
            Top             =   240
            Width           =   1275
         End
      End
      Begin VB.CommandButton cmdBuscar 
         Height          =   345
         Left            =   7200
         Picture         =   "FrmColPDuplicadocontratoSinCosto.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Buscar ..."
         Top             =   270
         Width           =   420
      End
      Begin SICMACT.ActXCodCta AXCodCta 
         Height          =   375
         Left            =   180
         TabIndex        =   13
         Top             =   270
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   661
         Texto           =   "Crédito"
         EnabledCta      =   -1  'True
         EnabledAge      =   -1  'True
      End
      Begin SICMACT.ActXColPDesCon AXDesCon 
         Height          =   3495
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   6165
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Costo Duplicado :  S/."
         Height          =   255
         Index           =   11
         Left            =   3840
         TabIndex        =   16
         Top             =   5280
         Visible         =   0   'False
         Width           =   1680
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
         Left            =   5520
         TabIndex        =   15
         Top             =   5280
         Visible         =   0   'False
         Width           =   1290
      End
   End
   Begin VB.Label lblMensaje 
      Caption         =   "<F12> Número Contrato Antiguo"
      Height          =   285
      Left            =   120
      TabIndex        =   20
      Top             =   5760
      Visible         =   0   'False
      Width           =   2655
   End
End
Attribute VB_Name = "FrmColPDuplicadocontratosinCosto"
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


'Permite inicializarlas variables del formulario
Private Sub Limpiar()
    Me.AXCodCta.NroCuenta = fgIniciaAxCuentaPignoraticio
    AXDesCon.Limpiar
    lblCostoDuplicado.Caption = Format(pCostoDuplicado, "#0.00")
    lblNroDuplic.Caption = ""
    Me.lblCostoTasacion = "0.00"
    Me.lblCostoCustodia = "0.00"
    Me.lblNroDuplic = ""
    Me.lblInteres = "0.00"
    Me.lblImpuesto = "0.00"
    Me.lblCostoDuplicado = "0.00"
End Sub

'Permite buscar el contrato ingresado
Private Sub BuscaContrato(ByVal psNroContrato As String)

Dim lbOk As Boolean
Dim lrValida As ADODB.Recordset
Dim loValContrato As nColPValida
'On Error GoTo ControlError

    'Valida Contrato
    Set lrValida = New ADODB.Recordset
    Set loValContrato = New nColPValida
        Set lrValida = loValContrato.nValidaDuplicadoContratoCredPignoraticio(psNroContrato)
    Set loValContrato = Nothing
    
    If lrValida Is Nothing Then ' Hubo un Error
        Limpiar
        Set lrValida = Nothing
        Exit Sub
    End If
    'Muestra Datos
    lbOk = fgMuestraCredPig_AXDesCon(psNroContrato, Me.AXDesCon, False)

    fnTasaInteresAdelantado = lrValida!nTasaInteres
    Me.lblInteres = Format(lrValida!nInteres, "#0.00")
    Me.lblImpuesto = Format(lrValida!nImpuesto, "#0.00")
    Me.lblCostoTasacion = Format(lrValida!nTasacion, "#0.00")
    Me.lblCostoCustodia = Format(lrValida!nCustodia, "#0.00")
    Me.lblNroDuplic = Format(lrValida!nNroDuplic + 1, "#0")
    Set lrValida = Nothing
    
    Me.lblCostoDuplicado = Format(pCostoDuplicado, "#0.00")
    AXCodCta.Enabled = False
    cmdImprimir.Enabled = True
    cmdImprimir.SetFocus

Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "

End Sub

Private Sub AXCodCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call BuscaContrato(AXCodCta.NroCuenta)
End Sub

Private Sub cmdBuscar_Click()
Dim loPers As UPersona
Dim lsPersCod As String, lsPersNombre As String
Dim lsEstados As String
Dim loPersContrato As dColPContrato
Dim lrContratos As ADODB.Recordset
Dim loCuentas As UProdPersona

On Error GoTo ControlError

Set loPers = New UPersona
    Set loPers = frmBuscaPersona.Inicio
    If loPers Is Nothing Then Exit Sub
    lsPersCod = loPers.sPersCod
    lsPersNombre = loPers.sPersNombre
Set loPers = Nothing

' Selecciona Estados
lsEstados = gColPEstRegis & "," & gColPEstDesem & "," & gColPEstVenci & "," & gColPEstPRema & "," & _
            gColPEstRenov & "," & gColPEstDifer & "," & gColPEstCance

If Trim(lsPersCod) <> "" Then
    Set loPersContrato = New dColPContrato
        Set lrContratos = loPersContrato.dObtieneCredPigDePersona(lsPersCod, lsEstados, Mid(gsCodAge, 4, 2))
    Set loPersContrato = Nothing
End If

Set loCuentas = New UProdPersona
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
    cmdImprimir.Enabled = False
    AXCodCta.Enabled = True
    AXCodCta.SetFocusCuenta
End Sub

Private Sub cmdImprimir_Click()

'On Error GoTo ControlError
Dim loContFunct As NContFunciones
Dim loGrabarDup As NColPContrato
Dim loImprime As NColPImpre
Dim loPrevio As Previo.clsPrevio

Dim lsMovNro As String
Dim lsFechaHoraGrab As String

Dim lnNumDuplicado As Integer
Dim lnMontoTransaccion As Currency
Dim lsCadImprimir As String
Dim lsNombreCliente As String
'Dim lsLote As String
Dim lrPersonas As ADODB.Recordset

lnNumDuplicado = Val(Me.lblNroDuplic.Caption)
'lnMontoTransaccion = CCur(Me.lblCostoDuplicado.Caption)


lsNombreCliente = AXDesCon.listaClientes.ListItems(1).ListSubItems.Item(1)
Set lrPersonas = fgGetCodigoPersonaListaRsNew(Me.AXDesCon.listaClientes)
'lsLote = fgEliminaEnters(Me.AXDesCon.DescLote) & vbCr

'If MsgBox(" Grabar Duplicado de Contrato Pignoraticio ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
'    cmdImprimir.Enabled = False
        
    'Genera el Mov Nro
    Set loContFunct = New NContFunciones
        lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Set loContFunct = Nothing
    
    lsFechaHoraGrab = fgFechaHoraGrab(lsMovNro)
    Set loGrabarDup = New NColPContrato
        'Grabar Duplicado de Contrato Pignoraticio
        Call loGrabarDup.nDuplicadoContratoCredPignoraticio(AXCodCta.NroCuenta, lnNumDuplicado, lsFechaHoraGrab, _
              lsMovNro, 0, False)
    Set loGrabarDup = Nothing

   ' *** Impresion


    If MsgBox("Imprimir Duplicado de Contrato Pignoraticio ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
        Set loImprime = New NColPImpre
            lsCadImprimir = loImprime.nPrintContratoPignoraticio(AXCodCta.NroCuenta, True, , , , , , , , , , _
                                    , , , , , , , , , , gsCodUser, lnNumDuplicado)
        Set loImprime = Nothing
        Set loPrevio = New Previo.clsPrevio
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

    Limpiar
    
    AXCodCta.Enabled = True
    AXCodCta.SetFocus
        
'Else
'    MsgBox " Grabación cancelada ", vbInformation, " Aviso "
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
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
    CargaParametros
    Limpiar
End Sub

'Procedimiento de impresión del duplicado del contrato
Private Sub ImprimirContrato(vNroImpresiones As Byte)
'    Dim I As Byte
'    Dim vDescLote As String
'    Dim vEspacio As Integer
'    Dim vEspaMedio As Integer
'    Dim x As Integer, vLongi As Integer
'    Dim vNombre As String
'    vEspacio = 7
'    vEspaMedio = 0
'    vDescLote = AXDesCon.DescLote
'    If AXDesCon.Oro14 > 0 Then vDescLote = vDescLote & vbCr & ImpreFormat(AXDesCon.Oro14, 8) & " grs. 14K."
'    If AXDesCon.Oro16 > 0 Then vDescLote = vDescLote & vbCr & ImpreFormat(AXDesCon.Oro16, 8) & " grs. 16K."
'    If AXDesCon.Oro18 > 0 Then vDescLote = vDescLote & vbCr & ImpreFormat(AXDesCon.Oro18, 8) & " grs. 18K."
'    If AXDesCon.Oro21 > 0 Then vDescLote = vDescLote & vbCr & ImpreFormat(AXDesCon.Oro21, 8) & " grs. 21K."
'    vDescLote = vDescLote & vbCr
'
'    For I = 1 To vNroImpresiones
'        ImpreBegChe False, 66
'        Print #ArcSal, " ": Print #ArcSal, " ": Print #ArcSal, " ": Print #ArcSal, " "
'        Print #ArcSal, " ": Print #ArcSal, " ": Print #ArcSal, " ": Print #ArcSal, " "
'
'        Print #ArcSal, Tab(22 + vEspaMedio); lstCliente.ListItems.Item(1); 'txtCodCliente;
'        Print #ArcSal, Tab(56 + vEspacio); Format(txtFechaEmpeno.Text, "dd/mm/yyyy");
'        Print #ArcSal, Chr$(27) & Chr$(69);    'Activa Negrita
'        Print #ArcSal, Tab(73 + vEspacio); Mid(vNroContrato, 1, 2) & "-" & Mid(vNroContrato, 3, 3) _
'        & "-" & Mid(vNroContrato, 6, 1) & "-" & Mid(vNroContrato, 7, 5) & "-" & Mid(vNroContrato, 12, 1) & "-"; InicialBoveda(gsCodAge)
'        Print #ArcSal, Chr$(27) & Chr$(70);    'Desactiva Negrita
'        Print #ArcSal, Tab(57 + vEspacio); Format(txtFechaVencimiento.Text, "dd/mm/yyyy")
'        Print #ArcSal, " "
'        Print #ArcSal, " "
'        vLongi = Switch(lstCliente.ListItems.Count = 1, 35, lstCliente.ListItems.Count = 2, 25, _
'            lstCliente.ListItems.Count = 3, 18, lstCliente.ListItems.Count = 4, 12, lstCliente.ListItems.Count > 4, 35)
'        If vLongi = 35 Then
'            vNombre = Left(lstCliente.ListItems(1).ListSubItems.Item(1), vLongi)
'        Else
'            For x = 1 To lstCliente.ListItems.Count
'                If x > 1 Then vNombre = vNombre & " / "
'                vNombre = vNombre & Left(lstCliente.ListItems(x).ListSubItems.Item(1), vLongi)
'            Next x
'        End If
'
'        Print #ArcSal, Tab(30 + vEspaMedio); ImpreCarEsp(vNombre)  'txtNomCliente.Text
'        Print #ArcSal, Tab(30 + vEspaMedio); lstCliente.ListItems(1).ListSubItems.Item(7) & "  " & lstCliente.ListItems(1).ListSubItems.Item(9) 'Left(txtNatural, 15) & "  " & Left(txtTributario, 15)
'        Print #ArcSal, Tab(30 + vEspaMedio); ImpreCarEsp(lstCliente.ListItems(1).ListSubItems.Item(2)); 'vDireccion;
'        Print #ArcSal, Tab(85 + vEspacio); AXDesCon.Plazo
'        Print #ArcSal, Tab(30 + vEspaMedio); ImpreCarEsp(lstCliente.ListItems(1).ListSubItems.Item(4)) 'vZona
'        Print #ArcSal, Tab(30 + vEspaMedio); ImpreCarEsp(lstCliente.ListItems(1).ListSubItems.Item(5)); Tab(81 + vEspacio); "DUPLICADO"  'vCiudad
'        Print #ArcSal, Tab(30 + vEspaMedio); lstCliente.ListItems(1).ListSubItems.Item(3) 'vTelefono
'        Print #ArcSal, " "
'        Print #ArcSal, Tab(83 + vEspacio - Len(Format(AXDesCon.Piezas, "#0"))); AXDesCon.Piezas
'        Print #ArcSal, Tab(86 + vEspacio - IIf(AXDesCon.OroBruto = 0, 4, Len(Format(AXDesCon.OroBruto, "#0.00")))); Format(AXDesCon.OroBruto, "#0.00")
'        Print #ArcSal, Tab(86 + vEspacio - IIf(AXDesCon.OroNeto = 0, 4, Len(Format(AXDesCon.OroNeto, "#0.00")))); Format(AXDesCon.OroNeto, "#0.00")
'        Print #ArcSal, Tab(86 + vEspacio - IIf(AXDesCon.ValTasa = 0, 4, Len(Format(AXDesCon.ValTasa, "#0.00")))); Format(AXDesCon.ValTasa, "#0.00")
'        Print #ArcSal, " "
'        Print #ArcSal, Tab(86 + vEspacio - IIf(AXDesCon.Prestamo = 0, 4, Len(Format(AXDesCon.Prestamo, "#0.00")))); Format(AXDesCon.Prestamo, "#0.00")
'        Print #ArcSal, Tab(86 + vEspacio - IIf(Txtinteres = 0, 4, Len(Format(Txtinteres, "#0.00")))); Format(Txtinteres.Text, "#0.00")
'        Print #ArcSal, " "
'        Print #ArcSal, Tab(13 + vEspaMedio); QuiebreTexto(vDescLote, 1);
'        Print #ArcSal, Tab(86 + vEspacio - IIf(txtCostoTasacion = 0, 4, Len(Format(txtCostoTasacion, "#0.00")))); Format(txtCostoTasacion, "#0.00")
'        Print #ArcSal, Tab(13 + vEspaMedio); QuiebreTexto(vDescLote, 2);
'        Print #ArcSal, Tab(86 + vEspacio - IIf(txtCostoCustodia = 0, 4, Len(Format(txtCostoCustodia, "#0.00")))); Format(txtCostoCustodia, "#0.00")
'        Print #ArcSal, Tab(13 + vEspaMedio); QuiebreTexto(vDescLote, 3);
'        Print #ArcSal, Tab(86 + vEspacio - IIf(txtImpuesto = 0, 4, Len(Format(txtImpuesto, "#0.00")))); Format(txtImpuesto.Text, "#0.00")
'        Print #ArcSal, Tab(13 + vEspaMedio); QuiebreTexto(vDescLote, 4) & " "
'        Print #ArcSal, Tab(13 + vEspaMedio); QuiebreTexto(vDescLote, 5) & " "
'        Print #ArcSal, Tab(13 + vEspaMedio); QuiebreTexto(vDescLote, 6) & " "
'        Print #ArcSal, Tab(13 + vEspaMedio); QuiebreTexto(vDescLote, 7);
'        Print #ArcSal, Tab(78 + vEspacio); Format(vNetoARecibir, "#0.00")
'        Print #ArcSal, Tab(13 + vEspaMedio); QuiebreTexto(vDescLote, 8) & " "
'        Print #ArcSal, Tab(13 + vEspaMedio); QuiebreTexto(vDescLote, 9) & " "
'        Print #ArcSal, Tab(13 + vEspaMedio); QuiebreTexto(vDescLote, 10) & " "
'        Print #ArcSal, Tab(13 + vEspaMedio); QuiebreTexto(vDescLote, 11) & " "
'        Print #ArcSal, Tab(13 + vEspaMedio); QuiebreTexto(vDescLote, 12) & " "
'        Print #ArcSal, Tab(13 + vEspaMedio); QuiebreTexto(vDescLote, 13) & " "
'        Print #ArcSal, Tab(13 + vEspaMedio); QuiebreTexto(vDescLote, 14) & " "
'        Print #ArcSal, Tab(13 + vEspaMedio); QuiebreTexto(vDescLote, 15) & " "
'        Print #ArcSal, Tab(13 + vEspaMedio); QuiebreTexto(vDescLote, 16) & " "
'        Print #ArcSal, Tab(13 + vEspaMedio); QuiebreTexto(vDescLote, 17) & " "
'        Print #ArcSal, Tab(13 + vEspaMedio); QuiebreTexto(vDescLote, 18) & " "
'        Print #ArcSal, Tab(13 + vEspaMedio); QuiebreTexto(vDescLote, 19) & " "
'        Print #ArcSal, Tab(13 + vEspaMedio); QuiebreTexto(vDescLote, 20) & " "
'        Print #ArcSal, " "
'        Print #ArcSal, Tab(12 + vEspaMedio); "TASA EFECTIVA ANUAL : " & Format(vTasaInteres, "#0.00%")
'        Print #ArcSal, Tab(12 + vEspaMedio); "Fecha Duplicado : " & gdFecSis & " " & Format(Time, "hh:mm:ss")
'        Print #ArcSal, Tab(12 + vEspaMedio); ImpreCarEsp("N° Duplicado    : ") & lblNroDuplic.Caption
'        Print #ArcSal, " ": Print #ArcSal, " ": Print #ArcSal, " ": Print #ArcSal, " "
'        Print #ArcSal, " ": Print #ArcSal, " ": Print #ArcSal, " "
'        Print #ArcSal, Chr$(27) & Chr$(107) & Chr$(2); 'Tipo letra : 0,1,2 - Roman,SansS,Courier
'        Print #ArcSal, Chr$(27) & Chr$(80);            'Tamaño  +...-: 80, 77, 103
'        Print #ArcSal, Chr$(27) & Chr$(69);    'Activa Negrita
'        Print #ArcSal, Tab(40); gsCodUser;
'        Print #ArcSal, Tab(61); Mid(vNroContrato, 1, 2) & "-" & Mid(vNroContrato, 3, 3) _
'        & "-" & Mid(vNroContrato, 6, 1) & "-" & Mid(vNroContrato, 7, 5) & "-" & Mid(vNroContrato, 12, 1) & "-"; InicialBoveda(gsCodAge)
'        Print #ArcSal, " "
'        Print #ArcSal, Tab(66); AXDesCon.Plazo
'        ImpreEnd
'    Next I
End Sub


Private Sub CargaParametros()
Dim loParam As DColPCalculos
Set loParam = New DColPCalculos
    pCostoDuplicado = loParam.dObtieneColocParametro(gConsColPCostoDuplicadoContrato)
Set loParam = Nothing
End Sub

