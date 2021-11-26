VERSION 5.00
Begin VB.Form frmCapDatosSueldosCTS 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registro de sueldos de clientes CTS"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8145
   Icon            =   "frmCapDatosSueldosCTS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   8145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6720
      TabIndex        =   19
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   " Histórico de Registro "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   2895
      Left            =   120
      TabIndex        =   13
      Top             =   3000
      Width           =   7875
      Begin SICMACT.FlexEdit FEHistorico 
         Height          =   2535
         Left            =   150
         TabIndex        =   14
         Top             =   240
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   4471
         Cols0           =   6
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "Nro-Fecha-Moneda-Total 4 Rem. Bruto-Agencia-Usuario"
         EncabezadosAnchos=   "500-1200-1200-1500-2000-800"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-C-R-L-C"
         FormatosEdit    =   "0-0-0-0-0-0"
         TextArray0      =   "Nro"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   495
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.Frame fraCuenta 
      Caption         =   " Selección de Cliente "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   2805
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   7875
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   6480
         TabIndex        =   7
         Top             =   2280
         Width           =   1215
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   5160
         TabIndex        =   6
         Top             =   2280
         Width           =   1215
      End
      Begin VB.ComboBox cboInstitucion 
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
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1440
         Width           =   6135
      End
      Begin VB.ComboBox cboMoneda 
         Height          =   315
         ItemData        =   "frmCapDatosSueldosCTS.frx":030A
         Left            =   1560
         List            =   "frmCapDatosSueldosCTS.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2200
         Width           =   1335
      End
      Begin SICMACT.EditMoney EdtSueldos 
         Height          =   315
         Left            =   1560
         TabIndex        =   4
         Top             =   1830
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
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
         Enabled         =   -1  'True
      End
      Begin SICMACT.TxtBuscar TxtBCodPers 
         Height          =   315
         Left            =   1560
         TabIndex        =   0
         Top             =   360
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   556
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TipoBusqueda    =   3
         sTitulo         =   ""
      End
      Begin VB.Label Label7 
         Caption         =   "Empresa:"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   1500
         Width           =   735
      End
      Begin VB.Label lblDOI 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1560
         TabIndex        =   2
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "DOI:"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1170
         Width           =   735
      End
      Begin VB.Label lblPersCod 
         AutoSize        =   -1  'True
         Caption         =   "Cod Cliente:"
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   390
         Width           =   855
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Moneda Sueldos:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label lblUltRemunBrutas 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   11
         Top             =   1875
         Width           =   3855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Total Sueldos:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1880
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente:"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   770
         Width           =   735
      End
      Begin VB.Label lblTitular 
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
         Left            =   1560
         TabIndex        =   1
         Top             =   720
         Width           =   6135
      End
   End
   Begin SICMACT.ActXCodCta txtCuenta 
      Height          =   435
      Left            =   240
      TabIndex        =   15
      Top             =   120
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   767
      Texto           =   "Cuenta N°"
      EnabledCMAC     =   -1  'True
      EnabledCta      =   -1  'True
      EnabledProd     =   -1  'True
      EnabledAge      =   -1  'True
   End
End
Attribute VB_Name = "frmCapDatosSueldosCTS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lsCtaCod As String
'JUEZ 20141112 Nuevos parámetros ******
Dim nParPorcTangRet As Double
Dim nParUltRemunBrutas As Integer
Dim nParSaldoMin As Double
'END JUEZ *****************************
Dim sMovNroAut As String 'APRI20170601 ERS033-2017
Dim sUltimaRemuneracion As Currency 'APRI20170601 ERS033-2017

'JUEZ 20130724 *************************************************
Private Sub cboInstitucion_Click()
    Call obtieneHistoricoCTS(TxtBCodPers.Text, Trim(Right(cboInstitucion.Text, 13)), Trim(Right(cboMoneda.Text, 2)))
    EdtSueldos.SetFocus
End Sub

Private Sub cboInstitucion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EdtSueldos.SetFocus
    End If
End Sub

Private Sub cboMoneda_Click()
    Call obtieneHistoricoCTS(TxtBCodPers.Text, Trim(Right(cboInstitucion.Text, 13)), Trim(Right(cboMoneda.Text, 2)))
    If lsCtaCod <> "" Then cmdAceptar.SetFocus
End Sub
'END JUEZ ******************************************************

Private Sub cboMoneda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdAceptar.SetFocus
    End If
End Sub

Private Sub cmdCancelar_Click()
   Call Limpiar
   TxtBCodPers.SetFocus 'JUEZ 20141112
   'Unload Me 'JUEZ 20141112
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

'***Agregado por ELRO el 20121010, según OYP-RFC101-2012
Private Sub EdtSueldos_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        'cmdAceptar.SetFocus
        cboMoneda.SetFocus 'JUEZ 20130724
    End If
End Sub
 '***Fin Agregado por ELRO el 20121010*******************

Private Sub Form_Load()
'JUEZ 20130724 *************************************************
'    txtCuenta.Prod = 234
'    txtCuenta.CMAC = 109
'    txtCuenta.EnabledCMAC = False
'    txtCuenta.EnabledProd = False
'    cboMoneda.Text = "SOLES"
    Dim rs As ADODB.Recordset
    Dim oCons As New COMDConstantes.DCOMConstantes
    Set rs = oCons.RecuperaConstantes(gMoneda)
    Call Llenar_Combo_con_Recordset(rs, cboMoneda)
    'lsCtaCod = ""
    Limpiar 'JUEZ 20141112
End Sub

Private Sub TxtBCodPers_EmiteDatos()
Dim oCred As COMDCredito.DCOMCredito
Dim oPers As comdpersona.DCOMPersonas
Dim R As ADODB.Recordset
    If Trim(TxtBCodPers.Text) = "" Then
        Exit Sub
    End If
    Set oPers = New comdpersona.DCOMPersonas
    Set R = oPers.RecuperaInstitucionesCTSPersona(TxtBCodPers.Text)
    If Not R.EOF Then
        Set oCred = New COMDCredito.DCOMCredito
        Set R = oCred.RecuperaDatosComision(TxtBCodPers.Text, 2)
        Set oCred = Nothing
        lblTitular.Caption = R!cPersNombre
        lblDOI.Caption = R!cPersIDnro
        Set oPers = New comdpersona.DCOMPersonas
        Set R = oPers.RecuperaInstitucionesCTSPersona(TxtBCodPers.Text)
        cboInstitucion.Clear
        While Not R.EOF
            cboInstitucion.AddItem R!cDescInst & Space(100) & R!cCodInst
            R.MoveNext
        Wend
        cboInstitucion.ListIndex = -1
        cboInstitucion.SetFocus
        TxtBCodPers.Enabled = False
    Else
        Limpiar
        MsgBox "Cliente no posee depósitos CTS", vbInformation, "Aviso"
    End If
End Sub
'END JUEZ ******************************************************

'Private Sub txtCuenta_KeyPress(KeyAscii As Integer) 'Comentado por JUEZ 20130823
'    If KeyAscii = 13 Then
'        Dim sCta As String
'        sCta = txtCuenta.NroCuenta
'        ObtieneDatosCuenta sCta
'        '***Agregado por ELRO el 20121010, según OYP-RFC101-2012
'        obtieneHistoricoCTS sCta
'        '***Fin Agregado por ELRO el 20121010*******************
'    End If
'End Sub
Private Sub ObtieneDatosCuenta(ByVal sCuenta As String)
    Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales
    Dim rsPers As ADODB.Recordset
    Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
    Set rsPers = New ADODB.Recordset
    Set rsPers = clsMant.GetPersonaCuenta(sCuenta)
    

    If Not (rsPers.EOF And rsPers.BOF) Then
        '***Modificado por ELRO el 20121123, según SATI INC1211220007
        If rsPers!nPrdEstado <> gCapEstCancelada And rsPers!nPrdEstado <> gCapEstAnulada Then
        lblTitular.Caption = UCase(PstaNombre(rsPers("Nombre")))
        Me.EdtSueldos.SetFocus
        '***Modificado por ELRO el 20121031, según OYP-RFC101-2012
        Me.txtCuenta.Enabled = False
        '***Fin Modificado por ELRO el 20121031*******************
        Else
            MsgBox "El Nro. de cuenta esta cancelada o anulada.", vbInformation, "¡Aviso!"
            Me.txtCuenta.SetFocus
            Exit Sub
        End If
        '***Modificado por ELRO el 20121123**************************
    Else
        MsgBox "No se ha encontrado información de la cuenta ingresada"
        Me.txtCuenta.SetFocus
    End If
End Sub
Private Sub Limpiar()
    Me.EdtSueldos.Text = "0"
    Me.lblTitular.Caption = ""
    'JUEZ 20130724 ****************
    'Me.txtCuenta.Age = ""
    'Me.txtCuenta.Cuenta = ""
    'Me.txtCuenta.Enabled = True
    'Me.txtCuenta.SetFocus
    TxtBCodPers.Text = ""
    cboInstitucion.Clear
    lblDOI.Caption = ""
    TxtBCodPers.Enabled = True
    'TxtBCodPers.SetFocus
    'END JUEZ *********************
    '***Agregado por ELRO el 20121010, según OYP-RFC101-2012
    Call LimpiaFlex(FEHistorico)
    '***Fin Agregado por ELRO el 20121010*******************
    'JUEZ 20141112 *******************
    lsCtaCod = ""
    cboMoneda.ListIndex = 0
    nParPorcTangRet = 0
    nParUltRemunBrutas = 0
    nParSaldoMin = 0
    lblUltRemunBrutas.Caption = ""
    'END JUEZ ************************
    sMovNroAut = "" 'APRI20170601 ERS033-2017
End Sub

Private Sub cmdAceptar_Click()
    Dim ClsMov As New NCOMCaptaMovimiento
    Dim clsMovN As New COMNContabilidad.NCOMContFunciones
    Dim clsCap As COMDCaptaGenerales.DCOMCaptaMovimiento
    Dim clsDef As NCOMCaptaDefinicion
    Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales
    Dim oGen As New COMDConstSistema.DCOMGeneral
    Dim rsCta As ADODB.Recordset
    Dim sMovNro As String
    Dim nPorcDisp As Double
    Dim nExcedente As Double
    Dim nIntSaldo As Double
    Dim dUltMov As Date
    Dim nTasa As Double
    Dim nDiasTranscurridos As Integer
    Dim nSaldoRetiro As Double '***Agregado por ELRO el 20121106, según OYP-RFC101-2012
    Dim lnSueldoMinimo As Currency '***Agregado por ELRO el 20121106, según OYP-RFC101-2012
    Dim lnUltimasRemuneracionesBruta As Integer '***Agregado por ELRO el 20121106, según OYP-RFC101-2012
    Dim lnSumaSueldosMinimoMN As Currency '***Agregado por ELRO el 20121106, según OYP-RFC101-2012
    Dim lnSumaSueldosMinimoME As Currency '***Agregado por ELRO el 20121106, según OYP-RFC101-2012
    

      
    If Me.lblTitular.Caption = "" Then
'        MsgBox "No se ha ingresado un N° de cuenta válido", vbExclamation + vbOKOnly, "Advertencia"
'        Me.txtCuenta.SetFocus
        MsgBox "No se ha ingresado el código de persona", vbInformation, "Aviso" 'JUEZ 20130724
        Exit Sub
    End If
    'JUEZ 20130724 **********************************************************************
    If Trim(cboMoneda.Text) = "" Then
        MsgBox "No se eligió la moneda", vbInformation, "Aviso"
        Me.cboMoneda.SetFocus
        Exit Sub
    End If
    If Trim(cboInstitucion.Text) = "" Then
        MsgBox "No se eligió la Institución", vbInformation, "Aviso"
        Me.cboInstitucion.SetFocus
        Exit Sub
    End If
    'END JUEZ ***************************************************************************
    If Me.EdtSueldos.value = 0 Then
        MsgBox "No se ingresó el monto total de los sueldos ", vbInformation, "Aviso"
        Me.EdtSueldos.SetFocus
        Exit Sub
    End If
    
    
    '***Agregado por ELRO el 20121106, según OYP-RFC101-2012
    Set clsDef = New NCOMCaptaDefinicion
    'JUEZ 20141112 Nuevos parámetros **************************
    'lnSueldoMinimo = CCur(clsDef.GetCapParametro(2128))
    'lnUltimasRemuneracionesBruta = CCur(clsDef.GetCapParametro(2129))
    lnSueldoMinimo = nParSaldoMin
    lnUltimasRemuneracionesBruta = nParUltRemunBrutas
    'END JUEZ ************************************************
    lnSumaSueldosMinimoMN = lnSueldoMinimo * lnUltimasRemuneracionesBruta
    lnSumaSueldosMinimoME = lnSumaSueldosMinimoMN / oGen.GetTipCambio(gdFecSis, TCFijoMes)
    Set clsDef = Nothing
    
    If Trim(cboMoneda.Text) = "" Then
        MsgBox "Debe elegir la moneda.", vbExclamation + vbOKOnly, "Advertencia"
        cboMoneda.SetFocus
        Exit Sub
    End If
    
    If Trim(Right(cboMoneda.Text, 2)) = "1" Then
        If lnSumaSueldosMinimoMN > CCur(EdtSueldos.value) Then
            MsgBox "El monto ingresado no es válido. Verifíquelo o inténtelo de nuevo o consulte con el Jefe de Ahorros.", vbExclamation + vbOKOnly, "Advertencia"
            Me.EdtSueldos.SetFocus
            Exit Sub
        End If
    Else
        If lnSumaSueldosMinimoME > CCur(EdtSueldos.value) Then
            MsgBox "El monto ingresado no es válido. Verifíquelo o inténtelo de nuevo o consulte con el Jefe de Ahorros.", vbExclamation + vbOKOnly, "Advertencia"
            Me.EdtSueldos.SetFocus
            Exit Sub
        End If
    End If
    '***Fin Agregado por ELRO*******************************
    
    '**********************APRI20170601 ERS033-2017***************************
    If Trim(Right(cboInstitucion.Text, 13)) = "1090100012521" Then
        If Me.EdtSueldos.value < sUltimaRemuneracion Then
            If MsgBox("El monto ingresado es menor al último registro de sus 4 sueldos,¿Desea Continuar? ", vbQuestion + vbYesNo, "Aviso") = vbNo Then
                Exit Sub
            End If
            Dim bRechazo As Boolean
            bRechazo = False
            If VerificarAutorizacion(bRechazo) = False Then
                sMovNroAut = ""
                Exit Sub
            End If
        End If
    End If
    '*********************END APRI20170601************************
    
    If MsgBox("Está seguro de grabar la información?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        Set ClsMov = New NCOMCaptaMovimiento
        Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
        Set clsMovN = New COMNContabilidad.NCOMContFunciones
        Set clsDef = New NCOMCaptaDefinicion
        Set clsCap = New COMDCaptaGenerales.DCOMCaptaMovimiento
        Dim nSaldoDisp As Double 'APRI20200330 POR COVID-19
        Dim nDU01 As Double 'APRI20200415 POR COVID-19
'        Set rsCta = clsMant.GetDatosCuentaCTS(Me.txtCuenta.NroCuenta)
        Set rsCta = clsMant.GetDatosCuentaCTS(lsCtaCod) 'JUEZ 20130724
        nSaldoRetiro = rsCta("nSaldRetiro")
        nTasa = rsCta("nTasaInteres")
        dUltMov = rsCta("dUltCierre")
        nSaldoDisp = rsCta("nSaldoDisp") * IIf(Mid(lsCtaCod, 9, 1) = "1", 1, oGen.GetTipCambio(gdFecSis, TCFijoMes)) 'APRI20200330 POR COVID-19
        nDU01 = rsCta("nDU01") 'APRI20200415 POR COVID-19
        nDiasTranscurridos = DateDiff("d", dUltMov, gdFecSis) - 1
        If nDiasTranscurridos < 0 Then
            nDiasTranscurridos = 0
        End If
        nIntSaldo = ClsMov.GetInteres(nSaldoRetiro, nTasa, nDiasTranscurridos, TpoCalcIntSimple)
        
        'nPorcDisp = clsDef.GetCapParametro(gPorRetCTS)
        nPorcDisp = nParPorcTangRet 'JUEZ 20141112 Nuevos Parámetros
        sMovNro = clsMovN.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
        nExcedente = 0
        
        'Call clsCap.AgregaDatosSueldosClientesCTS(sMovNro, Me.txtCuenta.NroCuenta, IIf(Trim(cboMoneda.Text) = "SOLES", 1, 2), Me.EdtSueldos.value)
        Call clsCap.AgregaDatosSueldosClientesCTS(sMovNro, lsCtaCod, CInt(Trim(Right(cboMoneda.Text, 2))), EdtSueldos.value) 'JUEZ 20130724

        'Set rsCta = clsCap.ObtenerCapSaldosCuentasCTS(Me.txtCuenta.NroCuenta, oGen.GetTipCambio(gdFecSis, TCFijoMes))
        Set rsCta = clsCap.ObtenerCapSaldosCuentasCTS(lsCtaCod, oGen.GetTipCambio(gdFecSis, TCFijoMes)) 'JUEZ 20130724
        nExcedente = rsCta!TotalSaldos - rsCta!TotalSueldos

        If nExcedente > 0 Then
            nSaldoRetiro = nExcedente * nPorcDisp / 100
        Else
            nSaldoRetiro = 0
        End If
        'APRI20200330 CULPA DEL COVID-19
        If gdFecSis <= "2020-04-12" Then
            nSaldoRetiro = nSaldoRetiro + IIf(nSaldoDisp < 2400, nSaldoDisp, 2400)
        End If
        'END APRI
        'clsCap.ActualizaSaldoRetiroCTS Me.txtCuenta.NroCuenta, nSaldoRetiro, nIntSaldo
        'clsCap.ActualizaSaldoRetiroCTS lsCtaCod, nSaldoRetiro, nIntSaldo 'JUEZ 20130724
        clsCap.ActualizaSaldoRetiroCTS lsCtaCod, nSaldoRetiro, nIntSaldo, nDU01 'APRI20200415 POR COVID-19
        
        MsgBox "Se ha realizado el registro de forma exitosa!", vbOKOnly + vbInformation, "Mensaje"
        Call Limpiar
        TxtBCodPers.SetFocus 'JUEZ 20141112
    End If
    Set clsDef = Nothing
    Set ClsMov = Nothing
    Set clsMovN = Nothing
    Set clsCap = Nothing
    Set clsMant = Nothing
End Sub
'***Agregado por ELRO el 20121010, según OYP-RFC101-2012
'Private Sub obtieneHistoricoCTS(ByVal psCuenta As String)
Private Sub obtieneHistoricoCTS(ByVal psPersCod As String, ByVal psInstitucion As String, ByVal pnMoneda As Integer) 'JUEZ 20130724
    Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales
    Dim clsDef As COMNCaptaGenerales.NCOMCaptaDefinicion 'JUEZ 20141112
    Dim rsPar As ADODB.Recordset 'JUEZ 20141112
    Dim rsHist As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim i As Integer
    
    If psPersCod = "" Or psInstitucion = "" Or CStr(pnMoneda) = "" Then Exit Sub
    Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
    'JUEZ 20130724 ****************************************************************************
    'Set rsHist = New ADODB.Recordset
    Set rs = clsMant.ObtenerCuentaCTS(psPersCod, psInstitucion, pnMoneda)
    If (rs.EOF And rs.BOF) Then Exit Sub
    lsCtaCod = rs!cCtaCod
    'JUEZ 20141112 Nuevos parámetros ****************
    Set clsDef = New COMNCaptaGenerales.NCOMCaptaDefinicion
    Set rsPar = clsDef.GetCapParametroNew(Mid(lsCtaCod, 6, 3), rs!nTpoPrograma)
    nParPorcTangRet = rsPar!nPorcTangRet
    nParSaldoMin = rsPar!nSueldoMin
    nParUltRemunBrutas = rsPar!nUltRemunBrutas
    lblUltRemunBrutas.Caption = "(Total de las últimas " & nParUltRemunBrutas & " remuneraciones brutas)"
    'END JUEZ ***************************************
    
    Set rsHist = clsMant.obtenerHistorialCaptacSueldosCTS(lsCtaCod, sUltimaRemuneracion)
    'END JUEZ *********************************************************************************
    
    
    Call LimpiaFlex(FEHistorico)
    i = 1
    FEHistorico.lbEditarFlex = True
        
    If Not (rsHist.EOF And rsHist.BOF) Then
        Do While Not rsHist.EOF
            FEHistorico.AdicionaFila
            FEHistorico.TextMatrix(i, 1) = rsHist!cFecha
            FEHistorico.TextMatrix(i, 2) = rsHist!cMoneda
            FEHistorico.TextMatrix(i, 3) = Format(rsHist!nSueldoTotal, gcFormView)
            FEHistorico.TextMatrix(i, 4) = rsHist!cAgeDescripcion
            FEHistorico.TextMatrix(i, 5) = rsHist!cUser
            i = i + 1
            rsHist.MoveNext
        Loop
    End If
    
    FEHistorico.lbEditarFlex = False
End Sub
'***Fin Agregado por ELRO el 20121010*******************
'******************APRI20170601 ERS033-2017*****************************
Private Function VerificarAutorizacion(ByRef pbRechazado As Boolean) As Boolean

Dim oCapAut As COMDCaptaGenerales.COMDCaptAutorizacion
Dim oCapAutN  As COMNCaptaGenerales.NCOMCaptAutorizacion
Dim rs As New ADODB.Recordset

Dim lsmensaje As String
Dim nMonto As Double
Dim cMoneda As String
'Dim lbRechazado As Boolean

nMonto = Me.EdtSueldos.value
cMoneda = Mid(lsCtaCod, 9, 1)
   
Set oCapAutN = New COMNCaptaGenerales.NCOMCaptAutorizacion
If sMovNroAut = "" Then 'Si es nueva, registra nueva solicitud
    
    oCapAutN.NuevaSolicitudOtrasOperaciones TxtBCodPers.Text, "4", gdFecSis, nMonto, cMoneda, "Autorización de Registro de 04 sueldos CTS", gsCodUser, gOpeAutorizacion04SueldoCTS, gsCodAge, sMovNroAut
        
    Do While VerificarAutorizacion = False
        If Not oCapAutN.VerificarAutorizacionOtrasOperaciones("4", nMonto, sMovNroAut, lsmensaje, pbRechazado) Then
            If lsmensaje = "Esta Operación Aun no esta Autorizada" Then
                If MsgBox("Para proceder con la operacion debe solicitar VºBº del Coordinador de Operaciones..." & vbNewLine & _
                          "Desea continuar esperando la Autorización?", vbYesNo) = vbNo Then
                    Exit Do
                Else
                    VerificarAutorizacion = False
                End If
            End If
            If lsmensaje = "Esta Operación fue Rechazada" Then
                MsgBox lsmensaje, vbInformation, "Aviso"
                Exit Do
            End If
        Else
            MsgBox lsmensaje, vbInformation, "Aviso"
            VerificarAutorizacion = True
        End If
    Loop
    
End If
Set oCapAutN = Nothing
End Function
'************************END APRI20170601****************************************

