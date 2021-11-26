VERSION 5.00
Begin VB.Form frmCapCambioEstadoCTS 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CTS - Cambio de Estado de Cliente"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9135
   Icon            =   "frmCapCambioEstadoCTS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   9135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   8040
      TabIndex        =   13
      Top             =   4680
      Width           =   975
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   6840
      TabIndex        =   12
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton cmdCeseLab 
      Caption         =   "Cese Laboral"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2400
      TabIndex        =   11
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton cmdRevCTSActivo 
      Caption         =   "Reversión a CTS Activo"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   4680
      Width           =   2175
   End
   Begin VB.Frame fraCliente 
      Caption         =   " Cliente "
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
      Height          =   1965
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8895
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
         TabIndex        =   2
         Top             =   1440
         Width           =   5655
      End
      Begin VB.CommandButton cmdSelecccionar 
         Caption         =   "Selecccionar"
         Height          =   375
         Left            =   7440
         TabIndex        =   1
         Top             =   1400
         Width           =   1215
      End
      Begin SICMACT.TxtBuscar TxtBCodPers 
         Height          =   315
         Left            =   1560
         TabIndex        =   3
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
         TabIndex        =   9
         Top             =   720
         Width           =   7095
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente:"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   770
         Width           =   735
      End
      Begin VB.Label lblPersCod 
         AutoSize        =   -1  'True
         Caption         =   "Cod Cliente:"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   390
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "DOI:"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1170
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
         TabIndex        =   5
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label7 
         Caption         =   "Empresa:"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1500
         Width           =   735
      End
   End
   Begin SICMACT.FlexEdit feCambEstado 
      Height          =   2415
      Left            =   120
      TabIndex        =   14
      Top             =   2160
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   4260
      Cols0           =   10
      HighLight       =   1
      AllowUserResizing=   1
      RowSizingMode   =   1
      EncabezadosNombres=   "#-Nº Cuenta-SubProducto-Moneda-Saldo-Tasa-Fec. Últ. Dep.-Fech. Migra-nTpoPrograma-bCeseLaboral"
      EncabezadosAnchos=   "300-1700-1700-800-1000-900-1080-1000-0-0"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X"
      TextStyleFixed  =   4
      ListaControles  =   "0-0-0-0-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-C-L-C-R-C-C-C-C-C"
      FormatosEdit    =   "0-0-0-0-2-0-0-0-0-0"
      CantEntero      =   12
      CantDecimales   =   4
      TextArray0      =   "#"
      SelectionMode   =   1
      lbUltimaInstancia=   -1  'True
      lbFormatoCol    =   -1  'True
      lbBuscaDuplicadoText=   -1  'True
      ColWidth0       =   300
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
End
Attribute VB_Name = "frmCapCambioEstadoCTS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'** Nombre : frmCapCambioEstadoCTS
'** Descripción : Formulario para cambiar de estado una cuenta CTS según TI-ERS013-2014
'** Creación : JUEZ, 20140305 12:20:00 AM
'*****************************************************************************************************

Option Explicit

Dim oDCapGen As COMDCaptaGenerales.DCOMCaptaGenerales
Dim oDCapMov As COMDCaptaGenerales.DCOMCaptaMovimiento
Dim R As ADODB.Recordset

Private Sub cmdCancelar_Click()
    HabilitaControles False
    TxtBCodPers.Text = ""
    lblTitular.Caption = ""
    lblDOI.Caption = ""
    cboInstitucion.Clear
    Call LimpiaFlex(feCambEstado)
    TxtBCodPers.SetFocus
End Sub

Private Sub cmdCeseLab_Click()
    If MsgBox("¿Desea realizar el cese laboral de las cuentas listadas?", vbYesNo, "Aviso") = vbNo Then Exit Sub
    Dim oGen As COMDConstSistema.DCOMGeneral
    Dim oNCapMov As COMNCaptaGenerales.NCOMCaptaMovimiento
    Dim oDCapGen As COMDCaptaGenerales.DCOMCaptaGenerales
    Dim oNCapDef As COMNCaptaGenerales.NCOMCaptaDefinicion
    Dim rsCta As ADODB.Recordset
    Dim bExito As Boolean
    Dim i As Integer
    Dim nTipoCamb As Double, nExcedente As Double, nSaldoRetiro As Double, nPorcDisp As Double, nIntSaldo As Double
    Dim nDiasTranscurridos As Integer, nTasa As Double, dUltMov As Date
    
    Set oDCapMov = New COMDCaptaGenerales.DCOMCaptaMovimiento
    oDCapMov.CTSCeseLaboral TxtBCodPers.Text, Trim(Right(cboInstitucion.Text, 13)), gsCodUser, gsCodAge, bExito

    If bExito Then
        Set oGen = New COMDConstSistema.DCOMGeneral
        Set oNCapMov = New COMNCaptaGenerales.NCOMCaptaMovimiento
        Set oDCapGen = New COMDCaptaGenerales.DCOMCaptaGenerales
        Set oNCapDef = New COMNCaptaGenerales.NCOMCaptaDefinicion
        
        For i = 1 To feCambEstado.rows - 1
            Set rsCta = oDCapGen.GetDatosCuentaCTS(feCambEstado.TextMatrix(i, 1))
            nSaldoRetiro = rsCta("nSaldRetiro")
            nTasa = rsCta("nTasaInteres")
            dUltMov = rsCta("dUltCierre")
            
            nDiasTranscurridos = DateDiff("d", dUltMov, gdFecSis) - 1
            If nDiasTranscurridos < 0 Then
                nDiasTranscurridos = 0
            End If
            nIntSaldo = oNCapMov.GetInteres(nSaldoRetiro, nTasa, nDiasTranscurridos, TpoCalcIntSimple)
            
            nPorcDisp = oNCapDef.GetCapParametro(gPorRetCTS)
            nExcedente = 0
            
            Set rsCta = oDCapMov.ObtenerCapSaldosCuentasCTS(feCambEstado.TextMatrix(i, 1), oGen.GetTipCambio(gdFecSis, TCFijoMes))
            nExcedente = rsCta!TotalSaldos - rsCta!TotalSueldos
            If nExcedente > 0 Then
                nSaldoRetiro = nExcedente * nPorcDisp / 100
            Else
                nSaldoRetiro = 0
            End If
            oDCapMov.ActualizaSaldoRetiroCTS feCambEstado.TextMatrix(i, 1), nSaldoRetiro, nIntSaldo
        Next i
        
        MsgBox "Se realizó el cese laboral de las cuentas con éxito", vbInformation, "Aviso"
        cmdCancelar_Click
    Else
        MsgBox "Hubo un error en la operación", vbInformation, "Aviso"
    End If
End Sub

Private Sub cmdRevCTSActivo_Click()
    If MsgBox("¿Desea migrar a CTS Activo a las cuentas listadas?", vbYesNo, "Aviso") = vbNo Then Exit Sub
    Dim oGen As COMDConstSistema.DCOMGeneral
    Dim oNCapMov As COMNCaptaGenerales.NCOMCaptaMovimiento
    Dim oDCapGen As COMDCaptaGenerales.DCOMCaptaGenerales
    Dim oNCapDef As COMNCaptaGenerales.NCOMCaptaDefinicion
    Dim rsCta As ADODB.Recordset
    Dim bExito As Boolean
    Dim i As Integer
    Dim nTipoCamb As Double, nExcedente As Double, nSaldoRetiro As Double, nPorcDisp As Double, nIntSaldo As Double
    Dim nDiasTranscurridos As Integer, nTasa As Double, dUltMov As Date
    
    Set oDCapMov = New COMDCaptaGenerales.DCOMCaptaMovimiento
    oDCapMov.RevertirCTSActivo TxtBCodPers.Text, Trim(Right(cboInstitucion.Text, 13)), gsCodUser, gsCodAge, bExito
    
    If bExito Then
        Set oGen = New COMDConstSistema.DCOMGeneral
        Set oNCapMov = New COMNCaptaGenerales.NCOMCaptaMovimiento
        Set oDCapGen = New COMDCaptaGenerales.DCOMCaptaGenerales
        Set oNCapDef = New COMNCaptaGenerales.NCOMCaptaDefinicion
        Dim nSaldoDisp As Double 'APRI20200330 POR COVID-19
        For i = 1 To feCambEstado.rows - 1
            Set rsCta = oDCapGen.GetDatosCuentaCTS(feCambEstado.TextMatrix(i, 1))
            nSaldoRetiro = rsCta("nSaldRetiro")
            nTasa = rsCta("nTasaInteres")
            dUltMov = rsCta("dUltCierre")
            nSaldoDisp = rsCta("nSaldoDisp") * IIf(Mid(feCambEstado.TextMatrix(i, 1), 9, 1) = "1", 1, oGen.GetTipCambio(gdFecSis, TCFijoMes)) 'APRI20200330 POR COVID-19
            nDiasTranscurridos = DateDiff("d", dUltMov, gdFecSis) - 1
            If nDiasTranscurridos < 0 Then
                nDiasTranscurridos = 0
            End If
            nIntSaldo = oNCapMov.GetInteres(nSaldoRetiro, nTasa, nDiasTranscurridos, TpoCalcIntSimple)
            
            nPorcDisp = oNCapDef.GetCapParametro(gPorRetCTS)
            nExcedente = 0
            
            Set rsCta = oDCapMov.ObtenerCapSaldosCuentasCTS(feCambEstado.TextMatrix(i, 1), oGen.GetTipCambio(gdFecSis, TCFijoMes))
            nExcedente = rsCta!TotalSaldos - rsCta!TotalSueldos
            If nExcedente > 0 Then
                nSaldoRetiro = nExcedente * nPorcDisp / 100
            Else
                nSaldoRetiro = 0
            End If
            'APRI20200330 CULPA DEL COVID-19
            Dim nDU01 As Double
            If gdFecSis <= "2020-06-10" Then
                nDU01 = (nDU01 + IIf(nSaldoDisp < 2400, nSaldoDisp, 2400)) / IIf(Mid(feCambEstado.TextMatrix(i, 1), 9, 1) = "1", 1, oGen.GetTipCambio(gdFecSis, TCFijoMes))
            End If
            'END APRI
            'oDCapMov.ActualizaSaldoRetiroCTS feCambEstado.TextMatrix(i, 1), nSaldoRetiro, nIntSaldo
            oDCapMov.ActualizaSaldoRetiroCTS feCambEstado.TextMatrix(i, 1), nSaldoRetiro, nIntSaldo, nDU01 'APRI20200415 POR COVID-19
        Next i
        
        MsgBox "Se revirtieron las cuentas con éxito", vbInformation, "Aviso"
        cmdCancelar_Click
    Else
        MsgBox "Hubo un error en la reversión", vbInformation, "Aviso"
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdSelecccionar_Click()
Dim lnFila As Integer

    Set oDCapGen = New COMDCaptaGenerales.DCOMCaptaGenerales
    Set R = oDCapGen.RecuperaDatosCuentasCTSInstitucion(TxtBCodPers.Text, Trim(Right(cboInstitucion.Text, 13)))
    Do While Not R.EOF
        feCambEstado.AdicionaFila
        lnFila = feCambEstado.row
        feCambEstado.TextMatrix(lnFila, 1) = R!cCtaCod
        feCambEstado.TextMatrix(lnFila, 2) = R!cTpoPrograma
        feCambEstado.TextMatrix(lnFila, 3) = R!cMoneda
        feCambEstado.TextMatrix(lnFila, 4) = Format(R!nSaldo, "#,##0.00")
        feCambEstado.TextMatrix(lnFila, 5) = R!nTEA
        feCambEstado.TextMatrix(lnFila, 6) = R!dFecUltDep
        feCambEstado.TextMatrix(lnFila, 7) = R!dFecMigra
        feCambEstado.TextMatrix(lnFila, 8) = R!nTpoPrograma
        If R!nTpoPrograma = 2 Then
            cmdRevCTSActivo.Enabled = True
        End If
        feCambEstado.TextMatrix(lnFila, 9) = R!bCeseLaboral
        If R!bCeseLaboral = 0 Then
            cmdCeseLab.Enabled = True
        End If
        R.MoveNext
    Loop
    R.Close
    Set oDCapGen = Nothing
    cmdSelecccionar.Enabled = False
End Sub

Private Sub Form_Load()
    HabilitaControles False
End Sub

Private Sub HabilitaControles(ByVal pbHabilita As Boolean)
    TxtBCodPers.Enabled = Not pbHabilita
    cboInstitucion.Enabled = pbHabilita
    cmdSelecccionar.Enabled = pbHabilita
    cmdRevCTSActivo.Enabled = pbHabilita
    cmdCeseLab.Enabled = pbHabilita
End Sub

Private Sub TxtBCodPers_EmiteDatos()
Dim oCred As COMDCredito.DCOMCredito
Dim oPers As comdpersona.DCOMPersonas

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
        TxtBCodPers.Enabled = False
        cmdSelecccionar.Enabled = True
        cboInstitucion.Enabled = True
        cboInstitucion.ListIndex = -1
        cboInstitucion.SetFocus
    Else
        cmdCancelar_Click
        MsgBox "Cliente no posee depósitos CTS vigentes", vbInformation, "Aviso"
    End If
End Sub
