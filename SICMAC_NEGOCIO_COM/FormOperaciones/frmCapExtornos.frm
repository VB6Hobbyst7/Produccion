VERSION 5.00
Begin VB.Form frmCapExtornos 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10545
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   10545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmMotExtorno 
      Caption         =   "Motivos del Extorno"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2700
      Left            =   3960
      TabIndex        =   23
      Top             =   1680
      Visible         =   0   'False
      Width           =   2845
      Begin VB.ComboBox cmbMotivos 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         ItemData        =   "frmCapExtornos.frx":0000
         Left            =   240
         List            =   "frmCapExtornos.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox txtDetExtorno 
         BackColor       =   &H00C0FFFF&
         Height          =   750
         Left            =   240
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   25
         Top             =   1200
         Width           =   2415
      End
      Begin VB.CommandButton cmdExtContinuar 
         Caption         =   "&Continuar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   960
         TabIndex        =   24
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Detalles del Extorno"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label lblExtCmb 
         BackStyle       =   0  'Transparent
         Caption         =   "Motivo"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame fraNroCredito 
      Height          =   540
      Left            =   2400
      TabIndex        =   17
      Top             =   600
      Visible         =   0   'False
      Width           =   3675
      Begin VB.TextBox txtCredito 
         Height          =   285
         Left            =   1155
         MaxLength       =   18
         TabIndex        =   19
         Top             =   180
         Width           =   2340
      End
      Begin VB.Label Label1 
         Caption         =   "Nro Credito"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   90
         TabIndex        =   18
         Top             =   232
         Width           =   1035
      End
   End
   Begin VB.Frame fraGlosa 
      Caption         =   "Glosa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1575
      Left            =   7560
      TabIndex        =   16
      Top             =   60
      Width           =   2895
      Begin VB.TextBox txtGlosa 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   1095
         Left            =   60
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   300
         Width           =   2655
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   60
      TabIndex        =   9
      Top             =   6060
      Width           =   1035
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   9420
      TabIndex        =   8
      Top             =   6060
      Width           =   1035
   End
   Begin VB.CommandButton cmdExtornar 
      Caption         =   "&Extornar"
      Height          =   375
      Left            =   8340
      TabIndex        =   7
      Top             =   6060
      Width           =   1035
   End
   Begin VB.Frame fraMovimientos 
      Caption         =   "Movimientos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   4215
      Left            =   60
      TabIndex        =   11
      Top             =   1740
      Width           =   10395
      Begin SICMACT.FlexEdit grdMov 
         Height          =   3855
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   10155
         _ExtentX        =   17912
         _ExtentY        =   6800
         Cols0           =   22
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosAnchos=   "250-2200-2300-1600-1200-1200-2500-0-0-0-0-1200-0-0-1200-1200-1200-1200-1200-0-0-0"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-L-C-R-C-L-C-C-C-L-R-C-C-R-C-R-R-R-C-C-C"
         FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0-0-2-0-0-2-2-0-0-0-0-0-0"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   255
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.Frame fraBuscar 
      Caption         =   "Datos Búsqueda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1575
      Left            =   60
      TabIndex        =   10
      Top             =   60
      Width           =   7395
      Begin VB.Frame Frame2 
         Height          =   1215
         Left            =   2220
         TabIndex        =   13
         Top             =   240
         Width           =   4995
         Begin VB.TextBox txtCodMov 
            Height          =   325
            Left            =   1440
            TabIndex        =   21
            Top             =   445
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.CommandButton cmdBuscar 
            Caption         =   "&Buscar"
            Height          =   375
            Left            =   3840
            TabIndex        =   4
            Top             =   420
            Width           =   1035
         End
         Begin VB.TextBox txtMovNro 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   1800
            TabIndex        =   2
            Top             =   420
            Width           =   1455
         End
         Begin SICMACT.ActXCodCta txtCuenta 
            Height          =   375
            Left            =   120
            TabIndex        =   3
            Top             =   420
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   661
            Texto           =   "Cuenta N°"
            EnabledCMAC     =   -1  'True
            EnabledCta      =   -1  'True
            EnabledProd     =   -1  'True
            EnabledAge      =   -1  'True
         End
         Begin VB.Label lblCodMov 
            Caption         =   "Cod. Mov.:"
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
            Left            =   240
            TabIndex        =   22
            Top             =   510
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label lblMov 
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
            Height          =   345
            Left            =   840
            TabIndex        =   15
            Top             =   435
            Width           =   975
         End
         Begin VB.Label lblNroMov 
            Caption         =   "# Mov :"
            Height          =   195
            Left            =   180
            TabIndex        =   14
            Top             =   450
            Width           =   675
         End
      End
      Begin VB.Frame Frame1 
         Height          =   1215
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1995
         Begin VB.OptionButton optTipoBus 
            Caption         =   "Codigo Movimiento"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   20
            Top             =   840
            Visible         =   0   'False
            Width           =   1755
         End
         Begin VB.OptionButton optTipoBus 
            Caption         =   "Número de &Cuenta"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   1
            Top             =   540
            Width           =   1755
         End
         Begin VB.OptionButton optTipoBus 
            Caption         =   "&Número Movimiento"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   0
            Top             =   240
            Width           =   1755
         End
      End
   End
End
Attribute VB_Name = "frmCapExtornos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nOperacion As COMDConstantes.CaptacOperacion
Dim nProducto As COMDConstantes.Producto
Dim lsCodExtOpc As String

Public Sub Inicia(ByVal nOpe As CaptacOperacion, ByVal sOperacion As String, ByVal nProd As Producto, Optional ByVal psCodExtOpc As String = "")
nOperacion = nOpe
Me.Caption = "Captaciones - Extornos - " & sOperacion

'*** PEAC 20081002
'*** SE AGREGO EL PARAMETRO psCodExtOpc PARA VALIDAR EL VISTO ELECTRONICO

lsCodExtOpc = psCodExtOpc

optTipoBus(0).value = True
cmdExtornar.Enabled = False
cmdCancelar.Enabled = False
nProducto = nProd
txtCuenta.Prod = nProducto
If nOperacion = "200104" Or nOperacion = "200904" Then
    optTipoBus(2).Visible = True
End If
'***Agregado por ELRO el 20121120, según OYP-RFC101-2012
If nOperacion = gCTSDepLotEfec Or nOperacion = gCTSDepLotChq Or nOperacion = gCTSDepLotTransf Then
    optTipoBus(1).Visible = False
End If
'***Fin Agregado por ELRO el 20121120*******************
'******* CTI3 ******************************************
If lsCodExtOpc = "270501" Or lsCodExtOpc = "270502" Or lsCodExtOpc = "270503" Or lsCodExtOpc = "270504" Then
    txtGlosa.Enabled = True
    txtGlosa.BackColor = &H80000005
Else
    txtGlosa.Enabled = False
    txtGlosa.BackColor = &H80000004
End If

Me.Show 1
End Sub

Private Sub AgregaMovGrid(ByVal rsMov As Recordset)
Dim nFila As Long
Do While Not rsMov.EOF
    grdMov.AdicionaFila
    nFila = grdMov.Rows - 1
    grdMov.TextMatrix(nFila, 1) = rsMov("cMovNro")
    grdMov.TextMatrix(nFila, 2) = rsMov("cOpeDesc")
    grdMov.TextMatrix(nFila, 3) = rsMov("cCtaCod")
    grdMov.TextMatrix(nFila, 4) = Format$(rsMov("nMonto"), "#,##0.00")
    grdMov.TextMatrix(nFila, 5) = rsMov("cDocNro")
    grdMov.TextMatrix(nFila, 6) = Trim(rsMov("cMovDesc"))
    grdMov.TextMatrix(nFila, 7) = Trim(rsMov("cOpeCod"))
    grdMov.TextMatrix(nFila, 8) = Trim(rsMov("nDocTpo"))
    grdMov.TextMatrix(nFila, 9) = Trim(rsMov("nMovNro"))
    grdMov.TextMatrix(nFila, 10) = Trim(rsMov("cPersCod"))
    If rsMov.Fields.count >= 11 Then
        grdMov.TextMatrix(nFila, 11) = Format$(rsMov("ITFCargo"), "#,##0.00")
        grdMov.TextMatrix(nFila, 12) = Trim(rsMov("ITFOperacion"))
        grdMov.TextMatrix(nFila, 13) = Trim(rsMov("ITFConcepto"))
          If nOperacion <> gServGiroApertEfec Then
           If nOperacion <> gServGiroCancEfec Then
            grdMov.TextMatrix(nFila, 14) = Format$(rsMov("RetOtraAge"), "#,##0.00")
            grdMov.TextMatrix(nFila, 15) = Format$(rsMov("RetxMaxOpe"), "#,##0.00")
            grdMov.TextMatrix(nFila, 16) = Format$(rsMov("DepOtraAge"), "#,##0.00")
            grdMov.TextMatrix(nFila, 17) = Format$(rsMov("RetSinTarj"), "#,##0.00")
          End If
         End If
        'RIRO20131212 ERS137
        If nOperacion = gAhoRetTransf Or nOperacion = gAhoCancTransfAbCtaBco Or _
            nOperacion = gPFRetIntAboCtaBanco Or nOperacion = gPFCancTransf Or _
            nOperacion = gCTSRetTransf Or nOperacion = gCTSCancTransfBco Then
                
                grdMov.TextMatrix(nFila, 18) = Format$(rsMov("comisionTrans"), "#,##0.00")
                grdMov.TextMatrix(nFila, 19) = rsMov("ComisionDebito")
                
        End If
        'END RIRO
        'CTI4 ERS0112020
        If nOperacion = gAhoRetEmiChq Then
            grdMov.TextMatrix(nFila, 20) = Format$(rsMov("ComisionEmiteCheque"), "#,##0.00")
            grdMov.TextMatrix(nFila, 21) = rsMov("ComisionEmiteChequeOpe")
        End If
        'CTI4 End
    End If
    rsMov.MoveNext
Loop
End Sub

Private Sub cmdBuscar_Click()
    Dim clsCap As COMNCaptaGenerales.NCOMCaptaMovimiento
        Dim oSeg As New COMNCaptaGenerales.NCOMSeguros 'RECO20160209 ERS073-2015
    Dim rsMov As ADODB.Recordset
    
    Dim sDatoBus As String
    Set clsCap = New COMNCaptaGenerales.NCOMCaptaMovimiento

    'If Left(nOperacion, 4) <> "2605" And nOperacion <> "107001" Then
    If Left(nOperacion, 4) <> "2605" And nOperacion <> "107001" And nOperacion <> "300151" And nOperacion <> "200380" Then 'RECO20160209 ERS073-2015
        If optTipoBus(0).value Then
            '***Modificado por ELRO el 20110919, según Acta 261-2011/TI-D
            'sDatoBus = lblMov & "%" & Trim(txtMovNro)  'comentado por ELRO el 20110919
            sDatoBus = lblMov & Trim(txtMovNro)
            '***Fin Modificado por ELRO
            If nOperacion = gServGiroApertEfec Or nOperacion = gServGiroCancEfec Then
                Set rsMov = clsCap.GetMovExtornoGiro(sDatoBus, gdFecSis, nOperacion, gsCodAge, 0)
            Else
                Set rsMov = clsCap.GetMovExtorno(sDatoBus, gdFecSis, nOperacion, gsCodAge, 0)
            End If
        ElseIf optTipoBus(1).value Then
            sDatoBus = txtCuenta.NroCuenta
            If nOperacion = gServGiroApertEfec Or nOperacion = gServGiroCancEfec Then
                Set rsMov = clsCap.GetMovExtornoGiro(sDatoBus, gdFecSis, nOperacion, gsCodAge, 1)
            Else
                Set rsMov = clsCap.GetMovExtorno(sDatoBus, gdFecSis, nOperacion, gsCodAge, 1)
            End If
        ElseIf optTipoBus(2).value Then
            sDatoBus = txtCodMov.Text
            If nOperacion = gServGiroApertEfec Or nOperacion = gServGiroCancEfec Then
                Set rsMov = clsCap.GetMovExtornoGiro(sDatoBus, gdFecSis, nOperacion, gsCodAge, 2)
            Else
                Set rsMov = clsCap.GetMovExtorno(sDatoBus, gdFecSis, nOperacion, gsCodAge, 2)
            End If
        End If
    End If

    If Left(nOperacion, 4) = "2605" Or nOperacion = "107001" Then
    
        If optTipoBus(0).value Then
            'sDatoBus = lblMov & "%" & Trim(txtMovNro)
            sDatoBus = Trim(txtMovNro)
            Set rsMov = clsCap.GetMovExtornoCMAC(sDatoBus, gdFecSis, nOperacion, 0)
        ElseIf optTipoBus(1).value And nOperacion <> "107001" Then
            sDatoBus = txtCuenta.NroCuenta
            Set rsMov = clsCap.GetMovExtornoCMAC(sDatoBus, gdFecSis, nOperacion, 1)
        ElseIf optTipoBus(1).value And nOperacion = "107001" Then
            sDatoBus = Trim(txtCredito.Text)
            Set rsMov = clsCap.GetMovExtornoCMAC(sDatoBus, gdFecSis, nOperacion, 1)
        End If
    End If
    'RECO20160209 ERS073-2015***************************************
    'SEGURO SEPELIO
    If nOperacion = "220315" Or nOperacion = "200380" Then
        If optTipoBus(0).value Then
            sDatoBus = lblMov & "%" & Trim(txtMovNro)
            Set rsMov = oSeg.SepelioDevolverMovimientoExtornar(sDatoBus, gdFecSis, nOperacion, gsCodAge, 0)
        ElseIf optTipoBus(1).value And nOperacion = "220315" Then
            sDatoBus = lblMov & "%" & Trim(txtMovNro)
            Set rsMov = oSeg.SepelioDevolverMovimientoExtornar(sDatoBus, gdFecSis, nOperacion, gsCodAge, 0)
        ElseIf optTipoBus(1).value And nOperacion = "200380" Then
            sDatoBus = txtCuenta.NroCuenta
            Set rsMov = oSeg.SepelioDevolverMovimientoExtornar(sDatoBus, gdFecSis, nOperacion, gsCodAge, 0)
        End If
    End If
    'RECO FIN ******************************************************
        
    fraBuscar.Enabled = False
    cmdExtornar.Enabled = True
    cmdCancelar.Enabled = True
    If Not (rsMov.EOF And rsMov.BOF) Then
        AgregaMovGrid rsMov
    Else
        MsgBox "No se registraron movimientos con el criterio de búsqueda", vbInformation, "Aviso"
        Call ExtDatosOculta
    End If '
    Set clsCap = Nothing
    rsMov.Close
    Set rsMov = Nothing
    End Sub

Private Sub cmdCancelar_Click()
grdMov.Clear
grdMov.Rows = 2
grdMov.FormaCabecera
fraBuscar.Enabled = True
cmdExtornar.Enabled = False
cmdCancelar.Enabled = False
optTipoBus(0).value = True

Call ExtDatosOculta '**** cti3

End Sub
'***CTI3 (FERIMORO) 03102018
Private Sub cmdExtornar_Click()

If Trim(grdMov.TextMatrix(1, 2)) <> "" Then
    If lsCodExtOpc = "270501" Or lsCodExtOpc = "270502" Or lsCodExtOpc = "270503" Or lsCodExtOpc = "270504" Then
        cmdExtContinuar_Click
        Exit Sub
    Else
        frmMotExtorno.Visible = True
        fraBuscar.Enabled = False
        cmdExtornar.Enabled = False
        grdMov.Enabled = False
        cmbMotivos.SetFocus
    End If
End If

End Sub

Sub ExtDatosOculta()
 If lsCodExtOpc = "270501" Or lsCodExtOpc = "270502" Or lsCodExtOpc = "270503" Or lsCodExtOpc = "270504" Then
    txtGlosa.Enabled = False
    txtGlosa.BackColor = &H80000004
    fraBuscar.Enabled = True
    cmdExtornar.Enabled = False
    grdMov.Enabled = True
 Else
    frmMotExtorno.Visible = False
    fraBuscar.Enabled = True
    cmdExtornar.Enabled = False
    Me.cmbMotivos.ListIndex = -1
    Me.txtDetExtorno.Text = ""
    grdMov.Enabled = True
 End If
End Sub

Private Sub cmdExtContinuar_Click()
Dim sGlosa As String
Dim lsCtaAhoExt As String 'CTI4 ERS0112020
Dim lsClienteAhoExt As String  'CTI4 ERS0112020
Dim lnMontoAhoExt As Currency 'CTI4 ERS0112020
Dim lsImpCargoCta As String 'CTI4 ERS0112020
Dim lsFechaHoraGrab As String 'CTI4 ERS0112020
If Trim(grdMov.TextMatrix(1, 2)) <> "" Then
    
    '*** PEAC 20081002
    Dim lbResultadoVisto As Boolean
    Dim sPersVistoCod  As String
    Dim sPersVistoCom As String
    Dim loVistoElectronico As frmVistoElectronico
    Set loVistoElectronico = New frmVistoElectronico
    Dim bestPago As Boolean 'MADM 20111220
    Dim oNCred As COMNCredito.NCOMCredito: Dim rsdev As ADODB.Recordset 'RIRO20140610 ERS017
    Dim oSegSep As New COMNCaptaGenerales.NCOMSeguros 'RECO20160209 ERS073-2015
        
    '****CTI3 (ferimoro)
    If lsCodExtOpc = "270501" Or lsCodExtOpc = "270502" Or lsCodExtOpc = "270503" Or lsCodExtOpc = "270504" Then
        sGlosa = UCase(txtDetExtorno.Text) 'Trim(txtGlosa)
    Else
        If cmbMotivos.ListIndex = -1 Or Len(txtDetExtorno.Text) <= 0 Then
            MsgBox "Debe ingresar el motivo y/o detalle del Extorno", vbInformation, "Aviso"
            Exit Sub
        End If
    
    '****cti3
    Dim DatosExtorna(1) As String
        DatosExtorna(0) = cmbMotivos.Text
        DatosExtorna(1) = txtDetExtorno.Text
        '********************
        frmMotExtorno.Visible = False
        sGlosa = UCase(txtDetExtorno.Text) ' CTI3 05112018
    End If
    
    'JUEZ 20131209 **********************************************************
'    If Trim(txtGlosa.Text) = "" Then
'        MsgBox "Es necesario que escriba la glosa", vbInformation, "Aviso"
'        txtGlosa.SetFocus
'        Exit Sub
'    End If
    'END JUEZ ***************************************************************
    
    'RIRO20131212 ERS137 ****************************************************
    If Not VerificarEstadoPendiente Then
        MsgBox "La opracion fue saldada por el módulo financiero, no es posible continuar con el proceso de extorno", vbExclamation, "Aviso"
        Call ExtDatosOculta '**** cti3
        Exit Sub
    End If
    'END RIRO ***************************************************************
    
    '*** PEAC 20081001 - visto electronico ******************************************************
    '*** en estos extornos de operaciones pedirá visto electrónico
    
    If Len(lsCodExtOpc) > 0 Then
    Select Case lsCodExtOpc
         Case "230101", "230201", "230301", "230302", "230401", "240101", _
              "240201", "240301", "250101", "250102", "250201", "270103", "270301", _
              "270302", "270501", "270503", "270504", "360101", "360201", "240203", _
              gCTSExtDepLotEfec, gCTSExtApeLoteEfec, gAhoExtCargoComDivAho, gCTSExtCargoComDivAho, gAhoExtDirectoClub, _
              "230303", "230411", "240204", "240304", "250302", "250403", _
              "100219", "100308", "100409", "100508", "100608", "100708", _
              "130207", "130306", "130406", "200110", "200267", "200103", _
              "200203", "210103", "210807", "220103", "220203", "230104", "230265", "230266", "230107" _
              , 290005, 290007, "250106", gCTSExtDepLotTransf, gPFExtApeLoteTransf, "230102", "230105", "230267", "240105" 'RECO20160209 ERS073-2015
             '***Parametro gCTSExtDepLotEfec,gCTSApeLoteEfec agregado por ELRO el 20121122, según OYP-RFC101-2012
             'JUEZ 20130906 Se agregó gAhoExtCargoComDivAho, gCTSExtCargoComDivAho
             ' RIRO20131212 ERS137 Se agregaron 230303,230411,240204,240304,250306,250406
             'RIRO20140610 ERS017, Se agrego "100219,100308,100409,100508,100608,100708,130207,130306,130406
             '                                200110,200267,200103,200203,210103,210807,220103,220203"
             'END RIRO
             ' *** RIRO SEGUN TI-ERS108-2013 ***
             Dim nMovNroOperacion As Long
             nMovNroOperacion = 0
             If grdMov.row >= 1 And Len(Trim(grdMov.TextMatrix(grdMov.row, 9))) > 0 Then
                 nMovNroOperacion = Val(grdMov.TextMatrix(grdMov.row, 9))
             End If
             ' *** FIN RIRO ***
             lbResultadoVisto = loVistoElectronico.Inicio(3, lsCodExtOpc, , , nMovNroOperacion) 'RIRO SEGUN TI-ERS108-2013 / Se agrego el campo nMovNroOperacion
             If Not lbResultadoVisto Then
                 Call ExtDatosOculta '**** cti3
                 Exit Sub
             End If
    End Select
    End If
    '*** FIN PEAC ************************************************************
       
    'RIRO20131212 ERS137 ****************************************************
    If Not VerificarEstadoPendiente Then
        MsgBox "La opracion fue saldada por el módulo financiero, no esposible continuar con el proceso de extorno", vbExclamation, "Aviso"
        Call ExtDatosOculta '**** cti3
        Exit Sub
    End If
    'END  RIRO

    'RIRO20140610 ERS017
    If InStr(1, "100219,100308,100409,100508,100608,100708,130207,130306,130406,200110,200267,200103,200203,210103,210807,220103,220203", CLng(grdMov.TextMatrix(grdMov.row, 7)), vbTextCompare) > 0 Then
        Set oNCred = New COMNCredito.NCOMCredito
        Set rsdev = oNCred.DevSobranteXope(CLng(grdMov.TextMatrix(grdMov.row, 9)))
        If Not rsdev Is Nothing Then
            If rsdev.RecordCount > 0 Then
                MsgBox "No se puede extornar la operacion porque ya se devolvio el sobrante del voucher", vbInformation, "Aviso"
                Set oNCred = Nothing: Set rsdev = Nothing: Call ExtDatosOculta  '**** cti3
                Exit Sub
            End If
        End If
        Set oNCred = Nothing: Set rsdev = Nothing
    End If
    'END RIRO
    
    'CTI2 20190405 ADD ****
    Dim bRespuesta As Boolean
    If lsCodExtOpc = gAhoExtDepositoHaberesEnLoteEfec Or _
       lsCodExtOpc = gAhoExtDepositoHaberesEnLoteTransf Then
       
        Dim oDMov As COMDCaptaGenerales.DCOMCaptaMovimiento
        Dim nMov As Long
        nMov = 0
        bRespuesta = False
        Set oDMov = New COMDCaptaGenerales.DCOMCaptaMovimiento
        nMov = CLng(grdMov.TextMatrix(grdMov.row, 9))
        
        bRespuesta = oDMov.ValidaExtornoLote(nMov)
        If Not bRespuesta Then
            MsgBox "No es posible realizar el extorno " & vbNewLine _
                   & "Hay cuentas que no cuentan con saldo para realizar el extorno", vbInformation, "Aviso - Extorno Depósito en Lote"
            Set oDMov = Nothing
            Exit Sub
        End If
        Set oDMov = Nothing
    End If
   'CTI2 20190405 END ****

    If MsgBox("¿Desea extornar la operación?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        Dim sMovNro As String, sMovNroBus As String ', sGlosa As String -- CTI3
        Dim nMovNroBus As Long
        Dim sCuenta As String, sNroDoc As String, sDescOpe As String, sPersCod As String
        Dim nTipoDoc As tpoDoc
        Dim bDocumento As Boolean
        Dim nOperacion As COMDConstantes.CaptacOperacion
        Dim clsMov As COMNContabilidad.NCOMContFunciones
        Dim nMonto As Double
        Dim nMontoITF As Double
        Dim clsCap As COMNCaptaGenerales.NCOMCaptaMovimiento
        Dim lnI As Integer, nFila As Integer
        Dim nITFOperacion As COMDConstantes.CaptacOperacion
        Dim nITFConcepto As COMDConstantes.CaptacConcepto
        Dim lsBoleta As String
        Dim lsBoletaITF As String
        Dim lnMovNro As Long 'RECO20160209 ERS073-2015
        'ALPA 20091119****************************
        Dim lsBoletaAbono As String
        '*****************************************
        Dim psBoletaExtorno As String
        Dim lsmensaje As String
        
        Dim nComiRetOtraAge As Double
        Dim nComiRetxMaxOpe As Double
        Dim nConfirmarConciliacion As Integer '***Agregado por ELRO el 20121221, según OYP-RFC0242012
        '***Agregado por ELRO el 20121203, según OYP-RFC101-2012
        Dim rsCuentasPorExtornar As ADODB.Recordset
        Dim oDCOMCaptaMovimiento As COMDCaptaGenerales.DCOMCaptaMovimiento
        '***Fin Agregado por ELRO el 20121203*******************
        
        nFila = grdMov.row
        sMovNroBus = grdMov.TextMatrix(nFila, 1)
        nOperacion = CLng(grdMov.TextMatrix(nFila, 7))
        nMonto = CDbl(grdMov.TextMatrix(nFila, 4))
        
        'madm 20111207
        bestPago = IIf(Trim(grdMov.TextMatrix(nFila, 6)) = "Pago Servcios", True, False)
        'end madm
        
        Dim pActualiza As COMDCaptaGenerales.DCOMCaptaMovimiento
          Set pActualiza = New COMDCaptaGenerales.DCOMCaptaMovimiento
        
        If grdMov.TextMatrix(nFila, 12) = "" Then
            nMontoITF = 0
            nITFOperacion = 0
            nITFConcepto = 0
        Else
            nMontoITF = CDbl(grdMov.TextMatrix(nFila, 11))
            nITFOperacion = CLng(grdMov.TextMatrix(nFila, 12))
            nITFConcepto = CInt(grdMov.TextMatrix(nFila, 13))
        End If
        
        If grdMov.TextMatrix(nFila, 15) = "" Then
            nComiRetxMaxOpe = 0
        Else
            nComiRetxMaxOpe = CDbl(grdMov.TextMatrix(nFila, 15))
        End If
        
         If grdMov.TextMatrix(nFila, 14) = "" Then
            nComiRetOtraAge = 0
        Else
            nComiRetOtraAge = CDbl(grdMov.TextMatrix(nFila, 14))
        End If
        Dim nComiDepOtraAge As Double
        If grdMov.TextMatrix(nFila, 16) = "" Then
            nComiDepOtraAge = 0
        Else
            nComiDepOtraAge = CDbl(grdMov.TextMatrix(nFila, 16))
        End If
        
        'add by GITU 11-12-2013
        Dim nComiRetSinTarj As Double
        If grdMov.TextMatrix(nFila, 17) = "" Then
            nComiRetSinTarj = 0
        Else
            nComiRetSinTarj = CDbl(grdMov.TextMatrix(nFila, 17))
        End If
        'end GITU
        
        'RIRO20131212 ERS137 *******
        Dim nComiTransf As Double
        Dim bComiDebito As Boolean 'CTI4 ERS0112020
        'If Len(Trim(grdMov.TextMatrix(nFila, 18))) = 0 Or Val(Trim(grdMov.TextMatrix(nFila, 19))) = 0 Then 'Comentado CTI4 ERS0112020
        If Len(Trim(grdMov.TextMatrix(nFila, 18))) = 0 And Val(Trim(grdMov.TextMatrix(nFila, 19))) = 0 Then 'CTI4 ERS0112020
            nComiTransf = 0
            bComiDebito = False 'CTI4 ERS0112020
        Else
            nComiTransf = CDbl(Trim(grdMov.TextMatrix(nFila, 18)))
            bComiDebito = Val(Trim(grdMov.TextMatrix(nFila, 19))) 'CTI4 ERS0112020
        End If
        'END RIRO ******************
        
        'CTI4 ERS0112020
        Dim nComEmiCheque As Double
        Dim psComEmiChequeOperacion As String
        If grdMov.TextMatrix(nFila, 20) = "" Then
            nComEmiCheque = 0
            psComEmiChequeOperacion = 0
        Else
            nComEmiCheque = grdMov.TextMatrix(nFila, 20)
            psComEmiChequeOperacion = grdMov.TextMatrix(nFila, 21)
        End If
        'CTI4 end
        
        sCuenta = grdMov.TextMatrix(nFila, 3)
        sNroDoc = grdMov.TextMatrix(nFila, 5)
        sDescOpe = grdMov.TextMatrix(nFila, 2)
        nMovNroBus = CLng(grdMov.TextMatrix(nFila, 9))
        If sNroDoc <> "" Then
            bDocumento = True
            nTipoDoc = CLng(grdMov.TextMatrix(nFila, 8))
            sPersCod = Trim(grdMov.TextMatrix(nFila, 10))
        Else
            bDocumento = False
        End If
        sGlosa = Trim(txtGlosa)
        
        Dim oCap As COMNCaptaGenerales.NCOMCaptaMovimiento
        Set oCap = New COMNCaptaGenerales.NCOMCaptaMovimiento
        
        
        'JUEZ 20131226 Se agregó gAhoApeCargoCta y gPFApeCargoCta
        If nOperacion = gAhoApeChq Or nOperacion = gAhoApeEfec Or nOperacion = gAhoApeTransf Or nOperacion = gAhoApeCargoCta _
            Or nOperacion = gPFApeChq Or nOperacion = gPFApeEfec Or nOperacion = gPFApeTransf Or nOperacion = gPFApeCargoCta _
            Or nOperacion = gCTSApeChq Or nOperacion = gCTSApeEfec Or nOperacion = gCTSApeTransf Then
            If oCap.TieneMovDespuesApertura(sCuenta) Then
                MsgBox "Cuenta " & sCuenta & " posee movimientos despues de la apertura, por favor extorne antes esos movimientos.", vbInformation, "Aviso"
                Call ExtDatosOculta '**** cti3
                Exit Sub
            End If
        ElseIf nOperacion = gAhoApeLoteChq Or nOperacion = gAhoApeLoteEfec _
            Or nOperacion = gPFApeLoteChq Or nOperacion = gPFApeLoteEfec _
            Or nOperacion = gCTSApeLoteChq Or nOperacion = gCTSApeLoteEfec Then
            For lnI = 1 To Me.grdMov.Rows - 1
                If sMovNroBus = grdMov.TextMatrix(lnI, 1) Then
                    If oCap.TieneMovDespuesApertura(grdMov.TextMatrix(lnI, 3)) Then
                        MsgBox "Cuenta " & grdMov.TextMatrix(lnI, 3) & " posee movimientos despues de la apertura, por favor extorne antes esos movimientos.", vbInformation, "Aviso"
                        Call ExtDatosOculta '**** cti3
                        Exit Sub
                    End If
                End If
            Next lnI
        End If
        
        '***Agregado por ELRO el 20121203, según OYP-RFC101-2012
        If nOperacion = gCTSApeLoteEfec Then
            Set rsCuentasPorExtornar = New ADODB.Recordset
            Set oDCOMCaptaMovimiento = New COMDCaptaGenerales.DCOMCaptaMovimiento
            Set rsCuentasPorExtornar = oDCOMCaptaMovimiento.devolverCuentasPorExtornarLote(nMovNroBus, CStr(nOperacion))
            Do While Not rsCuentasPorExtornar.EOF
                If oCap.TieneMovDespuesApertura(rsCuentasPorExtornar!cCtaCod) Then
                    MsgBox "Cuenta " & rsCuentasPorExtornar!cCtaCod & " posee movimientos despues de la apertura, por favor extorne antes esos movimientos.", vbInformation, "Aviso"
                    Call ExtDatosOculta '**** cti3
                    Exit Sub
                End If
                rsCuentasPorExtornar.MoveNext
            Loop
            Set rsCuentasPorExtornar = Nothing
             Set oDCOMCaptaMovimiento = Nothing
        End If
        If nOperacion = gCTSDepLotEfec Then
            Dim nMovNroUltDep As Long
            Set rsCuentasPorExtornar = New ADODB.Recordset
            Set oDCOMCaptaMovimiento = New COMDCaptaGenerales.DCOMCaptaMovimiento
            Set rsCuentasPorExtornar = oDCOMCaptaMovimiento.devolverCuentasPorExtornarLote(nMovNroBus, CStr(nOperacion))
            Do While Not rsCuentasPorExtornar.EOF
                nMovNroUltDep = oDCOMCaptaMovimiento.devolverUltimoMovimientoDeposito(rsCuentasPorExtornar!cCtaCod, Format(gdFecSis, "yyyyMMdd"))
                If nMovNroUltDep > 0 And nMovNroBus <> nMovNroUltDep Then
                    MsgBox "Cuenta " & rsCuentasPorExtornar!cCtaCod & " posee movimientos despues del depósito, por favor extorne antes esos movimientos.", vbInformation, "Aviso"
                    Call ExtDatosOculta '**** cti3
                    Exit Sub
                End If
                rsCuentasPorExtornar.MoveNext
            Loop
            Set rsCuentasPorExtornar = Nothing
            Set oDCOMCaptaMovimiento = Nothing
            nMovNroUltDep = 0
        End If
        '***Fin Agregado por ELRO el 20121203*******************
        '***Agregado por ELRO el 20130401, según TI-ERS011-2013****
        If nOperacion = gAhoMigracion Then
            Dim nMovNroUlt As Long
            Set rsCuentasPorExtornar = New ADODB.Recordset
            Set oDCOMCaptaMovimiento = New COMDCaptaGenerales.DCOMCaptaMovimiento
            Set rsCuentasPorExtornar = oDCOMCaptaMovimiento.devolverCuentasPorExtornarLote(nMovNroBus, CStr(nOperacion))
            Do While Not rsCuentasPorExtornar.EOF
                nMovNroUlt = oDCOMCaptaMovimiento.devolverUltimoMovimientoDeposito(rsCuentasPorExtornar!cCtaCod, Format(gdFecSis, "yyyyMMdd"))
                If nMovNroUlt > 0 And nMovNroBus <> nMovNroUlt Then
                    MsgBox "Cuenta " & rsCuentasPorExtornar!cCtaCod & " posee movimientos después de la migración, por favor extorne antes esos movimientos.", vbInformation, "Aviso"
                    Call ExtDatosOculta '**** cti3
                    Exit Sub
                End If
                rsCuentasPorExtornar.MoveNext
            Loop
            Set rsCuentasPorExtornar = Nothing
            Set oDCOMCaptaMovimiento = Nothing
            nMovNroUltDep = 0
        End If
        '***Fin Agregado por ELRO el 20130401, según TI-ERS011-2013
        
        Set clsMov = New COMNContabilidad.NCOMContFunciones
            sMovNro = clsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        Set clsMov = Nothing
        Set clsCap = New COMNCaptaGenerales.NCOMCaptaMovimiento
                
        lsFechaHoraGrab = fgFechaHoraGrab(sMovNro) 'CTI4 ERS0112020
                
        Select Case nOperacion
            Case gAhoApeChq 'Ahorro Apertura de Cheques gAhoExtApeChq = 230102
                clsCap.CapExtornoApertura nMovNroBus, gAhoExtApeChq, sCuenta, sMovNro, sGlosa, nMonto, bDocumento, nTipoDoc, sNroDoc, sPersCod, gsNomAge, sLpt, gsCodCMAC, , , , , lsBoleta, lsBoletaITF, gbImpTMU, DatosExtorna  '***CTI3 (ferimoro) 04102018
            Case gAhoApeTransf 'Ahorro Apertura Nota de Abono
                '***Agregado por ELRO el 20121221, según OYP-RFC024-2012
                nConfirmarConciliacion = clsCap.verificarConciliacionBancos(nMovNroBus)
                If nConfirmarConciliacion >= 1 Then
                   MsgBox "No puede extornar el movimiento porque ya fue conciliado por el área de Finanzas.", vbInformation, "!Aviso¡"
                   Call ExtDatosOculta '**** cti3
                   Exit Sub
                End If
                '***Fin Agregado por ELRO el 20121221******************* gAhoExtApeTransf = 230103
                clsCap.CapExtornoApertura nMovNroBus, gAhoExtApeTransf, sCuenta, sMovNro, sGlosa, nMonto, bDocumento, nTipoDoc, sNroDoc, , gsNomAge, sLpt, gsCodCMAC, , , , , lsBoleta, lsBoletaITF, , DatosExtorna '***CTI3 (ferimoro) 04102018
            Case gAhoApeEfec 'Ahorro Apertura Efectivo gAhoExtApeEfec = 230101
                clsCap.CapExtornoApertura nMovNroBus, gAhoExtApeEfec, sCuenta, sMovNro, sGlosa, nMonto, bDocumento, nTipoDoc, sNroDoc, , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsBoleta, lsBoletaITF, , DatosExtorna '***CTI3 (ferimoro) 04102018
            Case gAhoApeLoteEfec 'Ahorro Apertura Lote Efectivo gAhoExtApeLoteEfec = 230104
                clsCap.CapExtornoAperturaLote nMovNroBus, gAhoExtApeLoteEfec, sCuenta, sMovNro, sGlosa, nMonto, bDocumento, nTipoDoc, sNroDoc, , gsNomAge, sLpt, gsCodCMAC, lsBoleta, , DatosExtorna '***CTI3 (ferimoro) 04102018
                cmdCancelar_Click 'RIRO20140610 ERS017
            Case gAhoApeLoteChq  'Ahorro Apertura Lote Cheque  gAhoExtApeLoteChq = 230105
                clsCap.CapExtornoAperturaLote nMovNroBus, gAhoExtApeLoteChq, sCuenta, sMovNro, sGlosa, nMonto, bDocumento, nTipoDoc, sNroDoc, sPersCod, gsNomAge, sLpt, gsCodCMAC, , lsBoleta, DatosExtorna, gbImpTMU '***CTI3 (ferimoro) 04102018
            Case gAhoApeCargoCta 'Ahorro Apertura Cargo Cuenta 'JUEZ 20131226  gAhoExtApeCargoCta = 230106
                clsCap.CapExtornoApertura nMovNroBus, gAhoExtApeCargoCta, sCuenta, sMovNro, sGlosa, nMonto, bDocumento, nTipoDoc, sNroDoc, , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsBoleta, lsBoletaITF, , DatosExtorna '***CTI3 (ferimoro) 04102018
            Case gAhoApeLoteTransfBanco 'Ahorro Apertura Lote Transferencia 'RIRO 20140528 ERS017 gAhoExtApeLoteTransfBanco = 230107
                clsCap.CapExtornoAperturaLoteTransf nMovNroBus, gAhoExtApeLoteTransfBanco, sCuenta, sMovNro, sGlosa, nMonto, gsNomAge, sLpt, gsCodCMAC, gsCodAge, lsBoleta, DatosExtorna, gbImpTMU '***CTI3 (ferimoro) 04102018
                cmdCancelar_Click 'RIRO20140610 ERS017
'**********************************************************************************
            Case gCTSApeLoteTransfNew 'CTS Apertura Lote Transferencia 'CTI7 OPEv2
                clsCap.CapExtornoAperturaLoteTransf nMovNroBus, gCTSExtApeLoteTransf, sCuenta, sMovNro, sGlosa, nMonto, gsNomAge, sLpt, gsCodCMAC, gsCodAge, lsBoleta, DatosExtorna, gbImpTMU '***CTI3 (ferimoro) 04102018
                cmdCancelar_Click
'**********************************************************************************
            Case gPFApeChq 'Plazo Fijo Apertura Cheque gPFExtApeChq = 240102
                clsCap.CapExtornoApertura nMovNroBus, gPFExtApeChq, sCuenta, sMovNro, sGlosa, nMonto, bDocumento, nTipoDoc, sNroDoc, sPersCod, gsNomAge, sLpt, gsCodCMAC, , , , , lsBoleta, lsBoletaITF, gbImpTMU, DatosExtorna '***CTI3 (ferimoro) 04102018
            Case gPFApeTransf 'Plazo Fijo Apertura Nota Abono
                '***Agregado por ELRO el 20121221, según OYP-RFC024-2012
                nConfirmarConciliacion = clsCap.verificarConciliacionBancos(nMovNroBus)
                If nConfirmarConciliacion >= 1 Then
                   MsgBox "No puede extornar el movimiento porque ya fue conciliado por el área de Finanzas.", vbInformation, "!Aviso¡"
                   Call ExtDatosOculta '**** cti3
                   Exit Sub
                End If
                '***Fin Agregado por ELRO el 20121221******************* gPFExtApeTransf = 240103
                clsCap.CapExtornoApertura nMovNroBus, gPFExtApeTransf, sCuenta, sMovNro, sGlosa, nMonto, bDocumento, nTipoDoc, sNroDoc, , gsNomAge, sLpt, gsCodCMAC, , , , , lsBoleta, lsBoletaITF, gbImpTMU, DatosExtorna '***CTI3 (ferimoro) 04102018
            Case gPFApeEfec 'Plazo Fijo Apertura Efectivo gPFExtApeEfec = 240101
                clsCap.CapExtornoApertura nMovNroBus, gPFExtApeEfec, sCuenta, sMovNro, sGlosa, nMonto, bDocumento, nTipoDoc, sNroDoc, , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsBoleta, lsBoletaITF, , DatosExtorna '***CTI3 (ferimoro) 04102018
            Case gPFApeLoteEfec  'Plazo Fijo Apertura Lote Efectivo gPFExtApeLoteEfec = 240104
                clsCap.CapExtornoAperturaLote nMovNroBus, gPFExtApeLoteEfec, sCuenta, sMovNro, sGlosa, nMonto, bDocumento, nTipoDoc, sNroDoc, , gsNomAge, sLpt, gsCodCMAC, , lsBoleta, DatosExtorna '***CTI3 (ferimoro) 04102018
            Case gPFApeLoteChq 'Plazo Fijo Apertura Lote Cheque gPFExtApeLoteChq = 240105
                clsCap.CapExtornoAperturaLote nMovNroBus, gPFExtApeLoteChq, sCuenta, sMovNro, sGlosa, nMonto, bDocumento, nTipoDoc, sNroDoc, sPersCod, gsNomAge, sLpt, gsCodCMAC, , lsBoleta, DatosExtorna, gbImpTMU '***CTI3 (ferimoro) 04102018
            Case gPFApeCargoCta 'Plazo Fijo Apertura Cargo Cuenta 'JUEZ 20131226 gPFExtApeCargoCta = 240106
                clsCap.CapExtornoApertura nMovNroBus, gPFExtApeCargoCta, sCuenta, sMovNro, sGlosa, nMonto, bDocumento, nTipoDoc, sNroDoc, , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsBoleta, lsBoletaITF, gbImpTMU, DatosExtorna '***CTI3 (ferimoro) 04102018
            
'**********************************************************************************
            Case gPFApeLoteTransf 'CTS Apertura Lote Transferencia 'CTI7 OPEv2
                clsCap.CapExtornoAperturaLoteTransf nMovNroBus, gPFExtApeLoteTransf, sCuenta, sMovNro, sGlosa, nMonto, gsNomAge, sLpt, gsCodCMAC, gsCodAge, lsBoleta, DatosExtorna, gbImpTMU '***CTI3 (ferimoro) 04102018
                cmdCancelar_Click
'**********************************************************************************
            
            Case gCTSApeChq 'CTS Apertura Cheque
                                                    'gCTSExtApeChq = 250102
                clsCap.CapExtornoApertura nMovNroBus, gCTSExtApeChq, sCuenta, sMovNro, sGlosa, nMonto, bDocumento, nTipoDoc, sNroDoc, sPersCod, gsNomAge, sLpt, , , , , , lsBoleta, lsBoletaITF, gbImpTMU, DatosExtorna '***CTI3 (ferimoro) 04102018
            Case gCTSApeTransf 'CTS Apertura Nota Abono
                '***Agregado por ELRO el 20121221, según OYP-RFC024-2012
                nConfirmarConciliacion = clsCap.verificarConciliacionBancos(nMovNroBus)
                If nConfirmarConciliacion >= 1 Then
                   MsgBox "No puede extornar el movimiento porque ya fue conciliado por el área de Finanzas.", vbInformation, "!Aviso¡"
                   Call ExtDatosOculta '**** cti3
                   Exit Sub
                End If
                '***Fin Agregado por ELRO el 20121221******************* gCTSExtApeTransf = 250103
                clsCap.CapExtornoApertura nMovNroBus, gCTSExtApeTransf, sCuenta, sMovNro, sGlosa, nMonto, bDocumento, nTipoDoc, sNroDoc, , gsNomAge, sLpt, , , , , , lsBoleta, lsBoletaITF, gbImpTMU, DatosExtorna '***CTI3 (ferimoro) 04102018
            Case gCTSApeEfec 'CTS Apertura Efectivo  gCTSExtApeEfec = 250101
                clsCap.CapExtornoApertura nMovNroBus, gCTSExtApeEfec, sCuenta, sMovNro, sGlosa, nMonto, bDocumento, nTipoDoc, sNroDoc, , gsNomAge, sLpt, gsCodCMAC, , , , , lsBoleta, lsBoletaITF, gbImpTMU, DatosExtorna '***CTI3 (ferimoro) 04102018
            Case gCTSApeLoteEfec 'CTS Apertura Lote Efectivo gCTSExtApeLoteEfec = 250104
                clsCap.CapExtornoAperturaLote nMovNroBus, gCTSExtApeLoteEfec, sCuenta, sMovNro, sGlosa, nMonto, bDocumento, nTipoDoc, sNroDoc, , gsNomAge, sLpt, , , lsBoleta, DatosExtorna '***CTI3 (ferimoro) 04102018
            Case gCTSApeLoteChq 'CTS Apertura Lote Cheque gCTSExtApeLoteChq = 250105
                clsCap.CapExtornoAperturaLote nMovNroBus, gCTSExtApeLoteChq, sCuenta, sMovNro, sGlosa, nMonto, bDocumento, nTipoDoc, sNroDoc, sPersCod, gsNomAge, sLpt, , , lsBoleta, DatosExtorna, gbImpTMU '***CTI3 (ferimoro) 04102018
            
            Case gServGiroApertEfec, gServGiroApertCargoCta, gServGiroApertVoucher                   'gServExtGiroApertEfec = 360101
                clsCap.ServGiroExtornoApertura nMovNroBus, gServExtGiroApertEfec, sCuenta, sMovNro, sGlosa, nMonto, gsNomAge, sLpt, lsBoleta, gbImpTMU, DatosExtorna, nOperacion, psBoletaExtorno, gsNomAge '***CTI3 (ferimoro) 04102018
                
            Case gAhoDepEfec, "200243" 'Ahorro Deposito Efectivo     gAhoExtDepEfec = 230201
                'MADM 20111220 clsCap.CapExtornoAbonoAho nMovNroBus, gAhoExtDepEfec, sCuenta, sMovNro, sGlosa, nMonto, , , , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , , lsBoleta, lsBoletaITF, gbImpTMU, nComiDepOtraAge
                'MAVM 20120207 Se agrego la var: nComiDepOtraAge
                'clsCap.CapExtornoAbonoAho nMovNroBus, gAhoExtDepEfec, sCuenta, sMovNro, sGlosa, nMonto, , , , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsmensaje, lsBoleta, lsBoletaITF, gbImpTMU, nComiDepOtraAge, bestPago, , DatosExtorna 'CTI3 04102018 'JUEZ 20131209 Se agregò lsmensaje
                clsCap.CapExtornoAbonoAho nMovNroBus, gAhoExtDepEfec, sCuenta, sMovNro, sGlosa, nMonto, , , , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsmensaje, lsBoleta, lsBoletaITF, gbImpTMU, nComiDepOtraAge, bestPago, , DatosExtorna, nComiRetxMaxOpe 'APRI20190109 ERS077-2018
            'FRHU 20150128 ERS048-2014
            Case gCapNotaDeAbono 'Extorno de Nota de Abono gCapExtNotaDeAbono = 330002
                clsCap.CapExtornoAbonoAho nMovNroBus, gCapExtNotaDeAbono, sCuenta, sMovNro, sGlosa, nMonto, , , , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsmensaje, lsBoleta, lsBoletaITF, gbImpTMU, nComiDepOtraAge, bestPago, , DatosExtorna 'CTI3 04102018
            'FIN FRHU 20150128
            Case gAhoDepChq, "200244" 'Ahorro Depósito Cheque gAhoExtDepChq = 230202
                'clsCap.CapExtornoAbonoAho nMovNroBus, gAhoExtDepChq, sCuenta, sMovNro, sGlosa, nMonto, TpoDocCheque, sNroDoc, sPersCod, gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsmensaje, lsBoleta, lsBoletaITF, gbImpTMU, , , , DatosExtorna 'CTI3 04102018 'JUEZ 20131209 Se agregò lsmensaje
                clsCap.CapExtornoAbonoAho nMovNroBus, gAhoExtDepChq, sCuenta, sMovNro, sGlosa, nMonto, TpoDocCheque, sNroDoc, sPersCod, gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsmensaje, lsBoleta, lsBoletaITF, gbImpTMU, , , , DatosExtorna, nComiRetxMaxOpe 'APRI20190109 ERS077-2018
            Case gAhoDepTransf, "200245" 'Ahorro Depósito Nota Abono
                '***Agregado por ELRO el 20121221, según OYP-RFC024-2012
                nConfirmarConciliacion = clsCap.verificarConciliacionBancos(nMovNroBus)
                If nConfirmarConciliacion >= 1 Then
                   MsgBox "No puede extornar el movimiento porque ya fue conciliado por el área de Finanzas.", vbInformation, "!Aviso¡"
                   Call ExtDatosOculta '**** cti3
                   Exit Sub
                End If
                '***Fin Agregado por ELRO el 20121221******************* gAhoExtDepTransf = 230203
                'clsCap.CapExtornoAbonoAho nMovNroBus, gAhoExtDepTransf, sCuenta, sMovNro, sGlosa, nMonto, , , , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsmensaje, lsBoleta, lsBoletaITF, gbImpTMU, , , , DatosExtorna 'CTI3 04102018
                clsCap.CapExtornoAbonoAho nMovNroBus, gAhoExtDepTransf, sCuenta, sMovNro, sGlosa, nMonto, , , , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsmensaje, lsBoleta, lsBoletaITF, gbImpTMU, , , , DatosExtorna, nComiRetxMaxOpe 'APRI20190109 ERS077-2018
            Case gAhoDepPagServEdelnor  'Ahorro Depósito Pago de servicio edelnor gAhoExtDepPagServEdelnor = 230219
                clsCap.CapExtornoAbonoAho nMovNroBus, gAhoExtDepPagServEdelnor, sCuenta, sMovNro, sGlosa, nMonto, , , , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsmensaje, lsBoleta, lsBoletaITF, gbImpTMU, , , , DatosExtorna 'CTI3 04102018
            Case gAhoDepAboOtrosConceptos   'Ahorro Depósito Otros Conceptos gAhoExtDepOtrosConceptos = 230222
                clsCap.CapExtornoAbonoAho nMovNroBus, gAhoExtDepOtrosConceptos, sCuenta, sMovNro, sGlosa, nMonto, , , , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsmensaje, lsBoleta, lsBoletaITF, gbImpTMU, , , , DatosExtorna 'CTI3 04102018
            Case gAhoDepDevCredPersonales    'Ahorro Depósito Otros Conceptos gAhoExtDepDevCredPersonales = 230223
                clsCap.CapExtornoAbonoAho nMovNroBus, gAhoExtDepDevCredPersonales, sCuenta, sMovNro, sGlosa, nMonto, , , , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsmensaje, lsBoleta, lsBoletaITF, gbImpTMU, , , , DatosExtorna 'CTI3 04102018
            Case "200207"
                clsCap.CapExtornoAbonoAho nMovNroBus, "230207", sCuenta, sMovNro, sGlosa, nMonto, , , , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsmensaje, lsBoleta, lsBoletaITF, gbImpTMU, , , , DatosExtorna 'CTI3 04102018
            Case "200209"
                clsCap.CapExtornoAbonoAho nMovNroBus, "230209", sCuenta, sMovNro, sGlosa, nMonto, , , , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsmensaje, lsBoleta, lsBoletaITF, gbImpTMU, , , , DatosExtorna 'CTI3 04102018
            Case "200246"
                clsCap.CapExtornoAbonoAho nMovNroBus, 230246, sCuenta, sMovNro, sGlosa, nMonto, , , , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsmensaje, lsBoleta, lsBoletaITF, gbImpTMU, , , , DatosExtorna 'CTI3 04102018
            Case "200247"
                clsCap.CapExtornoAbonoAho nMovNroBus, 230247, sCuenta, sMovNro, sGlosa, nMonto, , , , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsmensaje, lsBoleta, lsBoletaITF, gbImpTMU, , , , DatosExtorna 'CTI3 04102018
            Case "200248"
                clsCap.CapExtornoAbonoAho nMovNroBus, 230248, sCuenta, sMovNro, sGlosa, nMonto, , , , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsmensaje, lsBoleta, lsBoletaITF, gbImpTMU, , , , DatosExtorna 'CTI3 04102018
            Case "200249"
                clsCap.CapExtornoAbonoAho nMovNroBus, 230249, sCuenta, sMovNro, sGlosa, nMonto, , , , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsmensaje, lsBoleta, lsBoletaITF, gbImpTMU, , , , DatosExtorna 'CTI3 04102018
            Case "200250"
                clsCap.CapExtornoAbonoAho nMovNroBus, 230250, sCuenta, sMovNro, sGlosa, nMonto, , , , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsmensaje, lsBoleta, lsBoletaITF, gbImpTMU, , , , DatosExtorna 'CTI3 04102018
            Case "200251"
                clsCap.CapExtornoAbonoAho nMovNroBus, 230251, sCuenta, sMovNro, sGlosa, nMonto, , , , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsmensaje, lsBoleta, lsBoletaITF, gbImpTMU, , , , DatosExtorna 'CTI3 04102018
            Case "200204"
                clsCap.CapExtornoAbonoAho nMovNroBus, 230252, sCuenta, sMovNro, sGlosa, nMonto, , , , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsmensaje, lsBoleta, lsBoletaITF, gbImpTMU, , , , DatosExtorna 'CTI3 04102018
            Case "200252"
                clsCap.CapExtornoAbonoAho nMovNroBus, 230254, sCuenta, sMovNro, sGlosa, nMonto, , , , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsmensaje, lsBoleta, lsBoletaITF, gbImpTMU, , , , DatosExtorna 'CTI3 04102018
            Case "200253"
                clsCap.CapExtornoAbonoAho nMovNroBus, 230255, sCuenta, sMovNro, sGlosa, nMonto, , , , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsmensaje, lsBoleta, lsBoletaITF, gbImpTMU, , , , DatosExtorna 'CTI3 04102018

            Case gAhoDepDirectoClub ' Deposito Club Trabajadores gAhoExtDirectoClub = 230264
                clsCap.CapExtornoAbonoAho nMovNroBus, gAhoExtDirectoClub, sCuenta, sMovNro, sGlosa, nMonto, , , , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsmensaje, lsBoleta, lsBoletaITF, gbImpTMU, , , , DatosExtorna 'CTI3 04102018
            Case gCMACOAAhoDepEfec 'Ahorro Deposito Efectivo Otra CMAC gCMACOAAhoExtDepEfec = 270101
                clsCap.CapExtornoAbonoAho nMovNroBus, gCMACOAAhoExtDepEfec, sCuenta, sMovNro, sGlosa, nMonto, , , , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsmensaje, lsBoleta, lsBoletaITF, gbImpTMU, , , , DatosExtorna 'CTI3 04102018
            Case gCMACOAAhoDepChq 'Ahorro Depósito Cheque Otra CMAC  gCMACOAAhoExtDepChq = 270102
                clsCap.CapExtornoAbonoAho nMovNroBus, gCMACOAAhoExtDepChq, sCuenta, sMovNro, sGlosa, nMonto, TpoDocCheque, sNroDoc, sPersCod, gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsmensaje, lsBoleta, lsBoletaITF, gbImpTMU, , , , DatosExtorna

            Case gCTSDepEfec 'CTS Depósito Efectivo gCTSExtDepEfec = 250201
                clsCap.CapExtornoAbonoCTS nMovNroBus, gCTSExtDepEfec, sCuenta, sMovNro, sGlosa, nMonto, , , , gsNomAge, sLpt, gsCodCMAC, , lsBoleta, gbImpTMU, DatosExtorna 'CTI3 04102018
            Case gCTSDepChq 'CTS Depósito Cheque     gCTSExtDepChq = 250202
                clsCap.CapExtornoAbonoCTS nMovNroBus, gCTSExtDepChq, sCuenta, sMovNro, sGlosa, nMonto, TpoDocCheque, sNroDoc, sPersCod, gsNomAge, sLpt, gsCodCMAC, , lsBoleta, gbImpTMU, DatosExtorna 'CTI3 04102018
            Case gCTSDepTransf 'CTS Depósito Nota Abono
                '***Agregado por ELRO el 20121221, según OYP-RFC024-2012
                nConfirmarConciliacion = clsCap.verificarConciliacionBancos(nMovNroBus)
                If nConfirmarConciliacion >= 1 Then
                   MsgBox "No puede extornar el movimiento porque ya fue conciliado por el área de Finanzas.", vbInformation, "!Aviso¡"
                   Call ExtDatosOculta '**** cti3
                   Exit Sub
                End If
                '***Fin Agregado por ELRO el 20121221******************* gCTSExtDepTransf = 250203
                clsCap.CapExtornoAbonoCTS nMovNroBus, gCTSExtDepTransf, sCuenta, sMovNro, sGlosa, nMonto, TpoDocNotaAbono, sMovNro, , gsNomAge, sLpt, gsCodCMAC, , lsBoleta, gbImpTMU, DatosExtorna 'CTI3 04102018
            Case gCMACOACTSDepEfec 'CTS Depósito Efectivo gCMACOACTSExtDepEfec = 270301
                clsCap.CapExtornoAbonoCTS nMovNroBus, gCMACOACTSExtDepEfec, sCuenta, sMovNro, sGlosa, nMonto, , , , gsNomAge, sLpt, gsCodCMAC, , lsBoleta, gbImpTMU, DatosExtorna 'CTI3 04102018
            '***Agregado por ELRO el 20121120, según OYP-RFC101-2012
            Case gCTSDepLotEfec, gCTSDepLotCargoCta 'CTS Depósito Efectivo Lote gCTSExtDepLotEfec = 250207
                'clsCap.CapExtornoAbonoCTSLote nMovNroBus, gCTSExtDepLotEfec, sMovNro, sGlosa, nMonto, gCTSDepLotEfec, , , , gsNomAge, sLpt, gsCodCMAC, , lsBoleta, gbImpTMU, DatosExtorna 'CTI3 04102018
                clsCap.CapExtornoAbonoCTSLote nMovNroBus, IIf(nOperacion = gCTSDepLotEfec, gCTSExtDepLotEfec, gCTSExtDepLotCargoCta), sMovNro, sGlosa, nMonto, nOperacion, , , , gsNomAge, sLpt, gsCodCMAC, , lsBoleta, gbImpTMU, DatosExtorna, lsCtaAhoExt, lsClienteAhoExt, lnMontoAhoExt 'CTI4 ERS0112020
            Case gCTSDepLotChq 'CTS Depósito Cheque Lote  gCTSExtDepLotChq = 250208
                'clsCap.CapExtornoAbonoCTSLote nMovNroBus, gCTSExtDepEfec, sCuenta, sMovNro, sGlosa, nMonto, , , , gsNomAge, sLpt, gsCodCMAC, , lsBoleta, gbImpTMU
                clsCap.CapExtornoAbonoCTSLote nMovNroBus, gCTSExtDepLotChq, sMovNro, sGlosa, nMonto, gCTSDepLotChq, , sNroDoc, sPersCod, gsNomAge, sLpt, gsCodCMAC, , lsBoleta, gbImpTMU, DatosExtorna 'PASI20140617
            Case gCTSDepLotTransf  'CTS Depósito Cheque Lote
                clsCap.CapExtornoAbonoCTSLote nMovNroBus, gCTSExtDepLotTransf, sMovNro, sGlosa, nMonto, nOperacion, , , , gsNomAge, sLpt, gsCodCMAC, , lsBoleta, gbImpTMU, DatosExtorna, lsCtaAhoExt, lsClienteAhoExt, lnMontoAhoExt
            '***Fin Agregado por ELRO el 20121120*******************
            Case gAhoRetEfec 'Ahorro Retiro Efectivo ---- add nComiRetSinTarj by GITU 11/12/2013  gAhoExtRetEfec = 230301
                clsCap.CapExtornoCargoAho nMovNroBus, gAhoExtRetEfec, sCuenta, sMovNro, sGlosa, nMonto, , , , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsBoleta, lsBoletaITF, nComiRetOtraAge, nComiRetxMaxOpe, gbImpTMU, , , nComiRetSinTarj, , , , DatosExtorna 'CTI3 04102018
            'FRHU 20150127 ERS048-2014
            Case gCapNotaDeCargo 'Extorno de Nota de Cargo gCapExtNotaDeCargo = 330001
                clsCap.CapExtornoCargoAho nMovNroBus, gCapExtNotaDeCargo, sCuenta, sMovNro, sGlosa, nMonto, , , , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsBoleta, lsBoletaITF, nComiRetOtraAge, nComiRetxMaxOpe, gbImpTMU, , , nComiRetSinTarj, , , , DatosExtorna 'CTI3 04102018
            'FIN FRHU
            Case 200316                             'gAhoExtRetRetencionJudicial = 230316
                clsCap.CapExtornoCargoAho nMovNroBus, gAhoExtRetRetencionJudicial, sCuenta, sMovNro, sGlosa, nMonto, , , , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsBoleta, lsBoletaITF, nComiRetOtraAge, nComiRetxMaxOpe, gbImpTMU, , , , , , , DatosExtorna 'CTI3 04102018
            
            Case 200601
                clsCap.CapExtornoCargoAho nMovNroBus, 230360, sCuenta, sMovNro, sGlosa, nMonto, , , , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsBoleta, lsBoletaITF, nComiRetOtraAge, nComiRetxMaxOpe, , , , , , , , DatosExtorna 'CTI3 04102018
            
            Case 200317                             'gAhoExtRetDuplicadoTarj = 230317
                clsCap.CapExtornoCargoAho nMovNroBus, gAhoExtRetDuplicadoTarj, sCuenta, sMovNro, sGlosa, nMonto, , , , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsBoleta, lsBoletaITF, nComiRetOtraAge, nComiRetxMaxOpe, gbImpTMU, , , , , , , DatosExtorna 'CTI3 04102018
            
            Case gAhoRetOP 'Ahorro Retiro con Orden de Pago gAhoExtRetOP = 230302
                clsCap.CapExtornoCargoAho nMovNroBus, gAhoExtRetOP, sCuenta, sMovNro, sGlosa, nMonto, TpoDocOrdenPago, sNroDoc, , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsBoleta, lsBoletaITF, nComiRetOtraAge, nComiRetxMaxOpe, gbImpTMU, , , , , , , DatosExtorna
            'RIRO20131212 ERS137
            Case gAhoRetTransf 'Ahorro Retiro Con Nota de Cargo gAhoExtRetTransf = 230303
                clsCap.CapExtornoCargoAho nMovNroBus, gAhoExtRetTransf, sCuenta, sMovNro, sGlosa, nMonto, , sNroDoc, , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsBoleta, lsBoletaITF, nComiRetOtraAge, nComiRetxMaxOpe, gbImpTMU, , , , , nComiTransf, , DatosExtorna, , , bComiDebito 'CTI3 09102018
            Case gAhoRetOPCanje 'Ahorro Retiro OP Canje gAhoExtRetOPCanje = 230304
                clsCap.CapExtornoCargoAho nMovNroBus, gAhoExtRetOPCanje, sCuenta, sMovNro, sGlosa, nMonto, TpoDocOrdenPago, sNroDoc, , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsBoleta, lsBoletaITF, nComiRetOtraAge, nComiRetxMaxOpe, gbImpTMU, , , , , , , DatosExtorna 'CTI3 04102018
            Case gAhoRetConsultaSaldos              'gAhoExtRetConsultaSaldos = 230326
                clsCap.CapExtornoCargoAho nMovNroBus, gAhoExtRetConsultaSaldos, sCuenta, sMovNro, sGlosa, nMonto, TpoDocOrdenPago, sNroDoc, , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsBoleta, lsBoletaITF, nComiRetOtraAge, nComiRetxMaxOpe, gbImpTMU, , , , , , , DatosExtorna 'CTI3 04102018
            Case gAhoRetOPCert 'Ahorro Retiro OP Certificada      gAhoExtRetOPCert = 230305
                clsCap.CapExtornoCargoAhoOPCertificada nMovNroBus, gAhoExtRetOPCert, sCuenta, sMovNro, sGlosa, nMonto, sNroDoc, gsNomAge, sLpt, gsCodCMAC, , lsBoleta, gbImpTMU
            Case gAhoRetOPCertCanje 'Ahorro Retiro OP Certificada Canje gAhoExtRetOPCertCanje = 230306
                clsCap.CapExtornoCargoAhoOPCertificada nMovNroBus, gAhoExtRetOPCertCanje, sCuenta, sMovNro, sGlosa, nMonto, sNroDoc, gsNomAge, sLpt, gsCodCMAC, , lsBoleta, gbImpTMU
            Case 200331
                clsCap.CapExtornoCargoAho nMovNroBus, 230332, sCuenta, sMovNro, sGlosa, nMonto, , , , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsBoleta, lsBoletaITF, nComiRetOtraAge, nComiRetxMaxOpe, gbImpTMU, , , , , , , DatosExtorna 'CTI3 04102018
            Case 200332
                clsCap.CapExtornoCargoAho nMovNroBus, 230333, sCuenta, sMovNro, sGlosa, nMonto, , , , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsBoleta, lsBoletaITF, nComiRetOtraAge, nComiRetxMaxOpe, gbImpTMU, , , , , , , DatosExtorna 'CTI3 04102018
            Case 200333
                clsCap.CapExtornoCargoAho nMovNroBus, 230334, sCuenta, sMovNro, sGlosa, nMonto, , , , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsBoleta, lsBoletaITF, nComiRetOtraAge, nComiRetxMaxOpe, gbImpTMU, , , , , , , DatosExtorna
            Case 200334
                clsCap.CapExtornoCargoAho nMovNroBus, 230335, sCuenta, sMovNro, sGlosa, nMonto, , , , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsBoleta, lsBoletaITF, nComiRetOtraAge, nComiRetxMaxOpe, gbImpTMU, , , , , , , DatosExtorna
            Case 200335
                clsCap.CapExtornoCargoAho nMovNroBus, 230336, sCuenta, sMovNro, sGlosa, nMonto, , , , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsBoleta, lsBoletaITF, nComiRetOtraAge, nComiRetxMaxOpe, gbImpTMU, , , , , , , DatosExtorna 'CTI3 04102018
            Case 200336
                clsCap.CapExtornoCargoAho nMovNroBus, 230337, sCuenta, sMovNro, sGlosa, nMonto, , , , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsBoleta, lsBoletaITF, nComiRetOtraAge, nComiRetxMaxOpe, gbImpTMU, , , , , , , DatosExtorna 'CTI3 04102018
            Case 200337
                clsCap.CapExtornoCargoAho nMovNroBus, 230338, sCuenta, sMovNro, sGlosa, nMonto, , , , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsBoleta, lsBoletaITF, nComiRetOtraAge, nComiRetxMaxOpe, gbImpTMU, , , , , , , DatosExtorna
            Case 200338
                clsCap.CapExtornoCargoAho nMovNroBus, 230339, sCuenta, sMovNro, sGlosa, nMonto, , , , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsBoleta, lsBoletaITF, nComiRetOtraAge, nComiRetxMaxOpe, gbImpTMU, , , , , , , DatosExtorna
            Case 200339
                clsCap.CapExtornoCargoAho nMovNroBus, 230340, sCuenta, sMovNro, sGlosa, nMonto, , , , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsBoleta, lsBoletaITF, nComiRetOtraAge, nComiRetxMaxOpe, gbImpTMU, , , , , , , DatosExtorna
            Case 200340
                clsCap.CapExtornoCargoAho nMovNroBus, 230341, sCuenta, sMovNro, sGlosa, nMonto, , , , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsBoleta, lsBoletaITF, nComiRetOtraAge, nComiRetxMaxOpe, gbImpTMU, , , , , , , DatosExtorna 'CTI3 04102018
            Case 200341
                clsCap.CapExtornoCargoAho nMovNroBus, 230342, sCuenta, sMovNro, sGlosa, nMonto, , , , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsBoleta, lsBoletaITF, nComiRetOtraAge, nComiRetxMaxOpe, gbImpTMU, , , , , , , DatosExtorna 'CTI3 04102018
            Case 200342
                clsCap.CapExtornoCargoAho nMovNroBus, 230343, sCuenta, sMovNro, sGlosa, nMonto, , , , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsBoleta, lsBoletaITF, nComiRetOtraAge, nComiRetxMaxOpe, gbImpTMU, , , , , , , DatosExtorna
            Case 200343
                clsCap.CapExtornoCargoAho nMovNroBus, 230344, sCuenta, sMovNro, sGlosa, nMonto, , , , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsBoleta, lsBoletaITF, nComiRetOtraAge, nComiRetxMaxOpe, gbImpTMU, , , , , , , DatosExtorna
            Case 200344
                clsCap.CapExtornoCargoAho nMovNroBus, 230345, sCuenta, sMovNro, sGlosa, nMonto, , , , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsBoleta, lsBoletaITF, nComiRetOtraAge, nComiRetxMaxOpe, gbImpTMU, , , , , , , DatosExtorna
            Case 200345
                clsCap.CapExtornoCargoAho nMovNroBus, 230346, sCuenta, sMovNro, sGlosa, nMonto, , , , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsBoleta, lsBoletaITF, nComiRetOtraAge, nComiRetxMaxOpe, gbImpTMU, , , , , , , DatosExtorna
            Case 200346
                clsCap.CapExtornoCargoAho nMovNroBus, 230347, sCuenta, sMovNro, sGlosa, nMonto, , , , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsBoleta, lsBoletaITF, nComiRetOtraAge, nComiRetxMaxOpe, gbImpTMU, , , , , , , DatosExtorna
            Case 200347
                clsCap.CapExtornoCargoAho nMovNroBus, 230348, sCuenta, sMovNro, sGlosa, nMonto, , , , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsBoleta, lsBoletaITF, nComiRetOtraAge, nComiRetxMaxOpe, gbImpTMU, , , , , , , DatosExtorna 'CTI3 04102018
            Case 200348
                clsCap.CapExtornoCargoAho nMovNroBus, 230349, sCuenta, sMovNro, sGlosa, nMonto, , , , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsBoleta, lsBoletaITF, nComiRetOtraAge, nComiRetxMaxOpe, gbImpTMU, , , , , , , DatosExtorna 'CTI3 04102018
            Case 200349
                clsCap.CapExtornoCargoAho nMovNroBus, 230350, sCuenta, sMovNro, sGlosa, nMonto, , , , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsBoleta, lsBoletaITF, nComiRetOtraAge, nComiRetxMaxOpe, gbImpTMU, , , , , , , DatosExtorna
            Case 200350
                clsCap.CapExtornoCargoAho nMovNroBus, 230351, sCuenta, sMovNro, sGlosa, nMonto, , , , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsBoleta, lsBoletaITF, nComiRetOtraAge, nComiRetxMaxOpe, gbImpTMU, , , , , , , DatosExtorna 'CTI3 04102018
            Case 200351
                clsCap.CapExtornoCargoAho nMovNroBus, 230352, sCuenta, sMovNro, sGlosa, nMonto, , , , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsBoleta, lsBoletaITF, nComiRetOtraAge, nComiRetxMaxOpe, gbImpTMU, , , , , , , DatosExtorna 'CTI3 04102018
            Case 200352
                clsCap.CapExtornoCargoAho nMovNroBus, 230353, sCuenta, sMovNro, sGlosa, nMonto, , , , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsBoleta, lsBoletaITF, nComiRetOtraAge, nComiRetxMaxOpe, gbImpTMU, , , , , , , DatosExtorna 'CTI3 04102018
           Case 200353
                clsCap.CapExtornoCargoAho nMovNroBus, 230354, sCuenta, sMovNro, sGlosa, nMonto, , , , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsBoleta, lsBoletaITF, nComiRetOtraAge, nComiRetxMaxOpe, gbImpTMU, , , , , , , DatosExtorna 'CTI3 04102018
           Case 200354
                clsCap.CapExtornoCargoAho nMovNroBus, 230355, sCuenta, sMovNro, sGlosa, nMonto, , , , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsBoleta, lsBoletaITF, nComiRetOtraAge, nComiRetxMaxOpe, gbImpTMU, , , , , , , DatosExtorna
           Case 200355
                clsCap.CapExtornoCargoAho nMovNroBus, 230356, sCuenta, sMovNro, sGlosa, nMonto, , , , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsBoleta, lsBoletaITF, nComiRetOtraAge, nComiRetxMaxOpe, gbImpTMU, , , , , , , DatosExtorna
           Case 200356
                clsCap.CapExtornoCargoAho nMovNroBus, 230357, sCuenta, sMovNro, sGlosa, nMonto, , , , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsBoleta, lsBoletaITF, nComiRetOtraAge, nComiRetxMaxOpe, gbImpTMU, , , , , , , DatosExtorna
           Case 200357
                clsCap.CapExtornoCargoAho nMovNroBus, 230358, sCuenta, sMovNro, sGlosa, nMonto, , , , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsBoleta, lsBoletaITF, nComiRetOtraAge, nComiRetxMaxOpe, gbImpTMU, , , , , , , DatosExtorna
            Case gAhoCargoCobroComDiversasAho 'JUEZ 20130906 gAhoExtCargoComDivAho = 230366
                clsCap.CapExtornoCargoAho nMovNroBus, gAhoExtCargoComDivAho, sCuenta, sMovNro, sGlosa, nMonto, , , , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsBoleta, lsBoletaITF, nComiRetOtraAge, nComiRetxMaxOpe, gbImpTMU, , , , , , , DatosExtorna 'CTI3 04102018
            Case gAhoRetComOrdPagDev                'gAhoExtRetComOrdPagDev = 230318
                 clsCap.CapExtornoCargoAho nMovNroBus, gAhoExtRetComOrdPagDev, sCuenta, sMovNro, sGlosa, nMonto, TpoDocOrdenPago, sNroDoc, , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsBoleta, lsBoletaITF, nComiRetOtraAge, nComiRetxMaxOpe, gbImpTMU, , , , , , , DatosExtorna 'CTI3 04102018
            Case gAhoRetAnulChq 'Ahorro Retiro Anulacion Cheque gAhoExtRetAnulChq = 230309
                clsCap.CapExtornoCargoAho nMovNroBus, gAhoExtRetAnulChq, sCuenta, sMovNro, sGlosa, nMonto, TpoDocCheque, sNroDoc, , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsBoleta, lsBoletaITF, nComiRetOtraAge, nComiRetxMaxOpe, gbImpTMU, , , , , , , DatosExtorna
            Case gAhoRetEmiChq  'Ahorro Retiro Emisión Cheque Simple/Gerencia gAhoExtRetEmiChq = 230310
                'clsCap.CapExtornoCargoAho nMovNroBus, gAhoExtRetEmiChq, sCuenta, sMovNro, sGlosa, nMonto, , , , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsBoleta, lsBoletaITF, nComiRetOtraAge, nComiRetxMaxOpe, gbImpTMU, , , , , , , DatosExtorna 'CTI3 04102018 'Comentado CTI4 ERS0112020
                clsCap.CapExtornoCargoAho nMovNroBus, gAhoExtRetEmiChq, sCuenta, sMovNro, sGlosa, nMonto, , , , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsBoleta, lsBoletaITF, nComiRetOtraAge, nComiRetxMaxOpe, gbImpTMU, , , nComiRetSinTarj, , , , DatosExtorna, nComEmiCheque, psComEmiChequeOperacion 'CTI3 04102018 /CTI4 ERS0112020 Add:nComEmiCheque,psComEmiChequeOperacion
            Case gAhoRetEmiChqCanjeOP  'Ahorro Retiro Emisión Cheque Canje OP gAhoExtRetEmiChqCanjeOP = 230311
                clsCap.CapExtornoCargoAho nMovNroBus, gAhoExtRetEmiChqCanjeOP, sCuenta, sMovNro, sGlosa, nMonto, TpoDocOrdenPago, sNroDoc, , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsBoleta, lsBoletaITF, nComiRetOtraAge, nComiRetxMaxOpe, gbImpTMU, , , , , , , DatosExtorna 'CTI3 04102018
            Case gAhoRetOtrosConceptos   'Ahorro Retiro Emisión Cheque Canje OP gAhoExtRetOtrosConceptos = 230330
                clsCap.CapExtornoCargoAho nMovNroBus, gAhoExtRetOtrosConceptos, sCuenta, sMovNro, sGlosa, nMonto, TpoDocOrdenPago, sNroDoc, , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsBoleta, lsBoletaITF, nComiRetOtraAge, nComiRetxMaxOpe, gbImpTMU, , , , , , , DatosExtorna 'CTI3 04102018
            Case gAhoRetChequeDevuelto   'gAhoExtRetChequeDevuelto = 230320
                clsCap.CapExtornoCargoAho nMovNroBus, gAhoExtRetChequeDevuelto, sCuenta, sMovNro, sGlosa, nMonto, TpoDocOrdenPago, sNroDoc, , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsBoleta, lsBoletaITF, nComiRetOtraAge, nComiRetxMaxOpe, gbImpTMU, , , , , , , DatosExtorna 'CTI3 04102018
            Case gAhoRetComTransferencia    'gAhoExtRetComTransferencia= 230323
                clsCap.CapExtornoCargoAho nMovNroBus, gAhoExtRetComTransferencia, sCuenta, sMovNro, sGlosa, nMonto, TpoDocOrdenPago, sNroDoc, , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsBoleta, lsBoletaITF, nComiRetOtraAge, nComiRetxMaxOpe, gbImpTMU, , , , , , , DatosExtorna 'CTI3 04102018
                
            Case gCMACOAAhoRetEfec 'Ahorro Retiro Efectivo Otra CMAC gCMACOAAhoExtRetEfec = 270103
                clsCap.CapExtornoCargoAho nMovNroBus, gCMACOAAhoExtRetEfec, sCuenta, sMovNro, sGlosa, nMonto, , , , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsBoleta, lsBoletaITF, nComiRetOtraAge, nComiRetxMaxOpe, gbImpTMU, , , , , , , DatosExtorna
            Case gCMACOAAhoRetOP 'Ahorro Retiro con Orden de Pago Otra CMAC gCMACOAAhoExtRetOP = 270104
                clsCap.CapExtornoCargoAho nMovNroBus, gCMACOAAhoExtRetOP, sCuenta, sMovNro, sGlosa, nMonto, TpoDocOrdenPago, sNroDoc, , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsBoleta, lsBoletaITF, nComiRetOtraAge, nComiRetxMaxOpe, gbImpTMU, , , , , , , DatosExtorna
            Case gCMACOAAhoRetOPCert 'Ahorro Retiro OP Certificada Otra CMAC  gCMACOAAhoExtRetOPCert = 270105
                clsCap.CapExtornoCargoAhoOPCertificada nMovNroBus, gCMACOAAhoExtRetOPCert, sCuenta, sMovNro, sGlosa, nMonto, sNroDoc, gsNomAge, sLpt, gsCodCMAC, , lsBoleta, gbImpTMU
            Case gAhoRetConsultaSaldos  'Ahorro Retiro Consulta de Saldos  gAhoExtRetConsultaSaldos = 230326
                clsCap.CapExtornoCargoAho nMovNroBus, gAhoExtRetConsultaSaldos, sCuenta, sMovNro, sGlosa, nMonto, , , , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsBoleta, lsBoletaITF, nComiRetOtraAge, nComiRetxMaxOpe, gbImpTMU, , , , , , , DatosExtorna 'CTI3 04102018
            Case gAhoRetPorteCargoCuenta 'Ahorro Porte Cargo Cuentas gAhoExtRetPorteCargoCuenta = 230327
                clsCap.CapExtornoCargoAho nMovNroBus, gAhoExtRetPorteCargoCuenta, sCuenta, sMovNro, sGlosa, nMonto, , , , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsBoleta, lsBoletaITF, nComiRetOtraAge, nComiRetxMaxOpe, gbImpTMU, , , , , , , DatosExtorna 'CTI3 04102018
            Case gAhoRetComVentaBases 'Ahorro Comision Venta de Bases   gAhoExtRetComVentaBases = 230328
                clsCap.CapExtornoCargoAho nMovNroBus, gAhoExtRetComVentaBases, sCuenta, sMovNro, sGlosa, nMonto, , , , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsBoleta, lsBoletaITF, nComiRetOtraAge, nComiRetxMaxOpe, gbImpTMU, , , , , , , DatosExtorna
            Case gAhoRetComServEDELNOR   'gAhoExtRetComServEdelnor = 230324
                clsCap.CapExtornoCargoAho nMovNroBus, gAhoExtRetComServEdelnor, sCuenta, sMovNro, sGlosa, nMonto, , , , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsBoleta, lsBoletaITF, nComiRetOtraAge, nComiRetxMaxOpe, gbImpTMU, , , , , , , DatosExtorna 'CTI3 04102018
            Case gAhoRetFondoFijo    'gAhoExtRetFondoFijo = 230307
                clsCap.CapExtornoCargoAho nMovNroBus, gAhoExtRetFondoFijo, sCuenta, sMovNro, sGlosa, nMonto, , , , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsBoleta, lsBoletaITF, nComiRetOtraAge, nComiRetxMaxOpe, gbImpTMU, , , , , , , DatosExtorna 'CTI3 04102018
            
            Case gAhoDctoEmiExt  'Ahorro Emision de Extracto gAhoExtDctoEmiExt = 230601
                clsCap.CapExtornoCargoAho nMovNroBus, gAhoExtDctoEmiExt, sCuenta, sMovNro, sGlosa, nMonto, , , , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsBoleta, lsBoletaITF, nComiRetOtraAge, nComiRetxMaxOpe, , , , , , , , DatosExtorna 'CTI3 04102018
            Case gAhoDctoEmiOP  'Ahorro Emisión Ordenes de Pago  gAhoExtDctoEmiOP = 230602
                clsCap.CapExtornoCargoAho nMovNroBus, gAhoExtDctoEmiOP, sCuenta, sMovNro, sGlosa, nMonto, , , , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsBoleta, lsBoletaITF, nComiRetOtraAge, nComiRetxMaxOpe, , , , , , , , DatosExtorna 'CTI3 04102018
            '***Agregado por ELRO el 20130401, según TI-ERS011-2013****
            Case gAhoMigracion           'gAhoExtMigracion = 231001
              Dim rsAux As ADODB.Recordset
              Dim i As Integer
              Set rsAux = New ADODB.Recordset
              'CTI3 : 06022018
              rsAux.Fields.Append "nMovNro", adVarChar, 400
              rsAux.Fields.Append "cuenta", adVarChar, 400
              rsAux.Fields.Append "Monto", adDouble, , adFldMayBeNull
              rsAux.Fields.Append "Mov", adVarChar, 400
              rsAux.Open
              'If grdMov.FlxGd.TextMatrix(1, 0) <> "" Then
                For i = 1 To grdMov.Rows - 1
                  rsAux.AddNew
                  rsAux.Fields("nMovNro") = grdMov.TextMatrix(i, 9)
                  rsAux.Fields("cuenta") = grdMov.TextMatrix(i, 3)
                  rsAux.Fields("Monto") = grdMov.TextMatrix(i, 4)
                  rsAux.Fields("Mov") = grdMov.TextMatrix(i, 1)
                Next i
'                clsCap.extornarMigracionCuentasAhorros sMovNro, nMovNroBus, gAhoExtMigracion, sGlosa, grdMov.GetRsNew, gsNomAge, gsCodCMAC, gbImpTMU, sLpt, lsBoleta, DatosExtorna 'CTI3 04102018
                clsCap.extornarMigracionCuentasAhorros sMovNro, nMovNroBus, gAhoExtMigracion, sGlosa, rsAux, gsNomAge, gsCodCMAC, gbImpTMU, sLpt, lsBoleta, DatosExtorna 'CTI3 04102018
                rsAux.Close
                Set rsAux = Nothing
            '***Fin Agregado por ELRO el 20130401, según TI-ERS011-2013
            Case gCTSRetEfec 'CTS Retiro en Efectivo   gCTSExtRetEfec = 250301
                clsCap.CapExtornoCargoCTS nMovNroBus, gCTSExtRetEfec, sCuenta, sMovNro, sGlosa, nMonto, , , gsNomAge, sLpt, gsCodCMAC, , lsBoleta, gbImpTMU, nComiRetSinTarj, , DatosExtorna 'CTI3 04102018
            'RIRO20131212 ERS137
            Case gCTSRetTransf 'CTS Retiro Nota de Cargo   gCTSExtRetTransf = 250302
                clsCap.CapExtornoCargoCTS nMovNroBus, gCTSExtRetTransf, sCuenta, sMovNro, sGlosa, nMonto, TpoDocNotaCargo, sNroDoc, gsNomAge, sLpt, gsCodCMAC, , lsBoleta, gbImpTMU, , nComiTransf, DatosExtorna 'CTI3 04102018
            Case "220303"
                clsCap.CapExtornoCargoCTS nMovNroBus, "250303", sCuenta, sMovNro, sGlosa, nMonto, , , gsNomAge, sLpt, gsCodCMAC, , lsBoleta, gbImpTMU, , , DatosExtorna
            Case gCTSCargoCobroComDiversasAho 'JUEZ 20130906   gCTSExtCargoComDivAho = 250305
                clsCap.CapExtornoCargoCTS nMovNroBus, gCTSExtCargoComDivAho, sCuenta, sMovNro, sGlosa, nMonto, , , gsNomAge, sLpt, gsCodCMAC, , lsBoleta, gbImpTMU, , , DatosExtorna 'CTI3 04102018
            
            Case gCMACOACTSRetEfec 'CTS Retiro en Efectivo Otra CMAC  gCMACOACTSExtRetEfec = 270302
                clsCap.CapExtornoCargoCTS nMovNroBus, gCMACOACTSExtRetEfec, sCuenta, sMovNro, sGlosa, nMonto, , , gsNomAge, sLpt, gsCodCMAC, , lsBoleta, gbImpTMU, , , DatosExtorna 'CTI3 04102018
                
            Case gPFRetInt 'Plazo Fijo Retiro de Intereses    gPFExtRetInt = 240201
                clsCap.CapExtornoRetiroIntPF nMovNroBus, gPFExtRetInt, sCuenta, sMovNro, sGlosa, nMonto, gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsBoleta, lsBoletaITF, gbImpTMU, , DatosExtorna 'CTI3 04102018
            Case gPFRetIntAboAho 'Plazo Fijo Retiro Intereses Abono Cuenta de Ahorros   gPFExtRetIntAboAho = 240202
                clsCap.CapExtornoRetiroIntPF nMovNroBus, gPFExtRetIntAboAho, sCuenta, sMovNro, sGlosa, nMonto, gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsBoleta, lsBoletaITF, gbImpTMU, , DatosExtorna 'CTI3 04102018
            Case gCMACOAPFRetInt 'Plazo Fijo Retiro de Intereses Otra CMAC   gCMACOAPFExtRetInt = 270201
                clsCap.CapExtornoRetiroIntPF nMovNroBus, gCMACOAPFExtRetInt, sCuenta, sMovNro, sGlosa, nMonto, gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsBoleta, lsBoletaITF, gbImpTMU, , DatosExtorna 'CTI3 04102018
            Case gPFRetIntAdelantado '*** PEAC 20091230  gPFExtRetIntCash = 240203
                clsCap.CapExtornoRetiroIntPF nMovNroBus, gPFExtRetIntCash, sCuenta, sMovNro, sGlosa, nMonto, gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsBoleta, lsBoletaITF, gbImpTMU, , DatosExtorna 'CTI3 04102018
            'RIRO20131212 ERS137 ******************
            Case gPFRetIntAboCtaBanco          'gPFExtRetIntAboCtaTransf = 240204
                clsCap.CapExtornoRetiroIntPF nMovNroBus, gPFExtRetIntAboCtaTransf, sCuenta, sMovNro, sGlosa, nMonto, gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsBoleta, lsBoletaITF, gbImpTMU, nComiTransf, DatosExtorna 'CTI3 04102018
            'END RIRO *****************************
            Case gAhoCancAct, gAhoCancInact 'Ahorro Cancelacion  gAhoExtCancAct = 230401
                clsCap.CapExtornoCancelacion nMovNroBus, gAhoExtCancAct, sCuenta, sMovNro, sGlosa, nMonto, , , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsBoleta, lsBoletaITF, gbImpTMU, nComiRetSinTarj, , DatosExtorna 'CTI3 04102018
            Case gAhoCancTransfAct, gAhoCancTransfInact     'gAhoExtCancTransfAct = 230402
                clsCap.CapExtornoCancelacion nMovNroBus, gAhoExtCancTransfAct, sCuenta, sMovNro, sGlosa, nMonto, , , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsBoleta, lsBoletaITF, gbImpTMU, , , DatosExtorna 'CTI3 04102018
            Case gPFCancEfec 'Plazo Fijo Cancelacion  gPFExtCancEfec = 240301
                clsCap.CapExtornoCancelacion nMovNroBus, gPFExtCancEfec, sCuenta, sMovNro, sGlosa, nMonto, , , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsBoleta, lsBoletaITF, gbImpTMU, , , DatosExtorna 'CTI3 04102018
            'RIRO20131212 ERS137
            Case gPFCancTransf  'Plazo Fijo Cancelacion  gPFExtCancTransfAbBco = 240304
                clsCap.CapExtornoCancelacion nMovNroBus, gPFExtCancTransfAbBco, sCuenta, sMovNro, sGlosa, nMonto, , , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsBoleta, lsBoletaITF, gbImpTMU, , nComiTransf, DatosExtorna 'CTI3 04102018
            
            Case gCTSCancEfec 'CTS Cancelación            gCTSExtCancEfec = 250401
                clsCap.CapExtornoCancelacion nMovNroBus, gCTSExtCancEfec, sCuenta, sMovNro, sGlosa, nMonto, , , gsNomAge, sLpt, gsCodCMAC, , , , , lsBoleta, lsBoletaITF, gbImpTMU, nComiRetSinTarj, , DatosExtorna
            Case gCTSCancTransf                         'gCTSExtCancTransf = 250402
                clsCap.CapExtornoCancelacion nMovNroBus, gCTSExtCancTransf, sCuenta, sMovNro, sGlosa, nMonto, , , gsNomAge, sLpt, gsCodCMAC, , , , , lsBoleta, lsBoletaITF, gbImpTMU, , , DatosExtorna
            'RIRO20131212 ERS137
            Case gCTSCancTransfBco                      'gCTSExtCancTransfAbCta = 250403
                clsCap.CapExtornoCancelacion nMovNroBus, gCTSExtCancTransfAbCta, sCuenta, sMovNro, sGlosa, nMonto, , , gsNomAge, sLpt, gsCodCMAC, , , , , lsBoleta, lsBoletaITF, gbImpTMU, , nComiTransf, DatosExtorna
            'RIRO20131212 ERS137
            Case gAhoCancTransfAbCtaBco                 'gAhoExtCanctransf = 230411
                clsCap.CapExtornoCancelacion nMovNroBus, gAhoExtCanctransf, sCuenta, sMovNro, sGlosa, nMonto, , , gsNomAge, sLpt, gsCodCMAC, , , , , lsBoleta, lsBoletaITF, gbImpTMU, , nComiTransf, DatosExtorna 'CTI3 04102018
            
            Case gServGiroCancEfec 'Giro Cancelación          gServExtGiroCancEfec = 360201
                clsCap.ServGiroExtornoCancelacion nMovNroBus, gServExtGiroCancEfec, sCuenta, sMovNro, sGlosa, nMonto, gsNomAge, sLpt, gsCodCMAC, lsBoleta, gbImpTMU, DatosExtorna 'CTI3 04102018
                
            Case gAhoTransCargo
                'ALPA 20091119****************
                'clsCap.CapExtornoTransfwerenciaAho nMovNroBus, sMovNro, sGlosa, gsNomAge, sLpt, gsCodCMAC, lsmensaje, txtCuenta.NroCuenta, lsBoleta, lsBoletaITF
                clsCap.CapExtornoTransferenciaAho nMovNroBus, sMovNro, sGlosa, gsNomAge, sLpt, gsCodCMAC, lsmensaje, txtCuenta.NroCuenta, lsBoleta, lsBoletaITF, lsBoletaAbono, DatosExtorna
                '*****************************
                'If Trim(lsMensaje) <> "" Then 'Comentado x JUEZ 20131209
                '    MsgBox lsMensaje, vbInformation, "Aviso"
                'End If
            'Add By GITU 07-11-2012
            'Case gAhoTransAbonoL Or gAhoTransCargoL RIRO20150907 Comentado
            Case gAhoTransAbonoL, gAhoTransCargoL 'RIRO20150907 ADD
                clsCap.CapExtornoTransferenciaAho nMovNroBus, sMovNro, sGlosa, gsNomAge, sLpt, gsCodCMAC, lsmensaje, txtCuenta.NroCuenta, lsBoleta, lsBoletaITF, lsBoletaAbono, DatosExtorna
            'End GITU

            'Extorno de Operaciones con CMACs LLamada
            Case gCMACOTAhoDepEfec 'Depósito Efectivo                   'gCMACOTAhoExtDepEfec = 270501
                clsCap.CapExtornoOpeAhoCMACLlamada nMovNroBus, sMovNro, gCMACOTAhoExtDepEfec, sGlosa, sCuenta, sDescOpe, nMonto, , gsNomAge, sLpt, gsCodCMAC, , lsBoleta, gbImpTMU
             
            Case gCMACOTAhoDepChq 'Depósito Cheque                     'gCMACOTAhoExtDepChq = 270502
                clsCap.CapExtornoOpeAhoCMACLlamada nMovNroBus, sMovNro, gCMACOTAhoExtDepChq, sGlosa, sCuenta, sDescOpe, nMonto, sNroDoc, gsNomAge, sLpt, gsCodCMAC, , lsBoleta, gbImpTMU
                
            Case gCMACOTAhoRetEfec 'Retiro Efectivo                 'gCMACOTAhoExtRetEfec = 270503
                clsCap.CapExtornoOpeAhoCMACLlamada nMovNroBus, sMovNro, gCMACOTAhoExtRetEfec, sGlosa, sCuenta, sDescOpe, nMonto, , gsNomAge, sLpt, gsCodCMAC, , lsBoleta, gbImpTMU
                
            Case gCMACOTAhoRetOP 'Retiro Orden Pago                    'gCMACOTAhoExtRetOP = 270504
                clsCap.CapExtornoOpeAhoCMACLlamada nMovNroBus, sMovNro, gCMACOTAhoExtRetOP, sGlosa, sCuenta, sDescOpe, nMonto, sNroDoc, gsNomAge, sLpt, gsCodCMAC, , lsBoleta, gbImpTMU
                
            Case "107001" 'Pago de Creditos
                clsCap.CapExtornoOpeAhoCMACLlamada nMovNroBus, sMovNro, "137000", sGlosa, sCuenta, sDescOpe, nMonto, sNroDoc, gsNomAge, sLpt, gsCodCMAC, , lsBoleta, gbImpTMU
            
            'Extornos de Aumento / Disminucion de Capital
            Case gPFAumCapEfec                                  'gPFExtAumCapEfec = 240801
                clsCap.CapExtornoCapAumDisPF sCuenta, nMovNroBus, gPFExtAumCapEfec, sMovNro, nMonto, nITFOperacion, nMontoITF, gsNomAge, gsCodCMAC, gsCodAge, , , , lsBoleta, gbImpTMU, DatosExtorna 'CTI3 04102018
            Case gPFAumCapTasaPactEfec                          'gPFExtAumCapTasaPactEfec = 240803
                clsCap.CapExtornoCapAumDisPF sCuenta, nMovNroBus, gPFExtAumCapTasaPactEfec, sMovNro, nMonto, nITFOperacion, nMontoITF, gsNomAge, gsCodCMAC, gsCodAge, , , , lsBoleta, gbImpTMU, DatosExtorna 'CTI3 04102018
            Case gPFAumCapchq                                   'gPFExtAumCapchq = 240802
                clsCap.CapExtornoCapAumDisPF sCuenta, nMovNroBus, gPFExtAumCapchq, sMovNro, nMonto, nITFOperacion, nMontoITF, gsNomAge, gsCodCMAC, gsCodAge, TpoDocCheque, sNroDoc, sPersCod, lsBoleta, gbImpTMU, DatosExtorna 'CTI3 04102018
            Case gPFAumCapTasaPactChq                           'gPFExtAumCapTasaPactChq = 240804
                clsCap.CapExtornoCapAumDisPF sCuenta, nMovNroBus, gPFExtAumCapTasaPactChq, sMovNro, nMonto, nITFOperacion, nMontoITF, gsNomAge, gsCodCMAC, gsCodAge, TpoDocCheque, sNroDoc, sPersCod, lsBoleta, gbImpTMU, DatosExtorna 'CTI3 04102018
            Case gPFAumCapTrans
                '***Agregado por ELRO el 20121221, según OYP-RFC024-2012
                nConfirmarConciliacion = clsCap.verificarConciliacionBancos(nMovNroBus)
                If nConfirmarConciliacion >= 1 Then
                   MsgBox "No puede extornar el movimiento porque ya fue conciliado por el área de Finanzas.", vbInformation, "!Aviso¡"
                   Exit Sub
                End If
                '***Fin Agregado por ELRO el 20121221*******************gPFExtAumCapTrans = 240807
                clsCap.CapExtornoCapAumDisPF sCuenta, nMovNroBus, gPFExtAumCapTrans, sMovNro, nMonto, nITFOperacion, nMontoITF, gsNomAge, gsCodCMAC, gsCodAge, , , , lsBoleta, gbImpTMU, DatosExtorna 'CTI3 04102018
            Case gPFAumCapTasaPactTrans                         'gPFExtAumCapTasaPactTrans = 240808
                clsCap.CapExtornoCapAumDisPF sCuenta, nMovNroBus, gPFExtAumCapTasaPactTrans, sMovNro, nMonto, nITFOperacion, nMontoITF, gsNomAge, gsCodCMAC, gsCodAge, , , , lsBoleta, gbImpTMU, DatosExtorna 'CTI3 04102018
            Case gPFDismCapEfec                                 'gPFExtDismCapEfec = 240805
                clsCap.CapExtornoCapAumDisPF sCuenta, nMovNroBus, gPFExtDismCapEfec, sMovNro, nMonto, nITFOperacion, nMontoITF, gsNomAge, gsCodCMAC, gsCodAge, , , , lsBoleta, gbImpTMU, DatosExtorna 'CTI3 04102018
            Case gPFAumCapCargoCta 'JUEZ 20131226               'gPFExtAumCapCargoCta = 240809
                clsCap.CapExtornoCapAumDisPF sCuenta, nMovNroBus, gPFExtAumCapCargoCta, sMovNro, nMonto, nITFOperacion, nMontoITF, gsNomAge, gsCodCMAC, gsCodAge, nTipoDoc, sNroDoc, , lsBoleta, gbImpTMU, DatosExtorna 'CTI3 04102018
            Case gAhoDepositoHaberesEnLoteEfec  'RIRO 20140530 ERS017 gAhoExtDepositoHaberesEnLoteEfec = 230265
                
                'CTI2 20190405 ADD ****
                Set oDCOMCaptaMovimiento = New COMDCaptaGenerales.DCOMCaptaMovimiento
                bRespuesta = False
                bRespuesta = oDCOMCaptaMovimiento.ValidaExtornoLote(nMovNroBus)
                If Not bRespuesta Then
                    MsgBox "No es posible realizar el extorno " & vbNewLine _
                           & "Hay cuentas que no cuentan con saldo para realizar el extorno", vbInformation, "Aviso - Extorno Depósito en Lote"
                    Exit Sub
                End If
                'CTI2 20190405 END ****
                Dim nResultado As Integer
                nResultado = 0
                nResultado = clsCap.CapExtornoAbonoLote(nMovNroBus, gAhoExtDepositoHaberesEnLoteEfec, sMovNro, sGlosa, gsNomAge, sLpt, gsCodCMAC, gsCodAge, lsmensaje, lsBoleta, gbImpTMU, DatosExtorna) 'CTI3 04102018
                If nResultado = -1 Then
                    MsgBox "Una de las cuentas a las que se aplicará el extorno, no cuenta con saldo suficiente, los cambios serán revertidos", vbInformation, "Aviso - Extorno Depósito en Lote"
                    Exit Sub
                End If
                cmdCancelar_Click
                
            Case gAhoDepositoHaberesEnLoteTransf 'RIRO 20140530 ERS017 gAhoExtDepositoHaberesEnLoteTransf = 230266
                
                'CTI2 20190405 ADD ****
                Set oDCOMCaptaMovimiento = New COMDCaptaGenerales.DCOMCaptaMovimiento
                bRespuesta = False
                bRespuesta = oDCOMCaptaMovimiento.ValidaExtornoLote(nMovNroBus)
                If Not bRespuesta Then
                    MsgBox "No es posible realizar el extorno " & vbNewLine _
                           & "Hay cuentas que no cuentan con saldo para realizar el extorno", vbInformation, "Aviso - Extorno Depósito en Lote"
                    Exit Sub
                End If
                'CTI2 20190405 END ****
                nResultado = 0
                nResultado = clsCap.CapExtornoAbonoLote(nMovNroBus, gAhoExtDepositoHaberesEnLoteTransf, sMovNro, sGlosa, gsNomAge, sLpt, gsCodCMAC, gsCodAge, lsmensaje, lsBoleta, gbImpTMU, DatosExtorna)  'CTI3 04102018
                If nResultado = -1 Then
                    MsgBox "Una de las cuentas a las que se aplicará el extorno, no cuenta con saldo suficiente, los cambios serán revertidos", vbInformation, "Aviso - Extorno Depósito en Lote"
                    Exit Sub
                End If
                cmdCancelar_Click
            Case gAhoDepositoHaberesEnLoteChq 'CTI6 20210606 gAhoExtDepositoHaberesEnLoteChq = 230267
                Set oDCOMCaptaMovimiento = New COMDCaptaGenerales.DCOMCaptaMovimiento
                bRespuesta = False
                bRespuesta = oDCOMCaptaMovimiento.ValidaExtornoLote(nMovNroBus)
                If Not bRespuesta Then
                    MsgBox "No es posible realizar el extorno " & vbNewLine _
                           & "Hay cuentas que no cuentan con saldo para realizar el extorno", vbInformation, "Aviso - Extorno Depósito Cheque"
                    Exit Sub
                End If
                
                nResultado = 0
                nResultado = clsCap.CapExtornoAbonoLote(nMovNroBus, gAhoExtDepositoHaberesEnLoteChq, sMovNro, sGlosa, gsNomAge, sLpt, gsCodCMAC, gsCodAge, lsmensaje, lsBoleta, gbImpTMU, DatosExtorna)
                If nResultado = -1 Then
                    MsgBox "Una de las cuentas a las que se aplicará el extorno, no cuenta con saldo suficiente, los cambios serán revertidos", vbInformation, "Aviso - Extorno Depósito en Lote"
                    Exit Sub
                End If
                cmdCancelar_Click
            'RECO20160209 ERS073-2015******************************
            'SEPELIO
            Case "200380" 'CARGO POR AFILIACION
                Dim lsNumCertif As String
                Set oDCOMCaptaMovimiento = New COMDCaptaGenerales.DCOMCaptaMovimiento
                If Mid(sCuenta, 6, 3) = "234" Then
                    clsCap.CapExtornoCargoCTS nMovNroBus, 290005, sCuenta, sMovNro, sGlosa, nMonto, , , gsNomAge, sLpt, gsCodCMAC, , lsBoleta, gbImpTMU, nComiRetSinTarj, , DatosExtorna 'CTI3 04102018
                Else
                    clsCap.CapExtornoCargoAho nMovNroBus, 290005, sCuenta, sMovNro, sGlosa, nMonto, , , , gsNomAge, sLpt, gsCodCMAC, nMontoITF, nITFOperacion, nITFConcepto, , lsBoleta, lsBoletaITF, nComiRetOtraAge, nComiRetxMaxOpe, gbImpTMU, , , nComiRetSinTarj, , , lnMovNro, DatosExtorna 'CTI3 04102018
                End If
                lsNumCertif = oSegSep.ActualizaEstadoSeguroSepelio("", "", nMovNroBus, 503) 'APRI20171020 ERS028-2017  4 -> 503
                Call oDCOMCaptaMovimiento.AgregaSegSepelioAfiliacionhis(lsNumCertif, "", gdFecSis, sMovNro, lnMovNro, grdMov.TextMatrix(grdMov.row, 10), gsCodAge, 503) 'APRI20171020 ERS028-2017  4 -> 503
            'RECO FIN**********************************************
        End Select

         
        'JUEZ 20131209 ****************************************
        If Trim(lsmensaje) <> "" Then
            MsgBox lsmensaje, vbInformation, "Aviso"
            Exit Sub
        End If
        'END JUEZ *********************************************
         
        '*** PEAC 20081001
            
        loVistoElectronico.RegistraVistoElectronico (nMovNroBus)
        
        '*** FIN PEAC
        
        Select Case nOperacion 'JUEZ 20130906
         Case gAhoExtCargoComDivAho, gCTSExtCargoComDivAho
            Dim oCredBol As COMNCredito.NCOMCredDoc
            Set oCredBol = New COMNCredito.NCOMCredDoc
                lsBoleta = oCredBol.ImprimeBoletaComision("EXTORNO COMISION", Left("Pago comision", 36), "", Str(nMonto), "", "", "________" & Mid(sCuenta, 9, 1), False, "", "", , gdFecSis, gsNomAge, gsCodUser, sLpt, , gbImpTMU)
            Set oCredBol = Nothing
        End Select
        
        'CTI4 ERS0112020
            Dim clsMovx As New COMDMov.DCOMMov, sCodUserBus As String
            sCodUserBus = Right(clsMovx.GetcMovNro(nMovNroBus), 4)
        'end CTI4
        
        If nOperacion = gCTSDepLotCargoCta Then 'CTI4 ERS0112020
            Dim oNCOMCaptaImpresion As New COMNCaptaGenerales.NCOMCaptaImpresion
            lsImpCargoCta = oNCOMCaptaImpresion.nPrintReciboExtorCargoCta(gsNomAge, lsFechaHoraGrab, "", lsCtaAhoExt, 0, PstaNombre(lsClienteAhoExt), _
            "", lnMontoAhoExt, 0, nMovNroBus, gsCodUser, "", "", sCodUserBus, gImpresora, gbImpTMU)
            Set oNCOMCaptaImpresion = Nothing
        End If
        
        If Trim(lsBoleta) <> "" Then
            nFicSal = FreeFile
            Open sLpt For Output As nFicSal
                Print #nFicSal, lsBoleta & lsImpCargoCta 'CTI4 ERS0112020
                Print #nFicSal, ""
            Close #nFicSal
        End If
    
        If Trim(lsBoletaITF) <> "" Then
            nFicSal = FreeFile
            Open sLpt For Output As nFicSal
                Print #nFicSal, lsBoletaITF
                Print #nFicSal, ""
            Close #nFicSal
        End If
        'ALPA 20091119************************************
        If Trim(lsBoletaAbono) <> "" Then
        MsgBox "Acomodar el papel de impresión", vbApplicationModal
            nFicSal = FreeFile
            Open sLpt For Output As nFicSal
                Print #nFicSal, lsBoletaAbono
                Print #nFicSal, ""
            Close #nFicSal
        End If
        '*************************************************
        If nOperacion = gServGiroApertCargoCta Then
            If Trim(psBoletaExtorno) <> "" Then
            MsgBox "Acomodar el papel de impresión", vbApplicationModal
                nFicSal = FreeFile
                Open sLpt For Output As nFicSal
                    Print #nFicSal, psBoletaExtorno
                    Print #nFicSal, ""
                Close #nFicSal
            End If
        End If
        '**************************************************
        
        
        If nOperacion = gAhoApeLoteChq Or nOperacion = gAhoApeLoteEfec _
            Or nOperacion = gPFApeLoteChq Or nOperacion = gPFApeLoteEfec _
            Or nOperacion = gCTSApeLoteChq Or nOperacion = gCTSApeLoteEfec _
            Or nOperacion = gAhoTransAbonoL Or nOperacion = gAhoTransCargoL Then
            cmdCancelar_Click
            cmdBuscar_Click
        Else
            grdMov.EliminaFila grdMov.row
        End If
    End If
    txtGlosa.Text = "" 'JUEZ 20131209
    Call ExtDatosOculta 'CTI3
    Set oCap = Nothing
    Set pActualiza = Nothing
    Set clsCap = Nothing
    Set clsMov = Nothing
Else
    MsgBox "No existen datos para realizar el Extorno", vbInformation, "Aviso"
End If
    
End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 If optTipoBus(1).value = True Then
    If KeyCode = vbKeyF12 And txtCuenta.Enabled = True Then 'F12
        Dim sCuenta As String
        sCuenta = frmValTarCodAnt.Inicia(nProducto, False)
        If sCuenta <> "" Then
            txtCuenta.NroCuenta = sCuenta
            txtCuenta.SetFocusCuenta
        End If
    End If
 End If
End Sub
'******CTI3 (ferimoro) 18102018
Private Sub CargaControles()
Dim oCons As COMDConstantes.DCOMConstantes
Dim R As ADODB.Recordset

Set oCons = New COMDConstantes.DCOMConstantes

Set R = oCons.ObtenerConstanteExtornoMotivo

Set oCons = Nothing
Call Llenar_Combo_MotivoExtorno(R, cmbMotivos) 'CTI3

End Sub
Private Sub Form_Load()
Me.Icon = LoadPicture(App.Path & gsRutaIcono)
End Sub

Private Sub grdMov_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Me.cmdExtornar.Enabled = True
  Me.cmdExtornar.SetFocus
  'txtGlosa.SetFocus
End If
End Sub

Private Sub optTipoBus_Click(Index As Integer)
lblNroMov.Visible = False
txtMovNro.Visible = False
lblMov.Visible = False
txtCuenta.Visible = False
'Add By Gitu 06-11-2012
lblCodMov.Visible = False
txtCodMov.Visible = False
'End Gitu
Select Case Index
    Case 0
        lblMov.Visible = True
        lblNroMov.Visible = True
        txtMovNro.Visible = True
        lblMov = Format$(gdFecSis, "yyyymmdd")
    Case 1
       If nOperacion <> "107001" Then
            txtCuenta.Visible = True
            If nOperacion = gCMACOTAhoRetEfec Or gCMACOTAhoDepEfec Or gCMACOTAhoDepChq Or gCMACOTAhoRetOP Then
                txtCuenta.CMAC = ""
                txtCuenta.Prod = ""
                txtCuenta.Age = ""
                txtCuenta.Cuenta = ""
                txtCuenta.EnabledCMAC = True
                txtCuenta.EnabledProd = True
            Else
               txtCuenta.CMAC = gsCodCMAC
               txtCuenta.Prod = ""
               txtCuenta.Age = ""
               txtCuenta.Cuenta = ""
               txtCuenta.EnabledCMAC = False
               txtCuenta.EnabledProd = False
            End If
        Else
            fraNroCredito.Visible = True
            txtCredito.Text = ""
            fraNroCredito.Enabled = True
        End If
    Case 2
        lblCodMov.Visible = True
        txtCodMov.Visible = True
End Select
End Sub

Private Sub optTipoBus_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    If Index = 0 Or Index = 2 Then
        txtMovNro.SetFocus
    ElseIf Index = 1 Then
        txtCuenta.SetFocus
    End If
End If
End Sub

Private Sub txtCodMov_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmdBuscar.SetFocus
    End If
End Sub

Private Sub txtCuenta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    CmdBuscar.SetFocus
    Exit Sub
End If
End Sub

Private Sub txtDetExtorno_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then KeyAscii = 0: Exit Sub
'If KeyAscii = 13 Then SendKeys "{TAB}": Exit Sub
End Sub

Private Sub txtGlosa_KeyPress(KeyAscii As Integer)
KeyAscii = fgIntfMayusculas(KeyAscii)
If KeyAscii = 13 Then
    cmdExtornar.SetFocus
End If
End Sub

Private Sub txtMovNro_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    CmdBuscar.SetFocus
    Exit Sub
End If
KeyAscii = Asc(UCase(Chr$(KeyAscii)))
End Sub

'RIRO20131212 ERS137
Private Function VerificarEstadoPendiente() As Boolean
    
    Dim bValor As Boolean
    Dim oDocRec As New NDocRec
    Dim rs As ADODB.Recordset
    bValor = True
    Select Case lsCodExtOpc
        Case "230303", "230411", "240204", "240304", "250306", "250406"
             Set rs = oDocRec.getPendientesTransf("", Mid(Trim(grdMov.TextMatrix(grdMov.row, 3)), 9, 1), CDbl(grdMov.TextMatrix(grdMov.row, 9)))
             If rs Is Nothing Then
                bValor = False
             Else
                bValor = True
             End If
             Set rs = Nothing
             Set oDocRec = Nothing
    End Select
    VerificarEstadoPendiente = bValor
End Function
'FIN RIRO

