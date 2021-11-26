VERSION 5.00
Begin VB.Form frmCredNewNivAutorizaResolver 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Autorizaciones requeridas del crédito"
   ClientHeight    =   10080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10395
   Icon            =   "frmCredNewNivAutorizaResolver.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10080
   ScaleWidth      =   10395
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSolMod 
      Caption         =   "Sol. Modif."
      Height          =   345
      Left            =   3060
      TabIndex        =   20
      ToolTipText     =   "Salir"
      Top             =   9600
      Width           =   1485
   End
   Begin VB.CommandButton cmdRechazar 
      Caption         =   "Rechazar"
      Height          =   345
      Left            =   1590
      TabIndex        =   19
      ToolTipText     =   "Salir"
      Top             =   9600
      Width           =   1485
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "Grabar"
      Height          =   345
      Left            =   120
      TabIndex        =   18
      ToolTipText     =   "Salir"
      Top             =   9600
      Width           =   1485
   End
   Begin VB.Frame Frame4 
      Caption         =   " Glosa [Sol. Modif.]"
      ForeColor       =   &H00FF0000&
      Height          =   1335
      Left            =   120
      TabIndex        =   16
      Top             =   8160
      Width           =   10095
      Begin VB.TextBox txtGlosa 
         Height          =   975
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   17
         Top             =   240
         Width           =   9855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " Autorizaciones "
      ForeColor       =   &H00FF0000&
      Height          =   3975
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   10095
      Begin SICMACT.FlexEdit FEAutorizaciones 
         Height          =   3615
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   6376
         Cols0           =   8
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Autorización-Saldo-Requiere-Detalle-Autorizar-cCodigo-cNivAprCod"
         EncabezadosAnchos=   "300-7050-800-800-0-730-0-0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-4-5-X-X"
         ListaControles  =   "0-0-0-0-1-4-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-C-C-C-C-C-C"
         FormatosEdit    =   "0-0-0-0-0-0-0-0"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   6
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   300
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Autorizaciones no contempladas"
      ForeColor       =   &H00FF0000&
      Height          =   2295
      Left            =   120
      TabIndex        =   4
      Top             =   5880
      Width           =   10095
      Begin SICMACT.FlexEdit FEExoneraciones 
         Height          =   1935
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   3413
         Cols0           =   6
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Exoneración-Descripción Exoneración-Requiere-Autorizar-nItem"
         EncabezadosAnchos=   "300-3200-4600-800-800-0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-4-X"
         ListaControles  =   "0-0-0-0-4-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-C-C-C"
         FormatosEdit    =   "0-0-0-0-0-0"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   300
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   " Datos del Crédito "
      ForeColor       =   &H00FF0000&
      Height          =   1815
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   10095
      Begin VB.CommandButton cmdListaAutorizacion 
         Caption         =   "Listado de Autorizaciones"
         Height          =   345
         Left            =   6480
         TabIndex        =   26
         ToolTipText     =   "Salir"
         Top             =   1080
         Width           =   2685
      End
      Begin VB.TextBox txtAgencia 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   285
         Left            =   1200
         TabIndex        =   23
         Top             =   1440
         Width           =   2055
      End
      Begin VB.TextBox txtTpoProducto 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   285
         Left            =   1200
         TabIndex        =   21
         Top             =   1080
         Width           =   2055
      End
      Begin SICMACT.ActXCodCta_New ActxCta 
         Height          =   735
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   1296
         Texto           =   " Crédito "
         EnabledCMAC     =   -1  'True
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
      Begin VB.TextBox txtNombre 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   285
         Left            =   4320
         TabIndex        =   5
         Top             =   180
         Width           =   4935
      End
      Begin VB.Label lblPeriodo 
         Caption         =   "Periodo:"
         Height          =   255
         Left            =   6720
         TabIndex        =   27
         Top             =   520
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "Detalle de Autorizaciones:"
         Height          =   255
         Left            =   4440
         TabIndex        =   25
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label6 
         Caption         =   "Agencia:"
         Height          =   255
         Left            =   480
         TabIndex        =   24
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Tipo Producto:"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lblTasa 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   285
         Left            =   8520
         TabIndex        =   12
         Top             =   500
         Width           =   735
      End
      Begin VB.Label lblCuotas 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   285
         Left            =   7440
         TabIndex        =   11
         Top             =   500
         Width           =   495
      End
      Begin VB.Label lblMoneda 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   285
         Left            =   4320
         TabIndex        =   10
         Top             =   500
         Width           =   465
      End
      Begin VB.Label lblMontoSol 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   285
         Left            =   4770
         TabIndex        =   9
         Top             =   500
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Monto :"
         Height          =   255
         Left            =   3600
         TabIndex        =   8
         Top             =   525
         Width           =   735
      End
      Begin VB.Label lblNroCuota 
         Caption         =   "Nº Cuotas:"
         Height          =   255
         Left            =   6600
         TabIndex        =   7
         Top             =   520
         Width           =   975
      End
      Begin VB.Label lblTasas 
         Caption         =   "Tasa:"
         Height          =   255
         Left            =   8040
         TabIndex        =   6
         Top             =   520
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente:"
         Height          =   255
         Left            =   3600
         TabIndex        =   3
         Top             =   210
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   345
      Left            =   8640
      TabIndex        =   0
      ToolTipText     =   "Salir"
      Top             =   9600
      Width           =   1485
   End
End
Attribute VB_Name = "frmCredNewNivAutorizaResolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************************************************************
'** Nombre : frmCredNewNivAutorizaVer
'** Descripción : Formulario para resolver las autorizaciones de créditos pendientes según ERS002-2016
'** Creación : RECO, 20160625 09:17:00 AM
'************************************************************************************************
Option Explicit
Dim fsNroCta As String 'FRHU 20160820
Public Sub inicia(ByVal psCtaCod As String)
    ActxCta.NroCuenta = psCtaCod
    fsNroCta = psCtaCod 'FRHU 20160820
    ActxCta.EnabledCMAC = False 'FRHU 20160820
    ActxCta.EnabledAge = False 'FRHU 20160820
    ActxCta.EnabledProd = False 'FRHU 20160820
    ActxCta.EnabledCta = False 'FRHU 20160820
    Call CargarDatos
    Me.Show 1
End Sub

Private Sub CargarDatos()
    Dim oDNiv As New COMDCredito.DCOMNivelAprobacion
    Dim oCred As New COMDCredito.DCOMCredito
    Dim oRs As New ADODB.Recordset, oRSCred As New ADODB.Recordset
    Dim nIndice As Integer, nCantidad As Integer
    
    
    'Set oRSCred = oCred.RecuperaDatosCredBasicos(ActxCta.NroCuenta)
    Set oRSCred = oCred.RecuperaDatosCredBasicos(fsNroCta) 'FRHU 20160820
    If Not (oRSCred.EOF And oRSCred.BOF) Then
        txtNombre.Text = oRSCred!cPersNombre
        LblMontoSol.Caption = Format(oRSCred!nMonto, gsFormatoNumeroView)
        lblmoneda.Caption = oRSCred!nmoneda
        lblCuotas.Caption = oRSCred!nCuotas
        lblTasa.Caption = Format(oRSCred!nTasaInteres, gsFormatoNumeroView)
        Me.txtAgencia.Text = oRSCred!CAgencia           'ARLO20170613 ERS0652016
        Me.txtTpoProducto.Text = oRSCred!cTpoCredCod    'ARLO20170613 ERS0652016
        
        '***** LUCV20171212, Agregó según observación SBS
        If (oRSCred!cTpoProdCod) = "514" Then
            lblPeriodo.Visible = True
            lblTasas.Visible = False
            lblTasa.Visible = False
        Else
            lblPeriodo.Visible = False
            lblTasas.Visible = True
            lblTasa.Visible = True
        End If
        '**** Fin LUCV20171212
    End If
    
    'Set oRS = oDNiv.ObtieneResultadoAutoExoNiv(ActxCta.NroCuenta, gsCodUser)
    Set oRs = oDNiv.ObtieneResultadoAutoExoNiv(fsNroCta, gsCodUser, gsCodAge) 'ARLO20170613 ERS0652016 se agrego variable : gsCodAge
    If Not (oRs.EOF And oRs.BOF) Then
        nCantidad = oRs.RecordCount
        FEAutorizaciones.Clear
        FormateaFlex FEAutorizaciones
        FEExoneraciones.Clear
        FormateaFlex FEExoneraciones
        For nIndice = 1 To nCantidad
            If oRs!cTipoReg = 2 Then Exit For
            FEAutorizaciones.AdicionaFila
            FEAutorizaciones.TextMatrix(nIndice, 1) = oRs!cDescripcion
            FEAutorizaciones.TextMatrix(nIndice, 2) = oRs!nTotal                        'ARLO20170613 ERS0652016
            FEAutorizaciones.TextMatrix(nIndice, 3) = IIf(oRs!nEsNivel = 1, "SI", "NO")
            FEAutorizaciones.TextMatrix(nIndice, 4) = "VER"
            FEAutorizaciones.TextMatrix(nIndice, 5) = IIf(oRs!nEstado = 1, "1", "")
            FEAutorizaciones.TextMatrix(nIndice, 6) = oRs!cCodigo
            FEAutorizaciones.TextMatrix(nIndice, 7) = oRs!cNivApreCod
            oRs.MoveNext
        Next
        nCantidad = nCantidad - (nIndice - 1)
        For nIndice = 1 To nCantidad
            FEExoneraciones.AdicionaFila
            FEExoneraciones.TextMatrix(nIndice, 1) = oRs!cCodigo
            FEExoneraciones.TextMatrix(nIndice, 2) = oRs!cDescripcion
            FEExoneraciones.TextMatrix(nIndice, 3) = IIf(oRs!nEsNivel = 1, "SI", "NO")
            FEExoneraciones.TextMatrix(nIndice, 4) = IIf(oRs!nEstado = 1, "1", "")
            FEExoneraciones.TextMatrix(nIndice, 5) = oRs!nItem
            oRs.MoveNext
        Next
    End If
    
End Sub

Private Sub cmdRechazar_Click()
    If MsgBox("¿Está seguro que desea rechazar el crédito?", vbQuestion + vbYesNo, "Alerta") = vbYes Then
        'Call frmCredRechazo.Rechazar(4, ActxCta.NroCuenta)
        Call frmCredRechazo.Rechazar(4, fsNroCta) 'FRHU 20160820
        Call LimpiarFomulario
        Unload Me
    End If
End Sub

Private Sub cmdSolMod_Click()
    Dim oNNiv As New COMNCredito.NCOMNivelAprobacion
    Dim sMsj As String, sMovNro As String
    
    sMsj = ValidaDatos
    If sMsj = "" Then
        If MsgBox("¿Está seguro que desea modificar la sugerencia?", vbQuestion + vbYesNo, "Alerta") = vbYes Then
            If sMsj = "" Then
                sMovNro = GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
                'Call oNNiv.dSolicitaModifAprobacionResultado(ActxCta.NroCuenta, Trim(txtGlosa.Text), gTipoNivelAuto, sMovNro)
                Call oNNiv.dSolicitaModifAprobacionResultado(fsNroCta, Trim(txtGlosa.Text), gTipoNivelAuto, sMovNro) 'FRHU 20160820
                Call LimpiarFomulario
                Unload Me
            Else
                MsgBox sMsj, vbInformation, "Alerta"
            End If
        End If
    Else
        MsgBox sMsj, vbInformation, "Alerta"
    End If
End Sub

Private Sub FEAutorizaciones_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
    If FEAutorizaciones.TextMatrix(pnRow, 2) = "NO" Then
        MsgBox "Nivel no corresponde", vbInformation, "Alerta"
        FEAutorizaciones.TextMatrix(pnRow, 4) = IIf(FEAutorizaciones.TextMatrix(pnRow, 4) = "", 1, "")
    End If
End Sub

Private Sub FEExoneraciones_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
    If FEExoneraciones.TextMatrix(pnRow, 3) = "NO" Then
        MsgBox "Nivel no corresponde", vbInformation, "Alerta"
        FEExoneraciones.TextMatrix(pnRow, 4) = IIf(FEExoneraciones.TextMatrix(pnRow, 4) = "", 1, "")
    End If
End Sub

Private Sub cmdGrabar_Click()
    Dim oCreNiv As New COMDCredito.DCOMNivelAprobacion
    Dim nIndice As Integer
    Dim sMsj As String, sMovNro As String
    sMsj = ValidaDatos
    sMovNro = GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Dim rs As ADODB.Recordset 'ARLO20170714
    If sMsj = "" Then
        If MsgBox("¿Está seguro que desea grabar?", vbQuestion + vbYesNo, "Alerta") = vbYes Then
            'ARLO 20170615 ERS0652016 INICIO
            For nIndice = 1 To FEAutorizaciones.rows - 1
                
                If FEAutorizaciones.TextMatrix(nIndice, 3) = "SI" Then
                Set rs = oCreNiv.ValidarSiExiteAutorizacion(fsNroCta, FEAutorizaciones.TextMatrix(nIndice, 6))

                        If (rs.RecordCount > 0) Then
                                If (rs!nSaldo = 0) Then
                                MsgBox "No le quedan mas Autorizaciones para el Tipo de Credito,Por favor Comunicarse con el Area de Riesgos", vbInformation, "Aviso"
                                Exit Sub
                                End If
                        End If
                End If
            Next
            'ARLO 20170615 ERS0652016 FIN
            For nIndice = 1 To FEAutorizaciones.rows - 1
                If FEAutorizaciones.TextMatrix(nIndice, 3) = "SI" Then 'ARLO 20170613 ERS0652016 CAMBIO DE LA COLUMNA 2 A 3
                    'Call oCreNiv.ActualizaResultadoAutoExo(ActxCta.NroCuenta, FEAutorizaciones.TextMatrix(nIndice, 5), FEAutorizaciones.TextMatrix(nIndice, 6), 0, sMovNro, _
                                                       IIf(FEAutorizaciones.TextMatrix(nIndice, 4) = "", 0, 1), txtGlosa.Text, 1)
                    Call oCreNiv.ActualizaResultadoAutoExo(fsNroCta, FEAutorizaciones.TextMatrix(nIndice, 6), FEAutorizaciones.TextMatrix(nIndice, 7), 0, sMovNro, _
                                                       IIf(FEAutorizaciones.TextMatrix(nIndice, 5) = "", 0, 1), txtGlosa.Text, 1) 'FRHU 20160820 'ARLO 20170613 ERS0652016 CAMBIO LAS POSICIONES DE LA COLUMNAS
                                                    
                End If
            Next
            For nIndice = 1 To FEExoneraciones.rows - 1
                If FEExoneraciones.TextMatrix(nIndice, 3) = "SI" Then
                    'Call oCreNiv.ActualizaResultadoAutoExo(ActxCta.NroCuenta, "", "", feExoneraciones.TextMatrix(nIndice, 5), sMovNro, IIf(feExoneraciones.TextMatrix(nIndice, 4) = "", 0, 1), "", 2)
                    Call oCreNiv.ActualizaResultadoAutoExo(fsNroCta, "", "", FEExoneraciones.TextMatrix(nIndice, 5), sMovNro, IIf(FEExoneraciones.TextMatrix(nIndice, 4) = "", 0, 1), "", 2) 'FRHU 20160820
                End If
            Next
            Call LimpiarFomulario
            Unload Me
        End If
    Else
        MsgBox sMsj, vbInformation, "Alerta"
    End If
End Sub

Private Function ValidaDatos() As String
    ValidaDatos = ""
    If txtGlosa.Text = "" Then
        ValidaDatos = "Debe registrar una glosa"
    End If
End Function

Private Sub LimpiarFomulario()
    FEAutorizaciones.Clear
    FEExoneraciones.Clear
    FormateaFlex FEAutorizaciones
    FormateaFlex FEExoneraciones
    txtGlosa.Text = ""
    ActxCta.NroCuenta = ""
End Sub

Private Sub cmdSalir_Click()
    Call LimpiarFomulario
    Unload Me
End Sub
'ARLO 20170613 ERS0652016 INICIO
Private Sub cmdListaAutorizacion_Click()
    frmCredRiesgosAutorizacionListado.InicioLectura
End Sub
'ARLO 20170613 ERS0652016 FIN
