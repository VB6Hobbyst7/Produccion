VERSION 5.00
Begin VB.Form frmCFEmision 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Carta Fianza - Emisión"
   ClientHeight    =   6210
   ClientLeft      =   2370
   ClientTop       =   405
   ClientWidth     =   7395
   Icon            =   "frmCFEmision.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   7395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cmbTpRiesgo 
      Height          =   315
      ItemData        =   "frmCFEmision.frx":030A
      Left            =   4320
      List            =   "frmCFEmision.frx":030C
      Style           =   2  'Dropdown List
      TabIndex        =   37
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Avalado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   555
      Left            =   240
      TabIndex        =   31
      Top             =   1800
      Width           =   7050
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Avalado"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   34
         Top             =   240
         Width           =   585
      End
      Begin VB.Label lblCodAvalado 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1140
         TabIndex        =   33
         Tag             =   "txtcodigo"
         Top             =   180
         Width           =   1305
      End
      Begin VB.Label lblNomAval 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2400
         TabIndex        =   32
         Tag             =   "txtnombre"
         Top             =   180
         Width           =   4470
      End
   End
   Begin VB.TextBox txtNumPoliza 
      Enabled         =   0   'False
      Height          =   390
      Left            =   1080
      TabIndex        =   30
      Top             =   4995
      Width           =   2115
   End
   Begin VB.Frame Frame2 
      Caption         =   "Acreedor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   555
      Left            =   180
      TabIndex        =   19
      Top             =   1140
      Width           =   7050
      Begin VB.Label lblNomAcreedor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2400
         TabIndex        =   22
         Tag             =   "txtnombre"
         Top             =   180
         Width           =   4470
      End
      Begin VB.Label lblCodAcreedor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1140
         TabIndex        =   21
         Tag             =   "txtcodigo"
         Top             =   180
         Width           =   1305
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Acreedor "
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   690
      End
   End
   Begin VB.CommandButton cmdExaminar 
      Caption         =   "E&xaminar..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5820
      TabIndex        =   18
      ToolTipText     =   "Buscar Credito"
      Top             =   120
      Width           =   1215
   End
   Begin VB.Frame Frame6 
      Height          =   690
      Left            =   120
      TabIndex        =   12
      Top             =   5400
      Width           =   7035
      Begin VB.CommandButton cmdGenerarPDF 
         Caption         =   "Vista Previa"
         Height          =   390
         Left            =   4200
         TabIndex        =   35
         Top             =   195
         Width           =   1215
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   390
         Left            =   5520
         TabIndex        =   17
         Top             =   195
         Width           =   1185
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar"
         Enabled         =   0   'False
         Height          =   390
         Left            =   225
         TabIndex        =   16
         ToolTipText     =   "Grabar Datos de Aprobacion de Credito"
         Top             =   195
         Width           =   1185
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   390
         Left            =   1620
         TabIndex        =   15
         ToolTipText     =   "Ir al Menu Principal"
         Top             =   195
         Width           =   1185
      End
   End
   Begin VB.Frame fraDatos 
      Caption         =   "Carta Fianza"
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
      ForeColor       =   &H8000000D&
      Height          =   2475
      Left            =   180
      TabIndex        =   5
      Top             =   2400
      Width           =   7065
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Monto"
         Height          =   195
         Index           =   6
         Left            =   4440
         TabIndex        =   28
         Top             =   240
         Width           =   450
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Vencimiento"
         Height          =   195
         Index           =   5
         Left            =   4380
         TabIndex        =   27
         Top             =   660
         Width           =   870
      End
      Begin VB.Label lblMontoApr 
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
         Left            =   5700
         TabIndex        =   26
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblFecVencApr 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   5700
         TabIndex        =   25
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Finalidad"
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   23
         Top             =   1380
         Width           =   630
      End
      Begin VB.Label lblAnalista 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   5700
         TabIndex        =   14
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Analista"
         Height          =   195
         Index           =   1
         Left            =   4440
         TabIndex        =   13
         Top             =   1020
         Width           =   555
      End
      Begin VB.Label lblApoderado 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   960
         TabIndex        =   0
         Top             =   960
         Width           =   2775
      End
      Begin VB.Label lblModalidad 
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
         Left            =   960
         TabIndex        =   11
         Top             =   600
         Width           =   2775
      End
      Begin VB.Label lblFinalidad 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   240
         TabIndex        =   10
         Top             =   1680
         Width           =   6735
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Apoderado"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   9
         Top             =   1020
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Modalidad"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   660
         Width           =   735
      End
      Begin VB.Label lblTipoCF 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   960
         TabIndex        =   7
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tipo "
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   6
         Top             =   300
         Width           =   360
      End
   End
   Begin VB.Frame fraCliente 
      Caption         =   "Afianzado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   600
      Left            =   180
      TabIndex        =   1
      Top             =   540
      Width           =   7050
      Begin VB.Label lblNomcli 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2400
         TabIndex        =   4
         Top             =   180
         Width           =   4485
      End
      Begin VB.Label lblCodcli 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1140
         TabIndex        =   3
         Top             =   180
         Width           =   1290
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   225
         Width           =   480
      End
   End
   Begin SICMACT.ActXCodCta ActXCodCta 
      Height          =   390
      Left            =   180
      TabIndex        =   24
      Top             =   120
      Width           =   3645
      _ExtentX        =   6429
      _ExtentY        =   688
   End
   Begin VB.Label Label5 
      Caption         =   "Tipo Riesgo:"
      Height          =   255
      Left            =   3360
      TabIndex        =   36
      Top             =   5040
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Num. Folio:"
      Height          =   315
      Left            =   225
      TabIndex        =   29
      Top             =   5040
      Width           =   990
   End
End
Attribute VB_Name = "frmCFEmision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'*  APLICACION : Carta Fianza
'*  ARCHIVO : frmCFEmision
'*  CREACION: 10/09/2002     - LAYG
'*************************************************************************
'*  RESUMEN: PERMITE EMITIR LA CARTA FIANZA
'*************************************************************************
Option Explicit
Dim vCodCta As String
Dim fpComision As Double
Dim fsEstado As String
Dim fnRenovacion As Integer
'WIOR 20120619 *****************
Dim fsCodEnvio As String
Dim fbRemesado As Boolean
'WIOR FIN **********************
Dim objPista As COMManejador.Pista
Dim ldFechaAsi As Date 'FRHU20131126

'*  VALIDACION DE DATOS DEL FORMULARIO ANTES DE GRABAR
Function ValidaDatos() As Boolean
Dim oValida As COMDCartaFianza.DCOMCartaFianza
Dim nMovNro As Long

    ValidaDatos = True
    Set oValida = New COMDCartaFianza.DCOMCartaFianza
    nMovNro = oValida.VerificaComision(ActXCodCta.NroCuenta)
    If nMovNro < 0 Then
        MsgBox "Aun no se ha efectuado el Pago de la Comision, no puede emitir la Carta Fianza", vbInformation, "Aviso"
        ValidaDatos = False
        Exit Function
    End If
    Set oValida = Nothing
    
    'ARCV 04-07-2007
    If fsEstado <> gColocEstRenovada Then   'Solo para las que no han sido renovadas
        If txtNumPoliza.Text = "" Then
            MsgBox "Ingrese el numero de Poliza", vbInformation, "Mensaje"
            ValidaDatos = False
        End If
    End If
    '------
    
End Function

'****************************************************************
'*  LIMPIA LOS DATOS DE LA PANTALLA PARA UNA NUEVA APROBACION
'****************************************************************
Sub LimpiaDatos()
    ActXCodCta.Enabled = True
    ActXCodCta.NroCuenta = fgIniciaAxCuentaCF
    lblNomcli.Caption = ""
    lblCodcli.Caption = ""
    lblNomcli.Caption = ""
    lblCodAcreedor.Caption = ""
    lblNomAcreedor.Caption = ""
    lblTipoCF.Caption = ""
    lblFinalidad.Caption = ""
'    lblFinalidad.Text = ""
    lblModalidad.Caption = ""
    lblMontoApr.Caption = ""
    lblFecVencApr.Caption = "__/__/____"
    lblAnalista.Caption = ""
    lblApoderado.Caption = ""
    cmdGrabar.Enabled = False
    fraDatos.Enabled = False
    txtNumPoliza.Text = "" 'ARCV 05-07-2007
    If lblCodAvalado.Visible Then
        lblCodAvalado.Caption = ""
        lblNomAval.Caption = ""
    End If
    cmdGenerarPDF.Enabled = False 'WIOR 20120613
    fbRemesado = False 'WIOR 20120619
    
    cmbTpRiesgo.Enabled = False 'JOEP20180622 Acta 122-2018
    cmbTpRiesgo.Clear  'JOEP20180622 Acta 122-2018
End Sub

'marg 07-06-2016------
Private Sub get_EstadoExpediente(ByVal psCta As String, ByRef Estado As Integer)

Dim oCF As COMDCartaFianza.DCOMCartaFianza 'DCartaFianza
Dim c As New ADODB.Recordset
Dim d As New ADODB.Recordset
Dim ubicacion As String
Dim observacion As String
'Dim estado As Integer
Dim count As Integer

    Set oCF = New COMDCartaFianza.DCOMCartaFianza
    Set c = oCF.get_CredAdmControlDesembolso(psCta)
    
    Set d = oCF.get_ControlCreditosObsAdmCred(psCta)
        If Not d.BOF And Not d.EOF Then
        For count = 1 To d.RecordCount
            If (d!nRegulariza = 0) Then
                 observacion = observacion & d!cDescripcion & vbCrLf
            End If
            d.MoveNext
        Next
    End If
    
    If Not c.BOF And Not c.EOF Then
       If (IsNull(c!dIngreso) And IsNull(c!dUltSalidaObs) And IsNull(c!dUltIngresoObs) And IsNull(c!dSalida)) Then
           ubicacion = "El Expediente aun se encuentra en Comité de Créditos y está pendiente de revisión por la Administración de Créditos"
           Estado = 0
       End If
       If (Not (IsNull(c!dIngreso)) And IsNull(c!dUltSalidaObs) And IsNull(c!dUltIngresoObs) And IsNull(c!dSalida)) Then
           ubicacion = "El Expediente ingresó al area de Administración de Créditos el " & Format(c!dIngreso, "dd/mm/yyyy") & IIf(observacion <> "", " y tiene las siguientes observaciones:", " para su respectiva observación")
           Estado = 1
       End If
       If (Not (IsNull(c!dIngreso)) And Not (IsNull(c!dUltSalidaObs)) And IsNull(c!dUltIngresoObs) And IsNull(c!dSalida)) Then
           ubicacion = "El Expediente salió del area de Administración de Créditos el " & Format(c!dIngreso, "dd/mm/yyyy") & " por las siguientes observaciónes:"
           Estado = 2
       End If
        If (Not (IsNull(c!dIngreso)) And (IsNull(c!dUltSalidaObs)) And Not (IsNull(c!dUltIngresoObs)) And IsNull(c!dSalida)) Then
           ubicacion = "El Expediente reigresó al area de Administración de Créditos el " & Format(c!dIngreso, "dd/mm/yyyy") & " para su respectivo levantamiento de observaciones:"
           Estado = 3
       End If
       If (Not (IsNull(c!dSalida))) Then
           ubicacion = "El Expediente salió del area de Administración de Créditos el " & Format(c!dIngreso, "dd/mm/yyyy") & " para su respectivo desembolso"
           Estado = 4
       End If
    Else
        ubicacion = "El Expediente aun se encuentra en Comité de Créditos y está pendiente de revisión por la Administración de Créditos"
        Estado = 0
    End If
    
    If (Estado <> 4) Then
        MsgBox ubicacion & vbCrLf & observacion, vbInformation, "Aviso"
    End If
End Sub
'</marg-------------------

'PROCEDIMIENTO QUE CARGA LOS DATOS QUE SE REQUIEREN PARA EL FORMULARIO
Private Sub CargaDatos(ByVal psCta As String)
    Dim oCF As COMDCartaFianza.DCOMCartaFianza 'DCartaFianza
    Dim R As New ADODB.Recordset
    Dim loCFCalculo As COMNCartaFianza.NCOMCartaFianzaCalculos 'NCartaFianzaCalculos
    Dim loConstante As COMDConstantes.DCOMConstantes 'DConstante
    Dim loCFValida As COMNCartaFianza.NCOMCartaFianzaValida 'NCartaFianzaValida
    Dim lbTienePermiso As Boolean
    Dim lnComisionPagada As Double
    Dim lnComisionCalculada As Double
    'Dim ldFechaAsi As Date 'FRHU20131126
    Dim rsCartaFianza As ADODB.Recordset 'WIOR 20120619
    Dim nNumEnvios As Long 'WIOR 20120619
    ActXCodCta.Enabled = False
    
    'JOEP20180622 Acta 122-2018
    Dim oComboTpRiesgo As COMDCartaFianza.DCOMCartaFianza
    Dim rsComboTpRiesgo As ADODB.Recordset
    Set oComboTpRiesgo = New COMDCartaFianza.DCOMCartaFianza
    Set rsComboTpRiesgo = oComboTpRiesgo.get_ObtieneDatosConstantexCod(20000)
    CargarComboBox rsComboTpRiesgo, cmbTpRiesgo
    Set oComboTpRiesgo = Nothing
    RSClose rsComboTpRiesgo
    'JOEP20180622 Acta 122-2018
    
    '***** LUCV20171212, Agregó según observación SBS *****
    Dim oDCOMCredito As New COMdCredito.DCOMCredito
     If oDCOMCredito.verificarExisteAutorizaciones(psCta) Then
        MsgBox "El crédito tiene una autorización/exoneración pendiente", vbInformation, "Alerta"
        Call frmCredNewNivAutorizaVer.Consultar(psCta)
        Exit Sub
    End If
    Set oDCOMCredito = Nothing
    '***** Fin LUCV20171212 *****
    
    Set oCF = New COMDCartaFianza.DCOMCartaFianza
    Set R = oCF.RecuperaCartaFianzaEmision(psCta)
    Set oCF = Nothing
    If Not R.BOF And Not R.EOF Then
        lblCodcli.Caption = R!cperscod
        lblNomcli.Caption = PstaNombre(R!cPersNombre)
        lblCodAcreedor.Caption = R!cPersAcreedor
        lblNomAcreedor.Caption = PstaNombre(R!cPersNomAcre)
        ldFechaAsi = R!dAsignacion
        
        'MAVM 20100606
        'If Mid(Trim(psCta), 6, 1) = "1" Then
            lblTipoCF = IIf(IsNull(R!cConsDescripcion), "", R!cConsDescripcion) 'IIf(Mid(Trim(psCta), 9, 1) = "1", "COMERCIALES - SOLES", "COMERCIALES - DOLARES")
        'ElseIf Mid(Trim(psCta), 6, 1) = "2" Then
            'lblTipoCF = IIf(Mid(Trim(psCta), 9, 1) = "1", "MICROEMPRESA - SOLES", "MICROEMPRESA - DOLARES")
        'End If
        lblAnalista.Caption = IIf(IsNull(R!cAnalista), "", R!cAnalista)
        lblApoderado.Caption = IIf(IsNull(R!cApoderado), "", R!cApoderado)
        
        'MADM 20111020
        lblCodAvalado.Caption = IIf(IsNull(R!cAvalCod), "", R!cAvalCod)
        If (R!cAvalNombre) <> "" Then
            Me.lblNomAval.Caption = PstaNombre(R!cAvalNombre)
        End If
        'END MADM
        
        lblFinalidad.Caption = IIf(IsNull(R!cfinalidad), "", R!cfinalidad)
        lblMontoApr = IIf(IsNull(R!nMontoApr), "", Format(R!nMontoApr, "#0.00"))
        lblFecVencApr = IIf(IsNull(R!dVencApr), "", Format(R!dVencApr, "dd/mm/yyyy"))
        fsEstado = R!nPrdEstado
        fnRenovacion = IIf(IsNull(R!nRenovacion), 0, R!nRenovacion)
        'By Capi Acta 035-2007
        If fsEstado <> gColocEstRenovada Then
            txtNumPoliza.Enabled = True
        End If
        
        Set loConstante = New COMDConstantes.DCOMConstantes
            'lblModalidad = loConstante.DameDescripcionConstante(gColCFModalidad, R!nModalidad)'comento JOEP20181222 CP
            If R!nModalidad = 13 Then
                lblModalidad = R!OtrsModalidades
            Else
                lblModalidad = loConstante.DameDescripcionConstante(gColCFModalidad, R!nModalidad)
            End If
        Set loConstante = Nothing
        'Verifica Fecha Vcto Aprobada es posterior a la Fecha Actual
        If Format(lblFecVencApr, "yyyy/mm/dd") < Format(gdFecSis, "yyyy/mm/dd") Then
            MsgBox "Fecha de Vencimiento de Carta Fianza es anterior a la Fecha Actual", vbInformation, "Aviso"
            LimpiaDatos
            Exit Sub
        End If
        
        Set loCFValida = New COMNCartaFianza.NCOMCartaFianzaValida
            'By capi 24022009 se modifico para enviar la fecha de proceso
            'lnComisionPagada = loCFValida.nCFPagoComision(psCta)
            lnComisionPagada = loCFValida.nCFPagoComision(psCta, gdFecSis)
            '
        Set loCFValida = Nothing
        Set loCFCalculo = New COMNCartaFianza.NCOMCartaFianzaCalculos
            lnComisionCalculada = Format(loCFCalculo.nCalculaComisionTrimestralCF(R!nMontoApr, DateDiff("d", ldFechaAsi, R!dVencApr), fpComision, Mid(psCta, 9, 1)), "####0.00")
        Set loCFCalculo = Nothing

        '*** PEAC 20090813
        'If lnComisionPagada < lnComisionCalculada Then
        If lnComisionPagada <= 0 Then
            MsgBox "No se ha pagado comision de Carta Fianza", vbInformation, "Aviso"
            LimpiaDatos
            Exit Sub
        End If
                
        'marg
        Dim Estado As Integer
        Call get_EstadoExpediente(psCta, Estado)
        If (Estado <> 4) Then
            LimpiaDatos
            Exit Sub
        End If
        
        '</marg
                
        'txtMontoApr.Text = IIf(IsNull(R!nMontoSug), "", Format(R!nMontoSug, "#0.00"))
        'TxtFecVenApr.Text = IIf(IsNull(R!dVencSug), "", Format(R!dVencSug, "dd/mm/yyyy"))
    
        fraDatos.Enabled = True
        cmdGrabar.Enabled = True
        cmdGenerarPDF.Enabled = True 'WIOR 20120613
        'JOEP20180622 Acta 122-2018
        cmbTpRiesgo.Enabled = True
        'JOEP20180622 Acta 122-2018
         'WIOR 20120619 ******************************
        Set loCFValida = New COMNCartaFianza.NCOMCartaFianzaValida
        Set rsCartaFianza = loCFValida.ExisteRemeseas(gsCodAge)
        If rsCartaFianza.RecordCount > 0 Then
        nNumEnvios = CLng(rsCartaFianza!nCodEnvio)
            If nNumEnvios > 0 Then
                Set rsCartaFianza = loCFValida.ObtenerEnvioFolios("1", gsCodAge)
                If rsCartaFianza.RecordCount > 0 Then
                    If fsEstado <> gColocEstRenovada Then
                        Set rsCartaFianza = loCFValida.ObtenerNumFolioAEmitir(gsCodAge, 0)
                        If rsCartaFianza.RecordCount > 0 Then
                            If Trim(rsCartaFianza!nNumFolio) <> "0" Then
                                Me.txtNumPoliza.Text = Format(rsCartaFianza!nNumFolio, "0000000")
                                Me.txtNumPoliza.Enabled = False
                                fbRemesado = True
                            Else
                                MsgBox "No cuenta con Folios Númerados para Cartas Fianza.", vbInformation, "Aviso"
                                LimpiaDatos
                                Exit Sub
                            End If
                        End If
                    Else
                        Set rsCartaFianza = loCFValida.ObtenerNumFolioAEmitir(gsCodAge, 1, True)
                        If rsCartaFianza.RecordCount <= 0 Then
                            MsgBox "No cuenta con Folios sin numeración para Renovacion de Cartas Fianza.", vbInformation, "Aviso"
                            LimpiaDatos
                            Exit Sub
                        Else
                            fsCodEnvio = rsCartaFianza!nCodEnvio
                            fbRemesado = True
                        End If
                    End If
                    Set loCFValida = Nothing
                    Set rsCartaFianza = Nothing
                Else
                    
                    Set rsCartaFianza = loCFValida.ObtenerEnvioFolios("1,2", gsCodAge)
                    If rsCartaFianza.RecordCount > 0 Then
                        MsgBox "Ud. no puede emitir la Carta Fianza, no cuenta con nuevos folios asignados.", vbInformation, "Aviso"
                        LimpiaDatos
                        Exit Sub
                    End If
                    
                    If nNumEnvios > 2 Then
                        MsgBox "Ud. no puede emitir la Carta Fianza, no cuenta con nuevos folios asignados.", vbInformation, "Aviso"
                        LimpiaDatos
                        Exit Sub
                    End If
                    Set rsCartaFianza = Nothing
                End If
            Else
                Me.txtNumPoliza.Enabled = True
            End If
        End If
        'WIOR FIN ***********************************
        'WIOR 20121114 ******************************
        If fsEstado = gColocEstRenovada Then
            txtNumPoliza.Enabled = False
        Else
            txtNumPoliza.Enabled = True
        End If
        'WIOR FIN ***********************************
    End If

End Sub

Private Sub ActXCodCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(Trim(ActXCodCta.NroCuenta)) > 0 Then
            Call CargaDatos(ActXCodCta.NroCuenta)
        Else
            Call LimpiaDatos
        End If
    End If
End Sub

Private Sub cmdCancelar_Click()
    LimpiaDatos
End Sub

Private Sub cmdExaminar_Click()
Dim lsCta As String
    'MAVM 20100605 BAS II
    lsCta = frmCFPersEstado.Inicio(Array(gColocEstAprob, gColocEstAprob, gColocEstRenovada), "Emision de Carta Fianza", Array(gColCFComercial, gColCFPYME, gColCFTpoProducto))
    If Len(Trim(lsCta)) > 0 Then
        ActXCodCta.NroCuenta = lsCta
        Call CargaDatos(lsCta)
    Else
        Call LimpiaDatos
    End If
End Sub

Private Sub cmdGrabar_Click()
Dim loNCartaFianza As COMNCartaFianza.NCOMCartaFianza 'NCartaFianza
Dim loImprime As COMNCartaFianza.NCOMCartaFianzaImpre 'NCartaFianzaImpre
Dim loCFExt As COMDCartaFianza.DCOMCartaFianza  'WIOR 20130319
Dim loPrevio As previo.clsprevio
Dim lsCadImprimir As String

Dim loContFunct As COMNContabilidad.NCOMContFunciones 'NContFunciones
Dim lsMovNro As String
Dim lsFechaHoraGrab As String

Dim lnMontoEmi As Currency
Dim ldVencEmi As Date
Dim ldAsigEmi As Date 'FRHU20131126
'APRI20170208
    Dim TipoCredito As ADODB.Recordset
    Dim nTipoCredito As Integer
    Dim nTotalGarantiaPF As Integer

'END APRI

vCodCta = ActXCodCta.NroCuenta
lnMontoEmi = Format(lblMontoApr, "#0.00")
ldVencEmi = Format(lblFecVencApr, "dd/mm/yyyy")
ldAsigEmi = Format(ldFechaAsi, "dd/mm/yyyy") 'FRHU20131126

If ValidaDatos = False Then
    Exit Sub
End If

If MsgBox("Desea Guardar Emision de Carta Fianza", vbInformation + vbYesNo, "Aviso") = vbYes Then

    'Genera el Mov Nro
    Set loContFunct = New COMNContabilidad.NCOMContFunciones
        lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Set loContFunct = Nothing
    lsFechaHoraGrab = fgFechaHoraGrab(lsMovNro)
    
        Set loNCartaFianza = New COMNCartaFianza.NCOMCartaFianza
        Set loCFExt = New COMDCartaFianza.DCOMCartaFianza 'WIOR 20130319
        Call loNCartaFianza.nCFEmision(vCodCta, lsFechaHoraGrab, lsMovNro, ldVencEmi, lnMontoEmi, ldAsigEmi, , (cmbTpRiesgo.ItemData(cmbTpRiesgo.ListIndex))) 'FRHU20131126 AGREGO: ldAsigEmi 'Agrego (cmbTpRiesgo.ItemData(cmbTpRiesgo.ListIndex)) 'JOEP20180622 Acta 122-2018
        
        
        'WIOR 20130319 **************************************************
        If fnRenovacion > 0 Then
            Call loCFExt.OperacionesCFRestaura(3, vCodCta, fnRenovacion - 1)
        End If
        'WIOR FIN ********************************************************
         
        Set objPista = New COMManejador.Pista
        objPista.InsertarPista gsOpeCod, lsMovNro, gsCodPersUser, GetMaquinaUsuario, gInsertar, "Emision de CF", vCodCta, gCodigoCuenta
        Set objPista = Nothing
        Set loNCartaFianza = Nothing
        Set loCFExt = Nothing 'WIOR 20130319
        
    Dim loImp As COMNCartaFianza.NCOMCartaFianzaReporte
    
    If Len(vCodCta) = 18 Then
'            Set loImp = New COMNCartaFianza.NCOMCartaFianzaReporte
'            loImp.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis
'            lsCadImprimir = loImp.nImprimeCartaFianza(vCodCta)
'            Set loImp = Nothing
      If fsEstado <> gColocEstRenovada Then
         Call ImpreDoc(vCodCta)
      Else
         Call ImpreDocRenovado(vCodCta)
      End If
    End If
    Call ImprimePagare
      'APRI 20170208
    Dim obj As COMNCredito.NCOMCredito
    
     Set obj = New COMNCredito.NCOMCredito
     Set TipoCredito = obj.ObtenerTipoCredito(vCodCta)
     Set obj = Nothing
    
    Do While Not TipoCredito.EOF
        nTipoCredito = TipoCredito!cTpoProdCod
        nTotalGarantiaPF = TipoCredito!nTotalGarantiaPF
        TipoCredito.MoveNext
    Loop
    
    If nTipoCredito = 514 And nTotalGarantiaPF > 0 And fsEstado <> gColocEstRenovada Then 'LUCV20171212, Agregó según observación SBS
        ImprimeCartaAfectacion vCodCta, nTipoCredito, lnMontoEmi, CLng(txtNumPoliza.Text)
    End If
    'END APRI
    cmdGrabar.Enabled = False
    cmdGenerarPDF.Enabled = False 'WIOR 20120613
    LimpiaDatos
End If
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub ActxCodCta_keypressEnter()
    vCodCta = ActXCodCta.NroCuenta
    If Len(vCodCta) > 0 Then
        Call CargaDatos(vCodCta)
        ActXCodCta.Enabled = False
    Else
        Call LimpiaDatos
    End If
End Sub

Private Sub Form_Load()
Dim loParam As COMDColocPig.DCOMColPCalculos 'DColPCalculos
Set loParam = New COMDColocPig.DCOMColPCalculos
    fpComision = loParam.dObtieneColocParametro(4001)
Set loParam = Nothing
Me.Icon = LoadPicture(App.Path & gsRutaIcono)
LimpiaDatos
gsOpeCod = gCredEmisionCF
fbRemesado = False 'WIOR 20120619

End Sub

Private Function UbiAgencia() As String

    Dim lszona As String
    Dim lscUbiGeoCod As String
    
    Dim lRstZona As ADODB.Recordset
    Dim OlZona  As COMDConstantes.DCOMZonas
    Set OlZona = New COMDConstantes.DCOMZonas

    Dim lRstAgencia As ADODB.Recordset
    Dim OlAgencia  As COMDConstantes.DCOMAgencias
    Set OlAgencia = New COMDConstantes.DCOMAgencias
    
    
    Set lRstAgencia = OlAgencia.RecuperaAgencias(gsCodAge)
        lscUbiGeoCod = lRstAgencia("cUbiGeoCod")
    Set lRstAgencia = Nothing
    Set lRstZona = OlZona.DameUnaZona(lscUbiGeoCod)
        lszona = Trim(lRstZona("cUbiGeoDescripcion"))
        UbiAgencia = lszona
    Set lRstZona = Nothing
    
End Function


Private Function DirAgencia() As String


    Dim lscAgeDireccion As String
    
    Dim lRstAgencia As ADODB.Recordset
    Dim OlAgencia  As COMDConstantes.DCOMAgencias
    Set OlAgencia = New COMDConstantes.DCOMAgencias
    
    
    Set lRstAgencia = OlAgencia.RecuperaAgencias(gsCodAge)
        lscAgeDireccion = lRstAgencia("cAgeDireccion")
        DirAgencia = lscAgeDireccion
    Set lRstAgencia = Nothing
    
End Function

Sub ImpreDoc(ByVal psCtaCod As String)
    Dim loRs As COMNCartaFianza.NCOMCartaFianzaValida 'NCartaFianzaValida
    
    Dim lrDataT As ADODB.Recordset
    Dim lrDataCF As ADODB.Recordset
    Dim lrDataCR As ADODB.Recordset
    Dim rsCartaFianza As ADODB.Recordset 'WIOR 20120619
    Dim nDias As Integer
    Dim dfechaini As Date
    Dim dfechafin As Date
    Dim lsFechas As String
    
    Dim nPoliza As Long
    Dim cDirecAgencia As String
    Dim oWord As Word.Application
    Dim oDoc As Word.Document
    Dim oRange As Word.Range
    Dim nCFPoliza As Long 'WIOR 20120407
    Set loRs = New COMNCartaFianza.NCOMCartaFianzaValida
    
    Set lrDataCF = loRs.RecuperaDatosGeneralesCF(psCtaCod)
    Set lrDataT = loRs.RecuperaDatosT(psCtaCod)
    Set lrDataCR = loRs.RecuperaDatosAcreedor(psCtaCod)
    
    'ARCV 04-07-2007
    'nPoliza = CLng(loRs.GetCF_Poliza(psCtaCod))
    cDirecAgencia = loRs.Get_Agencia_CF(psCtaCod)
    'If nPoliza = 0 Then
    '    nPoliza = CLng(loRs.GrabaCF_Poliza(psCtaCod, gdFecSis))
    'End If
    'Call loRs.GrabaCF_Poliza_NEW(psCtaCod, gdFecSis, CLng(txtNumPoliza.Text))
    nCFPoliza = loRs.GrabaCF_Poliza_NEW(psCtaCod, gdFecSis, CLng(txtNumPoliza.Text)) 'WIOR 20120407
    'Call frmCFImpresion.Inicio(vCodCta, IIf(Me.lblCodAvalado.Caption = "", False, True), nCFPoliza, 1) 'WIOR 20120427
    'WIOR 20120619 ********************************
    If fbRemesado Then
        Set rsCartaFianza = loRs.ObtenerEnvioFolios("1", gsCodAge)
        If rsCartaFianza.RecordCount > 0 Then
            Set rsCartaFianza = loRs.UltimoRegistroEnvio(, , CLng(txtNumPoliza.Text))
            If rsCartaFianza.RecordCount > 0 Then
                Call loRs.ActualizarEnvioFolios(Trim(rsCartaFianza!nCodEnvio), 2)
            End If
            Call loRs.ActualizarFolio(CLng(txtNumPoliza.Text), 1, psCtaCod, gdFecSis)
        End If
    End If
    'WIOR FIN **************************************
    Call ImprimirPDF(vCodCta, IIf(Me.lblCodAvalado.Caption = "", False, True), nCFPoliza, 2)
    Set loRs = Nothing
    
'    Set oWord = CreateObject("Word.Application")
'    'MADM 20110121 ---------------------------------------------------------------
'    oWord.Visible = True
'    'oWord.Visible = False
'    'MADM 20111020 20110121 ---------------------------------------------------------------
'    If lblCodAvalado.Caption = "" Then
'        Set oDoc = oWord.Documents.Open(App.path & "\FormatoCarta\CFMaynas.doc")
'    Else
'        Set oDoc = oWord.Documents.Open(App.path & "\FormatoCarta\CFMaynasGar.doc")
'    End If
'
'    'By Capi Acta 035-2007
'
'    Dim loAge As COMDConstantes.DCOMAgencias
'    Dim rs1 As ADODB.Recordset
'    Dim lsAgencia As String
'    Dim lsAgenciaDir As String
'    Dim lnPosicion As Integer
'
'    Set loAge = New COMDConstantes.DCOMAgencias
'    Set rs1 = New ADODB.Recordset
'        Set rs1 = loAge.RecuperaAgencias(gsCodAge)
'        If Not (rs1.EOF And rs1.BOF) Then
'            lsAgencia = Trim(rs1("cUbiGeoDescripcion"))
'            'By Capi Acta 035-2007
'            lnPosicion = InStr(lsAgencia, "(")
'            '**Modificado por DAOR 20080827 ****************************
'            'lsAgencia = Left(lsAgencia, lnPosicion - 1)
'            If lnPosicion > 0 Then
'                lsAgencia = Left(lsAgencia, lnPosicion - 1)
'            End If
'            '***********************************************************
'        End If
'    Set loAge = Nothing
'
'    With oWord.Selection.Find
'        .Text = "sAgencia"
'        .Replacement.Text = lsAgencia
'        .Forward = True
'        .Wrap = wdFindContinue
'        .Format = False
'        .Execute Replace:=wdReplaceAll
'    End With
'
'    'Numero de Cuenta
'    With oWord.Selection.Find
'        .Text = "<<CRED>>"
'        .Replacement.Text = Left(psCtaCod, 3) & "-" & Mid(psCtaCod, 4, 2) & "-" & Mid(psCtaCod, 6, 3) & "-" & Mid(psCtaCod, 9, 10)
'        .Forward = True
'        .Wrap = wdFindContinue
'        .Format = False
'        .Execute Replace:=wdReplaceAll
'    End With
'
'    If lblCodAvalado.Caption <> "" Then
'    'AVAL
'    With oWord.Selection.Find
'        .Text = "<<AVAL>>"
'        .Replacement.Text = PstaNombre(lblNomAval.Caption, True)
'        .Forward = True
'        .Wrap = wdFindContinue
'        .Format = False
'        .Execute Replace:=wdReplaceAll
'    End With
'    End If
'
'    'By Capi Acta 035-2007
'    lsAgenciaDir = cDirecAgencia
'    lnPosicion = InStr(lsAgenciaDir, "(")
'    cDirecAgencia = Left(lsAgenciaDir, lnPosicion - 2)
'    With oWord.Selection.Find
'        .Text = "<<DIRECCION>>"
'        .Replacement.Text = cDirecAgencia
'        .Forward = True
'        .Wrap = wdFindContinue
'        .Format = False
'        .Execute Replace:=wdReplaceAll
'    End With
'
'
'
'    With oWord.Selection.Find
'        .Text = "<<FOLIO>>"
'        'ARCV 05-07-2007
'        '.Replacement.Text = Format(nPoliza, "0000000")
'        '.Replacement.Text = Format(CLng(txtNumPoliza.Text), "0000000")
'        '-------
'        .Forward = True
'        .Wrap = wdFindContinue
'        .Format = False
'        .Execute Replace:=wdReplaceAll
'    End With
'
'
'    dfechafin = CDate(lrDataCF!Vence)
'    lsFechas = Format(dfechafin, "dd") & " de " & Format(dfechafin, "mmmm") & " del " & Format(dfechafin, "yyyy")
'    'Fecha Vencimineto
'    With oWord.Selection.Find
'        .Text = "<<VENCIMIENTO>>"
'        .Replacement.Text = lsFechas
'        .Forward = True
'        .Wrap = wdFindContinue
'        .Format = False
'        .Execute Replace:=wdReplaceAll
'    End With
'
'    'Fecha Actual
'    lsFechas = Format(gdFecSis, "dd") & " de " & Format(gdFecSis, "mmmm") & " del " & Format(gdFecSis, "yyyy")
'    With oWord.Selection.Find
'        .Text = "<<FECHA>>"
'        .Replacement.Text = lsFechas
'        .Forward = True
'        .Wrap = wdFindContinue
'        .Format = False
'        .Execute Replace:=wdReplaceAll
'    End With
'
''    With oWord.Selection.Find
''        .Text = "CcCliente"
''        .Replacement.Text = PstaNombre(lrDataCR!Nombre, True)
''        .Forward = True
''        .Wrap = wdFindContinue
''        .Format = False
''        .Execute Replace:=wdReplaceAll
''    End With
'
'    'ADREEDOR
'    With oWord.Selection.Find
'        .Text = "<<SEÑORES>>"
'        .Replacement.Text = PstaNombre(lblNomAcreedor, True)
'        .Forward = True
'        .Wrap = wdFindContinue
'        .Format = False
'        .Execute Replace:=wdReplaceAll
'    End With
'
'    'CLIENTE
'    With oWord.Selection.Find
'        .Text = "<<SOLICITANTE>>"
'        .Replacement.Text = PstaNombre(lblNomcli, True)
'        .Forward = True
'        .Wrap = wdFindContinue
'        .Format = False
'        .Execute Replace:=wdReplaceAll
'    End With
'
'
'    'Monto
'    With oWord.Selection.Find
'        .Text = "<<MONTO>>"
'        'ARCV 09-05-2007
'        '.Replacement.Text = IIf(Mid(psCtaCod, 9, 1) = "1", "S/.", "$.") & Format(lrDataT!nSaldo, "#,###0.00") & " " & "(" & UnNumero(lrDataT!nSaldo) & IIf(Mid(psCtaCod, 9, 1) = "1", "00/100 NUEVOS SOLES", "") & ")"
'        'By Capi Acta 035-2007 para que añada descripcion de moneda
'        .Replacement.Text = IIf(Mid(psCtaCod, 9, 1) = "1", "S/. ", "$. ") & Format(lrDataT!nSaldo, "#,###0.00") & " " & "(" & UCase(NumLet(lrDataT!nSaldo)) & IIf(Mid(psCtaCod, 9, 1) = "2", "", " Y " & IIf(InStr(1, lrDataT!nSaldo, ".") = 0, "00", Mid(lrDataT!nSaldo, InStr(1, lrDataT!nSaldo, ".") + 1, 2)) & "/100 ") & IIf(Mid(psCtaCod, 9, 1) = "1", "NUEVOS SOLES)", "US DOLARES)")
'        '-------------
'        .Forward = True
'        .Wrap = wdFindContinue
'        .Format = False
'        .Execute Replace:=wdReplaceAll
'    End With
'
'    'Finalidad
'    With oWord.Selection.Find
'        .Text = "<<Finalidad>>"
'        .Replacement.Text = Mid(lblFinalidad.Caption, 1, 250)
'        .Forward = True
'        .Wrap = wdFindContinue
'        .Format = False
'        .Execute Replace:=wdReplaceAll
'    End With
'
'    'Modalidad
'    With oWord.Selection.Find
'        .Text = "<<Modalidad>>"
'        .Replacement.Text = lblModalidad.Caption
'        .Forward = True
'        .Wrap = wdFindContinue
'        .Format = False
'        .Execute Replace:=wdReplaceAll
'    End With
'
'   oDoc.SaveAs App.path & "\SPOOLER\" & psCtaCod & ".doc"
  
  
   'MADM 20110121 ---------------------------------------------------------------
'   Dim x, sImpresora As String
'   Dim Prt As Printer
'   Dim xbol As Boolean
'   Dim Pred As String
'   xbol = False
'   sImpresora = Printer.DeviceName
'   x = App.path & "\SPOOLER\" & psCtaCod & ".doc"
'
'   'frmImpresora.Show 1
'   'frmImpresora.Inicia
'
' If sImpresora <> sLpt And sImpresora <> "" Then
'    oWord.Application.ActivePrinter = sLpt
'    xbol = True
'   End If
'
'    If oWord.Application.ActivePrinter = "" Then
'    Else
'        oWord.PrintOut Filename:=x, Range:=wdPrintAllDocument, iTem:= _
'        wdPrintDocumentContent, Copies:=1, Pages:="", PageType:=wdPrintAllPages, _
'        ManualDuplexPrint:=False, Collate:=True, Background:=True, PrintToFile:= _
'        False, PrintZoomColumn:=0, PrintZoomRow:=0, PrintZoomPaperWidth:=0, _
'        PrintZoomPaperHeight:=0
'    End If
'
'   Do While MsgBox("Desea Reimprimir el Comprobante de Pago?", vbInformation + vbYesNo, "Aviso") = vbYes
'
'        'frmImpresora.Inicia
'
'         If oWord.Application.ActivePrinter <> "" Then
'             oWord.PrintOut Filename:=x, Range:=wdPrintAllDocument, iTem:= _
'             wdPrintDocumentContent, Copies:=1, Pages:="", PageType:=wdPrintAllPages, _
'             ManualDuplexPrint:=False, Collate:=True, Background:=True, PrintToFile:= _
'             False, PrintZoomColumn:=0, PrintZoomRow:=0, PrintZoomPaperWidth:=0, _
'             PrintZoomPaperHeight:=0
'        End If
'   Loop
'   oWord.ActiveDocument.Close SaveChanges:=wdDoNotSaveChanges
'
'   Kill App.path & "\SPOOLER\" & psCtaCod & ".doc"
'
'   If xbol = True Then
'        oWord.Application.ActivePrinter = sImpresora
'   End If
'
'   oWord.Quit
   'MADM 20110121 ---------------------------------------------------------------
End Sub

Sub ImpreDocRenovado(ByVal psCtaCod As String)

    Dim lrDataT As ADODB.Recordset
    Dim lrDataCF As ADODB.Recordset
    Dim lrDataCR As ADODB.Recordset
    
    Dim nDias As Integer
    Dim dfechaini As Date
    Dim dfechafin As Date
    Dim lsFechas As String
    
    Dim nPoliza As Long
    Dim cDirecAgencia As String
    Dim oWord As Word.Application
    Dim oDoc As Word.Document
    Dim oRange As Word.Range
    Dim rsCartaFianza As ADODB.Recordset 'WIOR 20120619
    Dim loRs As COMNCartaFianza.NCOMCartaFianzaValida
    Set loRs = New COMNCartaFianza.NCOMCartaFianzaValida
    
    Set lrDataCF = loRs.RecuperaDatosGeneralesCF(psCtaCod)
    Set lrDataT = loRs.RecuperaDatosT(psCtaCod)
    Set lrDataCR = loRs.RecuperaDatosAcreedor(psCtaCod)
      
    'By Capi Acta 035-2007
    nPoliza = CLng(loRs.GetCF_Poliza(psCtaCod))
    'Call frmCFImpresion.Inicio(vCodCta, IIf(Me.lblCodAvalado.Caption = "", False, True), nPoliza, 2) 'WIOR 20120427
    'WIOR 20120619 **********************************************************
    If fbRemesado Then
        Set rsCartaFianza = loRs.ObtenerEnvioFolios("1", gsCodAge)
        If rsCartaFianza.RecordCount > 0 Then
            Set rsCartaFianza = loRs.UltimoRegistroEnvio(gsCodAge, 1, , 1)
            If rsCartaFianza.RecordCount > 0 Then
                Call loRs.ActualizarEnvioFolios(Trim(rsCartaFianza!nCodEnvio), 2)
            End If
            Call loRs.ActualizarFolio(nPoliza, 4, psCtaCod, gdFecSis)
            Call loRs.ActualizarFolioSinNumero(fsCodEnvio, 1)
        End If
    End If
    'WIOR FIN ***************************************************************
    Call ImprimirRenovacionPDF(vCodCta, IIf(Me.lblCodAvalado.Caption = "", False, True), nPoliza, 2)
    
'    Set oWord = CreateObject("Word.Application")
'    'MADM 20110121 ---------------------------------------------------------------
'    oWord.Visible = True
'    'oWord.Visible = False
'    'MADM 20110121 ---------------------------------------------------------------
'    If lblCodAvalado.Caption <> "" Then
'        Set oDoc = oWord.Documents.Open(App.path & "\FormatoCarta\CartaFianzaRenovacionGar.doc")
'    Else
'        Set oDoc = oWord.Documents.Open(App.path & "\FormatoCarta\CartaFianzaRenovacion.doc")
'    End If
'
'    Dim loAge As COMDConstantes.DCOMAgencias
'    Dim rs1 As ADODB.Recordset
'    Dim lsAgencia As String
'    Dim lnPosicion As Integer
'
'    Set loAge = New COMDConstantes.DCOMAgencias
'    Set rs1 = New ADODB.Recordset
'        Set rs1 = loAge.RecuperaAgencias(gsCodAge)
'        If Not (rs1.EOF And rs1.BOF) Then
'            lsAgencia = Trim(rs1("cUbiGeoDescripcion"))
'            'By Capi Acta 035-2007
'            lnPosicion = InStr(lsAgencia, "(")
'            '**Modificado por DAOR 20080827 ****************************
'            'lsAgencia = Left(lsAgencia, lnPosicion - 1)
'            If lnPosicion > 0 Then
'                lsAgencia = Left(lsAgencia, lnPosicion - 1)
'            End If
'            '***********************************************************
'        End If
'    Set loAge = Nothing
'
'    With oWord.Selection.Find
'        .Text = "sAgencia"
'        .Replacement.Text = lsAgencia
'        .Forward = True
'        .Wrap = wdFindContinue
'        .Format = False
'        .Execute Replace:=wdReplaceAll
'    End With
'    'By Capi Acta 035-2007
'    With oWord.Selection.Find
'        .Text = "NROSEG"
'        'ARCV 05-07-2007
'        .Replacement.Text = Format(nPoliza, "0000000")
'        '.Replacement.Text = Format(val(txtNumPoliza.Text), "0000000")
'        '-------
'        .Forward = True
'        .Wrap = wdFindContinue
'        .Format = False
'        .Execute Replace:=wdReplaceAll
'    End With
'
'    'Fecha
'    lsFechas = Format(lrDataCF!F_Asignacion, "dd") & " de " & Format(lrDataCF!F_Asignacion, "mmmm") & " del " & Format(lrDataCF!F_Asignacion, "yyyy")
'    With oWord.Selection.Find
'        .Text = "dFecha"
'        .Replacement.Text = lsFechas
'        .Forward = True
'        .Wrap = wdFindContinue
'        .Format = False
'        .Execute Replace:=wdReplaceAll
'    End With
'
'    'Titular
'    With oWord.Selection.Find
'        .Text = "cTitular"
'        .Replacement.Text = PstaNombre(lblNomcli, True)
'        .Forward = True
'        .Wrap = wdFindContinue
'        .Format = False
'        .Execute Replace:=wdReplaceAll
'    End With
'
'    'MADM 20111020
'    If lblCodAvalado.Caption <> "" Then
'        With oWord.Selection.Find
'            .Text = "AVAL"
'            .Replacement.Text = PstaNombre(lblNomAval, True)
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'    End If
'
'    'Cuenta
'    With oWord.Selection.Find
'        .Text = "cCtaCod"
'        .Replacement.Text = ActXCodCta.NroCuenta
'        .Forward = True
'        .Wrap = wdFindContinue
'        .Format = False
'        .Execute Replace:=wdReplaceAll
'    End With
'
'    'By Capi Acta 035-2007 pa que jale el numero de renovacion
'    With oWord.Selection.Find
'        .Text = "NroRen"
'        .Replacement.Text = fnRenovacion
'        .Forward = True
'        .Wrap = wdFindContinue
'        .Format = False
'        .Execute Replace:=wdReplaceAll
'    End With
'
'    'fecha de creacion
'    dfechaini = lrDataCF!dPrdEstado
'    With oWord.Selection.Find
'        .Text = "dFecCrea"
'        .Replacement.Text = Format(dfechaini, "DD/MM/YYYY")
'        .Forward = True
'        .Wrap = wdFindContinue
'        .Format = False
'        .Execute Replace:=wdReplaceAll
'    End With
'
'    'Monto
'    With oWord.Selection.Find
'        .Text = "nMonto"
'        .Replacement.Text = IIf(Mid(psCtaCod, 9, 1) = "1", "S/.", "$.") & Format(lblMontoApr, "#,###0.00") '& " " & UnNumero(lblMontoSol)
'        .Forward = True
'        .Wrap = wdFindContinue
'        .Format = False
'        .Execute Replace:=wdReplaceAll
'    End With
'
'    dfechafin = CDate(lrDataCF!dVenc)
'    'Fecha Vencimiento
'    With oWord.Selection.Find
'        .Text = "dFecVenA"
'        .Replacement.Text = Format(dfechafin, "DD/MM/YYYY")
'        .Forward = True
'        .Wrap = wdFindContinue
'        .Format = False
'        .Execute Replace:=wdReplaceAll
'    End With
'
'    'ADREEDOR
'    With oWord.Selection.Find
'        .Text = "dAcreedor"
'        .Replacement.Text = PstaNombre(lblNomAcreedor, True)
'        .Forward = True
'        .Wrap = wdFindContinue
'        .Format = False
'        .Execute Replace:=wdReplaceAll
'    End With
'
'    'Finalidad
'    With oWord.Selection.Find
'        .Text = "cFinalidad"
'        .Replacement.Text = Left(Trim(lrDataCF!Finalidad), 255)
'        .Forward = True
'        .Wrap = wdFindContinue
'        .Format = False
'        .Execute Replace:=wdReplaceAll
'    End With
'
'    'Nueva Fecha
'    With oWord.Selection.Find
'        .Text = "dFecVenN"
'        .Replacement.Text = lblFecVencApr
'        .Forward = True
'        .Wrap = wdFindContinue
'        .Format = False
'        .Execute Replace:=wdReplaceAll
'    End With
'
'   oDoc.SaveAs App.path & "\SPOOLER\" & psCtaCod & ".doc"
'   'MADM 20110121 ---------------------------------------------------------------
'   Dim x, sImpresora As String
'   Dim Prt As Printer
'   Dim xbol As Boolean
'
'   xbol = False
'   sImpresora = Printer.DeviceName
'
'   x = App.path & "\SPOOLER\" & psCtaCod & ".doc"
'
'   'frmImpresora.Show 1
'    frmImpresora.inicia
'
'   If sImpresora <> sLpt And sImpresora <> "" Then
'    oWord.Application.ActivePrinter = sLpt
'    xbol = True
'   End If
'
'    If oWord.Application.ActivePrinter = "" Then
'    Else
'        oWord.PrintOut Filename:=x, Range:=wdPrintAllDocument, iTem:= _
'        wdPrintDocumentContent, Copies:=1, Pages:="", PageType:=wdPrintAllPages, _
'        ManualDuplexPrint:=False, Collate:=True, Background:=True, PrintToFile:= _
'        False, PrintZoomColumn:=0, PrintZoomRow:=0, PrintZoomPaperWidth:=0, _
'        PrintZoomPaperHeight:=0
'    End If
'
'   Do While MsgBox("Desea Reimprimir el Comprobante de Pago?", vbInformation + vbYesNo, "Aviso") = vbYes
'
'        frmImpresora.inicia
'
'         If oWord.Application.ActivePrinter <> "" Then
'             oWord.PrintOut Filename:=x, Range:=wdPrintAllDocument, iTem:= _
'             wdPrintDocumentContent, Copies:=1, Pages:="", PageType:=wdPrintAllPages, _
'             ManualDuplexPrint:=False, Collate:=True, Background:=True, PrintToFile:= _
'             False, PrintZoomColumn:=0, PrintZoomRow:=0, PrintZoomPaperWidth:=0, _
'             PrintZoomPaperHeight:=0
'        End If
'   Loop
'   oWord.ActiveDocument.Close SaveChanges:=wdDoNotSaveChanges
'
'   Kill App.path & "\SPOOLER\" & psCtaCod & ".doc"
'
'   If xbol = True Then
'        oWord.Application.ActivePrinter = sImpresora
'   End If
'
'   oWord.Quit
'   'MADM 20110121 ---------------------------------------------------------------
End Sub


Private Sub txtNumPoliza_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
End Sub
'WIOR 20120613**********************************************
Private Sub cmdGenerarPDF_Click()
Dim oCF As COMNCartaFianza.NCOMCartaFianzaValida
Dim nPoliza As Long
On Error GoTo ErrorGenerarPdf
vCodCta = ActXCodCta.NroCuenta

If ValidaDatos = False Then
    Exit Sub
End If
If fsEstado <> gColocEstRenovada Then
    Call ImprimirPDF(vCodCta, IIf(Me.lblCodAvalado.Caption = "", False, True), Me.txtNumPoliza.Text, 1)
Else
    Set oCF = New COMNCartaFianza.NCOMCartaFianzaValida
    nPoliza = CLng(oCF.GetCF_Poliza(vCodCta))
    Call ImprimirRenovacionPDF(vCodCta, IIf(Me.lblCodAvalado.Caption = "", False, True), nPoliza, 1)
    Set oCF = Nothing
End If
MsgBox "Archivo Previo Generado Satisfacoriamente.", vbInformation, "Aviso"
Exit Sub
ErrorGenerarPdf:
    MsgBox Err.Description, vbInformation, "Aviso"
End Sub

Private Sub ImprimirPDF(ByVal psCodCta As String, ByVal pbAvalado As Boolean, ByVal psNumFolio As String, ByVal nTipo As Integer)
    On Error GoTo ErrorImprimirPDF
    Dim loRs As COMNCartaFianza.NCOMCartaFianzaValida
    Dim lrDataT As ADODB.Recordset
    Dim lrDataCF As ADODB.Recordset
    Dim lrDataCR As ADODB.Recordset
    Dim dfechaini As Date
    Dim nCFPoliza As Long
    Dim sParrafo1 As String
    Dim sParrafo2 As String
    Dim sParrafo3 As String
    Dim sParrafo4 As String
    Dim nTamano As Integer
    Dim nValidar As Double
    Dim nTop As Integer
    Dim sFechaActual As String
    Dim sSenores As String
    Dim sAval As String
    Dim sSolicitante As String
    Dim sMonto As String
    Dim sModalidad As String
    Dim sFinalidad As String
    Dim dfechainipdf As Date 'FRHU20131126
    Dim sVigenciapdf As String 'FRHU20131126
    Dim dfechafin As Date
    Dim sVencimiento As String
    Dim sDireccion As String
    Dim lnPosicion As Integer
    Dim oDoc  As cPDF
    
    Set loRs = New COMNCartaFianza.NCOMCartaFianzaValida
    Set lrDataCF = loRs.RecuperaDatosGeneralesCF(psCodCta)
    Set lrDataT = loRs.RecuperaDatosT(psCodCta)
    Set oDoc = New cPDF
    
    nCFPoliza = psNumFolio
    
    'Creacion de Archivo
    oDoc.Author = gsCodUser
    oDoc.Creator = "SICMACT - Negocio"
    oDoc.Producer = "Caja Municipal de Ahorros y Crédito de Maynas S.A."
    oDoc.Subject = "Carta Fianza Nº " & psCodCta
    oDoc.Title = "Carta Fianza Nº " & psCodCta
    If Not oDoc.PDFCreate(App.Path & "\Spooler\" & IIf(nTipo = 1, "Previo", "") & psCodCta & "_" & Format(gdFecSis, "YYYYMMDD") & "_" & Format(Time, "hhmmss") & ".pdf") Then
        Exit Sub
    End If
    
    oDoc.Fonts.Add "F1", "Times New Roman", TrueType, Normal, WinAnsiEncoding 'FRHU20131126
    oDoc.Fonts.Add "F2", "Times New Roman", TrueType, Bold, WinAnsiEncoding 'FRHU20131126
    
    oDoc.NewPage A4_Vertical

    sFechaActual = Format(gdFecSis, "dd") & " de " & Format(gdFecSis, "mmmm") & " del " & Format(gdFecSis, "yyyy")
    sSenores = PstaNombre(lblNomAcreedor, True)
    If pbAvalado Then
        sAval = PstaNombre(lblNomAval.Caption, True)
    End If
    sSolicitante = PstaNombre(lblNomcli, True)
    'WIOR 20120705
    Dim sSaldo As String
    sSaldo = Format(lrDataT!nSaldo, "#,###0.00")
    '''sMonto = IIf(Mid(psCodCta, 9, 1) = "1", "S/. ", "$. ") & sSaldo & " " & "(" & UCase(NumLet(sSaldo)) & IIf(Mid(psCodCta, 9, 1) = "2", "", " Y " & IIf(InStr(1, sSaldo, ".") = 0, "00", Mid(sSaldo, InStr(1, sSaldo, ".") + 1, 2)) & "/100 ") & IIf(Mid(psCodCta, 9, 1) = "1", " NUEVOS SOLES)", " DOLARES)") 'marg ers044-2016
    sMonto = IIf(Mid(psCodCta, 9, 1) = "1", gcPEN_SIMBOLO & " ", "$. ") & sSaldo & " " & "(" & UCase(NumLet(sSaldo)) & IIf(Mid(psCodCta, 9, 1) = "2", "", " Y " & IIf(InStr(1, sSaldo, ".") = 0, "00", Mid(sSaldo, InStr(1, sSaldo, ".") + 1, 2)) & "/100 ") & IIf(Mid(psCodCta, 9, 1) = "1", " " & StrConv(gcPEN_PLURAL, vbUpperCase) & ")", " DOLARES)") 'marg ers044-2016
    'sMonto = IIf(Mid(psCodCta, 9, 1) = "1", "S/. ", "$. ") & Format(lrDataT!nSaldo, "#,###0.00") & " " & "(" & UCase(NumLet(lrDataT!nSaldo)) & IIf(Mid(psCodCta, 9, 1) = "2", "", " Y " & IIf(InStr(1, str(Format(lrDataT!nSaldo, "#,###0.00")), ".") = 0, "00", Mid(str(Format(lrDataT!nSaldo, "#,###0.00")), InStr(1, str(Format(lrDataT!nSaldo, "#,###0.00")), ".") + 1, 2)) & "/100 ") & IIf(Mid(psCodCta, 9, 1) = "1", "NUEVOS SOLES)", "US DOLARES)")
    'sMonto = IIf(Mid(psCodCta, 9, 1) = "1", "S/. ", "$. ") & Format(lrDataT!nSaldo, "#,###0.00") & " " & "(" & UCase(NumLet(lrDataT!nSaldo)) & IIf(Mid(psCodCta, 9, 1) = "2", "", " Y " & IIf(InStr(1, lrDataT!nSaldo, ".") = 0, "00", Mid(lrDataT!nSaldo, InStr(1, lrDataT!nSaldo, ".") + 1, 2)) & "/100 ") & IIf(Mid(psCodCta, 9, 1) = "1", "NUEVOS SOLES)", "US DOLARES)")
    sModalidad = Trim(lblModalidad.Caption)
    sFinalidad = Trim(lblFinalidad.Caption)
    
    'F_Asignacion
    'FRHU20131126
    dfechainipdf = CDate(lrDataCF!F_Asignacion)
    sVigenciapdf = Format(dfechainipdf, "dd") & " de " & Format(dfechainipdf, "mmmm") & " del " & Format(dfechainipdf, "yyyy")
    'FIN FRHU20131126
    dfechafin = CDate(lrDataCF!Vence)
    sVencimiento = Format(dfechafin, "dd") & " de " & Format(dfechafin, "mmmm") & " del " & Format(dfechafin, "yyyy")
    sDireccion = loRs.Get_Agencia_CF(psCodCta)
    lnPosicion = InStr(sDireccion, "(")
    sDireccion = Left(sDireccion, lnPosicion - 2)
    
    oDoc.WTextBox 70, 50, 10, 450, Left(psCodCta, 3) & "-" & Mid(psCodCta, 4, 2) & "-" & Mid(psCodCta, 6, 3) & "-" & Mid(psCodCta, 9, 10), "F1", 12, hRight
    oDoc.WTextBox 120, 50, 10, 450, "CARTA FIANZA N° " & Format(nCFPoliza, "0000000"), "F1", 12, hCenter
    oDoc.WTextBox 170, 50, 10, 450, sFechaActual, "F2", 12, hRight 'FRHU20131126
    oDoc.WTextBox 220, 50, 10, 450, "Señores:", "F1", 12, hLeft
    oDoc.WTextBox 232, 50, 10, 450, sSenores, "F2", 12, hLeft 'FRHU20131126
    oDoc.WTextBox 260, 50, 10, 450, "Ciudad.-", "F1", 12, hLeft
    oDoc.WTextBox 280, 50, 10, 450, "Muy Señores Nuestros:", "F1", 12, hLeft
    sAval = " garantizando a " & sAval
    sParrafo1 = "A solicitud de " & sSolicitante & ", otorgamos por el presente " & _
                "documento una fianza solidaria, irrevocable, incondicional, de " & _
                "ejecución inmediata, con renuncia expresa al beneficio de " & _
                "excusión e indivisible, a favor de ustedes" & IIf(pbAvalado = True, sAval, "") & _
                ", hasta por la suma de " & sMonto & ", a fin de garantizar " & _
                "la Carta Fianza por " & sModalidad & ", objeto del proceso: " & sFinalidad & "."
    nTamano = Len(sParrafo1)
    nValidar = nTamano / 75
    nTamano = CInt(IIf(CInt(nValidar) > nValidar, CInt(nValidar), CInt(nValidar) + 1))
    nTop = 270
    'oDoc.WTextBox nTop, 50, nTamano * 10, 450, sParrafo1, "F1", 12, hjustify
    oDoc.WTextBox nTop, 0, nTamano * 20, 580, String(20, "-") & " " & sParrafo1, "F1", 11, hjustify, , , , , , 50
    oDoc.WTextBox nTop + 50, 50, 10, 75, "", "F1", 10, hjustify, , vbWhite, 1, vbWhite, True

    nTop = nTop + (nTamano * 10) + 12
      
    sParrafo2 = "Dejamos claramente establecido que la presente " & String(1, vbTab) & "Carta " & String(1, vbTab) & "Fianza no " & _
                "podrá ser usada " & String(1, vbTab) & "para operaciones comprendidas en la prohibición " & _
                "indicada en el inciso ''5'' del Articulo 217 de la " & String(1, vbTab) & "Ley  26702, Ley " & _
                "General del " & String(1, vbTab) & "Sistema " & String(1, vbTab) & "Financiero y del Sistema de Seguros y Orgánica " & _
                "de la Superintendencia de --- Banca y Seguros."
                
                
    nTamano = Len(sParrafo2)
    nValidar = nTamano / 80
    nTamano = CInt(IIf(CInt(nValidar) > nValidar, CInt(nValidar), CInt(nValidar) + 1))
    'oDoc.WTextBox nTop, 50, nTamano * 10, 450, sParrafo2, "F1", 12, hjustify
    oDoc.WTextBox nTop, 0, nTamano * 10, 580, String(20, "-") & " " & sParrafo2, "F1", 11, hjustify, , , , , , 50
    oDoc.WTextBox nTop + 50, 50, 10, 75, "", "F1", 10, hjustify, , vbWhite, 1, vbWhite, True
    oDoc.WTextBox nTop + 75, 520, 10, 20, "", "F1", 10, hjustify, , vbWhite, 1, vbWhite, True
    nTop = nTop + (nTamano * 12) + 12
    'JGPA20190614 Cambio razón social según Memorandum Nº 1037-2019-GM-DI/CMACM
    sParrafo3 = "Por efecto de este compromiso la CAJA MUNICIPAL DE AHORRO Y CRÉDITO MAYNAS S.A. " & _
                        "asume con su fiado las responsabilidades en que éste llegara a " & _
                        "incurrir siempre que el " & String(1, vbTab) & "monto de las  mismas  no " & String(1, vbTab) & "exceda por ningún " & _
                        "motivo de la suma antes mencionada y que estén estrictamente " & _
                        "vinculadas al cumplimiento de lo arriba indicado."
    
    nTamano = Len(sParrafo3)
    nValidar = nTamano / 90
    nTamano = CInt(IIf(CInt(nValidar) > nValidar, CInt(nValidar), CInt(nValidar) + 1))
    'oDoc.WTextBox nTop, 50, nTamano * 10, 450, sParrafo3, "F1", 12, hjustify, vMiddle, , , , False
    oDoc.WTextBox nTop, 0, nTamano * 10, 580, String(20, "-") & " " & sParrafo3, "F1", 11, hjustify, , , , , , 50
    oDoc.WTextBox nTop + 50, 50, 10, 75, "", "F1", 10, hjustify, , vbWhite, 1, vbWhite, True
    nTop = nTop + (nTamano * 12) + 13
    
    'FRHU20131126
    sParrafo4 = "" & _
                        "La presente garantía rige a partir del " & sVigenciapdf & " y vencerá " & _
                        "el " & sVencimiento & ". Cualquier  reclamo en virtud de esta " & _
                        "garantía deberá ceñirse estrictamente a lo estipulado por " & _
                        "el Art. 1898 del Código Civil y deberá ser formulado por vía " & _
                        "notarial y en nuestra oficina ubicada en " & sDireccion & "."
    nTamano = Len(sParrafo4)
    nValidar = nTamano / 90
    nTamano = CInt(IIf(CInt(nValidar) > nValidar, CInt(nValidar), CInt(nValidar) + 1))
    'oDoc.WTextBox nTop, 50, nTamano * 10, 450, sParrafo4, "F1", 12, hjustify, vMiddle, , , , False
    oDoc.WTextBox nTop, 0, nTamano * 10, 580, String(20, "-") & " " & sParrafo4, "F1", 11, hjustify, , , , , , 50
    oDoc.WTextBox nTop + 50, 50, 10, 75, "", "F1", 10, hjustify, , vbWhite, 1, vbWhite, True
    nTop = nTop + (nTamano * 13) + 80

    oDoc.WTextBox nTop, 50, 10, 450, "Atentamente,", "F1", 12, hCenter, vMiddle, , , , False
    oDoc.WTextBox nTop + 12, 50, 10, 450, "CAJA MUNICIPAL DE AHORRO Y CRÉDITO MAYNAS S.A.", "F1", 12, hCenter, vMiddle, , , , False 'JGPA20190614 Cambio razón social según Memorandum Nº 1037-2019-GM-DI/CMACM

    oDoc.PDFClose
    oDoc.Show
    Exit Sub
ErrorImprimirPDF:
    MsgBox Err.Description, vbInformation, "Aviso"
End Sub

Private Sub ImprimirRenovacionPDF(ByVal psCodCta As String, ByVal pbAvalado As Boolean, ByVal psNumFolio As String, ByVal nTipo As Integer)
    On Error GoTo ErrorImprimirRenovacionPDF
    Dim loRs As COMNCartaFianza.NCOMCartaFianzaValida
    Dim lrDataT As ADODB.Recordset
    Dim lrDataCF As ADODB.Recordset
    Dim loAge As COMDConstantes.DCOMAgencias
    Dim rsAge As ADODB.Recordset
    Dim lsAgencia As String
    Dim lnPosicion As Integer
    Dim dfechaini As Date
    Dim dfechafin As Date
    Dim lsFechas As String
    Dim nCFPoliza As Long
    Dim cDirecAgencia As String
    Dim sAcreedor As String
    Dim sAval As String
    Dim sSolicitante As String
    Dim sMonto As String
    Dim sFinalidad As String
    Dim sDireccion As String
    Dim sParrafo1 As String
    Dim sParrafo2 As String
    Dim sParrafo3 As String
    Dim sParrafo4 As String
    Dim sParrafo5 As String
    Dim sParrafo6 As String
    Dim nTamano As Integer
    Dim nValidar As Double
    Dim nTop As Integer
    
    Dim oDoc  As cPDF
    
    Set loRs = New COMNCartaFianza.NCOMCartaFianzaValida
    Set lrDataCF = loRs.RecuperaDatosGeneralesCF(psCodCta)
    Set lrDataT = loRs.RecuperaDatosT(psCodCta)

    Set loAge = New COMDConstantes.DCOMAgencias
    Set rsAge = New ADODB.Recordset
        Set rsAge = loAge.RecuperaAgencias(gsCodAge)
        If Not (rsAge.EOF And rsAge.BOF) Then
            lsAgencia = Trim(rsAge("cUbiGeoDescripcion"))
            lnPosicion = InStr(lsAgencia, "(")
            If lnPosicion > 0 Then
                lsAgencia = Left(lsAgencia, lnPosicion - 1)
            End If
        End If
    Set loAge = Nothing
    Set oDoc = New cPDF
    
    lsFechas = Format(lrDataCF!F_Asignacion, "dd") & " de " & Format(lrDataCF!F_Asignacion, "mmmm") & " del " & Format(lrDataCF!F_Asignacion, "yyyy")
    nCFPoliza = psNumFolio
    
    'Creacion del Archivo
    oDoc.Author = gsCodUser
    oDoc.Creator = "SICMACT - Negocio"
    oDoc.Producer = "Caja Municipal de Ahorros y Crédito de Maynas S.A."
    oDoc.Subject = "Renovación de Carta Fianza Nº " & psCodCta
    oDoc.Title = "Renovación Carta Fianza Nº " & psCodCta
    
    If Not oDoc.PDFCreate(App.Path & "\Spooler\" & IIf(nTipo = 1, "Previo", "") & "Renovacion" & psCodCta & "_" & Format(gdFecSis, "YYYYMMDD") & "_" & Format(Time, "hhmmss") & ".pdf") Then
        Exit Sub
    End If
    
    oDoc.Fonts.Add "F1", "Times New Roman", TrueType, Bold, WinAnsiEncoding
    oDoc.Fonts.Add "F2", "Times New Roman", TrueType, BoldItalic, WinAnsiEncoding
    
    
    'Tamaño de hoja A4
    oDoc.NewPage A4_Vertical

    sAcreedor = PstaNombre(lblNomAcreedor, True)
    sSolicitante = PstaNombre(lblNomcli, True)
    '''sMonto = IIf(Mid(psCodCta, 9, 1) = "1", "S/.", "$.") & Format(lrDataT!nSaldo, "#,###0.00") 'marg ers044-2016
    sMonto = IIf(Mid(psCodCta, 9, 1) = "1", gcPEN_SIMBOLO, "$.") & Format(lrDataT!nSaldo, "#,###0.00") 'marg ers044-2016
    sFinalidad = Trim(lblFinalidad.Caption)
    sDireccion = loRs.Get_Agencia_CF(psCodCta)
    lnPosicion = InStr(sDireccion, "(")
    sDireccion = Left(sDireccion, lnPosicion - 2)
    dfechaini = CDate(lrDataCF!dPrdEstado)
    dfechafin = CDate(lrDataCF!dVenc)
    
    oDoc.WTextBox 70, 50, 10, 450, "Renovación Nº " & fnRenovacion, "F1", 12, hRight
    oDoc.WTextBox 120, 50, 10, 450, lsAgencia & ",  " & lsFechas, "F1", 12, hRight
    oDoc.WTextBox 170, 50, 10, 450, "Señores:", "F1", 12, hLeft
    oDoc.WTextBox 200, 50, 10, 450, sAcreedor, "F2", 12, hCenter
    oDoc.WTextBox 250, 50, 10, 450, "Ciudad.-", "F1", 12, hLeft
    
    
    If Not pbAvalado Then
        
        sParrafo1 = "REF: Nuestra Carta Fianza Nº " & Format(nCFPoliza, "0000000") & " con crédito Nº " & psCodCta & " del " & _
                    Format(dfechaini, "DD/MM/YYYY") & " para la obra ''" & sFinalidad & "'', por " & sMonto & "  con vencimiento " & _
                    Format(dfechafin, "DD/MM/YYYY") & " a favor de ustedes y a/c. de:"
        nTamano = Len(sParrafo1)
        nValidar = nTamano / 79
        nTamano = CInt(IIf(CInt(nValidar) > nValidar, CInt(nValidar), CInt(nValidar) + 1))
        nTop = 250
        'oDoc.WTextBox nTop, 50, nTamano * 10, 450, sParrafo1, "F1", 12, hjustify
        oDoc.WTextBox nTop, 0, nTamano * 10, 580, String(20, "-") & " " & sParrafo1, "F1", 11, hjustify, , , , , , 50
        oDoc.WTextBox nTop + 50, 50, 10, 75, "", "F1", 10, hjustify, , vbWhite, 1, vbWhite, True
        
        nTop = nTop + (nTamano * 12) + 60
        
        sParrafo2 = sSolicitante
        nTamano = Len(sParrafo2)
        nValidar = nTamano / 90
        nTamano = CInt(IIf(CInt(nValidar) > nValidar, CInt(nValidar), CInt(nValidar) + 1))
        oDoc.WTextBox nTop, 50, nTamano * 10, 450, sParrafo2, "F1", 12, hCenter, vMiddle, , , , False
        nTop = nTop + (nTamano * 12) + 25
    Else
        sAval = PstaNombre(lblNomAval.Caption, True)
        sParrafo1 = "REF: Nuestra Carta Fianza Nº " & Format(nCFPoliza, "0000000") & " con crédito Nº " & psCodCta & " del " & _
                    Format(dfechaini, "DD/MM/YYYY") & " para la obra ''" & sFinalidad & "'', por " & sMonto & "  con vencimiento " & _
                    Format(dfechafin, "DD/MM/YYYY") & " a favor de ustedes y a/c. de " & sSolicitante & ", garantizando a:"
        nTamano = Len(sParrafo1)
        nValidar = nTamano / 78
        nTamano = CInt(IIf(CInt(nValidar) > nValidar, CInt(nValidar), CInt(nValidar) + 1))
        nTop = 250
        'oDoc.WTextBox nTop, 50, nTamano * 10, 450, sParrafo1, "F1", 12, hjustify
        oDoc.WTextBox nTop, 0, nTamano * 10, 580, String(20, "-") & " " & sParrafo1, "F1", 11, hjustify, , , , , , 50
        oDoc.WTextBox nTop + 50, 50, 10, 75, "", "F1", 10, hjustify, , vbWhite, 1, vbWhite, True
        nTop = nTop + (nTamano * 12) + 60
        
        sParrafo2 = sAval
        nTamano = Len(sParrafo2)
        nValidar = nTamano / 90
        nTamano = CInt(IIf(CInt(nValidar) > nValidar, CInt(nValidar), CInt(nValidar) + 1))
        oDoc.WTextBox nTop, 50, nTamano * 10, 450, sParrafo2, "F1", 12, hCenter, vMiddle, , , , False
        nTop = nTop + (nTamano * 12) + 25
    
    End If
    
    oDoc.WTextBox nTop, 50, 10, 450, "Estimados Señores:", "F1", 11, hLeft
    nTop = nTop - 25
    sParrafo3 = "Sírvase a tomar nota de que, a solicitud de nuestros " & _
                        "garantizados, hemos procedido a:"
    nTamano = Len(sParrafo3)
    nValidar = nTamano / 80
    nTamano = CInt(IIf(CInt(nValidar) > nValidar, CInt(nValidar), CInt(nValidar) + 1))
    'oDoc.WTextBox nTop, 50, nTamano * 10, 450, sParrafo3, "F1", 12, hjustify, vMiddle, , , , False
    oDoc.WTextBox nTop, 0, nTamano * 10, 580, String(20, "-") & " " & sParrafo3, "F1", 11, hjustify, , , , , , 50
    oDoc.WTextBox nTop + 50, 50, 10, 75, "", "F1", 10, hjustify, , vbWhite, 1, vbWhite, True
    nTop = nTop + (nTamano * 12)
    
    sParrafo4 = "Prorrogar el plazo de vencimiento hasta el " & Format(lblFecVencApr, "DD/MM/YYYY") & "."
    nTamano = Len(sParrafo4)
    nValidar = nTamano / 80
    nTamano = CInt(IIf(CInt(nValidar) > nValidar, CInt(nValidar), CInt(nValidar) + 1))
    'oDoc.WTextBox nTop, 50, nTamano * 10, 450, sParrafo4, "F1", 12, hjustify, vMiddle, , , , False
    oDoc.WTextBox nTop, 0, nTamano * 10, 580, String(20, "-") & " " & sParrafo4, "F1", 11, hjustify, , , , , , 50
    oDoc.WTextBox nTop + 50, 50, 10, 75, "", "F1", 10, hjustify, , vbWhite, 1, vbWhite, True
    nTop = nTop + (nTamano * 12) + 20
    
    sParrafo5 = "Manteniéndose vigente todos los demás términos y consideraciones de las misma."
    nTamano = Len(sParrafo5)
    nValidar = nTamano / 90
    nTamano = CInt(IIf(CInt(nValidar) > nValidar, CInt(nValidar), CInt(nValidar) + 1))
    'oDoc.WTextBox nTop, 50, nTamano * 10, 450, sParrafo5, "F1", 12, hjustify, vMiddle, , , , False
    oDoc.WTextBox nTop, 0, nTamano * 10, 580, String(20, "-") & " " & sParrafo5, "F1", 11, hjustify, , , , , , 50
    oDoc.WTextBox nTop + 50, 50, 10, 75, "", "F1", 10, hjustify, , vbWhite, 1, vbWhite, True
    nTop = nTop + (nTamano * 12) + 20
    
    sParrafo6 = "Cualquier reclamo en virtud de esta garantía deberá " & _
                        "ceñirse a lo estipulado por el Art. 1898 del Código " & _
                        "Civil y deberá ser formulado por vía notarial en el " & _
                        "horario de atención al público y en nuestra Oficina " & _
                        "Ubicada en " & sDireccion & "."
    nTamano = Len(sParrafo6)
    nValidar = nTamano / 80
    nTamano = CInt(IIf(CInt(nValidar) > nValidar, CInt(nValidar), CInt(nValidar) + 1))
    'oDoc.WTextBox nTop, 50, nTamano * 10, 450, sParrafo6, "F1", 12, hjustify, vMiddle, , , , False
    oDoc.WTextBox nTop, 0, nTamano * 10, 580, String(20, "-") & " " & sParrafo6, "F1", 11, hjustify, , , , , , 50
    oDoc.WTextBox nTop + 50, 50, 10, 75, "", "F1", 10, hjustify, , vbWhite, 1, vbWhite, True
    
    nTop = nTop + (nTamano * 12) + 80
    
    oDoc.WTextBox nTop, 50, 10, 450, "Atentamente,", "F1", 12, hCenter, vMiddle, , , , False
    oDoc.WTextBox nTop + 12, 50, 10, 450, "CAJA MUNICIPAL DE AHORRO Y CRÉDITO MAYNAS S.A.", "F1", 12, hCenter, vMiddle, , , , False
    'JGPA20190614 Cambio razón social según Memorandum Nº 1037-2019-GM-DI/CMACM
    oDoc.PDFClose
    oDoc.Show
    Exit Sub
ErrorImprimirRenovacionPDF:
    MsgBox Err.Description, vbInformation, "Aviso"
End Sub
Private Sub ImprimePagare()
    
    On Error GoTo ErrorImprimirPDF
    Dim oCF As New COMNCartaFianza.NCOMCartaFianza
    Dim objAge As New COMDConstantes.DCOMAgencias
    Dim oDoc  As New cPDF
    Dim oRs As New ADODB.Recordset
    Dim rsDatosAge As ADODB.Recordset
    Dim psCodCta As String
    Dim sAgeCiudad As String
    Dim nTasaIntCom As Currency
    Dim nTasaIntMor As Currency
    Dim nIndice As Integer
    Dim nConTit As Integer
    Dim nContAva As Integer
    
    Set rsDatosAge = objAge.RecuperaAgencias(Mid(ActXCodCta.NroCuenta, 4, 2))
    
    If Not (rsDatosAge.BOF And rsDatosAge.EOF) Then
        sAgeCiudad = Trim(rsDatosAge!cUbiGeoDescripcion)
    End If
    oDoc.Author = gsCodUser
    oDoc.Creator = "SICMACT - Negocio"
    oDoc.Producer = "Caja Municipal de Ahorros y Crédito de Maynas S.A."
    oDoc.Subject = "Hoja de Aprobación de Créditos Nº " & gsCodUser
    oDoc.Title = "Hoja de Aprobación de Créditos Nº " & gsCodUser
    If Not oDoc.PDFCreate(App.Path & "\Spooler\" & psCodCta & "_" & Format(gdFecSis, "YYYYMMDD") & "_" & Format(Time, "hhmmss") & ".pdf") Then
        Exit Sub
    End If

    'FUENTES
    Dim nFTabla As Integer
    Dim nFTablaCabecera As Integer
    Dim lnFontSizeBody As Integer
    Dim lnFuenteT As Integer
    oDoc.Fonts.Add "F1", "arial narrow", TrueType, Normal, WinAnsiEncoding 'MODIFICADO PTI1 20170315
    oDoc.Fonts.Add "F2", "arial narrow", TrueType, Bold, WinAnsiEncoding 'MODIFICADO PTI1 20170315
    oDoc.Fonts.Add "F3", "Arial", TrueType, Normal, WinAnsiEncoding
    oDoc.Fonts.Add "F4", "Arial", TrueType, Bold, WinAnsiEncoding
    
    nFTablaCabecera = 7
    nFTabla = 7
    lnFontSizeBody = 7
    lnFuenteT = 8
    'FIN FUENTES
    
    
    Set oRs = oCF.ObtieneTasaCompCF(ActXCodCta.NroCuenta)
    If Not (oRs.EOF And oRs.BOF) Then
        nTasaIntCom = oRs!nTasaIntComp
        nTasaIntMor = oRs!nTasaIntMora
    End If
    Set oRs = Nothing
    
oDoc.LoadImageFromFile App.Path & "\Logo_2015.jpg", "Logo" 'MODIFICADO PTI1 20170315
    
    oDoc.NewPage A4_Vertical
'----------------'COMENTADO PTI1 20170315--------------------
    Rem oDoc.WImage 60, 80, 50, 100, "Logo"
    
    Rem oDoc.WTextBox 50, 70, 10, 450, "PAGARE", "F4", 11, hCenter, , vbBlack
    Rem oDoc.WTextBox 65, 60, 730, 490, "", "F4", 11, hCenter, , vbBlack, 2
    
    Rem oDoc.WTextBox 80, 80, 13, 150, "LUGAR DE EMISION", "F3", lnFuenteT, hCenter, , vbBlack, 1
    Rem oDoc.WTextBox 95, 80, 13, 150, sAgeCiudad, "F3", lnFuenteT, hCenter, , vbBlack
    
    Rem oDoc.WTextBox 80, 230, 13, 150, "FECHA DE EMISION", "F3", lnFuenteT, hCenter, , vbBlack, 1
    Rem oDoc.WTextBox 95, 230, 13, 150, gdFecSis, "F4", lnFuenteT, hCenter, , vbBlack
    
    Rem oDoc.WTextBox 80, 380, 13, 150, "NUMERO", "F3", lnFuenteT, hCenter, , vbBlack, 1
    Rem oDoc.WTextBox 95, 380, 13, 150, ActXCodCta.NroCuenta, "F4", lnFuenteT, hCenter, , vbBlack
    
    Rem oDoc.WTextBox 93, 80, 20, 150, "", "F3", lnFuenteT, hCenter, , vbBlack, 1
    Rem oDoc.WTextBox 93, 230, 20, 150, "", "F3", lnFuenteT, hCenter, , vbBlack, 1
    Rem oDoc.WTextBox 93, 380, 20, 150, "", "F3", lnFuenteT, hCenter, , vbBlack, 1
    
    Rem oDoc.WTextBox 113, 80, 13, 150, "FECHA DE VENCIMIENTO", "F3", lnFuenteT, hCenter, , vbBlack, 1
    Rem oDoc.WTextBox 128, 80, 13, 150, lblFecVencApr.Caption, "F4", lnFuenteT, hCenter, , vbBlack
    
    Rem oDoc.WTextBox 113, 230, 13, 150, "MONEDA PAGARE", "F3", lnFuenteT, hCenter, , vbBlack, 1
    Rem oDoc.WTextBox 128, 230, 13, 150, IIf(Left(ActXCodCta.Cuenta, 1) = 1, "SOLES", "DOLARES"), "F4", lnFuenteT, hCenter, , vbBlack
    
    Rem oDoc.WTextBox 113, 380, 13, 150, "IMPORTE PAGARE", "F3", lnFuenteT, hCenter, , vbBlack, 1
    Rem oDoc.WTextBox 128, 380, 13, 150, Format(lblMontoApr.Caption, gsFormatoNumeroView), "F4", lnFuenteT, hCenter, , vbBlack
    
    Rem oDoc.WTextBox 126, 80, 20, 150, "", "F3", lnFuenteT, hCenter, , vbBlack, 1
    Rem oDoc.WTextBox 126, 230, 20, 150, "", "F3", lnFuenteT, hCenter, , vbBlack, 1
    Rem oDoc.WTextBox 126, 380, 20, 150, "", "F3", lnFuenteT, hCenter, , vbBlack, 1
    
    Rem oDoc.WTextBox 150, 80, 10, 450, "Por éste PAGARE prometo (emos) pagar incondicionalmente a la Orden de la CAJA MUNICIPAL DE AHORRO Y CREDITO DE MAYNAS S.A. (LA CAJA) la cantidad de:", "F3", lnFuenteT, hLeft, , vbBlack
    Rem oDoc.WTextBox 180, 80, 20, 320, "", "F3", lnFuenteT, hLeft, , vbBlack, 1
    Rem oDoc.WTextBox 180, 400, 20, 130, "", "F3", lnFuenteT, hLeft, , vbBlack, 1
    
    Rem oDoc.WTextBox 185, 95, 20, 320, ConvNumLet(lblMontoApr.Caption, , True), "F4", lnFuenteT, hLeft, , vbBlack
    Rem oDoc.WTextBox 185, 415, 20, 130, IIf(Left(ActXCodCta.Cuenta, 1) = 1, "SOLES", "DOLARES"), "F4", lnFuenteT, hLeft, , vbBlack
    
    Rem oDoc.WTextBox 210, 80, 10, 450, "Importe a debitar en la siguiente cuenta de la Empresa que se indica:", "F3", lnFuenteT, hLeft, , vbBlack
    Rem oDoc.WTextBox 220, 80, 10, 450, "EMPRESA : CAJA MUNICIPAL DE AHORRO Y CREDITO DE MAYNAS S.A.", "F4", lnFuenteT, hLeft, , vbBlack
    Rem oDoc.WTextBox 230, 80, 10, 450, "OFICINA", "F4", lnFuenteT, hLeft, , vbBlack
    Rem oDoc.WTextBox 240, 80, 10, 450, "NUMERO DE CUENTA", "F4", lnFuenteT, hLeft, , vbBlack
    Rem oDoc.WTextBox 250, 80, 10, 450, "D.C. :", "F4", lnFuenteT, hLeft, , vbBlack
    
    Rem oDoc.WTextBox 280, 80, 10, 430, "1.", "F3", lnFuenteT, hLeft, , vbBlack
    Rem oDoc.WTextBox 290, 80, 10, 430, "2.", "F3", lnFuenteT, hLeft, , vbBlack
    Rem oDoc.WTextBox 310, 80, 10, 430, "3.", "F3", lnFuenteT, hLeft, , vbBlack
    Rem oDoc.WTextBox 330, 80, 10, 430, "4.", "F3", lnFuenteT, hLeft, , vbBlack
    Rem oDoc.WTextBox 340, 80, 10, 430, "5.", "F3", lnFuenteT, hLeft, , vbBlack
    Rem oDoc.WTextBox 365, 80, 10, 430, "6.", "F3", lnFuenteT, hLeft, , vbBlack
    Rem oDoc.WTextBox 375, 80, 10, 430, "7.", "F3", lnFuenteT, hLeft, , vbBlack
    Rem oDoc.WTextBox 390, 80, 10, 430, "8.", "F3", lnFuenteT, hLeft, , vbBlack
    Rem oDoc.WTextBox 410, 80, 10, 430, "9.", "F3", lnFuenteT, hLeft, , vbBlack
    
    Rem oDoc.WTextBox 270, 70, 10, 450, "Cláusulas Especiales:", "F3", lnFuenteT, hLeft, , vbBlack
    Rem oDoc.WTextBox 280, 90, 10, 430, "Este Pagaré debe ser pagado sólo en la misma moneda que expresa este título valor.", "F3", lnFuenteT, hLeft, , vbBlack
    Rem oDoc.WTextBox 290, 90, 10, 430, "Desde su último vencimiento, su importe total y/o cuotas, generará los intereses compensatorios más moratorios a las tasas máximas autorizadas o permitidas a su último Tenedor.", "F3", lnFuenteT, hLeft, , vbBlack
    Rem oDoc.WTextBox 310, 90, 10, 430, "De conformidad con lo dispuesto en los arts. 52 y 81 de la ley No 27287 el presente Pagaré “NO REQUIERE PROTESTO”, pudiendo ejercitarse las acciones cambiarias por el solo mérito del vencimiento del plazo pactado, o de sus renovaciones.", "F3", lnFuenteT, hLeft, , vbBlack
    Rem oDoc.WTextBox 330, 90, 10, 430, "El importe de éste Pagaré, podrá ser amortizado parcial o totalmente, renovándose el mismo por el monto del saldo.", "F3", lnFuenteT, hLeft, , vbBlack
    Rem oDoc.WTextBox 340, 90, 10, 430, "El importe de éste Pagaré y/o de sus cuotas, generará desde la emisión de éste Pagaré hasta la fecha de su respectivo vencimiento, un interés compensatorio a la tasa de " & nTasaIntCom & "% por año y a partir de su vencimiento se cobrará adicionalmente un interés moratorio de " & nTasaIntMor & "% por año.", "F3", lnFuenteT, hLeft, , vbBlack
    Rem oDoc.WTextBox 365, 90, 10, 430, "Los pagos que correspondan, podrán ser realizados a través de los canales de pago que LA CAJA pone a su disposición", "F3", lnFuenteT, hLeft, , vbBlack
    Rem oDoc.WTextBox 375, 90, 10, 430, "El (los) emitente(s) y aval(es), autorizamos a retirar de mi (nuestra) cuenta(s) de Ahorro o Plazo que en cualquier moneda mantenga(mos) en LA CAJA la suma para amortizar o pagar la presente obligación.", "F3", lnFuenteT, hLeft, , vbBlack
    Rem oDoc.WTextBox 390, 90, 10, 430, "Los intereses compensatorios y moratorios podrán variar de acuerdo a lo convenido en el contrato de Mutuo suscrito por las partes.", "F3", lnFuenteT, hLeft, , vbBlack
    Rem oDoc.WTextBox 410, 90, 10, 430, "Este Pagaré tiene naturaleza mercantil y se sujeta a las disposiciones de la Ley 27287 de Títulos Valores (Artículos 158º en adelante), la Ley de Bancos y al proceso ejecutivo señalado en el Código Procesal Civil en su caso.", "F3", lnFuenteT, hLeft, , vbBlack
    
    Rem oDoc.WTextBox 440, 90, 10, 150, "Emitente(s)", "F3", lnFuenteT, hLeft, , vbBlack
    Rem oDoc.WTextBox 580, 90, 10, 150, "Aval (es)", "F3", lnFuenteT, hLeft, , vbBlack
    
    Rem oDoc.WTextBox 450, 83, 20, 50, "NOMBRE", "F3", lnFuenteT, hLeft, , vbBlack
    Rem oDoc.WTextBox 450, 80, 20, 50, "", "F3", lnFuenteT, hLeft, , vbBlack, 1
    '----------------'END COMENTADO PTI1 20170315--------------------
    
   oDoc.WImage 50, 494, 35, 73, "Logo"
    
oDoc.WTextBox 30, 40, 15, 500, "PAGARÉ", "F2", 12, hCenter
oDoc.WTextBox 60, 45, 15, 175, "LUGAR DE EMISIÓN", "F1", 10, hCenter, vMiddle, vbBlack, 1, vbBlack
oDoc.WTextBox 60, 220, 15, 160, "FECHA DE EMISIÓN", "F1", 10, hCenter, vMiddle, vbBlack, 1, vbBlack
oDoc.WTextBox 60, 380, 15, 187, "NÚMERO", "F1", 10, hCenter, vMiddle, vbBlack, 1, vbBlack
oDoc.WTextBox 75, 45, 15, 175, sAgeCiudad, "F2", 10, hCenter, vMiddle, vbBlack, 1, vbBlack
                                                                             
oDoc.WTextBox 75, 220, 15, 160, "", "F2", 10, hCenter, vMiddle, vbBlack, 1, vbBlack
oDoc.WTextBox 75, 380, 15, 187, "", "F2", 10, hCenter, vMiddle, vbBlack, 1, vbBlack

oDoc.WTextBox 90, 45, 15, 175, "FECHA DE VENCIMIENTO", "F1", 10, hCenter, vMiddle, vbBlack, 1, vbBlack
oDoc.WTextBox 90, 220, 15, 160, "MONEDA PAGARÉ", "F1", 10, hCenter, vMiddle, vbBlack, 1, vbBlack
oDoc.WTextBox 90, 380, 15, 187, "IMPORTE PAGARÉ", "F1", 10, hCenter, vMiddle, vbBlack, 1, vbBlack
    
oDoc.WTextBox 105, 45, 15, 175, "", "F2", 10, hCenter, vMiddle, vbBlack, 1, vbBlack
oDoc.WTextBox 105, 220, 15, 160, IIf(Left(ActXCodCta.Cuenta, 1) = 1, StrConv(gcPEN_PLURAL, vbUpperCase), "DOLARES"), "F2", 10, hCenter, vMiddle, vbBlack, 1, vbBlack
oDoc.WTextBox 105, 380, 15, 187, Format(lblMontoApr.Caption, gsFormatoNumeroView), "F2", 10, hCenter, vMiddle, vbBlack, 1, vbBlack
oDoc.WTextBox 130, 45, 20, 480, "Por este", "F1", 11, hjustify
oDoc.WTextBox 131, 80, 20, 480, "PAGARÉ", "F2", 10, hjustify
oDoc.WTextBox 130, 117, 20, 480, "prometo/prometemos  pagar solidariamente e incondicionalmente a la orden de la", "F1", 11, hjustify
oDoc.WTextBox 131, 447, 20, 480, "CAJA MUNICIPAL DE AHORRO", "F2", 10, hjustify
oDoc.WTextBox 141, 45, 30, 480, "Y CRÉDITO DE MAYNAS S.A.", "F2", 10, hjustify
oDoc.WTextBox 140, 158, 30, 480, ", con R.U.C N° 20103845328, en adelante ", "F1", 11, hjustify
oDoc.WTextBox 141, 330, 30, 480, "LA CAJA", "F2", 10, hjustify
oDoc.WTextBox 140, 363, 30, 480, ", en cualquiera de sus oficinas a nivel nacional, o a", "F1", 11, hjustify
oDoc.WTextBox 150, 45, 30, 480, "quien", "F1", 11, hjustify
oDoc.WTextBox 151, 68, 30, 480, "LA CAJA", "F2", 10, hjustify
oDoc.WTextBox 150, 104, 30, 480, "hubiera endosado el presente título valor", "F1", 11, hjustify
oDoc.WTextBox 150, 267, 30, 600, ", la suma de:", "F1", 11, hjustify
    
oDoc.WTextBox 165, 45, 15, 360, "", "F3", lnFuenteT, hLeft, , vbBlack, 1
oDoc.WTextBox 165, 405, 15, 160, "", "F3", lnFuenteT, hLeft, , vbBlack, 1
    
oDoc.WTextBox 167, 47, 20, 320, ConvNumLet(lblMontoApr.Caption, , True), "F2", 10, hjustify
oDoc.WTextBox 167, 415, 20, 130, IIf(Left(ActXCodCta.Cuenta, 1) = 1, StrConv(gcPEN_PLURAL, vbUpperCase), "DOLARES"), "F2", 10, hjustify
oDoc.WTextBox 180, 45, 30, 480, "importe" & String(0.55, vbTab) & " de" & String(0.55, vbTab) & " dinero" & String(0.55, vbTab) & " que" & String(0.55, vbTab) & " expresamente" & String(0.55, vbTab) & " declaro/declaramos adeudar a " & String(13, vbTab) & " y que me(nos) obligo/obligamos a pagar en la misma moneda antes expresada en la fecha de vencimiento consignada.", "F1", 11, hjustify
oDoc.WTextBox 181, 340, 30, 480, "LA CAJA", "F2", 10, hjustify
'-------------------------------------------------------------------------------------------
oDoc.WTextBox 210, 45, 30, 540, "Queda " & String(0.5, vbTab) & "expresamente" & String(0.55, vbTab) & " estipulado" & String(0.55, vbTab) & " que" & String(0.55, vbTab) & " el" & String(0.55, vbTab) & " importe" & String(0.55, vbTab) & " de" & String(0.55, vbTab) & " este" & String(0.55, vbTab) & " Pagaré " & String(0.55, vbTab) & "devengará" & String(0.55, vbTab) & " desde" & String(0.55, vbTab) & " su" & String(0.55, vbTab) & " fecha" & String(0.55, vbTab) & " de" & String(0.55, vbTab) & " emisión " & String(0.55, vbTab) & "hasta" & String(0.55, vbTab) & " la" & String(0.55, vbTab) & " fecha" & String(0.55, vbTab) & " de " & String(0.55, vbTab) & "", "F1", 11, hjustify
oDoc.WTextBox 220, 45, 30, 540, "su" & String(0.55, vbTab) & " vencimiento" & String(0.55, vbTab) & " un" & String(0.55, vbTab) & " interés" & String(0.55, vbTab) & " compensatorio" & String(0.55, vbTab) & " a" & String(0.55, vbTab) & " una" & String(0.55, vbTab) & " tasa" & String(0.55, vbTab) & " efectiva" & String(0.55, vbTab) & " anual" & String(0.55, vbTab) & " del" & String(0.55, vbTab) & "", "F1", 11, hjustify
oDoc.WTextBox 220, 394, 30, 520, "y" & String(0.55, vbTab) & " a" & String(0.55, vbTab) & "  partir" & String(0.55, vbTab) & "  de" & String(0.55, vbTab) & " su" & String(0.55, vbTab) & "  vencimiento" & String(0.55, vbTab) & " se" & String(0.55, vbTab) & " cobrará", "F1", 11, hjustify
oDoc.WTextBox 221, 358, 30, 520, "_______", "F2", 10, hjustify
oDoc.WTextBox 230, 45, 30, 520, "adicionalmente" & String(0.54, vbTab) & "un" & String(0.54, vbTab) & "interés" & String(0.54, vbTab) & "moratorio" & String(0.54, vbTab) & "a" & String(0.54, vbTab) & " una" & String(0.54, vbTab) & " tasa" & String(0.54, vbTab) & " efectiva" & String(0.54, vbTab) & " anual" & String(0.54, vbTab) & " del" & String(0.54, vbTab) & "", "F1", 11, hjustify
oDoc.WTextBox 231, 320, 30, 520, "_______", "F2", 10, hjustify
oDoc.WTextBox 230, 358, 30, 400, " Ambas" & String(0.54, vbTab) & "tasas" & String(0.54, vbTab) & "de" & String(0.54, vbTab) & "interés" & String(0.54, vbTab) & " continuarán" & String(0.54, vbTab) & "devengándose", "F1", 11, hjustify
oDoc.WTextBox 240, 45, 30, 520, "por todo el tiempo que demore el pago de la presente obligación.", "F1", 11, hjustify
'-------------------------------------------------------------------------------------------
oDoc.WTextBox 260, 45, 30, 515, "Asimismo " & String(0.51, vbTab) & " autorizo(amos) " & String(0.51, vbTab) & " de " & String(0.51, vbTab) & " manera" & String(0.51, vbTab) & " expresa " & String(0.51, vbTab) & " el cobro" & String(0.51, vbTab) & " de penalidades, seguros, gastos " & String(0.51, vbTab) & " notariales, de " & String(0.51, vbTab) & " cobranza judicial y " & String(300, vbTab) & "", "F1", 11, hjustify
oDoc.WTextBox 270, 45, 30, 520, String(2, vbTab) & " extrajudicial, y en" & String(0.54, vbTab) & " general" & String(0.54, vbTab) & "los gastos" & String(0.54, vbTab) & "y comisiones que pudiéramos adeudar derivados del crédito representado en este", "F1", 11, hjustify
oDoc.WTextBox 270, 535, 30, 520, "Pagaré,", "F1", 11, hjustify
oDoc.WTextBox 280, 45, 30, 540, "y que se pudieran generar desde la fecha de emisión del presente Pagaré hasta la cancelación total de la presente obligación,", "F1", 11, hjustify
oDoc.WTextBox 280, 554, 30, 540, "sin", "F1", 11, hjustify
oDoc.WTextBox 290, 45, 30, 540, "que" & String(0.55, vbTab) & "sea necesario" & String(0.55, vbTab) & " requerimiento" & String(0.55, vbTab) & " alguno" & String(0.55, vbTab) & " de" & String(0.55, vbTab) & " pago " & String(0.55, vbTab) & "para", "F1", 11, hjustify
oDoc.WTextBox 290, 278, 30, 540, "constituirme/constituirnos" & String(1, vbTab) & " en" & String(2, vbTab) & " mora," & String(0.54, vbTab) & " pues" & String(2, vbTab) & " es" & String(0.54, vbTab) & " entendido" & String(0.54, vbTab) & " que" & String(0.54, vbTab) & " ésta se ", "F1", 11, hjustify
oDoc.WTextBox 300, 45, 30, 540, "producirá de modo automático por el solo hecho del vencimiento de éste Pagaré.", "F1", 11, hjustify
'-------------------------------------------------------------------------------------------
oDoc.WTextBox 320, 45, 30, 540, "Expresamente" & String(0.55, vbTab) & " acepto(amos) toda" & String(1, vbTab) & " variación" & String(1, vbTab) & " de" & String(1, vbTab) & " las " & String(0.5, vbTab) & "tasas" & String(0.5, vbTab) & " de interés, dentro de los límites legales autorizados, las mismas que se ", "F1", 11, hjustify
oDoc.WTextBox 330, 45, 30, 540, "aplicarán" & String(0.55, vbTab) & " luego" & String(0.55, vbTab) & " de" & String(0.55, vbTab) & " la" & String(0.55, vbTab) & " comunicación" & String(0.55, vbTab) & " efectuada" & String(0.55, vbTab) & " por" & String(0.55, vbTab) & " la ", "F1", 11, hjustify
oDoc.WTextBox 331, 274, 30, 480, "LA CAJA", "F2", 10, hjustify
oDoc.WTextBox 330, 308, 30, 480, ", conforme a ley. Se" & String(0.55, vbTab) & " deja constancia que el presente Pagaré " & """" & "no", "F1", 11, hjustify
oDoc.WTextBox 340, 45, 30, 540, "requiere" & String(0.6, vbTab) & " ser" & String(0.55, vbTab) & " protestado" & """" & " por" & String(1.4, vbTab) & " falta" & String(1.4, vbTab) & " de" & String(1.4, vbTab) & " pago, procediendo" & String(1.4, vbTab) & "su ejecución" & String(0.55, vbTab) & " por el solo mérito del vencimiento del plazo pactado, o de", "F1", 11, hjustify
oDoc.WTextBox 350, 45, 30, 520, "sus renovaciones o prórrogas de ser el caso.", "F1", 11, hjustify
'-------------------------------------------------------------------------------------------
oDoc.WTextBox 370, 45, 30, 540, "De acuerdo" & String(0.55, vbTab) & "a" & String(0.55, vbTab) & " lo dispuesto en el numeral 11) del artículo 132° de la Ley General del Sistema Financiero y del Sistema de", "F1", 11, hjustify
oDoc.WTextBox 370, 533, 30, 540, "Seguros ", "F1", 11, hjustify
oDoc.WTextBox 380, 45, 30, 520, "y Orgánica " & String(0.5, vbTab) & "de" & String(0.5, vbTab) & " la" & String(0.55, vbTab) & "Superintendencia" & String(0.55, vbTab) & "de" & String(0.55, vbTab) & " Banca y Seguros, autorizo(amos) a la", "F1", 11, hjustify
oDoc.WTextBox 381, 355, 30, 480, "LA CAJA", "F2", 10, hjustify
oDoc.WTextBox 380, 392, 30, 480, "para" & String(0.55, vbTab) & " que compense entre mis acreencias y ", "F1", 11, hjustify
oDoc.WTextBox 390, 45, 30, 540, "activos (cuentas, valores, depósitos en general, entre otros) que" & String(0.55, vbTab) & " mantenga en su poder, hasta por el importe" & String(0.55, vbTab) & " de éste pagaré más", "F1", 11, hjustify
oDoc.WTextBox 400, 45, 30, 540, "los intereses compensatorios, moratorios, gastos y cualquier otro concepto antes detallado en el presente título valor.", "F1", 11, hjustify
'-------------------------------------------------------------------------------------------
oDoc.WTextBox 420, 45, 30, 530, "De" & String(0.55, vbTab) & " conformidad" & String(0.55, vbTab) & " con" & String(0.55, vbTab) & " el" & String(0.55, vbTab) & " artículo" & String(0.55, vbTab) & " 1233°" & String(0.55, vbTab) & " del" & String(0.55, vbTab) & " Código" & String(0.55, vbTab) & " Civil, acepto(amos)" & String(0.55, vbTab) & " que" & String(0.55, vbTab) & " en" & String(0.55, vbTab) & " la" & String(0.55, vbTab) & " eventualidad" & String(0.55, vbTab) & " de" & String(0.55, vbTab) & " que" & String(0.55, vbTab) & " el" & String(0.55, vbTab) & " presente" & String(0.55, vbTab) & " título se ", "F1", 11, hjustify
oDoc.WTextBox 430, 562, 30, 530, "o", "F1", 11, hjustify
oDoc.WTextBox 440, 45, 30, 540, "destrucción" & String(0.55, vbTab) & " parcial, deterioro" & String(0.55, vbTab) & " total, extravío" & String(0.55, vbTab) & " y sustracción, se aplicará lo dispuesto en los artículos 101° al 107° de la Ley No.27287, en lo que resultase pertinente.", "F1", 11, hjustify
oDoc.WTextBox 430, 45, 30, 525, "perjudicara" & String(0.55, vbTab) & "por" & String(0.55, vbTab) & "cualquier" & String(0.55, vbTab) & "causa, tal" & String(1, vbTab) & "hecho" & String(1, vbTab) & "no extinguirá la obligación primitiva" & String(0.51, vbTab) & "u original. Asimismo, en caso de deterioro notable", "F1", 11, hjustify
'-------------------------------------------------------------------------------------------
oDoc.WTextBox 470, 45, 30, 540, "Me(nos)" & String(0.55, vbTab) & " someto(emos) expresamente" & String(0.55, vbTab) & " a" & String(0.55, vbTab) & " la" & String(0.55, vbTab) & " competencia" & String(0.55, vbTab) & " y" & String(0.55, vbTab) & " tribunales" & String(0.55, vbTab) & "de" & String(0.55, vbTab) & "esta ciudad, en" & String(0.55, vbTab) & " cuyo" & String(0.55, vbTab) & " efecto" & String(0.55, vbTab) & " renuncio/renunciamos" & String(0.55, vbTab) & "al ", "F1", 11, hjustify
oDoc.WTextBox 480, 45, 30, 540, "fuero de mi/nuestro domicilio. Señalo(amos) como domicilio aquel" & String(0.55, vbTab) & " indicado" & String(0.55, vbTab) & " en" & String(0.55, vbTab) & " este pagaré, a donde se efectuarán las diligencias", "F1", 11, hjustify
oDoc.WTextBox 490, 45, 30, 540, "notariales, judiciales y demás que fuesen necesarias para lo que", "F1", 11, hjustify
oDoc.WTextBox 491, 306, 30, 480, "LA CAJA", "F2", 10, hjustify
oDoc.WTextBox 490, 342, 30, 520, "considere pertinente. Cualquier cambio de domicilio que", "F1", 11, hjustify
oDoc.WTextBox 500, 45, 30, 540, "haga(mos), para su validez, lo haré(mos) mediante carta notarial y conforme a lo dispuesto en el artículo 40° del Código Civil.", "F1", 11, hjustify
'-------------------------------------------------------------------------------------------
oDoc.WTextBox 520, 45, 30, 530, "Declaro(amos)" & String(2, vbTab) & " estar" & String(2, vbTab) & " plenamente" & String(2, vbTab) & " facultado(s)" & String(2, vbTab) & " para" & String(2, vbTab) & " suscribir" & String(2, vbTab) & " y" & String(2, vbTab) & " emitir" & String(2, vbTab) & "  el" & String(1, vbTab) & " presente" & String(1, vbTab) & " Pagaré, asumiendo", "F1", 11, hjustify
oDoc.WTextBox 520, 492, 30, 480, "en" & String(1, vbTab) & " caso" & String(1, vbTab) & " contrario", "F1", 11, hjustify
oDoc.WTextBox 530, 45, 30, 540, "responsabilidad civil y/o penal a que hubiera lugar. Se deja constancia que la información proporcionada por el(los) emitente(s) en", "F1", 11, hjustify
oDoc.WTextBox 540, 45, 30, 540, "el presente documento, tiene" & String(0.4, vbTab) & " el" & String(0.54, vbTab) & " carácter de declaración jurada, de acuerdo con el artículo 179° de la Ley No. 26702 - Ley General del Sistema Financiero y del Sistema de Seguros y Orgánica de la Superintendencia de Banca y Seguros.", "F1", 11, hjustify
oDoc.WTextBox 570, 45, 30, 520, "Suscribimos el presente en señal de conformidad.", "F1", 11, hjustify
'-------------------------------------------------------------------------------------------

oDoc.WTextBox 590, 45, 15, 520, "Nombres y Apellidos/Razón Social:", "F2", 11
Dim h As Integer
Dim nTitu As Integer
Dim nCode As Integer
Dim RelaGar As COMDPersona.DCOMPersonas
Set RelaGar = New COMDPersona.DCOMPersonas
Dim RrelGar As ADODB.Recordset
h = 160
Dim sPersCodR As String
Set oRs = oCF.ObtieneTitAvaPagare(ActXCodCta.NroCuenta)
If Not (oRs.EOF And oRs.BOF) Then
    Do Until oRs.EOF
        If oRs!nPrdPersRelac = gColRelPersTitular And (oRs!nPersPersoneria = 3 Or oRs!nPersPersoneria = 2) Then

            oDoc.WTextBox 460 + h, 45, 15, 50, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
            oDoc.WTextBox 470 + h, 45, 20, 250, "Apoderado:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
            oDoc.WTextBox 495 + h, 45, 15, 50, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
            oDoc.WTextBox 513 + h, 45, 35, 50, "Dirección:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
            oDoc.WTextBox 435 + h, 45, 35, 205, oRs!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
            oDoc.WTextBox 460 + h, 90, 15, 205, IIf(IsNull(oRs!cPersIDnro), oRs!Ruc, oRs!cPersIDnro), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
            oDoc.WTextBox 513 + h, 95, 35, 205, oRs!cPersDireccDomicilio, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 4

            oDoc.WTextBox 575 + h, 45, 15, 255, "Firma:_________________________", "F1", 11
            oDoc.WTextBox 545 + h, 220, 90, 65, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
            nTitu = 1
            
            ElseIf oRs!nPrdPersRelac = gColRelPersTitular Then
'
            oDoc.WTextBox 460 + h, 45, 15, 50, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3 '
            oDoc.WTextBox 475 + h, 45, 35, 50, "Dirección:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3 '

            oDoc.WTextBox 435 + h, 45, 35, 205, oRs!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
            oDoc.WTextBox 460 + h, 90, 15, 205, IIf(IsNull(oRs!cPersIDnro), oRs!Ruc, oRs!cPersIDnro), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
            oDoc.WTextBox 475 + h, 95, 35, 205, oRs!cPersDireccDomicilio, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 4

            oDoc.WTextBox 570 + h, 45, 15, 255, "Firma:_________________________", "F1", 11
            oDoc.WTextBox 540 + h, 220, 90, 65, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
            nTitu = 1

            nCode = 1

            ElseIf (oRs!nPrdPersRelac = gColRelPersConyugue Or oRs!nPrdPersRelac = gColRelPersCodeudor) And (oRs!nPersPersoneria = 3 Or oRs!nPersPersoneria = 2) Then
            sPersCodR = oRs!cperscod
                Set RrelGar = RelaGar.RecuperarDatosPersonaGar(sPersCodR)
                If Not (RrelGar.BOF And RrelGar.EOF) Then
                    While Not RrelGar.EOF
                       oDoc.WTextBox 475 + h, 330, 35, 205, RrelGar!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       oDoc.WTextBox 485 + h, 375, 35, 205, RrelGar!cPersIDnro, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       RrelGar.MoveNext
                    Wend
                Else
                   
                End If
            oDoc.WTextBox 590, 330, 15, 520, "Nombres y Apellidos/Razón Social:", "F2", 11
            oDoc.WTextBox 435 + h, 330, 35, 205, oRs!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
            oDoc.WTextBox 460 + h, 375, 15, 205, IIf(IsNull(oRs!cPersIDnro), oRs!Ruc, oRs!cPersIDnro), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
            oDoc.WTextBox 513 + h, 380, 35, 205, oRs!cPersDireccDomicilio, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 4
            
            oDoc.WTextBox 460 + h, 330, 15, 250, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
'
            oDoc.WTextBox 470 + h, 330, 20, 250, "Apoderado:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
'
            oDoc.WTextBox 495 + h, 330, 15, 50, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
'
                                                           
            oDoc.WTextBox 513 + h, 330, 35, 50, "Dirección:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
'

            oDoc.WTextBox 575 + h, 330, 15, 255, "Firma:_________________________", "F1", 11
            oDoc.WTextBox 545 + h, 490, 90, 65, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack

            nCode = 1
            
            ElseIf (oRs!nPrdPersRelac = gColRelPersConyugue Or oRs!nPrdPersRelac = gColRelPersCodeudor) Then
            sPersCodR = oRs!cperscod
                Set RrelGar = RelaGar.RecuperarDatosPersonaGar(sPersCodR)
                If Not (RrelGar.BOF And RrelGar.EOF) Then
                    While Not RrelGar.EOF
                       oDoc.WTextBox 475 + h, 330, 35, 205, RrelGar!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       oDoc.WTextBox 485 + h, 375, 35, 205, RrelGar!cPersIDnro, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       RrelGar.MoveNext
                    Wend
                Else
                   
                End If
            oDoc.WTextBox 590, 330, 15, 520, "Nombres y Apellidos/Razón Social:", "F2", 11
            oDoc.WTextBox 435 + h, 330, 35, 205, oRs!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
            oDoc.WTextBox 460 + h, 375, 15, 205, IIf(IsNull(oRs!cPersIDnro), oRs!Ruc, oRs!cPersIDnro), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
            oDoc.WTextBox 475 + h, 380, 35, 205, oRs!cPersDireccDomicilio, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 4
            
            
            oDoc.WTextBox 460 + h, 330, 15, 250, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
            oDoc.WTextBox 475 + h, 330, 35, 50, "Dirección:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3

            oDoc.WTextBox 570 + h, 330, 15, 255, "Firma:_________________________", "F1", 11
            oDoc.WTextBox 540 + h, 490, 90, 65, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack

            nCode = 1
'
            ElseIf oRs!nPrdPersRelac = gColRelPersRepresentante Then
            oDoc.WTextBox 473 + h, 45, 35, 205, oRs!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
            oDoc.WTextBox 485 + h, 90, 35, 250, IIf(IsNull(oRs!cPersIDnro), oRs!Ruc, oRs!cPersIDnro), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3

            nCode = 1
            Exit Do
        End If
        oRs.MoveNext
    Loop
End If
Dim sNroCuenta As String
sNroCuenta = ActXCodCta.NroCuenta
ImprimeGarantesPagare sNroCuenta, oDoc, RelaGar

    
    Set oCF = Nothing
    Set oRs = Nothing
    oDoc.PDFClose
    oDoc.Show
    Exit Sub
ErrorImprimirPDF:
    MsgBox Err.Description, vbInformation, "Aviso"
End Sub
'RECO FIN********************************************************'PTI1

'********************************* PTI1 *********************************************
'ActXCodCta.NroCuenta
Private Function ImprimeGarantesPagare(ByVal sNroCuenta As String, ByVal oDoc, ByVal RelaGar)
Dim RsGarantes As New ADODB.Recordset
Dim oDCred As COMdCredito.DCOMCredito
Set oDCred = New COMdCredito.DCOMCredito
Set RsGarantes = oDCred.RecuperaGarantes(sNroCuenta)
Dim sPersCodR As String
Dim RrelGar As ADODB.Recordset
Dim h As Integer
h = -160
           
Dim nGaran As Integer
nGaran = 0
'oDoc.NewPage A4_Vertical
If Not (RsGarantes.EOF And RsGarantes.BOF) Then
        While Not RsGarantes.EOF
        
       oDoc.WImage 50, 494, 35, 73, "Logo"
        
            '##################################### 1 Garante ##############################################
            If nGaran = 0 And (RsGarantes!nPersPersoneria = 3 Or RsGarantes!nPersPersoneria = 2) Then
            h = -170
            oDoc.NewPage A4_Vertical
            oDoc.WImage 50, 494, 35, 73, "Logo"
            
                sPersCodR = RsGarantes!cperscod
                Set RrelGar = RelaGar.RecuperarDatosPersonaGar(sPersCodR)
                If Not (RrelGar.BOF And RrelGar.EOF) Then
                    While Not RrelGar.EOF
                       oDoc.WTextBox 283 + h, 45, 35, 205, RrelGar!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       oDoc.WTextBox 295 + h, 90, 35, 205, RrelGar!cPersIDnro, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       RrelGar.MoveNext
                    Wend
                Else
                   
                End If
                            
                
             oDoc.WTextBox 245 + h, 45, 35, 205, RsGarantes!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 270 + h, 90, 15, 205, IIf(RsGarantes!DNI <> "", RsGarantes!DNI, RsGarantes!Ruc), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 320 + h, 95, 35, 205, IIf(IsNull(RsGarantes!cPersDireccDomicilio), "", RsGarantes!cPersDireccDomicilio), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 4
                '1
             oDoc.WTextBox 50, 45, 15, 450, "FIADOR SOLIDARIO", "F2", 11
             oDoc.WTextBox 70, 45, 15, 450, "Nombres y Apellidos/Razón Social:", "F2", 11
             oDoc.WTextBox 270 + h, 45, 15, 250, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
             oDoc.WTextBox 280 + h, 45, 20, 60, "Apoderado:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 305 + h, 45, 15, 50, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
             oDoc.WTextBox 320 + h, 45, 35, 50, "Dirección:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 380 + h, 45, 15, 255, "Firma:___________________________", "F1", 11
             oDoc.WTextBox 360 + h, 220, 90, 65, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
            ElseIf nGaran = 0 Then
            h = -170
            oDoc.NewPage A4_Vertical
            
            sPersCodR = RsGarantes!cperscod
                Set RrelGar = RelaGar.RecuperarDatosPersonaGar(sPersCodR)
                If Not (RrelGar.BOF And RrelGar.EOF) Then
                    While Not RrelGar.EOF
                       oDoc.WTextBox 283 + h, 45, 35, 205, RrelGar!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       oDoc.WTextBox 295 + h, 90, 35, 205, RrelGar!cPersIDnro, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       RrelGar.MoveNext
                    Wend
                Else
                   
                End If
                            
                
             oDoc.WTextBox 245 + h, 45, 35, 205, RsGarantes!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 270 + h, 90, 15, 205, IIf(RsGarantes!DNI <> "", RsGarantes!DNI, RsGarantes!Ruc), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 285 + h, 95, 35, 205, IIf(IsNull(RsGarantes!cPersDireccDomicilio), "", RsGarantes!cPersDireccDomicilio), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 4
                '1
             oDoc.WTextBox 50, 45, 15, 450, "FIADOR SOLIDARIO", "F2", 11
             oDoc.WTextBox 70, 45, 15, 450, "Nombres y Apellidos/Razón Social:", "F2", 11
             oDoc.WTextBox 270 + h, 45, 15, 250, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
             oDoc.WTextBox 285 + h, 45, 35, 50, "Dirección:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 380 + h, 45, 15, 255, "Firma:___________________________", "F1", 11
             oDoc.WTextBox 360 + h, 220, 90, 65, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
             '##################################### 2 Garante ##############################################
            ElseIf nGaran = 1 And (RsGarantes!nPersPersoneria = 3 Or RsGarantes!nPersPersoneria = 2) Then
            sPersCodR = RsGarantes!cperscod
                Set RrelGar = RelaGar.RecuperarDatosPersonaGar(sPersCodR)
                If Not (RrelGar.BOF And RrelGar.EOF) Then
                    While Not RrelGar.EOF
                       oDoc.WTextBox 283 + h, 300, 35, 205, RrelGar!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       oDoc.WTextBox 295 + h, 345, 35, 205, RrelGar!cPersIDnro, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       RrelGar.MoveNext
                    Wend
                Else
                    
                End If
                oDoc.WTextBox 245 + h, 300, 35, 205, RsGarantes!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 270 + h, 345, 15, 205, IIf(RsGarantes!DNI <> "", RsGarantes!DNI, RsGarantes!Ruc), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 320 + h, 350, 35, 205, IIf(IsNull(RsGarantes!cPersDireccDomicilio), "", RsGarantes!cPersDireccDomicilio), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 4
                 '2
             oDoc.WTextBox 50, 300, 15, 450, "FIADOR SOLIDARIO", "F2", 11
             oDoc.WTextBox 70, 300, 15, 450, "Nombres y Apellidos/Razón Social:", "F2", 11
             oDoc.WTextBox 270 + h, 300, 15, 250, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
             oDoc.WTextBox 280 + h, 300, 20, 60, "Apoderado:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 305 + h, 300, 15, 50, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
             oDoc.WTextBox 320 + h, 300, 35, 50, "Dirección:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 380 + h, 300, 15, 255, "Firma:___________________________", "F1", 11
             oDoc.WTextBox 360 + h, 490, 90, 65, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
             
            ElseIf nGaran = 1 Then
            
            sPersCodR = RsGarantes!cperscod
                Set RrelGar = RelaGar.RecuperarDatosPersonaGar(sPersCodR)
                If Not (RrelGar.BOF And RrelGar.EOF) Then
                    While Not RrelGar.EOF
                       oDoc.WTextBox 283 + h, 300, 35, 205, RrelGar!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       oDoc.WTextBox 295 + h, 345, 35, 205, RrelGar!cPersIDnro, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       RrelGar.MoveNext
                    Wend
                Else
                    
                End If
                oDoc.WTextBox 245 + h, 300, 35, 205, RsGarantes!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 270 + h, 345, 15, 205, IIf(RsGarantes!DNI <> "", RsGarantes!DNI, RsGarantes!Ruc), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 285 + h, 350, 35, 205, IIf(IsNull(RsGarantes!cPersDireccDomicilio), "", RsGarantes!cPersDireccDomicilio), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 4
                 '2
             oDoc.WTextBox 50, 300, 15, 450, "FIADOR SOLIDARIO", "F2", 11
             oDoc.WTextBox 70, 300, 15, 450, "Nombres y Apellidos/Razón Social:", "F2", 11
             oDoc.WTextBox 270 + h, 300, 15, 250, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
             oDoc.WTextBox 285 + h, 300, 35, 50, "Dirección:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 380 + h, 300, 15, 255, "Firma:___________________________", "F1", 11
             oDoc.WTextBox 360 + h, 490, 90, 65, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
             '##################################### 3 Garante ##############################################
            ElseIf nGaran = 2 And (RsGarantes!nPersPersoneria = 3 Or RsGarantes!nPersPersoneria = 2) Then
            
            sPersCodR = RsGarantes!cperscod
                Set RrelGar = RelaGar.RecuperarDatosPersonaGar(sPersCodR)
                If Not (RrelGar.BOF And RrelGar.EOF) Then
                    While Not RrelGar.EOF
                       oDoc.WTextBox 523 + h, 45, 35, 205, RrelGar!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       oDoc.WTextBox 535 + h, 90, 35, 205, RrelGar!cPersIDnro, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       RrelGar.MoveNext
                    Wend
                Else
                    
                End If
                oDoc.WTextBox 485 + h, 45, 35, 205, RsGarantes!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 510 + h, 90, 15, 205, IIf(RsGarantes!DNI <> "", RsGarantes!DNI, RsGarantes!Ruc), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 560 + h, 95, 35, 180, IIf(IsNull(RsGarantes!cPersDireccDomicilio), "", RsGarantes!cPersDireccDomicilio), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 4
                '2
              oDoc.WTextBox 290, 45, 15, 450, "FIADOR SOLIDARIO", "F2", 11
              oDoc.WTextBox 310, 45, 15, 450, "Nombres y Apellidos/Razón Social:", "F2", 11
             oDoc.WTextBox 510 + h, 45, 15, 50, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
             oDoc.WTextBox 520 + h, 45, 20, 60, "Apoderado:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 545 + h, 45, 15, 50, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
             oDoc.WTextBox 560 + h, 45, 35, 50, "Dirección:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 625 + h, 45, 15, 255, "Firma:___________________________", "F1", 11
             oDoc.WTextBox 605 + h, 220, 90, 65, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
            ElseIf nGaran = 2 Then
            h = -170
            sPersCodR = RsGarantes!cperscod
                Set RrelGar = RelaGar.RecuperarDatosPersonaGar(sPersCodR)
                If Not (RrelGar.BOF And RrelGar.EOF) Then
                    While Not RrelGar.EOF
                       oDoc.WTextBox 523 + h, 45, 35, 205, RrelGar!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       oDoc.WTextBox 535 + h, 90, 35, 205, RrelGar!cPersIDnro, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       RrelGar.MoveNext
                    Wend
                Else
                    
                End If
                oDoc.WTextBox 485 + h, 45, 35, 205, RsGarantes!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 510 + h, 90, 15, 205, IIf(RsGarantes!DNI <> "", RsGarantes!DNI, RsGarantes!Ruc), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 525 + h, 95, 35, 180, IIf(IsNull(RsGarantes!cPersDireccDomicilio), "", RsGarantes!cPersDireccDomicilio), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 4
                '2
              oDoc.WTextBox 290, 45, 15, 450, "FIADOR SOLIDARIO", "F2", 11
              oDoc.WTextBox 310, 45, 15, 450, "Nombres y Apellidos/Razón Social:", "F2", 11
             oDoc.WTextBox 510 + h, 45, 15, 50, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
             oDoc.WTextBox 525 + h, 45, 35, 50, "Dirección:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 625 + h, 45, 15, 255, "Firma:___________________________", "F1", 11
             oDoc.WTextBox 605 + h, 220, 90, 65, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
             '##################################### 4 Garante ##############################################
           ElseIf nGaran = 3 And (RsGarantes!nPersPersoneria = 3 Or RsGarantes!nPersPersoneria = 2) Then
            sPersCodR = RsGarantes!cperscod
            Set RrelGar = RelaGar.RecuperarDatosPersonaGar(sPersCodR)
                If Not (RrelGar.BOF And RrelGar.EOF) Then
                    While Not RrelGar.EOF
                       oDoc.WTextBox 523 + h, 300, 35, 205, RrelGar!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       oDoc.WTextBox 535 + h, 345, 35, 205, RrelGar!cPersIDnro, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       RrelGar.MoveNext
                    Wend
                Else
                    
                End If
                oDoc.WTextBox 485 + h, 300, 35, 205, RsGarantes!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 510 + h, 345, 15, 205, IIf(RsGarantes!DNI <> "", RsGarantes!DNI, RsGarantes!Ruc), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 560 + h, 350, 35, 205, IIf(IsNull(RsGarantes!cPersDireccDomicilio), "", RsGarantes!cPersDireccDomicilio), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 4
                
                '8
             oDoc.WTextBox 290, 300, 15, 450, "FIADOR SOLIDARIO", "F2", 11
             oDoc.WTextBox 310, 300, 15, 450, "Nombres y Apellidos/Razón Social:", "F2", 11
             oDoc.WTextBox 510 + h, 300, 15, 250, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
             oDoc.WTextBox 520 + h, 300, 20, 60, "Apoderado:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 545 + h, 300, 15, 50, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
             oDoc.WTextBox 560 + h, 300, 35, 50, "Dirección:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 625 + h, 300, 15, 255, "Firma:___________________________", "F1", 11
             oDoc.WTextBox 605 + h, 490, 90, 65, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
             
             ElseIf nGaran = 3 Then
            sPersCodR = RsGarantes!cperscod
            Set RrelGar = RelaGar.RecuperarDatosPersonaGar(sPersCodR)
                If Not (RrelGar.BOF And RrelGar.EOF) Then
                    While Not RrelGar.EOF
                       oDoc.WTextBox 523 + h, 300, 35, 205, RrelGar!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       oDoc.WTextBox 535 + h, 345, 35, 205, RrelGar!cPersIDnro, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       RrelGar.MoveNext
                    Wend
                Else
                    
                End If
                oDoc.WTextBox 485 + h, 300, 35, 205, RsGarantes!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 510 + h, 345, 15, 205, IIf(RsGarantes!DNI <> "", RsGarantes!DNI, RsGarantes!Ruc), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 525 + h, 350, 35, 205, IIf(IsNull(RsGarantes!cPersDireccDomicilio), "", RsGarantes!cPersDireccDomicilio), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 4

                '4
             oDoc.WTextBox 290, 300, 15, 450, "FIADOR SOLIDARIO", "F2", 11
             oDoc.WTextBox 310, 300, 15, 450, "Nombres y Apellidos/Razón Social:", "F2", 11
             oDoc.WTextBox 510 + h, 300, 15, 250, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
             oDoc.WTextBox 525 + h, 300, 35, 50, "Dirección:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 625 + h, 300, 15, 255, "Firma:___________________________", "F1", 11
             oDoc.WTextBox 605 + h, 490, 90, 65, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
              '##################################### 5 Garante ##############################################
            ElseIf nGaran = 4 And (RsGarantes!nPersPersoneria = 3 Or RsGarantes!nPersPersoneria = 2) Then
            h = 110
            sPersCodR = RsGarantes!cperscod
                Set RrelGar = RelaGar.RecuperarDatosPersonaGar(sPersCodR)
                If Not (RrelGar.BOF And RrelGar.EOF) Then
                    While Not RrelGar.EOF
                       oDoc.WTextBox 503 + h, 45, 35, 205, RrelGar!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       oDoc.WTextBox 515 + h, 90, 35, 205, RrelGar!cPersIDnro, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       RrelGar.MoveNext
                    Wend
                Else
                    
                End If
                oDoc.WTextBox 465 + h, 45, 35, 205, RsGarantes!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 490 + h, 90, 15, 205, IIf(RsGarantes!DNI <> "", RsGarantes!DNI, RsGarantes!Ruc), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 542 + h, 95, 35, 180, IIf(IsNull(RsGarantes!cPersDireccDomicilio), "", RsGarantes!cPersDireccDomicilio), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 4
                '2
              oDoc.WTextBox 550, 45, 15, 450, "FIADOR SOLIDARIO", "F2", 11
              oDoc.WTextBox 570, 45, 15, 450, "Nombres y Apellidos/Razón Social:", "F2", 11
             oDoc.WTextBox 490 + h, 45, 15, 50, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
             oDoc.WTextBox 500 + h, 45, 20, 60, "Apoderado:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 525 + h, 45, 15, 50, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
             oDoc.WTextBox 542 + h, 45, 35, 50, "Dirección:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 605 + h, 45, 15, 255, "Firma:___________________________", "F1", 11
             oDoc.WTextBox 585 + h, 220, 90, 65, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack


            ElseIf nGaran = 4 Then
            h = 110
            sPersCodR = RsGarantes!cperscod
                Set RrelGar = RelaGar.RecuperarDatosPersonaGar(sPersCodR)
                If Not (RrelGar.BOF And RrelGar.EOF) Then
                    While Not RrelGar.EOF
                       oDoc.WTextBox 503 + h, 45, 35, 205, RrelGar!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       oDoc.WTextBox 515 + h, 90, 35, 205, RrelGar!cPersIDnro, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       RrelGar.MoveNext
                    Wend
                Else
                    
                End If
                oDoc.WTextBox 465 + h, 45, 35, 205, RsGarantes!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 490 + h, 90, 15, 205, IIf(RsGarantes!DNI <> "", RsGarantes!DNI, RsGarantes!Ruc), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 505 + h, 95, 35, 180, IIf(IsNull(RsGarantes!cPersDireccDomicilio), "", RsGarantes!cPersDireccDomicilio), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 4
                '2
              oDoc.WTextBox 550, 45, 15, 450, "FIADOR SOLIDARIO", "F2", 11
              oDoc.WTextBox 570, 45, 15, 450, "Nombres y Apellidos/Razón Social:", "F2", 11
             oDoc.WTextBox 490 + h, 45, 15, 50, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
             oDoc.WTextBox 505 + h, 45, 35, 50, "Dirección:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 605 + h, 45, 15, 255, "Firma:___________________________", "F1", 11
             oDoc.WTextBox 585 + h, 220, 90, 65, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
           '##################################### 6 Garante ##############################################
            ElseIf nGaran = 5 And (RsGarantes!nPersPersoneria = 3 Or RsGarantes!nPersPersoneria = 2) Then
            sPersCodR = RsGarantes!cperscod
            Set RrelGar = RelaGar.RecuperarDatosPersonaGar(sPersCodR)
                If Not (RrelGar.BOF And RrelGar.EOF) Then
                    While Not RrelGar.EOF
                       oDoc.WTextBox 503 + h, 300, 35, 205, RrelGar!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       oDoc.WTextBox 515 + h, 345, 35, 205, RrelGar!cPersIDnro, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       RrelGar.MoveNext
                    Wend
                Else
                    
                End If
                oDoc.WTextBox 465 + h, 300, 35, 205, RsGarantes!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 490 + h, 345, 15, 205, IIf(RsGarantes!DNI <> "", RsGarantes!DNI, RsGarantes!Ruc), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 538 + h, 350, 35, 205, IIf(IsNull(RsGarantes!cPersDireccDomicilio), "", RsGarantes!cPersDireccDomicilio), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 4
                
                '4
             oDoc.WTextBox 550, 300, 15, 450, "FIADOR SOLIDARIO", "F2", 11
             oDoc.WTextBox 570, 300, 15, 450, "Nombres y Apellidos/Razón Social:", "F2", 11
             oDoc.WTextBox 490 + h, 300, 15, 250, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
             oDoc.WTextBox 500 + h, 300, 20, 60, "Apoderado:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 525 + h, 300, 15, 50, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
             oDoc.WTextBox 538 + h, 300, 35, 50, "Dirección:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 605 + h, 300, 15, 255, "Firma:___________________________", "F1", 11
             oDoc.WTextBox 585 + h, 490, 90, 65, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
             
             ElseIf nGaran = 5 Then
            sPersCodR = RsGarantes!cperscod
            Set RrelGar = RelaGar.RecuperarDatosPersonaGar(sPersCodR)
                If Not (RrelGar.BOF And RrelGar.EOF) Then
                    While Not RrelGar.EOF
                       oDoc.WTextBox 503 + h, 300, 35, 205, RrelGar!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       oDoc.WTextBox 515 + h, 345, 35, 205, RrelGar!cPersIDnro, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       RrelGar.MoveNext
                    Wend
                Else
                    
                End If
                oDoc.WTextBox 465 + h, 300, 35, 205, RsGarantes!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 490 + h, 345, 15, 205, IIf(RsGarantes!DNI <> "", RsGarantes!DNI, RsGarantes!Ruc), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 505 + h, 350, 35, 205, IIf(IsNull(RsGarantes!cPersDireccDomicilio), "", RsGarantes!cPersDireccDomicilio), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 4
                
                '4
             oDoc.WTextBox 550, 300, 15, 450, "FIADOR SOLIDARIO", "F2", 11
             oDoc.WTextBox 570, 300, 15, 450, "Nombres y Apellidos/Razón Social:", "F2", 11
             oDoc.WTextBox 490 + h, 300, 15, 250, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
             oDoc.WTextBox 505 + h, 300, 35, 50, "Dirección:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 605 + h, 300, 15, 255, "Firma:___________________________", "F1", 11
             oDoc.WTextBox 585 + h, 490, 90, 65, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
        '##################################### 7 Garante ##############################################
        
            ElseIf nGaran = 6 And (RsGarantes!nPersPersoneria = 3 Or RsGarantes!nPersPersoneria = 2) Then
            h = -160
            oDoc.NewPage A4_Vertical
                sPersCodR = RsGarantes!cperscod
                Set RrelGar = RelaGar.RecuperarDatosPersonaGar(sPersCodR) '----Recupera el Representante
                If Not (RrelGar.BOF And RrelGar.EOF) Then
                    While Not RrelGar.EOF
                       oDoc.WTextBox 282 + h, 45, 35, 205, RrelGar!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       oDoc.WTextBox 295 + h, 90, 35, 205, RrelGar!cPersIDnro, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       RrelGar.MoveNext
                    Wend
                Else
                End If
                
             oDoc.WTextBox 245 + h, 45, 35, 205, RsGarantes!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 270 + h, 90, 15, 205, IIf(RsGarantes!DNI <> "", RsGarantes!DNI, RsGarantes!Ruc), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 320 + h, 95, 35, 205, IIf(IsNull(RsGarantes!cPersDireccDomicilio), "", RsGarantes!cPersDireccDomicilio), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 4
                '1
             oDoc.WTextBox 60, 45, 15, 450, "FIADOR SOLIDARIO", "F2", 11
             oDoc.WTextBox 80, 45, 15, 450, "Nombres y Apellidos/Razón Social:", "F2", 11 '
             oDoc.WTextBox 270 + h, 45, 15, 250, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3 '
             oDoc.WTextBox 280 + h, 45, 20, 60, "Apoderado:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3 '
             oDoc.WTextBox 305 + h, 45, 15, 50, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3 '
             oDoc.WTextBox 320 + h, 45, 35, 50, "Dirección:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
'
             oDoc.WTextBox 385 + h, 45, 15, 255, "Firma:___________________________", "F1", 11
             oDoc.WTextBox 365 + h, 220, 90, 65, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
             
            ElseIf nGaran = 6 Then
            oDoc.NewPage A4_Vertical
            h = -160
            sPersCodR = RsGarantes!cperscod
                Set RrelGar = RelaGar.RecuperarDatosPersonaGar(sPersCodR)
                If Not (RrelGar.BOF And RrelGar.EOF) Then
                    While Not RrelGar.EOF
                       oDoc.WTextBox 285 + h, 45, 35, 205, RrelGar!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       oDoc.WTextBox 295 + h, 90, 35, 205, RrelGar!cPersIDnro, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       RrelGar.MoveNext
                    Wend
                Else
                   
                End If
                
             oDoc.WTextBox 245 + h, 45, 35, 205, RsGarantes!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 270 + h, 90, 15, 205, IIf(RsGarantes!DNI <> "", RsGarantes!DNI, RsGarantes!Ruc), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 285 + h, 95, 35, 205, IIf(IsNull(RsGarantes!cPersDireccDomicilio), "", RsGarantes!cPersDireccDomicilio), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 4
                '1
             oDoc.WTextBox 60, 45, 15, 450, "FIADOR SOLIDARIO", "F2", 11
             oDoc.WTextBox 80, 45, 15, 450, "Nombres y Apellidos/Razón Social:", "F2", 11

             oDoc.WTextBox 270 + h, 45, 15, 250, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
             oDoc.WTextBox 285 + h, 45, 35, 50, "Dirección:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 385 + h, 45, 15, 255, "Firma:___________________________", "F1", 11
             oDoc.WTextBox 365 + h, 220, 90, 65, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
              '##################################### 8 Garante ##############################################
            ElseIf nGaran = 7 And (RsGarantes!nPersPersoneria = 3 Or RsGarantes!nPersPersoneria = 2) Then
            sPersCodR = RsGarantes!cperscod
                Set RrelGar = RelaGar.RecuperarDatosPersonaGar(sPersCodR)
                If Not (RrelGar.BOF And RrelGar.EOF) Then
                    While Not RrelGar.EOF
                       oDoc.WTextBox 282 + h, 300, 35, 205, RrelGar!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       oDoc.WTextBox 295 + h, 345, 35, 205, RrelGar!cPersIDnro, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       RrelGar.MoveNext
                    Wend
                Else
                    
                End If
                oDoc.WTextBox 245 + h, 300, 35, 205, RsGarantes!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 270 + h, 345, 15, 205, IIf(RsGarantes!DNI <> "", RsGarantes!DNI, RsGarantes!Ruc), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 320 + h, 350, 35, 205, IIf(IsNull(RsGarantes!cPersDireccDomicilio), "", RsGarantes!cPersDireccDomicilio), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 4
                 '2
             oDoc.WTextBox 60, 300, 15, 450, "FIADOR SOLIDARIO", "F2", 11
             oDoc.WTextBox 80, 300, 15, 450, "Nombres y Apellidos/Razón Social:", "F2", 11
             oDoc.WTextBox 270 + h, 300, 15, 250, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
             oDoc.WTextBox 280 + h, 300, 20, 60, "Apoderado:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 305 + h, 300, 15, 50, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
             oDoc.WTextBox 320 + h, 300, 35, 50, "Dirección:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 385 + h, 300, 15, 255, "Firma:___________________________", "F1", 11
             oDoc.WTextBox 365 + h, 490, 90, 65, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
             
            ElseIf nGaran = 7 Then
            sPersCodR = RsGarantes!cperscod
                Set RrelGar = RelaGar.RecuperarDatosPersonaGar(sPersCodR)
                If Not (RrelGar.BOF And RrelGar.EOF) Then
                    While Not RrelGar.EOF
                       oDoc.WTextBox 285 + h, 300, 35, 205, RrelGar!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       oDoc.WTextBox 295 + h, 345, 35, 205, RrelGar!cPersIDnro, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       RrelGar.MoveNext
                    Wend
                Else
                    
                End If
                oDoc.WTextBox 245 + h, 300, 35, 205, RsGarantes!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 270 + h, 345, 15, 205, IIf(RsGarantes!DNI <> "", RsGarantes!DNI, RsGarantes!Ruc), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 285 + h, 350, 35, 205, IIf(IsNull(RsGarantes!cPersDireccDomicilio), "", RsGarantes!cPersDireccDomicilio), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 4
                 '2
             oDoc.WTextBox 60, 300, 15, 450, "FIADOR SOLIDARIO", "F2", 11
             oDoc.WTextBox 80, 300, 15, 450, "Nombres y Apellidos/Razón Social:", "F2", 11

             oDoc.WTextBox 270 + h, 300, 15, 250, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3

             oDoc.WTextBox 285 + h, 300, 35, 50, "Dirección:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3

             oDoc.WTextBox 385 + h, 300, 15, 255, "Firma:___________________________", "F1", 11
             oDoc.WTextBox 365 + h, 490, 90, 65, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
          '##################################### 9 Garante ##############################################
            ElseIf nGaran = 8 And (RsGarantes!nPersPersoneria = 3 Or RsGarantes!nPersPersoneria = 2) Then
            h = -140
            sPersCodR = RsGarantes!cperscod
                Set RrelGar = RelaGar.RecuperarDatosPersonaGar(sPersCodR)
                If Not (RrelGar.BOF And RrelGar.EOF) Then
                    While Not RrelGar.EOF
                       oDoc.WTextBox 502 + h, 45, 35, 205, RrelGar!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       oDoc.WTextBox 515 + h, 90, 35, 205, RrelGar!cPersIDnro, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       RrelGar.MoveNext
                    Wend
                Else
                    
                End If
                oDoc.WTextBox 465 + h, 45, 35, 205, RsGarantes!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 490 + h, 90, 15, 205, IIf(RsGarantes!DNI <> "", RsGarantes!DNI, RsGarantes!Ruc), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 540 + h, 95, 35, 180, IIf(IsNull(RsGarantes!cPersDireccDomicilio), "", RsGarantes!cPersDireccDomicilio), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 4
                '2
              oDoc.WTextBox 300, 45, 15, 450, "FIADOR SOLIDARIO", "F2", 11
              oDoc.WTextBox 320, 45, 15, 450, "Nombres y Apellidos/Razón Social:", "F2", 11
             oDoc.WTextBox 490 + h, 45, 15, 50, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
             oDoc.WTextBox 500 + h, 45, 20, 60, "Apoderado:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 525 + h, 45, 15, 50, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
             oDoc.WTextBox 540 + h, 45, 35, 50, "Dirección:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 605 + h, 45, 15, 255, "Firma:___________________________", "F1", 11
             oDoc.WTextBox 585 + h, 220, 90, 65, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack


            ElseIf nGaran = 8 Then
            h = -140
            sPersCodR = RsGarantes!cperscod
                Set RrelGar = RelaGar.RecuperarDatosPersonaGar(sPersCodR)
                If Not (RrelGar.BOF And RrelGar.EOF) Then
                    While Not RrelGar.EOF
                       oDoc.WTextBox 505 + h, 45, 35, 205, RrelGar!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       oDoc.WTextBox 515 + h, 90, 35, 205, RrelGar!cPersIDnro, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       RrelGar.MoveNext
                    Wend
                Else
                    
                End If
                oDoc.WTextBox 465 + h, 45, 35, 205, RsGarantes!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 490 + h, 90, 15, 205, IIf(RsGarantes!DNI <> "", RsGarantes!DNI, RsGarantes!Ruc), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 505 + h, 95, 35, 180, IIf(IsNull(RsGarantes!cPersDireccDomicilio), "", RsGarantes!cPersDireccDomicilio), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 4
                '2
              oDoc.WTextBox 300, 45, 15, 450, "FIADOR SOLIDARIO", "F2", 11
              oDoc.WTextBox 320, 45, 15, 450, "Nombres y Apellidos/Razón Social:", "F2", 11
             oDoc.WTextBox 490 + h, 45, 15, 50, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
             oDoc.WTextBox 505 + h, 45, 35, 50, "Dirección:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 605 + h, 45, 15, 255, "Firma:___________________________", "F1", 11
             oDoc.WTextBox 585 + h, 220, 90, 65, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
           '##################################### 10 Garante ##############################################
            ElseIf nGaran = 9 And (RsGarantes!nPersPersoneria = 3 Or RsGarantes!nPersPersoneria = 2) Then
            h = -140
            sPersCodR = RsGarantes!cperscod
            Set RrelGar = RelaGar.RecuperarDatosPersonaGar(sPersCodR)
                If Not (RrelGar.BOF And RrelGar.EOF) Then
                    While Not RrelGar.EOF
                       oDoc.WTextBox 502 + h, 300, 35, 205, RrelGar!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       oDoc.WTextBox 515 + h, 345, 35, 205, RrelGar!cPersIDnro, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       RrelGar.MoveNext
                    Wend
                Else
                    
                End If
                oDoc.WTextBox 465 + h, 300, 35, 205, RsGarantes!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 490 + h, 345, 15, 205, IIf(RsGarantes!DNI <> "", RsGarantes!DNI, RsGarantes!Ruc), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 538 + h, 350, 35, 205, IIf(IsNull(RsGarantes!cPersDireccDomicilio), "", RsGarantes!cPersDireccDomicilio), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 4
                
                '4
             oDoc.WTextBox 300, 300, 15, 450, "FIADOR SOLIDARIO", "F2", 11
             oDoc.WTextBox 320, 300, 15, 450, "Nombres y Apellidos/Razón Social:", "F2", 11
             oDoc.WTextBox 490 + h, 300, 15, 250, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
             oDoc.WTextBox 500 + h, 300, 20, 60, "Apoderado:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 525 + h, 300, 15, 50, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
             oDoc.WTextBox 538 + h, 300, 35, 50, "Dirección:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 605 + h, 300, 15, 255, "Firma:___________________________", "F1", 11
             oDoc.WTextBox 585 + h, 490, 90, 65, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
             
             ElseIf nGaran = 9 Then
             h = -140
            sPersCodR = RsGarantes!cperscod
            Set RrelGar = RelaGar.RecuperarDatosPersonaGar(sPersCodR)
                If Not (RrelGar.BOF And RrelGar.EOF) Then
                    While Not RrelGar.EOF
                       oDoc.WTextBox 505 + h, 300, 35, 205, RrelGar!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       oDoc.WTextBox 515 + h, 345, 35, 205, RrelGar!cPersIDnro, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                       RrelGar.MoveNext
                    Wend
                Else
                    
                End If
                oDoc.WTextBox 465 + h, 300, 35, 205, RsGarantes!cPersNombre, "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 490 + h, 345, 15, 205, IIf(RsGarantes!DNI <> "", RsGarantes!DNI, RsGarantes!Ruc), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
                oDoc.WTextBox 505 + h, 350, 35, 205, IIf(IsNull(RsGarantes!cPersDireccDomicilio), "", RsGarantes!cPersDireccDomicilio), "F2", 11, hLeft, vMiddle, vbBlack, vbBlack, , 4
                
                '4
             oDoc.WTextBox 300, 300, 15, 450, "FIADOR SOLIDARIO", "F2", 11
             oDoc.WTextBox 320, 300, 15, 450, "Nombres y Apellidos/Razón Social:", "F2", 11
             oDoc.WTextBox 490 + h, 300, 15, 250, "D.O.I N°:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, 3
             oDoc.WTextBox 505 + h, 300, 35, 50, "Dirección:", "F1", 11, hLeft, vMiddle, vbBlack, vbBlack, , 3
             oDoc.WTextBox 605 + h, 300, 15, 255, "Firma:___________________________", "F1", 11
             oDoc.WTextBox 585 + h, 490, 90, 65, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
            End If
            nGaran = nGaran + 1
            RsGarantes.MoveNext
        Wend
    Else
End If
End Function
'***************************** END PTI1 ***************************************



    '----------------'COMENTADO PTI1 20170315-----------------------------------------
    Rem oDoc.WTextBox 470, 83, 20, 50, "D.O.I.", "F3", lnFuenteT, hLeft, , vbBlack
    Rem oDoc.WTextBox 470, 80, 20, 50, "", "F3", lnFuenteT, hLeft, , vbBlack, 1
    Rem oDoc.WTextBox 490, 83, 20, 50, "DOMICILIO", "F3", lnFuenteT, hLeft, , vbBlack
    Rem oDoc.WTextBox 490, 80, 20, 50, "", "F3", lnFuenteT, hLeft, , vbBlack, 1
            
    Rem oDoc.WTextBox 450, 130, 20, 180, "", "F3", lnFuenteT, hLeft, , vbBlack, 1
    Rem oDoc.WTextBox 470, 130, 20, 180, "", "F3", lnFuenteT, hLeft, , vbBlack, 1
    Rem oDoc.WTextBox 490, 130, 20, 180, "", "F3", lnFuenteT, hLeft, , vbBlack, 1
    
    Rem oDoc.WTextBox 450, 313, 20, 50, "NOMBRE", "F3", lnFuenteT, hLeft, , vbBlack
    Rem oDoc.WTextBox 450, 310, 20, 50, "", "F3", lnFuenteT, hLeft, , vbBlack, 1
    Rem oDoc.WTextBox 470, 313, 20, 50, "D.O.I.", "F3", lnFuenteT, hLeft, , vbBlack
    Rem oDoc.WTextBox 470, 310, 20, 50, "", "F3", lnFuenteT, hLeft, , vbBlack, 1
    Rem oDoc.WTextBox 490, 313, 20, 50, "DOMICILIO", "F3", lnFuenteT, hLeft, , vbBlack
    Rem oDoc.WTextBox 490, 310, 20, 50, "", "F3", lnFuenteT, hLeft, , vbBlack, 1
    Rem oDoc.WTextBox 450, 360, 20, 180, "", "F3", lnFuenteT, hLeft, , vbBlack, 1
    Rem oDoc.WTextBox 470, 360, 20, 180, "", "F3", lnFuenteT, hLeft, , vbBlack, 1
    Rem oDoc.WTextBox 490, 360, 20, 180, "", "F3", lnFuenteT, hLeft, , vbBlack, 1
    
    Rem oDoc.WTextBox 590, 83, 20, 50, "NOMBRE", "F3", lnFuenteT, hLeft, , vbBlack
    Rem oDoc.WTextBox 590, 80, 20, 50, "", "F3", lnFuenteT, hLeft, , vbBlack, 1
    Rem oDoc.WTextBox 610, 83, 20, 50, "D.O.I.", "F3", lnFuenteT, hLeft, , vbBlack
    Rem oDoc.WTextBox 610, 80, 20, 50, "", "F3", lnFuenteT, hLeft, , vbBlack, 1
    Rem oDoc.WTextBox 630, 83, 20, 50, "DOMICILIO", "F3", lnFuenteT, hLeft, , vbBlack
    Rem oDoc.WTextBox 630, 80, 20, 50, "", "F3", lnFuenteT, hLeft, , vbBlack, 1
    Rem oDoc.WTextBox 590, 130, 20, 180, "", "F3", lnFuenteT, hLeft, , vbBlack, 1
    Rem oDoc.WTextBox 610, 130, 20, 180, "", "F3", lnFuenteT, hLeft, , vbBlack, 1
    Rem oDoc.WTextBox 630, 130, 20, 180, "", "F3", lnFuenteT, hLeft, , vbBlack, 1
    
    Rem oDoc.WTextBox 590, 313, 20, 50, "NOMBRE", "F3", lnFuenteT, hLeft, , vbBlack
    Rem oDoc.WTextBox 590, 310, 20, 50, "", "F3", lnFuenteT, hLeft, , vbBlack, 1
    Rem oDoc.WTextBox 610, 313, 20, 50, "D.O.I.", "F3", lnFuenteT, hLeft, , vbBlack
    Rem oDoc.WTextBox 610, 310, 20, 50, "", "F3", lnFuenteT, hLeft, , vbBlack, 1
    Rem oDoc.WTextBox 630, 313, 20, 50, "DOMICILIO", "F3", lnFuenteT, hLeft, , vbBlack
    Rem oDoc.WTextBox 630, 310, 20, 50, "", "F3", lnFuenteT, hLeft, , vbBlack, 1
    Rem oDoc.WTextBox 590, 360, 20, 180, "", "F3", lnFuenteT, hLeft, , vbBlack, 1
    Rem oDoc.WTextBox 610, 360, 20, 180, "", "F3", lnFuenteT, hLeft, , vbBlack, 1
    Rem oDoc.WTextBox 630, 360, 20, 180, "", "F3", lnFuenteT, hLeft, , vbBlack, 1
    
    Rem Set oRs = oCF.ObtieneTitAvaPagare(ActXCodCta.NroCuenta)
    Rem If Not (oRs.EOF And oRs.BOF) Then
        Rem For nIndice = 1 To oRs.RecordCount
            Rem If oRs!nPrdPersRelac = 20 And nConTit = 0 Then
                Rem oDoc.WTextBox 450, 133, 20, 180, oRs!cPersNombre, "F3", lnFuenteT, hLeft, , vbBlack
                Rem oDoc.WTextBox 470, 133, 20, 180, oRs!cPersIDnro, "F3", lnFuenteT, hLeft, , vbBlack
                Rem oDoc.WTextBox 490, 133, 20, 180, oRs!cPersDireccDomicilio, "F3", lnFuenteT, hLeft, , vbBlack
            Rem End If
            Rem If oRs!nPrdPersRelac = 20 And nConTit = 1 Then
                Rem oDoc.WTextBox 450, 363, 20, 180, oRs!cPersNombre, "F3", lnFuenteT, hLeft, , vbBlack
                Rem oDoc.WTextBox 470, 363, 20, 180, oRs!cPersIDnro, "F3", lnFuenteT, hLeft, , vbBlack
                Rem oDoc.WTextBox 490, 363, 20, 180, oRs!cPersDireccDomicilio, "F3", lnFuenteT, hLeft, , vbBlack
            Rem End If
            
            Rem If oRs!nPrdPersRelac = 25 And nContAva = 0 Then
                Rem oDoc.WTextBox 590, 133, 20, 180, oRs!cPersNombre, "F3", lnFuenteT, hLeft, , vbBlack
                Rem oDoc.WTextBox 610, 133, 20, 180, oRs!cPersIDnro, "F3", lnFuenteT, hLeft, , vbBlack
                Rem oDoc.WTextBox 630, 133, 20, 180, oRs!cPersDireccDomicilio, "F3", lnFuenteT, hLeft, , vbBlack
            Rem End If
            Rem If oRs!nPrdPersRelac = 25 And nContAva = 1 Then
                Rem oDoc.WTextBox 590, 363, 20, 180, oRs!cPersNombre, "F3", lnFuenteT, hLeft, , vbBlack
                Rem oDoc.WTextBox 610, 363, 20, 180, oRs!cPersIDnro, "F3", lnFuenteT, hLeft, , vbBlack
                Rem oDoc.WTextBox 630, 363, 20, 180, oRs!cPersDireccDomicilio, "F3", lnFuenteT, hLeft, , vbBlack
            Rem End If
            Rem oRs.MoveNext
            Rem nConTit = 1
            Rem nContAva = 1
        Rem Next
        
        Rem oDoc.WTextBox 550, 120, 10, 180, "..............................", "F3", lnFuenteT, hCenter, , vbBlack
        Rem oDoc.WTextBox 560, 120, 10, 180, "Firma Emitente", "F3", lnFuenteT, hCenter, , vbBlack
        
        Rem oDoc.WTextBox 550, 320, 10, 180, "..............................", "F3", lnFuenteT, hCenter, , vbBlack
        Rem oDoc.WTextBox 560, 320, 10, 180, "Firma Emitente", "F3", lnFuenteT, hCenter, , vbBlack
        
        Rem oDoc.WTextBox 690, 120, 10, 180, "..............................", "F3", lnFuenteT, hCenter, , vbBlack
        Rem oDoc.WTextBox 700, 120, 10, 180, "Firma Aval Firma Aval", "F3", lnFuenteT, hCenter, , vbBlack
        
        Rem oDoc.WTextBox 690, 320, 10, 180, "..............................", "F3", lnFuenteT, hCenter, , vbBlack
        Rem oDoc.WTextBox 700, 320, 10, 180, "Firma Aval Firma Aval", "F3", lnFuenteT, hCenter, , vbBlack
        
    Rem End If
    Rem oDoc.WTextBox 720, 70, 10, 430, "Oficina Principal: Jr. Próspero No 791 - Iquitos ; Ag. Calle Arequipa: Ca Arequipa Nº 428; Agencia Punchana Av. 28 de Julio 829 -" & _
                                    REM "Iquitos; Ag. Belén: Av. Grau Nº 1260 - Iquitos; Ag. San Juan Bautista- Avda. Abelardo Quiñones Nº 2670- Iquitos; Ag. Pucallpa: Jr. " & _
                                    REM "Ucayali No 850 - 852 ; Ag. Huánuco: Jr. General Prado No 836; Ag. Yurimaguas: Ca. Simón Bolívar Nº 113; Ag. Tingo María: Av. " & _
                                    REM "Antonio Raymondi Nº 246 ; Ag. Tarapoto: Jr San Martín Nº 205 ; Ag. Requena: Calle San Francisco Mz 28 Lt 07; Ag. Cajamarca: Jr. " & _
                                    REM "Amalia Puga Nº 417; Ag. Aguaytía- Jr. Rio Negro Nº 259; Ag. Cerro de Pasco; Plaza Carrión Nº 191; Ag. Minka: Ciudad Comercial " & _
                                    REM "Minka - Av. Argentina Nº 3093- Local 230 - Callao.", "F3", lnFuenteT, hLeft, , vbBlack
    Rem oDoc.NewPage A4_Vertical
        
    Rem oDoc.WTextBox 50, 70, 100, 200, "", "F3", lnFuenteT, hLeft, , vbBlack, 1
    Rem oDoc.WTextBox 50, 80, 10, 250, "Prorrogado en las mismas condiciones", "F3", lnFuenteT, hLeft, , vbBlack
    Rem oDoc.WTextBox 65, 80, 10, 250, "el saldo de ..............................................................", "F3", lnFuenteT, hLeft, , vbBlack
    Rem oDoc.WTextBox 80, 80, 10, 250, "................................................................................", "F3", lnFuenteT, hLeft, , vbBlack
    Rem oDoc.WTextBox 95, 80, 10, 250, "Hasta el ..................................................................", "F3", lnFuenteT, hLeft, , vbBlack
    Rem oDoc.WTextBox 110, 80, 10, 250, "................................................................................", "F3", lnFuenteT, hLeft, , vbBlack
    
    Rem Set oCF = Nothing
    Rem Set oRs = Nothing
    Rem oDoc.PDFClose
    Rem oDoc.Show
    Rem Exit Sub
Rem ErrorImprimirPDF:
    Rem MsgBox Err.Description, vbInformation, "Aviso"
Rem End Sub
Rem 'RECO FIN********************************************************
'----------------'END COMENTADO PTI1 20170315------------------------------



