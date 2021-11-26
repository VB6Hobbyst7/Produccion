VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmCredRiesgosInformeMod 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informe de Riesgos"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7755
   Icon            =   "frmCredRiesgosInformeMod.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   7755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdVerInforme 
      Caption         =   "Ver Informe"
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   5280
      Width           =   1335
   End
   Begin VB.CommandButton cmdDarBaja 
      Caption         =   "Dejar sin Efecto"
      Height          =   375
      Left            =   4800
      TabIndex        =   7
      Top             =   5280
      Width           =   1575
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   6480
      TabIndex        =   3
      Top             =   5280
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   8705
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Inform de Riesgos"
      TabPicture(0)   =   "frmCredRiesgosInformeMod.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "frm"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblNivel"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblGlosa"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblCliente"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label3"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblMoneda"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label6"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblMonto"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "ActxCta"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmdExaminar"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      Begin VB.CommandButton cmdExaminar 
         Caption         =   "..."
         Height          =   375
         Left            =   3840
         TabIndex        =   6
         Top             =   480
         Width           =   375
      End
      Begin SICMACT.ActXCodCta ActxCta 
         Height          =   420
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   3660
         _ExtentX        =   6456
         _ExtentY        =   741
         Texto           =   "Credito"
         EnabledCMAC     =   -1  'True
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
      Begin VB.Label Label1 
         Caption         =   "Conclusión General:"
         Height          =   435
         Left            =   480
         TabIndex        =   15
         Top             =   2760
         Width           =   825
      End
      Begin VB.Label lblMonto 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1440
         TabIndex        =   13
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Monto:"
         Height          =   195
         Left            =   720
         TabIndex        =   12
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label lblMoneda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   5040
         TabIndex        =   11
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Moneda:"
         Height          =   195
         Left            =   4320
         TabIndex        =   10
         Top             =   1680
         Width           =   630
      End
      Begin VB.Label lblCliente 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1440
         TabIndex        =   9
         Top             =   1200
         Width           =   5775
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   720
         TabIndex        =   8
         Top             =   1200
         Width           =   525
      End
      Begin VB.Label lblGlosa 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   2055
         Left            =   1440
         TabIndex        =   5
         Top             =   2640
         Width           =   5655
      End
      Begin VB.Label lblNivel 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1440
         TabIndex        =   4
         Top             =   2160
         Width           =   3615
      End
      Begin VB.Label frm 
         AutoSize        =   -1  'True
         Caption         =   "Nivel de Riesgo:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   2280
         Width           =   1170
      End
   End
End
Attribute VB_Name = "frmCredRiesgosInformeMod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bExisteDatos As Boolean 'RECO20150520 ERS010-2015
Dim fnInformeID As Long

Public Sub Inicio(ByVal psCtaCod As String)
    Call LlenarCampos(psCtaCod)
    If bExisteDatos = True Then 'RECO20150520 ERS010-2015
        Me.Show 1
    End If
End Sub

Private Sub ActxCta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    LlenarCampos (Trim(ActxCta.NroCuenta))
End If
End Sub

Private Sub cmdCerrar_Click()
Unload Me
End Sub

Public Sub LlenarCampos(ByVal pnNroCuenta As String)
Dim oCredito As COMDCredito.DCOMCredito
Dim rsCredito As ADODB.Recordset
Set oCredito = New COMDCredito.DCOMCredito
Set rsCredito = oCredito.ObtenerInformeRiesgo(Trim(pnNroCuenta), 1)

Me.ActxCta.NroCuenta = pnNroCuenta
ActxCta.Enabled = False

If Not (rsCredito.EOF And rsCredito.BOF) Then
    lblCliente.Caption = rsCredito!Cliente
    lblMonto.Caption = Format(rsCredito!Monto, "#000.00")
    lblMoneda.Caption = rsCredito!Moneda
    lblNivel.Caption = rsCredito!Nivel
    lblGlosa.Caption = rsCredito!Glosa
    fnInformeID = rsCredito!nInformeID
    bExisteDatos = True 'RECO20150520 ERS010-2015
    
    cmdVerInforme.Enabled = True 'JOEP-ERS064-20170608
    cmdDarBaja.Enabled = True 'JOEP ERS-064-20170608
Else
    MsgBox "No Existen Datos", vbInformation, "Aviso"
    bExisteDatos = False 'RECO20150520 ERS010-2015
    Call LimpiarDatos
End If



End Sub

Private Sub cmdDarBaja_Click()
    Dim oCredito As COMDCredito.DCOMCredito
    Set oCredito = New COMDCredito.DCOMCredito
    
        If Trim(ActxCta.Cuenta) = "" Or Len(Trim(ActxCta.NroCuenta)) < 18 Then
            MsgBox "Ingrese un numero de credito correctamente.", vbInformation, "Aviso"
            Call LimpiarDatos
        Else
            If MsgBox("Estas Seguro de Dejar sin Efecto el Informe?", vbYesNo + vbInformation, "AVISO") = vbYes Then
                'Call oCredito.OpeInformeRiesgo(Trim(ActxCta.NroCuenta), 2, , , , , 3)
                oCredito.ActualizarInformeRiesgo Trim(ActxCta.NroCuenta), fnInformeID, 3, , , , GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), , , , "Informe de Riesgo dado de baja"
                Call SalidaObs 'RECO20161019 ERS060-2016
                MsgBox "Se Dejo Sin Efecto el Informe Correctamente.", vbInformation, "Aviso"
                Call LimpiarDatos
            End If
        End If
    Set oCredito = Nothing
End Sub

Private Sub LimpiarDatos()
    ActxCta.Enabled = True
    ActxCta.NroCuenta = ""
    lblCliente.Caption = ""
    lblMonto.Caption = ""
    lblMoneda.Caption = ""
    lblNivel.Caption = ""
    lblGlosa.Caption = ""
    ActxCta.CMAC = "109"
    ActxCta.EnabledCMAC = False
    ActxCta.Age = gsCodAge
End Sub

Private Sub cmdExaminar_Click()
    Dim loPers As COMDPersona.UCOMPersona
    Dim lsPersCod As String, lsPersNombre As String
    Dim lsEstados As String
    Dim loPersCreditos As COMDCredito.DCOMCredito
    Dim lrCreditos As New ADODB.Recordset
    Dim loCuentas As COMDPersona.UCOMProdPersona
    
    On Error GoTo ControlError
    
    Set loPers = New COMDPersona.UCOMPersona
        Set loPers = frmBuscaPersona.Inicio
        If loPers Is Nothing Then Exit Sub
        lsPersCod = loPers.sPersCod
        lsPersNombre = loPers.sPersNombre
    Set loPers = Nothing
    
    If Trim(lsPersCod) <> "" Then
        Set loPersCreditos = New COMDCredito.DCOMCredito
        Set lrCreditos = loPersCreditos.CreditosInformeRiesgo(lsPersCod)
        Set loPersCreditos = Nothing
    End If
    
    Set loCuentas = New COMDPersona.UCOMProdPersona
        Set loCuentas = frmProdPersona.Inicio(lsPersNombre, lrCreditos)
        If loCuentas.sCtaCod <> "" Then
            ActxCta.NroCuenta = Mid(loCuentas.sCtaCod, 1, 18)
            Call ActxCta_KeyPress(13)
        End If
    Set loCuentas = Nothing
    Exit Sub
ControlError:
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Sub

Private Sub cmdVerInforme_Click()
    Dim rsVerInfRisgNewMod As ADODB.Recordset 'ERS-064 JOEP20170523
    Dim oVerfInfRisg As COMDCredito.DCOMCredito 'ERS-064 JOEP20170523
    
    Set oVerfInfRisg = New COMDCredito.DCOMCredito 'ERS-064 JOEP20170523
    
    If Trim(ActxCta.Cuenta) = "" Or Len(Trim(ActxCta.NroCuenta)) < 18 Then
        MsgBox "Ingrese un numero de credito correctamente.", vbInformation, "Aviso"
    Else
    Set rsVerInfRisgNewMod = oVerfInfRisg.VerificaInfRisgNewMod(Trim(ActxCta.NroCuenta)) 'ERS-064 JOEP20170523
    
        If Not (rsVerInfRisgNewMod.EOF And rsVerInfRisgNewMod.BOF) Then 'ERS-064 JOEP20170523
            If rsVerInfRisgNewMod!nValor = 1 Then 'ERS-064 JOEP20170523
                MsgBox "El Crédito tiene Informe de Riesgo con la Nueva Modalidad", vbInformation, "Aviso" 'ERS-064 JOEP20170523
            End If 'ERS-064 JOEP20170523
        Else
            frmCredRiesgosInforme.Inicio (Trim(ActxCta.NroCuenta))
        End If
        
    End If
    'Call ReporteAdeudadosCalendarioVigente
End Sub
'RECO20161019 ERS060-2016 **********************************************
Private Sub SalidaObs()
    Dim oNCOMColocEval As New NCOMColocEval
    Dim oNCOMContFunciones As New COMNContabilidad.NCOMContFunciones
    Dim lcMovNro As String
    lcMovNro = oNCOMContFunciones.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
    Call oNCOMColocEval.updateEstadoExpediente(ActxCta.NroCuenta, gTpoRegCtrlRiesgos)   'BY ARLO 20171017
    Call oNCOMColocEval.insEstadosExpediente(ActxCta.NroCuenta, "Créditos - Analista", "", lcMovNro, "", "", 1, 2001, gTpoRegCtrlRiesgos)
    MsgBox "Expediente Salio por Observación de Gerencia de Riesgos", vbInformation, "Aviso"
    Set oNCOMColocEval = Nothing
End Sub
'RECO FIN **************************************************************
Private Sub Form_Load()
ActxCta.CMAC = "109"
ActxCta.EnabledCMAC = False
ActxCta.Age = gsCodAge
cmdVerInforme.Enabled = False 'JOEP ERS-064-20170608
cmdDarBaja.Enabled = False 'JOEP ERS-064-20170608
End Sub
'Public Sub ReporteAdeudadosCalendarioVigente()
'    Dim fs As Scripting.FileSystemObject
'    Dim lbExisteHoja As Boolean
'    Dim lsArchivo1 As String
'    Dim lsNomHoja  As String
'    Dim lsNombreAgencia As String
'    Dim lsCodAgencia As String
'    Dim lsMes As String
'    Dim lnContador As Integer
'    Dim lsArchivo As String
'    Dim xlsAplicacion As Excel.Application
'    Dim xlsLibro As Excel.Workbook
'    Dim xlHoja1 As Excel.Worksheet
'
'
'
'    Dim oCreditos As New DCreditos
'
'    Dim sTexto As String
'    Dim sDocFecha As String
'    Dim nSaltoContador As Double
'    Dim sFecha As String
'    Dim sMov As String
'    Dim sDoc As String
'    Dim n As Integer
'    Dim pnLinPage As Integer
'    Dim nMES As Integer
'    Dim nSaldo12 As Currency
'    Dim nContTotal As Double
'    Dim nPase As Integer
'    Dim dFechaCP As Date
'    Dim lsCelda As String
'    Dim objDPersona As COMDPersona.DCOMPersona
'    Dim ors As ADODB.Recordset
'    'Dim lnContador As Integer
'    Set ors = New ADODB.Recordset
'    Set objDPersona = New COMDPersona.DCOMPersona
''On Error GoTo GeneraExcelErr
'    Dim clsTC As COMDConstSistema.NCOMTipoCambio
'    Dim nTC As Double
'    Set clsTC = New COMDConstSistema.NCOMTipoCambio
'    nTC = clsTC.EmiteTipoCambio(gdFecSis, TCFijoDia)
'    nTC = clsTC.EmiteTipoCambio(gdFecSis, TCFijoDia)
'    Set clsTC = Nothing
'
'    Set fs = New Scripting.FileSystemObject
'    Set xlsAplicacion = New Excel.Application
'    lsArchivo = "InformeRiesgo"
'    'Primera Hoja ******************************************************
'    lsNomHoja = "OPINIÓN DE RIESGOS"
'    '*******************************************************************
'    lsArchivo1 = "\spooler\" & lsArchivo & gsCodUser & "_" & Format(gdFecha, "YYYYMMDD") & "_" & Format$(Time(), "HHMMSS") & ".xls"
'    If fs.FileExists(App.path & "\FormatoCarta\" & lsArchivo & ".xls") Then
'        Set xlsLibro = xlsAplicacion.Workbooks.Open(App.path & "\FormatoCarta\" & lsArchivo & ".xls")
'    Else
'        MsgBox "No Existe Plantilla en Carpeta FormatoCarta, Consulte con el Area de  TI", vbInformation, "Advertencia"
'        Exit Sub
'    End If
'
'    For Each xlHoja1 In xlsLibro.Worksheets
'       If xlHoja1.Name = lsNomHoja Then
'            xlHoja1.Activate
'         lbExisteHoja = True
'        Exit For
'       End If
'    Next
'    If lbExisteHoja = False Then
'        Set xlHoja1 = xlsLibro.Worksheets
'        xlHoja1.Name = lsNomHoja
'    End If
'
'    nSaltoContador = 6
'
'    'Cuadro1
'    Set ors = objDPersona.ObtenerInformeRiesgoCuadro1(ActxCta.NroCuenta, gdFecSis)
'    nPase = 1
'    If (ors Is Nothing) Then
'        nPase = 0
'    End If
'    If nPase = 1 Then
'    Do While Not ors.EOF
'            xlHoja1.Cells(6, 8) = ors!cPersNombreTitular
'            xlHoja1.Cells(6, 36) = ors!cAgeDescripcion
'            xlHoja1.Cells(7, 5) = ors!cActividad
'            xlHoja1.Cells(7, 36) = ors!cPersNombreAnalista
'            xlHoja1.Cells(8, 8) = ors!Antiguedad_Neg
'            xlHoja1.Cells(8, 36) = ors!cModalidad
'            xlHoja1.Cells(9, 5) = ors!cSector
'            xlHoja1.Cells(9, 37) = "'" & ors!cCtaCod
'            xlHoja1.Cells(10, 5) = ors!cDestino
'            xlHoja1.Cells(10, 37) = ors!cTipoCredito
'
'            xlHoja1.Cells(15, 8) = IIf(IsNull(ors!Calificacion), "", ors!Calificacion)
'            xlHoja1.Cells(15, 23) = IIf(IsNull(ors!Antiguedad_CMACM), "", ors!Antiguedad_CMACM)
'            xlHoja1.Cells(15, 36) = IIf(IsNull(ors!CantidadCreditos), "", ors!CantidadCreditos)
'
'            xlHoja1.Cells(17, 7) = IIf(IsNull(ors!CalificacionSBS), "", ors!CalificacionSBS)
'            xlHoja1.Cells(17, 23) = IIf(Mid(ors!cCtaCod, 9, 1) = "1", "S/.", "$") & Format(IIf(IsNull(ors!nSumSalRCC), 0, ors!nSumSalRCC), "###,###,###,###0.00")
'            xlHoja1.Cells(17, 34) = IIf(IsNull(ors!Can_EntsRCC), 0, ors!Can_EntsRCC)
'
'            xlHoja1.Cells(19, 13) = IIf(IsNull(ors!cHisCredCMACM), "", ors!cHisCredCMACM)
'            xlHoja1.Cells(24, 13) = IIf(IsNull(ors!cEvolSistFina), "", ors!cEvolSistFina)
'
'            nSaltoContador = nSaltoContador + 1
'            ors.MoveNext
'        If ors.EOF Then
'           Exit Do
'        End If
'    Loop
'    End If
'
'    'Vinculados
'    Set ors = objDPersona.ObtenerInformeRiesgoCuadro1Vinculados(ActxCta.NroCuenta)
'    lnContador = 30
'    If Not (ors.BOF Or ors.EOF) Then
'    Do While Not ors.EOF
'        xlHoja1.Cells(lnContador, 2) = ors!cPersNombre
'        xlHoja1.Cells(lnContador, 3) = ors!cConsDescripcion
'        xlHoja1.Cells(lnContador, 4) = ors!Val_Saldo
'        xlHoja1.Cells(lnContador, 5) = ors!nCantidadESF
'        xlHoja1.Cells(lnContador, 6) = ors!cCalSBS
'        xlHoja1.Cells(lnContador, 7) = ors!cEvoEndeu
'        lnContador = lnContador + 1
'        ors.MoveNext
'    Loop
'    End If
'
'    Set ors = objDPersona.ObtenerInformeRiesgoCuadro2SaldoDeudor(ActxCta.NroCuenta)
'    lnContador = 38
'    If Not (ors.BOF Or ors.EOF) Then
'    Do While Not ors.EOF
'            xlHoja1.Cells(lnContador, 6) = ors!nSaldo
'            lnContador = lnContador + 1
'            ors.MoveNext
'    Loop
'    End If
'
'    Set ors = objDPersona.ObtenerInformeRiesgoCuadro2Creditos(ActxCta.NroCuenta, nTC)
'    If Not (ors.BOF Or ors.EOF) Then
'    Do While Not ors.EOF
'         xlHoja1.Cells(38, 15) = ors!nMontoPropuesto
'         xlHoja1.Cells(38, 24) = ors!nNroCuotas
'         xlHoja1.Cells(40, 23) = IIf(Mid(ActxCta.NroCuenta, 9, 1) = 1, "SOLES", "DOLARES")
'         xlHoja1.Cells(40, 31) = ors!nTEA
'         xlHoja1.Cells(42, 31) = IIf(IsNull(ors!nTEMA), 0#, ors!nTEMA)
'         xlHoja1.Cells(44, 52) = IIf(IsNull(ors!nVGET), 0, ors!nVGET)
'         xlHoja1.Cells(51, 34) = IIf(IsNull(ors!nCuoPropuesta), 0, ors!nCuoPropuesta)
'        ors.MoveNext
'    Loop
'    End If
'
'    Set ors = objDPersona.ObtenerInformeRiesgoCuadro2CartasFianzas(ActxCta.NroCuenta, nTC)
'    If Not (ors.BOF Or ors.EOF) Then
'    Do While Not ors.EOF
'        xlHoja1.Cells(38, 15) = ors!nMontoPropuesto
'         xlHoja1.Cells(42, 23) = ors!nTipoCambio
'         xlHoja1.Cells(38, 31) = ors!nComisionTr
'         xlHoja1.Cells(40, 23) = IIf(Mid(ActxCta.NroCuenta, 9, 1) = 1, "SOLES", "DOLARES")
'         xlHoja1.Cells(44, 52) = ors!nVGET
'        ors.MoveNext
'    Loop
'    End If
'
'    Set ors = objDPersona.ObtenerInformeRiesgoCuadro2Garantias(ActxCta.NroCuenta)
''    FormateaFlex FEGarantias2
''    lnContadorGada = 0
'    If Not (ors.BOF Or ors.EOF) Then
'    Do While Not ors.EOF
'            xlHoja1.Cells(44, 15) = ors!nVRM
'            xlHoja1.Cells(44, 22) = ors!nVGravamen
'            xlHoja1.Cells(44, 6) = ors!cConsDescripcion
'            xlHoja1.Cells(46, 9) = ors!cDescripcion
'            ors.MoveNext
'    Loop
'    End If
'
'
'    Set ors = objDPersona.ObtenerInformeRiesgoCuadro2EscaCuotasC2(ActxCta.NroCuenta)
''    FormateaFlex FECuoPro2
'    lnContador = 51
'
'    If Not (ors.BOF Or ors.EOF) Then
'    Do While Not ors.EOF
'            xlHoja1.Cells(lnContador, 21) = ors!nMontoPagado
'            ors.MoveNext
'    Loop
'    End If
'
'    Set ors = objDPersona.ObtenerInformeRiesgoCuadro3(ActxCta.NroCuenta, nTC)
'    If Not (ors.BOF Or ors.EOF) Then
'    Do While Not ors.EOF
'        xlHoja1.Cells(61, 5) = ors!nVentas
'        xlHoja1.Cells(62, 16) = ors!nUtilidades
'        xlHoja1.Cells(63, 5) = ors!nUnVentas
'        xlHoja1.Cells(64, 5) = ors!nRazonCTE '
'        xlHoja1.Cells(61, 16) = ors!nSaldoDisponible
'        xlHoja1.Cells(62, 16) = ors!nCapaPago
'        xlHoja1.Cells(64, 16) = ors!nCapiTrab
'        xlHoja1.Cells(61, 27) = ors!nPatrimonio
'        xlHoja1.Cells(61, 38) = ors!nPasivo
'        xlHoja1.Cells(62, 38) = ors!nLineaCred
'        xlHoja1.Cells(62, 27) = ors!nApalancamiento
'        xlHoja1.Cells(63, 16) = ors!nSensibli
'        xlHoja1.Cells(63, 27) = ors!nMoraDelSector
'        xlHoja1.Cells(63, 38) = ors!nOtrosIngr
'        xlHoja1.Cells(64, 27) = ors!nCapitalSocial
'        xlHoja1.Cells(68, 5) = ors!nMesAnterior1
'        xlHoja1.Cells(68, 9) = ors!nMesAnterior2
'        xlHoja1.Cells(68, 13) = ors!nMesAnterior3
'        xlHoja1.Cells(69, 22) = ors!nVentasAnter1
'        xlHoja1.Cells(69, 26) = ors!nUtilidAnter1
'        xlHoja1.Cells(69, 30) = ors!nVentasAnter2
'        xlHoja1.Cells(69, 34) = ors!nUtilidAnter2
'        xlHoja1.Cells(71, 9) = ors!cCondicionS
'        xlHoja1.Cells(73, 8) = ors!cCalidadEvS
'        ors.MoveNext
'    Loop
'    End If
'
'    'Cuadro4
'    Set ors = objDPersona.ObtenerInformeRiesgoCuadro4(ActxCta.NroCuenta)
'    nPase = 1
'    If (ors Is Nothing) Then
'        nPase = 0
'    End If
'    If nPase = 1 Then
'        Do While Not ors.EOF
'                xlHoja1.Cells(2, 2) = "INFORME DE OPINIÓN DE RIESGOS  Nº " & ors!cNumeroInforme
'                If ors!nNivelRiesgo = 1 Then
'                    xlHoja1.Cells(85, 6) = "SI"
'                ElseIf ors!nNivelRiesgo = 2 Then
'                    xlHoja1.Cells(85, 17) = "SI"
'                ElseIf ors!nNivelRiesgo = 3 Then
'                    xlHoja1.Cells(85, 28) = "SI"
'                ElseIf ors!nNivelRiesgo = 4 Then
'                    xlHoja1.Cells(85, 39) = "SI"
'                End If
'                xlHoja1.Cells(87, 6) = ors!cConclusione
'                xlHoja1.Cells(98, 7) = ors!cRecomendaci
'                xlHoja1.Cells(113, 7) = ors!cAnalisisRie
'                xlHoja1.Cells(113, 37) = Format(ors!dFechaNR, "YYYY/MM/DD")
'                nSaltoContador = nSaltoContador + 1
'            ors.MoveNext
'            If ors.EOF Then
'               Exit Do
'            End If
'        Loop
'    End If
'
'    Set objDPersona = Nothing
'    Set ors = Nothing
'
'    xlHoja1.SaveAs App.path & lsArchivo1
'    xlsAplicacion.Visible = True
'    xlsAplicacion.Windows(1).Visible = True
'    Set xlsAplicacion = Nothing
'    Set xlsLibro = Nothing
'    Set xlHoja1 = Nothing
'
'Exit Sub
'End Sub
