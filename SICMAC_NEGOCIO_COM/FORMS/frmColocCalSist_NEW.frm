VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmColocCalSist_NEW 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Colocaciones - Procesar Calificación 11356"
   ClientHeight    =   6510
   ClientLeft      =   4215
   ClientTop       =   2100
   ClientWidth     =   7095
   Icon            =   "frmColocCalSist_NEW.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCalifIndicadorAtrasoSBS 
      BackColor       =   &H000000C0&
      Caption         =   "Indicador Atraso - SBS"
      Height          =   375
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   3840
      Width           =   2295
   End
   Begin VB.CommandButton cmdCalifMetodologiaInterna 
      BackColor       =   &H00FFFF80&
      Caption         =   "Metodologia Interna"
      Height          =   375
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   3360
      Width           =   2295
   End
   Begin VB.CommandButton cmdCalifRiesgoCredCambiario 
      BackColor       =   &H0080FFFF&
      Caption         =   "Riesgo Crediticio Cambiario"
      Height          =   375
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   2880
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Height          =   6465
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   3795
      Begin VB.CommandButton cmdSeguimiento 
         BackColor       =   &H8000000B&
         Caption         =   "12.Obtener datos seguimiento"
         Height          =   420
         Left            =   200
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   5880
         Width           =   3420
      End
      Begin VB.CommandButton cmdProvisionProciclica 
         BackColor       =   &H8000000B&
         Caption         =   "11. Calcula Provision Prociclica"
         Height          =   420
         Left            =   200
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   5400
         Width           =   3420
      End
      Begin VB.CommandButton cmdCalculaProvSistF 
         BackColor       =   &H80000005&
         Caption         =   "10. Calcula Pr&ovision Según Calif. Sist. Finan."
         Height          =   420
         Left            =   200
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   4920
         Width           =   3420
      End
      Begin VB.CommandButton cmdCalculaProvSinAlinea 
         BackColor       =   &H80000005&
         Caption         =   "9. Calcula P&rovision Sin Alineamiento"
         Height          =   420
         Left            =   200
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   4440
         Width           =   3420
      End
      Begin VB.CommandButton cmdCalSinAlineamiento 
         BackColor       =   &H80000005&
         Caption         =   "8. Calificación Sin &Alineamiento"
         Height          =   420
         Left            =   200
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   3960
         Width           =   3420
      End
      Begin VB.CommandButton cmdActualizaGarantias 
         BackColor       =   &H80000005&
         Caption         =   "2. Califica Garantias"
         Height          =   420
         Left            =   200
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   720
         Width           =   3420
      End
      Begin VB.CommandButton cmdLLenaCreditoAudi 
         BackColor       =   &H80000005&
         Caption         =   "1. Preparar Archivo para Calificacion"
         Height          =   420
         Left            =   200
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Width           =   3420
      End
      Begin VB.CommandButton cmdCalculaProvision 
         BackColor       =   &H80000005&
         Caption         =   "7. Calcula &Provision"
         Height          =   420
         Left            =   200
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   3480
         Width           =   3420
      End
      Begin VB.CommandButton cmdCalSistF 
         BackColor       =   &H80000005&
         Caption         =   "5. Calificación  &Sistema Financiero"
         Height          =   420
         Left            =   200
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2520
         Width           =   3420
      End
      Begin VB.CommandButton cmdCalEvaluacion 
         BackColor       =   &H80000005&
         Caption         =   "4. Calificación  &Evaluacion Riesgos"
         Height          =   420
         Left            =   200
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2040
         Width           =   3420
      End
      Begin VB.CommandButton cmdCalCMAC 
         BackColor       =   &H80000005&
         Caption         =   "3. Calificación CMACM"
         Height          =   420
         Left            =   200
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1200
         Width           =   3420
      End
      Begin VB.CommandButton cmdCalGen 
         BackColor       =   &H80000005&
         Caption         =   "6. Calificación &General"
         Height          =   420
         Left            =   200
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   3000
         Width           =   3420
      End
      Begin VB.CommandButton cmdEndeudamientoSF 
         Caption         =   "Endeudamiento Sistema Financiero"
         Height          =   420
         Left            =   200
         TabIndex        =   15
         Top             =   2520
         Visible         =   0   'False
         Width           =   3420
      End
      Begin VB.CommandButton cmdCalRFA 
         Caption         =   "Calificación Creditos RFA"
         Height          =   420
         Left            =   200
         TabIndex        =   16
         Top             =   225
         Visible         =   0   'False
         Width           =   3420
      End
      Begin VB.CommandButton cmdCalRiesgoUnico 
         Caption         =   "Calificación Riesgo Unico"
         Height          =   420
         Left            =   200
         TabIndex        =   14
         Top             =   720
         Visible         =   0   'False
         Width           =   3420
      End
   End
   Begin VB.Frame FraFecha 
      Height          =   1410
      Left            =   3960
      TabIndex        =   1
      Top             =   0
      Width           =   2895
      Begin VB.TextBox txtTipoCambio 
         Alignment       =   2  'Center
         Height          =   330
         Left            =   1440
         TabIndex        =   11
         Top             =   900
         Width           =   1305
      End
      Begin MSMask.MaskEdBox txtFecha 
         Height          =   330
         Left            =   1440
         TabIndex        =   2
         Top             =   360
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   582
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tip.Cambio:"
         Height          =   195
         Left            =   480
         TabIndex        =   12
         Top             =   945
         Width           =   840
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha:"
         Height          =   240
         Left            =   840
         TabIndex        =   3
         Top             =   405
         Width           =   600
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   420
      Left            =   5400
      TabIndex        =   0
      Top             =   6000
      Width           =   1545
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000D&
      X1              =   0
      X2              =   3600
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label Label4 
      Caption         =   "Resolucion 1494 - Dic2006"
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
      Left            =   4680
      TabIndex        =   20
      Top             =   2520
      Width           =   2655
   End
   Begin VB.Label lblFecAlin 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   5775
      TabIndex        =   18
      Top             =   1635
      Width           =   1035
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Fecha Alineacion:"
      Height          =   195
      Left            =   4455
      TabIndex        =   17
      Top             =   1665
      Width           =   1275
   End
End
Attribute VB_Name = "frmColocCalSist_NEW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************
'* COLOCACIONES - CALIFICACION SISTEMA
'Archivo:  frmColocCalSist_NEW.frm
'LAYG   :  01/10/2002.-
'CAJA ICA : 26/06/2004 - 23/09/2004
'Cusco  : 2006/12 - LAYG
'Resumen:  Realiza el Proceso de Calificacion de la Cartera
Option Explicit

Dim fnTipoCambio  As Currency
Dim fdFechaFinMes  As Date
Dim fsServerConsol As String
Dim fsServerRCC As String
Dim fsBDRCC As String
Dim lsTablaTMP As String

'LUCV20190601, Comentó. Mejoras en el proceso de la cartera
'Private Sub cmdActualizaGarantias_Click()
'Dim oEval As COMNCredito.NCOMColocEval
'Set oEval = New COMNCredito.NCOMColocEval
'    Call oEval.nCalificaActualizarGarantias(gdFecSis, fnTipoCambio)
'Set oEval = Nothing
'
'MsgBox "Actualizacion Garantias Generada Correctamente , Termino : " & Time(), vbInformation, "Aviso"
'Me.barraestado.Panels(1).Text = ""
'Me.BarraProgreso.value = 0
'
'End Sub

'Private Sub cmdCalCMAC_Click()
'
'Dim oEval As COMNCredito.NCOMColocEval
'Set oEval = New COMNCredito.NCOMColocEval
'    Call oEval.nCalificaCMACT(fnTipoCambio)
'Set oEval = Nothing
'MsgBox "Calificacion Empresa Completada", vbInformation, "Aviso"
'Me.barraestado.Panels(1).Text = ""
'Me.BarraProgreso.value = 0
'
'End Sub

'Private Sub cmdCalculaProvision_Click()
'
'Dim oEval As COMNCredito.NCOMColocEval
'Set oEval = New COMNCredito.NCOMColocEval
'
'Screen.MousePointer = vbHourglass
'
'    Call oEval.nCalificaCalculaProvision(txtFecha.Text)
'
'Screen.MousePointer = vbDefault
'    MsgBox "Calculo de Provisión sin RCC terminado Correctamente ", vbInformation, "Aviso"
'
'Screen.MousePointer = vbHourglass
'
'    Call oEval.CalculaProvisionRCC(fsServerConsol, fnTipoCambio, txtFecha.Text)
'
'Screen.MousePointer = vbDefault
'
'    MsgBox "Calculo de Provisión con RCC terminado Correctamente ", vbInformation, "Aviso"
'Set oEval = Nothing
'
'MsgBox "Proceso de Calificacion terminado Correctamente ", vbInformation, "Aviso"
'
'End Sub

'Private Sub cmdCalculaProvSinAlinea_Click()
'Dim oEval As COMNCredito.NCOMColocEval
'    Set oEval = New COMNCredito.NCOMColocEval
'    Screen.MousePointer = vbHourglass
'    Call oEval.calculaProvisionSinAlineamiento(txtFecha.Text)
'    Screen.MousePointer = vbDefault
'    Set oEval = Nothing
'    MsgBox "El Calculo de Provisión Sin Alineamiento a Terminado Correctamente ", vbInformation, "Aviso"
'End Sub

'Private Sub cmdCalculaProvSistF_Click()
'Dim oEval As COMNCredito.NCOMColocEval
'    Set oEval = New COMNCredito.NCOMColocEval
'    Screen.MousePointer = vbHourglass
'    Call oEval.calculaProvisionSistemaFinanciero(txtFecha.Text)
'    Screen.MousePointer = vbDefault
'    Set oEval = Nothing
'    MsgBox "El Calculo de Provisión Según Calif. del Sistema Financiero a Terminado Correctamente ", vbInformation, "Aviso"
'
'End Sub

'Private Sub cmdCalEvaluacion_Click()
'
'Dim oEval As COMNCredito.NCOMColocEval
'Set oEval = New COMNCredito.NCOMColocEval
'    Call oEval.nCalificaEvaluacionRiesgos(txtFecha.Text)
'Set oEval = Nothing
'
'MsgBox "Calificacion Evaluacion de Cartera completado", vbInformation, "Aviso"
'Me.barraestado.Panels(1).Text = ""
'Me.BarraProgreso.value = 0
'End Sub

'Private Sub cmdCalGen_Click()
'Dim oEval As COMNCredito.NCOMColocEval
'Dim sMensaje As String
'
'Screen.MousePointer = vbHourglass
'
'Set oEval = New COMNCredito.NCOMColocEval
'    'ALPA 20100730**********************************************
'    'Call oEval.nCalificaGeneral(sMensaje)
'    Call oEval.nCalificaGeneral(sMensaje, txtFecha.Text)
'    '***********************************************************
'Set oEval = Nothing
'
'Screen.MousePointer = vbDefault
'
'If sMensaje <> "" Then
'    MsgBox sMensaje, vbInformation, "Mensaje"
'    Exit Sub
'End If
'
'MsgBox "Calificacion General completada ", vbInformation, "Aviso"
'End Sub

'Private Sub cmdCalifMetodologiaInterna_Click()
'Dim oEval As COMNCredito.NCOMColocEval
'Set oEval = New COMNCredito.NCOMColocEval
'    Call oEval.nCalificaMetodologiaInterna(fsServerConsol, txtFecha.Text)
'Set oEval = Nothing
'MsgBox "Calificacion Metodologia Interna Completada", vbInformation, "Aviso"
'Me.barraestado.Panels(1).Text = ""
'Me.BarraProgreso.value = 0
'
'End Sub

'Private Sub cmdCalifRiesgoCredCambiario_Click()
'
'Dim oEval As COMNCredito.NCOMColocEval
'Set oEval = New COMNCredito.NCOMColocEval
'    Call oEval.nCalificaRiesgoCredCambiario
'Set oEval = Nothing
'MsgBox "Calificacion Riesgo Cambiario Completada", vbInformation, "Aviso"
'Me.barraestado.Panels(1).Text = ""
'Me.BarraProgreso.value = 0
'
'End Sub

'Private Sub cmdCalSinAlineamiento_Click()
'Dim oEval As COMNCredito.NCOMColocEval
'Dim sMensaje As String
'
'    Screen.MousePointer = vbHourglass
'
'    Set oEval = New COMNCredito.NCOMColocEval
'        Call oEval.generarCalificacionSinAlineamiento(sMensaje)
'    Set oEval = Nothing
'
'    Screen.MousePointer = vbDefault
'
'    If sMensaje <> "" Then
'        MsgBox sMensaje, vbInformation, "Mensaje"
'        Exit Sub
'    End If
'
'    MsgBox "Calificacion Sin Alineamiento Completada ", vbInformation, "Aviso"
'End Sub

'Private Sub cmdCalSistF_Click()
'Dim oEval As COMNCredito.NCOMColocEval
'Dim sMensaje As String
'
'Screen.MousePointer = vbHourglass
'
'Set oEval = New COMNCredito.NCOMColocEval
'    Call oEval.nCalificaSistemaFinanciero(fsBDRCC, CDate(lblFecAlin.Caption), sMensaje, fsServerConsol)
'Set oEval = Nothing
'
'Screen.MousePointer = vbDefault
'
'If sMensaje <> "" Then
'    MsgBox sMensaje, vbInformation, "Mensaje"
'    Exit Sub
'End If

'Private Sub cmdLLenaCreditoAudi_Click()
'
''** Llena tabla ColocCalifProv (Contiene las Calificaciones de los creditos Vigentes)
'Dim sMensaje As String
'Dim oEval As COMNCredito.NCOMColocEval
'On Error GoTo ErrorConexion
'
'If VerificaDatosIngresados = False Then Exit Sub
'
'If MsgBox("Se Reprocesaran los Datos para la Calificacion, Desea Continuar ?", vbQuestion + vbYesNo, "Aviso") = vbNo Then
'    Exit Sub
'End If
'
''ARCV: se agrego el Tipo de Producto
'frmColocEvalPorProducto.Show 1
'
'Screen.MousePointer = vbHourglass
'
'Set oEval = New COMNCredito.NCOMColocEval
'    Call oEval.nCalificaPreparaArchivoCalificacion(fsServerConsol, txtFecha.Text, fnTipoCambio, sMensaje, frmColocEvalPorProducto.MatIndices)
'Set oEval = Nothing
'
'Screen.MousePointer = vbDefault
'
'If sMensaje <> "" Then
'    MsgBox sMensaje, vbInformation, "Mensaje"
'    Exit Sub
'End If
'MsgBox "Preparacion de Archivo para Calificacion completado ", vbInformation, "Aviso"
'Exit Sub
'ErrorConexion:
'    MsgBox "Error Nº[" & str(Err.Number) & "] " & Err.Description, vbInformation, "Aviso"
'End Sub

'Private Sub cmdProvisionProciclica_Click()
'Dim oEval As COMNCredito.NCOMColocEval
'Set oEval = New COMNCredito.NCOMColocEval
'
'Screen.MousePointer = vbHourglass
'
'    Call oEval.nCalificaCalculaProvisionProciclica(txtFecha.Text)
'
'Screen.MousePointer = vbDefault
'    MsgBox "Calculo de Provisión Prociclica terminado Correctamente ", vbInformation, "Aviso"
'
'Set oEval = Nothing
'
'MsgBox "Proceso de Calificacion terminado Correctamente ", vbInformation, "Aviso"
'
'End Sub

'Fin LUCV20190601

Private Sub cmdCalifIndicadorAtrasoSBS_Click()
    Dim oEval As COMNCredito.NCOMColocEval
    Set oEval = New COMNCredito.NCOMColocEval
    
    Call oEval.nCalificaDiasAtrasoSBS(fsServerConsol, txtFecha.Text)
    Set oEval = Nothing
    
    MsgBox "Calificacion Metodologia Interna Completada", vbInformation, "Aviso"
    'Me.barraEstado.Panels(1).Text = "" 'LUCV20190601, Comentó. Mejoras en el proceso de la cartera
    'Me.Barraprogreso.value = 0 'LUCV20190601, Comentó. Mejoras en el proceso de la cartera
End Sub

Private Sub cmdCalRFA_Click()
    Dim oEval As COMNCredito.NCOMColocEval
    Set oEval = New COMNCredito.NCOMColocEval
    Call oEval.CalificaRFANew(gsCodUser, fsServerConsol, CDate(txtFecha.Text), lsTablaTMP)
    Set oEval = Nothing
    MsgBox "Calificacion RFA Nueva Completada", vbInformation, "Aviso"
End Sub

Private Sub cmdCalRiesgoUnico_Click()
    'Calificacion mayor por Riesgo Unico
    Dim lsSQL As String
    Dim rs As New ADODB.Recordset
    Dim lnTotal As Long, j As Long
    Dim loConec As COMConecta.DCOMConecta
    Dim lrDat As ADODB.Recordset
    Dim lsCadConexion As String
    
'** Set loConec = New COMConecta.DCOMConecta
'**    loConec.AbreConexion 'lsCadConexion
    
'    lsSQL = " SELECT  CA1.cPersCod " _
'        & " FROM ColocCalifProv CA1 " _
'        & " Where  CA1.cPersCod IN ( SELECT  CA2.CPersCod " _
'        & "                          FROM ColocCalifProv CA2 " _
'        & "                          WHERE CA2.cPersCod =CA1.cPersCod " _
'        & "                          AND CA1.cCtaCod <> CA2.cCtaCod )" _
'        & " GROUP BY CA1.cPersCod "
'
'    Set lrDat = loConec.CargaRecordSet(lsSQL)
'    If Not (lrDat.BOF And lrDat.EOF) Then
'        lnTotal = lrDat.RecordCount
'        J = 0
'        Do While Not lrDat.EOF
'            J = J + 1
'
'            'lsSQL = "UPDATE ColocCalifProv " _
'                & " SET cCalRUnico = ( Select MAX(CA1.cCalNor)  " _
'                & "                  From ColocCalifProv CA1 Join " & fsServerConsol & "ProductoPersonaConsol PP1 " _
'                & "                  ON CA1.cCtaCod = PP1.cCtaCod " _
'                & "                  WHERE CA1.cPersCod= PP1.cPersCod " _
'                & "                  And PP1.nPrdPersRelac = 20 ) " _
'                & "WHERE cCtaCod ='" & Trim(lrDat!cCtaCod) & "'"
'            '12/01/2005- layg
'            lsSQL = "UPDATE ColocCalifProv " _
'                & " SET cCalRUnico = ( Select MAX(CA1.cCalNor)  " _
'                & "                  From ColocCalifProv CA1 Join ProductoPersona PP1 " _
'                & "                  ON CA1.cPersCod = PP1.cPersCod " _
'                & "                  WHERE PP1.nPrdPersRelac in(20,22,25) " _
'                & "                  And CA1.cPersCod='" & Trim(lrDat!cPersCod) & "'  ) " _
'                & "WHERE cPersCod ='" & Trim(lrDat!cPersCod) & "'"
'
'            loConec.Ejecutar lsSQL
'
'            Me.barraEstado.Panels(1).Text = "Cal. R. Unico :" & lrDat!cPersCod & " - " & Format(J / lnTotal * 100, "#,#0.00") & "%"
'            Me.Barraprogreso.value = Int(J / lnTotal * 100)
'            DoEvents
'            lrDat.MoveNext
'        Loop
'
'    End If
'    Set lrDat = Nothing
    
'** Set loConec = Nothing
MsgBox "Calificacion Riesgo Unico Completada", vbInformation, "Aviso"
'Me.barraestado.Panels(1).Text = ""
'Me.barraProgreso.value = 0
End Sub

Private Sub cmdEndeudamientoSF_Click()
    Dim oEval As COMNCredito.NCOMColocEval
    Set oEval = New COMNCredito.NCOMColocEval
    Call oEval.EndeudamientoSistFinanc(CDbl(txtTipoCambio.Text), fsBDRCC)
    Set oEval = Nothing
    MsgBox "Actualizacion Generada Correctamente , Termino : " & Time(), vbInformation, "Aviso"
    'MsgBox "Actualizacion Generada Correctamente , Termino : " & Time(), vbInformation, "Aviso"
    'Me.barraEstado.Panels(1).Text = ""
    'Me.Barraprogreso.value = 0

End Sub

Private Sub cmdsalir_Click()
Dim oEval As COMNCredito.NCOMColocEval
Set oEval = New COMNCredito.NCOMColocEval
    Call oEval.VerificaTablaTemporal(gsCodUser, lsTablaTMP)
Set oEval = Nothing
Unload Me
End Sub
'CTI3 ERS0032020
Private Sub cmdSeguimiento_Click()
 Dim sMensaje As String
    Dim oNCOMColocEval As COMNCredito.NCOMColocEval
    Set oNCOMColocEval = New COMNCredito.NCOMColocEval
    Dim sFecha As String
    Dim dHoraIni As String
    Dim dHoraFin As String
    
On Error GoTo ErrorProvisionProciclica
    
    Screen.MousePointer = vbHourglass
    dHoraIni = Time()
    sFecha = Format(txtFecha.Text, "YYYYMMDD")
    
    If MsgBox("Se relizará el proceso de obtener los datos del seguimiento de Sobre-Endeudados, Desea Continuar?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
    
    Call oNCOMColocEval.nCalificaCalculaDatosSeguimiento(sFecha, sMensaje)

    Set oNCOMColocEval = Nothing
    Screen.MousePointer = vbDefault

    If sMensaje <> "" Then
        MsgBox sMensaje, vbInformation, "Mensaje"
        cmdSeguimiento.BackColor = &HC0C0FF
        Exit Sub
    End If
        
    cmdSeguimiento.BackColor = &HC0FFC0
    dHoraFin = Time()
    
    MsgBox "El proceso: Obtención de datos de seguimiento de Sobre-Endeudados " & Chr(10) & _
           "Se realizó de manera satisfactoria. " & Chr(10) & Chr(10) & _
           "Hora Inicio: " & dHoraIni & " " & Chr(10) & _
           "Hora Fin   : " & dHoraFin & " ", vbInformation, "Aviso"
           
Exit Sub
ErrorProvisionProciclica:
    MsgBox "Error Nº[" & str(Err.Number) & "] " & Err.Description, vbInformation, "Aviso"
End Sub
'CTI3 END
Private Sub Form_Load()
Me.Icon = LoadPicture(App.Path & gsRutaIcono)
Dim loConstS As COMDConstSistema.NCOMConstSistema
Dim loTipCambio As COMDConstSistema.NCOMTipoCambio
Dim oEval As COMNCredito.NCOMColocEval

    Set loConstS = New COMDConstSistema.NCOMConstSistema
        fdFechaFinMes = CDate(loConstS.LeeConstSistema(gConstSistCierreMesNegocio))
        txtFecha.Text = fdFechaFinMes
        fsServerConsol = loConstS.LeeConstSistema(gConstSistServCentralRiesgos)
        fsServerRCC = loConstS.LeeConstSistema(143)
        fsBDRCC = loConstS.LeeConstSistema(144)
    Set loConstS = Nothing

    Set loTipCambio = New COMDConstSistema.NCOMTipoCambio
        fnTipoCambio = Format(loTipCambio.EmiteTipoCambio(gdFecSis, TCFijoMes), "0.###")
        txtTipoCambio.Text = fnTipoCambio
    Set loTipCambio = Nothing
    
    Set oEval = New COMNCredito.NCOMColocEval
    Me.lblFecAlin = oEval.GetFechaAlin(fsBDRCC)
    Set oEval = Nothing
End Sub

Private Function Valida() As Boolean
Dim i As Integer
Valida = True

If ValFecha(txtFecha) = False Then
    Valida = False
    Exit Function
End If
End Function

Private Function VerificaDatosIngresados() As Boolean
Dim lbOk As Boolean
lbOk = True
If Not IsDate(Me.txtFecha.Text) Then
    lbOk = False
End If

If Not IsNumeric(Me.txtTipoCambio.Text) Then
    lbOk = False
Else
    fnTipoCambio = Format(Me.txtTipoCambio.Text, "#,#0.000")
End If
VerificaDatosIngresados = lbOk
End Function

'LUCV20190601, Según Mejoras en el proceso de la cartera
Private Sub cmdLLenaCreditoAudi_Click()
    Dim sMensaje As String
    Dim oNCOMColocEval As COMNCredito.NCOMColocEval
    Dim sFecha As String
    Dim nTipoCambio As String
    Dim i As Integer
    Dim sTipoCredito As String
    Dim dHoraIni As String
    Dim dHoraFin As String
    
On Error GoTo ErrorConexion

    If VerificaDatosIngresados = False Then Exit Sub
    If MsgBox("Se Reprocesaran los Datos para la Calificacion, Desea Continuar?", vbQuestion + vbYesNo, "Aviso") = vbNo Then
        Exit Sub
    End If

    frmColocEvalPorProducto.Show 1
    dHoraIni = Time()
    Screen.MousePointer = vbHourglass

    'Valores de parámetros
    i = 0
    sFecha = Format(txtFecha.Text, "YYYYMMDD")
    nTipoCambio = CDbl(txtTipoCambio.Text)

    If Not IsArray(frmColocEvalPorProducto.MatIndices) Then
       sTipoCredito = "000000000"
    Else
        For i = 0 To UBound(frmColocEvalPorProducto.MatIndices) - 1
            sTipoCredito = sTipoCredito & "" & frmColocEvalPorProducto.MatIndices(i)
        Next i
    End If

    'Generar Datos
    Set oNCOMColocEval = New COMNCredito.NCOMColocEval
    Call oNCOMColocEval.nCalificaPreparaArchivoCalificacionNuevo(sFecha, nTipoCambio, sTipoCredito, sMensaje)

    Set oNCOMColocEval = Nothing
    Screen.MousePointer = vbDefault

    'Valida ejecución
    If sMensaje <> "" Then
        MsgBox sMensaje, vbInformation, "Mensaje"
        cmdLLenaCreditoAudi.BackColor = &HC0C0FF
        Exit Sub
    End If

    cmdLLenaCreditoAudi.BackColor = &HC0FFC0
    dHoraFin = Time()
    
    'Mensaje ejecución correcta
    MsgBox "El proceso: Preparar archivo para calificación. " & Chr(10) & _
           "Se realizó de manera satisfactoria." & Chr(10) & Chr(10) & _
           "Hora Inicio: " & dHoraIni & " " & Chr(10) & _
           "Hora Fin   : " & dHoraFin & " ", vbInformation, "Aviso"
Exit Sub
ErrorConexion:
    MsgBox "Error Nº[" & str(Err.Number) & "] " & Err.Description, vbInformation, "Aviso"
End Sub

Private Sub cmdActualizaGarantias_Click()
   Dim sMensaje As String
   Dim oNCOMColocEval As COMNCredito.NCOMColocEval
   Set oNCOMColocEval = New COMNCredito.NCOMColocEval
   Dim dHoraIni As String
   Dim dHoraFin As String
    
On Error GoTo ErrorActualizaGarantias
   
   Screen.MousePointer = vbHourglass
   dHoraIni = Time()
   Call oNCOMColocEval.nCalificaActualizarGarantiasNuevo(fnTipoCambio, sMensaje)
   
   Set oNCOMColocEval = Nothing
   Screen.MousePointer = vbDefault
    
   If sMensaje <> "" Then
        MsgBox sMensaje, vbInformation, "Mensaje"
        cmdActualizaGarantias.BackColor = &HC0C0FF
        Exit Sub
    End If
    
    cmdActualizaGarantias.BackColor = &HC0FFC0
    dHoraFin = Time()
    
    MsgBox "El proceso: Califica Garantías. " & Chr(10) & _
           "Se realizó de manera satisfactoria. " & Chr(10) & Chr(10) & _
           "Hora Inicio: " & dHoraIni & " " & Chr(10) & _
           "Hora Fin   : " & dHoraFin & " ", vbInformation, "Aviso"
Exit Sub
ErrorActualizaGarantias:
    MsgBox "Error Nº[" & str(Err.Number) & "] " & Err.Description, vbInformation, "Aviso"
End Sub

Private Sub cmdCalCMAC_Click()
    Dim sMensaje As String
    Dim oNCOMColocEval As COMNCredito.NCOMColocEval
    Set oNCOMColocEval = New COMNCredito.NCOMColocEval
    Dim dHoraIni As String
    Dim dHoraFin As String
    
On Error GoTo ErrorCMACM

    Screen.MousePointer = vbHourglass
    dHoraIni = Time()
    Call oNCOMColocEval.nCalificaCMACTNuevo(fnTipoCambio, sMensaje)
    
    Set oNCOMColocEval = Nothing
    Screen.MousePointer = vbDefault
    
    'Valida ejecución
    If sMensaje <> "" Then
        MsgBox sMensaje, vbInformation, "Mensaje"
        cmdCalCMAC.BackColor = &HC0C0FF
        Exit Sub
    End If
    
    cmdCalCMAC.BackColor = &HC0FFC0
    dHoraFin = Time()
    
    MsgBox "El proceso: Calificación CMACM " & Chr(10) & _
           "Se realizó de manera satisfactoria. " & Chr(10) & Chr(10) & _
           "Hora Inicio: " & dHoraIni & " " & Chr(10) & _
           "Hora Fin   : " & dHoraFin & " ", vbInformation, "Aviso"
    
Exit Sub
ErrorCMACM:
    MsgBox "Error Nº[" & str(Err.Number) & "] " & Err.Description, vbInformation, "Aviso"
End Sub

Private Sub cmdCalEvaluacion_Click()
    Dim sMensaje As String
    Dim oNCOMColocEval As COMNCredito.NCOMColocEval
    Set oNCOMColocEval = New COMNCredito.NCOMColocEval
    Dim sFecha As String
    Dim dHoraIni As String
    Dim dHoraFin As String
    
On Error GoTo ErrorEvaluacionRiesgos

    Screen.MousePointer = vbHourglass
    dHoraIni = Time()
    sFecha = Format(txtFecha.Text, "YYYYMMDD")
    Call oNCOMColocEval.nCalificaEvaluacionRiesgosNuevo(sFecha, sMensaje)
    
    Set oNCOMColocEval = Nothing
    Screen.MousePointer = vbDefault
    
    'Valida ejecución
    If sMensaje <> "" Then
        MsgBox sMensaje, vbInformation, "Mensaje"
        cmdCalEvaluacion.BackColor = &HC0C0FF
        Exit Sub
    End If
    
    cmdCalEvaluacion.BackColor = &HC0FFC0
    dHoraFin = Time()
    
    MsgBox "El proceso: Calificación Evaluación Riesgos " & Chr(10) & _
           "Se realizó de manera satisfactoria. " & Chr(10) & Chr(10) & _
           "Hora Inicio: " & dHoraIni & " " & Chr(10) & _
           "Hora Fin   : " & dHoraFin & " ", vbInformation, "Aviso"
    Exit Sub
ErrorEvaluacionRiesgos:
    MsgBox "Error Nº[" & str(Err.Number) & "] " & Err.Description, vbInformation, "Aviso"
End Sub

Private Sub cmdCalSistF_Click()
    Dim sMensaje As String
    Dim oNCOMColocEval As COMNCredito.NCOMColocEval
    Set oNCOMColocEval = New COMNCredito.NCOMColocEval
    Dim sFecha As String
    Dim dHoraIni As String
    Dim dHoraFin As String
    
On Error GoTo ErrorCalificacionSistemaFinanciero
    
    Screen.MousePointer = vbHourglass
    dHoraIni = Time()
    sFecha = Format(CDate(lblFecAlin.Caption), "YYYYMMDD")
    
    Call oNCOMColocEval.nCalificaSistemaFinancieroNuevo(sFecha, sMensaje)
    
    Set oNCOMColocEval = Nothing
    Screen.MousePointer = vbDefault
    
    If sMensaje <> "" Then
        MsgBox sMensaje, vbInformation, "Mensaje"
        cmdCalSistF.BackColor = &HC0C0FF
        Exit Sub
    End If

    cmdCalSistF.BackColor = &HC0FFC0
    dHoraFin = Time()
    
    MsgBox "El proceso: Calificación Sistema Financiero " & Chr(10) & _
           "Se realizó de manera satisfactoria. " & Chr(10) & Chr(10) & _
           "Hora Inicio: " & dHoraIni & " " & Chr(10) & _
           "Hora Fin   : " & dHoraFin & " ", vbInformation, "Aviso"
    Exit Sub
ErrorCalificacionSistemaFinanciero:
    MsgBox "Error Nº[" & str(Err.Number) & "] " & Err.Description, vbInformation, "Aviso"
End Sub

Private Sub cmdCalGen_Click()
    Dim sMensaje As String
    Dim oNCOMColocEval As COMNCredito.NCOMColocEval
    Set oNCOMColocEval = New COMNCredito.NCOMColocEval
    Dim sFecha As String
    Dim dHoraIni As String
    Dim dHoraFin As String

On Error GoTo ErrorCalificacionGeneral

    Screen.MousePointer = vbHourglass
    dHoraIni = Time()
    sFecha = Format(txtFecha.Text, "YYYYMMDD")

    Call oNCOMColocEval.nCalificaGeneralNuevo(sFecha, sMensaje)

    Set oNCOMColocEval = Nothing
    Screen.MousePointer = vbDefault

    If sMensaje <> "" Then
        MsgBox sMensaje, vbInformation, "Mensaje"
        cmdCalGen.BackColor = &HC0C0FF
        Exit Sub
    End If

    cmdCalGen.BackColor = &HC0FFC0
    dHoraFin = Time()

    MsgBox "El proceso: Calificación General " & Chr(10) & _
           "Se realizó de manera satisfactoria. " & Chr(10) & Chr(10) & _
           "Hora Inicio: " & dHoraIni & " " & Chr(10) & _
           "Hora Fin   : " & dHoraFin & " ", vbInformation, "Aviso"

    Exit Sub
ErrorCalificacionGeneral:
    MsgBox "Error Nº[" & str(Err.Number) & "] " & Err.Description, vbInformation, "Aviso"
End Sub

Private Sub cmdCalculaProvision_Click()
    Dim sMensaje As String
    Dim oNCOMColocEval As COMNCredito.NCOMColocEval
    Set oNCOMColocEval = New COMNCredito.NCOMColocEval
    Dim sFecha As String
    Dim sTipoCambio As String
    Dim dHoraIni As String
    Dim dHoraFin As String
    
On Error GoTo ErrorCalificaProvision

    Screen.MousePointer = vbHourglass
    dHoraIni = Time()
    sFecha = Format(txtFecha.Text, "YYYYMMDD")
    sTipoCambio = CStr(txtTipoCambio.Text)
    
    Call oNCOMColocEval.nCalificaCalculaProvisionNuevo(sFecha, sTipoCambio, sMensaje)
    
    Set oNCOMColocEval = Nothing
    Screen.MousePointer = vbDefault

    If sMensaje <> "" Then
        MsgBox sMensaje, vbInformation, "Mensaje"
        cmdCalculaProvision.BackColor = &HC0C0FF
        Exit Sub
    End If
        
    cmdCalculaProvision.BackColor = &HC0FFC0
    dHoraFin = Time()
    
    MsgBox "El proceso: Calcula Provisión " & Chr(10) & _
           "Se realizó de manera satisfactoria. " & Chr(10) & Chr(10) & _
           "Hora Inicio: " & dHoraIni & " " & Chr(10) & _
           "Hora Fin   : " & dHoraFin & " ", vbInformation, "Aviso"
           
Exit Sub
ErrorCalificaProvision:
    MsgBox "Error Nº[" & str(Err.Number) & "] " & Err.Description, vbInformation, "Aviso"
End Sub

Private Sub cmdCalSinAlineamiento_Click()
    Dim sMensaje As String
    Dim oNCOMColocEval As COMNCredito.NCOMColocEval
    Set oNCOMColocEval = New COMNCredito.NCOMColocEval
    Dim sFecha As String
    Dim sTipoCambio As String
    Dim dHoraIni As String
    Dim dHoraFin As String
    
On Error GoTo ErrorCalificaSinAlineamiento

    Screen.MousePointer = vbHourglass
    dHoraIni = Time()
    
    Call oNCOMColocEval.generarCalificacionSinAlineamientoNuevo(sMensaje)
    
    Set oNCOMColocEval = Nothing
    Screen.MousePointer = vbDefault

    If sMensaje <> "" Then
        MsgBox sMensaje, vbInformation, "Mensaje"
        cmdCalSinAlineamiento.BackColor = &HC0C0FF
        Exit Sub
    End If
    
    cmdCalSinAlineamiento.BackColor = &HC0FFC0
    dHoraFin = Time()
    
    MsgBox "El proceso: Calificación Sin Alineamiento " & Chr(10) & _
           "Se realizó de manera satisfactoria. " & Chr(10) & Chr(10) & _
           "Hora Inicio: " & dHoraIni & " " & Chr(10) & _
           "Hora Fin   : " & dHoraFin & " ", vbInformation, "Aviso"

Exit Sub
ErrorCalificaSinAlineamiento:
    MsgBox "Error Nº[" & str(Err.Number) & "] " & Err.Description, vbInformation, "Aviso"
End Sub

Private Sub cmdCalculaProvSinAlinea_Click()
    Dim sMensaje As String
    Dim oNCOMColocEval As COMNCredito.NCOMColocEval
    Set oNCOMColocEval = New COMNCredito.NCOMColocEval
    Dim sFecha As String
    Dim dHoraIni As String
    Dim dHoraFin As String
    
On Error GoTo ErrorProvisionSinAlineamiento

    Screen.MousePointer = vbHourglass
    dHoraIni = Time()
    sFecha = Format(txtFecha.Text, "YYYYMMDD")
    
    Call oNCOMColocEval.calculaProvisionSinAlineamientoNuevo(sFecha, sMensaje)
    
    Set oNCOMColocEval = Nothing
    Screen.MousePointer = vbDefault

    If sMensaje <> "" Then
        MsgBox sMensaje, vbInformation, "Mensaje"
        cmdCalculaProvSinAlinea.BackColor = &HC0C0FF
        Exit Sub
    End If
        
    cmdCalculaProvSinAlinea.BackColor = &HC0FFC0
    dHoraFin = Time()
    
    MsgBox "El proceso: Calcula Provisión Sin Alineamiento. " & Chr(10) & _
           "Se realizó de manera satisfactoria. " & Chr(10) & Chr(10) & _
           "Hora Inicio: " & dHoraIni & " " & Chr(10) & _
           "Hora Fin   : " & dHoraFin & " ", vbInformation, "Aviso"
    
Exit Sub
ErrorProvisionSinAlineamiento:
    MsgBox "Error Nº[" & str(Err.Number) & "] " & Err.Description, vbInformation, "Aviso"
End Sub

Private Sub cmdCalculaProvSistF_Click()
    Dim sMensaje As String
    Dim oNCOMColocEval As COMNCredito.NCOMColocEval
    Set oNCOMColocEval = New COMNCredito.NCOMColocEval
    Dim sFecha As String
    Dim dHoraIni As String
    Dim dHoraFin As String
    
On Error GoTo ErrorProvisionCalificacionSistFinanciero
    
    Screen.MousePointer = vbHourglass
    dHoraIni = Time()
    sFecha = Format(txtFecha.Text, "YYYYMMDD")
    
    Call oNCOMColocEval.calculaProvisionSistemaFinancieroNuevo(sFecha, sMensaje)
   
    Set oNCOMColocEval = Nothing
    Screen.MousePointer = vbDefault

    If sMensaje <> "" Then
        MsgBox sMensaje, vbInformation, "Mensaje"
        cmdCalculaProvSistF.BackColor = &HC0C0FF
        Exit Sub
    End If
        
    cmdCalculaProvSistF.BackColor = &HC0FFC0
    dHoraFin = Time()
    
    MsgBox "El proceso: Calcula Provisión Según Calificacion Sistema Finaciero. " & Chr(10) & _
           "Se realizó de manera satisfactoria. " & Chr(10) & Chr(10) & _
           "Hora Inicio: " & dHoraIni & " " & Chr(10) & _
           "Hora Fin   : " & dHoraFin & " ", vbInformation, "Aviso"
    
Exit Sub
ErrorProvisionCalificacionSistFinanciero:
    MsgBox "Error Nº[" & str(Err.Number) & "] " & Err.Description, vbInformation, "Aviso"
End Sub

Private Sub cmdProvisionProciclica_Click()
    Dim sMensaje As String
    Dim oNCOMColocEval As COMNCredito.NCOMColocEval
    Set oNCOMColocEval = New COMNCredito.NCOMColocEval
    Dim sFecha As String
    Dim dHoraIni As String
    Dim dHoraFin As String
    
On Error GoTo ErrorProvisionProciclica
    
    Screen.MousePointer = vbHourglass
    dHoraIni = Time()
    sFecha = Format(txtFecha.Text, "YYYYMMDD")
    
    If MsgBox("Se relizará el proceso de Cálculo Provisión Prociclica, Desea Continuar?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
    
    Call oNCOMColocEval.nCalificaCalculaProvisionProciclicaNuevo(sFecha, sMensaje)

    Set oNCOMColocEval = Nothing
    Screen.MousePointer = vbDefault

    If sMensaje <> "" Then
        MsgBox sMensaje, vbInformation, "Mensaje"
        cmdProvisionProciclica.BackColor = &HC0C0FF
        Exit Sub
    End If
        
    cmdProvisionProciclica.BackColor = &HC0FFC0
    dHoraFin = Time()
    
    MsgBox "El proceso: Calcula Provisión Prociclica " & Chr(10) & _
           "Se realizó de manera satisfactoria. " & Chr(10) & Chr(10) & _
           "Hora Inicio: " & dHoraIni & " " & Chr(10) & _
           "Hora Fin   : " & dHoraFin & " ", vbInformation, "Aviso"
           
Exit Sub
ErrorProvisionProciclica:
    MsgBox "Error Nº[" & str(Err.Number) & "] " & Err.Description, vbInformation, "Aviso"
End Sub

Private Sub cmdCalifRiesgoCredCambiario_Click()
    Dim sMensaje As String
    Dim oNCOMColocEval As COMNCredito.NCOMColocEval
    Set oNCOMColocEval = New COMNCredito.NCOMColocEval
    Dim dHoraIni As String
    Dim dHoraFin As String
    
On Error GoTo ErrorRiesgoCambiario
    
    Screen.MousePointer = vbHourglass
    dHoraIni = Time()
    Call oNCOMColocEval.nCalificaRiesgoCredCambiarioNuevo(sMensaje)
    
    Set oNCOMColocEval = Nothing
    Screen.MousePointer = vbDefault

    If sMensaje <> "" Then
        MsgBox sMensaje, vbInformation, "Mensaje"
        cmdCalifRiesgoCredCambiario.BackColor = &HC0C0FF
        Exit Sub
    End If
    
    cmdCalifRiesgoCredCambiario.BackColor = &HC0FFC0
    dHoraFin = Time()
    
    MsgBox "El proceso: Riesgo Crediticio Cambiario. " & Chr(10) & _
           "Se realizó de manera satisfactoria. " & Chr(10) & _
           "Hora Inicio: " & dHoraIni & " " & Chr(10) & _
           "Hora Fin: " & dHoraFin & " ", vbInformation, "Aviso"
    
    Exit Sub
ErrorRiesgoCambiario:
    MsgBox "Error Nº[" & str(Err.Number) & "] " & Err.Description, vbInformation, "Aviso"
End Sub

Private Sub cmdCalifMetodologiaInterna_Click()
    Dim sMensaje As String
    Dim oNCOMColocEval As COMNCredito.NCOMColocEval
    Set oNCOMColocEval = New COMNCredito.NCOMColocEval
    Dim sFecha As String
    Dim dHoraIni As String
    Dim dHoraFin As String
    
On Error GoTo ErrorMetodologíaInterna

    Screen.MousePointer = vbHourglass
    dHoraIni = Time()
    sFecha = Format(txtFecha.Text, "YYYYMMDD")
    
    Call oNCOMColocEval.nCalificaMetodologiaInternaNuevo(sFecha, sMensaje)
    
    Set oNCOMColocEval = Nothing
    Screen.MousePointer = vbDefault

    If sMensaje <> "" Then
        MsgBox sMensaje, vbInformation, "Mensaje"
        cmdCalifMetodologiaInterna.BackColor = &HC0C0FF
        Exit Sub
    End If
    
    cmdCalifMetodologiaInterna.BackColor = &HC0FFC0
    dHoraFin = Time()
    
    MsgBox "El proceso: Metodología Interna. " & Chr(10) & _
           "Se realizó de manera satisfactoria. " & Chr(10) & _
           "Hora Inicio: " & dHoraIni & " " & Chr(10) & _
           "Hora Fin: " & dHoraFin & " ", vbInformation, "Aviso"

    Exit Sub
ErrorMetodologíaInterna:
    MsgBox "Error Nº[" & str(Err.Number) & "] " & Err.Description, vbInformation, "Aviso"

End Sub
'Fin LUCV20190601


