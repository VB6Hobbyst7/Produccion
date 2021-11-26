VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmColocCalSist 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Colocaciones - Procesar Calificación 808"
   ClientHeight    =   5670
   ClientLeft      =   4215
   ClientTop       =   2100
   ClientWidth     =   7065
   Icon            =   "frmColocCalSist.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   7065
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ProgressBar Barraprogreso 
      Height          =   210
      Left            =   3060
      TabIndex        =   7
      Top             =   5430
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   370
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Frame Frame1 
      Height          =   5160
      Left            =   195
      TabIndex        =   5
      Top             =   120
      Width           =   3645
      Begin VB.CommandButton cmdCalRFA 
         Caption         =   "Calificación Creditos RFA"
         Height          =   420
         Left            =   195
         TabIndex        =   18
         Top             =   825
         Width           =   3165
      End
      Begin VB.CommandButton cmdCalRiesgoUnico 
         Caption         =   "Calificación Riesgo Unico"
         Height          =   420
         Left            =   195
         TabIndex        =   16
         Top             =   1845
         Width           =   3165
      End
      Begin VB.CommandButton cmdActualizaGarantias 
         Caption         =   "Califica Garantias"
         Height          =   420
         Left            =   195
         TabIndex        =   15
         Top             =   1320
         Width           =   3165
      End
      Begin VB.CommandButton cmdLLenaCreditoAudi 
         Caption         =   "Preparar Archivo para Calificacion"
         Height          =   420
         Left            =   195
         TabIndex        =   12
         Top             =   240
         Width           =   3165
      End
      Begin VB.CommandButton cmdCalculaProvision 
         Caption         =   "Calcula &Provision"
         Height          =   420
         Left            =   195
         TabIndex        =   11
         Top             =   4605
         Width           =   3165
      End
      Begin VB.CommandButton cmdCalSistF 
         Caption         =   "Calificación  &Sistema Financiero"
         Height          =   420
         Left            =   195
         TabIndex        =   10
         Top             =   3495
         Width           =   3165
      End
      Begin VB.CommandButton cmdCalEvaluacion 
         Caption         =   "Calificación  &Evaluacion Cartera"
         Height          =   420
         Left            =   195
         TabIndex        =   9
         Top             =   2955
         Width           =   3165
      End
      Begin VB.CommandButton cmdCalCMAC 
         Caption         =   "Calificación CMAC"
         Height          =   420
         Left            =   195
         TabIndex        =   8
         Top             =   2400
         Width           =   3165
      End
      Begin VB.CommandButton cmdCalGen 
         Caption         =   "Calificación &General"
         Height          =   420
         Left            =   195
         TabIndex        =   6
         Top             =   4050
         Width           =   3165
      End
      Begin VB.CommandButton cmdEndeudamientoSF 
         Caption         =   "Endeudamiento Sistema Financiero"
         Height          =   420
         Left            =   210
         TabIndex        =   17
         Top             =   4830
         Visible         =   0   'False
         Width           =   3165
      End
   End
   Begin VB.Frame FraFecha 
      Height          =   1410
      Left            =   3960
      TabIndex        =   2
      Top             =   240
      Width           =   2535
      Begin VB.TextBox txtTipoCambio 
         Alignment       =   2  'Center
         Height          =   330
         Left            =   1080
         TabIndex        =   13
         Top             =   900
         Width           =   1305
      End
      Begin MSMask.MaskEdBox txtFecha 
         Height          =   330
         Left            =   1080
         TabIndex        =   3
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
         Caption         =   "Tip.Cambio"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   900
         Width           =   795
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha "
         Height          =   240
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   840
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   420
      Left            =   5235
      TabIndex        =   1
      Top             =   4800
      Width           =   1545
   End
   Begin MSComctlLib.StatusBar barraEstado 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   5370
      Width           =   7065
      _ExtentX        =   12462
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6703
            MinWidth        =   6703
         EndProperty
      EndProperty
   End
   Begin VB.Label lblFecAlin 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   5535
      TabIndex        =   20
      Top             =   1875
      Width           =   1035
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Fecha de Alineacion:"
      Height          =   195
      Left            =   3975
      TabIndex        =   19
      Top             =   1905
      Width           =   1500
   End
End
Attribute VB_Name = "frmColocCalSist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************
'* COLOCACIONES - CALIFICACION SISTEMA
'Archivo:  frmColocCalSist.frm
'LAYG   :  01/10/2002.-
'CAJA ICA : 26/06/2004 - 23/09/2004
'Resumen:  Realiza el Proceso de Calificacion de la Cartera

'Option Explicit
'
'Dim fnTipoCambio  As Currency
'Dim fdFechaFinMes  As Date
'Dim fsServerConsol As String
'Dim fsServerRCC As String
'Dim fsBDRCC As String
'Dim lsTablaTMP As String
'
'Private Sub cmdActualizaGarantias_Click()
'Dim oEval As COMNCredito.NCOMColocEval
'
'Set oEval = New COMNCredito.NCOMColocEval
'Call oEval.ActualizarGarantias(gdFecSis, CDbl(txtTipoCambio.Text))
'Set oEval = Nothing
'MsgBox "Actualizacion Generada Correctamente , Termino : " & Time(), vbInformation, "Aviso"
'
'' Actualiza las Garantias Reales de los Creditos
''Dim loConec As COMConecta.DCOMConecta
''Dim loCalif As COMNCredito.NCOMColocEval
''Dim lsSQL As String
''Dim Rs As New ADODB.Recordset
''
''Dim lnTotal As Long, J As Long
''Dim lnGarant As Integer
''
''Set loConec = New COMConecta.DCOMConecta
''    loConec.AbreConexion
''
''    'Todo en Soles
''    lsSQL = " SELECT CA.cCtaCod, nPrdEstado, dFecVig, cCalNor, nGarAutoL, nGarMuyRR, nGarPref,   " ''        & " (Case When substring(ca.cCtaCod,9,1) = '1' then nSaldoCap " ''        & "       When substring(ca.cCtaCod,9,1) = '2' then nSaldoCap * " & CDbl(Me.txtTipoCambio.Text) ''        & "     End ) nSaldoCap " ''        & " FROM  ColocCalifProv CA "
''
''    Set Rs = loConec.CargaRecordSet(lsSQL)
''    lnTotal = Rs.RecordCount
''    J = 0
''    Do While Not Rs.EOF
''        J = J + 1
''        If (Rs!nGarAutoL >= Rs!nSaldoCap) And (Rs!nSaldoCap > 0) Then
''            lnGarant = 1
''        ElseIf (Rs!nGarMuyRR >= Rs!nSaldoCap) And Mid(Rs!cCtaCod, 6, 1) <> "3" Then
''            lnGarant = 2
''        'Se ha anulado por que se usa la tabla ColocCalifGarantia
''        'ElseIf rs!nGarPref >= rs!nSaldoCap And Mid(rs!cCtaCod, 6, 1) <> "3" Then
''        '    lnGarant = 3
''        Else
''            lnGarant = 4
''        End If
''        '
''        If Mid(Rs!cCtaCod, 6, 1) = "4" And lnGarant = 4 Then    ' A todos los hipotecarios => Garant Preferidas
''            lnGarant = 4  '3
''        End If
''
''        ' Si tienen garantia preferida y son judiciales
''        ' ==>> Es dudoso y  los dias de atraso > 24 meses =  720 dias de Ingreso a Judicial
''        ' ==>> Es perdida y  los dias de atraso > 36 meses =  1080 dias de Ingreso a Judicial
''        If lnGarant = 3 And Rs!nPrdEstado = gColocEstRecVigJud Then
''            If DateDiff("d", Rs!dFecVig, gdFecSis) > 720 Then
''                lnGarant = 4
''            End If
''            'If rs!cCalNor = "4" And DateDiff("d", rs!dFecVig, gdFecSis) > 1080 Then
''            '    lnGarant = 4
''            'End If
''        End If
''
''        lsSQL = "UPDATE ColocCalifProv " ''            & "SET nGarant = " & lnGarant ''            & "WHERE cCtaCod ='" & Trim(Rs!cCtaCod) & "'"
''
''        loConec.Ejecutar lsSQL
''
''        Me.barraEstado.Panels(1).Text = "Garantia :" & Rs!cCtaCod & " - " & Format(J / lnTotal * 100, "#,#0.00") & "%"
''        Me.Barraprogreso.value = Int(J / lnTotal * 100)
''        DoEvents
''        Rs.MoveNext
''    Loop
''    Rs.Close
''    Set Rs = Nothing
''    '*****************************************
''    '** Actualiza la Garantias Preferidas ICA 25/06/2004
''    lnGarant = 3
''    'lsSQL = " Update ColocCalifProv set nGarant = " & lnGarant & " Where cCtaCod in " ''          & "   (Select GCC.cCtaCod from ColocCalifGarantia GPref " ''          & "    join " & fsServerConsol & "GarantCredConsol GCC on GPref.cNumGarant = GCC.cNumGarant " ''          & "     Where nEstadoGar = 0 )"
''    ' 12/01/2005 - layg
''    lsSQL = " Update ColocCalifProv set nGarant = " & lnGarant & " Where cCtaCod in " ''          & "   (Select GCC.cCtaCod from ColocCalifGarantia GPref " ''          & "    join ColocGarantia GCC on GPref.cNumGarant = GCC.cNumGarant " ''          & "     Where nEstadoGar = 0 )"
''    loConec.Ejecutar lsSQL
''
''Set loConec = Nothing
''
''MsgBox "Actualizacion Generada Correctamente , Termino : " & Time(), vbInformation, "Aviso"
''Me.barraEstado.Panels(1).Text = ""
''Me.Barraprogreso.value = 0
'
'End Sub
'
'Private Sub cmdCalCMAC_Click()
'
'Dim oEval As COMNCredito.NCOMColocEval
'
'Set oEval = New COMNCredito.NCOMColocEval
'Call oEval.CalificaCMACT
'Set oEval = Nothing
'MsgBox "Calificacion CMAC Completada", vbInformation, "Aviso"
'
'' Calificacion mayor de los productos que tenga en la Cmact
''Dim lsSQL As String
''Dim Rs As New ADODB.Recordset
''Dim lnTotal As Long, J As Long
''Dim loConec As COMConecta.DCOMConecta
''Dim lrDat As ADODB.Recordset
''Dim lsCadConexion As String
''
''Set loConec = New COMConecta.DCOMConecta
''
''    loConec.AbreConexion 'lsCadConexion
''    '*** para utilizar cCalRUnico - ICA
''    'RFA - Actualiza Calif del RFA
''    lsSQL = "UPDATE ColocCalifProv SET cCalRUnico = cCalRFA WHERE cCalRFA is not Null "
''    loConec.Ejecutar lsSQL
''
''    'Actualiza Calif Riesgo Unico con la Calificacion por dias de atraso en caso no tenga una asignada
''    lsSQL = "UPDATE ColocCalifProv SET cCalRUnico = cCalNor WHERE (cCalRUnico is Null or cCalRUnico in (' ')) "
''    loConec.Ejecutar lsSQL
''    '****
''
''    lsSQL = " SELECT  CA1.cPersCod " ''        & " FROM ColocCalifProv CA1 " ''        & " Where  CA1.cPersCod IN ( SELECT  CA2.CPersCod " ''        & "                          FROM ColocCalifProv CA2 " ''        & "                          WHERE CA2.cPersCod =CA1.cPersCod " ''        & "                          AND CA1.cCtaCod <> CA2.cCtaCod )" ''        & " GROUP BY CA1.cPersCod "
''
''    Set lrDat = loConec.CargaRecordSet(lsSQL)
''    If Not (lrDat.BOF And lrDat.EOF) Then
''        lnTotal = lrDat.RecordCount
''        J = 0
''        Do While Not lrDat.EOF
''            J = J + 1
''            '*** Se cambio a cCalRUnico - ICA (26/06/2004 LAYG)
''            'lsSQL = "UPDATE ColocCalifProv " ''                & " SET cCalCMAC = ( Select MAX(CA1.cCalNor) as cCalCaja " ''                & "                  From ColocCalifProv CA1 WHERE CA1.cPersCod='" & Trim(lrDat!cPersCod) & "') " ''                & "WHERE cPersCod ='" & Trim(lrDat!cPersCod) & "'"
''            lsSQL = "UPDATE ColocCalifProv " ''                & " SET cCalCMAC = ( Select MAX(CA1.cCalRUnico) as cCalCaja " ''                & "                  From ColocCalifProv CA1 WHERE CA1.cPersCod='" & Trim(lrDat!cPersCod) & "') " ''                & "WHERE cPersCod ='" & Trim(lrDat!cPersCod) & "'"
''
''            loConec.Ejecutar lsSQL
''
''            Me.barraEstado.Panels(1).Text = "Cal. CMACT :" & lrDat!cPersCod & " - " & Format(J / lnTotal * 100, "#,#0.00") & "%"
''            Me.Barraprogreso.value = Int(J / lnTotal * 100)
''            DoEvents
''            lrDat.MoveNext
''        Loop
''
''    End If
''    Set lrDat = Nothing
''
''    lsSQL = "UPDATE ColocCalifProv SET cCalCMAC = cCalNor WHERE (cCalCMAC is Null or cCalCMAC in (' ')) "
''    loConec.Ejecutar lsSQL
''Set loConec = Nothing
''
''MsgBox "Calificacion CMAC Completada", vbInformation, "Aviso"
''Me.barraEstado.Panels(1).Text = ""
''Me.Barraprogreso.value = 0
'
'End Sub
'
'Private Sub cmdCalculaProvision_Click()
'
'Dim oEval As COMNCredito.NCOMColocEval
'Set oEval = New COMNCredito.NCOMColocEval
'Call oEval.CalculaProvisiones(fsServerConsol, fnTipoCambio) 'ARCV 14-08-2006
'Set oEval = Nothing
'
'MsgBox "Proceso de Calificacion terminado Correctamente ", vbInformation, "Aviso"
'
'End Sub
'
'Private Sub cmdCalEvaluacion_Click()
'
'Dim oEval As COMNCredito.NCOMColocEval
'Set oEval = New COMNCredito.NCOMColocEval
'Call oEval.CalculaEvaluacion(txtFecha.Text)
'Set oEval = Nothing
'
'MsgBox "Calificacion Evaluacion de Cartera completado", vbInformation, "Aviso"
''** Calificacion de Evaluacion ASIGNADA
''** Coge el valor de ColocEvalCalifDetalle
'
''Dim lsSQL As String
''
''Dim lnTotal As Long, J As Long
''Dim lsCalEvalua As String
''Dim loConec As COMConecta.DCOMConecta
''Dim lrDat As ADODB.Recordset
''Dim lrDatEval As ADODB.Recordset
''
'''*** Obtiene Calificacion en ultima Evaluacion de Colocaciones
'''*** Asignada en la Evaluacion de la Calificacion
'''*** de la tabla ColocEvalCalif
''
''lsSQL = " SELECT  distinct(cPersCod), " ''      & " cCalEval = (SELECT Max(a.cEvalCalifDet) cCalEval FROM ColocEvalCalifDetalle a " ''      & "             WHERE a.cPersCod = ca.cPersCod " ''      & "             AND nEvalTipo = 0 AND DateDiff(dd,dEval,'" & Format(Me.txtFecha.Text, "mm/dd/yyyy") & "') = 0 ) " ''      & " FROM ColocCalifProv ca "
''
''Set loConec = New COMConecta.DCOMConecta
''    loConec.AbreConexion
''    Set lrDat = loConec.CargaRecordSet(lsSQL)
''
''    If Not (lrDat.BOF And lrDat.EOF) Then
''        lnTotal = lrDat.RecordCount
''        J = 0
''        Do While Not lrDat.EOF
''            J = J + 1
''
''            lsCalEvalua = IIf(IsNull(lrDat!cCalEval), "", lrDat!cCalEval)
''
''            ' ******************
''            If lsCalEvalua <> "" Then
''                '*** Actualiza ColocCalifProv
''                lsSQL = "UPDATE ColocCalifProv SET cCalEval ='" & lsCalEvalua & "' " ''                      & "WHERE cPersCod ='" & Trim(lrDat!cPersCod) & "' "
''                loConec.Ejecutar lsSQL
''            End If
''            '************
''
''            Me.barraEstado.Panels(1).Text = "Cal. Evaluacion :" & lrDat!cPersCod & " - " & Format(J / lnTotal * 100, "#,#0.00") & "%"
''            Me.Barraprogreso.value = Int(J / lnTotal * 100)
''
''            DoEvents
''            lrDat.MoveNext
''        Loop
''    End If
''    Set lrDat = Nothing
''Set loConec = Nothing
''
''MsgBox "Calificacion Evaluacion de Cartera completado", vbInformation, "Aviso"
''Me.barraEstado.Panels(1).Text = ""
''Me.Barraprogreso.value = 0
'End Sub
'
'
'Private Sub cmdCalGen_Click()
'Dim oEval As COMNCredito.NCOMColocEval
'Dim sMensaje As String
'
'Set oEval = New COMNCredito.NCOMColocEval
'
'Call oEval.CalificacionGeneral(sMensaje)
'Set oEval = Nothing
'
'If sMensaje <> "" Then
'    MsgBox sMensaje, vbInformation, "Mensaje"
'    Exit Sub
'End If
'
'MsgBox "Calificacion General completada ", vbInformation, "Aviso"
''** Calificacion de Comparando la Calificacion Cmac y el Sist Financiero
''Dim lsSQL As String
''Dim lrs As ADODB.Recordset
''Dim lnTotal As Long, J As Long
''Dim lsCalifGen As String
''Dim lsCalCMAC As String
''Dim lsCalSistF As String, lsCalHist As String
''Dim lsCalEval As String
''Dim rsVerif As ADODB.Recordset
''Dim loConec As COMConecta.DCOMConecta
'''
''Dim lsCalObser As String ' Calific Observada ***
'''
''Set loConec = New COMConecta.DCOMConecta
''loConec.AbreConexion
''
''lsSQL = " SELECT  distinct(cPersCod), cCalCMAC, cCalSistF, cCalEval, cFlagProv " ''      & " FROM ColocCalifProv  "
''
''Set loConec = New COMConecta.DCOMConecta
''loConec.AbreConexion
''Set lrs = loConec.CargaRecordSet(lsSQL)
''    If lrs.BOF And lrs.EOF Then
''        MsgBox "No existen Datos para Calificar ", vbInformation, "Aviso"
''    Else
''        lnTotal = lrs.RecordCount
''        J = 0
''        Do While Not lrs.EOF
''            J = J + 1
''
''            '********************************************************************************
''            '** En caso la Ofic. Riesgos le haya asignado una calificacion, se asigna esta calificacion
''            '** Si el Cliente no tiene calificacion en el Sist. Financiero se le asigna la Calif de la CMAC
''            '** Si la Calificacion de la CMAC es mayor o igual que la del Sist. Financiero se le asigna la Calif CMAC
''            '** En caso de Calif Sist Financ sea mayor se le asigna la Calif del Sist.Financiero menos una categoria
''
''            lsCalifGen = ""
''            lsCalCMAC = lrs!cCalCMAC
''            lsCalSistF = IIf(IsNull(lrs!cCalSistF), "", lrs!cCalSistF)
''            lsCalHist = IIf(IsNull(lrs!cCalEval), "", lrs!cCalEval)
''
''            'If IsNull(lrs!cFlagProv) Then ' ***
''            '    lsCalObser = "-1"
''            'Else
''            '    lsCalObser = Trim(Val(lrs!cFlagProv))
''            'End If
''
''            If Trim(lsCalHist) <> "" Then
''                lsCalifGen = lsCalHist
''            'ElseIf Trim(lsCalObser) <> -1 And lsCalObser >= lsCalCMAC Then
''            '    lsCalifGen = lsCalObser  '  Calif Observado
''            ElseIf Trim(lsCalSistF) = "" Then
''                lsCalifGen = lsCalCMAC  '  No tiene cred en Sist Financ
''            ElseIf lsCalCMAC >= lsCalSistF Then
''                lsCalifGen = lsCalCMAC  ' Si la calif CMAC es mayor
''            Else  ' Si la CalSistF es mayor
''                lsCalifGen = lsCalSistF - 1  ' Una Calificacion menor
''            End If
''
''            '*** Actualiza ColocCalifProv
''            lsSQL = "UPDATE ColocCalifProv SET cCalGen = '" & lsCalifGen & "'" ''                  & "WHERE cPersCod ='" & Trim(lrs!cPersCod) & "'"
''            loConec.Ejecutar lsSQL
''            '*****************
''
''            Me.barraEstado.Panels(1).Text = "Cal. Gener :" & lrs!cPersCod & " - " & Format(J / lnTotal * 100, "#,#0.00") & "%"
''            Me.Barraprogreso.value = Int(J / lnTotal * 100)
''
''            DoEvents
''            lrs.MoveNext
''        Loop
''        Set lrs = Nothing
''        Set loConec = Nothing
''
''        MsgBox "Calificacion General completada ", vbInformation, "Aviso"
''        Me.barraEstado.Panels(1).Text = ""
''        Me.Barraprogreso.value = 0
''    End If
'End Sub
'
'Sub CalificaRFANEW()
'
'Dim oEval As COMNCredito.NCOMColocEval
'Set oEval = New COMNCredito.NCOMColocEval
'Call oEval.CalificaRFANEW(gsCodUser, fsServerConsol, CDate(txtFecha.Text), lsTablaTMP)
'Set oEval = Nothing
'MsgBox "Calificacion RFA Nueva Completada", vbInformation, "Aviso"
'
''' Calificacion Creditos RFA NUEVA - ACTUALIZADO EL 12 DE ABRIL DEL 2005
''Dim lsSQL As String
''Dim Rs As New ADODB.Recordset
''Dim lnTotal As Long, J As Long
''Dim loConec As COMConecta.DCOMConecta
''Dim lrDat As ADODB.Recordset
''Dim lrDatCalif As ADODB.Recordset
''Dim lsCalMaxActual As String, lsCalAnterior As String
''Dim lsCalRFA As String
''Set loConec = New COMConecta.DCOMConecta
''
''    VerificaTablaTemporal
''
''    loConec.AbreConexion
''
''    lsSQL = "         SELECT    C.CCTACOD, P.CPERSCOD, P.CPERSNOMBRE, CC.CRFA, C.NDIASATRASO, "
''    lsSQL = lsSQL & "           ISNULL(DATOS.nDiasAtrPag,C.NDIASATRASO) AS nDiasAtrPag, datos.dFecPag, datos.dFecVenc "
''    lsSQL = lsSQL & "  Into " & lsTablaTMP
''    lsSQL = lsSQL & "  FROM    " & fsServerConsol & "CREDITOCONSOL C"
''    lsSQL = lsSQL & "           FULL OUTER JOIN"
''    lsSQL = lsSQL & "               (SELECT P.cCtaCod,  P.dFecVenc, convert(datetime,Pago.dFecPag) AS dFecPag,"
''    lsSQL = lsSQL & "                       datediff(Day, P.dFecVenc, convert(DateTime, Pago.dFecPag)) As nDiasAtrPag"
''    lsSQL = lsSQL & "                From"
''    lsSQL = lsSQL & "                   ("
''    lsSQL = lsSQL & "                       SELECT  DISTINCT MD.cCtaCod,M.CMOVNRO,MD.nNroCuota,"
''    lsSQL = lsSQL & "                               CONVERT(DATETIME,SUBSTRING(M.CMOVNRO,5,2) + '/' + SUBSTRING(M.CMOVNRO,7,2) + '/' + SUBSTRING(M.CMOVNRO,1,4))  AS DFECPAG"
''    lsSQL = lsSQL & "                       FROM    MOV M "
''    lsSQL = lsSQL & "                               JOIN MOVCOL MC ON MC.NMOVNRO = M.NMOVNRO"
''    lsSQL = lsSQL & "                               JOIN MOVCOLDET MD ON MD.NMOVNRO = MC.NMOVNRO AND MD.COPECOD = MC.COPECOD AND MD.CCTACOD = MC.CCTACOD"
''    lsSQL = lsSQL & "                       WHERE   (MC.COPECOD LIKE '100[234567]%' OR MC.COPECOD LIKE '12[16][012]%') AND NMOVFLAG = 0"
''    lsSQL = lsSQL & "                               AND MD.CCTACOD IN (SELECT CCTACOD FROM COLOCACCRED WHERE CRFA IN ('RFA','RFC','DIF'))"
''    lsSQL = lsSQL & "                               AND LEFT(M.CMOVNRO,6)<= '" & Format(txtFecha, "yyyymm") & "'"
''    lsSQL = lsSQL & "                               AND M.CMOVNRO IN (  SELECT MAX(M1.CMOVNRO)"
''    lsSQL = lsSQL & "                                                   FROM    MOV M1"
''    lsSQL = lsSQL & "                                                           JOIN MOVCOL MC1 ON MC1.NMOVNRO = M1.NMOVNRO"
''    lsSQL = lsSQL & "                                                   WHERE   M1.NMOVFLAG=0 AND MC1.CCTACOD=MC.CCTACOD AND LEFT(M1.CMOVNRO,6)<= '" & Format(txtFecha, "yyyymm") & "'"
''    lsSQL = lsSQL & "                                                   AND (MC1.COPECOD LIKE '100[234567]%' OR MC1.COPECOD LIKE '12[16][012]%'))"
''    lsSQL = lsSQL & "                       ) AS PAGO"
''    lsSQL = lsSQL & "               JOIN " & fsServerConsol & "PLANDESPAGCONSOL P ON P.CCTACOD = PAGO.CCTACOD AND P.cNroCuo =  PAGO.NNROCUOTA  AND P.nTipo = 1 and P.nEstado = 1"
''    lsSQL = lsSQL & "               ) AS DATOS ON DATOS.CCTACOD = C.CCTACOD"
''    lsSQL = lsSQL & "           JOIN COLOCACCRED CC ON CC.CCTACOD = C.CCTACOD"
''    lsSQL = lsSQL & "           JOIN PRODUCTOPERSONA R ON R.CCTACOD = C.CCTACOD"
''    lsSQL = lsSQL & "           JOIN PERSONA P ON P.CPERSCOD = R.CPERSCOD AND R.NPRDPERSRELAC=20"
''    lsSQL = lsSQL & "    WHERE  CC.CRFA IN ('RFA','RFC','DIF') AND"
''    lsSQL = lsSQL & "           C.nPrdEstado in (2020,2021,2022,2030,2031,2032,2201,2205,2101,2104,2106,2107,2061,2060) "
''
''    loConec.Ejecutar (lsSQL)
''
''    lsSQL = "           SELECT  CPERSCOD, CPERSNOMBRE, AVG(nDiasAtrPag) AS nDiasAtrPag "
''    lsSQL = lsSQL & "   From "
''    lsSQL = lsSQL & "       (   SELECT *"
''    lsSQL = lsSQL & "           FROM   " & lsTablaTMP & " P"
''    lsSQL = lsSQL & "           WHERE   NOT EXISTS (SELECT P1.CPERSCOD"
''    lsSQL = lsSQL & "                               FROM " & lsTablaTMP & " P1"
''    lsSQL = lsSQL & "                               WHERE P1.CPERSCOD = P.CPERSCOD AND P1.nDiasAtrPag>0  )"
''    lsSQL = lsSQL & "                   AND NOT EXISTS (SELECT P1.CPERSCOD"
''    lsSQL = lsSQL & "                                   FROM " & lsTablaTMP & " P1"
''    lsSQL = lsSQL & "                                   WHERE P1.CPERSCOD = P.CPERSCOD AND P1.nDiasAtraso>0 )"
''    lsSQL = lsSQL & "                   AND NOT EXISTS (SELECT P1.CPERSCOD"
''    lsSQL = lsSQL & "                                   FROM " & lsTablaTMP & " P1"
''    lsSQL = lsSQL & "                                   WHERE P1.CPERSCOD = P.CPERSCOD AND P1.dFecPag is null )"
''    lsSQL = lsSQL & "                               ) AS X "
''    lsSQL = lsSQL & "   GROUP BY CPERSCOD, CPERSNOMBRE "
''    lsSQL = lsSQL & "   ORDER BY CPERSNOMBRE "
''
''    Set lrDat = loConec.CargaRecordSet(lsSQL)
''    If Not (lrDat.BOF And lrDat.EOF) Then
''        lnTotal = lrDat.RecordCount
''        J = 0
''        Do While Not lrDat.EOF
''            J = J + 1
''            'If lrDat!cPersCod = "1080100751979" Then Stop
''
''            lsSQL = "UPDATE ColocCalifProv " ''                & " SET cCalRFA = '0' " ''                & "WHERE cPersCod ='" & Trim(lrDat!cPersCod) & "'"
''
''            loConec.Ejecutar lsSQL
''
''            Me.barraEstado.Panels(1).Text = "Cal. RFA :" & lrDat!cPersCod & " - " & Format(J / lnTotal * 100, "#,#0.00") & "%"
''            Me.Barraprogreso.value = Int(J / lnTotal * 100)
''            DoEvents
''            lrDat.MoveNext
''        Loop
''    End If
''    Set lrDat = Nothing
''
''Set loConec = Nothing
''
''MsgBox "Calificacion RFA Nueva Completada", vbInformation, "Aviso"
''Me.barraEstado.Panels(1).Text = ""
''Me.Barraprogreso.value = 0
'
'End Sub
'
''Sub CALIFICARFAANT()
''' Calificacion Creditos RFA
''Dim lsSQL As String
''Dim Rs As New ADODB.Recordset
''Dim lnTotal As Long, J As Long
''Dim loConec As COMConecta.DCOMConecta
''Dim lrDat As ADODB.Recordset
''Dim lrDatCalif As ADODB.Recordset
''Dim lsCalMaxActual As String, lsCalAnterior As String
''Dim lsCalRFA As String
''Set loConec = New COMConecta.DCOMConecta
''
''    loConec.AbreConexion
''
''    lsSQL = " Select cPersCod From ColocCalifProv " ''          & " Where nprdestado in (2060,2061,2062,22261,2265) "
''
''
''    Set lrDat = loConec.CargaRecordSet(lsSQL)
''    If Not (lrDat.BOF And lrDat.EOF) Then
''        lnTotal = lrDat.RecordCount
''        J = 0
''        Do While Not lrDat.EOF
''            J = J + 1
''            '***************** Obtiene datos -
''            ' Mayor Calificacion de los creditos relacionados
''            lsSQL = " Select Max(cal.cCalNor) CalMaxActual From " & fsServerConsol & "CreditoConsol P " ''                  & "        Inner Join " & fsServerConsol & "ProductoPersonaConsol PP on P.cCtaCod=PP.cCtaCod " ''                  & "        Inner Join colocaccred CC on PP.cCtaCod=CC.cCtaCod " ''                  & "        Inner Join colocCalifProv  Cal on Cal.cCtaCod=CC.cCtaCod " ''                  & "        Where PP.cPersCod='" & lrDat!cPersCod & "' and PP.nPrdPersRelac=20 and " ''                  & "        CC.cRFA IN('RFC','RFA','DIF') and P.nPrdEstado<>2050 "
''
''            Set lrDatCalif = loConec.CargaRecordSet(lsSQL)
''            If Not (lrDatCalif.BOF And lrDatCalif.EOF) Then
''                lsCalMaxActual = lrDatCalif!CalMaxActual
''            End If
''            lrDatCalif.Close
''            ' Calificacion del Mes Anterior
''            lsSQL = " Select cCalGen CalAnterior from " & fsServerConsol & "ColocCalifProvTotal " ''                 & " Where cPersCod = '" & lrDat!cPersCod & "' " ''                 & " And datediff( d,dFecha,'" & Format(DateAdd("d", -1 * Day(Format(txtFecha.Text, "yyyy/mm/dd")), Format(txtFecha.Text, "yyyy/mm/dd")), "yyyy/mm/dd") & "') = 0 "
''            Set lrDatCalif = loConec.CargaRecordSet(lsSQL)
''            If Not (lrDatCalif.BOF And lrDatCalif.EOF) Then
''                lsCalAnterior = lrDatCalif!CalAnterior
''            End If
''            lrDatCalif.Close
''            '******************
''
''
''            If lsCalAnterior = "0" And lsCalMaxActual = "2" Then
''                lsCalRFA = lsCalAnterior
''            ElseIf lsCalMaxActual > lsCalAnterior Then
''                lsCalRFA = lsCalMaxActual
''            Else
''                lsCalRFA = lsCalAnterior
''            End If
''
''            lsSQL = "UPDATE ColocCalifProv " ''                & " SET cCalRFA = '" & Trim(lsCalRFA) & "'" ''                & "WHERE cPersCod ='" & Trim(lrDat!cPersCod) & "'"
''            loConec.Ejecutar lsSQL
''
''            Me.barraEstado.Panels(1).Text = "Cal. RFA :" & lrDat!cPersCod & " - " & Format(J / lnTotal * 100, "#,#0.00") & "%"
''            Me.Barraprogreso.value = Int(J / lnTotal * 100)
''            DoEvents
''            lrDat.MoveNext
''        Loop
''    End If
''    Set lrDat = Nothing
''
''
''Set loConec = Nothing
''
''MsgBox "Calificacion RFA Completada", vbInformation, "Aviso"
''Me.barraEstado.Panels(1).Text = ""
''Me.Barraprogreso.value = 0
''
''End Sub
'
'Private Sub cmdCalRFA_Click()
'    CalificaRFANEW
'End Sub
'
'Private Sub cmdCalRiesgoUnico_Click()
'' Calificacion mayor por Riesgo Unico
'Dim lsSQL As String
'Dim rs As New ADODB.Recordset
'Dim lnTotal As Long, J As Long
'Dim loConec As COMConecta.DCOMConecta
'Dim lrDat As ADODB.Recordset
'Dim lsCadConexion As String
'
'
''** Set loConec = New COMConecta.DCOMConecta
'
''**    loConec.AbreConexion 'lsCadConexion
'
''    lsSQL = " SELECT  CA1.cPersCod " ''        & " FROM ColocCalifProv CA1 " ''        & " Where  CA1.cPersCod IN ( SELECT  CA2.CPersCod " ''        & "                          FROM ColocCalifProv CA2 " ''        & "                          WHERE CA2.cPersCod =CA1.cPersCod " ''        & "                          AND CA1.cCtaCod <> CA2.cCtaCod )" ''        & " GROUP BY CA1.cPersCod "
''
''    Set lrDat = loConec.CargaRecordSet(lsSQL)
''    If Not (lrDat.BOF And lrDat.EOF) Then
''        lnTotal = lrDat.RecordCount
''        J = 0
''        Do While Not lrDat.EOF
''            J = J + 1
''
''            'lsSQL = "UPDATE ColocCalifProv " ''                & " SET cCalRUnico = ( Select MAX(CA1.cCalNor)  " ''                & "                  From ColocCalifProv CA1 Join " & fsServerConsol & "ProductoPersonaConsol PP1 " ''                & "                  ON CA1.cCtaCod = PP1.cCtaCod " ''                & "                  WHERE CA1.cPersCod= PP1.cPersCod " ''                & "                  And PP1.nPrdPersRelac = 20 ) " ''                & "WHERE cCtaCod ='" & Trim(lrDat!cCtaCod) & "'"
''            '12/01/2005- layg
''            lsSQL = "UPDATE ColocCalifProv " ''                & " SET cCalRUnico = ( Select MAX(CA1.cCalNor)  " ''                & "                  From ColocCalifProv CA1 Join ProductoPersona PP1 " ''                & "                  ON CA1.cPersCod = PP1.cPersCod " ''                & "                  WHERE PP1.nPrdPersRelac in(20,22,25) " ''                & "                  And CA1.cPersCod='" & Trim(lrDat!cPersCod) & "'  ) " ''                & "WHERE cPersCod ='" & Trim(lrDat!cPersCod) & "'"
''
''            loConec.Ejecutar lsSQL
''
''            Me.barraEstado.Panels(1).Text = "Cal. R. Unico :" & lrDat!cPersCod & " - " & Format(J / lnTotal * 100, "#,#0.00") & "%"
''            Me.Barraprogreso.value = Int(J / lnTotal * 100)
''            DoEvents
''            lrDat.MoveNext
''        Loop
''
''    End If
''    Set lrDat = Nothing
'
''** Set loConec = Nothing
'
'MsgBox "Calificacion Riesgo Unico Completada", vbInformation, "Aviso"
''Me.barraestado.Panels(1).Text = ""
''Me.barraProgreso.value = 0
'End Sub
'
'Private Sub cmdCalSistF_Click()
'Dim oEval As COMNCredito.NCOMColocEval
'Dim sMensaje As String
'If MsgBox(" Utilizar la Calificacion de RCC ? ", vbQuestion + vbYesNo + vbDefaultButton2, "Aviso") = vbYes Then
'    Set oEval = New COMNCredito.NCOMColocEval
'    Call oEval.CalificaSistemaFinanciero(fsBDRCC, CDate(lblFecAlin.Caption), sMensaje)
'    Set oEval = Nothing
'    If sMensaje <> "" Then
'        MsgBox sMensaje, vbInformation, "Mensaje"
'        Exit Sub
'    End If
'    MsgBox "Calificacion del Sistema Financiero completada", vbInformation, "Aviso"
'End If
'
'' Obtiene la Calificac del Sistema Financiero a CreditoAudi
''Dim lsSQL As String
''Dim lrDat As ADODB.Recordset
''Dim loConec As COMConecta.DCOMConecta
''Dim lnTotal As Long, j As Long
''
''If MsgBox(" Utilizar la Calificacion de RCC ? ", vbQuestion + vbYesNo + vbDefaultButton2, "Aviso") = vbYes Then
''
''    Set loConec = New COMConecta.DCOMConecta
''    loConec.AbreConexion
''
''    lsSQL = "Update ColocCalifProv Set cCalSistF = null "
''    loConec.Ejecutar lsSQL
''
''    '*********************************************
''    'Personas - Documento Civil
''    lsSQL = " Select CA.cPersCod, Can_Ents, " ''        & " Case when Calif_4 > 20 then 4 " ''        & "      when Calif_3 > 20 then 3 " ''        & "      when Calif_2 > 20 then 2 " ''        & "      when Calif_1 > 20 then 1 " ''        & "      else 0 end  as nCalSF " ''        & " FROM ColocCalifProv CA " ''        & " INNER join Persona P on CA.cPersCod = P.cPersCod " ''        & " INNER join PersId PI on P.cPersCod = PI.cPersCod " ''        & " INNER join " & fsBDRCC & "RCCTotal R on LTRIM(RTRIM(R.cod_Doc_Id)) = LTRIM(RTRIM(PI.cPersIDnro)) and CONVERT(CHAR(10),R.FEC_REP,112)='" & Format(lblFecAlin, "yyyymmdd") & "' " ''        & " WHERE PI.cPersIDTpo = " & gPersIdDNI & " " ''        & " AND ( LTRIM(RTRIM(PI.cPersIDnro)) <> '' or PI.cPersIDnro <> null ) " ''        & " And ( r.Can_Ents > 1     " ''        & "     ) " ''        & " Group by CA.cPersCod, Can_Ents, " ''        & " Case when Calif_4 > 20 then 4 " ''        & "      when Calif_3 > 20 then 3 " ''        & "      when Calif_2 > 20 then 2 "
''        & "      when Calif_1 > 20 then 1 " ''        & "      Else 0 end "
''
''        'Se saco la condicion
''        '& "        or (r.Can_Ents = 1 And  ca.dFecVig >= '" & Format(DateAdd("d", -1 * Day(gdFecData), gdFecData), "yyyy/mm/dd") & "' )  " ''
''    Set lrDat = loConec.CargaRecordSet(lsSQL)
''        If lrDat.BOF And lrDat.EOF Then
''            MsgBox "No existe Data para Calificar Personas Natural", vbInformation, "Aviso"
''        Else
''            lnTotal = lrDat.RecordCount
''            j = 0
''            Do While Not lrDat.EOF
''                j = j + 1
''
''                lsSQL = " UPDATE ColocCalifProv Set cCalSistF = '" & Trim(Str(lrDat!nCalSF)) & "' " ''                    & " WHERE cPersCod ='" & lrDat!cPersCod & "'  "
''                loConec.Ejecutar lsSQL
''
''                Me.barraestado.Panels(1).Text = "Cal. Sist Fin. :" & lrDat!cPersCod & " - " & Format(j / lnTotal * 100, "#,#0.00") & "%"
''                Me.Barraprogreso.value = Int(j / lnTotal * 100)
''                DoEvents
''                lrDat.MoveNext
''            Loop
''        End If
''    Set lrDat = Nothing
''
''    '*********************************************
''    'Personas - Documento Tributario
''    lsSQL = " Select CA.cPersCod, Can_Ents,  " ''        & " Case when Calif_4 > 20 then 4 " ''        & "      when Calif_3 > 20 then 3 " ''        & "      when Calif_2 > 20 then 2 " ''        & "      when Calif_1 > 20 then 1 " ''        & "      else 0 end  as nCalSF " ''        & " FROM ColocCalifProv CA " ''        & " INNER join Persona P on CA.cPersCod = P.cPersCod " ''        & " INNER join PersId PI on P.cPersCod = PI.cPersCod " ''        & " INNER join " & fsBDRCC & "RCCTotal R on LTRIM(RTRIM(R.Cod_Doc_Trib)) = LTRIM(RTRIM(PI.cPersIDnro)) " ''        & " WHERE PI.cPersIDTpo = " & gPersIdRUC & " " ''        & " AND ( LTRIM(RTRIM(PI.cPersIDnro)) <> '' or PI.cPersIDnro <> null ) " ''        & " And ( r.Can_Ents > 1     " ''        & "     ) " ''        & " Group by CA.cPersCod, Can_Ents,  " ''        & " Case when Calif_4 > 20 then 4 " ''        & "      when Calif_3 > 20 then 3 " ''        & "      when Calif_2 > 20 then 2 " ''        & "      when Calif_1 > 20 then 1 " ''        & "
''Else 0 end ""
''
''        'Se saco condicion
''        '& "        or (r.Can_Ents = 1 And  ca.dFecVig >= '" & Format(DateAdd("d", -1 * Day(gdFecData), gdFecData), "yyyy/mm/dd") & "' ) " ''
''    Set lrDat = loConec.CargaRecordSet(lsSQL)
''        If lrDat.BOF And lrDat.EOF Then
''            MsgBox "No existe Data para Calificar Persona Juridica ", vbInformation, "Aviso"
''        Else
''            lnTotal = lrDat.RecordCount
''            j = 0
''            Do While Not lrDat.EOF
''                j = j + 1
''
''                lsSQL = " UPDATE ColocCalifProv Set cCalSistF = '" & Trim(Str(lrDat!nCalSF)) & "' " ''                    & " WHERE cPersCod ='" & lrDat!cPersCod & "'  "
''
''                loConec.Ejecutar lsSQL
''
''                Me.barraestado.Panels(1).Text = "Cal. Sist Fin. :" & lrDat!cPersCod & " - " & Format(j / lnTotal * 100, "#,#0.00") & "%"
''                Me.Barraprogreso.value = Int(j / lnTotal * 100)
''                DoEvents
''                lrDat.MoveNext
''            Loop
''        End If
''    Set lrDat = Nothing
''    Set loConec = Nothing
''End If
''
''MsgBox "Calificacion del Sistema Financiero completada", vbInformation, "Aviso"
''Me.barraestado.Panels(1).Text = ""
''Me.Barraprogreso.value = 0
'End Sub
'
'Private Sub cmdEndeudamientoSF_Click()
'
'Dim oEval As COMNCredito.NCOMColocEval
'Set oEval = New COMNCredito.NCOMColocEval
'Call oEval.EndeudamientoSistFinanc(CDbl(txtTipoCambio.Text), fsBDRCC)
'Set oEval = Nothing
'MsgBox "Actualizacion Generada Correctamente , Termino : " & Time(), vbInformation, "Aviso"
''' Reclasifica los Creditos MES
''Dim lsSQL As String
''Dim rs As ADODB.Recordset
''Dim loConec As COMConecta.DCOMConecta
''Dim lnTotal As Long, j As Long
''
''lsSQL = " Select CA.cPersCod, " ''    & " Isnull( Sum(case when substring(cod_cuenta,3,1) = '2' And rd.cod_cuenta like '14_[1456]%' then Val_saldo / " & CDbl(Me.txtTipoCambio) ''    & "        End),0) nEndeudaDol " ''    & " FROM (select C.cPersCod From ColocCalifProv c where c.nPrdEstado in (2020,2021,2022,2030,2031,2032) group by c.cPersCod) ca " ''    & " INNER join Persona p on CA.cPersCod = P.cPersCod " ''    & " INNER join PersId PI on P.cPersCod = PI.cPersCod " ''    & " INNER join " & fsBDRCC & "rccTotal r on LTRIM(RTRIM(r.Cod_Doc_Id)) = LTRIM(RTRIM(PI.cPersIDnro)) " ''    & " INNER join " & fsBDRCC & "rccTotalDet rd on LTRIM(RTRIM(r.Cod_Sbs)) = LTRIM(RTRIM(rd.Cod_Sbs)) " ''    & " WHERE  PI.cPersIDTpo = " & gPersIdDNI & " " ''    & " AND( LTRIM(RTRIM(pi.cPersIDnro)) <>'' or pi.cPersIDnro <>null ) " ''    & " And rd.cod_cuenta like '14_[1456]%' " ''    & " Group by CA.cPersCod "
''
'' Set loConec = New COMConecta.DCOMConecta
'' loConec.AbreConexion
'' Set rs = loConec.CargaRecordSet(lsSQL)
''
''lnTotal = rs.RecordCount
''j = 0
''Do While Not rs.EOF
''    j = j + 1
''
''    lsSQL = "UPDATE ColocCalifProv SET nMESReclas =  " & Format(rs!nEndeudaDol, "#0.00") ''        & "WHERE cPersCod='" & Trim(rs!cPersCod) & "'"
''
''    loConec.Ejecutar lsSQL
''
''    Me.barraEstado.Panels(1).Text = "Endeud.  :" & rs!cPersCod & " - " & Format(j / lnTotal * 100, "#,#0.00") & "%"
''    Me.Barraprogreso.value = Int(j / lnTotal * 100)
''    DoEvents
''    rs.MoveNext
''Loop
''rs.Close
''
'''*** Personas Juridicas
''lsSQL = " Select CA.cPersCod, " ''    & " Isnull( Sum(case when substring(cod_cuenta,3,1) = '2' And rd.cod_cuenta like '14_[1456]%' then Val_saldo / " & CDbl(Me.txtTipoCambio) ''    & "        End),0) nEndeudaDol " ''    & " FROM (select C.cPersCod From ColocCalifProv c where c.nPrdEstado in (2020,2021,2022,2030,2031,2032) group by c.cPersCod) ca " ''    & " INNER join Persona p on CA.cPersCod = P.cPersCod " ''    & " INNER join PersId PI on P.cPersCod = PI.cPersCod " ''    & " INNER join " & fsBDRCC & "rccTotal r on LTRIM(RTRIM(r.Cod_Doc_Trib)) = LTRIM(RTRIM(PI.cPersIDnro)) " ''    & " INNER join " & fsBDRCC & "rccTotalDet rd on LTRIM(RTRIM(r.Cod_Sbs)) = LTRIM(RTRIM(rd.Cod_Sbs)) " ''    & " WHERE PI.cPersIDTpo = " & gPersIdRUC & " " ''    & " AND ( LTRIM(RTRIM(pi.cPersIDnro)) <>'' or pi.cPersIDnro <>null ) " ''    & " And rd.cod_cuenta like '14_[1456]%' " ''    & " Group by CA.cPersCod "
''
''Set rs = loConec.CargaRecordSet(lsSQL)
''
''lnTotal = rs.RecordCount
''j = 0
''Do While Not rs.EOF
''    j = j + 1
''
''    lsSQL = "UPDATE ColocCalifProv SET nMESReclas =  " & Format(rs!nEndeudaDol, "#0.00") ''        & "WHERE cPersCod='" & Trim(rs!cPersCod) & "'"
''
''    loConec.Ejecutar (lsSQL)
''
''    Me.barraEstado.Panels(1).Text = "Cal. CMACT :" & rs!cPersCod & " - " & Format(j / lnTotal * 100, "#,#0.00") & "%"
''    Me.Barraprogreso.value = Int(j / lnTotal * 100)
''    DoEvents
''    rs.MoveNext
''Loop
''rs.Close
''Set rs = Nothing
''Set loConec = Nothing
''
''MsgBox "Actualizacion Generada Correctamente , Termino : " & Time(), vbInformation, "Aviso"
''Me.barraEstado.Panels(1).Text = ""
''Me.Barraprogreso.value = 0
'
'End Sub
'
'Private Sub cmdLLenaCreditoAudi_Click()
'
''** Llena tabla ColocCalifProv (Contiene las Calificaciones de los creditos Vigentes)
'Dim oEval As COMNCredito.NCOMColocEval
'Dim sMensaje As String
'
'On Error GoTo ErrorConexion
'
'If VerificaDatosIngresados = False Then Exit Sub
'
'If MsgBox("Se Reprocesaran los Datos para la Calificacion, Desea Continuar ?", vbQuestion + vbYesNo, "Aviso") = vbNo Then
'    Exit Sub
'End If
'
'Set oEval = New COMNCredito.NCOMColocEval
'Call oEval.PreparaArchivoCalificacion(fsServerConsol, txtFecha.Text, fnTipoCambio, sMensaje)
'Set oEval = Nothing
'
'If sMensaje <> "" Then
'    MsgBox sMensaje, vbInformation, "Mensaje"
'    Exit Sub
'End If
'
'MsgBox "Preparacion de Archivo para Calificacion completado ", vbInformation, "Aviso"
'
'Exit Sub
'
'ErrorConexion:
'    MsgBox "Error Nº[" & Str(Err.Number) & "] " & Err.Description, vbInformation, "Aviso"
'
'''** Llena tabla ColocCalifProv (Contiene las Calificaciones de los creditos Vigentes)
''Dim loConec As COMConecta.DCOMConecta
''Dim loCalif As COMNCredito.NCOMColocEval
''Dim lrDatos As ADODB.Recordset
''Dim lsSQL As String
''
''Dim lsCalificacion As String
''Dim lsGarPrefProc As String
''
''Dim Total As Long
''Dim M As Long
''Dim lbGarPrefe As Boolean
'''On Error GoTo ErrorConexion
''
''If VerificaDatosIngresados = False Then Exit Sub
''
''If MsgBox("Se Reprocesaran los Datos para la Calificacion, Desea Continuar ?", vbYesNo, "Aviso") = vbNo Then
''    Exit Sub
''End If
''
''Set loCalif = New COMNCredito.NCOMColocEval
''    Set lrDatos = loCalif.nLLenaTablaCalifProv(fsServerConsol, txtFecha.Text, fnTipoCambio)
''Set loCalif = Nothing
''
''M = 0
''If lrDatos Is Nothing Then
''    MsgBox "No existen datos", vbInformation, "Aviso"
''    Exit Sub
''Else
''    Set loConec = New COMConecta.DCOMConecta
''        loConec.AbreConexion
''    Total = lrDatos.RecordCount
''    'Borra los datos de ColocCalifProv
''    lsSQL = "Delete ColocCalifProv "
''    loConec.Ejecutar lsSQL
''
''    Do While Not lrDatos.EOF
''        M = M + 1
''
''        'Aplica Reglas de Negocio
''        Set loCalif = New COMNCredito.NCOMColocEval
''            lsCalificacion = loCalif.nCorrigeCalifxDiaAtraso(lrDatos!cCalDias, lrDatos!nPrdEstado)
''        Set loCalif = Nothing
''
''        lsSQL = "INSERT INTO ColocCalifProv (cPersCod,cCtaCod, nPrdEstado, cCalNor, nSaldoCap, cRefinan, nDiasAtraso,nProvision," ''          & " nMontoApr,cLineaCred, cCodAnalista, dFecVig, nGarPref,nGarMuyRR, nGarAutoL) " ''          & "VALUES('" & lrDatos!cPersCod & "','" & lrDatos!cCtaCod & "'," & lrDatos!nPrdEstado & ",'" ''          & lsCalificacion & "'," & lrDatos!nSaldoCap & ",'" & lrDatos!cRefinan & "'," & lrDatos!nDiasAtraso & ",0," ''          & lrDatos!nMontoApr & ",'" & lrDatos!cLineaCred & "','" & lrDatos!cCodAnalista & "','" ''          & Format(lrDatos!dFecVig, "mm/dd/yyyy") & "'," & Format(lrDatos!nGarPref, "#.00") & "," ''          & Format(lrDatos!nGarMuyRR, "#.00") & "," & Format(lrDatos!nGarAutoL, "#.00") & " )"
''
''        loConec.Ejecutar lsSQL
''
''        Me.Barraprogreso.value = Int(M / Total * 100)
''        Me.barraEstado.Panels(1).Text = "Credito :" & lrDatos!cCtaCod & " - " & Format(M / Total * 100, "#0.00") & "%"
''
''        lrDatos.MoveNext
''        DoEvents
''    Loop
''    Set loConec = Nothing
''End If
''lrDatos.Close
''Set lrDatos = Nothing
''Set loCalif = Nothing
''
''MsgBox "Preparacion de Archivo para Calificacion completado ", vbInformation, "Aviso"
''
''Exit Sub
''
''ErrorConexion:
''    MsgBox "Error Nº[" & Str(Err.Number) & "] " & Err.Description, vbInformation, "Aviso"
''    Set lrDatos = Nothing
'End Sub
'
'
'Private Sub cmdSalir_Click()
'Dim oEval As COMNCredito.NCOMColocEval
'Set oEval = New COMNCredito.NCOMColocEval
'Call oEval.VerificaTablaTemporal(gsCodUser, lsTablaTMP)
'Set oEval = Nothing
'Unload Me
'End Sub
'
'Private Sub Form_Load()
'Me.Icon = LoadPicture(App.path & gsRutaIcono)
'Dim loConstS As COMDConstSistema.NCOMConstSistema
'Dim loTipCambio As COMDConstSistema.NCOMTipoCambio
'Dim oEval As COMNCredito.NCOMColocEval
'
'    Set loConstS = New COMDConstSistema.NCOMConstSistema
'        fdFechaFinMes = CDate(loConstS.LeeConstSistema(gConstSistCierreMesNegocio))
'        txtFecha.Text = fdFechaFinMes
'        fsServerConsol = loConstS.LeeConstSistema(gConstSistServCentralRiesgos)
'        fsServerRCC = loConstS.LeeConstSistema(143)
'        fsBDRCC = loConstS.LeeConstSistema(144)
'    Set loConstS = Nothing
'
'    Set loTipCambio = New COMDConstSistema.NCOMTipoCambio
'        fnTipoCambio = Format(loTipCambio.EmiteTipoCambio(gdFecSis, TCFijoMes), "0.###")
'        txtTipoCambio.Text = fnTipoCambio
'    Set loTipCambio = Nothing
'
'    Set oEval = New COMNCredito.NCOMColocEval
'    Me.lblFecAlin = oEval.GetFechaAlin(fsBDRCC)
'    Set oEval = Nothing
'End Sub
'
'Private Function Valida() As Boolean
'Dim i As Integer
'Valida = True
'
'If ValFecha(txtFecha) = False Then
'    Valida = False
'    Exit Function
'End If
'End Function
'
'
'Private Function VerificaDatosIngresados() As Boolean
'Dim lbOk As Boolean
'lbOk = True
'If Not IsDate(Me.txtFecha.Text) Then
'    lbOk = False
'End If
'
'If Not IsNumeric(Me.txtTipoCambio.Text) Then
'    lbOk = False
'Else
'    fnTipoCambio = Format(Me.txtTipoCambio.Text, "#,#0.000")
'End If
'VerificaDatosIngresados = lbOk
'End Function
'
''Private Function GetFechaAlin() As String
''Dim Sql As String
''Dim rs As ADODB.Recordset
''Dim oCon As COMConecta.DCOMConecta
''Set oCon = New COMConecta.DCOMConecta
''
''GetFechaAlin = ""
''oCon.AbreConexion
''Sql = "SELECT MAX(FEC_REP) as dFecAlin FROM " & fsBDRCC & "RCCTOTAL"
''Set rs = oCon.CargaRecordSet(Sql)
''If Not rs.EOF And Not rs.BOF Then
''    GetFechaAlin = Format(rs!dFecAlin, "dd/mm/yyyy")
''End If
''rs.Close
''Set rs = Nothing
''oCon.CierraConexion
''Set oCon = Nothing
''
''End Function
'
''Sub VerificaTablaTemporal()
''Dim Rs As ADODB.Recordset
''Dim Sql As String
''Dim oCon As COMConecta.DCOMConecta
''
''lsTablaTMP = "TMPPAGORFA" & gsCodUser
''
''Set oCon = New COMConecta.DCOMConecta
''oCon.AbreConexion
''
''Set Rs = oCon.CargaRecordSet("select * from sysobjects where name like '%" & lsTablaTMP & "%'")
''If Not Rs.EOF And Not Rs.BOF Then
''    Sql = "DROP TABLE " & lsTablaTMP
''    oCon.Ejecutar Sql
''End If
''Rs.Close
''Set Rs = Nothing
''
''End Sub
'
''Sub VerificaTablaTemporal()
''Dim rs As ADODB.Recordset
''Dim SQL As String
''Dim oCon As COMConecta.DCOMConecta
''
''lsTablaTMP = "TMPPAGORFA" & gsCodUser
''
''Set oCon = New COMConecta.DCOMConecta
''oCon.AbreConexion
''
''Set rs = oCon.CargaRecordSet("select * from sysobjects where name like '%" & lsTablaTMP & "%'")
''If Not rs.EOF And Not rs.BOF Then
''    SQL = "DROP TABLE " & lsTablaTMP
''    oCon.Ejecutar SQL
''End If
''rs.Close
''Set rs = Nothing
''
''End Sub
