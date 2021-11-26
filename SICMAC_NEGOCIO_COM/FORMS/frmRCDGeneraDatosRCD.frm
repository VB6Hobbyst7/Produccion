VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRCDGeneraDatosRCD 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informe RCD - Genera Datos para Informe RCD"
   ClientHeight    =   2055
   ClientLeft      =   3690
   ClientTop       =   4455
   ClientWidth     =   6240
   Icon            =   "frmRCDGeneraDatosRCD.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
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
      Left            =   4860
      TabIndex        =   4
      Top             =   1560
      Width           =   1140
   End
   Begin MSComctlLib.ProgressBar barra 
      Height          =   225
      Left            =   90
      TabIndex        =   3
      Top             =   780
      Width           =   5985
      _ExtentX        =   10557
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdConsolida 
      Caption         =   "&Consolida Data"
      CausesValidation=   0   'False
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
      Left            =   2280
      TabIndex        =   2
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Label lblDescripcion 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   135
      TabIndex        =   6
      Top             =   1080
      Width           =   2640
   End
   Begin VB.Label lblAvance 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   4140
      TabIndex        =   5
      Top             =   1080
      Width           =   1650
   End
   Begin VB.Label lblfecha 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
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
      Height          =   330
      Left            =   2160
      TabIndex        =   1
      Top             =   240
      Width           =   1380
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fecha  Consolidacion :"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1665
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmRCDGeneraDatosRCD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fnMontoMinimoRCD As Double
Dim fnTipCambio As Currency

Dim fsServConsol As String
Dim fsCodOfiInf As String
Dim fsUbicaGeoRCD As String

Dim vfcCodUnico As String * 10  ' Codigo asignado por la empresa
Dim vfcCodSBS As String * 10
Dim vfcNomPers As String * 120
Dim vfcNomPersComp As String * 120  ' Agregado 11/06/04
Dim vfcGenero As String * 1     ' Agregado 11/06/04
Dim vfcEstado As String * 1     ' Agregado 11/06/04
Dim vfcActEcon As String * 4
Dim vfcCodRegPub As String * 15
Dim vfcTiDoTr As String * 1
Dim vfcNuDoTr As String * 11
Dim vfcTiDoCi As String * 1
Dim vfcNuDoCi As String * 12
Dim vfcTipPers As String * 1
Dim vfcResid As String * 1
Dim vfcMagEmp As String * 1
Dim vfcAccionista As String * 1
Dim vfcRelInst As String * 1
Dim vfcPaisNac As String * 4
Dim vfcSiglas As String * 20
Dim vfcCalificacion As String * 1
Dim lsCtaPref As String
Dim lsCtaNoPref As String
Dim rsGarant As ADODB.Recordset

Dim lsPat  As String
Dim lsMat  As String
Dim lsCas  As String
Dim lsNom1 As String
Dim lsNom2 As String

Private Sub cmdConsolida_Click()
Dim loRCDproc As COMNCredito.NCOMRCD

If lsCtaNoPref = "" Then
    MsgBox "Plantilla contable de Garantias no se ha definido", vbInformation, "aviso"
    Exit Sub
End If
Me.Enabled = False
' Crea Tablas RCD del mes
Set loRCDproc = New COMNCredito.NCOMRCD
    Call loRCDproc.nCreaTablasRCD(Format(gdFecDataFM, "yyyymm"), fsServConsol)

' Llena Datos en tablas RCD
Call GeneraDatosRCD(Format(gdFecDataFM, "yyyymm"), loRCDproc)

Set loRCDproc = Nothing

Me.lblAvance = ""
Me.LblDescripcion = ""

Me.Enabled = True

MsgBox "El Proceso de Consolidacion de Cliente ha culminado", vbInformation, "Aviso"

End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Dim loRCDproc As COMNCredito.NCOMRCD
Dim lrPar As ADODB.Recordset
Dim loConstSist As COMDConstSistema.NCOMConstSistema
Dim oRCD As COMDCredito.DCOMRCD

Set loConstSist = New COMDConstSistema.NCOMConstSistema
    fsServConsol = loConstSist.LeeConstSistema(gConstSistServCentralRiesgos)
    fsCodOfiInf = loConstSist.LeeConstSistema(81)
    fsUbicaGeoRCD = loConstSist.LeeConstSistema(83)
Set loConstSist = Nothing

' Parametros de RCD
Set loRCDproc = New COMNCredito.NCOMRCD
    Set lrPar = loRCDproc.nCargaParametroRCD(Format(gdFecDataFM, "yyyymmdd"), fsServConsol)
Set loRCDproc = Nothing

    If lrPar.EOF And lrPar.BOF Then
        MsgBox "No se han ingresado Parametros de RCD", vbInformation, "Aviso"
        cmdConsolida.Enabled = False
    Else
        fnMontoMinimoRCD = lrPar!nMontoMin
        fnTipCambio = lrPar!nCambioFijo
    End If
Set lrPar = Nothing
Me.lblFecha.Caption = gdFecDataFM
Me.Icon = LoadPicture(App.path & gsRutaIcono)

Set oRCD = New COMDCredito.DCOMRCD
    lsCtaNoPref = oRCD.CargarCtaNoPref(fsServConsol)
Set oRCD = Nothing

End Sub

'**************************************************
'*** Genera los Datos para el RCD *****************

Private Sub GeneraDatosRCD(ByVal lsFecha As String, _
                            ByVal oRCD As COMNCredito.NCOMRCD)
Dim sMensaje() As String
Dim sMensaje1() As String
Dim I As Integer
                            
Call oRCD.GeneraDatosRCD(lsFecha, fsServConsol, gdFecDataFM, fnTipCambio, fsCodOfiInf, fsUbicaGeoRCD, fnMontoMinimoRCD, lsCtaNoPref, sMensaje, sMensaje1)

For I = 0 To UBound(sMensaje) - 1
    MsgBox sMensaje(I), vbInformation, "Mensaje"
Next

For I = 0 To UBound(sMensaje1) - 1
    MsgBox sMensaje1(I), vbInformation, "Mensaje"
Next

End Sub

'**************************************************
'*** Genera los Datos para el RCD *****************

'Private Sub GeneraDatosRCD(ByVal lsFecha As String)
'
'Dim lsSQL As String
'Dim rs As ADODB.Recordset
''Dim PObjConec as COMConecta.DCOMConecta
'Dim rsPers As ADODB.Recordset
'Dim rsRCD As ADODB.Recordset
'
'Dim lbEncuentroEnRCD  As Boolean
'Dim loRCDProceso As COMNCredito.NCOMRCD
'
'Dim lsRelacionInst  As String
'Dim rsCred As New ADODB.Recordset
'Dim lnTotal As Long
'Dim i As Long
'Dim J As Long
'Dim Agencia As String
'Dim lscadena As String
'Dim lsNombrePersona As String
'Dim lsNombrePersonaComp As String ' linea agregada
'Dim lsTipPers As String
'Dim lsTipoDocRCD As String * 1
'Dim lsTipoInfRCD As String * 1
'Dim lsCalifTabla As String
'Dim lsCodPers As String
'Dim lscTidoci As String
'Dim lscTidoTr As String
'Dim lscNudoCi As String
'Dim lscNudoRr As String
'Dim lsIndRCC  As String
'
'Dim lsNuDotr As String
'Dim lsTidotr As String
'Dim lsCredito As String
'Dim lnSaldoCap As Currency
'Dim lnCapVenc As Currency
'Dim lnDiasAtraso As Integer
'Dim lbCobJud As Boolean
'Dim lbDemanda As Boolean
'Dim lbCastigado As Boolean
'Dim lbRefinan As Boolean
'Dim lnIntDeveng As Currency
'
'Dim lnProvision As Currency
'
'Dim lnMontoGar1 As Currency  ' Garant Hipotecaria
'Dim lnMontoGar2 As Currency  ' Otras Garantias
'
'Dim lsEstadoCredito As String
'Dim lsCuentaCnt As String
'Dim PObjConec As COMConecta.DCOMConecta
'Set PObjConec = New COMConecta.DCOMConecta
'    PObjConec.AbreConexion
'
''** Carga los Creditos (Vigentes y Castigados)
'lsSQL = " SELECT C.cCtaCod, C.nPrdEstado, C.nDiasAtraso, C.cNumFuente, C.nMontoApr, " _
'    & "         C.nSaldoCap, ISNULL(C.nCapVencido,0) nCapVenc, ISNULL(C.nHipoSoles,0) nHipoSoles , " _
'    & "         ISNULL(C.nHipoDol,0) nHipoDol, " _
'    & "         ISNULL(C.nGarantSoles,0) nGarantSoles, ISNULL(C.nGarantDol,0) nGarantDol , " _
'    & "         ISNULL(C.nIntDev,0) nIntDeveng , ISNULL(C.nIntSusp, 0) nIntSusp, C.nDemanda, PP.cPersCod,   " _
'    & "         RMP.cCodUnico, RMP.cPersNom AS cPersNombre , RMP.cTidoci, " _
'    & "         RMP.cNudoci, RMP.cTidotr, RMP.cNudotr,  RMP.cTipPers, " _
'    & "         RMP.cCodSBS,  RMP.cActEcon, RMP.cCodRegPub, RMP.cResid, RMP.cMagEmp, RMP.cAccionista, RMP.cRelInst, RMP.cPaisNac, RMP.cSiglas," _
'    & "         nProvision = (Select nProvision From ColocCalifProv ca Where ca.cCtaCod = C.cCtaCod ),  " _
'    & "         cCalif =     (Select isnull(max(ca1.cCalGen),'') FROM ColocCalifProv ca1 WHERE ca1.cPersCod = PP.cPersCod ),  " _
'    & "         cActEcon1 =  (  SELECT  TOP 1 Right(Fte.cActEcon, 4)  " _
'    & "                         FROM " & fsServConsol & "FuenteIngresoConsol Fte " _
'    & "                         WHERE Fte.cPersCod = PP.cPersCod )" _
'    & "FROM " & fsServConsol & "CreditoConsol C " _
'    & "     INNER JOIN " & fsServConsol & "ProductoPersonaConsol PP ON PP.cCtaCod = C.cCtaCod " _
'    & "     LEFT JOIN " & fsServConsol & "RCDMaestroPersona RMP ON RMP.cPersCod = PP.cPersCod " _
'    & "WHERE C.nPrdEstado in( " & gColocEstVigNorm & "," & gColocEstVigVenc & "," & gColocEstVigMor _
'    & "," & gColocEstRefNorm & "," & gColocEstRefVenc & "," & gColocEstRefMor _
'    & "," & gColPEstDesem & "," & gColPEstVenci & "," & gColPEstPRema & "," & gColPEstRenov _
'    & "," & gColocEstRecVigJud & "," & gColocEstRecVigCast & ",2205,2206 ) " _
'    & "AND PP.nPrdPersRelac = 20 AND C.nSaldoCap > " & Trim(fnMontoMinimoRCD) _
'    & "AND PP.cPersCod = (Select max(p1.cPersCod) From " & fsServConsol & "ProductoPersonaConsol p1 " _
'    & "                where p1.cCtacod = c.cCtaCod and p1.nPrdPersRelac = 20  ) "
'
'i = 0
'
'lsTipoDocRCD = "1"
'lsTipoInfRCD = "1"
'
' Set rs = PObjConec.CargaRecordSet(lsSQL)
'
' lnTotal = rs.RecordCount
' i = 0
'
' 'cargamos total de garantias ejrs ICA
' Call CargaGarantiasTotal
'
' 'cargamos en un solo recordset el total de datos de maestro de personas
' lsSQL = "Select P.cPersCod, P.cPersNombre, P.dPersNacCreac, ISNULL(P.nPersPersoneria,'') As cTipPers, " _
'        & "IsNULL(P.cPersCIIU,'') AS cActEcon, ISNULL(P.nPersRelaInst,'') As cRelInst,  " _
'        & " NroDNI = (Select ISNULL(cPersIDnro,'') From PersID Where cPersCod = P.cPersCod and cPersIDTpo =  " & gPersIdDNI & " ), " _
'        & " NroRUC = (Select ISNULL(cPersIDnro,'') From PersID Where cPersCod = P.cPersCod and cPersIDTpo =  " & gPersIdRUC & " ), " _
'        & " ISNULL(PJ.cPersJurMagnitud,'') As cMagEmp,  ISNULL(PJ.cPersJurSigla,'') As cSiglas,  " _
'        & " IsNULL(PN.cPersNatSexo,'0') AS Genero,  IsNULL(PN.nPersNatEstCiv,'0') AS Estado  " _
'        & " From Persona P " _
'        & " LEFT JOIN PersonaJur PJ on P.cPersCod = PJ.cPersCod " _
'        & " LEFT JOIN PersonaNat PN on P.cPersCod = PN.cPersCod "
'Set rsPers = PObjConec.CargaRecordSet(lsSQL)
'rsPers.Sort = "cPersCod"
'
' If Not (rs.BOF And rs.EOF) Then
'    Do While Not rs.EOF
'        i = i + 1
'
'        'If rs!cPersCod = "1080100639991" Or rs!cPersCod = "1080400024104" Or rs!cPersCod = "1080100000761" Then Stop
'        '***********************************************
'        '**  DATOS DE LA PERSONA  -  RCDvcAAAAMM01   ***
'        '***********************************************
'        '***** limpio las variables
'        vfcCodUnico = "":         vfcCodSBS = ""
'        vfcNomPers = "":          vfcActEcon = ""
'        vfcNomPersComp = ""
'        vfcCodRegPub = "":        vfcTiDoTr = ""
'        vfcNuDoTr = "":           vfcTiDoCi = ""
'        vfcNuDoCi = "":           vfcTipPers = ""
'        vfcResid = "":            vfcMagEmp = ""
'        vfcAccionista = "":       vfcRelInst = ""
'        vfcPaisNac = "":          vfcSiglas = ""
'        vfcCalificacion = "0"
'        vfcGenero = ""
'        vfcEstado = ""
'        lsIndRCC = "0"
'        'indicar de riesgo crediticio cambiario modificado el 07 de julio del 2005
'        If Mid(rs!cCtaCod, 6, 3) = "423" Then
'            lsIndRCC = "2"
'        End If
'
'        lsCodPers = Trim(EmiteCodigoAux(Trim(rs!cPersCod), fsServConsol))
'        If Len(Trim(lsCodPers)) = 0 Then
'            lsCodPers = Trim(rs!cPersCod)
'        End If
'
'        'Verifica si se encuentra en el RCD
'        lsSQL = "Select cPersCod From " & fsServConsol & "RCDvc" & Format(gdFecDataFM, "yyyymm") & "01 " _
'              & " Where cPersCod ='" & Trim(lsCodPers) & "'"
'
'        Set rsRCD = PObjConec.CargaRecordSet(lsSQL)
'        lbEncuentroEnRCD = IIf((rsRCD.BOF And rsRCD.EOF) = True, False, True)
'        Set rsRCD = Nothing
'
'        If lbEncuentroEnRCD = False Then ' Lo inserta
'            If Not (IsNull(rs!cCodUnico)) Then ' Existe en Maestro Persona
'                ' Insertado el 11/06/2004
'                'lsSQL = "Select P.cPersCod, P.cPersNombre, P.dPersNacCreac, ISNULL(P.nPersPersoneria,'') As cTipPers, " _
'                        & "IsNULL(P.cPersCIIU,'') AS cActEcon, ISNULL(P.nPersRelaInst,'') As cRelInst,  " _
'                        & " NroDNI = (Select ISNULL(cPersIDnro,'') From PersID Where cPersCod = P.cPersCod and cPersIDTpo =  " & gPersIdDNI & " ), " _
'                        & " NroRUC = (Select ISNULL(cPersIDnro,'') From PersID Where cPersCod = P.cPersCod and cPersIDTpo =  " & gPersIdRUC & " ), " _
'                        & " ISNULL(PJ.cPersJurMagnitud,'') As cMagEmp,  ISNULL(PJ.cPersJurSigla,'') As cSiglas,  " _
'                        & " IsNULL(PN.cPersNatSexo,'0') AS Genero,  IsNULL(PN.nPersNatEstCiv,'0') AS Estado  " _
'                        & " From Persona P " _
'                        & " LEFT JOIN PersonaJur PJ on P.cPersCod = PJ.cPersCod " _
'                        & " LEFT JOIN PersonaNat PN on P.cPersCod = PN.cPersCod " _
'                        & " Where P.cPersCod ='" & Trim(rs!cPersCod) & "' "
'                'Set rsPers = PObjConec.CargaRecordSet(lsSQL)
'                rsPers.MoveFirst
'                rsPers.Find "cPersCod ='" & Trim(rs!cPersCod) & "'"
'
'                vfcNomPersComp = rsPers!cPersNombre
'
'                CambiaNombreRCD rsPers!cPersNombre
'
'                vfcGenero = rsPers!Genero
'                vfcEstado = rsPers!Estado
'                'Set rsPers = Nothing
'
'                ''''fin de insercion del 11/06/2004
'                ' Asigno los valores a las Variables
'                vfcCodUnico = rs!cCodUnico
'                vfcCodSBS = rs!cCodSBS
'                vfcNomPers = rs!cPersNombre
'                If InStr(1, vfcNomPers, "'", vbTextCompare) <> 0 Then
'                    vfcNomPers = Replace(vfcNomPers, "'", "''", , , vbTextCompare)
'                End If
'                vfcNomPers = Replace(vfcNomPers, "Ñ", "#", , , vbTextCompare)
'                vfcNomPers = Replace(vfcNomPers, "ñ", "#", , , vbTextCompare)
'                vfcNomPers = Replace(vfcNomPers, "-", "", , , vbTextCompare)
'                vfcNomPers = Replace(vfcNomPers, "|", "", , , vbTextCompare)
'                vfcNomPers = Replace(vfcNomPers, ".", " ", , , vbTextCompare)
'                'al final
'                vfcNomPers = Replace(vfcNomPers, "   ", " ", , , vbTextCompare)
'                vfcNomPers = Replace(vfcNomPers, "  ", " ", , , vbTextCompare)
'                vfcNomPers = Trim(vfcNomPers)
'                '------------------------
'                vfcNomPersComp = Replace(vfcNomPersComp, "Ñ", "#", , , vbTextCompare)
'                vfcNomPersComp = Replace(vfcNomPersComp, "ñ", "#", , , vbTextCompare)
'                vfcNomPersComp = Replace(vfcNomPersComp, "-", "", , , vbTextCompare)
'                vfcNomPersComp = Replace(vfcNomPersComp, "|", "", , , vbTextCompare)
'                vfcNomPersComp = Replace(vfcNomPersComp, ".", " ", , , vbTextCompare)
'                'al final
'                vfcNomPersComp = Replace(vfcNomPersComp, "   ", " ", , , vbTextCompare)
'                vfcNomPersComp = Replace(vfcNomPersComp, "  ", " ", , , vbTextCompare)
'                vfcNomPersComp = Trim(vfcNomPersComp)
'                '------------------------
'                vfcActEcon = rs!cActEcon
'                vfcCodRegPub = IIf(IsNull(rs!ccodregpub), "", rs!ccodregpub)
'                vfcTiDoTr = IIf(IsNull(rs!cTidoTr), "", rs!cTidoTr)
'                vfcNuDoTr = IIf(IsNull(rs!cNudOtr), "", rs!cNudOtr)
'                vfcTiDoCi = IIf(IsNull(rs!ctidoci), "", rs!ctidoci)
'                vfcNuDoCi = IIf(IsNull(rs!cnudoci), "", rs!cnudoci)
'                vfcTipPers = rs!cTipPers
'                If vfcTipPers = "3" Then
'                   vfcTipPers = "2"
'                End If
'                vfcResid = rs!cResid
'                vfcMagEmp = IIf(IsNull(rs!cMagEmp), "", rs!cMagEmp)
'                vfcAccionista = rs!cAccionista
'                vfcRelInst = rs!cRelInst
'                vfcPaisNac = rs!cPaisNac
'                vfcSiglas = IIf(IsNull(rs!cSiglas), "", rs!cSiglas)
'            Else ' Busco en Personas
'
''                lsSQL = "Select P.cPersCod, P.cPersNombre, P.dPersNacCreac, ISNULL(P.nPersPersoneria,'') As cTipPers, " _
''                        & "IsNULL(P.cPersCIIU,'') AS cActEcon, ISNULL(P.nPersRelaInst,'') As cRelInst,  " _
''                        & " NroDNI = (Select ISNULL(cPersIDnro,'') From PersID Where cPersCod = P.cPersCod and cPersIDTpo =  " & gPersIdDNI & " ), " _
''                        & " NroRUC = (Select ISNULL(cPersIDnro,'') From PersID Where cPersCod = P.cPersCod and cPersIDTpo =  " & gPersIdRUC & " ), " _
''                        & " ISNULL(PJ.cPersJurMagnitud,'') As cMagEmp,  ISNULL(PJ.cPersJurSigla,'') As cSiglas,  " _
''                        & " IsNULL(PN.cPersNatSexo,'0') AS Genero,  IsNULL(PN.nPersNatEstCiv,'0') AS Estado  " _
''                        & " From Persona P " _
''                        & " LEFT JOIN PersonaJur PJ on P.cPersCod = PJ.cPersCod " _
''                        & " LEFT JOIN PersonaNat PN on P.cPersCod = PN.cPersCod " _
''                        & " Where P.cPersCod ='" & Trim(rs!cPersCod) & "' "
'
'                'Set rsPers = PObjConec.CargaRecordSet(lsSQL)
'                    rsPers.MoveFirst
'                    rsPers.Find "cPersCod ='" & Trim(rs!cPersCod) & "'"
'
'                    vfcNomPersComp = rsPers!cPersNombre ' agregado el 11/06/2004
'
'                    CambiaNombreRCD rsPers!cPersNombre
'
'                    vfcGenero = rsPers!Genero
'                    vfcEstado = rsPers!Estado
'                    vfcNomPers = rsPers!cPersNombre
'                    If InStr(1, vfcNomPers, "'", vbTextCompare) <> 0 Then
'                        vfcNomPers = Replace(vfcNomPers, "'", "''", , , vbTextCompare)
'                    End If
'                    vfcNomPers = Replace(vfcNomPers, "Ñ", "#", , , vbTextCompare)
'                    vfcNomPers = Replace(vfcNomPers, "ñ", "#", , , vbTextCompare)
'                    vfcNomPers = Replace(vfcNomPers, "-", "", , , vbTextCompare)
'                    vfcNomPers = Replace(vfcNomPers, "|", "", , , vbTextCompare)
'                    vfcNomPers = Replace(vfcNomPers, ".", " ", , , vbTextCompare)
'                    'al final
'                    vfcNomPers = Replace(vfcNomPers, "   ", " ", , , vbTextCompare)
'                    vfcNomPers = Replace(vfcNomPers, "  ", " ", , , vbTextCompare)
'                    vfcNomPers = Trim(vfcNomPers)
'                    '-----------------------------
'                    vfcNomPersComp = Replace(vfcNomPersComp, "Ñ", "#", , , vbTextCompare)
'                    vfcNomPersComp = Replace(vfcNomPersComp, "ñ", "#", , , vbTextCompare)
'                    vfcNomPersComp = Replace(vfcNomPersComp, "-", "", , , vbTextCompare)
'                    vfcNomPersComp = Replace(vfcNomPersComp, "|", "", , , vbTextCompare)
'                    vfcNomPersComp = Replace(vfcNomPersComp, ".", " ", , , vbTextCompare)
'                    'al final
'                    vfcNomPersComp = Replace(vfcNomPersComp, "   ", " ", , , vbTextCompare)
'                    vfcNomPersComp = Replace(vfcNomPersComp, "  ", " ", , , vbTextCompare)
'                    vfcNomPersComp = Trim(vfcNomPersComp)
'                    '-----------------------------
'                    vfcActEcon = Right(rsPers!cActEcon, 4)
'                    'vfcCodRegPub = IIf(IsNull(rsPers!ccodregpub), "", rsPers!ccodregpub)
'                    'vfcTiDoTr = rsPers!cTidoTr
'                    'vfcNuDoTr = rsPers!cNudoTr
'                    'vfcTiDoCi = rsPers!ctidoci
'                    'vfcNuDoCi = rsPers!cnudoci
'                    vfcTipPers = rsPers!cTipPers
'                    'vfcResid = rsPers!cResid
'                    vfcMagEmp = rsPers!cMagEmp
'                    'vfcAccionista = rsPers!cAccionista
'                    vfcRelInst = rsPers!cRelInst
'                    'vfcPaisNac = rsPers!cPaisNac
'                    vfcSiglas = rsPers!cSiglas
'
'                    vfcCodUnico = lsCodPers
'                    vfcCodSBS = ""
'
'                    'vfcActEcon = fBuscaActividadEconomica(Trim(rs!cPersCod))
'                    vfcActEcon = IIf(IsNull(rs!cActEcon1), "9999", Right(rs!cActEcon1, 4))
'
'                    If vfcTipPers <> "1" Then
'                        'Call fCargaDatosJuridicos(rs!cPersCod, vfcCodRegPub, vfcMagEmp, vfcSiglas)
'                        'If lsProducto = "CREDITOS" Or lsProducto = "PERSONALES" Then
'                        '    'vfcMagEmp = fBuscaMagnitudEmpresarial(rs!cNumfuente)
'                        'End If
'                    Else
'                        vfcCodRegPub = ""
'                        vfcMagEmp = "0"
'                        vfcSiglas = ""
'                    End If
'                    'Insercion del 22/06/2004
'                    If vfcTipPers = "3" Then
'                        vfcTipPers = "2"
'                    End If
'                     'Fin Insercion del 22/06/2004
'
'                    If InStr(1, vfcNomPers, "Y/O", vbTextCompare) <> 0 Or InStr(1, vfcNomPers, " O ", vbTextCompare) <> 0 Then
'                        vfcTipPers = 3
'                    Else
'                        If Trim(rsPers!cTipPers) = 1 Then
'                            vfcTipPers = 1
'                        Else
'                            vfcTipPers = 2
'                        End If
'                    End If
'
'                    ' **** Documento Tributario
'                    vfcNuDoTr = IIf(IsNull(rsPers!NroRUC), "", Trim(rsPers!NroRUC))
'                    'If vfcTipPers = "2" And rs!cTidoTr = "2" And vfcNuDoTr <> "" And Len(vfcNuDoTr) <> 11 Then    ' RUC 11 digitos
'                    '    vfcNuDoTr = PersonaRUC11(rs!cCodPers)
'                    'End If
'                    vfcTiDoTr = IIf(IsNull(rsPers!NroRUC), "", 2)
'
'                    vfcTiDoCi = Trim(IIf(IsNull(rsPers!NroDNI), " ", "1"))
'
'                    vfcNuDoCi = Trim(IIf(IsNull(rsPers!NroDNI), "", rsPers!NroDNI))
'
'                    vfcResid = "1"
'
'                    If rsPers!cRelInst = "A" Then
'                        vfcAccionista = "1"
'                    Else
'                        vfcAccionista = "0"
'                    End If
'
'                    If Not IsNull(rsPers!cRelInst) Then
'                    Select Case Trim(rsPers!cRelInst)
'                        Case "N", "O"  'ninguno/ no indicado
'                            vfcRelInst = "0"
'                        Case "D"        'director
'                            vfcRelInst = "1"
'                        Case "F"        'funcionario
'                            vfcRelInst = "2"
'                        Case "T"        'trabajador
'                            vfcRelInst = "3"
'                        Case "A"
'                            vfcRelInst = "0"
'                        End Select
'                    Else
'                        vfcRelInst = "0"
'                    End If
'
'                    vfcPaisNac = "4028"
'                'Set rsPers = Nothing
'            End If
'
'            '******************** CALIFICACION DE LA PERSONA **************
'            'Set loRCDProceso = New COMNCredito.NCOMRCD
'                'vfcCalificacion = loRCDProceso.nObtieneCalificacionPersonaProcesada(rs!cPersCod, fsServConsol)
'
'            'Set loRCDProceso = Nothing
'            vfcCalificacion = rs!cCalif
'
'            '** Verifica calificacion
'            If vfcCalificacion = "" Then
'                MsgBox "Calificacion en blanco Encontrada en Cliente [" & rs!cPersCod & "] " & lsNombrePersona & Chr(13) & " por favor verifique los datos del cliente o consulte a sistemas", vbInformation, "Aviso"
'            End If
'
'            lsSQL = "INSERT INTO " & fsServConsol & "RCDvc" & lsFecha & "01  " _
'                & "(cTipoFor,cTipoInf,cNumSec, cCodSBS,cPersCod,cActEcon,cCodRegPub, " _
'                & " cTidoTr,cNudoTr,cTiDoci,cNuDoci, cTipPers, cResid,cCalifica,cMagEmp, " _
'                & " cAccionista,cRelInst,cPaisNac, cSiglas,cPersNom,cPersNomCom,cPersGenero,cPersEstado, " _
'                & " CAPEPAT, CAPEMAT,CAPECAS,CNOMBRE1,CNOMBRE2, cIndRCC ) " _
'                & " VALUES('" & lsTipoDocRCD & "','" & lsTipoInfRCD & "',Null," _
'                & "'" & vfcCodSBS & "','" & lsCodPers & "'," _
'                & "'" & vfcActEcon & "'" & "," _
'                & "'" & vfcCodRegPub & "'" & "," _
'                & IIf(vfcTiDoTr = "Null", "Null", "'" & Trim(vfcTiDoTr) & "'") & "," _
'                & IIf(vfcNuDoTr = "Null", "Null", "'" & Trim(vfcNuDoTr) & "'") & ",'" _
'                & Trim(IIf(IsNull(vfcTiDoCi), "1", vfcTiDoCi)) & "'," _
'                & Trim(IIf(IsNull(vfcNuDoCi), "NULL", "'" & Trim(vfcNuDoCi) & "'")) & ",'" _
'                & Trim(vfcTipPers) & "','" _
'                & vfcResid & "','" & vfcCalificacion & "','" & vfcMagEmp & "','" _
'                & vfcAccionista & "','" & vfcRelInst & "','" & vfcPaisNac & "'," _
'                & "'" & vfcSiglas & "'" & ",'" _
'                & Trim(vfcNomPers) & " ','" & Trim(vfcNomPersComp) & "','" & vfcGenero & "','" & vfcEstado & "','" _
'                & Trim(lsPat) & "','" & Trim(lsMat) & "','" & Trim(lsCas) & "','" & Trim(lsNom1) & "','" & Trim(lsNom2) & "','" & lsIndRCC & "')"
'
'            PObjConec.Ejecutar (lsSQL)
'
'            'insertamos la garantia por persona
'            Call GeneraContableGarantiaNEW(lsCodPers, lsFecha, Mid(rs!cCtaCod, 6, 1))
'        End If
'
'        '*******************************************
'        '**  DATOS DE SALDOS DE CUENTAS - RCDvc02 **
'        '*******************************************
'        lsCredito = rs!cCtaCod
'        lbRefinan = IIf((rs!nPrdEstado = gColocEstRefNorm Or rs!nPrdEstado = gColocEstRefMor Or rs!nPrdEstado = gColocEstRefVenc), True, False)
'        If rs!nPrdEstado = gColocEstRecVigJud Or rs!nPrdEstado = gColocEstRecVigCast _
'           Or rs!nPrdEstado = 2205 Or rs!nPrdEstado = 2206 Then
'            lbCobJud = True
'            'lnDiasAtraso = 0
'            lbCastigado = IIf(rs!nPrdEstado = gColocEstRecVigCast Or rs!nPrdEstado = 2206, True, False)
'            'lbDemanda = IIf(rs!nDemanda = "S", True, False)
'            lbDemanda = True
'            lnSaldoCap = rs!nSaldoCap
'            'lnDiasAtraso = 0
'            lnDiasAtraso = rs!nDiasAtraso
'        Else
'            lbCobJud = False
'            lbCastigado = False
'            lbDemanda = False
'            lnSaldoCap = rs!nSaldoCap
'            lnDiasAtraso = rs!nDiasAtraso
'        End If
'
'        ' ********
'
'        lsEstadoCredito = DevEstadoCredito(lsCredito, lnDiasAtraso, lbCobJud)
'        If lnDiasAtraso < 0 Then lnDiasAtraso = 0
'        '***************************************************************
'        ' Genero cuentas Contables de los Creditos que tenga el Cliente
'        '****************************************************************
'
'        ' ******** Cuenta de Saldos de Capital
'
'        If Mid(rs!cCtaCod, 6, 1) = "3" And Mid(rs!cCtaCod, 6, 3) <> "305" And lsEstadoCredito = "VEN" Then
'           'If lsEstadoCredito = "VEN" Then
'               ' Inserta la parte VENCIDA
'               If lnDiasAtraso <= 90 Then ' Vence el Capital de Cuota
'                    lsCuentaCnt = GeneraCuentaCapital(lsCredito, lsEstadoCredito, lbRefinan, False, False, False)
'                    Call GrabaEnRCDSaldos(lsFecha, lsCodPers, lsCuentaCnt, lnDiasAtraso, rs!nCapVenc, Mid(lsCredito, 6, 1))
'
'                    ' Inserta la parte VIGENTE
'                    lsCuentaCnt = GeneraCuentaCapital(lsCredito, "VIG", lbRefinan, False, False, False)
'                    Call GrabaEnRCDSaldos(lsFecha, lsCodPers, lsCuentaCnt, lnDiasAtraso, lnSaldoCap - rs!nCapVenc, Mid(lsCredito, 6, 1))
'               Else
'
'                    lsCuentaCnt = GeneraCuentaCapital(lsCredito, lsEstadoCredito, lbRefinan, False, False, False)
'                    Call GrabaEnRCDSaldos(lsFecha, lsCodPers, lsCuentaCnt, lnDiasAtraso, rs!nSaldoCap, Mid(lsCredito, 6, 1))
'
'               End If
'
'        ElseIf Mid(rs!cCtaCod, 6, 1) = "4" And lsEstadoCredito = "VEN" Then
'           'If lsEstadoCredito = "VEN" Then
'               ' Inserta la parte VENCIDA
'               If lnDiasAtraso <= 90 Then ' Vence el Capital de Cuota
'                    lsCuentaCnt = GeneraCuentaCapital(lsCredito, lsEstadoCredito, lbRefinan, False, False, False)
'                    Call GrabaEnRCDSaldos(lsFecha, lsCodPers, lsCuentaCnt, lnDiasAtraso, rs!nCapVenc, Mid(lsCredito, 6, 1))
'
'                    ' Inserta la parte VIGENTE
'                    lsCuentaCnt = GeneraCuentaCapital(lsCredito, "VIG", lbRefinan, False, False, False)
'                    Call GrabaEnRCDSaldos(lsFecha, lsCodPers, lsCuentaCnt, lnDiasAtraso, lnSaldoCap - rs!nCapVenc, Mid(lsCredito, 6, 1))
'               Else
'
'                    lsCuentaCnt = GeneraCuentaCapital(lsCredito, lsEstadoCredito, lbRefinan, False, False, False)
'                    Call GrabaEnRCDSaldos(lsFecha, lsCodPers, lsCuentaCnt, lnDiasAtraso, rs!nSaldoCap, Mid(lsCredito, 6, 1))
'
'               End If
'
'        Else ' Inserta Capital
'
'               lsCuentaCnt = GeneraCuentaCapital(lsCredito, lsEstadoCredito, lbRefinan, lbCobJud, lbDemanda, lbCastigado)
'               Call GrabaEnRCDSaldos(lsFecha, lsCodPers, lsCuentaCnt, lnDiasAtraso, lnSaldoCap, Mid(lsCredito, 6, 1))
'
'        End If
'
'        ' Cuenta de Intereses *******************
'
'        lnIntDeveng = IIf(IsNull(rs!nIntDeveng), 0, rs!nIntDeveng)
'
'        If lsEstadoCredito = "VIG" Then
'            lsCuentaCnt = "14" & Mid(lsCredito, 9, 1) & "80" & Mid(lsCredito, 6, 1)
'        Else
'            lsCuentaCnt = "81" & Mid(lsCredito, 9, 1) & "40" & Mid(lsCredito, 6, 1)
'        End If
'        Call GrabaEnRCDSaldos(lsFecha, lsCodPers, lsCuentaCnt, 0, lnIntDeveng, Mid(lsCredito, 6, 1))
'
'        ' Interes en Suspenso
'        If IIf(IsNull(rs!nIntSusp), 0, rs!nIntSusp) > 0 Then
'            lsCuentaCnt = "81" & Mid(lsCredito, 9, 1) & "40" & Mid(lsCredito, 6, 1)
'            Call GrabaEnRCDSaldos(lsFecha, lsCodPers, lsCuentaCnt, 0, lnIntDeveng, Mid(lsCredito, 6, 1))
'        End If
'
'
'        ' Provision Especifica
'            If Val(vfcCalificacion) > 0 And lbCastigado = False Then
'                lsCuentaCnt = "14" & Mid(lsCredito, 9, 1) & "90" & Mid(lsCredito, 6, 1) & "01"
'                'lnProvision = ObtieneProvisionCredito(lsCredito)
'                lnProvision = Format(IIf(IsNull(rs!nProvision), 0, rs!nProvision), "#.00")
'                Call GrabaEnRCDSaldos(lsFecha, lsCodPers, lsCuentaCnt, 0, lnProvision, Mid(lsCredito, 6, 1))
'            End If
'
'        ' Cuenta de Garantias  *****************
'
'        If Mid(lsCredito, 6, 3) = "305" Then ' Pignoraticio
'            lsCuentaCnt = "84140206000000"  ' G. Joyas
'            ' Monto en Soles
'            lnMontoGar2 = IIf(IsNull(rs!nGarantSoles), 0, rs!nGarantSoles)
'            If lnMontoGar2 > 0 Then
'                Call GrabaEnRCDSaldos(lsFecha, lsCodPers, lsCuentaCnt, 0, lnMontoGar2, Mid(lsCredito, 3, 1))
'            End If
'        Else
'            'Call GeneraContableGarantia(lsCodPers, lsFecha)
'            'Call GeneraContableGarantiaNEW(lsCodPers, lsFecha)
'        End If
'        ' Cuentas Garantias
'
'        rs.MoveNext
'        barra.value = Int(i / lnTotal * 100)
'        Me.lblAvance = "Avance :" & Format(i / lnTotal * 100, "#0.000") & "%"
'        Me.Caption = "Genera Datos para informe RCD - "
'        DoEvents
'    Loop
'End If
'rs.Close
'Set rs = Nothing
'
'Screen.MousePointer = 0
'barra.value = 0
'Me.lblAvance = ""
'
'End Sub

'************************************************
'*************************************************
'**  Devuelve la Cta contable de Cartera Saldos 14
'Function GeneraCuentaCapital(ByVal pCredito As String, ByVal pEstadoCredito As String, ByVal pRefinan As Boolean, ByVal pCobJud As Boolean, ByVal pDemanda As Boolean, ByVal pCastigado As Boolean) As String
'
'Dim lsCuenta As String
'
'Dim lsSituacion As String
'Dim lsCondicion As String
'
'If pCobJud = True Then    ' Cobranza Judicial
'
'    If pCastigado = True Then ' Si castigado
'        lsCuenta = "81" & Mid(pCredito, 9, 1) & "302"
'        GeneraCuentaCapital = lsCuenta
'        Exit Function
'    Else
'        If pDemanda = True Then
'            lsSituacion = "6"
'            'lsCondicion = IIf(pRefinan = False , "06",  IIf( Mid(pCredito, 3, 1) = "3", "19", "1929"))
'            lsCondicion = IIf(Mid(pCredito, 3, 3) = "423", "23", "06")
'        Else
'            lsSituacion = "5"
'            'lsCondicion = IIf(pRefinan = False, "06", IIf(Mid(pCredito, 3, 1) = "3", "19", "1929")) ' CAMBIO SILVITA
'            lsCondicion = IIf(Mid(pCredito, 3, 3) = "423", "23", "06")
'        End If
'    End If
'
'ElseIf Mid(pCredito, 6, 3) = "305" Then  ' Prendario
'    If pEstadoCredito = "VIG" Then
'        lsSituacion = "1"
'        lsCondicion = "13"
'    Else
'        lsSituacion = "5"
'        lsCondicion = "13"
'    End If
'ElseIf Mid(pCredito, 6, 3) = "320" Then ' Administrativos
'    If pEstadoCredito = "VIG" Then
'        lsSituacion = "1"
'        lsCondicion = "20"
'    Else
'        lsSituacion = "5"
'        lsCondicion = "20"
'    End If
'
'ElseIf Mid(pCredito, 6, 3) = "423" Then ' Administrativos
'    If pEstadoCredito = "VIG" Then
'        lsSituacion = "1"
'        lsCondicion = "23"
'    Else
'        lsSituacion = "5"
'        lsCondicion = "23"
'    End If
'
'Else   ' Comerc - Pyme - Consumo - HipoteCaja
'    If pEstadoCredito = "VIG" Then
'        If pRefinan = False Then
'            lsSituacion = "1"
'            lsCondicion = "06"
'        Else  ' Refinanciado
'            lsSituacion = "4"
'            lsCondicion = "06"
'        End If
'    Else    ' Vencido
'        If pRefinan = False Then
'            lsSituacion = "5"
'            lsCondicion = "06"
'        Else  ' Refinanciado
'            If Mid(pCredito, 6, 1) = "3" Or Mid(pCredito, 6, 1) = "4" Then ' Para Consumo no ha cambiado los refinanciados
'                lsSituacion = "5"
'                lsCondicion = "19"
'            Else
'                lsSituacion = "5"
'                lsCondicion = "1929"  ' Ojo
'            End If
'        End If
'    End If
'End If
'
''***  [14][M][Sit][TC][Cond]
'lsCuenta = "14" & Mid(pCredito, 9, 1) & lsSituacion & "0" & Mid(pCredito, 6, 1) & lsCondicion
'
'GeneraCuentaCapital = lsCuenta
'
'End Function


'**  Devuelva el Estado del Credito
'**  1 = Vigente // 2 = Vencido // 3 = Cob Judicial
'Function DevEstadoCredito(ByVal pCredito As String, ByVal pDiasAtraso As Integer, ByVal pCobJud As String) As String
'Dim lsEstado As String
'
'If pCobJud = True Then
'    lsEstado = "COJ"
'Else
'    Select Case Mid(pCredito, 6, 1)
'            Case "1"
'                If pDiasAtraso <= 15 Then
'                    lsEstado = "VIG"
'                Else
'                    lsEstado = "VEN"
'                End If
'            Case "2", "3", "4"
'                If pDiasAtraso <= 30 Then
'                    lsEstado = "VIG"
'                Else
'                    lsEstado = "VEN"
'                End If
'    End Select
'End If
'DevEstadoCredito = lsEstado
'End Function

'**************************************
'** Devuelve el Monto de la Garantia en la Moneda
'**************************************
Function DevGarantiaMoneda(ByVal pMoneda As String, ByVal pGarantSol As Currency, ByVal pGarantDol As Currency) As Currency
Dim lnValor As Currency
'If pMoneda = "1" Then  ' Soles
'    lnValor = pGarantSol + (pGarantDol * gnTipoCambio)
'Else  ' Dolares
'    If pGarantSol = 0 Then
'        lnValor = pGarantDol
'    Else
'        lnValor = (pGarantSol / gnTipoCambio) + pGarantDol
'    End If
'End If
'DevGarantiaMoneda = CCur(Format(lnValor, "#0.00"))
End Function

'***********************************
'**  Graba en RCD02 - RCD03 ********
'***********************************
'Sub GrabaEnRCDSaldos(ByVal pFecha As String, ByVal pcCodPers As String, _
'    ByVal pcCodCnt As String, ByVal pnDiasAtraso As Integer, ByVal pnSaldo As Currency, _
'    ByVal pTipoCredito As String)
'Dim lsSQL As String
'Dim rs As ADODB.Recordset
'Dim PObjConec As COMConecta.DCOMConecta
'Dim lbExiste As Boolean
'Dim lnSaldoSoles As Currency
'
'Dim rsVerif As New ADODB.Recordset ' Para Verificar
'
'If pnSaldo <= 0 Then
'    Exit Sub
'End If
'
''****  Cambio a Moneda Soles  ****
'If Mid(pcCodCnt, 3, 1) = "1" Then
'    lnSaldoSoles = Format(pnSaldo, "#0.00")
'Else
'    lnSaldoSoles = Format(pnSaldo * fnTipCambio, "#0.00")
'End If
'
''*******************************************
''**  SALDOS POR CUENTA CONTABLE  - RCDvc02 **
''*******************************************
'' -----** Graba en RCDvc02
'
'Set PObjConec = New COMConecta.DCOMConecta
'    PObjConec.AbreConexion
'lsSQL = "SELECT COUNT(cPersCod) hay From " & fsServConsol & "RCDvc" & pFecha & "02  WHERE cPersCod ='" & pcCodPers & "' " _
'      & "AND cCtaCnt ='" & pcCodCnt & "' AND nCondDias = " & pnDiasAtraso & " AND cTipoCred = '" & pTipoCredito & "' "
'Set rs = PObjConec.CargaRecordSet(lsSQL)
'
'lbExiste = IIf(rs!hay > 0, True, False)
'rs.Close
'Set rs = Nothing
'
'If lbExiste = False Then
'     lsSQL = "INSERT INTO " & fsServConsol & "RCDvc" & pFecha & "02 (cTipoFor,cTipoInf,cNumSec, " _
'      & "cCodAge,cUbicGeo, cCtaCnt, cTipoCred,nSaldo, nCondDias,cPersCod) " _
'      & " VALUES ('1','2',Null,'" & fsCodOfiInf & "','" & fsUbicaGeoRCD & "','" _
'      & pcCodCnt & "','" & pTipoCredito & "'," & lnSaldoSoles & "," & pnDiasAtraso & ",'" & pcCodPers & "') "
'Else
'     lsSQL = "UPDATE " & fsServConsol & "RCDvc" & pFecha & "02  SET nSaldo = nSaldo + " & lnSaldoSoles & "" _
'      & "WHERE cPersCod ='" & pcCodPers & "' AND  cCtaCnt = '" & pcCodCnt & "' " _
'      & "AND nCondDias =" & pnDiasAtraso & " AND cTipoCred = '" & pTipoCredito & "' "
'End If
'PObjConec.Ejecutar lsSQL
'
'
''*******************************************
''**  TOTALES DE CMACT X CUENTAS - RCDvc03 **
''*******************************************
'' -----** Graba en RCDvc03
'
'lsSQL = "SELECT COUNT(cCtaCnt) hay From " & fsServConsol & "RCDvc" & pFecha & "03 WHERE cCtaCnt ='" & pcCodCnt _
'     & "' AND cTipoCred = '" & pTipoCredito & "' AND nCondDias = " & pnDiasAtraso & " "
'Set rs = PObjConec.CargaRecordSet(lsSQL)
'
'lbExiste = IIf(rs!hay > 0, True, False)
'rs.Close
'Set rs = Nothing
'
'If lbExiste = False Then
'     lsSQL = "INSERT INTO " & fsServConsol & "RCDvc" & pFecha & "03 (cTipoFor,cTipoInf,cNumSec,cCtaCnt, cTipoCred,nSaldo, nCondDias) " _
'           & " VALUES ('2','2',Null,'" & pcCodCnt & "','" & pTipoCredito & "'," & lnSaldoSoles & "," & pnDiasAtraso & " )"
'Else
'     lsSQL = "UPDATE " & fsServConsol & "RCDvc" & pFecha & "03  SET nSaldo = nSaldo + " & lnSaldoSoles _
'           & " WHERE cCtaCnt = '" & pcCodCnt & "' AND cTipoCred = '" & pTipoCredito & "' AND nCondDias = " & pnDiasAtraso & " "
'End If
'
'PObjConec.Ejecutar lsSQL
'
'End Sub

'Private Function ExisteRCDMaestroPersona(ByVal psCodPers As String) As Boolean
'Dim lsSQL As String
'Dim lr As ADODB.Recordset
'lsSQL = "Select * from RCDMaestroPersona WHERE cCodPers='" & Trim(psCodPers) & "'"
'
''lr.Open lsSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
'
'If lr.BOF And lr.EOF Then
'    ExisteRCDMaestroPersona = False
'Else
'    ExisteRCDMaestroPersona = True
'    vfcCodUnico = lr!cCodUnico
'    vfcCodSBS = lr!cCodSBS
'    vfcNomPers = lr!cNomPers
'    If InStr(1, vfcNomPers, "'", vbTextCompare) <> 0 Then
'        vfcNomPers = Replace(vfcNomPers, "'", "''", , , vbTextCompare)
'    End If
'    vfcNomPers = Replace(vfcNomPers, "Ñ", "#", , , vbTextCompare)
'    vfcNomPers = Replace(vfcNomPers, "ñ", "#", , , vbTextCompare)
'    vfcNomPers = Replace(vfcNomPers, "-", "", , , vbTextCompare)
'    vfcNomPers = Replace(vfcNomPers, "|", "", , , vbTextCompare)
'    vfcNomPers = Replace(vfcNomPers, ".", " ", , , vbTextCompare)
'    'al final
'    vfcNomPers = Replace(vfcNomPers, "   ", " ", , , vbTextCompare)
'    vfcNomPers = Replace(vfcNomPers, "  ", " ", , , vbTextCompare)
'    vfcNomPers = Trim(vfcNomPers)
'
'    vfcActEcon = lr!cActEcon
'    vfcCodRegPub = IIf(IsNull(lr!ccodregpub), "", lr!ccodregpub)
'    vfcTiDoTr = lr!cTidoTr
'    vfcNuDoTr = lr!cNudOtr
'    vfcTiDoCi = lr!ctidoci
'    vfcNuDoCi = lr!cnudoci
'    vfcTipPers = lr!cTipPers
'    vfcResid = lr!cResid
'    vfcMagEmp = lr!cMagEmp
'    vfcAccionista = lr!cAccionista
'    vfcRelInst = lr!cRelInst
'    vfcPaisNac = lr!cPaisNac
'    vfcSiglas = lr!cSiglas
'End If
'lr.Close
'
'End Function

'Private Function fBuscaActividadEconomica(ByVal lsCodPers As String) As String
'Dim lsSQL As String
'Dim rsF As ADODB.Recordset
'Dim PObjConec As COMConecta.DCOMConecta
'Dim lsActividad As String
'lsActividad = ""
'    lsSQL = "SELECT FuenteIngreso.cNumFuente, FuenteIngreso.cActEcon, " _
'        & " FuenteIngreso.cSector, FuenteIngreso.nTipoFuente " _
'        & " FROM " & fsServConsol & "FuenteIngresoConsol FuenteIngreso " _
'        & " WHERE FuenteIngreso.cPersCod = '" & Trim(lsCodPers) _
'        & "' ORDER BY cActEcon Desc"
'
'    Set PObjConec = New COMConecta.DCOMConecta
'        PObjConec.AbreConexion
'        Set rsF = PObjConec.CargaRecordSet(lsSQL)
'    Set PObjConec = Nothing
'
'    If Not (rsF.BOF And rsF.EOF) Then
'        fBuscaActividadEconomica = IIf(IsNull(rsF!cActEcon), "", Right(rsF!cActEcon, 4))
'        If fBuscaActividadEconomica = "9999" Then
'            fBuscaActividadEconomica = ""
'        End If
'    Else
'        fBuscaActividadEconomica = ""
'    End If
'    rsF.Close
'    Set rsF = Nothing
'End Function

Private Sub fCargaDatosJuridicos(ByVal psCodPers As String, psCodRegPub As String, psMagEmp As String, psSiglas As String)
Dim lsSQL As String
Dim rsJ As New ADODB.Recordset
Dim lbEncuentroMaestroICC As Boolean

'lsSQL = "SELECT cCodPers, cCodRegPub, cMagEmp, cSigla FROM PersonaJur " _
'      & "WHERE cCodPers='" & Trim(psCodPers) & "'"
'Set rsJ = CargaRecord(lsSQL)
'If Not RSVacio(rsJ) Then
'    psCodRegPub = IIf(IsNull(rsJ!ccodregpub), "", rsJ!ccodregpub)
'    psMagEmp = IIf(IsNull(rsJ!cMagEmp), "", rsJ!cMagEmp)
'    psSiglas = IIf(IsNull(rsJ!csigla), "", rsJ!csigla)
'End If
'rsJ.Close
'Set rsJ = Nothing

End Sub

Public Function fBuscaMagnitudEmpresarial(ByVal psNumFuente As String) As String
Dim lsSQL As String
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim lnTotal As Currency
Dim lnSaldo As Currency
Dim lnValorUIT As Currency
Dim lsMagnitudEmpresarial As String

'lsSQL = "Select * From FuenteIngreso WHERE cNumFuente='" & psNumFuente & "'"
'rs.Open lsSQL, dbCmact, adOpenForwardOnly, adLockOptimistic, adCmdText
'If Not RSVacio(rs) Then
'    If rs!cTipoFuente = "I" Then
'        lsSQL = "select * From BALANCE  WHERE cNumFuente='" & psNumFuente & "' order by dFecBalanc desc"
'        rs1.Open lsSQL, dbCmact, adOpenForwardOnly, adLockOptimistic, adCmdText
'        If Not RSVacio(rs1) Then
'            lnSaldo = CCur(Format((IIf(IsNull(rs1!nventas), 0, rs1!nventas) + IIf(IsNull(rs1!nrecuctas), 0, rs1!nrecuctas) + IIf(IsNull(rs1!ningfam), 0, rs1!ningfam)) - (IIf(IsNull(rs1!ncostovent), 0, rs1!ncostovent) + IIf(IsNull(rs1!notrosegr), 0, rs1!notrosegr) + IIf(IsNull(rs1!nGasFam), 0, rs1!nGasFam)), "####0.00"))
'        End If
'        rs1.Close
'        Set rs1 = Nothing
'    Else
'        lsSQL = "select * From fdependiente  WHERE cNumFuente='" & psNumFuente & "' order by dFecEval desc"
'        rs1.Open lsSQL, dbCmact, adOpenForwardOnly, adLockOptimistic, adCmdText
'        If Not RSVacio(rs1) Then
'            lnSaldo = CCur(Format(IIf(IsNull(rs1!nIngClt), 0, rs1!nIngClt) + IIf(IsNull(rs1!nIngCon), 0, rs1!nIngCon) + IIf(IsNull(rs1!nOtrIng), 0, rs1!nOtrIng) - IIf(IsNull(rs1!nGasFam), 0, rs1!nGasFam), "####0.00"))
'        End If
'        rs1.Close
'        Set rs1 = Nothing
'    End If
'Else
'End If
'rs.Close
'Set rs = Nothing
'lnValorUIT = gnValorUIT
'lnTotal = lnSaldo * 12
'Select Case lnTotal
'    Case Is = 0
'        lsMagnitudEmpresarial = "4"
'    Case Is < 300 * lnValorUIT
'        lsMagnitudEmpresarial = "4"
'    Case Is < 600 * lnValorUIT
'        lsMagnitudEmpresarial = "3"
'    Case Is < 10000 * lnValorUIT
'        lsMagnitudEmpresarial = "2"
'    Case Else
'        lsMagnitudEmpresarial = "1"
'End Select

fBuscaMagnitudEmpresarial = lsMagnitudEmpresarial

End Function


'Private Function EmiteCodigoAux(ByVal lsCodPers As String, ByVal psServConsol As String) As String
'
'Dim PObjConec As COMConecta.DCOMConecta
'Dim SQL As String
'Dim rs As New ADODB.Recordset
'Set PObjConec = New COMConecta.DCOMConecta
'SQL = "Select * from " & psServConsol & "RCDCodigoAux where cPersCod ='" & Trim(lsCodPers) & "'"
'PObjConec.AbreConexion
'Set rs = PObjConec.CargaRecordSet(SQL)
'PObjConec.CierraConexion
''Rs.Open SQl, dbCmact, adOpenForwardOnly, adLockOptimistic, adCmdText
'If Not RSVacio(rs) Then
'    EmiteCodigoAux = Trim(rs!cCodAux)
'Else
'    EmiteCodigoAux = ""
'End If
'rs.Close
'Set rs = Nothing
'Set PObjConec = Nothing
'End Function

'Private Function GeneraContableGarantia(ByVal psPersCod As String, ByVal psFecha As String)
'Dim SQL As String
'Dim rs As ADODB.Recordset
'Dim lsCtaCont As String
'Dim lsDH As String
'Dim lsTipoPref As String
'Dim lsCodGaranCont As String
'Dim lsCodAge As String
'Dim lsTipoCred As String
'Dim lsMoneda As String
'Dim lnDebe As Currency
'Dim lnHaber As Currency
'Dim lsOpeCod As String
'Dim lxCta As String
'Dim lnMontoGar As Currency
'
'Dim ocon As COMConecta.DCOMConecta
'Set ocon = New COMConecta.DCOMConecta
'
'
'SQL = "         SELECT  G.cPersCod, G.nTipoGarant, G.nMoneda,  "
'SQL = SQL & "           SUBSTRING(C.cCtaCod,6,1) as nProducto, SUBSTRING(C.cCtaCod,4,2) AS cCodAge,"
'SQL = SQL & "           T.ccodcnt, T.ctippref, sum(G.nMontoRealiz) as nMonto"
'SQL = SQL & "   FROM    " & fsServConsol & "GarantiasConsol G"
'SQL = SQL & "           JOIN " & fsServConsol & "GarantCredConsol GC ON GC.cNumGarant = G.cNumGarant AND GC.cPersCod = G.cPersCod"
'SQL = SQL & "           JOIN " & fsServConsol & "CreditoConsol C ON C.cCtaCod = GC.cCtaCod"
'SQL = SQL & "           LEFT JOIN " & fsServConsol & "TIPOGARAN T ON T.nTpoGarantia = G.nTipoGarant"
'SQL = SQL & "   WHERE   C.nPrdEstado in (2020,2021,2022,2030,2031,2032,2201,2202,2205,2206,2101,2104,2106,2107)"
'SQL = SQL & "           and G.cPersCod ='" & psPersCod & "'"
'SQL = SQL & "           group by G.cPersCod, nTipoGarant,G.nMoneda, SUBSTRING(C.cCtaCod,6,1), SUBSTRING(C.cCtaCod,4,2), ccodcnt , ctippref"
'
'ocon.AbreConexion
'Set rs = ocon.CargaRecordSet(SQL)
'ocon.CierraConexion
'
'If Not rs.EOF And Not rs.BOF Then
'    Do While Not rs.EOF
'        lsCodAge = Trim(rs!cCodAge)
'        lsCtaCont = lsCtaNoPref
'        lsMoneda = Trim(rs!nmoneda)
'        lsTipoPref = "02" 'Right(rs!CPREF, 2)
'
'        If Not IsNull(rs!ccodcnt) Then
'            lsCodGaranCont = rs!ccodcnt
'        Else
'            MsgBox "No se ha establecido cuenta Contable para tipo de Garantía tipo [" & rs!nTipoGarant & "]", vbInformation, "Aviso"
'        End If
'
'        lsTipoCred = rs!nProducto
'        Select Case lsTipoCred
'            Case 2 'Microempresa
'                lsTipoCred = "02"
'            Case 3 'Consumo
'                lsTipoCred = "03"
'            Case 1 'Comerciales
'                lsTipoCred = "01"
'        End Select
'
'        lsCtaCont = Replace(lsCtaCont, "M", lsMoneda)
'        lsCtaCont = Replace(lsCtaCont, "PF", lsTipoPref)
'        lsCtaCont = Replace(lsCtaCont, "TG", lsCodGaranCont)
'        lsCtaCont = Replace(lsCtaCont, "TC", lsTipoCred)
'        lsCtaCont = Replace(lsCtaCont, "AG", lsCodAge)
'
'        lxCta = Left(Trim(lsCtaCont), 8)
'        If Left(lsCtaCont, 2) & "0" & Mid(lsCtaCont, 4, 3) = "840409" Or Left(lsCtaCont, 2) & "0" & Mid(lsCtaCont, 4, 3) = "840403" Then
'            lxCta = Left(Trim(lsCtaCont), 6)
'        End If
'        lxCta = lxCta + String(14 - Len(lxCta), "0")
'
'        lnMontoGar = rs!nMonto
'        If lnMontoGar > 0 Then
'            Call GrabaEnRCDSaldos(psFecha, psPersCod, lxCta, 0, lnMontoGar, rs!nProducto)
'        End If
'
'        rs.MoveNext
'    Loop
'Else
'    'MsgBox "No existen garantías registrados", vbInformation, "Aviso"
'End If
'rs.Close
'Set rs = Nothing
'
'End Function

'Sub CargaGarantiasTotal()
'Dim ocon As COMConecta.DCOMConecta
'Dim SQL As String
'Set ocon = New COMConecta.DCOMConecta
'
''sql = "         SELECT  G.cPersCod, G.nTipoGarant, G.nMoneda,  "
''sql = sql & "           SUBSTRING(C.cCtaCod,6,1) as nProducto, SUBSTRING(C.cCtaCod,4,2) AS cCodAge,"
''sql = sql & "           T.ccodcnt, T.ctippref, sum(G.nMontoRealiz) as nMonto"
''sql = sql & "   FROM    " & fsServConsol & "GarantiasConsol G"
''sql = sql & "           JOIN " & fsServConsol & "GarantCredConsol GC ON GC.cNumGarant = G.cNumGarant AND GC.cPersCod = G.cPersCod"
''sql = sql & "           JOIN " & fsServConsol & "CreditoConsol C ON C.cCtaCod = GC.cCtaCod"
''sql = sql & "           LEFT JOIN " & fsServConsol & "TIPOGARAN T ON T.nTpoGarantia = G.nTipoGarant"
''sql = sql & "   WHERE   C.nPrdEstado in (2020,2021,2022,2030,2031,2032,2201,2202,2205,2206,2101,2104,2106,2107)"
''SQL = SQL & "           and G.cPersCod ='" & psPersCod & "'"
''sql = sql & "           group by G.cPersCod, nTipoGarant,G.nMoneda, SUBSTRING(C.cCtaCod,6,1), SUBSTRING(C.cCtaCod,4,2), ccodcnt , ctippref"
'
''sql = "         SELECT  GC.cPersCod, G.cTipoGarant, G.nMoneda,"
''sql = sql & "           SUBSTRING(C.cCtaCod,6,1) as nProducto, SUBSTRING(C.cCtaCod,4,2) AS cCodAge,"
''sql = sql & "           T.ccodcnt, T.ctippref, sum(G.nMonto) as nMonto"
''sql = sql & "   from    " & fsServConsol & "CredGarantiasConsol G"
''sql = sql & "           JOIN " & fsServConsol & "GarantCredConsol GC ON GC.cctacod  = G.cctacod"
''sql = sql & "           JOIN " & fsServConsol & "CreditoConsol C ON C.cCtaCod = GC.cCtaCod"
''sql = sql & "           LEFT JOIN " & fsServConsol & "TIPOGARAN T ON T.nTpoGarantia = G.cTipoGarant"
''sql = sql & "   WHERE   C.nPrdEstado in (2020,2021,2022,2030,2031,2032,2201,2202,2205,2206,2101,2104,2106,2107)"
''sql = sql & "           AND CONVERT(VARCHAR(12),G.DFECHA,112) = '20050228' "
''sql = sql & "   GROUP BY GC.cPersCod, cTipoGarant,G.nMoneda, SUBSTRING(C.cCtaCod,6,1), SUBSTRING(C.cCtaCod,4,2), ccodcnt , ctippref"
'
'SQL = "         SELECT  G.cPersCod, G.nTipoGarant, G.nMoneda,"
'SQL = SQL & "           T.ccodcnt, T.ctippref, sum(G.nMontoRealiz) as nMonto"
'SQL = SQL & "           FROM    " & fsServConsol & "GarantiasConsol G"
'SQL = SQL & "                   LEFT JOIN " & fsServConsol & "TIPOGARAN T ON T.nTpoGarantia = G.nTipoGarant"
'SQL = SQL & "           WHERE   Exists (    select  GC.cPersCod"
'SQL = SQL & "                               from    " & fsServConsol & "GarantCredConsol GC"
'SQL = SQL & "                                       JOIN " & fsServConsol & "CreditoConsol C ON C.cCtaCod = GC.cCtaCod"
'SQL = SQL & "                               where   C.nPrdEstado in (2020,2021,2022,2030,2031,2032,2201,2202,2205,2206,2101,2104,2106,2107)"
'SQL = SQL & "                                       and GC.cPersCod = G.cPersCod AND GC.cNumGarant = G.cNumGarant )"
'SQL = SQL & "           group by G.cPersCod, nTipoGarant,G.nMoneda,  ccodcnt , ctippref"
'
'ocon.AbreConexion
'Set rsGarant = ocon.CargaRecordSet(SQL)
'rsGarant.Sort = "cPersCod"
'ocon.CierraConexion
'End Sub

'Private Function GeneraContableGarantiaNEW(ByVal psPersGarantCod As String, ByVal psFecha As String, ByVal psProducto As String)
'Dim lsCtaCont As String
'Dim lsDH As String
'Dim lsTipoPref As String
'Dim lsCodGaranCont As String
'Dim lsCodAge As String
'Dim lsTipoCred As String
'Dim lsMoneda As String
'Dim lnDebe As Currency
'Dim lnHaber As Currency
'Dim lsOpeCod As String
'Dim lxCta As String
'Dim lnMontoGar As Currency
'
'If rsGarant Is Nothing Then Exit Function
'
'rsGarant.MoveFirst
'rsGarant.Find "cPersCod ='" & psPersGarantCod & "'"
'If Not rsGarant.EOF And Not rsGarant.BOF Then
'
'    Do While Not rsGarant.EOF And rsGarant!cPersCod = psPersGarantCod
'        lsCodAge = "01" 'Trim(rsGarant!cCodAge)
'        lsCtaCont = lsCtaNoPref
'        lsMoneda = Trim(rsGarant!nmoneda)
'        lsTipoPref = "02" 'Right(rsGarant!CPREF, 2)
'
'        If Not IsNull(rsGarant!ccodcnt) Then
'            lsCodGaranCont = rsGarant!ccodcnt
'        Else
'            MsgBox "No se ha establecido cuenta Contable para tipo de Garantía tipo [" & rsGarant!nTipoGarant & "]", vbInformation, "Aviso"
'        End If
'
'        'lsTipoCred = rsGarant!nProducto
'        lsTipoCred = psProducto
'        Select Case lsTipoCred
'            Case 2 'Microempresa
'                lsTipoCred = "02"
'            Case 3 'Consumo
'                lsTipoCred = "03"
'            Case 1 'Comerciales
'                lsTipoCred = "01"
'        End Select
'
'        lsCtaCont = Replace(lsCtaCont, "M", lsMoneda)
'        lsCtaCont = Replace(lsCtaCont, "PF", lsTipoPref)
'        lsCtaCont = Replace(lsCtaCont, "TG", lsCodGaranCont)
'        lsCtaCont = Replace(lsCtaCont, "TC", lsTipoCred)
'        lsCtaCont = Replace(lsCtaCont, "AG", lsCodAge)
'
'        lxCta = Left(Trim(lsCtaCont), 8)
'        If Left(lsCtaCont, 2) & "0" & Mid(lsCtaCont, 4, 3) = "840409" Or Left(lsCtaCont, 2) & "0" & Mid(lsCtaCont, 4, 3) = "840403" Then
'            lxCta = Left(Trim(lsCtaCont), 6)
'        End If
'        lxCta = lxCta + String(14 - Len(lxCta), "0")
'
'        lnMontoGar = rsGarant!nMonto
'        If lnMontoGar > 0 Then
'            Call GrabaEnRCDSaldos(psFecha, psPersGarantCod, lxCta, 0, lnMontoGar, psProducto) 'rsGarant!nProducto
'        End If
'
'        rsGarant.MoveNext
'        If rsGarant.EOF Then Exit Do
'    Loop
'Else
'    'MsgBox "No existen garantías registrados", vbInformation, "Aviso"
'End If
''rsGarant.Close
''Set rsGarant = Nothing
'
'End Function

'Sub CambiaNombreRCD(ByVal lsNomCli As String)
'Dim lsNombre As String
'Dim lnPos As String
'Dim lsRazon As String
'
'lsNombre = Trim(lsNomCli)    ' && ELIMINAMOS BLANCOS A LA IZQUIERDA
'lnPos = 0
'lsPat = ""
'lsMat = ""
'lsCas = ""
'lsNom1 = ""
'lsNom2 = ""
'
'lsRazon = ""
'lnPos = InStr(1, lsNombre, "/")
'If lnPos > 0 Then
'     lsPat = Trim(Mid(lsNombre, 1, lnPos - 1))
'     lsPat = Trim(Replace(lsPat, "-", ""))
'     lsNombre = Trim(Mid(lsNombre, lnPos + 1, Len(lsNombre)))
'     lnPos = InStr(1, lsNombre, "\")
'     If lnPos > 0 Then
'         lsMat = Trim(Mid(lsNombre, 1, lnPos - 1))
'         lsMat = Trim(Replace(Replace(lsMat, "-", ""), "\", " "))
'         lsNombre = Trim(Mid(lsNombre, lnPos + 1, Len(lsNombre)))
'         lnPos = InStr(1, lsNombre, ",")
'         If lnPos > 0 Then
'             lsCas = Trim(Mid(lsNombre, 1, lnPos - 1))
'             lsCas = Trim(Replace(Replace(lsCas, "-", ""), "\", " "))
'             lsNombre = Trim(Mid(lsNombre, lnPos + 1, Len(lsNombre)))
'             lnPos = InStr(1, lsNombre, " ")
'             If lnPos > 0 Then
'                 lsNom1 = Mid(lsNombre, 1, lnPos - 1)
'                 lsNom2 = Mid(lsNombre, lnPos + 1, Len(lsNombre))
'             Else
'                 lsNom1 = Trim(lsNombre)
'             End If
'         End If
'     Else
'         lnPos = InStr(1, lsNombre, ",")
'         If lnPos > 0 Then
'             lsMat = Mid(lsNombre, 1, lnPos - 1)
'             lsMat = Trim(Replace(Replace(lsMat, "-", ""), "\", " "))
'             lsNombre = Trim(Mid(lsNombre, lnPos + 1, Len(lsNombre)))
'             lnPos = InStr(1, lsNombre, " ")
'             If lnPos > 0 Then
'                 lsNom1 = Mid(lsNombre, 1, lnPos - 1)
'                 lsNom2 = Mid(lsNombre, lnPos + 1, Len(lsNombre))
'             Else
'                 lsNom1 = Trim(lsNombre)
'             End If
'         End If
'     End If
' Else
'     lnPos = InStr(1, lsNombre, "\")
'     If lnPos > 0 Then
'         lsMat = Trim(Mid(lsNombre, 1, lnPos - 1))
'         lsMat = Trim(Replace(Replace(lsMat, "-", ""), "\", " "))
'         lsNombre = Trim(Mid(lsNombre, lnPos + 1, Len(lsNombre)))
'         lnPos = InStr(1, lsNombre, ",")
'         If lnPos > 0 Then
'             lsCas = Trim(Mid(lsNombre, 1, lnPos - 1))
'             lsCas = Trim(Replace(Replace(lsCas, "-", ""), "\", " "))
'             lsNombre = Trim(Mid(lsNombre, lnPos + 1, Len(lsNombre)))
'             lnPos = InStr(1, lsNombre, " ")
'             If lnPos > 0 Then
'                 lsNom1 = Mid(lsNombre, 1, lnPos - 1)
'                 lsNom2 = Mid(lsNombre, lnPos + 1, Len(lsNombre))
'             Else
'                 lsNom1 = Trim(lsNombre)
'             End If
'         End If
'     Else
'         lnPos = InStr(1, lsNombre, ",")
'         If lnPos > 0 Then
'             lsMat = Mid(lsNombre, 1, lnPos - 1)
'             lsMat = Trim(Replace(Replace(lsMat, "-", ""), "\", " "))
'             lsNombre = Trim(Mid(lsNombre, lnPos + 1, Len(lsNombre)))
'             lnPos = InStr(1, lsNombre, " ")
'             If lnPos > 0 Then
'                 lsNom1 = Mid(lsNombre, 1, lnPos - 1)
'                 lsNom2 = Mid(lsNombre, lnPos + 1, Len(lsNombre))
'             Else
'                 lsNom1 = Trim(lsNombre)
'             End If
'         Else
'             lsRazon = lsNombre
'         End If
'     End If
' End If
'
' If lsRazon <> "" Then
'     lsPat = Trim(lsRazon)
'     lsMat = ""
'     lsNom1 = ""
'     lsNom2 = ""
'     lsCas = ""
' Else
'     lsPat = Trim(IIf(lsPat = "", "XXXX", lsPat))
'     lsMat = Trim(IIf(lsMat = "", "XXXX", lsMat))
'     If InStr(1, Trim(lsCas), "DE LA ") = 0 Then
'         lnPos = InStr(1, Trim(lsCas), "DE ")
'         If lnPos > 0 Then
'             lsCas = Trim(Mid(lsCas, lnPos + 3, Len(lsCas)))
'         End If
'     End If
' End If
'
' lsPat = Replace(lsPat, "Ñ", "#")
' lsPat = Replace(lsPat, "-", "")
' lsPat = Replace(lsPat, "|", "")
' lsPat = Replace(lsPat, ".", " ")
' lsPat = Replace(lsPat, "   ", " ")
' lsPat = Replace(lsPat, "  ", " ")
'
' lsMat = Replace(lsMat, "Ñ", "#")
' lsMat = Replace(lsMat, "-", "")
' lsMat = Replace(lsMat, "|", "")
' lsMat = Replace(lsMat, ".", " ")
' lsMat = Replace(lsMat, "   ", " ")
' lsMat = Replace(lsMat, "  ", " ")
'
' lsCas = Replace(lsCas, "Ñ", "#")
' lsCas = Replace(lsCas, "-", "")
' lsCas = Replace(lsCas, "|", "")
' lsCas = Replace(lsCas, ".", " ")
' lsCas = Replace(lsCas, "   ", " ")
' lsCas = Replace(lsCas, "  ", " ")
'
' lsNom1 = Replace(lsNom1, "Ñ", "#")
' lsNom1 = Replace(lsNom1, "-", "")
' lsNom1 = Replace(lsNom1, "|", "")
' lsNom1 = Replace(lsNom1, ".", " ")
' lsNom1 = Replace(lsNom1, "   ", " ")
' lsNom1 = Replace(lsNom1, "  ", " ")
'
' lsNom2 = Replace(lsNom2, "Ñ", "#")
' lsNom2 = Replace(lsNom2, "-", "")
' lsNom2 = Replace(lsNom2, "|", "")
' lsNom2 = Replace(lsNom2, ".", " ")
' lsNom2 = Replace(lsNom2, "   ", " ")
' lsNom2 = Replace(lsNom2, "  ", " ")
'
'End Sub
