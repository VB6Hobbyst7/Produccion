VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRCDGeneraDatosRCD_NEW 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informe RCD - Genera Datos para Informe RCD"
   ClientHeight    =   3840
   ClientLeft      =   3690
   ClientTop       =   4080
   ClientWidth     =   6390
   Icon            =   "frmRCDGeneraDatosRCD_NEW.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   6390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdConsolidaRCDH 
      Caption         =   "Consolida Data RCD&H"
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
      Left            =   240
      TabIndex        =   15
      Top             =   3240
      Width           =   2535
   End
   Begin VB.CommandButton cmdConsolidaRCM 
      Caption         =   "Consolida Data RC&M"
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
      Left            =   240
      TabIndex        =   9
      Top             =   5520
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton cmdConsolidaGeneraSecuencia 
      Caption         =   "&Genera Secuencia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   240
      TabIndex        =   8
      Top             =   2040
      Width           =   2535
   End
   Begin VB.CommandButton cmdConsolidaRCA 
      Caption         =   "Consolida Data RC&A"
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
      Left            =   240
      TabIndex        =   7
      Top             =   1440
      Width           =   2535
   End
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
      Left            =   5040
      TabIndex        =   4
      Top             =   3240
      Width           =   1140
   End
   Begin MSComctlLib.ProgressBar barra 
      Height          =   225
      Left            =   240
      TabIndex        =   3
      Top             =   5160
      Width           =   5955
      _ExtentX        =   10504
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdConsolidaRCD 
      Caption         =   "Consolida Data &RCD"
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
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   2535
   End
   Begin VB.CommandButton cmdConsolidaRCDT 
      Caption         =   "Consolida Data RCD&T"
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
      Left            =   240
      TabIndex        =   14
      Top             =   2640
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "CON DISEÑO DE REGISTRO Resolucion SBS Nro 1699-2003"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   435
      Left            =   3000
      TabIndex        =   11
      Top             =   840
      Width           =   2775
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fecha  Consolidacion :"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   285
      Width           =   1875
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label5 
      Caption         =   "Layg-2007-2b"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   210
      Left            =   3000
      TabIndex        =   13
      Top             =   1920
      Width           =   3315
   End
   Begin VB.Label Label2 
      Caption         =   "Resolucion SBS (Diciemb2006) "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   210
      Left            =   3000
      TabIndex        =   12
      Top             =   1680
      Width           =   3315
   End
   Begin VB.Label Label4 
      Caption         =   "Resolucion SBS 426-2006 (Julio2006)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   210
      Left            =   3000
      TabIndex        =   10
      Top             =   1320
      Width           =   3315
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
      Left            =   2280
      TabIndex        =   1
      Top             =   240
      Width           =   1380
   End
End
Attribute VB_Name = "frmRCDGeneraDatosRCD_NEW"
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
Dim lsCtaNoPref As String
Dim rsGarant As ADODB.Recordset

Dim lsPat  As String
Dim lsMat  As String
Dim lsCas  As String
Dim lsNom1 As String
Dim lsNom2 As String
Dim lsObserv As String 'NAGL 20200718

'JIPR20210408 AGREGA PARÁMETRO RCDH CAMBIO TIPO DE DATO 1 RCDTvc,0  RCDvc , 2  RCDHvc
Private Sub cmdConsolidaRCDH_Click()
Dim loRCDproc As COMNCredito.NCOMRCD
Set loRCDproc = New COMNCredito.NCOMRCD
Me.Enabled = False
ActivaConfigFormBarra True
'Crea Tablas RCDH del mes
Call loRCDproc.nCreaTablasRCD(Format(gdFecDataFM, "yyyymm"), fsServConsol, 2)
barra.value = 4

'Obtiene Datos para procesar el RCDH
Call loRCDproc.ObtenerDatosParaGenerarRCDvc_Tvc(Format(gdFecDataFM, "yyyymm"), fsServConsol, fnTipCambio, fnMontoMinimoRCD, 2)
barra.value = 5

'Procesar Datos en tablas RCDH
Call GeneraDatosRCD(Format(gdFecDataFM, "yyyymm"), loRCDproc, 2)
barra.value = 7

Me.lblAvance = ""
Me.lblDescripcion = ""
Me.Enabled = True
If lsObserv = "" Then
    Call loRCDproc.GeneraSecuenciaRCD_RCA(Format(gdFecDataFM, "yyyymm"), fsServConsol, 2)
    Set loRCDproc = Nothing
    barra.value = 8
    MsgBox "El Proceso de RCDH ha culminado", vbInformation, "Aviso"
End If
lsObserv = ""
ActivaConfigFormBarra True
End Sub
'JIPR20210408 AGREGA PARÁMETRO RCDH CAMBIO TIPO DE DATO 1 RCDTvc,0  RCDvc , 2  RCDHvc

Private Sub Form_Load()
Dim loRCDproc As COMNCredito.NCOMRCD
Dim lrPar As ADODB.Recordset
Dim loConstSist As COMDConstSistema.NCOMConstSistema
Dim oRCD As COMDCredito.DCOMRCD
lsObserv = "" 'NAGL 20200718
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
        cmdConsolidaRCD.Enabled = False
        cmdConsolidaRCDT.Enabled = False 'JUEZ 20150310
    Else
        fnMontoMinimoRCD = lrPar!nMontoMin
        fnTipCambio = lrPar!nCambioFijo
    End If
Set lrPar = Nothing
Me.lblfecha.Caption = gdFecDataFM
Me.Icon = LoadPicture(App.Path & gsRutaIcono)

Set oRCD = New COMDCredito.DCOMRCD
    lsCtaNoPref = oRCD.CargarCtaNoPref(fsServConsol)
Set oRCD = Nothing

End Sub

Private Sub cmdConsolidaRCD_Click()
Dim loRCDproc As COMNCredito.NCOMRCD
Set loRCDproc = New COMNCredito.NCOMRCD
Me.Enabled = False
ActivaConfigFormBarra True 'NAGL 20200718
'Crea Tablas RCD del mes
'Call loRCDproc.nCreaTablasRCD(Format(gdFecDataFM, "yyyymm"), fsServConsol) JIPR20210408 AGREGA PARÁMETRO RCDH CAMBIO TIPO DE DATO 1 RCDTvc,0  RCDvc , 2  RCDHvc
Call loRCDproc.nCreaTablasRCD(Format(gdFecDataFM, "yyyymm"), fsServConsol, 0)
barra.value = 4 'NAGL 20200718

'Obtiene Datos para procesar el RCD
'Call loRCDproc.ObtenerDatosParaGenerarRCDvc_Tvc(Format(gdFecDataFM, "yyyymm"), fsServConsol, fnTipCambio, fnMontoMinimoRCD) 'Agregado by NAGL 20200703
 Call loRCDproc.ObtenerDatosParaGenerarRCDvc_Tvc(Format(gdFecDataFM, "yyyymm"), fsServConsol, fnTipCambio, fnMontoMinimoRCD, 0) 'Agregado by NAGL 20200703
'JIPR20210408 AGREGA PARÁMETRO RCDH CAMBIO TIPO DE DATO 1 RCDTvc,0  RCDvc , 2  RCDHvc

barra.value = 6 'NAGL 20200718

'Procesar Datos en tablas RCD
'Call GeneraDatosRCD(Format(gdFecDataFM, "yyyymm"), loRCDproc) JIPR20210408 AGREGA PARÁMETRO RCDH CAMBIO TIPO DE DATO 1 RCDTvc,0  RCDvc , 2  RCDHvc
Call GeneraDatosRCD(Format(gdFecDataFM, "yyyymm"), loRCDproc, 0)

Set loRCDproc = Nothing
Me.lblAvance = ""
Me.lblDescripcion = ""
Me.Enabled = True
If lsObserv = "" Then
    barra.value = 8
    MsgBox "El Proceso de RCD ha culminado", vbInformation, "Aviso"
End If 'NAGL 20200718
lsObserv = "" 'NAGL 20200718
'ActivaConfigFormBarra False 'NAGL 20200718
'JIPR20210408 AGREGA PARÁMETRO RCDH CAMBIO TIPO DE DATO 1 RCDTvc,0  RCDvc , 2  RCDHvc
ActivaConfigFormBarra True 'NAGL 20200718
End Sub

'*** Genera los Datos para el RCD *****************
Private Sub GeneraDatosRCD(ByVal lsFecha As String, _
                           ByVal oRCD As COMNCredito.NCOMRCD, _
                           Optional ByVal pbRCDTransf As Integer)
                           'JIPR20210408 AGREGA PARÁMETRO RCDH CAMBIO TIPO DE DATO 1 RCDTvc,0  RCDvc , 2  RCDHvc
                           'Optional ByVal pbRCDTransf As Boolean = False)
                           'JUEZ 20150310 Se agregó pbRCDTransf
Dim sMensaje() As String
Dim sMensaje1() As String
Dim i As Integer
                            
'Call oRCD.GeneraDatosRCD(lsFecha, fsServConsol, gdFecDataFM, fnTipCambio, fsCodOfiInf, fsUbicaGeoRCD, fnMontoMinimoRCD, lsCtaNoPref, sMensaje, sMensaje1, pbRCDTransf) 'Comentado by NAGL 20200703
Call oRCD.GeneraDatosRCD_New(lsFecha, fsServConsol, fnTipCambio, sMensaje, pbRCDTransf) 'Agregado by NAGL 20200703

For i = 0 To UBound(sMensaje) - 1
    If sMensaje(i) <> "" Then
        MsgBox sMensaje(i), vbInformation, "Mensaje"
        lsObserv = "Obs" 'NAGL 202007
    End If 'NAGL 20200703
Next
'For i = 0 To UBound(sMensaje1) - 1
    'MsgBox sMensaje1(i), vbInformation, "Mensaje"
'Next 'Comentado by NAGL 20200703
End Sub
'**************************************************

Private Sub cmdConsolidaRCA_Click()
Dim loRCDproc As COMNCredito.NCOMRCD
Set loRCDproc = New COMNCredito.NCOMRCD

If loRCDproc.ValidaExisteRCD_RCA(Format(gdFecDataFM, "yyyymm"), fsServConsol) = True Then
    Me.Enabled = False
    ActivaConfigFormBarra True 'NAGL 20200718
    'Crea Tablas RCA del mes
    Call loRCDproc.nCreaTablasRCA(Format(gdFecDataFM, "yyyymm"), fsServConsol)
    barra.value = 4 'NAGL 20200718
    
    'Obtiene Datos para procesar el RCA
    Call loRCDproc.ObtenerDatosParaGenerarRCA(Format(gdFecDataFM, "yyyymmdd")) 'Agregado by NAGL 20200703
    barra.value = 6 'NAGL 20200718
    
    'Procesar Datos en tablas RCA
    Call GeneraDatosRCA(Format(gdFecDataFM, "yyyymm"), loRCDproc)
    
    Set loRCDproc = Nothing
    Me.lblAvance = ""
    Me.lblDescripcion = ""
    Me.Enabled = True
    If lsObserv = "" Then
        barra.value = 8
        MsgBox "El Proceso de RCA ha culminado", vbInformation, "Aviso"
    End If 'NAGL 20200718
    lsObserv = "" 'NAGL 20200718
    'ActivaConfigFormBarra False 'NAGL 20200718
    'JIPR20210408 AGREGA PARÁMETRO RCDH CAMBIO TIPO DE DATO 1 RCDTvc,0  RCDvc , 2  RCDHvc
    ActivaConfigFormBarra True
Else
    MsgBox "Por favor, asegúrese primero de generar la Data RCD...!! ", vbInformation, "Atención"
End If 'NAGL 20200718
End Sub

'*** Genera los Datos para el RCA *****************
Private Sub GeneraDatosRCA(ByVal lsFecha As String, _
                           ByVal oRCD As COMNCredito.NCOMRCD)
Dim sMensaje() As String
Dim sMensaje1() As String
Dim i As Integer
                            
'Call oRCD.GeneraDatosRCA(lsFecha, fsServConsol, gdFecDataFM, fnTipCambio, fsCodOfiInf, fsUbicaGeoRCD, fnMontoMinimoRCD, lsCtaNoPref, sMensaje, sMensaje1)'Comentado by NAGL 20200710
Call oRCD.GeneraDatosRCA_New(lsFecha, fsServConsol, fnTipCambio, sMensaje) 'Agregado by NAGL 20200710

For i = 0 To UBound(sMensaje) - 1
    If sMensaje(i) <> "" Then
        MsgBox sMensaje(i), vbInformation, "Mensaje"
        lsObserv = "Obs" 'NAGL 202007
    End If 'NAGL 20200703
Next
'For i = 0 To UBound(sMensaje1) - 1
    'MsgBox sMensaje1(i), vbInformation, "Mensaje"
'Next'Comentado by NAGL 20200703
End Sub
'**************************************************

Private Sub cmdConsolidaGeneraSecuencia_Click()
'JUEZ 20130510 *****
'GeneraSecuencia
'EliminaSecuenciaSinDatos
'GeneraSecuencia
'END JUEZ ****Comentado by NAGL 20200716
'***Agregado by NAGL 20200716***'
Dim oRCD As COMNCredito.NCOMRCD
Set oRCD = New COMNCredito.NCOMRCD
'If oRCD.ValidaExisteRCD_RCA(Format(gdFecDataFM, "yyyymm"), fsServConsol, False, "RCA") = True Then
'JIPR20210408 AGREGA PARÁMETRO RCDH CAMBIO TIPO DE DATO 1 RCDTvc,0  RCDvc , 2  RCDHvc
If oRCD.ValidaExisteRCD_RCA(Format(gdFecDataFM, "yyyymm"), fsServConsol, 0, "RCA") = True Then
    ActivaConfigFormBarra True
    barra.value = 5
    'Call oRCD.GeneraSecuenciaRCD_RCA(Format(gdFecDataFM, "yyyymm"), fsServConsol)
    'JIPR20210408 AGREGA PARÁMETRO RCDH CAMBIO TIPO DE DATO 1 RCDTvc,0  RCDvc , 2  RCDHvc
    Call oRCD.GeneraSecuenciaRCD_RCA(Format(gdFecDataFM, "yyyymm"), fsServConsol, 0)
    barra.value = 8
    MsgBox "Se genero la Secuencia", vbInformation, "Aviso"
    'ActivaConfigFormBarra False
    'JIPR20210408 AGREGA PARÁMETRO RCDH CAMBIO TIPO DE DATO 1 RCDTvc,0  RCDvc , 2  RCDHvc
    ActivaConfigFormBarra True
Else
    MsgBox "Por favor, asegúrese primero de generar la Data RCD y RCA...!! ", vbInformation, "Atención"
End If
'*******END NAGL 20200716*******'
End Sub

Private Sub cmdConsolidaRCDT_Click()
Dim loRCDproc As COMNCredito.NCOMRCD
Set loRCDproc = New COMNCredito.NCOMRCD
Me.Enabled = False
ActivaConfigFormBarra True 'NAGL 20200718
'Crea Tablas RCDT del mes
'Call loRCDproc.nCreaTablasRCD(Format(gdFecDataFM, "yyyymm"), fsServConsol, True)
'JIPR20210408 AGREGA PARÁMETRO RCDH CAMBIO TIPO DE DATO 1 RCDTvc,0  RCDvc , 2  RCDHvc
Call loRCDproc.nCreaTablasRCD(Format(gdFecDataFM, "yyyymm"), fsServConsol, 1)
barra.value = 4 'NAGL 20200718

'Obtiene Datos para procesar el RCDT
'Call loRCDproc.ObtenerDatosParaGenerarRCDvc_Tvc(Format(gdFecDataFM, "yyyymm"), fsServConsol, fnTipCambio, fnMontoMinimoRCD, True) 'Agregado by NAGL 20200717
'JIPR20210408 AGREGA PARÁMETRO RCDH CAMBIO TIPO DE DATO 1 RCDTvc,0  RCDvc , 2  RCDHvc
Call loRCDproc.ObtenerDatosParaGenerarRCDvc_Tvc(Format(gdFecDataFM, "yyyymm"), fsServConsol, fnTipCambio, fnMontoMinimoRCD, 1)
barra.value = 5 'NAGL 20200718

'Procesar Datos en tablas RCDT
'Call GeneraDatosRCD(Format(gdFecDataFM, "yyyymm"), loRCDproc, True)
'JIPR20210408 AGREGA PARÁMETRO RCDH CAMBIO TIPO DE DATO 1 RCDTvc,0  RCDvc , 2  RCDHvc
Call GeneraDatosRCD(Format(gdFecDataFM, "yyyymm"), loRCDproc, 1)
barra.value = 7 'NAGL 20200718

'GeneraSecuencia True
'EliminaSecuenciaSinDatos True
'GeneraSecuencia True 'Comentado by NAGL 20200717
Me.lblAvance = ""
Me.lblDescripcion = ""
Me.Enabled = True
If lsObserv = "" Then
    'Call loRCDproc.GeneraSecuenciaRCD_RCA(Format(gdFecDataFM, "yyyymm"), fsServConsol, True) 'Agregado by NAGL 20200717
    'JIPR20210408 AGREGA PARÁMETRO RCDH CAMBIO TIPO DE DATO 1 RCDTvc,0  RCDvc , 2  RCDHvc
    Call loRCDproc.GeneraSecuenciaRCD_RCA(Format(gdFecDataFM, "yyyymm"), fsServConsol, 1) 'Agregado by NAGL 20200717
    Set loRCDproc = Nothing
    barra.value = 8
    'MsgBox "El Proceso de RCD Transferencia ha culminado", vbInformation, "Aviso"
    'JIPR20210408 AGREGA PARÁMETRO RCDH CAMBIO TIPO DE DATO 1 RCDTvc,0  RCDvc , 2  RCDHvc
     MsgBox "El Proceso de RCDT Transferencia ha culminado", vbInformation, "Aviso"
End If 'NAGL 20200718
lsObserv = "" 'NAGL 20200718
'ActivaConfigFormBarra False 'NAGL 20200718
'JIPR20210408 AGREGA PARÁMETRO RCDH CAMBIO TIPO DE DATO 1 RCDTvc,0  RCDvc , 2  RCDHvc
ActivaConfigFormBarra True
End Sub

Private Sub ActivaConfigFormBarra(psBarra As Boolean)
If psBarra = True Then
    'Me.Height = 4035
    'JIPR20210408 AGREGA PARÁMETRO RCDH CAMBIO TIPO DE DATO 1 RCDTvc,0  RCDvc , 2  RCDHvc
    Me.Height = 4275
    barra.Min = 0
    barra.Max = 8
    barra.value = 2
Else
    Me.Height = 3690
End If
End Sub 'NAGL 20200718

Private Sub cmdSalir_Click()
    Unload Me
End Sub



'************Boton cmdConsolidaRCM en Desuso*********
Private Sub cmdConsolidaRCM_Click()
Dim loRCDproc As COMNCredito.NCOMRCD
Me.Enabled = False
Set loRCDproc = New COMNCredito.NCOMRCD
    ' Crea Tablas RCM del mes
    Call loRCDproc.nCreaTablasRCM(Format(gdFecDataFM, "yyyymm"), fsServConsol)
    ' Llena Datos en tablas RCM
    Call GeneraDatosRCM(Format(gdFecDataFM, "yyyymm"), loRCDproc)
Set loRCDproc = Nothing
Me.lblAvance = ""
Me.lblDescripcion = ""
Me.Enabled = True
MsgBox "El Proceso de RCM ha culminado", vbInformation, "Aviso"
End Sub
'*** Genera los Datos para el RCM *******
Private Sub GeneraDatosRCM(ByVal lsFecha As String, _
                           ByVal oRCD As COMNCredito.NCOMRCD)
Dim sMensaje() As String
Dim sMensaje1() As String
Dim i As Integer
Call oRCD.GeneraDatosRCM(lsFecha, fsServConsol, gdFecDataFM, fnTipCambio, fsCodOfiInf, fsUbicaGeoRCD, fnMontoMinimoRCD, lsCtaNoPref, sMensaje, sMensaje1)
For i = 0 To UBound(sMensaje) - 1
    MsgBox sMensaje(i), vbInformation, "Mensaje"
Next
For i = 0 To UBound(sMensaje1) - 1
    MsgBox sMensaje1(i), vbInformation, "Mensaje"
Next
End Sub
'****************************************************

'***BEGIN Comentado by NAGL 20200716*****'
''JUEZ 20130510 ********************************************************************
'Private Sub GeneraSecuencia(Optional ByVal pbRCDTransf As Boolean = False) 'JUEZ 20150310 Se agregó pbRCDTransf
'Dim SQL1 As String
'Dim oconecta As COMConecta.DCOMConecta
'Dim lsNomTabla As String 'JUEZ 20150310
''sql1 = "CREATE  INDEX cNomPers ON dbo.RCDvc" & Format(gdFecSis, "yyyymm") & "01 ([cApePat], [cApeMat], [cApeCasada], [cNomPri])"
''dbCmact.Execute sql1
'Set oconecta = New COMConecta.DCOMConecta
'oconecta.AbreConexion
'oconecta.CommadTimeOut = 2000
'lsNomTabla = IIf(pbRCDTransf, "RCDTvc", "RCDvc") 'JUEZ 20150310
''Crea temporal
'SQL1 = "create table #tmpRCDOrden (codigo varchar(13), orden int identity(1,1) not null, secuencia varchar(8) ) "
'oconecta.Ejecutar SQL1
''Indexa
'SQL1 = "create index cod on codigo "
''oConecta.Ejecutar SQL1
''Inserta en la Temporal
'SQL1 = " insert into  #tmpRCDOrden (codigo) " _
     '& " select cPersCod   from " & fsServConsol & lsNomTabla & Format(gdFecData, "yyyymm") & "01 order by cApePat, cApeMat, cApeCasada, cNombre1, cNombre2 "
''oconecta.Ejecutar SQL1
''Actualiza la Secuencia
'SQL1 = " update #tmpRCDOrden set secuencia = replicate('0', 8-len(orden))+ ltrim(str(orden)) "
'oconecta.Ejecutar SQL1
''*** Actualiza RCDvc..01
'SQL1 = " update x set x.cnumsec = t.Secuencia " _
     '& " from " & fsServConsol & lsNomTabla & Format(gdFecData, "yyyymm") & "01 x " _
     '& " join #tmpRCDOrden t on " _
     '& " x.cPersCod = t.Codigo COLLATE SQL_Latin1_General_CP1_CI_AS "
'oconecta.Ejecutar SQL1
'*** Actualiza RCDvc..02
'SQL1 = " update x set x.cnumsec = t.Secuencia " _
     '& " from " & fsServConsol & lsNomTabla & Format(gdFecData, "yyyymm") & "02 x " _
     '& " join #tmpRCDOrden t on " _
     '& " x.cPersCod = t.Codigo COLLATE SQL_Latin1_General_CP1_CI_AS "
'oconecta.Ejecutar SQL1
'Elimina la temporal
'SQL1 = "DROP table #tmpRCDOrden  "
'oconecta.Ejecutar SQL1
'--- Actualiza RCA (Avales)
'If Not pbRCDTransf Then
  '** LUCV20170410, Comentó y Agregó Según Observacion SBS
    'SQL1 = "update " & fsServConsol & "RCAvc" & Format(gdFecData, "yyyymm") & "01  Set cNumSec = d.Secuencia " _
    '     & "From " & fsServConsol & "RCAvc" & Format(gdFecData, "yyyymm") & "01 m join " _
    '     & " (Select cPersCod, cNumSec Secuencia from " & fsServConsol & "RCDvc" & Format(gdFecData, "yyyymm") & "01 ) d " _
    '     & "ON m.cTitular = d.cPersCod COLLATE SQL_Latin1_General_CP1_CI_AS "
       
     'SQL1 = " UPDATE " & fsServConsol & "RCAvc" & Format(gdFecData, "yyyymm") & "01 " _
           '& " Set cNumSec = Rcd.Secuencia " _
           '& " FROM " & fsServConsol & "RCAvc" & Format(gdFecData, "yyyymm") & "01 RCA " _
           '& " INNER JOIN  (SELECT cPersCod, cNumSec Secuencia FROM " & fsServConsol & "RCDvc" & Format(gdFecData, "yyyymm") & "01) RCD ON RCA.cTitular = RCD.cPersCod " _
           '& " AND RCA.cPersCod <> RCA.cTitular COLLATE SQL_Latin1_General_CP1_CI_AS "
  ''Fin LUCV20170410
   'oconecta.Ejecutar SQL1
'End If
'Set oconecta = Nothing
'End Sub
'Private Sub EliminaSecuenciaSinDatos(Optional ByVal pbRCDTransf As Boolean = False) 'JUEZ 20150310 Se agregó pbRCDTransf
'Dim SQL1 As String
'Dim oconecta As COMConecta.DCOMConecta
'Set oconecta = New COMConecta.DCOMConecta
'oconecta.AbreConexion
'oconecta.CommadTimeOut = 2000
'SQL1 = "DELETE FROM " & fsServConsol & IIf(pbRCDTransf, "RCDTvc", "RCDVc") & Format(gdFecData, "yyyymm") & "01 " _
     '& "WHERE cNumSec NOT IN (Select cNumSec " _
     '& "                      From DBConsolidada.." & IIf(pbRCDTransf, "RCDTvc", "RCDVc") & Format(gdFecData, "yyyymm") & "02) "
'oconecta.Ejecutar SQL1
'Set oconecta = Nothing
'End Sub
'END JUEZ *************************************************************************
'***END Comentado by NAGL 20200716*****'

'***BEGIN Comentado by NAGL 20200716*****'
''** Devuelve el Monto de la Garantia en la Moneda
''**************************************
'Function DevGarantiaMoneda(ByVal pMoneda As String, ByVal pGarantSol As Currency, ByVal pGarantDol As Currency) As Currency
'Dim lnValor As Currency
''If pMoneda = "1" Then  ' Soles
''    lnValor = pGarantSol + (pGarantDol * gnTipoCambio)
''Else  ' Dolares
''    If pGarantSol = 0 Then
''        lnValor = pGarantDol
''    Else
''        lnValor = (pGarantSol / gnTipoCambio) + pGarantDol
''    End If
''End If
''DevGarantiaMoneda = CCur(Format(lnValor, "#0.00"))
'End Function
'
'Private Sub fCargaDatosJuridicos(ByVal psCodPers As String, psCodRegPub As String, psMagEmp As String, psSiglas As String)
'Dim lsSQL As String
'Dim rsJ As New ADODB.Recordset
'Dim lbEncuentroMaestroICC As Boolean
''lsSQL = "SELECT cCodPers, cCodRegPub, cMagEmp, cSigla FROM PersonaJur " _
''      & "WHERE cCodPers='" & Trim(psCodPers) & "'"
''Set rsJ = CargaRecord(lsSQL)
''If Not RSVacio(rsJ) Then
''    psCodRegPub = IIf(IsNull(rsJ!ccodregpub), "", rsJ!ccodregpub)
''    psMagEmp = IIf(IsNull(rsJ!cMagEmp), "", rsJ!cMagEmp)
''    psSiglas = IIf(IsNull(rsJ!csigla), "", rsJ!csigla)
''End If
''rsJ.Close
''Set rsJ = Nothing
'End Sub
'
'Public Function fBuscaMagnitudEmpresarial(ByVal psNumFuente As String) As String
'Dim lsSQL As String
'Dim rs As New ADODB.Recordset
'Dim rs1 As New ADODB.Recordset
'Dim lnTotal As Currency
'Dim lnSaldo As Currency
'Dim lnValorUIT As Currency
'Dim lsMagnitudEmpresarial As String
'fBuscaMagnitudEmpresarial = lsMagnitudEmpresarial
'End Function
'
'Private Function fgReemplazaCaracterEspecial(ByVal psNom As String) As String
'Dim lsNombrePers As String
'
'            lsNombrePers = CadDerecha(Trim(Replace(psNom, "-", "", , , vbTextCompare)), 80)
'            lsNombrePers = CadDerecha(Trim(Replace(psNom, ".", " ", , , vbTextCompare)), 80)
'            lsNombrePers = CadDerecha(Trim(Replace(psNom, "Ñ", "#", , , vbTextCompare)), 80)
'            lsNombrePers = CadDerecha(Trim(Replace(psNom, "ñ", "#", , , vbTextCompare)), 80)
'fgReemplazaCaracterEspecial = lsNombrePers
'End Function
'
'Private Function VerApostrofe(ByVal psCadena As String) As String
'    VerApostrofe = Replace(psCadena, "'", "''", , , vbTextCompare)
'End Function
'
''Separa el Nombre : ApePaterno, ApeMaterno, ApeConyugue, PrimerNombre, SegundNombre
'Public Function fgSeparaNombre(ByVal psNombre As String, ByRef lsApellido As String, ByRef lsMaterno As String, _
'        ByRef lsConyugue As String, ByRef lsNomPri As String, ByRef lsNomSeg As String, ByVal psTipPers As String)
'
'Dim lsSQL As String
'Dim lrs As New ADODB.Recordset
'
'Dim Total As Long
'Dim Pos As Long
'Dim CadAux As String
'Dim lsNombre As String
'Dim CadAux2 As String
'Dim posAux As Integer
'Dim lbVda As Boolean
'
'    lsApellido = "": lsMaterno = "": lsNombre = "": lsConyugue = "": lsNomPri = "": lsNomSeg = ""
'    lbVda = False
'    Total = Len(Trim(psNombre))
'    Pos = InStr(psNombre, "/")
'    If Pos <> 0 Then
'        lsApellido = Left(psNombre, Pos - 1)
'        'LAYG
'        If Len(lsApellido) = 0 Then
'            lsApellido = "XXXX"
'        End If
'        CadAux = Mid(psNombre, Pos + 1, Total)
'        Pos = InStr(CadAux, "\")
'        If Pos <> 0 Then
'            lsMaterno = Left(CadAux, Pos - 1)
'            lsMaterno = Replace(lsMaterno, "-", "")
'            CadAux = Mid(CadAux, Pos + 1, Total)
'            Pos = InStr(CadAux, ",")
'            If Pos > 0 Then
'                CadAux2 = Left(CadAux, Pos - 1)
'                posAux = InStr(CadAux, "VDA")
'                If posAux = 0 Then
'                    lsConyugue = CadAux2
'                Else
'                    lbVda = True
'                    lsConyugue = CadAux2
'                    lsConyugue = Replace(CadAux2, "VDA DE ", "")
'                End If
'            Else
'                lsMaterno = CadAux
'            End If
'        Else
'            CadAux = Mid(CadAux, Pos + 1, Total)
'            Pos = InStr(CadAux, ",")
'            If Pos <> 0 Then
'                lsMaterno = Left(CadAux, Pos - 1)
'                lsConyugue = ""
'            Else
'                lsMaterno = CadAux
'            End If
'        End If
'        lsNombre = Mid(CadAux, Pos + 1, Total)
'        'Nombre1 // Nombre2
'        Pos = InStr(lsNombre, " ")
'        If Pos = 0 Then
'            lsNomPri = lsNombre
'        Else
'            lsNomPri = Left(lsNombre, Pos - 1)
'            lsNomSeg = Mid(lsNombre, Pos + 1, Total)
'        End If
'
'        If Len(Trim(lsConyugue)) > 0 Then
'                If lbVda = True Then
'                    'PstaNombre = Trim(lsConyugue) & " " & Trim(lsNombre) & " " & Trim(lsApellido) & IIf(Len(Trim(lsMaterno)) = 0, " VDA DE", " " & Trim(lsMaterno) & "VDA DE")
'                Else
'                    'PstaNombre = Trim(lsConyugue) & " " & Trim(lsNombre) & " " & Trim(lsApellido) & IIf(Len(Trim(lsMaterno)) = 0, " DE", " " & Trim(lsMaterno) & " DE")
'                End If
'        Else
'            'PstaNombre = Trim(lsApellido) & " " & Trim(lsMaterno) & " " & Trim(lsNombre)
'        End If
'    Else
'        'PstaNombre = Trim(psNombre)
'        lsApellido = Trim(psNombre)
'    End If
'    If lsConyugue = "" And lsMaterno = "" And psTipPers = "1" Then
'        lsMaterno = "XXXX"
'    End If
'    lsApellido = VerApostrofe(lsApellido)
'    lsMaterno = VerApostrofe(lsMaterno)
'    lsConyugue = VerApostrofe(lsConyugue)
'    lsNomPri = VerApostrofe(lsNomPri)
'    lsNomSeg = VerApostrofe(lsNomSeg)
'    'If Trim(lrs!NomRCD) <> Trim(lsApellido & " " & lsMaterno & " " & lsNomPri & " " & lsNomSeg) Then
'    '   ' MsgBox lrs!NomRCD & " - " & lsApellido & " " & lsMaterno & " " & lsNomPri & " " & lsNomSeg
'    'End If
'End Function
'***END Comentado by NAGL 20200716*****'
Private Sub lblfecha_Click()

End Sub
