VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmAsientoDN 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cajero : Operaciones - Asiento contable diario"
   ClientHeight    =   4530
   ClientLeft      =   2850
   ClientTop       =   3210
   ClientWidth     =   5970
   FillColor       =   &H00808000&
   Icon            =   "frmAsientoDN.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   5970
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   390
      Left            =   4680
      TabIndex        =   11
      Top             =   3990
      Width           =   1020
   End
   Begin VB.PictureBox prgList 
      Appearance      =   0  'Flat
      Height          =   225
      Left            =   120
      ScaleHeight     =   195
      ScaleWidth      =   4395
      TabIndex        =   12
      Top             =   4110
      Visible         =   0   'False
      Width           =   4425
   End
   Begin RichTextLib.RichTextBox rtfImp 
      Height          =   315
      Left            =   6045
      TabIndex        =   13
      Top             =   2910
      Visible         =   0   'False
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   556
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmAsientoDN.frx":030A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSMask.MaskEdBox txtFecha 
      Height          =   315
      Left            =   2805
      TabIndex        =   15
      Top             =   1020
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Frame fraContenedor 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3330
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   570
      Width           =   5730
      Begin VB.CheckBox chkSinCambioSBS 
         Caption         =   "Sin Cambio SBS"
         Height          =   255
         Left            =   4080
         TabIndex        =   19
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Frame fraContenedor 
         Appearance      =   0  'Flat
         Caption         =   "Moneda "
         ForeColor       =   &H80000008&
         Height          =   1005
         Index           =   2
         Left            =   4110
         TabIndex        =   8
         Top             =   195
         Width           =   1380
         Begin VB.CheckBox chkMonDol 
            Appearance      =   0  'Flat
            Caption         =   "&Dolares"
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   180
            TabIndex        =   10
            Top             =   570
            Value           =   1  'Checked
            Width           =   870
         End
         Begin VB.CheckBox chkMonSol 
            Appearance      =   0  'Flat
            Caption         =   "S&oles"
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   180
            TabIndex        =   9
            Top             =   255
            Value           =   1  'Checked
            Width           =   825
         End
      End
      Begin VB.Frame fraContenedor 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   3015
         Index           =   1
         Left            =   195
         TabIndex        =   4
         Top             =   195
         Width           =   3750
         Begin VB.CommandButton cmdValidar 
            Caption         =   "&Validar"
            Height          =   330
            Left            =   120
            TabIndex        =   21
            Top             =   1560
            Width           =   3480
         End
         Begin VB.CommandButton cmdAsientos 
            Caption         =   "&Generar Asientos Contables"
            Height          =   330
            Left            =   120
            TabIndex        =   20
            Top             =   2040
            Width           =   3480
         End
         Begin VB.CommandButton cmdImprimir 
            Caption         =   "&Imprimir Asiento Contable"
            Height          =   330
            Left            =   120
            TabIndex        =   18
            Top             =   2520
            Width           =   3480
         End
         Begin VB.CheckBox chkPig 
            Appearance      =   0  'Flat
            Caption         =   "&Pignoraticio"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   2520
            TabIndex        =   7
            Top             =   975
            Value           =   1  'Checked
            Width           =   1155
         End
         Begin VB.CheckBox chkCre 
            Appearance      =   0  'Flat
            Caption         =   "C&rédito"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1365
            TabIndex        =   6
            Top             =   975
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.CheckBox chkAho 
            Appearance      =   0  'Flat
            Caption         =   "&Ahorros"
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   180
            TabIndex        =   5
            Top             =   975
            Value           =   1  'Checked
            Width           =   885
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00000080&
            BorderWidth     =   2
            X1              =   105
            X2              =   3615
            Y1              =   795
            Y2              =   795
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Elija el asiento a Generar "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   150
            TabIndex        =   17
            Top             =   285
            Width           =   2145
         End
      End
      Begin VB.Frame fraImpresion 
         Appearance      =   0  'Flat
         Caption         =   "Impresión"
         ForeColor       =   &H80000008&
         Height          =   945
         Left            =   4110
         TabIndex        =   1
         Top             =   1245
         Width           =   1380
         Begin VB.OptionButton optImpresion 
            Appearance      =   0  'Flat
            Caption         =   "Impresora"
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   1
            Left            =   120
            TabIndex        =   3
            Top             =   525
            Width           =   990
         End
         Begin VB.OptionButton optImpresion 
            Appearance      =   0  'Flat
            Caption         =   "Pantalla"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   2
            Top             =   255
            Value           =   -1  'True
            Width           =   960
         End
      End
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "Fecha :"
      Height          =   225
      Left            =   945
      TabIndex        =   16
      Top             =   750
      Width           =   585
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PLAN CONTABLE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   465
      Left            =   90
      TabIndex        =   14
      Top             =   90
      Width           =   5760
   End
End
Attribute VB_Name = "frmAsientoDN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim pAsiDia As Boolean
Dim pPrevioMax As Double
Dim pLineasMax As Double
Dim pHojaFiMax As Integer
Dim ssql As String
Dim vRTFImp As String
Dim dHoraGrab As Date
Dim gdHoraGrab As String
Dim gbAsientoDN As Boolean

Dim lTransActiva As Boolean
Dim sObjetoCod  As String
Dim sIFTpo As String
Dim SCtaIfCod As String

Dim oCon As COMConecta.DCOMConecta
Dim TCF As Currency, TCC As Currency, TCV As Currency, TCCPonderado As Currency
Dim gAsientoProcesoCapAho As Long
Dim gAsientoProcesoCapCTS As Long
Dim gAsientoProcesoNoAsiento As Long
Dim gAsientoProcesoSiAsiento As Long
Dim gAsientoProcesoNoAsientoCab3 As Long
Dim gAsientoProcesoNoAsientoCab4 As Long
Dim gAsientoProcesoSiAsientoCVME As Long
Dim gAsientoProcesoSiAsientoSofFal As Long
Dim gAsientoProcesoSiColocCargoCuenta As Long
Dim gAsientoProcesoCreditoSinAsiento As Long
Dim ldFechaAsiento As Date
Dim N1 As Integer ' PEAC 20141001, variable para realizar el Log del proceso de asnto cnt
Dim vObs As String
Dim vProduc As String

Dim nResult81 As Currency, nResult82 As Currency
Dim nResult85 As Currency, nResult86 As Currency
Dim vCta82D As Currency, vCta82H As Currency, vCta81D As Currency, vCta81H As Currency
Dim vCta83D As Currency, vCta83H As Currency, vCta84D As Currency, vCta84H As Currency, vDife8 As Double
Dim vCta86D As Currency, vCta86H As Currency, vCta85D As Currency, vCta85H As Currency
Dim vAgencia As String, vCodConta As String, vNoCtaCnt As String, vParche As String
Dim lsCVMEGanacia As String, lsCVMEPerdida As String, lsCajaSoles As String, lsCajaDolares As String
Dim lsCtaCtbPerdida As String, lsCtaCtbGanancia As String, lnMontoOperacion As Double
Dim lsCtaCtbPerdidaT As String, lsCtaCtbGananciaT As String
Dim lnMontoOperacionGP As Double
Dim lsCajaAgenciaSoles As String, lsCajaAgenciaDolares As String
Dim oAsi As COMDCajaGeneral.DCOMAsiento
Private Declare Function mciExecute Lib "winmm.dll" (ByVal lpstrCommand As String) As Long

Dim gsCtaCodFoncodes As String



'PEAC 20160701 - Procedimiento que ejecuta el comando con el Api mciExecute
'************************************************************
Private Sub ejecutar(comando As String)
    Dim cNombre As String
    
    cNombre = Dir$("c:\MusicaFinCierre.mp3")
    If cNombre <> "" Then
        mciExecute comando & "c:\" & cNombre
    End If
  
End Sub


'PEAC 20141001, Proceso para generar el log de Asiento contable
Private Sub GeneraLogAsiento(ByVal psTexto As String, Optional ByVal pnFecHor As Integer = 1)
    N1 = FreeFile()
    Open "C:\install\LOGASIENTO.TXT" For Append As #N1
    If pnFecHor <> 1 Then
        Print #N1, psTexto
    Else
        Print #N1, psTexto; Date; Time
    End If
    Close #1
End Sub

Private Sub GeneraRepoAsntoCnt(ByVal psTexto As String, ByVal psNomArchivo As String)
    N1 = FreeFile()
        
    Dim cNomArchivo As String
    cNomArchivo = "AC"
    
    'Open "C:\install\LOGASIENTO.TXT" For Append As #N1
    'Open "C:\install\prueba.TXT" For Output As #N1
    Open "C:\Install\Spooler\" & psNomArchivo For Output As #N1
    
    Print #N1, psTexto

    Close #1
End Sub

Private Sub CorreVariablesGlobales()
    gAsientoProcesoCapAho = 1
    gAsientoProcesoCapCTS = 2
    gAsientoProcesoNoAsiento = 3
    gAsientoProcesoSiAsiento = 4
    gAsientoProcesoNoAsientoCab3 = 5
    gAsientoProcesoNoAsientoCab4 = 6
    gAsientoProcesoSiAsientoCVME = 7
    gAsientoProcesoSiAsientoSofFal = 8
    gAsientoProcesoSiColocCargoCuenta = 9
    gAsientoProcesoCreditoSinAsiento = 10
End Sub
Private Sub IniciaVariables()

    lsCtaCtbPerdida = "41M80103AG"
    lsCtaCtbGanancia = "51M80103AG"
    lnMontoOperacion = 0
    '*****************************************************************

    nResult81 = 0: nResult82 = 0: nResult85 = 0: nResult86 = 0: vCta82D = 0: vCta82H = 0: vCta81D = 0: vCta81H = 0: vCta83D = 0: vCta83H = 0: vCta84D = 0: vCta84H = 0: vDife8 = 0:   vCta86D = 0: vCta86H = 0: vCta85D = 0: vCta85H = 0:    vAgencia = "": vCodConta = "": vNoCtaCnt = "": vParche = "":    lsCVMEGanacia = "": lsCVMEPerdida = "": lsCajaSoles = "": lsCajaDolares = ""
    '*****************************************************************

End Sub
'--PEAC 20160107
Private Sub CargaParametrosAsnto()
    Set oAsi = New COMDCajaGeneral.DCOMAsiento

    gsCtaCodFoncodes = oAsi.GetAsientoParametro(1)
    lsCVMEGanacia = oAsi.GetAsientoParametro(2)
    lsCVMEPerdida = oAsi.GetAsientoParametro(3)
    lsCajaSoles = oAsi.GetAsientoParametro(4)
    lsCajaDolares = oAsi.GetAsientoParametro(5)
    lsCajaAgenciaSoles = oAsi.GetAsientoParametro(6)
    lsCajaAgenciaDolares = oAsi.GetAsientoParametro(7)

End Sub

Private Sub BorraAsientoGenerado()
    Dim rsVer As ADODB.Recordset
    Dim lnRegAsi As Long
    lnRegAsi = -1
    Me.Caption = "Eliminando registros de asiento"
    Do While lnRegAsi <> 0
        ssql = "EXEC EliminaAsientoDB '" & Format(txtFecha, "yyyymmdd") & "'"
        oCon.ejecutar ssql

        'ssql = "SELECT ISNULL(COUNT(0),0) as nRegAsiento FROM ASIENTODN WHERE CONVERT(date,dFecha)='" & Format(txtFecha, "yyyymmdd") & "' "
        'Comentado by NAGL 202008
        ssql = "stp_sel_ConteoRegAsientoDN '" & Format(txtFecha, "yyyymmdd") & "'" 'NAGL 202008 Según Acta N°063-2020
        Set rsVer = oCon.CargaRecordSet(ssql)
        lnRegAsi = rsVer!nRegAsiento
        rsVer.Close
        Set rsVer = Nothing
    Loop
    Me.Caption = "Asientos Eliminados correctamente"
    'Consulta para obtener los precios del Oro
End Sub



Private Sub cmdAsientos_Click()


'para implementar
'    Dim lcPrtImp As String
'
'    If MsgBox("¿ Seguro desea generar Asiento Contable ?", vbQuestion + vbYesNo, "¡Confirmacion!") = vbNo Then Exit Sub
'
'    cmdAsientos.Enabled = False
'    txtFecha.Enabled = False
'    fraContenedor(1).Enabled = False
'    fraContenedor(2).Enabled = False
'    fraImpresion.Enabled = False
'
'    If chkMonSol.value = 0 And chkMonDol.value = 0 Then
'        'dbCmact.RollbackTrans ' Corta TRANSACCION
'        MsgBox " No se ha definido la Moneda ", vbInformation, " Aviso "
'        MousePointer = 0
'        Exit Sub
'    End If
'
'    Call GeneraAsientoDn(CDate(Me.txtFecha.Text), chkMonSol.value, chkMonDol.value, lcPrtImp)
'
''    ejecutar ("Play ") 'PEAC 20160701
'
'

'    'Activo Controles
'    cmdAsientos.Enabled = True
'    txtFecha.Enabled = True
'    fraContenedor(1).Enabled = True
'    fraContenedor(2).Enabled = True
'    fraImpresion.Enabled = True
'    'Fin Activos
'
'    Exit Sub
'
'ControlError:       ' Rutina de control de errores.
'        prgList.value = 0
'        'Activo Controles
'        cmdAsientos.Enabled = True
'        txtFecha.Enabled = True
'        fraContenedor(1).Enabled = True
'        fraContenedor(2).Enabled = True
'        fraImpresion.Enabled = True
'        'Fin Activos
'        MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
'           " Avise al Area de Sistemas ", vbInformation, " Aviso "
    
    
' fin implementar



    'On Error GoTo ControlError
    'Dim oAsi As COMDCajaGeneral.DCOMAsiento
    Dim lsPlantilla5114 As String 'ALPA20140303***********
    Dim lbPlantillaLogico As Boolean
    Dim lsNSIF As String 'DAOR 20100623, Proyecto Basilea II
    
    Dim cNomArchivoTxt As String
    cNomArchivoTxt = ""
    
    lsPlantilla5114 = "51M401TCCDNSIFSCAG" 'ALPA20140303***********
    '**DAOR 20081202, para operaciones por ATM***********************
    'Dim lsCtaCtbPerdida As String, lsCtaCtbGanancia As String, lnMontoOperacion As Double
    'Dim lsCtaCtbPerdidaT As String, lsCtaCtbGananciaT As String
    'Dim lnMontoOperacionGP As Double
    vRTFImp = ""
    Call IniciaVariables

'    lsCtaCtbPerdida = "41M80103AG"
'    lsCtaCtbGanancia = "51M80103AG"
'    lnMontoOperacion = 0
'    '*****************************************************************
'
'    nResult81 = 0: nResult82 = 0: nResult85 = 0: nResult86 = 0: vCta82D = 0: vCta82H = 0: vCta81D = 0: vCta81H = 0: vCta83D = 0: vCta83H = 0: vCta84D = 0: vCta84H = 0: vDife8 = 0:   vCta86D = 0: vCta86H = 0: vCta85D = 0: vCta85H = 0:    vAgencia = "": vCodConta = "": vNoCtaCnt = "": vParche = "":    lsCVMEGanacia = "": lsCVMEPerdida = "": lsCajaSoles = "": lsCajaDolares = ""
'    '*****************************************************************

    '*** PEAC 20191217
    'If MsgBox("¿ Seguro desea generar Asiento Contable ?", vbQuestion + vbYesNo, "¡Confirmacion!") = vbNo Then Exit Sub
    'If Date = "01/01/2020" Then
    'Do While Time() < "04:30:30 AM"
        'MsgBox "hora de ejecutar"
        'Exit Sub
    'Loop
    'End If



    Set oAsi = New COMDCajaGeneral.DCOMAsiento

    'Dim lsCVMEGanacia As String, lsCVMEPerdida As String, lsCajaSoles As String, lsCajaDolares As String
    'Dim lsCajaAgenciaSoles As String, lsCajaAgenciaDolares As String

    'Desactivo controles
    cmdAsientos.Enabled = False
    txtFecha.Enabled = False
    fraContenedor(1).Enabled = False
    fraContenedor(2).Enabled = False
    fraImpresion.Enabled = False
    'Fin Desactivo

    DoEvents

    Call CargaParametrosAsnto

'    gsCtaCodFoncodes = oAsi.GetAsientoParametro(1)
'    lsCVMEGanacia = oAsi.GetAsientoParametro(2)
'    lsCVMEPerdida = oAsi.GetAsientoParametro(3)
'    lsCajaSoles = oAsi.GetAsientoParametro(4)
'    lsCajaDolares = oAsi.GetAsientoParametro(5)
'    lsCajaAgenciaSoles = oAsi.GetAsientoParametro(6)
'    lsCajaAgenciaDolares = oAsi.GetAsientoParametro(7)

    Dim RegTran As ADODB.Recordset, RegOpeCta As ADODB.Recordset, RegTmp As ADODB.Recordset
    Dim tmpSql As String
    Dim vMoneda As Moneda
'    Dim vAgencia As String, vCodConta As String, vNoCtaCnt As String, vParche As String
    Dim vCont As Long, vCtaCnt As Long
    Dim vEst As String, vCad As String
    Dim x As Long
    Dim sTmp1 As String, sTmp3 As String
    Dim vespacio As Long, vLenNomb As Long
    Dim Arreglo() As Currency

    'ALPA 20120625**********************************
    Dim lnTipoCambioCompraL As Currency
    Dim lnTipoCambioVentaL As Currency
    Dim lnTipoCambioFijoL As Currency
    '***********************************************

    Call CorreVariablesGlobales

'    gAsientoProcesoCapAho = 1
'    gAsientoProcesoCapCTS = 2
'    gAsientoProcesoNoAsiento = 3
'    gAsientoProcesoSiAsiento = 4
'    gAsientoProcesoNoAsientoCab3 = 5
'    gAsientoProcesoNoAsientoCab4 = 6
'    gAsientoProcesoSiAsientoCVME = 7
'    gAsientoProcesoSiAsientoSofFal = 8
'    gAsientoProcesoSiColocCargoCuenta = 9
'    gAsientoProcesoCreditoSinAsiento = 10

    Set RegTran = New ADODB.Recordset
    Set RegTmp = New ADODB.Recordset
    Set RegOpeCta = New ADODB.Recordset
    Dim oPrevio As previo.clsprevio
    'Set oPrevio = New previo.clsprevio

    If pAsiDia Then
        ldFechaAsiento = gdFecSis
    Else
        ldFechaAsiento = CDate(Me.txtFecha.Text)
    End If

    '-----------------------------------------------------
    Call GeneraLogAsiento("Comienza a generar asiento contable del " & Format(ldFechaAsiento, "dd/MM/yyyy") & " - ")
    '-----------------------------------------------------

    ReDim Arreglo(1, 3)

    'Para la validacion Dia y hora de Generación
    'ARCV 14-03-2007
    'dHoraGrab = Format(gdFecSis & " " & Time, gsFormatoFechaHoraView)
    dHoraGrab = Format(Date & " " & Time, gsFormatoFechaHoraView)
    '-------
    'Carga fecha del Asiento y hora de grabación
    gdHoraGrab = Format(txtFecha & " " & Time, gsFormatoFechaHora)

    Dim vHH As Long, vMM As Long, vSS As Long, vSeg As Long
    Dim vTime As Variant
    vTime = Time

    Dim vCta28Debe As Currency, vCta28Haber As Currency
    Dim vCta19Debe As Currency, vCta19Haber As Currency, vCtaDolar As Currency
    '***********************************************************
    'PARTE 1
    Dim vMonAnt As String, vTipPer As String
    Dim pband28 As Boolean
    '***********************************************************
    'PARTE 2
    Dim vFondo As String, vPlazo As String, vTipoC As String, vRFA As String, vAG As String
    Dim vConceptoI As String 'JIPR 20200328
    
    Dim pBandFonc As Boolean
    Dim vCJ As String, vCR As String, vRE As String, vCE As String, vBanco As String
    Dim vTipoPD As String, vSC As String, vSB As String, vCD As String, lsCodAge As String
    Dim vTL As String
    '***********************************************************

    Dim vNS As String, vIF As String 'DAOR 20100623, Proyecto Basilea II

    MousePointer = 11
    vLenNomb = 70
    vespacio = vLenNomb + 54
    vRTFImp = ""

    'Insert registro en tabla para verificación del asiento
    ssql = "INSERT INTO AsientoValida (dAsientoFecha, cAsientoTipo, cAsientoEstado, cCodUsu, dAsientoModif)" & _
        " VALUES ('" & Format(txtFecha, gsFormatoFecha) & "','2','0','" & gsCodUser & "','" & Format(dHoraGrab, gsFormatoFechaHora) & "') "
    oCon.ejecutar ssql

    'Borra el asiento generado de este dia.
    'sSql = "DELETE AsientoDN Where convert(char(10),dfecha,112) = '" & Format(txtFecha, "yyyymmdd") & "'"

    Call BorraAsientoGenerado

'    Dim rsVer As ADODB.Recordset
'    Dim lnRegAsi As Long
'    lnRegAsi = -1
'    Me.Caption = "Eliminando registros de asiento"
'    Do While lnRegAsi <> 0
'        sSQL = "EXEC EliminaAsientoDB '" & Format(txtFecha, "mm/dd/yyyy") & "'"
'        oCon.ejecutar sSQL
'
'        sSQL = "SELECT ISNULL(COUNT(0),0) as nRegAsiento FROM ASIENTODN WHERE CONVERT(CHAR(10),dFecha,112)='" & Format(txtFecha, "yyyymmdd") & "' "
'        Set rsVer = oCon.CargaRecordSet(sSQL)
'        lnRegAsi = rsVer!nRegAsiento
'        rsVer.Close
'        Set rsVer = Nothing
'    Loop
'    Me.Caption = "Asientos Eliminados correctamente"
'    'Consulta para obtener los precios del Oro

    Dim oDGeneral As COMDConstSistema.NCOMTipoCambio
    Set oDGeneral = New COMDConstSistema.NCOMTipoCambio
    TCCPonderado = oDGeneral.EmiteTipoCambio(CDate(txtFecha.Text), TCPonderado)

    If Not GetTipCambio(CDate(txtFecha)) Then
        MsgBox " No se encuentran los Tipos de Cambio ", , " Aviso "
        MousePointer = 0
        Exit Sub
    Else
        With RegTmp
            TCV = gnTipCambioV
            TCC = gnTipCambioC
            TCF = gnTipCambio

            If TCF = 0 Then
                MsgBox " No se encuentra el tipo de cambio Fijo Diario", vbInformation, " Aviso "
                RegTmp.Close
                Set RegTmp = Nothing
                MousePointer = 0
                Exit Sub
            End If
        End With
    End If

    'Se carga TranDiaria para generar los asientos de acuerdo ha condiciones ingresadas
    'Para el Asiento
    'Se Agrego el Campo cCodAgeGir para generar el Asiento de Giros GITU 27/06/2008

    gAsientoProcesoNoAsiento = 3

    '**DAOR 20081202, se aumentó los campos MontoDispuesto y TipoCambioDispuesto
    '**para operaciones realizadas por cajeros automáticos

    sTmp1 = " exec stp_sel_ObtieneMovParaGenerarAsiento '" & Format(ldFechaAsiento, gsFormatoMovFecha) & "'"


    If chkMonSol.value = 0 And chkMonDol.value = 0 Then
        'dbCmact.RollbackTrans ' Corta TRANSACCION
        MsgBox " No se ha definido la Moneda ", vbInformation, " Aviso "
        MousePointer = 0
        Exit Sub
    End If

    Dim nPaso As Long
    vCta28Debe = 0: vCta28Haber = 0
    vCta19Debe = 0: vCta19Haber = 0
    vCont = 0: vCtaCnt = 0

    '**DAOR 20081204 **************************
    Call AsientoCapitlizacionAhorroCTS
    '******************************************
    '**PEAC 20170303 **************************
    Call AsientoOtrasOpesDeProcesoRapido
    '******************************************

    Dim lsRfa As String
    '**** Para la Generación del Asiento
    ssql = sTmp1
    RegTran.CursorLocation = adUseClient
    Set RegTran = oCon.CargaRecordSet(ssql)
    Set RegTran.ActiveConnection = Nothing
    If (RegTran.BOF Or RegTran.EOF) Then
        RegTran.Close
        Set RegTran = Nothing
    Else
'        prgList.Min = 0
'        prgList.Max = RegTran.RecordCount
'        prgList.Visible = True

        '' en este bucle se demora

        Do While Not RegTran.EOF

            lsRfa = IIf(IsNull(RegTran!cRFA), "", RegTran!cRFA)

            'DAOR 20081202, reemplazar a todos los RegTran!SumaMonto por lnMontoOperacion **
            lnMontoOperacion = RegTran!sumamonto
            '*******************************************************************************

            'If RegTran!cCtaCod = "109012321000377864" Or RegTran!cCtaCod = "109022321000108545" Then MsgBox "109012321000377864,109022321000108545"

            'Variables para cambio en el CodContable
            gdHoraGrab = Format$(RegTran!dFecTran, gsFormatoFechaHora)
            ssql = ""
            If Left(RegTran!cOpeCod, 4) = Left(gColPOpeCanceNorEnOtCjEFE, 4) Or _
                Left(RegTran!cOpeCod, 4) = Left(gServGiroCancEfec, 4) Or Left(RegTran!cOpeCod, 4) = Left(gITFGiroCancelEfect, 6) Then 'NAGL 20190306 Agregó Left(RegTran!cOpeCod, 4) = Left(gITFGiroCancelEfect, 6)
                vAgencia = Left(RegTran!cCtaCod, 2)
            Else
                If Len(RegTran!cCodAge) > 2 Then vAgencia = Mid(RegTran!cMovNro, 18, 2) Else vAgencia = Right(Trim(RegTran!cCodAge), 2)
            End If

            If Not IsNull(RegTran!cCtaCod) Then
                'Verifica Moneda de Cuentas Antiguas o Cuentas Nuevas
                If Len(Trim(RegTran!cCtaCod)) = 8 Then
                    vMonAnt = Mid(Trim(RegTran!cCtaCod), 2, 1)
                    If vMonAnt = "0" Then
                        vMoneda = Moneda.gMonedaNacional
                    ElseIf vMonAnt = Moneda.gMonedaNacional Then
                        vMoneda = Moneda.gMonedaExtranjera
                    Else
                        MsgBox "Moneda Antigua no reconocida " & RegTran!cCtaCod, , " Aviso "
                        'vNoCtaCnt = vNoCtaCnt & oImpresora.gPrnSaltoLinea & "     * " & RegTran!cCtaCod & " moneda antigua no reconocida."
                        vNoCtaCnt = vNoCtaCnt & vbCrLf & "     * " & RegTran!cCtaCod & " moneda antigua no reconocida."
                    End If
                Else
                    vMoneda = Mid(Trim(RegTran!cCtaCod), 9, 1)
                    If Not (vMoneda = Moneda.gMonedaNacional Or vMoneda = Moneda.gMonedaExtranjera) Then
                        'MsgBox " Código Errado, Moneda no definida " & RegTran!cCtaCod, vbInformation, " Aviso "
                        'vNoCtaCnt = vNoCtaCnt & oImpresora.gPrnSaltoLinea & "     * " & RegTran!cCtaCod & " Código Errado, Moneda no definida."
                        vNoCtaCnt = vNoCtaCnt & vbCrLf & "     * " & RegTran!cCtaCod & " Código Errado, Moneda no definida."

                    End If
                End If

                'If RegTran!cCtaCod = "109332321000026395" Then MsgBox RegTran!cCtaCod & "-" & RegTran!cOpeCod
'                    vTipPer = "0"
'                End If
' 53044065
'
'                If RegTran!nMovNro = "53044065" Then
'                    vTipPer = "0"
'                End If

                'En ahorros se busca el tipo de persona en ctas de Ahorro
                'If (Left(Trim(RegTran!cOpecod), 1) = "2" Or Trim(RegTran!cOpecod) = "990101" Or Trim(RegTran!cOpecod) = "990301") '***Modificado por ELRO el 20120213, según Acta N° 245-2011/TI-D
                'EJVG20120626 Se agregó la condición 990113 ITF Cargo Cta Ahorro x Pago Cuota EcoTaxi
                'RIRO20130401 Se agregaron operaciones 300111, 300120 y 990305 , segun proyecto de ahorros - Servicio de recaudo.
                'FRHU 20150203 ERS048-2014 NOTA DE CARGO(302001) Y NOTA DE ABONO(302002): Se agrego 302001 y 302002
                'VAPA 20170405 AGREGANDO OPERACIONES DE TRANSFERENCIA DE CCE 930026 930028 930100 990125 930031 990126 930016
                'NAGL AGREGADO EL COPECOD 300115 - COMISIÓN RECAUDO CC 30-05-2017
                'NAGL Agregó "3001[12]7" , "930[12]20" (PAGO DE SERVICIOS HB y CCE Banca por Internet) - 20191114
                'NAGL Agregó "102302" (PAGO MORATORIO CARGO CTA - BANCA POR INTERNET)- 20200811
                'APRI20200610 ADD "100949" POR COMISION REACTIVA
                If (Left(Trim(RegTran!cOpeCod), 1) = "2" Or Trim(RegTran!cOpeCod) = "990101" Or Trim(RegTran!cOpeCod) = "990301" Or Trim(RegTran!cOpeCod) = "990111" Or Trim(RegTran!cOpeCod) = "990113" _
                    Or Trim(RegTran!cOpeCod) = "300111" Or Trim(RegTran!cOpeCod) = "300120" Or Trim(RegTran!cOpeCod) = "300126" Or Trim(RegTran!cOpeCod) = "990305" Or Trim(RegTran!cOpeCod) = "100949" _
                    Or Trim(RegTran!cOpeCod) = "302001" Or Trim(RegTran!cOpeCod) = "302002" Or Trim(RegTran!cOpeCod) Like "3001[12]7" Or Trim(RegTran!cOpeCod) = "102202" _
                    Or Trim(RegTran!cOpeCod) = "930026" Or Trim(RegTran!cOpeCod) = "930016" Or Trim(RegTran!cOpeCod) = "930028" Or Trim(RegTran!cOpeCod) = "930100" Or Trim(RegTran!cOpeCod) = "990125" Or Trim(RegTran!cOpeCod) = "990306" _
                    Or Trim(RegTran!cOpeCod) = "930031" Or Trim(RegTran!cOpeCod) = "990126" Or Trim(RegTran!cOpeCod) = "300115" Or Trim(RegTran!cOpeCod) Like "10510[456789]" Or Trim(RegTran!cOpeCod) = "105120" Or Trim(RegTran!cOpeCod) Like "930[12]20" Or Trim(RegTran!cOpeCod) = "102302") _
                    And ExisTipPer(RegTran!cOpeCod, RegTran!Concepto, True, oCon.ConexionActiva) Then      'RERR: agregue las operaciones de ITF Cargo cuenta 'NAGL20190211 Agregó Trim(RegTran!cOpeCod) Like "10510[456789]" 'NAGL20190621 Agregó Trim(RegTran!cOpeCod) = "105120" 'NAGL 20190920 Agregó "990306"
                    vTipPer = ClienteTipoPers(RegTran!cCtaCod, oCon.ConexionActiva)
                    If Len(Trim(vTipPer)) = 0 Then
                        MsgBox " No existe Titular, Operac. Ahorros " & RegTran!cCtaCod, vbInformation, " Aviso "
                        'vNoCtaCnt = vNoCtaCnt & oImpresora.gPrnSaltoLinea & "     * " & RegTran!cCtaCod & " no existe Titular (Operac. Ahorros)."
                        vNoCtaCnt = vNoCtaCnt & vbCrLf & "     * " & RegTran!cCtaCod & " no existe Titular (Operac. Ahorros)."
                    End If
                    ssql = " AND nPersoneria = '" & vTipPer & "'"
                'En Crédito se busca Tipo de Persona en algunas Cuentas de Créditos
                'MARCA 01
                ElseIf Left(Trim(RegTran!cOpeCod), 4) = "0111" Or Left(Trim(RegTran!cOpeCod), 4) = "0118" Or _
                    Left(Trim(RegTran!cOpeCod), 4) = "0104" Or Left(Trim(RegTran!cOpeCod), 4) = "0132" Or _
                    Left(Trim(RegTran!cOpeCod), 4) = "0125" Or Left(Trim(RegTran!cOpeCod), 4) = "0139" Or _
                    Left(Trim(RegTran!cOpeCod), 4) = "0181" Or Left(Trim(RegTran!cOpeCod), 6) = "100202" Or Left(Trim(RegTran!cOpeCod), 6) = "100302" Or Left(Trim(RegTran!cOpeCod), 6) = "100402" Or Left(Trim(RegTran!cOpeCod), 6) = "100502" Or Left(Trim(RegTran!cOpeCod), 6) = "100602" Or Left(Trim(RegTran!cOpeCod), 6) = "100702" Or _
                    Trim(RegTran!cOpeCod) = "100102" Or Trim(RegTran!cOpeCod) = "100103" Or Trim(RegTran!cOpeCod) = "100302" Or _
                    Trim(RegTran!cOpeCod) = "100104" Or Trim(RegTran!cOpeCod) = "100104" Or _
                    Trim(RegTran!cOpeCod) = "100105" Or Trim(RegTran!cOpeCod) = "100106" Or _
                    Trim(RegTran!cOpeCod) = "100107" Or Trim(RegTran!cOpeCod) = "100108" Or _
                    Trim(RegTran!cOpeCod) = "100109" Or Left(Trim(RegTran!cOpeCod), 4) = "0188" Or Trim(RegTran!cOpeCod) = "100220" Then

                        If Left(Trim(RegTran!cOpeCod), 6) = "010440" Or Left(Trim(RegTran!cOpeCod), 4) = "0188" Or _
                            Left(Trim(RegTran!cOpeCod), 6) = "010105" Then
                            vTipPer = ClienteTipoPersCol(RegTran!nMovNro, oCon.ConexionActiva)
                        Else
                            If Left(RegTran!cOpeCod, 5) = "01043" Or Left(RegTran!cOpeCod, 5) = "01113" Or _
                            Left(RegTran!cOpeCod, 5) = "01183" Or Left(RegTran!cOpeCod, 5) = "01253" Or _
                            Left(RegTran!cOpeCod, 5) = "01393" Then
                                vTipPer = 0
                            Else
                                vTipPer = ClienteTipoPersCol(RegTran!nMovNro, oCon.ConexionActiva)
                            End If
                        End If
                        If Len(Trim(vTipPer)) = 0 Then
                            MsgBox " No existe Titular, Operac. Crédito " & RegTran!cCtaCod, vbInformation, " Aviso "
                            'vNoCtaCnt = vNoCtaCnt & oImpresora.gPrnSaltoLinea & "     * " & RegTran!cCtaCod & " no existe Titular (Operac. Créditos)."
                            vNoCtaCnt = vNoCtaCnt & vbCrLf & "     * " & RegTran!cCtaCod & " no existe Titular (Operac. Créditos)."

                        End If
                        ssql = " AND nPersoneria = " & vTipPer & " "
                ElseIf Left(Trim(RegTran!cOpeCod), 2) = Left(gColRecOpePasoARecup, 2) Then ' Arturo - Judicial
                    If RegTran!nSaldCnt = 1 Or RegTran!nSaldCnt = 2 Or RegTran!nSaldCnt = 3 Or RegTran!nSaldCnt = 4 Then
                        ssql = " AND nPersoneria = '" & RegTran!nSaldCnt & "'"
                    End If
                    '*** PEAC 20151201 se agrego las operaciones 300121 y 302002 porque duplican los montos en los asientos ya que no pedia personeria
                    '*** PEAC 20160217 - NOTA ABONO - PAGO DE UTILIDADES
                ElseIf RegTran!cOpeCod = "700104" Or RegTran!cOpeCod = "300121" Or (RegTran!cOpeCod = "302002" And (RegTran!Concepto = 209 Or RegTran!Concepto = 212 Or RegTran!Concepto = 207 Or RegTran!Concepto = 210)) Then
                    vTipPer = ClienteTipoPersCol(RegTran!nMovNro, oCon.ConexionActiva)
                    ssql = IIf(RegTran!cOpeCod = "302002" And (RegTran!Concepto = 207 Or RegTran!Concepto = 212), " AND (nPersoneria = 0 or nPersoneria = " & vTipPer & ") ", " AND nPersoneria = " & vTipPer & " ")

                End If
            Else
                If Not (Left(Trim(RegTran!cOpeCod), 2) = "23" Or Left(Trim(RegTran!cOpeCod), 2) = "24" Or Left(Trim(RegTran!cOpeCod), 2) = "25") Then
                    'MsgBox " No es un Código de Operac. Especial " & RegTran!cOpecod, vbInformation, " Aviso "
                    'vNoCtaCnt = vNoCtaCnt & oImpresora.gPrnSaltoLinea & "     * " & RegTran!cOpeCod & " no es un código de Operac. Especial (Cta.: " & RegTran!cCtaCod & ")."
                    vNoCtaCnt = vNoCtaCnt & vbCrLf & "     * " & RegTran!cOpeCod & " no es un código de Operac. Especial (Cta.: " & RegTran!cCtaCod & ")."

                End If


            End If
            'Para Asiento de Crédito, Foncodes pago en otra agencia
            pband28 = True

            'If RegTran!cOpeCod = "100505" Then Stop


            'Carga OPECUENTA para definir a que cuentas se ha de grabar
            'sSql = "SELECT cCtaContCod , cOpeCtaDH FROM OpeCtaNeg  " & _
                " WHERE cOpeCod = '" & RegTran!cOpecod & "' And nConcepto = " & RegTran!Concepto & " " & sSql   ' & " Order by cDebeHab "
             '*** PEAC 20161029
            ssql = "exec stp_sel_CargaOpeCuenta '" & RegTran!cOpeCod & "', " & RegTran!Concepto & "," & IIf(vTipPer = "", 0, vTipPer) & "," & IIf(Len(ssql) = 0, 0, 1)

            Set RegOpeCta = oCon.CargaRecordSet(ssql)
            If (RegOpeCta.BOF Or RegOpeCta.EOF) Then
                RegOpeCta.Close
                Set RegOpeCta = Nothing

                ' ESTE IF NO HACE NADA
'                If (Left(RegTran!cOpecod, 3) <> "107" And RegTran!cOpecod <> "120100") Then
'                    'MsgBox " Código de Operación No Reconocido " & RegTran!cOpeCod & "- Concepto " & RegTran!Concepto, vbInformation, " Aviso "
'                End If

                If RegTran!cOpeCod <> "120100" Then 'OBVIA LOS DEMAS CONCEPTOS DEL REGISTRO DE CREDITOS PIGNORATICIO
                    'vNoCtaCnt = vNoCtaCnt & oImpresora.gPrnSaltoLinea & "     * " & RegTran!cOpeCod & " operación no reconocida (Cta.: " & RegTran!cCtaCod & "- Concepto " & RegTran!Concepto & ")."
                    vNoCtaCnt = vNoCtaCnt & vbCrLf & "     * " & RegTran!cOpeCod & " operación no reconocida (Cta.: " & RegTran!cCtaCod & "- Concepto " & RegTran!Concepto & ")."

                End If
            Else

            '' este bucle demora

                Do While Not RegOpeCta.EOF
                    lbPlantillaLogico = False
                    'Mensaje de Cuenta
                    vCtaCnt = vCtaCnt + 1
                    If Len(Trim(RegTran!cCodAge)) > 2 Then vAgencia = Mid(RegTran!cMovNro, 18, 2) Else vAgencia = Right(Trim(RegTran!cCodAge), 2)
'                        vAgencia = Mid(RegTran!cMovNro, 18, 2)
'                    Else
'                        vAgencia = Right(Trim(RegTran!cCodAge), 2)
'                    End If

                    vSeg = DateDiff("s", vTime, Time)
                    vSS = vSeg - (Int(vSeg / 60) * 60)
                    vMM = Int(vSeg / 60)
                    vHH = Int(vSeg / 3600)
                    'If Mid(RegTran!cCtaCod, 6, 3) = "401" Or Mid(RegTran!cCtaCod, 6, 3) = "403" Or Mid(RegTran!cCtaCod, 6, 3) = "423" Then
                    '    MsgBox "Prueba si hay hipotecario"
                    'End If
                    Me.Caption = " CodCta.: " & RegTran!cCtaCod & " - Regis.: " & vCont + 1 & " - CtaCnt.: " & vCtaCnt & " - Time -> " & FillNum(str(vHH), 6, "0") & ":" & FillNum(str(vMM), 6, "0") & ":" & FillNum(str(vSS), 6, "0")

                    If IIf(IsNull(RegTran!cTpoCredCod), 0, 1) = 0 Then
                        vCodConta = GetPlantillaPuente(RegOpeCta!cCtaContCod, Mid(RegTran!cCtaCod, 6, 3), RegTran!Concepto, RegTran!cOpeCod, oCon.ConexionActiva)
                    Else
                        vCodConta = GetPlantillaPuente(RegOpeCta!cCtaContCod, RegTran!cTpoCredCod, RegTran!Concepto, RegTran!cOpeCod, oCon.ConexionActiva)
                    End If

                    If vCodConta = "" Then
                        vCodConta = Trim(RegOpeCta!cCtaContCod)
                    End If
                    'ALPA20140303***********************************
                    If Trim(lsPlantilla5114) = Trim(vCodConta) And IIf(IsNull(RegTran!cTpoCredCod), "", Trim(RegTran!cTpoCredCod)) = "755" Then
                        lbPlantillaLogico = True
                    End If
                    '***********************************************
                    If Not IsNull(RegTran!cCtaCod) Then
                        'Variable de cambios
                        vFondo = "": vPlazo = "": vTipoC = "": vRFA = ""

                         ''' validacion de creditos RFA (RFC-DIF)
                        '**14M4TC2501RFCD0PSCFFAG
                        If (lsRfa = "RFC" Or lsRfa = "DIF") And Left(vCodConta, 2) = "14" Then
                            'Stop
                            '' CAMBIAMOS PLANTILLAS PARA LOS CREDITOS rfa SI ES CUENTA 14
                            'If RegTran!nCredEstado = 2031 Then
                            If Left(vCodConta, 4) = "14M5" Then
                                vCodConta = Left(vCodConta, 6) + "1902RFCD0PSCFFAG"
                            Else
                                vCodConta = Left(vCodConta, 6) + "2501RF060PSCFFAG"
                            End If
                            'tipo de credito RFA (NEMMONICO RF)
                            Select Case lsRfa
                                Case "RFC"
                                    vRFA = "01"
                                Case "DIF"
                                    vRFA = "02"
                                Case Else
                                    vRFA = ""
                                    vCodConta = "15"
                            End Select
                            vCodConta = Replace(vCodConta, "RF", vRFA, , , vbTextCompare)
                        End If

                        'Arreglo de creditos de consumo vencidos refinanciados - EJRS
                        'If (Left(vCodConta, 4) = "14M5" Or Left(vCodConta, 4) = "14M6") And Left(RegTran!cOpecod, 4) = "1007" And Mid(RegTran!cCtaCod, 6, 1) = "3" Then
                        If (Left(vCodConta, 4) = "14M5" Or Left(vCodConta, 4) = "14M6") And Left(RegTran!cOpeCod, 4) = "1007" And (Mid(RegTran!cTpoCredCod, 1, 1) = "7" Or Mid(RegTran!cTpoCredCod, 1, 1) = "6") Then
                            'ELIMINAMOS EL PAR 29 PUESTO QUE EN CONSUMO NO SE CONSIDERA
                            vCodConta = Replace(vCodConta, "29", "")
                        End If

                        'Tipo de credito
                        If InStr(vCodConta, "TC") > 0 Then
                            If Mid(RegTran!cCtaCod, 6, 3) = Producto.gCapCTS Then
                                vTipoC = ClienteTipoCTS(RegTran!cCtaCod, oCon.ConexionActiva)
                                If vTipoC = "01" Or vTipoC = "02" Then
                                    vCodConta = Replace(vCodConta, "TC", vTipoC, , , vbTextCompare)
                                Else
                                    MsgBox " Tipo de Cliente - CTS Errado " & RegTran!cCtaCod, vbInformation, " Aviso "
                                    'vNoCtaCnt = vNoCtaCnt & oImpresora.gPrnSaltoLinea & "     * " & RegTran!cCtaCod & " tipo de cliente - CTS errado."
                                    vNoCtaCnt = vNoCtaCnt & vbCrLf & "     * " & RegTran!cCtaCod & " tipo de cliente - CTS errado."
                                End If
                            Else ' Para TC Tipo de Creditos
                                '**Modificado por DAOR 20100623, Proyecto Basilea II ***************************
                                vTipoC = VarAsientoProdEquiv(RegTran!cTpoCredCod, RegTran!cTpoProdCod, RegTran!nTpoInstCorp, "TC", oCon.ConexionActiva)
                                vCodConta = Replace(vCodConta, "TC", vTipoC, , , vbTextCompare)
                                '*******************************************************************************
                            End If
                        End If

                        'JIPR20200329 INCIO Reprogramación
                        If InStr(vCodConta, "CI") > 0 Then
                            If RegTran!cOpeCod = 100929 Or RegTran!cOpeCod = 100950 Or RegTran!cOpeCod = 107337 Or RegTran!cOpeCod = 300660 Then
                            '    Select Case RegTran!cTipo
                             Select Case RegTran!Concepto
                                    Case "1080"
                                        vConceptoI = "01"
                                    Case "1081"
                                        vConceptoI = "02"
                                    Case "1082"
                                        vConceptoI = "03"
'                                    Case "1083"
'                                        vConceptoI = "04"
'                                    Case "1084"
'                                        vConceptoI = "05"
                             End Select
                            vCodConta = Replace(vCodConta, "CI", vConceptoI, , , vbTextCompare)
                            End If
                        End If
                        'JIPR20200329 FIN



                        If InStr(vCodConta, "SC") > 0 Then
                            'ALPA20140303**********************************************
                            If lbPlantillaLogico = True Then
                                  vSC = "02"
                            Else
                                  vSC = VarAsientoProdEquiv(RegTran!cTpoCredCod, RegTran!cTpoProdCod, RegTran!nTpoInstCorp, "SC", oCon.ConexionActiva)
                            End If
                            If vSC = "" Then
                                'vNoCtaCnt = vNoCtaCnt & oImpresora.gPrnSaltoLinea & "     * " & RegTran!cCtaCod & " tipo de SC errado."
                                vNoCtaCnt = vNoCtaCnt & vbCrLf & "     * " & RegTran!cCtaCod & " tipo de SC errado."

                            Else
                                vCodConta = Replace(vCodConta, "SC", vSC, , , vbTextCompare)
                            End If
                        End If
                        If InStr(vCodConta, "IK") > 0 Then
                            If Me.chkSinCambioSBS.value = 1 Then
                                vCodConta = Replace(vCodConta, "IK", "", , , vbTextCompare)
                            Else
                                '**Modificado por DAOR 20100623, Proyecto Basilea II************************
                                 vCodConta = Replace(vCodConta, "IK", VarAsientoProdEquiv(RegTran!cTpoCredCod, RegTran!cTpoProdCod, RegTran!nTpoInstCorp, "IK", oCon.ConexionActiva), , , vbTextCompare)
                                '****************************************************************************
                            End If
                        End If
                        If InStr(vCodConta, "CD") > 0 Then
                            '**Modificado por DAOR 20100623, Proyecto Basilea II***************************
                            vCD = VarAsientoProdEquiv(RegTran!cTpoCredCod, RegTran!cTpoProdCod, RegTran!nTpoInstCorp, "CD", oCon.ConexionActiva)
                            '******************************************************************************
                            If vCD = "" Then
                            Else
                                vCodConta = Replace(vCodConta, "CD", vCD, , , vbTextCompare)
                            End If
                        End If
                        'ALPA 20111219***************************
                        If InStr(vCodConta, "TL") > 0 Then
                            vTL = VarAsientoProdEquiv(RegTran!cTpoCredCod, RegTran!cTpoProdCod, RegTran!nTpoInstCorp, "TL", oCon.ConexionActiva)
                            If vTL = "" Then
                            Else
                                vCodConta = Replace(vCodConta, "TL", vTL, , , vbTextCompare)
                            End If
                        End If
                        
                        '***************************************
                        '**DAOR 20100623, Proyecto Basilea II ******************************
                        If RegTran!nTpoInstCorp = 5 Then 'Empresas del Sistema Financiero, sólo para créditos
                            lsNSIF = VarAsientoEquivEmpSistFinanc(RegTran!cCtaCod, oCon.ConexionActiva)
                            If InStr(vCodConta, "IF") > 0 Then
                                vIF = IIf(Len(lsNSIF) > 3, Right(lsNSIF, 2), "")
                                vCodConta = IIf(vIF = "", "00", Replace(vCodConta, "IF", vIF, , , vbTextCompare))
                            End If

                            If InStr(vCodConta, "NS") > 0 Then
                                vNS = IIf(Len(lsNSIF) > 1, Left(lsNSIF, 2), "")
                                vCodConta = IIf(vNS = "", "00", Replace(vCodConta, "NS", vNS, , , vbTextCompare))
                            End If
                        Else
                            If InStr(vCodConta, "IF") > 0 Then
                                vCodConta = Replace(vCodConta, "IF", "", , , vbTextCompare)
                            End If

                            If InStr(vCodConta, "NS") > 0 Then
                                vCodConta = Replace(vCodConta, "NS", "", , , vbTextCompare)
                            End If
                        End If
                        '*******************************************************************
                        'Cambios para FF - Fondos
                        If InStr(vCodConta, "FF") > 0 Then
                            vFondo = VarFF(RegTran!cCtaCod, IIf(IsNull(RegTran!LineaC), "", RegTran!LineaC), "", oCon.ConexionActiva)
                            If vFondo = "" Then
                                'MsgBox " Tipo de FF Errado " & RegTran!cCtaCod, vbInformation, " Aviso "
                                'vNoCtaCnt = vNoCtaCnt & oImpresora.gPrnSaltoLinea & "     * " & RegTran!cCtaCod & " tipo de FF errado."
                                vNoCtaCnt = vNoCtaCnt & vbCrLf & "     * " & RegTran!cCtaCod & " tipo de FF errado."

                            End If
                            vCodConta = Replace(vCodConta, "FF", vFondo, 1, 1, vbTextCompare)
                        End If
                        'MARCA 02 ¿Que es fondo?
                        'Valida sólo para Crédito FF:02 - LC:01 y Cta = "28...."
                        If Left(Trim(RegTran!cOpeCod), 2) = "10" And vFondo = "02" _
                            And Left(vCodConta, 2) = "29" And Left(Trim(RegTran!cOpeCod), 3) <> "018" Then
                            vCodConta = ""
                            pband28 = False
                        End If
                        'Valida para Foncodes - Crédito
                        'MARCA 04
                        pBandFonc = False
                        If Left(Trim(RegTran!cOpeCod), 2) = "10" And vFondo = "02" And _
                            (Left(vCodConta, 2) = "14" Or Left(vCodConta, 2) = "51") Then
                            pBandFonc = True
                        End If
                        'Cambios para AO
                        'MARCA 05
                        If InStr(vCodConta, "AO") > 0 Then
                            If InStr(RegTran!cNumDoc, "@") = 1 Then
                                vCodConta = Replace(vCodConta, "AO", Mid(RegTran!cNumDoc, 5, 2), , , vbTextCompare)
                            Else
                                MsgBox " No se encontro AO - Créditos " & RegTran!cCtaCod, vbInformation, " Aviso "
                                'vNoCtaCnt = vNoCtaCnt & oImpresora.gPrnSaltoLinea & "     * " & RegTran!cCtaCod & " no se encontro AO - Créditos."
                                vNoCtaCnt = vNoCtaCnt & vbCrLf & "     * " & RegTran!cCtaCod & " no se encontro AO - Créditos."

                            End If
                        End If
                    End If
                    'Codificacion de los bancos para el canje de Ordenes de Pago
                    If InStr(vCodConta, "BC") > 0 Then
                        vBanco = VarBC(RegTran!nMovNro, vCodConta, oCon.ConexionActiva)
                        If vBanco = "" Then
                            MsgBox " Número de Banco no Identificado (Después de función) " & RegTran!cCtaCod, vbInformation, " Aviso "
                            'vNoCtaCnt = vNoCtaCnt & oImpresora.gPrnSaltoLinea & "     * " & RegTran!cCtaCod & " número de banco no identificado (Después de función) ."
                            vNoCtaCnt = vNoCtaCnt & vbCrLf & "     * " & RegTran!cCtaCod & " número de banco no identificado (Después de función) ."
                        Else
                            vCodConta = Replace(vCodConta, "BC", vBanco)
                        End If
                    End If
                    'Valida la CJ - Ahorros - Credito - Pignoraticio
                    'MARCA 06



                    If InStr(vCodConta, "CJ") > 0 Then
                        '************ OJO HAY QUE VERIFICAR ***************************
                        vCJ = VarInstitucionFinanciera(RegTran!cCtaCod, gTpoIFCmac, oCon.ConexionActiva)
                        If vCJ = "" Then
                           vCJ = VarInstitucionFinanciera(RegTran!cCtaCod, gTpoIFCooperativa, oCon.ConexionActiva)
                        End If
                        If vCJ = "" Then
                           vCJ = VarInstitucionFinanciera(RegTran!cCtaCod, gTpoIFEDPYME, oCon.ConexionActiva)
                        End If
                        If vCJ = "" Then
                           vCJ = VarInstitucionFinanciera(RegTran!cCtaCod, gTpoIFCrac, oCon.ConexionActiva)
                        End If
                        If vCJ = "" Then
                            ' parche para operaciones de recepcion CMAC
                            If (Left(RegTran!cOpeCod, 4) >= "2601" And Left(RegTran!cOpeCod, 4) <= "2603") Or _
                                RegTran!cOpeCod = "100205" Or Left(RegTran!cOpeCod, 4) = "9903" Or RegTran!cOpeCod = "100405" Or _
                                RegTran!cOpeCod = "100305" Or RegTran!cOpeCod = "136301" Or RegTran!cOpeCod = "136201" Or _
                                Left(RegTran!cOpeCod, 3) = "126" Or RegTran!cOpeCod = "100505" Then

                                'Los 126 Agregado por JHVP

                                vCJ = VarInstFinanMov(RegTran!nMovNro, gTpoIFCmac, oCon.ConexionActiva)
                                If vCJ = "" Then
                                    MsgBox " CJ no reconocido (Ahorros) " & RegTran!cCtaCod, vbInformation, " Aviso "
                                    'vNoCtaCnt = vNoCtaCnt & oImpresora.gPrnSaltoLinea & "     * " & RegTran!cCtaCod & " CJ no reconocido (Ahorros)."
                                    vNoCtaCnt = vNoCtaCnt & vbCrLf & "     * " & RegTran!cCtaCod & " CJ no reconocido (Ahorros)."
                                Else
                                    vCodConta = Replace(vCodConta, "CJ", vCJ, 1, 1, vbTextCompare)
                                End If
                            End If
                        Else
                            vCodConta = Replace(vCodConta, "CJ", vCJ, 1, 1, vbTextCompare)
                        End If
                    End If
                    
                    'Valida la CR - Otras Instituciones Financieras - Cajas Rurales (Raul)
                    If InStr(vCodConta, "CR") > 0 Then
                        If RegTran!cOpeCod = "700104" Then
                            vCR = VarInstitucionFinanciera(RegTran!cCtaCodReal, gTpoIFCrac, oCon.ConexionActiva)
                        Else
                            vCR = VarInstitucionFinanciera(RegTran!cCtaCod, gTpoIFCrac, oCon.ConexionActiva)
                        End If

                        If vCR = "" Then
                            MsgBox " CR no reconocido (Ahorros) " & RegTran!cCtaCod, vbInformation, " Aviso "
                            'vNoCtaCnt = vNoCtaCnt & oImpresora.gPrnSaltoLinea & "     * " & RegTran!cCtaCod & " CR no reconocido (Ahorros)."
                            vNoCtaCnt = vNoCtaCnt & vbCrLf & "     * " & RegTran!cCtaCod & " CR no reconocido (Ahorros)."

                        Else
                            vCodConta = Replace(vCodConta, "CR", vCR, 1, 1, vbTextCompare)
                        End If
                    End If
                    If InStr(vCodConta, "CO") > 0 Then
                        vCR = VarInstitucionFinanciera(RegTran!cCtaCod, gTpoIFCooperativa, oCon.ConexionActiva)
                        If vCR = "" Then
                            MsgBox " CO no reconocido (Ahorros) " & RegTran!cCtaCod, vbInformation, " Aviso "
                            'vNoCtaCnt = vNoCtaCnt & oImpresora.gPrnSaltoLinea & "     * " & RegTran!cCtaCod & " CR no reconocido (Ahorros)."
                            vNoCtaCnt = vNoCtaCnt & vbCrLf & "     * " & RegTran!cCtaCod & " CR no reconocido (Ahorros)."

                        Else
                            vCodConta = Replace(vCodConta, "CO", vCR, 1, 1, vbTextCompare)
                        End If
                    End If
                    If InStr(vCodConta, "ED") > 0 Then
                        If RegTran!cOpeCod = "700104" Then
                            vCR = VarInstitucionFinanciera(RegTran!cCtaCodReal, gTpoIFEDPYME, oCon.ConexionActiva)
                        Else
                            vCR = VarInstitucionFinanciera(RegTran!cCtaCod, gTpoIFEDPYME, oCon.ConexionActiva)
                        End If
                        If vCR = "" Then
                            MsgBox " ED no reconocido (Ahorros) " & RegTran!cCtaCod, vbInformation, " Aviso "
                            'vNoCtaCnt = vNoCtaCnt & oImpresora.gPrnSaltoLinea & "     * " & RegTran!cCtaCod & " CR no reconocido (Ahorros)."
                            vNoCtaCnt = vNoCtaCnt & vbCrLf & "     * " & RegTran!cCtaCod & " CR no reconocido (Ahorros)."

                        Else
                            vCodConta = Replace(vCodConta, "ED", vCR, 1, 1, vbTextCompare)
                        End If
                    End If

                    'Créditos
                    If InStr(vCodConta, "PD") > 0 Then
                        vTipoPD = VarPD(RegTran!cCtaCod, IIf(IsNull(RegTran!LineaC), "", RegTran!LineaC), "")
                        If vTipoPD = "01" Or vTipoPD = "02" Or vTipoPD = "03" Or vTipoPD = "05" Then
                            vCodConta = Replace(vCodConta, "PD", vTipoPD, , , vbTextCompare)
                        Else
                            MsgBox " Producto Pendiente Errado " & RegTran!cCtaCod, vbInformation, " Aviso "
                            'vNoCtaCnt = vNoCtaCnt & oImpresora.gPrnSaltoLinea & "     * " & RegTran!cCtaCod & " producto pendiente errado."
                            vNoCtaCnt = vNoCtaCnt & vbCrLf & "     * " & RegTran!cCtaCod & " producto pendiente errado."

                        End If
                    End If

                    'Cambios para P - Plazo
                    If InStr(vCodConta, "P") > 0 Then
                        vPlazo = Mid(Trim(RegTran!LineaC), 6, 1)
                        If vPlazo = "1" Or vPlazo = "2" Then
                            vCodConta = Replace(vCodConta, "P", vPlazo, 1, 1, vbTextCompare)
                        Else
                            MsgBox " Plazo de Crédito Errada " & RegTran!cCtaCod, vbInformation, " Aviso "
                            'vNoCtaCnt = vNoCtaCnt & oImpresora.gPrnSaltoLinea & "     * " & RegTran!cCtaCod & " plazo de crédito errado."
                            vNoCtaCnt = vNoCtaCnt & vbCrLf & "     * " & RegTran!cCtaCod & " plazo de crédito errado."

                        End If
                    End If

                    'Valida la agencia remota - AG - AL - AR

                    If InStr(vCodConta, "AR") > 0 Then
                        vCodConta = Replace(vCodConta, "AR", vAgencia, 1, 1, vbTextCompare)
                    End If
                    '------ LAYG 07/01/2005 ica
                    If InStr(vCodConta, "AG") > 0 Then

                        If Left(vCodConta, 2) = "11" Then
                            'vCodConta = Replace(vCodConta, "AG", Mid(RegTran!cMovNro, 15, 5), 1, 1, vbTextCompare)
                            vAG = VarAG(Mid(RegTran!cMovNro, 15, 5), vCodConta)
                        Else
                            If Left(RegTran!cOpeCod, 4) = Left(gPigOpeDespContNuevo, 4) Then
                                'vCodConta = Replace(vCodConta, "AG", vAgencia, 1, 1, vbTextCompare)
                                vAG = VarAG(gsCodCMAC & vAgencia, vCodConta)
                            Else
                                'vCodConta = Replace(vCodConta, "AG", Mid(RegTran!cCtaCod, 4, 2), 1, 1, vbTextCompare)
                                If RegTran!cOpeCod = gServGiroCancEfec Or RegTran!cOpeCod = gITFGiroCancelEfect Then 'NAGL 20190306 Agregó RegTran!cOpeCod = gITFGiroCancelEfect
                                    vAG = RegTran!cCodAgeGir
                                Else
                                    vAG = VarAG(RegTran!cCtaCod, vCodConta, oCon.ConexionActiva)
                                    'vAG = VarAG(RegTran!cCtaCod, vCodConta)
                                End If
                            End If
                        End If

                        'If RegTran!cOpecod = "100911" Or RegTran!cOpecod = "100912" Then
                        If RegTran!TipoDato = 1 And Left(vCodConta, 2) = "14" Then
                            vAG = RegTran!cAgeCodAct
                        End If
                        'CAMBIAMOS AGENCIA SI OPERACION ES DESEMBOLSO CON ABONO A CUENTA DE OTRA AGENCIA
                        'NUEVO PARCHE EJRS  26 DE MAYO 2005
                        If RegTran!cOpeCod = "100104" And Left(vCodConta, 2) = "21" Then
                           vAG = GetAgenciaCtaDesmAbonoCta(RegTran!nMovNro, RegTran!cOpeCod, oCon.ConexionActiva)
                           vAgencia = vAG
                        End If
                        If vAG = "" Then
                            MsgBox " AG no reconocida " & RegTran!cCtaCod, vbInformation, " Aviso "
                            'vNoCtaCnt = vNoCtaCnt & oImpresora.gPrnSaltoLinea & "     * " & RegTran!cCtaCod & " AG no reconocida."
                            vNoCtaCnt = vNoCtaCnt & vbCrLf & "     * " & RegTran!cCtaCod & " AG no reconocida."

                        Else
                            vCodConta = Replace(vCodConta, "AG", vAG, 1, 1, vbTextCompare)
                        End If

                    End If
                    '----------
                    'Cambios en Moneda y Agencia
                    vCodConta = Replace(vCodConta, "M", vMoneda, 1, 1, vbTextCompare)
                    'Asiento para ctas. de Foncodes en la Cta de Ahorro : 012321006254 - RAUL
                    'MARCA 08
                    'Para unificar varias cuentas en una - ARTURO
                    vParche = AsientoParche(vCodConta, True, oCon.ConexionActiva)
                    If Len(vParche) > 0 Then vCodConta = vParche

                    'Se valida la existencia de la Cta. para que se pueda utilizar
                    'ALPA 20130527
                    'If (Not ExisCtaCnt(vCodConta, True, oCon.ConexionActiva) And Len(Trim(vCodConta)) > 0) Or Len(Trim(vCodConta)) = 0 Then
                    If RegTran!cOpeCod <> "700104" Then
                        If ((Not ExisCtaCnt(vCodConta, True, oCon.ConexionActiva) And Len(Trim(vCodConta)) > 0) Or Len(Trim(vCodConta)) = 0) Then
                            'MsgBox "No existe la Cta. Contable y/o no es Cta de Asiento: " & vCodConta, vbInformation, " Aviso "
                            'vNoCtaCnt = vNoCtaCnt & oImpresora.gPrnSaltoLinea & "     * " & vCodConta & " cta. cnt. no existe (cta.: " & RegTran!cCtaCod & ", ope.: " & RegTran!cOpeCod & " - RFA = " & RegTran!cRFA & " - Concepto " & RegTran!Concepto & ")."
                            vNoCtaCnt = vNoCtaCnt & vbCrLf & "     * " & vCodConta & " cta. cnt. no existe (cta.: " & RegTran!cCtaCod & ", ope.: " & RegTran!cOpeCod & " - RFA = " & RegTran!cRFA & " - Concepto " & RegTran!Concepto & ")."

                        End If
                    End If

                    'If Trim(vCodConta) = "19180207" Then Stop

                    '****************************************************************************
                    'Verifica si el asiento ya fue creado
                    'Si existe lo actualiza; caso contrario lo agrega

                    Call ObtieneTipoCambioLeasing(RegTran!cCtaCod, lnTipoCambioCompraL, lnTipoCambioVentaL, lnTipoCambioFijoL)

'                    If Mid(RegTran!cCtaCod, 6, 3) = "515" Or Mid(RegTran!cCtaCod, 6, 3) = "516" Then
'                            Call ObtenerTipoCambioLeasing(RegTran!cCtaCod, oCon.ConexionActiva, lnTipoCambioCompraL, lnTipoCambioVentaL, lnTipoCambioFijoL)
'                    End If


                    If Len(Trim(vCodConta)) > 0 Then
                        If RegOpeCta!cOpeCtaDH = "D" Then

                            '**DAOR 20081202 para incluir operaciones por ATM ***********************************
                            If (RegTran!cOpeCod = "208021" Or RegTran!cOpeCod = "228021" Or RegTran!cOpeCod = "208031" Or RegTran!cOpeCod = "228031" Or RegTran!cOpeCod = "208035" Or RegTran!cOpeCod = "228035") And RegTran!nTipCambDispuesto > 0 And RegTran!Concepto = 1 Then   'Operación de ATM Compra/Venta 'AMDO20150401
                                'Mid(vCodConta, 3, 1) = IIf(Mid(vCodConta, 3, 1) = "1", "2", "1") 'invertir moneda
                                If Mid(vCodConta, 3, 1) = Moneda.gMonedaNacional Then 'Venta Dolares
                                    If (RegTran!nMontoDispuesto * TCF) > RegTran!sumamonto Then
                                        lnMontoOperacion = RegTran!sumamonto '+ ((RegTran!nMontoDispuesto * TCF) - RegTran!SumaMonto)
                                    Else
                                        lnMontoOperacion = RegTran!sumamonto '- (RegTran!SumaMonto - (RegTran!nMontoDispuesto * TCF))
                                    End If
                                Else 'Compra Dolares
                                    If RegTran!nMontoDispuesto > (RegTran!sumamonto * TCF) Then
                                        lnMontoOperacion = RegTran!sumamonto '+ ((RegTran!nMontoDispuesto * TCF) - RegTran!SumaMonto)
                                    Else
                                        lnMontoOperacion = RegTran!sumamonto '- (RegTran!SumaMonto - (RegTran!nMontoDispuesto * TCF))
                                    End If
                                End If
                            End If
                            '***************************************************************************************
                            Call IngresarAsientoDN(gdHoraGrab, vCodConta, lnMontoOperacion, 0, IIf(RegTran!Concepto = 8 Or RegTran!Concepto = 9, "3", "0"), vAgencia, RegTran!nMovNro, RegTran!cOpeCod, RegTran!cCtaCod, oCon.ConexionActiva, TCF)
'                            ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge, nMovNro, cOpeCod, cCtaCod) " & _
'                                " VALUES('" & gdHoraGrab & "','" & vCodConta & "'," & lnMontoOperacion & ",0,'" & IIf(RegTran!Concepto = 8 Or RegTran!Concepto = 9, "3", "0") & "','" & vAgencia & "'," & RegTran!nMovNro & ",'" & RegTran!cOpecod & "','" & RegTran!cCtaCod & "')"
'                            oCon.Ejecutar ssql
                            If Mid(vCodConta, 3, 1) = Moneda.gMonedaExtranjera And Not (RegTran!Concepto = 8 Or RegTran!Concepto = 9) Then
                                'ALPA 20120420***************************
                                If Left(vCodConta, 1) = "4" Or Left(vCodConta, 2) = "63" Or Left(vCodConta, 2) = "65" Then
                                Call IngresarAsientoDN(gdHoraGrab, vCodConta, lnMontoOperacion * TCV, 0, 3, vAgencia, RegTran!nMovNro, RegTran!cOpeCod, RegTran!cCtaCod, oCon.ConexionActiva, TCF)
'                                    ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge, nMovNro, cOpeCod, cCtaCod) " & _
'                                        " VALUES('" & gdHoraGrab & "','" & vCodConta & "'," & (lnMontoOperacion * TCV) & ",0,'3','" & vAgencia & "'," & RegTran!nMovNro & ",'" & RegTran!cOpecod & "','" & RegTran!cCtaCod & "')"
'                                    oCon.Ejecutar ssql
                                    'Ctas 282503 y 191503
                                    vCta28Debe = vCta28Debe + (lnMontoOperacion * TCF)
                                    vCta19Haber = vCta19Haber + (lnMontoOperacion * TCV)
                                'ALPA 20120420***************************
                                ElseIf Left(vCodConta, 1) = "5" Or Left(vCodConta, 2) = "62" Or Left(vCodConta, 2) = "64" Then
                                'ElseIf Left(vCodConta, 1) = "5" Then
                                Call IngresarAsientoDN(gdHoraGrab, vCodConta, lnMontoOperacion * TCC, 0, 3, vAgencia, RegTran!nMovNro, RegTran!cOpeCod, RegTran!cCtaCod, oCon.ConexionActiva, TCF)
'                                    ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge, nMovNro, cOpeCod, cCtaCod) " & _
'                                        " VALUES('" & gdHoraGrab & "','" & vCodConta & "'," & (lnMontoOperacion * TCC) & ",0,'3','" & vAgencia & "'," & RegTran!nMovNro & ",'" & RegTran!cOpecod & "','" & RegTran!cCtaCod & "')"
'                                    oCon.Ejecutar ssql
                                    'Ctas 282503 y 191503
                                    vCta28Debe = vCta28Debe + (lnMontoOperacion * TCF)
                                    vCta19Haber = vCta19Haber + (lnMontoOperacion * TCC)
                                Else
                                    Call IngresarAsientoDN(gdHoraGrab, vCodConta, lnMontoOperacion * TCF, 0, 3, vAgencia, RegTran!nMovNro, RegTran!cOpeCod, RegTran!cCtaCod, oCon.ConexionActiva, TCF)
'                                    ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge, nMovNro, cOpeCod, cCtaCod) " & _
'                                        " VALUES('" & gdHoraGrab & "','" & vCodConta & "'," & (lnMontoOperacion * TCF) & ",0,'3','" & vAgencia & "'," & RegTran!nMovNro & ",'" & RegTran!cOpecod & "','" & RegTran!cCtaCod & "')"
'                                    oCon.Ejecutar ssql
                                End If
                            End If
                        ElseIf RegOpeCta!cOpeCtaDH = "H" Then

                            '**DAOR 20081202 para incluir operaciones por ATM ***********************************
                            If (RegTran!cOpeCod = "208021" Or RegTran!cOpeCod = "228021" Or RegTran!cOpeCod = "208031" Or RegTran!cOpeCod = "228031" Or RegTran!cOpeCod = "208035" Or RegTran!cOpeCod = "228035") And RegTran!nTipCambDispuesto > 0 And RegTran!Concepto = 1 Then   'Operación de ATM Compra/Venta
                                lsCtaCtbPerdidaT = lsCtaCtbPerdida
                                lsCtaCtbGananciaT = lsCtaCtbGanancia
                                lsCtaCtbPerdidaT = Replace(lsCtaCtbPerdidaT, "M", "2", 1, 1, vbTextCompare)
                                lsCtaCtbPerdidaT = Replace(lsCtaCtbPerdidaT, "AG", Mid(RegTran!cCtaCod, 4, 2), 1, 1, vbTextCompare)

                                lsCtaCtbGananciaT = Replace(lsCtaCtbGananciaT, "M", "2", 1, 1, vbTextCompare)
                                lsCtaCtbGananciaT = Replace(lsCtaCtbGananciaT, "AG", Mid(RegTran!cCtaCod, 4, 2), 1, 1, vbTextCompare)

                                If Mid(vCodConta, 3, 1) = Moneda.gMonedaNacional Then 'Venta Dolares
                                    If (RegTran!nMontoDispuesto * TCF) > RegTran!sumamonto Then
                                        lnMontoOperacionGP = (RegTran!nMontoDispuesto * TCF) - RegTran!sumamonto
'                                        ssql = "insert into AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge, nMovNro, cOpeCod, cCtaCod) " & _
'                                                " values('" & gdHoraGrab & "','" & lsCtaCtbPerdidaT & "'," & lnMontoOperacionGP & ",0,'3','" & vAgencia & "'," & RegTran!nMovNro & ",'" & RegTran!cOpecod & "','" & RegTran!cCtaCod & "')"
                                        Call IngresarAsientoDN(gdHoraGrab, lsCtaCtbPerdidaT, lnMontoOperacionGP, 0, 3, vAgencia, RegTran!nMovNro, RegTran!cOpeCod, RegTran!cCtaCod, oCon.ConexionActiva, TCF)
                                    Else
                                        lnMontoOperacionGP = RegTran!sumamonto - (RegTran!nMontoDispuesto * TCF)
'                                        ssql = "insert into AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge, nMovNro, cOpeCod, cCtaCod) " & _
'                                                " values('" & gdHoraGrab & "','" & lsCtaCtbGananciaT & "',0," & lnMontoOperacionGP & ",'3','" & vAgencia & "'," & RegTran!nMovNro & ",'" & RegTran!cOpecod & "','" & RegTran!cCtaCod & "')"
                                         Call IngresarAsientoDN(gdHoraGrab, lsCtaCtbGananciaT, 0, lnMontoOperacionGP, 3, vAgencia, RegTran!nMovNro, RegTran!cOpeCod, RegTran!cCtaCod, oCon.ConexionActiva, TCF)
                                    End If
                                Else 'Compra Dolares
                                    If RegTran!nMontoDispuesto > (RegTran!sumamonto * TCF) Then
                                        lnMontoOperacionGP = RegTran!nMontoDispuesto - (RegTran!sumamonto * TCF)
'                                        ssql = "insert into AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge, nMovNro, cOpeCod, cCtaCod) " & _
'                                                " values('" & gdHoraGrab & "','" & lsCtaCtbPerdidaT & "'," & lnMontoOperacionGP & ",0,'3','" & vAgencia & "'," & RegTran!nMovNro & ",'" & RegTran!cOpecod & "','" & RegTran!cCtaCod & "')"
                                        Call IngresarAsientoDN(gdHoraGrab, lsCtaCtbPerdidaT, lnMontoOperacionGP, 0, 3, vAgencia, RegTran!nMovNro, RegTran!cOpeCod, RegTran!cCtaCod, oCon.ConexionActiva, TCF)
                                    Else
                                        lnMontoOperacionGP = (RegTran!sumamonto * TCF) - RegTran!nMontoDispuesto
'                                        ssql = "insert into AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge, nMovNro, cOpeCod, cCtaCod) " & _
'                                                " values('" & gdHoraGrab & "','" & lsCtaCtbGananciaT & "',0," & lnMontoOperacionGP & ",'3','" & vAgencia & "'," & RegTran!nMovNro & ",'" & RegTran!cOpecod & "','" & RegTran!cCtaCod & "')"
                                        Call IngresarAsientoDN(gdHoraGrab, lsCtaCtbGananciaT, 0, lnMontoOperacionGP, 3, vAgencia, RegTran!nMovNro, RegTran!cOpeCod, RegTran!cCtaCod, oCon.ConexionActiva, TCF)
                                    End If
                                End If
'                                oCon.Ejecutar ssql
                                'RegTran!cCtaCod
                                If Left(vCodConta, 2) = "19" Then
                                    Mid(vCodConta, 3, 1) = IIf(Mid(vCodConta, 3, 1) = "1", "2", "1") 'invertir moneda
                                    lnMontoOperacion = RegTran!nMontoDispuesto
                                End If
                                If Mid(RegTran!cCtaCod, 9, 1) = "2" Then
                                    Call IngresarAsientoDN(gdHoraGrab, vCodConta, 0, lnMontoOperacion, 3, vAgencia, RegTran!nMovNro, RegTran!cOpeCod, RegTran!cCtaCod, oCon.ConexionActiva, TCF)
'                                    ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge, nMovNro, cOpeCod, cCtaCod) " & _
'                                       " VALUES('" & gdHoraGrab & "','" & vCodConta & "',0," & lnMontoOperacion & ",'3','" & vAgencia & "'," & RegTran!nMovNro & ",'" & RegTran!cOpecod & "','" & RegTran!cCtaCod & "')"
'                                    oCon.Ejecutar ssql
                                Else
                                    Call IngresarAsientoDN(gdHoraGrab, vCodConta, 0, lnMontoOperacion, 0, vAgencia, RegTran!nMovNro, RegTran!cOpeCod, RegTran!cCtaCod, oCon.ConexionActiva, TCF)
'                                    ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge, nMovNro, cOpeCod, cCtaCod) " & _
'                                       " VALUES('" & gdHoraGrab & "','" & vCodConta & "',0," & lnMontoOperacion & ",'0','" & vAgencia & "'," & RegTran!nMovNro & ",'" & RegTran!cOpecod & "','" & RegTran!cCtaCod & "')"
'                                    oCon.Ejecutar ssql
                                    Call IngresarAsientoDN(gdHoraGrab, vCodConta, 0, lnMontoOperacion * TCF, 3, vAgencia, RegTran!nMovNro, RegTran!cOpeCod, RegTran!cCtaCod, oCon.ConexionActiva, TCF)
'                                    ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge, nMovNro, cOpeCod, cCtaCod) " & _
'                                       " VALUES('" & gdHoraGrab & "','" & vCodConta & "',0," & lnMontoOperacion * TCF & ",'3','" & vAgencia & "'," & RegTran!nMovNro & ",'" & RegTran!cOpecod & "','" & RegTran!cCtaCod & "')"
'                                    oCon.Ejecutar ssql

                                    Mid(vCodConta, 3, 1) = IIf(Mid(vCodConta, 3, 1) = "1", "2", "1") 'invertir moneda
                                    lnMontoOperacion = RegTran!sumamonto
                                    Call IngresarAsientoDN(gdHoraGrab, vCodConta, lnMontoOperacion, 0, 3, vAgencia, RegTran!nMovNro, RegTran!cOpeCod, RegTran!cCtaCod, oCon.ConexionActiva, TCF)
'                                    ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge, nMovNro, cOpeCod, cCtaCod) " & _
'                                       " VALUES('" & gdHoraGrab & "','" & vCodConta & "'," & lnMontoOperacion & ",0,'3','" & vAgencia & "'," & RegTran!nMovNro & ",'" & RegTran!cOpecod & "','" & RegTran!cCtaCod & "')"
'                                    oCon.Ejecutar ssql
                                    Call IngresarAsientoDN(gdHoraGrab, vCodConta, 0, lnMontoOperacion, 0, vAgencia, RegTran!nMovNro, RegTran!cOpeCod, RegTran!cCtaCod, oCon.ConexionActiva, TCF)
'                                    ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge, nMovNro, cOpeCod, cCtaCod) " & _
'                                       " VALUES('" & gdHoraGrab & "','" & vCodConta & "',0," & lnMontoOperacion & ",'0','" & vAgencia & "'," & RegTran!nMovNro & ",'" & RegTran!cOpecod & "','" & RegTran!cCtaCod & "')"
'                                    oCon.Ejecutar ssql

                                End If
                            Else
                                'Cualquier otra operación
                                Call IngresarAsientoDN(gdHoraGrab, vCodConta, 0, RegTran!sumamonto, IIf(RegTran!Concepto = 8 Or RegTran!Concepto = 9, "3", "0"), vAgencia, RegTran!nMovNro, RegTran!cOpeCod, RegTran!cCtaCod, oCon.ConexionActiva, TCF)
'                                    ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge, nMovNro, cOpeCod, cCtaCod) " & _
'                                       " VALUES('" & gdHoraGrab & "','" & vCodConta & "',0," & RegTran!SumaMonto & ",'" & IIf(RegTran!Concepto = 8 Or RegTran!Concepto = 9, "3", "0") & "','" & vAgencia & "'," & RegTran!nMovNro & ",'" & RegTran!cOpeCod & "','" & RegTran!cCtaCod & "')"
'                                    oCon.Ejecutar ssql
                            End If

                            If Mid(vCodConta, 3, 1) = Moneda.gMonedaExtranjera And Not (RegTran!Concepto = 8 Or RegTran!Concepto = 9) Then
                                If Left(vCodConta, 1) = "4" Then
                                    Call IngresarAsientoDN(gdHoraGrab, vCodConta, 0, (lnMontoOperacion * TCV), 3, vAgencia, RegTran!nMovNro, RegTran!cOpeCod, RegTran!cCtaCod, oCon.ConexionActiva)
'                                    ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge, nMovNro, cOpeCod, cCtaCod) " & _
'                                        " VALUES('" & gdHoraGrab & "','" & vCodConta & "',0," & (lnMontoOperacion * TCV) & ",'3','" & vAgencia & "'," & RegTran!nMovNro & ",'" & RegTran!cOpeCod & "','" & RegTran!cCtaCod & "')"
'                                    oCon.Ejecutar ssql
                                    'Ctas 282503 y 191503
                                    vCta28Haber = vCta28Haber + (lnMontoOperacion * TCF)
                                    vCta19Debe = vCta19Debe + (lnMontoOperacion * TCV)
                                ElseIf Left(vCodConta, 1) = "5" Then
                                    Call IngresarAsientoDN(gdHoraGrab, vCodConta, 0, (lnMontoOperacion * TCC), 3, vAgencia, RegTran!nMovNro, RegTran!cOpeCod, RegTran!cCtaCod, oCon.ConexionActiva, TCF)
'                                    ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge, nMovNro, cOpeCod, cCtaCod) " & _
'                                        " VALUES('" & gdHoraGrab & "','" & vCodConta & "',0," & (lnMontoOperacion * TCC) & ",'3','" & vAgencia & "'," & RegTran!nMovNro & ",'" & RegTran!cOpeCod & "','" & RegTran!cCtaCod & "')"
'                                    oCon.Ejecutar ssql
                                    'Ctas 282503 y 191503
                                    vCta28Haber = vCta28Haber + (lnMontoOperacion * TCF)
                                    vCta19Debe = vCta19Debe + (lnMontoOperacion * TCC)
                                Else
                                    If RegTran!cOpeCod = "100111" And Left(RegTran!cOpeCod, 2) = "16" Then
                                    Call IngresarAsientoDN(gdHoraGrab, vCodConta, 0, (lnMontoOperacion * lnTipoCambioVentaL), 3, vAgencia, RegTran!nMovNro, RegTran!cOpeCod, RegTran!cCtaCod, oCon.ConexionActiva, TCF)
'                                    ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge, nMovNro, cOpeCod, cCtaCod) " & _
'                                        " VALUES('" & gdHoraGrab & "','" & vCodConta & "',0," & (lnMontoOperacion * lnTipoCambioVentaL) & ",'3','" & vAgencia & "'," & RegTran!nMovNro & ",'" & RegTran!cOpeCod & "','" & RegTran!cCtaCod & "')"
'                                    oCon.Ejecutar ssql
                                    Else
                                    Call IngresarAsientoDN(gdHoraGrab, vCodConta, 0, (lnMontoOperacion * TCF), 3, vAgencia, RegTran!nMovNro, RegTran!cOpeCod, RegTran!cCtaCod, oCon.ConexionActiva, TCF)
'                                    ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge, nMovNro, cOpeCod, cCtaCod) " & _
'                                        " VALUES('" & gdHoraGrab & "','" & vCodConta & "',0," & (lnMontoOperacion * TCF) & ",'3','" & vAgencia & "'," & RegTran!nMovNro & ",'" & RegTran!cOpeCod & "','" & RegTran!cCtaCod & "')"
'                                    oCon.Ejecutar ssql
                                    End If
                                End If
                            End If
                        Else
                            MsgBox " Operación no reconocida al Insertar " & vbCr & _
                            " operación nro.: " & vCodConta, vbInformation, " Aviso "
                            'vNoCtaCnt = vNoCtaCnt & oImpresora.gPrnSaltoLinea & "     * " & RegOpeCta!cOpeCtaDH & " debe/haber no reconocido (cta.: " & RegTran!cCtaCod & ")."
                            vNoCtaCnt = vNoCtaCnt & vbCrLf & "     * " & RegOpeCta!cOpeCtaDH & " debe/haber no reconocido (cta.: " & RegTran!cCtaCod & ")."

                        End If
                    End If
                    RegOpeCta.MoveNext
                Loop
                RegOpeCta.Close
                Set RegOpeCta = Nothing
            End If
            DoEvents
            vCont = vCont + 1
'            prgList.value = vCont
            RegTran.MoveNext
        Loop

        'Modificado PASI20140513 TI-ERS004-2014
        'Dim vCta83D As Currency, vCta83H As Currency, vCta84D As Currency, vCta84H As Currency, vDife8 As Currency
        'Dim vCta83D As Currency, vCta83H As Currency, vCta84D As Currency, vCta84H As Currency, vDife8 As Double
        'end PASI

        'CUADRE DE CTAS DE ORDEN (SILVITA)   83 = 84
        'MARCA 11 CUENTAS DE ORDEN
        vCta83D = 0: vCta83H = 0
        tmpSql = "SELECT abs(round(SUM(ndebe),2)) nDebe, abs(round(SUM(nhaber),2)) nHaber FROM AsientoDN " & _
            " WHERE ctipo = '3' AND datediff(dd, dFecha ,'" & gdHoraGrab & "') = 0 " & _
            " AND cCtaCnt like '83%'" & _
            " Group BY cCtaCnt "
        Set RegTmp = oCon.CargaRecordSet(tmpSql)
        If (RegTmp.BOF Or RegTmp.EOF) Then
            'No se encuentra ctas en Dolares o Tipo 3
        Else
            Do While Not RegTmp.EOF
                vCta83D = vCta83D + RegTmp!nDebe
                vCta83H = vCta83H + RegTmp!nHaber
                RegTmp.MoveNext
            Loop
        End If
        RegTmp.Close
        Set RegTmp = Nothing
        vCta83D = Abs(vCta83D)
        vCta83H = Abs(vCta83H)

        vCta84D = 0: vCta84H = 0
        tmpSql = "SELECT abs(round(SUM(ndebe),2)) nDebe,  abs(round(SUM(nhaber),2)) nHaber FROM AsientoDN " & _
            " WHERE ctipo = '3' AND datediff(dd, dFecha ,'" & gdHoraGrab & "') = 0 " & _
            " AND cCtaCnt like '84%'" & _
            " Group BY cCtaCnt "
        Set RegTmp = oCon.CargaRecordSet(tmpSql)
        If (RegTmp.BOF Or RegTmp.EOF) Then
            'No se encuentra ctas en Dolares o Tipo 3
        Else
            Do While Not RegTmp.EOF
                vCta84D = vCta84D + RegTmp!nDebe
                vCta84H = vCta84H + RegTmp!nHaber
                RegTmp.MoveNext
            Loop
        End If
        RegTmp.Close
        Set RegTmp = Nothing
        vCta84D = Abs(vCta84D)
        vCta84H = Abs(vCta84H)

        'Validación
        'MsgBox (vCta83D + vCta83H) & "  " & (vCta84D + vCta84H)
        Dim nResult83 As Currency, nResult84 As Currency

        nResult83 = IIf(vCta83D > vCta83H, vCta83D - vCta83H, vCta83H - vCta83D)
        nResult84 = IIf(vCta84D > vCta84H, vCta84D - vCta84H, vCta84H - vCta84D)

        If nResult83 <> nResult84 Then
            vDife8 = 0
            vDife8 = IIf(nResult83 > nResult84, nResult83 - nResult84, nResult84 - nResult83)
            'vDife8 = Abs(nResult83 - nResult84)
            If Abs(nResult83) > Abs(nResult84) Then
                'SI 83 ES MAYOR
                If vCta83D > vCta83H Then
                    ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge) " & _
                        " VALUES('" & gdHoraGrab & "','8321',0," & vDife8 & ",'3','" & vAgencia & "')"
                    oCon.ejecutar ssql

                    ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge) " & _
                        " VALUES('" & gdHoraGrab & "','" & lsCVMEGanacia & "'," & vDife8 & ",0,'3','" & vAgencia & "')"
                    oCon.ejecutar ssql
                Else
                    ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge) " & _
                        " VALUES('" & gdHoraGrab & "','8321'," & vDife8 & ",0,'3','" & vAgencia & "')"
                    oCon.ejecutar ssql

                    ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge) " & _
                        " VALUES('" & gdHoraGrab & "','" & lsCVMEGanacia & "',0," & vDife8 & ",'3','" & vAgencia & "')"
                    oCon.ejecutar ssql
                End If
            Else
                'SI 84 ES MAYOR
                If vCta84D > vCta84H Then
                    ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge) " & _
                        " VALUES('" & gdHoraGrab & "','8321',0," & vDife8 & ",'3','" & vAgencia & "')"
                    oCon.ejecutar ssql

                    ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge) " & _
                        " VALUES('" & gdHoraGrab & "','" & lsCVMEPerdida & "'," & vDife8 & ",0,'3','" & vAgencia & "')"
                    oCon.ejecutar ssql
                Else
                    ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge) " & _
                        " VALUES('" & gdHoraGrab & "','8321'," & vDife8 & ",0,'3','" & vAgencia & "')"
                    oCon.ejecutar ssql

                    ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge) " & _
                        " VALUES('" & gdHoraGrab & "','" & lsCVMEPerdida & "',0," & vDife8 & ",'3','" & vAgencia & "')"
                    oCon.ejecutar ssql
                End If
            End If
        End If
        '**************************************************************************************
        'Dim vCta82D As Currency, vCta82H As Currency, vCta81D As Currency, vCta81H As Currency
        'CUADRE DE CTAS DE ORDEN (SILVITA)   82 = 81
        vCta82D = 0: vCta82H = 0
        tmpSql = "SELECT abs(round(SUM(ndebe),2)) nDebe, abs(round(SUM(nhaber),2)) nHaber FROM AsientoDN " & _
            " WHERE ctipo = '3' AND datediff(dd, dFecha ,'" & gdHoraGrab & "') = 0 " & _
            " AND cCtaCnt like '82%'" & _
            " Group BY cCtaCnt "
        Set RegTmp = oCon.CargaRecordSet(tmpSql)
        If (RegTmp.BOF Or RegTmp.EOF) Then
            'No se encuentra ctas en Dolares o Tipo 3
        Else
            Do While Not RegTmp.EOF
                vCta82D = vCta82D + RegTmp!nDebe
                vCta82H = vCta82H + RegTmp!nHaber
                RegTmp.MoveNext
            Loop
        End If
        RegTmp.Close
        Set RegTmp = Nothing
        vCta82D = Abs(vCta82D)
        vCta82H = Abs(vCta82H)

        vCta81D = 0: vCta81H = 0
        tmpSql = "SELECT abs(round(SUM(ndebe),2)) nDebe, abs(round(SUM(nhaber),2)) nHaber FROM AsientoDN " & _
            " WHERE ctipo = '3' AND datediff(dd, dFecha ,'" & gdHoraGrab & "') = 0 " & _
            " AND cCtaCnt like '81%'" & _
            " Group BY cCtaCnt "
        Set RegTmp = oCon.CargaRecordSet(tmpSql)
        If (RegTmp.BOF Or RegTmp.EOF) Then
            'No se encuentra ctas en Dolares o Tipo 3
        Else
            Do While Not RegTmp.EOF
                vCta81D = vCta81D + RegTmp!nDebe
                vCta81H = vCta81H + RegTmp!nHaber
                RegTmp.MoveNext
            Loop
        End If
        RegTmp.Close
        Set RegTmp = Nothing
        vCta81D = Abs(vCta81D)
        vCta81H = Abs(vCta81H)

        'Dim nResult81 As Currency, nResult82 As Currency

        nResult81 = IIf(vCta81D > vCta81H, vCta81D - vCta81H, vCta81H - vCta81D)
        nResult82 = IIf(vCta82D > vCta82H, vCta82D - vCta82H, vCta82H - vCta82D)

        'Validación

        'PEAC 20151205
        Call Validacion

'        If nResult82 <> nResult81 Then
'            vDife8 = 0
'            vDife8 = IIf(nResult81 > nResult82, nResult81 - nResult82, nResult82 - nResult81)
'            'vDife8 = Abs((vCta82D + vCta82H) - (vCta81D + vCta81H))
'            If Abs(nResult82) > Abs(nResult81) Then
'                'SI 82 MAYOR
'                If vCta82D > vCta82H Then
'                    ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge) " & _
'                        " VALUES('" & gdHoraGrab & "','8221',0," & vDife8 & ",'3','" & vAgencia & "')"
'                    oCon.Ejecutar ssql
'
'                    ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge) " & _
'                        " VALUES('" & gdHoraGrab & "','" & lsCVMEGanacia & "'," & vDife8 & ",0,'3','" & vAgencia & "')"
'                    oCon.Ejecutar ssql
'                Else
'                    ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge) " & _
'                        " VALUES('" & gdHoraGrab & "','8221'," & vDife8 & ",0,'3','" & vAgencia & "')"
'                    oCon.Ejecutar ssql
'
'                    ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge) " & _
'                        " VALUES('" & gdHoraGrab & "','" & lsCVMEGanacia & "',0," & vDife8 & ",'3','" & vAgencia & "')"
'                    oCon.Ejecutar ssql
'                End If
'            Else
'                'SI 81 MAYOR
'                If vCta81D > vCta81H Then
'                    ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge) " & _
'                        " VALUES('" & gdHoraGrab & "','8221',0," & vDife8 & ",'3','" & vAgencia & "')"
'                    oCon.Ejecutar ssql
'
'                    ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge) " & _
'                        " VALUES('" & gdHoraGrab & "','" & lsCVMEPerdida & "'," & vDife8 & ",0,'3','" & vAgencia & "')"
'                    oCon.Ejecutar ssql
'                Else
'                    ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge) " & _
'                        " VALUES('" & gdHoraGrab & "','8221'," & vDife8 & ",0,'3','" & vAgencia & "')"
'                    oCon.Ejecutar ssql
'
'                    ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge) " & _
'                        " VALUES('" & gdHoraGrab & "','" & lsCVMEPerdida & "',0," & vDife8 & ",'3','" & vAgencia & "')"
'                    oCon.Ejecutar ssql
'                End If
'            End If
'        End If
'
        'FIN PEAC

        '**************************************************************************************

        'Dim vCta86D As Currency, vCta86H As Currency, vCta85D As Currency, vCta85H As Currency
        'CUADRE DE CTAS DE ORDEN (SIMON) 86 = 85
        vCta86D = 0: vCta86H = 0
        tmpSql = "SELECT abs(round(SUM(ndebe),2)) nDebe, abs(round(SUM(nhaber),2)) nHaber FROM AsientoDN " & _
            " WHERE ctipo = '3' AND datediff(dd, dFecha ,'" & gdHoraGrab & "') = 0 " & _
            " AND cCtaCnt like '86%'" & _
            " Group BY cCtaCnt "
        Set RegTmp = oCon.CargaRecordSet(tmpSql)
        If (RegTmp.BOF Or RegTmp.EOF) Then
            'No se encuentra ctas en Dolares o Tipo 3
        Else
            Do While Not RegTmp.EOF
                vCta86D = vCta86D + RegTmp!nDebe
                vCta86H = vCta86H + RegTmp!nHaber
                RegTmp.MoveNext
            Loop
        End If
        RegTmp.Close
        Set RegTmp = Nothing
        vCta86D = Abs(vCta86D)
        vCta86H = Abs(vCta86H)

        vCta85D = 0: vCta85H = 0
        tmpSql = "SELECT abs(round(SUM(ndebe),2)) nDebe, abs(round(SUM(nhaber),2)) nHaber FROM AsientoDN " & _
            " WHERE ctipo = '3' AND datediff(dd, dFecha ,'" & gdHoraGrab & "') = 0 " & _
            " AND cCtaCnt like '85%'" & _
            " Group BY cCtaCnt "
            
        tmpSql = "stp_sel_CuadreCtasAsientoDN '" & gdHoraGrab & "' "
            
        Set RegTmp = oCon.CargaRecordSet(tmpSql)
        If (RegTmp.BOF Or RegTmp.EOF) Then
            'No se encuentra ctas en Dolares o Tipo 3
        Else
            Do While Not RegTmp.EOF
                vCta85D = vCta85D + RegTmp!nDebe
                vCta85H = vCta85H + RegTmp!nHaber
                RegTmp.MoveNext
            Loop
        End If
        RegTmp.Close
        Set RegTmp = Nothing
        vCta85D = Abs(vCta85D)
        vCta85H = Abs(vCta85H)

        'Dim nResult85 As Currency, nResult86 As Currency

        nResult85 = IIf(vCta85D > vCta85H, vCta85D - vCta85H, vCta85H - vCta85D)
        nResult86 = IIf(vCta86D > vCta86H, vCta86D - vCta86H, vCta86H - vCta86D)

        '*** PEAC 20151216
        'Validación
        Call validacion86

'        If nResult86 <> nResult85 Then
'            vDife8 = 0
'            vDife8 = IIf(nResult85 > nResult86, nResult85 - nResult86, nResult86 - nResult85)
'            If Abs(nResult85) > Abs(nResult86) Then
'                'SI 85 MAYOR
'                If vCta85D > vCta85H Then
'                    'Modificado PASI20140513 TI-ERS004-2014
''                    ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge) " & _
''                        " VALUES('" & gdHoraGrab & "','8528',0," & vDife8 & ",'3','" & vAgencia & "')"
''                    oCon.Ejecutar ssql
''
''                    ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge) " & _
''                        " VALUES('" & gdHoraGrab & "','" & lsCVMEGanacia & "'," & vDife8 & ",0,'3','" & vAgencia & "')"
''                    oCon.Ejecutar ssql
'                    Call IngresarAsientoDN(gdHoraGrab, "8528", 0, vDife8, 3, vAgencia, , , , oCon.ConexionActiva)
'                    Call IngresarAsientoDN(gdHoraGrab, lsCVMEGanacia, vDife8, 0, 3, vAgencia, , , , oCon.ConexionActiva)
'                Else
'                    'Modificado PASI20140513 TI-ERS004-2014
''                    ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge) " & _
''                        " VALUES('" & gdHoraGrab & "','8528'," & vDife8 & ",0,'3','" & vAgencia & "')"
''                    oCon.Ejecutar ssql
''
''                    ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge) " & _
''                        " VALUES('" & gdHoraGrab & "','" & lsCVMEGanacia & "',0," & vDife8 & ",'3','" & vAgencia & "')"
''                    oCon.Ejecutar ssql
'                    Call IngresarAsientoDN(gdHoraGrab, "8528", vDife8, 0, 3, vAgencia, , , , oCon.ConexionActiva)
'                    Call IngresarAsientoDN(gdHoraGrab, lsCVMEGanacia, 0, vDife8, 3, vAgencia, , , , oCon.ConexionActiva)
'                    'end PASI
'                End If
'            Else
'                'SI 86 MAYOR
'                If vCta86D > vCta86H Then
'                'Modificado PASI20140513 TI-ERS004-2014
''                    ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge) " & _
''                        " VALUES('" & gdHoraGrab & "','8528',0," & vDife8 & ",'3','" & vAgencia & "')"
''                    oCon.Ejecutar ssql
''
''                    ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge) " & _
''                        " VALUES('" & gdHoraGrab & "','" & lsCVMEPerdida & "'," & vDife8 & ",0,'3','" & vAgencia & "')"
''                    oCon.Ejecutar ssql
'                    Call IngresarAsientoDN(gdHoraGrab, "8528", 0, vDife8, 3, vAgencia, , , , oCon.ConexionActiva)
'                    Call IngresarAsientoDN(gdHoraGrab, lsCVMEPerdida, vDife8, 0, 3, vAgencia, , , , oCon.ConexionActiva)
'
'                'end PASI
'                Else
'                'Modificado PASI20141305 TI-ERS004-2014
''                    ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge) " & _
''                        " VALUES('" & gdHoraGrab & "','8528'," & vDife8 & ",0,'3','" & vAgencia & "')"
''                    oCon.Ejecutar ssql
''
''                    ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge) " & _
''                        " VALUES('" & gdHoraGrab & "','" & lsCVMEPerdida & "',0," & vDife8 & ",'3','" & vAgencia & "')"
''                    oCon.Ejecutar ssql
'                    Call IngresarAsientoDN(gdHoraGrab, "8528", vDife8, 0, 3, vAgencia, , , , oCon.ConexionActiva)
'                    Call IngresarAsientoDN(gdHoraGrab, lsCVMEPerdida, 0, vDife8, 3, vAgencia, , , , oCon.ConexionActiva)
'                    'end PASI
'                End If
'            End If
'        End If

'**** fin validacion86

        RegTran.Close
        Set RegTran = Nothing
'        prgList.Visible = False
'        prgList.value = 0
    End If

    'Rutina para obtener asientos de la Compra y Venta de Dolares
    Dim CVDebe As Currency, CVHaber As Currency

    'PEAC 20180927
    'tmpSql = " SELECT M.nMovNro, M.cMovNro, cOpeCod, nMovImporte AS Monto, MTC.nMovTpoCambio nTipCambio, " _
           & " SUBSTRING(M.cMovNro,18,2) Agencia, Convert(Decimal(20,2),nMovImporte*nMovTpoCambio) AS NMOVSOLES  FROM Mov M" _
           & " Inner Join MovCompraVenta MCV On M.nMovNro = MCV.nMovNro" _
           & " Inner Join MovTpoCambio MTC On M.nMovNro = MTC.nMovNro" _
           & " WHERE cOpeCod IN (Select cOpeCod From OpeTpoSinAsiento Where nProceso = " & gAsientoProcesoSiAsientoCVME & ")  AND (M.nMovFlag = " & MovFlag.gMovFlagVigente & ")" _
           & " And M.cMovNro Like '" & Format(ldFechaAsiento, gsFormatoMovFecha) & "%' AND MCV.nMovImporte <> 0 ORDER BY cOpeCod"
           
    'PEAC 20180927
    tmpSql = "stp_sel_ObtieneCompraVentaMEAsientoDN " & gAsientoProcesoSiAsientoCVME & "," & MovFlag.gMovFlagVigente & ",'" & Format(ldFechaAsiento, gsFormatoMovFecha) & "' "

    'tmpSql = "exec stp_sel_ObtieneMovParaOpeCompraVenta " & gAsientoProcesoSiAsientoCVME & "," & MovFlag.gMovFlagVigente & " ,' " & Format(ldFechaAsiento, gsFormatoMovFecha) & " '"  'JUCS20170918

    Set RegTmp = oCon.CargaRecordSet(tmpSql)
    If (RegTmp.BOF Or RegTmp.EOF) Then
        RegTmp.Close
        Set RegTmp = Nothing
    Else
        With RegTmp
            Do While Not .EOF
                ssql = " SELECT cCtaContCod , cOpeCtaDH  FROM OpeCtaNeg  " _
                     & " WHERE cOpeCod = '" & !cOpeCod & "'"
                Set RegOpeCta = oCon.CargaRecordSet(ssql)
                If (RegOpeCta.BOF Or RegOpeCta.EOF) Then
                    RegOpeCta.Close
                    Set RegOpeCta = Nothing
                    MsgBox " Código de Operación No Reconocido " & !cOpeCod, vbInformation, " Aviso "
                    'vNoCtaCnt = vNoCtaCnt & oImpresora.gPrnSaltoLinea & "     * " & !cOpeCod & " operación compra/venta no reconocida."
                    vNoCtaCnt = vNoCtaCnt & vbCrLf & "     * " & !cOpeCod & " operación compra/venta no reconocida."
                Else
                    Do While Not RegOpeCta.EOF
                        vCodConta = Trim(RegOpeCta!cCtaContCod)
                        vAG = VarAG(Mid(RegTmp!cMovNro, 15, 5), vCodConta)
                        vCodConta = Replace(vCodConta, "AG", vAG, 1, 1, vbTextCompare)
                        vParche = AsientoParche(vCodConta, True, oCon.ConexionActiva)
                        vAgencia = RegTmp!Agencia
                        If Len(vParche) > 0 Then
                            vCodConta = vParche
                        End If
                      '**************************************************************************
                      'Si existe lo actualiza; caso contrario lo agrega
                      'MARCA 12
                        'If Trim(vCodConta) = "19180207" Then Stop
                        If RegOpeCta!cOpeCtaDH = "D" Then
                            If Left(vCodConta, 4) = lsCajaSoles Or Left(vCodConta, 4) = "2514" Then 'APRI20170705 RFC1705220001 Add Left(vCodConta, 4) = "2514"
                                'sSql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge, nMovNro, cOpeCod) " & _
                                    " VALUES('" & gdHoraGrab & "','" & vCodConta & "'," & Format(!Monto * !nTipCambio, "#0.00") & ",0,'1','" & vAgencia & "'," & !nMovNro & ",'" & !cOpeCod & "' )"

                                'Modificado PASI20140513 TI-ERS004-2014
                                'ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge, nMovNro, cOpeCod) " & _
                                '    " VALUES('" & gdHoraGrab & "','" & vCodConta & "'," & !NMOVSOLES & ",0,'1','" & vAgencia & "'," & !nMovNro & ",'" & !cOpecod & "' )"
                                'oCon.Ejecutar ssql

                                'Call IngresarAsientoDN(gdHoraGrab, vCodConta, !NMOVSOLES, 0, 1, vAgencia, !nMovNro, !cOpeCod, , oCon.ConexionActiva, TCF)
                                Call IngresarAsientoDN(gdHoraGrab, vCodConta, Format(!NMOVSOLES, "#0.00"), 0, 1, vAgencia, !nMovNro, !cOpeCod, , oCon.ConexionActiva, TCF) 'Modificación JUCS 25012018 Agregó Format

                                'end PASI


                                'CVDebe = CVDebe + Format(!Monto * !nTipCambio, "#0.00")
                                CVDebe = CVDebe + Format(!NMOVSOLES, "#0.00")


                            ElseIf Left(vCodConta, 4) = lsCajaDolares Or Left(vCodConta, 4) = "2524" Then 'Or Left(vCodConta, 4) = "2825"
                                'Modificado PASI TI-ERS004-2014
                                'ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge, nMovNro, cOpeCod) " & _
                                '   " VALUES('" & gdHoraGrab & "','" & vCodConta & "'," & Round(!Monto * TCF, 2) & ",0,'1','" & vAgencia & "'," & !nMovNro & ",'" & !cOpecod & "')"
                                'oCon.Ejecutar ssql

                                 'Call IngresarAsientoDN(gdHoraGrab, vCodConta, Round(!Monto * TCF, 2), 0, 1, vAgencia, !nMovNro, !cOpeCod, , oCon.ConexionActiva, TCF) 'Cambiado RegTran!nMovNro >> !nMovNro ; RegTran!cOpecod>>!cOpecod

                                 Call IngresarAsientoDN(gdHoraGrab, vCodConta, Format(!Monto * TCF, "#0.00"), 0, 1, vAgencia, !nMovNro, !cOpeCod, , oCon.ConexionActiva, TCF) 'Cambiado RegTran!nMovNro >> !nMovNro ; RegTran!cOpecod>>!cOpecod
                                 'Modificacion JUCS 25012018 Agregó Format



                                'end PASI

                                'Modificado PASI TI-ERS004-2014
                                'ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge, nMovNro, cOpeCod) " & _
                                '   " VALUES('" & gdHoraGrab & "','" & vCodConta & "'," & !Monto & ",0,'2','" & vAgencia & "'," & !nMovNro & ",'" & !cOpecod & "')"
                                'oCon.Ejecutar ssql

                                'Call IngresarAsientoDN(gdHoraGrab, vCodConta, !Monto, 0, 2, vAgencia, !nMovNro, !cOpeCod, , oCon.ConexionActiva, TCF) 'Cambiado RegTran!nMovNro >> !nMovNro ; RegTran!cOpecod>>!cOpecod
                                Call IngresarAsientoDN(gdHoraGrab, vCodConta, Format(!Monto, "#0.00"), 0, 2, vAgencia, !nMovNro, !cOpeCod, , oCon.ConexionActiva, TCF) 'Cambiado RegTran!nMovNro >> !nMovNro ; RegTran!cOpecod>>!cOpecod
                                'Modificación JUCS 25012018 Agregó Format



                                'end PASI

                                'CVDebe = CVDebe + Round(!Monto * TCF, 2)
                                CVDebe = CVDebe + Format(!Monto * TCF, "#0.00") 'Modificación JUCS 25012018 Agregó Format


                            Else
                                MsgBox "Cuenta en Compra y Venta no reconocida", vbInformation, "Aviso"
                                'vNoCtaCnt = vNoCtaCnt & oImpresora.gPrnSaltoLinea & "     * " & vCodConta & " cta. de compra/venta no reconocida."
                                vNoCtaCnt = vNoCtaCnt & vbCrLf & "     * " & vCodConta & " cta. de compra/venta no reconocida."

                            End If
                        ElseIf RegOpeCta!cOpeCtaDH = "H" Then
                            If Left(vCodConta, 4) = lsCajaSoles Or Left(vCodConta, 4) = "2514" Then
                                'Modificado PASI TI-ERS004-2014
'                                ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge, nMovNro, cOpeCod) " & _
'                                    " VALUES('" & gdHoraGrab & "','" & vCodConta & "',0," & Round(!NMOVSOLES, 2) & ",'1','" & vAgencia & "'," & !nMovNro & ",'" & !cOpecod & "')"
'                                OCon.Ejecutar ssql

                                'Call IngresarAsientoDN(gdHoraGrab, vCodConta, 0, Round(!NMOVSOLES, 2), 1, vAgencia, !nMovNro, !cOpeCod, , oCon.ConexionActiva, TCF) 'Cambiado RegTran!nMovNro >> !nMovNro ; RegTran!cOpecod>>!cOpecod
                                Call IngresarAsientoDN(gdHoraGrab, vCodConta, 0, Format(!NMOVSOLES, "#0.00"), 1, vAgencia, !nMovNro, !cOpeCod, , oCon.ConexionActiva, TCF) 'Cambiado RegTran!nMovNro >> !nMovNro ; RegTran!cOpecod>>!cOpecod
                                'Modificación JUCS 25012018 Agregó Format

                                'end PASI

                                'CVHaber = CVHaber + Round(!Monto * !nTipCambio, 2)
                                CVHaber = CVHaber + Format(!Monto * !nTipCambio, "#0.00") 'Modificación JUCS 25012018 Agregó Format


                            'ElseIf Left(vCodConta, 4) = lsCajaDolares Or Left(vCodConta, 4) = "2524" Then 'Or Left(vCodConta, 4) = "2825" 'APRI20170705 RFC1705220001 Add Left(vCodConta, 4) = "2524"
                            ElseIf Left(vCodConta, 4) = lsCajaDolares Or Left(vCodConta, 4) = "2524" Then 'Or Left(vCodConta, 4) = "2825" 'APRI20170705 RFC1705220001 Add Left(vCodConta, 4) = "2524"
                                'Modificado PASI TI-ERS004-2014
'                                ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge, nMovNro, cOpeCod) " & _
'                                    " VALUES('" & gdHoraGrab & "','" & vCodConta & "',0," & Round(!Monto * TCF, 2) & ",'1','" & vAgencia & "'," & !nMovNro & ",'" & !cOpecod & "')"
'                                OCon.Ejecutar ssql

                                'Call IngresarAsientoDN(gdHoraGrab, vCodConta, 0, Round(!Monto * TCF, 2), 1, vAgencia, !nMovNro, !cOpeCod, , oCon.ConexionActiva, TCF) 'Cambiado RegTran!nMovNro >> !nMovNro ; RegTran!cOpecod>>!cOpecod
                                Call IngresarAsientoDN(gdHoraGrab, vCodConta, 0, Format(!Monto * TCF, "#0.00"), 1, vAgencia, !nMovNro, !cOpeCod, , oCon.ConexionActiva, TCF) 'Cambiado RegTran!nMovNro >> !nMovNro ; RegTran!cOpecod>>!cOpecod
                                'Modificación JUCS 25012018 Agregó Format


                                'end PASI

                                'Modificado PASI TI-ERS004-2014
'                                ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge, nMovNro, cOpeCod) " & _
'                                    " VALUES('" & gdHoraGrab & "','" & vCodConta & "',0," & (!Monto) & ",'2','" & vAgencia & "'," & !nMovNro & ",'" & !cOpecod & "')"
'                                OCon.Ejecutar ssql

                                Call IngresarAsientoDN(gdHoraGrab, vCodConta, 0, !Monto, 2, vAgencia, !nMovNro, !cOpeCod, , oCon.ConexionActiva, TCF)  'Cambiado RegTran!nMovNro >> !nMovNro ; RegTran!cOpecod>>!cOpecod

                                'end PASI

                                'CVHaber = CVHaber + Round(!Monto * TCF, 2)
                                CVHaber = CVHaber + Format(!Monto * TCF, "#0.00")   'ModificaciÃ³n JUCS 25012018 AgregÃ³ Format

                            Else
                                MsgBox "Cuenta en Compra y Venta no reconocida", vbInformation, "Aviso"
                                'vNoCtaCnt = vNoCtaCnt & oImpresora.gPrnSaltoLinea & "     * " & vCodConta & " cta. de compra/venta no reconocida."
                                vNoCtaCnt = vNoCtaCnt & vbCrLf & "     * " & vCodConta & " cta. de compra/venta no reconocida."

                            End If
                        Else
                            MsgBox " Operación no reconocida al Insertar " & vbCr & _
                            " operación nro.: " & vCodConta, vbInformation, " Aviso "
                            'vNoCtaCnt = vNoCtaCnt & oImpresora.gPrnSaltoLinea & "     * " & RegOpeCta!cOpeCtaDH & " debe/haber no reconocido (compra/venta)."
                            vNoCtaCnt = vNoCtaCnt & vbCrLf & "     * " & RegOpeCta!cOpeCtaDH & " debe/haber no reconocido (compra/venta)."

                        End If
                      RegOpeCta.MoveNext
                    Loop
                    RegOpeCta.Close
                    Set RegOpeCta = Nothing
                End If
                .MoveNext
            Loop
        End With
        RegTmp.Close
        Set RegTmp = Nothing
        'Validación de 0.01 por dolares - Silvita
        Dim nCtaDeb1 As Currency, nCtaHab1 As Currency

        'PEAC 20180927
        'tmpSql = "SELECT cCodAge, sum(round(ndebe,2)) nDebe , sum(round(nhaber,2)) nHaber " & _
            " FROM AsientodN " & _
            " WHERE cTipo = '1' AND datediff(dd, dFecha ,'" & gdHoraGrab & "') = 0 " & _
            " GROUP BY cCodAge, cCtaCnt Order by 1 "
            
        'PEAC 20180927
         tmpSql = "stp_sel_CuadreCtasAsientoDN2 '" & gdHoraGrab & "' "
            
        Set RegTmp = oCon.CargaRecordSet(tmpSql)
        If (RegTmp.BOF Or RegTmp.EOF) Then
            'No se encuentra ctas en Dolares o Tipo 1
        Else
            vAgencia = RegTmp!cCodAge
            Do While Not RegTmp.EOF
                Do While RegTmp!cCodAge = vAgencia
                    nCtaDeb1 = nCtaDeb1 + RegTmp!nDebe
                    nCtaHab1 = nCtaHab1 + RegTmp!nHaber
                    RegTmp.MoveNext
                    If RegTmp.EOF Then Exit Do
                Loop
                vCtaDolar = nCtaDeb1 - nCtaHab1


                If vCtaDolar > 0 Then
                    ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge) " & _
                        " VALUES('" & gdHoraGrab & "','" & lsCVMEGanacia & vAgencia & "',0," & Abs(vCtaDolar) & ",'1','" & vAgencia & "')"
                    oCon.ejecutar ssql
                ElseIf vCtaDolar < 0 Then
                    ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge) " & _
                        " VALUES('" & gdHoraGrab & "','" & lsCVMEPerdida & vAgencia & "'," & Abs(vCtaDolar) & ",0,'1','" & vAgencia & "')"
                    oCon.ejecutar ssql
                End If
                If RegTmp.EOF Then Exit Do
                vAgencia = RegTmp!cCodAge
                nCtaDeb1 = 0: nCtaHab1 = 0
            Loop
        End If
        RegTmp.Close
        Set RegTmp = Nothing
    End If

    'Asiento para sobrantes y faltantes
    Dim vCodConta1 As String, vCodConta2 As String
    'Add by gitu 16/09/2008
    Dim lsMonDebe As String, lsMonHaber As String, lsMonDebeD As String, lsMonHaberD As String
    'End Gitu

    'Mod. PASI20140429
'    tmpSql = "Select SUBSTRING(M.cMovNro,18,2) Agencia, ABS(O.nMovImporte) As Monto, M.cOpeCod as cCodOpe, O.nMoneda FROM Mov M " _
'            & "Inner Join MovOpeVarias O On M.nMovNro = O.nMovNro " _
'            & "Where M.cMovNro like '" & Format(ldFechaAsiento, gsFormatoMovFecha) & "%' And M.cOpeCod " _
'            & "IN (Select cOpeCod From OpeTpoSinAsiento Where nProceso = " & gAsientoProcesoSiAsientoSofFal & ")  and (M.nMovFlag = 0) AND nMovImporte <> 0"
     tmpSql = "stp_sel_ObtenerFaltantesSobrantesparaAsientoDN '" & Format(ldFechaAsiento, gsFormatoMovFecha) & "%' "
    'end PASI

    Set RegTmp = oCon.CargaRecordSet(tmpSql)
    If (RegTmp.BOF Or RegTmp.EOF) Then
        RegTmp.Close
        Set RegTmp = Nothing
    Else
        With RegTmp
            Do While Not .EOF
                'Variables para cambio en el CodContable
                ssql = ""
                vMoneda = !nmoneda
                vAgencia = !Agencia
                'If !Agencia = "07" Then Stop

                If Not (vMoneda = Moneda.gMonedaNacional Or vMoneda = Moneda.gMonedaExtranjera) Then
                    MsgBox " Código Errado, Moneda no definida " & !cCtaCod, vbInformation, " Aviso "
                    'vNoCtaCnt = vNoCtaCnt & oImpresora.gPrnSaltoLinea & "     * " & !cCtaCod & " moneda no definida(sobrantes/faltantes)."
                    vNoCtaCnt = vNoCtaCnt & vbCrLf & "     * " & !cCtaCod & " moneda no definida(sobrantes/faltantes)."
                End If
                If !cCodOpe = "901020" Or !cCodOpe = "901010" Or !cCodOpe = "901061" Or !cCodOpe = "300200" Then      ' Faltante de caja
                    vCodConta1 = "11M102AG"
                    'ARCV 31-03-2007
                    'vCodConta2 = IIf(!Monto > 1, "19M80202AG", "63M109090998")

                    'Comentado por gitu 11-08-2008 por que solo debe quedar con la cuenta 19
                    'vCodConta2 = IIf(!Monto > 1, "19M802AG", "42M229AG")
                    vCodConta2 = IIf(!cCodOpe = "901061", "29M201AG", "19M802AG") 'AGREGADO POR VAPA ERS035-2017
                    'Fin Gitu

                    '---------
                    vParche = AsientoParche(vCodConta2, True, oCon.ConexionActiva)
                    If Len(vParche) > 0 Then
                        vCodConta2 = vParche
                    End If
                    'Cambios en Moneda y Agencia
                    vCodConta1 = Replace(vCodConta1, "M", vMoneda, 1, 1, vbTextCompare)

                    'Modificado PASI20140429
                    'vAG = VarAG(gsCodCMAC & !Agencia, vCodConta1)
                    vAG = IIf(!cCodOpe = "300200", VarAG(gsCodCMAC & !AgenciaO, vCodConta1), VarAG(gsCodCMAC & !Agencia, vCodConta1))
                    'end PASI

                    vCodConta1 = Trim(Replace(vCodConta1, "AG", vAG, 1, 1, vbTextCompare))
                    vCodConta2 = Replace(vCodConta2, "M", vMoneda, 1, 1, vbTextCompare)
                    vCodConta2 = Trim(Replace(vCodConta2, "AG", !Agencia, 1, 1, vbTextCompare))

                    vParche = AsientoParche(vCodConta1, True, oCon.ConexionActiva)

                    If Len(vParche) > 0 Then vCodConta1 = vParche

                    'Cometado porque genera mal el asiento GITU 15-09-2008

                    'Comentado PASI TI-ERS004-2014
'                    If !cCodOpe = "300200" Then
'                        lsMonDebe = !Monto & ",0"
'                        lsMonDebeD = (!Monto * TCF) & ",0"
'                        lsMonHaber = " 0," & !Monto
'                        lsMonHaberD = " 0," & (!Monto * TCF)
'                    Else
'                        lsMonDebe = " 0," & !Monto
'                        lsMonDebeD = " 0," & (!Monto * TCF)
'                        lsMonHaber = !Monto & ",0 "
'                        lsMonHaberD = (!Monto * TCF) & ",0 "
'                    End If
                    'end PASI

                    'Modificado PASI TI-ERS004-2014
'                    ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge) " & _
'                           " VALUES('" & gdHoraGrab & "','" & vCodConta1 & "'," & lsMonDebe & ",'0','" & IIf(!cCodOpe = "300200", !AgenciaO, vAgencia) & "')" 'Mod. PASI20140409
'                    OCon.Ejecutar ssql
                    Call IngresarAsientoDN(gdHoraGrab, vCodConta1, IIf(!cCodOpe = "300200", !Monto, 0), IIf(!cCodOpe = "300200", 0, !Monto), 0, IIf(!cCodOpe = "300200", !AgenciaO, vAgencia), , !cCodOpe, , oCon.ConexionActiva, TCF)
                    'end PASI

                    If Mid(vCodConta1, 3, 1) = Moneda.gMonedaExtranjera Then

                    'Modificado PASI TI-ERS004-2014
'                    ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge) " & _
'                            " VALUES('" & gdHoraGrab & "','" & vCodConta1 & "'," & lsMonDebeD & ",'3','" & IIf(!cCodOpe = "300200", !AgenciaO, vAgencia) & "')" 'Mod. PASI20140409
'                    OCon.Ejecutar ssql
                    Call IngresarAsientoDN(gdHoraGrab, vCodConta1, IIf(!cCodOpe = "300200", (!Monto * TCF), 0), IIf(!cCodOpe = "300200", 0, (!Monto * TCF)), 3, IIf(!cCodOpe = "300200", !AgenciaO, vAgencia), , !cCodOpe, , oCon.ConexionActiva, TCF)
                    'End PASI

                    End If

                    'Modificado PASI TI-ERS004-2014
                    Call IngresarAsientoDN(gdHoraGrab, vCodConta2, IIf(!cCodOpe = "300200", 0, !Monto), IIf(!cCodOpe = "300200", !Monto, 0), 0, vAgencia, , !cCodOpe, , oCon.ConexionActiva, TCF)
'                    ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge) " & _
'                           " VALUES('" & gdHoraGrab & "','" & vCodConta2 & "'," & lsMonHaber & ",'0','" & vAgencia & "')"
'                    OCon.Ejecutar ssql
                    'END PASI

                    If Mid(vCodConta2, 3, 1) = Moneda.gMonedaExtranjera Then

                       'Modificado PASI TI-ERS004-2014
'                        ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge) " & _
'                               " VALUES('" & gdHoraGrab & "','" & vCodConta2 & "'," & lsMonHaberD & ",'3','" & vAgencia & "')"
'                        OCon.Ejecutar ssql
                    Call IngresarAsientoDN(gdHoraGrab, vCodConta2, IIf(!cCodOpe = "300200", 0, (!Monto * TCF)), IIf(!cCodOpe = "300200", (!Monto * TCF), 0), 3, vAgencia, , !cCodOpe, , oCon.ConexionActiva, TCF)  'Se Cambio Moneda de 0 a 3 PASI20140801
                    End If
                   'End GITU
                Else
                    vCodConta1 = "11M102AG"
                    'ARCV 31-03-2007
                    'vCodConta2 = IIf(!Monto > 5, "29M201AG", "52M2290909AG")

                    'Comentado por gitu 11-08-2008 por que solo debe quedar con la cuenta 29
                    'vCodConta2 = IIf(!Monto > 5, "29M201AG", "52M229AG")
                    vCodConta2 = "29M201AG"
                    'Fin Gitu

                    '---------
                    vParche = AsientoParche(vCodConta2, True, oCon.ConexionActiva)
                    If Len(vParche) > 0 Then vCodConta2 = vParche

                    'Cambios en Moneda y Agencia
                    vCodConta1 = Replace(vCodConta1, "M", vMoneda, 1, 1, vbTextCompare)
                    vAG = VarAG(gsCodCMAC & !Agencia, vCodConta1)
                    vCodConta1 = Trim(Replace(vCodConta1, "AG", vAG, 1, 1, vbTextCompare))
                    vCodConta2 = Replace(vCodConta2, "M", vMoneda, 1, 1, vbTextCompare)
                    vCodConta2 = Trim(Replace(vCodConta2, "AG", !Agencia, 1, 1, vbTextCompare))

                    vParche = AsientoParche(vCodConta1, True, oCon.ConexionActiva)
                    If Len(vParche) > 0 Then vCodConta1 = vParche

                    'Verifica si el asiento ya fue creado
                    'Si existe lo actualiza; caso contrario lo agrega
                        ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge) " & _
                            " VALUES('" & gdHoraGrab & "','" & vCodConta1 & "'," & !Monto & ",0,'0','" & vAgencia & "')"
                        oCon.ejecutar ssql
                        If Mid(vCodConta1, 3, 1) = Moneda.gMonedaExtranjera Then
                            ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge) " & _
                                " VALUES('" & gdHoraGrab & "','" & vCodConta1 & "'," & (!Monto * TCF) & ",0,'3','" & vAgencia & "')"
                            oCon.ejecutar ssql
                        End If
                        ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge) " & _
                            " VALUES('" & gdHoraGrab & "','" & vCodConta2 & "',0," & !Monto & ",'0','" & vAgencia & "')"
                        oCon.ejecutar ssql
                        If Mid(vCodConta2, 3, 1) = Moneda.gMonedaExtranjera Then
                            ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge) " & _
                                " VALUES('" & gdHoraGrab & "','" & vCodConta2 & "',0," & (!Monto * TCF) & ",'3','" & vAgencia & "')"
                            oCon.ejecutar ssql
                        End If
                End If
                .MoveNext
            Loop
        End With
        RegTmp.Close
        Set RegTmp = Nothing
    End If

    '-- Aca
       'Verifica que cuadre DEBE y HABER - Tipo '3'
        Dim vDife As Currency, vDifeSol As Currency, vMonDebe As Currency, vMonHaber As Currency
        Dim vMonDebeSol As Currency, vMonHaberSol As Currency
        'Carga suma de tipo soles para hallar la diferencia - RAUL

        'PEAC 20180927
        'tmpSql = "SELECT round(SUM(ndebe),2) AS MonDebe , round(SUM(nhaber),2) AS MonHaber FROM AsientoDN " & _
            " WHERE ctipo = '0' AND substring(cCtaCnt,3,1) = '" & Moneda.gMonedaNacional & "' AND datediff(dd, dFecha ,'" & gdHoraGrab & "') = 0 " & _
            " Group BY cCtaCnt "

        'tmpSql = "SELECT round(SUM(ndebe),2) AS MonDebe , round(SUM(nhaber),2) AS MonHaber FROM AsientoDN " & _
            " WHERE ctipo = '0'  and cOpeCod not in ('900022','900023') AND substring(cCtaCnt,3,1) = '" & Moneda.gMonedaNacional & "' AND datediff(dd, dFecha ,'" & gdHoraGrab & "') = 0 " & _
            " Group BY cCtaCnt " 'JUCS TI-ERS 002-2017

        'PEAC 20180927
        tmpSql = "stp_sel_CuadreCtasAsientoDN3 '" & gdHoraGrab & "','" & Moneda.gMonedaNacional & "' "

        vMonDebeSol = 0: vMonHaberSol = 0: vMonDebe = 0: vMonHaber = 0
        Set RegTmp = oCon.CargaRecordSet(tmpSql)
        If (RegTmp.BOF Or RegTmp.EOF) Then
            'No se encuentra ctas en Dolares o Tipo 3
        Else
            Do While Not RegTmp.EOF
                vMonDebeSol = vMonDebeSol + RegTmp!mondebe
                vMonHaberSol = vMonHaberSol + RegTmp!monhaber
                RegTmp.MoveNext
            Loop
        End If
        RegTmp.Close
        Set RegTmp = Nothing

        'Halla suma de diferencia de Dolares
        vDifeSol = vMonDebeSol - vMonHaberSol

        'PEAC 20180927
        tmpSql = "SELECT sUM(round(ndebe,2) - round(nhaber,2)) AS DiffDolares, cCodAge FROM AsientoDN " & _
            " WHERE ctipo = '3' AND datediff(dd, dFecha ,'" & gdHoraGrab & "') = 0 " & _
            " Group BY cCodAge "

         'tmpSql = "SELECT sUM(round(ndebe,2) - round(nhaber,2)) AS DiffDolares, cCodAge FROM AsientoDN " & _
            " WHERE ctipo = '3' and cOpeCod not in ('900022','900023') AND datediff(dd, dFecha ,'" & gdHoraGrab & "') = 0 " & _
            " Group BY cCodAge " 'JUCS TI-ERS 002-2017

        'PEAC 20180927
        tmpSql = "stp_sel_BuscaDifDolaresAsientoDN '" & gdHoraGrab & "' "

        Set RegTmp = oCon.CargaRecordSet(tmpSql)
        If (RegTmp.BOF Or RegTmp.EOF) Then
            'No se encuentra ctas en Dolares o Tipo 3
        Else
            Do While Not RegTmp.EOF
                vAgencia = RegTmp!cCodAge
                vDife = RegTmp!DiffDolares
                If vDifeSol <> 0 Then
                    vDife = vDife + vDifeSol
                    vDifeSol = 0
                End If

                vCtaDolar = -1 * vDife
                If vCtaDolar < 0 Then
                    'ARCV 31-03-2007
                     ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge) " & _
                        " VALUES('" & gdHoraGrab & "','" & "51280101" & vAgencia & "',0," & Abs(vCtaDolar) & ",'3','" & vAgencia & "')"
                    '----------
                    oCon.ejecutar ssql
                ElseIf vCtaDolar > 0 Then
                    'ARCV 31-03-2007
                    ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge) " & _
                        " VALUES('" & gdHoraGrab & "','" & "41280101" & vAgencia & "'," & Abs(vCtaDolar) & ",0,'3','" & vAgencia & "')"
                    '-------
                    oCon.ejecutar ssql
                End If

                RegTmp.MoveNext
            Loop
        End If
        RegTmp.Close
        Set RegTmp = Nothing

    'dbCmact.CommitTrans  'Finaliza TRANSACCION
    '*************************************************************************************
    '*************************************************************************************
    'CONVERSION DE ITF DE DOLARES A SOLES

    'GeneraAjuste ldFechaAsiento, TCCPonderado, oCon.ConexionActiva

    '*************************************************************************************
    '*************************************************************************************



    Me.Caption = "Generación de Asientos Contables"
    'GENERA IMPRESION DE LOS ASIENTOS
    Dim vIndice As Long, vLineas As Long, vPage As Long
    Dim vDolar As Boolean
    Dim vSumDebe As Currency, vSumHaber As Currency
    Dim vMoneAnte As String, sSqlDolar As String, vCabecera As String
    vDolar = False
    Dim vDiaSinOpe  As String
    'VALIDACION DE ASIENTOS
    'Dim vObs As String
    'vCabecera = "NRO. DE CUENTA                                 DESCRIPCION                                          DEBE          HABER" & oImpresora.gPrnSaltoLinea
    vCabecera = "NRO. DE CUENTA                                 DESCRIPCION                                          DEBE          HABER" & vbCrLf

    Dim rsAge As ADODB.Recordset
    Dim oAre As COMDConstantes.DCOMActualizaDatosArea
    Set rsAge = New ADODB.Recordset
    Set oAre = New COMDConstantes.DCOMActualizaDatosArea
    Set rsAge = oAre.getAgencias(, False)
    vObs = ""
    Do While Not rsAge.EOF
        Me.Caption = "Validando asientos de :" & Trim(rsAge!Descripcion)
        vObs = vObs & ValidaOk2(txtFecha, True, rsAge!codigo, oCon.ConexionActiva)
        rsAge.MoveNext
        DoEvents
    Loop

    Call BuscaCtasCntEnPlanCon

'    '*********** PEAC 20110705 - VERIFICA EXISTENCIA DE CTAS CONT EN EL PLAN CONT.
'    Dim rsCtas As ADODB.Recordset
'    Set rsCtas = New ADODB.Recordset
'    ssql = "exec stp_sel_BuscaCtaContPlan '" & Format(txtFecha, "yyyymmdd") & "'"
'    Set rsCtas = oCon.CargaRecordSet(ssql)
'    If Not (rsCtas.BOF Or rsCtas.EOF) Then
'        Do While Not rsCtas.EOF
'        'ALPA 20130527***********************************
'        'vObs = vObs & oImpresora.gPrnSaltoLinea & "** OJO ** Cta.Inexist.Plan Cont>> Cta.Cont:'" & rsCtas!cCtaCnt & "' Importe: '" & Trim(CStr(rsCtas!Importe)) & "' Age:'" & Trim(rsCtas!cCodAge) & "' Ope:'" & Trim(rsCtas!cOpecod) & "' Cod.Cta.'" & IIf(IsNull(rsCtas!cCtaCod), "", rsCtas!cCtaCod) & "'"
'        If rsCtas!cOpecod <> "700104" Then
'            vObs = vObs & oImpresora.gPrnSaltoLinea & "** OJO ** Cta.Inexist.Plan Cont>> Cta.Cont:'" & rsCtas!cctacnt & "' Importe: '" & Trim(CStr(rsCtas!Importe)) & "' Age:'" & Trim(rsCtas!cCodAge) & "' Ope:'" & Trim(rsCtas!cOpecod) & "' Cod.Cta.'" & IIf(IsNull(rsCtas!cCtaCod), "", rsCtas!cCtaCod) & "'"
'        End If
'        '************************************************
'        rsCtas.MoveNext
'        Loop
'    End If
'    rsCtas.Close
'
'    '*********** FIN PEAC


    ejecutar ("Play ") 'PEAC 20160701

'    Call VerificaVeracidadAsntoCnt

    'Verifica la veracidad del asiento contable
    Me.Caption = "Verificando Veracidad de asiento Contable"
    If Len(Trim(vObs)) > 0 Or Len(Trim(vNoCtaCnt)) > 0 Then

   '2020 peac comentar msgbox
        MsgBox "El asiento presenta Observaciones ! ", vbInformation, " Aviso "
        
        'vProduc = oImpresora.gPrnSaltoLinea & Space(40) & "  A S I E N T O     ¡   N O     V A L I D O   !" & oImpresora.gPrnSaltoLinea
        vProduc = vbCrLf & Space(40) & "  A S I E N T O     ¡   N O     V A L I D O   !" & vbCrLf

        'vProduc = vProduc & "   OBSERVACIONES DEL ASIENTO : " & oImpresora.gPrnSaltoLinea & "  ============================="
        vProduc = vProduc & "   OBSERVACIONES DEL ASIENTO : " & vbCrLf & "  ============================="

        'If Len(Trim(vNoCtaCnt)) > 0 Then vProduc = vProduc & vNoCtaCnt & oImpresora.gPrnSaltoLinea
        If Len(Trim(vNoCtaCnt)) > 0 Then vProduc = vProduc & vNoCtaCnt & vbCrLf

        'If Len(Trim(vObs)) > 0 Then vProduc = vProduc & vObs & oImpresora.gPrnSaltoLinea
        If Len(Trim(vObs)) > 0 Then vProduc = vProduc & vObs & vbCrLf

        'vProduc = vProduc & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
        vProduc = vProduc & vbCrLf & vbCrLf

        'vLineas = Repetido(vProduc, oImpresora.gPrnSaltoLinea)
        vLineas = Repetido(vProduc, vbCrLf)

    Else
        '*************************************************************************************
        If pAsiDia Then
            'Actualización de variable de control
            ssql = "UPDATE ConstSistema SET nConsSisValor = '1' WHERE nConsSisCod = '2'"
            oCon.ejecutar ssql
        End If
        ssql = "UPDATE AsientoValida SET cAsientoEstado = '1' " & _
             " WHERE dAsientoFecha = '" & Format(txtFecha, gsFormatoFecha) & "' AND cAsientoTipo = '2' AND " & _
             " cCodUsu = '" & gsCodUser & "' AND dAsientoModif = '" & Format(dHoraGrab, gsFormatoFechaHora) & "'"
        oCon.ejecutar ssql
        '*************************************************************************************
    End If
    'Determina Productos
    'vProduc = vProduc & "   DE : " & oImpresora.gPrnSaltoLinea & "  ======" & oImpresora.gPrnSaltoLinea
    vProduc = vProduc & "   DE : " & vbCrLf & "  ======" & vbCrLf

    'If chkAho.value = 1 Then vProduc = vProduc & "     *  AHORROS " & oImpresora.gPrnSaltoLinea
    If chkAho.value = 1 Then vProduc = vProduc & "     *  AHORROS " & vbCrLf

    'If chkCre.value = 1 Then vProduc = vProduc & "     *  CREDITOS " & oImpresora.gPrnSaltoLinea
    If chkCre.value = 1 Then vProduc = vProduc & "     *  CREDITOS " & vbCrLf

    'If chkPig.value = 1 Then vProduc = vProduc & "     *  CREDITO PIGNORATICIO " & oImpresora.gPrnSaltoLinea
    If chkPig.value = 1 Then vProduc = vProduc & "     *  CREDITO PIGNORATICIO " & vbCrLf

    ssql = "SELECT  cctacnt, round(sum(ndebe),2) nDebe, round(sum(nhaber),2) nHaber, substring(cctacnt,3,1) AS MONEDA " & _
        " FROM AsientoDN "

    'Verifica que condiciones se ha de tener en cuenta:  Soles - Dolares
    vEst = "": vCad = ""
    If chkMonSol.value = 1 Then vEst = vEst & Moneda.gMonedaNacional
    If chkMonDol.value = 1 Then vEst = vEst & Moneda.gMonedaExtranjera
    For x = 1 To Len(vEst) Step 1
        vCad = vCad & "'" & Mid(vEst, x, 1) & "'"
        If x <> Len(vEst) Then vCad = vCad & ","
    Next x

    If Len(vEst) > 0 Then
        '        sSqlDolar = sSql & " WHERE cTipo IN ('3') AND datediff(dd,dFecha,'" & gdHoraGrab & "') = 0  GROUP BY CTIPO, CCTACNT ORDER BY cctacnt "
        sSqlDolar = ssql & " WHERE cTipo IN ('3') AND datediff(dd,dFecha,'" & gdHoraGrab & "') = 0  And substring(cctacnt,3,1) in ('1','2')" & _
            " GROUP BY CTIPO, CCTACNT ORDER BY cctacnt "
        ssql = ssql & " WHERE cTipo IN ('0') AND substring(cctacnt,3,1) IN (" & vCad & ") AND datediff(dd,dFecha,'" & gdHoraGrab & "') = 0  "
        If chkMonSol.value = 1 And chkMonDol.value = 1 Then
            ssql = ssql & " GROUP BY CTIPO, CCTACNT ORDER BY moneda , cctacnt "
        Else
            ssql = ssql & " GROUP BY CTIPO, CCTACNT ORDER BY cctacnt "
        End If
    End If

    vMoneAnte = ""
    vLineas = vLineas + 12: vPage = 1
    'vRTFImp = gsNomCmac & Space(vLenNomb + 6) & Format(gdFecSis & " " & Time, "dd/mm/yyyy hh:mm:ss") & oImpresora.gPrnSaltoLinea & Space(vLenNomb + 16) & " Página :" & ImpreFormat(vPage, 5, 0) & oImpresora.gPrnSaltoLinea & ImpreFormat(UCase(gsNomAge), 25) & oImpresora.gPrnSaltoLinea & ImpreFormat("   ASIENTO CONTABLE DEL DIA " & Format(txtFecha, "dd/mm/yyyy"), 44, 43) & oImpresora.gPrnSaltoLinea & ImpreFormat(String(40, "="), 44, 42) & oImpresora.gPrnSaltoLinea
    vRTFImp = gsNomCmac & Space(vLenNomb + 6) & Format(gdFecSis & " " & Time, "dd/mm/yyyy hh:mm:ss") & vbCrLf & Space(vLenNomb + 16) & " Página :" & ImpreFormat(vPage, 5, 0) & vbCrLf & ImpreFormat(UCase(gsNomAge), 25) & vbCrLf & ImpreFormat("   ASIENTO CONTABLE DEL DIA " & Format(txtFecha, "dd/mm/yyyy"), 44, 43) & vbCrLf & ImpreFormat(String(40, "="), 44, 42) & vbCrLf

    If CDate(txtFecha) = gdFecSis Then
        If Not CierreRealizado2(1, oCon.ConexionActiva) Then
            'vRTFImp = vRTFImp & ImpreFormat(" NO VALIDO (ANTES DEL CIERRE DE DIA)", 40, 45) & oImpresora.gPrnSaltoLinea
            vRTFImp = vRTFImp & ImpreFormat(" NO VALIDO (ANTES DEL CIERRE DE DIA)", 40, 45) & vbCrLf

        Else
            'Para los dias que no hay movimientos
            vDiaSinOpe = "Select cMovNro From Mov M Where M.cMovNro Like '" & Format(gdHoraGrab, gsFormatoMovFecha) & "%' AND M.nMovFlag = " & MovFlag.gMovFlagVigente
            Set RegTmp = oCon.CargaRecordSet(vDiaSinOpe)
            If (RegTmp.BOF Or RegTmp.EOF) Then
                vRTFImp = vRTFImp & vbCr & Space(28) & "N O   S E   R E A L I Z A R O N   O P E R A C I O N E S   E N   E L   D I A"
            Else
                If VerificaDiaHabil(gdHoraGrab, 3) And Not CierreRealizado2(2, oCon.ConexionActiva) Then
                    'vRTFImp = vRTFImp & ImpreFormat(" NO VALIDO (ANTES DEL CIERRE DE MES)", 40, 45) & oImpresora.gPrnSaltoLinea
                    vRTFImp = vRTFImp & ImpreFormat(" NO VALIDO (ANTES DEL CIERRE DE MES)", 40, 45) & vbCrLf
                End If
            End If
            RegTmp.Close
            Set RegTmp = Nothing
        End If
    End If
    vRTFImp = vRTFImp & vProduc
    oCon.AbreConexion
    oCon.CommadTimeOut = 0
    Set RegTmp = oCon.CargaRecordSet(ssql)
    If (RegTmp.BOF Or RegTmp.EOF) Then
        RegTmp.Close
        Set RegTmp = Nothing
    Else
        With RegTmp
        Do While Not .EOF
            If !Moneda <> vMoneAnte Then
                If !Moneda = Moneda.gMonedaNacional Then
                    'vRTFImp = vRTFImp & oImpresora.gPrnSaltoLinea & "   MONEDA : SOLES" & oImpresora.gPrnSaltoLinea
                    vRTFImp = vRTFImp & vbCrLf & "   MONEDA : SOLES" & vbCrLf

                    vLineas = vLineas + 1
                Else
                    If chkMonSol.value = 1 And (vSumDebe > 0 Or vSumHaber > 0) Then
                        'vRTFImp = vRTFImp & String(vespacio, "-") & oImpresora.gPrnSaltoLinea
                        vRTFImp = vRTFImp & String(vespacio, "-") & vbCrLf

                        'vRTFImp = vRTFImp & Space(vLenNomb + 12) & ImpreFormat("TOTAL", 11, 0) & ImpreFormat(vSumDebe, 12, , True) & ImpreFormat(vSumHaber, 12, , True) & oImpresora.gPrnSaltoLinea
                        vRTFImp = vRTFImp & Space(vLenNomb + 12) & ImpreFormat("TOTAL", 11, 0) & ImpreFormat(vSumDebe, 12, , True) & ImpreFormat(vSumHaber, 12, , True) & vbCrLf

                        'vRTFImp = vRTFImp & String(vespacio, "-") & oImpresora.gPrnSaltoLinea
                        vRTFImp = vRTFImp & String(vespacio, "-") & vbCrLf

                        vSumDebe = 0: vSumHaber = 0
                        vLineas = vLineas + 3
                    End If
                    vDolar = True
                    If chkMonSol.value = 1 Then
                        vPage = vPage + 1
                        'vRTFImp = vRTFImp & oImpresora.gPrnSaltoPagina & oImpresora.gPrnSaltoLinea
                        vRTFImp = vRTFImp & vbCrLf & vbCrLf

                        'vRTFImp = vRTFImp & Space(vLenNomb + 35) & "Página :" & ImpreFormat(vPage, 5, 0) & oImpresora.gPrnSaltoLinea
                        vRTFImp = vRTFImp & Space(vLenNomb + 35) & "Página :" & ImpreFormat(vPage, 5, 0) & vbCrLf

                    End If
                    'vRTFImp = vRTFImp & oImpresora.gPrnSaltoLinea & "   MONEDA : DOLARES" & oImpresora.gPrnSaltoLinea
                    vRTFImp = vRTFImp & vbCrLf & "   MONEDA : DOLARES" & vbCrLf

                    vLineas = 5 'vLineas + 2
                End If
                'vRTFImp = vRTFImp & String(vespacio, "-") & oImpresora.gPrnSaltoLinea
                vRTFImp = vRTFImp & String(vespacio, "-") & vbCrLf

                vRTFImp = vRTFImp & Space(1) & vCabecera
                'vRTFImp = vRTFImp & String(vespacio, "-") & oImpresora.gPrnSaltoLinea
                vRTFImp = vRTFImp & String(vespacio, "-") & vbCrLf

                vLineas = vLineas + 3
            End If
            vMoneAnte = !Moneda
            'vRTFImp = vRTFImp & ImpreFormat(!cctacnt, 21, 1) & ImpreFormat(UCase(CuentaNombre(!cctacnt, True, oCon.ConexionActiva)), vLenNomb, 1) & ImpreFormat(Round(!nDebe, 2), 12, , True) & ImpreFormat(Round(!nHaber, 2), 12, , True) & oImpresora.gPrnSaltoLinea
            vRTFImp = vRTFImp & ImpreFormat(!cctacnt, 21, 1) & ImpreFormat(UCase(CuentaNombre(!cctacnt, True, oCon.ConexionActiva)), vLenNomb, 1) & ImpreFormat(Round(!nDebe, 2), 12, , True) & ImpreFormat(Round(!nHaber, 2), 12, , True) & vbCrLf

            vSumDebe = vSumDebe + Round(!nDebe, 2)
            vSumHaber = vSumHaber + Round(!nHaber, 2)
            vLineas = vLineas + 1
            If vLineas > pLineasMax Then
                vPage = vPage + 1
                'vRTFImp = vRTFImp & oImpresora.gPrnSaltoPagina & oImpresora.gPrnSaltoLinea & Space(vLenNomb + 35) & "Página :" & ImpreFormat(vPage, 5, 0) & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & String(vespacio, "-") & oImpresora.gPrnSaltoLinea & Space(1) & vCabecera & String(vespacio, "-") & oImpresora.gPrnSaltoLinea
                vRTFImp = vRTFImp & vbCrLf & vbCrLf & Space(vLenNomb + 35) & "Página :" & ImpreFormat(vPage, 5, 0) & vbCrLf & vbCrLf & String(vespacio, "-") & vbCrLf & Space(1) & vCabecera & String(vespacio, "-") & vbCrLf

                vLineas = 5
            End If
            .MoveNext
        Loop
        End With
        'vRTFImp = vRTFImp & String(vespacio, "-") & oImpresora.gPrnSaltoLinea
        vRTFImp = vRTFImp & String(vespacio, "-") & vbCrLf

        'vRTFImp = vRTFImp & Space(vLenNomb + 12) & ImpreFormat("TOTAL", 11, 0) & ImpreFormat(vSumDebe, 12, , True) & ImpreFormat(vSumHaber, 12, , True) & oImpresora.gPrnSaltoLinea
        vRTFImp = vRTFImp & Space(vLenNomb + 12) & ImpreFormat("TOTAL", 11, 0) & ImpreFormat(vSumDebe, 12, , True) & ImpreFormat(vSumHaber, 12, , True) & vbCrLf

        'vRTFImp = vRTFImp & String(vespacio, "-") & oImpresora.gPrnSaltoLinea
        vRTFImp = vRTFImp & String(vespacio, "-") & vbCrLf

        vLineas = vLineas + 3
        RegTmp.Close
        Set RegTmp = Nothing

        If vDolar Then
            Set RegTmp = oCon.CargaRecordSet(sSqlDolar)
            If (RegTmp.BOF Or RegTmp.EOF) Then
                RegTmp.Close
                Set RegTmp = Nothing
            Else
                If vLineas + 8 >= pLineasMax Then
                    vPage = vPage + 1
                    'vRTFImp = vRTFImp & oImpresora.gPrnSaltoPagina & oImpresora.gPrnSaltoLinea
                    vRTFImp = vRTFImp & vbCrLf & vbCrLf

                    'vRTFImp = vRTFImp & Space(vLenNomb + 35) & "Página :" & ImpreFormat(vPage, 5, 0) & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
                    vRTFImp = vRTFImp & Space(vLenNomb + 35) & "Página :" & ImpreFormat(vPage, 5, 0) & vbCrLf & vbCrLf

                    vLineas = 5
                End If
                'vRTFImp = vRTFImp & oImpresora.gPrnSaltoLinea & "   CONVERSION DE DOLARES" & oImpresora.gPrnSaltoLinea & "     *  TIPO CAMBIO FIJO   : " & ImpreFormat(TCF, 6, 3) & oImpresora.gPrnSaltoLinea & "     *  TIPO CAMBIO VENTA  : " & ImpreFormat(TCV, 6, 3) & oImpresora.gPrnSaltoLinea & "     *  TIPO CAMBIO COMPRA : " & ImpreFormat(TCC, 6, 3) & oImpresora.gPrnSaltoLinea & String(vespacio, "-") & oImpresora.gPrnSaltoLinea & Space(1) & vCabecera & String(vespacio, "-") & oImpresora.gPrnSaltoLinea
                vRTFImp = vRTFImp & vbCrLf & "   CONVERSION DE DOLARES" & vbCrLf & "     *  TIPO CAMBIO FIJO   : " & ImpreFormat(TCF, 6, 3) & vbCrLf & "     *  TIPO CAMBIO VENTA  : " & ImpreFormat(TCV, 6, 3) & vbCrLf & "     *  TIPO CAMBIO COMPRA : " & ImpreFormat(TCC, 6, 3) & vbCrLf & String(vespacio, "-") & vbCrLf & Space(1) & vCabecera & String(vespacio, "-") & vbCrLf

                vLineas = vLineas + 8
                vSumDebe = 0: vSumHaber = 0
                With RegTmp
                Do While Not .EOF
                    'vRTFImp = vRTFImp & ImpreFormat(!cctacnt, 21, 1) & ImpreFormat(UCase(CuentaNombre(!cctacnt, True, oCon.ConexionActiva)), vLenNomb, 1) & ImpreFormat(Round(!nDebe, 2), 12, , True) & ImpreFormat(Round(!nHaber, 2), 12, , True) & oImpresora.gPrnSaltoLinea
                    vRTFImp = vRTFImp & ImpreFormat(!cctacnt, 21, 1) & ImpreFormat(UCase(CuentaNombre(!cctacnt, True, oCon.ConexionActiva)), vLenNomb, 1) & ImpreFormat(Round(!nDebe, 2), 12, , True) & ImpreFormat(Round(!nHaber, 2), 12, , True) & vbCrLf

                    vSumDebe = vSumDebe + !nDebe
                    vSumHaber = vSumHaber + !nHaber
                    vLineas = vLineas + 1
                    If vLineas > pLineasMax Then
                        vPage = vPage + 1
                        'vRTFImp = vRTFImp & oImpresora.gPrnSaltoPagina & oImpresora.gPrnSaltoLinea & Space(vLenNomb + 35) & "Página :" & ImpreFormat(vPage, 5, 0) & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & String(vespacio, "-") & oImpresora.gPrnSaltoLinea & Space(1) & vCabecera & String(vespacio, "-") & oImpresora.gPrnSaltoLinea
                        vRTFImp = vRTFImp & vbCrLf & vbCrLf & Space(vLenNomb + 35) & "Página :" & ImpreFormat(vPage, 5, 0) & vbCrLf & vbCrLf & String(vespacio, "-") & vbCrLf & Space(1) & vCabecera & String(vespacio, "-") & vbCrLf

                        vLineas = 5
                    End If
                    .MoveNext
                Loop
                End With
                RegTmp.Close
                Set RegTmp = Nothing
                If vLineas + 3 > pLineasMax Then
                    vPage = vPage + 1
                    'vRTFImp = vRTFImp & oImpresora.gPrnSaltoPagina & oImpresora.gPrnSaltoLinea
                    vRTFImp = vRTFImp & vbCrLf & vbCrLf

                    'vRTFImp = vRTFImp & Space(vLenNomb + 35) & "Página :" & ImpreFormat(vPage, 5, 0) & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
                    vRTFImp = vRTFImp & Space(vLenNomb + 35) & "Página :" & ImpreFormat(vPage, 5, 0) & vbCrLf & vbCrLf

                    vLineas = 5
                End If
                'vRTFImp = vRTFImp & String(vespacio, "-") & oImpresora.gPrnSaltoLinea
                vRTFImp = vRTFImp & String(vespacio, "-") & vbCrLf

                'vRTFImp = vRTFImp & Space(vLenNomb + 12) & ImpreFormat("TOTAL", 11, 0) & ImpreFormat((vSumDebe), 12, , True) & ImpreFormat((vSumHaber), 12, , True) & oImpresora.gPrnSaltoLinea
                vRTFImp = vRTFImp & Space(vLenNomb + 12) & ImpreFormat("TOTAL", 11, 0) & ImpreFormat((vSumDebe), 12, , True) & ImpreFormat((vSumHaber), 12, , True) & vbCrLf

                'vRTFImp = vRTFImp & String(vespacio, "-") & oImpresora.gPrnSaltoLinea
                vRTFImp = vRTFImp & String(vespacio, "-") & vbCrLf

                vLineas = vLineas + 3
            End If
        End If
    End If

    '**************************************************************************************
    'CARGA DE : Compra y Venta en Dolares y en Soles
    Me.Caption = "Cargando asientos de compra - venta dolares y soles"

    Dim ComVta(6, 1) As Currency ' 0-Ctas; 1-Debe;  2-Haber
    Erase ComVta
    Dim ComVtaDolar As String
    ComVtaDolar = ""

    ssql = "SELECT  cctacnt, round(sum(ndebe),2) nDebe, round(sum(nhaber),2) nHaber, ctipo " & _
        " FROM AsientoDN WHERE ctipo IN ('1','2') AND datediff(dd, dFecha ,'" & gdHoraGrab & "') = 0 " & _
        " GROUP BY CTIPO, CCTACNT "

    'ssql = "SELECT  cctacnt, round(sum(ndebe),2) nDebe, round(sum(nhaber),2) nHaber, ctipo " & _
        " FROM AsientoDN WHERE ctipo IN ('1','2') AND cOpeCod NOT IN ('900022','900023') AND datediff(dd, dFecha ,'" & gdHoraGrab & "') = 0 " & _
        " GROUP BY CTIPO, CCTACNT " 'JUCS TI-ERS 002-2017

    Set RegTmp = oCon.CargaRecordSet(ssql)
    If (RegTmp.BOF Or RegTmp.EOF) Then
    Else
        If vLineas + 19 >= pLineasMax Then
            vPage = vPage + 1
            'vRTFImp = vRTFImp & oImpresora.gPrnSaltoPagina & oImpresora.gPrnSaltoLinea
            vRTFImp = vRTFImp & vbCrLf & vbCrLf

            'vRTFImp = vRTFImp & Space(vLenNomb + 35) & "Página :" & ImpreFormat(vPage, 5, 0) & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
            vRTFImp = vRTFImp & Space(vLenNomb + 35) & "Página :" & ImpreFormat(vPage, 5, 0) & vbCrLf & vbCrLf

            vLineas = 5
        End If

        With RegTmp
            Do While Not .EOF
                If !cTipo = "1" Then
                    If Left(!cctacnt, 4) = lsCajaDolares Or Left(vCodConta, 4) = "2524" Then 'APRI20170705 RFC1705220001 Add Left(vCodConta, 4) = "2524"
                        ComVta(1, 0) = !nDebe
                        ComVta(1, 1) = !nHaber
                    ElseIf Left(!cctacnt, 4) = lsCajaSoles Or Left(vCodConta, 4) = "2514" Then 'APRI20170705 RFC1705220001 Add Left(vCodConta, 4) = "2514"
                        ComVta(3, 0) = !nDebe
                        ComVta(3, 1) = !nHaber
                    ElseIf Left(!cctacnt, 4) = "4128" Then
                        ComVta(4, 0) = !nDebe
                    ElseIf Left(!cctacnt, 4) = "5128" Then
                        ComVta(4, 1) = !nHaber
                    Else
                        MsgBox "Cuenta no reconocida para Compra y Venta de dolares", vbInformation, " Aviso "
                    End If
                Else
                    If Left(!cctacnt, 4) = lsCajaDolares Or Left(vCodConta, 4) = "2524" Then 'APRI20170705 RFC1705220001 Add Left(vCodConta, 4) = "2524"
                        ComVta(5, 0) = !nDebe
                        ComVta(5, 1) = !nHaber
                    Else
                        MsgBox "Cuenta no reconocida para Compra y Venta de dolares", vbInformation, " Aviso "
                    End If
                End If
                .MoveNext
            Loop
        End With

        'ComVtaDolar = "" & oImpresora.gPrnSaltoLinea
        ComVtaDolar = "" & vbCrLf

        'ComVtaDolar = ComVtaDolar & "   DE : COMPRA Y VENTA EN DOLARES " & oImpresora.gPrnSaltoLinea
        ComVtaDolar = ComVtaDolar & "   DE : COMPRA Y VENTA EN DOLARES " & vbCrLf

        'ComVtaDolar = ComVtaDolar & String(vespacio, "-") & oImpresora.gPrnSaltoLinea
        ComVtaDolar = ComVtaDolar & String(vespacio, "-") & vbCrLf

        Dim sCajaME As String, sCajaMN As String
        Dim nDebeAcum As Double, nHaberAcum As Double

        nDebeAcum = 0
        nHaberAcum = 0
        RegTmp.MoveFirst

        Do While Not RegTmp.EOF
            If RegTmp("cTipo") = "2" Then
                sCajaME = RegTmp("cCtaCnt")
                nDebeAcum = nDebeAcum + RegTmp("nDebe")
                nHaberAcum = nHaberAcum + RegTmp("nHaber")
                'ComVtaDolar = ComVtaDolar & ImpreFormat(sCajaME, 21, 1) & ImpreFormat(UCase(CuentaNombre(sCajaME, True, oCon.ConexionActiva)), vLenNomb, 1) & ImpreFormat(Round(RegTmp("nDebe"), 2), 12) & ImpreFormat(Round(RegTmp("nHaber"), 2), 12) & oImpresora.gPrnSaltoLinea
                ComVtaDolar = ComVtaDolar & ImpreFormat(sCajaME, 21, 1) & ImpreFormat(UCase(CuentaNombre(sCajaME, True, oCon.ConexionActiva)), vLenNomb, 1) & ImpreFormat(Round(RegTmp("nDebe"), 2), 12) & ImpreFormat(Round(RegTmp("nHaber"), 2), 12) & vbCrLf

            End If
            RegTmp.MoveNext
        Loop

        'ComVtaDolar = ComVtaDolar & String(vespacio, "-") & oImpresora.gPrnSaltoLinea
        ComVtaDolar = ComVtaDolar & String(vespacio, "-") & vbCrLf

        'ComVtaDolar = ComVtaDolar & Space(vLenNomb + 12) & ImpreFormat("TOTAL", 11, 0) & ImpreFormat(Round(nDebeAcum, 2), 12, , True) & ImpreFormat(Round(nHaberAcum, 2), 12, , True) & oImpresora.gPrnSaltoLinea
        ComVtaDolar = ComVtaDolar & Space(vLenNomb + 12) & ImpreFormat("TOTAL", 11, 0) & ImpreFormat(Round(nDebeAcum, 2), 12, , True) & ImpreFormat(Round(nHaberAcum, 2), 12, , True) & vbCrLf

        'ComVtaDolar = ComVtaDolar & String(vespacio, "-") & oImpresora.gPrnSaltoLinea
        ComVtaDolar = ComVtaDolar & String(vespacio, "-") & vbCrLf

        'ComVtaDolar = ComVtaDolar & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
        ComVtaDolar = ComVtaDolar & vbCrLf & vbCrLf


        'ComVtaDolar = ComVtaDolar & "   DE : COMPRA Y VENTA EN SOLES " & oImpresora.gPrnSaltoLinea
        ComVtaDolar = ComVtaDolar & "   DE : COMPRA Y VENTA EN SOLES " & vbCrLf

        'ComVtaDolar = ComVtaDolar & String(vespacio, "-") & oImpresora.gPrnSaltoLinea
        ComVtaDolar = ComVtaDolar & String(vespacio, "-") & vbCrLf

        nDebeAcum = 0
        nHaberAcum = 0
        RegTmp.MoveFirst
        Do While Not RegTmp.EOF
            If RegTmp("cTipo") = "1" Then
                sCajaMN = RegTmp("cCtaCnt")
                nDebeAcum = nDebeAcum + RegTmp("nDebe")
                nHaberAcum = nHaberAcum + RegTmp("nHaber")
                'ComVtaDolar = ComVtaDolar & ImpreFormat(sCajaMN, 21, 1) & ImpreFormat(UCase(CuentaNombre(sCajaMN, True, oCon.ConexionActiva)), vLenNomb, 1) & ImpreFormat(RegTmp("nDebe"), 12) & ImpreFormat(RegTmp("nHaber"), 12) & oImpresora.gPrnSaltoLinea
                ComVtaDolar = ComVtaDolar & ImpreFormat(sCajaMN, 21, 1) & ImpreFormat(UCase(CuentaNombre(sCajaMN, True, oCon.ConexionActiva)), vLenNomb, 1) & ImpreFormat(RegTmp("nDebe"), 12) & ImpreFormat(RegTmp("nHaber"), 12) & vbCrLf

            End If
            RegTmp.MoveNext
        Loop

        'ComVtaDolar = ComVtaDolar & String(vespacio, "-") & oImpresora.gPrnSaltoLinea
        ComVtaDolar = ComVtaDolar & String(vespacio, "-") & vbCrLf

        'ComVtaDolar = ComVtaDolar & Space(vLenNomb + 12) & ImpreFormat("TOTAL", 11, 0) & ImpreFormat(nDebeAcum, 12, , True) & ImpreFormat(nHaberAcum, 12, , True) & oImpresora.gPrnSaltoLinea
        ComVtaDolar = ComVtaDolar & Space(vLenNomb + 12) & ImpreFormat("TOTAL", 11, 0) & ImpreFormat(nDebeAcum, 12, , True) & ImpreFormat(nHaberAcum, 12, , True) & vbCrLf

        'ComVtaDolar = ComVtaDolar & String(vespacio, "-") & oImpresora.gPrnSaltoLinea
        ComVtaDolar = ComVtaDolar & String(vespacio, "-") & vbCrLf

        'ComVtaDolar = ComVtaDolar & oImpresora.gPrnSaltoLinea
        ComVtaDolar = ComVtaDolar & vbCrLf

    End If
    RegTmp.Close
    Set RegTmp = Nothing

    Me.Caption = "Asiento finalizado"
    'Envia Asiento al Previo
    vRTFImp = vRTFImp & ComVtaDolar
    If Len(Trim(vRTFImp)) = 0 Then
        MsgBox " No existe ningún asiento generado en este día ", vbInformation, " Aviso "
    End If

    rtfImp.Text = vRTFImp

    'Graba asiento automaticamente
    If Not (Len(Trim(vObs)) > 0 Or Len(Trim(vNoCtaCnt)) > 0) And pAsiDia And CierreRealizado2(1, oCon.ConexionActiva) Then

        If gbAsientoDN Then
            'vRTFImp = oImpresora.gPrnCondensadaON & vRTFImp & oImpresora.gPrnCondensadaOFF
            vRTFImp = vRTFImp & vbCrLf
            rtfImp.Text = vRTFImp
            Me.rtfImp.SaveFile App.Path & "\Spooler\AP" & Right(gsCodAge, 2) & "_" & Format(gdFecSis, "yyyymmdd") & ".txt", 1
        End If
    End If

'*---------- fin

    '-----------------------------------------------------
    Call GeneraLogAsiento("Termina generacion de asiento contable del " & Format(ldFechaAsiento, "dd/MM/yyyy") & " - ")
    Call GeneraLogAsiento("===============================================================================", 0)
    '-----------------------------------------------------

    '*** PEAC 20130119
    If Len(Trim(vObs)) = 0 Or Len(Trim(vNoCtaCnt)) = 0 Then
        'MsgBox "Proceso Contable Finalizado OK, se mostrará el Asiento.", vbInformation, "Atención"
    End If

    cNomArchivoTxt = "AC" & Format(txtFecha, "yyyymmdd") & "_" & Trim(str(Int((9999 - 1) * Rnd() + 1) * 100)) & ".txt"

    '*** PEAC 20170905
    Call GeneraRepoAsntoCnt(vRTFImp, cNomArchivoTxt)

    If optImpresion(0).value = True Then
        'oPrevio.Show rtfImp.Text, " ASIENTOS CONTABLES DEL DIA " & Format(gdHoraGrab, "dd/mm/yyyy"), True, 66, gImpresora
        'Shell "notepad C:\install\prueba.txt", vbNormalFocus
        Shell "notepad C:\Install\Spooler\" & cNomArchivoTxt, vbNormalFocus
    Else
        ImpreBegin True, 66
            Print #ArcSal, rtfImp.Text
        ImpreEnd
    End If

    mciExecute "Close All" 'PEAC 20160701

'   ********COMENTADO POR APRI 20170428
'    'TRANSFERENCIA DEL ASIENTO - SOLO CUANDO ES EL DIA
'    'Verifica la veracidad del asiento contable
'    If CDate(txtFecha) = gdFecSis Then
'        If Not (Len(Trim(vObs)) > 0 Or Len(Trim(vNoCtaCnt)) > 0) And pAsiDia And CierreRealizado2(1, oCon.ConexionActiva) Then
'            If MsgBox(" Desea migrar el Asiento ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
'                Call TransfiereAsiento
'                MsgBox "Asiento Transferido satisfactoriamente...!", vbInformation, "¡Aviso!"
'            End If
'        End If
'    Else
'        If Not (Len(Trim(vObs)) > 0 Or Len(Trim(vNoCtaCnt)) > 0) And pAsiDia = False Then
'            If MsgBox(" Desea migrar el Asiento ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
'                Call TransfiereAsiento
'                MsgBox "Asiento Transferido satisfactoriamente...!", vbInformation, "¡Aviso!"
'            End If
'        End If
'    End If
'******************END APRI
    Call TransferenciaAsientoDN  'AGREGADO POR APRI 20170428
    MousePointer = 0

    'Activo Controles
    cmdAsientos.Enabled = True
    txtFecha.Enabled = True
    fraContenedor(1).Enabled = True
    fraContenedor(2).Enabled = True
    fraImpresion.Enabled = True
    'Fin Activos

    ' INICIO PEAC 20191216
    ' envia correo comunicando culminacion de asiento cnt
    ' y envio de asiento contable

    Call EnviaCorreo(IIf(Len(Trim(vObs)) > 0, Trim(vObs), "Asiento Contable sin observaciones."))
    'End
    
    ' FIN PEAC

    Exit Sub

ControlError:       ' Rutina de control de errores.
'        prgList.value = 0
        'Activo Controles
        cmdAsientos.Enabled = True
        txtFecha.Enabled = True
        fraContenedor(1).Enabled = True
        fraContenedor(2).Enabled = True
        fraImpresion.Enabled = True
        'Fin Activos
        MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
           " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub
'PEAC 20191216
Private Sub EnviaCorreo(Optional ByVal psOtroMensaje As String = "@")
    Dim oConst As COMDConstSistema.NCOMConstSistema
    Dim lsCorreoDestino As String, lsContenido As String
    Set oConst = New COMDConstSistema.NCOMConstSistema
    lsCorreoDestino = "pvalera@cajamaynas.pe" 'oConst.LeeConstSistema(93)

    If psOtroMensaje = "@" Then
        lsContenido = "El cierre de operaciones del día " & Trim(CStr(gdFecSis - 1)) & " ha culminado satisfactoriamente, " & _
                      " Así mismo se realizó el inicio del día " & Trim(CStr(gdFecSis)) & "."
    Else
        lsContenido = psOtroMensaje
    End If
    
    EnviarMail gsCorreoHost, gsCorreoEnvia, lsCorreoDestino, "Culminó la generación de Asiento Contable.", lsContenido
End Sub

Private Sub VerificaVeracidadAsntoCnt()
    'Verifica la veracidad del asiento contable
    Me.Caption = "Verificando Veracidad de asiento Contable"
    If Len(Trim(vObs)) > 0 Or Len(Trim(vNoCtaCnt)) > 0 Then

        MsgBox "El asiento presenta Observaciones ! ", vbInformation, " Aviso "
        'vProduc = oImpresora.gPrnSaltoLinea & Space(40) & "  A S I E N T O     ¡   N O     V A L I D O   !" & oImpresora.gPrnSaltoLinea
        vProduc = vbCrLf & Space(40) & "  A S I E N T O     ¡   N O     V A L I D O   !" & vbCrLf
        
        'vProduc = vProduc & "   OBSERVACIONES DEL ASIENTO : " & oImpresora.gPrnSaltoLinea & _
                            "  ============================="
                            
        vProduc = vProduc & "   OBSERVACIONES DEL ASIENTO : " & vbCrLf & "                   ============================="
                            
                            
        'If Len(Trim(vNoCtaCnt)) > 0 Then vProduc = vProduc & vNoCtaCnt & oImpresora.gPrnSaltoLinea
        If Len(Trim(vNoCtaCnt)) > 0 Then vProduc = vProduc & vNoCtaCnt & vbCrLf
                
        'If Len(Trim(vObs)) > 0 Then vProduc = vProduc & vObs & oImpresora.gPrnSaltoLinea
        If Len(Trim(vObs)) > 0 Then vProduc = vProduc & vObs & vbCrLf
        
        'vProduc = vProduc & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
        vProduc = vProduc & vbCrLf & vbCrLf
        
        'vLineas = Repetido(vProduc, oImpresora.gPrnSaltoLinea)
        vLineas = Repetido(vProduc, vbCrLf)
        
    Else
        '*************************************************************************************
        If pAsiDia Then
            'Actualización de variable de control
            ssql = "UPDATE ConstSistema SET nConsSisValor = '1' WHERE nConsSisCod = '2'"
            oCon.ejecutar ssql
        End If
        ssql = "UPDATE AsientoValida SET cAsientoEstado = '1' " & _
             " WHERE dAsientoFecha = '" & Format(txtFecha, gsFormatoFecha) & "' AND cAsientoTipo = '2' AND " & _
             " cCodUsu = '" & gsCodUser & "' AND dAsientoModif = '" & Format(dHoraGrab, gsFormatoFechaHora) & "'"
        oCon.ejecutar ssql
        '*************************************************************************************
    End If
    'Determina Productos
    'vProduc = vProduc & "   DE : " & oImpresora.gPrnSaltoLinea & "  ======" & oImpresora.gPrnSaltoLinea
    vProduc = vProduc & "   DE : " & vbCrLf & "  ======" & vbCrLf
    
    'If chkAho.value = 1 Then vProduc = vProduc & "     *  AHORROS " & oImpresora.gPrnSaltoLinea
    If chkAho.value = 1 Then vProduc = vProduc & "     *  AHORROS " & vbCrLf
    
    'If chkCre.value = 1 Then vProduc = vProduc & "     *  CREDITOS " & oImpresora.gPrnSaltoLinea
    If chkCre.value = 1 Then vProduc = vProduc & "     *  CREDITOS " & vbCrLf
    
    'If chkPig.value = 1 Then vProduc = vProduc & "     *  CREDITO PIGNORATICIO " & oImpresora.gPrnSaltoLinea
    If chkPig.value = 1 Then vProduc = vProduc & "     *  CREDITO PIGNORATICIO " & vbCrLf
    
    ssql = "SELECT  cctacnt, round(sum(ndebe),2) nDebe, round(sum(nhaber),2) nHaber, substring(cctacnt,3,1) AS MONEDA " & _
        " FROM AsientoDN "
    
    'Verifica que condiciones se ha de tener en cuenta:  Soles - Dolares
    vEst = "": vCad = ""
    If chkMonSol.value = 1 Then vEst = vEst & Moneda.gMonedaNacional
    If chkMonDol.value = 1 Then vEst = vEst & Moneda.gMonedaExtranjera
    For x = 1 To Len(vEst) Step 1
        vCad = vCad & "'" & Mid(vEst, x, 1) & "'"
        If x <> Len(vEst) Then vCad = vCad & ","
    Next x
    
    If Len(vEst) > 0 Then
        '        sSqlDolar = sSql & " WHERE cTipo IN ('3') AND datediff(dd,dFecha,'" & gdHoraGrab & "') = 0  GROUP BY CTIPO, CCTACNT ORDER BY cctacnt "
        sSqlDolar = ssql & " WHERE cTipo IN ('3') AND datediff(dd,dFecha,'" & gdHoraGrab & "') = 0  And substring(cctacnt,3,1) in ('1','2')" & _
            " GROUP BY CTIPO, CCTACNT ORDER BY cctacnt "
        ssql = ssql & " WHERE cTipo IN ('0') AND substring(cctacnt,3,1) IN (" & vCad & ") AND datediff(dd,dFecha,'" & gdHoraGrab & "') = 0  "
        If chkMonSol.value = 1 And chkMonDol.value = 1 Then
            ssql = ssql & " GROUP BY CTIPO, CCTACNT ORDER BY moneda , cctacnt "
        Else
            ssql = ssql & " GROUP BY CTIPO, CCTACNT ORDER BY cctacnt "
        End If
    End If
    
    vMoneAnte = ""
    vLineas = vLineas + 12: vPage = 1
    'vRTFImp = gsNomCmac & Space(vLenNomb + 6) & Format(gdFecSis & " " & Time, "dd/mm/yyyy hh:mm:ss") & oImpresora.gPrnSaltoLinea & Space(vLenNomb + 16) & " Página :" & ImpreFormat(vPage, 5, 0) & oImpresora.gPrnSaltoLinea & ImpreFormat(UCase(gsNomAge), 25) & oImpresora.gPrnSaltoLinea & ImpreFormat("   ASIENTO CONTABLE DEL DIA " & Format(txtFecha, "dd/mm/yyyy"), 44, 43) & oImpresora.gPrnSaltoLinea & ImpreFormat(String(40, "="), 44, 42) & oImpresora.gPrnSaltoLinea
    vRTFImp = gsNomCmac & Space(vLenNomb + 6) & Format(gdFecSis & " " & Time, "dd/mm/yyyy hh:mm:ss") & vbCrLf & Space(vLenNomb + 16) & " Página :" & ImpreFormat(vPage, 5, 0) & vbCrLf & ImpreFormat(UCase(gsNomAge), 25) & vbCrLf & ImpreFormat("   ASIENTO CONTABLE DEL DIA " & Format(txtFecha, "dd/mm/yyyy"), 44, 43) & vbCrLf & ImpreFormat(String(40, "="), 44, 42) & vbCrLf
    
    
    If CDate(txtFecha) = gdFecSis Then
        If Not CierreRealizado2(1, oCon.ConexionActiva) Then
            'vRTFImp = vRTFImp & ImpreFormat(" NO VALIDO (ANTES DEL CIERRE DE DIA)", 40, 45) & oImpresora.gPrnSaltoLinea
            vRTFImp = vRTFImp & ImpreFormat(" NO VALIDO (ANTES DEL CIERRE DE DIA)", 40, 45) & vbCrLf
            
        Else
            'Para los dias que no hay movimientos
            vDiaSinOpe = "Select cMovNro From Mov M Where M.cMovNro Like '" & Format(gdHoraGrab, gsFormatoMovFecha) & "%' AND M.nMovFlag = " & MovFlag.gMovFlagVigente
            Set RegTmp = oCon.CargaRecordSet(vDiaSinOpe)
            If (RegTmp.BOF Or RegTmp.EOF) Then
                vRTFImp = vRTFImp & vbCr & Space(28) & "N O   S E   R E A L I Z A R O N   O P E R A C I O N E S   E N   E L   D I A"
            Else
                If VerificaDiaHabil(gdHoraGrab, 3) And Not CierreRealizado2(2, oCon.ConexionActiva) Then
                    'vRTFImp = vRTFImp & ImpreFormat(" NO VALIDO (ANTES DEL CIERRE DE MES)", 40, 45) & oImpresora.gPrnSaltoLinea
                    vRTFImp = vRTFImp & ImpreFormat(" NO VALIDO (ANTES DEL CIERRE DE MES)", 40, 45) & vbCrLf
                    
                End If
            End If
            RegTmp.Close
            Set RegTmp = Nothing
        End If
    End If
    vRTFImp = vRTFImp & vProduc
    oCon.AbreConexion
    oCon.CommadTimeOut = 0
    Set RegTmp = oCon.CargaRecordSet(ssql)
    If (RegTmp.BOF Or RegTmp.EOF) Then
        RegTmp.Close
        Set RegTmp = Nothing
    Else
        With RegTmp
        Do While Not .EOF
            If !Moneda <> vMoneAnte Then
                If !Moneda = Moneda.gMonedaNacional Then
                    'vRTFImp = vRTFImp & oImpresora.gPrnSaltoLinea & "   MONEDA : SOLES" & oImpresora.gPrnSaltoLinea
                    vRTFImp = vRTFImp & vbCrLf & "   MONEDA : SOLES" & vbCrLf
                    
                    vLineas = vLineas + 1
                Else
                    If chkMonSol.value = 1 And (vSumDebe > 0 Or vSumHaber > 0) Then
                        'vRTFImp = vRTFImp & String(vespacio, "-") & oImpresora.gPrnSaltoLinea
                        vRTFImp = vRTFImp & String(vespacio, "-") & vbCrLf
                        
                        'vRTFImp = vRTFImp & Space(vLenNomb + 12) & ImpreFormat("TOTAL", 11, 0) & ImpreFormat(vSumDebe, 12, , True) & ImpreFormat(vSumHaber, 12, , True) & oImpresora.gPrnSaltoLinea
                        vRTFImp = vRTFImp & Space(vLenNomb + 12) & ImpreFormat("TOTAL", 11, 0) & ImpreFormat(vSumDebe, 12, , True) & ImpreFormat(vSumHaber, 12, , True) & vbCrLf
                        
                        'vRTFImp = vRTFImp & String(vespacio, "-") & oImpresora.gPrnSaltoLinea
                        vRTFImp = vRTFImp & String(vespacio, "-") & vbCrLf
                        
                        vSumDebe = 0: vSumHaber = 0
                        vLineas = vLineas + 3
                    End If
                    vDolar = True
                    If chkMonSol.value = 1 Then
                        vPage = vPage + 1
                        'vRTFImp = vRTFImp & oImpresora.gPrnSaltoPagina & oImpresora.gPrnSaltoLinea
                        vRTFImp = vRTFImp & vbCrLf & vbCrLf
                        
                        
                        'vRTFImp = vRTFImp & Space(vLenNomb + 35) & "Página :" & ImpreFormat(vPage, 5, 0) & oImpresora.gPrnSaltoLinea
                        vRTFImp = vRTFImp & Space(vLenNomb + 35) & "Página :" & ImpreFormat(vPage, 5, 0) & vbCrLf
                        
                    End If
                    'vRTFImp = vRTFImp & oImpresora.gPrnSaltoLinea & "   MONEDA : DOLARES" & oImpresora.gPrnSaltoLinea
                    vRTFImp = vRTFImp & vbCrLf & "   MONEDA : DOLARES" & vbCrLf
                    
                    vLineas = 5 'vLineas + 2
                End If
                'vRTFImp = vRTFImp & String(vespacio, "-") & oImpresora.gPrnSaltoLinea
                vRTFImp = vRTFImp & String(vespacio, "-") & vbCrLf
                
                vRTFImp = vRTFImp & Space(1) & vCabecera
                'vRTFImp = vRTFImp & String(vespacio, "-") & oImpresora.gPrnSaltoLinea
                vRTFImp = vRTFImp & String(vespacio, "-") & vbCrLf
                
                vLineas = vLineas + 3
            End If
            vMoneAnte = !Moneda
            'vRTFImp = vRTFImp & ImpreFormat(!cctacnt, 21, 1) & ImpreFormat(UCase(CuentaNombre(!cctacnt, True, oCon.ConexionActiva)), vLenNomb, 1) & ImpreFormat(Round(!nDebe, 2), 12, , True) & ImpreFormat(Round(!nHaber, 2), 12, , True) & oImpresora.gPrnSaltoLinea
            vRTFImp = vRTFImp & ImpreFormat(!cctacnt, 21, 1) & ImpreFormat(UCase(CuentaNombre(!cctacnt, True, oCon.ConexionActiva)), vLenNomb, 1) & ImpreFormat(Round(!nDebe, 2), 12, , True) & ImpreFormat(Round(!nHaber, 2), 12, , True) & vbCrLf
            
            vSumDebe = vSumDebe + Round(!nDebe, 2)
            vSumHaber = vSumHaber + Round(!nHaber, 2)
            vLineas = vLineas + 1
            If vLineas > pLineasMax Then
                vPage = vPage + 1
                'vRTFImp = vRTFImp & oImpresora.gPrnSaltoPagina & oImpresora.gPrnSaltoLinea & Space(vLenNomb + 35) & "Página :" & ImpreFormat(vPage, 5, 0) & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & String(vespacio, "-") & oImpresora.gPrnSaltoLinea & Space(1) & vCabecera & String(vespacio, "-") & oImpresora.gPrnSaltoLinea
                vRTFImp = vRTFImp & vbCrLf & vbCrLf & Space(vLenNomb + 35) & "Página :" & ImpreFormat(vPage, 5, 0) & vbCrLf & vbCrLf & String(vespacio, "-") & vbCrLf & Space(1) & vCabecera & String(vespacio, "-") & vbCrLf
                
                vLineas = 5
            End If
            .MoveNext
        Loop
        End With
        'vRTFImp = vRTFImp & String(vespacio, "-") & oImpresora.gPrnSaltoLinea
        vRTFImp = vRTFImp & String(vespacio, "-") & vbCrLf
        
        'vRTFImp = vRTFImp & Space(vLenNomb + 12) & ImpreFormat("TOTAL", 11, 0) & ImpreFormat(vSumDebe, 12, , True) & ImpreFormat(vSumHaber, 12, , True) & oImpresora.gPrnSaltoLinea
        vRTFImp = vRTFImp & Space(vLenNomb + 12) & ImpreFormat("TOTAL", 11, 0) & ImpreFormat(vSumDebe, 12, , True) & ImpreFormat(vSumHaber, 12, , True) & vbCrLf
        
        'vRTFImp = vRTFImp & String(vespacio, "-") & oImpresora.gPrnSaltoLinea
        vRTFImp = vRTFImp & String(vespacio, "-") & vbCrLf
        
        vLineas = vLineas + 3
        RegTmp.Close
        Set RegTmp = Nothing
        
        If vDolar Then
            Set RegTmp = oCon.CargaRecordSet(sSqlDolar)
            If (RegTmp.BOF Or RegTmp.EOF) Then
                RegTmp.Close
                Set RegTmp = Nothing
            Else
                If vLineas + 8 >= pLineasMax Then
                    vPage = vPage + 1
                    'vRTFImp = vRTFImp & oImpresora.gPrnSaltoPagina & oImpresora.gPrnSaltoLinea
                    vRTFImp = vRTFImp & vbCrLf & vbCrLf
                    
                    'vRTFImp = vRTFImp & Space(vLenNomb + 35) & "Página :" & ImpreFormat(vPage, 5, 0) & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
                    vRTFImp = vRTFImp & Space(vLenNomb + 35) & "Página :" & ImpreFormat(vPage, 5, 0) & vbCrLf & vbCrLf
                    
                    vLineas = 5
                End If
                
                'vRTFImp = vRTFImp & oImpresora.gPrnSaltoLinea & "   CONVERSION DE DOLARES" & oImpresora.gPrnSaltoLinea & "     *  TIPO CAMBIO FIJO   : " & ImpreFormat(TCF, 6, 3) & oImpresora.gPrnSaltoLinea & "     *  TIPO CAMBIO VENTA  : " & ImpreFormat(TCV, 6, 3) & oImpresora.gPrnSaltoLinea & "     *  TIPO CAMBIO COMPRA : " & ImpreFormat(TCC, 6, 3) & oImpresora.gPrnSaltoLinea & String(vespacio, "-") & oImpresora.gPrnSaltoLinea & Space(1) & vCabecera & String(vespacio, "-") & oImpresora.gPrnSaltoLinea
                vRTFImp = vRTFImp & vbCrLf & "   CONVERSION DE DOLARES" & vbCrLf & "     *  TIPO CAMBIO FIJO   : " & ImpreFormat(TCF, 6, 3) & vbCrLf & "     *  TIPO CAMBIO VENTA  : " & ImpreFormat(TCV, 6, 3) & vbCrLf & "     *  TIPO CAMBIO COMPRA : " & ImpreFormat(TCC, 6, 3) & vbCrLf & String(vespacio, "-") & vbCrLf & Space(1) & vCabecera & String(vespacio, "-") & vbCrLf
                
                vLineas = vLineas + 8
                vSumDebe = 0: vSumHaber = 0
                With RegTmp
                Do While Not .EOF
                    'vRTFImp = vRTFImp & ImpreFormat(!cctacnt, 21, 1) & ImpreFormat(UCase(CuentaNombre(!cctacnt, True, oCon.ConexionActiva)), vLenNomb, 1) & ImpreFormat(Round(!nDebe, 2), 12, , True) & ImpreFormat(Round(!nHaber, 2), 12, , True) & oImpresora.gPrnSaltoLinea
                    vRTFImp = vRTFImp & ImpreFormat(!cctacnt, 21, 1) & ImpreFormat(UCase(CuentaNombre(!cctacnt, True, oCon.ConexionActiva)), vLenNomb, 1) & ImpreFormat(Round(!nDebe, 2), 12, , True) & ImpreFormat(Round(!nHaber, 2), 12, , True) & vbCrLf
                    
                    vSumDebe = vSumDebe + !nDebe
                    vSumHaber = vSumHaber + !nHaber
                    vLineas = vLineas + 1
                    If vLineas > pLineasMax Then
                        vPage = vPage + 1
                        'vRTFImp = vRTFImp & oImpresora.gPrnSaltoPagina & oImpresora.gPrnSaltoLinea & Space(vLenNomb + 35) & "Página :" & ImpreFormat(vPage, 5, 0) & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & String(vespacio, "-") & oImpresora.gPrnSaltoLinea & Space(1) & vCabecera & String(vespacio, "-") & oImpresora.gPrnSaltoLinea
                        vRTFImp = vRTFImp & vbCrLf & vbCrLf & Space(vLenNomb + 35) & "Página :" & ImpreFormat(vPage, 5, 0) & vbCrLf & vbCrLf & String(vespacio, "-") & vbCrLf & Space(1) & vCabecera & String(vespacio, "-") & vbCrLf
                        
                        vLineas = 5
                    End If
                    .MoveNext
                Loop
                End With
                RegTmp.Close
                Set RegTmp = Nothing
                If vLineas + 3 > pLineasMax Then
                    vPage = vPage + 1
                    'vRTFImp = vRTFImp & oImpresora.gPrnSaltoPagina & oImpresora.gPrnSaltoLinea
                    vRTFImp = vRTFImp & vbCrLf & vbCrLf
                    
                    'vRTFImp = vRTFImp & Space(vLenNomb + 35) & "Página :" & ImpreFormat(vPage, 5, 0) & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
                    vRTFImp = vRTFImp & Space(vLenNomb + 35) & "Página :" & ImpreFormat(vPage, 5, 0) & vbCrLf & vbCrLf
                    
                    vLineas = 5
                End If
                'vRTFImp = vRTFImp & String(vespacio, "-") & oImpresora.gPrnSaltoLinea
                vRTFImp = vRTFImp & String(vespacio, "-") & vbCrLf
                
                'vRTFImp = vRTFImp & Space(vLenNomb + 12) & ImpreFormat("TOTAL", 11, 0) & ImpreFormat((vSumDebe), 12, , True) & ImpreFormat((vSumHaber), 12, , True) & oImpresora.gPrnSaltoLinea
                vRTFImp = vRTFImp & Space(vLenNomb + 12) & ImpreFormat("TOTAL", 11, 0) & ImpreFormat((vSumDebe), 12, , True) & ImpreFormat((vSumHaber), 12, , True) & vbCrLf
                    
                'vRTFImp = vRTFImp & String(vespacio, "-") & oImpresora.gPrnSaltoLinea
                vRTFImp = vRTFImp & String(vespacio, "-") & vbCrLf
                
                vLineas = vLineas + 3
            End If
        End If
    End If
    
    '**************************************************************************************
    'CARGA DE : Compra y Venta en Dolares y en Soles
    Me.Caption = "Cargando asientos de compra - venta dolares y soles"
    
    Dim ComVta(6, 1) As Currency ' 0-Ctas; 1-Debe;  2-Haber
    Erase ComVta
    Dim ComVtaDolar As String
    ComVtaDolar = ""
    
    ssql = "SELECT  cctacnt, round(sum(ndebe),2) nDebe, round(sum(nhaber),2) nHaber, ctipo " & _
        " FROM AsientoDN WHERE ctipo IN ('1','2') AND datediff(dd, dFecha ,'" & gdHoraGrab & "') = 0 " & _
        " GROUP BY CTIPO, CCTACNT "
        
    'ssql = "SELECT  cctacnt, round(sum(ndebe),2) nDebe, round(sum(nhaber),2) nHaber, ctipo " & _
        " FROM AsientoDN WHERE ctipo IN ('1','2') AND cOpeCod NOT IN ('900022','900023')  AND datediff(dd, dFecha ,'" & gdHoraGrab & "') = 0 " & _
        " GROUP BY CTIPO, CCTACNT " 'JUCS TI-ERS 002-2017
        
    Set RegTmp = oCon.CargaRecordSet(ssql)
    If (RegTmp.BOF Or RegTmp.EOF) Then
    Else
        If vLineas + 19 >= pLineasMax Then
            vPage = vPage + 1
            'vRTFImp = vRTFImp & oImpresora.gPrnSaltoPagina & oImpresora.gPrnSaltoLinea
            vRTFImp = vRTFImp & vbCrLf & vbCrLf
            
            'vRTFImp = vRTFImp & Space(vLenNomb + 35) & "Página :" & ImpreFormat(vPage, 5, 0) & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
            vRTFImp = vRTFImp & Space(vLenNomb + 35) & "Página :" & ImpreFormat(vPage, 5, 0) & vbCrLf & vbCrLf
            
            vLineas = 5
        End If
    
        With RegTmp
            Do While Not .EOF
                If !cTipo = "1" Then
                    If Left(!cctacnt, 4) = lsCajaDolares Then
                        ComVta(1, 0) = !nDebe
                        ComVta(1, 1) = !nHaber
                    ElseIf Left(!cctacnt, 4) = lsCajaSoles Then
                        ComVta(3, 0) = !nDebe
                        ComVta(3, 1) = !nHaber
                    ElseIf Left(!cctacnt, 4) = "4128" Then
                        ComVta(4, 0) = !nDebe
                    ElseIf Left(!cctacnt, 4) = "5128" Then
                        ComVta(4, 1) = !nHaber
                    Else
                        MsgBox "Cuenta no reconocida para Compra y Venta de dolares", vbInformation, " Aviso "
                    End If
                Else
                    If Left(!cctacnt, 4) = lsCajaDolares Then
                        ComVta(5, 0) = !nDebe
                        ComVta(5, 1) = !nHaber
                    Else
                        MsgBox "Cuenta no reconocida para Compra y Venta de dolares", vbInformation, " Aviso "
                    End If
                End If
                .MoveNext
            Loop
        End With
        
        'ComVtaDolar = "" & oImpresora.gPrnSaltoLinea
        ComVtaDolar = "" & vbCrLf
        
        'ComVtaDolar = ComVtaDolar & "   DE : COMPRA Y VENTA EN DOLARES " & oImpresora.gPrnSaltoLinea
        ComVtaDolar = ComVtaDolar & "   DE : COMPRA Y VENTA EN DOLARES " & vbCrLf
        
        'ComVtaDolar = ComVtaDolar & String(vespacio, "-") & oImpresora.gPrnSaltoLinea
        ComVtaDolar = ComVtaDolar & String(vespacio, "-") & vbCrLf
        
        Dim sCajaME As String, sCajaMN As String
        Dim nDebeAcum As Double, nHaberAcum As Double
        
        nDebeAcum = 0
        nHaberAcum = 0
        RegTmp.MoveFirst
        
        Do While Not RegTmp.EOF
            If RegTmp("cTipo") = "2" Then
                sCajaME = RegTmp("cCtaCnt")
                nDebeAcum = nDebeAcum + RegTmp("nDebe")
                nHaberAcum = nHaberAcum + RegTmp("nHaber")
                'ComVtaDolar = ComVtaDolar & ImpreFormat(sCajaME, 21, 1) & ImpreFormat(UCase(CuentaNombre(sCajaME, True, oCon.ConexionActiva)), vLenNomb, 1) & ImpreFormat(Round(RegTmp("nDebe"), 2), 12) & ImpreFormat(Round(RegTmp("nHaber"), 2), 12) & oImpresora.gPrnSaltoLinea
                ComVtaDolar = ComVtaDolar & ImpreFormat(sCajaME, 21, 1) & ImpreFormat(UCase(CuentaNombre(sCajaME, True, oCon.ConexionActiva)), vLenNomb, 1) & ImpreFormat(Round(RegTmp("nDebe"), 2), 12) & ImpreFormat(Round(RegTmp("nHaber"), 2), 12) & vbCrLf
                
            End If
            RegTmp.MoveNext
        Loop
        
        'ComVtaDolar = ComVtaDolar & String(vespacio, "-") & oImpresora.gPrnSaltoLinea
        ComVtaDolar = ComVtaDolar & String(vespacio, "-") & vbCrLf
        
        'ComVtaDolar = ComVtaDolar & Space(vLenNomb + 12) & ImpreFormat("TOTAL", 11, 0) & ImpreFormat(Round(nDebeAcum, 2), 12, , True) & ImpreFormat(Round(nHaberAcum, 2), 12, , True) & oImpresora.gPrnSaltoLinea
        ComVtaDolar = ComVtaDolar & Space(vLenNomb + 12) & ImpreFormat("TOTAL", 11, 0) & ImpreFormat(Round(nDebeAcum, 2), 12, , True) & ImpreFormat(Round(nHaberAcum, 2), 12, , True) & vbCrLf
        
        'ComVtaDolar = ComVtaDolar & String(vespacio, "-") & oImpresora.gPrnSaltoLinea
        ComVtaDolar = ComVtaDolar & String(vespacio, "-") & vbCrLf
        
        'ComVtaDolar = ComVtaDolar & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
        ComVtaDolar = ComVtaDolar & vbCrLf & vbCrLf
        
        'ComVtaDolar = ComVtaDolar & "   DE : COMPRA Y VENTA EN SOLES " & oImpresora.gPrnSaltoLinea
        ComVtaDolar = ComVtaDolar & "   DE : COMPRA Y VENTA EN SOLES " & vbCrLf
        
        'ComVtaDolar = ComVtaDolar & String(vespacio, "-") & oImpresora.gPrnSaltoLinea
        ComVtaDolar = ComVtaDolar & String(vespacio, "-") & vbCrLf
        
        
        nDebeAcum = 0
        nHaberAcum = 0
        RegTmp.MoveFirst
        Do While Not RegTmp.EOF
            If RegTmp("cTipo") = "1" Then
                sCajaMN = RegTmp("cCtaCnt")
                nDebeAcum = nDebeAcum + RegTmp("nDebe")
                nHaberAcum = nHaberAcum + RegTmp("nHaber")
                'ComVtaDolar = ComVtaDolar & ImpreFormat(sCajaMN, 21, 1) & ImpreFormat(UCase(CuentaNombre(sCajaMN, True, oCon.ConexionActiva)), vLenNomb, 1) & ImpreFormat(RegTmp("nDebe"), 12) & ImpreFormat(RegTmp("nHaber"), 12) & oImpresora.gPrnSaltoLinea
                ComVtaDolar = ComVtaDolar & ImpreFormat(sCajaMN, 21, 1) & ImpreFormat(UCase(CuentaNombre(sCajaMN, True, oCon.ConexionActiva)), vLenNomb, 1) & ImpreFormat(RegTmp("nDebe"), 12) & ImpreFormat(RegTmp("nHaber"), 12) & vbCrLf
                
            End If
            RegTmp.MoveNext
        Loop

        'ComVtaDolar = ComVtaDolar & String(vespacio, "-") & oImpresora.gPrnSaltoLinea
        ComVtaDolar = ComVtaDolar & String(vespacio, "-") & vbCrLf
        
        'ComVtaDolar = ComVtaDolar & Space(vLenNomb + 12) & ImpreFormat("TOTAL", 11, 0) & ImpreFormat(nDebeAcum, 12, , True) & ImpreFormat(nHaberAcum, 12, , True) & oImpresora.gPrnSaltoLinea
        ComVtaDolar = ComVtaDolar & Space(vLenNomb + 12) & ImpreFormat("TOTAL", 11, 0) & ImpreFormat(nDebeAcum, 12, , True) & ImpreFormat(nHaberAcum, 12, , True) & vbCrLf
        
        'ComVtaDolar = ComVtaDolar & String(vespacio, "-") & oImpresora.gPrnSaltoLinea
        ComVtaDolar = ComVtaDolar & String(vespacio, "-") & vbCrLf
        
        'ComVtaDolar = ComVtaDolar & oImpresora.gPrnSaltoLinea
        ComVtaDolar = ComVtaDolar & vbCrLf
        
    End If
    RegTmp.Close
    Set RegTmp = Nothing

    Me.Caption = "Asiento finalizado"
    'Envia Asiento al Previo
    vRTFImp = vRTFImp & ComVtaDolar
    If Len(Trim(vRTFImp)) = 0 Then
        MsgBox " No existe ningún asiento generado en este día ", vbInformation, " Aviso "
    End If
    
    rtfImp.Text = vRTFImp
    
    'Graba asiento automaticamente
    If Not (Len(Trim(vObs)) > 0 Or Len(Trim(vNoCtaCnt)) > 0) And pAsiDia And CierreRealizado2(1, oCon.ConexionActiva) Then
                
        If gbAsientoDN Then
            'vRTFImp = oImpresora.gPrnCondensadaON & vRTFImp & oImpresora.gPrnCondensadaOFF
            vRTFImp = vRTFImp
            
            rtfImp.Text = vRTFImp
            Me.rtfImp.SaveFile App.Path & "\Spooler\AP" & Right(gsCodAge, 2) & "_" & Format(gdFecSis, "yyyymmdd") & ".txt", 1
        End If
    End If

End Sub

Private Sub ObtieneTipoCambioLeasing(ByVal pcta As String, ByVal pnTipoCambioCompraL As Currency, ByVal pnTipoCambioVentaL As Currency, ByVal pnTipoCambioFijoL As Currency)
    If Mid(pcta, 6, 3) = "515" Or Mid(pcta, 6, 3) = "516" Then
            Call ObtenerTipoCambioLeasing(pcta, oCon.ConexionActiva, pnTipoCambioCompraL, pnTipoCambioVentaL, pnTipoCambioFijoL)
    End If
End Sub


Private Sub validacion86()
        If nResult86 <> nResult85 Then
            vDife8 = 0
            vDife8 = IIf(nResult85 > nResult86, nResult85 - nResult86, nResult86 - nResult85)
            If Abs(nResult85) > Abs(nResult86) Then
                'SI 85 MAYOR
                If vCta85D > vCta85H Then
                    'Modificado PASI20140513 TI-ERS004-2014
                    Call IngresarAsientoDN(gdHoraGrab, "8528", 0, vDife8, 3, vAgencia, , , , oCon.ConexionActiva)
                    Call IngresarAsientoDN(gdHoraGrab, lsCVMEGanacia, vDife8, 0, 3, vAgencia, , , , oCon.ConexionActiva)
                Else
                    'Modificado PASI20140513 TI-ERS004-2014
                    Call IngresarAsientoDN(gdHoraGrab, "8528", vDife8, 0, 3, vAgencia, , , , oCon.ConexionActiva)
                    Call IngresarAsientoDN(gdHoraGrab, lsCVMEGanacia, 0, vDife8, 3, vAgencia, , , , oCon.ConexionActiva)
                    'end PASI
                End If
            Else
                'SI 86 MAYOR
                If vCta86D > vCta86H Then
                'Modificado PASI20140513 TI-ERS004-2014
                    Call IngresarAsientoDN(gdHoraGrab, "8528", 0, vDife8, 3, vAgencia, , , , oCon.ConexionActiva)
                    Call IngresarAsientoDN(gdHoraGrab, lsCVMEPerdida, vDife8, 0, 3, vAgencia, , , , oCon.ConexionActiva)
                    
                'end PASI
                Else
                'Modificado PASI20141305 TI-ERS004-2014
                    Call IngresarAsientoDN(gdHoraGrab, "8528", vDife8, 0, 3, vAgencia, , , , oCon.ConexionActiva)
                    Call IngresarAsientoDN(gdHoraGrab, lsCVMEPerdida, 0, vDife8, 3, vAgencia, , , , oCon.ConexionActiva)
                    'end PASI
                End If
            End If
        End If
End Sub

Private Sub Validacion()

        If nResult82 <> nResult81 Then
            vDife8 = 0
            vDife8 = IIf(nResult81 > nResult82, nResult81 - nResult82, nResult82 - nResult81)
            'vDife8 = Abs((vCta82D + vCta82H) - (vCta81D + vCta81H))
            If Abs(nResult82) > Abs(nResult81) Then
                'SI 82 MAYOR
                If vCta82D > vCta82H Then
                    ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge) " & _
                        " VALUES('" & gdHoraGrab & "','8221',0," & vDife8 & ",'3','" & vAgencia & "')"
                    oCon.ejecutar ssql

                    ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge) " & _
                        " VALUES('" & gdHoraGrab & "','" & lsCVMEGanacia & "'," & vDife8 & ",0,'3','" & vAgencia & "')"
                    oCon.ejecutar ssql
                Else
                    ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge) " & _
                        " VALUES('" & gdHoraGrab & "','8221'," & vDife8 & ",0,'3','" & vAgencia & "')"
                    oCon.ejecutar ssql

                    ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge) " & _
                        " VALUES('" & gdHoraGrab & "','" & lsCVMEGanacia & "',0," & vDife8 & ",'3','" & vAgencia & "')"
                    oCon.ejecutar ssql
                End If
            Else
                'SI 81 MAYOR
                If vCta81D > vCta81H Then
                    ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge) " & _
                        " VALUES('" & gdHoraGrab & "','8221',0," & vDife8 & ",'3','" & vAgencia & "')"
                    oCon.ejecutar ssql

                    ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge) " & _
                        " VALUES('" & gdHoraGrab & "','" & lsCVMEPerdida & "'," & vDife8 & ",0,'3','" & vAgencia & "')"
                    oCon.ejecutar ssql
                Else
                    ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge) " & _
                        " VALUES('" & gdHoraGrab & "','8221'," & vDife8 & ",0,'3','" & vAgencia & "')"
                    oCon.ejecutar ssql

                    ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge) " & _
                        " VALUES('" & gdHoraGrab & "','" & lsCVMEPerdida & "',0," & vDife8 & ",'3','" & vAgencia & "')"
                    oCon.ejecutar ssql
                End If
            End If
        End If

End Sub

Private Sub BuscaCtasCntEnPlanCon()
    '*********** PEAC 20110705 - VERIFICA EXISTENCIA DE CTAS CONT EN EL PLAN CONT.
    Dim rsCtas As ADODB.Recordset
    Set rsCtas = New ADODB.Recordset
    ssql = "exec stp_sel_BuscaCtaContPlan '" & Format(txtFecha, "yyyymmdd") & "'"
    Set rsCtas = oCon.CargaRecordSet(ssql)
    If Not (rsCtas.BOF Or rsCtas.EOF) Then
        Do While Not rsCtas.EOF
        'ALPA 20130527***********************************
        'vObs = vObs & oImpresora.gPrnSaltoLinea & "** OJO ** Cta.Inexist.Plan Cont>> Cta.Cont:'" & rsCtas!cCtaCnt & "' Importe: '" & Trim(CStr(rsCtas!Importe)) & "' Age:'" & Trim(rsCtas!cCodAge) & "' Ope:'" & Trim(rsCtas!cOpecod) & "' Cod.Cta.'" & IIf(IsNull(rsCtas!cCtaCod), "", rsCtas!cCtaCod) & "'"
        If rsCtas!cOpeCod <> "700104" Then
            'vObs = vObs & oImpresora.gPrnSaltoLinea & "** OJO ** Cta.Inexist.Plan Cont>> Cta.Cont:'" & rsCtas!cctacnt & "' Importe: '" & Trim(CStr(rsCtas!Importe)) & "' Age:'" & Trim(rsCtas!cCodAge) & "' Ope:'" & Trim(rsCtas!cOpeCod) & "' Cod.Cta.'" & IIf(IsNull(rsCtas!cCtaCod), "", rsCtas!cCtaCod) & "'"
            vObs = vObs & vbCrLf & "** OJO ** Cta.Inexist.Plan Cont>> Cta.Cont:'" & rsCtas!cctacnt & "' Importe: '" & Trim(CStr(rsCtas!Importe)) & "' Age:'" & Trim(rsCtas!cCodAge) & "' Ope:'" & Trim(rsCtas!cOpeCod) & "' Cod.Cta.'" & IIf(IsNull(rsCtas!cCtaCod), "", rsCtas!cCtaCod) & "'"
            
        End If
        '************************************************
        rsCtas.MoveNext
        Loop
    End If
    rsCtas.Close
    
    '*********** FIN PEAC
End Sub

Private Sub cmdImprimir_Click()

    Dim vNoCtaCnt As String
    Dim vEst As String
    Dim vCad As String
    Dim x As Long
    Dim vespacio As Long, vLenNomb As Long
    Dim RegTmp As ADODB.Recordset
    Dim tmpSql As String
    Dim TCF As Currency, TCC As Currency, TCV As Currency
    Dim lsCajaDolares As String
    Dim lsCajaSoles As String
    
    Dim cNomArchivoTxt As String
    cNomArchivoTxt = ""
    
    'Dim oAsi As COMDCajaGeneral.DCOMAsiento
    Set oAsi = New COMDCajaGeneral.DCOMAsiento
    Dim oPrevio As previo.clsprevio
    'Set oPrevio = New previo.clsprevio
    vLenNomb = 70
    vespacio = vLenNomb + 54
    
    lsCajaSoles = oAsi.GetAsientoParametro(4)
    lsCajaDolares = oAsi.GetAsientoParametro(5)
    
    '*************************************************************************************
    '*************************************************************************************
    Me.Caption = "Impresión de Asientos Contables"
    'GENERA IMPRESION DE LOS ASIENTOS
    Dim vIndice As Long  'contador de Item
    Dim vLineas As Long
    Dim vPage As Long
    Dim vMoneAnte As String
    Dim vSumDebe As Currency
    Dim vSumHaber As Currency
    Dim vProduc As String
    Dim vDolar As Boolean
    Dim sSqlDolar As String
    Dim vCabecera As String
    
    vRTFImp = ""
    
    'Para la validacion Dia y hora de Generación
    dHoraGrab = Format(gdFecSis & " " & Time, gsFormatoFechaHoraView)
    'Carga fecha del Asiento y hora de grabación
    gdHoraGrab = Format(txtFecha & " " & Time, gsFormatoFechaHora)
    
    vDolar = False
    Dim vDiaSinOpe  As String
    'VALIDACION DE ASIENTOS
    'Dim vObs As String
    vCabecera = "NRO. DE CUENTA                                 DESCRIPCION                                          DEBE          HABER" & vbCrLf
    Dim rsAge As ADODB.Recordset
    Dim oAre As COMDConstantes.DCOMActualizaDatosArea
    Set rsAge = New ADODB.Recordset
    Set oAre = New COMDConstantes.DCOMActualizaDatosArea
    Set rsAge = oAre.getAgencias(, False)
    'Set rsAge = oAre.GetAgenciasValidado(, False)
    vObs = ""
    Do While Not rsAge.EOF
        Me.Caption = "Validando asientos de :" & Trim(rsAge!Descripcion)
        vObs = vObs & ValidaOk(txtFecha, True, rsAge!codigo)
        rsAge.MoveNext
        DoEvents
    Loop
    'Verifica la veracidad del asiento contable
    Me.Caption = "Verificando Veracidad de asiento Contable"
    If Len(Trim(vObs)) > 0 Or Len(Trim(vNoCtaCnt)) > 0 Then
        MsgBox "El asiento presenta Observaciones ! ", vbInformation, " Aviso "
        vProduc = vbCrLf & Space(40) & "  A S I E N T O     ¡   N O     V A L I D O   !" & vbCrLf
        vProduc = vProduc & "   OBSERVACIONES DEL ASIENTO : " & vbCrLf & _
                            "  ============================="
        If Len(Trim(vNoCtaCnt)) > 0 Then vProduc = vProduc & vNoCtaCnt & vbCrLf
        If Len(Trim(vObs)) > 0 Then vProduc = vProduc & vObs & vbCrLf
        vProduc = vProduc & vbCrLf & vbCrLf
        vLineas = Repetido(vProduc, vbCrLf)
    End If
    'Determina Productos
    vProduc = vProduc & "   DE : " & vbCrLf & "  ======" & vbCrLf
    If chkAho.value = 1 Then vProduc = vProduc & "     *  AHORROS " & vbCrLf
    If chkCre.value = 1 Then vProduc = vProduc & "     *  CREDITOS " & vbCrLf
    If chkPig.value = 1 Then vProduc = vProduc & "     *  CREDITO PIGNORATICIO " & vbCrLf
    ssql = "SELECT  cctacnt, round(sum(ndebe),2) nDebe, round(sum(nhaber),2) nHaber, substring(cctacnt,3,1) AS MONEDA " & _
        " FROM AsientoDN "
    
    'Verifica que condiciones se ha de tener en cuenta:  Soles - Dolares
    vEst = "": vCad = ""
    If chkMonSol.value = 1 Then vEst = vEst & Moneda.gMonedaNacional
    If chkMonDol.value = 1 Then vEst = vEst & Moneda.gMonedaExtranjera
    For x = 1 To Len(vEst) Step 1
        vCad = vCad & "'" & Mid(vEst, x, 1) & "'"
        If x <> Len(vEst) Then vCad = vCad & ","
    Next x
    
    If Len(vEst) > 0 Then
        '        sSqlDolar = sSql & " WHERE cTipo IN ('3') AND datediff(dd,dFecha,'" & gdHoraGrab & "') = 0  GROUP BY CTIPO, CCTACNT ORDER BY cctacnt "
        sSqlDolar = ssql & " WHERE cTipo IN ('3') AND datediff(dd,dFecha,'" & gdHoraGrab & "') = 0  And substring(cctacnt,3,1) in ('1','2')" & _
            " GROUP BY CTIPO, CCTACNT ORDER BY cctacnt "
        ssql = ssql & " WHERE cTipo IN ('0') AND substring(cctacnt,3,1) IN (" & vCad & ") AND datediff(dd,dFecha,'" & gdHoraGrab & "') = 0  "
        If chkMonSol.value = 1 And chkMonDol.value = 1 Then
            ssql = ssql & " GROUP BY CTIPO, CCTACNT ORDER BY moneda , cctacnt "
        Else
            ssql = ssql & " GROUP BY CTIPO, CCTACNT ORDER BY cctacnt "
        End If
    End If
    
    vMoneAnte = ""
    vLineas = vLineas + 12: vPage = 1
    vRTFImp = gsNomCmac & Space(vLenNomb + 6) & Format(gdFecSis & " " & Time, "dd/mm/yyyy hh:mm:ss") & vbCrLf
    vRTFImp = vRTFImp & Space(vLenNomb + 16) & " Página :" & ImpreFormat(vPage, 5, 0) & vbCrLf
    vRTFImp = vRTFImp & ImpreFormat(UCase(gsNomAge), 25) & vbCrLf
    vRTFImp = vRTFImp & ImpreFormat("   ASIENTO CONTABLE DEL DIA " & Format(txtFecha, "dd/mm/yyyy"), 44, 43) & vbCrLf
    vRTFImp = vRTFImp & ImpreFormat(String(40, "="), 44, 42) & vbCrLf
    If CDate(txtFecha) = gdFecSis Then
        If Not CierreRealizado(1) Then
            vRTFImp = vRTFImp & ImpreFormat(" NO VALIDO (ANTES DEL CIERRE DE DIA)", 40, 45) & vbCrLf
        Else
            'Para los dias que no hay movimientos
            vDiaSinOpe = "Select cMovNro From Mov M Where M.cMovNro Like '" & Format(gdHoraGrab, gsFormatoMovFecha) & "%' AND M.nMovFlag = " & MovFlag.gMovFlagVigente
            Set RegTmp = oCon.CargaRecordSet(vDiaSinOpe)
            If (RegTmp.BOF Or RegTmp.EOF) Then
                vRTFImp = vRTFImp & vbCr & Space(28) & "N O   S E   R E A L I Z A R O N   O P E R A C I O N E S   E N   E L   D I A"
            Else
                If VerificaDiaHabil(gdHoraGrab, 3) And Not CierreRealizado(2) Then
                    vRTFImp = vRTFImp & ImpreFormat(" NO VALIDO (ANTES DEL CIERRE DE MES)", 40, 45) & vbCrLf
                End If
            End If
            RegTmp.Close
            Set RegTmp = Nothing
        End If
    End If
    vRTFImp = vRTFImp & vProduc
    
    Set RegTmp = oCon.CargaRecordSet(ssql)
    If (RegTmp.BOF Or RegTmp.EOF) Then
        RegTmp.Close
        Set RegTmp = Nothing
    Else
        With RegTmp
        Do While Not .EOF
            If !Moneda <> vMoneAnte Then
                If !Moneda = Moneda.gMonedaNacional Then
                    vRTFImp = vRTFImp & vbCrLf & "   MONEDA : SOLES" & vbCrLf
                    vLineas = vLineas + 1
                Else
                    If chkMonSol.value = 1 And (vSumDebe > 0 Or vSumHaber > 0) Then
                        vRTFImp = vRTFImp & String(vespacio, "-") & vbCrLf
                        vRTFImp = vRTFImp & Space(vLenNomb + 12) & ImpreFormat("TOTAL", 11, 0) & _
                            ImpreFormat(vSumDebe, 12, , True) & ImpreFormat(vSumHaber, 12, , True) & vbCrLf
                        vRTFImp = vRTFImp & String(vespacio, "-") & vbCrLf
                        vSumDebe = 0: vSumHaber = 0
                        vLineas = vLineas + 3
                    End If
                    vDolar = True
                    If chkMonSol.value = 1 Then
                        vPage = vPage + 1
                        vRTFImp = vRTFImp & vbCrLf & vbCrLf
                        vRTFImp = vRTFImp & Space(vLenNomb + 35) & "Página :" & ImpreFormat(vPage, 5, 0) & vbCrLf
                    End If
                    vRTFImp = vRTFImp & vbCrLf & "   MONEDA : DOLARES" & vbCrLf
                    vLineas = 5 'vLineas + 2
                End If
                vRTFImp = vRTFImp & String(vespacio, "-") & vbCrLf
                vRTFImp = vRTFImp & Space(1) & vCabecera
                vRTFImp = vRTFImp & String(vespacio, "-") & vbCrLf
                vLineas = vLineas + 3
            End If
            vMoneAnte = !Moneda
            vRTFImp = vRTFImp & ImpreFormat(!cctacnt, 21, 1) & ImpreFormat(UCase(CuentaNombre(!cctacnt, True, oCon.ConexionActiva)), vLenNomb, 1) & ImpreFormat(Round(!nDebe, 2), 12, , True) & ImpreFormat(Round(!nHaber, 2), 12, , True) & vbCrLf
            vSumDebe = vSumDebe + Round(!nDebe, 2)
            vSumHaber = vSumHaber + Round(!nHaber, 2)
            vLineas = vLineas + 1
            If vLineas > pLineasMax Then
                vPage = vPage + 1
                vRTFImp = vRTFImp & vbCrLf & vbCrLf
                vRTFImp = vRTFImp & Space(vLenNomb + 35) & "Página :" & ImpreFormat(vPage, 5, 0) & vbCrLf & vbCrLf
                vRTFImp = vRTFImp & String(vespacio, "-") & vbCrLf
                vRTFImp = vRTFImp & Space(1) & vCabecera
                vRTFImp = vRTFImp & String(vespacio, "-") & vbCrLf
                vLineas = 5
            End If
            .MoveNext
        Loop
        End With
        vRTFImp = vRTFImp & String(vespacio, "-") & vbCrLf
        vRTFImp = vRTFImp & Space(vLenNomb + 12) & ImpreFormat("TOTAL", 11, 0) & _
            ImpreFormat(vSumDebe, 12, , True) & ImpreFormat(vSumHaber, 12, , True) & vbCrLf
        vRTFImp = vRTFImp & String(vespacio, "-") & vbCrLf
        vLineas = vLineas + 3
        RegTmp.Close
        Set RegTmp = Nothing
        
        If vDolar Then
            Set RegTmp = oCon.CargaRecordSet(sSqlDolar)
            If (RegTmp.BOF Or RegTmp.EOF) Then
                RegTmp.Close
                Set RegTmp = Nothing
            Else
                If vLineas + 8 >= pLineasMax Then
                    vPage = vPage + 1
                    vRTFImp = vRTFImp & vbCrLf & vbCrLf
                    vRTFImp = vRTFImp & Space(vLenNomb + 35) & "Página :" & ImpreFormat(vPage, 5, 0) & vbCrLf & vbCrLf
                    vLineas = 5
                End If
                vRTFImp = vRTFImp & vbCrLf & "   CONVERSION DE DOLARES" & vbCrLf
                vRTFImp = vRTFImp & "     *  TIPO CAMBIO FIJO   : " & ImpreFormat(TCF, 6, 3) & vbCrLf
                vRTFImp = vRTFImp & "     *  TIPO CAMBIO VENTA  : " & ImpreFormat(TCV, 6, 3) & vbCrLf
                vRTFImp = vRTFImp & "     *  TIPO CAMBIO COMPRA : " & ImpreFormat(TCC, 6, 3) & vbCrLf
                vRTFImp = vRTFImp & String(vespacio, "-") & vbCrLf
                vRTFImp = vRTFImp & Space(1) & vCabecera
                vRTFImp = vRTFImp & String(vespacio, "-") & vbCrLf
                vLineas = vLineas + 8
                vSumDebe = 0: vSumHaber = 0
                With RegTmp
                Do While Not .EOF
                    vRTFImp = vRTFImp & ImpreFormat(!cctacnt, 21, 1) & ImpreFormat(UCase(CuentaNombre(!cctacnt, True, oCon.ConexionActiva)), vLenNomb, 1) & ImpreFormat(Round(!nDebe, 2), 12, , True) & ImpreFormat(Round(!nHaber, 2), 12, , True) & vbCrLf
                    vSumDebe = vSumDebe + !nDebe
                    vSumHaber = vSumHaber + !nHaber
                    vLineas = vLineas + 1
                    If vLineas > pLineasMax Then
                        vPage = vPage + 1
                        vRTFImp = vRTFImp & vbCrLf & vbCrLf
                        vRTFImp = vRTFImp & Space(vLenNomb + 35) & "Página :" & ImpreFormat(vPage, 5, 0) & vbCrLf & vbCrLf
                        vRTFImp = vRTFImp & String(vespacio, "-") & vbCrLf
                        vRTFImp = vRTFImp & Space(1) & vCabecera
                        vRTFImp = vRTFImp & String(vespacio, "-") & vbCrLf
                        vLineas = 5
                    End If
                    .MoveNext
                Loop
                End With
                RegTmp.Close
                Set RegTmp = Nothing
                If vLineas + 3 > pLineasMax Then
                    vPage = vPage + 1
                    vRTFImp = vRTFImp & vbCrLf & vbCrLf
                    vRTFImp = vRTFImp & Space(vLenNomb + 35) & "Página :" & ImpreFormat(vPage, 5, 0) & vbCrLf & vbCrLf
                    vLineas = 5
                End If
                vRTFImp = vRTFImp & String(vespacio, "-") & vbCrLf
                vRTFImp = vRTFImp & Space(vLenNomb + 12) & ImpreFormat("TOTAL", 11, 0) & _
                    ImpreFormat((vSumDebe), 12, , True) & ImpreFormat((vSumHaber), 12, , True) & vbCrLf
                vRTFImp = vRTFImp & String(vespacio, "-") & vbCrLf
                vLineas = vLineas + 3
            End If
        End If
    End If
    
    ejecutar ("Play ") 'PEAC 20160701
    
    '**************************************************************************************
    'CARGA DE : Compra y Venta en Dolares y en Soles
    Me.Caption = "Cargando asientos de compra - venta dolares y soles"
    Dim ComVta(6, 1) As Currency ' 0-Ctas; 1-Debe;  2-Haber
    Erase ComVta
    Dim ComVtaDolar As String
    ComVtaDolar = ""
    
    ssql = "SELECT  cctacnt, round(sum(ndebe),2) nDebe, round(sum(nhaber),2) nHaber, ctipo " & _
        " FROM AsientoDN WHERE ctipo IN ('1','2') AND datediff(dd, dFecha ,'" & gdHoraGrab & "') = 0 " & _
        " GROUP BY CTIPO, CCTACNT "
        
    'ssql = "SELECT  cctacnt, round(sum(ndebe),2) nDebe, round(sum(nhaber),2) nHaber, ctipo " & _
        " FROM AsientoDN WHERE ctipo IN ('1','2') AND cOpeCod NOT IN ('900022','900023') AND datediff(dd, dFecha ,'" & gdHoraGrab & "') = 0 " & _
        " GROUP BY CTIPO, CCTACNT " 'JUCS TI-ERS-002-2017

        
    Set RegTmp = oCon.CargaRecordSet(ssql)
    If (RegTmp.BOF Or RegTmp.EOF) Then
    Else
        If vLineas + 19 >= pLineasMax Then
            vPage = vPage + 1
            vRTFImp = vRTFImp & vbCrLf & vbCrLf
            vRTFImp = vRTFImp & Space(vLenNomb + 35) & "Página :" & ImpreFormat(vPage, 5, 0) & vbCrLf & vbCrLf
            vLineas = 5
        End If
    
        With RegTmp
            Do While Not .EOF
                If !cTipo = "1" Then
                    If Left(!cctacnt, 4) = lsCajaDolares Or Left(vCodConta, 4) = "2524" Then 'APRI20170705 RFC1705220001 Add Left(vCodConta, 4) = "2524"
                        ComVta(1, 0) = !nDebe
                        ComVta(1, 1) = !nHaber
                    ElseIf Left(!cctacnt, 4) = lsCajaSoles Or Left(vCodConta, 4) = "2514" Then 'APRI20170705 RFC1705220001 Add Left(vCodConta, 4) = "2514"
                        ComVta(3, 0) = !nDebe
                        ComVta(3, 1) = !nHaber
                    ElseIf Left(!cctacnt, 4) = "4128" Then
                        ComVta(4, 0) = !nDebe
                    ElseIf Left(!cctacnt, 4) = "5128" Then
                        ComVta(4, 1) = !nHaber
                    Else
                        MsgBox "Cuenta no reconocida para Compra y Venta de dolares", vbInformation, " Aviso "
                    End If
                Else
                    If Left(!cctacnt, 4) = lsCajaDolares Or Left(vCodConta, 4) = "2524" Then 'APRI20170705 RFC1705220001 Add Left(vCodConta, 4) = "2524"
                        ComVta(5, 0) = !nDebe
                        ComVta(5, 1) = !nHaber
                    Else
                        MsgBox "Cuenta no reconocida para Compra y Venta de dolares", vbInformation, " Aviso "
                    End If
                End If
                .MoveNext
            Loop
        End With
        
        ComVtaDolar = "" & vbCrLf
        ComVtaDolar = ComVtaDolar & "   DE : COMPRA Y VENTA EN DOLARES " & vbCrLf
        ComVtaDolar = ComVtaDolar & String(vespacio, "-") & vbCrLf
        Dim sCajaME As String, sCajaMN As String
        Dim nDebeAcum As Double, nHaberAcum As Double
        
        nDebeAcum = 0
        nHaberAcum = 0
        RegTmp.MoveFirst
        
        Do While Not RegTmp.EOF
            If RegTmp("cTipo") = "2" Then
                sCajaME = RegTmp("cCtaCnt")
                nDebeAcum = nDebeAcum + RegTmp("nDebe")
                nHaberAcum = nHaberAcum + RegTmp("nHaber")
                ComVtaDolar = ComVtaDolar & ImpreFormat(sCajaME, 21, 1) & ImpreFormat(UCase(CuentaNombre(sCajaME, True, oCon.ConexionActiva)), vLenNomb, 1) & ImpreFormat(Round(RegTmp("nDebe"), 2), 12) & ImpreFormat(Round(RegTmp("nHaber"), 2), 12) & vbCrLf
            End If
            RegTmp.MoveNext
        Loop
        
        ComVtaDolar = ComVtaDolar & String(vespacio, "-") & vbCrLf
        ComVtaDolar = ComVtaDolar & Space(vLenNomb + 12) & ImpreFormat("TOTAL", 11, 0) & _
                    ImpreFormat(Round(nDebeAcum, 2), 12, , True) & _
                    ImpreFormat(Round(nHaberAcum, 2), 12, , True) & vbCrLf
        ComVtaDolar = ComVtaDolar & String(vespacio, "-") & vbCrLf
        
        ComVtaDolar = ComVtaDolar & vbCrLf & vbCrLf
        ComVtaDolar = ComVtaDolar & "   DE : COMPRA Y VENTA EN SOLES " & vbCrLf
        ComVtaDolar = ComVtaDolar & String(vespacio, "-") & vbCrLf
        
        nDebeAcum = 0
        nHaberAcum = 0
        RegTmp.MoveFirst
        Do While Not RegTmp.EOF
            If RegTmp("cTipo") = "1" Then
                sCajaMN = RegTmp("cCtaCnt")
                nDebeAcum = nDebeAcum + RegTmp("nDebe")
                nHaberAcum = nHaberAcum + RegTmp("nHaber")
                ComVtaDolar = ComVtaDolar & ImpreFormat(sCajaMN, 21, 1) & ImpreFormat(UCase(CuentaNombre(sCajaMN, True, oCon.ConexionActiva)), vLenNomb, 1) & ImpreFormat(RegTmp("nDebe"), 12) & ImpreFormat(RegTmp("nHaber"), 12) & vbCrLf
            End If
            RegTmp.MoveNext
        Loop

        ComVtaDolar = ComVtaDolar & String(vespacio, "-") & vbCrLf
        ComVtaDolar = ComVtaDolar & Space(vLenNomb + 12) & ImpreFormat("TOTAL", 11, 0) & _
                    ImpreFormat(nDebeAcum, 12, , True) & _
                    ImpreFormat(nHaberAcum, 12, , True) & vbCrLf
        ComVtaDolar = ComVtaDolar & String(vespacio, "-") & vbCrLf
        ComVtaDolar = ComVtaDolar & vbCrLf
    End If
    RegTmp.Close
    Set RegTmp = Nothing
    Me.Caption = "Asiento finalizado"
    'Envia Asiento al Previo
    vRTFImp = vRTFImp & ComVtaDolar
    If Len(Trim(vRTFImp)) = 0 Then
        MsgBox " No existe ningún asiento generado en este día ", vbInformation, " Aviso "
    End If
    
    rtfImp.Text = vRTFImp
    
    'Graba asiento automaticamente
    If Not (Len(Trim(vObs)) > 0 Or Len(Trim(vNoCtaCnt)) > 0) And pAsiDia And CierreRealizado(1) Then
        If gbAsientoDN Then
            vRTFImp = oImpresora.gPrnCondensadaON & vRTFImp & oImpresora.gPrnCondensadaOFF
            rtfImp.Text = vRTFImp
            Me.rtfImp.SaveFile App.Path & "\Spooler\AP" & Right(gsCodAge, 2) & "_" & Format(gdFecSis, "yyyymmdd") & ".txt", 1
        End If
    End If
    
    cNomArchivoTxt = "AC" & Format(txtFecha, "yyyymmdd") & "_" & Trim(str(Int((9999 - 1) * Rnd() + 1) * 100)) & ".txt"
    
    Call GeneraRepoAsntoCnt(vRTFImp, cNomArchivoTxt)
    
    If optImpresion(0).value = True Then
'        oPrevio.Show rtfImp.Text, " ASIENTOS CONTABLES DEL DIA " & Format(gdHoraGrab, "dd/mm/yyyy"), True, 66
        'Shell "notepad C:\install\prueba.txt", vbNormalFocus
        Shell "notepad C:\Install\Spooler\" & cNomArchivoTxt, vbNormalFocus
        
    Else
        ImpreBegin True, 66
            Print #ArcSal, rtfImp.Text
        ImpreEnd
    End If

    mciExecute "Close All" 'PEAC 20170201

    'TRANSFERENCIA DEL ASIENTO - SOLO CUANDO ES EL DIA
    'Verifica la veracidad del asiento contable
    If CDate(txtFecha) = gdFecSis Then
        If Not (Len(Trim(vObs)) > 0 Or Len(Trim(vNoCtaCnt)) > 0) And CierreRealizado(1) Then
            If MsgBox(" ¿ Desea migrar el Asiento ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
                Call TransfiereAsiento
                MsgBox "Asiento Transferido satisfactoriamente...!", vbInformation, "¡Aviso!"
            End If
        End If
    Else
        If Not (Len(Trim(vObs)) > 0 Or Len(Trim(vNoCtaCnt)) > 0) Then
            If MsgBox(" ¿ Desea migrar el Asiento ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
                Call TransfiereAsiento
                MsgBox "Asiento Transferido satisfactoriamente...!", vbInformation, "¡Aviso!"
            End If
        End If
    End If
    MousePointer = 0
    
    
    Exit Sub



    

End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

'Parametros para el formulario
Private Sub CargaParametros()
    
    Dim oConst As COMDConstSistema.NCOMConstSistema
    Set oConst = New COMDConstSistema.NCOMConstSistema
    
    gdFecSis = oConst.LeeConstSistema(16)
    
    Set oConst = Nothing
    
    'gdFecSis
    gsCodAge = "01"
    gsCodCMAC = "109"
    Me.txtFecha.Text = gdFecSis
    
    pPrevioMax = 4000
    pLineasMax = 56
    pHojaFiMax = 66
    
    
   'Call correasiento
      
    'cmdAsientos_Click
    'cmdImprimir_Click
    
End Sub
Private Sub correasiento()
    MsgBox "aqui comienza el asnto cnt"
End Sub
 
Private Sub Form_Load()
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    oCon.CommadTimeOut = 0
    CargaParametros
    txtFecha.Text = gdFecSis
    If pAsiDia Then
        txtFecha.Visible = False
        lblEtiqueta.Visible = False
    Else
        fraContenedor(0).Top = fraContenedor(0).Top + 350
        prgList.Top = prgList.Top + 350
        cmdSalir.Top = cmdSalir.Top + 350
        Me.Height = Me.Height + 350
    End If
    
    '*** PEAC 20191217
    'Call cmdAsientos_Click
    
    'Me.Icon = LoadPicture(App.path & gsRutaIcono)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'CierraConexion
End Sub

Private Sub txtFecha_GotFocus()
    fEnfoque txtFecha
End Sub

Private Sub txtFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If ValFecha(txtFecha) Then cmdAsientos.SetFocus
    End If
End Sub

Private Sub txtFecha_Validate(Cancel As Boolean)
    If Not ValFecha(txtFecha) Then
        Cancel = True
    End If
End Sub

Public Sub Inicio(Optional ByVal pAsientoDiario As Boolean = True)
    pAsiDia = pAsientoDiario
    Me.Show 1
End Sub

Private Sub TransfiereAsiento()
    Dim ssql As String
    Dim rs As New ADODB.Recordset
    Dim aAsientoS() As String
    Dim aAsientoD() As String
    Dim aAsientoV() As String
    Dim vdFecha     As Date
    Dim nPos        As Integer
    ReDim aAsientoS(1 To 3, 0 To 0)
    ReDim aAsientoD(1 To 5, 0 To 0)
    ReDim aAsientoV(1 To 5, 0 To 0)
    vdFecha = CDate(txtFecha)
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    
   If Not oCon.AbreConexion Then
      MsgBox "No se puede conectar con Ag. Sede." & vbCr & "No se trasfirió Asiento Contable", vbInformation, "¡Aviso!"
      Exit Sub
   End If
   
   
'   'JUCS TI-ERS-002-2017
'   ssql = "delete from AsientoDN where convert(date,dFecha)>'20171018' AND cCtaCnt LIKE '51280101%'"
'   oCon.CargaRecordSet (ssql)
'   ssql = "delete from AsientoDN where convert(date,dFecha)>'20171018' AND cCtaCnt LIKE '41280101%'"
'   oCon.CargaRecordSet (ssql)
'   ssql = "delete from AsientoDN where convert(date,dFecha)>'20171018' AND cOpeCod IN('900022','900023')"
'   oCon.CargaRecordSet (ssql)
'   'JUCS TI-ERS-002-2017
   
   ssql = "delete from AsientoDN where datediff(dd,dfecha,'" & Format(vdFecha, "mm/dd/yyyy") & "')= 0 and cOpeCod='700104'"
   oCon.CargaRecordSet (ssql)
   
   'ssql = " SELECT cCtaCnt, cTipo, Round(SUM(nDebe),2) as nDebe, Round(SUM(nHaber),2) as nHaber " _
        & " FROM AsientoDN WHERE datediff(dd,dfecha,'" & Format(vdFecha, "mm/dd/yyyy") & "')= 0 " _
        & " GROUP BY cCtaCnt, cTipo ORDER BY cCtaCnt "
        
    'JUCS - 20171020
    'ssql = " SELECT cCtaCnt, cTipo, Round(SUM(nDebe),2) as nDebe, Round(SUM(nHaber),2) as nHaber " _
        & " FROM AsientoDN WHERE datediff(dd,dfecha,'" & Format(vdFecha, "mm/dd/yyyy") & "')= 0  " _
        & "AND cOpeCod NOT IN('900022','900023') AND cCtaCnt NOT like '51280101%' AND cCtaCnt NOT like '41280101%' " _
        & " GROUP BY cCtaCnt, cTipo ORDER BY cCtaCnt "
        
' CAMBIO JUCS TI-ERS-002-2017 - 20171022
   If Format(vdFecha, "yyyyMMdd") < "20171019" Then
        ssql = " SELECT cCtaCnt, cTipo, Round(SUM(nDebe),2) as nDebe, Round(SUM(nHaber),2) as nHaber " _
        & " FROM AsientoDN WHERE datediff(dd,dfecha,'" & Format(vdFecha, "mm/dd/yyyy") & "')= 0  " _
        & " GROUP BY cCtaCnt, cTipo ORDER BY cCtaCnt "
   Else
       ssql = "EXEC stp_sel_ObtenerRegistrosAsientoDN  '" & Format(vdFecha, "yyyymmdd") & "' "
   End If
        
   If rs.State = adStateOpen Then rs.Close: Set rs = Nothing
   rs.CursorLocation = adUseClient
   
   Set rs = oCon.CargaRecordSet(ssql)
   
   rs.ActiveConnection = Nothing
   If rs.EOF Then
      MsgBox "No se registraron Movimientos en Agencia el día " & vdFecha & "...!", vbInformation, "¡AViso!"
   Else
      Do While Not rs.EOF

         Select Case rs!cTipo
         Case "0"
            If Mid(rs!cctacnt, 3, 1) = "1" Then
               ReDim Preserve aAsientoS(1 To 3, 0 To UBound(aAsientoS, 2) + 1)
               nPos = UBound(aAsientoS, 2)
               aAsientoS(1, nPos) = rs!cctacnt
               aAsientoS(2, nPos) = Round(rs!nDebe, 2)
               aAsientoS(3, nPos) = Round(rs!nHaber, 2)
            Else
               nPos = BuscaMatriz(aAsientoD, rs!cctacnt, 2)
               If nPos = -1 Then
                  ReDim Preserve aAsientoD(1 To 5, 0 To UBound(aAsientoD, 2) + 1)
                  nPos = UBound(aAsientoD, 2)
               End If
               aAsientoD(1, nPos) = rs!cctacnt
               aAsientoD(4, nPos) = Round(rs!nDebe, 2)
               aAsientoD(5, nPos) = Round(rs!nHaber, 2)
            End If
         Case "1"
            nPos = BuscaMatriz(aAsientoV, rs!cctacnt, 2)
            If nPos = -1 Then
               ReDim Preserve aAsientoV(1 To 5, 0 To UBound(aAsientoV, 2) + 1)
               nPos = UBound(aAsientoV, 2)
            End If
            aAsientoV(1, nPos) = rs!cctacnt
            aAsientoV(2, nPos) = Round(rs!nDebe, 2)
            aAsientoV(3, nPos) = Round(rs!nHaber, 2)
         Case "2"
            nPos = BuscaMatriz(aAsientoV, rs!cctacnt, 2)
            If nPos = -1 Then
               ReDim Preserve aAsientoV(1 To 5, 0 To UBound(aAsientoV, 2) + 1)
               nPos = UBound(aAsientoV, 2)
            End If
            aAsientoV(1, nPos) = rs!cctacnt
            aAsientoV(4, nPos) = Round(rs!nDebe, 2)
            aAsientoV(5, nPos) = Round(rs!nHaber, 2)
         Case "3"
            nPos = BuscaMatriz(aAsientoD, rs!cctacnt, 2)
            If nPos = -1 Then
               ReDim Preserve aAsientoD(1 To 5, 0 To UBound(aAsientoD, 2) + 1)
               nPos = UBound(aAsientoD, 2)
            End If
            aAsientoD(1, nPos) = rs!cctacnt
            aAsientoD(2, nPos) = Round(rs!nDebe, 2)
            aAsientoD(3, nPos) = Round(rs!nHaber, 2)
         End Select
         rs.MoveNext
      Loop
      If Not GrabaAsientoMigra(aAsientoS, True, , 1) Then
         Exit Sub
      End If
      If Not GrabaAsientoMigra(aAsientoD, False, , 2) Then
         Exit Sub
      End If
      GrabaAsientoMigra aAsientoV, False, "Compra - Venta de M.E.", 3
   End If
End Sub

Private Function GrabaAsientoMigra(paAsiento() As String, plMN As Boolean, Optional psMsg As String = "", Optional pnTipo As Integer = 0) As Boolean
  Dim n As Integer
Dim sCtaCod As String
Dim nItem As Integer
Dim gcOpeCod As String
Dim gcMovNro As String
Dim gcGlosa  As String
Dim rs       As ADODB.Recordset
Dim oCon As COMConecta.DCOMConecta
Set oCon = New COMConecta.DCOMConecta
Dim oMov As COMDMov.DCOMMov
Set oMov = New COMDMov.DCOMMov
Dim lnMovNro As Long
On Error GoTo ErrMigra


Set rs = New ADODB.Recordset

oCon.AbreConexion
oCon.CommadTimeOut = 0
GrabaAsientoMigra = False

gcMovNro = Format(txtFecha, "yyyymmdd") & String(6, "0") & gsCodCMAC & gsCodAge & "00XXX" & pnTipo
If Len(gcMovNro) <> 25 Then
   MsgBox "Error en definición de Agencia...!", vbInformation, "Aviso"
   Exit Function
End If
ssql = "SELECT cMovNro FROM Mov WHERE cMovNro like '" & Mid(gcMovNro, 1, 19) & "__" & Right(gcMovNro, 4) & "' and nMovEstado = " & MovEstado.gMovEstContabMovContable & " And nMovFlag <> " & MovFlag.gMovFlagEliminado
If rs.State = adStateOpen Then rs.Close: Set rs = Nothing

Set rs = oCon.CargaRecordSet(ssql)

If Not rs.EOF Then
   MsgBox "Ya se realizó la Migración de Asiento de " & gsNomAge & oImpresora.gPrnSaltoLinea & "Por favor verificar...!", vbInformation, "Advertencia"
   Exit Function
End If
If pnTipo = 3 Then
   gcOpeCod = "701108"
Else
   gcOpeCod = "701107"
End If
gcMovNro = oMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser, gcMovNro)
gcGlosa = "Asiento Contable " & gsNomAge & " en " & IIf(plMN, "M.N.", "M.E.") & " del " & txtFecha & " " & psMsg
If lTransActiva Then
   oCon.RollbackTrans
End If
lTransActiva = True
oCon.BeginTrans      'Iniciamos Transaccion

'Grabamos Mov
'sSql = "INSERT INTO Mov (cMovNro, cOpeCod, cMovDesc, cMovEstado) VALUES ('" & gcMovNro & "','" & gcOpeCod & "','" & gcGlosa & "','0')"
oMov.InsertaMov gcMovNro, gcOpeCod, gcGlosa

lnMovNro = oMov.GetnMovNro(gcMovNro)

nItem = 0
For n = 1 To UBound(paAsiento, 2)
   'Verificamos la existencia de Cuenta Contable
   sCtaCod = VerificaCuenta(paAsiento(1, n), oCon.ConexionActiva)
   'Grabamos MovCta
   If sCtaCod <> "" Then
      If val(paAsiento(2, n)) > 0 Then
         nItem = nItem + 1
         oMov.InsertaMovCta lnMovNro, nItem, sCtaCod, paAsiento(2, n)
         If sObjetoCod <> "" Then  'Grabamos MovObj
            oMov.InsertaMovObj lnMovNro, nItem, 1, Format(ObjEntidadesFinancieras, "00")
            oMov.InsertaMovObjIF lnMovNro, nItem, 1, sObjetoCod, sIFTpo, SCtaIfCod
         End If
         If Not plMN Then
            If val(paAsiento(4, n)) > 0 Then
               oMov.InsertaMovMe lnMovNro, nItem, paAsiento(4, n)
            End If
         End If
      End If
      If val(paAsiento(3, n)) > 0 Then
         nItem = nItem + 1
         oMov.InsertaMovCta lnMovNro, nItem, sCtaCod, val(paAsiento(3, n)) * -1
         If sObjetoCod <> "" Then  'Grabamos MovObj
            oMov.InsertaMovObj lnMovNro, nItem, 1, Format(ObjEntidadesFinancieras, "00")
            oMov.InsertaMovObjIF lnMovNro, nItem, 1, sObjetoCod, sIFTpo, SCtaIfCod
         End If
         If Not plMN Then
            If val(paAsiento(5, n)) > 0 Then
               oMov.InsertaMovMe lnMovNro, nItem, val(paAsiento(5, n)) * -1
            End If
         End If
      End If
   Else
      MsgBox "Cuenta Contable " & paAsiento(1, n) & " del día " & txtFecha & " Tipo " & Format(pnTipo, "00") & " no existe. Por favor verificar", vbInformation, "Error"
      oCon.RollbackTrans
      lTransActiva = False
      Exit Function
   End If
Next
oCon.CommitTrans
lTransActiva = False
GrabaAsientoMigra = True

Exit Function
ErrMigra:
   If lTransActiva Then
      oCon.RollbackTrans
      lTransActiva = False
   End If
   MsgBox Err.Description, vbInformation, "¡Aviso!"
End Function

Private Function VerificaCuenta(sCtaCod As String, oCon As ADODB.Connection) As String
Dim lsCta As String
Dim CtaCor As String
Dim SubCta As String
Dim rs     As New ADODB.Recordset
'Dim oCon As COMConecta.DCOMConecta
'Set oCon = New COMConecta.DCOMConecta

'oCon.AbreConexion

lsCta = sCtaCod
sObjetoCod = ""
'ssql = "SELECT cCtaContCod, MAX(cPersCod+cIFTpo+cCtaIFCod) as cObjetoCod FROM CtaIFFiltro WHERE cCtaContCod + cCtaIFSubCta = '" & sCtaCod & "' GROUP BY cCtaContCod " 'Comments PASI20170630
ssql = "stp_sel_VerificaCuentaxIFI '" & sCtaCod & "'" 'PASI20170630
If rs.State = adStateOpen Then rs.Close: Set rs = Nothing
'Set rs = oCon.CargaRecordSet(ssql)
'ALPA 20080827****************************************************
'rs.Open sSql, oCon, adOpenStatic, adLockReadOnly, adCmdText
Set rs = oCon.Execute(ssql)
'*****************************************************************
If Not rs.EOF Then
   sObjetoCod = Left(rs!cObjetoCod, 13)
   sIFTpo = Mid(rs!cObjetoCod, 14, 2)
   SCtaIfCod = Mid(rs!cObjetoCod, 16, 7)
End If
rs.Close: Set rs = Nothing
VerificaCuenta = lsCta
End Function

Private Function BuscaMatriz(paM() As String, psDato As String, Optional pnDimen As Integer = 1) As Integer
Dim n As Integer
BuscaMatriz = -1
For n = LBound(paM, pnDimen) To UBound(paM, pnDimen)
   If paM(1, n) = psDato Then
      BuscaMatriz = n
      Exit Function
   End If
Next
End Function

Function GetAgenciaCtaDesmAbonoCta(ByVal lnMovNro As Long, ByVal lsOpeCod As String, oCon As ADODB.Connection) As String
'Dim oCon As COMConecta.DCOMConecta
Dim rs As ADODB.Recordset
Dim sql As String

'Set oCon = New COMConecta.DCOMConecta
GetAgenciaCtaDesmAbonoCta = ""
'oCon.AbreConexion
sql = "select SUBSTRING(cCtaCod,4,2) as cCodAge from movcap where nmovnro =" & lnMovNro & " and copecod ='" & lsOpeCod & "'"
'Set rs = oCon.CargaRecordSet(sql)
'ALPA 20080827*****************************************************
Set rs = New ADODB.Recordset
'rs.Open Sql, oCon, adOpenStatic, adLockReadOnly, adCmdText
Set rs = oCon.Execute(sql)
'******************************************************************
If Not rs.EOF And Not rs.BOF Then
    GetAgenciaCtaDesmAbonoCta = rs!cCodAge
End If
rs.Close
Set rs = Nothing
'oCon.CierraConexion
End Function


Private Sub GeneraAjuste(pdFechaAsiento As Date, pnTCCPonderado As Currency, oCon As ADODB.Connection)
    Dim sqlITF As String
    Dim rsITF As ADODB.Recordset
    Set rsITF = New ADODB.Recordset
    Dim lnAjusCV As Currency
    Dim ssql As String
    
    sqlITF = " Select cctacnt, Sum(case ctipo when 0 then nhaber else 0 end) Dolares," _
           & " Convert(Decimal(20,2),Sum(case ctipo when 0 then nhaber else 0 end) * " & pnTCCPonderado & ") DolaresPond," _
           & " Sum(case ctipo when 3 then nhaber else 0 end) DolaresSoles," _
           & " Convert(Decimal(20,2),Sum(case ctipo when 0 then nhaber else 0 end) * " & pnTCCPonderado & ")  -" _
           & " Sum(case ctipo when 3 then nhaber else 0 end) DiffCambio" _
           & " From asientodn where dfecha >= '" & Format(pdFechaAsiento, gsFormatoFecha) & "' and dfecha < '" & Format(pdFechaAsiento + 1, gsFormatoFecha) & "'" _
           & " and cctacnt like '21240206%'  group by cctacnt"
    'ALPA 20080827************************************************
    'rsITF.Open sqlITF, oCon, adOpenStatic, adLockReadOnly, adCmdText
    Set rsITF = oCon.Execute(sqlITF)
    '*************************************************************
    lnAjusCV = 0
    If Not (rsITF.EOF And rsITF.BOF) Then
        While Not rsITF.EOF
            ssql = " INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge,nMovNro) " & _
                  "  VALUES('" & gdHoraGrab & "','" & rsITF!cctacnt & "'," & Abs(rsITF!DolaresSoles) & ",0,'3','" & gsCodAge & "',1)"
            oCon.Execute ssql
            ssql = " INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge,nMovNro) " & _
                  "  VALUES('" & gdHoraGrab & "','" & Left(rsITF!cctacnt, 2) & "1" & Mid(rsITF!cctacnt, 4, 50) & "',0," & Abs(rsITF!DolaresPond) & ",'3','" & gsCodAge & "',2)"
            oCon.Execute ssql
            lnAjusCV = lnAjusCV + rsITF!DiffCambio
            rsITF.MoveNext
        Wend
        If lnAjusCV < 0 Then
            ssql = " INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge) " & _
                   "  VALUES('" & gdHoraGrab & "','5128010201',0," & Abs(lnAjusCV) & ",'3','" & gsCodAge & "')"
            oCon.Execute ssql
        ElseIf lnAjusCV > 0 Then
            ssql = " INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge) " & _
                   "  VALUES('" & gdHoraGrab & "','4128010201'," & Abs(lnAjusCV) & ",0,'3','" & gsCodAge & "')"
            oCon.Execute ssql
        End If
        rsITF.Close
        Set rsITF = Nothing
    End If

End Sub

'*** PEAC 20170303 - En este proceso están operaciones rapidas sin muchas validaciones al
'                    momento de crear las plantillas contables.
Sub AsientoOtrasOpesDeProcesoRapido()
Dim tmpSql As String, vNoCtaCnt As String
Dim vMoneda  As String, vCodConta As String, lsCodAge As String, vAgencia As String
Dim nPaso As Integer
Dim vCtaDolar As Currency

'    Me.Caption = "Generando asientos contables Pruebas."
'    ssql = "exec stp_ins_OperacionesVariasAsientoDN '" & Format(ldFechaAsiento, gsFormatoMovFecha) & "'," & TCV & "," & TCC & "," & TCF
'    oCon.ejecutar ssql


End Sub


'**DAOR 20081204, se pasó parte del codigo de generación de asientos, debidoa que mostraba mensaje : Procedimiento demasiado largo
Sub AsientoCapitlizacionAhorroCTS()
Dim tmpSql As String, vNoCtaCnt As String
Dim vMoneda  As String, vCodConta As String, lsCodAge As String, vAgencia As String
Dim nPaso As Integer
Dim vCta28Debe As Currency, vCta28Haber As Currency
Dim vCta19Debe As Currency, vCta19Haber As Currency, vCtaDolar As Currency
Dim RegOpeCta As ADODB.Recordset
Dim MensajeError As String

    vCta28Debe = 0: vCta28Haber = 0
    vCta19Debe = 0: vCta19Haber = 0: MensajeError = ""
    'vCont = 0: vCtaCnt = 0

    '*** PEAC 20170228 - OPTIMIZADO CAPITALIZACION DE AHORROS
        

        Me.Caption = "Verificando Capitalizaciones de Ahorros Personeria. 123"
        ssql = "exec stp_ins_Capitaliza_Ahorro_Cts_Persone_123 '" & Format(ldFechaAsiento, gsFormatoMovFecha) & "'," & TCV & "," & TCC & "," & TCF & ", " & MensajeError & " output "
        oCon.ejecutar ssql
        
        If Len(Trim(MensajeError)) > 0 Then
            MsgBox "Error stp_ins_Capitaliza_Ahorro_Cts_Persone_123"
        End If
        MensajeError = ""

        Me.Caption = "Verificando Capitalizaciones de Ahorros Personeria. 45678"
        ssql = "exec stp_ins_Capitaliza_Ahorro_Cts_Persone_45678 '" & Format(ldFechaAsiento, gsFormatoMovFecha) & "'," & TCV & "," & TCC & "," & TCF & ", " & MensajeError & " output "
        oCon.ejecutar ssql

        If Len(Trim(MensajeError)) > 0 Then
            MsgBox "Error stp_ins_Capitaliza_Ahorro_Cts_Persone_123"
        End If
        MensajeError = ""

        Me.Caption = "Verificando Capitalizaciones de Ahorros Institucion"
        ssql = "exec stp_ins_Capitaliza_Ahorro_Cts_Persone_Institucion '" & Format(ldFechaAsiento, gsFormatoMovFecha) & "'," & TCV & "," & TCC & "," & TCF & ",'" & gsCodPersCMACT & "', " & MensajeError & " output "
        oCon.ejecutar ssql

        If Len(Trim(MensajeError)) > 0 Then
            MsgBox "Error stp_ins_Capitaliza_Ahorro_Cts_Persone_123"
        End If
        MensajeError = ""

    'Asiento para Capitalización de Ahorros y CTS
'    For nPaso = 1 To 3
'        'If pAsiDia Then
'            If nPaso = 1 Then
'                'Personerias 1,2,3
'                tmpSql = " Select   A.cCtaCod, SUBSTRING(A.cCtaCod,4,2) cAgencia, SUBSTRING(A.cCtaCod,9,1) cMoneda, M.cOpeCod, A.nPersoneria, " _
'                       & "          OC.cCtaContCod, OC.cOpeCtaDH, SUM(Abs(MC.nMonto)) Monto" _
'                       & " FROM OpectaNeg OC" _
'                       & "          INNER JOIN Mov M" _
'                       & "          INNER JOIN MovCap MC ON M.nMovNro = MC.nMovNro" _
'                       & "          INNER JOIN Captaciones A ON MC.cCtaCod = A.cCtaCod ON OC.cOpeCod = M.cOpeCod AND OC.nPersoneria = A.nPersoneria" _
'                       & " Where M.cOpeCod IN (Select cOpeCod From OpeTpoSinAsiento Where nProceso = " & gAsientoProcesoCapAho & ")" _
'                       & "       And (M.nMovFlag = 0) " _
'                       & "       AND A.nPersoneria NOT IN (4,5,6,7,8) " _
'                       & "       And M.cMovNro Like '" & Format(ldFechaAsiento, gsFormatoMovFecha) & "%'" _
'                       & "  Group by A.cCtaCod,SUBSTRING(A.cCtaCod,4,2), substring(A.cCtaCod,9,1), M.cOpeCod, A.nPersoneria, OC.cCtaContCod, OC.cOpeCtaDH" _
'                       & " Order by A.cCtaCod,cAgencia, cMoneda, M.cOpeCod, A.nPersoneria, OC.cOpeCtaDH"
'
'            ElseIf nPaso = 2 Then
'                'Personeria 4,5,7,8
'                tmpSql = " Select   A.cCtaCod,SUBSTRING(A.cCtaCod,4,2) cAgencia, SUBSTRING(A.cCtaCod,9,1) cMoneda, M.cOpeCod, A.nPersoneria, I.cSubCtaContCod, " _
'                       & "          OC.cCtaContCod, OC.cOpeCtaDH, SUM(Abs(MC.nMonto)) Monto" _
'                       & " FROM OpectaNeg OC " _
'                       & "      JOIN Mov M JOIN MovCap MC " _
'                       & "      JOIN Captaciones A JOIN ProductoPersona PP " _
'                       & "      JOIN InstitucionFinanc I ON PP.cPersCod = I.cPersCod ON A.cCtaCod = PP.cCtaCod ON " _
'                       & "      MC.cCtaCod = A.cCtaCod ON M.nMovNro = MC.nMovNro ON OC.cOpeCod = M.cOpeCod AND OC.nPersoneria = A.nPersoneria " _
'                       & " Where M.cOpeCod IN (Select cOpeCod From OpeTpoSinAsiento Where nProceso = " & gAsientoProcesoCapAho & ")" _
'                       & "      And (M.nMovFlag = 0) And PP.nPrdPersRelac = 10 And A.nPersoneria NOT IN (1,2,3) " _
'                       & "      And M.cMovNro Like '" & Format(ldFechaAsiento, gsFormatoMovFecha) & "%'" _
'                       & " Group by A.cCtaCod,SUBSTRING(A.cCtaCod,4,2), substring(A.cCtaCod,9,1), M.cOpeCod, A.nPersoneria, I.cSubCtaContCod, OC.cCtaContCod, OC.cOpeCtaDH" _
'                       & " Order by A.cCtaCod,cAgencia, cMoneda, M.cOpeCod, A.nPersoneria, I.cSubCtaContCod, OC.cOpeCtaDH"
'
'            ElseIf nPaso = 3 Then
'                'CTS
'                tmpSql = "Select    T1.cCtaCod, T1.cAgencia, T1.cTipoCliente, T1.cMoneda, T1.cCtaContCod, T1.cOpeCtaDH, Sum(T1.Monto) Monto, T1.cOpeCod " _
'                    & "From " _
'                    & "(    Select  CaptacCTS.cCtaCod, cTipoCliente = CASE WHEN LTRIM(RTRIM(CaptacCTS.cCodInst)) = '" & gsCodPersCMACT & "' " _
'                    & "             THEN '02' Else '01' END, " _
'                    & "             SUBSTRING(CaptacCTS.cCtaCod,9,1) cMoneda, SUBSTRING(CaptacCTS.cCtaCod,4,2) cAgencia, " _
'                    & "             LTRIM(RTRIM(OC.cCtaContCod)) cCtaContCod, OC.cOpeCtaDH, Abs(MC.nMonto) Monto, T.cOpeCod " _
'                    & "     FROM    Mov T " _
'                    & "             INNER JOIN  MovCap MC " _
'                    & "             INNER JOIN  MovCapDet MCD " _
'                    & "             INNER JOIN  OpeCtaNeg OC ON MCD.cOpeCod = OC.cOpeCod And MCD.nConceptoCod = nConcepto ON MC.nMovNro = MCD.nMovNro And " _
'                    & "                         MC.cOpeCod = MCD.cOpeCod And MC.cCtaCod = MCD.cCtaCod ON T.nMovNro = MC.nMovNro " _
'                    & "             INNER JOIN CaptacCTS ON MC.cCtaCod = CaptacCTS.cCtaCod " _
'                    & "     WHERE   T.cMovnro like '" & Format(ldFechaAsiento, gsFormatoMovFecha) & "%' " _
'                    & "             And T.cOpeCod IN (Select cOpeCod From OpeTpoSinAsiento " _
'                    & "                       Where nProceso = " & gAsientoProcesoCapCTS & ") AND (T.nMovFlag = 0) ) T1 " _
'                    & "GROUP BY    T1.cCtaCod, cAgencia, cTipoCliente, cMoneda, cCtaContCod, cOpeCtaDH,cOpeCod Order by cAgencia, cTipoCliente, cMoneda, cCtaContCod, cOpeCtaDH"
'
'            Else
'                MsgBox " Tipo operación Ahorros no reconocida ", vbInformation, " Aviso "
'                vNoCtaCnt = vNoCtaCnt & oImpresora.gPrnSaltoLinea & "     * Tipo operación Ahorros no reconocida."
'            End If
'
'        Set RegOpeCta = New ADODB.Recordset
'        RegOpeCta.CursorLocation = adUseClient
'        Set RegOpeCta = oCon.CargaRecordSet(tmpSql)
'
'        Set RegOpeCta.ActiveConnection = Nothing
'        If (RegOpeCta.BOF Or RegOpeCta.EOF) Then
'            RegOpeCta.Close
'            Set RegOpeCta = Nothing
'        Else
'            Me.prgList.Max = RegOpeCta.RecordCount
'            With RegOpeCta
'                Do While Not .EOF
'                    If !cOpecod = "101301" Then
'                        vMoneda = vMoneda
'                    End If
'
'                    Me.Caption = "Verificando Capitalizaciones de Ahorros : Registro " & .Bookmark & " de " & .RecordCount
'                    'Variables para cambio en el CodContable
'                    vMoneda = Trim(!cMoneda)
'                    vCodConta = Trim(!cCtaContCod)
'                    vCodConta = Replace(vCodConta, "M", vMoneda, , , vbTextCompare)
'                    '**** LAYG - 07/01/05 ica
'                    lsCodAge = VarAG(!cCtaCod, !cCtaContCod)
'                    vAgencia = lsCodAge
'                    vCodConta = Replace(vCodConta, "AG", lsCodAge, , , vbTextCompare)
'                    '***************
'                    If nPaso = 2 Then
'                        vCodConta = Replace(vCodConta, "CJ", Trim(!cSubCtaContCod), , , vbTextCompare)
'                        vCodConta = Replace(vCodConta, "CR", Trim(!cSubCtaContCod), , , vbTextCompare)
'                        vCodConta = Replace(vCodConta, "ED", Trim(!cSubCtaContCod), , , vbTextCompare)
'                    ElseIf nPaso = 3 Then
'                        vCodConta = Replace(vCodConta, "TC", !cTipoCliente, , , vbTextCompare)
'                    End If
'                    'Se valida la existencia de la Cta. para que se pueda utilizar
''                    If Not ExisCtaCnt(vCodConta, True, oCon) And Len(Trim(vCodConta)) > 0 Then
''                        MsgBox "No existe la Cta. Contable y/o no es Cta de Asiento: " & vCodConta, vbInformation, " Aviso "
''                        vNoCtaCnt = vNoCtaCnt & oImpresora.gPrnSaltoLinea & "     * " & vCodConta & " cta. cnt. no existe. Operacion: " & !cOpeCod & ")."
''                    End If
'
'                    If !cOpeCtaDH = "D" Then
'                        sSQL = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge,cOpeCod, cCtaCod ) " & _
'                            " VALUES('" & gdHoraGrab & "','" & vCodConta & "'," & !Monto & ",0,'0','" & vAgencia & "','" & !cOpecod & "','" & !cCtaCod & "')"
'
'                        oCon.ejecutar sSQL
'                        If Mid(vCodConta, 3, 1) = Moneda.gMonedaExtranjera Then
'                            If Left(vCodConta, 1) = "4" Then
'                                sSQL = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge,cOpeCod, cCtaCod ) " & _
'                                    " VALUES('" & gdHoraGrab & "','" & vCodConta & "'," & (!Monto * TCV) & ",0,'3','" & vAgencia & "','" & !cOpecod & "','" & !cCtaCod & "')"
'                                oCon.ejecutar sSQL
'                                'Ctas 282503 y 191503
'                                vCta28Debe = vCta28Debe + (!Monto * TCF)
'                                vCta19Haber = vCta19Haber + (!Monto * TCV)
'                            ElseIf Left(vCodConta, 1) = "5" Then
'                                sSQL = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge, cOpeCod, cCtaCod ) " & _
'                                    " VALUES('" & gdHoraGrab & "','" & vCodConta & "'," & (!Monto * TCC) & ",0,'3','" & vAgencia & "','" & !cOpecod & "','" & !cCtaCod & "')"
'                                oCon.ejecutar sSQL
'                                'Ctas 282503 y 191503
'                                vCta28Debe = vCta28Debe + (!Monto * TCF)
'                                vCta19Haber = vCta19Haber + (!Monto * TCC)
'                            Else
'                                sSQL = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge,cOpeCod, cCtaCod ) " & _
'                                    " VALUES('" & gdHoraGrab & "','" & vCodConta & "'," & (!Monto * TCF) & ",0,'3','" & vAgencia & "','" & !cOpecod & "','" & !cCtaCod & "')"
'                                oCon.ejecutar sSQL
'                            End If
'                        End If
'                    ElseIf !cOpeCtaDH = "H" Then
'                        sSQL = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge, cOpeCod, cCtaCod) " & _
'                            " VALUES('" & gdHoraGrab & "','" & vCodConta & "',0," & !Monto & ",'0','" & vAgencia & "','" & !cOpecod & "','" & !cCtaCod & "')"
'
'                        oCon.ejecutar sSQL
'                        If Mid(vCodConta, 3, 1) = Moneda.gMonedaExtranjera Then
'                            If Left(vCodConta, 1) = "4" Then
'                                sSQL = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge,cOpeCod, cCtaCod) " & _
'                                    " VALUES('" & gdHoraGrab & "','" & vCodConta & "',0," & (!Monto * TCV) & ",'3','" & vAgencia & "','" & !cOpecod & "','" & !cCtaCod & "')"
'                                oCon.ejecutar sSQL
'                                'Ctas 282503 y 191503
'                                vCta28Haber = vCta28Haber + (!Monto * TCF)
'                                vCta19Debe = vCta19Debe + (!Monto * TCV)
'                            ElseIf Left(vCodConta, 1) = "5" Then
'                                sSQL = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge,cOpeCod, cCtaCod) " & _
'                                    " VALUES('" & gdHoraGrab & "','" & vCodConta & "',0," & (!Monto * TCC) & ",'3','" & vAgencia & "','" & !cOpecod & "','" & !cCtaCod & "')"
'                                oCon.ejecutar sSQL
'                                'Ctas 282503 y 191503
'                                vCta28Haber = vCta28Haber + (!Monto * TCF)
'                                vCta19Debe = vCta19Debe + (!Monto * TCC)
'                            Else
'                                sSQL = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge,cOpeCod, cCtaCod) " & _
'                                    " VALUES('" & gdHoraGrab & "','" & vCodConta & "',0," & (!Monto * TCF) & ",'3','" & vAgencia & "','" & !cOpecod & "','" & !cCtaCod & "')"
'                                oCon.ejecutar sSQL
'                            End If
'                        End If
'                    Else
'                        MsgBox " Operación no reconocida al Insertar " & vbCr & _
'                        " operación nro.: " & vCodConta, vbInformation, " Aviso "
'                        vNoCtaCnt = vNoCtaCnt & oImpresora.gPrnSaltoLinea & "     * " & !cOpeCtaDH & " debe/haber no reconocido (cta.: " & !cOpecod & ")."
'                    End If
'                    prgList.Visible = True
'                    prgList.value = RegOpeCta.Bookmark
'                    .MoveNext
'                    DoEvents
'                Loop
'            End With
'            RegOpeCta.Close
'            Set RegOpeCta = Nothing
'        End If
'    Next

'*** FIN PEAC
    
End Sub
Private Sub cmdValidar_Click()

  
   'On Error GoTo ControlError
    'Dim oAsi As COMDCajaGeneral.DCOMAsiento
    Dim lsPlantilla5114 As String 'ALPA20140303***********
    Dim lbPlantillaLogico As Boolean
    Dim lsNSIF As String 'DAOR 20100623, Proyecto Basilea II
    lsPlantilla5114 = "51M401TCCDNSIFSCAG" 'ALPA20140303***********
    '**DAOR 20081202, para operaciones por ATM***********************
    Dim lsCtaCtbPerdida As String, lsCtaCtbGanancia As String, lnMontoOperacion As Double
    Dim lsCtaCtbPerdidaT As String, lsCtaCtbGananciaT As String
    Dim lnMontoOperacionGP As Double

    lsCtaCtbPerdida = "41M80103AG"
    lsCtaCtbGanancia = "51M80103AG"
    lnMontoOperacion = 0
    '*****************************************************************
        
    If MsgBox("¿ Seguro desea validar cuentas y operaciones para el Asiento Contable ?", vbQuestion + vbYesNo, "¡Confirmacion!") = vbNo Then Exit Sub
    
    Set oAsi = New COMDCajaGeneral.DCOMAsiento
    
    Dim lsCVMEGanacia As String, lsCVMEPerdida As String, lsCajaSoles As String, lsCajaDolares As String
    Dim lsCajaAgenciaSoles As String, lsCajaAgenciaDolares As String
    
    'Desactivo controles
    cmdAsientos.Enabled = False
    cmdImprimir.Enabled = False
    txtFecha.Enabled = False
    fraContenedor(1).Enabled = False
    fraContenedor(2).Enabled = False
    fraImpresion.Enabled = False
    'Fin Desactivo
    
    DoEvents

 '   Call CargaParametrosAsnto
        
    gsCtaCodFoncodes = oAsi.GetAsientoParametro(1)
    lsCVMEGanacia = oAsi.GetAsientoParametro(2)
    lsCVMEPerdida = oAsi.GetAsientoParametro(3)
    lsCajaSoles = oAsi.GetAsientoParametro(4)
    lsCajaDolares = oAsi.GetAsientoParametro(5)
    lsCajaAgenciaSoles = oAsi.GetAsientoParametro(6)
    lsCajaAgenciaDolares = oAsi.GetAsientoParametro(7)
    
'    gsCtaCodFoncodes = oAsi.GetAsientoParametro(1)
'    lsCVMEGanacia = oAsi.GetAsientoParametro(2)
'    lsCVMEPerdida = oAsi.GetAsientoParametro(3)
'    lsCajaSoles = oAsi.GetAsientoParametro(4)
'    lsCajaDolares = oAsi.GetAsientoParametro(5)
'    lsCajaAgenciaSoles = oAsi.GetAsientoParametro(6)
'    lsCajaAgenciaDolares = oAsi.GetAsientoParametro(7)
    
    Dim RegTran As ADODB.Recordset, RegOpeCta As ADODB.Recordset, RegTmp As ADODB.Recordset
    Dim tmpSql As String
    Dim vMoneda As Moneda
'    Dim vAgencia As String, vCodConta As String, vNoCtaCnt As String, vParche As String
    Dim vCont As Long, vCtaCnt As Long
    Dim vEst As String, vCad As String
    Dim x As Long
    Dim sTmp1 As String, sTmp3 As String
    Dim vespacio As Long, vLenNomb As Long
    Dim Arreglo() As Currency
    
    'ALPA 20120625**********************************
    Dim lnTipoCambioCompraL As Currency
    Dim lnTipoCambioVentaL As Currency
    Dim lnTipoCambioFijoL As Currency
    '***********************************************
 
    Call CorreVariablesGlobales
 
'    gAsientoProcesoCapAho = 1
'    gAsientoProcesoCapCTS = 2
'    gAsientoProcesoNoAsiento = 3
'    gAsientoProcesoSiAsiento = 4
'    gAsientoProcesoNoAsientoCab3 = 5
'    gAsientoProcesoNoAsientoCab4 = 6
'    gAsientoProcesoSiAsientoCVME = 7
'    gAsientoProcesoSiAsientoSofFal = 8
'    gAsientoProcesoSiColocCargoCuenta = 9
'    gAsientoProcesoCreditoSinAsiento = 10
    
    Set RegTran = New ADODB.Recordset
    Set RegTmp = New ADODB.Recordset
    Set RegOpeCta = New ADODB.Recordset
    Dim oPrevio As previo.clsprevio
 '   Set oPrevio = New previo.clsprevio
    
    
    If pAsiDia Then
        ldFechaAsiento = gdFecSis
    Else
        ldFechaAsiento = CDate(Me.txtFecha.Text)
    End If
    
    ReDim Arreglo(1, 3)
    
    'Para la validacion Dia y hora de Generación
    'ARCV 14-03-2007
    'dHoraGrab = Format(gdFecSis & " " & Time, gsFormatoFechaHoraView)
    dHoraGrab = Format(Date & " " & Time, gsFormatoFechaHoraView)
    '-------
    'Carga fecha del Asiento y hora de grabación
    gdHoraGrab = Format(txtFecha & " " & Time, gsFormatoFechaHora)
    
    Dim vHH As Long, vMM As Long, vSS As Long, vSeg As Long
    Dim vTime As Variant
    vTime = Time
    
    Dim vCta28Debe As Currency, vCta28Haber As Currency
    Dim vCta19Debe As Currency, vCta19Haber As Currency, vCtaDolar As Currency
    '***********************************************************
    'PARTE 1
    Dim vMonAnt As String, vTipPer As String
    Dim pband28 As Boolean
    '***********************************************************
    'PARTE 2
    Dim vFondo As String, vPlazo As String, vTipoC As String, vRFA As String, vAG As String
    Dim pBandFonc As Boolean
    Dim vCJ As String, vCR As String, vRE As String, vCE As String, vBanco As String
    Dim vTipoPD As String, vSC As String, vSB As String, vCD As String, lsCodAge As String
    Dim vTL As String
    '***********************************************************
    
    Dim vNS As String, vIF As String 'DAOR 20100623, Proyecto Basilea II
    
    MousePointer = 11
    vLenNomb = 70
    vespacio = vLenNomb + 54
    vRTFImp = ""
     
    Dim oDGeneral As COMDConstSistema.NCOMTipoCambio
    Set oDGeneral = New COMDConstSistema.NCOMTipoCambio
    TCCPonderado = oDGeneral.EmiteTipoCambio(CDate(txtFecha.Text), TCPonderado)
     
    If Not GetTipCambio(CDate(txtFecha)) Then
        MsgBox " No se encuentran los Tipos de Cambio ", , " Aviso "
        MousePointer = 0
        Exit Sub
    Else
        With RegTmp
            TCV = gnTipCambioV
            TCC = gnTipCambioC
            TCF = gnTipCambio
            
            If TCF = 0 Then
                MsgBox " No se encuentra el tipo de cambio Fijo Diario", vbInformation, " Aviso "
                RegTmp.Close
                Set RegTmp = Nothing
                MousePointer = 0
                Exit Sub
            End If
        End With
    End If
    
    gAsientoProcesoNoAsiento = 3

    sTmp1 = " exec stp_sel_ObtieneMovParaGenerarAsiento '" & Format(ldFechaAsiento, gsFormatoMovFecha) & "'"
    

    If chkMonSol.value = 0 And chkMonDol.value = 0 Then
        MsgBox " No se ha definido la Moneda ", vbInformation, " Aviso "
        MousePointer = 0
        Exit Sub
    End If
    
    Dim nPaso As Long
    vCta28Debe = 0: vCta28Haber = 0
    vCta19Debe = 0: vCta19Haber = 0
    vCont = 0: vCtaCnt = 0
        
    Call AsientoCapitlizacionAhorroCTSValida
    
    Dim lsRfa As String
    
    ssql = sTmp1
    RegTran.CursorLocation = adUseClient
    Set RegTran = oCon.CargaRecordSet(ssql)
    Set RegTran.ActiveConnection = Nothing
    If (RegTran.BOF Or RegTran.EOF) Then
        RegTran.Close
        Set RegTran = Nothing
    Else
        'prgList.Min = 0
        'prgList.Max = RegTran.RecordCount
        'prgList.Visible = True
        Do While Not RegTran.EOF

            lsRfa = IIf(IsNull(RegTran!cRFA), "", RegTran!cRFA)
     
            lnMontoOperacion = RegTran!sumamonto
            
            'If RegTran!cCtaCod = "109012321000822043" Then MsgBox "109012321000822043"

            'Variables para cambio en el CodContable
            gdHoraGrab = Format$(RegTran!dFecTran, gsFormatoFechaHora)
            ssql = ""
            If Left(RegTran!cOpeCod, 4) = Left(gColPOpeCanceNorEnOtCjEFE, 4) Or _
                Left(RegTran!cOpeCod, 4) = Left(gServGiroCancEfec, 4) Then
                vAgencia = Left(RegTran!cCtaCod, 2)
            Else
                If Len(RegTran!cCodAge) > 2 Then vAgencia = Mid(RegTran!cMovNro, 18, 2) Else vAgencia = Right(Trim(RegTran!cCodAge), 2)
            End If
            
            If Not IsNull(RegTran!cCtaCod) Then
                'Verifica Moneda de Cuentas Antiguas o Cuentas Nuevas
                If Len(Trim(RegTran!cCtaCod)) = 8 Then
                    vMonAnt = Mid(Trim(RegTran!cCtaCod), 2, 1)
                    If vMonAnt = "0" Then
                        vMoneda = Moneda.gMonedaNacional
                    ElseIf vMonAnt = Moneda.gMonedaNacional Then
                        vMoneda = Moneda.gMonedaExtranjera
                    Else
                        MsgBox "Moneda Antigua no reconocida " & RegTran!cCtaCod, , " Aviso "
                        vNoCtaCnt = vNoCtaCnt & vbCrLf & "     * " & RegTran!cCtaCod & " moneda antigua no reconocida."
                    End If
                Else
                    vMoneda = Mid(Trim(RegTran!cCtaCod), 9, 1)
                    If Not (vMoneda = Moneda.gMonedaNacional Or vMoneda = Moneda.gMonedaExtranjera) Then
                        MsgBox " Código Errado, Moneda no definida " & RegTran!cCtaCod, vbInformation, " Aviso "
                        vNoCtaCnt = vNoCtaCnt & vbCrLf & "     * " & RegTran!cCtaCod & " moneda no reconocida."
                    End If
                End If
                
                If (Left(Trim(RegTran!cOpeCod), 1) = "2" Or Trim(RegTran!cOpeCod) = "990101" Or Trim(RegTran!cOpeCod) = "990301" Or Trim(RegTran!cOpeCod) = "990111" Or Trim(RegTran!cOpeCod) = "990113" _
                    Or Trim(RegTran!cOpeCod) = "300111" Or Trim(RegTran!cOpeCod) = "300120" Or Trim(RegTran!cOpeCod) = "300126" Or Trim(RegTran!cOpeCod) = "990305" _
                    Or Trim(RegTran!cOpeCod) = "302001" Or Trim(RegTran!cOpeCod) = "302002" _
                    Or Trim(RegTran!cOpeCod) = "930026" Or Trim(RegTran!cOpeCod) = "930016" Or Trim(RegTran!cOpeCod) = "930028" Or Trim(RegTran!cOpeCod) = "930100" Or Trim(RegTran!cOpeCod) = "990125" _
                    Or Trim(RegTran!cOpeCod) = "930031" Or Trim(RegTran!cOpeCod) = "990126" Or Trim(RegTran!cOpeCod) = "300115") _
                    And ExisTipPer(RegTran!cOpeCod, RegTran!Concepto, True, oCon.ConexionActiva) Then      'RERR: agregue las operaciones de ITF Cargo cuenta
                    vTipPer = ClienteTipoPers(RegTran!cCtaCod, oCon.ConexionActiva)
                    If Len(Trim(vTipPer)) = 0 Then
                        MsgBox " No existe Titular, Operac. Ahorros " & RegTran!cCtaCod, vbInformation, " Aviso "
                        'vNoCtaCnt = vNoCtaCnt & oImpresora.gPrnSaltoLinea & "     * " & RegTran!cCtaCod & " no existe Titular (Operac. Ahorros)."
                        vNoCtaCnt = vNoCtaCnt & vbCrLf & "     * " & RegTran!cCtaCod & " no existe Titular (Operac. Ahorros)."
                    End If
                    ssql = " AND nPersoneria = '" & vTipPer & "'"
                'En Crédito se busca Tipo de Persona en algunas Cuentas de Créditos
                'MARCA 01
                ElseIf Left(Trim(RegTran!cOpeCod), 4) = "0111" Or Left(Trim(RegTran!cOpeCod), 4) = "0118" Or _
                    Left(Trim(RegTran!cOpeCod), 4) = "0104" Or Left(Trim(RegTran!cOpeCod), 4) = "0132" Or _
                    Left(Trim(RegTran!cOpeCod), 4) = "0125" Or Left(Trim(RegTran!cOpeCod), 4) = "0139" Or _
                    Left(Trim(RegTran!cOpeCod), 4) = "0181" Or Left(Trim(RegTran!cOpeCod), 6) = "100202" Or _
                    Left(Trim(RegTran!cOpeCod), 6) = "100302" Or Left(Trim(RegTran!cOpeCod), 6) = "100402" Or _
                    Left(Trim(RegTran!cOpeCod), 6) = "100502" Or Left(Trim(RegTran!cOpeCod), 6) = "100602" Or _
                    Left(Trim(RegTran!cOpeCod), 6) = "100702" Or _
                    Trim(RegTran!cOpeCod) = "100102" Or Trim(RegTran!cOpeCod) = "100103" Or Trim(RegTran!cOpeCod) = "100302" Or _
                    Trim(RegTran!cOpeCod) = "100104" Or Trim(RegTran!cOpeCod) = "100104" Or _
                    Trim(RegTran!cOpeCod) = "100105" Or Trim(RegTran!cOpeCod) = "100106" Or _
                    Trim(RegTran!cOpeCod) = "100107" Or Trim(RegTran!cOpeCod) = "100108" Or _
                    Trim(RegTran!cOpeCod) = "100109" Or Left(Trim(RegTran!cOpeCod), 4) = "0188" Or Trim(RegTran!cOpeCod) = "100220" Then
                        
                        If Left(Trim(RegTran!cOpeCod), 6) = "010440" Or Left(Trim(RegTran!cOpeCod), 4) = "0188" Or _
                            Left(Trim(RegTran!cOpeCod), 6) = "010105" Then
                            vTipPer = ClienteTipoPersCol(RegTran!nMovNro, oCon.ConexionActiva)
                        Else
                            If Left(RegTran!cOpeCod, 5) = "01043" Or Left(RegTran!cOpeCod, 5) = "01113" Or _
                            Left(RegTran!cOpeCod, 5) = "01183" Or Left(RegTran!cOpeCod, 5) = "01253" Or _
                            Left(RegTran!cOpeCod, 5) = "01393" Then
                                vTipPer = 0
                            Else
                                vTipPer = ClienteTipoPersCol(RegTran!nMovNro, oCon.ConexionActiva)
                            End If
                        End If
                        If Len(Trim(vTipPer)) = 0 Then
                            MsgBox " No existe Titular, Operac. Crédito " & RegTran!cCtaCod, vbInformation, " Aviso "
                            vNoCtaCnt = vNoCtaCnt & vbCrLf & "     * " & RegTran!cCtaCod & " no existe Titular (Operac. Créditos)."
                        End If
                        ssql = " AND nPersoneria = " & vTipPer & " "
                ElseIf Left(Trim(RegTran!cOpeCod), 2) = Left(gColRecOpePasoARecup, 2) Then ' Arturo - Judicial
                    If RegTran!nSaldCnt = 1 Or RegTran!nSaldCnt = 2 Or RegTran!nSaldCnt = 3 Or RegTran!nSaldCnt = 4 Then
                        ssql = " AND nPersoneria = '" & RegTran!nSaldCnt & "'"
                    End If
                ElseIf RegTran!cOpeCod = "700104" Or RegTran!cOpeCod = "300121" Or (RegTran!cOpeCod = "302002" And (RegTran!Concepto = 209 Or RegTran!Concepto = 212 Or RegTran!Concepto = 207 Or RegTran!Concepto = 210)) Then
                    vTipPer = ClienteTipoPersCol(RegTran!nMovNro, oCon.ConexionActiva)
                    ssql = IIf(RegTran!cOpeCod = "302002" And (RegTran!Concepto = 207 Or RegTran!Concepto = 212), " AND (nPersoneria = 0 or nPersoneria = " & vTipPer & ") ", " AND nPersoneria = " & vTipPer & " ")
                    
                End If
            Else
                If Not (Left(Trim(RegTran!cOpeCod), 2) = "23" Or Left(Trim(RegTran!cOpeCod), 2) = "24" Or Left(Trim(RegTran!cOpeCod), 2) = "25") Then
                    'MsgBox " No es un Código de Operac. Especial " & RegTran!cOpeCod, vbInformation, " Aviso "
                    vNoCtaCnt = vNoCtaCnt & vbCrLf & "     * " & RegTran!cOpeCod & " no es un código de Operac. Especial (Cta.: " & RegTran!cCtaCod & ")."
                End If
                
                
            End If
            'Para Asiento de Crédito, Foncodes pago en otra agencia
            pband28 = True
                    
            'Carga OPECUENTA para definir a que cuentas se ha de grabar
            'ssql = "SELECT cCtaContCod , cOpeCtaDH FROM OpeCtaNeg  " & _
                " WHERE cOpeCod = '" & RegTran!cOpeCod & "' And nConcepto = " & RegTran!Concepto & " " & ssql   ' & " Order by cDebeHab "
                
             '*** PEAC 20161029
            ssql = "exec stp_sel_CargaOpeCuenta '" & RegTran!cOpeCod & "', " & RegTran!Concepto & "," & IIf(vTipPer = "", 0, vTipPer) & "," & IIf(Len(ssql) = 0, 0, 1)
                
                
            Set RegOpeCta = oCon.CargaRecordSet(ssql)
            If (RegOpeCta.BOF Or RegOpeCta.EOF) Then
                RegOpeCta.Close
                Set RegOpeCta = Nothing
'                If (Left(RegTran!cOpeCod, 3) <> "107" And RegTran!cOpeCod <> "120100") Then
                    'MsgBox " Código de Operación No Reconocido " & RegTran!cOpeCod & "- Concepto " & RegTran!Concepto, vbInformation, " Aviso "
'                End If
                If RegTran!cOpeCod <> "120100" Then 'OBVIA LOS DEMAS CONCEPTOS DEL REGISTRO DE CREDITOS PIGNORATICIO
                    vNoCtaCnt = vNoCtaCnt & vbCrLf & "     * " & RegTran!cOpeCod & " operación no reconocida (Cta.: " & RegTran!cCtaCod & "- Concepto " & RegTran!Concepto & ")."
                End If
            Else
                Do While Not RegOpeCta.EOF
                    lbPlantillaLogico = False
                    'Mensaje de Cuenta
                    vCtaCnt = vCtaCnt + 1
                     If Len(Trim(RegTran!cCodAge)) > 2 Then vAgencia = Mid(RegTran!cMovNro, 18, 2) Else vAgencia = Right(Trim(RegTran!cCodAge), 2)
'                        vAgencia = Mid(RegTran!cMovNro, 18, 2)
'                    Else
'                        vAgencia = Right(Trim(RegTran!cCodAge), 2)
'                    End If
                    
                    vSeg = DateDiff("s", vTime, Time)
                    vSS = vSeg - (Int(vSeg / 60) * 60)
                    vMM = Int(vSeg / 60)
                    vHH = Int(vSeg / 3600)

                    Me.Caption = " Validando CodCta.: " & RegTran!cCtaCod & " - Regis.: " & vCont + 1 & " - CtaCnt.: " & vCtaCnt & " - Time -> " & FillNum(str(vHH), 6, "0") & ":" & FillNum(str(vMM), 6, "0") & ":" & FillNum(str(vSS), 6, "0")
                   
                    If IIf(IsNull(RegTran!cTpoCredCod), 0, 1) = 0 Then
                        vCodConta = GetPlantillaPuente(RegOpeCta!cCtaContCod, Mid(RegTran!cCtaCod, 6, 3), RegTran!Concepto, RegTran!cOpeCod, oCon.ConexionActiva)
                    Else
                        vCodConta = GetPlantillaPuente(RegOpeCta!cCtaContCod, RegTran!cTpoCredCod, RegTran!Concepto, RegTran!cOpeCod, oCon.ConexionActiva)
                    End If
                                        
                    If vCodConta = "" Then
                        vCodConta = Trim(RegOpeCta!cCtaContCod)
                    End If
                    'ALPA20140303***********************************
                    If Trim(lsPlantilla5114) = Trim(vCodConta) And IIf(IsNull(RegTran!cTpoCredCod), "", Trim(RegTran!cTpoCredCod)) = "755" Then
                        lbPlantillaLogico = True
                    End If
                    '***********************************************
                    If Not IsNull(RegTran!cCtaCod) Then
                        'Variable de cambios
                        vFondo = "": vPlazo = "": vTipoC = "": vRFA = ""
                        
                         ''' validacion de creditos RFA (RFC-DIF)
                        '**14M4TC2501RFCD0PSCFFAG
                        If (lsRfa = "RFC" Or lsRfa = "DIF") And Left(vCodConta, 2) = "14" Then
                            'Stop
                            '' CAMBIAMOS PLANTILLAS PARA LOS CREDITOS rfa SI ES CUENTA 14
                            'If RegTran!nCredEstado = 2031 Then
                            If Left(vCodConta, 4) = "14M5" Then
                                vCodConta = Left(vCodConta, 6) + "1902RFCD0PSCFFAG"
                            Else
                                vCodConta = Left(vCodConta, 6) + "2501RF060PSCFFAG"
                            End If
                            'tipo de credito RFA (NEMMONICO RF)
                            Select Case lsRfa
                                Case "RFC"
                                    vRFA = "01"
                                Case "DIF"
                                    vRFA = "02"
                                Case Else
                                    vRFA = ""
                                    vCodConta = "15"
                            End Select
                            vCodConta = Replace(vCodConta, "RF", vRFA, , , vbTextCompare)
                        End If
                        
                        'Arreglo de creditos de consumo vencidos refinanciados - EJRS
                        'If (Left(vCodConta, 4) = "14M5" Or Left(vCodConta, 4) = "14M6") And Left(RegTran!cOpecod, 4) = "1007" And Mid(RegTran!cCtaCod, 6, 1) = "3" Then
                        If (Left(vCodConta, 4) = "14M5" Or Left(vCodConta, 4) = "14M6") And Left(RegTran!cOpeCod, 4) = "1007" And (Mid(RegTran!cTpoCredCod, 1, 1) = "7" Or Mid(RegTran!cTpoCredCod, 1, 1) = "6") Then
                            'ELIMINAMOS EL PAR 29 PUESTO QUE EN CONSUMO NO SE CONSIDERA
                            vCodConta = Replace(vCodConta, "29", "")
                        End If
                         
                        'Tipo de credito
                        If InStr(vCodConta, "TC") > 0 Then
                            If Mid(RegTran!cCtaCod, 6, 3) = Producto.gCapCTS Then
                                vTipoC = ClienteTipoCTS(RegTran!cCtaCod, oCon.ConexionActiva)
                                If vTipoC = "01" Or vTipoC = "02" Then
                                    vCodConta = Replace(vCodConta, "TC", vTipoC, , , vbTextCompare)
                                Else
 '                                   MsgBox " Tipo de Cliente - CTS Errado " & RegTran!cCtaCod, vbInformation, " Aviso "
                                    vNoCtaCnt = vNoCtaCnt & vbCrLf & "     * " & RegTran!cCtaCod & " tipo de cliente - CTS errado."
                                End If
                            Else ' Para TC Tipo de Creditos
                                '**Modificado por DAOR 20100623, Proyecto Basilea II ***************************
                                vTipoC = VarAsientoProdEquiv(RegTran!cTpoCredCod, RegTran!cTpoProdCod, RegTran!nTpoInstCorp, "TC", oCon.ConexionActiva)
                                vCodConta = Replace(vCodConta, "TC", vTipoC, , , vbTextCompare)
                                '*******************************************************************************
                            End If
                        End If
                         
                        If InStr(vCodConta, "SC") > 0 Then
                            'ALPA20140303**********************************************
                            If lbPlantillaLogico = True Then
                                  vSC = "02"
                            Else
                                  vSC = VarAsientoProdEquiv(RegTran!cTpoCredCod, RegTran!cTpoProdCod, RegTran!nTpoInstCorp, "SC", oCon.ConexionActiva)
                            End If
                            If vSC = "" Then
                                vNoCtaCnt = vNoCtaCnt & vbCrLf & "     * " & RegTran!cCtaCod & " tipo de SC errado."
                            Else
                                vCodConta = Replace(vCodConta, "SC", vSC, , , vbTextCompare)
                            End If
                        End If
                        If InStr(vCodConta, "IK") > 0 Then
                            If Me.chkSinCambioSBS.value = 1 Then
                                vCodConta = Replace(vCodConta, "IK", "", , , vbTextCompare)
                            Else
                                '**Modificado por DAOR 20100623, Proyecto Basilea II************************
                                 vCodConta = Replace(vCodConta, "IK", VarAsientoProdEquiv(RegTran!cTpoCredCod, RegTran!cTpoProdCod, RegTran!nTpoInstCorp, "IK", oCon.ConexionActiva), , , vbTextCompare)
                                '****************************************************************************
                            End If
                        End If
                        If InStr(vCodConta, "CD") > 0 Then
                            '**Modificado por DAOR 20100623, Proyecto Basilea II***************************
                            vCD = VarAsientoProdEquiv(RegTran!cTpoCredCod, RegTran!cTpoProdCod, RegTran!nTpoInstCorp, "CD", oCon.ConexionActiva)
                            '******************************************************************************
                            If vCD = "" Then
                            Else
                                vCodConta = Replace(vCodConta, "CD", vCD, , , vbTextCompare)
                            End If
                        End If
                        'ALPA 20111219***************************
                        If InStr(vCodConta, "TL") > 0 Then
                            vTL = VarAsientoProdEquiv(RegTran!cTpoCredCod, RegTran!cTpoProdCod, RegTran!nTpoInstCorp, "TL", oCon.ConexionActiva)
                            If vTL = "" Then
                            Else
                                vCodConta = Replace(vCodConta, "TL", vTL, , , vbTextCompare)
                            End If
                        End If
                        '***************************************
                        '**DAOR 20100623, Proyecto Basilea II ******************************
                        If RegTran!nTpoInstCorp = 5 Then 'Empresas del Sistema Financiero, sólo para créditos
                            lsNSIF = VarAsientoEquivEmpSistFinanc(RegTran!cCtaCod, oCon.ConexionActiva)
                            If InStr(vCodConta, "IF") > 0 Then
                                vIF = IIf(Len(lsNSIF) > 3, Right(lsNSIF, 2), "")
                                vCodConta = IIf(vIF = "", "00", Replace(vCodConta, "IF", vIF, , , vbTextCompare))
                            End If
                            
                            If InStr(vCodConta, "NS") > 0 Then
                                vNS = IIf(Len(lsNSIF) > 1, Left(lsNSIF, 2), "")
                                vCodConta = IIf(vNS = "", "00", Replace(vCodConta, "NS", vNS, , , vbTextCompare))
                            End If
                        Else
                            If InStr(vCodConta, "IF") > 0 Then
                                vCodConta = Replace(vCodConta, "IF", "", , , vbTextCompare)
                            End If
                            
                            If InStr(vCodConta, "NS") > 0 Then
                                vCodConta = Replace(vCodConta, "NS", "", , , vbTextCompare)
                            End If
                        End If
                        '*******************************************************************
                        'Cambios para FF - Fondos
                        If InStr(vCodConta, "FF") > 0 Then
                            vFondo = VarFF(RegTran!cCtaCod, IIf(IsNull(RegTran!LineaC), "", RegTran!LineaC), "", oCon.ConexionActiva)
                            If vFondo = "" Then
                                'MsgBox " Tipo de FF Errado " & RegTran!cCtaCod, vbInformation, " Aviso "
                                vNoCtaCnt = vNoCtaCnt & vbCrLf & "     * " & RegTran!cCtaCod & " tipo de FF errado."
                            End If
                            vCodConta = Replace(vCodConta, "FF", vFondo, 1, 1, vbTextCompare)
                        End If
                        'MARCA 02 ¿Que es fondo?
                        'Valida sólo para Crédito FF:02 - LC:01 y Cta = "28...."
                        If Left(Trim(RegTran!cOpeCod), 2) = "10" And vFondo = "02" _
                            And Left(vCodConta, 2) = "29" And Left(Trim(RegTran!cOpeCod), 3) <> "018" Then
                            vCodConta = ""
                            pband28 = False
                        End If
                        'Valida para Foncodes - Crédito
                        'MARCA 04
                        pBandFonc = False
                        If Left(Trim(RegTran!cOpeCod), 2) = "10" And vFondo = "02" And _
                            (Left(vCodConta, 2) = "14" Or Left(vCodConta, 2) = "51") Then
                            pBandFonc = True
                        End If
                        'Cambios para AO
                        'MARCA 05
                        If InStr(vCodConta, "AO") > 0 Then
                            If InStr(RegTran!cNumDoc, "@") = 1 Then
                                vCodConta = Replace(vCodConta, "AO", Mid(RegTran!cNumDoc, 5, 2), , , vbTextCompare)
                            Else
  '                              MsgBox " No se encontro AO - Créditos " & RegTran!cCtaCod, vbInformation, " Aviso "
                                vNoCtaCnt = vNoCtaCnt & vbCrLf & "     * " & RegTran!cCtaCod & " no se encontro AO - Créditos."
                            End If
                        End If
                    End If
                    'Codificacion de los bancos para el canje de Ordenes de Pago
                    If InStr(vCodConta, "BC") > 0 Then
                        vBanco = VarBC(RegTran!nMovNro, vCodConta, oCon.ConexionActiva)
                        If vBanco = "" Then
   '                         MsgBox " Número de Banco no Identificado (Después de función) " & RegTran!cCtaCod, vbInformation, " Aviso "
                            vNoCtaCnt = vNoCtaCnt & vbCrLf & "     * " & RegTran!cCtaCod & " número de banco no identificado (Después de función) ."
                        Else
                            vCodConta = Replace(vCodConta, "BC", vBanco)
                        End If
                    End If
                    'Valida la CJ - Ahorros - Credito - Pignoraticio
                    'MARCA 06
                    If InStr(vCodConta, "CJ") > 0 Then
                        '************ OJO HAY QUE VERIFICAR ***************************
                        vCJ = VarInstitucionFinanciera(RegTran!cCtaCod, gTpoIFCmac, oCon.ConexionActiva)
                        If vCJ = "" Then
                           vCJ = VarInstitucionFinanciera(RegTran!cCtaCod, gTpoIFCooperativa, oCon.ConexionActiva)
                        End If
                        If vCJ = "" Then
                           vCJ = VarInstitucionFinanciera(RegTran!cCtaCod, gTpoIFEDPYME, oCon.ConexionActiva)
                        End If
                        If vCJ = "" Then
                           vCJ = VarInstitucionFinanciera(RegTran!cCtaCod, gTpoIFCrac, oCon.ConexionActiva)
                        End If
                        If vCJ = "" Then
                            ' parche para operaciones de recepcion CMAC
                            If (Left(RegTran!cOpeCod, 4) >= "2601" And Left(RegTran!cOpeCod, 4) <= "2603") Or _
                                RegTran!cOpeCod = "100205" Or Left(RegTran!cOpeCod, 4) = "9903" Or RegTran!cOpeCod = "100405" Or _
                                RegTran!cOpeCod = "100305" Or RegTran!cOpeCod = "136301" Or RegTran!cOpeCod = "136201" Or _
                                Left(RegTran!cOpeCod, 3) = "126" Or RegTran!cOpeCod = "100505" Then
                                
                                'Los 126 Agregado por JHVP
                                
                                vCJ = VarInstFinanMov(RegTran!nMovNro, gTpoIFCmac, oCon.ConexionActiva)
                                If vCJ = "" Then
                                    MsgBox " CJ no reconocido (Ahorros) " & RegTran!cCtaCod, vbInformation, " Aviso "
                                    vNoCtaCnt = vNoCtaCnt & vbCrLf & "     * " & RegTran!cCtaCod & " CJ no reconocido (Ahorros)."
                                Else
                                    vCodConta = Replace(vCodConta, "CJ", vCJ, 1, 1, vbTextCompare)
                                End If
                            End If
                        Else
                            vCodConta = Replace(vCodConta, "CJ", vCJ, 1, 1, vbTextCompare)
                        End If
                    End If
                    'Valida la CR - Otras Instituciones Financieras - Cajas Rurales (Raul)
                    If InStr(vCodConta, "CR") > 0 Then
                        If RegTran!cOpeCod = "700104" Then
                            vCR = VarInstitucionFinanciera(RegTran!cCtaCodReal, gTpoIFCrac, oCon.ConexionActiva)
                        Else
                            vCR = VarInstitucionFinanciera(RegTran!cCtaCod, gTpoIFCrac, oCon.ConexionActiva)
                        End If
                        
                        If vCR = "" Then
                            MsgBox " CR no reconocido (Ahorros) " & RegTran!cCtaCod, vbInformation, " Aviso "
                            vNoCtaCnt = vNoCtaCnt & vbCrLf & "     * " & RegTran!cCtaCod & " CR no reconocido (Ahorros)."
                        Else
                            vCodConta = Replace(vCodConta, "CR", vCR, 1, 1, vbTextCompare)
                        End If
                    End If
                    If InStr(vCodConta, "CO") > 0 Then
                        vCR = VarInstitucionFinanciera(RegTran!cCtaCod, gTpoIFCooperativa, oCon.ConexionActiva)
                        If vCR = "" Then
                            MsgBox " CO no reconocido (Ahorros) " & RegTran!cCtaCod, vbInformation, " Aviso "
                            vNoCtaCnt = vNoCtaCnt & vbCrLf & "     * " & RegTran!cCtaCod & " CR no reconocido (Ahorros)."
                        Else
                            vCodConta = Replace(vCodConta, "CO", vCR, 1, 1, vbTextCompare)
                        End If
                    End If
                    If InStr(vCodConta, "ED") > 0 Then
                        If RegTran!cOpeCod = "700104" Then
                            vCR = VarInstitucionFinanciera(RegTran!cCtaCodReal, gTpoIFEDPYME, oCon.ConexionActiva)
                        Else
                            vCR = VarInstitucionFinanciera(RegTran!cCtaCod, gTpoIFEDPYME, oCon.ConexionActiva)
                        End If
                        If vCR = "" Then
                            MsgBox " ED no reconocido (Ahorros) " & RegTran!cCtaCod, vbInformation, " Aviso "
                            vNoCtaCnt = vNoCtaCnt & vbCrLf & "     * " & RegTran!cCtaCod & " CR no reconocido (Ahorros)."
                        Else
                            vCodConta = Replace(vCodConta, "ED", vCR, 1, 1, vbTextCompare)
                        End If
                    End If
                    
                    'Créditos
                    If InStr(vCodConta, "PD") > 0 Then
                        vTipoPD = VarPD(RegTran!cCtaCod, IIf(IsNull(RegTran!LineaC), "", RegTran!LineaC), "")
                        If vTipoPD = "01" Or vTipoPD = "02" Or vTipoPD = "03" Or vTipoPD = "05" Then
                            vCodConta = Replace(vCodConta, "PD", vTipoPD, , , vbTextCompare)
                        Else
                            MsgBox " Producto Pendiente Errado " & RegTran!cCtaCod, vbInformation, " Aviso "
                            vNoCtaCnt = vNoCtaCnt & vbCrLf & "     * " & RegTran!cCtaCod & " producto pendiente errado."
                        End If
                    End If
    
                    'Cambios para P - Plazo
                    If InStr(vCodConta, "P") > 0 Then
                        vPlazo = Mid(Trim(RegTran!LineaC), 6, 1)
                        If vPlazo = "1" Or vPlazo = "2" Then
                            vCodConta = Replace(vCodConta, "P", vPlazo, 1, 1, vbTextCompare)
                        Else
                            MsgBox " Plazo de Crédito Errada " & RegTran!cCtaCod, vbInformation, " Aviso "
                            vNoCtaCnt = vNoCtaCnt & vbCrLf & "     * " & RegTran!cCtaCod & " plazo de crédito errado."
                        End If
                    End If
                    
                    'Valida la agencia remota - AG - AL - AR
                    If InStr(vCodConta, "AR") > 0 Then
                        vCodConta = Replace(vCodConta, "AR", vAgencia, 1, 1, vbTextCompare)
                    End If
                    '------ LAYG 07/01/2005 ica
                    If InStr(vCodConta, "AG") > 0 Then
                        If Left(vCodConta, 2) = "11" Then
                            'vCodConta = Replace(vCodConta, "AG", Mid(RegTran!cMovNro, 15, 5), 1, 1, vbTextCompare)
                            vAG = VarAG(Mid(RegTran!cMovNro, 15, 5), vCodConta)
                        Else
                            If Left(RegTran!cOpeCod, 4) = Left(gPigOpeDespContNuevo, 4) Then
                                'vCodConta = Replace(vCodConta, "AG", vAgencia, 1, 1, vbTextCompare)
                                vAG = VarAG(gsCodCMAC & vAgencia, vCodConta)
                            Else
                                'vCodConta = Replace(vCodConta, "AG", Mid(RegTran!cCtaCod, 4, 2), 1, 1, vbTextCompare)
                                If RegTran!cOpeCod = gServGiroCancEfec Then
                                    vAG = RegTran!cCodAgeGir
                                Else
                                    vAG = VarAG(RegTran!cCtaCod, vCodConta, oCon.ConexionActiva)
                                    'vAG = VarAG(RegTran!cCtaCod, vCodConta)
                                End If
                            End If
                        End If
                        
                        'If RegTran!cOpecod = "100911" Or RegTran!cOpecod = "100912" Then
                        If RegTran!TipoDato = 1 And Left(vCodConta, 2) = "14" Then
                            vAG = RegTran!cAgeCodAct
                        End If
                        'CAMBIAMOS AGENCIA SI OPERACION ES DESEMBOLSO CON ABONO A CUENTA DE OTRA AGENCIA
                        'NUEVO PARCHE EJRS  26 DE MAYO 2005
                        If RegTran!cOpeCod = "100104" And Left(vCodConta, 2) = "21" Then
                           vAG = GetAgenciaCtaDesmAbonoCta(RegTran!nMovNro, RegTran!cOpeCod, oCon.ConexionActiva)
                           vAgencia = vAG
                        End If
                        If vAG = "" Then
                            MsgBox " AG no reconocida " & RegTran!cCtaCod, vbInformation, " Aviso "
                            vNoCtaCnt = vNoCtaCnt & vbCrLf & "     * " & RegTran!cCtaCod & " AG no reconocida."
                        Else
                            vCodConta = Replace(vCodConta, "AG", vAG, 1, 1, vbTextCompare)
                        End If
                        
                    End If
                    '----------
                    'Cambios en Moneda y Agencia
                    vCodConta = Replace(vCodConta, "M", vMoneda, 1, 1, vbTextCompare)
                    vParche = AsientoParche(vCodConta, True, oCon.ConexionActiva)
                    If Len(vParche) > 0 Then vCodConta = vParche
                    
                    If RegOpeCta!cCtaContCod = "2312010401" Then
                        MsgBox "verifica cuenta que no existe en plan contable - pacuva"
                    End If
                    
                    'Se valida la existencia de la Cta. para que se pueda utilizar
                    'ALPA 20130527
                    'If (Not ExisCtaCnt(vCodConta, True, oCon.ConexionActiva) And Len(Trim(vCodConta)) > 0) Or Len(Trim(vCodConta)) = 0 Then
                    If RegTran!cOpeCod <> "700104" Then
                        If ((Not ExisCtaCnt(vCodConta, True, oCon.ConexionActiva) And Len(Trim(vCodConta)) > 0) Or Len(Trim(vCodConta)) = 0) Then
                            'MsgBox "No existe la Cta. Contable y/o no es Cta de Asiento: " & vCodConta, vbInformation, " Aviso "
                            'vNoCtaCnt = vNoCtaCnt & vbCrLf & "     * " & vCodConta & " cta. cnt. no existe (cta.: " & RegTran!cCtaCod & ", ope.: " & RegTran!cOpeCod & " - RFA = " & RegTran!cRFA & " - Concepto " & RegTran!Concepto & ")."
                            vNoCtaCnt = vNoCtaCnt & vbCrLf & "     * " & vCodConta & " cta. cnt. no existe (cta.: " & RegTran!cCtaCod & ", ope.: " & RegTran!cOpeCod & " - RFA = " & RegTran!cRFA & " - Concepto " & RegTran!Concepto & ")."
                        End If
                    End If
                    
                    'If Trim(vCodConta) = "19180207" Then Stop
                    
                    '****************************************************************************
                    'Verifica si el asiento ya fue creado
                    'Si existe lo actualiza; caso contrario lo agrega
                    If Mid(RegTran!cCtaCod, 6, 3) = "515" Or Mid(RegTran!cCtaCod, 6, 3) = "516" Then
                            Call ObtenerTipoCambioLeasing(RegTran!cCtaCod, oCon.ConexionActiva, lnTipoCambioCompraL, lnTipoCambioVentaL, lnTipoCambioFijoL)
                    End If
            
                    
                    If Len(Trim(vCodConta)) > 0 Then
                        If Not (RegOpeCta!cOpeCtaDH = "D" Or RegOpeCta!cOpeCtaDH = "H") Then
                            MsgBox " Operación no reconocida al Insertar " & vbCr & _
                            " operación nro.: " & vCodConta, vbInformation, " Aviso "
                            vNoCtaCnt = vNoCtaCnt & vbCrLf & "     * " & RegOpeCta!cOpeCtaDH & " debe/haber no reconocido (cta.: " & RegTran!cCtaCod & ")."
                        End If
                    End If
                    RegOpeCta.MoveNext
                Loop
                RegOpeCta.Close
                Set RegOpeCta = Nothing
            End If
            DoEvents
            vCont = vCont + 1
            'prgList.value = vCont
            RegTran.MoveNext
        Loop
                      
        Dim vCta83D As Currency, vCta83H As Currency, vCta84D As Currency, vCta84H As Currency, vDife8 As Double

        
        'CUADRE DE CTAS DE ORDEN (SILVITA)   83 = 84
        'MARCA 11 CUENTAS DE ORDEN
        vCta83D = 0: vCta83H = 0
        tmpSql = "SELECT abs(round(SUM(ndebe),2)) nDebe, abs(round(SUM(nhaber),2)) nHaber FROM AsientoDN " & _
            " WHERE ctipo = '3' AND datediff(dd, dFecha ,'" & gdHoraGrab & "') = 0 " & _
            " AND cCtaCnt like '83%'" & _
            " Group BY cCtaCnt "
        Set RegTmp = oCon.CargaRecordSet(tmpSql)
        If (RegTmp.BOF Or RegTmp.EOF) Then
            'No se encuentra ctas en Dolares o Tipo 3
        Else
            Do While Not RegTmp.EOF
                vCta83D = vCta83D + RegTmp!nDebe
                vCta83H = vCta83H + RegTmp!nHaber
                RegTmp.MoveNext
            Loop
        End If
        RegTmp.Close
        Set RegTmp = Nothing
        vCta83D = Abs(vCta83D)
        vCta83H = Abs(vCta83H)
        
        vCta84D = 0: vCta84H = 0
        tmpSql = "SELECT abs(round(SUM(ndebe),2)) nDebe,  abs(round(SUM(nhaber),2)) nHaber FROM AsientoDN " & _
            " WHERE ctipo = '3' AND datediff(dd, dFecha ,'" & gdHoraGrab & "') = 0 " & _
            " AND cCtaCnt like '84%'" & _
            " Group BY cCtaCnt "
        Set RegTmp = oCon.CargaRecordSet(tmpSql)
        If (RegTmp.BOF Or RegTmp.EOF) Then
            'No se encuentra ctas en Dolares o Tipo 3
        Else
            Do While Not RegTmp.EOF
                vCta84D = vCta84D + RegTmp!nDebe
                vCta84H = vCta84H + RegTmp!nHaber
                RegTmp.MoveNext
            Loop
        End If
        RegTmp.Close
        Set RegTmp = Nothing
        vCta84D = Abs(vCta84D)
        vCta84H = Abs(vCta84H)
        
        'Validación
        'MsgBox (vCta83D + vCta83H) & "  " & (vCta84D + vCta84H)
        Dim nResult83 As Currency, nResult84 As Currency
        
        nResult83 = IIf(vCta83D > vCta83H, vCta83D - vCta83H, vCta83H - vCta83D)
        nResult84 = IIf(vCta84D > vCta84H, vCta84D - vCta84H, vCta84H - vCta84D)
            
        If nResult83 <> nResult84 Then

        End If
        '**************************************************************************************
        Dim vCta82D As Currency, vCta82H As Currency, vCta81D As Currency, vCta81H As Currency
        'CUADRE DE CTAS DE ORDEN (SILVITA)   82 = 81
        vCta82D = 0: vCta82H = 0
        tmpSql = "SELECT abs(round(SUM(ndebe),2)) nDebe, abs(round(SUM(nhaber),2)) nHaber FROM AsientoDN " & _
            " WHERE ctipo = '3' AND datediff(dd, dFecha ,'" & gdHoraGrab & "') = 0 " & _
            " AND cCtaCnt like '82%'" & _
            " Group BY cCtaCnt "
        Set RegTmp = oCon.CargaRecordSet(tmpSql)
        If (RegTmp.BOF Or RegTmp.EOF) Then
            'No se encuentra ctas en Dolares o Tipo 3
        Else
            Do While Not RegTmp.EOF
                vCta82D = vCta82D + RegTmp!nDebe
                vCta82H = vCta82H + RegTmp!nHaber
                RegTmp.MoveNext
            Loop
        End If
        RegTmp.Close
        Set RegTmp = Nothing
        vCta82D = Abs(vCta82D)
        vCta82H = Abs(vCta82H)
        
        vCta81D = 0: vCta81H = 0
        tmpSql = "SELECT abs(round(SUM(ndebe),2)) nDebe, abs(round(SUM(nhaber),2)) nHaber FROM AsientoDN " & _
            " WHERE ctipo = '3' AND datediff(dd, dFecha ,'" & gdHoraGrab & "') = 0 " & _
            " AND cCtaCnt like '81%'" & _
            " Group BY cCtaCnt "
        Set RegTmp = oCon.CargaRecordSet(tmpSql)
        If (RegTmp.BOF Or RegTmp.EOF) Then
            'No se encuentra ctas en Dolares o Tipo 3
        Else
            Do While Not RegTmp.EOF
                vCta81D = vCta81D + RegTmp!nDebe
                vCta81H = vCta81H + RegTmp!nHaber
                RegTmp.MoveNext
            Loop
        End If
        RegTmp.Close
        Set RegTmp = Nothing
        vCta81D = Abs(vCta81D)
        vCta81H = Abs(vCta81H)
        
        Dim nResult81 As Currency, nResult82 As Currency
        
        nResult81 = IIf(vCta81D > vCta81H, vCta81D - vCta81H, vCta81H - vCta81D)
        nResult82 = IIf(vCta82D > vCta82H, vCta82D - vCta82H, vCta82H - vCta82D)
        
        'Validación

        
        '**************************************************************************************
        Dim vCta86D As Currency, vCta86H As Currency, vCta85D As Currency, vCta85H As Currency
        'CUADRE DE CTAS DE ORDEN (SIMON) 86 = 85
        vCta86D = 0: vCta86H = 0
        tmpSql = "SELECT abs(round(SUM(ndebe),2)) nDebe, abs(round(SUM(nhaber),2)) nHaber FROM AsientoDN " & _
            " WHERE ctipo = '3' AND datediff(dd, dFecha ,'" & gdHoraGrab & "') = 0 " & _
            " AND cCtaCnt like '86%'" & _
            " Group BY cCtaCnt "
        Set RegTmp = oCon.CargaRecordSet(tmpSql)
        If (RegTmp.BOF Or RegTmp.EOF) Then
            'No se encuentra ctas en Dolares o Tipo 3
        Else
            Do While Not RegTmp.EOF
                vCta86D = vCta86D + RegTmp!nDebe
                vCta86H = vCta86H + RegTmp!nHaber
                RegTmp.MoveNext
            Loop
        End If
        RegTmp.Close
        Set RegTmp = Nothing
        vCta86D = Abs(vCta86D)
        vCta86H = Abs(vCta86H)
        
        vCta85D = 0: vCta85H = 0
        tmpSql = "SELECT abs(round(SUM(ndebe),2)) nDebe, abs(round(SUM(nhaber),2)) nHaber FROM AsientoDN " & _
            " WHERE ctipo = '3' AND datediff(dd, dFecha ,'" & gdHoraGrab & "') = 0 " & _
            " AND cCtaCnt like '85%'" & _
            " Group BY cCtaCnt "
        Set RegTmp = oCon.CargaRecordSet(tmpSql)
        If (RegTmp.BOF Or RegTmp.EOF) Then
            'No se encuentra ctas en Dolares o Tipo 3
        Else
            Do While Not RegTmp.EOF
                vCta85D = vCta85D + RegTmp!nDebe
                vCta85H = vCta85H + RegTmp!nHaber
                RegTmp.MoveNext
            Loop
        End If
        RegTmp.Close
        Set RegTmp = Nothing
        vCta85D = Abs(vCta85D)
        vCta85H = Abs(vCta85H)
        
        Dim nResult85 As Currency, nResult86 As Currency
        
        nResult85 = IIf(vCta85D > vCta85H, vCta85D - vCta85H, vCta85H - vCta85D)
        nResult86 = IIf(vCta86D > vCta86H, vCta86D - vCta86H, vCta86H - vCta86D)
        
        
        RegTran.Close
        Set RegTran = Nothing
        'prgList.Visible = False
        'prgList.value = 0
    End If
        
    'Rutina para obtener asientos de la Compra y Venta de Dolares
    Dim CVDebe As Currency, CVHaber As Currency
    tmpSql = " SELECT M.nMovNro, M.cMovNro, cOpeCod, nMovImporte AS Monto, MTC.nMovTpoCambio nTipCambio, " _
           & " SUBSTRING(M.cMovNro,18,2) Agencia, Convert(Decimal(20,2),nMovImporte*nMovTpoCambio) AS NMOVSOLES  FROM Mov M" _
           & " Inner Join MovCompraVenta MCV On M.nMovNro = MCV.nMovNro" _
           & " Inner Join MovTpoCambio MTC On M.nMovNro = MTC.nMovNro" _
           & " WHERE cOpeCod IN (Select cOpeCod From OpeTpoSinAsiento Where nProceso = " & gAsientoProcesoSiAsientoCVME & ")  AND (M.nMovFlag = " & MovFlag.gMovFlagVigente & ")" _
           & " And M.cMovNro Like '" & Format(ldFechaAsiento, gsFormatoMovFecha) & "%' AND MCV.nMovImporte <> 0 ORDER BY cOpeCod"
    
'    tmpSql = "exec stp_sel_ObtieneMovParaOpeCompraVenta " & gAsientoProcesoSiAsientoCVME & "," & MovFlag.gMovFlagVigente & " ,' " & Format(ldFechaAsiento, gsFormatoMovFecha) & " '"  'JUCS20170918
    
    Set RegTmp = oCon.CargaRecordSet(tmpSql)
    If (RegTmp.BOF Or RegTmp.EOF) Then
        RegTmp.Close
        Set RegTmp = Nothing
    Else
        With RegTmp
            Do While Not .EOF
                ssql = " SELECT cCtaContCod , cOpeCtaDH  FROM OpeCtaNeg  " _
                     & " WHERE cOpeCod = '" & !cOpeCod & "'"
                Set RegOpeCta = oCon.CargaRecordSet(ssql)
                If (RegOpeCta.BOF Or RegOpeCta.EOF) Then
                    RegOpeCta.Close
                    Set RegOpeCta = Nothing
                    MsgBox " Código de Operación No Reconocido " & !cOpeCod, vbInformation, " Aviso "
                    vNoCtaCnt = vNoCtaCnt & vbCrLf & "     * " & !cOpeCod & " operación compra/venta no reconocida."
                Else
                    Do While Not RegOpeCta.EOF
                        vCodConta = Trim(RegOpeCta!cCtaContCod)
                        vAG = VarAG(Mid(RegTmp!cMovNro, 15, 5), vCodConta)
                        vCodConta = Replace(vCodConta, "AG", vAG, 1, 1, vbTextCompare)
                        vParche = AsientoParche(vCodConta, True, oCon.ConexionActiva)
                        vAgencia = RegTmp!Agencia
                        If Len(vParche) > 0 Then
                            vCodConta = vParche
                        End If
                      '**************************************************************************
                      'Si existe lo actualiza; caso contrario lo agrega
                      'MARCA 12
                        'If Trim(vCodConta) = "19180207" Then Stop
                        If RegOpeCta!cOpeCtaDH = "D" Then
                            If Left(vCodConta, 4) = lsCajaSoles Or Left(vCodConta, 4) = "2514" Then 'APRI20170705 RFC1705220001 Add Left(vCodConta, 4) = "2514"
                            
                            ElseIf Left(vCodConta, 4) = lsCajaDolares Or Left(vCodConta, 4) = "2524" Then
                            
                            Else
                                vNoCtaCnt = vNoCtaCnt & vbCrLf & "     * " & vCodConta & " cta. de compra/venta no reconocida."
                            End If
                        ElseIf RegOpeCta!cOpeCtaDH = "H" Then
                            If Left(vCodConta, 4) = lsCajaSoles Or Left(vCodConta, 4) = "2514" Then
                            
                            ElseIf Left(vCodConta, 4) = lsCajaDolares Or Left(vCodConta, 4) = "2524" Then 'Or Left(vCodConta, 4) = "2825" 'APRI20170705 RFC1705220001 Add Left(vCodConta, 4) = "2524"
                            
                            Else
                                MsgBox "Cuenta en Compra y Venta no reconocida", vbInformation, "Aviso"
                                vNoCtaCnt = vNoCtaCnt & vbCrLf & "     * " & vCodConta & " cta. de compra/venta no reconocida."
                            End If
                        Else
                            MsgBox " Operación no reconocida al Insertar " & vbCr & _
                            " operación nro.: " & vCodConta, vbInformation, " Aviso "
                            vNoCtaCnt = vNoCtaCnt & vbCrLf & "     * " & RegOpeCta!cOpeCtaDH & " debe/haber no reconocido (compra/venta)."
                        End If
                      RegOpeCta.MoveNext
                    Loop
                    RegOpeCta.Close
                    Set RegOpeCta = Nothing
                End If
                .MoveNext
            Loop
        End With
        RegTmp.Close
        Set RegTmp = Nothing
        'Validación de 0.01 por dolares - Silvita
        Dim nCtaDeb1 As Currency, nCtaHab1 As Currency
        
        tmpSql = "SELECT cCodAge, sum(round(ndebe,2)) nDebe , sum(round(nhaber,2)) nHaber " & _
            " FROM AsientodN " & _
            " WHERE cTipo = '1' AND datediff(dd, dFecha ,'" & gdHoraGrab & "') = 0 " & _
            " GROUP BY cCodAge, cCtaCnt Order by 1 "
            
        'tmpSql = "SELECT cCodAge, sum(round(ndebe,2)) nDebe , sum(round(nhaber,2)) nHaber " & _
            " FROM AsientodN " & _
            " WHERE cTipo = '1' AND cOpecod NOT IN ('900022','900023') AND datediff(dd, dFecha ,'" & gdHoraGrab & "') = 0 " & _
            " GROUP BY cCodAge, cCtaCnt Order by 1 " 'JUCS TI-ERS-002-2017
            
        Set RegTmp = oCon.CargaRecordSet(tmpSql)
        If (RegTmp.BOF Or RegTmp.EOF) Then
            'No se encuentra ctas en Dolares o Tipo 1
        Else
            vAgencia = RegTmp!cCodAge
            Do While Not RegTmp.EOF
                Do While RegTmp!cCodAge = vAgencia
                    nCtaDeb1 = nCtaDeb1 + RegTmp!nDebe
                    nCtaHab1 = nCtaHab1 + RegTmp!nHaber
                    RegTmp.MoveNext
                    If RegTmp.EOF Then Exit Do
                Loop
                vCtaDolar = nCtaDeb1 - nCtaHab1

                If RegTmp.EOF Then Exit Do
                vAgencia = RegTmp!cCodAge
                nCtaDeb1 = 0: nCtaHab1 = 0
            Loop
        End If
        RegTmp.Close
        Set RegTmp = Nothing
    End If
    
    'Asiento para sobrantes y faltantes
    Dim vCodConta1 As String, vCodConta2 As String
    'Add by gitu 16/09/2008
    Dim lsMonDebe As String, lsMonHaber As String, lsMonDebeD As String, lsMonHaberD As String
    'End Gitu

    

    
    '-- Aca
       'Verifica que cuadre DEBE y HABER - Tipo '3'
        Dim vDife As Currency, vDifeSol As Currency, vMonDebe As Currency, vMonHaber As Currency
        Dim vMonDebeSol As Currency, vMonHaberSol As Currency
        'Carga suma de tipo soles para hallar la diferencia - RAUL
        
        tmpSql = "SELECT round(SUM(ndebe),2) AS MonDebe , round(SUM(nhaber),2) AS MonHaber FROM AsientoDN " & _
            " WHERE ctipo = '0' AND substring(cCtaCnt,3,1) = '" & Moneda.gMonedaNacional & "' AND datediff(dd, dFecha ,'" & gdHoraGrab & "') = 0 " & _
            " Group BY cCtaCnt "
            
        'tmpSql = "SELECT round(SUM(ndebe),2) AS MonDebe , round(SUM(nhaber),2) AS MonHaber FROM AsientoDN " & _
            " WHERE ctipo = '0' AND substring(cCtaCnt,3,1) = '" & Moneda.gMonedaNacional & "' AND cOpeCod NOT IN ('900022','900023') AND datediff(dd, dFecha ,'" & gdHoraGrab & "') = 0 " & _
            " Group BY cCtaCnt " 'JUCS TI-ERS-002-2017
            
        vMonDebeSol = 0: vMonHaberSol = 0: vMonDebe = 0: vMonHaber = 0
        Set RegTmp = oCon.CargaRecordSet(tmpSql)
        If (RegTmp.BOF Or RegTmp.EOF) Then
            'No se encuentra ctas en Dolares o Tipo 3
        Else
            Do While Not RegTmp.EOF
                vMonDebeSol = vMonDebeSol + RegTmp!mondebe
                vMonHaberSol = vMonHaberSol + RegTmp!monhaber
                RegTmp.MoveNext
            Loop
        End If
        RegTmp.Close
        Set RegTmp = Nothing

        'Halla suma de diferencia de Dolares
        vDifeSol = vMonDebeSol - vMonHaberSol


'        RegTmp.Close
'        Set RegTmp = Nothing
    
    'dbCmact.CommitTrans  'Finaliza TRANSACCION
    '*************************************************************************************
    '*************************************************************************************
    'CONVERSION DE ITF DE DOLARES A SOLES
    
    'GeneraAjuste ldFechaAsiento, TCCPonderado, oCon.ConexionActiva
    
    '*************************************************************************************
    '*************************************************************************************

    
    
    '*********** PEAC 20110705 - VERIFICA EXISTENCIA DE CTAS CONT EN EL PLAN CONT.
    Dim rsCtas As ADODB.Recordset
    Set rsCtas = New ADODB.Recordset
    Dim vIndice As Long, vLineas As Long, vPage As Long
    Dim vDolar As Boolean
    Dim vSumDebe As Currency, vSumHaber As Currency
    Dim vMoneAnte As String, vProduc As String, sSqlDolar As String, vCabecera As String
    vDolar = False
    Dim vDiaSinOpe  As String
    'VALIDACION DE ASIENTOS
    'Dim vObs As String
    ssql = "exec stp_sel_BuscaCtaContPlan '" & Format(txtFecha, "yyyymmdd") & "'"
    Set rsCtas = oCon.CargaRecordSet(ssql)
    If Not (rsCtas.BOF Or rsCtas.EOF) Then
        Do While Not rsCtas.EOF
        'ALPA 20130527***********************************
        'vObs = vObs & vbCrLf & "** OJO ** Cta.Inexist.Plan Cont>> Cta.Cont:'" & rsCtas!cCtaCnt & "' Importe: '" & Trim(CStr(rsCtas!Importe)) & "' Age:'" & Trim(rsCtas!cCodAge) & "' Ope:'" & Trim(rsCtas!cOpecod) & "' Cod.Cta.'" & IIf(IsNull(rsCtas!cCtaCod), "", rsCtas!cCtaCod) & "'"
        If rsCtas!cOpeCod <> "700104" Then
            vObs = vObs & vbCrLf & "** OJO ** Cta.Inexist.Plan Cont>> Cta.Cont:'" & rsCtas!cctacnt & "' Importe: '" & Trim(CStr(rsCtas!Importe)) & "' Age:'" & Trim(rsCtas!cCodAge) & "' Ope:'" & Trim(rsCtas!cOpeCod) & "' Cod.Cta.'" & IIf(IsNull(rsCtas!cCtaCod), "", rsCtas!cCtaCod) & "'"
        End If
        '************************************************
        rsCtas.MoveNext
        Loop
    End If
    rsCtas.Close
    
    '*********** FIN PEAC
    
    
    'Verifica la veracidad del asiento contable
    Me.Caption = "Verificando Veracidad de asiento Contable"
    If Len(Trim(vObs)) > 0 Or Len(Trim(vNoCtaCnt)) > 0 Then
        MsgBox "La Validación del asiento presenta Observaciones ! ", vbInformation, " Aviso "
        vProduc = vbCrLf & Space(40) & "  A S I E N T O     ¡   N O     V A L I D O   !" & vbCrLf
        vProduc = vProduc & "   OBSERVACIONES DEL ASIENTO : " & vbCrLf & _
                            "  ============================="
        If Len(Trim(vNoCtaCnt)) > 0 Then vProduc = vProduc & vNoCtaCnt & vbCrLf
        If Len(Trim(vObs)) > 0 Then vProduc = vProduc & vObs & vbCrLf
        vProduc = vProduc & vbCrLf & vbCrLf
        vLineas = Repetido(vProduc, vbCrLf)
    Else
        '*************************************************************************************
        If pAsiDia Then
            'Actualización de variable de control
            ssql = "UPDATE ConstSistema SET nConsSisValor = '1' WHERE nConsSisCod = '2'"
            oCon.ejecutar ssql
        End If
        ssql = "UPDATE AsientoValida SET cAsientoEstado = '1' " & _
             " WHERE dAsientoFecha = '" & Format(txtFecha, gsFormatoFecha) & "' AND cAsientoTipo = '2' AND " & _
             " cCodUsu = '" & gsCodUser & "' AND dAsientoModif = '" & Format(dHoraGrab, gsFormatoFechaHora) & "'"
        oCon.ejecutar ssql
        '*************************************************************************************
    End If
    'Determina Productos
    vProduc = vProduc & "   DE : " & vbCrLf & "  ======" & vbCrLf
    If chkAho.value = 1 Then vProduc = vProduc & "     *  AHORROS " & vbCrLf
    If chkCre.value = 1 Then vProduc = vProduc & "     *  CREDITOS " & vbCrLf
    If chkPig.value = 1 Then vProduc = vProduc & "     *  CREDITO PIGNORATICIO " & vbCrLf
    ssql = "SELECT  cctacnt, round(sum(ndebe),2) nDebe, round(sum(nhaber),2) nHaber, substring(cctacnt,3,1) AS MONEDA " & _
        " FROM AsientoDN "
    
    'Verifica que condiciones se ha de tener en cuenta:  Soles - Dolares
    vEst = "": vCad = ""
    If chkMonSol.value = 1 Then vEst = vEst & Moneda.gMonedaNacional
    If chkMonDol.value = 1 Then vEst = vEst & Moneda.gMonedaExtranjera
    For x = 1 To Len(vEst) Step 1
        vCad = vCad & "'" & Mid(vEst, x, 1) & "'"
        If x <> Len(vEst) Then vCad = vCad & ","
    Next x
    
    If Len(vEst) > 0 Then
        '        sSqlDolar = sSql & " WHERE cTipo IN ('3') AND datediff(dd,dFecha,'" & gdHoraGrab & "') = 0  GROUP BY CTIPO, CCTACNT ORDER BY cctacnt "
        sSqlDolar = ssql & " WHERE cTipo IN ('3') AND datediff(dd,dFecha,'" & gdHoraGrab & "') = 0  And substring(cctacnt,3,1) in ('1','2')" & _
            " GROUP BY CTIPO, CCTACNT ORDER BY cctacnt "
        ssql = ssql & " WHERE cTipo IN ('0') AND substring(cctacnt,3,1) IN (" & vCad & ") AND datediff(dd,dFecha,'" & gdHoraGrab & "') = 0  "
        If chkMonSol.value = 1 And chkMonDol.value = 1 Then
            ssql = ssql & " GROUP BY CTIPO, CCTACNT ORDER BY moneda , cctacnt "
        Else
            ssql = ssql & " GROUP BY CTIPO, CCTACNT ORDER BY cctacnt "
        End If
    End If
    
    vMoneAnte = ""
    vLineas = vLineas + 12: vPage = 1
    vRTFImp = gsNomCmac & Space(vLenNomb + 6) & Format(gdFecSis & " " & Time, "dd/mm/yyyy hh:mm:ss") & vbCrLf _
     & Space(vLenNomb + 16) & " Página :" & ImpreFormat(vPage, 5, 0) & vbCrLf _
     & ImpreFormat(UCase(gsNomAge), 25) & vbCrLf _
     & ImpreFormat("   VALIDACION DEL ASIENTO CONTABLE DEL DIA " & Format(txtFecha, "dd/mm/yyyy"), 44, 43) & vbCrLf _
     & ImpreFormat(String(40, "="), 44, 42) & vbCrLf
    If CDate(txtFecha) = gdFecSis Then
        If Not CierreRealizado2(1, oCon.ConexionActiva) Then
            vRTFImp = vRTFImp & ImpreFormat(" NO VALIDO (ANTES DEL CIERRE DE DIA)", 40, 45) & vbCrLf
        Else
            'Para los dias que no hay movimientos
            vDiaSinOpe = "Select cMovNro From Mov M Where M.cMovNro Like '" & Format(gdHoraGrab, gsFormatoMovFecha) & "%' AND M.nMovFlag = " & MovFlag.gMovFlagVigente
            Set RegTmp = oCon.CargaRecordSet(vDiaSinOpe)
            If (RegTmp.BOF Or RegTmp.EOF) Then
                vRTFImp = vRTFImp & vbCr & Space(28) & "N O   S E   R E A L I Z A R O N   O P E R A C I O N E S   E N   E L   D I A"
            Else
                If VerificaDiaHabil(gdHoraGrab, 3) And Not CierreRealizado2(2, oCon.ConexionActiva) Then
                    vRTFImp = vRTFImp & ImpreFormat(" NO VALIDO (ANTES DEL CIERRE DE MES)", 40, 45) & vbCrLf
                End If
            End If
            RegTmp.Close
            Set RegTmp = Nothing
        End If
    End If
    vRTFImp = vRTFImp & vProduc
    oCon.AbreConexion
    oCon.CommadTimeOut = 0
    Set RegTmp = oCon.CargaRecordSet(ssql)
    If (RegTmp.BOF Or RegTmp.EOF) Then
        RegTmp.Close
        Set RegTmp = Nothing
    Else
        With RegTmp
        Do While Not .EOF
            If !Moneda <> vMoneAnte Then
                If !Moneda = Moneda.gMonedaNacional Then
                    vRTFImp = vRTFImp & vbCrLf & "   MONEDA : SOLES" & vbCrLf
                    vLineas = vLineas + 1
                Else
                    If chkMonSol.value = 1 And (vSumDebe > 0 Or vSumHaber > 0) Then
                        vRTFImp = vRTFImp & String(vespacio, "-") & vbCrLf
                        vRTFImp = vRTFImp & Space(vLenNomb + 12) & ImpreFormat("TOTAL", 11, 0) & _
                            ImpreFormat(vSumDebe, 12, , True) & ImpreFormat(vSumHaber, 12, , True) & vbCrLf
                        vRTFImp = vRTFImp & String(vespacio, "-") & vbCrLf
                        vSumDebe = 0: vSumHaber = 0
                        vLineas = vLineas + 3
                    End If
                    vDolar = True
                    If chkMonSol.value = 1 Then
                        vPage = vPage + 1
                        vRTFImp = vRTFImp & oImpresora.gPrnSaltoPagina & vbCrLf
                        vRTFImp = vRTFImp & Space(vLenNomb + 35) & "Página :" & ImpreFormat(vPage, 5, 0) & vbCrLf
                    End If
                    vRTFImp = vRTFImp & vbCrLf & "   MONEDA : DOLARES" & vbCrLf
                    vLineas = 5 'vLineas + 2
                End If
                vRTFImp = vRTFImp & String(vespacio, "-") & vbCrLf
                vRTFImp = vRTFImp & Space(1) & vCabecera
                vRTFImp = vRTFImp & String(vespacio, "-") & vbCrLf
                vLineas = vLineas + 3
            End If
            vMoneAnte = !Moneda
            vRTFImp = vRTFImp & ImpreFormat(!cctacnt, 21, 1) & ImpreFormat(UCase(CuentaNombre(!cctacnt, True, oCon.ConexionActiva)), vLenNomb, 1) & ImpreFormat(Round(!nDebe, 2), 12, , True) & ImpreFormat(Round(!nHaber, 2), 12, , True) & vbCrLf
            vSumDebe = vSumDebe + Round(!nDebe, 2)
            vSumHaber = vSumHaber + Round(!nHaber, 2)
            vLineas = vLineas + 1
            If vLineas > pLineasMax Then
                vPage = vPage + 1
                vRTFImp = vRTFImp & oImpresora.gPrnSaltoPagina & vbCrLf _
                 & Space(vLenNomb + 35) & "Página :" & ImpreFormat(vPage, 5, 0) & vbCrLf & vbCrLf _
                 & String(vespacio, "-") & vbCrLf _
                 & Space(1) & vCabecera _
                 & String(vespacio, "-") & vbCrLf
                vLineas = 5
            End If
            .MoveNext
        Loop
        End With
        vRTFImp = vRTFImp & String(vespacio, "-") & vbCrLf
        vRTFImp = vRTFImp & Space(vLenNomb + 12) & ImpreFormat("TOTAL", 11, 0) & _
            ImpreFormat(vSumDebe, 12, , True) & ImpreFormat(vSumHaber, 12, , True) & vbCrLf
        vRTFImp = vRTFImp & String(vespacio, "-") & vbCrLf
        vLineas = vLineas + 3
        RegTmp.Close
        Set RegTmp = Nothing
        
        If vDolar Then
            Set RegTmp = oCon.CargaRecordSet(sSqlDolar)
            If (RegTmp.BOF Or RegTmp.EOF) Then
                RegTmp.Close
                Set RegTmp = Nothing
            Else
                If vLineas + 8 >= pLineasMax Then
                    vPage = vPage + 1
                    vRTFImp = vRTFImp & oImpresora.gPrnSaltoPagina & vbCrLf
                    vRTFImp = vRTFImp & Space(vLenNomb + 35) & "Página :" & ImpreFormat(vPage, 5, 0) & vbCrLf & vbCrLf
                    vLineas = 5
                End If
                vRTFImp = vRTFImp & vbCrLf & "   CONVERSION DE DOLARES" & vbCrLf _
                 & "     *  TIPO CAMBIO FIJO   : " & ImpreFormat(TCF, 6, 3) & vbCrLf _
                 & "     *  TIPO CAMBIO VENTA  : " & ImpreFormat(TCV, 6, 3) & vbCrLf _
                 & "     *  TIPO CAMBIO COMPRA : " & ImpreFormat(TCC, 6, 3) & vbCrLf _
                 & String(vespacio, "-") & vbCrLf _
                 & Space(1) & vCabecera _
                 & String(vespacio, "-") & vbCrLf
                vLineas = vLineas + 8
                vSumDebe = 0: vSumHaber = 0
                With RegTmp
                Do While Not .EOF
                    vRTFImp = vRTFImp & ImpreFormat(!cctacnt, 21, 1) & ImpreFormat(UCase(CuentaNombre(!cctacnt, True, oCon.ConexionActiva)), vLenNomb, 1) & ImpreFormat(Round(!nDebe, 2), 12, , True) & ImpreFormat(Round(!nHaber, 2), 12, , True) & vbCrLf
                    vSumDebe = vSumDebe + !nDebe
                    vSumHaber = vSumHaber + !nHaber
                    vLineas = vLineas + 1
                    If vLineas > pLineasMax Then
                        vPage = vPage + 1
                        vRTFImp = vRTFImp & oImpresora.gPrnSaltoPagina & vbCrLf _
                         & Space(vLenNomb + 35) & "Página :" & ImpreFormat(vPage, 5, 0) & vbCrLf & vbCrLf _
                         & String(vespacio, "-") & vbCrLf _
                         & Space(1) & vCabecera _
                         & String(vespacio, "-") & vbCrLf
                        vLineas = 5
                    End If
                    .MoveNext
                Loop
                End With
                RegTmp.Close
                Set RegTmp = Nothing
                If vLineas + 3 > pLineasMax Then
                    vPage = vPage + 1
                    vRTFImp = vRTFImp & oImpresora.gPrnSaltoPagina & vbCrLf
                    vRTFImp = vRTFImp & Space(vLenNomb + 35) & "Página :" & ImpreFormat(vPage, 5, 0) & vbCrLf & vbCrLf
                    vLineas = 5
                End If
                vRTFImp = vRTFImp & String(vespacio, "-") & vbCrLf
                vRTFImp = vRTFImp & Space(vLenNomb + 12) & ImpreFormat("TOTAL", 11, 0) & _
                    ImpreFormat((vSumDebe), 12, , True) & ImpreFormat((vSumHaber), 12, , True) & vbCrLf
                vRTFImp = vRTFImp & String(vespacio, "-") & vbCrLf
                vLineas = vLineas + 3
            End If
        End If
    End If
    
    '**************************************************************************************
    'CARGA DE : Compra y Venta en Dolares y en Soles
    Me.Caption = "Validando carga de asientos de compra - venta dolares y soles"
    
    Dim ComVta(6, 1) As Currency ' 0-Ctas; 1-Debe;  2-Haber
    Erase ComVta
    Dim ComVtaDolar As String
    ComVtaDolar = ""
    
    ssql = "SELECT  cctacnt, round(sum(ndebe),2) nDebe, round(sum(nhaber),2) nHaber, ctipo " & _
        " FROM AsientoDN WHERE ctipo IN ('1','2') AND datediff(dd, dFecha ,'" & gdHoraGrab & "') = 0 " & _
        " GROUP BY CTIPO, CCTACNT "
    Set RegTmp = oCon.CargaRecordSet(ssql)
    If (RegTmp.BOF Or RegTmp.EOF) Then
    Else
        If vLineas + 19 >= pLineasMax Then
            vPage = vPage + 1
            vRTFImp = vRTFImp & oImpresora.gPrnSaltoPagina & vbCrLf
            vRTFImp = vRTFImp & Space(vLenNomb + 35) & "Página :" & ImpreFormat(vPage, 5, 0) & vbCrLf & vbCrLf
            vLineas = 5
        End If
    
        With RegTmp
            Do While Not .EOF
                If !cTipo = "1" Then
                    If Left(!cctacnt, 4) = lsCajaDolares Or Left(vCodConta, 4) = "2524" Then 'APRI20170705 RFC1705220001 Add Left(vCodConta, 4) = "2524"
                        ComVta(1, 0) = !nDebe
                        ComVta(1, 1) = !nHaber
                    ElseIf Left(!cctacnt, 4) = lsCajaSoles Or Left(vCodConta, 4) = "2514" Then 'APRI20170705 RFC1705220001 Add Left(vCodConta, 4) = "2514"
                        ComVta(3, 0) = !nDebe
                        ComVta(3, 1) = !nHaber
                    ElseIf Left(!cctacnt, 4) = "4128" Then
                        ComVta(4, 0) = !nDebe
                    ElseIf Left(!cctacnt, 4) = "5128" Then
                        ComVta(4, 1) = !nHaber
                    Else
                        MsgBox "Cuenta no reconocida para Compra y Venta de dolares", vbInformation, " Aviso "
                    End If
                Else
                    If Left(!cctacnt, 4) = lsCajaDolares Or Left(vCodConta, 4) = "2524" Then 'APRI20170705 RFC1705220001 Add Left(vCodConta, 4) = "2524"
                        ComVta(5, 0) = !nDebe
                        ComVta(5, 1) = !nHaber
                    Else
                        MsgBox "Cuenta no reconocida para Compra y Venta de dolares", vbInformation, " Aviso "
                    End If
                End If
                .MoveNext
            Loop
        End With
        
        ComVtaDolar = "" & vbCrLf
        ComVtaDolar = ComVtaDolar & "   DE : COMPRA Y VENTA EN DOLARES " & vbCrLf
        ComVtaDolar = ComVtaDolar & String(vespacio, "-") & vbCrLf
        Dim sCajaME As String, sCajaMN As String
        Dim nDebeAcum As Double, nHaberAcum As Double
        
        nDebeAcum = 0
        nHaberAcum = 0
        RegTmp.MoveFirst
        
        Do While Not RegTmp.EOF
            If RegTmp("cTipo") = "2" Then
                sCajaME = RegTmp("cCtaCnt")
                nDebeAcum = nDebeAcum + RegTmp("nDebe")
                nHaberAcum = nHaberAcum + RegTmp("nHaber")
                ComVtaDolar = ComVtaDolar & ImpreFormat(sCajaME, 21, 1) & ImpreFormat(UCase(CuentaNombre(sCajaME, True, oCon.ConexionActiva)), vLenNomb, 1) & ImpreFormat(Round(RegTmp("nDebe"), 2), 12) & ImpreFormat(Round(RegTmp("nHaber"), 2), 12) & vbCrLf
            End If
            RegTmp.MoveNext
        Loop
        
        ComVtaDolar = ComVtaDolar & String(vespacio, "-") & vbCrLf
        ComVtaDolar = ComVtaDolar & Space(vLenNomb + 12) & ImpreFormat("TOTAL", 11, 0) & _
                    ImpreFormat(Round(nDebeAcum, 2), 12, , True) & _
                    ImpreFormat(Round(nHaberAcum, 2), 12, , True) & vbCrLf
        ComVtaDolar = ComVtaDolar & String(vespacio, "-") & vbCrLf
        
        ComVtaDolar = ComVtaDolar & vbCrLf & vbCrLf
        ComVtaDolar = ComVtaDolar & "   DE : COMPRA Y VENTA EN SOLES " & vbCrLf
        ComVtaDolar = ComVtaDolar & String(vespacio, "-") & vbCrLf
        
        nDebeAcum = 0
        nHaberAcum = 0
        RegTmp.MoveFirst
        Do While Not RegTmp.EOF
            If RegTmp("cTipo") = "1" Then
                sCajaMN = RegTmp("cCtaCnt")
                nDebeAcum = nDebeAcum + RegTmp("nDebe")
                nHaberAcum = nHaberAcum + RegTmp("nHaber")
                ComVtaDolar = ComVtaDolar & ImpreFormat(sCajaMN, 21, 1) & ImpreFormat(UCase(CuentaNombre(sCajaMN, True, oCon.ConexionActiva)), vLenNomb, 1) & ImpreFormat(RegTmp("nDebe"), 12) & ImpreFormat(RegTmp("nHaber"), 12) & vbCrLf
            End If
            RegTmp.MoveNext
        Loop

        ComVtaDolar = ComVtaDolar & String(vespacio, "-") & vbCrLf
        ComVtaDolar = ComVtaDolar & Space(vLenNomb + 12) & ImpreFormat("TOTAL", 11, 0) & _
                    ImpreFormat(nDebeAcum, 12, , True) & _
                    ImpreFormat(nHaberAcum, 12, , True) & vbCrLf
        ComVtaDolar = ComVtaDolar & String(vespacio, "-") & vbCrLf
        ComVtaDolar = ComVtaDolar & vbCrLf
    End If
    RegTmp.Close
    Set RegTmp = Nothing
    
    Me.Caption = "Validación de asiento finalizada"
    'Envia Asiento al Previo
    vRTFImp = vRTFImp & ComVtaDolar
    If Len(Trim(vRTFImp)) = 0 Then
        MsgBox " No existe ningún asiento generado en este día ", vbInformation, " Aviso "
    End If
    
    rtfImp.Text = vRTFImp
    
    'Graba asiento automaticamente
    If Not (Len(Trim(vObs)) > 0 Or Len(Trim(vNoCtaCnt)) > 0) And pAsiDia And CierreRealizado2(1, oCon.ConexionActiva) Then
                
        If gbAsientoDN Then
            vRTFImp = oImpresora.gPrnCondensadaON & vRTFImp & oImpresora.gPrnCondensadaOFF
            rtfImp.Text = vRTFImp
            Me.rtfImp.SaveFile App.Path & "\Spooler\AP" & Right(gsCodAge, 2) & "_" & Format(gdFecSis, "yyyymmdd") & ".txt", 1
        End If
    End If
    
        '*** PEAC 20130119
    If Len(Trim(vObs)) = 0 And Len(Trim(vNoCtaCnt)) = 0 Then
        MsgBox "Proceso Contable Finalizado OK, se mostrará el Asiento.", vbInformation, "Atención"
        cmdAsientos.Enabled = True
        cmdImprimir.Enabled = True
        cmdValidar.Enabled = False
    End If
    
    
    Call GeneraRepoAsntoCnt(vRTFImp)
    
    If optImpresion(0).value = True Then
        'oPrevio.Show rtfImp.Text, " VERIFICACION DE LOS ASIENTOS CONTABLES DEL DIA " & Format(gdHoraGrab, "dd/mm/yyyy"), True, 66, gImpresora
        
        'oPrevio.Show rtfImp.Text, " ASIENTOS CONTABLES DEL DIA " & Format(gdHoraGrab, "dd/mm/yyyy"), True, 66, gImpresora
        Shell "notepad C:\install\prueba2.txt", vbNormalFocus
        
        
    Else
        ImpreBegin True, 66
            Print #ArcSal, rtfImp.Text
        ImpreEnd
    End If
    MousePointer = 0

    
    'Activo Controles
'    cmdAsientos.Enabled = True
    txtFecha.Enabled = True
    fraContenedor(1).Enabled = True
    fraContenedor(2).Enabled = True
    fraImpresion.Enabled = True
    'Fin Activos
    
    Exit Sub
    
ControlError:       ' Rutina de control de errores.
'        prgList.value = 0
'        'Activo Controles
'        cmdAsientos.Enabled = True
'        txtFecha.Enabled = True
'        fraContenedor(1).Enabled = True
'        fraContenedor(2).Enabled = True
'        fraImpresion.Enabled = True
        'Fin Activos
        MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
           " Avise al Area de Sistemas ", vbInformation, " Aviso "

End Sub

'ALPA 20140802**************************************************************************
Sub AsientoCapitlizacionAhorroCTSValida()
Dim tmpSql As String, vNoCtaCnt As String
Dim vMoneda  As String, vCodConta As String, lsCodAge As String, vAgencia As String
Dim nPaso As Integer
Dim vCta28Debe As Currency, vCta28Haber As Currency
Dim vCta19Debe As Currency, vCta19Haber As Currency, vCtaDolar As Currency
Dim RegOpeCta As ADODB.Recordset

    vCta28Debe = 0: vCta28Haber = 0
    vCta19Debe = 0: vCta19Haber = 0
    'vCont = 0: vCtaCnt = 0
    
    'Asiento para Capitalización de Ahorros y CTS
    For nPaso = 1 To 3
        'If pAsiDia Then
            If nPaso = 1 Then
                'Personerias 1,2,3
                tmpSql = " Select   A.cCtaCod, SUBSTRING(A.cCtaCod,4,2) cAgencia, SUBSTRING(A.cCtaCod,9,1) cMoneda, M.cOpeCod, A.nPersoneria, " _
                       & "          OC.cCtaContCod, OC.cOpeCtaDH, SUM(Abs(MC.nMonto)) Monto" _
                       & " FROM OpectaNeg OC" _
                       & "          INNER JOIN Mov M" _
                       & "          INNER JOIN MovCap MC ON M.nMovNro = MC.nMovNro" _
                       & "          INNER JOIN Captaciones A ON MC.cCtaCod = A.cCtaCod ON OC.cOpeCod = M.cOpeCod AND OC.nPersoneria = A.nPersoneria" _
                       & " Where M.cOpeCod IN (Select cOpeCod From OpeTpoSinAsiento Where nProceso = " & gAsientoProcesoCapAho & ")" _
                       & "       And (M.nMovFlag = 0) " _
                       & "       AND A.nPersoneria NOT IN (4,5,6,7,8) " _
                       & "       And M.cMovNro Like '" & Format(ldFechaAsiento, gsFormatoMovFecha) & "%'" _
                       & "  Group by A.cCtaCod,SUBSTRING(A.cCtaCod,4,2), substring(A.cCtaCod,9,1), M.cOpeCod, A.nPersoneria, OC.cCtaContCod, OC.cOpeCtaDH" _
                       & " Order by A.cCtaCod,cAgencia, cMoneda, M.cOpeCod, A.nPersoneria, OC.cOpeCtaDH"

            ElseIf nPaso = 2 Then
                'Personeria 4,5,7,8
                tmpSql = " Select   A.cCtaCod,SUBSTRING(A.cCtaCod,4,2) cAgencia, SUBSTRING(A.cCtaCod,9,1) cMoneda, M.cOpeCod, A.nPersoneria, I.cSubCtaContCod, " _
                       & "          OC.cCtaContCod, OC.cOpeCtaDH, SUM(Abs(MC.nMonto)) Monto" _
                       & " FROM OpectaNeg OC " _
                       & "      JOIN Mov M JOIN MovCap MC " _
                       & "      JOIN Captaciones A JOIN ProductoPersona PP " _
                       & "      JOIN InstitucionFinanc I ON PP.cPersCod = I.cPersCod ON A.cCtaCod = PP.cCtaCod ON " _
                       & "      MC.cCtaCod = A.cCtaCod ON M.nMovNro = MC.nMovNro ON OC.cOpeCod = M.cOpeCod AND OC.nPersoneria = A.nPersoneria " _
                       & " Where M.cOpeCod IN (Select cOpeCod From OpeTpoSinAsiento Where nProceso = " & gAsientoProcesoCapAho & ")" _
                       & "      And (M.nMovFlag = 0) And PP.nPrdPersRelac = 10 And A.nPersoneria NOT IN (1,2,3) " _
                       & "      And M.cMovNro Like '" & Format(ldFechaAsiento, gsFormatoMovFecha) & "%'" _
                       & " Group by A.cCtaCod,SUBSTRING(A.cCtaCod,4,2), substring(A.cCtaCod,9,1), M.cOpeCod, A.nPersoneria, I.cSubCtaContCod, OC.cCtaContCod, OC.cOpeCtaDH" _
                       & " Order by A.cCtaCod,cAgencia, cMoneda, M.cOpeCod, A.nPersoneria, I.cSubCtaContCod, OC.cOpeCtaDH"
                        
            ElseIf nPaso = 3 Then
                'CTS
                tmpSql = "Select    T1.cCtaCod, T1.cAgencia, T1.cTipoCliente, T1.cMoneda, T1.cCtaContCod, T1.cOpeCtaDH, Sum(T1.Monto) Monto, T1.cOpeCod " _
                    & "From " _
                    & "(    Select  CaptacCTS.cCtaCod, cTipoCliente = CASE WHEN LTRIM(RTRIM(CaptacCTS.cCodInst)) = '" & gsCodPersCMACT & "' " _
                    & "             THEN '02' Else '01' END, " _
                    & "             SUBSTRING(CaptacCTS.cCtaCod,9,1) cMoneda, SUBSTRING(CaptacCTS.cCtaCod,4,2) cAgencia, " _
                    & "             LTRIM(RTRIM(OC.cCtaContCod)) cCtaContCod, OC.cOpeCtaDH, Abs(MC.nMonto) Monto, T.cOpeCod " _
                    & "     FROM    Mov T " _
                    & "             INNER JOIN  MovCap MC " _
                    & "             INNER JOIN  MovCapDet MCD " _
                    & "             INNER JOIN  OpeCtaNeg OC ON MCD.cOpeCod = OC.cOpeCod And MCD.nConceptoCod = nConcepto ON MC.nMovNro = MCD.nMovNro And " _
                    & "                         MC.cOpeCod = MCD.cOpeCod And MC.cCtaCod = MCD.cCtaCod ON T.nMovNro = MC.nMovNro " _
                    & "             INNER JOIN CaptacCTS ON MC.cCtaCod = CaptacCTS.cCtaCod " _
                    & "     WHERE   T.cMovnro like '" & Format(ldFechaAsiento, gsFormatoMovFecha) & "%' " _
                    & "             And T.cOpeCod IN (Select cOpeCod From OpeTpoSinAsiento " _
                    & "                       Where nProceso = " & gAsientoProcesoCapCTS & ") AND (T.nMovFlag = 0) ) T1 " _
                    & "GROUP BY    T1.cCtaCod, cAgencia, cTipoCliente, cMoneda, cCtaContCod, cOpeCtaDH,cOpeCod Order by cAgencia, cTipoCliente, cMoneda, cCtaContCod, cOpeCtaDH"
                  
            Else
                MsgBox " Tipo operación Ahorros no reconocida ", vbInformation, " Aviso "
                vNoCtaCnt = vNoCtaCnt & oImpresora.gPrnSaltoLinea & "     * Tipo operación Ahorros no reconocida."
            End If
        
        Set RegOpeCta = New ADODB.Recordset
        RegOpeCta.CursorLocation = adUseClient
        Set RegOpeCta = oCon.CargaRecordSet(tmpSql)
        
        Set RegOpeCta.ActiveConnection = Nothing
        If (RegOpeCta.BOF Or RegOpeCta.EOF) Then
            RegOpeCta.Close
            Set RegOpeCta = Nothing
        Else
            Me.prgList.Max = RegOpeCta.RecordCount
            With RegOpeCta
                Do While Not .EOF
                    If !cOpeCod = "101301" Then
                        vMoneda = vMoneda
                    End If
                    
                    Me.Caption = "Capitalizaciones de Ahorros : Registro " & .Bookmark & " de " & .RecordCount
                    'Variables para cambio en el CodContable
                    vMoneda = Trim(!cMoneda)
                    vCodConta = Trim(!cCtaContCod)
                    vCodConta = Replace(vCodConta, "M", vMoneda, , , vbTextCompare)
                    '**** LAYG - 07/01/05 ica
                    lsCodAge = VarAG(!cCtaCod, !cCtaContCod)
                    vAgencia = lsCodAge
                    vCodConta = Replace(vCodConta, "AG", lsCodAge, , , vbTextCompare)
                    '***************
                    If nPaso = 2 Then
                        vCodConta = Replace(vCodConta, "CJ", Trim(!cSubCtaContCod), , , vbTextCompare)
                        vCodConta = Replace(vCodConta, "CR", Trim(!cSubCtaContCod), , , vbTextCompare)
                        vCodConta = Replace(vCodConta, "CJ", Trim(!cSubCtaContCod), , , vbTextCompare)
                    ElseIf nPaso = 3 Then
                        vCodConta = Replace(vCodConta, "TC", !cTipoCliente, , , vbTextCompare)
                    End If
                        
                    If Not (!cOpeCtaDH = "D" Or !cOpeCtaDH = "H") Then
                        MsgBox " Operación no reconocida al Insertar " & vbCr & _
                        " operación nro.: " & vCodConta, vbInformation, " Aviso "
                        vNoCtaCnt = vNoCtaCnt & oImpresora.gPrnSaltoLinea & "     * " & !cOpeCtaDH & " debe/haber no reconocido (cta.: " & !cOpeCod & ")."
                    End If
                    prgList.Visible = True
                    prgList.value = RegOpeCta.Bookmark
                    .MoveNext
                    DoEvents
                Loop
            End With
            RegOpeCta.Close
            Set RegOpeCta = Nothing
        End If
    Next
End Sub
'APRI20170428
Private Sub TransferenciaAsientoDN()
    'TRANSFERENCIA DEL ASIENTO - SOLO CUANDO ES EL DIA
    'Verifica la veracidad del asiento contable
    
    '2020 peac migra si o si
    'Call TransfiereAsiento
    'Exit Sub
    
    If CDate(txtFecha) = gdFecSis Then
        If Not (Len(Trim(vObs)) > 0 Or Len(Trim(vNoCtaCnt)) > 0) And pAsiDia And CierreRealizado2(1, oCon.ConexionActiva) Then
            If MsgBox(" Desea migrar el Asiento ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
                Call TransfiereAsiento
                MsgBox "Asiento Transferido satisfactoriamente...!", vbInformation, "¡Aviso!"
            End If
        End If
    Else
        If Not (Len(Trim(vObs)) > 0 Or Len(Trim(vNoCtaCnt)) > 0) And pAsiDia = False Then
            If MsgBox(" Desea migrar el Asiento ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
                Call TransfiereAsiento
                MsgBox "Asiento Transferido satisfactoriamente...!", vbInformation, "¡Aviso!"
            End If
        End If
    End If
    
    
End Sub
'END APRI


