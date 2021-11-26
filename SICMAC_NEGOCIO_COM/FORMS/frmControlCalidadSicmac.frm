VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmControlCalidadSicmac 
   Caption         =   "Control de Calidad de Operaciones SICMAC-I"
   ClientHeight    =   5910
   ClientLeft      =   3390
   ClientTop       =   2130
   ClientWidth     =   5970
   Icon            =   "frmControlCalidadSicmac.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5910
   ScaleWidth      =   5970
   Begin VB.CheckBox chkInst 
      Caption         =   "Por Instituciones"
      Height          =   300
      Left            =   3750
      TabIndex        =   23
      Top             =   1620
      Width           =   1800
   End
   Begin VB.CommandButton cmdValidar 
      Caption         =   "&Validar"
      Height          =   405
      Left            =   4215
      TabIndex        =   9
      Top             =   3735
      Width           =   1530
   End
   Begin VB.Frame Frame2 
      Height          =   4170
      Left            =   105
      TabIndex        =   6
      Top             =   1545
      Width           =   3480
      Begin VB.OptionButton optCalidad 
         Caption         =   "Retiros de Ahorros"
         Height          =   375
         Index           =   8
         Left            =   135
         TabIndex        =   25
         Top             =   2985
         Width           =   2145
      End
      Begin VB.OptionButton optCalidad 
         Caption         =   "Depositos de Ahorros"
         Height          =   375
         Index           =   7
         Left            =   105
         TabIndex        =   24
         Top             =   2625
         Width           =   2145
      End
      Begin VB.CommandButton cmdCalend 
         Caption         =   "Verifica Calendarios"
         Height          =   360
         Left            =   960
         TabIndex        =   22
         Top             =   3600
         Width           =   1980
      End
      Begin VB.OptionButton optCalidad 
         Caption         =   "Cancelaciones de Ahorros y PF"
         Height          =   375
         Index           =   6
         Left            =   105
         TabIndex        =   20
         Top             =   2250
         Width           =   3150
      End
      Begin VB.OptionButton optCalidad 
         Caption         =   "Pagos Creditos Prendario"
         Height          =   300
         Index           =   5
         Left            =   105
         TabIndex        =   19
         Top             =   1905
         Width           =   2325
      End
      Begin VB.OptionButton optCalidad 
         Caption         =   "Desembolsos Creditos Prendarios"
         Height          =   360
         Index           =   4
         Left            =   105
         TabIndex        =   18
         Top             =   1530
         Width           =   3150
      End
      Begin VB.OptionButton optCalidad 
         Caption         =   "Pago de Creditos Judiciales"
         Height          =   300
         Index           =   3
         Left            =   105
         TabIndex        =   11
         Top             =   1200
         Width           =   2325
      End
      Begin VB.OptionButton optCalidad 
         Caption         =   "Pago de Creditos RFA"
         Height          =   300
         Index           =   2
         Left            =   105
         TabIndex        =   10
         Top             =   855
         Width           =   2325
      End
      Begin VB.OptionButton optCalidad 
         Caption         =   "Pagos de Creditos normales "
         Height          =   300
         Index           =   1
         Left            =   105
         TabIndex        =   8
         Top             =   510
         Width           =   2325
      End
      Begin VB.OptionButton optCalidad 
         Caption         =   "Desembolsos de Creditos por Productos"
         Height          =   375
         Index           =   0
         Left            =   105
         TabIndex        =   7
         Top             =   165
         Value           =   -1  'True
         Width           =   3285
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos Principales"
      Height          =   1440
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   5715
      Begin VB.OptionButton optMoneda 
         Caption         =   "Dolares"
         Height          =   240
         Index           =   1
         Left            =   2835
         TabIndex        =   17
         Top             =   1065
         Width           =   945
      End
      Begin VB.OptionButton optMoneda 
         Caption         =   "Soles"
         Height          =   255
         Index           =   0
         Left            =   2085
         TabIndex        =   5
         Top             =   1035
         Value           =   -1  'True
         Width           =   915
      End
      Begin MSMask.MaskEdBox txtFecha 
         Height          =   330
         Left            =   795
         TabIndex        =   4
         Top             =   990
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   582
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin SICMACT.TxtBuscar TxtBuscarUser 
         Height          =   345
         Left            =   960
         TabIndex        =   12
         Top             =   570
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   609
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
         sTitulo         =   ""
         ForeColor       =   12582912
      End
      Begin SICMACT.TxtBuscar TxtBuscarAge 
         Height          =   345
         Left            =   960
         TabIndex        =   13
         Top             =   225
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   609
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
         sTitulo         =   ""
         ForeColor       =   12582912
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "User :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   165
         TabIndex        =   16
         Top             =   675
         Width           =   480
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         Caption         =   "Agencia :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   165
         TabIndex        =   15
         Top             =   285
         Width           =   750
      End
      Begin VB.Label lblDescAge 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   1935
         TabIndex        =   14
         Top             =   225
         Width           =   3570
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         Height          =   195
         Left            =   180
         TabIndex        =   3
         Top             =   1065
         Width           =   495
      End
      Begin VB.Label lblUsuSIAFC 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   3300
         TabIndex        =   2
         Top             =   600
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Usuario SIAFC:"
         Height          =   195
         Left            =   2070
         TabIndex        =   1
         Top             =   645
         Width           =   1080
      End
   End
   Begin SICMACT.Usuario Usuario 
      Left            =   5250
      Top             =   1950
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.Label lblmensaje 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   330
      Left            =   3765
      TabIndex        =   21
      Top             =   3030
      Width           =   2055
   End
End
Attribute VB_Name = "frmControlCalidadSicmac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dbBase As ADODB.Connection
Dim gsBaseCred As String
Dim gsBaseCli As String
Dim gsBaseAho As String
Dim gsBaseKPR As String
Dim gsBaseTCOS As String
Dim lscadena As String
Dim lnLinea As String
Dim oGen As COMDConstSistema.DCOMGeneral
Dim oPrev As clsprevio
Dim lsTmpPagos As String

Private Sub cmdCalend_Click()
DeterminaRutas
frmVerCalendarios.Inicio gsBaseCli, gsBaseCred, dbBase
End Sub

Private Sub cmdValidar_Click()
Dim lsMoneda As String
Dim lsFecha As String
If optCalidad(7).value = False And optCalidad(8).value = False And optCalidad(1).value = False Then
    If txtBuscarUser = "" Then
        MsgBox "Usuario no valido", vbInformation, "Aviso"
        Exit Sub
    End If
    If Me.lblUsuSIAFC = "" Then
        MsgBox "Usuario del SIAFC no valido", vbInformation, "Aviso"
        Exit Sub
    End If
End If
If Me.TxtBuscarAge = "" Then
    MsgBox "Agencia no valida", vbInformation, "Aviso"
    Exit Sub
End If

If ValFecha(Me.txtfecha) = False Then
    Exit Sub
End If
lsMoneda = IIf(Me.optmoneda(0).value = True, "1", "2")
lsFecha = Me.txtfecha.Text
Me.lblMensaje = "POR FAVOR ESPERE..."
Me.lblMensaje.Refresh
If optCalidad(0).value = True Then
    ValidaDesembolsos lsFecha, lblUsuSIAFC, lsMoneda, txtBuscarUser
End If
If optCalidad(1).value = True Then
    ProcesaPagosSICMAC lsFecha, lsMoneda, Me.lblUsuSIAFC, txtBuscarUser, False, False
End If
If optCalidad(2).value = True Then
    ProcesaPagosSICMAC lsFecha, lsMoneda, Me.lblUsuSIAFC, txtBuscarUser, True, False
End If
If optCalidad(3).value = True Then
    ProcesaPagosSICMAC lsFecha, lsMoneda, Me.lblUsuSIAFC, txtBuscarUser, False, True
End If
If optCalidad(4).value = True Then
    ValidaDesembolsosPren lsFecha, lblUsuSIAFC, lsMoneda, txtBuscarUser
End If
If optCalidad(5).value = True Then
    ProcesaPagosPRENDSICMAC lsFecha, lsMoneda, Me.lblUsuSIAFC, txtBuscarUser
End If

If optCalidad(6).value = True Then
    CancelacionesAhorros lsFecha, lsMoneda, Me.lblUsuSIAFC, txtBuscarUser
End If

If optCalidad(7).value = True Then
    MovimientosAhorros lsFecha, lsMoneda, Me.lblUsuSIAFC, txtBuscarUser, True
End If

If optCalidad(8).value = True Then
    MovimientosAhorros lsFecha, lsMoneda, Me.lblUsuSIAFC, txtBuscarUser, False
End If

Me.lblMensaje = ""
Me.lblMensaje.Refresh
End Sub
Function GetUserSiac(ByVal psCodUsu As String)
Dim sql As String
Dim rs As ADODB.Recordset

DeterminaRutas

Set rs = New ADODB.Recordset
sql = "SELECT CCODUSU FROM " & gsBaseTCOS & "TCOTUSU WHERE CCODUSUWIN ='" & psCodUsu & "'"
Set rs = CargaRecordFox(sql)
If Not rs.EOF And Not rs.BOF Then
    GetUserSiac = rs!cCodUsu
Else
    GetUserSiac = psCodUsu
End If
rs.Close
Set rs = Nothing


End Function
Private Sub Form_Load()

Set oPrev = New clsprevio
CentraForm Me
AbreConexionFox
Usuario.Inicio gsCodUser
txtfecha = gdFecSis
Set oGen = New COMDConstSistema.DCOMGeneral
TxtBuscarAge.psRaiz = "AGENCIAS"
Me.TxtBuscarAge.rs = oGen.GetAgenciasArbol
TxtBuscarAge.Text = gsCodAge
'TxtBuscarAge.Enabled = False
lblDescAge = gsNomAge
txtBuscarUser.psRaiz = "USUARIOS"
txtBuscarUser.rs = oGen.GetUserAreaAgencia("026", TxtBuscarAge.Text, "", False)
txtBuscarUser.Enabled = True
txtBuscarUser.Text = ""
End Sub

Sub AbreConexionFox()
On Error GoTo ErrorConex

    psConexion = "Driver={Microsoft Visual FoxPro Driver};UID=;SourceDB=F:\APL\TCOS\;SourceType=DBF;Exclusive=No;Collate=GENERAL"
    
    Set dbBase = New ADODB.Connection
    dbBase.Open psConexion
    dbBase.CommandTimeout = 7200
    dbBase.Execute "SET DELETE ON"
    
    Exit Sub
ErrorConex:
    MsgBox "Error:" & Err.Description & " [" & Err.Number & "] ", vbInformation, "Aviso"
    End
End Sub
Function CargaRecordFox(ByVal sql As String) As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseServer
    rs.Open sql, dbBase, adOpenKeyset, adLockOptimistic, adCmdText
    
    Set CargaRecordFox = rs
    'rs.ActiveConnection = Nothing
End Function
Public Sub CierraConexionFox()
  If nSalida = 1 Then
     If dbBase Is Nothing Then Exit Sub
     If dbBase.State = adStateOpen Then
        dbBase.Close
        Set dbBase = Nothing
     End If
  End If
End Sub
Sub DeterminaRutas()
    Select Case Trim(TxtBuscarAge)
        Case "01" 'ica
            gsBaseCred = "F:\APL.FIN\KPY.ICA\"
            gsBaseCli = "F:\APL.FIN\CLI.ICA\"
            gsBaseAho = "F:\APL.FIN\AHO.ICA\"
            gsBaseKPR = "F:\APL.FIN\KPR.ICA\"
            gsBaseTCOS = "F:\APL.FIN\TCOS.ICA\"
            
            gsBaseCred = "F:\APL\KPY\"
            gsBaseCli = "F:\APL\CLI\"
            gsBaseAho = "F:\APL\AHO\"
            gsBaseKPR = "F:\APL\KPR\"
            gsBaseTCOS = "F:\APL\TCOS\"
        Case "13" 'CAÑETE
            gsBaseCred = "F:\APL.FIN\KPY.CAN\"
            gsBaseCli = "F:\APL.FIN\CLI.CAN\"
            gsBaseAho = "F:\APL.FIN\AHO.CAN\"
            gsBaseKPR = "F:\APL.FIN\KPR.CAN\"
            gsBaseTCOS = "F:\APL.FIN\TCOS.CAN\"
            
            gsBaseCred = "F:\AGENCIAS\CANETE\APL\KPY\"
            gsBaseCli = "F:\AGENCIAS\CANETE\APL\CLI\"
            gsBaseAho = "F:\AGENCIAS\CANETE\APL\AHO\"
            gsBaseKPR = "F:\AGENCIAS\CANETE\APL\KPR\"
            gsBaseTCOS = "F:\AGENCIAS\CANETE\APL\TCOS\"
            
        Case "02" 'CHINCHA
            gsBaseCred = "F:\APL.FIN\KPY.CHI\"
            gsBaseCli = "F:\APL.FIN\CLI.CHI\"
            gsBaseAho = "F:\APL.FIN\AHO.ICA\"
            gsBaseKPR = "F:\APL.FIN\KPR.ICA\"
            gsBaseTCOS = "F:\APL.FIN\TCOS.CHI\"
        Case "04" 'NASCA
            gsBaseCred = "F:\APL.FIN\KPY.NAS\"
            gsBaseCli = "F:\APL.FIN\CLI.NAS\"
            gsBaseAho = "F:\APL.FIN\AHO.NAS\"
            gsBaseKPR = "F:\APL.FIN\KPR.NAS\"
            gsBaseTCOS = "F:\APL.FIN\TCOS.NAS\"
            
            gsBaseCred = "F:\AGENCIAS\NASCA\APL\KPY\"
            gsBaseCli = "F:\AGENCIAS\NASCA\APL\CLI\"
            gsBaseAho = "F:\AGENCIAS\NASCA\APL\AHO\"
            gsBaseKPR = "F:\AGENCIAS\NASCA\APL\KPR\"
            gsBaseTCOS = "F:\AGENCIAS\NASCA\APL\TCOS\"
        Case "07" 'PUQUIO
            gsBaseCred = "F:\APL.FIN\KPY.PUQ\"
            gsBaseCli = "F:\APL.FIN\CLI.PUQ\"
            gsBaseAho = "F:\APL.FIN\AHO.PUQ\"
            gsBaseKPR = "F:\APL.FIN\KPR.PUQ\"
            gsBaseTCOS = "F:\APL.FIN\TCOS.PUQ\"
            
            gsBaseCred = "F:\AGENCIAS\PUQUIO\APL\KPY\"
            gsBaseCli = "F:\AGENCIAS\PUQUIO\APL\CLI\"
            gsBaseAho = "F:\AGENCIAS\PUQUIO\APL\AHO\"
            gsBaseKPR = "F:\AGENCIAS\PUQUIO\APL\KPR\"
            gsBaseTCOS = "F:\AGENCIAS\PUQUIO\APL\TCOS\"
        Case "08" 'HUAMANGA
            gsBaseCred = "F:\APL.FIN\KPY.HUA\"
            gsBaseCli = "F:\APL.FIN\CLI.HUA\"
            gsBaseAho = "F:\APL.FIN\AHO.HUA\"
            gsBaseKPR = "F:\APL.FIN\KPR.HUA\"
            gsBaseTCOS = "F:\APL.FIN\TCOS.HUA\"
            
            gsBaseCred = "F:\AGENCIAS\HUAMANGA\APL\KPY\"
            gsBaseCli = "F:\AGENCIAS\HUAMANGA\APL\CLI\"
            gsBaseAho = "F:\AGENCIAS\HUAMANGA\APL\AHO\"
            gsBaseKPR = "F:\AGENCIAS\HUAMANGA\APL\KPR\"
            gsBaseTCOS = "F:\AGENCIAS\HUAMANGA\APL\TCOS\"
        Case "06" 'MALA
            gsBaseCred = "F:\APL.FIN\KPY.MAL\"
            gsBaseCli = "F:\APL.FIN\CLI.MAL\"
            gsBaseAho = "F:\APL.FIN\AHO.MAL\"
            gsBaseKPR = "F:\APL.FIN\KPR.MAL\"
            gsBaseTCOS = "F:\APL.FIN\TCOS.MAL\"
            
            gsBaseCred = "F:\AGENCIAS\MALA\APL\KPY\"
            gsBaseCli = "F:\AGENCIAS\MALA\APL\CLI\"
            gsBaseAho = "F:\AGENCIAS\MALA\APL\AHO\"
            gsBaseKPR = "F:\AGENCIAS\MALA\APL\KPR\"
            gsBaseTCOS = "F:\AGENCIAS\MALA\APL\TCOS\"
        Case "05" 'PALPA
            gsBaseCred = "F:\APL.FIN\KPY.PAL\"
            gsBaseCli = "F:\APL.FIN\CLI.PAL\"
            gsBaseAho = "F:\APL.FIN\AHO.PAL\"
            gsBaseKPR = "F:\APL.FIN\KPR.PAL\"
            gsBaseTCOS = "F:\APL.FIN\TCOS.PAL\"
            
            gsBaseCred = "F:\AGENCIAS\PALPA\APL\KPY\"
            gsBaseCli = "F:\AGENCIAS\PALPA\APL\CLI\"
            gsBaseAho = "F:\AGENCIAS\PALPA\APL\AHO\"
            gsBaseKPR = "F:\AGENCIAS\PALPA\APL\KPR\"
            gsBaseTCOS = "F:\AGENCIAS\PALPA\APL\TCOS\"
        Case "03" 'IMPERIAL
            gsBaseCred = "F:\APL.FIN\KPY.IMP\"
            gsBaseCli = "F:\APL.FIN\CLI.IMP\"
            gsBaseAho = "F:\APL.FIN\AHO.IMP\"
            gsBaseKPR = "F:\APL.FIN\KPR.IMP\"
            gsBaseTCOS = "F:\APL.FIN\TCOS.IMP\"
            
            gsBaseCred = "F:\AGENCIAS\IMPERIAL\APL\KPY\"
            gsBaseCli = "F:\AGENCIAS\IMPERIAL\APL\CLI\"
            gsBaseAho = "F:\AGENCIAS\IMPERIAL\APL\AHO\"
            gsBaseKPR = "F:\AGENCIAS\IMPERIAL\APL\KPR\"
            gsBaseTCOS = "F:\AGENCIAS\IMPERIAL\APL\TCOS\"
    End Select
End Sub
Function CredDesembolsadosSIAFC(ByVal pdFecha As Date, ByVal psMoneda As String, ByVal psCodUsu As String) As ADODB.Recordset
Dim sql As String
'Dim rs As ADODB.Recordset
DeterminaRutas
sql = "SELECT   IIF(Substr(KD.cCodCta,7,2)='01','MES      ', " _
    & "         IIF(Substr(KD.cCodCta,7,2)='04','COMERCIAL',  " _
    & "         IIF(Substr(KD.cCodCta,7,2)='03','AGRICOLA ',  " _
    & "         IIF(Substr(KD.cCodCta,7,2)='02','CONSUMO  ',  " _
    & "         IIF(Substr(KD.cCodCta,7,2)='05','HIPOTECAR','******************'))) ) ) as cProducto,  " _
    & "         KD.cCodCta, C.CNOMCLI, KD.CNROCUO,   " _
    & "         SUM(IIF(KD.cConcep$'KP,KS,KC', KD.NMONTO,00000000.00 )) AS CAPITAL,  " _
    & "         SUM(IIF(KD.cConcep='IN', KD.NMONTO,00000000.00 )) AS INTERES,  " _
    & "         SUM(IIF(KD.cConcep='MO', KD.NMONTO,00000000.00 )) AS MORA,   " _
    & "         SUM(IIF(KD.cConcep='01', KD.NMONTO,00000000.00 )) AS PROTESTO,  " _
    & "         SUM(IIF(KD.cConcep='05', KD.NMONTO,00000000.00 )) AS NCOMISION,   " _
    & "         SUM(IIF(KD.cConcep='IT', KD.NMONTO,00000000.00 )) AS ITF,  " _
    & "         SUM(IIF(!KD.cConcep$'KP,IN,MO,01,05,KS,KC,AC', NMONTO,00000000.00 )) AS NOTROS,   " _
    & "         SUM(IIF(KD.cConcep$'CJ,AC', KD.NMONTO,00000000.00 )) AS CAJA, KD.CCODUSU   " _
    & " FROM    " & gsBaseCred & "KPYDKAR KD " _
    & "         JOIN " & gsBaseCred & "KPYMCRE K ON K.CCODCTA =KD.CCODCTA  " _
    & "         JOIN " & gsBaseCli & "CLIMIDE C  ON C.CCODCLI = K.CCODCLI  " _
    & " WHERE   KD.cDescOb ='D' and KD.dFecPro = CTOD('" & Format(pdFecha, "mm/dd/yyyy") & "')  " _
    & "         AND KD.CMONEDA ='" & psMoneda & "' AND KD.CCODUSU ='" & psCodUsu & "' and empty(KD.cEstado)  " _
    & " GROUP BY KD.CCODCTA, KD.CNROCUO, KD.CCODUSU   " _
    & " ORDER BY CPRODUCTO, CNOMCLI "

'ORDER BY CPRODUCTO, CNOMCLI
Set CredDesembolsadosSIAFC = CargaRecordFox(sql)

End Function
Function CredPagosSIAFC(ByVal pdFecha As Date, ByVal psMoneda As String, ByVal psCodUsu As String, Optional lbRfa As Boolean = False, Optional lbJudicial As Boolean = False) As ADODB.Recordset
Dim sql As String
Dim lsCondJud As String
Dim lsCondRfa  As String
Dim lsFiltroUser As String
'Dim rs As ADODB.Recordset
DeterminaRutas
If lbRfa Then
    lsCondRfa = " and KD.cCodRec$'120,121,122' "
Else
    lsCondRfa = " and !KD.cCodRec$'120,121,122' "
End If
If lbJudicial Then
    lsCondJud = " and KD.cNroCuo$'CAS,JUD' "
Else
    lsCondJud = " and !KD.cNroCuo$'CAS,JUD' "
End If

If psCodUsu = "" Then
    lsFiltroUser = ""
Else
    lsFiltroUser = " AND KD.CCODUSU ='" & psCodUsu & "' "
End If

If Me.chkInst.value = 0 Then
sql = "SELECT  'SIAFC' as cSistema, KD.cCodCta as cCtaCod, K.NCAPDES-K.NCAPPAG AS NSALDO,  " _
    & "         iif(K.CESTADO='F','VIGENTE',  " _
    & "         IIF(K.CESTADO ='H','JUDICIAL', " _
    & "         IIF(K.CESTADO='G','CANCELADO',IIF(K.CESTADO='O','CASTIGADO','*********')))) AS cESTADO,  " _
    & "         C.CNOMCLI as cPersNombre, KD.CCODREC as cRfa, KD.cCodCta as cCtaCodAnt, " _
    & "         SUM(IIF(KD.cConcep$'KP,KS,KC,DK,KD', KD.NMONTO,00000000.00 )) AS CAPITAL, " _
    & "         SUM(IIF(KD.cConcep$'IN,IS,IC', KD.NMONTO,00000000.00 )) AS INTERES, " _
    & "         SUM(IIF(KD.cConcep='MO', KD.NMONTO,00000000.00 )) AS MORA,  " _
    & "         SUM(IIF(KD.cConcep='ID', KD.NMONTO,00000000.00 )) AS DESAGIO,  " _
    & "         SUM(IIF(KD.cConcep='IT', KD.NMONTO,00000000.00 )) AS ITF,  " _
    & "         SUM(IIF(KD.cConcep='01', KD.NMONTO,00000000.00 )) AS PROTESTO,  " _
    & "         SUM(IIF(KD.cConcep='05', KD.NMONTO,00000000.00 )) AS NCOMISION,  " _
    & "         SUM(IIF(!KD.cConcep$'KP,KS,KC,IN,IS,IC,MO,01,05,IT,CJ,ID,AC,DK,KD', NMONTO,00000000.00 )) AS NOTROS, " _
    & "         SUM(IIF(KD.cConcep$'CJ,AC', KD.NMONTO,00000000.00 )) AS CAJA, space(3) as cCodIns, Space(40) as cInst, KD.cCodusu as cUser " _
    & " FROM    " & gsBaseCred & "KPYDKAR KD " _
    & "         JOIN " & gsBaseCred & "KPYMCRE K ON K.CCODCTA =KD.CCODCTA " _
    & "         JOIN " & gsBaseCli & "CLIMIDE  C  ON C.CCODCLI = K.CCODCLI " _
    & " WHERE   KD.cDescOb !='D' and KD.dFecPro = CTOD('" & Format(pdFecha, "mm/dd/yyyy") & "')  " _
    & "         AND KD.CMONEDA ='" & psMoneda & "' " & lsFiltroUser & " and empty(KD.cEstado) " & lsCondRfa & lsCondJud _
    & " GROUP BY KD.CCODCTA, KD.CCODUSU  " _
    & " ORDER BY CNOMCLI"
    
Else
sql = "SELECT  'SIAFC' as cSistema, KD.cCodCta as cCtaCod, K.NCAPDES-K.NCAPPAG AS NSALDO,  " _
    & "         iif(K.CESTADO='F','VIGENTE',  " _
    & "         IIF(K.CESTADO ='H','JUDICIAL', " _
    & "         IIF(K.CESTADO='G','CANCELADO',IIF(K.CESTADO='O','CASTIGADO','*********')))) AS cESTADO,  " _
    & "         C.CNOMCLI as cPersNombre, KD.CCODREC as cRfa, KD.cCodCta as cCtaCodAnt, " _
    & "         SUM(IIF(KD.cConcep$'KP,KS,KC,DK,KD', KD.NMONTO,00000000.00 )) AS CAPITAL, " _
    & "         SUM(IIF(KD.cConcep$'IN,IS,IC', KD.NMONTO,00000000.00 )) AS INTERES, " _
    & "         SUM(IIF(KD.cConcep='MO', KD.NMONTO,00000000.00 )) AS MORA,  " _
    & "         SUM(IIF(KD.cConcep='ID', KD.NMONTO,00000000.00 )) AS DESAGIO,  " _
    & "         SUM(IIF(KD.cConcep='IT', KD.NMONTO,00000000.00 )) AS ITF,  " _
    & "         SUM(IIF(KD.cConcep='01', KD.NMONTO,00000000.00 )) AS PROTESTO,  " _
    & "         SUM(IIF(KD.cConcep='05', KD.NMONTO,00000000.00 )) AS NCOMISION,  " _
    & "         SUM(IIF(!KD.cConcep$'KP,KS,KC,IN,IS,IC,MO,01,05,IT,CJ,ID,AC,DK,KD', NMONTO,00000000.00 )) AS NOTROS, " _
    & "         SUM(IIF(KD.cConcep$'CJ,AC', KD.NMONTO,00000000.00 )) AS CAJA, K.CCODINS, LEFT(T.CDESCRI,40) AS CINST, KD.cCodusu as cUser " _
    & " FROM    " & gsBaseCred & "KPYDKAR KD " _
    & "         JOIN " & gsBaseCred & "KPYMCRE K ON K.CCODCTA =KD.CCODCTA " _
    & "         JOIN " & gsBaseCli & "CLIMIDE  C  ON C.CCODCLI = K.CCODCLI " _
    & "         JOIN F:\APL\TCO\TCOTTAB T ON T.CCODIGO = K.CCODINS AND T.CCODTAB = '077' " _
    & " WHERE   KD.cDescOb !='D' and KD.dFecPro = CTOD('" & Format(pdFecha, "mm/dd/yyyy") & "')  " _
    & "         AND KD.CMONEDA ='" & psMoneda & "' " & lsFiltroUser & " and empty(KD.cEstado) " & lsCondRfa & lsCondJud _
    & " GROUP BY KD.CCODCTA, KD.CCODUSU  " _
    & " ORDER BY CNOMCLI"
End If

Set CredPagosSIAFC = CargaRecordFox(sql)
End Function

Sub ValidaDesembolsos(ByVal pdFecha As Date, ByVal psCodUsu As String, ByVal psMoneda As String, ByVal psUsuSICMAC As String)
Dim rs As ADODB.Recordset
Dim lsCadena1 As String
Set rs = CredDesembolsadosSIAFC(pdFecha, psMoneda, psCodUsu)

lsCadena1 = ReporteDesembolsos(rs, pdFecha, psCodUsu, psMoneda, "INFORMACION SIAFC", 0)
Set rs = GetDesembolsoSicmact(pdFecha, psUsuSICMAC, psMoneda)
lsCadena1 = lsCadena1 + ReporteDesembolsos(rs, pdFecha, psUsuSICMAC, psMoneda, "INFORMACION SICMAC-I", 1)

'oPrev.Show lsCadena1, "", True
oPrev.Show lsCadena1, "", True, , gImpresora
End Sub
Function ReporteDesembolsos(ByVal rs As ADODB.Recordset, ByVal pdFecha As Date, ByVal psCodUsu As String, ByVal psMoneda As String, ByVal lsInfoSicmac_Siafc As String, lnPagina As Integer) As String
Dim lnTotalProdCap As Currency
Dim lnTotalProdITF As Currency
Dim lnTotalProdCaj As Currency
Dim lscadena As String
Dim lnTotalCaso  As Long


If Not rs.EOF And Not rs.BOF Then
lscadena = CabeRepo(gsNomCmac, lblDescAge, "", IIf(psMoneda = "1", "SOLES", "DOLARES"), Format(pdFecha, "dd/mm/yyyy"), "Control Calidad Desembolsos", "USUARIO :" & psCodUsu, lsInfoSicmac_Siafc, "", lnPagina, 64)
lscadena = lscadena + Chr(10)
lnTotalProd = 0
lnLinea = 0
lsProducto = ""
lscadena = lscadena & String(100, "-") & Chr(10)
lscadena = lscadena & ImpreFormat("Producto", 15) & _
                       ImpreFormat("N° Credito", 18) & _
                       ImpreFormat("Nombre Cliente", 40) & _
                       ImpreFormat("CAPITAL", 12, 2) & _
                       ImpreFormat("I.T.F.", 12, 2) & _
                       ImpreFormat("CAJA", 12, 2) & Chr(10)

lscadena = lscadena & String(100, "-") & Chr(10)
lnTotalProdCap = 0
lnTotalProdITF = 0
lnTotalProdCaj = 0
lnTotalCaso = 0
Do While Not rs.EOF
    
    If lsProducto <> Trim(rs!cProducto) Then
        lsProducto = Trim(rs!cProducto)
        If lnTotalProdCap > 0 Then
            lscadena = lscadena & String(55, "-") & Chr(10)
            lscadena = lscadena & ImpreFormat("TOTAL PRODUCTO      N°", 15) & _
                       ImpreFormat(lnTotalCaso, 18) & _
                       ImpreFormat("", 40) & _
                       ImpreFormat(lnTotalProdCap, 12, 2) & _
                       ImpreFormat(lnTotalProdITF, 12, 2) & _
                       ImpreFormat(lnTotalProdCaj, 12, 2) & Chr(10)
            
            lscadena = lscadena & String(55, "-") & Chr(10) & Chr(10)
            lnTotalProdCap = 0
            lnTotalProdITF = 0
            lnTotalProdCaj = 0
            lnTotalCaso = 0
        End If
    End If
    lnLinea = lnLinea + 1
    lscadena = lscadena & ImpreFormat(rs!cProducto, 15) & _
                        ImpreFormat(rs!cCodCta, 18) & _
                       ImpreFormat(Trim(rs!cNomCli), 40) & _
                        ImpreFormat(rs!capital, 12, 2) & _
                       ImpreFormat(rs!ITF, 12, 2) & _
                       ImpreFormat(rs!CAJA, 12, 2) & Chr(10)
    
    lnTotalProdCap = lnTotalProdCap + rs!capital
    lnTotalProdITF = lnTotalProdITF + rs!ITF
    lnTotalProdCaj = lnTotalProdCaj + rs!CAJA
    lnTotalCaso = lnTotalCaso + 1
    
    If lnLinea > 63 Then
        lnLinea = 1
        lscadena = lscadena & Chr(12)
        lscadena = lscadena & ImpreFormat("Producto", 15) & _
                       ImpreFormat("N° Credito", 18) & _
                       ImpreFormat("Nombre Cliente", 40) & _
                       ImpreFormat("CAPITAL", 12, 2) & _
                       ImpreFormat("I.T.F.", 12, 2) & _
                       ImpreFormat("CAJA", 12, 2) & Chr(10)
        
    End If
    rs.MoveNext
Loop
    If lnTotalProdCap > 0 Then
            lscadena = lscadena & String(55, "-") & Chr(10)
            lscadena = lscadena & ImpreFormat("TOTAL PRODUCTO    N°", 15) & _
                       ImpreFormat(lnTotalCaso, 18) & _
                       ImpreFormat("", 40) & _
                       ImpreFormat(lnTotalProdCap, 12, 2) & _
                       ImpreFormat(lnTotalProdITF, 12, 2) & _
                       ImpreFormat(lnTotalProdCaj, 12, 2) & Chr(10)
            
            lscadena = lscadena & String(55, "-") & Chr(10) & Chr(10)
            lnTotalProdCap = 0
            lnTotalProdITF = 0
            lnTotalProdCaj = 0
            lnTotalCaso = 0
        End If
End If

rs.Close
Set rs = Nothing
ReporteDesembolsos = lscadena
End Function
Function GetDesembolsoSicmact(ByVal pdFecha As Date, ByVal psCodUsu As String, ByVal psMoneda As String) As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim oCon As COMConecta.DCOMConecta

Set oCon = New COMConecta.DCOMConecta

sql = " SELECT * " _
    & " FROM (      " _
    & "         SELECT  CASE    WHEN SUBSTRING(T.CCODCTA,6,3) = '202' THEN 'AGRICOLA' " _
    & "                 Else " _
    & "                     CASE WHEN SUBSTRING(T.CCODCTA,6,1) = '1' THEN 'COMERCIAL' " _
    & "                     Else " _
    & "                         CASE WHEN SUBSTRING(T.CCODCTA,6,1) = '2' AND SUBSTRING(T.CCODCTA,6,3) <> '202' THEN 'MES     ' " _
    & "                             Else " _
    & "                                 CASE WHEN SUBSTRING(T.CCODCTA,6,1) = '3' AND SUBSTRING(T.CCODCTA,6,3) <> '305' THEN 'CONSUMO' " _
    & "                                 Else " _
    & "                                         CASE WHEN SUBSTRING(T.CCODCTA,6,1) = '4' THEN 'HIPOTECARIO' " _
    & "                                         Else " _
    & "                                             '**************' " _
    & "                                         End " _
    & "                                 End " _
    & "                             End " _
    & "                     End " _
    & "             END AS cPRODUCTO, " _
    & "             SUBSTRING(T.CCODCTA,6,3) AS CCODPROD, " _
    & "             P.CPERSNOMBRE AS CNOMCLI,T.*, T.CAPITAL - T.ITF AS CAJA " _
    & "             From  " _
    & "                 (SELECT MD.CCTACOD AS CCODCTA, M.nmovnro, " _
    & "                         SUM(CASE WHEN MD.cOpeCod ='100101' AND nPrdConceptoCod=1000 THEN MD.nMonto ELSE 0 END) AS CAPITAL, " _
    & "                         SUM(CASE WHEN MD.cOpeCod ='990106' THEN MD.nMonto ELSE 0 END) AS ITF " _
    & "                     FROM    MOV M "
sql = sql + "               JOIN MOVCOL     MC ON MC.NMOVNRO = M.NMOVNRO" _
    & "                     JOIN MOVCOLDET  MD ON MD.NMOVNRO = MC.NMOVNRO AND MD.CCTACOD = MC.CCTACOD AND MD.COPECOD = MC.COPECOD " _
    & "                 WHERE   LEFT(M.CMOVNRO,8) ='" & Format(pdFecha, "yyyymmdd") & "' AND RIGHT(CMOVNRO,4)='" & psCodUsu & "' AND SUBSTRING(MC.CCTACOD,9,1)='" & psMoneda & "' " _
    & "                         AND M.COPECOD IN ('100101','100102','100103','100104','990106') AND M.nMovFlag ='0' " _
    & "                 GROUP BY MD.CCTACOD, M.nmovnro) AS T " _
    & "             JOIN PRODUCTOPERSONA PP ON PP.CCTACOD = T.CCODCTA AND nPrdPersRelac= 20 " _
    & " JOIN PERSONA P ON P.CPERSCOD = PP.CPERSCOD ) AS X " _
    & " ORDER BY cPRODUCTO, CNOMCLI "

oCon.AbreConexion
Set GetDesembolsoSicmact = oCon.CargaRecordSet(sql)
oCon.CierraConexion

End Function
Private Sub Form_Unload(Cancel As Integer)
CierraConexionFox
End Sub





Private Sub TxtBuscarAge_EmiteDatos()
lblDescAge = TxtBuscarAge.psDescripcion
If TxtBuscarAge <> "" And lblDescAge <> "" Then
    txtBuscarUser = ""
    txtBuscarUser.psRaiz = "USUARIOS " & TxtBuscarAge.psDescripcion
    txtBuscarUser.Enabled = True
    txtBuscarUser.rs = oGen.GetUserAreaAgencia("026", TxtBuscarAge)
End If
End Sub

Private Sub TxtBuscarUser_EmiteDatos()
If Me.txtBuscarUser.Text <> "" Then
    Me.lblUsuSIAFC = GetUserSiac(Me.txtBuscarUser)
End If
End Sub
Function ReportePagos(ByVal rs As ADODB.Recordset, ByVal pdFecha As Date, ByVal psCodUsu As String, ByVal psMoneda As String, ByVal lsInfoSicmac_Siafc As String, lnPagina As Integer) As String
Dim lscadena As String
Dim lnPag As Integer
Dim lnTotaCapSICMAC As Currency
Dim lnTotaCapSIAFC As Currency

Dim lnTotIntSICMAC As Currency
Dim lnTotIntSIAFC As Currency

Dim lnTotaMoraSICMAC As Currency
Dim lnTotaMoraSIAFC As Currency

Dim lnTotaITFSICMAC As Currency
Dim lnTotaITFSIAFC As Currency

Dim lnTotalCredSicmac As Long
Dim lnTotalCredSiafc As Long
Dim lsCodInst As String

lnPag = 0
If Not rs.EOF And Not rs.BOF Then
lscadena = CabeRepo(gsNomCmac, lblDescAge, "", IIf(psMoneda = "1", "SOLES", "DOLARES"), Format(pdFecha, "dd/mm/yyyy"), "Control Calidad PAGOS DE CREDITOS", "USUARIO :" & psCodUsu, lsInfoSicmac_Siafc, "", lnPag, 64)
lscadena = lscadena + Chr(10)
lnTotalProd = 0
lnLinea = 11
lsProducto = ""
lscadena = lscadena & String(100, "-") & Chr(10)
                        
lscadena = lscadena & ImpreFormat("Sistema", 6) & _
                       ImpreFormat("N° Credito", 18) & _
                       ImpreFormat("Nombre Cliente", 16) & _
                       ImpreFormat("Est", 3) & _
                       ImpreFormat("RFA", 3) & _
                       ImpreFormat("USER", 4) & _
                       ImpreFormat("SaldoCap", 10) & _
                       ImpreFormat("CAPITAL", 11) & _
                       ImpreFormat("INTERES", 8) & _
                       ImpreFormat("MORA", 6) & _
                       ImpreFormat("DESAGIO", 8) & _
                       ImpreFormat("PROTESTO", 8) & _
                       ImpreFormat("COMISION", 6) & _
                       ImpreFormat("TOTAL", 8) & _
                       ImpreFormat("I.T.F.", 8) & _
                       ImpreFormat("CAJA", 4) & Chr(10)

lscadena = lscadena & String(100, "-") & Chr(10)

lnTotaCapSICMAC = 0
lnTotaCapSIAFC = 0

lnTotIntSICMAC = 0
lnTotIntSIAFC = 0

lnTotaMoraSICMAC = 0
lnTotaMoraSIAFC = 0

lnTotaITFSICMAC = 0
lnTotaITFSIAFC = 0


rs.MoveFirst
lsCodInst = ""
Do While Not rs.EOF
    If Me.chkInst.value = 1 Then
        If lsCodInst <> rs!cCodIns Then
            lscadena = lscadena & "INSTITUCION :" & Trim(rs!cInst) & Chr(10)
            lsCodInst = rs!cCodIns
        End If
    End If
    lnLinea = lnLinea + 1
    lscadena = lscadena & ImpreFormat(rs!Sistema, 6) & _
                       ImpreFormat(rs!cCtaCod, 18) & _
                       ImpreFormat(rs!CPERSNOMBRE, 16) & _
                       ImpreFormat(rs!cEstado, 3) & _
                       ImpreFormat(rs!CRFA, 3) & _
                       ImpreFormat(rs!Cuser, 4) & _
                       ImpreFormat(rs!nSaldo, 8) & _
                       ImpreFormat(rs!capital, 8) & _
                       ImpreFormat(rs!Interes, 8) & _
                       ImpreFormat(rs!MORA, 6) & _
                       ImpreFormat(rs!DESAGIO, 6) & _
                       ImpreFormat(rs!PROTESTO, 6) & _
                       ImpreFormat(rs!COMISION, 6) & _
                       ImpreFormat(rs!capital + rs!Interes + rs!MORA + rs!DESAGIO + rs!PROTESTO + rs!COMISION, 6) & _
                       ImpreFormat(rs!ITF, 6) & _
                       ImpreFormat(rs!CAJA, 8) & Chr(10)

'ImpreFormat(rs!otros, 6) & _

    lnTotalCaso = lnTotalCaso + 1
    If Trim(rs!Sistema) = "SIAFC" Then
        lnTotaCapSIAFC = lnTotaCapSIAFC + rs!capital
        lnTotIntSIAFC = lnTotIntSIAFC + rs!Interes
        lnTotaMoraSIAFC = lnTotaMoraSIAFC + rs!MORA
        lnTotaITFSIAFC = lnTotaITFSIAFC + rs!ITF
        lnTotalCredSiafc = lnTotalCredSiafc + 1
    Else
        lnTotaCapSICMAC = lnTotaCapSICMAC + rs!capital
        lnTotIntSICMAC = lnTotIntSICMAC + rs!Interes
        lnTotaITFSICMAC = lnTotaITFSICMAC + rs!ITF
        lnTotaMoraSICMAC = lnTotaMoraSICMAC + rs!MORA
        lnTotalCredSicmac = lnTotalCredSicmac + 1
    End If
    If lnLinea > 61 Then
        lnLinea = 11
        lnPag = lnPag + 1
        lscadena = lscadena & CabeRepo(gsNomCmac, lblDescAge, "", IIf(psMoneda = "1", "SOLES", "DOLARES"), Format(pdFecha, "dd/mm/yyyy"), "Control Calidad PAGOS DE CREDITOS", "USUARIO :" & psCodUsu, lsInfoSicmac_Siafc, "", lnPag, 64)
        lscadena = lscadena & String(100, "-") & Chr(10)

        lscadena = lscadena & ImpreFormat("Sistema", 6) & _
                       ImpreFormat("N° Credito", 18) & _
                       ImpreFormat("Nombre Cliente", 16) & _
                       ImpreFormat("Est", 3) & _
                       ImpreFormat("RFA", 3) & _
                       ImpreFormat("User", 4) & _
                       ImpreFormat("SaldoCap", 10) & _
                       ImpreFormat("CAPITAL", 11) & _
                       ImpreFormat("INTERES", 8) & _
                       ImpreFormat("MORA", 6) & _
                       ImpreFormat("DESAGIO", 8) & _
                       ImpreFormat("PROTESTO", 8) & _
                       ImpreFormat("COMISION", 6) & _
                       ImpreFormat("TOTAL", 8) & _
                       ImpreFormat("I.T.F.", 8) & _
                       ImpreFormat("CAJA", 4) & Chr(10)

        lscadena = lscadena & String(100, "-") & Chr(10)

    End If
    rs.MoveNext
Loop
End If
lscadena = lscadena & String(100, "-") & Chr(10)

lscadena = lscadena & ImpreFormat("TOTALES :", 6) & _
                        ImpreFormat("CAP.SIAFC:", 10) & _
                        ImpreFormat(lnTotaCapSIAFC, 11, 2) & _
                        ImpreFormat("CAP.SICMAC:", 10) & _
                        ImpreFormat(lnTotaCapSICMAC, 11, 2) & _
                        ImpreFormat("INT.SIAFC:", 10) & _
                        ImpreFormat(lnTotIntSIAFC, 11, 2) & _
                        ImpreFormat("INT.SICMAC:", 10) & _
                        ImpreFormat(lnTotIntSICMAC, 11, 2) & Chr(10) & ImpreFormat("", 6) & _
                        ImpreFormat("MORA.SIAFC:", 10) & _
                        ImpreFormat(lnTotaMoraSIAFC, 11, 2) & _
                        ImpreFormat("MORA.SICMAC:", 10) & _
                        ImpreFormat(lnTotaMoraSICMAC, 11, 2) & _
                        ImpreFormat("ITF.SIAFC:", 10) & _
                        ImpreFormat(lnTotaITFSIAFC, 11, 2) & _
                        ImpreFormat("ITF.SICMAC:", 10) & _
                        ImpreFormat(lnTotaITFSICMAC, 11, 2) & Chr(10)

lscadena = lscadena & ImpreFormat("TOTALES CASOS:", 20) & _
                        ImpreFormat("SIAFC:", 10) & _
                        ImpreFormat(lnTotalCredSiafc, 11, 2) & _
                        ImpreFormat("SICMAC-I:", 10) & _
                        ImpreFormat(lnTotalCredSicmac, 11, 2) & Chr(10)


rs.Close
Set rs = Nothing
ReportePagos = lscadena

End Function

Sub ProcesaPagosSICMAC(ByVal pdFecha As Date, ByVal psMoneda As String, ByVal psCodUsu As String, ByVal psUsuSICMAC As String, Optional lbRfa As Boolean = False, Optional lbJudicial As Boolean = False)
Dim sql As String
Dim oCon As COMConecta.DCOMConecta
Dim lsCondJud As String
Dim lsCondRfa  As String
Dim lsTitulo As String
Dim lsFiltroUser As String

lsTmpPagos = "TMPREPPAGOS" & psCodUsu

If lbRfa Then
    lsCondRfa = " and C.cRFA IN ('RFA','RFC','DIF') "
    lsTitulo = "REPORTE VALIDACION PAGOS SIAFC - SICMACI (CRED.RFA)"
Else
    lsCondRfa = " and (C.cRFA IS NULL OR CRFA ='NOR') "
    lsTitulo = "REPORTE VALIDACION PAGOS SIAFC-SICMACI(CRED.NORMALES)"
End If
If lbJudicial Then
    lsCondJud = " and M.COPECOD like '13%' "
    lsTitulo = "REPORTE VALIDACION PAGOS SIAFC-SICMACI (CRED.JUD)"
Else
    lsCondJud = " and NOT M.COPECOD like '13%' "
    lsTitulo = "REPORTE VALIDACION PAGOS SIAFC-SICMACI(CRED.NORMALES)"
End If

If psCodUsu = "" Then
    lsFiltroUser = ""
Else
    lsFiltroUser = " AND RIGHT(CMOVNRO,4)='" & psUsuSICMAC & "' "
End If

Set oCon = New COMConecta.DCOMConecta
oCon.AbreConexion

Dim rs As ADODB.Recordset
Set rs = oCon.CargaRecordSet("select * from sysobjects where name like '%" & lsTmpPagos & "%'")
If Not rs.EOF And Not rs.BOF Then
    sql = "DROP TABLE " & lsTmpPagos
    oCon.Ejecutar sql
End If
rs.Close
Set rs = Nothing

sql = "SELECT   'SICMAC-I' as SISTEMA,PAGOS.CCTACOD,  P.NSALDO, c.cConsDescripcion AS cEstado, P1.CPERSNOMBRE ," _
    & "         PAGOS.CRFA, ISNULL(R.CCTACODANT,PAGOS.CCTACOD) AS CCTACODANT , PAGOS.CUSER, " _
    & "         SUM(PAGOS.CAPITAL) AS CAPITAL, " _
    & "         SUM(PAGOS.INTERES) AS INTERES, " _
    & "         SUM(PAGOS.MORA) AS MORA, " _
    & "         SUM(PAGOS.DESAGIO) AS DESAGIO, " _
    & "         SUM(PAGOS.ITF) AS ITF, " _
    & "         SUM(PAGOS.PROTESTO) AS PROTESTO, " _
    & "         SUM(PAGOS.COMISION) AS COMISION, " _
    & "         SUM(PAGOS.OTROS) AS OTROS, " _
    & "         SUM(PAGOS.nMonto) As CAJA, SPACE(3) AS cCodIns, space(40) as cInst " _
    & " INTO    " & lsTmpPagos _
    & " FROM    (SELECT M.NMOVNRO,MD.CCTACOD, MC.nMonto , isnull(C.CRFA,space(3)) as CRFA , RIGHT(M.CMOVNRO,4) AS CUSER,  " _
    & "                 SUM(CASE WHEN nPrdConceptoCod in (1000,124353,3000,1109,1010,1110) THEN MD.NMONTO ELSE 0 END) AS CAPITAL, " _
    & "                 SUM(CASE WHEN nPrdConceptoCod in (1100,3100) THEN MD.NMONTO ELSE 0 END) AS INTERES, " _
    & "                 SUM(CASE WHEN nPrdConceptoCod in (1101,3101) THEN MD.NMONTO ELSE 0 END) AS MORA, " _
    & "                 SUM(CASE WHEN nPrdConceptoCod=1106 THEN MD.NMONTO ELSE 0 END) AS DESAGIO, " _
    & "                 SUM(CASE WHEN nPrdConceptoCod IN (20,21) THEN MD.NMONTO ELSE 0 END) AS ITF, " _
    & "                 SUM(CASE WHEN nPrdConceptoCod IN (1219,1220) THEN MD.NMONTO ELSE 0 END) AS PROTESTO, " _
    & "                 SUM(CASE WHEN nPrdConceptoCod IN (124350) THEN MD.NMONTO ELSE 0 END) AS COMISION, " _
    & "                 SUM(CASE WHEN nPrdConceptoCod NOT IN (1000,1100,1101,20,21,1219,1220,124350,1106,124353,3000,3100,3101) THEN MD.NMONTO ELSE 0 END) AS OTROS " _
    & "         FROM    MOV M " _
    & "                 JOIN MOVCOL     MC ON MC.NMOVNRO = M.NMOVNRO " _
    & "                 JOIN MOVCOLDET  MD ON MD.NMOVNRO = MC.NMOVNRO AND MD.CCTACOD = MC.CCTACOD AND MD.COPECOD = MC.COPECOD " _
    & "                 JOIN COLOCACCRED C ON C.CCTACOD = MC.CCTACOD "
sql = sql & "  WHERE    LEFT(M.CMOVNRO,8) ='" & Format(pdFecha, "yyyymmdd") & "' " & lsFiltroUser & " AND SUBSTRING(MC.CCTACOD,9,1)='" & psMoneda & "' " _
    & "                 AND (NOT M.COPECOD IN ('100101','100102','100103','100104') AND NOT MD.COPECOD LIKE '107[1,2,3,4]%' AND MD.COPECOD<>'' " & lsCondJud & " ) AND M.nMovFlag ='0' " _
    & "                 AND SUBSTRING(MC.CCTACOD,6,3)<>'305' AND SUBSTRING(MC.CCTACOD,4,2)='" & Trim(TxtBuscarAge) & "'  " & lsCondRfa _
    & "          GROUP BY M.NMOVNRO,MD.CCTACOD, MC.nMonto,C.cRFA, RIGHT(M.CMOVNRO,4) ) AS PAGOS " _
    & "             JOIN PRODUCTO P ON P.CCTACOD= PAGOS.CCTACOD " _
    & "             JOIN PRODUCTOPERSONA PP ON PP.CCTACOD = P.CCTACOD AND PP.nPrdPersRelac= 20 " _
    & "             JOIN PERSONA P1 ON P1.CPERSCOD = PP.CPERSCOD " _
    & "             LEFT JOIN RELCUENTAS R ON R.CCTACOD = P.CCTACOD " _
    & "             JOIN CONSTANTE C ON C.nConsValor = P.NPRDESTADO AND C.nConsCod =3001 " _
    & " GROUP BY PAGOS.CCTACOD,  P.NSALDO, P.NPRDESTADO, P1.CPERSNOMBRE,PAGOS.CRFA,c.cConsDescripcion, R.CCTACODANT,PAGOS.CUSER " _
    & " ORDER BY P1.CPERSNOMBRE "

oCon.Ejecutar sql

' pagos de fox
Set rs = CredPagosSIAFC(pdFecha, psMoneda, psCodUsu, lbRfa, lbJudicial)
Do While Not rs.EOF
    
    sql = "INSERT INTO " & lsTmpPagos & " (SISTEMA,CCTACOD, NSALDO , CESTADO, CPERSNOMBRE, CRFA,  " _
        & " CCTACODANT, CAPITAL, INTERES, MORA, DESAGIO, ITF, PROTESTO, COMISION, OTROS, CAJA, cCodIns, cInst, cUser) " _
        & " VALUES('" & rs!cSistema & "','" & rs!cCtaCod & "'," & rs!nSaldo & ",'" & Trim(rs!cEstado) & "','" _
        & Trim(Replace(rs!CPERSNOMBRE, "'", "''")) & "'," & IIf(Trim(rs!CRFA) = "", " space(3)", "'" & Trim(rs!CRFA) & "'") & ",'" & rs!CCTACODANT & "'," & rs!capital & "," _
        & rs!Interes & "," & rs!MORA & "," & rs!DESAGIO & "," & rs!ITF & "," & rs!PROTESTO & "," _
        & rs!nComision & "," & rs!nOtros & "," & rs!CAJA & ",'" & rs!cCodIns & "','" & rs!cInst & "','" & Trim(rs!Cuser) & "')"
    
    oCon.Ejecutar sql
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing
If Me.chkInst.value = 0 Then
sql = "select * from " & lsTmpPagos _
 & " order by CCTACODANT,CPERSNOMBRE, SISTEMA "
Else
sql = "select * from " & lsTmpPagos _
 & " order by cCodIns, CCTACODANT, CPERSNOMBRE, SISTEMA "
End If
Set rs = oCon.CargaRecordSet(sql)
Dim lsCad As String

lsCad = ReportePagos(rs, pdFecha, psCodUsu, psMoneda, lsTitulo, 0)
'oPrev.Show lsCad, "", True
oPrev.Show lsCad, "", True, , gImpresora
oCon.CierraConexion
End Sub
Function GetDesembolsoPrendarioSiafc(ByVal pdFecha As Date, ByVal psMoneda As String, ByVal psCodUsu As String) As ADODB.Recordset
Dim sql As String
DeterminaRutas

sql = "select   kd.CCODCTA, C.CNOMCLI, K.CESTADO, K.NMSALDO AS NSALDO, " _
    & "         SUM(IIF(kd.cCodTrx ='002' , kd.nMonCre , 00000.00 )) as Capital,  " _
    & "         SUM(IIF(kd.cCodTrx ='002' , kd.nIntere , 00000.00 )) as Interes,  " _
    & "         SUM(IIF(kd.cCodTrx <>'002', kd.nMonTrx , 00000.00 )) as Itf, sum(iif(cCodTrx <>'002', nMonTrx*-1 , nMonTrx)) as CAJA  " _
    & " from    " & gsBaseKPR & "kprdtrx kd " _
    & "         join " & gsBaseKPR & "kprmpre K ON K.CCODCTA = KD.CCODCTA " _
    & "         JOIN " & gsBaseCli & "CLIMIDE C ON C.CCODCLI = K.CCODCLI  " _
    & " where kd.dfecpro = CTOD('" & Format(pdFecha, "mm/dd/yyyy") & "') and kd.cEstado ='D' and kd.cMoneda ='" & psMoneda & "' AND kd.CCODUSU ='" & psCodUsu & "' " _
    & " GROUP BY KD.cCodCta ORDER BY C.CNOMCLI"

Set GetDesembolsoPrendarioSiafc = CargaRecordFox(sql)

End Function

Function GetDesembolsoPrendarioSICMAC(ByVal pdFecha As Date, ByVal psMoneda As String, ByVal psCodUsu As String) As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim oCon As COMConecta.DCOMConecta

sql = "SELECT   PAGOS.CCTACOD AS CCODCTA, P.NSALDO, c.cConsDescripcion AS CESTADO, P1.CPERSNOMBRE AS CNOMCLI," _
    & "         ISNULL(R.CCTACODANT,SPACE(18)) AS CCTACODANT , " _
    & "         SUM(PAGOS.CAPITAL +PAGOS.INTERES + PAGOS.ITF) AS CAPITAL, " _
    & "         SUM(PAGOS.INTERES) AS INTERES, " _
    & "         SUM(PAGOS.MORA) AS MORA, " _
    & "         SUM(PAGOS.ITF) AS ITF, " _
    & "         SUM(PAGOS.OTROS) AS OTROS, " _
    & "         SUM(PAGOS.nMonto - PAGOS.ITF) As CAJA " _
    & " FROM    (SELECT MD.CCTACOD, MC.nMonto, " _
    & "                 SUM(CASE WHEN nPrdConceptoCod IN (2000,2001,8103,8000,2271) THEN MD.NMONTO ELSE 0 END) AS CAPITAL," _
    & "                 SUM(CASE WHEN nPrdConceptoCod IN (2100,8104,8100) THEN MD.NMONTO ELSE 0 END) AS INTERES, " _
    & "                 SUM(CASE WHEN nPrdConceptoCod IN (2101,8101) THEN MD.NMONTO ELSE 0 END) AS MORA, " _
    & "                 SUM(CASE WHEN nPrdConceptoCod IN (20,21,22) THEN MD.NMONTO ELSE 0 END) AS ITF, " _
    & "                 SUM(CASE WHEN nPrdConceptoCod NOT IN (2000,2001,8103,8000,2100,8104,8100,2101,8101,2271,20,21,22) THEN MD.NMONTO ELSE 0 END) AS OTROS " _
    & "         FROM    MOV M " _
    & "                 JOIN MOVCOL     MC ON MC.NMOVNRO = M.NMOVNRO " _
    & "                 JOIN MOVCOLDET  MD ON MD.NMOVNRO = MC.NMOVNRO AND MD.CCTACOD = MC.CCTACOD AND MD.COPECOD = MC.COPECOD " _
    & "         WHERE   LEFT(M.CMOVNRO,8) ='" & Format(pdFecha, "yyyymmdd") & "' AND RIGHT(CMOVNRO,4)='" & psCodUsu & "' AND SUBSTRING(MC.CCTACOD,9,1)='" & psMoneda & "'  " _
    & "                 AND M.COPECOD ='120201' AND M.nMovFlag ='0' AND SUBSTRING(MC.CCTACOD,6,3)='305' " _
    & "         GROUP BY MD.CCTACOD, MC.nMonto) AS PAGOS " _
    & "         JOIN PRODUCTO P ON P.CCTACOD= PAGOS.CCTACOD " _
    & "         JOIN PRODUCTOPERSONA PP ON PP.CCTACOD = P.CCTACOD AND PP.nPrdPersRelac= 20 " _
    & "         JOIN PERSONA P1 ON P1.CPERSCOD = PP.CPERSCOD " _
    & "         LEFT JOIN RELCUENTAS R ON R.CCTACOD = P.CCTACOD " _
    & "         JOIN CONSTANTE C ON C.nConsValor = P.NPRDESTADO AND C.nConsCod =3001 "
sql = sql & " GROUP BY PAGOS.CCTACOD,  P.NSALDO, P.NPRDESTADO, P1.CPERSNOMBRE,c.cConsDescripcion, R.CCTACODANT " _
    & " ORDER BY P1.CPERSNOMBRE"

Set oCon = New COMConecta.DCOMConecta
oCon.AbreConexion
Set GetDesembolsoPrendarioSICMAC = oCon.CargaRecordSet(sql)
oCon.CierraConexion
End Function
Function ReporteDesemPrendario(ByVal rs As ADODB.Recordset, ByVal pdFecha As Date, ByVal psCodUsu As String, ByVal psMoneda As String, ByVal lsInfoSicmac_Siafc As String, lnPagina As Integer) As String
Dim lnTotalProdCap As Currency
Dim lnTotalProdITF As Currency
Dim lnTotalProdCaj As Currency
Dim lscadena As String
Dim lnTotalCaso  As Long

If Not rs.EOF And Not rs.BOF Then
lscadena = CabeRepo(gsNomCmac, lblDescAge, "", IIf(psMoneda = "1", "SOLES", "DOLARES"), Format(pdFecha, "dd/mm/yyyy"), "Control Calidad Desembolsos PRENDARIO", "USUARIO :" & psCodUsu, lsInfoSicmac_Siafc, "", lnPagina, 64)
lscadena = lscadena + Chr(10)
lnTotalProd = 0
lnLinea = 0
lsProducto = ""
lscadena = lscadena & String(100, "-") & Chr(10)
lscadena = lscadena & ImpreFormat("N° Credito", 18) & _
                       ImpreFormat("Nombre Cliente", 48) & _
                       ImpreFormat("SALDO", 12, 2) & _
                       ImpreFormat("CAPITAL", 12, 2) & _
                       ImpreFormat("INTERES", 12, 2) & _
                       ImpreFormat("I.T.F.", 12, 2) & _
                       ImpreFormat("CAJA", 12, 2) & Chr(10)

lscadena = lscadena & String(100, "-") & Chr(10)
lnTotalProdCap = 0
lnTotalProdITF = 0
lnTotalProdCaj = 0
lnTotalCaso = 0
Do While Not rs.EOF
    lnLinea = lnLinea + 1
    lscadena = lscadena & ImpreFormat(rs!cCodCta, 18) & _
                        ImpreFormat(Trim(rs!cNomCli), 40) & _
                        ImpreFormat(rs!nSaldo, 12, 2) & _
                        ImpreFormat(rs!capital, 12, 2) & _
                        ImpreFormat(rs!Interes, 12, 2) & _
                       ImpreFormat(rs!ITF, 12, 2) & _
                       ImpreFormat(rs!CAJA, 12, 2) & Chr(10)
    
    lnTotalProdCap = lnTotalProdCap + rs!capital
    lnTotalProdITF = lnTotalProdITF + rs!ITF
    lnTotalProdCaj = lnTotalProdCaj + rs!CAJA
    lnTotalCaso = lnTotalCaso + 1
    
    If lnLinea > 63 Then
        lnLinea = 1
        lscadena = lscadena & Chr(12)
        lscadena = lscadena & ImpreFormat("N° Credito", 18) & _
                       ImpreFormat("Nombre Cliente", 40) & _
                       ImpreFormat("SALDO", 12, 2) & _
                       ImpreFormat("CAPITAL", 12, 2) & _
                       ImpreFormat("INTERES", 12, 2) & _
                       ImpreFormat("I.T.F.", 12, 2) & _
                       ImpreFormat("CAJA", 12, 2) & Chr(10)
        
    End If
    rs.MoveNext
Loop
End If

rs.Close
Set rs = Nothing
ReporteDesemPrendario = lscadena
End Function
Sub ValidaDesembolsosPren(ByVal pdFecha As Date, ByVal psCodUsu As String, ByVal psMoneda As String, ByVal psUsuSICMAC As String)
Dim rs As ADODB.Recordset
Dim lsCadena1 As String
Set rs = GetDesembolsoPrendarioSiafc(pdFecha, psMoneda, psCodUsu)

lsCadena1 = ReporteDesemPrendario(rs, pdFecha, psCodUsu, psMoneda, "INFORMACION SIAFC", 0)
Set rs = GetDesembolsoPrendarioSICMAC(pdFecha, psMoneda, psUsuSICMAC)
lsCadena1 = lsCadena1 + ReporteDesemPrendario(rs, pdFecha, psUsuSICMAC, psMoneda, "INFORMACION SICMAC-I", 1)

'oPrev.Show lsCadena1, "", True
oPrev.Show lsCadena1, "", True, , gImpresora
End Sub
Function GetPagosPrendarioSIAFC(ByVal pdFecha As Date, ByVal psMoneda As String, ByVal psCodUsu As String) As ADODB.Recordset
Dim sql As String

sql = "   select    kd.CCODCTA, C.CNOMCLI, K.CESTADO, K.NMSALDO AS NSALDO,  " _
    & "         SUM(IIF(kd.cCodTrx <>'080' , kd.nMonTrx - kd.nMonMor - KD.NiNTERE , 00000.00 )) as Capital,  " _
    & "         SUM(IIF(kd.cCodTrx <>'080' , kd.nIntere , 00000.00 )) as Interes,  " _
    & "         SUM(IIF(kd.cCodTrx <>'080' , kd.nMonMor , 00000.00 )) as Mora, " _
    & "         SUM(IIF(kd.cCodTrx ='080', kd.nMonTrx , 00000.00 )) as Itf,  " _
    & "         sum(iif(cCodTrx ='080', nMonTrx, nMonTrx)) as CAJA, KD.CCODUSU  " _
    & "     from    " & gsBaseKPR & "kprdtrx kd " _
    & "             join " & gsBaseKPR & "kprmpre K ON K.CCODCTA = KD.CCODCTA " _
    & "             JOIN " & gsBaseCli & "CLIMIDE C ON C.CCODCLI = K.CCODCLI  " _
    & "     where KD.dfecpro = CTOD('" & Format(pdFecha, "mm/dd/yyyy") & "') and kd.cEstado $'R,A,C' and kd.cMoneda ='" & psMoneda & "'  and KD.cCodUsu='" & psCodUsu & "' " _
    & " GROUP BY KD.cCodCta " _
    & " ORDER BY C.CNOMCLI "

Set GetPagosPrendarioSIAFC = CargaRecordFox(sql)

End Function
Sub ProcesaPagosPRENDSICMAC(ByVal pdFecha As Date, ByVal psMoneda As String, ByVal psCodUsu As String, ByVal psUsuSICMAC As String)
Dim sql As String
Dim oCon As COMConecta.DCOMConecta

lsTmpPagos = "TMPPAGOPREND" & psCodUsu

Set oCon = New COMConecta.DCOMConecta
oCon.AbreConexion

Dim rs As ADODB.Recordset
Set rs = oCon.CargaRecordSet("select * from sysobjects where name like '%" & lsTmpPagos & "%'")
If Not rs.EOF And Not rs.BOF Then
    sql = "DROP TABLE " & lsTmpPagos
    oCon.Ejecutar sql
End If
rs.Close
Set rs = Nothing


sql = " SELECT  'SICMAC-I' as SISTEMA,PAGOS.CCTACOD, P.NSALDO, c.cConsDescripcion AS CESTADO, P1.CPERSNOMBRE, " _
    & "         ISNULL(R.CCTACODANT,SPACE(18)) AS CCTACODANT , " _
    & "         SUM(PAGOS.CAPITAL) AS CAPITAL, " _
    & "         SUM(PAGOS.INTERES) AS INTERES, " _
    & "         SUM(PAGOS.MORA) AS MORA, " _
    & "         SUM(PAGOS.ITF) AS ITF, " _
    & "         SUM(PAGOS.OTROS) AS OTROS, " _
    & "         SUM(PAGOS.nMonto) As CAJA " _
    & "         INTO " & lsTmpPagos _
    & "         FROM    (SELECT MD.CCTACOD, MC.nMonto, " _
    & "                 SUM(CASE WHEN nPrdConceptoCod IN (2000,2001,8103,8000) THEN MD.NMONTO ELSE 0 END) AS CAPITAL, " _
    & "                 SUM(CASE WHEN nPrdConceptoCod IN (2100,8104,8100) THEN MD.NMONTO ELSE 0 END) AS INTERES, " _
    & "                 SUM(CASE WHEN nPrdConceptoCod IN (2101,8101) THEN MD.NMONTO ELSE 0 END) AS MORA, " _
    & "                 SUM(CASE WHEN nPrdConceptoCod IN (20,21,22) THEN MD.NMONTO ELSE 0 END) AS ITF, " _
    & "                 SUM(CASE WHEN nPrdConceptoCod NOT IN (2000,2001,8103,8000,2100,8104,8100,2101,8101,2271,20,21,22) THEN MD.NMONTO ELSE 0 END) AS OTROS " _
    & "                 FROM    MOV M " _
    & "                         JOIN MOVCOL     MC ON MC.NMOVNRO = M.NMOVNRO " _
    & "                         JOIN MOVCOLDET  MD ON MD.NMOVNRO = MC.NMOVNRO AND MD.CCTACOD = MC.CCTACOD AND MD.COPECOD = MC.COPECOD " _
    & "                 WHERE   LEFT(M.CMOVNRO,8) ='" & Format(pdFecha, "yyyymmdd") & "' AND RIGHT(CMOVNRO,4)='" & Trim(psUsuSICMAC) & "' AND SUBSTRING(MC.CCTACOD,9,1)='" & psMoneda & "' " _
    & "                         AND M.COPECOD like '12[16][012]%'  AND M.nMovFlag ='0' " _
    & "                         AND SUBSTRING(MC.CCTACOD,6,3)='305' " _
    & "                 GROUP BY MD.CCTACOD, MC.nMonto) AS PAGOS " _
    & "         JOIN PRODUCTO P ON P.CCTACOD= PAGOS.CCTACOD " _
    & "         JOIN PRODUCTOPERSONA PP ON PP.CCTACOD = P.CCTACOD AND PP.nPrdPersRelac= 20 " _
    & "     JOIN PERSONA P1 ON P1.CPERSCOD = PP.CPERSCOD "
sql = sql + "     LEFT JOIN RELCUENTAS R ON R.CCTACOD = P.CCTACOD " _
    & "JOIN CONSTANTE C ON C.nConsValor = P.NPRDESTADO AND C.nConsCod =3001 " _
    & "     GROUP BY PAGOS.CCTACOD,  P.NSALDO, P.NPRDESTADO, P1.CPERSNOMBRE,c.cConsDescripcion, R.CCTACODANT " _
    & " ORDER BY P1.CPERSNOMBRE "

oCon.Ejecutar sql

' pagos de fox
Set rs = GetPagosPrendarioSIAFC(pdFecha, psMoneda, psCodUsu)
Do While Not rs.EOF
    sql = "INSERT INTO " & lsTmpPagos & " (SISTEMA, CCTACOD, NSALDO, CESTADO, CPERSNOMBRE, CCTACODANT, CAPITAL, INTERES, MORA,ITF, OTROS, CAJA) " _
        & " VALUES('SIAFC','" & rs!cCodCta & "'," & rs!nSaldo & ",'" & Trim(rs!cEstado) & "','" _
        & Trim(rs!cNomCli) & "','" & rs!cCodCta & "'," & rs!capital & "," _
        & rs!Interes & "," & rs!MORA & "," & rs!ITF & ",0," & rs!CAJA & ")"
    
    oCon.Ejecutar sql
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing

sql = "select * from " & lsTmpPagos _
 & " order by CCTACODANT,CPERSNOMBRE, SISTEMA "
Set rs = oCon.CargaRecordSet(sql)

Dim lsCad As String
Dim lsTitulo As String
lsTitulo = "REPORTE DE VALIDACION DE PAGOS SIAFC-SICMACI(PRENDARIO)"
lsCad = ReportePagosPREND(rs, pdFecha, psCodUsu, psMoneda, lsTitulo, 0)
'oPrev.Show lsCad, "", True
oPrev.Show lsCad, "", True, , gImpresora
oCon.CierraConexion


End Sub
Function ReportePagosPREND(ByVal rs As ADODB.Recordset, ByVal pdFecha As Date, ByVal psCodUsu As String, ByVal psMoneda As String, ByVal lsInfoSicmac_Siafc As String, lnPagina As Integer) As String
Dim lscadena As String
Dim lnPag As Integer
Dim lnTotaCapSICMAC As Currency
Dim lnTotaCapSIAFC As Currency

Dim lnTotIntSICMAC As Currency
Dim lnTotIntSIAFC As Currency

Dim lnTotaMoraSICMAC As Currency
Dim lnTotaMoraSIAFC As Currency

Dim lnTotaITFSICMAC As Currency
Dim lnTotaITFSIAFC As Currency

Dim lnTotalCredSicmac As Long
Dim lnTotalCredSiafc As Long


lnPag = 0
If Not rs.EOF And Not rs.BOF Then
lscadena = CabeRepo(gsNomCmac, lblDescAge, "", IIf(psMoneda = "1", "SOLES", "DOLARES"), Format(pdFecha, "dd/mm/yyyy"), "Control Calidad PAGOS CREDITO PRENDARIO", "USUARIO :" & psCodUsu, lsInfoSicmac_Siafc, "", lnPag, 64)
lscadena = lscadena + Chr(10)
lnTotalProd = 0
lnLinea = 11
lsProducto = ""
lscadena = lscadena & String(100, "-") & Chr(10)
                        
lscadena = lscadena & ImpreFormat("Sistema", 6) & _
                       ImpreFormat("N° Credito", 18) & _
                       ImpreFormat("Nombre Cliente", 20) & _
                       ImpreFormat("Est", 3) & _
                       ImpreFormat("SaldoCap", 10) & _
                       ImpreFormat("CAPITAL", 11) & _
                       ImpreFormat("INTERES", 8) & _
                       ImpreFormat("MORA", 6) & _
                       ImpreFormat("I.T.F.", 8) & _
                       ImpreFormat("OTROS", 8) & _
                       ImpreFormat("CAJA", 4) & Chr(10)

lscadena = lscadena & String(100, "-") & Chr(10)

lnTotaCapSICMAC = 0
lnTotaCapSIAFC = 0

lnTotIntSICMAC = 0
lnTotIntSIAFC = 0

lnTotaMoraSICMAC = 0
lnTotaMoraSIAFC = 0

lnTotaITFSICMAC = 0
lnTotaITFSIAFC = 0


rs.MoveFirst
Do While Not rs.EOF
    lnLinea = lnLinea + 1
    lscadena = lscadena & ImpreFormat(rs!Sistema, 6) & _
                       ImpreFormat(rs!cCtaCod, 18) & _
                       ImpreFormat(rs!CPERSNOMBRE, 20) & _
                       ImpreFormat(rs!cEstado, 3) & _
                       ImpreFormat(rs!nSaldo, 8) & _
                       ImpreFormat(rs!capital, 8) & _
                       ImpreFormat(rs!Interes, 8) & _
                       ImpreFormat(rs!MORA, 6) & _
                       ImpreFormat(rs!ITF, 6) & _
                       ImpreFormat(rs!otros, 6) & _
                       ImpreFormat(rs!CAJA, 8) & Chr(10)

    lnTotalCaso = lnTotalCaso + 1
    If Trim(rs!Sistema) = "SIAFC" Then
        lnTotaCapSIAFC = lnTotaCapSIAFC + rs!capital
        lnTotIntSIAFC = lnTotIntSIAFC + rs!Interes
        lnTotaMoraSIAFC = lnTotaMoraSIAFC + rs!MORA
        lnTotaITFSIAFC = lnTotaITFSIAFC + rs!ITF
        lnTotalCredSiafc = lnTotalCredSiafc + 1
    Else
        lnTotaCapSICMAC = lnTotaCapSICMAC + rs!capital
        lnTotIntSICMAC = lnTotIntSICMAC + rs!Interes
        lnTotaITFSICMAC = lnTotaITFSICMAC + rs!ITF
        lnTotaMoraSICMAC = lnTotaMoraSICMAC + rs!MORA
        lnTotalCredSicmac = lnTotalCredSicmac + 1
    End If
    If lnLinea > 61 Then
        lnLinea = 11
        lnPag = lnPag + 1
        lscadena = lscadena & CabeRepo(gsNomCmac, lblDescAge, "", IIf(psMoneda = "1", "SOLES", "DOLARES"), Format(pdFecha, "dd/mm/yyyy"), "Control Calidad PAGOS DE CREDITOS", "USUARIO :" & psCodUsu, lsInfoSicmac_Siafc, "", lnPag, 64)
        lscadena = lscadena & String(100, "-") & Chr(10)

        lscadena = lscadena & ImpreFormat("Sistema", 6) & _
                       ImpreFormat("N° Credito", 18) & _
                       ImpreFormat("Nombre Cliente", 20) & _
                       ImpreFormat("Est", 3) & _
                       ImpreFormat("SaldoCap", 10) & _
                       ImpreFormat("CAPITAL", 11) & _
                       ImpreFormat("INTERES", 8) & _
                       ImpreFormat("MORA", 6) & _
                       ImpreFormat("I.T.F.", 8) & _
                       ImpreFormat("OTROS", 8) & _
                       ImpreFormat("CAJA", 4) & Chr(10)

        lscadena = lscadena & String(100, "-") & Chr(10)

    End If
    rs.MoveNext
Loop
End If
lscadena = lscadena & String(100, "-") & Chr(10)

lscadena = lscadena & ImpreFormat("TOTALES :", 6) & _
                        ImpreFormat("CAP.SIAFC:", 10) & _
                        ImpreFormat(lnTotaCapSIAFC, 11, 2) & _
                        ImpreFormat("CAP.SICMAC:", 10) & _
                        ImpreFormat(lnTotaCapSICMAC, 11, 2) & _
                        ImpreFormat("INT.SIAFC:", 10) & _
                        ImpreFormat(lnTotIntSIAFC, 11, 2) & _
                        ImpreFormat("INT.SICMAC:", 10) & _
                        ImpreFormat(lnTotIntSICMAC, 11, 2) & Chr(10) & ImpreFormat("", 6) & _
                        ImpreFormat("MORA.SIAFC:", 10) & _
                        ImpreFormat(lnTotaMoraSIAFC, 11, 2) & _
                        ImpreFormat("MORA.SICMAC:", 10) & _
                        ImpreFormat(lnTotaMoraSICMAC, 11, 2) & _
                        ImpreFormat("ITF.SIAFC:", 10) & _
                        ImpreFormat(lnTotaITFSIAFC, 11, 2) & _
                        ImpreFormat("ITF.SICMAC:", 10) & _
                        ImpreFormat(lnTotaITFSICMAC, 11, 2) & Chr(10)

lscadena = lscadena & ImpreFormat("TOTALES CASOS:", 20) & _
                        ImpreFormat("SIAFC:", 10) & _
                        ImpreFormat(lnTotalCredSiafc, 11, 2) & _
                        ImpreFormat("SICMAC-I:", 10) & _
                        ImpreFormat(lnTotalCredSicmac, 11, 2) & Chr(10)


rs.Close
Set rs = Nothing
ReportePagosPREND = lscadena

End Function
Public Function GetCancAhorrosSIAFC(ByVal pdFecha As Date, ByVal psMoneda As String, ByVal psCodUsu As String)
Dim sql As String
Dim rs As ADODB.Recordset

DeterminaRutas

sql = "SELECT   substr(kd.ccodcta,7,2) as Producto, KD.CCODCTA, KD.CNRODOC, c.cnomcli,  " _
    & "         SUM(IIF(CTIPOPE$'1300,1600,1700,2300,2600,2700', ABS(NMONTO), 0000000.000 )) AS CAPITAL,  " _
    & "         SUM(IIF(CTIPOPE$'1310,1610,1710,2310,2600,2710', ABS(NMONTO), 0000000.000 )) AS INTERES,  " _
    & "         SUM(IIF(CTIPOPE$'8000,8001,8004,8005', ABS(NMONTO), 0000000.000 )) AS ITF  " _
    & "  FROM    " & gsBaseAho & "AVIDKARD KD , " & gsBaseAho & "avimctas a , " & gsBaseCli & "climide c   " _
    & "  WHERE   KD.CTIPOPE $'1300,1310,1600,1610,1700,1710,2300,2310,2600,2610,2700,2710,8000,8001,8004,8005'  " _
    & "         AND KD.CESTADO='N' AND KD.DFECPRO = CTOD('" & Format(pdFecha, "mm/dd/yyyy") & "')  AND KD.CCODUSU='" & psCodUsu & "' AND SUBSTR(KD.CCODCTA,9,1)='" & psMoneda & "' " _
    & "         and a.ccodcta = kd.ccodcta  and c.ccodcli = a.ccodcli   " _
    & "         AND EXIST ( SELECT *  " _
    & "                     FROM " & gsBaseAho & "AVIDKARD D1  " _
    & "                     WHERE D1.CTIPOPE $'1300,1310,1600,1610,1700,1710,2300,2310,2600,2610,2700,2710'  " _
    & "                     AND D1.CESTADO='N' AND D1.DFECPRO= KD.DFECPRO AND KD.CCODCTA = D1.CCODCTA AND D1.CNRODOC=KD.CNRODOC)  " _
    & " GROUP BY kd.CCODCTA, kd.CNRODOC " _
    & " union "
sql = sql + " SELECT  substr(kd.ccodcta,7,2) as Producto, KD.CCODCTA, KD.CNRODOC, c.cnomcli,   " _
    & "             SUM(IIF(CTIPOPE$'1300,1600,1700,2300,2600,2700', ABS(NMONTO), 0000000.000 )) AS CAPITAL,  " _
    & "             SUM(IIF(CTIPOPE$'1310,1610,1710,2310,2600,2710', ABS(NMONTO), 0000000.000 )) AS INTERES,  " _
    & "             SUM(IIF(CTIPOPE$'8000,8001,8004,8005', ABS(NMONTO), 0000000.000 )) AS ITF  " _
    & "         FROM    " & gsBaseAho & "AVIDKARD KD , " & gsBaseAho & "apfmctas a , " & gsBaseCli & "climide c   " _
    & "         WHERE   KD.CTIPOPE $'1300,1310,1600,1610,1700,1710,2300,2310,2600,2610,2700,2710,8000,8001,8004,8005'  " _
    & "                 AND KD.CESTADO='N' AND KD.DFECPRO = CTOD('" & Format(pdFecha, "mm/dd/yyyy") & "')  AND KD.CCODUSU='" & psCodUsu & "' AND SUBSTR(KD.CCODCTA,9,1)='" & psMoneda & "'  " _
    & "                 and a.ccodcta = kd.ccodcta  and c.ccodcli = a.ccodcli   " _
    & "                 AND EXIST ( SELECT *  " _
    & "                             FROM " & gsBaseAho & "AVIDKARD D1  " _
    & "                             WHERE D1.CTIPOPE $'1300,1310,1600,1610,1700,1710,2300,2310,2600,2610,2700,2710'  " _
    & "                 AND D1.CESTADO='N' AND D1.DFECPRO= KD.DFECPRO AND KD.CCODCTA = D1.CCODCTA AND D1.CNRODOC=KD.CNRODOC)  " _
    & " GROUP BY kd.CCODCTA, kd.CNRODOC"

Set GetCancAhorrosSIAFC = CargaRecordFox(sql)

End Function

Sub CancelacionesAhorros(ByVal pdFecha As Date, ByVal psMoneda As String, ByVal psCodUsu As String, ByVal psUsuSICMAC As String)
Dim sql As String
Dim rs As ADODB.Recordset
Dim oCon As COMConecta.DCOMConecta

Set oCon = New COMConecta.DCOMConecta
oCon.AbreConexion

lsTmpPagos = "TMPAHORROS" & psCodUsu

Set rs = oCon.CargaRecordSet("select * from sysobjects where name like '%" & lsTmpPagos & "%'")
If Not rs.EOF And Not rs.BOF Then
    sql = "DROP TABLE " & lsTmpPagos
    oCon.Ejecutar sql
End If
rs.Close
Set rs = Nothing

sql = "select  'SICMAC-I' AS SISTEMA, SUBSTRING(MD.CCTACOD,6,3) AS cProducto, MD.CCTACOD, P.CPERSNOMBRE, PP.nPrdPersRelac," _
    & "         SUM(CASE WHEN MD.nConceptoCod =1 THEN MD.NMONTO ELSE 0 END) AS CAPITAL, " _
    & "         SUM(CASE WHEN MD.nConceptoCod =2 THEN MD.NMONTO ELSE 0 END) AS INTERES, " _
    & "         SUM(CASE WHEN MD.nConceptoCod IN (20,21,22) THEN MD.NMONTO ELSE 0 END) AS ITF " _
    & " into    " & lsTmpPagos _
    & " from    mov M " _
    & "         JOIN MOVCAP MC ON MC.NMOVNRO = M.NMOVNRO " _
    & "         JOIN MOVCAPDET MD ON MD.NMOVNRO = MC.NMOVNRO AND MD.CCTACOD = MC.CCTACOD AND MD.COPECOD = MC.COPECOD " _
    & "         JOIN PRODUCTOPERSONA PP ON PP.CCTACOD = MC.CCTACOD  AND PP.nPrdPersRelac= 10 " _
    & "         JOIN PERSONA P ON P.CPERSCOD = PP.CPERSCOD " _
    & " where   M.COPECOD IN ('200401','200402','200403','200404','200405','200406','200408','210300','210301','210302') " _
    & "         and left(cmovnro,8)='" & Format(pdFecha, "yyyymmdd") & "' AND M.NMOVFLAG=0  AND RIGHT(M.CMOVNRO,4)='" & psUsuSICMAC & "' AND SUBSTRING(MD.CCTACOD,9,1)='" & psMoneda & "' " _
    & " GROUP BY MD.CCTACOD, P.CPERSNOMBRE, PP.nPrdPersRelac"


oCon.Ejecutar sql

' CANCELACIONES FOX
Set rs = GetCancAhorrosSIAFC(pdFecha, psMoneda, psCodUsu)
Do While Not rs.EOF
    sql = "INSERT INTO " & lsTmpPagos & " (SISTEMA, cProducto, CCTACOD, CPERSNOMBRE, nPrdPersRelac, CAPITAL, INTERES, ITF) " _
        & " VALUES('SIAFC','" & rs!Producto & "','" & rs!cCodCta & "','" & Trim(rs!cNomCli) & "',10," & rs!capital & "," _
        & rs!Interes & "," & rs!ITF & ")"
    
    oCon.Ejecutar sql
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing

sql = "select * from " & lsTmpPagos _
 & " order by CPERSNOMBRE, CCTACOD, SISTEMA "
Set rs = oCon.CargaRecordSet(sql)

Dim lsCad As String
Dim lsTitulo As String
lsTitulo = "REPORTE DE VALIDACION CANCELACIONES AHORROS "
lsCad = RepCancelacionAhorros(rs, pdFecha, psCodUsu, psMoneda, lsTitulo, 0)
'oPrev.Show lsCad, "", True
oPrev.Show lsCad, "", True, , gImpresora
oCon.CierraConexion

End Sub
Sub MovimientosAhorros(ByVal pdFecha As Date, ByVal psMoneda As String, ByVal psCodUsu As String, ByVal psUsuSICMAC As String, ByVal pbDepositos As Boolean)
Dim sql As String
Dim rs As ADODB.Recordset
Dim oCon As COMConecta.DCOMConecta
Dim lsFiltroUsu As String

Set oCon = New COMConecta.DCOMConecta
oCon.AbreConexion

lsTmpPagos = "TMPDEPOSITOS" & IIf(psCodUsu = "", "XXXX", psCodUsu)

Set rs = oCon.CargaRecordSet("select * from sysobjects where name like '%" & lsTmpPagos & "%'")
If Not rs.EOF And Not rs.BOF Then
    sql = "DROP TABLE " & lsTmpPagos
    oCon.Ejecutar sql
End If
rs.Close
Set rs = Nothing

If psUsuSICMAC <> "" Then
    lsFiltroUsu = " AND RIGHT (M.CMOVNRO,4)='" & psUsuSICMAC & "'"
Else
    lsFiltroUsu = ""
End If

If pbDepositos Then
    sql = "SELECT 'SICMAC-I' AS SISTEMA, CPERSNOMBRE, CCTACOD, NMONTO , CUSER, cOpeCod, cOpeDesc " _
        & " into    " & lsTmpPagos _
        & " From " _
        & "         (SELECT  (SELECT TOP 1 P.CPERSNOMBRE " _
        & "                    FROM    PRODUCTOPERSONA R " _
        & "                            JOIN PERSONA P ON P.CPERSCOD = R.CPERSCOD AND R.nPrdPersRelac = 10 " _
        & "                    Where R.cCtaCod = MC.cCtaCod " _
        & "                    ORDER BY P.CPERSNOMBRE) AS CPERSNOMBRE, MC.*, RIGHT(M.CMOVNRO,4) AS CUSER, O.cOpeDesc  " _
        & " FROM    MOV M " _
        & "         JOIN MOVCAP MC ON MC.NMOVNRO = M.NMOVNRO " _
        & "         JOIN MOVCAPDET MD ON MD.NMOVNRO = MC.NMOVNRO AND MD.COPECOD = MC.COPECOD AND MD.CCTACOD = MC.CCTACOD " _
        & "         JOIN OPETPO O ON O.COPECOD = MC.COPECOD " _
        & " WHERE   LEFT(M.CMOVNRO,8)='" & Format(pdFecha, "yyyymmdd") & "' AND SUBSTRING(MC.CCTACOD,4,2)='" & Trim(TxtBuscarAge) & "' " _
        & "         AND SUBSTRING(MC.CCTACOD,9,1)='" & psMoneda & "' " _
        & "         AND SUBSTRING(MC.CCTACOD,6,3)='232' AND (MC.COPECOD LIKE '2002%' OR MC.COPECOD IN ('260101','260102','200901')) " _
        & " AND M.NMOVFLAG =0 " & lsFiltroUsu & "  ) AS X " _
        & " ORDER BY CPERSNOMBRE"
Else
    sql = "SELECT 'SICMAC-I' AS SISTEMA, CPERSNOMBRE, CCTACOD, NMONTO , CUSER, cOpeCod, cOpeDesc " _
        & " into    " & lsTmpPagos _
        & " From " _
        & "         (SELECT  (SELECT TOP 1 P.CPERSNOMBRE " _
        & "                    FROM    PRODUCTOPERSONA R " _
        & "                            JOIN PERSONA P ON P.CPERSCOD = R.CPERSCOD AND R.nPrdPersRelac = 10 " _
        & "                    Where R.cCtaCod = MC.cCtaCod " _
        & "                    ORDER BY P.CPERSNOMBRE) AS CPERSNOMBRE, MC.*, RIGHT(M.CMOVNRO,4) AS CUSER, O.cOpeDesc " _
        & " FROM    MOV M " _
        & "         JOIN MOVCAP MC ON MC.NMOVNRO = M.NMOVNRO " _
        & "         JOIN MOVCAPDET MD ON MD.NMOVNRO = MC.NMOVNRO AND MD.COPECOD = MC.COPECOD AND MD.CCTACOD = MC.CCTACOD " _
        & "         JOIN OPETPO O ON O.COPECOD = MC.COPECOD " _
        & " WHERE   LEFT(M.CMOVNRO,8)='" & Format(pdFecha, "yyyymmdd") & "' AND SUBSTRING(MC.CCTACOD,4,2)='" & Trim(TxtBuscarAge) & "' " _
        & "         AND SUBSTRING(MC.CCTACOD,9,1)='" & psMoneda & "' " _
        & "         AND SUBSTRING(MC.CCTACOD,6,3)='232' AND (MC.COPECOD LIKE '2003%' OR MC.COPECOD IN ('260103','260104','200602','200901')) " _
        & " AND M.NMOVFLAG =0 " & lsFiltroUsu & "  ) AS X " _
        & " ORDER BY CPERSNOMBRE"
    
End If
oCon.Ejecutar sql


' MOVIMIENTOS FOX
Set rs = New ADODB.Recordset
Set rs = GetMovSIAF(pdFecha, psCodUsu, psMoneda)
Do While Not rs.EOF
    If pbDepositos Then
        If rs!nDepositos > 0 Then
            sql = "INSERT INTO " & lsTmpPagos & " (SISTEMA, CPERSNOMBRE, CCTACOD, NMONTO , CUSER, COPECOD, COPEDESC ) " _
                & " VALUES('SIAFC','" & rs!cNomCli & "','" & rs!cCodCta & "'," & rs!nDepositos & ",'" & rs!cCodUsu & "','" & rs!cTipOpe & "','" & Trim(rs!cDescri) & "')"
        
            oCon.Ejecutar sql
        End If
    Else
        If rs!nRetiros < 0 Then
            sql = "INSERT INTO " & lsTmpPagos & " (SISTEMA, CPERSNOMBRE, CCTACOD, NMONTO , CUSER, COPECOD, COPEDESC ) " _
                & " VALUES('SIAFC','" & rs!cNomCli & "','" & rs!cCodCta & "'," & Abs(rs!nRetiros) & ",'" & rs!cCodUsu & "','" & rs!cTipOpe & "','" & Trim(rs!cDescri) & "')"
                
            oCon.Ejecutar sql
        End If
    End If
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing

sql = "select * from " & lsTmpPagos _
 & " order by CPERSNOMBRE, NMONTO, SISTEMA "
Set rs = oCon.CargaRecordSet(sql)

Dim lsCad As String
Dim lsTitulo As String
If pbDepositos Then
    lsTitulo = "REPORTE DE DEPOSITOS DE AHORRROS "
Else
    lsTitulo = "REPORTE DE RETIROS DE AHORRROS "
End If

lsCad = RepMOVAhorros(rs, pdFecha, psCodUsu, psMoneda, lsTitulo, 0)
'oPrev.Show lsCad, "", True
oCon.CierraConexion

End Sub
Function GetMovSIAF(ByVal pdFecha As Date, ByVal psCodUsu As String, ByVal psMoneda As String)
Dim sql As String
Dim rs As ADODB.Recordset

DeterminaRutas
If psCodUsu <> "" Then
    lsFiltroUsu = " and kd.cCodUsu ='" & psCodUsu & "' "
Else
    lsFiltroUsu = ""
End If

sql = "select  C1.CNOMCLI, kd.CCODCTA, kd.NMONTO, kd.CTIPOPE, kd.CNRODOC, IIF(inlist(cTipOpe,'1760','1761','1763','1773'), NMONTO, 0000000.000) AS NMONINAC,  " _
    & "         IIF(inlist(cTipOpe,'1770','1771'),NMONTO,00000000.000) AS nCanInac, " _
    & "         IIF(cTipOpe='1780', NMONTO, 00000000.00 ) AS nDeschq , " _
    & "         iif(cTipOpe='1100', nmonto, 00000000.00) as nAperturas,  " _
    & "         iif((INLIST(cTipOpe,'1200','1500','1660','1661','1901','1902','2108','3108','1250','1255','1870','1501','1905','1906') .OR. LEFT(CTIPOPE,2)$'51,52') and (nMonto>0), nmonto, 0000.000) as nDepositos, " _
    & "         iif((INLIST(cTipOpe,'1200','1500','1660','1661','1901','1902','2108','3108','1250','1255','1870','1501','1905','1906') .OR. LEFT(CTIPOPE,2)$'51,52') and (nMonto<0), nmonto, 0000.000) as nRetiros, " _
    & "         iif(cTipOpe$'1300,1700', nmonto, 00000000.00) as nMontoCance,  " _
    & "         iif(cTipOpe$'1310,1710', nmonto, 00000000.00) as nIntCan,  " _
    & "         iif(INLIST(cTipOpe,'8000'),nmonto,000000.00) as nCargoITF, " _
    & "         IIF(INLIST(cTipOpe,'8001'),NMONTO, 000000.00) AS nItfEfec, " _
    & "         iif(INLIST(cTipOpe,'8003'), nmonto, 000000.00) as nRedITF, kd.cCodUsu, t.cDescri  " _
    & " from    '" & gsBaseAho & "AVIDKARD' kd  " _
    & "         JOIN '" & gsBaseAho & "AVIMCTAS' A  ON A.CCODCTA =KD.CCODCTA  " _
    & "         JOIN '" & gsBaseCli & "CLIMIDE' C1 ON C1.CCODCLI = A.CCODCLI  " _
    & "         JOIN 'F:\APL\TCO\TCOTTAB' T ON T.cCodigo = kd.cTipOpe and T.cCodTab = '135' " _
    & " where   KD.cestado ='N' and KD.dFecpro = CTOD('" & Format(pdFecha, "mm/dd/yyyy") & "')  " _
    & "         AND substr(KD.ccodcta,9,1)='" & psMoneda & "' " & lsFiltroUsu _
    & " ORDER BY C1.CNOMCLI"

'CTOD('" & Format(pdFecha, "mm/dd/yyyy") & "')
'KD.dFecpro = CTOD('" & Format(pdFecha, "mm/dd/yyyy") & "') and

Set GetMovSIAF = CargaRecordFox(sql)

End Function
Function RepCancelacionAhorros(ByVal rs As ADODB.Recordset, ByVal pdFecha As Date, ByVal psCodUsu As String, ByVal psMoneda As String, ByVal lsInfoSicmac_Siafc As String, lnPagina As Integer) As String
Dim lscadena As String
Dim lnPag As Integer
Dim lnTotaCapSICMAC As Currency
Dim lnTotaCapSIAFC As Currency

Dim lnTotIntSICMAC As Currency
Dim lnTotIntSIAFC As Currency

Dim lnTotaMoraSICMAC As Currency
Dim lnTotaMoraSIAFC As Currency

Dim lnTotaITFSICMAC As Currency
Dim lnTotaITFSIAFC As Currency

Dim lnTotalCredSicmac As Long
Dim lnTotalCredSiafc As Long

lnPag = 0
If Not rs.EOF And Not rs.BOF Then
lscadena = CabeRepo(gsNomCmac, lblDescAge, "", IIf(psMoneda = "1", "SOLES", "DOLARES"), Format(pdFecha, "dd/mm/yyyy"), "Control Calidad CANCELACIONES AHORROS", "USUARIO :" & psCodUsu, lsInfoSicmac_Siafc, "", lnPag, 64)
lscadena = lscadena + Chr(10)
lnTotalProd = 0
lnLinea = 11
lsProducto = ""
lscadena = lscadena & String(100, "-") & Chr(10)
                        
lscadena = lscadena & ImpreFormat("Sistema", 6) & _
                       ImpreFormat("N° Credito", 18) & _
                       ImpreFormat("Nombre Cliente", 20) & _
                       ImpreFormat("PROD", 4) & _
                       ImpreFormat("CAPITAL", 11) & _
                       ImpreFormat("INTERES", 8) & _
                       ImpreFormat("I.T.F.", 8) & Chr(10)

lscadena = lscadena & String(100, "-") & Chr(10)

lnTotaCapSICMAC = 0
lnTotaCapSIAFC = 0

lnTotIntSICMAC = 0
lnTotIntSIAFC = 0

lnTotaMoraSICMAC = 0
lnTotaMoraSIAFC = 0

lnTotaITFSICMAC = 0
lnTotaITFSIAFC = 0


rs.MoveFirst
Dim lsCodCta As String
Do While Not rs.EOF
    If lsCodCta = rs!cCtaCod Then
        GoTo Continua
    Else
        lsCodCta = rs!cCtaCod
    End If
    lnLinea = lnLinea + 1
    lscadena = lscadena & ImpreFormat(rs!Sistema, 6) & _
                       ImpreFormat(rs!cCtaCod, 18) & _
                       ImpreFormat(rs!CPERSNOMBRE, 20) & _
                       ImpreFormat(rs!cProducto, 3) & _
                       ImpreFormat(rs!capital, 8) & _
                       ImpreFormat(rs!Interes, 8) & _
                       ImpreFormat(rs!ITF, 6) & Chr(10)

    lnTotalCaso = lnTotalCaso + 1
    If Trim(rs!Sistema) = "SIAFC" Then
        lnTotaCapSIAFC = lnTotaCapSIAFC + rs!capital
        lnTotIntSIAFC = lnTotIntSIAFC + rs!Interes
        lnTotaITFSIAFC = lnTotaITFSIAFC + rs!ITF
        lnTotalCredSiafc = lnTotalCredSiafc + 1
    Else
        lnTotaCapSICMAC = lnTotaCapSICMAC + rs!capital
        lnTotIntSICMAC = lnTotIntSICMAC + rs!Interes
        lnTotaITFSICMAC = lnTotaITFSICMAC + rs!ITF
        lnTotalCredSicmac = lnTotalCredSicmac + 1
    End If
    If lnLinea > 61 Then
        lnLinea = 11
        lnPag = lnPag + 1
        lscadena = lscadena & CabeRepo(gsNomCmac, lblDescAge, "", IIf(psMoneda = "1", "SOLES", "DOLARES"), Format(pdFecha, "dd/mm/yyyy"), "Control Calidad PAGOS DE CREDITOS", "USUARIO :" & psCodUsu, lsInfoSicmac_Siafc, "", lnPag, 64)
        lscadena = lscadena & String(100, "-") & Chr(10)

        lscadena = lscadena & ImpreFormat("Sistema", 6) & _
                       ImpreFormat("N° Credito", 18) & _
                       ImpreFormat("Nombre Cliente", 20) & _
                       ImpreFormat("PROD", 4) & _
                       ImpreFormat("CAPITAL", 11) & _
                       ImpreFormat("INTERES", 8) & _
                       ImpreFormat("I.T.F.", 8) & Chr(10)

        lscadena = lscadena & String(100, "-") & Chr(10)

    End If
Continua:
    lsCodCta = rs!cCtaCod
    rs.MoveNext
Loop
End If
lscadena = lscadena & String(100, "-") & Chr(10)

lscadena = lscadena & ImpreFormat("TOTALES :", 6) & _
                        ImpreFormat("CAP.SIAFC:", 10) & _
                        ImpreFormat(lnTotaCapSIAFC, 11, 2) & _
                        ImpreFormat("CAP.SICMAC:", 10) & _
                        ImpreFormat(lnTotaCapSICMAC, 11, 2) & _
                        ImpreFormat("INT.SIAFC:", 10) & _
                        ImpreFormat(lnTotIntSIAFC, 11, 2) & _
                        ImpreFormat("INT.SICMAC:", 10) & _
                        ImpreFormat(lnTotIntSICMAC, 11, 2) & Chr(10) & ImpreFormat("", 6) & _
                        ImpreFormat("ITF.SIAFC:", 10) & _
                        ImpreFormat(lnTotaITFSIAFC, 11, 2) & _
                        ImpreFormat("ITF.SICMAC:", 10) & _
                        ImpreFormat(lnTotaITFSICMAC, 11, 2) & Chr(10)

lscadena = lscadena & ImpreFormat("TOTALES CASOS:", 20) & _
                        ImpreFormat("SIAFC:", 10) & _
                        ImpreFormat(lnTotalCredSiafc, 11, 2) & _
                        ImpreFormat("SICMAC-I:", 10) & _
                        ImpreFormat(lnTotalCredSicmac, 11, 2) & Chr(10)


rs.Close
Set rs = Nothing
RepCancelacionAhorros = lscadena

End Function
Function RepMOVAhorros(ByVal rs As ADODB.Recordset, ByVal pdFecha As Date, ByVal psCodUsu As String, ByVal psMoneda As String, ByVal lsInfoSicmac_Siafc As String, lnPagina As Integer) As String
Dim lscadena As String
Dim lnPag As Integer
Dim lnTotaCapSICMAC As Currency
Dim lnTotaCapSIAFC As Currency

Dim lnTotIntSICMAC As Currency
Dim lnTotIntSIAFC As Currency

Dim lnTotaMoraSICMAC As Currency
Dim lnTotaMoraSIAFC As Currency

Dim lnTotaITFSICMAC As Currency
Dim lnTotaITFSIAFC As Currency

Dim lnTotalCredSicmac As Long
Dim lnTotalCredSiafc As Long
Dim lnTotalCtasCMAC As Currency

Dim vExcelObj As Excel.Application

Dim vNHC As String

lnPag = 0
If Not rs.EOF And Not rs.BOF Then

'lsCadena = CabeRepo(gsNomCmac, lblDescAge, "", IIf(psMoneda = "1", "SOLES", "DOLARES"), Format(pdFecha, "dd/mm/yyyy"), "Control Calidad MOVIMIENTOS AHORROS", "USUARIO :" & psCodUsu, lsInfoSicmac_Siafc, "", lnPag, 64)
'lsCadena = lsCadena + Chr(10)
                               
vNHC = App.path & "\spooler\CONTROLOPE" & Format(txtfecha, "yyyymmdd") & IIf(Me.optCalidad(7).value = True, "DEP", "RET") & IIf(psMoneda = "1", "SOLES", "DOLARES") & ".XLS"

lnTotalProd = 0
lnLinea = 11
lsProducto = ""
lscadena = lscadena & String(100, "-") & Chr(10)
                        
Set vExcelObj = New Excel.Application  '   = CreateObject("Excel.Application")
vExcelObj.DisplayAlerts = True

vExcelObj.Workbooks.Add
vExcelObj.Sheets("Hoja1").Select
vExcelObj.Sheets("Hoja1").Name = "CONTROL"

vExcelObj.Range("A1:IV65536").Font.Name = "Arial Narrow"
vExcelObj.Range("A1:IV65536").Font.Size = 8
vExcelObj.Columns("A:IV").Select
vExcelObj.Selection.VerticalAlignment = 3

vExcelObj.Columns("A").Select
vExcelObj.Selection.HorizontalAlignment = 1
vExcelObj.Columns("B:H").Select
vExcelObj.Selection.HorizontalAlignment = 1
'vExcelObj.Columns("D:H").Select
'vExcelObj.Selection.HorizontalAlignment = 2

vExcelObj.Range("A1").Select
vExcelObj.Range("A1").Font.Bold = True
vExcelObj.Range("A1").HorizontalAlignment = 1
vExcelObj.ActiveCell.value = UCase(Trim(gsNomCmac))

vExcelObj.Range("D1").Select
vExcelObj.Range("D1").Font.Bold = True
vExcelObj.Range("D1").HorizontalAlignment = 1
vExcelObj.ActiveCell.value = Format(Me.txtfecha, "dd/mm/yyyy")

vExcelObj.Range("A2").Select
vExcelObj.Range("A2").Font.Bold = True
vExcelObj.Range("A2").HorizontalAlignment = 1
vExcelObj.ActiveCell.value = UCase(Trim(lblDescAge)) & " - " & IIf(psMoneda = "1", "SOLES", "DOLARES")

vExcelObj.Range("A4").Select
vExcelObj.Range("A4").Font.Bold = True
vExcelObj.Range("A4").HorizontalAlignment = 1
If psCodUsu <> "" Then
    vExcelObj.ActiveCell.value = "Control Calidad MOVIMIENTOS AHORROS - USUARIO :" & psCodUsu
Else
    vExcelObj.ActiveCell.value = "Control Calidad MOVIMIENTOS AHORROS - CONSOLIDADO"
End If

vExcelObj.Range("A5").Select
vExcelObj.Range("A5").Font.Bold = True
vExcelObj.Range("A5").HorizontalAlignment = 1
vExcelObj.ActiveCell.value = lsInfoSicmac_Siafc

vExcelObj.Range("A6").Select
vExcelObj.Range("A6").Font.Bold = True
vExcelObj.Range("A6").ColumnWidth = 10
vExcelObj.ActiveCell.value = "Sistema"

vExcelObj.Range("B6").Select
vExcelObj.Range("B6").Font.Bold = True
vExcelObj.Range("B6").ColumnWidth = 30
vExcelObj.ActiveCell.value = "Nombre Cliente"

vExcelObj.Range("C6").Select
vExcelObj.Range("C6").Font.Bold = True
vExcelObj.Range("C6").ColumnWidth = 10
vExcelObj.ActiveCell.value = "N° Cuenta"

vExcelObj.Range("D6").Select
vExcelObj.Range("D6").Font.Bold = True
vExcelObj.Range("D6").ColumnWidth = 10
vExcelObj.ActiveCell.value = "Cod.Operac."

vExcelObj.Range("E6").Select
vExcelObj.Range("E6").Font.Bold = True
vExcelObj.Range("E6").ColumnWidth = 10
vExcelObj.ActiveCell.value = "OPERACION"
                        
vExcelObj.Range("F6").Select
vExcelObj.Range("F6").Font.Bold = True
vExcelObj.Range("F6").ColumnWidth = 10
vExcelObj.ActiveCell.value = "MONTO"
                        
vExcelObj.Range("G6").Select
vExcelObj.Range("G6").Font.Bold = True
vExcelObj.Range("G6").ColumnWidth = 10
vExcelObj.ActiveCell.value = "USER"


lnTotaCapSICMAC = 0
lnTotaCapSIAFC = 0

lnTotIntSICMAC = 0
lnTotIntSIAFC = 0

lnTotaMoraSICMAC = 0
lnTotaMoraSIAFC = 0

lnTotaITFSICMAC = 0
lnTotaITFSIAFC = 0

rs.MoveFirst
Dim lsCodCta As String
vIni = 6
vItem = vIni
lnTotalCtasCMAC = 0
Do While Not rs.EOF
         vItem = vItem + 1
    
          If rs!cCtaCod = "108001211001152892" Or rs!cCtaCod = "10800121110140002X" Or rs!cCtaCod = "108001212000482992" Then
                lnTotalCtasCMAC = lnTotalCtasCMAC + rs!nMonto
          End If
    
         vCel = "A" + Trim(Str(vItem))
         vExcelObj.Range(vCel).Select
         vExcelObj.ActiveCell.value = "'" + rs!Sistema
    
         vCel = "B" + Trim(Str(vItem))
         vExcelObj.Range(vCel).Select
         vExcelObj.ActiveCell.value = "'" + rs!CPERSNOMBRE
    
         vCel = "C" + Trim(Str(vItem))
         vExcelObj.Range(vCel).Select
         vExcelObj.ActiveCell.value = "'" + rs!cCtaCod
    
         vCel = "D" + Trim(Str(vItem))
         vExcelObj.Range(vCel).Select
         vExcelObj.ActiveCell.value = "'" + rs!cOpecod
    
         vCel = "E" + Trim(Str(vItem))
         vExcelObj.Range(vCel).Select
         vExcelObj.ActiveCell.value = "'" + rs!cOpedesc
    
         vCel = "F" + Trim(Str(vItem))
         vExcelObj.Range(vCel).Select
         vExcelObj.ActiveCell.value = Format(rs!nMonto, "#0.00")
    
         vCel = "G" + Trim(Str(vItem))
         vExcelObj.Range(vCel).Select
         vExcelObj.ActiveCell.value = "'" + rs!Cuser
    
    
    lnLinea = lnLinea + 1
    
    lnTotalCaso = lnTotalCaso + 1
    If Trim(rs!Sistema) = "SIAFC" Then
        lnTotaCapSIAFC = lnTotaCapSIAFC + Abs(rs!nMonto)
        lnTotalCredSiafc = lnTotalCredSiafc + 1
    Else
        lnTotaCapSICMAC = lnTotaCapSICMAC + Abs(rs!nMonto)
        lnTotalCredSicmac = lnTotalCredSicmac + 1
    End If
    lsCodCta = rs!cCtaCod
    rs.MoveNext
Loop
vItem = vItem + 2
vCel = "A" + Trim(Str(vItem))
vExcelObj.Range(vCel).Select
vExcelObj.ActiveCell.value = "TOTALES"

vCel = "B" + Trim(Str(vItem))
vExcelObj.Range(vCel).Select
vExcelObj.ActiveCell.value = "MOV.SIAFC:"

vCel = "C" + Trim(Str(vItem))
vExcelObj.Range(vCel).Select
vExcelObj.ActiveCell.value = Format(lnTotaCapSIAFC, "#0.00")

vCel = "D" + Trim(Str(vItem))
vExcelObj.Range(vCel).Select
vExcelObj.ActiveCell.value = "MOV.SICMACT:"

vCel = "E" + Trim(Str(vItem))
vExcelObj.Range(vCel).Select
vExcelObj.ActiveCell.value = Format(lnTotaCapSICMAC, "#0.00")

vCel = "F" + Trim(Str(vItem))
vExcelObj.Range(vCel).Select
vExcelObj.ActiveCell.value = "DIFERENCIA"

vCel = "G" + Trim(Str(vItem))
vExcelObj.Range(vCel).Select
vExcelObj.ActiveCell.value = Format(lnTotaCapSIAFC - lnTotaCapSICMAC, "#0.00")

vCel = "H" + Trim(Str(vItem))
vExcelObj.Range(vCel).Select
vExcelObj.ActiveCell.value = "TOTAL MONTO CTAS CMAC ICA: "

vCel = "I" + Trim(Str(vItem))
vExcelObj.Range(vCel).Select
vExcelObj.ActiveCell.value = Format(lnTotalCtasCMAC, "#0.00")




vItem = vItem + 1
vCel = "A" + Trim(Str(vItem))
vExcelObj.Range(vCel).Select
vExcelObj.ActiveCell.value = "TOTAL CASOS"

vCel = "B" + Trim(Str(vItem))
vExcelObj.Range(vCel).Select
vExcelObj.ActiveCell.value = "SIAFC:"

vCel = "C" + Trim(Str(vItem))
vExcelObj.Range(vCel).Select
vExcelObj.ActiveCell.value = lnTotalCredSiafc

vCel = "D" + Trim(Str(vItem))
vExcelObj.Range(vCel).Select
vExcelObj.ActiveCell.value = "SICMAC-I"

vCel = "E" + Trim(Str(vItem))
vExcelObj.Range(vCel).Select
vExcelObj.ActiveCell.value = lnTotalCredSicmac



If Dir(vNHC) <> "" Then
   If MsgBox("Archivo Ya Existe ...  Desea Reemplazarlo ??", vbQuestion + vbYesNo + vbDefaultButton1, " Mensaje del Sistema ...") = vbNo Then
      Exit Function
   End If
End If
vExcelObj.Range("A1").Select
vExcelObj.ActiveWorkbook.SaveAs (vNHC)
vExcelObj.ActiveWorkbook.Close
vExcelObj.Workbooks.Open (vNHC)
vExcelObj.Visible = True

Set vExcelObj = Nothing

MsgBox "SE HA GENERADO CON ÉXITO EL ARCHIVO !!  ", vbInformation, " Mensaje del Sistema ..."

End If


rs.Close
Set rs = Nothing
RepMOVAhorros = lscadena

End Function

