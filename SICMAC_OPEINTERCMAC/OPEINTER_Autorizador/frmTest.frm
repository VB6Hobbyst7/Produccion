VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "Form1"
   ClientHeight    =   7935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   ScaleHeight     =   7935
   ScaleWidth      =   4590
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   1320
      TabIndex        =   11
      Top             =   7440
      Width           =   2055
   End
   Begin VB.CommandButton CmdTestGlobal 
      Caption         =   "Test Globalnet"
      Height          =   495
      Left            =   660
      TabIndex        =   10
      Top             =   6600
      Width           =   3255
   End
   Begin VB.CommandButton CmdCambClave 
      Caption         =   "Test de Consulta Cambio Clave"
      Height          =   495
      Left            =   600
      TabIndex        =   9
      Top             =   5760
      Width           =   3255
   End
   Begin VB.CommandButton CmdConsTipCam 
      Caption         =   "Test de Consulta Tipo Cambio"
      Height          =   495
      Left            =   600
      TabIndex        =   8
      Top             =   5160
      Width           =   3255
   End
   Begin VB.CommandButton CmdConsMov 
      Caption         =   "Test de Consulta Movimientos"
      Height          =   495
      Left            =   600
      TabIndex        =   7
      Top             =   4560
      Width           =   3255
   End
   Begin VB.CommandButton CmdConsIntegrada 
      Caption         =   "Test de Consulta Integrada"
      Height          =   495
      Left            =   600
      TabIndex        =   6
      Top             =   3960
      Width           =   3255
   End
   Begin VB.CommandButton CmdConsCta 
      Caption         =   "Test de Consulta de Cuenta"
      Height          =   495
      Left            =   600
      TabIndex        =   5
      Top             =   3360
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Test de Extorno de Transferencia"
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Top             =   2760
      Width           =   3255
   End
   Begin VB.CommandButton CmdTestTra2 
      Caption         =   "Test de Tranferencia en Diferente  Moneda"
      Height          =   495
      Left            =   600
      TabIndex        =   3
      Top             =   2160
      Width           =   3255
   End
   Begin VB.CommandButton CmdTestTra 
      Caption         =   "Test de Tranferencia en la misma Moneda"
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   1560
      Width           =   3255
   End
   Begin VB.CommandButton CmdTestExtorno 
      Caption         =   "Test de Extorno de Retiro"
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   960
      Width           =   3255
   End
   Begin VB.CommandButton CmdTest 
      Caption         =   "Test de Retiro de Efectivo"
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   375
      Width           =   3255
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public INXml As String
Public INXmlExt As String
Public INXmlTransfer As String
Public INXmlTransfDifMon As String
Public INXmlExtTransf As String
Public INXmlConsCta As String
Public INXmlConsInteg As String
Public INXmlConsMov As String
Public INXmlConsTipCam As String
Public INXmlCambClave As String

Private Sub CmdCambClave_Click()
Dim T As AutorizadorIT.ClsAutorizador
Dim S As String
Set T = New AutorizadorIT.ClsAutorizador
Call T.Ejecutor(INXmlCambClave)
Set T = Nothing

MsgBox S

End Sub

Private Sub CmdConsCta_Click()

Dim T As AutorizadorIT.ClsAutorizador
Dim S As String
Dim sTramaSalida As String
Set T = New AutorizadorIT.ClsAutorizador

'    MESSAGE_TYPE = "0200"
'    TRACE = "045986"
'    PRCODE = "500035"
'    PAN = "8109000000000060"
'    TIME_LOCAL = "104456"
'    DATE_LOCAL = "0717"
'    TERMINAL_ID = "00000370"
'    ACCT_1 = "2321000000019 26154    001"
'    CARD_ACCEPTOR = "000000000000000"
'    ACQ_INST = "426154"
'    POS_COND_CODE = "02"
'    TXN_AMOUNT = "000000010000"
'    CUR_CODE = "604"
'    ACCT_2 = ""
'    DATE_EXP = "1111"
'    CARD_LOCATION = ""
'
'
'   sTramaSalida = T.EjecutorGlobalNet(MESSAGE_TYPE, TRACE, PRCODE, PAN, TIME_LOCAL, DATE_LOCAL, TERMINAL_ID, _
'                ACCT_1, CARD_ACCEPTOR, ACQ_INST, POS_COND_CODE, TXN_AMOUNT, CUR_CODE, ACCT_2, _
'                DATE_EXP, CARD_LOCATION, "", 0)

'Call T.Ejecutor(INXmlConsCta)

Set T = Nothing

'MsgBox S
MsgBox sTramaSalida

End Sub

Private Sub CmdConsIntegrada_Click()

Dim T As AutorizadorIT.ClsAutorizador
Dim S As String
Set T = New AutorizadorIT.ClsAutorizador
Call T.Ejecutor(INXmlConsInteg)
Set T = Nothing

MsgBox S


End Sub

Private Sub CmdConsMov_Click()

Dim T As AutorizadorIT.ClsAutorizador
Dim S As String
Set T = New AutorizadorIT.ClsAutorizador
Call T.Ejecutor(INXmlConsMov)
Set T = Nothing

MsgBox S

End Sub

Private Sub CmdConsTipCam_Click()

Dim T As AutorizadorIT.ClsAutorizador
Dim S As String
Set T = New AutorizadorIT.ClsAutorizador
Call T.Ejecutor(INXmlConsTipCam)
Set T = Nothing

MsgBox S

End Sub

Private Sub CmdTest_Click()

Dim T As AutorizadorIT.ClsAutorizador
Dim S As String
Set T = New AutorizadorIT.ClsAutorizador

Dim A As AutorizadorIT.ClsAutorizador
Set A = New AutorizadorIT.ClsAutorizador

A.Prueba "8109000000000060", 200#

INXml = "[<?xml version=""1.0""?> <Messages> <TXN_FIN_REQ> <MESSAGE_TYPE value=""0200""/>  <PAN value=""4261540000000127""/>  " & _
"<PRCODE value=""311000""/>  <TXN_AMOUNT         value=""000000000000""/>  <CARDISS_AMOUNT     value=""null""/>  <TXN_DATE_TIME value=""0409180257""/>  <CONVRATE           value=""4261540000000142""/>  <TRACE value=""011968""/>  <TIME_LOCAL value=""180257""/>  <DATE_LOCAL value=""0409""/>  <DATE_EXP           value=""1110""/>  <DATE_STTL          value=""0409""/>  <DATE_CAPTURE       value=""0409""/>  <MERCHANT           value=""0000""/>  <COUNTRY_CODE       value=""null""/>  <POS_ENTRY_MODE     value=""000""/>  <POS_COND_CODE      value=""02""/>  <ACQ_INST value=""440822""/>  <ISS_INST value=""426154""/>  <PAN_EXT  value=""null""/>  <TRACK2             value=""4261540000000127=11101201947551600000""/>  <REFNUM             value=""000000000000""/>  <AUTH_CODE          value=""null""/>  <RESP_CODE value=""null""/>  <TERMINAL_ID value=""00000011""/>  <CARD_ACCEPTOR      " & _
"value=""000000000000000""/>  <CARD_LOCATION      value=""    RED UNICARD         AV. GRAU 422    ""/>  <ADD_RESP_DATA value=""""/>  <CUR_CODE           value=""840""/>  <CUR_CODE_CARDISS   value=""null""/>  <PIN_BLOCK          value=""B1912D2834F1F19E""/>  <ADD_POS_INFO       value=""null""/>  <NET_INF            value=""null""/>  <ORG_DATA           value=""null""/>  <REP_AMOUNT         value=""null""/>  <REQ_INST           value=""426154""/>  <ACCT_1             value=""2322000000019 426154     001""/>  <ACCT_2             value=""""/>  <CUSTOMER_INF_RESP  value=""null""/>  <PRIV_USE           value=""null""/>  </TXN_FIN_REQ> </Messages>]"

INXml = "[<?xml version=""1.0""?> <Messages> <TXN_FIN_REQ> <MESSAGE_TYPE value=""0200""/>  <PAN value=""4261540000000127""/>  <PRCODE value=""011000""/>  <TXN_AMOUNT         value=""000000002000""/>  <CARDISS_AMOUNT     value=""null""/>  <TXN_DATE_TIME value=""0411110333""/>  <CONVRATE           value=""4261540000000127""/>  <TRACE value=""370567""/>  <TIME_LOCAL value=""110333""/>  <DATE_LOCAL value=""0411""/>  <DATE_EXP           value=""1110""/>  <DATE_STTL          value=""0411""/>  <DATE_CAPTURE       value=""0411""/>  " & _
"<MERCHANT           value=""0000""/>  <COUNTRY_CODE       value=""null""/>  <POS_ENTRY_MODE     value=""000""/>  <POS_COND_CODE      value=""02""/>  <ACQ_INST value=""426154""/>  <ISS_INST value=""426154""/>  <PAN_EXT  value=""null""/>  <TRACK2             value=""4261540000000127=11101201113882300000""/>  <REFNUM             value=""000000000000""/>  <AUTH_CODE          value=""null""/>  <RESP_CODE value=""null""/>  <TERMINAL_ID value=""00000370""/> " & _
" <CARD_ACCEPTOR      value=""000000000000000""/>  <CARD_LOCATION      value=""  RED UNICARD       CAJERO TEST CUZCO   ""/>  <ADD_RESP_DATA value=""""/>  <CUR_CODE           value=""840""/>  <CUR_CODE_CARDISS   value=""null""/>  <PIN_BLOCK          value=""2FF9F41253ED5309""/>  <ADD_POS_INFO       value=""null""/>  <NET_INF            value=""null""/>  <ORG_DATA           value=""null""/>  <REP_AMOUNT         value=""null""/>  <REQ_INST           value=""426154""/>  <ACCT_1             value=""2322002002892 426154     001""/>  <ACCT_2             value=""""/>  <CUSTOMER_INF_RESP  value=""null""/>  <PRIV_USE           value=""null""/>  </TXN_FIN_REQ> </Messages>]"

'pruebas
INXml = "[<?xml version=""1.0""?> <Messages> <TXN_FIN_REQ> <MESSAGE_TYPE value=""0200""/>  <PAN value=""4261540000000127""/>  <PRCODE value=""311000""/>  <TXN_AMOUNT         value=""000000000000""/>  <CARDISS_AMOUNT     value=""null""/>  <TXN_DATE_TIME value=""0519124420""/>  <CONVRATE           value=""4261540000000127""/>  <TRACE value=""370614""/>  <TIME_LOCAL value=""124420""/>  <DATE_LOCAL value=""0519""/>  <DATE_EXP           value=""1110""/>  <DATE_STTL          value=""0519""/>  " & _
"<DATE_CAPTURE       value=""0519""/>  <MERCHANT           value=""0000""/>  <COUNTRY_CODE       value=""null""/>  <POS_ENTRY_MODE     value=""000""/>  <POS_COND_CODE      value=""02""/>  <ACQ_INST value=""426154""/>  <ISS_INST value=""426154""/>  <PAN_EXT  value=""null""/>  <TRACK2             value=""4261540000000127=11101201113882300000""/>  <REFNUM             value=""000000000000""/>  <AUTH_CODE          value=""null""/>  <RESP_CODE value=""null""/>  <TERMINAL_ID value=""00000370""/>  <CARD_ACCEPTOR      value=""000000000000000""/>  <CARD_LOCATION      value=""  RED UNICARD       CAJERO TEST CUZCO   ""/>  <ADD_RESP_DATA value=""""/>   " & _
"<CUR_CODE           value=""840""/>  <CUR_CODE_CARDISS   value=""null""/>  <PIN_BLOCK          value=""B9701399AA4F1BDC""/>  <ADD_POS_INFO       value=""null""/>  <NET_INF            value=""null""/>  <ORG_DATA           value=""null""/>  <REP_AMOUNT         value=""null""/>  <REQ_INST           value=""426154""/>  <ACCT_1             value=""2322002002892 426154     001""/>  <ACCT_2             value=""""/>  <CUSTOMER_INF_RESP  value=""null""/>  <PRIV_USE           value=""null""/>  </TXN_FIN_REQ> </Messages>]"

INXml = "[<?xml version=""1.0""?> <Messages> <TXN_FIN_REQ> <MESSAGE_TYPE value=""0200""/>  <PAN value=""4261540000000027""/>  <PRCODE value=""011000""/>  <TXN_AMOUNT         value=""000000005000""/>  <CARDISS_AMOUNT     value=""null""/>  <TXN_DATE_TIME value=""0523170219""/>  <CONVRATE           value=""4261540000000027""/>  <TRACE value=""370654""/>  <TIME_LOCAL value=""170219""/>  <DATE_LOCAL value=""0523""/>  <DATE_EXP           value=""1004""/>  <DATE_STTL          value=""0523""/>  <DATE_CAPTURE       value=""0523""/>  <MERCHANT           value=""0000""/>  <COUNTRY_CODE       value=""null""/>  <POS_ENTRY_MODE     value=""000""/>  <POS_COND_CODE" & _
" value=""02""/>  <ACQ_INST value=""426154""/>  <ISS_INST value=""426154""/>  <PAN_EXT  value=""null""/>  <TRACK2             value=""4261540000000027=10041201195363600000""/>  <REFNUM             value=""000000000000""/>  <AUTH_CODE          value=""null""/>  <RESP_CODE value=""null""/>  <TERMINAL_ID value=""00000370""/>  <CARD_ACCEPTOR      value=""000000000000000""/>  <CARD_LOCATION      value=""  RED UNICARD       CAJERO TEST CUZCO   ""/>  <ADD_RESP_DATA value=""""/>  <CUR_CODE           value=""604""/>  <CUR_CODE_CARDISS   value=""null""/>  <PIN_BLOCK          value=""1659E4CE1F0B7BAC""/>  <ADD_POS_INFO       value=""null""/>  " & _
" <NET_INF            value=""null""/>  <ORG_DATA           value=""null""/>  <REP_AMOUNT         value=""null""/>  <REQ_INST           value=""426154""/>  <ACCT_1             value=""23210029116630426154     001""/>  <ACCT_2             value=""""/>  <CUSTOMER_INF_RESP  value=""null""/>  <PRIV_USE           value=""null""/>  </TXN_FIN_REQ> </Messages>]"


INXml = "[<?xml version=""1.0""?> <Messages> <TXN_FIN_REQ> <MESSAGE_TYPE value=""0200""/>  <PAN value=""4261540000000027""/>  <PRCODE value=""391000""/>  <TXN_AMOUNT         value=""000000000000""/>  <CARDISS_AMOUNT     value=""null""/>  <TXN_DATE_TIME value=""0524124605""/>  <CONVRATE           value=""4261540000000027""/>  <TRACE value=""370665""/>  <TIME_LOCAL value=""124605""/>  <DATE_LOCAL value=""0524""/>  <DATE_EXP           value=""1004""/>  <DATE_STTL          value=""0524""/>  <DATE_CAPTURE       value=""0524""/>  <MERCHANT           value=""0000""/>  <COUNTRY_CODE       value=""null""/>  <POS_ENTRY_MODE     value=""000""/>  <POS_COND_CODE      value=""02""/>  <ACQ_INST value=""426154""/>  <ISS_INST value=""426154""/>  " & _
"<PAN_EXT  value=""null""/>  <TRACK2             value=""4261540000000027=10041201195363600000""/>  <REFNUM             value=""000000000000""/>  <AUTH_CODE          value=""null""/>  <RESP_CODE value=""null""/>  <TERMINAL_ID value=""00000370""/>  <CARD_ACCEPTOR      value=""000000000000000""/>  <CARD_LOCATION      value=""  RED UNICARD       CAJERO TEST CUZCO   ""/>  <ADD_RESP_DATA value=""""/>  <CUR_CODE           value=""604""/>  <CUR_CODE_CARDISS   value=""null""/>  <PIN_BLOCK          value=""F31C71E9CDEB0181""/>  <ADD_POS_INFO       value=""null""/>  <NET_INF            value=""null""/>  <ORG_DATA           value=""null""/>  <REP_AMOUNT         value=""null""/>  <REQ_INST           value=""426154""/>  " & _
"<ACCT_1             value=""23210029116630426154     001""/>  <ACCT_2             value=""""/>  <CUSTOMER_INF_RESP  value=""null""/>  <PRIV_USE           value=""null""/>  </TXN_FIN_REQ> </Messages>]"

INXml = "[<?xml version=""1.0""?> <Messages> <TXN_FIN_REQ> <MESSAGE_TYPE value=""0200""/>  <PAN value=""4261540000000027""/>  <PRCODE value=""930099""/>  <TXN_AMOUNT         value=""000000000000""/>  <CARDISS_AMOUNT     value=""null""/>  <TXN_DATE_TIME value=""0526114119""/>  <CONVRATE           value=""4261540000000027""/>  <TRACE value=""370670""/>  <TIME_LOCAL value=""114119""/>  <DATE_LOCAL value=""0526""/>  <DATE_EXP           value=""1004""/>  <DATE_STTL          value=""0526""/>  <DATE_CAPTURE       value=""0526""/>  <MERCHANT           value=""0000""/>  <COUNTRY_CODE       value=""null""/>  <POS_ENTRY_MODE     value=""000""/>  <POS_COND_CODE      value=""02""/>  " & _
"<ACQ_INST value=""426154""/>  <ISS_INST value=""426154""/>  <PAN_EXT  value=""null""/>  <TRACK2             value=""4261540000000027=10041201195363600000""/>  <REFNUM             value=""000000000000""/>  <AUTH_CODE          value=""null""/>  <RESP_CODE value=""null""/>  <TERMINAL_ID value=""00000370""/>  <CARD_ACCEPTOR      value=""000000000000000""/>  <CARD_LOCATION      value=""  RED UNICARD       CAJERO TEST CUZCO   ""/>  <ADD_RESP_DATA value=""""/>  <CUR_CODE           value=""604""/>  <CUR_CODE_CARDISS   value=""null""/>  <PIN_BLOCK          value=""1659E4CE1F0B7BAC""/>  <ADD_POS_INFO       value=""null""/>  <NET_INF            value=""null""/>  " & _
"<ORG_DATA           value=""null""/>  <REP_AMOUNT         value=""null""/>  <REQ_INST           value=""426154""/>  <ACCT_1             value=""426154     1""/>  <ACCT_2             value=""""/>  <CUSTOMER_INF_RESP  value=""null""/>  <PRIV_USE           value=""null""/>  </TXN_FIN_REQ> </Messages>]"

INXml = "[<?xml version=""1.0""?> <Messages> <TXN_FIN_REQ> <MESSAGE_TYPE value=""0200""/>  <PAN value=""4261540000000159""/>  <PRCODE value=""311000""/>  <TXN_AMOUNT         value=""000000000000""/>  <CARDISS_AMOUNT     value=""null""/>  <TXN_DATE_TIME value=""0526114057""/>  <CONVRATE           value=""4261540000000159""/>  <TRACE value=""370669""/>  <TIME_LOCAL value=""114057""/>  <DATE_LOCAL value=""0526""/>  <DATE_EXP           value=""1111""/>  <DATE_STTL          value=""0526""/>  <DATE_CAPTURE       value=""0526""/>  <MERCHANT           value=""0000""/>  <COUNTRY_CODE       value=""null""/>  <POS_ENTRY_MODE     value=""000""/>  <POS_COND_CODE      value=""02""/>  <ACQ_INST value=""426154""/>  <ISS_INST value=""426154""/> " & _
" <PAN_EXT  value=""null""/>  <TRACK2             value=""4261540000000159=11111201705242600000""/>  <REFNUM             value=""000000000000""/>  <AUTH_CODE          value=""null""/>  <RESP_CODE value=""null""/>  <TERMINAL_ID value=""00000370""/>  <CARD_ACCEPTOR      value=""000000000000000""/>  <CARD_LOCATION      value=""  RED UNICARD       CAJERO TEST CUZCO   ""/>  <ADD_RESP_DATA value=""""/>  <CUR_CODE           value=""604""/>  <CUR_CODE_CARDISS   value=""null""/>  <PIN_BLOCK          value=""07886DF826EE2965""/>  <ADD_POS_INFO       value=""null""/>  <NET_INF            value=""null""/>  <ORG_DATA           value=""null""/>  <REP_AMOUNT         value=""null""/>  <REQ_INST           value=""426154""/>   " & _
" <ACCT_1             value=""23210006570930426154     001""/>  <ACCT_2             value=""""/>  <CUSTOMER_INF_RESP  value=""null""/>  <PRIV_USE           value=""null""/>  </TXN_FIN_REQ> </Messages>] "


'compra POS
INXml = "[<?xml version=""1.0""?> <Messages> <TXN_FIN_REQ> <MESSAGE_TYPE value=""0200""/>  <PAN value=""4261540000000127""/>  <PRCODE value=""971000""/>  <TXN_AMOUNT         value=""000000005000""/>  <CARDISS_AMOUNT     value=""null""/>  <TXN_DATE_TIME value=""0507225356""/>  <CONVRATE           value=""4261540000000127""/>  <TRACE value=""000228""/>  <TIME_LOCAL value=""175356""/>  <DATE_LOCAL value=""0507""/>  <DATE_EXP           value=""1110""/>  <DATE_STTL          value=""0507""/>  <DATE_CAPTURE       value=""0507""/>  <MERCHANT           value=""5411""/>  <COUNTRY_CODE       value=""null""/>  <POS_ENTRY_MODE     value=""901""/>  <POS_COND_CODE      value=""51""/>  <ACQ_INST value=""457043""/>  <ISS_INST value=""426154""/>  " & _
"<PAN_EXT  value=""null""/>  <TRACK2             value=""4261540000000127=11101201947551600000""/>  <REFNUM             value=""812822000228""/>  <AUTH_CODE          value=""null""/>  <RESP_CODE value=""null""/>  <TERMINAL_ID value=""37616720""/>  <CARD_ACCEPTOR      value=""122077201      ""/>  <CARD_LOCATION      value=""PIER LUIGI ASCILI S.     LIMA         PE""/>  <ADD_RESP_DATA value=""""/>  <CUR_CODE           value=""840""/>  <CUR_CODE_CARDISS   value=""null""/>  <PIN_BLOCK          value=""0000000000000000""/>  <ADD_POS_INFO       value=""null""/>  <NET_INF            value=""null""/>  <ORG_DATA           value=""null""/>  <REP_AMOUNT         value=""null""/>  <REQ_INST           value=""426154""/>  <ACCT_1     " & _
"value=""2321000000019 26154    001""/>  <ACCT_2             value=""""/>  <CUSTOMER_INF_RESP  value=""null""/>  <PRIV_USE           value=""null""/>  </TXN_FIN_REQ> </Messages>]"


'transferencia
INXml = "[<?xml version=""1.0""?> <Messages> <TXN_FIN_REQ> <MESSAGE_TYPE value=""0200""/>  <PAN value=""4261540000000127""/>  <PRCODE value=""401010""/>  <TXN_AMOUNT         value=""000000001000""/>  <CARDISS_AMOUNT     value=""null""/>  <TXN_DATE_TIME value=""0529170618""/>  <CONVRATE           value=""4261540000000159""/>  <TRACE value=""370718""/>  <TIME_LOCAL value=""170618""/>  <DATE_LOCAL value=""0529""/>  <DATE_EXP           value=""1111""/>  <DATE_STTL          value=""0529""/>  <DATE_CAPTURE       value=""0529""/>  <MERCHANT           value=""0000""/>  <COUNTRY_CODE       value=""null""/>  <POS_ENTRY_MODE     value=""000""/>  <POS_COND_CODE      value=""02""/>  <ACQ_INST value=""426154""/>  <ISS_INST value=""426154""/>  " & _
"<PAN_EXT  value=""null""/>  <TRACK2             value=""4261540000000159=11111201705242600000""/>  <REFNUM             value=""000000000000""/>  <AUTH_CODE          value=""null""/>  <RESP_CODE value=""null""/>  <TERMINAL_ID value=""00000370""/>  <CARD_ACCEPTOR      value=""000000000000000""/>  <CARD_LOCATION      value=""  RED UNICARD       CAJERO TEST CUZCO   ""/>  <ADD_RESP_DATA value=""""/>  <CUR_CODE           value=""604""/>  <CUR_CODE_CARDISS   value=""null""/>  <PIN_BLOCK          value=""53F7843C13366494""/>  <ADD_POS_INFO       value=""null""/>  <NET_INF            value=""null""/>  <ORG_DATA           value=""null""/>  <REP_AMOUNT         value=""null""/>  <REQ_INST           value=""426154""/>  " & _
"<ACCT_1             value=""2321000000019 26154    001""/>  <ACCT_2             value=""2322000000019 26154    001""/>  <CUSTOMER_INF_RESP  value=""null""/>  <PRIV_USE           value=""null""/>  </TXN_FIN_REQ> </Messages>]"

'retiro
INXml = "[<?xml version=""1.0""?> <Messages> <TXN_FIN_REQ> <MESSAGE_TYPE value=""0200""/>  <PAN value=""4261540000000142""/>  <PRCODE value=""011000""/>  <TXN_AMOUNT         value=""000000005000""/>  <CARDISS_AMOUNT     value=""null""/>  <TXN_DATE_TIME value=""0529170532""/>  <CONVRATE           value=""4261540000000159""/>  <TRACE value=""370717""/>  <TIME_LOCAL value=""170532""/>  <DATE_LOCAL value=""0529""/>  <DATE_EXP           value=""1111""/>  <DATE_STTL          value=""0529""/>  <DATE_CAPTURE       value=""0529""/>  <MERCHANT           value=""0000""/>  <COUNTRY_CODE       value=""null""/>  <POS_ENTRY_MODE     value=""000""/>  <POS_COND_CODE      value=""02""/>  <ACQ_INST value=""426154""/>  " & _
"<ISS_INST value=""426154""/>  <PAN_EXT  value=""null""/>  <TRACK2             value=""4261540000000142=11111201705242600000""/>  <REFNUM             value=""000000000000""/>  <AUTH_CODE          value=""null""/>  <RESP_CODE value=""null""/>  <TERMINAL_ID value=""00000370""/>  <CARD_ACCEPTOR      value=""000000000000000""/>  <CARD_LOCATION      value=""  RED UNICARD       CAJERO TEST CUZCO   ""/>  <ADD_RESP_DATA value=""""/>  <CUR_CODE           value=""604""/>  <CUR_CODE_CARDISS   value=""null""/>  <PIN_BLOCK          value=""53F7843C13366494""/>  <ADD_POS_INFO       value=""null""/>  <NET_INF            value=""null""/>  <ORG_DATA           value=""null""/>  " & _
"<REP_AMOUNT         value=""null""/>  <REQ_INST           value=""426154""/>  <ACCT_1             value=""2321000000019 26154    001""/>  <ACCT_2             value=""""/>  <CUSTOMER_INF_RESP  value=""null""/>  <PRIV_USE           value=""null""/>  </TXN_FIN_REQ> </Messages>]"

'Retiro
INXml = "<Messages> <TXN_FIN_REQ> <MESSAGE_TYPE value=""0200""/>  <PAN value=""4261540000000127""/>  <PRCODE value=""011000""/>  <TXN_AMOUNT value=""000000002200""/>  <CARDISS_AMOUNT     value=""null""/>  <TXN_DATE_TIME value=""0708132127""/>  <CONVRATE           value=""4261540000000127""/>  <TRACE value=""370814""/>  <TIME_LOCAL value=""132127""/>  <DATE_LOCAL value=""0708""/> <DATE_EXP           value=""1110""/>  <DATE_STTL          value=""0708""/> <DATE_CAPTURE       value=""0708""/>  <MERCHANT           value=""0000""/> <COUNTRY_CODE       value=""null""/>  <POS_ENTRY_MODE     value=""000""/> <POS_COND_CODE      value=""02""/>  <ACQ_INST value=""426154""/>  <ISS_INST value=""426154""/>  <PAN_EXT  value=""null""/>  <TRACK2 value=""4261540000000127=11101201113882300000""/>  <REFNUM value=""000000000000""/>  <AUTH_CODE          value=""null""/>  <RESP_CODE value=""null""/>  <TERMINAL_ID value=""00000370""/>  <CARD_ACCEPTOR value=""000000000000000""/>  " & _
"<CARD_LOCATION      value=""  RED UNICARD       CAJERO TEST CUZCO   ""/>  <ADD_RESP_DATA value=""""/>  <CUR_CODE           value=""604""/>" & _
"<CUR_CODE_CARDISS   value=""null""/>  <PIN_BLOCK          value=""52A35C8D4E01B958""/>" & _
"<ADD_POS_INFO       value=""null""/>  <NET_INF            value=""null""/>  <ORG_DATA  value=""null""/>  <REP_AMOUNT         value=""null""/>  " & _
"<REQ_INST value=""426154""/>  <ACCT_1             value=""2321000000019 26154    001""/>" & _
"<ACCT_2             value=""""/>  <CUSTOMER_INF_RESP  value=""null""/>  <PRIV_USE   value=""null""/>  </TXN_FIN_REQ> </Messages>]"

'Retiro graba mal ctrama
INXml = " [<?xml version=""1.0""?> <Messages> <TXN_FIN_REQ> <MESSAGE_TYPE value=""0200""/>  <PAN "
INXml = INXml & " value=""4261540000000127""/>  <PRCODE value=""011000""/>  <TXN_AMOUNT "
INXml = INXml & " value=""000000010000""/>  <CARDISS_AMOUNT     value=""null""/>  <TXN_DATE_TIME "
INXml = INXml & " value=""0717104456""/>  <CONVRATE           value=""4261540000000159""/>  <TRACE "
INXml = INXml & " value=""370846""/>  <TIME_LOCAL value=""104456""/>  <DATE_LOCAL value=""0717""/> "
INXml = INXml & " <DATE_EXP           value=""1111""/>  <DATE_STTL          value=""0717""/> "
INXml = INXml & " <DATE_CAPTURE       value=""0717""/>  <MERCHANT           value=""0000""/> "
INXml = INXml & " <COUNTRY_CODE       value=""null""/>  <POS_ENTRY_MODE     value=""000""/> "
INXml = INXml & " <POS_COND_CODE      value=""02""/>  <ACQ_INST value=""426154""/>  <ISS_INST "
INXml = INXml & " value=""426154""/>  <PAN_EXT  value=""null""/>  <TRACK2 "
INXml = INXml & " value=""4261540000000159=11111201705242600000""/>  <REFNUM "
INXml = INXml & " value=""000000000000""/>  <AUTH_CODE          value=""null""/>  <RESP_CODE"
INXml = INXml & " value=""null""/>  <TERMINAL_ID value=""00000370""/>  <CARD_ACCEPTOR"
INXml = INXml & " value=""000000000000000""/>  <CARD_LOCATION      value=""  RED UNICARD       CAJERO"
INXml = INXml & " TEST CUZCO   ""/>  <ADD_RESP_DATA value=""""/>  <CUR_CODE           value=""604""/>"
INXml = INXml & " <CUR_CODE_CARDISS   value=""null""/>  <PIN_BLOCK          value=""E2ADA72FBC4EB9FD""/>"
INXml = INXml & " <ADD_POS_INFO       value=""null""/>  <NET_INF            value=""null""/>  <ORG_DATA "
INXml = INXml & "        value=""null""/>  <REP_AMOUNT         value=""null""/>  <REQ_INST"
INXml = INXml & "  value=""426154""/>  <ACCT_1             value=""2321000000019 26154    001""/>"
INXml = INXml & "  <ACCT_2             value=""""/>  <CUSTOMER_INF_RESP  value=""null""/>  <PRIV_USE"
INXml = INXml & "    value=""null""/>  </TXN_FIN_REQ> </Messages>]"

'Retiro graba mal ctrama
'INXml = "<?xml version=""1.0""?> <Messages> <TXN_FIN_REQ> <MESSAGE_TYPE value=""0200""/>  "
'INXml = INXml & " <PAN value=""4261540000000127""/>  <PRCODE value=""391000""/>  <TXN_AMOUNT         value=""000000000000""/>  <CARDISS_AMOUNT     value=""null""/>  <TXN_DATE_TIME value=""0717184959""/>  <CONVRATE           value=""4261540000000159""/>  "
'INXml = INXml & " <TRACE value=""370886""/>  <TIME_LOCAL value=""184959""/>  <DATE_LOCAL value=""0717""/>  <DATE_EXP           value=""1111""/>  <DATE_STTL          value=""0717""/>  <DATE_CAPTURE       value=""0717""/>  <MERCHANT           value=""0000""/>  <COUNTRY_CODE       value=""null""/>  <POS_ENTRY_MODE     value=""000""/>  <POS_COND_CODE      value=""02""/>  <ACQ_INST value=""426154""/>  <ISS_INST value=""426154""/>  <PAN_EXT  value=""null""/>  <TRACK2             value=""4261540000000159=11111201705242600000""/>  <REFNUM             value=""000000000000""/>  <AUTH_CODE          value=""null""/>  <RESP_CODE value=""null""/>  <TERMINAL_ID value=""00000370""/>  <CARD_ACCEPTOR      value=""000000000000000""/>  "
'INXml = INXml & " <CARD_LOCATION      value=""""  RED UNICARD       CAJERO TEST CUZCO   ""/>  <ADD_RESP_DATA value=""""/>  <CUR_CODE           value=""604""/>  <CUR_CODE_CARDISS   value=""null""/>  <PIN_BLOCK          value=""6E02523B8151EF82""/>  <ADD_POS_INFO       value=""null""/>  <NET_INF            value=""null""/>  <ORG_DATA           value=""null""/>  <REP_AMOUNT         value=""null""/>  <REQ_INST           value=""426154""/>  <ACCT_1             value=""2321000000019 26154    001""/>  <ACCT_2             value=""""/>  <CUSTOMER_INF_RESP  value=""null""/>  <PRIV_USE           value=""null""/>  </TXN_FIN_REQ> </Messages>"
'

S = T.Ejecutor(INXml)
Set T = Nothing

MsgBox S

End Sub

Private Sub CmdTestExtorno_Click()
Dim T As AutorizadorIT.ClsAutorizador
Dim S As String
Set T = New AutorizadorIT.ClsAutorizador
Call T.Ejecutor(INXmlExt)
Set T = Nothing

MsgBox S

End Sub

Private Sub CmdTestGlobal_Click()
Dim A As AutorizadorIT.ClsAutorizador
Dim sTramaSalida As String
Dim sCampoP039 As String
Dim sCampoS125 As String
Dim sCampoS102 As String
Dim lsCtaCod As String
Dim lsDNI As String

Set A = New AutorizadorIT.ClsAutorizador

    gsMESSAGE_TYPE = "0200"
    gsTRACE = "727514003717"
    gsPRCODE = "500035"
    gsPAN = "8109000000001829"
    gsTIME_LOCAL = "104456"
    gsDATE_LOCAL = "0717"
    gsTERMINAL_ID = "00000370"
    ACCT_1 = "2321000000019 26154    001"
    gsCARD_ACCEPTOR = "000000000000000"
    gsACQ_INST = "426154"
    gsPOS_COND_CODE = "02"
    gsTXN_AMOUNT = "000000010000"
    gsCUR_CODE = "604"
    ACCT_2 = ""
    gsDATE_EXP = "1111"
    gsCARD_LOCATION = ""
    lsCtaCod = "109013201003015025"
    lsDNI = ""

   sTramaSalida = A.EjecutorGlobalNet(gsMESSAGE_TYPE, gsTRACE, gsPRCODE, gsPAN, gsTIME_LOCAL, gsDATE_LOCAL, gsTERMINAL_ID, _
                ACCT_1, gsCARD_ACCEPTOR, gsACQ_INST, gsPOS_COND_CODE, gsTXN_AMOUNT, gsCUR_CODE, ACCT_2, _
                gsDATE_EXP, gsCARD_LOCATION, "604", "200911041109001090100CMAC", 0, lsCtaCod, lsDNI)
                
    '109033011003330745
    '109012321000607452
    
    sCampoS102 = Left(Right(sTramaSalida, 18) & String(28, " "), 28)
    sTramaSalida = UCase(sTramaSalida)
    sCampoP039 = Mid(sTramaSalida, 1, 2)
    sCampoS125 = Right(sTramaSalida, Len(sTramaSalida) - 2)
    
Set A = Nothing

End Sub

Private Sub CmdTestTra_Click()

Dim T As AutorizadorIT.ClsAutorizador
Dim S As String
Set T = New AutorizadorIT.ClsAutorizador
Call T.Ejecutor(INXmlTransfer)
MsgBox S
Set T = Nothing


End Sub

Private Sub CmdTestTra2_Click()
Dim T As AutorizadorIT.ClsAutorizador
Dim S As String
Set T = New AutorizadorIT.ClsAutorizador
Call T.Ejecutor(INXmlTransfDifMon)
MsgBox S
Set T = Nothing

End Sub

Private Sub Command1_Click()

Dim T As AutorizadorIT.ClsAutorizador
Dim S As String
Set T = New AutorizadorIT.ClsAutorizador
Call T.Ejecutor(INXmlExtTransf)
MsgBox S
Set T = Nothing

End Sub



'Private Sub Command2_Click()
'Dim T As AutorizadorIT.ClsAutorizador
'Dim S As String
'Set T = New AutorizadorIT.ClsAutorizador
'
'        If Not T.PIT_ValidaExtorno("", 31057544, 1) Then
'            MsgBox "Entro"
'            Exit Sub
'        End If
'
'Set T = Nothing
'End Sub

Private Sub Form_Load()
Dim sIDTRAMA As String

sIDTRAMA = "186034"

'*************************************************
'TRAMA DE RETIRO
'*************************************************
INXml = "<Messages> "
INXml = INXml & "<TXN_FIN_REQ>"
INXml = INXml & "<MESSAGE_TYPE   value=""200"" />"
INXml = INXml & "<PAN    value=""0106000100000127"" />"
INXml = INXml & "<PRCODE     value=""011000"" />"
INXml = INXml & "<TXN_AMOUNT value=""00000005000"" />"
INXml = INXml & "<CARDISS_AMOUNT  value="""" />"
INXml = INXml & "<TXN_DATE_TIME      value=""0410105413"" />"
INXml = INXml & "<CONVRATE   value="""" />"
INXml = INXml & "<TRACE  value=""" & sIDTRAMA & """" & " />"
INXml = INXml & "<TIME_LOCAl value=""055604"" />"
INXml = INXml & "<DATE_LOCAL     value=""0410"" />"
INXml = INXml & "<DATE_EXP   value=""0907"" />"
INXml = INXml & "<DATE_STTL   value=""0410"" />"
INXml = INXml & "<DATE_CAPTURE   value=""0410"" />"
INXml = INXml & "<MERCHANT   value=""6011"" />"
INXml = INXml & "<COUNTRY_CODE   value="""" />"
INXml = INXml & "<POS_ENTRY_MODE     value=""901"" />"
INXml = INXml & "<POS_COND_CODE  value=""51"" />"
INXml = INXml & "<ACQ_INST   value=""457043"" />"
INXml = INXml & "<ISS_INST value="""" /><PAN_EXT    value="""" />"
INXml = INXml & "<TRACK2 value="""" />"
INXml = INXml & "<REFNUM  value=""710010185027"" />"
INXml = INXml & "<AUTH_CODE      value="""" />"
INXml = INXml & "<RESP_CODE    value="""" />"
INXml = INXml & "<TERMINAL_ID    value=""00008055"" />"
INXml = INXml & "<CARD_ACCEPTOR      value=""Citibank Russia "" />"
INXml = INXml & "<CARD_LOCATION  value=""SADOVAYACH 13/3 MOSCOW   RU "" />"
INXml = INXml & "<ADD_RESP_DATA      value="""" />"
INXml = INXml & "<CUR_CODE    value=""840"" />"
INXml = INXml & "<CUR_CODE_CARDISS   value="""" />"
INXml = INXml & "<PINBLOCK   value=""0000000000000000"" />"
INXml = INXml & "<ADD_POS_INFOvalue="""" />"
INXml = INXml & "<NET_INFvalue="""" />"
INXml = INXml & "<ORG_DATA   value="""" />"
INXml = INXml & "<REP_AMOUNT    value="""" />"
INXml = INXml & "<REQ_INST   value=""492491"" />"
INXml = INXml & "<ACCT_1     value=""23220020028920001"" />" '-> Cuenta de Ahorros
INXml = INXml & "<ACCT_2     value="" "" />"
INXml = INXml & "<CUSTOMER_INF_RESP  value="""" />"
INXml = INXml & "<PRIV_USE   value="""" />"
INXml = INXml & "</TXN_FIN_REQ>"
INXml = INXml & "</Messages>"

'***************************************************
'TRAMA 2
'*****************************************************

'************************************************
'TRAMA DE EXTORNO
'************************************************

INXmlExt = "<Messages> "
INXmlExt = INXmlExt & "<TXN_FIN_REQ>"
INXmlExt = INXmlExt & "<MESSAGE_TYPE   value=""400"" />"
INXmlExt = INXmlExt & "<PAN    value=""4924910000000001"" />"
INXmlExt = INXmlExt & "<PRCODE     value=""011000"" />"
INXmlExt = INXmlExt & "<TXN_AMOUNT value=""000000005000"" />"
INXmlExt = INXmlExt & "<CARDISS_AMOUNT  value="""" />"
INXmlExt = INXmlExt & "<TXN_DATE_TIME      value=""0410105413"" />"
INXmlExt = INXmlExt & "<CONVRATE   value="""" />"
INXmlExt = INXmlExt & "<TRACE  value=""" & sIDTRAMA & """" & " />"
INXmlExt = INXmlExt & "<TIME_LOCAl value=""055604"" />"
INXmlExt = INXmlExt & "<DATE_LOCAL     value=""0410"" />"
INXmlExt = INXmlExt & "<DATE_EXP   value=""0907"" />"
INXmlExt = INXmlExt & "<DATE_STTL   value=""0410"" />"
INXmlExt = INXmlExt & "<DATE_CAPTURE   value=""0410"" />"
INXmlExt = INXmlExt & "<MERCHANT   value=""6011"" />"
INXmlExt = INXmlExt & "<COUNTRY_CODE   value="""" />"
INXmlExt = INXmlExt & "<POS_ENTRY_MODE     value=""901"" />"
INXmlExt = INXmlExt & "<POS_COND_CODE  value=""51"" />"
INXmlExt = INXmlExt & "<ACQ_INST   value=""457043"" />"
INXmlExt = INXmlExt & "<ISS_INST value="""" /><PAN_EXT    value="""" />"
INXmlExt = INXmlExt & "<TRACK2 value="""" />"
INXmlExt = INXmlExt & "<REFNUM  value=""710010185027"" />"
INXmlExt = INXmlExt & "<AUTH_CODE      value="""" />"
INXmlExt = INXmlExt & "<RESP_CODE    value="""" />"
INXmlExt = INXmlExt & "<TERMINAL_ID    value=""00008055"" />"
INXmlExt = INXmlExt & "<CARD_ACCEPTOR      value=""Citibank Russia "" />"
INXmlExt = INXmlExt & "<CARD_LOCATION  value=""SADOVAYACH 13/3 MOSCOW   RU "" />"
INXmlExt = INXmlExt & "<ADD_RESP_DATA      value="""" />"
INXmlExt = INXmlExt & "<CUR_CODE    value=""840"" />"
INXmlExt = INXmlExt & "<CUR_CODE_CARDISS   value="""" />"
INXmlExt = INXmlExt & "<PINBLOCK   value=""0000000000000000"" />"
INXmlExt = INXmlExt & "<ADD_POS_INFOvalue="""" />"
INXmlExt = INXmlExt & "<NET_INFvalue="""" />"
INXmlExt = INXmlExt & "<ORG_DATA   value="""" />"
INXmlExt = INXmlExt & "<REP_AMOUNT    value="""" />"
INXmlExt = INXmlExt & "<REQ_INST   value=""492491"" />"
INXmlExt = INXmlExt & "<ACCT_1     value=""109012321000001422"" />" '-> Cuenta de Ahorros
INXmlExt = INXmlExt & "<ACCT_2     value="" "" />"
INXmlExt = INXmlExt & "<CUSTOMER_INF_RESP  value="""" />"
INXmlExt = INXmlExt & "<PRIV_USE   value="""" />"
INXmlExt = INXmlExt & "</TXN_FIN_REQ>"
INXmlExt = INXmlExt & "</Messages>"

'*******************************************************************
'TRAMA DE TRANSFERENCIA
'*******************************************************************
INXmlTransfer = "<Messages> "
INXmlTransfer = INXmlTransfer & "<TXN_FIN_REQ>"
INXmlTransfer = INXmlTransfer & "<MESSAGE_TYPE   value=""200"" />"
INXmlTransfer = INXmlTransfer & "<PAN    value=""4924910000000001"" />"
INXmlTransfer = INXmlTransfer & "<PRCODE     value=""401010"" />"
INXmlTransfer = INXmlTransfer & "<TXN_AMOUNT value=""000000005000"" />"
INXmlTransfer = INXmlTransfer & "<CARDISS_AMOUNT  value="""" />"
INXmlTransfer = INXmlTransfer & "<TXN_DATE_TIME      value=""0410105413"" />"
INXmlTransfer = INXmlTransfer & "<CONVRATE   value="""" />"
INXmlTransfer = INXmlTransfer & "<TRACE  value=""" & sIDTRAMA & """" & " />"
INXmlTransfer = INXmlTransfer & "<TIME_LOCAl value=""055604"" />"
INXmlTransfer = INXmlTransfer & "<DATE_LOCAL     value=""0410"" />"
INXmlTransfer = INXmlTransfer & "<DATE_EXP   value=""0907"" />"
INXmlTransfer = INXmlTransfer & "<DATE_STTL   value=""0410"" />"
INXmlTransfer = INXmlTransfer & "<DATE_CAPTURE   value=""0410"" />"
INXmlTransfer = INXmlTransfer & "<MERCHANT   value=""6011"" />"
INXmlTransfer = INXmlTransfer & "<COUNTRY_CODE   value="""" />"
INXmlTransfer = INXmlTransfer & "<POS_ENTRY_MODE     value=""901"" />"
INXmlTransfer = INXmlTransfer & "<POS_COND_CODE  value=""51"" />"
INXmlTransfer = INXmlTransfer & "<ACQ_INST   value=""457043"" />"
INXmlTransfer = INXmlTransfer & "<ISS_INST value="""" /><PAN_EXT    value="""" />"
INXmlTransfer = INXmlTransfer & "<TRACK2 value="""" />"
INXmlTransfer = INXmlTransfer & "<REFNUM  value=""710010185027"" />"
INXmlTransfer = INXmlTransfer & "<AUTH_CODE      value="""" />"
INXmlTransfer = INXmlTransfer & "<RESP_CODE    value="""" />"
INXmlTransfer = INXmlTransfer & "<TERMINAL_ID    value=""00008055"" />"
INXmlTransfer = INXmlTransfer & "<CARD_ACCEPTOR      value=""Citibank Russia "" />"
INXmlTransfer = INXmlTransfer & "<CARD_LOCATION  value=""SADOVAYACH 13/3 MOSCOW   RU "" />"
INXmlTransfer = INXmlTransfer & "<ADD_RESP_DATA      value="""" />"
INXmlTransfer = INXmlTransfer & "<CUR_CODE    value=""604"" />"
INXmlTransfer = INXmlTransfer & "<CUR_CODE_CARDISS   value="""" />"
INXmlTransfer = INXmlTransfer & "<PINBLOCK   value=""0000000000000000"" />"
INXmlTransfer = INXmlTransfer & "<ADD_POS_INFOvalue="""" />"
INXmlTransfer = INXmlTransfer & "<NET_INFvalue="""" />"
INXmlTransfer = INXmlTransfer & "<ORG_DATA   value="""" />"
INXmlTransfer = INXmlTransfer & "<REP_AMOUNT    value="""" />"
INXmlTransfer = INXmlTransfer & "<REQ_INST   value=""492491"" />"
INXmlTransfer = INXmlTransfer & "<ACCT_1     value=""109012321000001422"" />" '-> Cuenta de Ahorros
INXmlTransfer = INXmlTransfer & "<ACCT_2     value=""109012321000239526"" />"
INXmlTransfer = INXmlTransfer & "<CUSTOMER_INF_RESP  value="""" />"
INXmlTransfer = INXmlTransfer & "<PRIV_USE   value="""" />"
INXmlTransfer = INXmlTransfer & "</TXN_FIN_REQ>"
INXmlTransfer = INXmlTransfer & "</Messages>"


'*******************************************************************
'TRAMA DE TRANSFERENCIA DIFERENTE MONEDA
'*******************************************************************
INXmlTransfDifMon = "<Messages> "
INXmlTransfDifMon = INXmlTransfDifMon & "<TXN_FIN_REQ>"
INXmlTransfDifMon = INXmlTransfDifMon & "<MESSAGE_TYPE   value=""200"" />"
INXmlTransfDifMon = INXmlTransfDifMon & "<PAN    value=""4924910000000001"" />"
INXmlTransfDifMon = INXmlTransfDifMon & "<PRCODE     value=""401010"" />"
INXmlTransfDifMon = INXmlTransfDifMon & "<TXN_AMOUNT value=""000000005000"" />"
INXmlTransfDifMon = INXmlTransfDifMon & "<CARDISS_AMOUNT  value="""" />"
INXmlTransfDifMon = INXmlTransfDifMon & "<TXN_DATE_TIME      value=""0410105413"" />"
INXmlTransfDifMon = INXmlTransfDifMon & "<CONVRATE   value="""" />"
INXmlTransfDifMon = INXmlTransfDifMon & "<TRACE  value=""" & sIDTRAMA & """" & " />"
INXmlTransfDifMon = INXmlTransfDifMon & "<TIME_LOCAl value=""055604"" />"
INXmlTransfDifMon = INXmlTransfDifMon & "<DATE_LOCAL     value=""0410"" />"
INXmlTransfDifMon = INXmlTransfDifMon & "<DATE_EXP   value=""0907"" />"
INXmlTransfDifMon = INXmlTransfDifMon & "<DATE_STTL   value=""0410"" />"
INXmlTransfDifMon = INXmlTransfDifMon & "<DATE_CAPTURE   value=""0410"" />"
INXmlTransfDifMon = INXmlTransfDifMon & "<MERCHANT   value=""6011"" />"
INXmlTransfDifMon = INXmlTransfDifMon & "<COUNTRY_CODE   value="""" />"
INXmlTransfDifMon = INXmlTransfDifMon & "<POS_ENTRY_MODE     value=""901"" />"
INXmlTransfDifMon = INXmlTransfDifMon & "<POS_COND_CODE  value=""51"" />"
INXmlTransfDifMon = INXmlTransfDifMon & "<ACQ_INST   value=""457043"" />"
INXmlTransfDifMon = INXmlTransfDifMon & "<ISS_INST value="""" /><PAN_EXT    value="""" />"
INXmlTransfDifMon = INXmlTransfDifMon & "<TRACK2 value="""" />"
INXmlTransfDifMon = INXmlTransfDifMon & "<REFNUM  value=""710010185027"" />"
INXmlTransfDifMon = INXmlTransfDifMon & "<AUTH_CODE      value="""" />"
INXmlTransfDifMon = INXmlTransfDifMon & "<RESP_CODE    value="""" />"
INXmlTransfDifMon = INXmlTransfDifMon & "<TERMINAL_ID    value=""00008055"" />"
INXmlTransfDifMon = INXmlTransfDifMon & "<CARD_ACCEPTOR      value=""Citibank Russia "" />"
INXmlTransfDifMon = INXmlTransfDifMon & "<CARD_LOCATION  value=""SADOVAYACH 13/3 MOSCOW   RU "" />"
INXmlTransfDifMon = INXmlTransfDifMon & "<ADD_RESP_DATA      value="""" />"
INXmlTransfDifMon = INXmlTransfDifMon & "<CUR_CODE    value=""604"" />"
INXmlTransfDifMon = INXmlTransfDifMon & "<CUR_CODE_CARDISS   value="""" />"
INXmlTransfDifMon = INXmlTransfDifMon & "<PINBLOCK   value=""0000000000000000"" />"
INXmlTransfDifMon = INXmlTransfDifMon & "<ADD_POS_INFOvalue="""" />"
INXmlTransfDifMon = INXmlTransfDifMon & "<NET_INFvalue="""" />"
INXmlTransfDifMon = INXmlTransfDifMon & "<ORG_DATA   value="""" />"
INXmlTransfDifMon = INXmlTransfDifMon & "<REP_AMOUNT    value="""" />"
INXmlTransfDifMon = INXmlTransfDifMon & "<REQ_INST   value=""492491"" />"
INXmlTransfDifMon = INXmlTransfDifMon & "<ACCT_1     value=""109012321000001422"" />" '-> Cuenta de Ahorros
INXmlTransfDifMon = INXmlTransfDifMon & "<ACCT_2     value=""109012322000048836"" />"
INXmlTransfDifMon = INXmlTransfDifMon & "<CUSTOMER_INF_RESP  value="""" />"
INXmlTransfDifMon = INXmlTransfDifMon & "<PRIV_USE   value="""" />"
INXmlTransfDifMon = INXmlTransfDifMon & "</TXN_FIN_REQ>"
INXmlTransfDifMon = INXmlTransfDifMon & "</Messages>"


'*******************************************************************
'TRAMA DE EXTORNO DE TRANSFERENCIA
'*******************************************************************
INXmlExtTransf = "<Messages> "
INXmlExtTransf = INXmlExtTransf & "<TXN_FIN_REQ>"
INXmlExtTransf = INXmlExtTransf & "<MESSAGE_TYPE   value=""400"" />"
INXmlExtTransf = INXmlExtTransf & "<PAN    value=""4924910000000001"" />"
INXmlExtTransf = INXmlExtTransf & "<PRCODE     value=""401010"" />"
INXmlExtTransf = INXmlExtTransf & "<TXN_AMOUNT value=""000000005000"" />"
INXmlExtTransf = INXmlExtTransf & "<CARDISS_AMOUNT  value="""" />"
INXmlExtTransf = INXmlExtTransf & "<TXN_DATE_TIME      value=""0410105413"" />"
INXmlExtTransf = INXmlExtTransf & "<CONVRATE   value="""" />"
INXmlExtTransf = INXmlExtTransf & "<TRACE  value=""" & sIDTRAMA & """" & " />"
INXmlExtTransf = INXmlExtTransf & "<TIME_LOCAl value=""055604"" />"
INXmlExtTransf = INXmlExtTransf & "<DATE_LOCAL     value=""0410"" />"
INXmlExtTransf = INXmlExtTransf & "<DATE_EXP   value=""0907"" />"
INXmlExtTransf = INXmlExtTransf & "<DATE_STTL   value=""0410"" />"
INXmlExtTransf = INXmlExtTransf & "<DATE_CAPTURE   value=""0410"" />"
INXmlExtTransf = INXmlExtTransf & "<MERCHANT   value=""6011"" />"
INXmlExtTransf = INXmlExtTransf & "<COUNTRY_CODE   value="""" />"
INXmlExtTransf = INXmlExtTransf & "<POS_ENTRY_MODE     value=""901"" />"
INXmlExtTransf = INXmlExtTransf & "<POS_COND_CODE  value=""51"" />"
INXmlExtTransf = INXmlExtTransf & "<ACQ_INST   value=""457043"" />"
INXmlExtTransf = INXmlExtTransf & "<ISS_INST value="""" /><PAN_EXT    value="""" />"
INXmlExtTransf = INXmlExtTransf & "<TRACK2 value="""" />"
INXmlExtTransf = INXmlExtTransf & "<REFNUM  value=""710010185027"" />"
INXmlExtTransf = INXmlExtTransf & "<AUTH_CODE      value="""" />"
INXmlExtTransf = INXmlExtTransf & "<RESP_CODE    value="""" />"
INXmlExtTransf = INXmlExtTransf & "<TERMINAL_ID    value=""00008055"" />"
INXmlExtTransf = INXmlExtTransf & "<CARD_ACCEPTOR      value=""Citibank Russia "" />"
INXmlExtTransf = INXmlExtTransf & "<CARD_LOCATION  value=""SADOVAYACH 13/3 MOSCOW   RU "" />"
INXmlExtTransf = INXmlExtTransf & "<ADD_RESP_DATA      value="""" />"
INXmlExtTransf = INXmlExtTransf & "<CUR_CODE    value=""604"" />"
INXmlExtTransf = INXmlExtTransf & "<CUR_CODE_CARDISS   value="""" />"
INXmlExtTransf = INXmlExtTransf & "<PINBLOCK   value=""0000000000000000"" />"
INXmlExtTransf = INXmlExtTransf & "<ADD_POS_INFOvalue="""" />"
INXmlExtTransf = INXmlExtTransf & "<NET_INFvalue="""" />"
INXmlExtTransf = INXmlExtTransf & "<ORG_DATA   value="""" />"
INXmlExtTransf = INXmlExtTransf & "<REP_AMOUNT    value="""" />"
INXmlExtTransf = INXmlExtTransf & "<REQ_INST   value=""492491"" />"
INXmlExtTransf = INXmlExtTransf & "<ACCT_1     value=""109012321000001422"" />" '-> Cuenta de Ahorros
INXmlExtTransf = INXmlExtTransf & "<ACCT_2     value=""109012322000048836"" />"
INXmlExtTransf = INXmlExtTransf & "<CUSTOMER_INF_RESP  value="""" />"
INXmlExtTransf = INXmlExtTransf & "<PRIV_USE   value="""" />"
INXmlExtTransf = INXmlExtTransf & "</TXN_FIN_REQ>"
INXmlExtTransf = INXmlExtTransf & "</Messages>"

'******************************************************************
'TRAMA CONSULTA DE SALDO
'******************************************************************
INXmlConsCta = INXmlConsCta & "<?xml version=""1.0""?> <Messages>"
INXmlConsCta = INXmlConsCta & "<TXN_FIN_REQ>"
INXmlConsCta = INXmlConsCta & "<MESSAGE_TYPE   value=""0200"" />"
INXmlConsCta = INXmlConsCta & "<PAN    value= ""8109000000000060"" />"
INXmlConsCta = INXmlConsCta & "<PRCODE     value=""311000"" />"
INXmlConsCta = INXmlConsCta & "<TXN_AMOUNT value=""000000000000"" />"
INXmlConsCta = INXmlConsCta & "<CARDISS_AMOUNT  value="""" />"
INXmlConsCta = INXmlConsCta & "<TXN_DATE_TIME      value=""0411165142"" />"
INXmlConsCta = INXmlConsCta & "<CONVRATE   value="""" />"
INXmlConsCta = INXmlConsCta & "<TRACE  value=""045986"" />"
INXmlConsCta = INXmlConsCta & "<TIME_LOCAl value=""170032"" />"
INXmlConsCta = INXmlConsCta & "<DATE_LOCAL     value=""0411"" />"
INXmlConsCta = INXmlConsCta & "<DATE_EXP   value=""1507"" />"
INXmlConsCta = INXmlConsCta & "<DATE_STTL   value=""0411"" />"
INXmlConsCta = INXmlConsCta & "<DATE_CAPTURE   value=""0411"" />"
INXmlConsCta = INXmlConsCta & "<MERCHANT   value=""6011"" />"
INXmlConsCta = INXmlConsCta & "<COUNTRY_CODE   value="""" />"
INXmlConsCta = INXmlConsCta & "<POS_ENTRY_MODE     value=""901"" />"
INXmlConsCta = INXmlConsCta & "<POS_COND_CODE  value=""51"" />"
INXmlConsCta = INXmlConsCta & "<ACQ_INST   value=""457043"" />"
INXmlConsCta = INXmlConsCta & "<ISS_INST value="""" />"
INXmlConsCta = INXmlConsCta & "<PAN_EXT    value="""" />"
INXmlConsCta = INXmlConsCta & "<TRACK2 value="""" />"
INXmlConsCta = INXmlConsCta & "<REFNUM  value=""710010185027"" />"
INXmlConsCta = INXmlConsCta & "<AUTH_CODE      value="""" />"
INXmlConsCta = INXmlConsCta & "<RESP_CODE    value="""" />"
INXmlConsCta = INXmlConsCta & "<TERMINAL_ID    value=""00000000"" />"
INXmlConsCta = INXmlConsCta & "<CARD_ACCEPTOR      value=""TDA.052 CENTROB "" />"
INXmlConsCta = INXmlConsCta & "<CARD_LOCATION  value=""CENTROBANK SNTA.ANITA IIILIMA     PE "" />"""
INXmlConsCta = INXmlConsCta & "<ADD_RESP_DATA      value="""" />"
INXmlConsCta = INXmlConsCta & "<CUR_CODE    value=""604"" />"
INXmlConsCta = INXmlConsCta & "<CUR_CODE_CARDISS   value="""" />"
INXmlConsCta = INXmlConsCta & "<PINBLOCK   value=""0000000000000000"" />"
INXmlConsCta = INXmlConsCta & "<ADD_POS_INFOvalue="""" />"
INXmlConsCta = INXmlConsCta & "<NET_INF value="""" />"
INXmlConsCta = INXmlConsCta & "<ORG_DATA   value="""" />"
INXmlConsCta = INXmlConsCta & "<REP_AMOUNT    value="""" />"
INXmlConsCta = INXmlConsCta & "<REQ_INST   value=""491249"" />"
INXmlConsCta = INXmlConsCta & "<ACCT_1     value=""2321000000019 26154    001"" />"
INXmlConsCta = INXmlConsCta & "<ACCT_2     value="" "" />"
INXmlConsCta = INXmlConsCta & "<CUSTOMER_INF_RESP  value="""" />"
INXmlConsCta = INXmlConsCta & "<PRIV_USE   value="""" />"
INXmlConsCta = INXmlConsCta & "</TXN_FIN_REQ>"
INXmlConsCta = INXmlConsCta & "</Messages>"

'*****************************************************************************************
'CONSULTA INTEGRADA
'*****************************************************************************************
INXmlConsInteg = "<?xml version=""1.0""?> <Messages>"
INXmlConsInteg = INXmlConsInteg & "<TXN_FIN_REQ>"
INXmlConsInteg = INXmlConsInteg & "<MESSAGE_TYPE   value=""200"" />"
INXmlConsInteg = INXmlConsInteg & "<PAN    value= ""4924910000000001"" />"
INXmlConsInteg = INXmlConsInteg & "<PRCODE     value=""930099"" />"
INXmlConsInteg = INXmlConsInteg & "<TXN_AMOUNT value=""000000000000"" />"
INXmlConsInteg = INXmlConsInteg & "<CARDISS_AMOUNT  value="""" />"
INXmlConsInteg = INXmlConsInteg & "<TXN_DATE_TIME      value=""0405080109"" />"
INXmlConsInteg = INXmlConsInteg & "<CONVRATE   value="""" />"
INXmlConsInteg = INXmlConsInteg & "<TRACE  value=""185027"" />"
INXmlConsInteg = INXmlConsInteg & "<TIME_LOCAl value=""080109"" />"
INXmlConsInteg = INXmlConsInteg & "<DATE_LOCAL     value=""0405"" />"
INXmlConsInteg = INXmlConsInteg & "<DATE_EXP   value=""1405"" />"
INXmlConsInteg = INXmlConsInteg & "<DATE_STTL   value=""0405"" />"
INXmlConsInteg = INXmlConsInteg & "<DATE_CAPTURE   value=""0405"" />"
INXmlConsInteg = INXmlConsInteg & "<MERCHANT   value=""0000"" />"
INXmlConsInteg = INXmlConsInteg & "<COUNTRY_CODE   value="""" />"
INXmlConsInteg = INXmlConsInteg & "<POS_ENTRY_MODE     value=""000"" />"
INXmlConsInteg = INXmlConsInteg & "<POS_COND_CODE  value=""02"" />"
INXmlConsInteg = INXmlConsInteg & "<ACQ_INST   value=""522301"" />"
INXmlConsInteg = INXmlConsInteg & "<ISS_INST value="""" />"
INXmlConsInteg = INXmlConsInteg & "<PAN_EXT    value="""" />"
INXmlConsInteg = INXmlConsInteg & "<TRACK2 value="""" />"
INXmlConsInteg = INXmlConsInteg & "<REFNUM  value=""000000000000"" />"
INXmlConsInteg = INXmlConsInteg & "<AUTH_CODE      value="""" />"
INXmlConsInteg = INXmlConsInteg & "<RESP_CODE    value="""" />"
INXmlConsInteg = INXmlConsInteg & "<TERMINAL_ID    value=""00000251"" />"
INXmlConsInteg = INXmlConsInteg & "<CARD_ACCEPTOR      value=""000000000000000"" />"
INXmlConsInteg = INXmlConsInteg & "<CARD_LOCATION  value=""RED UNICARD  C.C.AURORA 156160 "" />"
INXmlConsInteg = INXmlConsInteg & "<ADD_RESP_DATA      value="""" />"
INXmlConsInteg = INXmlConsInteg & "<CUR_CODE    value=""604"" />"
INXmlConsInteg = INXmlConsInteg & "<CUR_CODE_CARDISS   value="""" />"
INXmlConsInteg = INXmlConsInteg & "<PINBLOCK   value=""BAC9FA5BBA75AC0A"" />"
INXmlConsInteg = INXmlConsInteg & "<ADD_POS_INFOvalue="""" />"
INXmlConsInteg = INXmlConsInteg & "<NET_INF value="""" />"
INXmlConsInteg = INXmlConsInteg & "<ORG_DATA   value="""" />"
INXmlConsInteg = INXmlConsInteg & "<REP_AMOUNT    value="""" />"
INXmlConsInteg = INXmlConsInteg & "<REQ_INST   value=""492491"" />"
INXmlConsInteg = INXmlConsInteg & "<ACCT_1     value=""2321000240958            001"" />"
INXmlConsInteg = INXmlConsInteg & "<ACCT_2     value="" "" />"
INXmlConsInteg = INXmlConsInteg & "<CUSTOMER_INF_RESP  value="""" />"
INXmlConsInteg = INXmlConsInteg & "<PRIV_USE   value="""" /> </TXN_FIN_REQ>"
INXmlConsInteg = INXmlConsInteg & "</Messages>"

'***********************************************************************************
'CONSULTA DE ULTIMOS MOVIMIENTOS
'***********************************************************************************

INXmlConsMov = "<?xml version=""1.0""?> <Messages>"
INXmlConsMov = INXmlConsMov & "<TXN_FIN_REQ>"
INXmlConsMov = INXmlConsMov & "<MESSAGE_TYPE   value=""200"" />"
INXmlConsMov = INXmlConsMov & "<PAN    value=""4261540000000127"" />"
INXmlConsMov = INXmlConsMov & "<PRCODE     value=""391010"" />"
INXmlConsMov = INXmlConsMov & "<TXN_AMOUNT value=""000000000000"" />"
INXmlConsMov = INXmlConsMov & "<CARDISS_AMOUNT  value="""" />"
INXmlConsMov = INXmlConsMov & "<TXN_DATE_TIME      value=""0410084938"" />"
INXmlConsMov = INXmlConsMov & "<CONVRATE   value="""" />"
INXmlConsMov = INXmlConsMov & "<TRACE  value=""045986"" />"
INXmlConsMov = INXmlConsMov & "<TIME_LOCAl value=""084938"" />"
INXmlConsMov = INXmlConsMov & "<DATE_LOCAL     value=""0410"" />"
INXmlConsMov = INXmlConsMov & "<DATE_EXP   value=""1501"" />"
INXmlConsMov = INXmlConsMov & "<DATE_STTL   value=""0410"" />"
INXmlConsMov = INXmlConsMov & "<DATE_CAPTURE   value=""0410"" />"
INXmlConsMov = INXmlConsMov & "<MERCHANT   value=""0000"" />"
INXmlConsMov = INXmlConsMov & "<COUNTRY_CODE   value="""" />"
INXmlConsMov = INXmlConsMov & "<POS_ENTRY_MODE     value=""000"" />"
INXmlConsMov = INXmlConsMov & "<POS_COND_CODE  value=""02"" />"
INXmlConsMov = INXmlConsMov & "<ACQ_INST   value=""522301"" />"
INXmlConsMov = INXmlConsMov & "<ISS_INST value="""" />"
INXmlConsMov = INXmlConsMov & "<PAN_EXT    value="""" />"
INXmlConsMov = INXmlConsMov & "<TRACK2 value="""" />"
INXmlConsMov = INXmlConsMov & "<REFNUM  value=""000000000000"" />"
INXmlConsMov = INXmlConsMov & "<AUTH_CODE      value="""" />"
INXmlConsMov = INXmlConsMov & "<RESP_CODE    value="""" />"
INXmlConsMov = INXmlConsMov & "<TERMINAL_ID    value=""00000257"" />"
INXmlConsMov = INXmlConsMov & "<CARD_ACCEPTOR      value=""000000000000000"" />"
INXmlConsMov = INXmlConsMov & "<CARD_LOCATION  value=""RED UNICARD AV.FAUCETT 525 S.M. "" />"
INXmlConsMov = INXmlConsMov & "<ADD_RESP_DATA      value="""" />"
INXmlConsMov = INXmlConsMov & "<CUR_CODE    value=""604"" />"
INXmlConsMov = INXmlConsMov & "<CUR_CODE_CARDISS   value="""" />"
INXmlConsMov = INXmlConsMov & "<PINBLOCK   value=""FF0992CF47CB42F2"" />"
INXmlConsMov = INXmlConsMov & "<ADD_POS_INFOvalue="""" />"
INXmlConsMov = INXmlConsMov & "<NET_INF value="""" />"
INXmlConsMov = INXmlConsMov & "<ORG_DATA   value="""" />"
INXmlConsMov = INXmlConsMov & "<REP_AMOUNT    value="""" />"
INXmlConsMov = INXmlConsMov & "<REQ_INST   value=""491249"" />"
INXmlConsMov = INXmlConsMov & "<ACCT_1     value=""2321000000019 26154    001"" />"
INXmlConsMov = INXmlConsMov & "<ACCT_2     value="" "" />"
INXmlConsMov = INXmlConsMov & "<CUSTOMER_INF_RESP  value="""" />"
INXmlConsMov = INXmlConsMov & "<PRIV_USE   value="""" />"
INXmlConsMov = INXmlConsMov & "</TXN_FIN_REQ>"
INXmlConsMov = INXmlConsMov & "</Messages>"

'*******************************************************************************************
'CONSULTA DE TIPO CAMBIO
'*******************************************************************************************
INXmlConsTipCam = "<?xml version=""1.0""?> <Messages>"
INXmlConsTipCam = INXmlConsTipCam & "<TXN_FIN_REQ>"
INXmlConsTipCam = INXmlConsTipCam & "<MESSAGE_TYPE   value=""200"" />"
INXmlConsTipCam = INXmlConsTipCam & "<PAN    value=""4924910000000001"" />"
INXmlConsTipCam = INXmlConsTipCam & "<PRCODE     value=""980000"" />"
INXmlConsTipCam = INXmlConsTipCam & "<TXN_AMOUNT value=""000000000000"" />"
INXmlConsTipCam = INXmlConsTipCam & "<CARDISS_AMOUNT  value="""" />"
INXmlConsTipCam = INXmlConsTipCam & "<TXN_DATE_TIME      value=""0602112533"" />"
INXmlConsTipCam = INXmlConsTipCam & "<CONVRATE   value="""" />"
INXmlConsTipCam = INXmlConsTipCam & "<TRACE  value=""045986"" />"
INXmlConsTipCam = INXmlConsTipCam & "<TIME_LOCAl value=""170032"" />"
INXmlConsTipCam = INXmlConsTipCam & "<DATE_LOCAL     value=""0602"" />"
INXmlConsTipCam = INXmlConsTipCam & "<DATE_EXP   value=""1507"" />"
INXmlConsTipCam = INXmlConsTipCam & "<DATE_STTL   value=""0602"" />"
INXmlConsTipCam = INXmlConsTipCam & "<DATE_CAPTURE   value=""0602"" />"
INXmlConsTipCam = INXmlConsTipCam & "<MERCHANT   value=""0000"" />"
INXmlConsTipCam = INXmlConsTipCam & "<COUNTRY_CODE   value="""" />"
INXmlConsTipCam = INXmlConsTipCam & "<POS_ENTRY_MODE     value=""000"" />"
INXmlConsTipCam = INXmlConsTipCam & "<POS_COND_CODE  value=""02"" />"
INXmlConsTipCam = INXmlConsTipCam & "<ACQ_INST   value=""458102"" />"
INXmlConsTipCam = INXmlConsTipCam & "<ISS_INST value="""" />"
INXmlConsTipCam = INXmlConsTipCam & "<PAN_EXT    value="""" />"
INXmlConsTipCam = INXmlConsTipCam & "<TRACK2 value="""" />"
INXmlConsTipCam = INXmlConsTipCam & "<REFNUM  value=""000000000000"" />"
INXmlConsTipCam = INXmlConsTipCam & "<AUTH_CODE      value="""" />"
INXmlConsTipCam = INXmlConsTipCam & "<RESP_CODE    value="""" />"
INXmlConsTipCam = INXmlConsTipCam & "<TERMINAL_ID    value=""00000268"" />"
INXmlConsTipCam = INXmlConsTipCam & "<CARD_ACCEPTOR      value=""000000000000000"" />"
INXmlConsTipCam = INXmlConsTipCam & "<CARD_LOCATION  value=""RED UNICARD SAN BORJA NORTE 996"" />"
INXmlConsTipCam = INXmlConsTipCam & "<ADD_RESP_DATA      value="""" />"
INXmlConsTipCam = INXmlConsTipCam & "<CUR_CODE    value=""840"" />"
INXmlConsTipCam = INXmlConsTipCam & "<CUR_CODE_CARDISS   value="""" />"
INXmlConsTipCam = INXmlConsTipCam & "<PINBLOCK   value=""1B873464A8886456"" />"
INXmlConsTipCam = INXmlConsTipCam & "<ADD_POS_INFOvalue="""" />"
INXmlConsTipCam = INXmlConsTipCam & "<NET_INF value="""" />"
INXmlConsTipCam = INXmlConsTipCam & "<ORG_DATA   value="""" />"
INXmlConsTipCam = INXmlConsTipCam & "<REP_AMOUNT    value="""" />"
INXmlConsTipCam = INXmlConsTipCam & "<REQ_INST   value=""491249"" />"
INXmlConsTipCam = INXmlConsTipCam & "<ACCT_1     value=""2321000240958            001"" />"
INXmlConsTipCam = INXmlConsTipCam & "<ACCT_2     value="" "" />"
INXmlConsTipCam = INXmlConsTipCam & "<CUSTOMER_INF_RESP  value="""" />"
INXmlConsTipCam = INXmlConsTipCam & "<PRIV_USE   value="""" /> </TXN_FIN_REQ>"
INXmlConsTipCam = INXmlConsTipCam & "</Messages>"

'***********************************************************************************
'CAMBIO DE CLAVE
'***********************************************************************************
INXmlCambClave = "<?xml version=""1.0""?>"
INXmlCambClave = INXmlCambClave & "<Messages>"
INXmlCambClave = INXmlCambClave & "<TXN_FIN_REQ>"
INXmlCambClave = INXmlCambClave & "<MESSAGE_TYPE   value=""0200"" />"
INXmlCambClave = INXmlCambClave & "<PAN    value=""4261540000000142"" />"
INXmlCambClave = INXmlCambClave & "<PRCODE     value=""910000"" />"
INXmlCambClave = INXmlCambClave & "<TXN_AMOUNT value=""000000000000"" />"
INXmlCambClave = INXmlCambClave & "<CARDISS_AMOUNT  value="""" />"
INXmlCambClave = INXmlCambClave & "<TXN_DATE_TIME      value=""0405081926"" />"
INXmlCambClave = INXmlCambClave & "<CONVRATE   value="""" />"
INXmlCambClave = INXmlCambClave & "<TRACE  value=""045986"" />"
INXmlCambClave = INXmlCambClave & "<TIME_LOCAl value=""170032"" />"
INXmlCambClave = INXmlCambClave & "<DATE_LOCAL     value=""0405"" />"
INXmlCambClave = INXmlCambClave & "<DATE_EXP   value=""1011"" />"
INXmlCambClave = INXmlCambClave & "<DATE_STTL   value=""0405"" />"
INXmlCambClave = INXmlCambClave & "<DATE_CAPTURE   value=""0405"" />"
INXmlCambClave = INXmlCambClave & "<MERCHANT   value=""0000"" />"
INXmlCambClave = INXmlCambClave & "<COUNTRY_CODE   value="""" />"
INXmlCambClave = INXmlCambClave & "<POS_ENTRY_MODE     value=""000"" />"
INXmlCambClave = INXmlCambClave & "<POS_COND_CODE  value=""02"" />"
INXmlCambClave = INXmlCambClave & "<ACQ_INST   value=""492491"" />"
INXmlCambClave = INXmlCambClave & "<ISS_INST value="""" />"
INXmlCambClave = INXmlCambClave & "<PAN_EXT    value="""" />"
INXmlCambClave = INXmlCambClave & "<TRACK2 value="""" />"
INXmlCambClave = INXmlCambClave & "<REFNUM  value=""000000000000"" />"
INXmlCambClave = INXmlCambClave & "<AUTH_CODE      value="""" />"
INXmlCambClave = INXmlCambClave & "<RESP_CODE    value="""" />"
INXmlCambClave = INXmlCambClave & "<TERMINAL_ID    value=""00000279"" />"
INXmlCambClave = INXmlCambClave & "<CARD_ACCEPTOR      value=""000000000000000"" />"
INXmlCambClave = INXmlCambClave & "<CARD_LOCATION  value=""RED UNICARD SAN BORJA NORTE 996"" />"
INXmlCambClave = INXmlCambClave & "<ADD_RESP_DATA      value="""" />"
INXmlCambClave = INXmlCambClave & "<CUR_CODE    value=""604"" />"
INXmlCambClave = INXmlCambClave & "<CUR_CODE_CARDISS   value="""" />"
INXmlCambClave = INXmlCambClave & "<PINBLOCK   value=""83420FFFFFFFFFFF"" />"
INXmlCambClave = INXmlCambClave & "<ADD_POS_INFOvalue="""" />"
INXmlCambClave = INXmlCambClave & "<NET_INF value="""" />"
INXmlCambClave = INXmlCambClave & "<ORG_DATA   value="""" />"
INXmlCambClave = INXmlCambClave & "<REP_AMOUNT    value="""" />"
INXmlCambClave = INXmlCambClave & "<REQ_INST   value=""492491"" />"
INXmlCambClave = INXmlCambClave & "<ACCT_1     value=""2321000000019 26154    001"" />"
INXmlCambClave = INXmlCambClave & "<ACCT_2     value="" "" />"
INXmlCambClave = INXmlCambClave & "<CUSTOMER_INF_RESP  value="""" />"
INXmlCambClave = INXmlCambClave & "<PRIV_USE   value="""" />"
INXmlCambClave = INXmlCambClave & "</TXN_FIN_REQ>"
INXmlCambClave = INXmlCambClave & "</Messages>"


End Sub
