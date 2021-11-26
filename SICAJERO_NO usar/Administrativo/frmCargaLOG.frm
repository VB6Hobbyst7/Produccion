VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmCargaLOG 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Conciliacion:Cargar Archivo LOG"
   ClientHeight    =   1860
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7470
   Icon            =   "frmCargaLOG.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1860
   ScaleWidth      =   7470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog dlgArchivo 
      Left            =   570
      Top             =   1485
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCargar 
      Caption         =   "Cargar"
      Height          =   360
      Left            =   4890
      TabIndex        =   8
      Top             =   1470
      Width           =   1230
   End
   Begin MSComctlLib.ProgressBar pB 
      Height          =   195
      Left            =   90
      TabIndex        =   7
      Top             =   1185
      Width           =   7305
      _ExtentX        =   12885
      _ExtentY        =   344
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Frame fraArchivoLOG 
      Caption         =   "Archivo"
      Height          =   1080
      Left            =   60
      TabIndex        =   1
      Top             =   30
      Width           =   7350
      Begin MSMask.MaskEdBox mskFecha 
         Height          =   300
         Left            =   765
         TabIndex        =   4
         Top             =   660
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "..."
         Height          =   345
         Left            =   6870
         TabIndex        =   3
         Top             =   285
         Width           =   420
      End
      Begin VB.TextBox txtArchivo 
         Height          =   315
         Left            =   765
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   300
         Width           =   6045
      End
      Begin VB.Label lblRura 
         Caption         =   "Archivo :"
         Height          =   180
         Left            =   75
         TabIndex        =   6
         Top             =   375
         Width           =   795
      End
      Begin VB.Label lblFecha 
         Caption         =   "Fecha :"
         Height          =   180
         Left            =   75
         TabIndex        =   5
         Top             =   720
         Width           =   795
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   360
      Left            =   6165
      TabIndex        =   0
      Top             =   1455
      Width           =   1230
   End
End
Attribute VB_Name = "frmCargaLOG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Cargar()
Dim Cmd As New ADODB.Command
Dim Prm As New ADODB.Parameter
Dim sResp As String

Dim lsRuta As String
Dim cad As String
Dim i As Integer
Dim TIPO_REG, LNET_TID_PSE, INST_ID_PSE, TERM_ID_PSE, LNET_ID_EMISOR, INST_ID_EMISOR As String
Dim PAN, COD_SUCURSAL, COD_REGION, TIPO_TRAN, TIPO_MSG, STATUS_MSG, ORIGEN_TRAN, ORIGEN_MSG As String
Dim ENTRY_TIME, EXIT_TIME, RE_ENTRY_TIME, TRAN_DATE, TRAN_TIME, PROC_DATE, ACQ_PROC_DATE, ISS_PROC_DATE As String
Dim NUM_TRACE, TIPO_TERMINAL, OFFSET_TIME, NROUTE_ACQ, NROUTE_ISS, TRAN_CODE, TIPO_CTA_FROM, TIPO_CTA_TO As String
Dim NRO_CTA_FROM, SOSPECHA_EXTORNO, NRO_CTA_TO, IND_MULT_ACCT, TIPO_DEPOSITO, IND_RETENCION, COD_RESP As String
Dim DIRECCION_TERM, NOM_INST_TERM, CIUDAD_TERM, ESTADO_TERM, PAIS_TERM, NUM_SEQ_ORIG, DATE_TRAN_ORIG As String
Dim TIME_TRAN_ORIG, DATE_PROC_ORIG, COD_MON_ORIG, COD_MON_AUTH, TIPO_CAMBIO_AUTH, COD_MON_CARGO As String
Dim TIPO_CAMBIO_CARGO, TCAM_DATE_TIME, IND_MOTIVO_EXT, SHARING_GRUOP, DEST_ORDER, HOST_COD_AUTH As String
Dim FORWARD_INST_ID, CARD_ACQ_ID, CARD_ISS_ID, ID_TOKEN, LONG_TOKEN, FILLER1, COD_MON_SOLIC As String
Dim COD_MON_FROM_ACCT, COD_MON_TO_ACCT, FILLER2 As String
Dim IMPORTE1, IMPORTE2, IMPORTE3, SALDO_CRED_DEP As Double
Dim US_IMPORTE1, US_IMPORTE2, US_IMPORTE3 As Double
Dim TIPO_CAMBIO As String
Dim loConec As New DConecta

If Not IsDate(Me.mskFecha.Text) Then
    MsgBox "Fecha no valida.", vbInformation, "Aviso"
    mskFecha.SetFocus
    Exit Sub
End If

Dim rsF As New ADODB.Recordset
loConec.AbreConexion
rsF.Open "dbo.ATM_ConsultaFechaLOG '" & Format(CDate(Me.mskFecha.Text), "YYYY-MM-DD") & "'", loConec.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText


If Not (rsF.EOF And rsF.BOF) Then
    MsgBox "No puede insertar un archivo LOG con la misma fecha dos veces", vbInformation, "Aviso"
    Exit Sub
End If


lsRuta = Me.txtArchivo.Text
i = -1

Dim ban As Boolean
Dim cadenita As String
Dim lnTotal As Integer
lnTotal = 0

Open lsRuta For Input As #1
Do Until EOF(1)
    Input #1, cad
'    MsgBox cad
    lnTotal = lnTotal + 1
    If lnTotal = 1 Then
        If Format(CDate(Me.mskFecha.Text), "YYYYMMDD") <> Trim(Left(cad, 8)) Then
            MsgBox "Fecha no valida. La fecha del archivo LOG es : " & cad, vbInformation, "Aviso"
            Close #1
            mskFecha.SetFocus
            Exit Sub
        End If
    End If
Loop
Close #1

Me.pB.Max = lnTotal + 2
Me.pB.Min = 0
Me.pB.Value = 0

ban = True
 Open lsRuta For Input As #1
    Rem Recuerda que Fichero.txt será la lista de componentes
    Do Until EOF(1)
        Set Cmd = New ADODB.Command
        
        Input #1, cad
        
        Me.pB.Value = Me.pB.Value + 1
        
        If Not ban Then
            Dim CADENA As String
            loConec.AbreConexion
            Cmd.ActiveConnection = loConec.ConexionActiva ' AbrirConexion
            'ConflictDetection = "OverwriteChanges"
            TIPO_REG = Mid(cad, 1, 2)
            LNET_TID_PSE = Mid(cad, 3, 4)
            INST_ID_PSE = Mid(cad, 7, 4)
            TERM_ID_PSE = Mid(cad, 11, 16)
            LNET_ID_EMISOR = Mid(cad, 27, 4)
            INST_ID_EMISOR = Mid(cad, 31, 4)
            PAN = Mid(cad, 35, 19)
            COD_SUCURSAL = Mid(cad, 54, 4)
            COD_REGION = Mid(cad, 58, 4)
            TIPO_TRAN = Mid(cad, 62, 2)
            TIPO_MSG = Mid(cad, 64, 4)
            STATUS_MSG = Mid(cad, 68, 2)
            ORIGEN_TRAN = Mid(cad, 70, 1)
            ORIGEN_MSG = Mid(cad, 71, 1)
            ENTRY_TIME = Mid(cad, 72, 19)
            EXIT_TIME = Mid(cad, 91, 19)
            'MsgBox EXIT_TIME
            RE_ENTRY_TIME = Mid(cad, 110, 19)
            'MsgBox RE_ENTRY_TIME
            TRAN_DATE = Mid(cad, 129, 6)
            TRAN_TIME = Mid(cad, 135, 8)
            PROC_DATE = Mid(cad, 143, 6)
            ACQ_PROC_DATE = Mid(cad, 149, 6)
            'MsgBox "ACQ_PROC_DATE" & ACQ_PROC_DATE
            ISS_PROC_DATE = Mid(cad, 155, 6)
            NUM_TRACE = Mid(cad, 161, 12)
            TIPO_TERMINAL = Mid(cad, 173, 2)
            OFFSET_TIME = Mid(cad, 175, 5)
            NROUTE_ACQ = Mid(cad, 180, 11)
            NROUTE_ISS = Mid(cad, 191, 11)
            TRAN_CODE = Mid(cad, 202, 2)
            TIPO_CTA_FROM = Mid(cad, 204, 2)
            TIPO_CTA_TO = Mid(cad, 206, 2)
            NRO_CTA_FROM = Mid(cad, 208, 19)
            SOSPECHA_EXTORNO = Mid(cad, 227, 1)
            NRO_CTA_TO = Mid(cad, 228, 19)
            IND_MULT_ACCT = Mid(cad, 247, 1)
            IMPORTE1 = CDbl(Mid(cad, 248, 15)) / 100
            IMPORTE2 = CDbl(Mid(cad, 263, 15)) / 100
            IMPORTE3 = CDbl(Mid(cad, 278, 15)) / 100
            SALDO_CRED_DEP = CDbl(Mid(cad, 293, 10))
            TIPO_DEPOSITO = Mid(cad, 303, 1)
            IND_RETENCION = Mid(cad, 304, 1)
            COD_RESP = Mid(cad, 305, 2)
            DIRECCION_TERM = Mid(cad, 307, 25)
            NOM_INST_TERM = Mid(cad, 332, 22)
            CIUDAD_TERM = Mid(cad, 354, 13)
            ESTADO_TERM = Mid(cad, 367, 3)
            PAIS_TERM = Mid(cad, 370, 2)
            NUM_SEQ_ORIG = Mid(cad, 372, 12)
            DATE_TRAN_ORIG = Mid(cad, 384, 4)
            TIME_TRAN_ORIG = Mid(cad, 388, 8)
            DATE_PROC_ORIG = Mid(cad, 396, 4)
            COD_MON_ORIG = Mid(cad, 400, 3)
            COD_MON_AUTH = Mid(cad, 403, 3)
            TIPO_CAMBIO_AUTH = CDbl(Mid(cad, 406, 8))
            COD_MON_CARGO = Mid(cad, 414, 3)
            TIPO_CAMBIO_CARGO = Mid(cad, 417, 8)
            TCAM_DATE_TIME = Mid(cad, 425, 19)
            IND_MOTIVO_EXT = Mid(cad, 444, 2)
            SHARING_GRUOP = Mid(cad, 446, 1)
            DEST_ORDER = Mid(cad, 447, 1)
            HOST_COD_AUTH = Mid(cad, 448, 6)
            FORWARD_INST_ID = Mid(cad, 454, 11)
            CARD_ACQ_ID = Mid(cad, 465, 11)
            CARD_ISS_ID = Mid(cad, 476, 11)
            ID_TOKEN = Mid(cad, 487, 2)
            LONG_TOKEN = Mid(cad, 489, 5)
            FILLER1 = Mid(cad, 494, 1)
            COD_MON_SOLIC = Mid(cad, 495, 3)
            COD_MON_FROM_ACCT = Mid(cad, 498, 3)
            COD_MON_TO_ACCT = Mid(cad, 501, 3)
            TIPO_CAMBIO = Mid(cad, 504, 12)
            FILLER2 = Mid(cad, 516, 1)
            'MsgBox Mid(cad, 517, 12)
            US_IMPORTE1 = CDbl(Mid(cad, 517, 12))
            US_IMPORTE2 = CDbl(Mid(cad, 529, 12))
            US_IMPORTE3 = CDbl(Mid(cad, 541, 12))
            
                
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@TIPO_REG", adVarChar, adParamInput, 2, TIPO_REG)
        Cmd.Parameters.Append Prm
    
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@LNET_ID_PSE", adVarChar, adParamInput, 4, LNET_TID_PSE)
        Cmd.Parameters.Append Prm
    
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@INST_ID_PSE", adVarChar, adParamInput, 4, INST_ID_PSE)
        Cmd.Parameters.Append Prm
    
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@TERM_ID_PSE", adVarChar, adParamInput, 16, TERM_ID_PSE)
        Cmd.Parameters.Append Prm
    
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@LNET_ID_EMISOR", adVarChar, adParamInput, 4, LNET_ID_EMISOR)
        Cmd.Parameters.Append Prm
    
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@INST_ID_EMISOR", adVarChar, adParamInput, 4, INST_ID_EMISOR)
        Cmd.Parameters.Append Prm
    
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@PAN", adVarChar, adParamInput, 19, PAN)
        Cmd.Parameters.Append Prm
    
    
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@COD_SUCURSAL", adVarChar, adParamInput, 4, COD_SUCURSAL)
        Cmd.Parameters.Append Prm
    
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@COD_REGION", adVarChar, adParamInput, 4, COD_REGION)
        Cmd.Parameters.Append Prm
    
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@TIPO_TRAN", adVarChar, adParamInput, 2, TIPO_TRAN)
        Cmd.Parameters.Append Prm
    
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@TIPO_MSG", adVarChar, adParamInput, 4, TIPO_MSG)
        Cmd.Parameters.Append Prm
    
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@STATUS_MSG", adVarChar, adParamInput, 2, STATUS_MSG)
        Cmd.Parameters.Append Prm
    
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@ORIGEN_TRAN", adVarChar, adParamInput, 1, ORIGEN_TRAN)
        Cmd.Parameters.Append Prm
    
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@ORIGEN_MSG", adVarChar, adParamInput, 1, ORIGEN_MSG)
        Cmd.Parameters.Append Prm
    
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@ENTRY_TIME", adVarChar, adParamInput, 19, ENTRY_TIME)
        Cmd.Parameters.Append Prm
    
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@EXIT_TIME", adVarChar, adParamInput, 19, EXIT_TIME)
        Cmd.Parameters.Append Prm
    
        Set Prm = New ADODB.Parameter
        'MsgBox "RE ENTRY TIME: " & RE_ENTRY_TIME
        Set Prm = Cmd.CreateParameter("@RE_ENTRY_TIME", adVarChar, adParamInput, 19, RE_ENTRY_TIME)
        Cmd.Parameters.Append Prm
    
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@TRAN_DATE", adVarChar, adParamInput, 50, TRAN_DATE)
        Cmd.Parameters.Append Prm
    
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@TRAN_TIME", adVarChar, adParamInput, 8, TRAN_TIME)
        Cmd.Parameters.Append Prm
    
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@PROC_DATE", adVarChar, adParamInput, 6, PROC_DATE)
        Cmd.Parameters.Append Prm
    
    
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@ACQ_PROC_DATE", adVarChar, adParamInput, 6, ACQ_PROC_DATE)
        Cmd.Parameters.Append Prm
    
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@ISS_PROC_DATE", adVarChar, adParamInput, 6, ISS_PROC_DATE)
        Cmd.Parameters.Append Prm
    
    
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@NUM_TRACE", adVarChar, adParamInput, 12, NUM_TRACE)
        Cmd.Parameters.Append Prm
    
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@TIPO_TERMINAL", adVarChar, adParamInput, 2, TIPO_TERMINAL)
        Cmd.Parameters.Append Prm
    
    
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@OFFSET_TIME", adVarChar, adParamInput, 5, OFFSET_TIME)
        Cmd.Parameters.Append Prm
    
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@NROUTE_ACQ", adVarChar, adParamInput, 11, NROUTE_ACQ)
        Cmd.Parameters.Append Prm
    
    
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@NROUTE_ISS", adVarChar, adParamInput, 11, NROUTE_ISS)
        Cmd.Parameters.Append Prm
    
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@TRAN_CODE", adVarChar, adParamInput, 2, TRAN_CODE)
        Cmd.Parameters.Append Prm
    
    
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@TIPO_CTA_FROM", adVarChar, adParamInput, 2, TIPO_CTA_FROM)
        Cmd.Parameters.Append Prm
    
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@TIPO_CTA_TO", adVarChar, adParamInput, 2, TIPO_CTA_FROM)
        Cmd.Parameters.Append Prm
    
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@NRO_CTA_FROM", adVarChar, adParamInput, 19, NRO_CTA_FROM)
        Cmd.Parameters.Append Prm
    
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@SOSPECHA_EXTORNO", adVarChar, adParamInput, 1, SOSPECHA_EXTORNO)
        Cmd.Parameters.Append Prm
    
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@NRO_CTA_TO", adVarChar, adParamInput, 19, NRO_CTA_FROM)
        Cmd.Parameters.Append Prm
    
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@IND_MULT_ACCT", adVarChar, adParamInput, 1, IND_MULT_ACCT)
        Cmd.Parameters.Append Prm
    
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@IMPORTE1", adCurrency, adParamInput, , CDbl(IMPORTE1))
        Cmd.Parameters.Append Prm
    
    
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@IMPORTE2", adCurrency, adParamInput, , CDbl(IMPORTE2))
        Cmd.Parameters.Append Prm
    
    
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@IMPORTE3", adCurrency, adParamInput, , CDbl(IMPORTE3))
        Cmd.Parameters.Append Prm
    
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@SALDO_CRED_DEP", adCurrency, adParamInput, , CDbl(SALDO_CRED_DEP))
        Cmd.Parameters.Append Prm
    
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@TIPO_DEPOSITO", adVarChar, adParamInput, 1, TIPO_DEPOSITO)
        Cmd.Parameters.Append Prm
    
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@IND_RETENCION", adVarChar, adParamInput, 1, IND_RETENCION)
        Cmd.Parameters.Append Prm
    
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@COD_RESP", adVarChar, adParamInput, 2, COD_RESP)
        Cmd.Parameters.Append Prm
    
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@DIRECCION_TERM", adVarChar, adParamInput, 25, DIRECCION_TERM)
        Cmd.Parameters.Append Prm
    
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@NOM_INST_TERM", adVarChar, adParamInput, 22, NOM_INST_TERM)
        Cmd.Parameters.Append Prm
    
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@CIUDAD_TERM", adVarChar, adParamInput, 13, CIUDAD_TERM)
        Cmd.Parameters.Append Prm
    
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@ESTADO_TERM", adVarChar, adParamInput, 3, ESTADO_TERM)
        Cmd.Parameters.Append Prm
    
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@PAIS_TERM", adVarChar, adParamInput, 2, PAIS_TERM)
        Cmd.Parameters.Append Prm
    
    
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@NUM_SEQ_ORIG", adVarChar, adParamInput, 12, NUM_SEQ_ORIG)
        Cmd.Parameters.Append Prm
    
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@DATE_TRAN_ORIG", adVarChar, adParamInput, 4, DATE_TRAN_ORIG)
        Cmd.Parameters.Append Prm
    
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@TIME_TRAN_ORIG", adVarChar, adParamInput, 8, TIME_TRAN_ORIG)
        Cmd.Parameters.Append Prm
    
    
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@DATE_PROC_ORIG", adVarChar, adParamInput, 4, DATE_PROC_ORIG)
        Cmd.Parameters.Append Prm
    
    
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@COD_MON_ORIG", adVarChar, adParamInput, 3, COD_MON_ORIG)
        Cmd.Parameters.Append Prm
    
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@COD_MON_AUTH", adVarChar, adParamInput, 3, COD_MON_AUTH)
        Cmd.Parameters.Append Prm
    
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@TIPO_CAMBIO_AUTH", adVarChar, adParamInput, 8, TIPO_CAMBIO_AUTH)
        Cmd.Parameters.Append Prm
    
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@COD_MON_CARGO", adVarChar, adParamInput, 3, COD_MON_CARGO)
        Cmd.Parameters.Append Prm
    
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@TIPO_CAMBIO_CARGO", adVarChar, adParamInput, 8, TIPO_CAMBIO_CARGO)
        Cmd.Parameters.Append Prm
    
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@TCAM_DATE_TIME", adVarChar, adParamInput, 19, TCAM_DATE_TIME)
        Cmd.Parameters.Append Prm
    
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@IND_MOTIVO_EXT", adVarChar, adParamInput, 2, IND_MOTIVO_EXT)
        Cmd.Parameters.Append Prm
    
    
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@SHARING_GRUOP", adVarChar, adParamInput, 1, SHARING_GRUOP)
        Cmd.Parameters.Append Prm
    
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@DEST_ORDER", adVarChar, adParamInput, 1, DEST_ORDER)
        Cmd.Parameters.Append Prm
    
    
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@HOST_COD_AUTH", adVarChar, adParamInput, 6, HOST_COD_AUTH)
        Cmd.Parameters.Append Prm
    
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@FORWARD_INST_ID", adVarChar, adParamInput, 11, FORWARD_INST_ID)
        Cmd.Parameters.Append Prm
    
    
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@CARD_ACQ_ID", adVarChar, adParamInput, 11, CARD_ACQ_ID)
        Cmd.Parameters.Append Prm
    
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@CARD_ISS_ID", adVarChar, adParamInput, 11, CARD_ISS_ID)
        Cmd.Parameters.Append Prm
    
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@ID_TOKEN", adVarChar, adParamInput, 2, ID_TOKEN)
        Cmd.Parameters.Append Prm
    
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@LONG_TOKEN", adVarChar, adParamInput, 5, LONG_TOKEN)
        Cmd.Parameters.Append Prm
    
    
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@FILLER1", adVarChar, adParamInput, 1, FILLER1)
        Cmd.Parameters.Append Prm
    
    
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@COD_MON_SOLIC", adVarChar, adParamInput, 3, COD_MON_SOLIC)
        Cmd.Parameters.Append Prm
    
    
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@COD_MON_FROM_ACCT", adVarChar, adParamInput, 3, COD_MON_FROM_ACCT)
        Cmd.Parameters.Append Prm
    
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@COD_MON_TO_ACCT", adVarChar, adParamInput, 3, COD_MON_TO_ACCT)
        Cmd.Parameters.Append Prm
    
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@TIPO_CAMBIO", adDouble, adParamInput, , TIPO_CAMBIO)
        Cmd.Parameters.Append Prm
    
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@FILLER2", adVarChar, adParamInput, 1, FILLER2)
        Cmd.Parameters.Append Prm
    
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@US_IMPORTE1", adDouble, adParamInput, , US_IMPORTE2)
        Cmd.Parameters.Append Prm
    
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@US_IMPORTE2", adDouble, adParamInput, , US_IMPORTE2)
        Cmd.Parameters.Append Prm
    
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@US_IMPORTE3", adDouble, adParamInput, , US_IMPORTE3)
        Cmd.Parameters.Append Prm
        
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@dFechaArchivo", adDate, adParamInput, , CDate(Me.mskFecha.Text))
        Cmd.Parameters.Append Prm
        
    
        Cmd.CommandType = adCmdStoredProc
        Cmd.CommandText = "ATM_CargaLog"
        Cmd.Execute
        
        'CerrarConexion
        loConec.CierraConexion
    
    End If
    
    ban = False
    i = i + 1
    Set Cmd = Nothing
    
    Loop
    Close #1
    
    
    Dim CmdF As New ADODB.Command
    Dim PrmF As New ADODB.Parameter
    
    loConec.AbreConexion
    CmdF.ActiveConnection = loConec.ConexionActiva ' AbrirConexion
    
    Set PrmF = New ADODB.Parameter
    Set PrmF = CmdF.CreateParameter("@dateFecha", adDate, adParamInput, , CDate(Me.mskFecha.Text))
    CmdF.Parameters.Append PrmF
    
    CmdF.CommandType = adCmdStoredProc
    CmdF.CommandText = "ATM_InsertaFechaLOG"
    CmdF.Execute

    'CerrarConexion
    loConec.CierraConexion
    Set loConec = Nothing
    
    MsgBox "Se cargaron un total de " & Format(i, "#,##0") & " registros.", vbInformation, "Aviso"
    Unload Me
End Sub

Private Sub cmdBuscar_Click()
    txtArchivo.Text = Empty
    
    dlgArchivo.InitDir = "C:\"
    dlgArchivo.Filter = "Archivos de Texto (*.txt)|*.txt|Todos los Archivo (*.*)|*.*"
    dlgArchivo.ShowOpen
    If dlgArchivo.FileName <> Empty Then
        txtArchivo.Text = dlgArchivo.FileName
    Else
        txtArchivo.Text = "NO SE ABRIO NINGUN ARCHIVO"
    End If
End Sub

Private Sub cmdCargar_Click()
    Cargar
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub mskFecha_GotFocus()
mskFecha.SelStart = 0
mskFecha.SelLength = 50
End Sub
