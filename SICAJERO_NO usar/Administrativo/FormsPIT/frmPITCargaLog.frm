VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmPITCargaLog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Conciliación de Operaciones InterCajas: Carga de Archivo LOG Diario"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   7605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   675
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   7410
      Begin VB.Label Label3 
         Caption         =   "Usuario:"
         Height          =   165
         Left            =   120
         TabIndex        =   12
         Top             =   285
         Width           =   600
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha de Carga:"
         Height          =   285
         Left            =   2880
         TabIndex        =   11
         Top             =   300
         Width           =   1440
      End
      Begin VB.Label lblUsuario 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1200
         TabIndex        =   10
         Top             =   240
         Width           =   1320
      End
      Begin VB.Label lblFecha 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   4440
         TabIndex        =   9
         Top             =   240
         Width           =   1320
      End
   End
   Begin MSComDlg.CommonDialog dlgArchivo 
      Left            =   120
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdIniciar 
      Caption         =   "&Iniciar Carga"
      Height          =   360
      Left            =   4680
      TabIndex        =   7
      Top             =   2520
      Width           =   1230
   End
   Begin VB.Frame fraArchivoLOG 
      Caption         =   "Archivo Log Diario de Operaciones "
      Height          =   1320
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   7410
      Begin MSMask.MaskEdBox mskFechaLog 
         Height          =   300
         Left            =   1200
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
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   300
         Width           =   5640
      End
      Begin MSComctlLib.ProgressBar pgbLog 
         Height          =   195
         Left            =   50
         TabIndex        =   13
         Top             =   1080
         Width           =   7305
         _ExtentX        =   12885
         _ExtentY        =   344
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label lblRuta 
         Caption         =   "Archivo :"
         Height          =   180
         Left            =   75
         TabIndex        =   6
         Top             =   375
         Width           =   795
      End
      Begin VB.Label lblFechaLog 
         Caption         =   "Fecha de Log :"
         Height          =   300
         Left            =   75
         TabIndex        =   5
         Top             =   720
         Width           =   1140
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   360
      Left            =   6240
      TabIndex        =   0
      Top             =   2520
      Width           =   1230
   End
End
Attribute VB_Name = "frmPITCargaLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim nTotalRegLog As Long, nLogOpeId As Long
Dim sRuta As String

Private Sub RegistraLogOpeIB()
Dim lsTipoReg As String, lsPtoServ_LnetIdPse As String, lsPtoServ_InstIdPse As String, lsPtoServ_TermIdPse  As String
Dim lsTarjeta_LnetIdEmisor As String, lsTarjeta_InstIdEmisor As String, lsTarjeta_PAN As String
Dim lsTarjeta_CodSucursal As String, lsTarjeta_CodRegion As String
Dim lsAuthData_TipoTran As String, lsAuthData_TipoMsg As String, lsAuthData_StatusMsg As String
Dim lsAuthData_OrigenTran As String, lsAuthData_OrigenMsg As String, lsAuthData_EntryTime As String
Dim lsAuthData_ExitTime As String, lsAuthData_ReEntryTime As String, lsAuthData_TranDate As String
Dim lsAuthData_TranTime  As String, lsAuthData_ProcDate As String, lsAuthData_AcqProcDate As String
Dim lsAuthData_IssProcDate As String, lsAuthData_NumTrace  As String, lsAuthData_TipoTerminal As String
Dim lsAuthData_OffSetTime As String, lsAuthData_NRouteAcq As String, lsAuthData_NRouteIss  As String
Dim lsCodTransac_TranCode As String, lsCodTransac_TipoCtaFrom As String, lsCodTransac_TipoCtaTo As String
Dim lsNroCtaFrom As String, lsSospechaExtorno As String, lsNroCtaTo As String, lsIndMultAcct As String
Dim lnImporte1  As Double, lnImporte2 As Double, lnImporte3 As Double, lnSaldoCredDep As Double
Dim lsTipoDeposito As String, lsRespuesta_IndRetencion As String, lsRespuesta_CodResp As String
Dim lsUbicacion_DireccionTerm As String, lsUbicacion_NomInstTerm As String, lsUbicacion_CiudadTerm As String
Dim lsUbicacion_EstadoTerm As String, lsUbicacion_PaisTerm As String, lsDataOrigTran_NumSeqOrig As String
Dim lsDataOrigTran_DateTranOrig As String, lsDataOrigTran_TimeTranOrig As String, lsDataOrigTran_DateProcOrig As String
Dim lsDataOrigTran_CodMonOrig As String, lsMultiMoneda_CodMonAuth As String, lsMultiMoneda_TipoCambioAuth As String
Dim lsMultiMoneda_CodMonCargo As String, lsMultiMoneda_TipoCambioCargo As String, lsMultiMoneda_TCamDateTime As String
Dim lsIndMotivoExt As String, lsSharingGruop As String, lsDestOrder As String, lsHostCodAuth As String
Dim lsForwardInstId  As String, lsCardAcqId As String, lsCardIssId As String
Dim lsTokenMulti_IdToken  As String, lsTokenMulti_LongToken As String, lsTokenMulti_Filler1 As String
Dim lsTokenMulti_CodMonSolic  As String, lsTokenMulti_CodMonFromAcct  As String, lsTokenMulti_CodMonToAcct As String
Dim lnTokenMulti_TipoCambio   As Double, lsTokenMulti_Filler2 As String
Dim lnTokenMulti_UsImporte1   As Double, lnTokenMulti_UsImporte2 As Double, lnTokenMulti_UsImporte3 As Double
Dim lsIDAdquiriente As String, lsIDEmisor As String
    

Dim lsCad As String, lsSQL As String
Dim lnNumReg As Long

Dim loConec As DConecta
Dim loOIC As dPITFunciones
    
        
    nTotalRegLog = 0
    Open sRuta For Input As #1
    Do Until EOF(1)
        Input #1, lsCad
        nTotalRegLog = nTotalRegLog + 1
    Loop
    Close #1
        
    Me.pgbLog.Max = nTotalRegLog + 2
    Me.pgbLog.Min = 0
    Me.pgbLog.Value = 0
    
    Set loOIC = New dPITFunciones
    nLogOpeId = loOIC.nRegistraLogOpeIB(gdFecSis, CDate(Me.mskFechaLog.Text), gsCodUser)
    
    lnNumReg = 0
    
    Set loConec = New DConecta
    Open sRuta For Input As #1
        Input #1, lsCad 'Obviar primera fila
        
        loConec.AbreConexion
        Do Until EOF(1)
            
            Input #1, lsCad
            
            Me.pgbLog.Value = Me.pgbLog.Value + 1
            
            lsTipoReg = Mid(lsCad, 1, 2)
            lsPtoServ_LnetIdPse = Mid(lsCad, 3, 4)
            lsPtoServ_InstIdPse = Mid(lsCad, 7, 4)
            lsPtoServ_TermIdPse = Mid(lsCad, 11, 16)
            lsTarjeta_LnetIdEmisor = Mid(lsCad, 27, 4)
            lsTarjeta_InstIdEmisor = Mid(lsCad, 31, 4)
            lsTarjeta_PAN = Mid(lsCad, 35, 19)
            lsTarjeta_CodSucursal = Mid(lsCad, 54, 4)
            lsTarjeta_CodRegion = Mid(lsCad, 58, 4)
            lsAuthData_TipoTran = Mid(lsCad, 62, 2)
            lsAuthData_TipoMsg = Mid(lsCad, 64, 4)
            lsAuthData_StatusMsg = Mid(lsCad, 68, 2)
            lsAuthData_OrigenTran = Mid(lsCad, 70, 1)
            lsAuthData_OrigenMsg = Mid(lsCad, 71, 1)
            lsAuthData_EntryTime = Mid(lsCad, 72, 19)
            lsAuthData_ExitTime = Mid(lsCad, 91, 19)
            lsAuthData_ReEntryTime = Mid(lsCad, 110, 19)
            lsAuthData_TranDate = Mid(lsCad, 129, 6)
            lsAuthData_TranTime = Mid(lsCad, 135, 8)
            lsAuthData_ProcDate = Mid(lsCad, 143, 6)
            lsAuthData_AcqProcDate = Mid(lsCad, 149, 6)
            lsAuthData_IssProcDate = Mid(lsCad, 155, 6)
            lsAuthData_NumTrace = Mid(lsCad, 161, 12)
            lsAuthData_TipoTerminal = Mid(lsCad, 173, 2)
            lsAuthData_OffSetTime = Mid(lsCad, 175, 5)
            lsAuthData_NRouteAcq = Mid(lsCad, 180, 11)
            lsAuthData_NRouteIss = Mid(lsCad, 191, 11)
            lsCodTransac_TranCode = Mid(lsCad, 202, 2)
            lsCodTransac_TipoCtaFrom = Mid(lsCad, 204, 2)
            lsCodTransac_TipoCtaTo = Mid(lsCad, 206, 2)
            lsNroCtaFrom = Mid(lsCad, 208, 19)
            lsSospechaExtorno = Mid(lsCad, 227, 1)
            lsNroCtaTo = Mid(lsCad, 228, 19)
            lsIndMultAcct = Mid(lsCad, 247, 1)
            lnImporte1 = CDbl(Mid(lsCad, 248, 15)) / 100
            lnImporte2 = CDbl(Mid(lsCad, 263, 15)) / 100
            lnImporte3 = CDbl(Mid(lsCad, 278, 15)) / 100
            lnSaldoCredDep = CDbl(Mid(lsCad, 293, 10))
            lsTipoDeposito = Mid(lsCad, 303, 1)
            lsRespuesta_IndRetencion = Mid(lsCad, 304, 1)
            lsRespuesta_CodResp = Mid(lsCad, 305, 2)
            lsUbicacion_DireccionTerm = Mid(lsCad, 307, 25)
            lsUbicacion_NomInstTerm = Mid(lsCad, 332, 22)
            lsUbicacion_CiudadTerm = Mid(lsCad, 354, 13)
            lsUbicacion_EstadoTerm = Mid(lsCad, 367, 3)
            lsUbicacion_PaisTerm = Mid(lsCad, 370, 2)
            lsDataOrigTran_NumSeqOrig = Mid(lsCad, 372, 12)
            lsDataOrigTran_DateTranOrig = Mid(lsCad, 384, 4)
            lsDataOrigTran_TimeTranOrig = Mid(lsCad, 388, 8)
            lsDataOrigTran_DateProcOrig = Mid(lsCad, 396, 4)
            lsDataOrigTran_CodMonOrig = Mid(lsCad, 400, 3)
            lsMultiMoneda_CodMonAuth = Mid(lsCad, 403, 3)
            lsMultiMoneda_TipoCambioAuth = CDbl(Mid(lsCad, 406, 8))
            lsMultiMoneda_CodMonCargo = Mid(lsCad, 414, 3)
            lsMultiMoneda_TipoCambioCargo = Mid(lsCad, 417, 8)
            lsMultiMoneda_TCamDateTime = Mid(lsCad, 425, 19)
            lsIndMotivoExt = Mid(lsCad, 444, 2)
            lsSharingGruop = Mid(lsCad, 446, 1)
            lsDestOrder = Mid(lsCad, 447, 1)
            lsHostCodAuth = Mid(lsCad, 448, 6)
            lsForwardInstId = Mid(lsCad, 454, 11)
            lsCardAcqId = Mid(lsCad, 465, 11)
            lsCardIssId = Mid(lsCad, 476, 11)
            
            lsTokenMulti_IdToken = Mid(lsCad, 487, 2)
            lsTokenMulti_LongToken = Mid(lsCad, 489, 5)
            lsTokenMulti_Filler1 = Mid(lsCad, 494, 1)
            lsTokenMulti_CodMonSolic = Mid(lsCad, 495, 3)
            lsTokenMulti_CodMonFromAcct = Mid(lsCad, 498, 3)
            lsTokenMulti_CodMonToAcct = Mid(lsCad, 501, 3)
            lnTokenMulti_TipoCambio = 0 'Mid(lsCad, 504, 12)
            lsTokenMulti_Filler2 = Mid(lsCad, 516, 1)
            lnTokenMulti_UsImporte1 = 0 ' CDbl(Mid(lsCad, 517, 12))
            lnTokenMulti_UsImporte2 = 0 ' CDbl(Mid(lsCad, 529, 12))
            lnTokenMulti_UsImporte3 = 0 ' CDbl(Mid(lsCad, 541, 12))
            
            lsIDAdquiriente = Mid(lsCad, 571, 6)
            lsIDEmisor = Mid(lsCad, 577, 6)
                

            lsSQL = " exec PIT_stp_ins_LogDetOpeIB '" & lsTipoReg & "','" & lsPtoServ_LnetIdPse & "','" & lsPtoServ_InstIdPse & "','" & lsPtoServ_TermIdPse & "','" & lsTarjeta_LnetIdEmisor & "','" & _
                         lsTarjeta_InstIdEmisor & "','" & lsTarjeta_PAN & "','" & lsTarjeta_CodSucursal & "','" & lsTarjeta_CodRegion & "','" & lsAuthData_TipoTran & "','" & _
                         lsAuthData_TipoMsg & "','" & lsAuthData_StatusMsg & "','" & lsAuthData_OrigenTran & "','" & lsAuthData_OrigenMsg & "','" & lsAuthData_EntryTime & "','" & _
                         lsAuthData_ExitTime & "','" & lsAuthData_ReEntryTime & "','" & lsAuthData_TranDate & "','" & lsAuthData_TranTime & "','" & lsAuthData_ProcDate & "','" & _
                         lsAuthData_AcqProcDate & "','" & lsAuthData_IssProcDate & "','" & lsAuthData_NumTrace & "','" & lsAuthData_TipoTerminal & "','" & lsAuthData_OffSetTime & "','" & _
                         lsAuthData_NRouteAcq & "','" & lsAuthData_NRouteIss & "','" & lsCodTransac_TranCode & "','" & lsCodTransac_TipoCtaFrom & "','" & lsCodTransac_TipoCtaTo & "','" & _
                         lsNroCtaFrom & "','" & lsSospechaExtorno & "','" & lsNroCtaTo & "','" & lsIndMultAcct & "'," & lnImporte1 & "," & lnImporte2 & "," & lnImporte3 & "," & lnSaldoCredDep & ",'" & lsTipoDeposito & "','" & _
                         lsRespuesta_IndRetencion & "','" & lsRespuesta_CodResp & "','" & lsUbicacion_DireccionTerm & "','" & lsUbicacion_NomInstTerm & "','" & lsUbicacion_CiudadTerm & "','" & _
                         lsUbicacion_EstadoTerm & "','" & lsUbicacion_PaisTerm & "','" & lsDataOrigTran_NumSeqOrig & "','" & lsDataOrigTran_DateTranOrig & "','" & lsDataOrigTran_TimeTranOrig & "','" & _
                         lsDataOrigTran_DateProcOrig & "','" & lsDataOrigTran_CodMonOrig & "','" & lsMultiMoneda_CodMonAuth & "','" & lsMultiMoneda_TipoCambioAuth & "','" & _
                         lsMultiMoneda_CodMonCargo & "','" & lsMultiMoneda_TipoCambioCargo & "','" & lsMultiMoneda_TCamDateTime & "','" & lsIndMotivoExt & "','" & lsSharingGruop & "','" & _
                         lsDestOrder & "','" & lsHostCodAuth & "','" & lsForwardInstId & "','" & lsCardAcqId & "','" & lsCardIssId & "','" & lsTokenMulti_IdToken & "','" & lsTokenMulti_LongToken & "','" & _
                         lsTokenMulti_Filler1 & "','" & lsTokenMulti_CodMonSolic & "','" & lsTokenMulti_CodMonFromAcct & "','" & lsTokenMulti_CodMonToAcct & "','" & lnTokenMulti_TipoCambio & "','" & _
                         lsTokenMulti_Filler2 & "'," & lnTokenMulti_UsImporte1 & "," & lnTokenMulti_UsImporte2 & "," & lnTokenMulti_UsImporte3 & ",'" & _
                         lsIDAdquiriente & "','" & lsIDEmisor & "'," & nLogOpeId
        
            loConec.Ejecutar (lsSQL)
                  
            lnNumReg = lnNumReg + 1
        
        Loop
    Close #1
    loConec.CierraConexion
        
    
    MsgBox "Se cargaron un total de " & Format(lnNumReg, "#,##0") & " registros.", vbInformation, "Aviso"
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
        txtArchivo.Text = ""
        Call MsgBox("NO SE ABRIO NINGUN ARCHIVO", vbInformation)
    End If
End Sub

Private Sub cmdIniciar_Click()
Dim lsCad As String
Dim lRs As ADODB.Recordset
Dim loOIC As dPITFunciones

    sRuta = Me.txtArchivo.Text

    If Not IsDate(Me.mskFechaLog.Text) Then
        MsgBox "Fecha no valida.", vbInformation, "Aviso"
        mskFechaLog.SetFocus
        Exit Sub
    End If
    
    If Trim(sRuta) = "" Then
        MsgBox "No ha seleccionado ningún archivo", vbInformation, "Aviso"
        cmdBuscar.SetFocus
        Exit Sub
    End If
        
    Open sRuta For Input As #1
    If Not EOF(1) Then
        Input #1, lsCad

        If Format(CDate(Me.mskFechaLog.Text), "YYYYMMDD") <> Trim(Left(lsCad, 8)) Then
            MsgBox "Fecha no valida. La fecha del archivo LOG es : " & lsCad, vbInformation, "Aviso"
            Close #1
            mskFechaLog.SetFocus
            Exit Sub
        End If
    End If
    Close #1
    
    Set loOIC = New dPITFunciones
    Set lRs = loOIC.recuperaLogOpeIBPorFechaLog(CDate(mskFechaLog.Text))
    Set loOIC = Nothing
    If Not (lRs.EOF And lRs.BOF) Then
        MsgBox "No puede conciliarse con un archivo LOG con la misma fecha dos veces", vbInformation, "Aviso"
        Exit Sub
    End If
    
    
    Call RegistraLogOpeIB
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    lblFecha.Caption = Format(gdFecSis, "DD/MM/YYYY")
    lblUsuario.Caption = gsCodUser
    
End Sub

Private Sub mskFechaLog_GotFocus()
mskFechaLog.SelStart = 0
mskFechaLog.SelLength = 50
End Sub
