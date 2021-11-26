VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Genera la Consolidada"
   ClientHeight    =   4800
   ClientLeft      =   4830
   ClientTop       =   2715
   ClientWidth     =   4890
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   4890
   Begin VB.TextBox txtTCMes 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   2325
      TabIndex        =   9
      Text            =   "0.00"
      Top             =   1140
      Width           =   1020
   End
   Begin VB.CommandButton cmdgarantias 
      Caption         =   "Actualiza Garantias"
      Height          =   435
      Left            =   735
      TabIndex        =   7
      Top             =   3225
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   405
      Left            =   3570
      TabIndex        =   5
      Top             =   4215
      Width           =   1215
   End
   Begin VB.CommandButton cmdDTSCalcula 
      Caption         =   "DTS Calcula"
      Height          =   435
      Left            =   690
      TabIndex        =   2
      Top             =   2115
      Width           =   3000
   End
   Begin VB.CommandButton cmdDTSTransfiere 
      Caption         =   "DTS Transferencia"
      Height          =   435
      Left            =   705
      TabIndex        =   1
      Top             =   1680
      Width           =   3000
   End
   Begin VB.CommandButton cmdActualizaCapitalVencido 
      Caption         =   "Actualiza Datos Calculos "
      Height          =   435
      Left            =   690
      TabIndex        =   0
      Top             =   2565
      Width           =   3000
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Tipo de Cambio Mes:"
      Height          =   195
      Left            =   660
      TabIndex        =   8
      Top             =   1215
      Width           =   1500
   End
   Begin VB.Label lblServidor 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   135
      TabIndex        =   6
      Top             =   645
      Width           =   4605
   End
   Begin VB.Label lblFechaConsol 
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
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   2730
      TabIndex        =   4
      Top             =   180
      Width           =   1995
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha de Consolidada"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   195
      TabIndex        =   3
      Top             =   180
      Width           =   2475
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*** PARA LA CONSOLIDACION DE LA DATA
'*** 1. SE DEBE TENER UN DSNCONSOLIDADA ORIENTADO A LA BASE CONSOLIDADA

Dim objDTS As DTS.Package
Dim objSteps As DTS.Step
Dim strServerName As String, strUsuarioSQL As String, strPasswordSQL As String, strBaseSQL As String
Dim strNameDTS As String
Dim nError As Long
Dim sSource As String, sDesc As String
Dim bExito As Boolean
Dim WithEvents V As BDConsolAux.ClsConsolida
Attribute V.VB_VarHelpID = -1
Dim oBarra As clsProgressBar
Private Sub cmdActualizaCapitalVencido_Click()

    Set V = New BDConsolAux.ClsConsolida
    V.ActualizaInteresDevengado
''   V.ActualizaInteresDevengadoPrendario
    V.ActualizaGarantiasCredito
''    V.ActualizaCredGarantiasPrendario
    V.ActualizaGarantiasCreditoDetallado
 ''   V.ActualizaCredGarantiasPrendarioDetalle
    
 'ARCV 04-07-2006
 '   V.ActualizaGarantiasJudicialDetallado
 '-------------------------------
 
    V.ConsolidaCreditosConexion

'NO SON NECESARIOS
''    V.ActualizarEstadoRFA
''    V.CorrigeDescoberturaGarantia txtTCMes
    
    'V.ActualizaGarantiasJudicial
    'V.ActualizarCapitalVencido
'    V.ActualizaEstadisticaMensualCredito
'    V.ActualizaEstadisticaMensualJudicial
'    V.ActualizaEstadisticasMensualPrendario
    
    
    Set V = Nothing
End Sub

Private Sub cmdDTSCalcula_Click()
    Call DTSConsolidadaCalculos
End Sub

Private Sub cmdDTSTransfiere_Click()
    Call DTSConsolidadaTransfiere
End Sub

Private Sub cmdgarantias_Click()
Set V = New BDConsolAux.ClsConsolida
    V.ActualizaGarantiasCredito
    V.ActualizaCredGarantiasPrendario
    V.ActualizaGarantiasCreditoDetallado
    V.ActualizaCredGarantiasPrendarioDetalle
    V.ActualizaGarantiasJudicialDetallado
Set V = Nothing
End Sub

Private Sub cmdSalir_Click()
    End
End Sub

Private Sub Form_Load()
'Verifica DNSConsolidada (conexion a Base Datos Consolidada)
'"DSN=DSNConsolidada;UID=sa;PWD=cmacica"
Dim lcCon As BDConsolAux.ClsConsolida
    Set lcCon = New BDConsolAux.ClsConsolida
        lcCon.Abreconexion (True)
    Set lcCon = Nothing

'strServerName = "01srvSicmac01"
'strUsuarioSQL = "sa"
'strPasswordSQL = "cmacica"
    
'strServerName = "HYO-SRV-DESA01"
'strUsuarioSQL = "sa"
'strPasswordSQL = "confianza"
    
'CUSCO
strServerName = "10.0.0.8"
strUsuarioSQL = "sa"
strPasswordSQL = "migrasa"
'FIN CUSCO
    
    
'** Fecha de Data a Consolidar
Set lcCon = New BDConsolAux.ClsConsolida
    Me.lblFechaConsol.Caption = Format(lcCon.ObtieneFechaConsolida, "dd/mm/yyyy")
Set lcCon = Nothing
Me.lblServidor = "Server : " & strServerName
End Sub

Private Sub DTSConsolidadaTransfiere()
    strNameDTS = "DTSConsolidadaTransferencia"

Set objDTS = New DTS.Package

    objDTS.LoadFromSQLServer strServerName, strUsuarioSQL, strPasswordSQL, DTSSQLStgFlag_Default, _
                              , , , strNameDTS
    'objDTS.GlobalVariables("gFechaHora").Value = Format(gdFecSis, "mm/dd/yyyy")
    objDTS.Execute
    bExito = True
    For Each objSteps In objDTS.Steps
        objSteps.ExecuteInMainThread = True
        If objSteps.ExecutionResult = DTSStepExecResult_Failure Then
            objSteps.GetExecutionErrorInfo nError, sSource, sDesc
            MsgBox "Error en Transferencia :" & objSteps.Description & " " & sDesc, vbExclamation, "Error"
            bExito = False
            Exit For
        End If
    Next
    objDTS.UnInitialize
    Set objDTS = Nothing
End Sub


Private Sub DTSConsolidadaCalculos()
    strNameDTS = "DTSConsolidadaCalculos"

Set objDTS = New DTS.Package

    objDTS.LoadFromSQLServer strServerName, strUsuarioSQL, strPasswordSQL, DTSSQLStgFlag_Default, _
                              , , , strNameDTS
    'objDTS.GlobalVariables("gFechaHora").Value = Format(gdFecSis, "mm/dd/yyyy")
                       
    objDTS.Execute
    bExito = True
    For Each objSteps In objDTS.Steps
        objSteps.ExecuteInMainThread = True
        If objSteps.ExecutionResult = DTSStepExecResult_Failure Then
            objSteps.GetExecutionErrorInfo nError, sSource, sDesc
            MsgBox "Error en Transferencia :" & objSteps.Description & " " & sDesc, vbExclamation, "Error"
            bExito = False
            Exit For
        End If
    Next
    objDTS.UnInitialize
    Set objDTS = Nothing
End Sub


Private Sub V_finBarra()
oBarra.CloseForm Me
End Sub

Private Sub V_IniciaBarra(ByVal pnMaxValor As Long)
Set oBarra = New clsProgressBar
oBarra.Max = pnMaxValor
oBarra.ShowForm Me
End Sub

Private Sub V_Progreso(ByVal i As Long, lsTitulo As String, lsSubtitulo As String)
oBarra.Progress i, lsTitulo, lsSubtitulo
DoEvents
End Sub
