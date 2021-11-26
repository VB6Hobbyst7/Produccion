VERSION 5.00
Begin VB.Form frmCuadreConsolAge 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cuadre de Tarjeta Consolidado"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5865
   Icon            =   "frmCuadreConsolAge.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   5865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   1230
      Left            =   45
      TabIndex        =   18
      Top             =   45
      Width           =   5820
      Begin VB.CommandButton CmdProcesar 
         Caption         =   "Procesar"
         Height          =   330
         Left            =   75
         TabIndex        =   19
         Top             =   705
         Width           =   5595
      End
      Begin VB.Label Label4 
         Caption         =   "Usuario :"
         Height          =   285
         Left            =   2730
         TabIndex        =   23
         Top             =   270
         Width           =   765
      End
      Begin VB.Label lblFecha 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   795
         TabIndex        =   22
         Top             =   285
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha :"
         Height          =   285
         Left            =   105
         TabIndex        =   21
         Top             =   300
         Width           =   630
      End
      Begin VB.Label lblUsuario 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Usu"
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
         Left            =   3510
         TabIndex        =   20
         Top             =   255
         Width           =   1320
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2505
      Left            =   30
      TabIndex        =   3
      Top             =   1380
      Width           =   5805
      Begin VB.Label lblRemSal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
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
         Left            =   4230
         TabIndex        =   27
         Top             =   1065
         Width           =   840
      End
      Begin VB.Label Label7 
         Caption         =   "Remesas Salida :"
         Height          =   285
         Left            =   2940
         TabIndex        =   26
         Top             =   1080
         Width           =   1245
      End
      Begin VB.Label Label5 
         Caption         =   "Remesas Entrada:"
         Height          =   285
         Left            =   540
         TabIndex        =   25
         Top             =   1080
         Width           =   1395
      End
      Begin VB.Label lblRemEnt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
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
         Left            =   1950
         TabIndex        =   24
         Top             =   1065
         Width           =   840
      End
      Begin VB.Label lblDiferencia 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
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
         Left            =   4245
         TabIndex        =   17
         Top             =   2070
         Width           =   840
      End
      Begin VB.Label Label14 
         Caption         =   "Diferencia :"
         Height          =   285
         Left            =   3090
         TabIndex        =   16
         Top             =   2085
         Width           =   1035
      End
      Begin VB.Line Line1 
         X1              =   135
         X2              =   5580
         Y1              =   1980
         Y2              =   1980
      End
      Begin VB.Label lblStockAct 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
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
         Left            =   4230
         TabIndex        =   15
         Top             =   1515
         Width           =   840
      End
      Begin VB.Label Label12 
         Caption         =   "Stock Actual :"
         Height          =   285
         Left            =   2940
         TabIndex        =   14
         Top             =   1530
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "Saldo Final :"
         Height          =   285
         Left            =   780
         TabIndex        =   13
         Top             =   1530
         Width           =   1035
      End
      Begin VB.Label Label8 
         Caption         =   "Saldo Anterior :"
         Height          =   285
         Left            =   600
         TabIndex        =   12
         Top             =   240
         Width           =   1260
      End
      Begin VB.Label lblDevoluciones 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
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
         Left            =   4245
         TabIndex        =   11
         Top             =   630
         Width           =   840
      End
      Begin VB.Label Label6 
         Caption         =   "Devoluciones :"
         Height          =   285
         Left            =   2970
         TabIndex        =   10
         Top             =   645
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Habilitaciones :"
         Height          =   285
         Left            =   600
         TabIndex        =   9
         Top             =   645
         Width           =   1260
      End
      Begin VB.Label lblHabilitaciones 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
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
         Left            =   1965
         TabIndex        =   8
         Top             =   630
         Width           =   840
      End
      Begin VB.Label lblSaldoAnt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
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
         Left            =   1965
         TabIndex        =   7
         Top             =   225
         Width           =   840
      End
      Begin VB.Label lblSaldoF 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
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
         Left            =   1950
         TabIndex        =   6
         Top             =   1515
         Width           =   840
      End
      Begin VB.Label LblMovim 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
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
         Left            =   4245
         TabIndex        =   5
         Top             =   240
         Width           =   840
      End
      Begin VB.Label LBL 
         Caption         =   "Movimiento :"
         Height          =   285
         Left            =   3015
         TabIndex        =   4
         Top             =   255
         Width           =   960
      End
   End
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   45
      TabIndex        =   0
      Top             =   3915
      Width           =   5790
      Begin VB.CommandButton CmdImprimir 
         Caption         =   "Imprimir"
         Enabled         =   0   'False
         Height          =   390
         Left            =   90
         TabIndex        =   2
         Top             =   210
         Width           =   1305
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Salir"
         Height          =   390
         Left            =   4275
         TabIndex        =   1
         Top             =   210
         Width           =   1305
      End
   End
End
Attribute VB_Name = "frmCuadreConsolAge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oConec As DConecta

Private Sub CboUsu_Click()
    CmdProcesar.Enabled = True
End Sub

Private Sub CmdImprimir_Click()
Dim P As Previo.clsPrevio
Dim R As ADODB.Recordset
    
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter

Dim sSQL As String
Dim sCadRep As String
    
    Set Cmd = New ADODB.Command


    Set R = New ADODB.Recordset
    sCadRep = "."

    'Cabecera
    sCadRep = sCadRep & Space(5) & "CMAC MAYNAS S.A." & Space(50) & "FECHA : " & Format(Now(), "dd/mm/yyyy hh:mm:ss") & Chr(10)
    sCadRep = sCadRep & Space(5) & "SIMACC-Tarjeta de Debito" & Chr(10) & Chr(10)
    sCadRep = sCadRep & Space(30) & "Reporte Del Cuadre de Tarjetas CONSOLIDADO " & Chr(10) & Chr(10) & Chr(10)
    sCadRep = sCadRep & Space(5) & String(100, "-") & Chr(10) & Chr(10)
        
    'Cuerpo
    
    sCadRep = sCadRep & Space(5) & Left("Saldo Anterior     : " & Space(20), 20) & Right(Space(15) & lblSaldoAnt.Caption, 15) & Chr(10) & Chr(10)
    sCadRep = sCadRep & Space(5) & Left("Habilitaciones     : " & Space(20), 20) & Right(Space(15) & lblHabilitaciones.Caption, 15) & Chr(10)
    sCadRep = sCadRep & Space(5) & Left("Devoluciones       : " & Space(20), 20) & Right(Space(15) & lblDevoluciones.Caption, 15) & Chr(10)
    
    sCadRep = sCadRep & Space(5) & Left("Remesa de Entrada  : " & Space(20), 20) & Right(Space(15) & lblRemEnt.Caption, 15) & Chr(10)
    sCadRep = sCadRep & Space(5) & Left("Remesa de Salida   : " & Space(20), 20) & Right(Space(15) & lblRemSal.Caption, 15) & Chr(10)
    
    sCadRep = sCadRep & Space(5) & Left("Movimientos        : " & Space(20), 20) & Right(Space(15) & LblMovim.Caption, 15) & Chr(10)
    sCadRep = sCadRep & Space(5) & Left("Saldo Final        : " & Space(20), 20) & Right(Space(15) & lblSaldoF.Caption, 15) & Chr(10) & Chr(10)
    
    sCadRep = sCadRep & Space(5) & Left("Stock Actual       : " & Space(20), 20) & Right(Space(15) & lblStockAct.Caption, 15) & Chr(10) & Chr(10)
    sCadRep = sCadRep & Space(5) & Left("Diferencia         : " & Space(20), 20) & Right(Space(15) & lblDiferencia.Caption, 15) & Chr(10) & Chr(10)
    
'
    
'    'sCadRep = sCadRep & Space(5) & Right(Space(16) & lblSaldoAnt.Caption, 16) & Space(5) & lblIngresos.Caption & Space(5) & Left(lblSalidas.Caption & Space(30), 25) & Space(2) & Left(lblRemesas.Caption & Space(20), 10) & Space(2) & Left(lblConfirmaciones & Space(30), 20) & Space(2) & Chr(10)
'    sCadRep = sCadRep & Space(5) & Space(20) & Left("Saldo Anterior: " & Space(25), 16) & Right(lblSaldoAnt.Caption, 6) & Space(16) & Left("Total de Movimiento: " & Space(25), 21) & Right(LblMovim.Caption, 6) & Chr(10)
'    'sCadRep = sCadRep & Space(5) & "Total de Movimiento: " & LblMovim.Caption & Chr(10) & Chr(10)
'    sCadRep = sCadRep & Space(5) & String(100, " ") & Chr(10)
'    sCadRep = sCadRep & Space(5) & Space(20) & Left("Total de Habilitaciones: " & Space(25), 25) & Right(lblHabilitaciones.Caption, 6) & Space(6) & Left("Total de Devoluciones: " & Space(25), 23) & Right(lblDevoluciones.Caption, 6) & Chr(10)
'    'sCadRep = sCadRep & Space(5) & Right(Space(36) & "Total de Habilitaciones: " & lblHabilitaciones.Caption, 36) & Right(Space(10) & "Total de Devoluciones: " & lblDevoluciones.Caption, 25) & Chr(10)
'    sCadRep = sCadRep & Space(5) & String(100, " ") & Chr(10)
'    'sCadRep = sCadRep & Space(5) & "Total de Devoluciones: " & lblDevoluciones.Caption & Chr(10) & Chr(10)
'    'sCadRep = sCadRep & Space(5) & String(100, " ") & Chr(10)
'    sCadRep = sCadRep & Space(5) & Space(20) & Left("Total de Saldo Final: " & Space(25), 23) & Right(lblSaldoF.Caption, 6) & Space(8) & Left("Total Stock Actual: " & Space(25), 20) & Right(lblStockAct.Caption, 6) & Chr(10)
'    'sCadRep = sCadRep & Space(5) & Right(Space(36) & "Total de Saldo Final: " & lblSaldoF.Caption, 36) & Right(Space(10) & "Total Stock Actual: " & lblStockAct.Caption, 25) & Chr(10)
'    sCadRep = sCadRep & Space(5) & Space(20) & String(65, "_") & Chr(10)
'    sCadRep = sCadRep & Space(5) & String(100, " ") & Chr(10)
'    'sCadRep = sCadRep & Space(5) & "Total Stock Actual: " & lblStockAct.Caption & Chr(10) & Chr(10)
'    'sCadRep = sCadRep & Space(5) & String(100, " ") & Chr(10)
'    sCadRep = sCadRep & Space(5) & Space(53) & Left("Diferencia: " & Space(25), 12) & Right(lblDiferencia.Caption, 6)
'    sCadRep = sCadRep & Space(5) & String(100, " ") & Chr(10) & Chr(10)
    sCadRep = sCadRep & Space(5) & String(100, "-") & Chr(10) & Chr(10)

    Set P = New Previo.clsPrevio
    Call P.Show(sCadRep, "REPORTE")
    Set P = Nothing
    



End Sub

Private Sub CmdProcesar_Click()
    Call CargaValores
    CmdImprimir.Enabled = True
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Set oConec = New DConecta
    Call CargaDatos
    lblUsuario.Caption = gsCodUser
End Sub

Private Sub CargaValores()
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter
Dim R As ADODB.Recordset
    
    Set R = New ADODB.Recordset
    
    oConec.AbreConexion
    Cmd.ActiveConnection = oConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCodUsu", adVarChar, adParamInput, 20, lblUsuario.Caption)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@dFecha", adDBDate, adParamInput, , lblFecha.Caption)
    Cmd.Parameters.Append Prm
            
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nSaldoAnt", adInteger, adParamOutput, 8, 0)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nHabilitaciones", adInteger, adParamOutput, 8, 0)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nDevoluciones", adInteger, adParamOutput, 8, 0)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nRemEntrada", adInteger, adParamOutput, , 0)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nRemSalida", adInteger, adParamOutput, , 0)
    Cmd.Parameters.Append Prm
    
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nStockActual", adInteger, adParamOutput, 8, 0)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nSaldoFinal", adInteger, adParamOutput, , 0)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nDiferencia", adInteger, adParamOutput, , 0)
    Cmd.Parameters.Append Prm
            
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nMovimiento", adInteger, adParamOutput, , 0)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCodAge", adChar, adParamInput, 4, gsCodAge)
    Cmd.Parameters.Append Prm
    
    Cmd.CommandText = "ATM_RegistraCuadreTarjetaConsol"
    Dim i As Integer
    
    R.CursorType = adOpenStatic
    R.LockType = adLockReadOnly
    Set R = Cmd.Execute
    
    '        For i = 0 To 8 Step 1
    '            MsgBox Cmd.Parameters(i).Value
    '        Next
    lblSaldoAnt = Cmd.Parameters(2).Value
    lblHabilitaciones = Cmd.Parameters(3).Value
    lblDevoluciones = Cmd.Parameters(4).Value
    
    lblRemEnt = Cmd.Parameters(5).Value
    lblRemSal = Cmd.Parameters(6).Value
    
    lblStockAct = Cmd.Parameters(7).Value
    lblSaldoF = Cmd.Parameters(8).Value
    lblDiferencia = Cmd.Parameters(9).Value
    LblMovim = Cmd.Parameters(10).Value
    
    'CerrarConexion
    oConec.CierraConexion
    Set oConec = Nothing
    Set Cmd = Nothing
    Set Prm = Nothing
    'CmdAct.Enabled = True
End Sub

Private Sub CargaDatos()
 
    Dim Cmd As New Command
    Dim Prm As New ADODB.Parameter
    Dim R As ADODB.Recordset
    Set Prm = New ADODB.Parameter
'
'    Set Prm = New ADODB.Parameter
'    Set Prm = Cmd.CreateParameter("@pnCodAge", adInteger, adParamInput, , CInt(gsCodAge))
'    Cmd.Parameters.Append Prm
'
'
'    Cmd.ActiveConnection = AbrirConexion
'    Cmd.CommandType = adCmdStoredProc
'    Cmd.CommandText = "ATM_RecuperaUsuarios"
'
'    Set R = Cmd.Execute
'    'CboCtas.Clear
'    Do While Not R.EOF
'         CboUsu.AddItem R!Codigo & "-" & R!Nombre
'       R.MoveNext
'    Loop
'    R.Close
'    Set R = Nothing
'
'    CerrarConexion
    
    'FECHA DEL SISTEMA
    '
    Set Cmd = New Command
    Set Prm = New ADODB.Parameter
    
    Set Prm = New ADODB.Parameter
    
    oConec.AbreConexion
    Cmd.ActiveConnection = oConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_RecuperaFechaSistema"
    
    Set R = Cmd.Execute
    Me.lblFecha.Caption = Format(CDate(Format(R!FechaSistema, "dd/mm/yyyy")), "dd/mm/yyyy")
    R.Close
    Set Cmd = Nothing
    
    Set R = Nothing
    
    'CerrarConexion
    oConec.CierraConexion
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set oConec = Nothing
End Sub
