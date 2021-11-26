VERSION 5.00
Begin VB.Form frmCuadreBvda 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cuadre Boveda "
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5925
   Icon            =   "frmCuadreBvda.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   5925
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   90
      TabIndex        =   18
      Top             =   3630
      Width           =   5790
      Begin VB.CommandButton CmdImprimir 
         Caption         =   "Imprimir"
         Enabled         =   0   'False
         Height          =   390
         Left            =   75
         TabIndex        =   20
         Top             =   210
         Width           =   1305
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Salir"
         Height          =   390
         Left            =   4275
         TabIndex        =   19
         Top             =   210
         Width           =   1305
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2385
      Left            =   75
      TabIndex        =   4
      Top             =   1230
      Width           =   5790
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
         Left            =   4260
         TabIndex        =   23
         Top             =   1395
         Width           =   840
      End
      Begin VB.Label Label5 
         Caption         =   "Stock Actual:"
         Height          =   285
         Left            =   2955
         TabIndex        =   22
         Top             =   1395
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Rem.Entrada:"
         Height          =   285
         Left            =   2940
         TabIndex        =   21
         Top             =   615
         Width           =   1215
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
         Left            =   4260
         TabIndex        =   17
         Top             =   1980
         Width           =   840
      End
      Begin VB.Label Label14 
         Caption         =   "Diferencia :"
         Height          =   285
         Left            =   3090
         TabIndex        =   16
         Top             =   1980
         Width           =   1035
      End
      Begin VB.Line Line1 
         X1              =   150
         X2              =   5595
         Y1              =   1830
         Y2              =   1830
      End
      Begin VB.Label lblRemSalida 
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
         Left            =   4260
         TabIndex        =   15
         Top             =   1020
         Width           =   840
      End
      Begin VB.Label Label12 
         Caption         =   "Rem.Salida:"
         Height          =   285
         Left            =   3060
         TabIndex        =   14
         Top             =   1005
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "Saldo Final :"
         Height          =   285
         Left            =   795
         TabIndex        =   13
         Top             =   1020
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
      Begin VB.Label lblRemEntrada 
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
         Left            =   4260
         TabIndex        =   11
         Top             =   630
         Width           =   840
      End
      Begin VB.Label Label6 
         Caption         =   "Devoluciones :"
         Height          =   285
         Left            =   615
         TabIndex        =   10
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Habilitaciones :"
         Height          =   285
         Left            =   2865
         TabIndex        =   9
         Top             =   270
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
         Left            =   1935
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
         Left            =   1935
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
         Left            =   1935
         TabIndex        =   6
         Top             =   1005
         Width           =   840
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
         Left            =   4260
         TabIndex        =   5
         Top             =   240
         Width           =   840
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1230
      Left            =   90
      TabIndex        =   0
      Top             =   -15
      Width           =   5820
      Begin VB.CommandButton CmdProcesar 
         Caption         =   "Procesar"
         Height          =   330
         Left            =   90
         TabIndex        =   1
         Top             =   705
         Width           =   5595
      End
      Begin VB.Label Label9 
         Caption         =   "Usuario :"
         Height          =   285
         Left            =   3060
         TabIndex        =   25
         Top             =   210
         Width           =   765
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "B O V E D A"
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
         Left            =   3900
         TabIndex        =   24
         Top             =   210
         Width           =   1320
      End
      Begin VB.Label lblFecha 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   780
         TabIndex        =   3
         Top             =   225
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha :"
         Height          =   285
         Left            =   105
         TabIndex        =   2
         Top             =   240
         Width           =   630
      End
   End
End
Attribute VB_Name = "frmCuadreBvda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oConec As DConecta

Private Sub CargaFecha()
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
    sCadRep = sCadRep & Space(40) & "Reporte Del Cuadre de Boveda" & Chr(10) & Chr(10) & Chr(10)
    sCadRep = sCadRep & Space(5) & String(100, "-") & Chr(10) & Chr(10)
    
    
    'Cuerpo
    sCadRep = sCadRep & Space(5) & Left("Saldo Anterior     : " & Space(20), 20) & Right(Space(15) & lblSaldoAnt.Caption, 15) & Chr(10) & Chr(10)
    sCadRep = sCadRep & Space(5) & Left("Habilitaciones     : " & Space(20), 20) & Right(Space(15) & lblHabilitaciones.Caption, 15) & Chr(10)
    sCadRep = sCadRep & Space(5) & Left("Remesas de Salida  : " & Space(20), 20) & Right(Space(15) & lblRemSalida.Caption, 15) & Chr(10)
    sCadRep = sCadRep & Space(5) & Left("Remesas de Entrada : " & Space(20), 20) & Right(Space(15) & lblRemEntrada.Caption, 15) & Chr(10)
    sCadRep = sCadRep & Space(5) & Left("Devoluciones       : " & Space(20), 20) & Right(Space(15) & lblDevoluciones.Caption, 15) & Chr(10)
    sCadRep = sCadRep & Space(5) & Left("Saldo Final        : " & Space(20), 20) & Right(Space(15) & lblSaldoF.Caption, 15) & Chr(10) & Chr(10)
    
    sCadRep = sCadRep & Space(5) & Left("Stock Actual       : " & Space(20), 20) & Right(Space(15) & lblStockAct.Caption, 15) & Chr(10) & Chr(10)
    sCadRep = sCadRep & Space(5) & Left("Diferencia         : " & Space(20), 20) & Right(Space(15) & lblDiferencia.Caption, 15) & Chr(10) & Chr(10)
    
'
'
'
'    sCadRep = sCadRep & Space(5) & Space(20) & Left("Saldo Anterior: " & Space(28), 16) & Right(lblSaldoAnt.Caption, 6) & Space(22) & Left("Total de Habilitaciones: " & Space(28), 22) & Right(lblHabilitaciones.Caption, 6) & Chr(10)
'    sCadRep = sCadRep & Space(5) & String(100, " ") & Chr(10)
'    'sCadRep = sCadRep & Space(5) & "Total de Habilitaciones: " & lblHabilitaciones.Caption & Chr(10) & Chr(10)
'    'sCadRep = sCadRep & Space(5) & String(100, " ") & Chr(10)
'    sCadRep = sCadRep & Space(5) & Space(20) & Left("Total de Devoluciones: " & Space(28), 23) & Right(lblDevoluciones.Caption, 6) & Space(15) & Left("Total de Remesas Entrantes: " & Space(28), 28) & Right(lblRemEntrada.Caption, 6) & Chr(10)
'    sCadRep = sCadRep & Space(5) & String(100, " ") & Chr(10)
'    'sCadRep = sCadRep & Space(5) & "Total de Remesas Entrantes: " & lblRemEntrada.Caption & Chr(10) & Chr(10)
'    'sCadRep = sCadRep & Space(5) & String(100, " ") & Chr(10)
'    sCadRep = sCadRep & Space(5) & Space(20) & Left("Saldo Final: " & Space(28), 11) & Right(lblSaldoF.Caption, 6) & Space(25) & Left("Total de Remesas Salientes: " & Space(28), 28) & Right(lblStockAct.Caption, 6) & Chr(10)
'    'sCadRep = sCadRep & Space(5) & String(100, " ") & Chr(10)
'    'sCadRep = sCadRep & Space(5) & "Total Stock Actual: " & lblStockAct.Caption & Chr(10) & Chr(10)
'    sCadRep = sCadRep & Space(5) & String(100, " ") & Chr(10)
'    sCadRep = sCadRep & Space(5) & Space(59) & Left("Stock Actual: " & Space(28), 14) & Right(lblDiferencia.Caption, 6) & Chr(10)
'    sCadRep = sCadRep & Space(5) & Space(20) & String(75, "_") & Chr(10)
'    sCadRep = sCadRep & Space(5) & String(100, " ") & Chr(10)
'    sCadRep = sCadRep & Space(5) & Space(59) & Left("Diferencia: " & Space(28), 12) & Right(lblDiferencia.Caption, 6)
'    sCadRep = sCadRep & Space(5) & String(100, " ") & Chr(10) & Chr(10)
    sCadRep = sCadRep & Space(5) & String(100, "-") & Chr(10) & Chr(10)
    
    Set P = New Previo.clsPrevio
    Call P.Show(sCadRep, "REPORTE")
    Set P = Nothing
    


End Sub

Private Sub CmdProcesar_Click()
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter
Dim R As ADODB.Recordset
    
    Set R = New ADODB.Recordset
    
    oConec.AbreConexion
    Cmd.ActiveConnection = oConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@dFecha", adDate, adParamInput, , Format(lblFecha.Caption, "dd/mm/yyyy"))
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
    Set Prm = Cmd.CreateParameter("@nRemEntrada", adInteger, adParamOutput, 8, 0)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nRemSalida", adInteger, adParamOutput, 8, 0)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nStockActual", adInteger, adParamOutput, 8, 0)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nDiferencia", adInteger, adParamOutput, , 0)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nSaldoFinal", adInteger, adParamOutput, , 0)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nCodAge", adChar, adParamInput, 4, gsCodAge)
    Cmd.Parameters.Append Prm
    
    Cmd.CommandText = "ATM_RegistraCuadreBvda"
    Dim i As Integer
    
    R.CursorType = adOpenStatic
    R.LockType = adLockReadOnly
    Set R = Cmd.Execute
    
    '        For i = 0 To 8 Step 1
    '            MsgBox Cmd.Parameters(i).Value
    '        Next
    lblSaldoAnt = Cmd.Parameters(1).Value
    lblHabilitaciones = Cmd.Parameters(2).Value
    lblDevoluciones = Cmd.Parameters(3).Value
    lblRemEntrada = Cmd.Parameters(4).Value
    lblRemSalida = Cmd.Parameters(5).Value
    lblStockAct = Cmd.Parameters(6).Value
    lblDiferencia = Cmd.Parameters(7).Value
    lblSaldoF = Cmd.Parameters(8).Value
    
    
    'CerrarConexion
    oConec.CierraConexion
    Set oConec = Nothing
    Set Cmd = Nothing
    Set Prm = Nothing
    CmdImprimir.Enabled = True
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Set oConec = New DConecta
    Call CargaFecha
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set oConec = Nothing
End Sub
