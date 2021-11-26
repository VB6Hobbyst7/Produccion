VERSION 5.00
Begin VB.Form frmCuadreBvda 
   Caption         =   "Cuadre Boveda General"
   ClientHeight    =   4455
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6015
   Icon            =   "frmCuadreBvdaGnral.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   6015
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   90
      TabIndex        =   18
      Top             =   3630
      Width           =   5790
      Begin VB.CommandButton CmdImprimir 
         Caption         =   "Imprimir"
         Height          =   390
         Left            =   90
         TabIndex        =   20
         Top             =   210
         Visible         =   0   'False
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
      Left            =   105
      TabIndex        =   4
      Top             =   1260
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
         Top             =   1965
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
         Left            =   4245
         TabIndex        =   15
         Top             =   1005
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
         Left            =   4245
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
         Left            =   1950
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
         Left            =   4230
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
Private Sub CargaFecha()
    Set Cmd = New Command
    Set Prm = New ADODB.Parameter
    
    Set Prm = New ADODB.Parameter
    
    
    Cmd.ActiveConnection = AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_RecuperaFechaSistema"
    
    Set R = Cmd.Execute
    Me.lblFecha.Caption = R!FechaSistema
    R.Close
    Set Cmd = Nothing
    
    Set R = Nothing
    
    CerrarConexion
End Sub

Private Sub CmdProcesar_Click()
    Dim Cmd As New Command
    Dim Prm As New ADODB.Parameter
    Dim R As ADODB.Recordset
    
        Set R = New ADODB.Recordset
        
        Cmd.ActiveConnection = AbrirConexion
        Cmd.CommandType = adCmdStoredProc
        
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@dFecha", adDate, adParamInput, , CDate(lblFecha.Caption))
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
        
        Cmd.CommandText = "ATM_RegistraCuadreBvdaGnral"
        Dim i As Integer

        R.CursorType = adOpenStatic
        R.LockType = adLockReadOnly
        Set R = Cmd.Execute
        
        For i = 0 To 8 Step 1
            MsgBox Cmd.Parameters(i).Value
        Next
        lblSaldoAnt = Cmd.Parameters(1).Value
        lblHabilitaciones = Cmd.Parameters(2).Value
        lblDevoluciones = Cmd.Parameters(3).Value
        lblRemEntrada = Cmd.Parameters(4).Value
        lblRemSalida = Cmd.Parameters(5).Value
        lblStockAct = Cmd.Parameters(6).Value
        lblDiferencia = Cmd.Parameters(7).Value
        lblSaldoF = Cmd.Parameters(8).Value

        
        CerrarConexion
        Set Cmd = Nothing
        Set Prm = Nothing
        'CmdAct.Enabled = True
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call CargaFecha
End Sub
