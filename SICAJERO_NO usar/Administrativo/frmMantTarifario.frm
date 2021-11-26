VERSION 5.00
Begin VB.Form frmMantTarifario 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento de Tarifario"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6240
   Icon            =   "frmMantTarifario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Height          =   675
      Left            =   30
      TabIndex        =   19
      Top             =   1260
      Width           =   6150
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   330
         Left            =   4800
         TabIndex        =   11
         Top             =   195
         Width           =   1230
      End
      Begin VB.CommandButton CmdGrabar 
         Caption         =   "Grabar"
         Height          =   330
         Left            =   60
         TabIndex        =   10
         Top             =   210
         Width           =   1230
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Comision por Exceso de Retiros"
      Height          =   1740
      Left            =   30
      TabIndex        =   18
      Top             =   1320
      Visible         =   0   'False
      Width           =   6150
      Begin VB.Frame Frame6 
         Caption         =   "Dolares"
         Height          =   780
         Left            =   105
         TabIndex        =   21
         Top             =   900
         Width           =   5925
         Begin VB.TextBox TxtMonRetDol 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   5115
            TabIndex        =   9
            Text            =   "0"
            Top             =   270
            Width           =   645
         End
         Begin VB.TextBox TxtNumRetMaxDol 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   3765
            TabIndex        =   8
            Text            =   "0"
            Top             =   270
            Width           =   645
         End
         Begin VB.TextBox TxtNumRetMinDol 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   1515
            TabIndex        =   7
            Text            =   "0"
            Top             =   270
            Width           =   645
         End
         Begin VB.Label Label10 
            Caption         =   "Monto :"
            Height          =   195
            Left            =   4515
            TabIndex        =   27
            Top             =   330
            Width           =   600
         End
         Begin VB.Label Label9 
            Caption         =   "Num. Retir. Maximo :"
            Height          =   195
            Left            =   2250
            TabIndex        =   26
            Top             =   330
            Width           =   1500
         End
         Begin VB.Label Label8 
            Caption         =   "Num. Retir. Minimo :"
            Height          =   195
            Left            =   45
            TabIndex        =   25
            Top             =   330
            Width           =   1455
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Soles"
         Height          =   675
         Left            =   105
         TabIndex        =   20
         Top             =   225
         Width           =   5925
         Begin VB.TextBox TxtMonRetSol 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   5160
            TabIndex        =   6
            Text            =   "0"
            Top             =   225
            Width           =   645
         End
         Begin VB.TextBox TxtNumRetMaxSol 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   3810
            TabIndex        =   5
            Text            =   "0"
            Top             =   225
            Width           =   645
         End
         Begin VB.TextBox TxtNumRetMinSol 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   1560
            TabIndex        =   4
            Text            =   "0"
            Top             =   225
            Width           =   645
         End
         Begin VB.Label Label7 
            Caption         =   "Monto :"
            Height          =   195
            Left            =   4560
            TabIndex        =   24
            Top             =   285
            Width           =   600
         End
         Begin VB.Label Label6 
            Caption         =   "Num. Retir. Maximo :"
            Height          =   195
            Left            =   2295
            TabIndex        =   23
            Top             =   285
            Width           =   1500
         End
         Begin VB.Label Label5 
            Caption         =   "Num. Retir. Minimo :"
            Height          =   195
            Left            =   90
            TabIndex        =   22
            Top             =   285
            Width           =   1455
         End
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Comisiones para Operac. No Monetarias"
      Height          =   1140
      Left            =   3015
      TabIndex        =   15
      Top             =   75
      Width           =   3150
      Begin VB.TextBox TxtOpeNoMonComSol 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   1815
         TabIndex        =   2
         Text            =   "0.00"
         Top             =   315
         Width           =   960
      End
      Begin VB.TextBox TxtOpeNoMonComDol 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   1815
         TabIndex        =   3
         Text            =   "0.00"
         Top             =   675
         Width           =   960
      End
      Begin VB.Label Label4 
         Caption         =   "Comision en Soles    :"
         Height          =   255
         Left            =   135
         TabIndex        =   17
         Top             =   300
         Width           =   1590
      End
      Begin VB.Label Label3 
         Caption         =   "Comision en Dolares :"
         Height          =   255
         Left            =   135
         TabIndex        =   16
         Top             =   675
         Width           =   1590
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Comisiones para Operac. Monetarias"
      Height          =   1140
      Left            =   15
      TabIndex        =   12
      Top             =   75
      Width           =   2955
      Begin VB.TextBox TxtOpeMonComDol 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   1815
         TabIndex        =   1
         Text            =   "0.00"
         Top             =   675
         Width           =   960
      End
      Begin VB.TextBox TxtOpeMonComSol 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   1815
         TabIndex        =   0
         Text            =   "0.00"
         Top             =   315
         Width           =   960
      End
      Begin VB.Label Label2 
         Caption         =   "Comision en Dolares :"
         Height          =   255
         Left            =   135
         TabIndex        =   14
         Top             =   675
         Width           =   1590
      End
      Begin VB.Label Label1 
         Caption         =   "Comision en Soles    :"
         Height          =   255
         Left            =   135
         TabIndex        =   13
         Top             =   300
         Width           =   1590
      End
   End
End
Attribute VB_Name = "frmMantTarifario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oConec As DConecta


Private Sub CmdGrabar_Click()

Dim Cmd As New Command
Dim Prm As New ADODB.Parameter
Dim nOpeNOMonComSol As Double
Dim nOpeNOMonComDol As Double
Dim nOpeMonComSol As Double
Dim nOpeMonComDol As Double
Dim nSolNumRetMin As Double
Dim nSolNumRetMax As Double
Dim nSolExeRetMonto As Double
Dim nDolNumRetMin As Double
Dim nDolNumRetMax As Double
Dim nDolExeRetMonto As Double

    If Not IsNumeric(TxtOpeMonComSol.Text) Then
        MsgBox "Monto Incorrecto"
        TxtOpeMonComSol.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(TxtOpeMonComDol.Text) Then
        MsgBox "Monto Incorrecto"
        TxtOpeMonComDol.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(TxtOpeNoMonComSol.Text) Then
        MsgBox "Monto Incorrecto"
        TxtOpeNoMonComSol.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(TxtOpeNoMonComDol.Text) Then
        MsgBox "Monto Incorrecto"
        TxtOpeNoMonComDol.SetFocus
        Exit Sub
    End If

    If Not IsNumeric(TxtNumRetMinSol.Text) Then
        MsgBox "Numero Incorrecto"
        TxtNumRetMinSol.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(TxtNumRetMaxSol.Text) Then
        MsgBox "Numero Incorrecto"
        TxtNumRetMaxSol.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(TxtMonRetSol.Text) Then
        MsgBox "Numero Incorrecto"
        TxtNumRetMaxSol.SetFocus
        Exit Sub
    End If
    
    If Not IsNumeric(TxtNumRetMinDol.Text) Then
        MsgBox "Numero Incorrecto"
        TxtNumRetMinSol.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(TxtNumRetMaxDol.Text) Then
        MsgBox "Numero Incorrecto"
        TxtNumRetMaxSol.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(TxtMonRetDol.Text) Then
        MsgBox "Numero Incorrecto"
        TxtNumRetMaxSol.SetFocus
        Exit Sub
    End If

    nOpeMonComSol = CDbl(TxtOpeMonComSol.Text)
    nOpeMonComDol = CDbl(TxtOpeMonComDol.Text)
    nOpeNOMonComSol = CDbl(TxtOpeNoMonComSol.Text)
    nOpeNOMonComDol = CDbl(TxtOpeNoMonComDol.Text)
    nSolNumRetMin = CInt(TxtNumRetMinSol.Text)
    nSolNumRetMax = CInt(TxtNumRetMaxSol.Text)
    nSolExeRetMonto = CDbl(TxtMonRetSol.Text)
    nDolNumRetMin = CInt(TxtNumRetMinDol.Text)
    nDolNumRetMax = CInt(TxtNumRetMaxDol.Text)
    nDolExeRetMonto = CDbl(TxtMonRetDol.Text)


    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnOpeNOMonComSol", adDouble, adParamInput, , nOpeNOMonComSol)
    Cmd.Parameters.Append Prm
     
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnOpeNOMonComDol", adDouble, adParamInput, , nOpeNOMonComDol)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnOpeMonComSol", adDouble, adParamInput, , nOpeMonComSol)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnOpeMonComDol", adDouble, adParamInput, , nOpeMonComDol)
    Cmd.Parameters.Append Prm
    
     Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnSolNumRetMin", adInteger, adParamInput, , nSolNumRetMin)
    Cmd.Parameters.Append Prm
    
     Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnSolNumRetMax", adInteger, adParamInput, , nSolNumRetMax)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnSolExeRetMonto", adDouble, adParamInput, , nSolExeRetMonto)
    Cmd.Parameters.Append Prm
    
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnDolNumRetMin", adInteger, adParamInput, , nDolNumRetMin)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnDolNumRetMax", adInteger, adParamInput, , nDolNumRetMax)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnDolExeRetMonto", adDouble, adParamInput, , nDolExeRetMonto)
    Cmd.Parameters.Append Prm
    
    oConec.AbreConexion
    Cmd.ActiveConnection = oConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_RegistraTarifario "
    Cmd.Execute
    
    'Call CerrarConexion
    oConec.CierraConexion

    Set Cmd = Nothing
    Set Prm = Nothing
    
    MsgBox "Datos Grabados"
    

End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter
Dim nOpeNOMonComSol As Double
Dim nOpeNOMonComDol As Double
Dim nOpeMonComSol As Double
Dim nOpeMonComDol As Double
Dim nSolNumRetMin As Double
Dim nSolNumRetMax As Double
Dim nSolExeRetMonto As Double
Dim nDolNumRetMin As Double
Dim nDolNumRetMax As Double
Dim nDolExeRetMonto As Double

    Set oConec = New DConecta
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnOpeNOMonComSol", adDouble, adParamOutput)
    Cmd.Parameters.Append Prm
     
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnOpeNOMonComDol", adDouble, adParamOutput)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnOpeMonComSol", adDouble, adParamOutput)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnOpeMonComDol", adDouble, adParamOutput)
    Cmd.Parameters.Append Prm
    
     Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnSolNumRetMin", adInteger, adParamOutput)
    Cmd.Parameters.Append Prm
    
     Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnSolNumRetMax", adInteger, adParamOutput)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnSolExeRetMonto", adDouble, adParamOutput)
    Cmd.Parameters.Append Prm
    
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnDolNumRetMin", adInteger, adParamOutput)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnDolNumRetMax", adInteger, adParamOutput)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnDolExeRetMonto", adDouble, adParamOutput)
    Cmd.Parameters.Append Prm
    
    oConec.AbreConexion
    Cmd.ActiveConnection = oConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_RecuperaTarifario "
    Cmd.Execute
    
    nOpeNOMonComSol = Cmd.Parameters(0).Value
    nOpeNOMonComDol = Cmd.Parameters(1).Value
    nOpeMonComSol = Cmd.Parameters(2).Value
    nOpeMonComDol = Cmd.Parameters(3).Value
    nSolNumRetMin = Cmd.Parameters(4).Value
    nSolNumRetMax = Cmd.Parameters(5).Value
    nSolExeRetMonto = Cmd.Parameters(6).Value
    nDolNumRetMin = Cmd.Parameters(7).Value
    nDolNumRetMax = Cmd.Parameters(8).Value
    nDolExeRetMonto = Cmd.Parameters(9).Value
    
    TxtOpeMonComSol.Text = Format(nOpeMonComSol, "#,0.00")
    TxtOpeMonComDol.Text = Format(nOpeMonComDol, "#,0.00")
    TxtOpeNoMonComSol.Text = Format(nOpeNOMonComSol, "#,0.00")
    TxtOpeNoMonComDol.Text = Format(nOpeNOMonComDol, "#,0.00")
    TxtNumRetMinSol.Text = Format(nSolNumRetMin, "#,0")
    TxtNumRetMaxSol.Text = Format(nSolNumRetMax, "#,0")
    TxtMonRetSol.Text = Format(nSolExeRetMonto, "#,0.00")
    TxtNumRetMinDol.Text = Format(nDolNumRetMin, "#,0")
    TxtNumRetMaxDol.Text = Format(nDolNumRetMax, "#,0")
    TxtMonRetDol.Text = Format(nDolExeRetMonto, "#,0.00")
  
    'Call CerrarConexion
    oConec.CierraConexion
    
    Set Cmd = Nothing
    Set Prm = Nothing
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set oConec = Nothing
End Sub
