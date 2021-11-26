VERSION 5.00
Object = "{F9AB04EF-FCD4-4161-99E1-9F65F8191D72}#14.0#0"; "OCXTarjeta.ocx"
Begin VB.Form frmRegCta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Vincular Cuenta - F12 para Digitar Tarjeta"
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10800
   Icon            =   "frmRegCta.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   10800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin OCXTarjeta.CtrlTarjeta CtrlTarjeta1 
      Height          =   420
      Left            =   6210
      TabIndex        =   17
      Top             =   270
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   741
   End
   Begin VB.Frame Frame3 
      Height          =   720
      Left            =   0
      TabIndex        =   5
      Top             =   3840
      Width           =   10755
      Begin VB.CommandButton CmdTarjCta 
         Caption         =   "Tarjeta Cuenta"
         Height          =   360
         Left            =   3645
         TabIndex        =   13
         Top             =   210
         Width           =   1545
      End
      Begin VB.CommandButton CmdNewreg 
         Caption         =   "Nuevo Registro"
         Height          =   360
         Left            =   2400
         TabIndex        =   10
         Top             =   210
         Width           =   1245
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   360
         Left            =   9345
         TabIndex        =   8
         Top             =   210
         Width           =   1320
      End
      Begin VB.CommandButton CmdReg 
         Caption         =   "Registrar"
         Enabled         =   0   'False
         Height          =   360
         Left            =   45
         TabIndex        =   7
         Top             =   210
         Width           =   1155
      End
      Begin VB.CommandButton CmdAct 
         Caption         =   "Eliminar"
         Enabled         =   0   'False
         Height          =   360
         Left            =   1245
         TabIndex        =   6
         Top             =   210
         Width           =   1140
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   15
      TabIndex        =   0
      Top             =   -15
      Width           =   5580
      Begin VB.TextBox TxtNumTarj 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   810
         MaxLength       =   16
         TabIndex        =   12
         Top             =   225
         Visible         =   0   'False
         Width           =   3240
      End
      Begin VB.CommandButton CmdLecTarj 
         Caption         =   "Leer Tarjeta"
         Height          =   345
         Left            =   4065
         TabIndex        =   1
         Top             =   255
         Width           =   1290
      End
      Begin VB.Label Label2 
         Caption         =   "Tarjeta :"
         Height          =   255
         Left            =   60
         TabIndex        =   3
         Top             =   300
         Width           =   735
      End
      Begin VB.Label LblNumTarj 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   825
         TabIndex        =   2
         Top             =   240
         Width           =   3225
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Cuentas"
      Height          =   2955
      Left            =   -15
      TabIndex        =   4
      Top             =   885
      Width           =   10785
      Begin VB.ListBox LstCta 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2370
         ItemData        =   "frmRegCta.frx":030A
         Left            =   150
         List            =   "frmRegCta.frx":030C
         Style           =   1  'Checkbox
         TabIndex        =   11
         Top             =   435
         Width           =   10530
      End
      Begin VB.Label Label6 
         Caption         =   "Tipo de Cuenta"
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
         Height          =   255
         Left            =   8430
         TabIndex        =   18
         Top             =   180
         Width           =   1470
      End
      Begin VB.Label Label5 
         Caption         =   "Registrada"
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
         Height          =   255
         Left            =   6570
         TabIndex        =   9
         Top             =   165
         Width           =   915
      End
      Begin VB.Label Label4 
         Caption         =   "Nro. Titular"
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
         Height          =   255
         Left            =   5265
         TabIndex        =   16
         Top             =   165
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo de Cuenta"
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
         Height          =   255
         Left            =   3285
         TabIndex        =   15
         Top             =   165
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Cuenta"
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
         Height          =   255
         Left            =   795
         TabIndex        =   14
         Top             =   165
         Width           =   660
      End
   End
End
Attribute VB_Name = "frmRegCta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sbIn As String
Dim sResp As String
Dim R As ADODB.Recordset
Dim oConec As DConecta


Private Sub CmdTarjCta_Click()
    frmAdicTarjetaCta.Show 1

End Sub

Private Sub Form_Load()
    Set oConec = New DConecta
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set oConec = Nothing
End Sub

Private Sub TxtNumTarj_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
            LblNumTarj.Caption = TxtNumTarj.Text
            TxtNumTarj.Visible = False
            Me.LblNumTarj.Visible = True
            Me.CmdLecTarj.Visible = True
            Me.Caption = "Vincular Cuenta - F12 para Digitar Tarjeta"

            Dim nCond As Integer
            Dim nRetenerTar As Integer
            Dim nNOOpeMonExt As Integer
            Dim nSuspOper As Integer
            Dim dFecVenc As Date
            Dim sEstadoDesc As String
            
            If Not ExisteTarjeta(Me.LblNumTarj.Caption) Then
                MsgBox "Tarjeta No Existe", vbInformation, "Aviso"
                Exit Sub
            End If
            
            
            Call RecuperaDatosDETarjetas(Me.LblNumTarj.Caption, nCond, nRetenerTar, nNOOpeMonExt, nSuspOper, dFecVenc, sEstadoDesc)
            If nCond <> 1 Then
                MsgBox "Tarjeta NO esta Activa"
                LblNumTarj.Caption = ""
            End If


            If Len(Trim(LblNumTarj.Caption)) > 0 Then
                Call CargaDatos
                Me.CmdReg.Enabled = True
                CmdAct.Enabled = True
            Else
                Me.CmdReg.Enabled = False
                CmdAct.Enabled = False
            End If
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 123 Then
            TxtNumTarj.Text = ""
            TxtNumTarj.Visible = True
            Me.LblNumTarj.Visible = False
            Me.CmdLecTarj.Visible = False
            Me.Caption = "Vincular Cuenta - F12 para Digitar Tarjeta"
            TxtNumTarj.SetFocus
    End If
End Sub
Private Sub CmdAct_Click()

Dim sResp As String
Dim sTramaResp As String

Dim Cmd As New Command
Dim Prm As New ADODB.Parameter
Dim i As Integer

        For i = 0 To Me.LstCta.ListCount - 1
        If Me.LstCta.Selected(i) = True Then
                                               
                Set Prm = New ADODB.Parameter
                Set Prm = Cmd.CreateParameter("@cNumTarjeta", adVarChar, adParamInput, 50, LblNumTarj.Caption)
                Cmd.Parameters.Append Prm
                
                Set Prm = New ADODB.Parameter
                Set Prm = Cmd.CreateParameter("@cCtaCod", adVarChar, adParamInput, 50, Mid(LstCta.List(i), 1, 18))
                Cmd.Parameters.Append Prm
                
                oConec.AbreConexion
                Cmd.ActiveConnection = oConec.ConexionActiva ' AbrirConexion
                Cmd.CommandType = adCmdStoredProc
                Cmd.CommandText = "ATM_EliminaRelacionTarjCta"
                
                Call Cmd.Execute
                
                oConec.CierraConexion
                
                Set Cmd = Nothing
                Set Prm = Nothing
         End If
        Next i
        
        MsgBox "Se ha eliminado correctamente las Cuentas Seleccionadas"
        
        Call CargaDatos
        

End Sub

Private Sub CargaDatos()
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter


    Set R = New ADODB.Recordset
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psNumTarj", adVarChar, adParamInput, 20, Me.LblNumTarj.Caption)
    Cmd.Parameters.Append Prm
    
    oConec.AbreConexion
    Cmd.ActiveConnection = oConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    
    Cmd.CommandText = "ATM_RecuperaCtasTitularTotal"
    
    R.CursorType = adOpenStatic
    R.LockType = adLockReadOnly
    Set R = Cmd.Execute
    LstCta.Clear
    Do While Not R.EOF
        LstCta.AddItem Left(R!Cuenta & Space(25), 25) & Left(R!TipoCta & Space(25), 25) & Left(R!NroTitular & Space(20), 20) & Left(R!cVinculado & Space(15), 15) & Left(R!cTipoPrograma & Space(20), 20)
        R.MoveNext
    Loop
    Me.CmdReg.Enabled = True
    CmdAct.Enabled = True
    oConec.CierraConexion
    
End Sub
Private Sub CmdLecTarj_Click()

Me.Caption = "Registro de Cuenta - PASE LA TARJETA"

LblNumTarj.Caption = Mid(CtrlTarjeta1.LeerTarjeta("PASE LA TARJETA", gnTipoPinPad, gnPinPadPuerto), 2, 16)
Me.Caption = "Registro de Cuenta - F12 para Digitar Tarjeta"
If Not ExisteTarjeta(LblNumTarj.Caption) Then
    LblNumTarj.Caption = ""
    Exit Sub
End If

Dim nCond As Integer
Dim nRetenerTar As Integer
Dim nNOOpeMonExt As Integer
Dim nSuspOper As Integer
Dim dFecVenc As Date
Dim sEstadoDesc As String



Call RecuperaDatosDETarjetas(Me.LblNumTarj.Caption, nCond, nRetenerTar, nNOOpeMonExt, nSuspOper, dFecVenc, sEstadoDesc)
If nCond <> 1 Then
    MsgBox "Tarjeta NO esta Activa"
    LblNumTarj.Caption = ""
    Exit Sub
End If

Call CargaDatos

End Sub

Private Sub LimpiaPantalla()
    LblNumTarj.Caption = ""
    Me.LstCta.Clear
    Me.CmdReg.Enabled = False
End Sub

Private Sub CmdNewreg_Click()
    Call LimpiaPantalla
End Sub

Private Sub CmdReg_Click()
Dim sResp As String
Dim sTramaResp As String
Dim i As Integer

    For i = 0 To Me.LstCta.ListCount - 1
        If Me.LstCta.Selected(i) = True Then
        
            sResp = "00"
            If sResp = "00" Then
                    
                If CuentaVinculada(LblNumTarj.Caption, Mid(LstCta.List(i), 1, 18)) Then
                    MsgBox "Cuenta : " & Mid(LstCta.List(i), 1, 18) & ", Ya esta Vinculada, Se continuará el proceso con las otras cuentas Seleccionadas", vbInformation
                End If
                
                Dim Cmd As New Command
                Dim Prm As New ADODB.Parameter
                   
                Set Prm = New ADODB.Parameter
                Set Prm = Cmd.CreateParameter("@cNumTarjeta", adVarChar, adParamInput, 50, LblNumTarj.Caption)
                Cmd.Parameters.Append Prm
                
                Set Prm = New ADODB.Parameter
                Set Prm = Cmd.CreateParameter("@cCtaCod", adVarChar, adParamInput, 50, Mid(LstCta.List(i), 1, 18))
                Cmd.Parameters.Append Prm
                    
                Set Prm = New ADODB.Parameter
                Set Prm = Cmd.CreateParameter("@nEstado", adInteger, adParamInput, , 0)
                Cmd.Parameters.Append Prm
                    
                Set Prm = New ADODB.Parameter
                Set Prm = Cmd.CreateParameter("@dFecha", adDate, adParamInput, , Now())
                Cmd.Parameters.Append Prm
                    
                Set Prm = New ADODB.Parameter
                Set Prm = Cmd.CreateParameter("@nPrior", adInteger, adParamInput, , 0)
                Cmd.Parameters.Append Prm
                    
                Set Prm = New ADODB.Parameter
                Set Prm = Cmd.CreateParameter("@nRelacionada", adInteger, adParamInput, , 0)
                Cmd.Parameters.Append Prm
                
                oConec.AbreConexion
                Cmd.ActiveConnection = oConec.ConexionActiva ' AbrirConexion
                Cmd.CommandType = adCmdStoredProc
                Cmd.CommandText = "ATM_RegistraTarjetaCuenta"
                
                Call Cmd.Execute
                Set Prm = Nothing
                Set Cmd = Nothing
                
                oConec.CierraConexion
                
            End If
        End If
    Next i
    
    MsgBox "Cuentas Registradas Correctamente"
    Call CargaDatos
    
    CmdAct.Enabled = False
    CmdReg.Enabled = False
End Sub



Private Sub CmdSalir_Click()
    Unload Me
End Sub
