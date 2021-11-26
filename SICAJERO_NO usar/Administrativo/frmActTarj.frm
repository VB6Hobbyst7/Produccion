VERSION 5.00
Object = "{F9AB04EF-FCD4-4161-99E1-9F65F8191D72}#18.0#0"; "OCXTarjeta.ocx"
Begin VB.Form frmActTarj 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Activación de Tarjeta - F12 para Digitar Tarjeta"
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5610
   Icon            =   "frmActTarj.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   5610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin OCXTarjeta.CtrlTarjeta Tarjeta 
      Height          =   660
      Left            =   6390
      TabIndex        =   32
      Top             =   390
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   1164
   End
   Begin VB.Frame Frame3 
      Caption         =   "Persona"
      Enabled         =   0   'False
      Height          =   4950
      Left            =   15
      TabIndex        =   7
      Top             =   960
      Width           =   5535
      Begin VB.CheckBox ChkComi 
         Caption         =   "Cobrar Comision de Reposicion"
         Height          =   225
         Left            =   2655
         TabIndex        =   41
         Top             =   3975
         Width           =   2625
      End
      Begin VB.Frame Frame4 
         Enabled         =   0   'False
         Height          =   900
         Left            =   2490
         TabIndex        =   36
         Top             =   4005
         Width           =   3000
         Begin VB.CommandButton CmdSelecCuenta 
            Caption         =   "Seleccionar Cuenta"
            Height          =   285
            Left            =   240
            TabIndex        =   40
            Top             =   555
            Width           =   2520
         End
         Begin VB.OptionButton OptMoneda 
            Caption         =   "Dolares"
            Height          =   195
            Index           =   1
            Left            =   1920
            TabIndex        =   39
            Top             =   315
            Width           =   885
         End
         Begin VB.OptionButton OptMoneda 
            Caption         =   "Soles"
            Height          =   195
            Index           =   0
            Left            =   1020
            TabIndex        =   38
            Top             =   300
            Value           =   -1  'True
            Width           =   720
         End
         Begin VB.Label Label3 
            Caption         =   "Moneda :"
            Height          =   180
            Left            =   135
            TabIndex        =   37
            Top             =   300
            Width           =   720
         End
      End
      Begin VB.TextBox TxtFecExp 
         Enabled         =   0   'False
         Height          =   360
         Left            =   1050
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   3975
         Width           =   1290
      End
      Begin VB.CommandButton CmdBuscarDNI 
         Caption         =   "Buscar Persona"
         Height          =   390
         Left            =   2955
         TabIndex        =   15
         Top             =   225
         Width           =   1545
      End
      Begin VB.TextBox TxtDNI 
         Height          =   360
         Left            =   1065
         MaxLength       =   8
         TabIndex        =   14
         Top             =   270
         Width           =   1755
      End
      Begin VB.TextBox TxtTelef 
         Enabled         =   0   'False
         Height          =   360
         Left            =   1065
         TabIndex        =   13
         Top             =   2535
         Width           =   1425
      End
      Begin VB.OptionButton OptSex 
         Caption         =   "Masculino"
         Enabled         =   0   'False
         Height          =   225
         Index           =   0
         Left            =   3150
         TabIndex        =   12
         Top             =   2160
         Value           =   -1  'True
         Width           =   1185
      End
      Begin VB.OptionButton OptSex 
         Caption         =   "Femenino"
         Enabled         =   0   'False
         Height          =   225
         Index           =   1
         Left            =   4350
         TabIndex        =   11
         Top             =   2175
         Width           =   1035
      End
      Begin VB.TextBox TxtFecNac 
         Enabled         =   0   'False
         Height          =   360
         Left            =   3615
         TabIndex        =   10
         Top             =   2565
         Width           =   1290
      End
      Begin VB.ComboBox CboEstCiv 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmActTarj.frx":030A
         Left            =   1065
         List            =   "frmActTarj.frx":031A
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   3060
         Width           =   2310
      End
      Begin VB.TextBox TxtDirecc 
         Enabled         =   0   'False
         Height          =   345
         Left            =   1065
         TabIndex        =   8
         Top             =   3525
         Width           =   4080
      End
      Begin VB.Label Label1 
         Caption         =   "Fec. Exp   :"
         Height          =   255
         Left            =   45
         TabIndex        =   35
         Top             =   4035
         Width           =   870
      End
      Begin VB.Label Label4 
         Caption         =   "A. Paterno :"
         Height          =   255
         Left            =   75
         TabIndex        =   16
         Top             =   840
         Width           =   915
      End
      Begin VB.Label LblApePat 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   1065
         TabIndex        =   29
         Top             =   765
         Width           =   2370
      End
      Begin VB.Label Label6 
         Caption         =   "A. Materno :"
         Height          =   255
         Left            =   75
         TabIndex        =   28
         Top             =   1275
         Width           =   915
      End
      Begin VB.Label LblApeMat 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   1065
         TabIndex        =   27
         Top             =   1200
         Width           =   2370
      End
      Begin VB.Label Label8 
         Caption         =   "Nombres    :"
         Height          =   255
         Left            =   75
         TabIndex        =   26
         Top             =   1710
         Width           =   915
      End
      Begin VB.Label LblNom 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   1065
         TabIndex        =   25
         Top             =   1635
         Width           =   4200
      End
      Begin VB.Label Label10 
         Caption         =   "DNI :"
         Height          =   255
         Left            =   90
         TabIndex        =   24
         Top             =   345
         Width           =   480
      End
      Begin VB.Label Label11 
         Caption         =   "DNI            :"
         Height          =   255
         Left            =   75
         TabIndex        =   23
         Top             =   2115
         Width           =   915
      End
      Begin VB.Label LblDNI 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   1065
         TabIndex        =   22
         Top             =   2070
         Width           =   1320
      End
      Begin VB.Label Label13 
         Caption         =   "Sexo :"
         Height          =   255
         Left            =   2565
         TabIndex        =   21
         Top             =   2145
         Width           =   525
      End
      Begin VB.Label Label15 
         Caption         =   "Telefono :"
         Height          =   255
         Left            =   75
         TabIndex        =   20
         Top             =   2580
         Width           =   840
      End
      Begin VB.Label Label14 
         Caption         =   "Fec. Nacim:"
         Height          =   255
         Left            =   2655
         TabIndex        =   19
         Top             =   2610
         Width           =   870
      End
      Begin VB.Label Label16 
         Caption         =   "Estado Civil :"
         Height          =   255
         Left            =   75
         TabIndex        =   18
         Top             =   3090
         Width           =   945
      End
      Begin VB.Label Label17 
         Caption         =   "Direccion   :"
         Height          =   255
         Left            =   75
         TabIndex        =   17
         Top             =   3600
         Width           =   915
      End
   End
   Begin VB.Frame Frame2 
      Height          =   645
      Left            =   30
      TabIndex        =   4
      Top             =   5955
      Width           =   5505
      Begin VB.CommandButton CmdNuevaAct 
         Caption         =   "Nueva Activac."
         Enabled         =   0   'False
         Height          =   375
         Left            =   3120
         TabIndex        =   33
         Top             =   150
         Width           =   1305
      End
      Begin VB.CommandButton CmdRegCta 
         Caption         =   "Registro de Cuenta"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1455
         TabIndex        =   30
         Top             =   150
         Width           =   1620
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   375
         Left            =   4440
         TabIndex        =   6
         Top             =   150
         Width           =   1005
      End
      Begin VB.CommandButton CmdActTar 
         Caption         =   "Activar Tarjeta"
         Enabled         =   0   'False
         Height          =   375
         Left            =   105
         TabIndex        =   5
         Top             =   150
         Width           =   1305
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   30
      TabIndex        =   0
      Top             =   90
      Width           =   5505
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
         Left            =   780
         MaxLength       =   16
         TabIndex        =   31
         Top             =   210
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
      Begin VB.Label Lblnumtarjeta 
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
         Left            =   795
         TabIndex        =   2
         Top             =   240
         Width           =   3225
      End
   End
End
Attribute VB_Name = "frmActTarj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sPersCod As String
Dim psCtaCod As String
Dim sNumTarjAnt As String
Dim oConec As DConecta

Private Function VerificaSITarjetaActiva() As Boolean


If TarjetaActiva(Me.Lblnumtarjeta.Caption) Then
    VerificaSITarjetaActiva = True
Else
    VerificaSITarjetaActiva = False
End If

End Function

Private Sub ChkComi_Click()
    If ChkComi.Value = 1 Then
        Frame4.Enabled = True
    Else
        Frame4.Enabled = False
        
    End If
End Sub

Public Function CobrarComision() As Integer
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter
Dim nRes As Integer

                
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psCtaCod", adVarChar, adParamInput, 18, psCtaCod)
    Cmd.Parameters.Append Prm
    
         
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psOpeCod", adVarChar, adParamInput, 6, "208025")
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psOpeCodITF", adVarChar, adParamInput, 6, "208023")
    Cmd.Parameters.Append Prm
        
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnResultado", adInteger, adParamOutput)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psUser", adVarChar, adParamInput, 20, gsCodUser)
    Cmd.Parameters.Append Prm
    
    oConec.AbreConexion
    
    Cmd.ActiveConnection = oConec.ConexionActiva 'AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    
    Cmd.CommandText = "ATM_RetiroReposicion"

    Cmd.Execute
        
        
    nRes = Cmd.Parameters(3).Value
    
    
    CobrarComision = nRes

    Set Cmd = Nothing
    Set Prm = Nothing
        
    oConec.CierraConexion

    
End Function

Public Function VerificarSaldoParaCobrarComision() As Integer
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter
Dim nRes As Integer
                
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psCtaCod", adVarChar, adParamInput, 18, psCtaCod)
    Cmd.Parameters.Append Prm
         
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psOpeCod", adVarChar, adParamInput, 6, "")
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psOpeCodITF", adVarChar, adParamInput, 6, "")
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnResultado", adInteger, adParamOutput)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psUser", adVarChar, adParamInput, 20, gsCodUser)
    Cmd.Parameters.Append Prm
    
    oConec.AbreConexion
    Cmd.ActiveConnection = oConec.ConexionActiva 'AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    
    Cmd.CommandText = "ATM_RetiroReposicion"
    
    Cmd.Execute
    
    nRes = Cmd.Parameters(3).Value
    
    VerificarSaldoParaCobrarComision = nRes
    
    Set Cmd = Nothing
    Set Prm = Nothing
    
    oConec.CierraConexion
        
End Function



Private Sub CmdActTar_Click()
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter
Dim sResp As String
Dim sTramaResp As String
Dim nResulComi As Integer
Dim dFecExp As Date 'DAOR 20081219

    If Not IsDate(Me.TxtFecExp.Text) Then
        MsgBox "Fecha Invalida"
        Exit Sub
    End If
    
    If VerificaSITarjetaActiva Then
        MsgBox "Tarjeta ya Esta Activa", vbInformation
            Me.Frame3.Enabled = False
            CmdActTar.Enabled = False
            CmdRegCta.Enabled = False
            Me.TxtFecExp.Text = ""
            Me.TxtFecExp.Enabled = False
        LimpiaDatos
        Exit Sub
    End If
    
    If Len(Trim(psCtaCod)) = 0 And Me.ChkComi.Value = 1 Then
        MsgBox "Cuenta de Ahorros para cobro de Comision de Reposición es Invalida", vbInformation, "Aviso"
          Me.Frame3.Enabled = False
            CmdActTar.Enabled = False
            CmdRegCta.Enabled = False
            Me.TxtFecExp.Text = ""
            Me.TxtFecExp.Enabled = False
        LimpiaDatos
        Exit Sub
    End If
    

    If VerificarSaldoParaCobrarComision = 2 Then
        MsgBox "Cuenta de Ahorros No Tiene Saldo para cobro de Comision de Reposición ", vbInformation, "Aviso"
        Me.Frame3.Enabled = False
        CmdActTar.Enabled = False
        CmdRegCta.Enabled = False
        Me.TxtFecExp.Text = ""
        Me.TxtFecExp.Enabled = False
        LimpiaDatos
        Exit Sub
    End If
    
    sResp = "00"
    
    If sResp = "00" Then
        Set Cmd = New ADODB.Command
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@cNumtarjeta", adVarChar, adParamInput, 16, Lblnumtarjeta.Caption)
        Cmd.Parameters.Append Prm
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@cPersCod", adVarChar, adParamInput, 20, sPersCod)
        Cmd.Parameters.Append Prm
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@dFecActivacion", adDate, adParamInput, , Now)
        Cmd.Parameters.Append Prm
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("cUserActiv", adChar, adParamInput, 4, gsCodUser)
        Cmd.Parameters.Append Prm
        
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@pnResult", adInteger, adParamOutput)
        Cmd.Parameters.Append Prm
        
        '**DAOR 20081219 ****************************************************************
        dFecExp = DateAdd("d", -1, DateAdd("m", 1, CDate("01/" & Left(Me.TxtFecExp.Text, 2) & "/20" & Mid(Me.TxtFecExp.Text, 4, 2))))
        '********************************************************************************

        Set Prm = New ADODB.Parameter
        '**Modifcado por DAOR 20081219 ****************************************************************
        'Set Prm = Cmd.CreateParameter("@pdFecExp", adDate, adParamInput, , CDate(Me.TxtFecExp.Text))
        Set Prm = Cmd.CreateParameter("@pdFecExp", adDate, adParamInput, , dFecExp)
        '**********************************************************************************************
        Cmd.Parameters.Append Prm
        
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@nCodAge", adInteger, adParamInput, , CInt(gsCodAge))
        Cmd.Parameters.Append Prm
        
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@cCodUsu", adVarChar, adParamInput, 20, gsCodUser)
        Cmd.Parameters.Append Prm
        
        oConec.AbreConexion
        Cmd.ActiveConnection = oConec.ConexionActiva 'AbrirConexion
        Cmd.CommandType = adCmdStoredProc
        Cmd.CommandText = "ATM_ActivarTarjeta"
        Cmd.Execute
        
        If Cmd.Parameters(4).Value = 0 Then
            If sResp = "00" Then
                nResulComi = CobrarComision
                If nResulComi <> 0 And Me.ChkComi.Value = 1 Then
                    If nResulComi <> 2 Then
                        MsgBox "No se pudo Realizar el Cargo de la Comisión, Comuniquese con Sistema", vbInformation, "Aviso"
                    Else
                        MsgBox "Cuenta NO Tienes Saldo para Realizar el Cargo de la Operación", vbInformation, "Aviso"
                    End If
                    LimpiaDatos
                    Exit Sub
                End If
                Set Prm = Nothing
                Set Cmd = Nothing
                
                CmdActTar.Enabled = False
                
                MsgBox "Tarjeta Activada"
                
            Else
                MsgBox "NO SE PUEDO ACTIVAR LA TARJETA, Por favor Comuniquese con TI"
            End If
        Else
            MsgBox "Tarjeta ya se encuentra Activada"
        End If
        
        'Call CerrarConexion
        oConec.CierraConexion
    Else
        MsgBox "NO SE PUEDO ACTIVAR LA TARJETA, Por favor Comuniquese con TI"
    End If
    
Set Prm = Nothing
Set Cmd = Nothing
                
End Sub

Private Sub CmdBuscarDNI_Click()
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter

    psCtaCod = ""
    
    If Trim(TxtDNI.Text) = "" Or Len(Trim(TxtDNI.Text)) <> 8 Then
      MsgBox "Debe ingresar un número de DNI Valido. Verifique", vbInformation
      Exit Sub
    End If
    
    Set Cmd = New ADODB.Command
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psDNI", adVarChar, adParamInput, 50, TxtDNI.Text)
    Cmd.Parameters.Append Prm

    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psApePat", adVarChar, adParamOutput, 50)
    Cmd.Parameters.Append Prm

    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psApeMat", adVarChar, adParamOutput, 50)
    Cmd.Parameters.Append Prm

    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psNombres", adVarChar, adParamOutput, 50)
    Cmd.Parameters.Append Prm

    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psSexo", adVarChar, adParamOutput, 50)
    Cmd.Parameters.Append Prm

    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psTelef", adVarChar, adParamOutput, 50)
    Cmd.Parameters.Append Prm

    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pdFecNac", adDate, adParamOutput)
    Cmd.Parameters.Append Prm

    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psEstCiv", adVarChar, adParamOutput, 50)
    Cmd.Parameters.Append Prm

    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psDirecc", adVarChar, adParamOutput, 50)
    Cmd.Parameters.Append Prm

    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psPersCodTar", adVarChar, adParamOutput, 50)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psNumTarjeta", adVarChar, adParamOutput, 20)
    Cmd.Parameters.Append Prm

    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@pnNumTarjAcum", adInteger, adParamOutput)
    Cmd.Parameters.Append Prm

    oConec.AbreConexion
    Cmd.ActiveConnection = oConec.ConexionActiva 'AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_RecuperaDatosPersona"
    Cmd.Execute
    
    If Trim(Cmd.Parameters(9).Value) <> "" Then
        Me.LblApePat.Caption = Replace(UCase(Cmd.Parameters(1).Value), "Ñ", "#")
        Me.LblApeMat.Caption = Replace(UCase(Cmd.Parameters(2).Value), "Ñ", "#")
        Me.LblNom.Caption = Replace(UCase(Cmd.Parameters(3).Value), "Ñ", "#")
        Me.OptSex(0).Value = IIf(Cmd.Parameters(4).Value = "M", True, False)
        Me.OptSex(1).Value = IIf(Cmd.Parameters(4).Value = "F", True, False)
        Me.TxtTelef.Text = Cmd.Parameters(5).Value
        Me.TxtFecNac.Text = Cmd.Parameters(6).Value
        Me.CboEstCiv.Text = IIf(Cmd.Parameters(7).Value = "1", "SOLTERO", _
                IIf(Cmd.Parameters(7).Value = "2", "CASADO", _
                IIf(Cmd.Parameters(7).Value = "3", "VIUDO", "DIVORCIADO")))
        LblDNI.Caption = Cmd.Parameters(0).Value
        Me.TxtDirecc.Text = Replace(UCase(Cmd.Parameters(8).Value), "Ñ", "#")
        sPersCod = Cmd.Parameters(9).Value
        CmdActTar.Enabled = True
        CmdRegCta.Enabled = True
        
        sNumTarjAnt = ""
        
        '@psNumTarjeta con Valor
        If Len(Trim(Cmd.Parameters(10).Value)) > 0 Then
            sNumTarjAnt = Cmd.Parameters(10).Value
            MsgBox "Cliente ya Posee Tarjeta"
            
            Call LimpiaDatos
            CmdActTar.Enabled = False
            CmdRegCta.Enabled = False
            Me.TxtFecExp.Enabled = False
            
            Set Cmd = Nothing
            Set Prm = Nothing
            
            Exit Sub
        End If
        
        
        If Cmd.Parameters(11).Value > 0 Then
            MsgBox "Cliente ya a solicitado Tarjeta anteriormente, Se le cobrará el cargo de Reposición al Activar la Nueva Tarjeta "
            Me.ChkComi.Value = 1
            Me.ChkComi.Enabled = False
            
        Else
            Me.ChkComi.Value = 0
            Me.ChkComi.Enabled = False
            
        End If
                                    
        Me.TxtFecExp.Enabled = False
        CmdActTar.Enabled = True
        CmdRegCta.Enabled = True
        
    Else
        MsgBox "El DNI puesto no existe en el sistema. Verifique", vbInformation
        CmdActTar.Enabled = False
        CmdRegCta.Enabled = False
        Me.TxtFecExp.Enabled = False
    End If
    
    'Call CerrarConexion
    oConec.CierraConexion

    Set Cmd = Nothing
    Set Prm = Nothing


End Sub


Private Sub LimpiaDatos()
    Me.LblApeMat.Caption = ""
    Me.LblApePat.Caption = ""
    Me.LblDNI.Caption = ""
    Me.LblNom.Caption = ""
    Me.Lblnumtarjeta.Caption = ""
    Me.TxtDirecc.Text = ""
    Me.TxtDNI.Text = ""
    Me.TxtFecNac.Text = ""
    Me.TxtTelef.Text = ""
    
    psCtaCod = ""
    

End Sub

Private Function RecuperafechaExp() As String
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter
Dim nRes As Integer
        
                
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psNumTarjeta", adVarChar, adParamInput, 20, Lblnumtarjeta.Caption)
    Cmd.Parameters.Append Prm
    
         
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psFechaVenc", adVarChar, adParamOutput, 10)
    Cmd.Parameters.Append Prm
    
    oConec.AbreConexion

    Cmd.ActiveConnection = oConec.ConexionActiva 'AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    
    Cmd.CommandText = "ATM_RecuperaFechaVencTarjeta"
    
    Cmd.Execute
        
    RecuperafechaExp = Cmd.Parameters(1).Value

    oConec.CierraConexion
    
    Set Cmd = Nothing
    Set Prm = Nothing

End Function
Private Sub CmdLecTarj_Click()
   
Me.Caption = "Activación de Tarjeta - PASE LA TARJETA"


Lblnumtarjeta.Caption = Mid(Tarjeta.LeerTarjeta("PASE LA TARJETA", gnTipoPinPad, gnPinPadPuerto), 2, 16)
Frame3.Enabled = True
Me.Caption = "Activación de Tarjeta - F12 para Digitar Tarjeta"

If Not ExisteTarjeta(Lblnumtarjeta.Caption) Then
    Lblnumtarjeta.Caption = ""
    Exit Sub
End If

If VerificaSITarjetaActiva Then
    MsgBox ("Tarjeta ya Esta Activa")
      Me.Frame3.Enabled = False
            CmdActTar.Enabled = False
            CmdRegCta.Enabled = False
        LimpiaDatos
        Exit Sub
        
End If

If Not ExisteTarjetaEmitida(Lblnumtarjeta.Caption) Then
    Lblnumtarjeta.Caption = ""
    MsgBox "La tarjeta no se encuentra emitida o está pendiente de confirmación, consulte con el supervisor de agencia", vbInformation
    Exit Sub
End If




Me.TxtFecExp.Text = RecuperafechaExp


End Sub

Private Sub CmdNuevaAct_Click()
    Me.Frame3.Enabled = False
    CmdActTar.Enabled = False
    CmdRegCta.Enabled = False
    LimpiaDatos
End Sub

Private Sub CmdRegCta_Click()
    frmRegCta.Show 1
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub



Private Sub CmdSelecCuenta_Click()
    
    psCtaCod = frmSelectCuenta.Seleccionar(IIf(Me.OptMoneda(0).Value, 1, 2), sPersCod)
    
    If Len(Trim(psCtaCod)) = 0 Then
        MsgBox "Cuenta de Ahorros para Cargo de Comsion Invalida", vbInformation, "Aviso"
        
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 123 Then
            TxtNumTarj.Text = ""
            TxtNumTarj.Visible = True
            Me.Lblnumtarjeta.Visible = False
            Me.CmdLecTarj.Visible = False
            Frame3.Enabled = False
            Me.Caption = "Activación de Tarjeta - F12 para Digitar Tarjeta"
            TxtNumTarj.SetFocus
    End If
End Sub


Private Sub Form_Load()
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter
Dim R As ADODB.Recordset

    Set oConec = New DConecta
    
    
    Set Cmd = New ADODB.Command
    oConec.AbreConexion
    Cmd.ActiveConnection = oConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_RecuperaEstadoCivil"
    Set R = Cmd.Execute
    Me.CboEstCiv.Clear
    
    Do While Not R.EOF
        CboEstCiv.AddItem R!cConsDescripcion
        R.MoveNext
    Loop
    R.Close
    
    'Call CerrarConexion
    oConec.CierraConexion
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set oConec = Nothing
End Sub


Private Sub TxtNumTarj_KeyPress(KeyAscii As Integer)
    
    
    If KeyAscii = 13 Then
            Lblnumtarjeta.Caption = TxtNumTarj.Text
            TxtNumTarj.Visible = False
            Me.Lblnumtarjeta.Visible = True
            Me.CmdLecTarj.Visible = True
            Me.Caption = "Activación de Tarjeta - F12 para Digitar Tarjeta"
            If Not ExisteTarjeta(Lblnumtarjeta.Caption) Then
                Lblnumtarjeta.Caption = ""
                Exit Sub
            End If
            If Len(Trim(Lblnumtarjeta.Caption)) > 0 Then
            
                If VerificaSITarjetaActiva Then
                    MsgBox ("Tarjeta ya Esta Activa")
                    Me.Frame3.Enabled = False
                    CmdActTar.Enabled = False
                    CmdRegCta.Enabled = False
                    LimpiaDatos
                    Exit Sub
                End If
                
                If Not ExisteTarjetaEmitida(Lblnumtarjeta.Caption) Then
                        MsgBox "La tarjeta no se encuentra emitida o está pendiente de confirmación, consulte con el supervisor de agencia", vbInformation
                    Exit Sub
                End If
                Frame3.Enabled = True
                CmdBuscarDNI.SetFocus
                Me.TxtFecExp.Text = RecuperafechaExp
            Else
                Frame3.Enabled = False
            End If
    End If
End Sub
