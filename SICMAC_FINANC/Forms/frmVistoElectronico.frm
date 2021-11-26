VERSION 5.00
Begin VB.Form frmVistoElectronico 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Visto Electrónico"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5730
   ForeColor       =   &H8000000D&
   Icon            =   "frmVistoElectronico.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   5730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraVisto 
      Caption         =   " Visto Electrónico "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5475
      Begin VB.TextBox txtCargo 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C16A0B&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1080
         TabIndex        =   11
         ToolTipText     =   "Ingrese su Usuario"
         Top             =   1920
         Width           =   4215
      End
      Begin VB.TextBox txtNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C16A0B&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1080
         TabIndex        =   10
         ToolTipText     =   "Ingrese su Usuario"
         Top             =   1560
         Width           =   4215
      End
      Begin VB.TextBox txtComentario 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C16A0B&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   3
         ToolTipText     =   "Ingrese su Usuario"
         Top             =   960
         Width           =   4215
      End
      Begin VB.TextBox txtUsuario 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C16A0B&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1080
         MaxLength       =   4
         TabIndex        =   1
         ToolTipText     =   "Ingrese su Usuario"
         Top             =   240
         Width           =   2430
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "&Validar"
         Height          =   375
         Left            =   4320
         TabIndex        =   4
         Top             =   360
         Width           =   1000
      End
      Begin VB.TextBox TxtClave 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C16A0B&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1080
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   2
         ToolTipText     =   "Ingrese su Clave Secreta"
         Top             =   600
         Width           =   2430
      End
      Begin VB.Label lblCargo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cargo :"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   1920
         Width           =   510
      End
      Begin VB.Label lblNombre 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre :"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   1560
         Width           =   645
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000000&
         X1              =   120
         X2              =   5280
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Label lblComentario 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Comentario :"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   885
      End
      Begin VB.Label lblClave 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Clave     :"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   675
      End
      Begin VB.Label lblusuario 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario :"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   255
         Width           =   630
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4560
      TabIndex        =   8
      Top             =   2520
      Width           =   1000
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Aceptar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3120
      TabIndex        =   7
      Top             =   2520
      Width           =   1000
   End
End
Attribute VB_Name = "frmVistoElectronico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************
'***     Rutina:              frmVistoElectronico
'***     Descripcion:         Permite registrar el visto o aprobacion de una operacion por el supervisor o jefe inmediato
'***     Creado por:          PEAC
'***     Maquina:             PCSIS010020
'***     Fecha-Tiempo:        20080807 10:00:00 AM
'***     Ultima Modificacion: Creacion del Formulario
'*****************************************************************************************
Option Explicit

Dim sVisPersCod As String
Dim sVisPersNom As String
Dim sVisPersCar As String

Dim sMensaje As String 'Add Gitu and Peac 22-08-08

Private fsVistoUserCod As String
Private fnVistoCod As Long
Private fsOpeCod As String
Private fsClientePersCod As String
Private fsVistoPersCod As String
Private fsVistoComentario As String
Private fsVistoMovNro As String
Private fnMovNro As Long
Private fnCasoCod As Integer
'madm 2010909 -----------------------------------------
Public VistoxOpera As Boolean
Public lsCodUserOpera As String
'madm 2010909 -----------------------------------------
'**Valores para fnCasoCod: *****************************
' 00:
' 01: Visto para dudosos  e historial negativo
' 02:
'*******************************************************

Public bResultadoVisto As Boolean

Property Let VistoCodigo(pnVistoCod As Long)
   fnVistoCod = pnVistoCod
End Property
Property Get VistoCodigo() As Long
    VistoCodigo = fnVistoCod
End Property

Property Let ClientePersCod(psClientePersCod As String)
   fsClientePersCod = psClientePersCod
End Property
Property Get ClientePersCod() As String
    ClientePersCod = fsClientePersCod
End Property

Property Let CasoCodigo(pnCasoCod As Integer)
   fnCasoCod = pnCasoCod
End Property
Property Get CasoCodigo() As Integer
    CasoCodigo = fnCasoCod
End Property

Property Let OperacionCodigo(psOpecod As String)
   fsOpeCod = psOpecod
End Property
Property Get OperacionCodigo() As String
    OperacionCodigo = fsOpeCod
End Property

Property Let PersonaCodigo(psVistoPersCod As String)
   fsVistoPersCod = psVistoPersCod
End Property
Property Get PersonaCodigo() As String
    PersonaCodigo = fsVistoPersCod
End Property

Property Let VistoComentario(psVistoComentario As String)
   fsVistoComentario = psVistoComentario
End Property
Property Get VistoComentario() As String
    VistoComentario = fsVistoComentario
End Property

Property Let VistoMovNro(psVistoMovNro As String)
   fsVistoMovNro = psVistoMovNro
End Property
Property Get VistoMovNro() As String
    VistoMovNro = fsVistoMovNro
End Property

Property Let MovNro(pnMovNro As Long)
   fnMovNro = pnMovNro
End Property
Property Get MovNro() As Long
    MovNro = fnMovNro
End Property

Private Sub cmdAceptar_Click()
Dim oAcceso As COMDPersona.UCOMAcceso
Set oAcceso = New COMDPersona.UCOMAcceso

    'Select Case fnCasoCod
    '    Case 0
            
    '    Case 1
    '***Modificado por ELRO el 20120403
    'If Not oAcceso.VistoElectronicoEsCorrecto(txtUsuario.Text, TxtClave.Text, sVisPersCod, fnCasoCod, sVisPersNom, sVisPersCar, sMensaje) Then
    '        MsgBox (sMensaje)
    '        Me.txtNombre.Text = ""
    '        Me.txtCargo.Text = ""
    '        fsVistoPersCod = ""
    '        fsVistoUserCod = ""
    '    Else
    '        MsgBox ("Visto satisfactorio, proceda a registrar ")
    '        cmdGrabar.Enabled = True
    '
    '        Me.txtNombre.Text = sVisPersNom
    '        Me.txtCargo.Text = sVisPersCar
    '        fsVistoPersCod = sVisPersCod
    '        fsVistoUserCod = UCase(txtUsuario.Text)
    '    End If
    If Len(Trim(txtUsuario)) > 0 And Len(Trim(TxtClave)) > 0 Then
        If Not oAcceso.VistoElectronicoEsCorrecto(txtUsuario.Text, TxtClave.Text, sVisPersCod, fnCasoCod, sVisPersNom, sVisPersCar, sMensaje) Then
            MsgBox (sMensaje)
            Me.txtNombre.Text = ""
            Me.txtCargo.Text = ""
            fsVistoPersCod = ""
            fsVistoUserCod = ""
        Else
            MsgBox ("Visto satisfactorio, proceda a registrar ")
            cmdGrabar.Enabled = True
                                    
            Me.txtNombre.Text = sVisPersNom
            Me.txtCargo.Text = sVisPersCar
            fsVistoPersCod = sVisPersCod
            fsVistoUserCod = UCase(txtUsuario.Text)
        End If
    ElseIf Len(Trim(txtUsuario)) = 0 Then
        Call MsgBox("Falta ingregar su usuario", vbInformation, "Aviso")
        txtUsuario.SetFocus
        Set oAcceso = Nothing
        Exit Sub
    ElseIf Len(Trim(TxtClave)) = 0 Then
        Call MsgBox("Falta ingregar su clave", vbInformation, "Aviso")
        txtUsuario.SetFocus
        Set oAcceso = Nothing
        Exit Sub
    End If
    '***Fin Modificado por ELRO********
    CmdAceptar_KeyPress (13)
    Set oAcceso = Nothing
            
    'End Select
End Sub

Private Sub CmdAceptar_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{Tab}", True
End If
End Sub

Private Sub cmdCancelar_Click()
    
    Call CancelaVisto
    
End Sub

Private Sub cmdGrabar_Click()
    'Select Case fnCasoCod
    '    Case 0
            
    '    Case 1
    'Modify by Gitu
    If fnCasoCod = 0 Then
    
    ElseIf fnCasoCod = 3 Or fnCasoCod = 4 Then
        If Len(Me.txtComentario.Text) > 0 Then
        fsVistoComentario = RTrim(Me.txtComentario.Text)
        'madm 20100915 - operacion vb peps   Or fsOpeCod = "300503"
        'MIOL 20120914, SEGUN RQ12270 SE AGREGO Or fsOpeCod = "901013" Y "901017"************************
        If fsOpeCod = "910000" Or fsOpeCod = "300503" Or fsOpeCod = "901013" Or fsOpeCod = "901017" Then
        'END MIOL ***************************************************************************************
            RegistraVistoElectronico (123456789) 'MADM 20100909
        End If
        'end madm
        Else
            MsgBox "Complete el Comentario a Registrar"
            bResultadoVisto = False
            Exit Sub
        End If
        bResultadoVisto = True
        Unload Me
    Else
        If Len(Me.txtNombre.Text) > 0 And Len(Me.txtCargo.Text) > 0 Then
            bResultadoVisto = True
            fsVistoComentario = UCase(Trim(txtComentario.Text))
            Me.txtNombre.Text = ""
            Me.txtCargo.Text = ""
            Unload Me
        Else
            bResultadoVisto = False
            fsVistoPersCod = ""
            fsVistoComentario = ""
            Me.txtNombre.Text = ""
            Me.txtCargo.Text = ""
        End If
    End If
    'End gitu
    'End Select
End Sub


Private Sub Form_Load()
If VistoxOpera Then
      Dim oAcceso As COMDPersona.UCOMAcceso 'MADM 20100909
      Set oAcceso = New COMDPersona.UCOMAcceso 'MADM 20100909

      If oAcceso.VistoElectronicoUser(lsCodUserOpera, TxtClave.Text, sVisPersCod, fnCasoCod, sVisPersNom, sVisPersCar, sMensaje) Then
                CmdAceptar.Enabled = False
                TxtClave.Enabled = False
                txtUsuario.Enabled = False
                Me.txtComentario.Enabled = True
                txtUsuario.Text = lsCodUserOpera
                TxtClave.Text = "****"
                
                Me.txtNombre.Text = sVisPersNom
                Me.txtCargo.Text = sVisPersCar
                
                fsVistoPersCod = sVisPersCod
                fsVistoUserCod = UCase(txtUsuario.Text)
                
                cmdGrabar.Enabled = True
        End If
    Set oAcceso = Nothing
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    VistoxOpera = False
End Sub

Private Sub TxtClave_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{Tab}", True
End If
End Sub

Private Sub txtComentario_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{Tab}", True
End If
End Sub

Private Sub txtUsuario_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{Tab}", True
End If
End Sub

Public Function Inicio(ByVal pnCodCaso As Integer, Optional ByVal psOpecod As String = "", Optional ByVal psClientePersCod As String = "", Optional psCodUser As String = "") As Boolean
Dim i As Integer
Dim loDCred As COMDCredito.DCOMCredito
Dim lrs As ADODB.Recordset

    fnCasoCod = pnCodCaso
    fsClientePersCod = psClientePersCod
    fsOpeCod = psOpecod
    
    Select Case fnCasoCod
        Case 0
        
        Case 1:
            If Len(fsClientePersCod) = 0 Then
                CancelaVisto
            End If
        
            Set loDCred = New COMDCredito.DCOMCredito
            Set lrs = loDCred.RecuperaClienteHisNegativo(fsClientePersCod)
            Set loDCred = Nothing
            
            If Not (lrs.BOF And lrs.EOF) Then
                If lrs!cOrden = "1" Then ' BD Dudosos lav dinero
                    If lrs!nEstado = 2 Then ' condicion fraudulento
                        MsgBox "La persona " & Trim(lrs!cpersnombre) & " está en condición " & Trim(lrs!cEstado) & ", NO se podrá continuar.", vbInformation, "Aviso"
                        Call CancelaVisto
                    Else
                        MsgBox "La persona " & Trim(lrs!cpersnombre) & " está en condición " & Trim(lrs!cEstado) & ", para continuar deberá pasar por el Visto Electrónico.", vbInformation, "Aviso"
                        Me.Show 1
                        If Not Me.bResultadoVisto Then
                            Call CancelaVisto
                        End If
                    End If
                Else '2 judi/casti
                    MsgBox "La persona " & Trim(lrs!cpersnombre) & " tiene créditos en estado " & Trim(lrs!cEstado) & ", para continuar deberá pasar por el Visto Electrónico.", vbInformation, "Aviso"
                    Me.Show 1
                    If Not Me.bResultadoVisto Then
                        Call CancelaVisto
                    End If
                End If
            Else
                Call PasaVisto
            End If
            
                Inicio = Me.bResultadoVisto
        'Add By Gitu 22-08-08
        Case 2
            If fsOpeCod = "100902" Then 'JUEZ 20130417
                MsgBox "Dias a reprogramar mayores a los 29 dias, para continuar deberá pasar por el Visto Electrónico.", vbInformation, "Aviso"
            End If
            Me.Show 1
            If Not Me.bResultadoVisto Then
                Call CancelaVisto
            End If
            Inicio = Me.bResultadoVisto
        'End Gitu
        
        Case 3, 6 '*** PEAC 20081002
            Me.Show 1
            If Not Me.bResultadoVisto Then
                Call CancelaVisto
            End If
            Inicio = Me.bResultadoVisto
       Case 4 '*** MADM 20100909
            lsCodUserOpera = psCodUser
            VistoxOpera = True
            Me.Show 1
            Inicio = Me.bResultadoVisto
       Case 5 '*** GITU 23-08-2011 Para control de operaciones sin tarjeta
            lsCodUserOpera = psCodUser
            Me.Show 1
            If Not Me.bResultadoVisto Then
                Call CancelaVisto
            End If
            Inicio = Me.bResultadoVisto
'        Case 6 '** Juez 20120829
'            Me.Show 1
'            If Not Me.bResultadoVisto Then
'                Call CancelaVisto
'            End If
'            Inicio = Me.bResultadoVisto
        Case 7 'EJVG20131115
            lsCodUserOpera = psCodUser
            Me.Show 1
            If Not Me.bResultadoVisto Then
                Call CancelaVisto
            End If
            Inicio = Me.bResultadoVisto
    End Select
End Function

Private Sub CancelaVisto()
    Me.txtNombre.Text = ""
    Me.txtCargo.Text = ""
    bResultadoVisto = False
    fsVistoUserCod = ""
    fnVistoCod = 0
    fsOpeCod = ""
    fsClientePersCod = ""
    fsVistoPersCod = ""
    fsVistoComentario = ""
    fsVistoMovNro = ""
    fnMovNro = 0
    fnCasoCod = 0
    Unload Me
End Sub

Private Sub PasaVisto()

    Me.txtNombre.Text = ""
    Me.txtCargo.Text = ""
    fsVistoUserCod = ""
    fnVistoCod = 0
    fsOpeCod = ""
    fsClientePersCod = ""
    fsVistoPersCod = ""
    fsVistoComentario = ""
    fsVistoMovNro = ""
    fnMovNro = 0
    fnCasoCod = 0
    
    bResultadoVisto = True
    Unload Me
    
End Sub


Public Sub RegistraVistoElectronico(ByVal pnMovNro As Long)

Dim lsSQL As String
Dim loMov As DMov 'COMDMov.DCOMMov
Dim lConn As DConecta 'COMConecta.DCOMConecta

Dim nMovNro As Long
Dim sMovNro As String
Set lConn = New DConecta 'COMConecta.DCOMConecta
Set loMov = New DMov 'COMDMov.DCOMMov
            
    Select Case fnCasoCod
        Case 0
        
        Case 1, 2, 3
            fnMovNro = pnMovNro
            fsVistoMovNro = loMov.GeneraMovNro(gdFecSis, gsCodAge, fsVistoUserCod)
            
            'madm 20101012
            If fsOpeCod = "300503" Then
                 Call loMov.InsertaMov(fsVistoMovNro, fsOpeCod, "Devolucion Creditos Convenio", gMovEstContabNoContable, gMovFlagVigente)
                 fnMovNro = loMov.GetnMovNro(fsVistoMovNro)
            ElseIf fsOpeCod = "910000" Then
                Call loMov.InsertaMov(fsVistoMovNro, fsOpeCod, "Visto Elctronino para Clientes", gMovEstContabNoContable, gMovFlagVigente)
                 fnMovNro = loMov.GetnMovNro(fsVistoMovNro)
            End If
            'madm 20101012
            
            Set loMov = Nothing
            
            fsClientePersCod = IIf((fsClientePersCod = ""), "", fsClientePersCod)
            
            lsSQL = "exec stp_ins_VistoElectronico '" & fsOpeCod & "','" & fsVistoPersCod & "','" & fsVistoComentario & "','" & fsVistoMovNro & "'," & fnMovNro & ",'" & fsClientePersCod & "'"
            
            lConn.AbreConexion
            Call lConn.Ejecutar(lsSQL)
            lConn.CierraConexion
            Set lConn = Nothing
            
        Case 4
            'madm 20100915
            sMovNro = loMov.GeneraMovNro(gdFecSis, gsCodAge, fsVistoUserCod)
            Call loMov.InsertaMov(sMovNro, "910000", "Visto Elctronino para Clientes", gMovEstContabNoContable, gMovFlagVigente)
            nMovNro = loMov.GetnMovNro(sMovNro)
            
            Set loMov = Nothing
            
            lsSQL = "exec stp_ins_VistoElectronico '910000','" & fsVistoPersCod & "' ,'" & fsVistoComentario & "','" & sMovNro & "'," & nMovNro & ", '" & fsClientePersCod & "'"
            
            lConn.AbreConexion
            Call lConn.Ejecutar(lsSQL)
            lConn.CierraConexion
            Set lConn = Nothing
        Case 5
            'GITU 23-08-2011
            sMovNro = loMov.GeneraMovNro(gdFecSis, gsCodAge, fsVistoUserCod)
            Call loMov.InsertaMov(sMovNro, "910000", "Visto Elctronino para Control Operaciones sin Tarjetas", gMovEstContabNoContable, gMovFlagVigente)
            nMovNro = loMov.GetnMovNro(sMovNro)
            
            Set loMov = Nothing
            
            lsSQL = "exec stp_ins_VistoElectronico '910000','" & fsVistoPersCod & "' ,'" & fsVistoComentario & "','" & sMovNro & "'," & nMovNro & ", '" & fsClientePersCod & "'"
            
            lConn.AbreConexion
            Call lConn.Ejecutar(lsSQL)
            lConn.CierraConexion
            Set lConn = Nothing
            'End GITU
            
        Case 6 '** Juez 20120829
            sMovNro = loMov.GeneraMovNro(gdFecSis, gsCodAge, fsVistoUserCod)
            Call loMov.InsertaMov(sMovNro, "910000", "Visto Electronico para el Proceso de Arqueos", gMovEstContabNoContable, gMovFlagVigente)
            nMovNro = loMov.GetnMovNro(sMovNro)
            
            Set loMov = Nothing
            
            lsSQL = "exec stp_ins_VistoElectronico '910000','" & fsVistoPersCod & "' ,'" & fsVistoComentario & "','" & sMovNro & "'," & nMovNro & ", '" & fsClientePersCod & "'"
            
            lConn.AbreConexion
            Call lConn.Ejecutar(lsSQL)
            lConn.CierraConexion
            Set lConn = Nothing
        Case 7 'EJVG20131116
            sMovNro = loMov.GeneraMovNro(gdFecSis, gsCodAge, fsVistoUserCod)
            Call loMov.InsertaMov(sMovNro, "910000", "Visto Electronico para Provisiones de Contabilidad", gMovEstContabNoContable, gMovFlagVigente)
            nMovNro = loMov.GetnMovNro(sMovNro)
            Set loMov = Nothing
            lsSQL = "exec stp_ins_VistoElectronico '910000','" & fsVistoPersCod & "' ,'" & fsVistoComentario & "','" & sMovNro & "'," & nMovNro & ", '" & fsClientePersCod & "'"
            lConn.AbreConexion
            Call lConn.Ejecutar(lsSQL)
            lConn.CierraConexion
            Set lConn = Nothing
    End Select

End Sub

'** Juez 2012809 **********
Public Function ObtieneUsuarioVisto() As String
    ObtieneUsuarioVisto = fsVistoUserCod
End Function
