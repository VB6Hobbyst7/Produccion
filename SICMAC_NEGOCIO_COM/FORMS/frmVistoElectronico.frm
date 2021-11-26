VERSION 5.00
Begin VB.Form frmVistoElectronico 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Visto Electrónico"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5745
   ForeColor       =   &H8000000D&
   Icon            =   "frmVistoElectronico.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   5745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin SICMACT.Usuario ctlUsuario 
      Left            =   105
      Top             =   3570
      _ExtentX        =   820
      _ExtentY        =   820
   End
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
      Height          =   2400
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5475
      Begin VB.TextBox TxtCargo 
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
         MaxLength       =   25
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
      Left            =   4595
      TabIndex        =   8
      Top             =   2600
      Width           =   1000
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Aceptar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3120
      TabIndex        =   7
      Top             =   2600
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
Private fnMovNroOperacion As Long ' *** RIRO SEGUN TI-ERS108-2013 ***
Private fnCasoCod As Integer

Private paraArqueo As Boolean 'GIPO 22-11-2016

'madm 2010909 -----------------------------------------
Public VistoxOpera As Boolean
Public lsCodUserOpera As String
'madm 2010909 -----------------------------------------
'**Valores para fnCasoCod: *****************************
' 00:
' 01: Visto para dudosos  e historial negativo
' 02:
'*******************************************************
Public cMovNroRetencion As String 'JOEP20210910 campana prendario retencion
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

Property Let OperacionCodigo(psOpeCod As String)
   fsOpeCod = psOpeCod
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


Private Sub CmdAceptar_Click()
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
    'GIPO ERS051-2016
    Dim aCOM As COMDPersona.UCOMAcceso
    Dim rs As ADODB.Recordset
    Set aCOM = New COMDPersona.UCOMAcceso
    Set rs = aCOM.obtenerPermisoAccesoArqueo(gsCodUser, txtUsuario.Text)
    
    If Len(Trim(txtUsuario)) > 0 And Len(Trim(TxtClave)) > 0 Then
        If Not ValidacionRFIII Then Exit Sub ' RIRO SEGUN TI-ERS108-2013
            If Not oAcceso.VistoElectronicoEsCorrecto(txtUsuario.Text, TxtClave.Text, sVisPersCod, fnCasoCod, sVisPersNom, sVisPersCar, sMensaje) Then
            MsgBox (sMensaje)
            Me.txtNombre.Text = ""
            Me.TxtCargo.Text = ""
            fsVistoPersCod = ""
            fsVistoUserCod = ""
        Else
            If rs!acceso = "PERMITIDO" Or paraArqueo = False Then 'GIPO ERS051-2016 Validación de cargos. Un Arqueador debe tener un cargo más alto
                MsgBox ("Visto satisfactorio, proceda a registrar ")
                cmdGrabar.Enabled = True
                Me.txtNombre.Text = sVisPersNom
                Me.TxtCargo.Text = sVisPersCar
                fsVistoPersCod = sVisPersCod
                fsVistoUserCod = UCase(txtUsuario.Text)
            Else
                MsgBox "El usuario arqueador debe tener un cargo más elevado para poder efectuar el Arqueo", vbInformation, "Mensaje"
            End If
            
          
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
    
    'ande
    gcUsuarioVistoArqExpAho = txtUsuario.Text
    'end ande
    
    'Select Case fnCasoCod
    '    Case 0
            
    '    Case 1
    'Modify by Gitu
    
    If Not ValidacionRFIII Then Exit Sub ' RIRO SEGUN TI-ERS108-2013
    If fnCasoCod = 0 Then
    
    'ElseIf fnCasoCod = 3 Or fnCasoCod = 4 Then
    'ElseIf fnCasoCod = 3 Or fnCasoCod = 4 Or fnCasoCod = 10 Or fnCasoCod = 16 Or fnCasoCod = 17 Or fnCasoCod = 22 Then 'APRI20171018  ERS028-2017 ADD 17 y 22
    ElseIf fnCasoCod = 3 Or fnCasoCod = 4 Or fnCasoCod = 10 Or fnCasoCod = 16 Or fnCasoCod = 17 Or fnCasoCod = 22 Or fnCasoCod = 23 Then 'JOEP20210910 Campana Prendario rentencion 23
        If Len(Me.txtComentario.Text) > 0 Then
        fsVistoComentario = RTrim(Me.txtComentario.Text)
        'madm 20100915 - operacion vb peps   Or fsOpeCod = "300503"
        'MIOL 20120914, SEGUN RQ12270 SE AGREGO Or fsOpeCod = "901013" Y "901017"************************
        'If fsOpeCod = "910000" Or fsOpeCod = "300503" Or fsOpeCod = "901013" Or fsOpeCod = "901017" Or Mid(fsOpeCod, 1, 5) = "12020" Or fsOpeCod = "401596" Then ''COMMENT BY MARG ERS052-2016
        If fsOpeCod = "910000" Or fsOpeCod = "901013" Or fsOpeCod = "901017" Or Mid(fsOpeCod, 1, 5) = "12020" Or fsOpeCod = "401596" Then 'MARG ERS052-2016 -->SE QUITÓ sOpeCod = "300503" PARA REGISTRAR EL VISTO DESPUES DE OBTENER EL nMovNro de LA OPERACION
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
        If Len(Me.txtNombre.Text) > 0 And Len(Me.TxtCargo.Text) > 0 Then
            bResultadoVisto = True
            fsVistoComentario = UCase(Trim(txtComentario.Text))
            Me.txtNombre.Text = ""
            Me.TxtCargo.Text = ""
            Unload Me
        Else
            bResultadoVisto = False
            fsVistoPersCod = ""
            fsVistoComentario = ""
            Me.txtNombre.Text = ""
            Me.TxtCargo.Text = ""
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
                Me.TxtCargo.Text = sVisPersCar
                
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

Public Function Inicio(ByVal pnCodCaso As Integer, Optional ByVal psOpeCod As String = "", Optional ByVal psClientePersCod As String = "", Optional psCodUser As String = "", Optional pMovNroOperacion As Long = 0) As Boolean
    Dim i As Integer
    Dim loDCred As COMDCredito.DCOMCredito
    Dim lrs As ADODB.Recordset
    paraArqueo = False 'GIPO 22-11-2016
    fnCasoCod = pnCodCaso
    fsClientePersCod = psClientePersCod
    fsOpeCod = psOpeCod
    fnMovNroOperacion = pMovNroOperacion
    
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
                        MsgBox "La persona " & Trim(lrs!cPersNombre) & " está en condición " & Trim(lrs!cEstado) & ", NO se podrá continuar.", vbInformation, "Aviso"
                        Call CancelaVisto
                    Else
                        MsgBox "La persona " & Trim(lrs!cPersNombre) & " está en condición " & Trim(lrs!cEstado) & ", para continuar deberá pasar por el Visto Electrónico.", vbInformation, "Aviso"
                        Me.Show 1
                        If Not Me.bResultadoVisto Then
                            Call CancelaVisto
                        End If
                    End If
                Else '2 judi/casti
                    MsgBox "La persona " & Trim(lrs!cPersNombre) & " tiene créditos en estado " & Trim(lrs!cEstado) & ", para continuar deberá pasar por el Visto Electrónico.", vbInformation, "Aviso"
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
        Case 9 'EJVG20140217 Otros Egresos x Devolución Sobrante de Otras Operaciones con Cheque
            Me.Show 1
            If Not Me.bResultadoVisto Then
                Call CancelaVisto
            End If
            Inicio = Me.bResultadoVisto
        Case 10 'RECO20140408 ERS160-2013 Superar el umbral de monto establecido para operaciones pignoraticios en un determinado periodo
            Me.Show 1
            If Not Me.bResultadoVisto Then
                Call CancelaVisto
            End If
            Inicio = Me.bResultadoVisto
        Case 11 'FRHU20140505 ERS063-2014 Otros Egresos en Efectivo
            Me.Show 1
            If Not Me.bResultadoVisto Then
                Call CancelaVisto
            End If
            Inicio = Me.bResultadoVisto
        Case 12 'FRHU20140701 ERS048-2014 Visto electronico Notas de cargo y abono
            Me.Show 1
            If Not Me.bResultadoVisto Then
                Call CancelaVisto
            End If
            Inicio = Me.bResultadoVisto
        Case 13 'RIRO 20150605 ERS162-2014 Visto Electrónico Para Pagos de Utilidades
            Me.Show 1
            If Not Me.bResultadoVisto Then
                Call CancelaVisto
            End If
            Inicio = Me.bResultadoVisto
        Case 14 'FRHU 20151130 ERS-077-2015 Visto Electrónico Para Actualización y Autorización de Datos
            Me.Show 1
            If Not Me.bResultadoVisto Then
                Call CancelaVisto
            End If
            Inicio = Me.bResultadoVisto
        Case 15 'PASI20151221 ers0692015 Visto Electrónico Para Arqueo de Tarjetas de Debito
            paraArqueo = True 'GIPO 22-11-2016
            Me.Show 1
            If Not Me.bResultadoVisto Then
                Call CancelaVisto
            End If
            Inicio = Me.bResultadoVisto
        Case 16 'MARG ERS052-2017
            paraArqueo = True
            Me.Show 1
            If Not Me.bResultadoVisto Then
                Call CancelaVisto
            End If
            Inicio = Me.bResultadoVisto
            
        Case 17, 22 'RECO20160714 Visto para Anulacion de Seguro Sepelio
                    'APRI20171027 ERS028-2017 ADD 22 ANULACION SEG. TARJETA
            Me.Show 1
            If Not Me.bResultadoVisto Then
                Call CancelaVisto
            End If
            Inicio = Me.bResultadoVisto
        Case 18 'GIPO ERS070-2016
            Me.Show 1
            If Not Me.bResultadoVisto Then
                 Call CancelaVisto
            End If
            Inicio = Me.bResultadoVisto
        Case 19, 20, 21 'ANDE ERS021-2017 = 19: Arqueo de expedientes de ahorro, 20: extorno de Arqueo de expedientes de ahorro
                        'RIRO: Extorno de aprobación de créditos = 21
            Me.Show 1
            If Not Me.bResultadoVisto Then
                 Call CancelaVisto
            End If
            Inicio = Me.bResultadoVisto
        'JOEP20210909 campaña prendario
        Case 23
            Me.Show 1
            If Not Me.bResultadoVisto Then
                 Call CancelaVisto
            End If
            Inicio = Me.bResultadoVisto
    'JOEP20210909 campaña prendario
    End Select
End Function

Private Sub CancelaVisto()
    Me.txtNombre.Text = ""
    Me.TxtCargo.Text = ""
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
    Me.TxtCargo.Text = ""
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


'Public Sub RegistraVistoElectronico(ByVal pnMovNro As Long) 'FRHU 20140505 ERS063-2014
'Public Sub RegistraVistoElectronico(ByVal pnMovNro As Long, Optional ByRef pnMovVisto As Long = 0) 'FRHU 20140505 ERS063-2014 'COMMENT BY MARG ERS065-2017
Public Sub RegistraVistoElectronico(ByVal pnMovNro As Long, Optional ByRef pnMovVisto As Long = 0, Optional ByVal pcUserSupervisado As String = "", Optional ByVal pnMovNroOperacion As Long = 0) 'FRHU 20140505 ERS063-2014 'MARG ERS065-2017

Dim lsSQL As String
Dim loMov As COMDMov.DCOMMov
Dim lConn As COMConecta.DCOMConecta

Dim nMovNro As Long
Dim sMovNro As String
Set lConn = New COMConecta.DCOMConecta
Set loMov = New COMDMov.DCOMMov
            
'MARG ERS052-2017 -------------------
Dim cUserSupervisado As String
Dim nMovNroOperacion As Long

If pcUserSupervisado = "" Then
    cUserSupervisado = gsCodUser
Else
    cUserSupervisado = pcUserSupervisado
End If

nMovNroOperacion = pnMovNroOperacion
'END MARG -----------------------------

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
                Call loMov.InsertaMov(fsVistoMovNro, fsOpeCod, "Visto Elctrónico para Clientes", gMovEstContabNoContable, gMovFlagVigente)
                 fnMovNro = loMov.GetnMovNro(fsVistoMovNro)
            End If
            'madm 20101012
            
            Set loMov = Nothing
            
            fsClientePersCod = IIf((fsClientePersCod = ""), "", fsClientePersCod)
            
            'lsSQL = "exec stp_ins_VistoElectronico '" & fsOpeCod & "','" & fsVistoPersCod & "','" & fsVistoComentario & "','" & fsVistoMovNro & "'," & fnMovNro & ",'" & fsClientePersCod & "'" 'COMMENT BY MARG ERS065-2017
            lsSQL = "exec stp_ins_VistoElectronico '" & fsOpeCod & "','" & fsVistoPersCod & "','" & fsVistoComentario & "','" & fsVistoMovNro & "'," & fnMovNro & ",'" & fsClientePersCod & "','" & cUserSupervisado & "'," & nMovNroOperacion  'ADD BY MARG ERS065-2017
            
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
            
            'lsSQL = "exec stp_ins_VistoElectronico '910000','" & fsVistoPersCod & "' ,'" & fsVistoComentario & "','" & sMovNro & "'," & nMovNro & ", '" & fsClientePersCod & "'" 'COMMENT BY MARG ERS052-2017
            lsSQL = "exec stp_ins_VistoElectronico '910000','" & fsVistoPersCod & "' ,'" & fsVistoComentario & "','" & sMovNro & "'," & nMovNro & ", '" & fsClientePersCod & "','" & cUserSupervisado & "'," & nMovNroOperacion  'ADD BY MARG ERS052-2017
            
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
            
            'lsSQL = "exec stp_ins_VistoElectronico '910000','" & fsVistoPersCod & "' ,'" & fsVistoComentario & "','" & sMovNro & "'," & nMovNro & ", '" & fsClientePersCod & "'" 'COMMENT BY MARG ERS065-2017
            lsSQL = "exec stp_ins_VistoElectronico '910000','" & fsVistoPersCod & "' ,'" & fsVistoComentario & "','" & sMovNro & "'," & nMovNro & ", '" & fsClientePersCod & "','" & cUserSupervisado & "'," & nMovNroOperacion 'ADD BY MARG ERS065-2017
            
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
            
            'lsSQL = "exec stp_ins_VistoElectronico '910000','" & fsVistoPersCod & "' ,'" & fsVistoComentario & "','" & sMovNro & "'," & nMovNro & ", '" & fsClientePersCod & "'" 'COMMENT BY MARG ERS052-2017
            lsSQL = "exec stp_ins_VistoElectronico '910000','" & fsVistoPersCod & "' ,'" & fsVistoComentario & "','" & sMovNro & "'," & nMovNro & ", '" & fsClientePersCod & "','" & cUserSupervisado & "'," & nMovNroOperacion  'ADD BY MARG ERS052-2017
            
            lConn.AbreConexion
            Call lConn.Ejecutar(lsSQL)
            lConn.CierraConexion
            Set lConn = Nothing
'        Case 2
'            Set lConn = New COMConecta.DCOMConecta
'
'            fnMovNro = pnMovNro
'
'            Set loMov = New COMDMov.DCOMMov
'            fsVistoMovNro = loMov.GeneraMovNro(gdFecSis, gsCodAge, fsVistoUserCod)
'            Set loMov = Nothing
'
'            lsSQL = "exec stp_ins_VistoElectronico '" & fsOpeCod & "','" & fsVistoPersCod & "','" & fsVistoComentario & "','" & fsVistoMovNro & "'," & fnMovNro
'
'            lConn.AbreConexion
'            Call lConn.Ejecutar(lsSQL)
'            lConn.CierraConexion
'            Set lConn = Nothing
'
'        Case 3 '*** PEAC 20081001
'
'            Set lConn = New COMConecta.DCOMConecta
'
'            fnMovNro = pnMovNro
'
'            Set loMov = New COMDMov.DCOMMov
'            fsVistoMovNro = loMov.GeneraMovNro(gdFecSis, gsCodAge, fsVistoUserCod)
'            Set loMov = Nothing
'
'            lsSQL = "exec stp_ins_VistoElectronico '" & fsOpeCod & "','" & fsVistoPersCod & "','" & fsVistoComentario & "','" & fsVistoMovNro & "'," & fnMovNro
'
'            lConn.AbreConexion
'            Call lConn.Ejecutar(lsSQL)
'            lConn.CierraConexion
'            Set lConn = Nothing
        Case 9 '** EJVG20140219
            sMovNro = loMov.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), fsVistoUserCod)
            Call loMov.InsertaMov(sMovNro, "910000", "Visto Electronico para devolución de sobrantes otras operaciones con cheque", gMovEstContabNoContable, gMovFlagVigente)
            Set loMov = Nothing
            'lsSQL = "exec stp_ins_VistoElectronico '910000','" & fsVistoPersCod & "' ,'" & fsVistoComentario & "','" & sMovNro & "'," & pnMovNro & ", '" & fsClientePersCod & "'" ''COMMENT BY MARG ERS052-2017
            lsSQL = "exec stp_ins_VistoElectronico '910000','" & fsVistoPersCod & "' ,'" & fsVistoComentario & "','" & sMovNro & "'," & pnMovNro & ", '" & fsClientePersCod & "','" & cUserSupervisado & "'," & nMovNroOperacion  'ADD BY MARG ERS052-2017
            lConn.AbreConexion
            Call lConn.Ejecutar(lsSQL)
            lConn.CierraConexion
            Set lConn = Nothing
        Case 10 'RECO20140408 ERS160-2013
            sMovNro = loMov.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), fsVistoUserCod)
            Call loMov.InsertaMov(sMovNro, "910000", "Visto Electronico para Operaciones que superen el umbral de monto establecido pignoraticio en un periodo determinado", gMovEstContabNoContable, gMovFlagVigente)
            Set loMov = Nothing
            'lsSQL = "exec stp_ins_VistoElectronico '" & fsOpeCod & "','" & fsVistoPersCod & "' ,'" & fsVistoComentario & "','" & VistoMovNro & "'," & pnMovNro & ", '" & fsClientePersCod & "'" 'COMMENT BY MARG ERS052-2017
            lsSQL = "exec stp_ins_VistoElectronico '" & fsOpeCod & "','" & fsVistoPersCod & "' ,'" & fsVistoComentario & "','" & VistoMovNro & "'," & pnMovNro & ", '" & fsClientePersCod & "','" & cUserSupervisado & "'," & nMovNroOperacion  'ADD BY MARG ERS052-2017
            lConn.AbreConexion
            Call lConn.Ejecutar(lsSQL)
            lConn.CierraConexion
            Set lConn = Nothing
        Case 11 'FRHU20140505 ERS063-2014
            sMovNro = loMov.GeneraMovNro(gdFecSis, gsCodAge, fsVistoUserCod)
            Call loMov.InsertaMov(sMovNro, "910000", "Visto Electronico para Operaciones de Otros Egresos en Efectivo", gMovEstContabNoContable, gMovFlagVigente)
            nMovNro = loMov.GetnMovNro(sMovNro)
            
            Set loMov = Nothing
            
            'lsSQL = "exec stp_ins_VistoElectronico '910000','" & fsVistoPersCod & "' ,'" & fsVistoComentario & "','" & sMovNro & "'," & nMovNro & ", '" & fsClientePersCod & "'" 'COMMENT BY MARG ERS052-2017
            lsSQL = "exec stp_ins_VistoElectronico '910000','" & fsVistoPersCod & "' ,'" & fsVistoComentario & "','" & sMovNro & "'," & nMovNro & ", '" & fsClientePersCod & "','" & cUserSupervisado & "'," & nMovNroOperacion  'ADD BY MARG ERS052-2017
            
            lConn.AbreConexion
            Call lConn.Ejecutar(lsSQL)
            pnMovVisto = nMovNro
            lConn.CierraConexion
            Set lConn = Nothing
        Case 12 'FRHU20140701 ERS048-2014
            sMovNro = loMov.GeneraMovNro(gdFecSis, gsCodAge, fsVistoUserCod)
            Call loMov.InsertaMov(sMovNro, "910000", "Visto Electronico para Operaciones de Notas de Cargo y Abono", gMovEstContabNoContable, gMovFlagVigente)
            nMovNro = loMov.GetnMovNro(sMovNro)
            
            Set loMov = Nothing
            
            'lsSQL = "exec stp_ins_VistoElectronico '910000','" & fsVistoPersCod & "' ,'" & fsVistoComentario & "','" & sMovNro & "'," & nMovNro & ", '" & fsClientePersCod & "'" 'COMMENT BY MARG ERS052-2017
            lsSQL = "exec stp_ins_VistoElectronico '910000','" & fsVistoPersCod & "' ,'" & fsVistoComentario & "','" & sMovNro & "'," & nMovNro & ", '" & fsClientePersCod & "','" & cUserSupervisado & "'," & nMovNroOperacion  'ADD BY MARG ERS052-2017
            
            lConn.AbreConexion
            Call lConn.Ejecutar(lsSQL)
            pnMovVisto = nMovNro
            lConn.CierraConexion
            Set lConn = Nothing
            
        Case 13 '** RIRO 20150605 ERS162-2014
            sMovNro = loMov.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), fsVistoUserCod)
            Call loMov.InsertaMov(sMovNro, "910000", "Visto Electronico para Pago de Utilidades y Otros", gMovEstContabNoContable, gMovFlagVigente)
            Set loMov = Nothing
            'lsSQL = "exec stp_ins_VistoElectronico '910000','" & fsVistoPersCod & "' ,'" & fsVistoComentario & "','" & sMovNro & "'," & pnMovNro & ", '" & fsClientePersCod & "'" 'COMMENT BY MARG ERS052-2017
            lsSQL = "exec stp_ins_VistoElectronico '910000','" & fsVistoPersCod & "' ,'" & fsVistoComentario & "','" & sMovNro & "'," & pnMovNro & ", '" & fsClientePersCod & "','" & cUserSupervisado & "'," & nMovNroOperacion  'ADD BY MARG ERS052-2017
            lConn.AbreConexion
            Call lConn.Ejecutar(lsSQL)
            lConn.CierraConexion
            Set lConn = Nothing
        Case 14 '** FRHU 20151130 ERS077-2015
            sMovNro = loMov.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), fsVistoUserCod)
            Call loMov.InsertaMov(sMovNro, "910000", "Visto Electronico para Actualización y Autorizacón de Datos", gMovEstContabNoContable, gMovFlagVigente)
            Set loMov = Nothing
            'lsSQL = "exec stp_ins_VistoElectronico '910000','" & fsVistoPersCod & "' ,'" & fsVistoComentario & "','" & sMovNro & "'," & pnMovNro & ", '" & fsClientePersCod & "'" 'COMMENT BY MARG ERS052-2017
            lsSQL = "exec stp_ins_VistoElectronico '910000','" & fsVistoPersCod & "' ,'" & fsVistoComentario & "','" & sMovNro & "'," & pnMovNro & ", '" & fsClientePersCod & "','" & cUserSupervisado & "'," & nMovNroOperacion  'ADD BY MARG ERS052-2017
            lConn.AbreConexion
            Call lConn.Ejecutar(lsSQL)
            lConn.CierraConexion
            Set lConn = Nothing
        'MARG ERS052-2017 ---------------------------
        Case 15
            sMovNro = loMov.GeneraMovNro(gdFecSis, gsCodAge, fsVistoUserCod)
            Call loMov.InsertaMov(sMovNro, "910000", "Visto Electronico para Arqueo de Stock de Tarjetas de Débito - Bóveda", gMovEstContabNoContable, gMovFlagVigente)
            nMovNro = loMov.GetnMovNro(sMovNro)
            
            Set loMov = Nothing
            lsSQL = "exec stp_ins_VistoElectronico '900800','" & fsVistoPersCod & "' ,'" & fsVistoComentario & "','" & sMovNro & "'," & nMovNro & ", '" & fsClientePersCod & "','" & cUserSupervisado & "'," & nMovNroOperacion  'ADD BY MARG ERS052-2017
            
            lConn.AbreConexion
            Call lConn.Ejecutar(lsSQL)
            lConn.CierraConexion
            Set lConn = Nothing
        Case 16
            sMovNro = loMov.GeneraMovNro(gdFecSis, gsCodAge, fsVistoUserCod)
            Call loMov.InsertaMov(sMovNro, "910000", "Visto Electronico para Arqueo de Stock de Tarjetas de Débito - Ventanilla", gMovEstContabNoContable, gMovFlagVigente)
            nMovNro = loMov.GetnMovNro(sMovNro)
            
            Set loMov = Nothing
            lsSQL = "exec stp_ins_VistoElectronico '900900','" & fsVistoPersCod & "' ,'" & fsVistoComentario & "','" & sMovNro & "'," & nMovNro & ", '" & fsClientePersCod & "','" & cUserSupervisado & "'," & nMovNroOperacion  'ADD BY MARG ERS052-2017
            
            lConn.AbreConexion
            Call lConn.Ejecutar(lsSQL)
            lConn.CierraConexion
            Set lConn = Nothing
        'END MARG -------------------------------------
        'Case 16 'RECO20160714
        Case 17, 22 'EJVG20161111
            sMovNro = loMov.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), fsVistoUserCod)
            'APRI20171027 ERS028-2017
            Dim cGlosa As String
            If fnCasoCod = 17 Then
                cGlosa = "Visto Electrónico por Anulación Seguro Sepelio"
            ElseIf fnCasoCod = 22 Then
                cGlosa = "Visto Electrónico por Anulación Seguro Tarjeta"
            End If
            'END APRI
            Call loMov.InsertaMov(sMovNro, "910000", cGlosa, gMovEstContabNoContable, gMovFlagVigente) 'APRI20171027 ADD cGlosa
            Set loMov = Nothing
            'lsSQL = "exec stp_ins_VistoElectronico '910000','" & fsVistoPersCod & "' ,'" & fsVistoComentario & "','" & sMovNro & "'," & pnMovNro & ", '" & fsClientePersCod & "'" 'COMMENT BY MARG ERS052-2017
            lsSQL = "exec stp_ins_VistoElectronico '910000','" & fsVistoPersCod & "' ,'" & fsVistoComentario & "','" & sMovNro & "'," & pnMovNro & ", '" & fsClientePersCod & "','" & cUserSupervisado & "'," & nMovNroOperacion  'ADD BY MARG ERS052-2017
            lConn.AbreConexion
            Call lConn.Ejecutar(lsSQL)
            lConn.CierraConexion
            Set lConn = Nothing
        'MARG ERS052-2017 ---------------------------
        Case 18
            sMovNro = loMov.GeneraMovNro(gdFecSis, gsCodAge, fsVistoUserCod)
            Call loMov.InsertaMov(sMovNro, "910000", "Visto Electronico para la Devolucion de Sobrantes de Adjudicados", gMovEstContabNoContable, gMovFlagVigente)
            nMovNro = loMov.GetnMovNro(sMovNro)
            
            Set loMov = Nothing
            lsSQL = "exec stp_ins_VistoElectronico '910000','" & fsVistoPersCod & "' ,'" & fsVistoComentario & "','" & sMovNro & "'," & nMovNro & ", '" & fsClientePersCod & "','" & cUserSupervisado & "'," & nMovNroOperacion  'ADD BY MARG ERS052-2017
            
            lConn.AbreConexion
            Call lConn.Ejecutar(lsSQL)
            lConn.CierraConexion
            Set lConn = Nothing
        'END MARG -------------------------------------
        Case 19, 20, 21 'ANDE 20170530 Visto para arqueo y extorno de expedientes de arqueo
                        'RIRO 20171006 Visto electrónico para extorno de aprobación de crédito
            Dim cDes As String
            If fnCasoCod = 19 Then
                cDes = "Visto Electrónico por Arqueo expediente de ahorro"
            ElseIf fnCasoCod = 20 Then
                cDes = "Visto Electrónico por Extorno de Arqueo expediente de ahorro"
            ElseIf fnCasoCod = 21 Then
                cDes = "Visto Electrónico por Extorno de Aprobación de crédito"
            End If
            sMovNro = loMov.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), fsVistoUserCod)
            Call loMov.InsertaMov(sMovNro, "910000", cDes, gMovEstContabNoContable, gMovFlagVigente)
            Set loMov = Nothing
            'lsSQL = "exec stp_ins_VistoElectronico '910000','" & fsVistoPersCod & "' ,'" & fsVistoComentario & "','" & sMovNro & "'," & pnMovNro & ", '" & fsClientePersCod & "'" 'COMMENT BY MARG ERS052-2017
            lsSQL = "exec stp_ins_VistoElectronico '910000','" & fsVistoPersCod & "' ,'" & fsVistoComentario & "','" & sMovNro & "'," & pnMovNro & ", '" & fsClientePersCod & "','" & cUserSupervisado & "'," & nMovNroOperacion  'ADD BY MARG ERS052-2017
            lConn.AbreConexion
            Call lConn.Ejecutar(lsSQL)
            lConn.CierraConexion
            Set lConn = Nothing
    'joep20210927 campana prendario
        Case 23
            sMovNro = loMov.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), fsVistoUserCod)
            Set loMov = Nothing
            lsSQL = "exec stp_ins_VistoElectronico '910000','" & fsVistoPersCod & "' ,'" & fsVistoComentario & "','" & sMovNro & "'," & pnMovNro & ", '" & fsClientePersCod & "','" & cUserSupervisado & "'," & nMovNroOperacion
            lConn.AbreConexion
            Call lConn.Ejecutar(lsSQL)
            lConn.CierraConexion
            Set lConn = Nothing
            cMovNroRetencion = sMovNro
    'joep20210927 campana prendario
    End Select

End Sub

'** Juez 2012809 **********
Public Function ObtieneUsuarioVisto() As String
    ObtieneUsuarioVisto = fsVistoUserCod
End Function

' *** RIRO SEGUN TI-ERS108-2013 ***
Public Function ValidacionRFIII() As Boolean
    Dim rsRF3 As New ADODB.Recordset
    Dim oDCOMMov As COMDMov.DCOMMov
    Dim sMensaje As String
    Dim sCodCargo As String
    Dim sMovNro As String
   
        
    If fnMovNroOperacion > 0 Then
        Set oDCOMMov = New COMDMov.DCOMMov
        sMovNro = oDCOMMov.GetcMovNro(fnMovNroOperacion)
        Set oDCOMMov = Nothing
    End If
    
    ctlUsuario.Inicio (Trim(txtUsuario.Text))
    sCodCargo = ctlUsuario.PersCargoCod
    Set rsRF3 = ValidarRFIII
    sMensaje = ""
    
    '*** VERIFICANDO SI RFFIII ESTA EN MODO SUPERVISOR
        Dim oAcceso As UCOMAcceso
        Dim clsGen As COMDConstSistema.DCOMGeneral
        Dim sGrupos, sTemporal, sGrupoRF3 As String
        Dim bModoSupervisor As Boolean
        
        Set oAcceso = New UCOMAcceso
        Set clsGen = New COMDConstSistema.DCOMGeneral
        sGrupoRF3 = clsGen.GetConstante(10027, , "100", "1")!cDescripcion
        sTemporal = ""
        sGrupos = ""
        
        If oAcceso.VerificarUsuarioExistaEnRRHH(Trim(txtUsuario.Text)) Then
            Call oAcceso.CargaGruposUsuario(Trim(txtUsuario.Text), gsDominio)
            sTemporal = oAcceso.DameGrupoUsuario
            Do While Len(sTemporal) > 0
                sGrupos = sGrupos & sTemporal & ","
                sTemporal = oAcceso.DameGrupoUsuario
            Loop
        End If

        sGrupos = Mid(sGrupos, 1, Len(sGrupos) - IIf(Len(sGrupos) > 0, 1, 0))
        Set oAcceso = Nothing
        Set clsGen = Nothing
        If InStr(1, sGrupos, sGrupoRF3) > 0 Then
        
            Dim rsRF3User As New ADODB.Recordset
            Dim oPersona As New COMDPersona.DCOMPersonas
            Set rsRF3User = oPersona.RecuperarGruposRF3(Trim(txtUsuario.Text))
            
            If Not rsRF3User Is Nothing Then
                If Not rsRF3User.BOF And Not rsRF3User.EOF Then
                    If rsRF3User!nEstado = 1 Then
                        bModoSupervisor = True
                    Else
                        bModoSupervisor = False
                    End If
                Else
                    bModoSupervisor = False
                End If
            Else
                bModoSupervisor = False
            End If
        
        Else
            bModoSupervisor = False
        End If
    '*** FIN VERIFICACION
    
    If Not (rsRF3.EOF Or rsRF3.BOF) And rsRF3.RecordCount > 0 Then
        If sCodCargo = "006005" Then     ' *** SI ES "SUPERVISOR"
            If Not rsRF3!bOpcionesSimultaneas And rsRF3!bModoSupervisor Then
                sMensaje = "No es posible emitir el VB virtual porque el RFIII se encuentra activo en modo supervisor si desea dar VB usted debe " & vbNewLine & _
                "desactivar al RFIII en este perfil, de lo contrario solicite a él que emita el VB correspondiente"
            End If
        ElseIf sCodCargo = "007026" Then ' *** SI ES "RFIII"
            If Not bModoSupervisor Then
                sMensaje = "Actualmente no cuenta con permisos para dar VB por lo que el VB virtual debe darlo el supervisor " & _
                           "de operaciones."
            End If
        End If
    End If
    If Len(sMovNro) > 0 Then
        If UCase(Right(sMovNro, 4)) = UCase(Trim(txtUsuario.Text)) Then
            sMensaje = "No puede dar su VB a operaciones realizadas por usted mismo"
        End If
    End If
    If Len(sMensaje) > 0 Then
        MsgBox sMensaje, vbExclamation, "Aviso"
         ValidacionRFIII = False
        Exit Function
    End If
    ValidacionRFIII = True
End Function
' *** FIN RIRO ***
'MARG ERS052-2017-----------------------------------
Public Sub RegistrarVistoElectronicoLavDinero(ByVal pcOpeCod As String, ByVal pcVistoPersCod As String, ByVal pcVistoComentario As String, ByVal pcClientePersCod As String, ByVal pcUserSupervisado As String, ByVal pnMovNroOperacion As Long)
            Dim sMovNro As String
            Dim nMovNro As Long
            Dim lsSQL As String
            Dim loMov As COMDMov.DCOMMov
            Dim lConn As COMConecta.DCOMConecta
            Set loMov = New COMDMov.DCOMMov
            Set lConn = New COMConecta.DCOMConecta
            
            sMovNro = loMov.GeneraMovNro(gdFecSis, gsCodAge, fsVistoUserCod)
            Call loMov.InsertaMov(sMovNro, pcOpeCod, "Visto Electronico para el Registro de Operaciones(REU)", gMovEstContabNoContable, gMovFlagVigente)
            nMovNro = loMov.GetnMovNro(sMovNro)
            
            Set loMov = Nothing
            lsSQL = "exec stp_ins_VistoElectronico '" & pcOpeCod & "','" & pcVistoPersCod & "' ,'" & pcVistoComentario & "','" & sMovNro & "'," & nMovNro & ", '" & pcClientePersCod & "','" & pcUserSupervisado & "'," & pnMovNroOperacion
            
            lConn.AbreConexion
            Call lConn.Ejecutar(lsSQL)
            lConn.CierraConexion
            Set lConn = Nothing
End Sub
'END MARG ---------------------------------------------
