VERSION 5.00
Begin VB.Form frmCapVBCancelaCTS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CONFIRMACION DE CESE"
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5880
   Icon            =   "frmCapVBCancelaCTS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   5880
   StartUpPosition =   2  'CenterScreen
   Begin SICMACT.Usuario ctlUsuario 
      Left            =   5280
      Top             =   2880
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox txtPersContac 
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   240
      Width           =   3495
   End
   Begin VB.CommandButton cmdValidar 
      Caption         =   "Validar"
      Height          =   375
      Left            =   4440
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox txtClave 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2160
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox txtUsuario 
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   720
      Width           =   1815
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   5640
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Persona de Contacto:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   12
      Top             =   330
      Width           =   1860
   End
   Begin VB.Label lblCargo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2160
      TabIndex        =   11
      Top             =   2400
      Width           =   3495
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Cargo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   10
      Top             =   2490
      Width           =   570
   End
   Begin VB.Label lblNombre 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2160
      TabIndex        =   9
      Top             =   1980
      Width           =   3495
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Nombre:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   2070
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Clave"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   1290
      Width           =   495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Usuario: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   810
      Width           =   780
   End
End
Attribute VB_Name = "frmCapVBCancelaCTS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sVisPersCod As String
Dim sVisPersNom As String
Dim sVisPersCar As String

Dim sMensaje As String 'Add Gitu and Peac 22-08-08

Private fsVistoUserCod As String
Private fnVistoCod As Long
Private fsVistoPersCod As String
Private fsVistoComentario As String
Private fsVistoMovNro As String
Private fnCasoCod As Integer

Private lsCtaCod As String
Private lsCodAge As String
Private lnMonCanc As Double
Private lsCodUsu As String
Private lsPersCod As String

Public fbValidaVisto As Boolean

Private Sub cmdAceptar_Click()
Dim ldFechaHora As String

    'ldFechaHora = Format(Now(), "yyyy-mm-dd HH:MM:SS")
    ldFechaHora = Right(gdFecSis, 4) & "-" & Mid(gdFecSis, 4, 2) & "-" & Left(gdFecSis, 2) & " " & Format(Time, "hh:mm:ss")

    If Len(txtPersContac) > 2 Then
    
        Call RegistraVistoCancelacion(lsCtaCod, lsCodAge, ldFechaHora, lnMonCanc, lsCodUsu, txtUsuario.Text, txtPersContac.Text, lsPersCod)
        fbValidaVisto = True
        Unload Me
    Else
        MsgBox "El nombre del CONTACTO debe tener mas de 2 caracteres", vbInformation + vbOKOnly, "Aviso"
        txtPersContac.SetFocus
        Exit Sub
    End If
End Sub

Private Sub cmdCancelar_Click()
    fbValidaVisto = False
    Unload Me
End Sub

Private Sub cmdValidar_Click()
    Dim oAcceso As comdpersona.UCOMAcceso
    Set oAcceso = New comdpersona.UCOMAcceso
    
    fnCasoCod = 16

    If Len(Trim(txtUsuario)) > 0 And Len(Trim(txtClave)) > 0 Then
        If gsCodUser <> txtUsuario.Text Then
            If Not ValidacionRFIII Then Exit Sub ' RIRO SEGUN TI-ERS108-2013
            If Not oAcceso.VistoElectronicoEsCorrecto(txtUsuario.Text, txtClave.Text, sVisPersCod, fnCasoCod, sVisPersNom, sVisPersCar, sMensaje) Then
                MsgBox (sMensaje)
                Me.lblNombre.Caption = ""
                Me.lblCargo.Caption = ""
                fsVistoPersCod = ""
                fsVistoUserCod = ""
            Else
                MsgBox ("Visto satisfactorio, proceda a registrar")
                cmdAceptar.Enabled = True
                cmdValidar.Enabled = False
                txtUsuario.Enabled = False
                txtClave.Enabled = False
                
                Me.lblNombre.Caption = sVisPersNom
                Me.lblCargo.Caption = sVisPersCar
            End If
        Else
            MsgBox "El Usuario que realiza la Operación NO puede validar el Visto Bueno, comuníquese con el Supervisor", vbInformation, "Aviso"
            txtUsuario.SetFocus
            Exit Sub
        End If
    End If
End Sub

Private Sub txtPersContac_KeyPress(KeyAscii As Integer)
    Call SoloLetras(KeyAscii, True)
    KeyAscii = fgIntfMayusculas(KeyAscii)
End Sub
Private Sub txtUsuario_KeyPress(KeyAscii As Integer)
    KeyAscii = fgIntfMayusculas(KeyAscii)
End Sub
Public Function Inicia(ByVal psCtaCod As String, ByVal psCodAge As String, ByVal pnMonCanc As Double, ByVal psCodUsu As String, ByVal psPersCod As String) As Boolean

    lsCtaCod = psCtaCod
    lsCodAge = psCodAge
    lnMonCanc = pnMonCanc
    lsCodUsu = psCodUsu
    lsPersCod = psPersCod
    
    Me.Show 1
    
    If Not Me.fbValidaVisto Then
        Call CancelaVisto
    End If
    Inicia = Me.fbValidaVisto
    
End Function
Private Sub CancelaVisto()
    txtPersContac.Text = ""
    txtUsuario.Text = ""
    txtClave.Text = ""
    
    lblNombre.Caption = ""
    lblNombre.Caption = ""
    
    cmdAceptar.Enabled = False
    
    Unload Me
End Sub

Private Sub RegistraVistoCancelacion(ByVal psCtaCod As String, ByVal psCodAge As String, ByVal pdFecHorOpe As String, ByVal pnMontoCanc As Double, ByVal psCodUsu As String, _
                                     ByVal psCodUsuAut As String, ByVal psPersContac As String, ByVal pcPersCod As String)
Dim lsSQL As String
Dim lConn As COMConecta.DCOMConecta
    
    Set lConn = New COMConecta.DCOMConecta
    
    lsSQL = "exec stp_ins_RegistaVistoBuenoCancelacionCTS '" & psCtaCod & "','" & psCodAge & "' ,'" & pdFecHorOpe & "'," & pnMontoCanc & ",'" & psCodUsu & "', '" & psCodUsuAut & "', '" & psPersContac & "', '" & pcPersCod & "'"
    lConn.AbreConexion
    Call lConn.Ejecutar(lsSQL)
    lConn.CierraConexion
    Set lConn = Nothing

End Sub
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
        sGrupoRF3 = clsGen.GetConstante(10027, , "100", "1")!cdescripcion
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
            Dim oPersona As New comdpersona.DCOMPersonas
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

