VERSION 5.00
Begin VB.Form frmCajaChicaCambioEncargado 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Caja Chica: Cambio de encargado"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8535
   Icon            =   "frmCajaChicaCambioEncargado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   8535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraSaldo 
      Caption         =   "Saldo Actual Caja Chica"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   855
      Left            =   120
      TabIndex        =   19
      Top             =   3960
      Width           =   8355
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   375
         Left            =   6960
         TabIndex        =   24
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdCambiar 
         Caption         =   "Cambiar"
         Height          =   375
         Left            =   5520
         TabIndex        =   23
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblSaldo 
         AutoSize        =   -1  'True
         Caption         =   "S/. :"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   315
      End
      Begin VB.Label lblSaldoActual 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   330
         Left            =   600
         TabIndex        =   20
         Top             =   360
         Width           =   1380
      End
   End
   Begin VB.Frame fraNuevoEncargado 
      Caption         =   "Datos: Nuevo Encargado"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1425
      Left            =   120
      TabIndex        =   13
      Top             =   2520
      Width           =   8355
      Begin Sicmact.TxtBuscar TxtBuscarPersCod 
         Height          =   360
         Left            =   1200
         TabIndex        =   22
         Top             =   360
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   423
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TipoBusqueda    =   7
         TipoBusPers     =   1
         EnabledText     =   0   'False
      End
      Begin VB.Label lblNomNuevoEncargado 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   330
         Left            =   1200
         TabIndex        =   18
         Top             =   840
         Width           =   4545
      End
      Begin VB.Label lblCodNuevo 
         AutoSize        =   -1  'True
         Caption         =   "Código :"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   480
         Width           =   585
      End
      Begin VB.Label lblNomNuevo 
         AutoSize        =   -1  'True
         Caption         =   "Nombre :"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   840
         Width           =   645
      End
      Begin VB.Label lblDniNUevo 
         Caption         =   "DNI :"
         Height          =   255
         Left            =   6120
         TabIndex        =   15
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lblNroDniNuevo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   330
         Left            =   6720
         TabIndex        =   14
         Top             =   840
         Width           =   1380
      End
   End
   Begin VB.Frame fraEncActual 
      Caption         =   "Datos: Actual Encargado"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1545
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   8355
      Begin Sicmact.Usuario Usu 
         Left            =   7800
         Top             =   240
         _ExtentX        =   820
         _ExtentY        =   820
      End
      Begin VB.Label lblNroDni 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   330
         Left            =   6720
         TabIndex        =   12
         Top             =   960
         Width           =   1380
      End
      Begin VB.Label lblDni 
         Caption         =   "DNI :"
         Height          =   255
         Left            =   6120
         TabIndex        =   11
         Top             =   960
         Width           =   495
      End
      Begin VB.Label lblNombre 
         AutoSize        =   -1  'True
         Caption         =   "Nombre :"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   645
      End
      Begin VB.Label lblCodigo 
         AutoSize        =   -1  'True
         Caption         =   "Código :"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   585
      End
      Begin VB.Label lblPerscod 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   330
         Left            =   1200
         TabIndex        =   8
         Top             =   480
         Width           =   1380
      End
      Begin VB.Label lblPersNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   330
         Left            =   1200
         TabIndex        =   7
         Top             =   960
         Width           =   4545
      End
   End
   Begin VB.Frame fraCajaChica 
      Caption         =   "Caja Chica"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   825
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8355
      Begin Sicmact.TxtBuscar txtBuscarAreaCH 
         Height          =   345
         Left            =   1200
         TabIndex        =   1
         Top             =   270
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   609
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         EnabledText     =   0   'False
      End
      Begin VB.Label lblAreaAge 
         AutoSize        =   -1  'True
         Caption         =   "Area/Agencia : "
         Height          =   210
         Left            =   90
         TabIndex        =   5
         Top             =   330
         Width           =   1140
      End
      Begin VB.Label lblCajaChicaDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2640
         TabIndex        =   4
         Top             =   270
         Width           =   4485
      End
      Begin VB.Label lblNro 
         AutoSize        =   -1  'True
         Caption         =   "N° :"
         Height          =   210
         Left            =   7380
         TabIndex        =   3
         Top             =   330
         Width           =   255
      End
      Begin VB.Label lblNroProcCH 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   7665
         TabIndex        =   2
         Top             =   270
         Width           =   525
      End
   End
End
Attribute VB_Name = "frmCajaChicaCambioEncargado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oCajaCH As nCajaChica
Dim oArendir As NARendir

'ARLO20170208****
Dim objPista As COMManejador.Pista
'************

Private Sub cmdCambiar_Click()
Dim oCon As NContFunciones
Dim lsMovNro As String
Dim lsTexto As String

    If MsgBox("¿Desea realizar Definir al Cambio de Responsable de Caja Chica?", vbYesNo + vbQuestion, "Aviso") = vbYes Then
    
        Set oCon = New NContFunciones
        lsMovNro = oCon.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        
        lsTexto = oCajaCH.GrabaResponsableNuevoCH(lsMovNro, gsOpeCod, "Cambio Responsable Caja Chica - Definido", Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2), Val(lblNroProcCH), lblPerscod, TxtBuscarPersCod, lblSaldoActual, 0)
        
        If lsTexto = "" Then
            MsgBox "No se realizo la operación" & vbCrLf & "Verifique los datos del nuevo responsable", vbInformation, "Aviso"
        Else
            MsgBox "El proceso se realizo con exito", vbInformation, "Aviso"
            cmdCambiar.Enabled = False
            
            'ARLO20170208
            Set objPista = New COMManejador.Pista
            'gsOpeCod = LogPistaCierreDiarioCont
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", Me.Caption & " Se Cambio de Encargado Caja Chica del  Area : " & lblCajaChicaDesc & " |Encargado Actual: " & lblPersNombre _
            & "|Nuevo Encargado : " & lblNomNuevoEncargado & "|Monto : " & lblSaldoActual
            Set objPista = Nothing
            '*******
        End If
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Set oArendir = New NARendir
Set oCajaCH = New nCajaChica

    txtBuscarAreaCH.psRaiz = "Areas - Agencias"
    txtBuscarAreaCH.rs = oArendir.EmiteCajasChicas
    cmdCambiar.Enabled = False
    TxtBuscarPersCod.Enabled = False
    
End Sub

Private Sub txtBuscarAreaCH_EmiteDatos()
    lblCajaChicaDesc = txtBuscarAreaCH.psDescripcion
    lblNroProcCH = oCajaCH.GetDatosCajaChica(Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2), NroCajaChica)
    Call CargaDatosCajaChicaxCambiar(Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2), Val(lblNroProcCH))
End Sub

Private Sub TxtBuscarPersCod_EmiteDatos()
    lblNomNuevoEncargado = ""
    lblNroDniNuevo = ""
    cmdCambiar.Enabled = False
    If TxtBuscarPersCod.Text = "" Then Exit Sub
    usu.DatosPers TxtBuscarPersCod.Text
    If usu.PersCod = "" Then
        MsgBox "Persona no Válida o no se Encuentra Registrada como Trabajador en la Institución", vbInformation, "Aviso"
    Else
        If (Len(txtBuscarAreaCH.Text) = 3 And Mid(txtBuscarAreaCH.Text, 1, 3) <> Trim(usu.cAreaCodAct)) Or (Len(txtBuscarAreaCH.Text) = 5 And Mid(txtBuscarAreaCH.Text, 1, 3) <> Trim(usu.cAreaCodAct)) Then
            MsgBox "Persona no pertenece al área Seleccionada", vbInformation, "Aviso"
            TxtBuscarPersCod.Text = ""
            usu.DatosPers TxtBuscarPersCod.Text
            cmdCambiar.Enabled = False
        End If
    End If
    AsignaValores 2
End Sub

Private Sub AsignaValores(ByVal nTpoPer As Integer)
    If nTpoPer = 1 Then
        lblPersNombre = PstaNombre(usu.UserNom)
        lblNroDni = IIf(usu.NroDNIUser = "", usu.NroRucUser, usu.NroDNIUser)
        TxtBuscarPersCod.Enabled = True
        If TxtBuscarPersCod.Text <> "" Then
            TxtBuscarPersCod.Text = ""
            lblNomNuevoEncargado = ""
            lblNroDniNuevo = ""
            cmdCambiar.Enabled = False
        End If
    ElseIf nTpoPer = 2 Then
        If TxtBuscarPersCod.Text = "" Then
            cmdCambiar.Enabled = False
        Else
            cmdCambiar.Enabled = True
        End If
        lblNomNuevoEncargado = PstaNombre(usu.UserNom)
        lblNroDniNuevo = IIf(usu.NroDNIUser = "", usu.NroRucUser, usu.NroDNIUser)
    End If
End Sub

Private Sub CargaDatosCajaChicaxCambiar(ByVal psAreaCh As String, ByVal psAgeCh As String, ByVal pnProcNro As Integer)
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    Set rs = oCajaCH.devolverDatosCajaChicaxCambiar(psAreaCh, psAgeCh, pnProcNro)
    If Not rs.EOF And Not rs.BOF Then
        lblPerscod = rs!cPersCod
        usu.DatosPers rs!cPersCod
        AsignaValores 1
        lblSaldoActual.Caption = Format(rs!nSaldo, "#,#0.00")
    Else
        MsgBox "El Cambio de responsable de una Caja Chica no se puede realizar consulte con sistemas.", vbInformation, "Aviso"
        lblCajaChicaDesc = ""
        lblNroProcCH = ""
        txtBuscarAreaCH = ""
        txtBuscarAreaCH.psCodigoPersona = ""
        TxtBuscarPersCod = ""
        TxtBuscarPersCod.psCodigoPersona = ""
        lblPerscod = ""
        lblPersNombre = ""
        lblNroDni = ""
    End If
    rs.Close
    Set rs = Nothing
End Sub

