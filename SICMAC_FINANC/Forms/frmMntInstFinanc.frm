VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMntInstFinanc 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento de Instituciones Financieras "
   ClientHeight    =   6390
   ClientLeft      =   1665
   ClientTop       =   1470
   ClientWidth     =   8295
   Icon            =   "frmMntInstFinanc.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   8295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   6345
      Left            =   60
      TabIndex        =   11
      Top             =   0
      Width           =   8205
      Begin VB.PictureBox Picture1 
         Height          =   585
         Left            =   225
         ScaleHeight     =   525
         ScaleWidth      =   7755
         TabIndex        =   12
         Top             =   5610
         Width           =   7815
         Begin VB.CommandButton cmdSalir 
            Cancel          =   -1  'True
            Caption         =   "&Salir"
            CausesValidation=   0   'False
            Height          =   400
            Left            =   6630
            TabIndex        =   9
            Top             =   60
            Width           =   1100
         End
         Begin VB.CommandButton cmdAceptar 
            Caption         =   "&Aceptar"
            Height          =   400
            Left            =   5490
            TabIndex        =   8
            Top             =   60
            Visible         =   0   'False
            Width           =   1100
         End
         Begin VB.CommandButton cmdModificar 
            Caption         =   "&Modificar"
            Height          =   400
            Left            =   1140
            TabIndex        =   7
            Top             =   75
            Width           =   1100
         End
         Begin VB.CommandButton cmdEliminar 
            Caption         =   "&Eliminar"
            Height          =   400
            Left            =   2250
            TabIndex        =   6
            Top             =   75
            Width           =   1100
         End
         Begin VB.CommandButton cmdNuevo 
            Caption         =   "&Nuevo"
            Height          =   400
            Left            =   30
            TabIndex        =   5
            Top             =   75
            Width           =   1100
         End
         Begin VB.CommandButton cmdCancelar 
            Caption         =   "&Cancelar"
            CausesValidation=   0   'False
            Height          =   400
            Left            =   6630
            TabIndex        =   10
            Top             =   60
            Visible         =   0   'False
            Width           =   1100
         End
      End
      Begin MSDataGridLib.DataGrid GrdInstFinanc 
         Height          =   3630
         Left            =   240
         TabIndex        =   0
         Top             =   300
         Width           =   7845
         _ExtentX        =   13838
         _ExtentY        =   6403
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   2
         RowHeight       =   17
         RowDividerStyle =   6
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   4
         BeginProperty Column00 
            DataField       =   "cPersCod"
            Caption         =   "Codigo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "cPersNombre"
            Caption         =   "Nombre"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "sTipo"
            Caption         =   "Tipo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "cSubCtaContCod"
            Caption         =   "Sub Cuenta"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   3
            ScrollBars      =   2
            BeginProperty Column00 
               Alignment       =   2
               DividerStyle    =   6
               ColumnWidth     =   1319.811
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3644.788
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1530.142
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               ColumnWidth     =   750.047
            EndProperty
         EndProperty
      End
      Begin Sicmact.TxtBuscar TxtBPersInstFinanc 
         Height          =   315
         Left            =   330
         TabIndex        =   1
         Top             =   4350
         Width           =   2385
         _ExtentX        =   4207
         _ExtentY        =   556
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TipoBusqueda    =   3
         sTitulo         =   ""
      End
      Begin VB.ComboBox CboTipoInstFinanc 
         Height          =   315
         Left            =   330
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   4980
         Width           =   4875
      End
      Begin VB.TextBox TxtSubCta 
         Height          =   315
         Left            =   5235
         TabIndex        =   4
         Top             =   4980
         Width           =   2715
      End
      Begin VB.Label LblInstituc 
         Caption         =   "Institucion Financiera :"
         Height          =   210
         Left            =   345
         TabIndex        =   15
         Top             =   4065
         Width           =   1620
      End
      Begin VB.Label LblSubCta 
         Caption         =   "Sub Cuenta :"
         Height          =   210
         Left            =   5235
         TabIndex        =   14
         Top             =   4740
         Width           =   1020
      End
      Begin VB.Label LblTipoUnstFinan 
         Caption         =   "Tipo de Institucion :"
         Height          =   210
         Left            =   315
         TabIndex        =   13
         Top             =   4740
         Width           =   1500
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00E0E0E0&
         Height          =   1515
         Left            =   240
         Top             =   3975
         Width           =   7815
      End
      Begin VB.Label LblNomPers 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2700
         TabIndex        =   2
         Top             =   4350
         Width           =   5265
      End
   End
End
Attribute VB_Name = "frmMntInstFinanc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oIFinan As DInstFinanc
Dim ComandoEjec As Integer
Dim R As ADODB.Recordset
Dim nTipoInicio As Integer
Dim objPista As COMManejador.Pista 'ARLO20170217
Dim lsAccion As String

Public Sub InicioActualizar()
    nTipoInicio = 1 ' 1 para Actualizacion
    Me.Show 1
End Sub
Public Sub InicioConsulta()
    nTipoInicio = 2 ' 2 para consulta
    Me.Show 1
End Sub

Private Function ValidaDatos() As Boolean
    ValidaDatos = True
    If Len(Trim(TxtBPersInstFinanc.Text)) = 0 Then
        MsgBox "Falta Ingresar la Institucion Financiera", vbInformation, "Aviso"
        ValidaDatos = False
        TxtBPersInstFinanc.SetFocus
        Exit Function
    End If
    
    If CboTipoInstFinanc.ListIndex = -1 Then
        MsgBox "Falta Seleccionar el Tipo de Institucion Financiera", vbInformation, "Aviso"
        ValidaDatos = False
        CboTipoInstFinanc.SetFocus
        Exit Function
    End If
    
'    If Len(Trim(TxtSubCta.Text)) = 0 Then
'        MsgBox "Falta Ingresar la Sub Cuenta Contable para la Institucion Financiera", vbInformation, "Aviso"
'        ValidaDatos = False
'        TxtSubCta.SetFocus
'        Exit Function
'    End If
        
End Function

Private Sub AcitvaControles(ByVal pbHabilita As Boolean)
    GrdInstFinanc.Enabled = Not pbHabilita
    TxtBPersInstFinanc.Enabled = pbHabilita
    CboTipoInstFinanc.Enabled = pbHabilita
    txtSubCta.Enabled = pbHabilita
    cmdNuevo.Visible = IIf(pbHabilita, False, True)
    cmdModificar.Visible = IIf(pbHabilita, False, True)
    cmdEliminar.Visible = IIf(pbHabilita, False, True)
    cmdSalir.Visible = IIf(pbHabilita, False, True)
    cmdAceptar.Visible = pbHabilita
    cmdCancelar.Visible = pbHabilita
End Sub
Private Sub cargaControles()
    Call CargaComboConstante(gCGTipoIF, CboTipoInstFinanc)
End Sub
Private Sub CargaDatos()

    Set oIFinan = New DInstFinanc
    Set R = oIFinan.CargaInstituciones
    Set GrdInstFinanc.DataSource = R
    Set oIFinan = Nothing
    GrdInstFinanc.Refresh
End Sub
Private Sub LimpiaControles()
    TxtBPersInstFinanc.Text = ""
    lblNomPers.Caption = ""
    CboTipoInstFinanc.ListIndex = -1
    txtSubCta.Text = ""
End Sub


Private Sub CboTipoInstFinanc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtSubCta.SetFocus
    End If
End Sub

Private Sub cmdAceptar_Click()
Dim oNInstFin As NInstFinanc
Dim sError As String

On Error GoTo ErrorAceptar
    If Not ValidaDatos Then
        Exit Sub
    End If
    
    sError = ""
    Set oNInstFin = New NInstFinanc
    Select Case ComandoEjec
        Case 1 'Nuevo
            sError = oNInstFin.NuevaInstitucion(Trim(TxtBPersInstFinanc.Text), Trim(Right(CboTipoInstFinanc.Text, 15)), Trim(txtSubCta.Text))
            If Len(sError) > 0 Then
                MsgBox sError, vbInformation, "Aviso"
                Set oNInstFin = Nothing
                Exit Sub
            End If
            Call CargaDatos
            R.Find "cPersCod = '" & TxtBPersInstFinanc.Text & "'"
        Case 2 'Modificar
            sError = oNInstFin.ActualizaInstitucion(Trim(TxtBPersInstFinanc.Text), Trim(Right(CboTipoInstFinanc.Text, 15)), Trim(txtSubCta.Text), Right(R!sTipo, 2))
            If Len(sError) > 0 Then
                MsgBox sError, vbInformation, "Aviso"
                Set oNInstFin = Nothing
                Exit Sub
            End If
            Call CargaDatos
            R.Find "cPersCod = '" & TxtBPersInstFinanc.Text & "'"
    End Select
    Set oIFinan = Nothing
    GrdInstFinanc.Height = 5190
    Call LimpiaControles
    Call AcitvaControles(False)
    GrdInstFinanc.SetFocus
    ComandoEjec = -1
                'ARLO20170217
                If (ComandoEjec = 1) Then
                lsAccion = "1"
                Else
                lsAccion = "2"
                End If
                Set objPista = New COMManejador.Pista
                objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, lsAccion, Me.Caption & " Se Grabo la Operación "
                Set objPista = Nothing
                '****
    Exit Sub
    
ErrorAceptar:
    MsgBox Err.Description, vbQuestion, "Aviso"

End Sub

Private Sub cmdCancelar_Click()
    GrdInstFinanc.Height = 5190
    Call LimpiaControles
    Call AcitvaControles(False)
    GrdInstFinanc.SetFocus
    ComandoEjec = -1
    Exit Sub
End Sub

Private Sub cmdEliminar_Click()
Dim oNInstFin As NInstFinanc
Dim sError As String

On Error GoTo ERRORcmdEliminar
    If MsgBox("Se Va a Eliminar la Institucion Financiera " & R!cPersNombre & ", Desea Continuar ?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        Set oNInstFin = New NInstFinanc
        sError = oNInstFin.EliminaInstitucion(R!cPersCod, Right(R!sTipo, 2))
        If Len(sError) > 0 Then
            MsgBox sError, vbInformation, "Aviso"
        End If
        Set oNInstFin = Nothing
        Call CargaDatos
    End If
    GrdInstFinanc.SetFocus
                'ARLO20170217
                Set objPista = New COMManejador.Pista
                objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "3", Me.Caption & " Se Elimino la Institucion Financiera "
                Set objPista = Nothing
                '****
    Exit Sub
ERRORcmdEliminar:
    MsgBox Err.Description, vbExclamation, "Aviso"
End Sub

Private Sub cmdModificar_Click()
        
    Call LimpiaControles
    Call AcitvaControles(True)
    TxtBPersInstFinanc.Enabled = False
    TxtBPersInstFinanc.Text = R!cPersCod
    lblNomPers.Caption = R!cPersNombre
    CboTipoInstFinanc.ListIndex = IndiceListaCombo(CboTipoInstFinanc, Trim(Str(CInt(Trim(Right(R!sTipo, 15))))))
    txtSubCta.Text = Trim(R!cSubCtaContCod)
    GrdInstFinanc.Height = 3630
    CboTipoInstFinanc.SetFocus
    ComandoEjec = 2
End Sub

Private Sub cmdNuevo_Click()
    GrdInstFinanc.Height = 3630
    Call LimpiaControles
    Call AcitvaControles(True)
    TxtBPersInstFinanc.SetFocus
    ComandoEjec = 1
End Sub

Private Sub cmdSalir_Click()
    Set oIFinan = Nothing
    Unload Me
End Sub

Private Sub Form_Load()
    Call LimpiaControles
    Call cargaControles
    Call AcitvaControles(False)
    Call CargaDatos
    ComandoEjec = -1
    GrdInstFinanc.Height = 5190
    If nTipoInicio = 2 Then
        cmdNuevo.Enabled = False
        cmdModificar.Enabled = False
        cmdEliminar.Enabled = False
    End If
    CentraForm Me
End Sub


Private Sub GrdInstFinanc_DblClick()
    If cmdModificar.Enabled Then
        Call cmdModificar_Click
    End If
End Sub

Private Sub GrdInstFinanc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cmdModificar.Enabled Then
            Call cmdModificar_Click
        End If
    End If
End Sub

Private Sub TxtBPersInstFinanc_EmiteDatos()
    lblNomPers.Caption = TxtBPersInstFinanc.psDescripcion
End Sub

Private Sub TxtBPersInstFinanc_GotFocus()
    fEnfoque TxtBPersInstFinanc
End Sub

Private Sub TxtBPersInstFinanc_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        CboTipoInstFinanc.SetFocus
    End If
End Sub

Private Sub TxtSubCta_GotFocus()
    fEnfoque txtSubCta
End Sub

Private Sub TxtSubCta_KeyPress(KeyAscii As Integer)

    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        cmdAceptar.SetFocus
    End If
End Sub
