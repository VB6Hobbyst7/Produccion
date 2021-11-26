VERSION 5.00
Begin VB.Form frmRCDMantMaestroPersona 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informe RCD - Mantenimiento de Maestro Persona"
   ClientHeight    =   5010
   ClientLeft      =   2145
   ClientTop       =   2025
   ClientWidth     =   7260
   Icon            =   "frmRCDMantMaestroPersona.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   390
      Left            =   5700
      TabIndex        =   29
      Top             =   4515
      Width           =   1320
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   390
      Left            =   1485
      TabIndex        =   16
      Top             =   4515
      Width           =   1320
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   390
      Left            =   165
      TabIndex        =   15
      Top             =   4515
      Width           =   1320
   End
   Begin VB.Frame Frame2 
      Caption         =   "Datos del Cliente"
      Height          =   3240
      Left            =   135
      TabIndex        =   17
      Top             =   1200
      Width           =   6900
      Begin VB.Frame frajuridicos 
         Height          =   960
         Left            =   1980
         TabIndex        =   30
         Top             =   2055
         Width           =   4605
         Begin VB.TextBox txtRegPub 
            Height          =   315
            Left            =   1860
            MaxLength       =   15
            TabIndex        =   33
            Top             =   195
            Width           =   1935
         End
         Begin VB.TextBox txtSiglas 
            Height          =   315
            Left            =   2970
            MaxLength       =   15
            TabIndex        =   32
            Top             =   555
            Width           =   1470
         End
         Begin VB.TextBox txtMagEmp 
            Height          =   315
            Left            =   1860
            MaxLength       =   1
            TabIndex        =   31
            Top             =   540
            Width           =   405
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Nº Registros Publicos :"
            Height          =   195
            Left            =   210
            TabIndex        =   36
            Top             =   225
            Width           =   1620
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Siglas :"
            Height          =   195
            Left            =   2355
            TabIndex        =   35
            Top             =   615
            Width           =   510
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Magnitud Empresarial :"
            Height          =   195
            Left            =   210
            TabIndex        =   34
            Top             =   555
            Width           =   1605
         End
      End
      Begin VB.TextBox txtActecon 
         Height          =   315
         Left            =   930
         MaxLength       =   4
         TabIndex        =   14
         Top             =   2175
         Width           =   825
      End
      Begin VB.TextBox txtCodSBS 
         Height          =   315
         Left            =   5385
         MaxLength       =   10
         TabIndex        =   8
         Top             =   285
         Width           =   1380
      End
      Begin VB.Frame Frame3 
         Caption         =   "Documento Jurídico"
         Height          =   1035
         Left            =   3465
         TabIndex        =   22
         Top             =   1020
         Width           =   2580
         Begin VB.TextBox txtdocJur 
            Height          =   315
            Left            =   750
            MaxLength       =   15
            TabIndex        =   13
            Top             =   555
            Width           =   1620
         End
         Begin VB.TextBox txttipoJur 
            Height          =   315
            Left            =   750
            MaxLength       =   1
            TabIndex        =   12
            Top             =   210
            Width           =   405
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Numero :"
            Height          =   195
            Left            =   90
            TabIndex        =   26
            Top             =   555
            Width           =   645
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Tipo :"
            Height          =   195
            Left            =   120
            TabIndex        =   25
            Top             =   255
            Width           =   405
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Documento Natural "
         Height          =   1035
         Left            =   270
         TabIndex        =   21
         Top             =   1020
         Width           =   2535
         Begin VB.TextBox txtdocnat 
            Height          =   315
            Left            =   825
            MaxLength       =   15
            TabIndex        =   11
            Top             =   570
            Width           =   1560
         End
         Begin VB.TextBox txtTipoNat 
            Height          =   315
            Left            =   825
            MaxLength       =   1
            TabIndex        =   10
            Top             =   225
            Width           =   405
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Numero :"
            Height          =   195
            Left            =   165
            TabIndex        =   24
            Top             =   570
            Width           =   645
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Tipo :"
            Height          =   195
            Left            =   195
            TabIndex        =   23
            Top             =   270
            Width           =   405
         End
      End
      Begin VB.TextBox txtTipoPers 
         Height          =   315
         Left            =   3795
         MaxLength       =   1
         TabIndex        =   7
         Top             =   285
         Width           =   405
      End
      Begin VB.TextBox txtNombre 
         Height          =   315
         Left            =   1020
         MaxLength       =   150
         TabIndex        =   9
         Top             =   645
         Width           =   5760
      End
      Begin VB.TextBox txtcodigo 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1005
         MaxLength       =   13
         TabIndex        =   6
         Top             =   270
         Width           =   1380
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Ciiu :"
         Height          =   195
         Left            =   300
         TabIndex        =   28
         Top             =   2220
         Width           =   345
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Codigo SBS :"
         Height          =   195
         Left            =   4365
         TabIndex        =   27
         Top             =   345
         Width           =   945
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Persona :"
         Height          =   195
         Left            =   2715
         TabIndex        =   20
         Top             =   315
         Width           =   1035
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nombre :"
         Height          =   195
         Left            =   330
         TabIndex        =   19
         Top             =   705
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Codigo :"
         Height          =   195
         Left            =   330
         TabIndex        =   18
         Top             =   330
         Width           =   585
      End
   End
   Begin VB.ListBox lstClientes 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   5505
      TabIndex        =   5
      Top             =   150
      Width           =   1485
   End
   Begin VB.Frame fraBuscar 
      Caption         =   "Buscar por :"
      Height          =   960
      Left            =   120
      TabIndex        =   0
      Top             =   75
      Width           =   3180
      Begin VB.OptionButton optOpcion 
         Caption         =   "O&tros"
         Height          =   255
         Index           =   1
         Left            =   210
         TabIndex        =   3
         Top             =   555
         Width           =   930
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
         Height          =   345
         Left            =   1545
         TabIndex        =   4
         Top             =   555
         Width           =   1320
      End
      Begin VB.TextBox txtBuscar 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1560
         MaxLength       =   9
         TabIndex        =   2
         Top             =   180
         Width           =   1260
      End
      Begin VB.OptionButton optOpcion 
         Caption         =   "Có&digo SBS"
         Height          =   255
         Index           =   0
         Left            =   210
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmRCDMantMaestroPersona"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fsServConsol As String

Private Sub cmdBuscar_Click()
Dim loPers As COMDPersona.UCOMPersona
Dim lsPersCod As String, lsPersNombre As String

Set loPers = New COMDPersona.UCOMPersona
    Set loPers = frmBuscaPersona.Inicio
    If loPers Is Nothing Then Exit Sub
    lsPersCod = loPers.sPersCod
    lsPersNombre = loPers.sPersNombre
Set loPers = Nothing

If Len(Trim(lsPersCod)) > 0 Then
    CargaClientes 0, lsPersCod
End If

End Sub

Private Sub cmdCancelar_Click()

On Error GoTo ErrorCancelar
Limpiar
lstClientes.Clear
txtBuscar = ""
If Me.txtBuscar.Visible Then
    Me.txtBuscar.SetFocus
Else
    Me.cmdBuscar.SetFocus
End If
Exit Sub
ErrorCancelar:
    MsgBox "Error Nº : [" & Err.Number & " ] " & Err.Description, vbInformation, "Aviso"
End Sub

Private Sub cmdGrabar_Click()
Dim oRCD As COMDCredito.DCOMRCD

If Not Valida Then Exit Sub
If MsgBox("Desea Grabar la Información?", vbYesNo + vbQuestion, "Aviso") = vbNo Then Exit Sub

Set oRCD = New COMDCredito.DCOMRCD
Call oRCD.ModificarRCDMaestroPersona(fsServConsol, Trim(Replace(txtNombre.Text, "'", "''")), _
                                     Trim(txtCodSBS.Text), Trim(txtActecon.Text), Trim(txttipoJur.Text), _
                                     Trim(txtdocJur.Text), Trim(txtTipoNat.Text), Trim(txtdocnat.Text), _
                                     Trim(txtRegPub.Text), Trim(txtMagEmp.Text), Trim(txtTipoPers.Text), _
                                     Trim(txtSiglas.Text), Trim(txtcodigo.Text))

Set oRCD = Nothing
Me.cmdCancelar.value = True
Exit Sub
ErrorGrabaMaestro:
    MsgBox "Error Nº [" & Err.Number & "] " & Err.Description, vbInformation, "Aviso"

'Dim lsSQL  As String
'Dim loBase As COMConecta.DCOMConecta
'On Error GoTo ErrorGrabaMaestro
'If Valida Then
'    If MsgBox("Desea Grabar la Información", vbYesNo + vbQuestion, "Aviso") = vbYes Then
'        lsSQL = "UPDATE " & fsServConsol & "RCDMaestroPersona SET " _
'            & "cpersnom='" & Trim(Replace(txtNombre, "'", "''")) & "'," _
'            & "cCodSBS='" & Trim(txtcodsbs) & "'," _
'            & "cActEcon='" & Trim(txtActecon) & "'," _
'            & "cTidoTr=" & IIf(Len(Trim(txttipoJur)) = 0, "Null", "'" & Trim(txttipoJur) & "'") & "," _
'            & "cNudoTr=" & IIf(Len(Trim(txtDocJur)) = 0, "Null", "'" & Trim(txtDocJur) & "'") & "," _
'            & "cTiDoci=" & IIf(Len(Trim(txtTipoNat)) = 0, "Null", "'" & Trim(txtTipoNat) & "'") & "," _
'            & "cNuDoci=" & IIf(Len(Trim(txtdocnat)) = 0, "Null", "'" & Trim(txtdocnat) & "'") & "," _
'            & "cCodRegPub=" & IIf(Len(Trim(txtRegPub)) = 0, "Null", "'" & Trim(txtRegPub) & "'") & "," _
'            & "cMagEmp=" & IIf(Len(Trim(txtMagEmp)) = 0, "Null", "'" & Trim(txtMagEmp) & "'") & "," _
'            & "cTipPers='" & Trim(txtTipoPers) & "'," _
'            & "cSiglas=" & IIf(Len(Trim(TxtSiglas)) = 0, "Null", "'" & Trim(TxtSiglas) & "'") & " " _
'            & "WHERE cPersCod='" & Trim(txtCodigo) & "'"
'        Set loBase = New COMConecta.DCOMConecta
'            loBase.AbreConexion
'            loBase.Ejecutar (lsSQL)
'        Set loBase = Nothing
'        Me.CmdCancelar.value = True
'    End If
'End If
'Exit Sub
'ErrorGrabaMaestro:
'    MsgBox "Error Nº [" & Err.Number & "] " & Err.Description, vbInformation, "Aviso"
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim loConstSist As COMDConstSistema.NCOMConstSistema
    Set loConstSist = New COMDConstSistema.NCOMConstSistema
        fsServConsol = loConstSist.LeeConstSistema(gConstSistServCentralRiesgos)
    Set loConstSist = Nothing
Me.Icon = LoadPicture(App.path & gsRutaIcono)
End Sub

Private Sub lstClientes_DblClick()
If Me.lstClientes.ListCount > 0 Then
    DatosClientes Trim(lstClientes)
End If
End Sub

Private Sub lstClientes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Me.lstClientes.ListCount > 0 Then
        If Len(Trim(lstClientes)) > 0 Then
            DatosClientes Trim(lstClientes)
            If txtcodigo.Enabled Then Me.txtcodigo.SetFocus
        Else
            MsgBox "Seleccione Código de Cliente", vbInformation, "Aviso"
            Me.lstClientes.SetFocus
        End If
    End If
End If
End Sub

Private Sub optOpcion_Click(Index As Integer)
Select Case Index
    Case 0
        txtBuscar.Visible = True
        Me.cmdBuscar.Visible = False
    Case 1
        txtBuscar.Visible = False
        Me.cmdBuscar.Visible = True
End Select
End Sub

Private Sub optOpcion_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    txtBuscar.SetFocus
End If
End Sub

Private Sub txtActecon_GotFocus()
fEnfoque txtActecon
End Sub

Private Sub txtActecon_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
If KeyAscii = 13 Then
    If txtRegPub.Enabled And txtRegPub.Visible Then
        txtRegPub.SetFocus
    Else
        Me.cmdGrabar.SetFocus
    End If
End If
End Sub

Private Sub txtBuscar_GotFocus()
fEnfoque txtBuscar
End Sub

Private Sub txtBuscar_KeyPress(KeyAscii As Integer)
Dim Opcion As Integer
KeyAscii = NumerosEnteros(KeyAscii)
If KeyAscii = 13 Then
    If Len(Trim(txtBuscar)) >= 8 Then
        CargaClientes 1, Trim(Me.txtBuscar)
    Else
        MsgBox "Código no Válido", vbInformation, "Aviso"
    End If
End If
End Sub

Private Sub txtcodigo_GotFocus()
fEnfoque txtcodigo
End Sub

Private Sub txtcodigo_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
If KeyAscii = 13 Then
    Me.txtTipoPers.SetFocus
End If
End Sub

Private Sub txtCodSBS_GotFocus()
fEnfoque txtCodSBS
End Sub

Private Sub txtCodSBS_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
If KeyAscii = 13 Then
    Me.txtNombre.SetFocus
End If

End Sub

Private Sub txtdocJur_GotFocus()
fEnfoque txtdocJur
End Sub

Private Sub txtdocJur_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
If KeyAscii = 13 Then
    Me.txtActecon.SetFocus
End If
End Sub

Private Sub txtdocnat_GotFocus()
fEnfoque txtdocnat
End Sub

Private Sub txtdocnat_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
If KeyAscii = 13 Then
    Me.txttipoJur.SetFocus
End If
End Sub

Private Sub txtMagEmp_GotFocus()
fEnfoque txtMagEmp
End Sub

Private Sub txtMagEmp_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
If KeyAscii = 13 Then
     Me.txtSiglas.SetFocus
End If
End Sub

Private Sub txtNombre_GotFocus()
fEnfoque txtNombre
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
KeyAscii = UCase(KeyAscii)
If KeyAscii = 13 Then
    Me.txtTipoNat.SetFocus
End If

End Sub

Private Sub txtRegPub_GotFocus()
fEnfoque txtRegPub
End Sub

Private Sub txtRegPub_KeyPress(KeyAscii As Integer)
KeyAscii = UCase(KeyAscii)
If KeyAscii = 13 Then
    If Me.txtSiglas.Visible Then
        Me.txtMagEmp.SetFocus
    End If
End If
End Sub

Private Sub txtSiglas_GotFocus()
fEnfoque txtSiglas
End Sub

Private Sub txtSiglas_KeyPress(KeyAscii As Integer)
KeyAscii = UCase(KeyAscii)
If KeyAscii = 13 Then
    If Me.txtMagEmp.Visible Then
          Me.cmdGrabar.SetFocus
    End If
End If
End Sub

Private Sub txttipoJur_GotFocus()
fEnfoque txttipoJur
End Sub

Private Sub txttipoJur_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
If KeyAscii = 13 Then
    Me.txtdocJur.SetFocus
End If
End Sub

Private Sub txtTipoNat_GotFocus()
fEnfoque txtTipoNat
End Sub

Private Sub txtTipoNat_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
If KeyAscii = 13 Then
    Me.txtdocnat.SetFocus
End If

End Sub

Private Sub txtTipoPers_Change()
If txtTipoPers = "1" Then
    frajuridicos.Visible = False
Else
    frajuridicos.Visible = True
End If
End Sub

Private Sub txtTipoPers_GotFocus()
fEnfoque txtTipoPers
End Sub

Private Sub txtTipoPers_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
If KeyAscii = 13 Then
    Me.txtCodSBS.SetFocus
End If

End Sub
Private Sub CargaClientes(Seleccion As Integer, lsDato As String)
Dim lrs As ADODB.Recordset
Dim oRCD As COMDCredito.DCOMRCD

Set oRCD = New COMDCredito.DCOMRCD
Set lrs = oRCD.CargaClientes(fsServConsol, Seleccion, lsDato)
Set oRCD = Nothing
Me.lstClientes.Clear

    If Not (lrs.BOF And lrs.EOF) Then
        Do While Not lrs.EOF
            Me.lstClientes.AddItem Trim(lrs!cPersCod)
            lrs.MoveNext
        Loop
        lstClientes.SetFocus
    Else
        Limpiar
        MsgBox "Datos no encontrados", vbInformation, "Aviso"
   End If
    lrs.Close
    Set lrs = Nothing

'Dim lsSQL  As String
'Dim lrs As ADODB.Recordset
'Dim loBase As COMConecta.DCOMConecta
'
'Select Case Seleccion
'    Case 0
'        lsSQL = "Select cPersCod From " & fsServConsol & "RCDMaestroPersona Where cPersCod= '" & lsDato & "'"
'    Case 1
'        lsSQL = "Select cPersCod From " & fsServConsol & "RCDMaestroPersona Where cCodSBS = '" & lsDato & "'"
'End Select
'
'Me.lstClientes.Clear
'Set loBase = New COMConecta.DCOMConecta
'    loBase.AbreConexion
'    Set lrs = loBase.CargaRecordSet(lsSQL)
'Set loBase = Nothing
'
'    If Not (lrs.BOF And lrs.EOF) Then
'        Do While Not lrs.EOF
'            Me.lstClientes.AddItem Trim(lrs!cPersCod)
'            lrs.MoveNext
'        Loop
'        lstClientes.SetFocus
'    Else
'        Limpiar
'        MsgBox "Datos no encontrados", vbInformation, "Aviso"
'    End If
'    lrs.Close
'    Set lrs = Nothing
'
End Sub

Private Sub DatosClientes(lsCodPers As String)
Dim lrs As ADODB.Recordset
Dim oRCD As COMDCredito.DCOMRCD

Set oRCD = New COMDCredito.DCOMRCD
Set lrs = oRCD.CargaDatosCliente(fsServConsol, lsCodPers)
Set oRCD = Nothing

If Not (lrs.BOF And lrs.EOF) Then
    txtcodigo = Trim(lrs!cPersCod)
    txtCodSBS = IIf(IsNull(lrs!cCodSBS), "", Trim(lrs!cCodSBS))
    txtNombre = Trim(lrs!cPersNom)
    txtActecon = IIf(IsNull(lrs!cActEcon), "", lrs!cActEcon)
    txttipoJur = IIf(IsNull(lrs!cTidoTr), "", lrs!cTidoTr)
    txtdocJur = IIf(IsNull(lrs!cNudoTr), "", lrs!cNudoTr)
    txtTipoNat = IIf(IsNull(lrs!ctidoci), "", lrs!ctidoci)
    txtdocnat = IIf(IsNull(lrs!cnudoci), "", lrs!cnudoci)
    txtTipoPers = IIf(IsNull(lrs!cTipPers), "", lrs!cTipPers)
    If lrs!cTipPers <> "1" Then
        txtRegPub = IIf(IsNull(lrs!ccodregpub), "", lrs!ccodregpub)
        txtMagEmp = IIf(IsNull(lrs!cMagEmp), "", lrs!cMagEmp)
        txtSiglas = IIf(IsNull(lrs!cSiglas), "", lrs!cSiglas)
    Else
        frajuridicos.Visible = False
    End If
End If
lrs.Close
Set lrs = Nothing

'Dim lsSQL As String
'Dim lrs As ADODB.Recordset
'Dim loBase As COMConecta.DCOMConecta
'
'Limpiar
'
'lsSQL = "Select * From " & fsServConsol & "RCDMaestroPersona where cPersCod='" & lsCodPers & "'"
'
'Set loBase = New COMConecta.DCOMConecta
'    loBase.AbreConexion
'    Set lrs = loBase.CargaRecordSet(lsSQL)
'Set loBase = Nothing
'
'If Not (lrs.BOF And lrs.EOF) Then
'    txtcodigo = Trim(lrs!cPersCod)
'    txtCodSBS = IIf(IsNull(lrs!cCodSBS), "", Trim(lrs!cCodSBS))
'    txtNombre = Trim(lrs!cPersNom)
'    txtActecon = IIf(IsNull(lrs!cActEcon), "", lrs!cActEcon)
'    txttipoJur = IIf(IsNull(lrs!cTidoTr), "", lrs!cTidoTr)
'    txtdocJur = IIf(IsNull(lrs!cNudoTr), "", lrs!cNudoTr)
'    txtTipoNat = IIf(IsNull(lrs!ctidoci), "", lrs!ctidoci)
'    txtdocnat = IIf(IsNull(lrs!cnudoci), "", lrs!cnudoci)
'    txtTipoPers = IIf(IsNull(lrs!cTipPers), "", lrs!cTipPers)
'    If lrs!cTipPers <> "1" Then
'        txtRegPub = IIf(IsNull(lrs!ccodregpub), "", lrs!ccodregpub)
'        txtMagEmp = IIf(IsNull(lrs!cMagEmp), "", lrs!cMagEmp)
'        txtSiglas = IIf(IsNull(lrs!cSiglas), "", lrs!cSiglas)
'    Else
'        frajuridicos.Visible = False
'    End If
'End If
'lrs.Close
'Set lrs = Nothing
End Sub
Private Sub Limpiar()
    txtcodigo = ""
    txtCodSBS = ""
    txtNombre = ""
    txtActecon = ""
    txttipoJur = ""
    txtdocJur = ""
    txtTipoNat = ""
    txtdocnat = ""
    txtRegPub = ""
    txtMagEmp = ""
    txtTipoPers = ""
    txtSiglas = ""
End Sub

Private Function Valida() As Boolean
If Len(Trim(Me.txtcodigo)) < 8 Then
    MsgBox "Codigo de Cliente no válido", vbInformation, "Aviso"
    Me.txtcodigo.SetFocus
    Valida = False
    Exit Function
Else
    Valida = True

End If
If Len(Trim(txtTipoPers)) = 0 Then
    MsgBox "Tipo de Persona no válido", vbInformation, "Aviso"
    txtTipoPers.SetFocus
    Valida = False
    Exit Function
Else
    Valida = True
End If
If Len(Trim(txtCodSBS)) < 9 Then
    MsgBox "Código de SBS no válido", vbInformation, "Aviso"
    txtCodSBS.SetFocus
    Valida = False
    Exit Function
Else
    Valida = True
End If
If Len(Trim(txtNombre)) = 0 Then
    MsgBox "Nombre no válido", vbInformation, "Aviso"
    txtNombre.SetFocus
    Valida = False
    Exit Function
Else
    Valida = True
End If

If txtTipoPers <> "1" Then
    If Len(Trim(txttipoJur)) = 0 Then
        If MsgBox("Tipo de Documento Jurídico no válido. Desea Proseguir? ", vbQuestion + vbYesNo, "Aviso") = vbYes Then
            Valida = True
        Else
            txttipoJur.SetFocus
            Valida = False
            Exit Function
        End If
    Else
        Valida = True
    End If
    If Len(Trim(txtdocJur)) = 0 Then
        If MsgBox("Documento Jurídico no válido. Desea Proseguir? ", vbQuestion + vbYesNo, "Aviso") = vbYes Then
            Valida = True
        Else
            txtdocJur.SetFocus
            Valida = False
            Exit Function
        End If
    Else
        Valida = True
    End If
Else
    If Len(Trim(txtTipoNat)) = 0 Then
        MsgBox "Tipo de Documento Natural no válido", vbInformation, "Aviso"
        txtTipoNat.SetFocus
        Valida = False
        Exit Function
    Else
        Valida = True
    End If

    If Len(Trim(txtActecon)) > 0 And Len(Trim(txtActecon)) < 4 Then
        MsgBox "Actividad Económica no Válida ", vbInformation, "Aviso"
        txtActecon.SetFocus
        Valida = False
        Exit Function
    Else
        Valida = True
    End If
End If
End Function
