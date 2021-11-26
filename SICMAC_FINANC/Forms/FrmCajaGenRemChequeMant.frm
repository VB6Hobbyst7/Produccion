VERSION 5.00
Begin VB.Form FrmCajaGenRemChequeMant 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento de Remesas con Cheque"
   ClientHeight    =   5865
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5640
   Icon            =   "FrmCajaGenRemChequeMant.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraOrigen 
      Caption         =   "Nueva Agencia:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   690
      Left            =   60
      TabIndex        =   21
      Top             =   4260
      Width           =   5385
      Begin Sicmact.TxtBuscar txtCtaOrig 
         Height          =   345
         Left            =   240
         TabIndex        =   22
         Top             =   225
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   609
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
         sTitulo         =   ""
      End
      Begin VB.Label lblDescCtaOrig 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2145
         TabIndex        =   23
         Top             =   240
         Width           =   3105
      End
   End
   Begin VB.Frame Frame3 
      Height          =   585
      Left            =   90
      TabIndex        =   15
      Top             =   5100
      Width           =   5385
      Begin VB.CheckBox ChkAnular 
         Caption         =   "Anular"
         Height          =   345
         Left            =   2640
         TabIndex        =   24
         Top             =   150
         Width           =   1425
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   315
         Left            =   4260
         TabIndex        =   18
         Top             =   180
         Width           =   1065
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "Cancelar"
         Height          =   315
         Left            =   1350
         TabIndex        =   17
         Top             =   180
         Width           =   1065
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "Aceptar"
         Height          =   315
         Left            =   300
         TabIndex        =   16
         Top             =   180
         Width           =   915
      End
   End
   Begin VB.Frame Frame2 
      Height          =   705
      Left            =   60
      TabIndex        =   1
      Top             =   0
      Width           =   5325
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Buscar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3630
         TabIndex        =   4
         Top             =   180
         Width           =   975
      End
      Begin VB.TextBox txtCheque 
         Appearance      =   0  'Flat
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
         Height          =   375
         Left            =   870
         TabIndex        =   3
         Top             =   180
         Width           =   2595
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cheque:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   30
         TabIndex        =   2
         Top             =   240
         Width           =   750
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3435
      Left            =   60
      TabIndex        =   0
      Top             =   690
      Width           =   5355
      Begin VB.TextBox txtInstitucion 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   990
         Width           =   4215
      End
      Begin VB.TextBox txtGlosa 
         Height          =   1365
         Left            =   150
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   14
         Top             =   1950
         Width           =   4635
      End
      Begin VB.OptionButton OptDolares 
         Caption         =   "Dolares"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3120
         TabIndex        =   12
         Top             =   1650
         Width           =   1005
      End
      Begin VB.OptionButton OptSoles 
         Caption         =   "Soles"
         Enabled         =   0   'False
         Height          =   255
         Left            =   4200
         TabIndex        =   11
         Top             =   1650
         Width           =   765
      End
      Begin VB.TextBox txtMonto 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "0.00"
         Top             =   1380
         Width           =   1905
      End
      Begin VB.TextBox txtAgencia 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   600
         Width           =   2895
      End
      Begin VB.TextBox txtNroCuenta 
         Appearance      =   0  'Flat
         Height          =   390
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   180
         Width           =   2925
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Institucion:"
         Height          =   195
         Left            =   180
         TabIndex        =   20
         Top             =   1080
         Width           =   765
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Glosa:"
         Height          =   195
         Left            =   210
         TabIndex        =   13
         Top             =   1620
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Monto:"
         Height          =   195
         Left            =   180
         TabIndex        =   9
         Top             =   1380
         Width           =   495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Agencia:"
         Height          =   195
         Left            =   180
         TabIndex        =   7
         Top             =   660
         Width           =   630
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nro Cuenta:"
         Height          =   195
         Left            =   180
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "FrmCajaGenRemChequeMant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nMovNroRef As Long
Dim objPista As COMManejador.Pista 'ARLO20170217

Private Sub cmdAceptar_Click()
    'Dim onCredito As NCredito
'    If txtCheque <> "" And txtAgencia <> "" And txtmonto <> "" And txtGlosa <> "" Then
'        Set onCredito = New NCredito
'           If onCredito.InsertaRemesaCheque(gsCodAge, gsCodUser, gdFecSis, txtmonto, txtCheque, IIf(OptSoles.value = True, 1, 2), nMovNroRef, txtGlosa) = True Then
'                MsgBox "Se registro correctamente"
'                'Call Imprimir
'                Call cmdCancelar_Click
'           Else
'                MsgBox "Error en el registro"
'                Call cmdCancelar_Click
'           End If
'    Else
'        MsgBox "Los datos esta incompletos"
'        Call cmdCancelar_Click
'    End If
Dim oMov As DMov

Dim pnMoneda As Integer



           
If Mid(gsOpeCod, 3, 1) = "1" Then
    pnMoneda = 1
Else
    pnMoneda = 2
End If

Set oMov = New DMov

If MsgBox("Desea Grabar el movimiento respectivo??", vbYesNo + vbQuestion, "Aviso") = vbYes Then
   
   If ChkAnular.value = 0 Then
        oMov.ActualizaOpeCheq nMovNroRef, 1, 1, "", txtCtaOrig
        oMov.ActualizaMovCheq nMovNroRef, txtGlosa
                
                'ARLO20170217
                Set objPista = New COMManejador.Pista
                objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", Me.Caption & " Se Grabo la Operación "
                Set objPista = Nothing
                '****

        
        If MsgBox("Desea realizar otra operación??", vbYesNo + vbQuestion, "Aviso") = vbYes Then
            Call cmdCancelar_Click
        Else
            Unload Me
        End If
    Else
       oMov.AnulaMovCheq nMovNroRef
    End If
    
End If
    Call cmdCancelar_Click
End Sub

Private Sub cmdBuscar_Click()
    Call CargarDatos
    txtCheque.Enabled = False
End Sub

Private Sub cmdCancelar_Click()
    txtCheque = ""
    txtNroCuenta = ""
    txtAgencia = ""
    txtMonto = Format("0.00", "#0.00")
    OptSoles.value = False
    OptDolares.value = False
    txtGlosa = ""
    CmdAceptar.Enabled = True
    CmdCancelar.Enabled = True
    txtCheque.Enabled = True
    txtInstitucion = ""
    txtCtaOrig = ""
    lblDescCtaOrig = ""
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Sub CargarDatos()
    Dim rs As ADODB.Recordset
    Dim oMov As DMov
    Dim cAgencia As String
    
    Set oMov = New DMov
    Set rs = oMov.GetDatosCheque(Trim(txtCheque))
    Set oMov = Nothing
    
    If Not rs.EOF And Not rs.BOF Then
        nMovNroRef = rs!nMovNro
        txtNroCuenta = rs!cCtaIfCod
        txtAgencia = rs!cAgencia
        txtMonto = Format(rs!nMonto, "#0.00")
        If rs!nMoneda = "1" Then
            txtMonto.BackColor = vbWhite
            OptSoles.value = True
        Else
            txtMonto.BackColor = vbGreen
            OptDolares.value = True
        End If
        txtInstitucion = rs!cPersNombre
        cAgencia = rs!cAgeCod
    End If
    Set rs = Nothing
    'If cAgencia <> gsCodAge Then
    '    MsgBox "Este cheque no le corresponde a la agencia", vbInformation, "AVISO"
    '    CmdAceptar.Enabled = False
        'CmdCancelar.Enabled = False
    'End If
End Sub



Private Sub Form_Load()
Dim oAgencias As DActualizaDatosArea
Set oAgencias = New DActualizaDatosArea
CentraForm Me
txtCtaOrig.rs = oAgencias.GetAgencias

End Sub


Private Sub txtCheque_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdBuscar_Click
    End If
End Sub

Sub Imprimir()
Dim lsCaption
'Dim oBol As NCapImpBoleta

'lsCaption = "Remesa con cheque "
'
'        Set oBol = New NCapImpBoleta
'        Do
'
'            oBol.ImprimeBoleta "OTRAS OPERACIONES", Left(lsCaption, 15), "00000000" & IIf(OptSoles.value = True, "1", "2"), CStr(Format(txtmonto, "#0.00")), gsNomAge, "00000000" & IIf(OptSoles.value = True, "1", "2"), txtCheque, 0, "0", "", 0, 0, False, False, , "ICA", "LMMD", , , , Mid(txtGlosa, 1, 30) & " " & CStr(Format(txtmonto, "#0.00")), gdFecSis, gsNomAge, gsCodUser, sLPT, False, False
'
'            ''If gbITFAplica And CCur(Me.LblITF.Caption) > 0 Then
'            ''    fgITFImprimeBoleta gsNomAge, 0, Me.Caption, 0, , , , , , False
'           '' End If
'
'        Loop While MsgBox("Desea Re Imprimir ?", vbQuestion + vbYesNo, "Aviso") = vbYes
'        Set oBol = Nothing
End Sub
Private Sub txtCtaOrig_EmiteDatos()
lblDescCtaOrig = txtCtaOrig.psDescripcion
CmdAceptar.SetFocus
End Sub


