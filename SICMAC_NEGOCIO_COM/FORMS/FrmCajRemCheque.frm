VERSION 5.00
Begin VB.Form FrmCajRemCheque 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Remesas con Cheque "
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   5565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   585
      Left            =   90
      TabIndex        =   15
      Top             =   4230
      Width           =   5385
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
         Left            =   1320
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
         Top             =   1920
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
Attribute VB_Name = "FrmCajRemCheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nMovNroRef As Long
Dim sCadenaRef As String

Private Sub CmdAceptar_Click()

    Dim onCredito As COMNCredito.NCOMCredito
    If txtCheque <> "" And txtAgencia <> "" And txtMonto <> "" And txtGlosa <> "" Then
        Set onCredito = New COMNCredito.NCOMCredito
           If onCredito.InsertaRemesaCheque(gsCodAge, gsCodUser, gdFecSis, Val(txtMonto.Text), txtCheque.Text, IIf(OptSoles.value = True, 1, 2), nMovNroRef, txtGlosa, "") = True Then
                MsgBox "Se registro correctamente"
                Call Imprimir
                Call cmdCancelar_Click
           Else
                MsgBox "Error en el registro"
                Call cmdCancelar_Click
           End If
    Else
        MsgBox "Los datos esta incompletos"
        Call cmdCancelar_Click
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
    sCadenaRef = ""
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Sub CargarDatos()
    Dim rs As ADODB.Recordset
    Dim oCredito As COMDCredito.DCOMCredito
    Dim cAgencia As String

    Set oCredito = New COMDCredito.DCOMCredito
    Set rs = oCredito.GetDatosCheque(Trim(txtCheque.Text))
    Set oCredito = Nothing

    If Not rs.EOF And Not rs.BOF Then
        nMovNroRef = rs!nMovNro
        txtNroCuenta = rs!cCtaIfCod
        txtAgencia = rs!cAgencia
        txtMonto = Format(rs!nMonto, "#0.00")
        If rs!nmoneda = "1" Then
            txtMonto.BackColor = vbWhite
            OptSoles.value = True
        Else
            txtMonto.BackColor = vbGreen
            OptDolares.value = True
        End If
        txtInstitucion = rs!cPersNombre
        cAgencia = rs!cAgeCod
        sCadenaRef = "01." & rs!cPersCod & "." & rs!cCtaIfCod
    End If
    Set rs = Nothing
    If cAgencia <> gsCodAge Then
        sCadenaRef = ""
        MsgBox "Este cheque no le corresponde a la agencia", vbInformation, "AVISO"
        CmdAceptar.Enabled = False
        'CmdCancelar.Enabled = False
    End If
End Sub

Private Sub Form_Load()
    sCadenaRef = ""
End Sub

Private Sub txtCheque_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdBuscar_Click
    End If
End Sub

Sub Imprimir()
Dim lsCaption
Dim oBol As COMNCaptaGenerales.NCOMCaptaImpresion
Dim lsBoleta As String
Dim nFicSal As Integer
lsCaption = "Remesa con cheque "

        Set oBol = New COMNCaptaGenerales.NCOMCaptaImpresion
            lsBoleta = oBol.ImprimeBoleta("OTRAS OPERACIONES", Left(lsCaption, 15), "00000000" & IIf(OptSoles.value = True, "1", "2"), CStr(Format(Val(txtMonto.Text), "#0.00")), gsNomAge, "00000000" & IIf(OptSoles.value = True, "1", "2"), txtCheque.Text, 0, "0", "", 0, 0, False, False, , "ICA", "LMMD", , , , Mid(txtGlosa, 1, 30) & " " & CStr(Format(Val(txtMonto.Text), "#0.00")), gdFecSis, gsNomAge, gsCodUser, sLpt, False, False)
        Set oBol = Nothing
        Do
            
            If Trim(lsBoleta) <> "" Then
                nFicSal = FreeFile
                Open sLpt For Output As nFicSal
                    Print #nFicSal, lsBoleta
                    Print #nFicSal, ""
                Close #nFicSal
            End If
            
            'If gbITFAplica And CCur(Me.LblITF.Caption) > 0 Then
            '    fgITFImprimeBoleta gsNomAge, 0, Me.Caption, 0, , , , , , False
            ' End If
            
        Loop While MsgBox("Desea Re Imprimir ?", vbQuestion + vbYesNo, "Aviso") = vbYes
        
End Sub

Private Sub txtGlosa_KeyPress(KeyAscii As Integer)
    KeyAscii = fgIntfMayusculas(KeyAscii)
End Sub
