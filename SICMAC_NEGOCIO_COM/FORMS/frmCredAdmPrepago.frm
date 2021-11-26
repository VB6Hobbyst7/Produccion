VERSION 5.00
Begin VB.Form frmCredAdmPrepago 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Activacion de Prepago - Hipotecario"
   ClientHeight    =   3765
   ClientLeft      =   3330
   ClientTop       =   2730
   ClientWidth     =   4830
   Icon            =   "frmCredAdmPrepago.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame4 
      Caption         =   "Generar Calendario Con :"
      Height          =   615
      Left            =   75
      TabIndex        =   16
      Top             =   2520
      Width           =   4695
      Begin VB.OptionButton OptTipoCalPre 
         Caption         =   "Menor Cuota Mismo Plazo"
         Height          =   240
         Index           =   1
         Left            =   2400
         TabIndex        =   18
         Top             =   255
         Width           =   2190
      End
      Begin VB.OptionButton OptTipoCalPre 
         Caption         =   "Misma Cuota Menor Plazo"
         Height          =   240
         Index           =   0
         Left            =   180
         TabIndex        =   17
         Top             =   240
         Value           =   -1  'True
         Width           =   2190
      End
   End
   Begin VB.Frame Frame3 
      Height          =   660
      Left            =   75
      TabIndex        =   12
      Top             =   3075
      Width           =   4695
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   345
         Left            =   3435
         TabIndex        =   15
         Top             =   210
         Width           =   1170
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         Enabled         =   0   'False
         Height          =   345
         Left            =   1305
         TabIndex        =   14
         Top             =   210
         Width           =   1170
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "&Aceptar"
         Enabled         =   0   'False
         Height          =   345
         Left            =   105
         TabIndex        =   13
         Top             =   210
         Width           =   1170
      End
   End
   Begin VB.Frame Frame2 
      Height          =   660
      Left            =   75
      TabIndex        =   10
      Top             =   1845
      Width           =   4695
      Begin VB.CheckBox ChkActivaPP 
         Caption         =   "Activacion de Prepago para Mi Vivienda"
         Enabled         =   0   'False
         Height          =   225
         Left            =   750
         TabIndex        =   11
         Top             =   270
         Width           =   3165
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos"
      Height          =   1290
      Left            =   75
      TabIndex        =   0
      Top             =   555
      Width           =   4695
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Titular :"
         Height          =   195
         Left            =   90
         TabIndex        =   8
         Top             =   285
         Width           =   525
      End
      Begin VB.Label lblTitular 
         Height          =   195
         Left            =   705
         TabIndex        =   7
         Top             =   285
         Width           =   3795
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Analista :"
         Height          =   195
         Left            =   90
         TabIndex        =   6
         Top             =   540
         Width           =   645
      End
      Begin VB.Label LblAnalista 
         Height          =   195
         Left            =   855
         TabIndex        =   5
         Top             =   540
         Width           =   3645
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Prestamo :"
         Height          =   195
         Left            =   105
         TabIndex        =   4
         Top             =   885
         Width           =   750
      End
      Begin VB.Label LblPrestamo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   945
         TabIndex        =   3
         Top             =   885
         Width           =   900
      End
      Begin VB.Label Saldo 
         AutoSize        =   -1  'True
         Caption         =   "Saldo :"
         Height          =   195
         Left            =   2070
         TabIndex        =   2
         Top             =   900
         Width           =   495
      End
      Begin VB.Label LblSaldo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   2655
         TabIndex        =   1
         Top             =   885
         Width           =   915
      End
   End
   Begin SICMACT.ActXCodCta ActxCta 
      Height          =   480
      Left            =   75
      TabIndex        =   9
      Top             =   150
      Width           =   3705
      _ExtentX        =   6535
      _ExtentY        =   847
      Texto           =   "Credito :"
      EnabledCMAC     =   -1  'True
      EnabledCta      =   -1  'True
      EnabledProd     =   -1  'True
      EnabledAge      =   -1  'True
      CMAC            =   "112"
   End
End
Attribute VB_Name = "frmCredAdmPrepago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objPista As COMManejador.Pista


Private Sub ActxCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call CargaDatos(ActxCta.NroCuenta)
    End If
End Sub

Private Sub CmdAceptar_Click()
Dim oCred As COMDCredito.DCOMCredActBD
    
    If MsgBox("Se va Actualizar los Datos, Desea Continuar?", vbInformation + vbYesNo, "Aviso") = vbNo Then
        Exit Sub
    End If
    Set oCred = New COMDCredito.DCOMCredActBD
    Call oCred.dUpdateColocacCred(ActxCta.NroCuenta, , , , , , , , , , , , , , , , , , , , , , , ChkActivaPP.value, IIf(OptTipoCalPre(0).value, 1, 2))
        
    ''*** PEAC 20090219
    objPista.InsertarPista gsOpeCod, GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, gModificar, "Activacion de Prepago Hipotecario.", ActxCta.NroCuenta, gCodigoCuenta
    
    Set oCred = Nothing
    LimpiaPantalla
End Sub

Private Sub cmdCancelar_Click()
    Call LimpiaPantalla
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 And ActxCta.Enabled = True Then 'F12
        Dim sCuenta As String
        sCuenta = frmValTarCodAnt.Inicia(gColPYMEEmp, False)
        If sCuenta <> "" Then
            ActxCta.NroCuenta = sCuenta
            ActxCta.SetFocusCuenta
        End If
    End If
End Sub

Private Sub Form_Load()
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
    CentraForm Me
    ActxCta.CMAC = gsCodCMAC
    ActxCta.Age = gsCodAge
    
    Set objPista = New COMManejador.Pista
    gsOpeCod = gCredAdminPrepagosHipoteca
    
End Sub

Private Sub HabilitaControles(ByVal pbHab As Boolean)
    ChkActivaPP.Enabled = pbHab
    CmdAceptar.Enabled = pbHab
    CmdCancelar.Enabled = pbHab
    CmdSalir.Enabled = Not pbHab
    ActxCta.Enabled = Not pbHab
End Sub

Private Sub LimpiaPantalla()
    ActxCta.NroCuenta = ""
    ActxCta.CMAC = gsCodCMAC
    ActxCta.Age = gsCodAge
    lblTitular.Caption = ""
    LblAnalista.Caption = ""
    LblPrestamo.Caption = "0.00"
    LblSaldo.Caption = "0.00"
    ChkActivaPP.value = 0
    HabilitaControles False
End Sub

Private Sub CargaDatos(ByVal psCtaCod As String)
Dim oCredito As COMDCredito.DCOMCredito
Dim R2 As ADODB.Recordset
    Set oCredito = New COMDCredito.DCOMCredito
    Set R2 = oCredito.RecuperaDatosComunes(psCtaCod, False)
    If R2.EOF Or R2.BOF Then
        MsgBox "Credito No Existe, o No Esta Vigente, o No es un Credito Hipotecario", vbInformation, "Aviso"
        LimpiaPantalla
        Exit Sub
    End If
    'If R2.RecordCount > 0 And CInt(Mid(psCtaCod, 6, 1)) = CInt(Mid(Trim(str(gColHipoMiVivienda)), 1, 1)) Then
    If R2.RecordCount > 0 And CInt(Mid(R2!cTpoCredCod, 1, 1)) = CInt(Mid(Trim(str(gColCredHipot)), 1, 1)) Then
        lblTitular.Caption = PstaNombre(R2!cTitular)
        LblAnalista.Caption = PstaNombre(R2!cAnalista)
        LblSaldo.Caption = Format(R2!nSaldo, "#0.00")
        LblPrestamo.Caption = Format(R2!nMontoCol, "#0.00")
        ChkActivaPP.value = IIf(IIf(IsNull(R2!bPrepago), 0, R2!bPrepago), 1, 0)
        'OptCalDimTipo(R2!nCalendDinamTipo - 1).Value = True
        HabilitaControles True
        If R2!nCalendDinamTipo = 1 Then
            OptTipoCalPre(0).value = True
        Else
            OptTipoCalPre(1).value = True
        End If
    Else
        MsgBox "Credito No Existe, o No Esta Vigente, o No es un Credito Hipotecario", vbInformation, "Aviso"
        LimpiaPantalla
    End If
    R2.Close
    Set oCredito = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objPista = Nothing
End Sub
