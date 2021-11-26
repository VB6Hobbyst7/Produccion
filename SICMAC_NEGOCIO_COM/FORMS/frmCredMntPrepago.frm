VERSION 5.00
Begin VB.Form frmCredMntPrepago 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Prepagos"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4860
   Icon            =   "frmCredMntPrepago.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   4860
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Tipo Cuota:"
      Height          =   975
      Left            =   75
      TabIndex        =   17
      Top             =   2340
      Width           =   4695
      Begin VB.OptionButton OptTipoCuota 
         Caption         =   "Reducción del número de cuotas"
         Height          =   240
         Index           =   1
         Left            =   180
         TabIndex        =   19
         Top             =   600
         Width           =   2655
      End
      Begin VB.OptionButton OptTipoCuota 
         Caption         =   "Reducción del monto de las cuotas"
         Height          =   240
         Index           =   0
         Left            =   180
         TabIndex        =   18
         Top             =   240
         Value           =   -1  'True
         Width           =   2910
      End
      Begin VB.Label lblColocCalendCod 
         Height          =   255
         Left            =   3480
         TabIndex        =   21
         Top             =   240
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblCuotasPend 
         Caption         =   "0"
         Height          =   255
         Left            =   3240
         TabIndex        =   20
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos"
      Height          =   1290
      Left            =   75
      TabIndex        =   7
      Top             =   405
      Width           =   4695
      Begin VB.Label lblSaldo 
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
         TabIndex        =   15
         Top             =   885
         Width           =   915
      End
      Begin VB.Label Saldo 
         AutoSize        =   -1  'True
         Caption         =   "Saldo :"
         Height          =   195
         Left            =   2070
         TabIndex        =   14
         Top             =   900
         Width           =   495
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
         TabIndex        =   13
         Top             =   885
         Width           =   900
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Prestamo :"
         Height          =   195
         Left            =   105
         TabIndex        =   12
         Top             =   885
         Width           =   750
      End
      Begin VB.Label LblAnalista 
         Height          =   195
         Left            =   855
         TabIndex        =   11
         Top             =   540
         Width           =   3645
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Analista :"
         Height          =   195
         Left            =   90
         TabIndex        =   10
         Top             =   540
         Width           =   645
      End
      Begin VB.Label lblTitular 
         Height          =   195
         Left            =   705
         TabIndex        =   9
         Top             =   285
         Width           =   3795
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Titular :"
         Height          =   195
         Left            =   90
         TabIndex        =   8
         Top             =   285
         Width           =   525
      End
   End
   Begin VB.Frame Frame3 
      Height          =   660
      Left            =   75
      TabIndex        =   3
      Top             =   3300
      Width           =   4695
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "&Aceptar"
         Enabled         =   0   'False
         Height          =   345
         Left            =   105
         TabIndex        =   6
         Top             =   210
         Width           =   1170
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         Enabled         =   0   'False
         Height          =   345
         Left            =   1305
         TabIndex        =   5
         Top             =   210
         Width           =   1170
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   345
         Left            =   3435
         TabIndex        =   4
         Top             =   210
         Width           =   1170
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Activar tipo de Calendario"
      Height          =   615
      Left            =   75
      TabIndex        =   0
      Top             =   1725
      Width           =   4695
      Begin VB.OptionButton OptTipoCalPre 
         Caption         =   "Dinamico"
         Height          =   240
         Index           =   0
         Left            =   180
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   2190
      End
      Begin VB.OptionButton OptTipoCalPre 
         Caption         =   "Normal"
         Height          =   240
         Index           =   1
         Left            =   2400
         TabIndex        =   1
         Top             =   255
         Width           =   2190
      End
   End
   Begin SICMACT.ActXCodCta ActxCta 
      Height          =   480
      Left            =   75
      TabIndex        =   16
      Top             =   0
      Width           =   3705
      _ExtentX        =   6535
      _ExtentY        =   847
      Texto           =   "Credito :"
      EnabledCMAC     =   -1  'True
      EnabledCta      =   -1  'True
      EnabledProd     =   -1  'True
      EnabledAge      =   -1  'True
      CMAC            =   "109"
   End
End
Attribute VB_Name = "frmCredMntPrepago"
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
Dim sMovNro As String

    If MsgBox("Se va Actualizar los Datos, Desea Continuar?", vbInformation + vbYesNo, "Aviso") = vbNo Then
        Exit Sub
    End If
    Set oCred = New COMDCredito.DCOMCredActBD
    Call oCred.dUpdateColocacCred(ActxCta.NroCuenta, , , , , , , , , , , , IIf(OptTipoCalPre(0).value, 1, 0))
     
    'MAVM 20100826 ***
    If OptTipoCalPre(0).value = True Then
        sMovNro = GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        oCred.dInsertCredMantPrepago ActxCta.NroCuenta, sMovNro, IIf(OptTipoCalPre(0).value, 1, 0), IIf(OptTipoCuota(0).value, 1, 0), 0, CInt(lblCuotasPend.Caption)
        If lblColocCalendCod.Caption = gColocCalendCodFFCF Or lblColocCalendCod.Caption = gColocCalendCodFFCFPG Then
            Call oCred.dUpdateColocacEstadoCD(ActxCta.NroCuenta)
        End If
    End If
    '***
    
    '*** PEAC 20090217
    objPista.InsertarPista gsOpeCod, sMovNro, gsCodPersUser, GetMaquinaUsuario, gModificar, "Cambio a dinamico o normal.", ActxCta.NroCuenta, gCodigoCuenta
    
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
        sCuenta = frmValTarCodAnt.inicia(gColPYMEEmp, False)
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
    gsOpeCod = gCredAdminPrepagosNormales
   
End Sub

Private Sub HabilitaControles(ByVal pbHab As Boolean)
    
    OptTipoCalPre(0).Enabled = pbHab
    OptTipoCalPre(1).Enabled = pbHab
    'MAVM 20100826 ***
    OptTipoCuota(0).Enabled = pbHab
    OptTipoCuota(1).Enabled = pbHab
    '***
    cmdAceptar.Enabled = pbHab
    CmdCancelar.Enabled = pbHab
    cmdSalir.Enabled = Not pbHab
    ActxCta.Enabled = Not pbHab
End Sub

Private Sub LimpiaPantalla()
    ActxCta.NroCuenta = ""
    ActxCta.CMAC = gsCodCMAC
    ActxCta.Age = gsCodAge
    lblTitular.Caption = ""
    lblanalista.Caption = ""
    lblPrestamo.Caption = "0.00"
    lblSaldo.Caption = "0.00"
    OptTipoCalPre(0).value = False
    OptTipoCalPre(1).value = False
    'MAVM 20100826 ***
    OptTipoCuota(0).value = False
    OptTipoCuota(1).value = False
    lblCuotasPend.Caption = 0
    lblColocCalendCod.Caption = ""
    '***
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
    'MAVM 20100823 ***
    'If R2.RecordCount > 0 And gdFecSis >= R2!dVenc_Cuota Then
    If R2.RecordCount > 0 And gdFecSis > R2!dVenc_Cuota Then 'JUEZ 20150108
        MsgBox "El Credito se encuentra con Cuota Vencida", vbInformation, "Aviso"
        LimpiaPantalla
        Exit Sub
    End If
    'MAVM 20100823 ***
    'By Capi 18112008 para que permita a hipotecarios con recursos propios
    'If R2.RecordCount > 0 And CInt(Mid(psCtaCod, 6, 1)) <> CInt(Mid(Trim(Str(gColHipoMiVivienda)), 1, 1)) Then
    'If R2.RecordCount > 0 And CInt(Mid(psCtaCod, 6, 3)) <> CInt(Trim(str(gColHipoMiVivienda))) Then
    'If R2.RecordCount > 0 And CInt(R2!cTpoCredCod) <> CInt(Trim(Str(gColProHipoteMiVivienda))) Then
    If R2.RecordCount > 0 Then 'JUEZ 20150505
        If CInt(R2!cTpoCredCod) = "853" Or CInt(R2!cTpoCredCod) = "854" Then 'JUEZ 20150505
            lblTitular.Caption = PstaNombre(R2!cTitular)
            lblanalista.Caption = PstaNombre(R2!cAnalista)
            lblSaldo.Caption = Format(R2!nSaldo, "#0.00")
            lblPrestamo.Caption = Format(R2!nMontoCol, "#0.00")
            'MAVM 20100908 ***
            lblCuotasPend.Caption = IIf(IsNull(R2!nCuotasPend), 0, R2!nCuotasPend)
            lblColocCalendCod.Caption = IIf(IsNull(R2!nColocCalendCod), 0, R2!nColocCalendCod)
            '***
            If R2!nCalendDinamico = 1 Then
                OptTipoCalPre(0).value = True
            Else
                'OptTipoCalPre(1).value = False
                OptTipoCalPre(1).value = True
            End If
            'OptCalDimTipo(R2!nCalendDinamTipo - 1).Value = True
            HabilitaControles True
        Else
            MsgBox "Credito No es MiVivienda ni Techo Propio", vbInformation, "Aviso"
            LimpiaPantalla
        End If
    Else
        'MsgBox "Credito No Existe, o No Esta Vigente, o No es un Credito Hipotecario", vbInformation, "Aviso"
        MsgBox "Credito No Existe, o No Esta Vigente", vbInformation, "Aviso"
        LimpiaPantalla
    End If
    R2.Close
    Set oCredito = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objPista = Nothing
End Sub

Private Sub OptTipoCalPre_Click(Index As Integer)
    If OptTipoCalPre(1).value = True Then
        OptTipoCuota(0).Enabled = False
        OptTipoCuota(1).Enabled = False
    Else
        OptTipoCuota(0).Enabled = True
        OptTipoCuota(1).Enabled = True
    End If
End Sub
