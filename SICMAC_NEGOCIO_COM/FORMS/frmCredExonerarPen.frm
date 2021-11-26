VERSION 5.00
Begin VB.Form frmCredExonerarPen 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Exonerar Penalidad de Cancelacion de Credito"
   ClientHeight    =   5385
   ClientLeft      =   2505
   ClientTop       =   1890
   ClientWidth     =   7200
   Icon            =   "frmCredExonerarPen.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   2025
      Left            =   60
      TabIndex        =   24
      Top             =   2760
      Width           =   7020
      Begin VB.CommandButton CmdIzq 
         Caption         =   "<<"
         Height          =   375
         Left            =   3315
         TabIndex        =   28
         Top             =   900
         Width           =   510
      End
      Begin VB.CommandButton CmdDer 
         Caption         =   ">>"
         Height          =   375
         Left            =   3315
         TabIndex        =   27
         Top             =   450
         Width           =   510
      End
      Begin VB.ListBox LstGastosExo 
         Height          =   1425
         Left            =   3915
         TabIndex        =   26
         Top             =   420
         Width           =   2970
      End
      Begin VB.ListBox LstGastos 
         Height          =   1425
         Left            =   105
         TabIndex        =   25
         Top             =   450
         Width           =   3090
      End
      Begin VB.Label Label5 
         Caption         =   "Gastos Exonerados"
         Height          =   240
         Left            =   3930
         TabIndex        =   30
         Top             =   165
         Width           =   1470
      End
      Begin VB.Label Label1 
         Caption         =   "Gastos"
         Height          =   240
         Left            =   150
         TabIndex        =   29
         Top             =   165
         Width           =   1335
      End
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   3645
      TabIndex        =   23
      Top             =   4935
      Width           =   1275
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   3645
      TabIndex        =   22
      Top             =   4935
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2340
      TabIndex        =   21
      Top             =   4935
      Width           =   1275
   End
   Begin VB.Frame Frame3 
      Caption         =   "Credito"
      Height          =   1245
      Left            =   75
      TabIndex        =   16
      Top             =   0
      Width           =   6990
      Begin VB.Frame FraListaCred 
         Caption         =   "&Lista Creditos"
         Height          =   960
         Left            =   5025
         TabIndex        =   19
         Top             =   180
         Width           =   1875
         Begin VB.ListBox LstCred 
            Height          =   645
            ItemData        =   "frmCredExonerarPen.frx":030A
            Left            =   75
            List            =   "frmCredExonerarPen.frx":030C
            TabIndex        =   20
            Top             =   225
            Width           =   1725
         End
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "&Buscar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3915
         TabIndex        =   18
         Top             =   540
         Width           =   1035
      End
      Begin SICMACT.ActXCodCta ActxCta 
         Height          =   435
         Left            =   195
         TabIndex        =   17
         Top             =   495
         Width           =   3660
         _ExtentX        =   6456
         _ExtentY        =   767
         Texto           =   "Credito :"
         EnabledCMAC     =   -1  'True
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1455
      Left            =   60
      TabIndex        =   3
      Top             =   1275
      Width           =   7020
      Begin VB.Label LblMonCred 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1905
         TabIndex        =   15
         Top             =   255
         Width           =   1020
      End
      Begin VB.Label lblmonto 
         AutoSize        =   -1  'True
         Caption         =   "Prestamo "
         Height          =   195
         Left            =   645
         TabIndex        =   14
         Top             =   285
         Width           =   705
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Saldo Capital"
         Height          =   195
         Left            =   3810
         TabIndex        =   13
         Top             =   285
         Width           =   930
      End
      Begin VB.Label LblSalCap 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4980
         TabIndex        =   12
         Top             =   255
         Width           =   1020
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Interes"
         Height          =   195
         Left            =   3810
         TabIndex        =   11
         Top             =   585
         Width           =   480
      End
      Begin VB.Label LblIntCred 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4980
         TabIndex        =   10
         Top             =   555
         Width           =   1020
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Analista  "
         Height          =   195
         Left            =   645
         TabIndex        =   9
         Top             =   1125
         Width           =   645
      End
      Begin VB.Label lblAnalista 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1905
         TabIndex        =   8
         Top             =   1095
         Width           =   4125
      End
      Begin VB.Label lblProximaCuota 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1905
         TabIndex        =   7
         Top             =   555
         Width           =   1020
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Prox Cuota"
         Height          =   195
         Left            =   645
         TabIndex        =   6
         Top             =   585
         Width           =   780
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Nota    "
         Height          =   195
         Left            =   645
         TabIndex        =   5
         Top             =   855
         Width           =   525
      End
      Begin VB.Label lblNota 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1905
         TabIndex        =   4
         Top             =   825
         Width           =   1020
      End
   End
   Begin VB.Frame fraNuevoMet 
      Height          =   825
      Left            =   60
      TabIndex        =   0
      Top             =   2760
      Visible         =   0   'False
      Width           =   7020
      Begin VB.CheckBox chkCobrarPenalidad 
         Height          =   285
         Left            =   3780
         TabIndex        =   1
         Top             =   300
         Width           =   375
      End
      Begin VB.Label Label9 
         Caption         =   "Exonerar de  Cobro de Penalidad por Pago Adelantado ?"
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
         Height          =   405
         Left            =   255
         TabIndex        =   2
         Top             =   255
         Width           =   3135
      End
   End
End
Attribute VB_Name = "frmCredExonerarPen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objPista As COMManejador.Pista


Private Sub LimpiaPantalla()
    LimpiaControles Me
    chkCobrarPenalidad.value = 0
    ActxCta.NroCuenta = ""
    ActxCta.CMAC = gsCodCMAC
    ActxCta.Age = gsCodAge
    LstCred.Clear
    LstGastos.Clear
    LstGastosExo.Clear
End Sub

Private Sub HabilitaActualizar(ByVal pbHabilita As Boolean)
    fraNuevoMet.Enabled = pbHabilita
    Frame3.Enabled = Not pbHabilita
    CmdAceptar.Enabled = pbHabilita
    cmdCancelar.Visible = pbHabilita
    CmdSalir.Visible = Not pbHabilita
End Sub

Private Function CargaDatos() As Boolean
'Dim oCredito As COMDCredito.DCOMCredito
Dim R As ADODB.Recordset
Dim RGastosExon As ADODB.Recordset
Dim RGastos As ADODB.Recordset
Dim oGastos As COMDCredito.DCOMGasto

    On Error GoTo ErrorCargaDatos
    
    Set oGastos = New COMDCredito.DCOMGasto
    Call oGastos.CargarDatosPenalidad(ActxCta.NroCuenta, R, RGastos, RGastosExon)
    Set oGastos = Nothing
    
    'Set oCredito = New COMDCredito.DCOMCredito
    'Set R = oCredito.RecuperaDatosComunes(ActxCta.NroCuenta, False)
    'Set RGastosExon = oCredito.RecuperaGastosExonerados(ActxCta.NroCuenta)
    'Set oCredito = Nothing
    
    'Set oGastos = New COMDCredito.DCOMGasto
    'Set RGastos = oGastos.RecuperaGastosAplicablesCuotas(CInt(Mid(ActxCta.NroCuenta, 9, 1)), "PA")
    'Set oGastos = Nothing
    
    If Not R.BOF And Not R.EOF Then
        CargaDatos = True
        LblMonCred.Caption = Format(R!nMontoCol, "#0.00")
        lblProximaCuota.Caption = IIf(IsNull(R!nNroProxCuota), 0, R!nNroProxCuota)
        LblSalCap.Caption = Format(IIf(IsNull(R!nSaldo), 0, R!nSaldo), "#0.00")
        LblIntCred.Caption = Format(IIf(IsNull(R!nTasaInteres), 0, R!nTasaInteres), "#0.00")
        lblAnalista.Caption = IIf(IsNull(R!cAnalista), "", R!cAnalista)
        lblNota.Caption = IIf(IsNull(R!nNota), "", R!nNota)
        'chkCobrarPenalidad.Value = IIf(IsNull(R!nExoPenalidad), 0, R!nExoPenalidad)
        chkCobrarPenalidad.value = 0
        
        LstGastos.Clear
        LstGastosExo.Clear
        
        Do While Not RGastos.EOF
            RGastosExon.Find "nPrdConceptoCod = " & RGastos!nPrdConceptoCod, , adSearchForward, 1
            If RGastosExon.EOF Then
                LstGastos.AddItem Trim(RGastos!cdescripcion) & Space(150) & RGastos!nPrdConceptoCod
                LstGastos.ToolTipText = Trim(RGastos!cdescripcion)
            Else
               LstGastosExo.AddItem Trim(RGastos!cdescripcion) & Space(150) & RGastos!nPrdConceptoCod
               LstGastosExo.ToolTipText = Trim(RGastos!cdescripcion)
            End If
            RGastos.MoveNext
        Loop
    Else
        CargaDatos = False
    End If
    R.Close
    Set R = Nothing
    RGastos.Close
    Set RGastos = Nothing
    RGastosExon.Close
    Set RGastosExon = Nothing
    Exit Function

ErrorCargaDatos:
        MsgBox Err.Description, vbCritical, "Aviso"

End Function


Private Sub CmdAceptar_Click()
Dim oCredito As COMDCredito.DCOMCredActBD
Dim i As Integer
Dim MatGastos() As String

    If MsgBox("Se va ha Grabar la Exoneracion de Penalidad, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then
        Exit Sub
    End If
    
    Set oCredito = New COMDCredito.DCOMCredActBD
    'Call oCredito.dUpdateColocacCred(ActxCta.NroCuenta, , , , , , , , , , IIf(chkCobrarPenalidad.Value = 0, 0, 1), , , , , False)
    'oCredito.dBeginTrans
    'Call oCredito.EliminaGastosExonerados(ActxCta.NroCuenta)
    'For i = 0 To Me.LstGastosExo.ListCount - 1
    '    Call oCredito.dInsertColocCredGastosExon(ActxCta.NroCuenta, CLng(Trim(Right(LstGastosExo.List(i), 10))))
    'Next i
    'oCredito.dCommitTrans
    ReDim MatGastos(Me.LstGastosExo.ListCount)
    For i = 0 To Me.LstGastosExo.ListCount - 1
        MatGastos(i) = Trim(Right(LstGastosExo.List(i), 10))
    Next i
    Call oCredito.GrabarGastosExonerados(ActxCta.NroCuenta, MatGastos)
    
    ''*** PEAC 20090126
    objPista.InsertarPista gsOpeCod, GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, gInsertar, , ActxCta.NroCuenta, gCodigoCuenta
    
    Set oCredito = Nothing
    HabilitaActualizar False
    Call LimpiaPantalla
    
End Sub

Private Sub cmdBuscar_Click()
Dim oCredito As COMDCredito.DCOMCredito
Dim R As ADODB.Recordset
Dim oPers As COMDPersona.UCOMPersona

    On Error GoTo ErrorCmdBuscar_Click
    LstCred.Clear
    Set oPers = frmBuscaPersona.Inicio()
    If Not oPers Is Nothing Then
        Set oCredito = New COMDCredito.DCOMCredito
        Set R = oCredito.RecuperaCreditosVigentes(oPers.sPersCod, , Array(gColocEstVigMor, gColocEstVigNorm, gColocEstVigNorm))
        Do While Not R.EOF
            LstCred.AddItem R!cCtaCod
            R.MoveNext
        Loop
        R.Close
        Set R = Nothing
        Set oCredito = Nothing
    End If
    If LstCred.ListCount = 0 Then
        MsgBox "El Cliente No Tiene Creditos Vigentes", vbInformation, "Aviso"
    End If
    
    Exit Sub

ErrorCmdBuscar_Click:
    MsgBox Err.Description, vbCritical, "Aviso"

End Sub

Private Sub ActxCta_KeyPress(KeyAscii As Integer)
On Error GoTo ErrorActxCta_KeyPress
    If KeyAscii = 13 Then
        If Not CargaDatos() Then
            HabilitaActualizar False
            MsgBox "No se pudo encontrar el Credito, o el Credito No esta Vigente", vbInformation, "Aviso"
        Else
            HabilitaActualizar True
        End If
    End If
    Exit Sub

ErrorActxCta_KeyPress:
        MsgBox Err.Description, vbCritical, "Aviso"
End Sub

Private Sub cmdCancelar_Click()
     HabilitaActualizar False
    Call LimpiaPantalla
End Sub

Private Sub CmdDer_Click()
    If LstGastos.ListCount > 0 Then
        LstGastosExo.AddItem LstGastos.Text
        LstGastos.RemoveItem LstGastos.ListIndex
    Else
        MsgBox "No Existen Gastos a Exonerar", vbInformation, "Mensaje"
    End If
End Sub

Private Sub CmdIzq_Click()
    If LstGastosExo.ListCount > 0 Then
        LstGastos.AddItem LstGastosExo.Text
        LstGastosExo.RemoveItem LstGastosExo.ListIndex
    Else
        MsgBox "No Existen Gastos ", vbInformation, "Mensaje"
    End If
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
    CentraForm Me
    ActxCta.NroCuenta = ""
    ActxCta.CMAC = gsCodCMAC
    ActxCta.Age = gsCodAge
    Set objPista = New COMManejador.Pista
    gsOpeCod = gCredExonerarPenalidadCancel
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objPista = Nothing
End Sub

Private Sub LstCred_Click()
    If LstCred.ListCount > 0 And LstCred.ListIndex <> -1 Then
        ActxCta.NroCuenta = LstCred.Text
    End If
End Sub
