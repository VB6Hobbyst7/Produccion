VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCajeroPagoServ 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2985
   ClientLeft      =   3285
   ClientTop       =   2970
   ClientWidth     =   4170
   HelpContextID   =   400
   Icon            =   "frmCajeroPagoServ.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   4170
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   2925
      TabIndex        =   1
      Top             =   2520
      Width           =   1170
   End
   Begin VB.Frame FraServicios 
      Height          =   2430
      Left            =   105
      TabIndex        =   2
      Top             =   0
      Width           =   3975
      Begin VB.TextBox txtMovDesc 
         Height          =   705
         Left            =   105
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   1590
         Width           =   3735
      End
      Begin MSMask.MaskEdBox txtReferencia 
         Height          =   315
         Left            =   1680
         TabIndex        =   8
         Top             =   510
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtNroRecibo 
         Height          =   315
         Left            =   1680
         TabIndex        =   9
         Top             =   165
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   " "
      End
      Begin Sicmact.EditMoney txtImporte 
         Height          =   360
         Left            =   1680
         TabIndex        =   10
         Top             =   855
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
         Enabled         =   -1  'True
      End
      Begin VB.Label lblSimbolo1 
         AutoSize        =   -1  'True
         Caption         =   "S/."
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
         Left            =   3495
         TabIndex        =   11
         Top             =   960
         Width           =   285
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Glosa"
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
         Left            =   120
         TabIndex        =   7
         Top             =   1380
         Width           =   495
      End
      Begin VB.Label lblReferencia 
         AutoSize        =   -1  'True
         Caption         =   "Referencia"
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
         Left            =   105
         TabIndex        =   5
         Top             =   585
         Width           =   945
      End
      Begin VB.Label lblRecibo 
         AutoSize        =   -1  'True
         Caption         =   "Nro de Recibo:"
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
         Left            =   105
         TabIndex        =   4
         Top             =   225
         Width           =   1305
      End
      Begin VB.Label lblmontorecibo 
         AutoSize        =   -1  'True
         Caption         =   "Neto a Pagar"
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
         Left            =   105
         TabIndex        =   3
         Top             =   960
         Width           =   1140
      End
   End
   Begin VB.CommandButton CmdGuardar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   1755
      TabIndex        =   0
      Top             =   2520
      Width           =   1170
   End
End
Attribute VB_Name = "frmCajeroPagoServ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lsNumRecibo As String
Dim lsCodigoRec As String
Dim lnMontoRec As Currency
Dim oCajero As nCajero

Private Sub Limpiar()
    Me.txtImporte.value = 0
    txtNroRecibo.Mask = ""
    txtReferencia.Mask = ""
    txtNroRecibo.Text = ""
    txtReferencia.Text = ""
    Select Case gsOpeCod
        Case gOpeCajeroVarHIDRANDINAMN, gOpeCajeroVarHIDRANDINAME
            txtNroRecibo.Mask = "###-##-##-########"
            txtReferencia.Mask = "########"
            lblReferencia.Caption = "Num.Medidor"
        Case gOpeCajeroVarSEDALIBMN, gOpeCajeroVarSEDALIBME
            txtNroRecibo.Mask = "###-########-##"
            txtReferencia.Mask = "###########"
            lblReferencia.Caption = "Código"
        Case gOpeCajeroVarTELEFONICAMN, gOpeCajeroVarTELEFONICAME
            txtNroRecibo.Mask = "###############"
            txtReferencia.Mask = "###########"
            lblReferencia.Caption = "Código"
    End Select
    If txtNroRecibo.Visible And txtNroRecibo.Enabled Then txtNroRecibo.SetFocus
End Sub
Private Sub cmdGuardar_Click()
Dim lsMovNro As String
Dim oContF As NContFunciones
Set oContF = New NContFunciones

If Valida = False Then Exit Sub
   
If MsgBox("Desea Grabar Operación de Servicio??", vbYesNo + vbQuestion, "Aviso") = vbYes Then
    lsMovNro = oContF.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    If oCajero.GrabaCajeroPagoServicios(lsMovNro, gsOpeCod, Trim(txtMovDesc), CCur(txtImporte), txtNroRecibo, txtReferencia) = 0 Then
        Dim oImp As NContImprimir
        Dim lsTexto As String
        Dim lbReimp As Boolean
        Set oImp = New NContImprimir
        
        lbReimp = True
        Do While lbReimp
            oImp.ImprimeBoletaGeneral "OPERACIONES VARIAS - " + gsOpeDescPadre, gsOpeDescHijo, gsOpeCod, txtImporte, gsNomAge, lsMovNro, _
                     sLpt, , txtNroRecibo, txtReferencia, Trim(txtMovDesc)
        
            If MsgBox("Desea Reimprimir boleta de Operación", vbYesNo + vbQuestion, "Aviso") = vbNo Then
                lbReimp = False
            End If
        Loop
        Set oImp = Nothing
        Limpiar
    End If
End If
Set oContF = Nothing
Exit Sub
ERROR:
    MsgBox "Existe un error " & TextErr(Err.Description), vbCritical, Me.Caption
End Sub
Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set oCajero = Nothing
End Sub

Private Sub txtImporte_GotFocus()
    txtImporte.MarcaTexto
End Sub
Private Sub txtImporte_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    CmdGuardar.SetFocus
End If
End Sub
Private Sub Form_Load()
    Set oCajero = New nCajero
    txtNroRecibo.Text = ""
    txtReferencia = ""
    CentraForm Me
    Me.Caption = gsOpeDesc
    Limpiar
End Sub
Private Sub txtMovDesc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    CmdGuardar.SetFocus
End If
End Sub
Private Sub txtNroRecibo_GotFocus()
    fEnfoque txtNroRecibo
End Sub
Private Sub TxtNroRecibo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtReferencia.SetFocus
End If
End Sub

Private Sub txtReferencia_GotFocus()
    fEnfoque txtReferencia
End Sub
Private Sub txtReferencia_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        'If Right(txtReferencia, 4) = "    " Then txtReferencia = Left(txtReferencia, 4) & "       "
            If txtImporte.Enabled Then txtImporte.SetFocus
    End If
End Sub
Private Function DigChequeo(ByVal psRecibo As String, ByVal psCodigoCliente As String, ByVal pcMonto As Currency) As Integer
    Dim sFacSerNro As String, sFacNro As String, sDigCheq As String
    Dim sValor As String
    Dim iLongValor As Integer, iValTotal As Integer, N As Integer
    sFacSerNro = Mid(psRecibo, 1, 3)
    sFacNro = Mid(psRecibo, 4, 8)
    sDigCheq = Right(psRecibo, 2)
    sValor = psCodigoCliente & sFacSerNro & sFacNro & Right(String((10 - Len(Str(pcMonto * 100))), " ") & Str(pcMonto * 100), 8)
    iLongValor = Len(sValor)
    N = 0
    iValTotal = 0
    Do While N < iLongValor
        N = N + 1
        iValTotal = iValTotal + N * Val(Mid(sValor, N, 1))
    Loop
    DigChequeo = 97 - (iValTotal - Int(iValTotal / 97) * 97)
End Function
Private Function ValSedalib(ByVal psRecibo As String, ByVal psCodigoCliente As String, ByVal pcMonto As Currency) As Boolean
    If DigChequeo(psRecibo, psCodigoCliente, pcMonto) = Val(Right(psRecibo, 2)) Then
        ValSedalib = True
    Else
        ValSedalib = False
    End If
End Function
Function Valida() As Boolean
Valida = True
If txtImporte.value = 0 Then
    MsgBox "Monto de Operación no válido", vbInformation, "Aviso"
    Valida = False
    Me.txtImporte.SetFocus
    Exit Function
ElseIf Len(txtReferencia.Text) = 0 Then
    MsgBox "Debe indicar un el " & lblReferencia.Caption & " del Recibo", vbInformation, "Aviso"
    Valida = False
    txtReferencia.SetFocus
    Exit Function
    ElseIf Trim(Trim(txtNroRecibo.Text)) = "" Then
        MsgBox "El Recibo no puede ser Vacio", vbInformation, "Aviso"
        Valida = False
        txtNroRecibo.SetFocus
        Exit Function
    ElseIf oCajero.ValidaReciboServ(gsOpeCod, txtNroRecibo.Text, txtReferencia) Then
        MsgBox "El Recibo ya fue ingresado.", vbInformation, "Aviso"
        Valida = False
        txtReferencia.SetFocus
        Exit Function
    End If
    'SOLO PARA SEDALIB
If (gsOpeCod = gOpeCajeroVarSEDALIBMN Or gsOpeCod = gOpeCajeroVarSEDALIBME) Then
    'If ReadVarSis("AHO", "CtrValCobr") = "1" Then
    If Not ValSedalib(Left(txtNroRecibo, 3) & Mid(txtNroRecibo.Text, 5, 8) & Right(txtNroRecibo, 2), txtReferencia.Text, txtImporte.value) And (lsNumRecibo <> txtNroRecibo Or lsCodigoRec <> txtReferencia.Text Or lnMontoRec <> Me.txtImporte.value) Then
        lsNumRecibo = txtNroRecibo.Text
        lsCodigoRec = txtReferencia.Text
        lnMontoRec = txtImporte.value
        MsgBox "Códigos Errados.", vbInformation, "Aviso"
        Limpiar
        Valida = False
        Exit Function
    End If
    'End If
End If
If Len(Trim(txtMovDesc)) = 0 Then
    MsgBox "Glosa de Operación no ha sido Ingresada", vbInformation, "Aviso"
    txtMovDesc.SetFocus
    Valida = False
    Exit Function
End If
End Function
