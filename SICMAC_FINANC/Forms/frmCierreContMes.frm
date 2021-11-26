VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCierreContMes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cierre Mesual de Contabilidad"
   ClientHeight    =   5070
   ClientLeft      =   3360
   ClientTop       =   2550
   ClientWidth     =   4845
   Icon            =   "frmCierreContMes.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   4845
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdValidaBalance 
      Caption         =   "Generación de &Balances"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   360
      TabIndex        =   9
      Top             =   3420
      Width           =   4125
   End
   Begin VB.CommandButton cmdValidaAsiento 
      Caption         =   "Validación de &Asientos y Saldos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   360
      TabIndex        =   8
      Top             =   2970
      Width           =   4125
   End
   Begin MSMask.MaskEdBox txtFechacierre 
      Height          =   330
      Left            =   2370
      TabIndex        =   6
      Top             =   2415
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton cmdsalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   360
      TabIndex        =   4
      Top             =   4320
      Width           =   4125
   End
   Begin VB.CommandButton cmdEjecutar 
      Caption         =   "&Ejecutar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   360
      TabIndex        =   3
      Top             =   3870
      Width           =   4125
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "ADVERTENCIA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   1470
      TabIndex        =   7
      Top             =   210
      Width           =   1905
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Fecha de Cierre :"
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
      Left            =   795
      TabIndex        =   5
      Top             =   2460
      Width           =   1485
   End
   Begin VB.Label Label2 
      Caption         =   $"frmCierreContMes.frx":030A
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1260
      Left            =   300
      TabIndex        =   2
      Top             =   645
      Width           =   4530
   End
   Begin VB.Image imgAlerta 
      Height          =   480
      Left            =   285
      Picture         =   "frmCierreContMes.frx":03D0
      Top             =   105
      Width           =   480
   End
   Begin VB.Label lblFechaUltcierre 
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H00C00000&
      Height          =   330
      Left            =   2370
      TabIndex        =   1
      Top             =   2010
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Ultimo Cierre :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   810
      TabIndex        =   0
      Top             =   2070
      Width           =   1215
   End
End
Attribute VB_Name = "frmCierreContMes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim R As ADODB.Recordset
Dim lbValidaOk  As Boolean
Dim lbBalanceOk As Boolean
Dim dFecIni As Date
Dim dFecFin As Date

'ARLO20170208****
Dim objPista As COMManejador.Pista
'************

Private Sub cmdEjecutar_Click()
On Error GoTo ErrorCierreCont
If Not lbValidaOk Then
   If MsgBox("Validación de Asientos no realizado o presenta observaciones." & Chr(10) & " ¿ Desea continuar ? ", vbQuestion + vbYesNo, "¡Aviso!") = vbNo Then
      Exit Sub
   End If
End If
If Not lbBalanceOk Then
   MsgBox "Es necesario generar los Balances antes del Cierre de Mes", vbInformation, "¡Aviso!"
   Exit Sub
End If

If Valida() Then
   dFecIni = CDate("01/" & Mid(txtFechacierre, 4, 7))
   dFecFin = CDate(txtFechacierre)
   
   If MsgBox("¿ Desea Realizar Cierre Mensual de Contabilidad ?", vbYesNo + vbQuestion, "Aviso") = vbYes Then
      Me.Enabled = False
      MousePointer = 11
      Dim oSdo As New NCtasaldo
      oSdo.CierreContableMensual Format(dFecIni, gsFormatoMovFecha), Format(dFecFin, gsFormatoMovFecha), gsCodUser, Format(GetFechaHoraServer, gsFormatoFechaHora)
      
      Me.Enabled = True
      MousePointer = 0
      MsgBox "Cierre de mes Realizado Satisfactoriamente", vbInformation, "Aviso"
            'ARLO20170208
            Set objPista = New COMManejador.Pista
            gsOpeCod = LogPistaCierreDiarioCont
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", "Cierre de mes Realizado Satisfactoriamente con Fecha de Ultimo Cierre " & dFecIni & " y Fecha de Cierre : " & dFecFin
            Set objPista = Nothing
            '*******
      Unload Me
   End If
End If
Exit Sub
ErrorCierreCont:
    MsgBox TextErr(Err.Description) & Chr(13) & "Consulte al Area de Sistemas", vbInformation, "Aviso"
    Enabled = True
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub cmdValidaAsiento_Click()
Dim oBalance As New NBalanceCont
Dim sTexto  As String
Dim lnTipoCambio As Currency
Dim oTC As New nTipoCambio
On Error GoTo ErrorValida
lbValidaOk = False
If Valida() Then
   dFecIni = CDate("01/" & Mid(txtFechacierre, 4, 7))
   dFecFin = CDate(txtFechacierre)
   lnTipoCambio = oTC.EmiteTipoCambio(dFecFin, TCFijoMes)
   sTexto = ""
   sTexto = sTexto & oBalance.ImprimeValidaBalance(Format(dFecIni, gsFormatoMovFecha), Format(dFecFin, gsFormatoMovFecha), gnLinPage, gValidaCuadreAsiento, , True)
   sTexto = sTexto & oBalance.ImprimeValidaBalance(Format(dFecIni, gsFormatoMovFecha), Format(dFecFin, gsFormatoMovFecha), gnLinPage, gValidaConvesionME, , True, lnTipoCambio)
   sTexto = sTexto & oBalance.ImprimeValidaBalance(Format(dFecIni, gsFormatoMovFecha), Format(dFecFin, gsFormatoMovFecha), gnLinPage, gValidaCuentasNoExistentes, , True)
   sTexto = sTexto & oBalance.ImprimeValidaBalance(Format(dFecIni, gsFormatoMovFecha), Format(dFecFin, gsFormatoMovFecha), gnLinPage, gValidaCuentasNoExistentes2, , True)
   sTexto = sTexto & oBalance.ImprimeValidaBalance(Format(dFecIni, gsFormatoMovFecha), Format(dFecFin, gsFormatoMovFecha), gnLinPage, gValidaCuentasAnaliticas, , True)
   sTexto = sTexto & oBalance.ImprimeValidaBalance(Format(dFecIni, gsFormatoMovFecha), Format(dFecFin, gsFormatoMovFecha), gnLinPage, gValidaCuentasDeOrden, , True)
   sTexto = sTexto & oBalance.ImprimeValidaBalance(Format(dFecIni, gsFormatoFecha), Format(dFecFin, gsFormatoFecha), gnLinPage, gValidaSaldosContables, , True)
   sTexto = sTexto & oBalance.ImprimeValidaBalance(Format(dFecIni, gsFormatoFecha), Format(dFecFin, gsFormatoFecha), gnLinPage, gValidaCuentasSinPadre, , True)
   If sTexto = "" Then
      MsgBox "Asientos y Saldos registrados Correctamente", vbInformation, "¡Aviso!"
      lbValidaOk = True
            'ARLO20170208
            Set objPista = New COMManejador.Pista
            gsOpeCod = LogPistaCierreDiarioCont
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "", "Validaion de Asientos y Saldos Mensual con Fecha de Ultimo Cierre " & dFecIni & " y Fecha de Cierre : " & dFecFin
            Set objPista = Nothing
            '*******
      Exit Sub
   Else
      EnviaPrevio sTexto, "VALIDACION DE ASIENTOS Y SALDOS", gnLinPage, False
   End If
End If
Set oBalance = Nothing
Set oTC = Nothing

If lbValidaOk And lbBalanceOk Then
   cmdEjecutar.Enabled = True
End If

Exit Sub
ErrorValida:
   MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
End Sub

Private Sub cmdValidaBalance_Click()
Dim nTipoMone  As Integer
Dim nTipoBala  As Integer
Dim nMoneda    As Integer
Dim dBalance   As dBalanceCont
Dim nTipCambio As Currency
Dim oBalance   As New NBalanceCont

nTipCambio = oBalance.GetTipCambioBalance(Format(dFecFin, gsFormatoMovFecha))
If nTipCambio = 0 Then
   nTipCambio = gnTipCambio
End If
Set oBalance = Nothing

lbBalanceOk = False
dFecIni = CDate("01/" & Mid(txtFechacierre, 4, 7))
dFecFin = CDate(txtFechacierre)
Set dBalance = New dBalanceCont
cmdValidaBalance.Enabled = False
For nTipoBala = 1 To 2
   For nTipoMone = 0 To 6
      MousePointer = 11
      If nTipoMone = 0 Then gsSimbolo = "": nMoneda = 0
      If nTipoMone = 1 Then gsSimbolo = gcMN: nMoneda = 1
      If nTipoMone = 2 Then gsSimbolo = gcME: nMoneda = 2
      If nTipoMone = 3 Then gsSimbolo = "": nMoneda = 3
      If nTipoMone = 4 Then gsSimbolo = "": nMoneda = 4
      If nTipoMone = 5 Then gsSimbolo = "": nMoneda = 6
      If nTipoMone = 6 Then gsSimbolo = "": nMoneda = 9
      DoEvents
      dBalance.EliminaBalance nTipoBala, nMoneda, Month(dFecIni), Year(dFecIni)
      DoEvents
      dBalance.EliminaBalanceTemp nTipoBala, nMoneda
      DoEvents
      dBalance.InsertaSaldosIniciales nTipoBala, nMoneda, Format(dFecIni, gsFormatoFecha)    ', True
      DoEvents
      dBalance.InsertaMovimientosMes nTipoBala, nMoneda, Format(dFecIni, gsFormatoMovFecha), Format(dFecFin, gsFormatoMovFecha), , False
      DoEvents
      dBalance.MayorizacionBalance nTipoBala, nMoneda, Month(dFecIni), Year(dFecIni), False, gsCodCMAC
      MousePointer = 0
   Next
Next
MsgBox "Balances Generados Satisfactoriamente", vbInformation, "¡Aviso!"
cmdValidaBalance.Enabled = True
Set dBalance = Nothing
lbBalanceOk = True

            'ARLO20170208
            Set objPista = New COMManejador.Pista
            gsOpeCod = LogPistaCierreDiarioCont
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "", "Generacion de Balance Mensual con Fecha de Ultimo Cierre " & dFecIni & " y Fecha de Cierre : " & dFecFin
            Set objPista = Nothing
            '*******

If lbValidaOk And lbBalanceOk Then
   cmdEjecutar.Enabled = True
End If

End Sub

Private Sub Form_Load()
Dim oGen As New DGeneral
CentraForm Me
lblFechaUltcierre = LeeConstanteSist(gConstSistCierreMensualCont)
txtFechacierre = gdFecSis
Set oGen = Nothing
End Sub

Private Function Valida() As Boolean
Valida = False
If ValidaFecha(txtFechacierre) <> "" Then
    Exit Function
End If
If CDate(txtFechacierre) < CDate(lblFechaUltcierre) Then
    MsgBox "Fecha Ingresada menor que la del Ultimo Cierre Realizado", vbInformation, "Aviso"
    Valida = False
    Exit Function
End If
Valida = True
End Function

Private Sub txtFechacierre_GotFocus()
fEnfoque Me.txtFechacierre
End Sub

Private Sub txtFechacierre_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If ValidaFecha(txtFechacierre) <> "" Then
       MsgBox "¡ Fecha no válida ! ", vbInformation, "Aviso"
       txtFechacierre.SetFocus
    Else
       cmdValidaAsiento.SetFocus
    End If
End If
End Sub
