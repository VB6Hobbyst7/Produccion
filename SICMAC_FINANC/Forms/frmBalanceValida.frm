VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmBalanceValida 
   Caption         =   "  PROCESO DE VALIDACIÓN DE BALANCE"
   ClientHeight    =   4815
   ClientLeft      =   2400
   ClientTop       =   1935
   ClientWidth     =   4935
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBalanceValida.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Caption         =   "Rango de Fechas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   765
      Left            =   150
      TabIndex        =   13
      Top             =   90
      Width           =   3540
      Begin MSMask.MaskEdBox txtFechaDel 
         Height          =   345
         Left            =   585
         TabIndex        =   0
         Top             =   300
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   609
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtFechaAl 
         Height          =   345
         Left            =   2205
         TabIndex        =   1
         Top             =   300
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   609
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "DEL"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   195
         TabIndex        =   15
         Top             =   360
         Width           =   300
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "AL"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1860
         TabIndex        =   14
         Top             =   375
         Width           =   210
      End
   End
   Begin VB.CommandButton cmdSalir 
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
      Height          =   420
      Left            =   3765
      TabIndex        =   11
      Top             =   4035
      Width           =   1140
   End
   Begin VB.CommandButton cmdProcesar 
      Caption         =   "&Procesar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3765
      TabIndex        =   10
      Top             =   3555
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      Caption         =   "Verificar ..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   3570
      Left            =   150
      TabIndex        =   12
      Top             =   900
      Width           =   3525
      Begin VB.CheckBox chkDecimales 
         Caption         =   "Asientos con mas de dos Decimales"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   60
         TabIndex        =   18
         Top             =   3150
         Width           =   3360
      End
      Begin VB.CheckBox chkCtasConsolidada 
         Caption         =   "Cuentas Contables sin Cta Consolidada"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   90
         TabIndex        =   17
         Top             =   2790
         Width           =   3360
      End
      Begin VB.CheckBox chkPapa 
         Caption         =   "Asientos sin Cuentas Padre"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   9
         Top             =   2475
         Width           =   3360
      End
      Begin VB.CheckBox chkAsientos 
         Caption         =   "Asientos Migrados Por Día"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   8
         Top             =   2175
         Width           =   3360
      End
      Begin VB.CheckBox chkSaldos 
         Caption         =   "Validación de Saldos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   7
         Top             =   1875
         Width           =   3360
      End
      Begin VB.CheckBox chkCuentaAsiento 
         Caption         =   "Asientos con Cuentas A&nalíticas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   5
         Top             =   1260
         Width           =   3360
      End
      Begin VB.CheckBox chkOrden 
         Caption         =   "C&uentas de Orden"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   6
         Top             =   1560
         Width           =   3360
      End
      Begin VB.CheckBox chkBala 
         Caption         =   "Asientos con Cuentas &No Existentes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   4
         Top             =   960
         Width           =   3360
      End
      Begin VB.CheckBox chkME 
         Caption         =   "&Conversión Moneda Extranjera"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   3
         Top             =   645
         Width           =   3360
      End
      Begin VB.CheckBox chkCuadra 
         Caption         =   "&Asientos Descuadrados"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   2
         Top             =   345
         Width           =   3360
      End
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   795
      Left            =   3840
      Picture         =   "frmBalanceValida.frx":08CA
      Stretch         =   -1  'True
      Top             =   150
      Width           =   780
   End
   Begin VB.Label lblmsg 
      AutoSize        =   -1  'True
      Caption         =   "Validando Balance ..."
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   150
      TabIndex        =   16
      Top             =   4530
      Visible         =   0   'False
      Width           =   1515
   End
End
Attribute VB_Name = "frmBalanceValida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Function ValidaDatos() As Boolean
ValidaDatos = False
   If Len(Trim(txtfechaDel)) <> 10 Then
      MsgBox "Falta indicar Fecha Inicial", vbInformation, "¡Aviso!"
      txtfechaDel.SetFocus
      Exit Function
   End If

   If Len(Trim(txtFechaAl)) <> 10 Then
      MsgBox "Falta indicar Fecha Final", vbInformation, "¡Aviso!"
      txtFechaAl.SetFocus
      Exit Function
   End If
ValidaDatos = True
End Function

 
Private Sub cmdProcesar_Click()
Dim nLin As Integer
Dim P    As Integer
Dim sTit As String
Dim sTexto As String
Dim lnTipoCambio As Currency

Dim I As Integer
Dim sFecha1 As String
Dim sFecha2 As String
Dim dFecha As Date
Dim nDias As Integer
Dim sTitulo As String
Dim lsCabecera As String
Dim sTemp As String
Dim sTemp2 As String

If Not ValidaDatos() Then
   Exit Sub
End If
nLin = gnLinPage
P = 0
If MsgBox(" ¿ Seguro de Iniciar Validación de Asientos ? ", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmación") = vbNo Then
   Exit Sub
End If
lblmsg.Visible = True
sTexto = ""
DoEvents
MousePointer = 11
Dim oTC As nTipoCambio
Set oTC = New nTipoCambio
lnTipoCambio = oTC.EmiteTipoCambio(txtfechaDel, TCFijoMes)
Set oTC = Nothing

DoEvents
Dim oBalance As New NBalanceCont

If chkCuadra.value = vbChecked Then
   lblmsg.Caption = "Validando asientos Descuadrados ... Espere un momento"
   DoEvents
   sTexto = sTexto & oBalance.ImprimeValidaBalance(Format(txtfechaDel, gsFormatoMovFecha), Format(txtFechaAl, gsFormatoMovFecha), gnLinPage, gValidaCuadreAsiento)
End If
If chkME.value = vbChecked Then
   lblmsg.Caption = "Validando asientos de Moneda Extranjera ... Espere un momento"
   DoEvents
   sTexto = sTexto & oBalance.ImprimeValidaBalance(Format(txtfechaDel, gsFormatoMovFecha), Format(txtFechaAl, gsFormatoMovFecha), gnLinPage, gValidaConvesionME, , , lnTipoCambio)
End If
If chkBala.value = vbChecked Then
   lblmsg.Caption = "Validando Cuentas Contables... Espere un momento"
   DoEvents
   sTexto = sTexto & oBalance.ImprimeValidaBalance(Format(txtfechaDel, gsFormatoMovFecha), Format(txtFechaAl, gsFormatoMovFecha), gnLinPage, gValidaCuentasNoExistentes)
   DoEvents
   sTexto = sTexto & oBalance.ImprimeValidaBalance(Format(txtfechaDel, gsFormatoMovFecha), Format(txtFechaAl, gsFormatoMovFecha), gnLinPage, gValidaCuentasNoExistentes2)
End If
If chkCuentaAsiento.value = vbChecked Then
   lblmsg.Caption = "Validando asientos con Cuentas Analíticas... Espere un momento"
   DoEvents
   sTexto = sTexto & oBalance.ImprimeValidaBalance(Format(txtfechaDel, gsFormatoMovFecha), Format(txtFechaAl, gsFormatoMovFecha), gnLinPage, gValidaCuentasAnaliticas)
End If
If chkOrden.value = vbChecked Then
   lblmsg.Caption = "Validando Cuentas de Orden por Agencia... Espere un momento"
   DoEvents
   sTexto = sTexto & oBalance.ImprimeValidaBalance(Format(txtfechaDel, gsFormatoMovFecha), Format(txtFechaAl, gsFormatoMovFecha), gnLinPage, gValidaCuentasDeOrden)
End If
If chkSaldos.value = vbChecked Then
   lblmsg.Caption = "Validando Saldos Contables... Espere un momento"
   DoEvents
   sTexto = sTexto & oBalance.ImprimeValidaBalance(Format(txtfechaDel, gsFormatoFecha), Format(txtFechaAl, gsFormatoFecha), gnLinPage, gValidaSaldosContables)
End If
 
If chkAsientos.value = vbChecked Then
    lblmsg.Caption = "Validando Migración de Asientos Contables... Espere un momento"
    
    sFecha1 = txtfechaDel.Text
    sFecha2 = txtFechaAl.Text
    
    nDias = DateDiff("d", sFecha1, sFecha2) + 1
    
    For I = 1 To nDias
        dFecha = sFecha1
        sTemp2 = sTemp2 & ValidaMigracion(dFecha)
        sFecha1 = DateAdd("d", 1, sFecha1)
    Next
    
    sTitulo = "V A L I D A C I O N   D E   A S I E N T O S"
    Linea lsCabecera, Cabecera(sTitulo, P, "", 50, , , "CMACT")
    sTemp = lsCabecera
    
    lsCabecera = ""
    Linea lsCabecera, "VALIDACION DE MIGRACION DE ASIENTOS CONTABLES ", 2
    sTemp = sTemp & lsCabecera
    
    
    
    If Len(Trim(sTemp2)) > 0 Then
        sTemp = sTemp & Chr(10) & sTemp2
   Else
        sTemp = sTemp & Chr(10) & " *** NO SE ENCONTRARON OBSERVACIONES *** "
   End If
   
   sTexto = sTexto & sTemp
   
End If
If chkPapa.value = vbChecked Then
   lblmsg.Caption = "Validando Cuentas Sin Padre... Espere un momento"
   DoEvents
   sTexto = sTexto & oBalance.ImprimeValidaBalance(Format(txtfechaDel, gsFormatoFecha), Format(txtFechaAl, gsFormatoFecha), gnLinPage, gValidaCuentasSinPadre)
End If

If Me.chkCtasConsolidada.value = vbChecked Then
   lblmsg.Caption = "Validando Cuentas Sin Creacion de cuentas Consolidada ... Espere un momento"
   DoEvents
   sTexto = sTexto & oBalance.ImprimeValidaBalance(Format(txtfechaDel, gsFormatoFecha), Format(txtFechaAl, gsFormatoFecha), gnLinPage, gValidaCuentasConsolidadas)
End If

If Me.chkDecimales.value = vbChecked Then
   lblmsg.Caption = "Validando Asientos con importes de mas de dos digitos ... Espere un momento"
   DoEvents
   sTexto = sTexto & oBalance.ImprimeValidaBalance(Format(txtfechaDel, gsFormatoFecha), Format(txtFechaAl, gsFormatoFecha), gnLinPage, gValidaAsientosDigitos)
End If


Set oBalance = Nothing
lblmsg.Visible = False
MousePointer = 0
EnviaPrevio sTexto, "VALIDACION DE BALANCE: Resultado", gnLinPage, False
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
CentraForm Me
End Sub

Private Sub txtFechaAl_GotFocus()
fEnfoque txtFechaAl
End Sub

Private Sub txtFechaAl_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If ValidaFecha(txtFechaAl.Text) <> "" Then
      MsgBox "Fecha no válida...!", vbInformation, "Error"
      txtFechaAl.SetFocus
   End If
   cmdProcesar.SetFocus
End If
End Sub

Private Sub txtFechaDel_GotFocus()
fEnfoque txtfechaDel
End Sub

Private Sub txtFechaDel_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If ValidaFecha(txtfechaDel.Text) <> "" Then
      MsgBox "Fecha no válida...!", vbInformation, "Error"
      txtfechaDel.SetFocus
   End If
   txtFechaAl.SetFocus
End If
End Sub


