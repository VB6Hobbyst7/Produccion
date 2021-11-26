VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmBalanceHisto 
   Caption         =   "Balance Histórico"
   ClientHeight    =   6660
   ClientLeft      =   630
   ClientTop       =   2025
   ClientWidth     =   11265
   Icon            =   "frmBalanceHisto.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   6660
   ScaleWidth      =   11265
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkSoloAnaliticas 
      Caption         =   "Generar sólo Cuentas Analíticas"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   6420
      Width           =   2745
   End
   Begin VB.CheckBox chkCierreAnio 
      Caption         =   "&Incluir Asiento de Cierre de Ejercicio"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   6180
      Width           =   3165
   End
   Begin VB.CheckBox chkSele 
      Caption         =   "Imprimir sólo S&elección"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   5700
      Width           =   2745
   End
   Begin VB.CheckBox chkFecha 
      Caption         =   "&Utilizar fecha final de Balance"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   5940
      Width           =   2745
   End
   Begin VB.CommandButton cmdSituacion 
      Caption         =   "Balance de &Situación"
      Enabled         =   0   'False
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
      Left            =   7260
      TabIndex        =   18
      Top             =   6000
      Width           =   2640
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
      Height          =   390
      Left            =   3360
      TabIndex        =   17
      Top             =   6000
      Width           =   1200
   End
   Begin VB.Frame frmMoneda 
      Caption         =   "Moneda"
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
      Height          =   765
      Left            =   6480
      TabIndex        =   12
      Top             =   30
      Width           =   4665
      Begin VB.ComboBox cboMoneda 
         Height          =   315
         ItemData        =   "frmBalanceHisto.frx":030A
         Left            =   120
         List            =   "frmBalanceHisto.frx":0323
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   300
         Width           =   4425
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      CausesValidation=   0   'False
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
      Left            =   9930
      TabIndex        =   5
      Top             =   6000
      Width           =   1200
   End
   Begin VB.Frame Frame4 
      Height          =   765
      Left            =   4320
      TabIndex        =   9
      Top             =   30
      Width           =   2115
      Begin VB.ComboBox cboDig 
         Height          =   315
         ItemData        =   "frmBalanceHisto.frx":03DB
         Left            =   1320
         List            =   "frmBalanceHisto.frx":0403
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   300
         Width           =   645
      End
      Begin VB.Label lblDig 
         Caption         =   "Nro de Dígitos"
         Height          =   225
         Left            =   150
         TabIndex        =   10
         Top             =   360
         Width           =   1065
      End
   End
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
      ForeColor       =   &H8000000D&
      Height          =   765
      Left            =   90
      TabIndex        =   6
      Top             =   30
      Width           =   4170
      Begin MSMask.MaskEdBox txtFechaDel 
         Height          =   345
         Left            =   690
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
         Left            =   2760
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
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "AL"
         Height          =   195
         Left            =   2370
         TabIndex        =   8
         Top             =   390
         Width           =   195
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "DEL"
         Height          =   195
         Left            =   150
         TabIndex        =   7
         Top             =   390
         Width           =   315
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid fg 
      Height          =   4515
      Left            =   90
      TabIndex        =   3
      Top             =   840
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   7964
      _Version        =   393216
      Cols            =   8
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      TextStyleFixed  =   3
      FocusRect       =   0
      FillStyle       =   1
      AllowUserResizing=   1
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   8
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin RichTextLib.RichTextBox rtf 
      Height          =   315
      Left            =   1410
      TabIndex        =   11
      Top             =   4710
      Visible         =   0   'False
      Width           =   645
      _ExtentX        =   1138
      _ExtentY        =   556
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmBalanceHisto.frx":0432
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtDebe 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   10250
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   255
      Left            =   7980
      TabIndex        =   15
      Top             =   5385
      Width           =   1470
   End
   Begin VB.TextBox txtHaber 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   10250
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   255
      Left            =   9465
      TabIndex        =   14
      Top             =   5385
      Width           =   1470
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Balance de Comprobación"
      Enabled         =   0   'False
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
      Left            =   4590
      TabIndex        =   4
      Top             =   6000
      Width           =   2640
   End
   Begin VB.Label lblTotal 
      BackColor       =   &H00E0E0E0&
      Caption         =   "TOTALES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   6630
      TabIndex        =   16
      Top             =   5430
      Width           =   855
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000C&
      X1              =   9450
      X2              =   9450
      Y1              =   5370
      Y2              =   5670
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000C&
      X1              =   7965
      X2              =   7965
      Y1              =   5370
      Y2              =   5670
   End
   Begin VB.Label lblmsg 
      Caption         =   "Procesando...Por favor espere un momento."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   150
      TabIndex        =   13
      Top             =   5430
      Visible         =   0   'False
      Width           =   4245
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000C&
      Height          =   300
      Left            =   6480
      Top             =   5370
      Width           =   4470
   End
   Begin VB.Menu mnuBala 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuBuscar 
         Caption         =   "Buscar"
      End
   End
End
Attribute VB_Name = "frmBalanceHisto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sSql       As String
Dim rs         As New ADODB.Recordset
Dim nTipoBala  As Integer
Dim nTipCambio As Currency
Dim nMoneda    As Integer
Dim lbValidaBalance As Boolean

Public Sub Inicio(pnTipoBala As Integer)
nTipoBala = pnTipoBala
Me.Show 0, frmMdiMain
End Sub

Private Sub FormatoFlex()
If fg.Rows = 1 Then
   fg.Rows = 2
   fg.Row = 1
End If
If cboMoneda.ListIndex = 2 Then
   fg.Cols = 8
   fg.TextMatrix(0, 7) = "SALDO FINAL ME"
   fg.ColWidth(7) = 1450
   fg.ColAlignment(7) = 7
End If
fg.TextMatrix(0, 1) = "Cuenta Contable"
fg.TextMatrix(0, 2) = "Descripción"
fg.TextMatrix(0, 3) = "SALDO INICIAL"
fg.TextMatrix(0, 4) = "DEBE"
fg.TextMatrix(0, 5) = "HABER"
fg.TextMatrix(0, 6) = "SALDO FINAL"
fg.ColWidth(0) = 200
fg.ColWidth(1) = 1400
fg.ColWidth(2) = 3400
fg.ColWidth(3) = 1450
fg.ColWidth(4) = 1450
fg.ColWidth(5) = 1450
fg.ColWidth(6) = 1450
fg.ColAlignment(1) = 1
fg.ColAlignment(3) = 7
fg.ColAlignment(4) = 7
fg.ColAlignment(5) = 7
fg.ColAlignment(6) = 7
fg.RowHeight(-1) = 285
End Sub
Private Sub CambiaFormatoFlex()
Dim N As Integer
Dim c As Integer
Dim nDebe As Currency, nHaber As Currency
nDebe = 0: nHaber = 0
For N = 1 To fg.Rows - 1
   If cboMoneda.ListIndex = 2 Then
      fg.TextMatrix(N, 7) = Format(Round(Val(fg.TextMatrix(N, 6)) / nTipCambio, 2), gsFormatoNumeroView)
   End If
   For c = 3 To 6
      fg.TextMatrix(N, c) = Format(fg.TextMatrix(N, c), gsFormatoNumeroView)
   Next
Next
End Sub

Private Sub cboDig_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cboMoneda.SetFocus
End If
End Sub

Private Sub CboMoneda_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdProcesar.SetFocus
End If
End Sub

Private Sub chkCierreAnio_Click()
If chkCierreAnio.value = vbChecked Then
    If Not (Month(txtFechaDel) = 12 And Month(txtFechaAl) = 12) Then
        MsgBox "Mes de Balance debe ser Diciembre", vbInformation, "!Aviso¡"
        chkCierreAnio.value = vbUnchecked
    End If
End If
End Sub

Private Sub cmdImprimir_Click()
Dim sTexto As String
lblmsg.Visible = True
lblmsg.Caption = "Procesando Balance de Comprobación..."
Me.MousePointer = 11
fg.MousePointer = 11

Dim nPosIni As Integer, nPosFin As Integer
Dim lsCtaIni As String, lsCtaFin As String
nPosIni = 1
nPosFin = fg.Rows - 1
lsCtaIni = ""
lsCtaFin = ""
If chkSele.value = vbChecked Then
   nPosIni = fg.Row
   nPosFin = fg.Row
   If fg.RowSel > fg.Row Then
      nPosFin = fg.RowSel
   Else
      nPosIni = fg.RowSel
   End If
   lsCtaIni = fg.TextMatrix(nPosIni, 1)
   lsCtaFin = fg.TextMatrix(nPosFin, 1)
End If

Dim oBalance As New NBalanceCont
sTexto = oBalance.ImprimeBalanceComprobacion(CDate(txtFechaDel), CDate(txtFechaAl), IIf(chkFecha.value = 1, txtFechaAl, gdFecSis), nTipoBala, nMoneda, gnLinPage, nVal(txtDebe), nVal(txtHaber), lsCtaIni, lsCtaFin, Val(cboDig), chkSoloAnaliticas.value, Me.chkCierreAnio.value)
Set oBalance = Nothing
EnviaPrevio sTexto, "BALANCE DE COMPROBACION", gnLinPage, False
Me.MousePointer = 0
fg.MousePointer = 0
lblmsg.Visible = False
fg.SetFocus
End Sub
Private Function PermiteGenerarBalance() As Boolean
Dim oGen As New DGeneral
Dim oBlq As New DBloqueos
PermiteGenerarBalance = False
   Set rs = oBlq.CargaBloqueo(gBloqueoBalance)
   If Not rs.EOF Then
      If Trim(rs!cVarValor) = "1" Then
         Set rs = oGen.GetDataUser(Right(rs!cUltimaActualizacion, 4))
         If Not rs.EOF Then
            If MsgBox(Trim(rs!cPersNombre) & " esta Generando Balance. ¿Desea de todas maneras generar el Balance?", vbQuestion + vbYesNo + vbDefaultButton2, "Aviso") = vbNo Then
               RSClose rs
               Exit Function
            End If
         Else
            MsgBox "Existe otro usuario Generando Balance. Por favor procesar posteriormente", vbInformation, "Aviso"
            RSClose rs
            Exit Function
         End If
      End If
   End If
   oBlq.ActualizaBloqueo gBloqueoBalance, "1", GeneraMovNroActualiza(Format(GetFechaHoraServer(), gsFormatoFecha), gsCodUser, gsCodCMAC, gsCodAge)
Set oGen = Nothing
Set oBlq = Nothing
PermiteGenerarBalance = True
End Function


Private Sub cmdProcesar_Click()
Dim sBalance As String
Dim sCodUsu  As String
Dim sCondBala As String, sCondBala2 As String
Dim sCond1 As String, sCond2 As String
Dim sCta   As String
Dim N      As Integer
Dim nPos   As Variant
Dim sCtaGrp As String

On Error GoTo ErrGeneraBalance

If cboMoneda.ListIndex = 0 Then gsSimbolo = "": nMoneda = 0
If cboMoneda.ListIndex = 1 Then gsSimbolo = gcMN: nMoneda = 1
If cboMoneda.ListIndex = 2 Then gsSimbolo = gcME: nMoneda = 2
If cboMoneda.ListIndex = 5 Then gsSimbolo = gcME: nMoneda = 6
If cboMoneda.ListIndex = 6 Then gsSimbolo = "": nMoneda = 9
If Trim(txtFechaDel) = "/  /" Or Trim(txtFechaAl) = "/  /" Or cboDig.ListIndex = -1 Then
   Exit Sub
End If

If Day(CDate(txtFechaDel)) <> 1 Then
   MsgBox "Balance debe generarse desde el primer día del Mes", vbInformation, "¡Aviso!"
   txtFechaDel.SetFocus
   Exit Sub
End If

If Not PermiteGenerarBalance() Then
   Exit Sub
End If


Dim oBalance As New NBalanceCont
Dim oBlq     As New DBloqueos
lblmsg.Visible = True
lblmsg.Caption = "Procesando... Por favor espere un momento"
DoEvents
MousePointer = 11
Dim oFun As NContFunciones
Set oFun = New NContFunciones
gsMovNro = oFun.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)

N = vbYes

If oBalance.BalanceGeneradoHisto(nTipoBala, nMoneda, Month(CDate(txtFechaDel)) + CInt(Me.chkCierreAnio), Year(CDate(txtFechaDel))) Then
    If oFun.PermiteModificarAsiento(Format(txtFechaDel, gsFormatoMovFecha), False) Then
       MousePointer = 0
       N = MsgBox("Balance ya fue generado. ¿Desea volver a Procesar?", vbQuestion + vbYesNoCancel + vbDefaultButton2, "Aviso")
       If N = vbCancel Then
          oBlq.ActualizaBloqueo gBloqueoBalance, "0", gsMovNro
          lblmsg.Caption = ""
          Exit Sub
       End If
       MousePointer = 11
       DoEvents
    Else
        N = vbNo
    End If
End If
Set oFun = Nothing


If N = vbYes Then
   Dim dBalance As New DbalanceCont
   Dim lsMsgErr As String
   lbValidaBalance = oBalance.BalanceGeneradoHisto(nTipoBala, nMoneda, Month(CDate(txtFechaDel) - 1) + CInt(Me.chkCierreAnio), Year(CDate(txtFechaDel) - 1))
   If lbValidaBalance Then
      lblmsg.Caption = "Validando Saldos Iniciales... Espere un momento"
      DoEvents
   
      lsMsgErr = dBalance.ValidaSaldosIniciales(Month(txtFechaDel), Year(CDate(txtFechaDel)))
      If lsMsgErr <> "" Then
         If MsgBox(TextErr(lsMsgErr), vbQuestion + vbYesNo + vbDefaultButton2, "¿Desea continuar? ") = vbNo Then
            lblmsg.Caption = ""
            Exit Sub
         End If
      End If
   End If
   oBlq.ActualizaBloqueo gBloqueoBalance, "1", gsMovNro
   'ELIMINA BALANCE DEL MES
   dBalance.EliminaBalance nTipoBala, nMoneda, Month(CDate(txtFechaDel)) + CInt(Me.chkCierreAnio), Year(CDate(txtFechaDel))
   dBalance.EliminaBalanceTemp nTipoBala, nMoneda
   
   lblmsg.Caption = "Cálculo de Movimientos..."
   DoEvents
   
   'SALDOS INICIALES
   dBalance.InsertaSaldosIniciales nTipoBala, nMoneda, Format(txtFechaDel, gsFormatoFecha)    ', True
   
   'MOVIMIENTOS DEL MES
'OJO no considerar este tipo de Operacion de Cierre Anual
'
'cOpeCod LIKE '" & Left(gContCierreAnual, 5) & "%' "
'--------------------

   dBalance.InsertaMovimientosMes nTipoBala, nMoneda, Format(txtFechaDel, gsFormatoMovFecha), Format(txtFechaAl, gsFormatoMovFecha), , chkCierreAnio.value = vbChecked
   
   lblmsg.Caption = "Mayorizando... Espere un momento"
   DoEvents
   If cboMoneda.ListIndex = 2 Then
      nTipCambio = oBalance.GetTipCambioBalance(Format(txtFechaAl, gsFormatoMovFecha))
      If nTipCambio = 0 Then
         MsgBox "No se definio Tipo de Cambio para fecha Final del Balance. Se usara TC.Fijo del día", vbInformation, "¡Aviso!"
         nTipCambio = gnTipCambio
      End If
   End If
   
   'MAYORIZACION
   dBalance.MayorizacionBalance nTipoBala, nMoneda, Month(CDate(txtFechaDel)), Year(CDate(txtFechaDel)), chkCierreAnio.value = vbChecked, gsCodCMAC
   'dBalance.EjecutaBatch
   Set dBalance = Nothing
   
   'VALIDACION DE BALANCE
   lblmsg.Caption = "Validando Balance... Por favor espere un momento"
   DoEvents
   sBalance = oBalance.ValidaBalance(False, CDate(txtFechaDel), CDate(txtFechaAl), nTipoBala, nMoneda)
   EnviaPrevio sBalance, "Generación de Balance: Validación", gnLinPage, False
Else
   If cboMoneda.ListIndex = 2 Then
      nTipCambio = oBalance.GetTipCambioBalance(Format(txtFechaAl, gsFormatoMovFecha))
   End If
End If
Set rs = oBalance.TotalizaBalanceHisto(nTipoBala, nMoneda, Month(CDate(txtFechaDel)) + CInt(Me.chkCierreAnio), Year(CDate(txtFechaDel)))
txtDebe = Format(rs!nDebe, gsFormatoNumeroView)
txtHaber = Format(rs!nHaber, gsFormatoNumeroView)

oBlq.ActualizaBloqueo gBloqueoBalance, "0", gsMovNro

lblmsg.Caption = "Mostrando Balance... Espere un momento"
DoEvents
Set rs = oBalance.LeeBalanceHisto(nTipoBala, nMoneda, Month(CDate(txtFechaDel)) + CInt(Me.chkCierreAnio), Year(CDate(txtFechaDel)), , , cboDig, chkSoloAnaliticas)
Set fg.Recordset = rs
If fg.Rows - 1 < rs.RecordCount Then
   rs.Move fg.Rows - 1
   Do While Not rs.EOF
      fg.AddItem ""
      nPos = fg.Rows - 1
      For N = 0 To rs.Fields.Count - 1
         fg.TextMatrix(nPos, N + 1) = rs.Fields(N)
      Next
      rs.MoveNext
   Loop
End If
RSClose rs
Set oBlq = Nothing
Set oBalance = Nothing

FormatoFlex
CambiaFormatoFlex
lblmsg.Visible = False
MousePointer = 0
cmdImprimir.Enabled = True
cmdSituacion.Enabled = True

Exit Sub
ErrGeneraBalance:
   MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
   RSClose rs
   Set oBlq = Nothing
   Set oBalance = Nothing
   
   MousePointer = 0
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub cmdSituacion_Click()
Dim sTexto As String
MousePointer = 11
lblmsg.Visible = True
lblmsg.Caption = "Procesando Balance de Situación..."
Dim oBalance As New NBalanceCont
sTexto = oBalance.ImprimeBalanceSituacion(CDate(txtFechaDel), CDate(txtFechaAl), IIf(chkFecha.value = 1, txtFechaAl, gdFecSis), nTipoBala, nMoneda, nVal(txtDebe), nVal(txtHaber), Me.chkCierreAnio.value)
Set oBalance = Nothing
lblmsg.Visible = False
MousePointer = 0
lblmsg.Caption = ""
EnviaPrevio sTexto, "Balance de Situación", gnLinPage, False
End Sub

Private Sub fg_KeyUp(KeyCode As Integer, Shift As Integer)
Flex_PresionaKey fg, KeyCode, Shift
End Sub

Private Sub fg_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
   PopupMenu mnuBala
End If
End Sub

Private Sub Form_Load()
CentraForm Me
Me.Caption = "Balance " & IIf(nTipoBala = 1, "Histórico", "Ajustado")
txtFechaDel = "01/" & Format(Month(gdFecSis), "00") & "/" & Format(Year(gdFecSis), "0000")
txtFechaAl = gdFecSis
cboDig.ListIndex = cboDig.ListCount - 1
cboMoneda.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
CierraConexion
End Sub

Private Sub mnuBuscar_Click()
Dim sTexto As String
Dim N      As Integer
sTexto = InputBox("Cuenta Contable: ", "Busqueda de Cuenta")
If sTexto <> "" Then
   For N = 1 To fg.Rows - 1
      If fg.TextMatrix(N, 1) = sTexto Then
         fg.Row = N
         fg.TopRow = N
      End If
   Next
End If
fg.SetFocus
End Sub

Private Sub optMoneda_Click(Index As Integer)
If (Index = 0 And fg.Cols <> 7) Or (Index = 1 And fg.Cols <> 8) Then
   cmdImprimir.Enabled = False
   cmdSituacion.Enabled = False
End If
End Sub

Private Sub txtFechaAl_GotFocus()
txtFechaAl.SelStart = 0
txtFechaAl.SelLength = Len(txtFechaAl)
End Sub

Private Sub txtFechaAl_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If ValidaFecha(txtFechaAl.Text) <> "" Then
      MsgBox "Fecha no válida...!", vbInformation, "Error"
      txtFechaAl.SetFocus
   End If
   cboDig.SetFocus
End If
End Sub

Private Sub txtFechaAl_Validate(Cancel As Boolean)
If ValidaFecha(txtFechaAl.Text) <> "" Then
   MsgBox "Fecha no válida...!", vbInformation, "Error"
   Cancel = True
End If
End Sub

Private Sub txtFechaDel_GotFocus()
txtFechaDel.SelStart = 0
txtFechaDel.SelLength = Len(txtFechaDel)
End Sub

Private Sub txtFechaDel_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If ValidaFecha(txtFechaDel.Text) <> "" Then
      MsgBox "Fecha no válida...!", vbInformation, "Error"
      txtFechaDel.SetFocus
   End If
   txtFechaAl.SetFocus
End If
End Sub

Private Sub txtFechaDel_Validate(Cancel As Boolean)
If ValidaFecha(txtFechaDel.Text) <> "" Then
   MsgBox "Fecha no válida...!", vbInformation, "Error"
   Cancel = True
End If
End Sub
