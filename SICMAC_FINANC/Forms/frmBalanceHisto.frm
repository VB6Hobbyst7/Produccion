VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmBalanceHisto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Balance Histórico"
   ClientHeight    =   7035
   ClientLeft      =   615
   ClientTop       =   2010
   ClientWidth     =   11880
   Icon            =   "frmBalanceHisto.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDetGsto 
      Caption         =   "Balance Detallado del Gasto"
      Height          =   375
      Left            =   6480
      TabIndex        =   27
      Top             =   6600
      Width           =   2655
   End
   Begin VB.CommandButton cmdGsto 
      Caption         =   "Balance Comparativo del Gasto"
      Height          =   375
      Left            =   6480
      TabIndex        =   26
      Top             =   6240
      Width           =   2655
   End
   Begin VB.CheckBox ChkCuentaCero 
      Caption         =   "Considerar todas las  Cuentas con Movimiento Histórico"
      Height          =   375
      Left            =   90
      TabIndex        =   25
      Top             =   6300
      Width           =   3105
   End
   Begin VB.CheckBox chkExcel 
      Caption         =   "Imprimir en Excel"
      Height          =   255
      Left            =   3300
      TabIndex        =   24
      Top             =   5700
      Width           =   1695
   End
   Begin VB.CheckBox chkSoloAnaliticas 
      Caption         =   "Generar sólo Cuentas Analíticas"
      Height          =   255
      Left            =   90
      TabIndex        =   23
      Top             =   6000
      Width           =   2745
   End
   Begin VB.CheckBox chkCierreAnio 
      Caption         =   "&Incluir Asiento de Cierre de Ejercicio"
      Height          =   255
      Left            =   90
      TabIndex        =   21
      Top             =   5700
      Width           =   3165
   End
   Begin VB.CheckBox chkSele 
      Caption         =   "Imprimir sólo S&elección"
      Height          =   255
      Left            =   3300
      TabIndex        =   20
      Top             =   6000
      Width           =   2745
   End
   Begin VB.CheckBox chkFecha 
      Caption         =   "&Utilizar fecha final de Balance"
      Height          =   255
      Left            =   3300
      TabIndex        =   19
      Top             =   6300
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
      Left            =   9150
      TabIndex        =   18
      Top             =   5760
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
      Left            =   10560
      TabIndex        =   17
      Top             =   270
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
      Left            =   5850
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
      Left            =   10560
      TabIndex        =   5
      Top             =   6240
      Width           =   1200
   End
   Begin VB.Frame Frame4 
      Height          =   765
      Left            =   3720
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
      Width           =   3600
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
         Left            =   2310
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
         Left            =   2040
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
      Width           =   11685
      _ExtentX        =   20611
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
      Left            =   6480
      TabIndex        =   4
      Top             =   5760
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

Dim lsArchivo As String
Dim lbExcel As Boolean
Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet
Dim ApExcel As Variant

Public Sub ImprimeBalanceComprobacionExcel(pdFechaIni As Date, pdFechaFin As Date, pdFecha As Date, pnTipoBala As Integer, pnMoneda As Integer, pnLinPage As Integer, nTotDebe As Currency, nTotHaber As Currency, Optional psCtaIni As String = "", Optional psCtaFin As String = "", Optional pnDigitos As Integer = 0, Optional pbSoloAnaliticas As Boolean = False, Optional pnCierreAnio As Integer = 0, Optional nTipo As Integer)
Dim lsImpre As String
On Error GoTo GeneraEstadError
   lsArchivo = App.path & "\SPOOLER\" & "Balance_" & pnMoneda & "_" & Mid(Format(pdFechaIni, "yyyymmdd"), 5, 2) & ".XLS"
   lbExcel = ExcelBegin(lsArchivo, xlAplicacion, xlLibro, False)
   If Not lbExcel Then
      Exit Sub
   End If
   ExcelAddHoja "B_" & Format(pdFechaIni, "yyyymmdd") & "_" & Right(Format(pdFechaFin, "yyyymmdd"), 2), xlLibro, xlHoja1, False

   Dim oBalance As New NBalanceCont
    oBalance.ImprimeBalanceComprobacion pdFechaIni, pdFechaFin, pdFecha, pnTipoBala, pnMoneda, pnLinPage, nTotDebe, nTotHaber, psCtaIni, psCtaFin, pnDigitos, pbSoloAnaliticas, pnCierreAnio, True, xlHoja1, nTipo
   Set oBalance = Nothing
    
    ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1, True
    If lsArchivo <> "" Then
       CargaArchivo lsArchivo, App.path & "\SPOOLER\"
    End If
   
Exit Sub
GeneraEstadError:
    MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
    If lbExcel = True Then
      ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1, False
    End If
End Sub


Public Sub Inicio(pnTipoBala As Integer)
nTipoBala = pnTipoBala
Me.Show 1 ', frmMdiMain
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
Dim N As Long
Dim c As Integer
Dim nDebe As Currency, nHaber As Currency

Dim oBalance As NBalanceCont
Set oBalance = New NBalanceCont
Dim rsCtasME As ADODB.Recordset
Set rsCtasME = New ADODB.Recordset

If cboMoneda.ListIndex = 2 Then
    Set rsCtasME = oBalance.GetCtasSaldoME(CDate(txtFechaAl.Text))
End If

nDebe = 0: nHaber = 0
For N = 1 To fg.Rows - 1
   If cboMoneda.ListIndex = 2 Then
      rsCtasME.MoveFirst
      rsCtasME.Find "cCtaContCod Like '" & fg.TextMatrix(N, 1) & "'"
      If rsCtasME.EOF And rsCtasME.EOF Then
         fg.TextMatrix(N, 7) = Format(Round(Val(fg.TextMatrix(N, 6)) / nTipCambio, 2), gsFormatoNumeroView)
      Else
         fg.TextMatrix(N, 7) = Format(rsCtasME.Fields(1), gsFormatoNumeroView)
      End If
   End If
   For c = 3 To 6
      fg.TextMatrix(N, c) = Format(fg.TextMatrix(N, c), gsFormatoNumeroView)
   Next
Next

Set oBalance = Nothing
Set rsCtasME = Nothing
End Sub

Private Sub cboDig_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cboMoneda.SetFocus
End If
End Sub

Private Sub cboMoneda_KeyPress(KeyAscii As Integer)
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

Private Sub cmdDetGsto_Click()
Dim lcAnio As String, lcMes As String

If Mid(txtFechaDel.Text, 4, 2) <> Mid(txtFechaAl.Text, 4, 2) Or Mid(txtFechaDel.Text, 7, 4) <> Mid(txtFechaAl.Text, 7, 4) Then
    MsgBox "El rango de fecha debe ser del mismo mes y año.", vbOKOnly + vbInformation, "Atención"
    Exit Sub
End If

ImprimeBalanceDetalladoDelGastoExcel Mid(txtFechaDel.Text, 7, 4), Mid(txtFechaDel.Text, 4, 2)

End Sub

Private Sub cmdGsto_Click()
Dim lcAnio As String, lcMes As String

If Mid(txtFechaDel.Text, 4, 2) <> Mid(txtFechaAl.Text, 4, 2) Or Mid(txtFechaDel.Text, 7, 4) <> Mid(txtFechaAl.Text, 7, 4) Then
    MsgBox "El rango de fecha debe ser del mismo mes y año.", vbOKOnly + vbInformation, "Atención"
    Exit Sub
End If
ImprimeBalanceComparativoDelGastoExcel Mid(txtFechaDel.Text, 7, 4), Mid(txtFechaDel.Text, 4, 2)

End Sub

Private Sub cmdImprimir_Click()
Dim sTexto As String
lblMsg.Visible = True
lblMsg.Caption = "Procesando Balance de Comprobación..."
Me.MousePointer = 11
fg.MousePointer = 11

Dim nPosIni As Long, nPosFin As Long
Dim lsCtaIni As String, lsCtaFin As String
nPosIni = 1
nPosFin = fg.Rows - 1
lsCtaIni = ""
lsCtaFin = ""

'If CDate(txtFechaAl.Text) >= gdFecSis Then
'   MsgBox "Fecha Final mayor o igual a Fecha Actual "
'   Exit Sub

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

    If chkExcel.value = vbChecked Then
       ImprimeBalanceComprobacionExcel CDate(txtFechaDel), CDate(txtFechaAl), IIf(chkFecha.value = 1, txtFechaAl, gdFecSis), nTipoBala, nMoneda, gnLinPage, nVal(txtDebe), nVal(txtHaber), lsCtaIni, lsCtaFin, Val(cboDig), chkSoloAnaliticas.value, Me.chkCierreAnio.value, nMoneda
    Else
       Dim oBalance As New NBalanceCont
       sTexto = oBalance.ImprimeBalanceComprobacion(CDate(txtFechaDel), CDate(txtFechaAl), IIf(chkFecha.value = 1, txtFechaAl, gdFecSis), nTipoBala, nMoneda, gnLinPage, nVal(txtDebe), nVal(txtHaber), lsCtaIni, lsCtaFin, Val(cboDig), chkSoloAnaliticas.value, Me.chkCierreAnio.value, , , nMoneda)
       EnviaPrevio sTexto, "BALANCE DE COMPROBACION", gnLinPage, False
       Set oBalance = Nothing
    End If
    Me.MousePointer = 0
    fg.MousePointer = 0
    lblMsg.Visible = False
    fg.SetFocus
'End If
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
Dim bConHistorico As Boolean

On Error GoTo ErrGeneraBalance

If cboMoneda.ListIndex = 0 Then gsSimbolo = "": nMoneda = 0
If cboMoneda.ListIndex = 1 Then gsSimbolo = gcMN: nMoneda = 1
If cboMoneda.ListIndex = 2 Then gsSimbolo = gcME: nMoneda = 2
If cboMoneda.ListIndex = 3 Then gsSimbolo = "": nMoneda = 3
If cboMoneda.ListIndex = 4 Then gsSimbolo = "": nMoneda = 4
If cboMoneda.ListIndex = 5 Then gsSimbolo = "": nMoneda = 6
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
lblMsg.Visible = True
lblMsg.Caption = "Procesando... Por favor espere un momento"
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
          lblMsg.Caption = ""
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
      lblMsg.Caption = "Validando Saldos Iniciales... Espere un momento"
      DoEvents
   
      lsMsgErr = dBalance.ValidaSaldosIniciales(Month(txtFechaDel), Year(CDate(txtFechaDel)))
      If lsMsgErr <> "" Then
         If MsgBox(TextErr(lsMsgErr), vbQuestion + vbYesNo + vbDefaultButton2, "¿Desea continuar? ") = vbNo Then
            lblMsg.Caption = ""
            Exit Sub
         End If
      End If
   End If
   oBlq.ActualizaBloqueo gBloqueoBalance, "1", gsMovNro
   'ELIMINA BALANCE DEL MES
   dBalance.EliminaBalance nTipoBala, nMoneda, Month(CDate(txtFechaDel)) + CInt(Me.chkCierreAnio), Year(CDate(txtFechaDel))
   dBalance.EliminaBalanceTemp nTipoBala, nMoneda
   
   lblMsg.Caption = "Cálculo de Movimientos..."
   DoEvents
   
   'SALDOS INICIALES
   If Me.ChkCuentaCero.value = 1 Then bConHistorico = True Else bConHistorico = False
   dBalance.InsertaSaldosIniciales nTipoBala, nMoneda, Format(txtFechaDel, gsFormatoFecha), False  ', bConHistorico
   
   'MOVIMIENTOS DEL MES
'OJO no considerar este tipo de Operacion de Cierre Anual
'
'cOpeCod LIKE '" & Left(gContCierreAnual, 5) & "%' "
'--------------------

   dBalance.InsertaMovimientosMes nTipoBala, nMoneda, Format(txtFechaDel, gsFormatoMovFecha), Format(txtFechaAl, gsFormatoMovFecha), , chkCierreAnio.value = vbChecked
   
   lblMsg.Caption = "Mayorizando... Espere un momento"
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
   lblMsg.Caption = "Validando Balance... Por favor espere un momento"
   DoEvents
   ValidaBalanceEXCEL False, CDate(txtFechaDel), CDate(txtFechaAl), nTipoBala, nMoneda
   'sBalance = oBalance.ValidaBalance(False, CDate(txtFechaDel), CDate(txtFechaAl), nTipoBala, nMoneda)
'   EnviaPrevio sBalance, "Generación de Balance: Validación", gnLinPage, False
Else
   If cboMoneda.ListIndex = 2 Then
      nTipCambio = oBalance.GetTipCambioBalance(Format(txtFechaAl, gsFormatoMovFecha))
   End If
End If
Set rs = oBalance.TotalizaBalanceHisto(nTipoBala, nMoneda, Month(CDate(txtFechaDel)) + CInt(Me.chkCierreAnio), Year(CDate(txtFechaDel)))
txtDebe = Format(rs!nDebe, gsFormatoNumeroView)
txtHaber = Format(rs!nHaber, gsFormatoNumeroView)

oBlq.ActualizaBloqueo gBloqueoBalance, "0", gsMovNro

lblMsg.Caption = "Mostrando Balance... Espere un momento"
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
lblMsg.Visible = False
MousePointer = 0
CmdImprimir.Enabled = True
cmdSituacion.Enabled = True

Exit Sub
ErrGeneraBalance:
   MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
   'RSClose
   rs.Close
   Set oBlq = Nothing
   Set oBalance = Nothing
   
   MousePointer = 0
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub cmdSituacion_Click()
Dim sTexto As String
Dim sTexto2 As String
MousePointer = 11
lblMsg.Visible = True
lblMsg.Caption = "Procesando Balance de Situación..."
Dim oBalance As New NBalanceCont
sTexto = oBalance.ImprimeBalanceSituacion(CDate(txtFechaDel), CDate(txtFechaAl), IIf(chkFecha.value = 1, txtFechaAl, gdFecSis), nTipoBala, nMoneda, nVal(txtDebe), nVal(txtHaber), Me.chkCierreAnio.value)

sTexto2 = AgregaUtilidad(False, CDate(txtFechaDel), CDate(txtFechaAl), nTipoBala, nMoneda)

sTexto = sTexto & sTexto2

Set oBalance = Nothing
lblMsg.Visible = False
MousePointer = 0
lblMsg.Caption = ""
EnviaPrevio sTexto, "Balance de Situación", gnLinPage, False
End Sub

Private Sub Command1_Click()

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

'*** PEAC 20110412
If nTipoBala = 1 Then
    cmdGsto.Enabled = True
    cmdDetGsto.Enabled = True
Else
    cmdGsto.Enabled = False
    cmdDetGsto.Enabled = False
End If
'*** FIN PEAC

End Sub

Private Sub Form_Unload(Cancel As Integer)
CierraConexion
End Sub

Private Sub mnuBuscar_Click()
Dim sTexto As String
Dim N      As Long
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
   CmdImprimir.Enabled = False
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


Private Function AgregaUtilidad(lSoloUtilidad As Boolean, pdFechaIni As Date, pdFechaFin As Date, pnTipoBala As Integer, pnMoneda As Integer) As String
Dim nUtilidad As Currency
Dim nUtilidadMes As Currency
Dim nRei As Currency
Dim nDeduccion As Currency
Dim nDeduccion1 As Currency
Dim sValida    As String
Dim n5 As Currency, n4 As Currency
Dim n62 As Currency, n63 As Currency, n64 As Currency, n65 As Currency, n66 As Currency
Dim oBal As NBalanceCont
Dim glsarchivo As String

'********************************************
nUtilidad = 0
nUtilidadMes = 0
Set oBal = New NBalanceCont

If Month(pdFechaIni) > 1 Then
   nUtilidad = oBal.GetUtilidadAcumulada(Format(pnTipoBala, "#"), pnMoneda, Format(Month(pdFechaIni - 1), "00"), Format(Year(pdFechaIni - 1), "0000"))
End If

n5 = oBal.getImporteBalanceMes("5", pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni))
n62 = oBal.getImporteBalanceMes("62", pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni))
n63 = oBal.getImporteBalanceMes("63", pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni))
n64 = oBal.getImporteBalanceMes("64", pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni))
n65 = oBal.getImporteBalanceMes("65", pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni))
n66 = oBal.getImporteBalanceMes("66", pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni))
n4 = oBal.getImporteBalanceMes("4", pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni))
nUtilidadMes = n5 + n62 + n64 - (n4 + n63 + n65)

'69
nRei = oBal.getImporteBalanceMes("69", pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni))
If gsCodCMAC = "102" Then
   nRei = nRei * -1
End If
nDeduccion = oBal.getImporteBalanceMes("67", pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni))
nDeduccion1 = oBal.getImporteBalanceMes("68", pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni))
nDeduccion = nDeduccion * -1
nDeduccion1 = nDeduccion1 * -1

'If Not lSoloUtilidad Then
'   glsArchivo = "C A L C U L O   D E   L A   U T I L I D A D" & " " & " " & Format(gdFecSis, "ddmmyyyy") & "_" & Format(Time(), "HHMMSS") & ".XLS"
'   If pnMoneda = 0 Then
'      glsArchivo1 = "C O N S O L I D A D O" & " " & "AL " & pdFechaFin
'   End If
'End If
    
If Not lSoloUtilidad Then

   Dim nActivo As Currency
   Dim nPasivo As Currency
   Dim nPatri  As Currency

   'Eliminamos si Existe la Utilidad Acumulada del Mes
   Dim dBalance As New DbalanceCont
   dBalance.EliminaUtilidadAcumulada pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni), True
   dBalance.InsertaUtilidadAcumulada pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni), nUtilidadMes, True
   dBalance.EjecutaBatch

   nActivo = oBal.getImporteBalanceMes("1", pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni))
   nPasivo = oBal.getImporteBalanceMes("2", pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni))
   nPatri = oBal.getImporteBalanceMes("3", pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni))


'   Select Case pnMoneda
'      Case 0: xlHoja1.Cells(liLineas + 15, 2) = " ( CONSOLIDADO ) "
'      Case 1: xlHoja1.Cells(liLineas + 15, 2) = " ( MONEDA NACIONAL ) "
'      Case 2: xlHoja1.Cells(liLineas + 15, 2) = " ( MONEDA EXTRANJERA ) "
'   End Select
   
   glsarchivo = glsarchivo & Space(20) & "CONSTANCIA DE CUADRE DE BALANCE" & oImpresora.gPrnSaltoLinea
   If pnMoneda = 0 Then
      glsarchivo = glsarchivo & Space(20) & "        (CONSOLIDADO)          " & oImpresora.gPrnSaltoLinea
   ElseIf pnMoneda = 2 Then
      glsarchivo = glsarchivo & Space(20) & "     (MONEDA EXTRANJERA)       " & oImpresora.gPrnSaltoLinea
   Else
      glsarchivo = glsarchivo & Space(20) & "      (MONEDA NACIONAL)        " & oImpresora.gPrnSaltoLinea
   End If
   glsarchivo = glsarchivo & Space(20) & String(35, "-") & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
   
   glsarchivo = glsarchivo & Space(5) & "ACTIVO" & Space(35) & PrnVal(nActivo, 16, 2) & oImpresora.gPrnSaltoLinea
   
   glsarchivo = glsarchivo & Space(5) & "PASIVO" & Space(55) & PrnVal(nPasivo, 16, 2) & oImpresora.gPrnSaltoLinea
   
   glsarchivo = glsarchivo & Space(5) & "PATRIMONIO" & Space(50) & PrnVal(nPatri, 16, 2) & oImpresora.gPrnSaltoLinea
   
   glsarchivo = glsarchivo & Space(5) & "UTILIDAD (PERDIDA) NETA" & Space(37) & PrnVal(nRei + nUtilidadMes + nDeduccion + nDeduccion1, 16, 2) & oImpresora.gPrnSaltoLinea
   
   glsarchivo = glsarchivo & Space(47) & String(38, "-") & oImpresora.gPrnSaltoLinea
      
   glsarchivo = glsarchivo & Space(45) & PrnVal(nActivo, 16, 2) & Space(5) & PrnVal(nPasivo + nPatri + nRei + nUtilidadMes + nDeduccion + nDeduccion1, 16, 2) & oImpresora.gPrnSaltoLinea
   
   glsarchivo = glsarchivo & Space(47) & String(38, "-") & oImpresora.gPrnSaltoLinea
   
   glsarchivo = glsarchivo & Space(5) & "DIFERENCIA" & Space(30) & PrnVal(nActivo - (nPasivo + nPatri + nRei + nUtilidadMes + nDeduccion + nDeduccion1), 16, 2)

   AgregaUtilidad = glsarchivo
End If


End Function

'*** PEAC 20110412
Public Sub ImprimeBalanceComparativoDelGastoExcel(ByVal psAnio As String, ByVal psMes As String)

Dim pnHayDatos As Integer, lnTotColumnas As Integer
Dim I As Integer, lcMeses As String
Dim j As Integer, lnFila As Integer, lcOrden As String
Dim lnTotFila As Double

Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset

Dim dBalance As New NBalanceCont
Set rs = dBalance.LeeBalanceComparativoGastoAdmOpe(psAnio, psMes)
Set dBalance = Nothing

    If (rs.BOF And rs.EOF) Then
        pnHayDatos = 0
    Else
        pnHayDatos = 1
    End If
        
    If pnHayDatos = 0 Then
     MsgBox "No existe información para este reporte.", vbOKOnly + vbInformation, "Atención"
     lblMsg = ""
     lblMsg.Visible = False
     Exit Sub
    End If

    lblMsg.Visible = True
    lblMsg.Caption = "Generando Archivo Excel, espere por favor...."

    lnTotColumnas = rs.Fields.Count

    Set ApExcel = CreateObject("Excel.application")

    Call CreaBalanceComparativoGastoAdmOpeExcel(psAnio, lnTotColumnas, rs)

    ApExcel.Visible = True
    Set ApExcel = Nothing
    
    lblMsg.Visible = False
    lblMsg.Caption = ""
    
End Sub

Public Sub ImprimeBalanceDetalladoDelGastoExcel(ByVal psAnio As String, ByVal psMes As String)

Dim pnHayDatos As Integer, lnTotColumnas As Integer, lnTotColumnas1 As Integer, lnTotColumnas2 As Integer
Dim I As Integer, lcMeses As String
Dim j As Integer, lnFila As Integer, lcOrden As String
Dim lnTotFila As Double

Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset

Dim rs1 As ADODB.Recordset
Set rs1 = New ADODB.Recordset

Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset


Dim dBalance As New NBalanceCont
Set rs = dBalance.LeeBalanceDetalladoGastoAdmOpe(psAnio, psMes)
Set rs1 = dBalance.BalanceResumenGastosMensuales(psAnio, psMes)
Set rs2 = dBalance.BalanceResumenGastosMensualesDetalle(psAnio, psMes)
Set dBalance = Nothing

    If (rs.BOF And rs.EOF) Then
        pnHayDatos = 0
        Exit Sub
    Else
        pnHayDatos = 1
    End If
        
    If pnHayDatos = 0 Then
     MsgBox "No existe información para este reporte.", vbOKOnly + vbInformation, "Atención"
     lblMsg = ""
     lblMsg.Visible = False
     Exit Sub
    End If

    lblMsg.Visible = True
    lblMsg.Caption = "Generando Archivo Excel, espere por favor...."

    lnTotColumnas = rs.Fields.Count
    lnTotColumnas1 = rs1.Fields.Count
    lnTotColumnas2 = rs2.Fields.Count

    Set ApExcel = CreateObject("Excel.application")

    Call CreaHoja1(psAnio, lnTotColumnas, rs)
    Call CreaHoja2(psAnio, lnTotColumnas1, rs1)
    Call CreaHoja3(psAnio, lnTotColumnas2, rs2)

    ApExcel.Visible = True
    Set ApExcel = Nothing

    lblMsg.Visible = False
    lblMsg.Caption = ""

End Sub

'*** PEAC 20110415
Public Sub CreaHoja1(ByVal psAnio As String, ByVal pnTotColumnas As Integer, rs As ADODB.Recordset)
'CreaBalanceComparativoGastoAdmOpeExcel
Dim I As Integer
Dim lcMeses As String
Dim j As Integer
Dim lnFila As Integer
Dim lcOrden As String
Dim lnTotFila As Double

'Agrega un nuevo Libro
ApExcel.Workbooks.Add
   
'Poner Titulos

ApExcel.Cells(3, 1) = "CAJA MAYNAS S.A."

ApExcel.Cells(3, 2).Interior.Color = RGB(255, 255, 128)
ApExcel.Cells(3, 2).Font.Bold = True
ApExcel.Cells(3, 2).Font.Color = RGB(255, 0, 0)
ApExcel.Cells(3, 2) = "CUADRO 11"

ApExcel.Cells(4, 1) = "RESUMEN DE GASTOS E INGRESOS MENSUALES " & Trim(psAnio)

ApExcel.Cells(8, 1) = "DESCRIPCION"

ApExcel.Cells(6, 2) = "TOTAL"
ApExcel.Cells(7, 2) = "ACUMULADO"
ApExcel.Cells(8, 2) = "DICIEMBRE"
ApExcel.Cells(9, 2) = "'" + CStr((CInt(psAnio) - 1))

ApExcel.Cells(7, 3) = "PROMEDIO"
ApExcel.Cells(8, 3) = "'" + CStr((CInt(psAnio) - 1))

For I = 1 To pnTotColumnas - 5

    lcMeses = Choose(I, "ENERO", "FEBRERO", "MARZO", "ABRIL", _
                            "MAYO", "JUNIO", "JULIO", "AGOSTO", _
                            "SETIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE")
    j = I + 3
    ApExcel.Cells(7, j) = lcMeses
    ApExcel.Cells(8, j) = "'" + psAnio

Next I
I = I + 3
ApExcel.Cells(7, I) = "TOTAL"
ApExcel.Cells(8, I) = "ACUMULADO"
ApExcel.Cells(9, I) = lcMeses & " " & psAnio
I = I + 1
ApExcel.Cells(7, I) = "PROMEDIO"
ApExcel.Cells(8, I) = "'" + psAnio
I = I + 1
ApExcel.Cells(8, I) = "   %"

ApExcel.Range(ApExcel.Cells(6, 1), ApExcel.Cells(9, I)).Cells.Interior.Color = RGB(255, 203, 151)

ApExcel.Range(ApExcel.Cells(12, 1), ApExcel.Cells(12, I)).Cells.Interior.Color = RGB(255, 255, 128)
ApExcel.Range(ApExcel.Cells(26, 1), ApExcel.Cells(26, I)).Cells.Interior.Color = RGB(255, 255, 128)

ApExcel.Range(ApExcel.Cells(14, 1), ApExcel.Cells(14, I)).Cells.Font.Bold = True
ApExcel.Range(ApExcel.Cells(21, 1), ApExcel.Cells(21, I)).Cells.Font.Bold = True
ApExcel.Range(ApExcel.Cells(28, 1), ApExcel.Cells(28, I)).Cells.Font.Bold = True
ApExcel.Range(ApExcel.Cells(33, 1), ApExcel.Cells(33, I)).Cells.Font.Bold = True
ApExcel.Range(ApExcel.Cells(37, 1), ApExcel.Cells(37, I)).Cells.Font.Bold = True

'----------- FIN CABECERA

lnFila = 11

Dim lnArrayVar(1 To 12) As Double
lcOrden = ""
Do While Not rs.EOF
  
   If rs!corden <> lcOrden Then
    lnFila = lnFila + 1
   End If
   
   lcOrden = rs!corden
   
    If lcOrden = "11" And lnFila = 12 Then
        ApExcel.Cells(lnFila, 1) = "PROVISIONES Y DEPRECIACION"
        lnFila = lnFila + 2
    ElseIf lcOrden = "21" And lnFila = 26 Then
        ApExcel.Cells(lnFila, 1) = "OTROS INGRESOS Y GASTOS"
        lnFila = lnFila + 2
    End If
    
    If lcOrden = "11" And lnFila = 14 Then
        ApExcel.Cells(lnFila, 1) = "DEPRECIACION"
        lnFila = lnFila + 1
    ElseIf lcOrden = "12" And lnFila = 21 Then
        ApExcel.Cells(lnFila, 1) = "OTRAS PROVISIONES"
        lnFila = lnFila + 1
    ElseIf lcOrden = "21" And lnFila = 28 Then
        ApExcel.Cells(lnFila, 1) = "INGRESOS NETOS/GASTOS NETOS"
        lnFila = lnFila + 1
    ElseIf lcOrden = "22" And lnFila = 33 Then
        ApExcel.Cells(lnFila, 1) = "OTROS INGRESOS"
        lnFila = lnFila + 1
    ElseIf lcOrden = "23" And lnFila = 37 Then
        ApExcel.Cells(lnFila, 1) = "OTROS GASTOS"
        lnFila = lnFila + 1
    End If
    
    ApExcel.Cells(lnFila, 1) = rs!cCtaContDesc

    ApExcel.Cells(lnFila, 2).NumberFormat = "#,##0"
    ApExcel.Cells(lnFila, 2) = rs!A00
    
    ApExcel.Cells(lnFila, 3).NumberFormat = "#,##0"
    ApExcel.Cells(lnFila, 3) = rs!A00 / 12

    lnTotFila = 0
    For I = 4 To pnTotColumnas - 2
        ApExcel.Cells(lnFila, I).NumberFormat = "#,##0"
        ApExcel.Cells(lnFila, I) = rs.Fields(I)
'        lnArrayVar(I - 3) = lnArrayVar(I - 3) + rs.Fields(I)
        lnTotFila = lnTotFila + rs.Fields(I)
    Next I
    ApExcel.Cells(lnFila, I).NumberFormat = "#,##0"
    ApExcel.Cells(lnFila, I) = lnTotFila ''rs!A13
    I = I + 1
    ApExcel.Cells(lnFila, I).NumberFormat = "#,##0"
    ApExcel.Cells(lnFila, I) = lnTotFila / (pnTotColumnas - 5)
    
    I = I + 1
    ApExcel.Cells(lnFila, I).NumberFormat = "0.00%"
    If Left(lcOrden, 1) = "1" Then
        ApExcel.Cells(lnFila, I).FormulaR1C1 = "=+RC[-1]/R12C" & Trim(Str(I - 1))
    Else
        ApExcel.Cells(lnFila, I).FormulaR1C1 = "=+RC[-1]/R26C" & Trim(Str(I - 1))
    End If
    
    lnFila = lnFila + 1
    rs.MoveNext
Loop

'--- INI SUB TOTALES
ApExcel.Cells(14, 2).NumberFormat = "#,##0"
ApExcel.Cells(14, 2).FormulaR1C1 = "=+SUM(R[1]C:R[5]C)"
ApExcel.Cells(14, 3).NumberFormat = "#,##0"
ApExcel.Cells(14, 3).FormulaR1C1 = "=+SUM(R[1]C:R[5]C)"
For I = 4 To pnTotColumnas - 2
    ApExcel.Cells(14, I).NumberFormat = "#,##0"
    ApExcel.Cells(14, I).FormulaR1C1 = "=+SUM(R[1]C:R[5]C)"
Next I
ApExcel.Cells(14, I).NumberFormat = "#,##0"
ApExcel.Cells(14, I).FormulaR1C1 = "=+SUM(R[1]C:R[5]C)"
I = I + 1
ApExcel.Cells(14, I).NumberFormat = "#,##0"
ApExcel.Cells(14, I).FormulaR1C1 = "=+SUM(R[1]C:R[5]C)"
I = I + 1
ApExcel.Cells(14, I).NumberFormat = "0.00%"
ApExcel.Cells(14, I).FormulaR1C1 = "=+SUM(R[1]C:R[5]C)"


'-------
ApExcel.Cells(21, 2).NumberFormat = "#,##0"
ApExcel.Cells(21, 2).FormulaR1C1 = "=+SUM(R[1]C:R[3]C)"
ApExcel.Cells(21, 3).NumberFormat = "#,##0"
ApExcel.Cells(21, 3).FormulaR1C1 = "=+SUM(R[1]C:R[3]C)"
For I = 4 To pnTotColumnas - 2
    ApExcel.Cells(21, I).NumberFormat = "#,##0"
    ApExcel.Cells(21, I).FormulaR1C1 = "=+SUM(R[1]C:R[3]C)"
Next I
ApExcel.Cells(21, I).NumberFormat = "#,##0"
ApExcel.Cells(21, I).FormulaR1C1 = "=+SUM(R[1]C:R[3]C)"
I = I + 1
ApExcel.Cells(21, I).NumberFormat = "#,##0"
ApExcel.Cells(21, I).FormulaR1C1 = "=+SUM(R[1]C:R[3]C)"
I = I + 1
ApExcel.Cells(21, I).NumberFormat = "0.00%"
ApExcel.Cells(21, I).FormulaR1C1 = "=+SUM(R[1]C:R[3]C)"

'-------
ApExcel.Cells(28, 2).NumberFormat = "#,##0"
ApExcel.Cells(28, 2).FormulaR1C1 = "=+SUM(R[1]C:R[3]C)"
ApExcel.Cells(28, 3).NumberFormat = "#,##0"
ApExcel.Cells(28, 3).FormulaR1C1 = "=+SUM(R[1]C:R[3]C)"
For I = 4 To pnTotColumnas - 2
    ApExcel.Cells(28, I).NumberFormat = "#,##0"
    ApExcel.Cells(28, I).FormulaR1C1 = "=+SUM(R[1]C:R[3]C)"
Next I
ApExcel.Cells(28, I).NumberFormat = "#,##0"
ApExcel.Cells(28, I).FormulaR1C1 = "=+SUM(R[1]C:R[3]C)"
I = I + 1
ApExcel.Cells(28, I).NumberFormat = "#,##0"
ApExcel.Cells(28, I).FormulaR1C1 = "=+SUM(R[1]C:R[3]C)"
I = I + 1
ApExcel.Cells(28, I).NumberFormat = "0.00%"
ApExcel.Cells(28, I).FormulaR1C1 = "=+SUM(R[1]C:R[3]C)"

'-------
ApExcel.Cells(33, 2).NumberFormat = "#,##0"
ApExcel.Cells(33, 2).FormulaR1C1 = "=+SUM(R[1]C:R[2]C)"
ApExcel.Cells(33, 3).NumberFormat = "#,##0"
ApExcel.Cells(33, 3).FormulaR1C1 = "=+SUM(R[1]C:R[2]C)"
For I = 4 To pnTotColumnas - 2
    ApExcel.Cells(33, I).NumberFormat = "#,##0"
    ApExcel.Cells(33, I).FormulaR1C1 = "=+SUM(R[1]C:R[2]C)"
Next I
ApExcel.Cells(33, I).NumberFormat = "#,##0"
ApExcel.Cells(33, I).FormulaR1C1 = "=+SUM(R[1]C:R[2]C)"
I = I + 1
ApExcel.Cells(33, I).NumberFormat = "#,##0"
ApExcel.Cells(33, I).FormulaR1C1 = "=+SUM(R[1]C:R[2]C)"
I = I + 1
ApExcel.Cells(33, I).NumberFormat = "0.00%"
ApExcel.Cells(33, I).FormulaR1C1 = "=+SUM(R[1]C:R[2]C)"

'-------
ApExcel.Cells(37, 2).NumberFormat = "#,##0"
ApExcel.Cells(37, 2).FormulaR1C1 = "=+SUM(R[1]C:R[2]C)"
ApExcel.Cells(37, 3).NumberFormat = "#,##0"
ApExcel.Cells(37, 3).FormulaR1C1 = "=+SUM(R[1]C:R[2]C)"
For I = 4 To pnTotColumnas - 2
    ApExcel.Cells(37, I).NumberFormat = "#,##0"
    ApExcel.Cells(37, I).FormulaR1C1 = "=+SUM(R[1]C:R[2]C)"
Next I
ApExcel.Cells(37, I).NumberFormat = "#,##0"
ApExcel.Cells(37, I).FormulaR1C1 = "=+SUM(R[1]C:R[2]C)"
I = I + 1
ApExcel.Cells(37, I).NumberFormat = "#,##0"
ApExcel.Cells(37, I).FormulaR1C1 = "=+SUM(R[1]C:R[2]C)"
I = I + 1
ApExcel.Cells(37, I).NumberFormat = "0.00%"
ApExcel.Cells(37, I).FormulaR1C1 = "=+SUM(R[1]C:R[2]C)"

'--- FIN SUB TOTALES

'--- INI TOTALES
ApExcel.Cells(12, 2).NumberFormat = "#,##0"
ApExcel.Cells(12, 2).FormulaR1C1 = "=+R[2]C+R[9]C"
ApExcel.Cells(12, 3).NumberFormat = "#,##0"
ApExcel.Cells(12, 3).FormulaR1C1 = "=+R[2]C+R[9]C"
For I = 4 To pnTotColumnas - 2
    ApExcel.Cells(12, I).NumberFormat = "#,##0"
    ApExcel.Cells(12, I).FormulaR1C1 = "=+R[2]C+R[9]C"
Next I
ApExcel.Cells(12, I).NumberFormat = "#,##0"
ApExcel.Cells(12, I).FormulaR1C1 = "=+R[2]C+R[9]C"
I = I + 1
ApExcel.Cells(12, I).NumberFormat = "#,##0"
ApExcel.Cells(12, I).FormulaR1C1 = "=+R[2]C+R[9]C"
I = I + 1
ApExcel.Cells(12, I).NumberFormat = "0.00%"
ApExcel.Cells(12, I).FormulaR1C1 = "=+R[2]C+R[9]C"

'-----------

ApExcel.Cells(26, 2).NumberFormat = "#,##0"
ApExcel.Cells(26, 2).FormulaR1C1 = "=+R[2]C+R[7]C+R[11]C"
ApExcel.Cells(26, 3).NumberFormat = "#,##0"
ApExcel.Cells(26, 3).FormulaR1C1 = "=+R[2]C+R[7]C+R[11]C"
For I = 4 To pnTotColumnas - 2
    ApExcel.Cells(26, I).NumberFormat = "#,##0"
    ApExcel.Cells(26, I).FormulaR1C1 = "=+R[2]C+R[7]C+R[11]C"
Next I
ApExcel.Cells(26, I).NumberFormat = "#,##0"
ApExcel.Cells(26, I).FormulaR1C1 = "=+R[2]C+R[7]C+R[11]C"
I = I + 1
ApExcel.Cells(26, I).NumberFormat = "#,##0"
ApExcel.Cells(26, I).FormulaR1C1 = "=+R[2]C+R[7]C+R[11]C"
I = I + 1
ApExcel.Cells(26, I).NumberFormat = "0.00%"
ApExcel.Cells(26, I).FormulaR1C1 = "=+R[2]C+R[7]C+R[11]C"


'--- FIN TOTALES

ApExcel.Cells(42, 1) = "Fuente: Balance de Comprobación Consolidado"

ApExcel.Cells.Select
ApExcel.Cells.EntireColumn.AutoFit

ApExcel.Cells.Select
ApExcel.Cells.Font.Size = 8

ApExcel.Cells.Range("A1").Select

End Sub
'*** PEAC 20110415
Public Sub CreaHoja2(ByVal psAnio As String, ByVal lnTotColumnas As Integer, rs As ADODB.Recordset)

Dim I As Integer
Dim lcMeses As String
Dim j As Integer
Dim lnFila As Integer
Dim lcOrden As String
Dim lnTotFila As Double

ApExcel.Sheets("Hoja2").Select

ApExcel.Cells(3, 1) = "CAJA MAYNAS S.A."

ApExcel.Cells(3, 2).Interior.Color = RGB(255, 255, 128)
ApExcel.Cells(3, 2).Font.Bold = True
ApExcel.Cells(3, 2).Font.Color = RGB(255, 0, 0)
ApExcel.Cells(3, 2) = "CUADRO 10"

ApExcel.Cells(4, 1) = "RESUMEN DE GASTOS MENSUALES " & Trim(psAnio)

ApExcel.Cells(8, 1) = "DESCRIPCION"

ApExcel.Cells(6, 2) = "TOTAL"
ApExcel.Cells(7, 2) = "ACUMULADO"
ApExcel.Cells(8, 2) = "DICIEMBRE"
ApExcel.Cells(9, 2) = "'" + CStr((CInt(psAnio) - 1))

ApExcel.Cells(7, 3) = "PROMEDIO"
ApExcel.Cells(8, 3) = "'" + CStr((CInt(psAnio) - 1))

For I = 1 To lnTotColumnas - 5

    lcMeses = Choose(I, "ENERO", "FEBRERO", "MARZO", "ABRIL", _
                            "MAYO", "JUNIO", "JULIO", "AGOSTO", _
                            "SETIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE")
    j = I + 3
    ApExcel.Cells(7, j) = lcMeses
    ApExcel.Cells(8, j) = "'" + psAnio

Next I
I = I + 3
ApExcel.Cells(7, I) = "TOTAL"
ApExcel.Cells(8, I) = "ACUMULADO"
ApExcel.Cells(9, I) = lcMeses & " " & psAnio
I = I + 1
ApExcel.Cells(7, I) = "PROMEDIO"
ApExcel.Cells(8, I) = "'" + psAnio
I = I + 1
ApExcel.Cells(8, I) = "   %"

ApExcel.Range(ApExcel.Cells(6, 1), ApExcel.Cells(9, I)).Cells.Interior.Color = RGB(255, 203, 151)

ApExcel.Range(ApExcel.Cells(12, 1), ApExcel.Cells(12, I)).Cells.Interior.Color = RGB(255, 255, 128)
ApExcel.Range(ApExcel.Cells(31, 1), ApExcel.Cells(31, I)).Cells.Interior.Color = RGB(255, 255, 128)
ApExcel.Range(ApExcel.Cells(66, 1), ApExcel.Cells(66, I)).Cells.Interior.Color = RGB(255, 255, 128)

ApExcel.Range(ApExcel.Cells(14, 1), ApExcel.Cells(14, I)).Cells.Font.Bold = True
ApExcel.Range(ApExcel.Cells(15, 1), ApExcel.Cells(15, I)).Cells.Font.Bold = True
ApExcel.Range(ApExcel.Cells(16, 1), ApExcel.Cells(16, I)).Cells.Font.Bold = True
ApExcel.Range(ApExcel.Cells(24, 1), ApExcel.Cells(24, I)).Cells.Font.Bold = True
ApExcel.Range(ApExcel.Cells(33, 1), ApExcel.Cells(33, I)).Cells.Font.Bold = True
ApExcel.Range(ApExcel.Cells(57, 1), ApExcel.Cells(57, I)).Cells.Font.Bold = True
ApExcel.Range(ApExcel.Cells(63, 1), ApExcel.Cells(63, I)).Cells.Font.Bold = True

'----------- FIN CABECERA

lnFila = 11

Dim lnArrayVar(1 To 12) As Double
lcOrden = ""
Do While Not rs.EOF
  
   If rs!corden <> lcOrden Then
    lnFila = lnFila + 1
   End If
   
   lcOrden = rs!corden
   
    If lnFila = 12 Then
        ApExcel.Cells(lnFila, 1) = "GASTOS DE ADMINISTRACION"
        lnFila = lnFila + 2
    ElseIf lnFila = 31 Then
        ApExcel.Cells(lnFila, 1) = "GASTOS OPERATIVOS"
        lnFila = lnFila + 2
    ElseIf lnFila = 66 Then
        ApExcel.Cells(lnFila, 1) = "GASTOS DEL DIRECTORIO"
        lnFila = lnFila + 2
    End If
    
    If lnFila = 14 Then
        ApExcel.Cells(lnFila, 1) = "GASTOS DE PERSONAL"
        lnFila = lnFila + 1
        ApExcel.Cells(lnFila, 1) = "Número de Personas"
        lnFila = lnFila + 1
        ApExcel.Cells(lnFila, 1) = "Remuneraciones"
        lnFila = lnFila + 1
        
'    ElseIf lnFila = 16 Then
'        ApExcel.Cells(lnFila, 1) = "Remuneraciones"
'        lnFila = lnFila + 1
        
    ElseIf lnFila = 24 Then
        ApExcel.Cells(lnFila, 1) = "Otros Gastos del Personal"
        lnFila = lnFila + 1
    ElseIf lnFila = 33 Then
        ApExcel.Cells(lnFila, 1) = "Servicios Recibidos de Terceros"
        lnFila = lnFila + 1
    ElseIf lnFila = 57 Then
        ApExcel.Cells(lnFila, 1) = "Tributos y Contribuciones"
        lnFila = lnFila + 1
    ElseIf lnFila = 63 Then
        ApExcel.Cells(lnFila, 1) = "Gastos Ejercicios Anteriores"
        lnFila = lnFila + 1
    End If
    
    ApExcel.Cells(lnFila, 1) = rs!cCtaContDesc

    ApExcel.Cells(lnFila, 2).NumberFormat = "#,##0"
    ApExcel.Cells(lnFila, 2) = rs!A00
    
    ApExcel.Cells(lnFila, 3).NumberFormat = "#,##0"
    ApExcel.Cells(lnFila, 3) = rs!A00 / 12

    lnTotFila = 0
    For I = 4 To lnTotColumnas - 2
        ApExcel.Cells(lnFila, I).NumberFormat = "#,##0"
        ApExcel.Cells(lnFila, I) = rs.Fields(I)
        lnTotFila = lnTotFila + rs.Fields(I)
    Next I
    ApExcel.Cells(lnFila, I).NumberFormat = "#,##0"
    ApExcel.Cells(lnFila, I) = lnTotFila
    I = I + 1
    ApExcel.Cells(lnFila, I).NumberFormat = "#,##0"
    ApExcel.Cells(lnFila, I) = lnTotFila / (lnTotColumnas - 5)
    
    I = I + 1
    ApExcel.Cells(lnFila, I).NumberFormat = "0.00%"
    If Left(lcOrden, 1) = "1" Then
        ApExcel.Cells(lnFila, I).FormulaR1C1 = "=+RC[-1]/R14C" & Trim(Str(I - 1))
    ElseIf Left(lcOrden, 1) = "2" Then
        ApExcel.Cells(lnFila, I).FormulaR1C1 = "=+RC[-1]/R31C" & Trim(Str(I - 1))
    ElseIf Left(lcOrden, 1) = "3" Then
        ApExcel.Cells(lnFila, I).FormulaR1C1 = "=+RC[-1]/R66C" & Trim(Str(I - 1))
    End If
    
    lnFila = lnFila + 1
    rs.MoveNext
Loop

'--- INI SUB TOTALES
ApExcel.Cells(16, 2).NumberFormat = "#,##0"
ApExcel.Cells(16, 2).FormulaR1C1 = "=+SUM(R[1]C:R[6]C)"
ApExcel.Cells(16, 3).NumberFormat = "#,##0"
ApExcel.Cells(16, 3).FormulaR1C1 = "=+SUM(R[1]C:R[6]C)"
For I = 4 To lnTotColumnas - 2
    ApExcel.Cells(16, I).NumberFormat = "#,##0"
    ApExcel.Cells(16, I).FormulaR1C1 = "=+SUM(R[1]C:R[6]C)"
Next I
ApExcel.Cells(16, I).NumberFormat = "#,##0"
ApExcel.Cells(16, I).FormulaR1C1 = "=+SUM(R[1]C:R[6]C)"
I = I + 1
ApExcel.Cells(16, I).NumberFormat = "#,##0"
ApExcel.Cells(16, I).FormulaR1C1 = "=+SUM(R[1]C:R[6]C)"
I = I + 1
ApExcel.Cells(16, I).NumberFormat = "0.00%"
ApExcel.Cells(16, I).FormulaR1C1 = "=+SUM(R[1]C:R[6]C)"

'-------
ApExcel.Cells(24, 2).NumberFormat = "#,##0"
ApExcel.Cells(24, 2).FormulaR1C1 = "=+SUM(R[1]C:R[5]C)"
ApExcel.Cells(24, 3).NumberFormat = "#,##0"
ApExcel.Cells(24, 3).FormulaR1C1 = "=+SUM(R[1]C:R[5]C)"
For I = 4 To lnTotColumnas - 2
    ApExcel.Cells(24, I).NumberFormat = "#,##0"
    ApExcel.Cells(24, I).FormulaR1C1 = "=+SUM(R[1]C:R[5]C)"
Next I
ApExcel.Cells(24, I).NumberFormat = "#,##0"
ApExcel.Cells(24, I).FormulaR1C1 = "=+SUM(R[1]C:R[5]C)"
I = I + 1
ApExcel.Cells(24, I).NumberFormat = "#,##0"
ApExcel.Cells(24, I).FormulaR1C1 = "=+SUM(R[1]C:R[5]C)"
I = I + 1
ApExcel.Cells(24, I).NumberFormat = "0.00%"
ApExcel.Cells(24, I).FormulaR1C1 = "=+SUM(R[1]C:R[5]C)"

'-------
ApExcel.Cells(33, 2).NumberFormat = "#,##0"
ApExcel.Cells(33, 2).FormulaR1C1 = "=+SUM(R[1]C:R[22]C)"
ApExcel.Cells(33, 3).NumberFormat = "#,##0"
ApExcel.Cells(33, 3).FormulaR1C1 = "=+SUM(R[1]C:R[22]C)"
For I = 4 To lnTotColumnas - 2
    ApExcel.Cells(33, I).NumberFormat = "#,##0"
    ApExcel.Cells(33, I).FormulaR1C1 = "=+SUM(R[1]C:R[22]C)"
Next I
ApExcel.Cells(33, I).NumberFormat = "#,##0"
ApExcel.Cells(33, I).FormulaR1C1 = "=+SUM(R[1]C:R[22]C)"
I = I + 1
ApExcel.Cells(33, I).NumberFormat = "#,##0"
ApExcel.Cells(33, I).FormulaR1C1 = "=+SUM(R[1]C:R[22]C)"
I = I + 1
ApExcel.Cells(33, I).NumberFormat = "0.00%"
ApExcel.Cells(33, I).FormulaR1C1 = "=+SUM(R[1]C:R[22]C)"

'-------
ApExcel.Cells(57, 2).NumberFormat = "#,##0"
ApExcel.Cells(57, 2).FormulaR1C1 = "=+SUM(R[1]C:R[4]C)"
ApExcel.Cells(57, 3).NumberFormat = "#,##0"
ApExcel.Cells(57, 3).FormulaR1C1 = "=+SUM(R[1]C:R[4]C)"
For I = 4 To lnTotColumnas - 2
    ApExcel.Cells(57, I).NumberFormat = "#,##0"
    ApExcel.Cells(57, I).FormulaR1C1 = "=+SUM(R[1]C:R[4]C)"
Next I
ApExcel.Cells(57, I).NumberFormat = "#,##0"
ApExcel.Cells(57, I).FormulaR1C1 = "=+SUM(R[1]C:R[4]C)"
I = I + 1
ApExcel.Cells(57, I).NumberFormat = "#,##0"
ApExcel.Cells(57, I).FormulaR1C1 = "=+SUM(R[1]C:R[4]C)"
I = I + 1
ApExcel.Cells(57, I).NumberFormat = "0.00%"
ApExcel.Cells(57, I).FormulaR1C1 = "=+SUM(R[1]C:R[4]C)"

'-------
ApExcel.Cells(63, 2).NumberFormat = "#,##0"
ApExcel.Cells(63, 2).FormulaR1C1 = "=+SUM(R[1]C:R[1]C)"
ApExcel.Cells(63, 3).NumberFormat = "#,##0"
ApExcel.Cells(63, 3).FormulaR1C1 = "=+SUM(R[1]C:R[1]C)"
For I = 4 To lnTotColumnas - 2
    ApExcel.Cells(63, I).NumberFormat = "#,##0"
    ApExcel.Cells(63, I).FormulaR1C1 = "=+SUM(R[1]C:R[1]C)"
Next I
ApExcel.Cells(63, I).NumberFormat = "#,##0"
ApExcel.Cells(63, I).FormulaR1C1 = "=+SUM(R[1]C:R[1]C)"
I = I + 1
ApExcel.Cells(63, I).NumberFormat = "#,##0"
ApExcel.Cells(63, I).FormulaR1C1 = "=+SUM(R[1]C:R[1]C)"
I = I + 1
ApExcel.Cells(63, I).NumberFormat = "0.00%"
ApExcel.Cells(63, I).FormulaR1C1 = "=+SUM(R[1]C:R[1]C)"

'-------
ApExcel.Cells(66, 2).NumberFormat = "#,##0"
ApExcel.Cells(66, 2).FormulaR1C1 = "=+SUM(R[2]C:R[4]C)"
ApExcel.Cells(66, 3).NumberFormat = "#,##0"
ApExcel.Cells(66, 3).FormulaR1C1 = "=+SUM(R[2]C:R[4]C)"
For I = 4 To lnTotColumnas - 2
    ApExcel.Cells(66, I).NumberFormat = "#,##0"
    ApExcel.Cells(66, I).FormulaR1C1 = "=+SUM(R[2]C:R[4]C)"
Next I
ApExcel.Cells(66, I).NumberFormat = "#,##0"
ApExcel.Cells(66, I).FormulaR1C1 = "=+SUM(R[2]C:R[4]C)"
I = I + 1
ApExcel.Cells(66, I).NumberFormat = "#,##0"
ApExcel.Cells(66, I).FormulaR1C1 = "=+SUM(R[2]C:R[4]C)"
I = I + 1
ApExcel.Cells(66, I).NumberFormat = "0.00%"
ApExcel.Cells(66, I).FormulaR1C1 = "=+SUM(R[2]C:R[4]C)"

'--- FIN SUB TOTALES

'--- INI PRIMER TOTALES
ApExcel.Cells(14, 2).NumberFormat = "#,##0"
ApExcel.Cells(14, 2).FormulaR1C1 = "=+R[2]C+R[10]C"
ApExcel.Cells(14, 3).NumberFormat = "#,##0"
ApExcel.Cells(14, 3).FormulaR1C1 = "=+R[2]C+R[10]C"
For I = 4 To lnTotColumnas - 2
    ApExcel.Cells(14, I).NumberFormat = "#,##0"
    ApExcel.Cells(14, I).FormulaR1C1 = "=+R[2]C+R[10]C"
Next I
ApExcel.Cells(14, I).NumberFormat = "#,##0"
ApExcel.Cells(14, I).FormulaR1C1 = "=+R[2]C+R[10]C"
I = I + 1
ApExcel.Cells(14, I).NumberFormat = "#,##0"
ApExcel.Cells(14, I).FormulaR1C1 = "=+R[2]C+R[10]C"
I = I + 1
ApExcel.Cells(14, I).NumberFormat = "0.00%"
ApExcel.Cells(14, I).FormulaR1C1 = "=+R[2]C+R[10]C"
'--- FIN PRIMER TOTALES

'--- INI TOTALES
ApExcel.Cells(12, 2).NumberFormat = "#,##0"
ApExcel.Cells(12, 2).FormulaR1C1 = "=+R[2]C+R[19]C+R[54]C"
ApExcel.Cells(12, 3).NumberFormat = "#,##0"
ApExcel.Cells(12, 3).FormulaR1C1 = "=+R[2]C+R[19]C+R[54]C"
For I = 4 To lnTotColumnas - 2
    ApExcel.Cells(12, I).NumberFormat = "#,##0"
    ApExcel.Cells(12, I).FormulaR1C1 = "=+R[2]C+R[19]C+R[54]C"
Next I
ApExcel.Cells(12, I).NumberFormat = "#,##0"
ApExcel.Cells(12, I).FormulaR1C1 = "=+R[2]C+R[19]C+R[54]C"
I = I + 1
ApExcel.Cells(12, I).NumberFormat = "#,##0"
ApExcel.Cells(12, I).FormulaR1C1 = "=+R[2]C+R[19]C+R[54]C"
I = I + 1
ApExcel.Cells(12, I).NumberFormat = "0.00%"
ApExcel.Cells(12, I).FormulaR1C1 = "=+R[2]C"
'----------------
ApExcel.Cells(31, 2).NumberFormat = "#,##0"
ApExcel.Cells(31, 2).FormulaR1C1 = "=+R[2]C+R[26]C+R[32]C"
ApExcel.Cells(31, 3).NumberFormat = "#,##0"
ApExcel.Cells(31, 3).FormulaR1C1 = "=+R[2]C+R[26]C+R[32]C"
For I = 4 To lnTotColumnas - 2
    ApExcel.Cells(31, I).NumberFormat = "#,##0"
    ApExcel.Cells(31, I).FormulaR1C1 = "=+R[2]C+R[26]C+R[32]C"
Next I
ApExcel.Cells(31, I).NumberFormat = "#,##0"
ApExcel.Cells(31, I).FormulaR1C1 = "=+R[2]C+R[26]C+R[32]C"
I = I + 1
ApExcel.Cells(31, I).NumberFormat = "#,##0"
ApExcel.Cells(31, I).FormulaR1C1 = "=+R[2]C+R[26]C+R[32]C"
I = I + 1
ApExcel.Cells(31, I).NumberFormat = "0.00%"
ApExcel.Cells(31, I).FormulaR1C1 = "=+R[2]C+R[26]C+R[32]C"

'--- FIN TOTALES

ApExcel.Cells(74, 1) = "Fuente: Balance de Comprobación Consolidado"

ApExcel.Cells.Select
ApExcel.Cells.EntireColumn.AutoFit

ApExcel.Cells.Select
ApExcel.Cells.Font.Size = 8

ApExcel.Cells.Range("A1").Select

End Sub

'*** PEAC 20110415
Public Sub CreaHoja3(ByVal psAnio As String, ByVal lnTotColumnas As Integer, rs As ADODB.Recordset)
    
Dim I As Integer
Dim lcMeses As String
Dim j As Integer
Dim lnFila As Integer
Dim lcOrden As String
Dim lnTotFila As Double

ApExcel.Sheets("Hoja3").Select

ApExcel.Cells(3, 1) = "CAJA MAYNAS S.A."

ApExcel.Cells(3, 2).Interior.Color = RGB(255, 255, 128)
ApExcel.Cells(3, 2).Font.Bold = True
ApExcel.Cells(3, 2).Font.Color = RGB(255, 0, 0)
ApExcel.Cells(3, 2) = "CUADRO 3.5"

ApExcel.Cells(4, 1) = "RESUMEN DE GASTOS MENSUALES " & Trim(psAnio)

ApExcel.Cells(8, 1) = "DESCRIPCION"

ApExcel.Cells(6, 2) = "TOTAL"
ApExcel.Cells(7, 2) = "ACUMULADO"
ApExcel.Cells(8, 2) = "DICIEMBRE"
ApExcel.Cells(9, 2) = "'" + CStr((CInt(psAnio) - 1))

ApExcel.Cells(7, 3) = "PROMEDIO"
ApExcel.Cells(8, 3) = "'" + CStr((CInt(psAnio) - 1))

For I = 1 To lnTotColumnas - 5

    lcMeses = Choose(I, "ENERO", "FEBRERO", "MARZO", "ABRIL", _
                            "MAYO", "JUNIO", "JULIO", "AGOSTO", _
                            "SETIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE")
    j = I + 3
    ApExcel.Cells(7, j) = lcMeses
    ApExcel.Cells(8, j) = "'" + psAnio

Next I
I = I + 3
ApExcel.Cells(7, I) = "TOTAL"
ApExcel.Cells(8, I) = "ACUMULADO"
ApExcel.Cells(9, I) = lcMeses & " " & psAnio
I = I + 1
ApExcel.Cells(7, I) = "PROMEDIO"
ApExcel.Cells(8, I) = "'" + psAnio
I = I + 1
ApExcel.Cells(8, I) = "   %"

'--cabecera
ApExcel.Range(ApExcel.Cells(6, 1), ApExcel.Cells(9, I)).Cells.Interior.Color = RGB(255, 203, 151)

''--titulos color amarillo
ApExcel.Range(ApExcel.Cells(12, 1), ApExcel.Cells(12, I)).Cells.Interior.Color = RGB(255, 255, 128)
ApExcel.Range(ApExcel.Cells(40, 1), ApExcel.Cells(40, I)).Cells.Interior.Color = RGB(255, 255, 128)
ApExcel.Range(ApExcel.Cells(99, 1), ApExcel.Cells(99, I)).Cells.Interior.Color = RGB(255, 255, 128)

''-- negrita
ApExcel.Range(ApExcel.Cells(14, 1), ApExcel.Cells(16, I)).Cells.Font.Bold = True
ApExcel.Range(ApExcel.Cells(27, 1), ApExcel.Cells(27, I)).Cells.Font.Bold = True
ApExcel.Range(ApExcel.Cells(42, 1), ApExcel.Cells(42, I)).Cells.Font.Bold = True
ApExcel.Range(ApExcel.Cells(90, 1), ApExcel.Cells(90, I)).Cells.Font.Bold = True
ApExcel.Range(ApExcel.Cells(96, 1), ApExcel.Cells(96, I)).Cells.Font.Bold = True
'
''-- letras rojas
ApExcel.Range(ApExcel.Cells(18, 1), ApExcel.Cells(19, 2)).Cells.Font.Color = RGB(255, 0, 0)
ApExcel.Range(ApExcel.Cells(34, 1), ApExcel.Cells(38, 2)).Cells.Font.Color = RGB(255, 0, 0)
ApExcel.Range(ApExcel.Cells(44, 1), ApExcel.Cells(45, 2)).Cells.Font.Color = RGB(255, 0, 0)
ApExcel.Range(ApExcel.Cells(48, 1), ApExcel.Cells(49, 2)).Cells.Font.Color = RGB(255, 0, 0)
ApExcel.Range(ApExcel.Cells(52, 1), ApExcel.Cells(55, 2)).Cells.Font.Color = RGB(255, 0, 0)
ApExcel.Range(ApExcel.Cells(58, 1), ApExcel.Cells(60, 2)).Cells.Font.Color = RGB(255, 0, 0)
ApExcel.Range(ApExcel.Cells(63, 1), ApExcel.Cells(64, 2)).Cells.Font.Color = RGB(255, 0, 0)
ApExcel.Range(ApExcel.Cells(67, 1), ApExcel.Cells(68, 2)).Cells.Font.Color = RGB(255, 0, 0)
ApExcel.Range(ApExcel.Cells(71, 1), ApExcel.Cells(72, 2)).Cells.Font.Color = RGB(255, 0, 0)

'----------- FIN CABECERA

lnFila = 11

Dim lnArrayVar(1 To 12) As Double
lcOrden = ""
Do While Not rs.EOF
  
    If rs!corden <> lcOrden Then
        lnFila = lnFila + 1
    End If
   
    If lnFila = 12 Then
        ApExcel.Cells(lnFila, 1) = "GASTOS DE ADMINISTRACION"
        lnFila = lnFila + 2
    ElseIf lnFila = 40 Then
        ApExcel.Cells(lnFila, 1) = "GASTOS OPERATIVOS"
        lnFila = lnFila + 2
    ElseIf lnFila = 99 Then
        ApExcel.Cells(lnFila, 1) = "GASTOS DEL DIRECTORIO"
        lnFila = lnFila + 2
    End If

    If lnFila = 14 Then
        ApExcel.Cells(lnFila, 1) = "GASTOS DE PERSONAL"
        lnFila = lnFila + 1
        ApExcel.Cells(lnFila, 1) = "Número de Personas"
        lnFila = lnFila + 1
        ApExcel.Cells(lnFila, 1) = "Remuneraciones"
        lnFila = lnFila + 1
        ApExcel.Cells(lnFila, 1) = "Fijas (Básico/Gratificación/Asignac. Famil.)"
        lnFila = lnFila + 1
    ElseIf lnFila = 27 Then
        ApExcel.Cells(lnFila, 1) = "Otros Gastos del Personal"
        lnFila = lnFila + 1
    ElseIf lnFila = 33 Then
        ApExcel.Cells(lnFila, 1) = "Otros gastos personal (Atenciones,canasta,ref,etc)"
        lnFila = lnFila + 1
    ElseIf lnFila = 42 Then
        ApExcel.Cells(lnFila, 1) = "Servicios Recibidos de Terceros"
        lnFila = lnFila + 1
        ApExcel.Cells(lnFila, 1) = "Publicidad, Campañas y Gastos de Promoción"
        lnFila = lnFila + 1
    ElseIf lnFila = 47 Then
        ApExcel.Cells(lnFila, 1) = "Comunicaciones"
        lnFila = lnFila + 1
    ElseIf lnFila = 51 Then
        ApExcel.Cells(lnFila, 1) = "HHPP/ Consultorías / Est.y Proy/Soc.Audit."
        lnFila = lnFila + 1
    ElseIf lnFila = 57 Then
        ApExcel.Cells(lnFila, 1) = "Utiles de Oficina, Impresos y Bienes no Deprec."
        lnFila = lnFila + 1
    ElseIf lnFila = 62 Then
        ApExcel.Cells(lnFila, 1) = "Transporte+ TUUA"
        lnFila = lnFila + 1
    ElseIf lnFila = 66 Then
        ApExcel.Cells(lnFila, 1) = "Otros Servicios y Gastos"
        lnFila = lnFila + 1
    ElseIf lnFila = 70 Then
        ApExcel.Cells(lnFila, 1) = "Gastos Judiciales, Notariales y de registro"
        lnFila = lnFila + 1
    ElseIf lnFila = 90 Then
        ApExcel.Cells(lnFila, 1) = "Tributos y Contribuciones"
        lnFila = lnFila + 1
    ElseIf lnFila = 96 Then
        ApExcel.Cells(lnFila, 1) = "Gastos Ejercicios Anteriores"
        lnFila = lnFila + 1
    End If

    ApExcel.Cells(lnFila, 1) = rs!cCtaContDesc

    ApExcel.Cells(lnFila, 2).NumberFormat = "#,##0"
    ApExcel.Cells(lnFila, 2) = rs!A00
    
    ApExcel.Cells(lnFila, 3).NumberFormat = "#,##0"
    ApExcel.Cells(lnFila, 3) = rs!A00 / 12

    lnTotFila = 0
    For I = 4 To lnTotColumnas - 2
        ApExcel.Cells(lnFila, I).NumberFormat = "#,##0"
        ApExcel.Cells(lnFila, I) = rs.Fields(I)
        lnTotFila = lnTotFila + rs.Fields(I)
    Next I
    ApExcel.Cells(lnFila, I).NumberFormat = "#,##0"
    ApExcel.Cells(lnFila, I) = lnTotFila
    I = I + 1
    ApExcel.Cells(lnFila, I).NumberFormat = "#,##0"
    ApExcel.Cells(lnFila, I) = lnTotFila / (lnTotColumnas - 5)
    
    I = I + 1
'    ApExcel.Cells(lnFila, I).NumberFormat = "0.00%"
'    If Left(lcOrden, 1) = "1" Then
'        ApExcel.Cells(lnFila, I).FormulaR1C1 = "=+RC[-1]/R14C" & Trim(Str(I - 1))
'    ElseIf Left(lcOrden, 1) = "2" Then
'        ApExcel.Cells(lnFila, I).FormulaR1C1 = "=+RC[-1]/R40C" & Trim(Str(I - 1))
'    ElseIf Left(lcOrden, 1) = "3" Then
'        ApExcel.Cells(lnFila, I).FormulaR1C1 = "=+RC[-1]/R99C" & Trim(Str(I - 1))
'    End If
'
    lcOrden = rs!corden
    
    lnFila = lnFila + 1
    rs.MoveNext
Loop

''--- INI TOTALES DETALLES

ApExcel.Cells(17, 2).NumberFormat = "#,##0"
ApExcel.Cells(17, 2).FormulaR1C1 = "=+SUM(R[1]C:R[2]C)"
ApExcel.Cells(17, 3).NumberFormat = "#,##0"
ApExcel.Cells(17, 3).FormulaR1C1 = "=+SUM(R[1]C:R[2]C)"
For I = 4 To lnTotColumnas - 2
    ApExcel.Cells(17, I).NumberFormat = "#,##0"
    ApExcel.Cells(17, I).FormulaR1C1 = "=+SUM(R[1]C:R[2]C)"
Next I
ApExcel.Cells(17, I).NumberFormat = "#,##0"
ApExcel.Cells(17, I).FormulaR1C1 = "=+SUM(R[1]C:R[2]C)"
I = I + 1
ApExcel.Cells(17, I).NumberFormat = "#,##0"
ApExcel.Cells(17, I).FormulaR1C1 = "=+SUM(R[1]C:R[2]C)"
I = I + 1
ApExcel.Cells(17, I).NumberFormat = "0.00%"
ApExcel.Cells(17, I).FormulaR1C1 = "=+SUM(R[1]C:R[2]C)"
'-------------

ApExcel.Cells(33, 2).NumberFormat = "#,##0"
ApExcel.Cells(33, 2).FormulaR1C1 = "=+SUM(R[1]C:R[5]C)"
ApExcel.Cells(33, 3).NumberFormat = "#,##0"
ApExcel.Cells(33, 3).FormulaR1C1 = "=+SUM(R[1]C:R[5]C)"
For I = 4 To lnTotColumnas - 2
    ApExcel.Cells(33, I).NumberFormat = "#,##0"
    ApExcel.Cells(33, I).FormulaR1C1 = "=+SUM(R[1]C:R[5]C)"
Next I
ApExcel.Cells(33, I).NumberFormat = "#,##0"
ApExcel.Cells(33, I).FormulaR1C1 = "=+SUM(R[1]C:R[5]C)"
I = I + 1
ApExcel.Cells(33, I).NumberFormat = "#,##0"
ApExcel.Cells(33, I).FormulaR1C1 = "=+SUM(R[1]C:R[5]C)"
I = I + 1
ApExcel.Cells(33, I).NumberFormat = "0.00%"
ApExcel.Cells(33, I).FormulaR1C1 = "=+SUM(R[1]C:R[5]C)"
'-------------

ApExcel.Cells(43, 2).NumberFormat = "#,##0"
ApExcel.Cells(43, 2).FormulaR1C1 = "=+SUM(R[1]C:R[2]C)"
ApExcel.Cells(43, 3).NumberFormat = "#,##0"
ApExcel.Cells(43, 3).FormulaR1C1 = "=+SUM(R[1]C:R[2]C)"
For I = 4 To lnTotColumnas - 2
    ApExcel.Cells(43, I).NumberFormat = "#,##0"
    ApExcel.Cells(43, I).FormulaR1C1 = "=+SUM(R[1]C:R[2]C)"
Next I
ApExcel.Cells(43, I).NumberFormat = "#,##0"
ApExcel.Cells(43, I).FormulaR1C1 = "=+SUM(R[1]C:R[2]C)"
I = I + 1
ApExcel.Cells(43, I).NumberFormat = "#,##0"
ApExcel.Cells(43, I).FormulaR1C1 = "=+SUM(R[1]C:R[2]C)"
I = I + 1
ApExcel.Cells(43, I).NumberFormat = "0.00%"
ApExcel.Cells(43, I).FormulaR1C1 = "=+SUM(R[1]C:R[2]C)"
'-------------

ApExcel.Cells(47, 2).NumberFormat = "#,##0"
ApExcel.Cells(47, 2).FormulaR1C1 = "=+SUM(R[1]C:R[2]C)"
ApExcel.Cells(47, 3).NumberFormat = "#,##0"
ApExcel.Cells(47, 3).FormulaR1C1 = "=+SUM(R[1]C:R[2]C)"
For I = 4 To lnTotColumnas - 2
    ApExcel.Cells(47, I).NumberFormat = "#,##0"
    ApExcel.Cells(47, I).FormulaR1C1 = "=+SUM(R[1]C:R[2]C)"
Next I
ApExcel.Cells(47, I).NumberFormat = "#,##0"
ApExcel.Cells(47, I).FormulaR1C1 = "=+SUM(R[1]C:R[2]C)"
I = I + 1
ApExcel.Cells(47, I).NumberFormat = "#,##0"
ApExcel.Cells(47, I).FormulaR1C1 = "=+SUM(R[1]C:R[2]C)"
I = I + 1
ApExcel.Cells(47, I).NumberFormat = "0.00%"
ApExcel.Cells(47, I).FormulaR1C1 = "=+SUM(R[1]C:R[2]C)"
'-------------

ApExcel.Cells(51, 2).NumberFormat = "#,##0"
ApExcel.Cells(51, 2).FormulaR1C1 = "=+SUM(R[1]C:R[4]C)"
ApExcel.Cells(51, 3).NumberFormat = "#,##0"
ApExcel.Cells(51, 3).FormulaR1C1 = "=+SUM(R[1]C:R[4]C)"
For I = 4 To lnTotColumnas - 2
    ApExcel.Cells(51, I).NumberFormat = "#,##0"
    ApExcel.Cells(51, I).FormulaR1C1 = "=+SUM(R[1]C:R[4]C)"
Next I
ApExcel.Cells(51, I).NumberFormat = "#,##0"
ApExcel.Cells(51, I).FormulaR1C1 = "=+SUM(R[1]C:R[4]C)"
I = I + 1
ApExcel.Cells(51, I).NumberFormat = "#,##0"
ApExcel.Cells(51, I).FormulaR1C1 = "=+SUM(R[1]C:R[4]C)"
I = I + 1
ApExcel.Cells(51, I).NumberFormat = "0.00%"
ApExcel.Cells(51, I).FormulaR1C1 = "=+SUM(R[1]C:R[4]C)"
'-------------

ApExcel.Cells(57, 2).NumberFormat = "#,##0"
ApExcel.Cells(57, 2).FormulaR1C1 = "=+SUM(R[1]C:R[3]C)"
ApExcel.Cells(57, 3).NumberFormat = "#,##0"
ApExcel.Cells(57, 3).FormulaR1C1 = "=+SUM(R[1]C:R[3]C)"
For I = 4 To lnTotColumnas - 2
    ApExcel.Cells(57, I).NumberFormat = "#,##0"
    ApExcel.Cells(57, I).FormulaR1C1 = "=+SUM(R[1]C:R[3]C)"
Next I
ApExcel.Cells(57, I).NumberFormat = "#,##0"
ApExcel.Cells(57, I).FormulaR1C1 = "=+SUM(R[1]C:R[3]C)"
I = I + 1
ApExcel.Cells(57, I).NumberFormat = "#,##0"
ApExcel.Cells(57, I).FormulaR1C1 = "=+SUM(R[1]C:R[3]C)"
I = I + 1
ApExcel.Cells(57, I).NumberFormat = "0.00%"
ApExcel.Cells(57, I).FormulaR1C1 = "=+SUM(R[1]C:R[3]C)"
'-------------

ApExcel.Cells(62, 2).NumberFormat = "#,##0"
ApExcel.Cells(62, 2).FormulaR1C1 = "=+SUM(R[1]C:R[2]C)"
ApExcel.Cells(62, 3).NumberFormat = "#,##0"
ApExcel.Cells(62, 3).FormulaR1C1 = "=+SUM(R[1]C:R[2]C)"
For I = 4 To lnTotColumnas - 2
    ApExcel.Cells(62, I).NumberFormat = "#,##0"
    ApExcel.Cells(62, I).FormulaR1C1 = "=+SUM(R[1]C:R[2]C)"
Next I
ApExcel.Cells(62, I).NumberFormat = "#,##0"
ApExcel.Cells(62, I).FormulaR1C1 = "=+SUM(R[1]C:R[2]C)"
I = I + 1
ApExcel.Cells(62, I).NumberFormat = "#,##0"
ApExcel.Cells(62, I).FormulaR1C1 = "=+SUM(R[1]C:R[2]C)"
I = I + 1
ApExcel.Cells(62, I).NumberFormat = "0.00%"
ApExcel.Cells(62, I).FormulaR1C1 = "=+SUM(R[1]C:R[2]C)"
'-------------

ApExcel.Cells(66, 2).NumberFormat = "#,##0"
ApExcel.Cells(66, 2).FormulaR1C1 = "=+SUM(R[1]C:R[2]C)"
ApExcel.Cells(66, 3).NumberFormat = "#,##0"
ApExcel.Cells(66, 3).FormulaR1C1 = "=+SUM(R[1]C:R[2]C)"
For I = 4 To lnTotColumnas - 2
    ApExcel.Cells(66, I).NumberFormat = "#,##0"
    ApExcel.Cells(66, I).FormulaR1C1 = "=+SUM(R[1]C:R[2]C)"
Next I
ApExcel.Cells(66, I).NumberFormat = "#,##0"
ApExcel.Cells(66, I).FormulaR1C1 = "=+SUM(R[1]C:R[2]C)"
I = I + 1
ApExcel.Cells(66, I).NumberFormat = "#,##0"
ApExcel.Cells(66, I).FormulaR1C1 = "=+SUM(R[1]C:R[2]C)"
I = I + 1
ApExcel.Cells(66, I).NumberFormat = "0.00%"
ApExcel.Cells(66, I).FormulaR1C1 = "=+SUM(R[1]C:R[2]C)"
'-------------

ApExcel.Cells(70, 2).NumberFormat = "#,##0"
ApExcel.Cells(70, 2).FormulaR1C1 = "=+SUM(R[1]C:R[2]C)"
ApExcel.Cells(70, 3).NumberFormat = "#,##0"
ApExcel.Cells(70, 3).FormulaR1C1 = "=+SUM(R[1]C:R[2]C)"
For I = 4 To lnTotColumnas - 2
    ApExcel.Cells(70, I).NumberFormat = "#,##0"
    ApExcel.Cells(70, I).FormulaR1C1 = "=+SUM(R[1]C:R[2]C)"
Next I
ApExcel.Cells(70, I).NumberFormat = "#,##0"
ApExcel.Cells(70, I).FormulaR1C1 = "=+SUM(R[1]C:R[2]C)"
I = I + 1
ApExcel.Cells(70, I).NumberFormat = "#,##0"
ApExcel.Cells(70, I).FormulaR1C1 = "=+SUM(R[1]C:R[2]C)"
I = I + 1
ApExcel.Cells(70, I).NumberFormat = "0.00%"
ApExcel.Cells(70, I).FormulaR1C1 = "=+SUM(R[1]C:R[2]C)"

'--- FIN TOTALES DETALLES


'--- INI SUB TOTALES
ApExcel.Cells(16, 2).NumberFormat = "#,##0"
ApExcel.Cells(16, 2).FormulaR1C1 = "=+SUM(R[5]C:R[9]C)+R[1]C"
ApExcel.Cells(16, 3).NumberFormat = "#,##0"
ApExcel.Cells(16, 3).FormulaR1C1 = "=+SUM(R[5]C:R[9]C)+R[1]C"
For I = 4 To lnTotColumnas - 2
    ApExcel.Cells(16, I).NumberFormat = "#,##0"
    ApExcel.Cells(16, I).FormulaR1C1 = "=+SUM(R[5]C:R[9]C)+R[1]C"
Next I
ApExcel.Cells(16, I).NumberFormat = "#,##0"
ApExcel.Cells(16, I).FormulaR1C1 = "=+SUM(R[5]C:R[9]C)+R[1]C"
I = I + 1
ApExcel.Cells(16, I).NumberFormat = "#,##0"
ApExcel.Cells(16, I).FormulaR1C1 = "=+SUM(R[5]C:R[9]C)+R[1]C"
I = I + 1
ApExcel.Cells(16, I).NumberFormat = "0.00%"
ApExcel.Cells(16, I).FormulaR1C1 = "=+SUM(R[5]C:R[9]C)+R[1]C"

'-------
ApExcel.Cells(27, 2).NumberFormat = "#,##0"
ApExcel.Cells(27, 2).FormulaR1C1 = "=+SUM(R[1]C:R[4]C)+R[6]C"
ApExcel.Cells(27, 3).NumberFormat = "#,##0"
ApExcel.Cells(27, 3).FormulaR1C1 = "=+SUM(R[1]C:R[4]C)+R[6]C"
For I = 4 To lnTotColumnas - 2
    ApExcel.Cells(27, I).NumberFormat = "#,##0"
    ApExcel.Cells(27, I).FormulaR1C1 = "=+SUM(R[1]C:R[4]C)+R[6]C"
Next I
ApExcel.Cells(27, I).NumberFormat = "#,##0"
ApExcel.Cells(27, I).FormulaR1C1 = "=+SUM(R[1]C:R[4]C)+R[6]C"
I = I + 1
ApExcel.Cells(27, I).NumberFormat = "#,##0"
ApExcel.Cells(27, I).FormulaR1C1 = "=+SUM(R[1]C:R[4]C)+R[6]C"
I = I + 1
ApExcel.Cells(27, I).NumberFormat = "0.00%"
ApExcel.Cells(27, I).FormulaR1C1 = "=+SUM(R[1]C:R[4]C)+R[6]C"

'-------
ApExcel.Cells(42, 2).NumberFormat = "#,##0"
ApExcel.Cells(42, 2).FormulaR1C1 = "=+R[1]C+R[5]C+R[9]C+R[15]C+R[20]C+R[24]C+R[28]C+SUM(R[32]C:R[46]C)"
ApExcel.Cells(42, 3).NumberFormat = "#,##0"
ApExcel.Cells(42, 3).FormulaR1C1 = "=+R[1]C+R[5]C+R[9]C+R[15]C+R[20]C+R[24]C+R[28]C+SUM(R[32]C:R[46]C)"
For I = 4 To lnTotColumnas - 2
    ApExcel.Cells(42, I).NumberFormat = "#,##0"
    ApExcel.Cells(42, I).FormulaR1C1 = "=+R[1]C+R[5]C+R[9]C+R[15]C+R[20]C+R[24]C+R[28]C+SUM(R[32]C:R[46]C)"
Next I
ApExcel.Cells(42, I).NumberFormat = "#,##0"
ApExcel.Cells(42, I).FormulaR1C1 = "=+R[1]C+R[5]C+R[9]C+R[15]C+R[20]C+R[24]C+R[28]C+SUM(R[32]C:R[46]C)"
I = I + 1
ApExcel.Cells(42, I).NumberFormat = "#,##0"
ApExcel.Cells(42, I).FormulaR1C1 = "=+R[1]C+R[5]C+R[9]C+R[15]C+R[20]C+R[24]C+R[28]C+SUM(R[32]C:R[46]C)"
I = I + 1
ApExcel.Cells(42, I).NumberFormat = "0.00%"
ApExcel.Cells(42, I).FormulaR1C1 = "=+R[1]C+R[5]C+R[9]C+R[15]C+R[20]C+R[24]C+R[28]C+SUM(R[32]C:R[46]C)"

'-------
ApExcel.Cells(90, 2).NumberFormat = "#,##0"
ApExcel.Cells(90, 2).FormulaR1C1 = "=+SUM(R[1]C:R[4]C)"
ApExcel.Cells(90, 3).NumberFormat = "#,##0"
ApExcel.Cells(90, 3).FormulaR1C1 = "=+SUM(R[1]C:R[4]C)"
For I = 4 To lnTotColumnas - 2
    ApExcel.Cells(90, I).NumberFormat = "#,##0"
    ApExcel.Cells(90, I).FormulaR1C1 = "=+SUM(R[1]C:R[4]C)"
Next I
ApExcel.Cells(90, I).NumberFormat = "#,##0"
ApExcel.Cells(90, I).FormulaR1C1 = "=+SUM(R[1]C:R[4]C)"
I = I + 1
ApExcel.Cells(90, I).NumberFormat = "#,##0"
ApExcel.Cells(90, I).FormulaR1C1 = "=+SUM(R[1]C:R[4]C)"
I = I + 1
ApExcel.Cells(90, I).NumberFormat = "0.00%"
ApExcel.Cells(90, I).FormulaR1C1 = "=+SUM(R[1]C:R[4]C)"

'-------
ApExcel.Cells(96, 2).NumberFormat = "#,##0"
ApExcel.Cells(96, 2).FormulaR1C1 = "=+SUM(R[1]C:R[1]C)"
ApExcel.Cells(96, 3).NumberFormat = "#,##0"
ApExcel.Cells(96, 3).FormulaR1C1 = "=+SUM(R[1]C:R[1]C)"
For I = 4 To lnTotColumnas - 2
    ApExcel.Cells(96, I).NumberFormat = "#,##0"
    ApExcel.Cells(96, I).FormulaR1C1 = "=+SUM(R[1]C:R[1]C)"
Next I
ApExcel.Cells(96, I).NumberFormat = "#,##0"
ApExcel.Cells(96, I).FormulaR1C1 = "=+SUM(R[1]C:R[1]C)"
I = I + 1
ApExcel.Cells(96, I).NumberFormat = "#,##0"
ApExcel.Cells(96, I).FormulaR1C1 = "=+SUM(R[1]C:R[1]C)"
I = I + 1
ApExcel.Cells(96, I).NumberFormat = "0.00%"
ApExcel.Cells(96, I).FormulaR1C1 = "=+SUM(R[1]C:R[1]C)"

'-------
ApExcel.Cells(99, 2).NumberFormat = "#,##0"
ApExcel.Cells(99, 2).FormulaR1C1 = "=+SUM(R[2]C:R[4]C)"
ApExcel.Cells(99, 3).NumberFormat = "#,##0"
ApExcel.Cells(99, 3).FormulaR1C1 = "=+SUM(R[2]C:R[4]C)"
For I = 4 To lnTotColumnas - 2
    ApExcel.Cells(99, I).NumberFormat = "#,##0"
    ApExcel.Cells(99, I).FormulaR1C1 = "=+SUM(R[2]C:R[4]C)"
Next I
ApExcel.Cells(99, I).NumberFormat = "#,##0"
ApExcel.Cells(99, I).FormulaR1C1 = "=+SUM(R[2]C:R[4]C)"
I = I + 1
ApExcel.Cells(99, I).NumberFormat = "#,##0"
ApExcel.Cells(99, I).FormulaR1C1 = "=+SUM(R[2]C:R[4]C)"
I = I + 1
ApExcel.Cells(99, I).NumberFormat = "0.00%"
ApExcel.Cells(99, I).FormulaR1C1 = "=+SUM(R[2]C:R[4]C)"

'--- FIN SUB TOTALES

'--- INI PRIMER TOTALES
ApExcel.Cells(14, 2).NumberFormat = "#,##0"
ApExcel.Cells(14, 2).FormulaR1C1 = "=+R[2]C+R[13]C"
ApExcel.Cells(14, 3).NumberFormat = "#,##0"
ApExcel.Cells(14, 3).FormulaR1C1 = "=+R[2]C+R[13]C"
For I = 4 To lnTotColumnas - 2
    ApExcel.Cells(14, I).NumberFormat = "#,##0"
    ApExcel.Cells(14, I).FormulaR1C1 = "=+R[2]C+R[13]C"
Next I
ApExcel.Cells(14, I).NumberFormat = "#,##0"
ApExcel.Cells(14, I).FormulaR1C1 = "=+R[2]C+R[13]C"
I = I + 1
ApExcel.Cells(14, I).NumberFormat = "#,##0"
ApExcel.Cells(14, I).FormulaR1C1 = "=+R[2]C+R[13]C"
I = I + 1
ApExcel.Cells(14, I).NumberFormat = "0.00%"
ApExcel.Cells(14, I).FormulaR1C1 = "=+R[2]C+R[13]C"
'--- FIN PRIMER TOTALES

''--- INI TOTALES
ApExcel.Cells(12, 2).NumberFormat = "#,##0"
ApExcel.Cells(12, 2).FormulaR1C1 = "=+R[2]C+R[28]C+R[87]C"
ApExcel.Cells(12, 3).NumberFormat = "#,##0"
ApExcel.Cells(12, 3).FormulaR1C1 = "=+R[2]C+R[28]C+R[87]C"
For I = 4 To lnTotColumnas - 2
    ApExcel.Cells(12, I).NumberFormat = "#,##0"
    ApExcel.Cells(12, I).FormulaR1C1 = "=+R[2]C+R[28]C+R[87]C"
Next I
ApExcel.Cells(12, I).NumberFormat = "#,##0"
ApExcel.Cells(12, I).FormulaR1C1 = "=+R[2]C+R[28]C+R[87]C"
I = I + 1
ApExcel.Cells(12, I).NumberFormat = "#,##0"
ApExcel.Cells(12, I).FormulaR1C1 = "=+R[2]C+R[28]C+R[87]C"
I = I + 1
ApExcel.Cells(12, I).NumberFormat = "0.00%"
ApExcel.Cells(12, I).FormulaR1C1 = "=+R[2]C"
'----------------

ApExcel.Cells(40, 2).NumberFormat = "#,##0"
ApExcel.Cells(40, 2).FormulaR1C1 = "=+R[2]C+R[50]C+R[56]C"
ApExcel.Cells(40, 3).NumberFormat = "#,##0"
ApExcel.Cells(40, 3).FormulaR1C1 = "=+R[2]C+R[50]C+R[56]C"
For I = 4 To lnTotColumnas - 2
    ApExcel.Cells(40, I).NumberFormat = "#,##0"
    ApExcel.Cells(40, I).FormulaR1C1 = "=+R[2]C+R[50]C+R[56]C"
Next I
ApExcel.Cells(40, I).NumberFormat = "#,##0"
ApExcel.Cells(40, I).FormulaR1C1 = "=+R[2]C+R[50]C+R[56]C"
I = I + 1
ApExcel.Cells(40, I).NumberFormat = "#,##0"
ApExcel.Cells(40, I).FormulaR1C1 = "=+R[2]C+R[50]C+R[56]C"
I = I + 1
ApExcel.Cells(40, I).NumberFormat = "0.00%"
ApExcel.Cells(40, I).FormulaR1C1 = "=+R[2]C+R[50]C+R[56]C"

''--- FIN TOTALES

ApExcel.Cells(107, 1) = "Fuente: Balance de Comprobación Consolidado"

ApExcel.Cells.Select
ApExcel.Cells.EntireColumn.AutoFit

ApExcel.Cells.Select
ApExcel.Cells.Font.Size = 8

ApExcel.Cells.Range("A1").Select
    
End Sub

'*** PEAC 20110415
Public Sub CreaBalanceComparativoGastoAdmOpeExcel(ByVal psAnio As String, ByVal pnTotColumnas As Integer, rs As ADODB.Recordset)

Dim I As Integer
Dim lcMeses As String
Dim j As Integer
Dim lnFila As Integer
Dim lcOrden As String
Dim lnTotFila As Double
Dim lnMontoIni As Double
Dim lnTot45 As Double

'Agrega un nuevo Libro
ApExcel.Workbooks.Add
   
'Poner Titulos

ApExcel.Cells(1, 1) = "CAJA MAYNAS S.A."
ApExcel.Cells(2, 2) = " BALANCE COMPRATIVO DEL GASTO DE PERSONAL Y OPERATIVO"

ApExcel.Cells(5, 1) = "CUENTA"
ApExcel.Cells(5, 2) = "DESCRIPCION"
ApExcel.Cells(4, 3) = "TOTAL"
ApExcel.Cells(5, 3) = "ACUMULADO"
ApExcel.Cells(6, 3) = "'" + CStr((CInt(psAnio) - 1))

For I = 1 To pnTotColumnas - 5

    lcMeses = Choose(I, "ENERO", "FEBRERO", "MARZO", "ABRIL", _
                            "MAYO", "JUNIO", "JULIO", "AGOSTO", _
                            "SETIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE")
    j = I + 3
    ApExcel.Cells(5, j) = lcMeses
    ApExcel.Cells(6, j) = "'" + psAnio

Next I
I = I + 3
ApExcel.Cells(4, I) = "TOTAL"
ApExcel.Cells(5, I) = "ACUMULADO"
ApExcel.Cells(6, I) = "'" + psAnio
I = I + 1
ApExcel.Cells(4, I) = "ANALISIS"
ApExcel.Cells(5, I) = "VERTICAL"
ApExcel.Cells(6, I) = "EN %"

ApExcel.Range(ApExcel.Cells(4, 1), ApExcel.Cells(6, I)).Cells.Interior.Color = RGB(223, 223, 223)

'----------- FIN CABECERA

lnFila = 7

Dim lnArrayVar(1 To 12) As Double

Do While Not rs.EOF
   
    If Len(rs!cCtaContCod) = 4 And lnFila <> 7 Then
        lnFila = lnFila + 1
    End If
   
    ApExcel.Cells(lnFila, 1).NumberFormat = "@"
    ApExcel.Cells(lnFila, 1) = rs!cCtaContCod
    
    ApExcel.Cells(lnFila, 2) = "'" + rs!cCtaContDesc
    
    ApExcel.Cells(lnFila, 3).NumberFormat = "#,##0.00"
    ApExcel.Cells(lnFila, 3) = rs!A00
    
    For I = 4 To pnTotColumnas - 2
        ApExcel.Cells(lnFila, I).NumberFormat = "#,##0.00"
        ApExcel.Cells(lnFila, I) = rs.Fields(I)
        If Len(rs!cCtaContCod) = 4 Then
            lnArrayVar(I - 3) = lnArrayVar(I - 3) + rs.Fields(I)
        End If
    Next I
    ApExcel.Cells(lnFila, I).NumberFormat = "#,##0.00"
    ApExcel.Cells(lnFila, I) = rs!A13
    
    I = I + 1
    
    If rs!cCtaContCod = "4501" Then
        lnMontoIni = rs!A13
    End If

    ApExcel.Cells(lnFila, I).NumberFormat = "0.00%"
    ApExcel.Cells(lnFila, I) = IIf(rs!cCtaContCod = "4501", rs!A13 / IIf(rs!A13 = 0, 1, rs!A13), rs!A13 / IIf(lnMontoIni = 0, 1, lnMontoIni))

    If Len(rs!cCtaContCod) = 4 Then
        ApExcel.Range(ApExcel.Cells(lnFila, 1), ApExcel.Cells(lnFila, I)).Cells.Interior.Color = RGB(247, 240, 206)
    ElseIf Len(rs!cCtaContCod) = 6 Then
        ApExcel.Range(ApExcel.Cells(lnFila, 1), ApExcel.Cells(lnFila, I)).Cells.Interior.Color = RGB(213, 240, 228)
    End If
   
    lnFila = lnFila + 1
    rs.MoveNext
Loop

lnFila = lnFila + 1
ApExcel.Cells(lnFila, 1) = "'45"
ApExcel.Cells(lnFila, 2) = "TOTAL GASTOS ADMINISTRATIVOS"
lnTot45 = 0
For I = 4 To pnTotColumnas - 2
    ApExcel.Cells(lnFila, I).NumberFormat = "#,##0.00"
    ApExcel.Cells(lnFila, I) = lnArrayVar(I - 3)
    lnTot45 = lnTot45 + lnArrayVar(I - 3)
Next I

ApExcel.Cells(lnFila, I).NumberFormat = "#,##0.00"
ApExcel.Cells(lnFila, I) = lnTot45
ApExcel.Range(ApExcel.Cells(lnFila, 1), ApExcel.Cells(lnFila, I)).Cells.Interior.Color = RGB(247, 240, 206)

ApExcel.Cells.Select
ApExcel.Cells.EntireColumn.AutoFit

ApExcel.Cells.Select
ApExcel.Cells.Font.Size = 8
ApExcel.Cells.Range("A1").Select

End Sub


