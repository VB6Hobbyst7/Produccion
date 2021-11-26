VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmBalanceHisto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Balance Histórico"
   ClientHeight    =   6660
   ClientLeft      =   615
   ClientTop       =   2010
   ClientWidth     =   11880
   Icon            =   "frmBalanceHisto.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
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
      Left            =   9120
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
      Left            =   10590
      TabIndex        =   5
      Top             =   6210
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
Dim nTipoBala As Integer
Dim nTipCambio As Currency
Dim nmoneda As Integer
Dim lbValidaBalance As Boolean
Dim lsArchivo As String
Dim lbExcel As Boolean
Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet

Public Function ExcelBegin(psArchivo As String, _
        xlAplicacion As Excel.Application, _
        xlLibro As Excel.Workbook, Optional pbBorraExiste As Boolean = True) As Boolean
Dim fs As New Scripting.FileSystemObject
On Error GoTo ErrBegin
Set fs = New Scripting.FileSystemObject
Set xlAplicacion = New Excel.Application
If fs.FileExists(psArchivo) Then
   If pbBorraExiste Then
      fs.DeleteFile psArchivo, True
      Set xlLibro = xlAplicacion.Workbooks.Add
   Else
      Set xlLibro = xlAplicacion.Workbooks.Open(psArchivo)
   End If
Else
   Set xlLibro = xlAplicacion.Workbooks.Add
End If
ExcelBegin = True
Exit Function
ErrBegin:
  MsgBox TextErr(Err.Description), vbInformation, "Aviso"
  ExcelBegin = False
End Function

Public Sub ExcelEnd(psArchivo As String, xlAplicacion As Excel.Application, xlLibro As Excel.Workbook, xlHoja1 As Excel.Worksheet, Optional plSave As Boolean = True)
On Error GoTo ErrEnd
   If plSave Then
        xlHoja1.SaveAs psArchivo
   End If
   xlLibro.Close
   xlAplicacion.Quit
   Set xlAplicacion = Nothing
   Set xlLibro = Nothing
   Set xlHoja1 = Nothing
Exit Sub
ErrEnd:
   MsgBox TextErr(Err.Description), vbInformation, "Aviso"
End Sub

Public Sub ImprimeBalanceComprobacionExcel(pdFechaIni As Date, pdFechaFin As Date, pdFecha As Date, pnTipoBala As Integer, pnMoneda As Integer, pnLinPage As Integer, nTotDebe As Currency, nTotHaber As Currency, Optional psCtaIni As String = "", Optional psCtaFin As String = "", Optional pnDigitos As Integer = 0, Optional pbSoloAnaliticas As Boolean = False, Optional pnCierreAnio As Integer = 0, Optional nTipo As Integer)
Dim lsImpre As String
On Error GoTo GeneraEstadError
    lsArchivo = App.path & "\SPOOLER\" & "Balance_" & pnMoneda & "_" & Mid(Format(pdFechaIni, "yyyymmdd"), 5, 2) & ".XLS"
    lbExcel = ExcelBegin(lsArchivo, xlAplicacion, xlLibro, False)
    If Not lbExcel Then
        Exit Sub
    End If
    ExcelAddHoja "B_" & Format(pdFechaIni, "yyyymmdd") & "_" & Right(Format(pdFechaFin, "yyyymmdd"), 2), xlLibro, xlHoja1, False
    ImprimeBalanceComprobacion pdFechaIni, pdFechaFin, pdFecha, pnTipoBala, pnMoneda, pnLinPage, nTotDebe, nTotHaber, psCtaIni, psCtaFin, pnDigitos, pbSoloAnaliticas, pnCierreAnio, True, xlHoja1, nTipo
    ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1, True
    If lsArchivo <> "" Then
       gFunContab.CargaArchivo lsArchivo, App.path & "\SPOOLER\"
    End If
Exit Sub
GeneraEstadError:
    MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
    If lbExcel = True Then
      ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1, False
    End If

End Sub

Public Sub ExcelAddHoja(psHojName As String, xlLibro As Excel.Workbook, xlHoja1 As Excel.Worksheet, Optional pbActivaHoja As Boolean = True)
Dim lbExisteHoja As Boolean
Dim lbBorrarRangos As Boolean
On Error Resume Next
lbExisteHoja = False
lbBorrarRangos = False
activaHoja:
For Each xlHoja1 In xlLibro.Worksheets
    If UCase(xlHoja1.Name) = UCase(psHojName) Then
        If Not pbActivaHoja Then
            SendKeys "{ENTER}"
            xlHoja1.Delete
        Else
            xlHoja1.Activate
            If lbBorrarRangos Then xlHoja1.Range("A1:BZ1").EntireColumn.Delete
            lbExisteHoja = True
        End If
       Exit For
    End If
Next
If Not lbExisteHoja Then
    Set xlHoja1 = xlLibro.Worksheets.Add
    xlHoja1.Name = psHojName
    If Err Then
        Err.Clear
        pbActivaHoja = True
        lbBorrarRangos = True
        GoTo activaHoja
    End If
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
Dim n As Long
Dim c As Integer
Dim nDebe As Currency, nHaber As Currency

Dim oBalance1 As COMNAuditoria.NBalanceCont
Set oBalance1 = New COMNAuditoria.NBalanceCont

Dim rsCtasME1 As ADODB.Recordset
Set rsCtasME1 = New ADODB.Recordset

If cboMoneda.ListIndex = 2 Then
    Set rsCtasME1 = oBalance1.GetCtasSaldoME(CDate(txtFechaAl.Text))
End If

nDebe = 0: nHaber = 0
For n = 1 To fg.Rows - 1
   If cboMoneda.ListIndex = 2 Then
      rsCtasME1.MoveFirst
      rsCtasME1.Find "cCtaContCod Like '" & fg.TextMatrix(n, 1) & "'"
      If rsCtasME1.EOF And rsCtasME1.EOF Then
         fg.TextMatrix(n, 7) = Format(Round(Val(fg.TextMatrix(n, 6)) / nTipCambio, 2), gsFormatoNumeroView)
      Else
         fg.TextMatrix(n, 7) = Format(rsCtasME1.Fields(1), gsFormatoNumeroView)
      End If
   End If
   For c = 3 To 6
      fg.TextMatrix(n, c) = Format(fg.TextMatrix(n, c), gsFormatoNumeroView)
   Next
Next

Set oBalance1 = Nothing
Set rsCtasME1 = Nothing
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

Private Sub CmdImprimir_Click()
On Error GoTo cmdImprimir
    Dim MM As Integer
                                        MM = 1
    Dim sTexto As String
                                        MM = 2
    lblmsg.Visible = True
                                        MM = 3
    lblmsg.Caption = "Procesando Balance de Comprobación..."
                                        MM = 4
    Me.MousePointer = 11
                                        MM = 5
    fg.MousePointer = 11
                                        MM = 6
    Dim nPosIni As Long, nPosFin As Long
                                        MM = 7
    Dim lsCtaIni As String, lsCtaFin As String
                                        MM = 8
    nPosIni = 1
                                        MM = 9
    nPosFin = fg.Rows - 1
                                        MM = 10
    lsCtaIni = ""
                                        MM = 11
    lsCtaFin = ""
                                        MM = 12
        If chkSele.value = vbChecked Then
                                        MM = 13
           nPosIni = fg.Row
                                        MM = 14
           nPosFin = fg.Row
                                        MM = 15
            If fg.RowSel > fg.Row Then
                                        MM = 16
               nPosFin = fg.RowSel
                                        MM = 17
            Else
                                        MM = 18
               nPosIni = fg.RowSel
                                        MM = 19
           End If
                                        MM = 20
           lsCtaIni = fg.TextMatrix(nPosIni, 1)
                                        MM = 21
           lsCtaFin = fg.TextMatrix(nPosFin, 1)
                                        MM = 22
        End If
                                        MM = 23
        If chkExcel.value = vbChecked Then
                                        MM = 24
           ImprimeBalanceComprobacionExcel CDate(txtFechaDel), CDate(txtFechaAl), IIf(chkFecha.value = 1, txtFechaAl, gdFecSis), nTipoBala, nmoneda, gnLinPage, nVal(txtDebe), nVal(txtHaber), lsCtaIni, lsCtaFin, Val(cboDig), chkSoloAnaliticas.value, Me.chkCierreAnio.value, nmoneda
                                        MM = 25
        Else
                                        MM = 26
           sTexto = gFunGeneralContabilidad.ImprimeBalanceComprobacion(CDate(txtFechaDel), CDate(txtFechaAl), IIf(chkFecha.value = 1, txtFechaAl, gdFecSis), nTipoBala, nmoneda, gnLinPage, nVal(txtDebe), nVal(txtHaber), lsCtaIni, lsCtaFin, Val(cboDig), chkSoloAnaliticas.value, Me.chkCierreAnio.value, , , nmoneda)
                                        MM = 27
           EnviaPrevio sTexto, "BALANCE DE COMPROBACION", gnLinPage, False
                                        MM = 28
        End If
                                        MM = 29
        Me.MousePointer = 0
                                        MM = 30
        fg.MousePointer = 0
                                        MM = 31
        lblmsg.Visible = False
                                        MM = 32
        fg.SetFocus
                                        MM = 33
        Exit Sub
cmdImprimir:
   MsgBox TextErr(Err.Number & " " & Err.Description & " " & Err.Source & " " & MM), vbInformation, "¡BALANCE COMPROBACION!"
End Sub

'Private Function PermiteGenerarBalance() As Boolean
'Dim rs1 As ADODB.Recordset
'Set rs1 = New ADODB.Recordset
'
'Dim oGen As DGeneral
'Set oGen = New DGeneral
'Dim oBlq As COMNAuditoria.DBloqueos
'Set oBlq = New COMNAuditoria.DBloqueos
'
'PermiteGenerarBalance = False
'   Set rs1 = oBlq.CargaBloqueo(CGBloqueos.gBloqueoBalance)
'   If Not rs1.EOF Then
'      If Trim(rs1!cVarValor) = "1" Then
'         Set rs1 = oGen.GetDataUser(Right(rs1!cUltimaActualizacion, 4))
'         If Not rs1.EOF Then
'            If MsgBox(Trim(rs1!cPersNombre) & " esta Generando Balance. ¿Desea de todas maneras generar el Balance?", vbQuestion + vbYesNo + vbDefaultButton2, "Aviso") = vbNo Then
'               RSClose rs1
'               Exit Function
'            End If
'         Else
'            MsgBox "Existe otro usuario Generando Balance. Por favor procesar posteriormente", vbInformation, "Aviso"
'            RSClose rs1
'            Exit Function
'         End If
'      End If
'   End If
'   oBlq.ActualizaBloqueo CGBloqueos.gBloqueoBalance, "1", GeneraMovNroActualiza(Format(GetFechaHoraServer(), gsFormatoFecha), gsCodUser, gsCodCMAC, gsCodAge)
'
'Set rs1 = Nothing
'Set oGen = Nothing
'Set oBlq = Nothing
'PermiteGenerarBalance = True
'End Function

Private Sub cmdProcesar_Click()
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim n As Integer
Dim nPos As Variant
Dim bConHistorico As Boolean
Dim m As Integer

On Error GoTo ErrGeneraBalance

    If cboMoneda.ListIndex = 0 Then gsSimbolo = "": nmoneda = 0
    If cboMoneda.ListIndex = 1 Then gsSimbolo = gcMN: nmoneda = 1
    If cboMoneda.ListIndex = 2 Then gsSimbolo = gcME: nmoneda = 2
    If cboMoneda.ListIndex = 3 Then gsSimbolo = "": nmoneda = 3
    If cboMoneda.ListIndex = 4 Then gsSimbolo = "": nmoneda = 4
    If cboMoneda.ListIndex = 5 Then gsSimbolo = "": nmoneda = 6
    If cboMoneda.ListIndex = 6 Then gsSimbolo = "": nmoneda = 9
    
    If Trim(txtFechaDel) = "/  /" Or Trim(txtFechaAl) = "/  /" Or cboDig.ListIndex = -1 Then
       Exit Sub
    End If
    
    If Day(CDate(txtFechaDel)) <> 1 Then
       MsgBox "Balance debe generarse desde el primer día del Mes", vbInformation, "¡Aviso!"
       txtFechaDel.SetFocus
       Exit Sub
    End If
                                            m = 1
'    If Not PermiteGenerarBalance() Then
'                                            m = 2
'       Exit Sub
'                                            m = 3
'    End If
                                            m = 4
    Dim oBalance As New COMNAuditoria.NBalanceCont
                                            m = 5
    'Set oBalance = New COMNAuditoria.NBalanceCont
                                            m = 6
    Dim oBlq1 As New COMNAuditoria.DBloqueos
                                            m = 7
    'Set oBlq1 = New COMNAuditoria.DBloqueos
                                            m = 8
    lblmsg.Visible = True
                                            m = 9
    lblmsg.Caption = "Procesando... Por favor espere un momento"
                                            m = 10
    DoEvents
                                            m = 11
    MousePointer = 11
                                            m = 12
    Dim oFun As NContFunciones
                                            m = 13
    Set oFun = New NContFunciones
                                            m = 14
    gsMovNro = oFun.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
                                            m = 15
    n = vbYes
                                            m = 16
    If oBalance.BalanceGeneradoHisto(nTipoBala, nmoneda, Month(CDate(txtFechaDel.Text)) + CInt(Me.chkCierreAnio.value), Year(CDate(txtFechaDel.Text))) Then
                                            m = 17
        If oFun.PermiteModificarAsiento(Format(txtFechaDel.Text, gsFormatoMovFecha), False) Then
            MousePointer = 0
                                            m = 18
            n = MsgBox("Balance ya fue generado. ¿Desea volver a Procesar?", vbQuestion + vbYesNoCancel + vbDefaultButton2, "Aviso")
                                            m = 19
            If n = vbCancel Then
                                            m = 20
                oBlq1.ActualizaBloqueo gBloqueoBalance, "0", gsMovNro
                                            m = 21
                lblmsg.Caption = ""
                                            m = 22
                Exit Sub
                                            m = 23
            End If
                                            m = 24
            MousePointer = 11
                                            m = 25
            DoEvents
                                            m = 26
        Else
                                            m = 27
            n = vbNo
                                            m = 28
        End If
                                            m = 29
    End If
                                            m = 30
    Set oFun = Nothing
                                            m = 31
    If n = vbYes Then
                                            m = 32
        Dim dBalance As COMNAuditoria.DbalanceCont
                                            m = 33
        Set dBalance = New COMNAuditoria.DbalanceCont
                                            m = 34
        Dim lsMsgErr As String
                                            m = 35
        lbValidaBalance = oBalance.BalanceGeneradoHisto(nTipoBala, nmoneda, Month(CDate(txtFechaDel) - 1) + CInt(Me.chkCierreAnio), Year(CDate(txtFechaDel) - 1))
                                            m = 36
        If lbValidaBalance Then
                                            m = 37
            lblmsg.Caption = "Validando Saldos Iniciales... Espere un momento"
                                            m = 38
            DoEvents
                                            m = 39
            lsMsgErr = dBalance.ValidaSaldosIniciales(Month(txtFechaDel), Year(CDate(txtFechaDel)))
                                            m = 40
            If lsMsgErr <> "" Then
                                            m = 41
                If MsgBox(TextErr(lsMsgErr), vbQuestion + vbYesNo + vbDefaultButton2, "¿Desea continuar? ") = vbNo Then
                                            m = 42
                    lblmsg.Caption = ""
                                            m = 43
                    Exit Sub
                                            m = 44
                End If
                                            m = 45
            End If
                                            m = 46
        End If
                                            m = 47
        oBlq1.ActualizaBloqueo gBloqueoBalance, "1", gsMovNro
                                            m = 48
        dBalance.EliminaBalance nTipoBala, nmoneda, Month(CDate(txtFechaDel)) + CInt(Me.chkCierreAnio), Year(CDate(txtFechaDel))
                                            m = 49
        dBalance.EliminaBalanceTemp nTipoBala, nmoneda
                                            m = 50
        lblmsg.Caption = "Cálculo de Movimientos..."
                                            m = 51
        DoEvents
                                            m = 52
        If Me.ChkCuentaCero.value = 1 Then bConHistorico = True Else bConHistorico = False
                                            m = 53
            dBalance.InsertaSaldosIniciales nTipoBala, nmoneda, Format(txtFechaDel, gsFormatoFecha), False  ', bConHistorico
                                            m = 54
            dBalance.InsertaMovimientosMes nTipoBala, nmoneda, Format(txtFechaDel, gsFormatoMovFecha), Format(txtFechaAl, gsFormatoMovFecha), , chkCierreAnio.value = vbChecked
                                            m = 55
            lblmsg.Caption = "Mayorizando... Espere un momento"
                                            m = 56
            DoEvents
                                            m = 57
            If cboMoneda.ListIndex = 2 Then
                                            m = 58
                nTipCambio = oBalance.GetTipCambioBalance(Format(txtFechaAl, gsFormatoMovFecha))
                                            m = 59
                If nTipCambio = 0 Then
                                            m = 60
                    MsgBox "No se definio Tipo de Cambio para fecha Final del Balance. Se usara TC.Fijo del día", vbInformation, "¡Aviso!"
                                            m = 61
                    nTipCambio = gnTipCambio
                                            m = 62
                End If
                                            m = 63
            End If
                                            m = 64
        dBalance.MayorizacionBalance nTipoBala, nmoneda, Month(CDate(txtFechaDel)), Year(CDate(txtFechaDel)), chkCierreAnio.value = vbChecked, gsCodCMAC
                                            m = 65
        Set dBalance = Nothing
                                            m = 66
        lblmsg.Caption = "Validando Balance... Por favor espere un momento"
                                            m = 67
        DoEvents
                                            m = 68
        ValidaBalanceEXCEL False, CDate(txtFechaDel), CDate(txtFechaAl), nTipoBala, nmoneda
                                            m = 69
    Else
                                            m = 70
        If cboMoneda.ListIndex = 2 Then
                                            m = 71
            nTipCambio = oBalance.GetTipCambioBalance(Format(txtFechaAl, gsFormatoMovFecha))
                                            m = 72
        End If
                                            m = 73
    End If
                                            m = 74
    Set rs = oBalance.TotalizaBalanceHisto(nTipoBala, nmoneda, Month(CDate(txtFechaDel)) + CInt(Me.chkCierreAnio), Year(CDate(txtFechaDel)))
                                            m = 75
    txtDebe.Text = Format(rs!nDebe, gsFormatoNumeroView)
                                            m = 76
    txtHaber.Text = Format(rs!nHaber, gsFormatoNumeroView)
                                            m = 77
    oBlq1.ActualizaBloqueo CGBloqueos.gBloqueoBalance, "0", gsMovNro
    
    Set oBlq1 = Nothing
                                            m = 78
    lblmsg.Caption = "Mostrando Balance... Espere un momento"
                                            m = 79
    DoEvents
                                            m = 80
    Set rs = oBalance.LeeBalanceHisto(nTipoBala, nmoneda, Month(CDate(txtFechaDel)) + CInt(Me.chkCierreAnio), Year(CDate(txtFechaDel)), , , cboDig, chkSoloAnaliticas)
                                            m = 81
    Set fg.Recordset = rs
                                            m = 82
    If fg.Rows - 1 < rs.RecordCount Then
                                            m = 83
        rs.Move fg.Rows - 1
                                            m = 84
        Do While Not rs.EOF
                                            m = 85
            fg.AddItem ""
                                            m = 86
            nPos = fg.Rows - 1
                                            m = 87
            For n = 0 To rs.Fields.Count - 1
                                            m = 88
                fg.TextMatrix(nPos, n + 1) = rs.Fields(n)
                                            m = 89
            Next
                                            m = 90
            rs.MoveNext
                                            m = 91
        Loop
                                            m = 92
    End If
                                            m = 93
    'RSClose rs
    
    Set rs = Nothing
                                            m = 94
    Set oBalance = Nothing
                                            m = 95
    FormatoFlex
                                            m = 96
    'CambiaFormatoFlex
                                            m = 97
    lblmsg.Visible = False
                                            m = 98
    MousePointer = 0
                                            m = 99
    cmdImprimir.Enabled = True
                                            m = 100
    cmdSituacion.Enabled = True
                                            m = 101
    Exit Sub
                                            m = 102
ErrGeneraBalance:
   MsgBox TextErr(Err.Number & " " & Err.Description & " " & Err.Source & " " & m), vbInformation, "¡Aviso!"
   'rs.Close
   Set oBlq1 = Nothing
   Set oBalance = Nothing
   Set dBalance = Nothing
   MousePointer = 0
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub cmdSituacion_Click()
On Error GoTo cmdSituacion
    Dim MMM As Integer
                                        MMM = 1
    Dim sTexto As String
                                        MMM = 2
    Dim sTexto2 As String
                                        MMM = 3
    MousePointer = 11
                                        MMM = 4
    lblmsg.Visible = True
                                        MMM = 5
    lblmsg.Caption = "Procesando Balance de Situación..."
                                        MMM = 6
    sTexto = gFunGeneralLogistica.ImprimeBalanceSituacion(CDate(txtFechaDel), CDate(txtFechaAl), IIf(chkFecha.value = 1, txtFechaAl, gdFecSis), nTipoBala, nmoneda, nVal(txtDebe), nVal(txtHaber), Me.chkCierreAnio.value)
                                        MMM = 7
    sTexto2 = gFunGeneralTesoreria.AgregaUtilidad(False, CDate(txtFechaDel.Text), CDate(txtFechaAl.Text), nTipoBala, nmoneda)
                                        MMM = 8
    sTexto = sTexto & sTexto2
                                        MMM = 9
    lblmsg.Visible = False
                                        MMM = 10
    MousePointer = 0
                                        MMM = 11
    lblmsg.Caption = ""
                                        MMM = 12
    EnviaPrevio sTexto, "Balance de Situación", gnLinPage, False
                                        MMM = 13
    Unload Me
    Exit Sub
cmdSituacion:
   MsgBox TextErr(Err.Number & " " & Err.Description & " " & Err.Source & " " & MMM), vbInformation, "¡BALANCE SITUACION!"

End Sub

Private Sub fg_KeyUp(KeyCode As Integer, Shift As Integer)
Flex_PresionaKey fg, KeyCode, Shift
End Sub

Private Sub fg_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
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
Dim ObjConeccion As COMConecta.DCOMConecta
Set ObjConeccion = New COMConecta.DCOMConecta
ObjConeccion.CierraConexion
End Sub

Private Sub mnuBuscar_Click()
Dim sTexto As String
Dim n As Long
sTexto = InputBox("Cuenta Contable: ", "Busqueda de Cuenta")
If sTexto <> "" Then
   For n = 1 To fg.Rows - 1
      If fg.TextMatrix(n, 1) = sTexto Then
         fg.Row = n
         fg.TopRow = n
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
