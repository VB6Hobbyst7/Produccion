VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmRepResumenOp 
   Caption         =   "Resumen Operaciones"
   ClientHeight    =   1545
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5145
   Icon            =   "frmRepResumenOp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   5145
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   810
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   5115
      Begin MSMask.MaskEdBox txtFechaFinal 
         Height          =   330
         Left            =   3645
         TabIndex        =   2
         Top             =   225
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   582
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFechaInicial 
         Height          =   300
         Left            =   1335
         TabIndex        =   1
         Top             =   255
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha de Fin :"
         Height          =   225
         Left            =   2625
         TabIndex        =   7
         Top             =   300
         Width           =   1305
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha de Inicio :"
         Height          =   225
         Left            =   120
         TabIndex        =   6
         Top             =   285
         Width           =   1305
      End
   End
   Begin VB.Frame Frame2 
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   810
      Width           =   5070
      Begin VB.CommandButton cmdGeneraReporte 
         Caption         =   "Generar"
         Height          =   375
         Left            =   105
         TabIndex        =   3
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   3480
         TabIndex        =   4
         Top             =   240
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmRepResumenOp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oConec As DConecta

'Dim lnMoneda As Integer

Public Sub Ini()
    'lnMoneda = pnMoneda
    Me.Show 1
End Sub

Private Sub cmdGeneraReporte_Click()
    Dim P As Previo.clsPrevio
    Dim sCad As String
    Dim sCadRep As String
    sCadRep = GeneraReporteB()
    If sCadRep = "" Then
        Exit Sub
    End If
    sCadRep = sCadRep & GeneraReporteA()
    
    Set P = New Previo.clsPrevio
    Call P.Show(sCadRep, "REPORTE")
    Set P = Nothing
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Set oConec = New DConecta
    Me.txtFechaFinal.Text = Format(gdFecSis, "DD/MM/YYYY")
    Me.txtFechaInicial.Text = Format(gdFecSis, "DD/MM/YYYY")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set oConec = Nothing
End Sub

Private Sub txtFechaFinal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdGeneraReporte.SetFocus
    End If
End Sub

Private Sub txtFechaInicial_GotFocus()
    txtFechaInicial.SelStart = 0
    txtFechaInicial.SelLength = 50
End Sub

Private Sub txtFechaFinal_GotFocus()
    txtFechaFinal.SelStart = 0
    txtFechaFinal.SelLength = 50
End Sub

Private Sub txtFechaInicial_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtFechaFinal.SetFocus
    End If
End Sub

Private Function GeneraReporteA() As String
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter
Dim R As ADODB.Recordset
Dim lnContador As Integer

Dim sSQL As String
Dim sCadRep As String
    
    If Not IsDate(txtFechaInicial.Text) Then
        MsgBox "Debe incluir una fecha Inicial valida.", vbInformation, "Reporte Operaciones Confirmadas"
        txtFechaInicial.SetFocus
        Exit Function
    End If
    
    If Not IsDate(txtFechaFinal.Text) Then
        MsgBox "Debe incluir una fecha final valida.", vbInformation, "Reporte Operaciones Confirmadas"
        txtFechaFinal.SetFocus
        Exit Function
    End If
    
    'sSql = "ATM_ReporteLOGResumen '" & Format(txtFechaInicial.Text, "yyyyMMdd") & "','" & Format(txtFechaFinal.Text, "yyyyMMdd") & "', '" & Trim(Str(lnMoneda)) & "',0,0,0,0,0,0,0,0,0"
    'MsgBox sSql
    
    Set R = New ADODB.Recordset
    
    oConec.AbreConexion
    Cmd.ActiveConnection = oConec.ConexionActiva 'AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@cFechaInicial", adVarChar, adParamInput, 8, Format(txtFechaInicial.Text, "yyyyMMdd"))
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@cFechaFinal", adVarChar, adParamInput, 8, Format(txtFechaFinal.Text, "yyyyMMdd"))
    Cmd.Parameters.Append Prm
    
                    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nConfLOGME", adInteger, adParamOutput, 8, 0)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nConfLOGMN", adInteger, adParamOutput, 8, 0)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nConfLOGTotal", adInteger, adParamOutput, 8, 0)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nConfxRegMN", adInteger, adParamOutput, 8, 0)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nConfxRegME", adInteger, adParamOutput, 8, 0)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nConfxRegTotal", adInteger, adParamOutput, 8, 0)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nConfConcilMN", adInteger, adParamOutput, , 0)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nConfConcilME", adInteger, adParamOutput, , 0)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nConfConcilTotal", adInteger, adParamOutput, , 0)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nConfLOGCC", adInteger, adParamOutput, 8, 0)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nConfxRegCC", adInteger, adParamOutput, 8, 0)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nConfConcilCC", adInteger, adParamOutput, , 0)
    Cmd.Parameters.Append Prm
    
    Cmd.CommandText = "ATM_ReporteLOGResumen"
    
    R.CursorType = adOpenStatic
    R.LockType = adLockReadOnly
    Set R = Cmd.Execute
        
    
    'MsgBox (Cmd.Parameters(14).Value)
    'Exit Sub
    Dim lnConfLOGMN As String * 5
    Dim lnConfLOGME As String * 5
    Dim lnConfLOGTotal As String * 5
    
    Dim lnConfxRegMN As String * 5
    Dim lnConfxRegME As String * 5
    Dim lnConfxRegTotal As String * 5
    
    Dim lnConfConcilMN As String * 5
    Dim lnConfConcilME As String * 5
    Dim lnConfConcilTotal As String * 5
    
    Dim lnConfLOGCC As String * 5
    Dim lnConfxRegCC As String * 5
    Dim lnConffConcilCC As String * 5
    sCadRep = "."
    
    'Cabecera
'''    sCadRep = sCadRep & Space(5) & "CMAC MAYNAS S.A." & Space(50) & "FECHA : " & Format(Now(), "dd/mm/yyyy hh:mm:ss") & Chr(10)
'''    sCadRep = sCadRep & Space(5) & "SIMACC-Tarjeta de Debito" & Chr(10) & Chr(10)
'''
'''    sCadRep = sCadRep & Space(35) & "Reporte Resumen de Operaciones - Entre " & Me.txtFechaInicial & " y " & Me.txtFechaFinal.Text & Chr(10) & Chr(10) & Chr(10)
'''
    sCadRep = sCadRep & Chr(10)
    sCadRep = sCadRep & Chr(10)
    sCadRep = sCadRep & Chr(10)

    sCadRep = sCadRep & Space(5) & String(75, "-") & Chr(10)
    sCadRep = sCadRep & Space(20) & "  " & "Confirmadas" & "  " & "Confirmadas LOG" & "  " & "Confirmadas por" & Chr(10)
    sCadRep = sCadRep & Space(20) & "  " & "    LOG    " & "  " & "  Conciliadas  " & "  " & "Regularizar LOG" & Chr(10)
    sCadRep = sCadRep & Space(5) & String(75, "-") & Chr(10)

    RSet lnConfLOGMN = Format(Cmd.Parameters(2).Value, "#,##0")
    RSet lnConfxRegMN = Format(Cmd.Parameters(5).Value, "#,##0")
    RSet lnConfConcilMN = Format(Cmd.Parameters(8).Value, "#,##0")
    sCadRep = sCadRep & Space(5) & "Moneda Nacional   " & "   " & Right(lnConfLOGMN, 5) & Space(10) & Right(lnConfConcilMN, 5) & Space(10) & Right(lnConfxRegMN, 5) & Chr(10)
    
    RSet lnConfLOGME = Format(Cmd.Parameters(3).Value, "#,##0")
    RSet lnConfxRegME = Format(Cmd.Parameters(6).Value, "#,##0")
    RSet lnConfConcilME = Format(Cmd.Parameters(9).Value, "#,##0")
    sCadRep = sCadRep & Space(5) & "Moneda Extranjera " & "   " & Right(lnConfLOGME, 5) & Space(10) & Right(lnConfConcilME, 5) & Space(10) & Right(lnConfxRegME, 5) & Chr(10)
    
    RSet lnConfLOGCC = Format(Cmd.Parameters(11).Value, "#,##0")
    RSet lnConfxRegCC = Format(Cmd.Parameters(12).Value, "#,##0")
    RSet lnConffConcilCC = Format(Cmd.Parameters(13).Value, "#,##0")
    sCadRep = sCadRep & Space(5) & "Cambio de Clave   " & "   " & Right(lnConfLOGCC, 5) & Space(10) & Right(lnConffConcilCC, 5) & Space(10) & Right(lnConfxRegCC, 5) & Chr(10)
    
    RSet lnConfLOGTotal = Format(Cmd.Parameters(4).Value, "#,##0")
    RSet lnConfxRegTotal = Format(Cmd.Parameters(7).Value, "#,##0")
    RSet lnConfConcilTotal = Format(Cmd.Parameters(10).Value, "#,##0")
    sCadRep = sCadRep & Space(5) & "TOTAL             " & "   " & Left(lnConfLOGTotal, 5) & Space(10) & Right(lnConfConcilTotal, 5) & Space(10) & Right(lnConfxRegTotal, 5) & Chr(10)
    
    
    sCadRep = sCadRep & Space(5) & String(75, "-") & Chr(10)
'
'    lsItem = ""
'    lsTarjeta = Format(lnContador, "#0")
'    lsHora = ""
'    lsDiaMes = ""
'    RSet lsImporte = Format(lnImporte, "#,##0.00")
'    lsTipoMoneda = ""
'    lsDesc = ""
    
    
'    sCadRep = sCadRep & Space(5) & lsItem & "  " & lsTarjeta & "  " & lsHora & "  " & lsDiaMes & "  " & lsImporte & "  " & lsTipoMoneda & "  " & lsDesc & Chr(10)
    
    'R.Close
    'CerrarConexion
    oConec.CierraConexion
    
    Set Cmd = Nothing
    Set Prm = Nothing
    Set R = Nothing
    
    'Cuerpo
'    sCadRep = sCadRep & Space(5) & Space(20) & Left("Saldo Anterior: " & Space(23), 16) & Right(lblSaldoAnt.Caption, 6) & Space(16) & Left("Total de Ingresos: " & Space(23), 21) & Right(lblIngresos.Caption, 6) & Chr(10)
'    sCadRep = sCadRep & Space(5) & String(100, " ") & Chr(10)
'    sCadRep = sCadRep & Space(5) & Space(20) & Left("Total de Salidas: " & Space(23), 18) & Right(lblSalidas.Caption, 6) & Space(16) & Left("Total de Remesas Conf.: " & Space(23), 23) & Right(lblRemesas.Caption, 6) & Chr(10)
'    sCadRep = sCadRep & Space(5) & String(100, " ") & Chr(10)
'    sCadRep = sCadRep & Space(5) & Space(20) & Left("Total de Devoluciones: " & Space(23), 23) & Right(lblDevoluciones.Caption, 6) & Space(12) & Left("Total Stock Actual: " & Space(23), 21) & Right(lblStockActual.Caption, 6) & Chr(10)
'    sCadRep = sCadRep & Space(5) & String(100, " ") & Chr(10) & Chr(10)
'    sCadRep = sCadRep & Space(5) & String(100, "-") & Chr(10) & Chr(10)
    
    GeneraReporteA = sCadRep
    
End Function

Private Function GeneraReporteB() As String
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter
Dim R As ADODB.Recordset
Dim lnContador As Integer

Dim sSQL As String
Dim sCadRep As String

    If Not IsDate(txtFechaInicial.Text) Then
        MsgBox "Debe incluir una fecha Inicial valida.", vbInformation, "Reporte Operaciones Confirmadas"
        txtFechaInicial.SetFocus
        Exit Function
    End If
    
    If Not IsDate(txtFechaFinal.Text) Then
        MsgBox "Debe incluir una fecha final valida.", vbInformation, "Reporte Operaciones Confirmadas"
        txtFechaFinal.SetFocus
        Exit Function
    End If
    
    'sSql = "ATM_ReporteLOGResumen '" & Format(txtFechaInicial.Text, "yyyyMMdd") & "','" & Format(txtFechaFinal.Text, "yyyyMMdd") & "', '" & Trim(Str(lnMoneda)) & "',0,0,0,0,0,0,0,0,0"
    'MsgBox sSql
    
    Set R = New ADODB.Recordset
        
    oConec.AbreConexion
    Cmd.ActiveConnection = oConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@cFechaInicial", adVarChar, adParamInput, 8, Format(txtFechaInicial.Text, "yyyyMMdd"))
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@cFechaFinal", adVarChar, adParamInput, 8, Format(txtFechaFinal.Text, "yyyyMMdd"))
    Cmd.Parameters.Append Prm
    
                    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nConfCAJAME", adInteger, adParamOutput, 8, 0)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nConfCAJAMN", adInteger, adParamOutput, 8, 0)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nConfCAJATotal", adInteger, adParamOutput, 8, 0)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nConfxRegCAJAMN", adInteger, adParamOutput, 8, 0)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nConfxRegCAJAME", adInteger, adParamOutput, 8, 0)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nConfxRegCAJATotal", adInteger, adParamOutput, 8, 0)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nConfConcilCAJAMN", adInteger, adParamOutput, , 0)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nConfConcilCAJAME", adInteger, adParamOutput, , 0)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nConfConcilCAJATotal", adInteger, adParamOutput, , 0)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nConfCAJACC", adInteger, adParamOutput, 8, 0)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nConfxRegCAJACC", adInteger, adParamOutput, 8, 0)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nConfConciCAJAlCC", adInteger, adParamOutput, , 0)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nDenegadasSol", adInteger, adParamOutput, , 0)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@nDenegadasDol", adInteger, adParamOutput, , 0)
    Cmd.Parameters.Append Prm
    
    Cmd.CommandText = "ATM_ReporteCAJAResumen"
    
    R.CursorType = adOpenStatic
    R.LockType = adLockReadOnly
    Set R = Cmd.Execute
        
        'MsgBox Cmd.Parameters(11).Value
    
        'Dim i As Integer
        
'        For i = 0 To 14 Step 1
'        MsgBox Cmd.Parameters(i).Value
'        Next
    
    'Exit Sub
    Dim lnConfLOGMN As String * 5
    Dim lnConfLOGME As String * 5
    Dim lnConfLOGTotal As String * 5
    
    Dim lnConfxRegMN As String * 5
    Dim lnConfxRegME As String * 5
    Dim lnConfxRegTotal As String * 5
    
    Dim lnConfConcilMN As String * 5
    Dim lnConfConcilME As String * 5
    Dim lnConfConcilTotal As String * 5
    Dim lnDevolucionesMN As String * 5
    Dim lnDevolucionesME As String * 5
    Dim lnDevoluciones As String * 5
    
    Dim lnConfLOGCC As String * 5
    Dim lnConfxRegCC As String * 5
    Dim lnConffConcilCC As String * 5
    sCadRep = "."
    
    'Cabecera
    sCadRep = sCadRep & Space(5) & "CMAC MAYNAS S.A." & Space(50) & "FECHA : " & Format(Now(), "dd/mm/yyyy hh:mm:ss") & Chr(10)
    sCadRep = sCadRep & Space(5) & "SIMACC-Tarjeta de Debito" & Chr(10) & Chr(10)
    
    sCadRep = sCadRep & Space(35) & "Reporte Resumen de Operaciones - Entre " & Me.txtFechaInicial & " y " & Me.txtFechaFinal.Text & Chr(10) & Chr(10) & Chr(10)
    
    sCadRep = sCadRep & Space(5) & String(75, "-") & Chr(10)
    sCadRep = sCadRep & Space(20) & "  " & " Realizadas" & "  " & "Confirmadas    " & "  " & "Pendientes     " & "  Denegadas " & Chr(10)
    sCadRep = sCadRep & Space(20) & "  " & "           " & "  " & "               " & "  " & "               " & Chr(10)
    sCadRep = sCadRep & Space(5) & String(75, "-") & Chr(10)

    RSet lnConfLOGMN = Format(Cmd.Parameters(2).Value, "#,##0")
    RSet lnConfxRegMN = Format(Cmd.Parameters(5).Value, "#,##0")
    RSet lnConfConcilMN = Format(Cmd.Parameters(8).Value, "#,##0")
    RSet lnDevolucionesMN = Format(Cmd.Parameters(14).Value, "#,##0")
    sCadRep = sCadRep & Space(5) & "Moneda Nacional   " & "   " & lnConfLOGMN & Space(10) & lnConfxRegMN & Space(10) & lnConfConcilMN & Space(10) & lnDevolucionesMN & Chr(10)
    
    RSet lnConfLOGME = Format(Cmd.Parameters(3).Value, "#,##0")
    RSet lnConfxRegME = Format(Cmd.Parameters(6).Value, "#,##0")
    RSet lnConfConcilME = Format(Cmd.Parameters(9).Value, "#,##0")
    RSet lnDevolucionesME = Format(Cmd.Parameters(15).Value, "#,##0")
    sCadRep = sCadRep & Space(5) & "Moneda Extranjera " & "   " & lnConfLOGME & Space(10) & lnConfxRegME & Space(10) & lnConfConcilME & Space(10) & lnDevolucionesME & Chr(10)
    
    lnConfLOGCC = Cmd.Parameters(11).Value
    lnConfxRegCC = Cmd.Parameters(12).Value
    lnConffConcilCC = Cmd.Parameters(13).Value
    'sCadRep = sCadRep & Space(5) & "Cambio de Clave   " & "   " & Right(lnConfLOGCC, 5) & Space(10) & Right(lnConfxRegCC, 5) & Space(10) & Right(lnConffConcilCC, 5) & Chr(10)
    
    'sCadRep = sCadRep & Space(5) & "Denegadas         " & "   " & "     " & Space(10) & "     " & Space(10) & "     " & Space(10) & Cmd.Parameters(14).Value & Chr(10)
    
    RSet lnConfLOGTotal = Format(Cmd.Parameters(4).Value, "#,##0")
    RSet lnConfxRegTotal = Format(Cmd.Parameters(7).Value, "#,##0")
    RSet lnConfConcilTotal = Format(Cmd.Parameters(10).Value, "#,##0")
    RSet lnDevoluciones = Format(Cmd.Parameters(14).Value + Cmd.Parameters(15).Value, "#,##0")
    sCadRep = sCadRep & Space(5) & "TOTAL             " & "   " & lnConfLOGTotal & Space(10) & lnConfxRegTotal & Space(10) & lnConfConcilTotal & Space(10) & lnDevoluciones & Chr(10)
    
    
    sCadRep = sCadRep & Space(5) & String(75, "-") & Chr(10)
'
'    lsItem = ""
'    lsTarjeta = Format(lnContador, "#0")
'    lsHora = ""
'    lsDiaMes = ""
'    RSet lsImporte = Format(lnImporte, "#,##0.00")
'    lsTipoMoneda = ""
'    lsDesc = ""
    
    
'    sCadRep = sCadRep & Space(5) & lsItem & "  " & lsTarjeta & "  " & lsHora & "  " & lsDiaMes & "  " & lsImporte & "  " & lsTipoMoneda & "  " & lsDesc & Chr(10)
    
    'R.Close
    'CerrarConexion
    oConec.CierraConexion
    Set Cmd = Nothing
    Set Prm = Nothing
    Set R = Nothing
    
    'Cuerpo
'    sCadRep = sCadRep & Space(5) & Space(20) & Left("Saldo Anterior: " & Space(23), 16) & Right(lblSaldoAnt.Caption, 6) & Space(16) & Left("Total de Ingresos: " & Space(23), 21) & Right(lblIngresos.Caption, 6) & Chr(10)
'    sCadRep = sCadRep & Space(5) & String(100, " ") & Chr(10)
'    sCadRep = sCadRep & Space(5) & Space(20) & Left("Total de Salidas: " & Space(23), 18) & Right(lblSalidas.Caption, 6) & Space(16) & Left("Total de Remesas Conf.: " & Space(23), 23) & Right(lblRemesas.Caption, 6) & Chr(10)
'    sCadRep = sCadRep & Space(5) & String(100, " ") & Chr(10)
'    sCadRep = sCadRep & Space(5) & Space(20) & Left("Total de Devoluciones: " & Space(23), 23) & Right(lblDevoluciones.Caption, 6) & Space(12) & Left("Total Stock Actual: " & Space(23), 21) & Right(lblStockActual.Caption, 6) & Chr(10)
'    sCadRep = sCadRep & Space(5) & String(100, " ") & Chr(10) & Chr(10)
'    sCadRep = sCadRep & Space(5) & String(100, "-") & Chr(10) & Chr(10)
    
    GeneraReporteB = sCadRep
    
End Function
