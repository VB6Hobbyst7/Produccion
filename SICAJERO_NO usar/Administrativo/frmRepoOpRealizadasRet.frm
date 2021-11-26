VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmRepoOpRealizadasRet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte Operaciones Log Caja"
   ClientHeight    =   1410
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5010
   Icon            =   "frmRepoOpRealizadasRet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   5010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Detalle"
      Height          =   765
      Left            =   90
      TabIndex        =   4
      Top             =   -30
      Width           =   4860
      Begin MSMask.MaskEdBox txtFechaFinal 
         Height          =   330
         Left            =   3525
         TabIndex        =   1
         Top             =   255
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
         Left            =   1230
         TabIndex        =   0
         Top             =   270
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Final:"
         Height          =   225
         Left            =   2610
         TabIndex        =   7
         Top             =   315
         Width           =   930
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Inicial:"
         Height          =   225
         Left            =   240
         TabIndex        =   6
         Top             =   315
         Width           =   1125
      End
   End
   Begin VB.Frame Frame2 
      Height          =   660
      Left            =   105
      TabIndex        =   5
      Top             =   720
      Width           =   4860
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   315
         Left            =   3135
         TabIndex        =   3
         Top             =   210
         Width           =   1410
      End
      Begin VB.CommandButton cmdGeneraReporte 
         Caption         =   "Generar Reporte"
         Height          =   315
         Left            =   300
         TabIndex        =   2
         Top             =   225
         Width           =   1425
      End
   End
End
Attribute VB_Name = "frmRepoOpRealizadasRet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lnTipo As Integer
Dim lnMoneda As Integer

Public Sub Ini(pnTipo As Integer)
    lnTipo = pnTipo
    Me.Show 1
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdGeneraReporte_Click()
    Dim P As Previo.clsPrevio
    Dim sCadRep As String
    
    sCadRep = GeneraReporte(1)
    If sCadRep = "" Then
        Exit Sub
    End If
    sCadRep = sCadRep & Chr(12) & GeneraReporte(2)
    
    
    
    Set P = New Previo.clsPrevio
    Call P.Show(sCadRep, "REPORTE")
    Set P = Nothing
End Sub

Private Sub Form_Load()
    Me.txtFechaFinal.Text = Format(gdFecSis, "DD/MM/YYYY")
    Me.txtFechaInicial.Text = Format(gdFecSis, "DD/MM/YYYY")
'
'    If lnTipo = 1 Then
'        Me.Caption = "Reporte Operaciones Realizadas"
'    ElseIf lnTipo = 2 Then
'        Me.Caption = "Reporte Operaciones Confirmadas"
'    ElseIf lnTipo = 3 Then
'        Me.Caption = "Reporte Operaciones Pendientes"
'    End If
    
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




Private Function GeneraReporte(pnMoneda As Integer) As String
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter
Dim R As ADODB.Recordset
Dim lnContador As Integer
Dim lnCont As Integer

Dim sSQL As String
Dim sCadRep As String

Dim lsItem As String * 4
Dim lsMovNro As String * 8
Dim lsTarjeta As String * 16
Dim lsCuenta As String * 18
Dim lsHora As String * 8
Dim lsDiaMes As String * 5
Dim lsMondaCta As String * 7
Dim lsMonedaDisp As String * 7
Dim lsTC As String * 8
Dim lsMDisp As String * 10
Dim lsMOpe As String * 10
Dim lsMCom As String * 10
Dim lsMITF As String * 6
Dim lsFecha As String * 10
Dim lsFechaLog As String * 10

Dim lsDesc As String
Dim lsEstado As String * 9

Dim lnMDispSoles As Currency
Dim lnMDispDolares As Currency
Dim lnMontoCom As Currency
Dim lnITF As Currency
Dim lnContSoles As Currency
Dim lnContDolares As Currency
Dim lnMOp As Currency

Dim loConec As New DConecta

    Set R = New ADODB.Recordset
    
    If Not IsDate(txtFechaInicial.Text) Then
        MsgBox "Debe incluir una fecha Inicial valida.", vbInformation, "Reporte Operaciones Realizadas"
        txtFechaInicial.SetFocus
        Exit Function
    End If
    
    If Not IsDate(txtFechaFinal.Text) Then
        MsgBox "Debe incluir una fecha final valida.", vbInformation, "Reporte Operaciones Realizadas"
        txtFechaFinal.SetFocus
        Exit Function
    End If
    
    sSQL = "ATM_ReporteOperacionesRealizadas " & Format(lnTipo) & ",'" & Format(txtFechaInicial.Text, "yyyymmdd") & "','" & Format(txtFechaFinal.Text, "yyyymmdd") & "', '" & Trim(Str(pnMoneda)) & "'"
    'MsgBox sSql
    sCadRep = "."

    
    
    lnMDispSoles = 0
    lnMDispDolares = 0
    lnMontoCom = 0
    lnITF = 0
    lnCont = 1
    lnContSoles = 0
    lnContDolares = 0
    lnMOp = 0
    
    lsItem = "Item"
    lsFecha = "  Fecha "
    lsMovNro = "    Id"
    lsTarjeta = "    Tarjeta"
    lsCuenta = "    Cuenta  "
    lsHora = "Hora"
    lsDiaMes = "M/Dia"
    lsMondaCta = "Mon.Cta"
    lsMonedaDisp = "Mon.Dis"
    lsTC = "   TC"
    lsMDisp = " Mont.Disp"
    lsMOpe = "  Mont.Ope"
    lsMCom = "  Mont.Com"
    lsMITF = "  ITF"
    lsDesc = "   Mov.Desc"
    lsFechaLog = "Fecha LOG"
    lsEstado = "   Estado"
    
    'Cabecera
    sCadRep = sCadRep & Space(5) & "CMAC MAYNAS S.A." & Space(50) & "FECHA : " & Format(Now(), "dd/mm/yyyy hh:mm:ss") & Chr(10)
    sCadRep = sCadRep & Space(5) & "SIMACC-Tarjeta de Debito - " & IIf(pnMoneda = 1, " Soles", " Dolares") & Chr(10) & Chr(10)
    
    If lnTipo = 1 Then
        If Me.txtFechaInicial = Me.txtFechaFinal.Text Then
            sCadRep = sCadRep & Space(40) & "Reporte Log Caja - Retencion el " & Me.txtFechaInicial & Chr(10) & Chr(10) & Chr(10)
        Else
            sCadRep = sCadRep & Space(40) & "Reporte Log Caja -  Retencion entre " & Me.txtFechaInicial & " y " & Me.txtFechaFinal.Text & Chr(10) & Chr(10) & Chr(10)
        End If
    ElseIf lnTipo = 2 Then
        If Me.txtFechaInicial = Me.txtFechaFinal.Text Then
            sCadRep = sCadRep & Space(40) & "Reporte Log Caja - Conciliadas el " & Me.txtFechaInicial & Chr(10) & Chr(10) & Chr(10)
        Else
            sCadRep = sCadRep & Space(40) & "Reporte Log Caja - Conciliadas entre " & Me.txtFechaInicial & " y " & Me.txtFechaFinal.Text & Chr(10) & Chr(10) & Chr(10)
        End If
    ElseIf lnTipo = 3 Then
        If Me.txtFechaInicial = Me.txtFechaFinal.Text Then
            sCadRep = sCadRep & Space(40) & "Reporte Log Caja -  Pendientes el " & Me.txtFechaInicial & Chr(10) & Chr(10) & Chr(10)
        Else
            sCadRep = sCadRep & Space(40) & "Reporte Log Caja -  Pendientes entre " & Me.txtFechaInicial & " y " & Me.txtFechaFinal.Text & Chr(10) & Chr(10) & Chr(10)
        End If
    End If
    
    If lnTipo = 2 Then
        sCadRep = sCadRep & Space(5) & String(189, "-") & Chr(10)
        sCadRep = sCadRep & Space(5) & lsItem & "  " & lsFecha & "  " & lsMovNro & "  " & lsTarjeta & "  " & lsCuenta & "  " & lsHora & "  " & lsDiaMes & "  " & lsMondaCta & "  " & lsMonedaDisp & "  " & lsTC & "  " & lsMDisp & "  " & lsMOpe & "  " & lsMCom & "  " & lsMITF & "  " & lsDesc & "      " & lsFechaLog & "     " & lsEstado & Chr(10)
        sCadRep = sCadRep & Space(5) & String(189, "-") & Chr(10)
    Else
        sCadRep = sCadRep & Space(5) & String(183, "-") & Chr(10)
        sCadRep = sCadRep & Space(5) & lsItem & "  " & lsFecha & "  " & lsMovNro & "  " & lsTarjeta & "  " & lsCuenta & "  " & lsHora & "  " & lsDiaMes & "  " & lsMondaCta & "  " & lsMonedaDisp & "  " & lsTC & "  " & lsMDisp & "  " & lsMOpe & "  " & lsMCom & "  " & lsMITF & "  " & lsDesc & "     " & lsEstado & Chr(10)
        sCadRep = sCadRep & Space(5) & String(183, "-") & Chr(10)
    End If
    
    'AbrirConexion
    loConec.AbreConexion
    R.Open sSQL, loConec.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
     
    Do While Not R.EOF
        lsItem = Format(lnCont, "#0000")
        lsMovNro = Format(R!nMOvNro, "00000000")
        lsTarjeta = R!cPAN
        lsCuenta = R!cCtaCod
        lsHora = Left(R!cHora, 2) & ":" & Mid(R!cHora, 3, 2) & ":" & Right(R!cHora, 2)
        lsDiaMes = Left(R!cMesDia, 2) & "/" & Right(R!cMesDia, 2)
        lsMondaCta = R!MonedaCta
        lsMonedaDisp = R!MonedaDisp
        RSet lsTC = Format(R!TipoCambio, "#0.0000")
        RSet lsMDisp = Format(R!MDispuesto, "#0.00")
        RSet lsMOpe = Format(R!MontoOpe, "#0.00")
        RSet lsMCom = Format(R!Comision, "#0.00")
        RSet lsMITF = Format(R!ITF, "#0.00")
        lsFecha = Mid(R!cMovNro, 7, 2) & "/" & Mid(R!cMovNro, 5, 2) & "/" & Left(R!cMovNro, 4)
        lsDesc = R!cOpeDesc
        lsEstado = R!Estado
    
        
        If lnTipo = 2 Then
            lsFechaLog = Format(R!dfechaArchivo, "DD/MM/YYYY")
            sCadRep = sCadRep & Space(5) & lsItem & "  " & lsFecha & "  " & lsMovNro & "  " & lsTarjeta & "  " & lsCuenta & "  " & lsHora & "  " & lsDiaMes & "  " & lsMondaCta & "  " & lsMonedaDisp & "  " & lsTC & "  " & lsMDisp & "  " & lsMOpe & "  " & lsMCom & "  " & lsMITF & "  " & lsDesc & "  " & lsFechaLog & "  " & lsEstado & Chr(10)
        Else
            sCadRep = sCadRep & Space(5) & lsItem & "  " & lsFecha & "  " & lsMovNro & "  " & lsTarjeta & "  " & lsCuenta & "  " & lsHora & "  " & lsDiaMes & "  " & lsMondaCta & "  " & lsMonedaDisp & "  " & lsTC & "  " & lsMDisp & "  " & lsMOpe & "  " & lsMCom & "  " & lsMITF & "  " & lsDesc & "  " & lsEstado & Chr(10)
        End If
    
        lnContador = lnContador + 1
        lnCont = lnCont + 1
        
        If R!Estado = "Realizada" Then
            lnMOp = lnMOp + R!MontoOpe
            If UCase(R!MonedaDisp) = "DOLARES" Then
                lnContDolares = lnContDolares + R!MDispuesto
            Else
                lnContSoles = lnContSoles + R!MDispuesto
            End If
        End If
        
            lnMontoCom = lnMontoCom + R!Comision
            
            lnITF = lnITF + R!ITF
        
        R.MoveNext
    Loop
    
    sCadRep = sCadRep & Space(5) & String(183, "-") & Chr(10)
    
    lsItem = ""
    RSet lsMovNro = Format(lnContador, "#0")
    lsTarjeta = "Disp.  Cap / ITF"
    lsCuenta = ""
    lsHora = ""
    lsDiaMes = ""
    lsMondaCta = ""
    lsMonedaDisp = ""
    lsTC = ""
    lsMDisp = "" 'Format(lnMDispSoles, "#,##0.00") '"" '"Monto Soles: " & Format(lnContSoles, "#,##0.00") & "/n"
    lsMOpe = Format(lnMOp, "#,##0.00")
    lsMCom = Format(lnMontoCom, "#,##0.00")
    lsMITF = Format(lnITF, "#,##0.00")
    lsFecha = ""
    lsDesc = ""
    
    'sCadRep = sCadRep & Space(5) & lsItem & "  " & lsFecha & "  " & lsMovNro & "  " & lsTarjeta & "  " & lsCuenta & "  " & lsHora & "  " & lsDiaMes & "  " & lsMondaCta & "  " & lsMonedaDisp & "  " & lsTC & Space(2) & lsMDisp & "  " & lsMOpe & "  " & lsMCom & "  " & lsMITF & "  " & lsDesc & Chr(10)
    
    sCadRep = sCadRep & Space(5) & lsItem & "  " & lsFecha & "  " & lsMovNro & "  " & lsTarjeta & "  " & lsCuenta & "  " & lsHora & "  " & lsDiaMes & "  " & lsMondaCta & "  " & lsMonedaDisp & "  " & lsTC & "    " & lsMDisp & "  " & lsMOpe & "    " & lsMCom & lsMITF & Chr(10)
    
    lsItem = ""
    RSet lsMovNro = ""
    lsTarjeta = ""
    lsCuenta = ""
    lsHora = ""
    lsDiaMes = ""
    lsMondaCta = ""
    lsMonedaDisp = ""
    lsTC = "Soles: "
    RSet lsMDisp = Format(lnContSoles, "#,##0.00")
    lsMOpe = ""
    lsMCom = ""
    lsMITF = ""
    lsFecha = ""
    lsDesc = ""
    
    'sCadRep = sCadRep & Space(5) & lsItem & "  " & lsFecha & "  " & lsMovNro & "  " & lsTarjeta & "  " & lsCuenta & "  " & lsHora & "  " & lsDiaMes & "  " & lsMondaCta & "  " & lsMonedaDisp & "  " & lsTC & Space(2) & lsMDisp & "  " & lsMOpe & "  " & lsMCom & "  " & lsMITF & "  " & lsDesc & Chr(10)
    
    sCadRep = sCadRep & Space(5) & lsItem & "  " & lsFecha & "  " & lsMovNro & "  " & lsTarjeta & "  " & lsCuenta & "  " & lsHora & "  " & lsDiaMes & "  " & lsMondaCta & "  " & lsMonedaDisp & "  " & lsTC & "    " & lsMDisp & "  " & lsMOpe & "    " & lsMCom & lsMITF & Chr(10)
    
    lsItem = ""
    RSet lsMovNro = ""
    lsTarjeta = ""
    lsCuenta = ""
    lsHora = ""
    lsDiaMes = ""
    lsMondaCta = ""
    lsMonedaDisp = ""
    lsTC = "Dolares: "
    RSet lsMDisp = Format(lnContDolares, "#,##0.00")
    lsMOpe = ""
    lsMCom = ""
    lsMITF = ""
    lsFecha = ""
    lsDesc = ""
    
    'sCadRep = sCadRep & Space(5) & lsItem & "  " & lsFecha & "  " & lsMovNro & "  " & lsTarjeta & "  " & lsCuenta & "  " & lsHora & "  " & lsDiaMes & "  " & lsMondaCta & "  " & lsMonedaDisp & "  " & lsTC & Space(2) & lsMDisp & "  " & lsMOpe & "  " & lsMCom & "  " & lsMITF & "  " & lsDesc & Chr(10)
    
    sCadRep = sCadRep & Space(5) & lsItem & "  " & lsFecha & "  " & lsMovNro & "  " & lsTarjeta & "  " & lsCuenta & "  " & lsHora & "  " & lsDiaMes & "  " & lsMondaCta & "  " & lsMonedaDisp & "  " & lsTC & "    " & lsMDisp & "  " & lsMOpe & "    " & lsMCom & lsMITF & Chr(10)
    
    
    R.Close
    'CerrarConexion
    loConec.CierraConexion
    Set loConec = Nothing
    Set R = Nothing

    'Cuerpo
'    sCadRep = sCadRep & Space(5) & Space(20) & Left("Saldo Anterior: " & Space(23), 16) & Right(lblSaldoAnt.Caption, 6) & Space(16) & Left("Total de Ingresos: " & Space(23), 21) & Right(lblIngresos.Caption, 6) & Chr(10)
'    sCadRep = sCadRep & Space(5) & String(100, " ") & Chr(10)
'    sCadRep = sCadRep & Space(5) & Space(20) & Left("Total de Salidas: " & Space(23), 18) & Right(lblSalidas.Caption, 6) & Space(16) & Left("Total de Remesas Conf.: " & Space(23), 23) & Right(lblRemesas.Caption, 6) & Chr(10)
'    sCadRep = sCadRep & Space(5) & String(100, " ") & Chr(10)
'    sCadRep = sCadRep & Space(5) & Space(20) & Left("Total de Devoluciones: " & Space(23), 23) & Right(lblDevoluciones.Caption, 6) & Space(12) & Left("Total Stock Actual: " & Space(23), 21) & Right(lblStockActual.Caption, 6) & Chr(10)
'    sCadRep = sCadRep & Space(5) & String(100, " ") & Chr(10) & Chr(10)
'    sCadRep = sCadRep & Space(5) & String(100, "-") & Chr(10) & Chr(10)

    GeneraReporte = sCadRep
    
End Function

Private Function GeneraReporteDenegadas()
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter
Dim R As ADODB.Recordset
Dim lnContador As Integer

Dim sSQL As String
Dim sCadRep As String

Dim lsTarjeta As String * 16
Dim lsHora As String * 8
Dim lsDiaMes As String * 8
Dim lsImporte As String * 7
Dim lsTipoMoneda As String * 7
Dim lsItem As String * 4
Dim lsDesc As String
Dim lnCont As Integer
Dim lnContSoles As Integer
Dim lnContDolares As Integer

Dim lnImporte As Currency
Dim loConec As New DConecta


    Set R = New ADODB.Recordset
    
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
    
    sSQL = "ATM_ActualizaTramas '" & Format(txtFechaInicial.Text, "yyyyMMdd") & "','" & Format(txtFechaFinal.Text, "yyyyMMdd") & "' "
    'MsgBox sSql
    sCadRep = "."
    

    
    lnImporte = 0
    lsItem = "Item"
    lsTarjeta = "    Tarjeta"
    lsHora = "  Hora"
    lsDiaMes = "M/Dia"
    lsImporte = "Importe"
    lsTipoMoneda = "Moneda"
    lsDesc = "   Mov.Desc"
    lnCont = 1
    
    'Cabecera
    sCadRep = sCadRep & Space(5) & "CMAC MAYNAS S.A." & Space(50) & "FECHA : " & Format(Now(), "dd/mm/yyyy hh:mm:ss") & Chr(10)
    sCadRep = sCadRep & Space(5) & "SIMACC-Tarjeta de Debito" & Chr(10) & Chr(10)
    
    sCadRep = sCadRep & Space(35) & "Reporte Operaciones Denegadas - Retencion entre " & Me.txtFechaInicial & " y " & Me.txtFechaFinal.Text & Chr(10) & Chr(10) & Chr(10)
    
    sCadRep = sCadRep & Space(5) & String(76, "-") & Chr(10)
    sCadRep = sCadRep & Space(5) & lsItem & "  " & lsTarjeta & "  " & lsHora & "  " & lsDiaMes & "  " & lsImporte & "  " & lsTipoMoneda & "  " & lsDesc & Chr(10)
    sCadRep = sCadRep & Space(5) & String(76, "-") & Chr(10)
    
    'AbrirConexion
    loConec.AbreConexion
    R.Open sSQL, loConec.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
     
    Do While Not R.EOF
        lsItem = Format(lnCont, "#0000")
        lsTarjeta = R!PAN
        lsHora = Left(R!TimeLocal, 2) & ":" & Mid(R!TimeLocal, 3, 2) & ":" & Right(R!TimeLocal, 2)
        lsDiaMes = Mid(R!DateLocal, 3, 2) & "/" & Right(R!DateLocal, 2)
        RSet lsImporte = Format(R!Txn_Amount / 100, "#,##0.00")
        RSet lsImporte = Format(R!Txn_Amount, "#0.00")
        lsTipoMoneda = R!TipoMoneda
        lsDesc = R!TipoMov
        
        'lnImporte = lnImporte + R!IMPORTE1
        sCadRep = sCadRep & Space(5) & lsItem & "  " & lsTarjeta & "  " & lsHora & "  " & lsDiaMes & "  " & lsImporte & "  " & lsTipoMoneda & "  " & lsDesc & Chr(10)
        
        If UCase(R!TipoMoneda) = "DOLARES" Then
            lnContDolares = lnContDolares + R!Txn_Amount
            
        Else
            lnContSoles = lnContSoles + R!Txn_Amount
            
        End If
                
        'lnImporte = lnImporte + R!Txn_Amount '/ 100
        R.MoveNext
        lnContador = lnContador + 1
        lnCont = lnCont + 1
    Loop
    
    sCadRep = sCadRep & Space(5) & String(76, "-") & Chr(10)
    
    lsItem = ""
    lsTarjeta = Format(lnContador, "#0")
    lsHora = ""
    lsDiaMes = "Soles: "
    RSet lsImporte = Format(lnContSoles, "#,##0.00")
    lsTipoMoneda = ""
    lsDesc = ""
    
    
    sCadRep = sCadRep & Space(5) & lsItem & "  " & lsTarjeta & "  " & lsHora & "  " & lsDiaMes & "  " & lsImporte & "  " & lsTipoMoneda & "  " & lsDesc & Chr(10)
    
    lsItem = ""
    lsTarjeta = ""
    lsHora = ""
    lsDiaMes = "Dolares: "
    RSet lsImporte = Format(lnContDolares, "#,##0.00")
    lsTipoMoneda = ""
    lsDesc = ""
    
    
    sCadRep = sCadRep & Space(5) & lsItem & "  " & lsTarjeta & "  " & lsHora & "  " & lsDiaMes & "  " & lsImporte & "  " & lsTipoMoneda & "  " & lsDesc & Chr(10)
    
    
    R.Close
    'CerrarConexion
    loConec.CierraConexion
    Set loConec = Nothing
    Set R = Nothing
    
    'Cuerpo
'    sCadRep = sCadRep & Space(5) & Space(20) & Left("Saldo Anterior: " & Space(23), 16) & Right(lblSaldoAnt.Caption, 6) & Space(16) & Left("Total de Ingresos: " & Space(23), 21) & Right(lblIngresos.Caption, 6) & Chr(10)
'    sCadRep = sCadRep & Space(5) & String(100, " ") & Chr(10)
'    sCadRep = sCadRep & Space(5) & Space(20) & Left("Total de Salidas: " & Space(23), 18) & Right(lblSalidas.Caption, 6) & Space(16) & Left("Total de Remesas Conf.: " & Space(23), 23) & Right(lblRemesas.Caption, 6) & Chr(10)
'    sCadRep = sCadRep & Space(5) & String(100, " ") & Chr(10)
'    sCadRep = sCadRep & Space(5) & Space(20) & Left("Total de Devoluciones: " & Space(23), 23) & Right(lblDevoluciones.Caption, 6) & Space(12) & Left("Total Stock Actual: " & Space(23), 21) & Right(lblStockActual.Caption, 6) & Chr(10)
'    sCadRep = sCadRep & Space(5) & String(100, " ") & Chr(10) & Chr(10)
'    sCadRep = sCadRep & Space(5) & String(100, "-") & Chr(10) & Chr(10)

    GeneraReporteDenegadas = sCadRep
    
End Function
