VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmRepoOpeNoConcil 
   Caption         =   "Reporte Operaciones Conciliadas"
   ClientHeight    =   1725
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4980
   Icon            =   "frmRepoOpeNoConcil.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   4980
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   660
      Left            =   105
      TabIndex        =   7
      Top             =   1005
      Width           =   4860
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   315
         Left            =   3120
         TabIndex        =   3
         Top             =   195
         Width           =   1410
      End
      Begin VB.CommandButton cmdGeneraReporte 
         Caption         =   "Generar Reporte"
         Height          =   315
         Left            =   285
         TabIndex        =   2
         Top             =   225
         Width           =   1425
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Detalle"
      Height          =   990
      Left            =   90
      TabIndex        =   4
      Top             =   -15
      Width           =   4860
      Begin VB.CheckBox chkFechaMov 
         Caption         =   "Fecha movimientos"
         Height          =   195
         Left            =   420
         TabIndex        =   8
         Top             =   675
         Width           =   3225
      End
      Begin MSMask.MaskEdBox txtFechaFinal 
         Height          =   330
         Left            =   3525
         TabIndex        =   1
         Top             =   255
         Width           =   1185
         _ExtentX        =   2090
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
         Top             =   240
         Width           =   1200
         _ExtentX        =   2117
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
         TabIndex        =   6
         Top             =   270
         Width           =   930
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Inicial:"
         Height          =   225
         Left            =   240
         TabIndex        =   5
         Top             =   285
         Width           =   1125
      End
   End
End
Attribute VB_Name = "frmRepoOpeNoConcil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lnMoneda As Integer

Public Sub Inic()
    'lnMoneda = pnMoneda
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

Private Function GeneraReporte(lnMoneda As Integer) As String
Dim P As Previo.clsPrevio
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter
Dim R As ADODB.Recordset
Dim lnContador As Integer

Dim sSQL As String
Dim sCadRep As String

Dim lsTarjeta As String * 16
Dim lsHora As String * 8
Dim lsDiaMes As String * 5
Dim lsImporte As String * 7
Dim lsTipoMoneda As String * 7
Dim lsItem As String * 4
Dim lsDesc As String * 20
Dim lsEstado As String * 9
Dim lsFechaSistema As String * 19
Dim lnCont As Integer

Dim lnImporte As Currency

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
    
    If Me.chkFechaMov.Value = 0 Then
        sSQL = "ATM_ReporteOpConfirmadasConcialiadas '" & Format(txtFechaInicial.Text, "yymmdd") & "','" & Format(txtFechaFinal.Text, "yymmdd") & "', '" & Trim(Str(lnMoneda)) & "', '1' "
    Else
        sSQL = "ATM_ReporteOpConfirmadasConcialiadas '" & Format(txtFechaInicial.Text, "YYYYMMDD") & "','" & Format(txtFechaFinal.Text, "YYYYMMDD") & "', '" & Trim(Str(lnMoneda)) & "', '2' "
    End If
    
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
    lsEstado = " Estado"
    lsFechaSistema = "Fecha Sistema"
    lnCont = 1
    
    'Cabecera
    sCadRep = sCadRep & Space(5) & "CMAC MAYNAS S.A." & Space(50) & "FECHA : " & Format(Now(), "dd/mm/yyyy hh:mm:ss") & Chr(10)
    sCadRep = sCadRep & Space(5) & "SIMACC-Tarjeta de Debito" & IIf(lnMoneda, " Soles", " Dolares") & Chr(10) & Chr(10)
    
    sCadRep = sCadRep & Space(35) & "Reporte De Operaciones Concilidas - Retencion entre " & Me.txtFechaInicial & " y " & Me.txtFechaFinal.Text & Chr(10) & Chr(10) & Chr(10)
    
    sCadRep = sCadRep & Space(5) & String(108, "-") & Chr(10)
    sCadRep = sCadRep & Space(5) & lsItem & "  " & lsTarjeta & "  " & lsHora & "  " & lsDiaMes & "  " & lsImporte & "  " & lsTipoMoneda & "  " & lsDesc & "  " & lsEstado & "  " & lsFechaSistema & Chr(10)
    sCadRep = sCadRep & Space(5) & String(108, "-") & Chr(10)
    
    'AbrirConexion
    loConec.AbreConexion
    R.Open sSQL, loConec.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
     
    Do While Not R.EOF
        lsItem = Format(lnCont, "#0000")
        lsTarjeta = R!PAN
        lsHora = Left(R!Hora, 2) & ":" & Mid(R!Hora, 3, 2) & ":" & Right(R!Hora, 2)
        lsDiaMes = Mid(R!Fecha, 3, 2) & "/" & Right(R!Fecha, 2)
        RSet lsImporte = Format(R!IMPORTE1 / 100, "#,##0.00")
        RSet lsImporte = Format(R!IMPORTE1, "#0.00")
        lsTipoMoneda = R!TipoMoneda
        lsDesc = R!TipoMov
        lsEstado = R!Condicion
        If R!cMovNro & "" = "" Then
            lsFechaSistema = ""
        Else
            lsFechaSistema = Left(R!cMovNro & "", 4) & "/" & Mid(R!cMovNro & "", 5, 2) & "/" & Mid(R!cMovNro & "", 7, 2) & " " & Mid(R!cMovNro & "", 9, 2) & ":" & Mid(R!cMovNro & "", 11, 2) & ":" & Mid(R!cMovNro & "", 13, 2)
        End If
        
        
        'lnImporte = lnImporte + R!IMPORTE1
        sCadRep = sCadRep & Space(5) & lsItem & "  " & lsTarjeta & "  " & lsHora & "  " & lsDiaMes & "  " & lsImporte & "  " & lsTipoMoneda & "  " & lsDesc & "  " & lsEstado & "  " & lsFechaSistema & Chr(10)
        If R!Condicion = "Realizada" Then
            lnImporte = lnImporte + R!IMPORTE1 '/ 100
        End If
        R.MoveNext
        lnContador = lnContador + 1
        lnCont = lnCont + 1
    Loop
    
    sCadRep = sCadRep & Space(5) & String(108, "-") & Chr(10)
    
    lsItem = ""
    lsTarjeta = Format(lnContador, "#0")
    lsHora = ""
    lsDiaMes = ""
    RSet lsImporte = Format(lnImporte, "#,##0.00")
    lsTipoMoneda = ""
    lsDesc = ""
    lsEstado = ""
    lsFechaSistema = ""
    
    sCadRep = sCadRep & Space(5) & lsItem & "  " & lsTarjeta & "  " & lsHora & "  " & lsDiaMes & "  " & lsImporte & "  " & lsTipoMoneda & "  " & lsDesc & "  " & lsEstado & "  " & lsFechaSistema & Chr(10)

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
