VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmCierreDiario 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cierre de Dia"
   ClientHeight    =   4680
   ClientLeft      =   3495
   ClientTop       =   2610
   ClientWidth     =   4950
   Icon            =   "frmCierreDiario.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPrecuadre 
      Caption         =   "&PreCuadre"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   210
      TabIndex        =   5
      Top             =   4200
      Width           =   1425
   End
   Begin VB.PictureBox Pic2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   900
      Left            =   120
      ScaleHeight     =   870
      ScaleWidth      =   4665
      TabIndex        =   12
      Top             =   120
      Width           =   4695
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CIERRE DE OPERACIONES"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   960
         TabIndex        =   13
         Top             =   240
         Width           =   3225
      End
      Begin VB.Image imgAlerta 
         Height          =   480
         Left            =   120
         Picture         =   "frmCierreDiario.frx":030A
         Top             =   120
         Width           =   480
      End
   End
   Begin VB.Frame fraFechas 
      Caption         =   "Fechas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   750
      Left            =   120
      TabIndex        =   9
      Top             =   1080
      Width           =   4605
      Begin MSMask.MaskEdBox txtFecFin 
         Height          =   375
         Left            =   2640
         TabIndex        =   1
         Top             =   270
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFecIni 
         Height          =   375
         Left            =   735
         TabIndex        =   0
         Top             =   270
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Al :"
         Height          =   195
         Left            =   2280
         TabIndex        =   11
         Top             =   360
         Width           =   225
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Del :"
         Height          =   195
         Left            =   180
         TabIndex        =   10
         Top             =   360
         Width           =   330
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   120
      ScaleHeight     =   2025
      ScaleWidth      =   4620
      TabIndex        =   8
      Top             =   1920
      Width           =   4650
      Begin VB.CheckBox ChkColoacSaldo 
         Caption         =   "Guarda Saldos Diarios"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   300
         Left            =   510
         TabIndex        =   15
         Top             =   405
         Width           =   3450
      End
      Begin VB.CheckBox CheckConsol 
         Caption         =   "Guarda Temporales"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   300
         Left            =   510
         TabIndex        =   14
         Top             =   90
         Width           =   3450
      End
      Begin VB.CheckBox ChkCierreCred 
         Caption         =   "Cierre de Creditos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   300
         Left            =   510
         TabIndex        =   3
         Top             =   1095
         Width           =   3165
      End
      Begin VB.CheckBox ChkCierrePig 
         Caption         =   "Cierre de Pignoraticio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   510
         TabIndex        =   4
         Top             =   1440
         Width           =   3450
      End
      Begin VB.CheckBox ChkCierreAho 
         Caption         =   "Cierre de Ahorros"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   300
         Left            =   510
         TabIndex        =   2
         Top             =   735
         Width           =   3450
      End
   End
   Begin VB.CommandButton CmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3240
      TabIndex        =   7
      Top             =   4200
      Width           =   1545
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1725
      TabIndex        =   6
      Top             =   4200
      Width           =   1425
   End
End
Attribute VB_Name = "frmCierreDiario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oConec As DConecta
Dim rsPersoneria As Recordset
Dim bCierreMes As Boolean
Dim lnFiltro As Integer
Dim lnBitUsarValidacion As Integer

Dim objDTS As DTS.Package
Dim objSteps As DTS.Step
Dim strServerName As String, strUsuarioSQL As String, strPasswordSQL As String, strBaseSQL As String
Dim strNameDTS As String
Dim nError As Long
Dim sSource As String, sDesc As String
Dim bExito As Boolean

Function ObtenerPassword() As String
Dim oConec As DConecta
Dim strCadenaConexion As String
Dim intPosI As Integer
Dim intPosF As Integer
Set oConec = New DConecta
    oConec.AbreConexion
    strCadenaConexion = oConec.CadenaConexion
    oConec.CierraConexion
    intPosI = InStr(1, strCadenaConexion, "Password")
    intPosF = InStr(intPosI, strCadenaConexion, ";")
    ObtenerPassword = Mid(strCadenaConexion, intPosI + Len("Password="), intPosF - (intPosI + Len("Password=")))
    Set oConec = Nothing
End Function

Function ObtenerUsuarioDatos() As String
Dim oConec As DConecta
Dim strCadenaConexion As String
Dim intPos As Integer
Dim intPosF As Integer
Set oConec = New DConecta
    oConec.AbreConexion
    strCadenaConexion = oConec.CadenaConexion
    oConec.CierraConexion
    intPos = InStr(1, strCadenaConexion, "User ID=")
    intPosF = InStr(intPos, strCadenaConexion, ";")
    ObtenerUsuarioDatos = Mid(strCadenaConexion, intPos + Len("User ID="), intPosF - (intPos + Len("User ID=")))
Set oConec = Nothing
End Function

Function ObtenerServidorDatos() As String
Dim oConec As DConecta
Dim strCadenaConexion As String
Dim intPos As Integer
Dim intPosF As Integer
Set oConec = New DConecta
    oConec.AbreConexion
    strCadenaConexion = oConec.CadenaConexion
    oConec.CierraConexion
    intPos = InStr(1, strCadenaConexion, "Data Source=")
    intPosF = InStr(intPos, strCadenaConexion, ";")
    ObtenerServidorDatos = Mid(strCadenaConexion, intPos + Len("Data Source="), intPosF - (intPos + Len("Data Source=")))
    
Set oConec = Nothing
End Function

Function ObtenerBaseDatos() As String
Dim oConec As DConecta
Dim strCadenaConexion As String
Dim intPos As Integer
Dim intPosF As Integer
Set oConec = New DConecta
    oConec.AbreConexion
    strCadenaConexion = oConec.CadenaConexion
    oConec.CierraConexion
    intPos = InStr(1, strCadenaConexion, "Initial Catalog=")
    intPosF = InStr(intPos, strCadenaConexion, ";")
    ObtenerBaseDatos = Mid(strCadenaConexion, intPos + Len("Initial Catalog="), intPosF - (intPos + Len("Initial Catalog=")))
    
Set oConec = Nothing
End Function

Private Sub GetDatosConexionDTS()
strServerName = ObtenerServidorDatos
strUsuarioSQL = ObtenerUsuarioDatos
strPasswordSQL = ObtenerPassword

End Sub
Public Sub CierreMes()
    bCierreMes = True
    Me.Caption = "Cierre de Mes"
    Me.Show 1
End Sub

Public Sub CierreDia()
    bCierreMes = False
    Me.Show 1
End Sub

'Private Sub CapGeneraEstadSaldoPersoneria()
'Dim nPers As PersPersoneria
'rsPersoneria.MoveFirst
'Do While Not rsPersoneria.EOF
'    nPers = rsPersoneria("nConsValor")
'    CapGeneraEstadSaldo sAge, sFec, nProd, nMon, sCMACT, sUsu, nPers
'    rsPersoneria.MoveNext
'Loop
'End Sub

Private Sub CapGeneraEstadMov(ByVal sAge As String, ByVal dFecha As Date, _
        ByVal nProd As Producto, ByVal nMon As Moneda, ByVal sUsu As String)

Dim nNumAper As Long, nNumCancAct As Long, nNumCancInact As Long
Dim nMonAper As Double, nMonCancAct As Double, nMonCancInact As Double
Dim nMonRetInact As Double, nMonRetInt As Double, nMonRet As Double
Dim nNumRet As Long, nNumDep As Long
Dim nMonDep As Double, nMonIntCap As Double, nSaldo As Double, nMonChq As Double
Dim sFecha As String, sCondicion As String
Dim oCap As DCapMantenimiento
Dim rsEstad As ADODB.Recordset
Dim sSql As String

sFecha = Format$(dFecha, "yyyymmdd") & "%"
sCondicion = "___" & sAge & Trim(nProd) & Trim(nMon) & "%"

nNumAper = 0: nNumCancAct = 0: nNumCancInact = 0: nMonAper = 0
nMonCancAct = 0: nMonCancInact = 0: nMonRetInact = 0: nMonRetInt = 0
nMonRet = 0: nNumRet = 0: nNumDep = 0: nMonDep = 0: nMonIntCap = 0
nSaldo = 0: nMonChq = 0

Set oCap = New DCapMantenimiento
'Calculamos el número y monto de cada tipo de movimiento
Set rsEstad = oCap.GetCapEstadMovTipo(sCondicion, sFecha)
If Not (rsEstad.EOF And rsEstad.BOF) Then
    nNumAper = rsEstad("nNumAper")
    nNumCancAct = rsEstad("nNumCancAct")
    nNumCancInact = rsEstad("nNumCancInact")
    nMonAper = rsEstad("nMonAper")
    nMonCancAct = rsEstad("nMonCancAct")
    nMonCancInact = rsEstad("nMonCancInact")
    nMonRetInact = rsEstad("nMonRetInact")
    nMonRetInt = rsEstad("nMonRetInt")
    nMonRet = rsEstad("nMonRet")
    nNumRet = rsEstad("nNumRet")
    nNumDep = rsEstad("nNumDep")
    nMonDep = rsEstad("nMonDep")
    nMonIntCap = rsEstad("nMonIntCap")
End If
'Calculamos El Saldo por Producto
Set rsEstad = oCap.GetCapEstadSaldoProd(sCondicion)
If Not (rsEstad.EOF And rsEstad.BOF) Then
    nSaldo = rsEstad("nSaldo")
End If
'Calculamos el monto en cheques movidos en la fecha indicada
Set rsEstad = oCap.GetCapEstadMovCheques(sFecha)
If Not (rsEstad.EOF And rsEstad.BOF) Then
    nMonChq = rsEstad("nMonChq")
End If
rsEstad.Close
Set rsEstad = Nothing
Set oCap = Nothing

'Ingresamos el Registro en la tabla de Estadistica
sSql = "Insert CapEstadMovimiento (cCodAge,dEstad,nProducto,nMoneda,nNumAper,nMonAper,nNumCanc,nMonCanc, " _
    & "nRetInt,nRetInac,nNumCancInac,nMonCancInac,nNumDep,nMonDep,nNumRet,nMonRet, nIntCap,nSaldo,nMonChq,cCodUsu) " _
    & "Values ('" & sAge & "','" & Format$(dFecha, "mm/dd/yyyy hh:mm:ss") & "'," & Trim(nProd) & "," & Trim(nMon) & ", " _
    & Trim(nNumAper) & ", " & Trim(nMonAper) & "," & Trim(nNumCancAct) & "," & Trim(nMonCancAct) & ", " _
    & Trim(nMonRetInt) & "," & Trim(nMonRetInact) & "," & Trim(nNumCancInact) & "," & Trim(nMonCancInact) & ", " _
    & Trim(nNumDep) & "," & Trim(nMonDep) & "," & Trim(nNumRet) & "," & Trim(nMonRet) & "," & Trim(nMonIntCap) & ", " _
    & Trim(nSaldo) & "," & Trim(nMonChq) & ",'" & sUsu & "')"
oConec.ConexionActiva.Execute sSql

End Sub

Private Sub CapGeneraEstadistica()
Dim rsAgencia As Recordset
Dim sAgencia As String, sSql As String, sFecha As String
Dim clsGen As DGeneral
Dim cmd As ADODB.Command
Dim prm As ADODB.Parameter
Dim bTrans As Boolean
Dim dFecha As Date
On Error GoTo ErrStore
Dim SqlSent As String


bTrans = False
Set clsGen = New DGeneral
Set rsAgencia = clsGen.GetAgencias()
Set clsGen = Nothing
sFecha = Format$(gdFecSis & " " & GetHoraServer(), "dd/mm/yyyy hh:mm:ss")
dFecha = CDate(sFecha)
bTrans = True
    
Do While Not rsAgencia.EOF
    sAgencia = rsAgencia("cAgeCod")
    ' Miguel 24 - Marzo - 2004
'    SqlSent = "Exec CapGeneraEstadMov '" & sAgencia & "', '" & sFecha & "', " & gCapAhorros & ", " & gMonedaNacional & ", '" & gsCodUser & "'"
'    oConec.Ejecutar SqlSent
'    SqlSent = "Exec CapGeneraEstadMov '" & sAgencia & "', '" & sFecha & "', " & gCapAhorros & ", " & gMonedaExtranjera & ", '" & gsCodUser & "'"
'    oConec.Ejecutar SqlSent
'    SqlSent = "Exec CapGeneraEstadMov '" & sAgencia & "', '" & sFecha & "', " & gCapPlazoFijo & ", " & gMonedaNacional & ", '" & gsCodUser & "'"
'    oConec.Ejecutar SqlSent
'    SqlSent = "Exec CapGeneraEstadMov '" & sAgencia & "', '" & sFecha & "', " & gCapPlazoFijo & ", " & gMonedaExtranjera & ", '" & gsCodUser & "'"
'    oConec.Ejecutar SqlSent
'    SqlSent = "Exec CapGeneraEstadMov '" & sAgencia & "', '" & sFecha & "', " & gCapCTS & ", " & gMonedaNacional & ", '" & gsCodUser & "'"
'    oConec.Ejecutar SqlSent
'    SqlSent = "Exec CapGeneraEstadMov '" & sAgencia & "', '" & sFecha & "', " & gCapCTS & ", " & gMonedaExtranjera & ", '" & gsCodUser & "'"
'    oConec.Ejecutar SqlSent
        CapGeneraEstadMov sAgencia, sFecha, gCapAhorros, gMonedaNacional, gsCodUser
        CapGeneraEstadMov sAgencia, sFecha, gCapAhorros, gMonedaExtranjera, gsCodUser
        CapGeneraEstadMov sAgencia, sFecha, gCapPlazoFijo, gMonedaNacional, gsCodUser
        CapGeneraEstadMov sAgencia, sFecha, gCapPlazoFijo, gMonedaExtranjera, gsCodUser
        CapGeneraEstadMov sAgencia, sFecha, gCapCTS, gMonedaNacional, gsCodUser
        CapGeneraEstadMov sAgencia, sFecha, gCapCTS, gMonedaExtranjera, gsCodUser
    rsAgencia.MoveNext
Loop

rsAgencia.Close
Set rsAgencia = Nothing

Set cmd = Nothing
Set cmd = New ADODB.Command
cmd.CommandText = "CapGeneraEstadSaldo"
cmd.CommandType = adCmdStoredProc
cmd.Name = "CapGeneraEstadSaldo"
Set prm = cmd.CreateParameter("dEstad", adDate, adParamInput)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("sCodCMACT", adChar, adParamInput, 4)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("sUsuario", adChar, adParamInput, 4)
cmd.Parameters.Append prm
Set cmd.ActiveConnection = oConec.ConexionActiva
cmd.CommandTimeout = 720
cmd.Parameters.Refresh

oConec.ConexionActiva.CapGeneraEstadSaldo sFecha, gsCodPersCMACT, gsCodUser

'rsAgencia.MoveFirst
'Do While Not rsAgencia.EOF
'    sAgencia = rsAgencia("cAgeCod")
'    CapGeneraEstadSaldoPersoneria sAgencia, sFecha, gCapAhorros, gMonedaNacional, gsCodPersCMACT, gsCodUser
'    CapGeneraEstadSaldoPersoneria sAgencia, sFecha, gCapAhorros, gMonedaExtranjera, gsCodPersCMACT, gsCodUser
'    CapGeneraEstadSaldoPersoneria sAgencia, sFecha, gCapPlazoFijo, gMonedaNacional, gsCodPersCMACT, gsCodUser
'    CapGeneraEstadSaldoPersoneria sAgencia, sFecha, gCapPlazoFijo, gMonedaExtranjera, gsCodPersCMACT, gsCodUser
'    CapGeneraEstadSaldoPersoneria sAgencia, sFecha, gCapCTS, gMonedaNacional, gsCodPersCMACT, gsCodUser
'    CapGeneraEstadSaldoPersoneria sAgencia, sFecha, gCapCTS, gMonedaExtranjera, gsCodPersCMACT, gsCodUser
'    rsAgencia.MoveNext
'Loop


Set cmd = Nothing
Set cmd = New ADODB.Command
cmd.CommandText = "CapGeneraSaldosDiarios"
cmd.CommandType = adCmdStoredProc
cmd.Name = "CapGeneraSaldosDiarios"
Set prm = cmd.CreateParameter("dEstad", adDate, adParamInput)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("sUsuario", adChar, adParamInput, 4)
cmd.Parameters.Append prm
Set cmd.ActiveConnection = oConec.ConexionActiva
cmd.CommandTimeout = 720
cmd.Parameters.Refresh
oConec.ConexionActiva.CapGeneraSaldosDiarios sFecha, gsCodUser

Set prm = Nothing
Set cmd = Nothing


bTrans = False
Exit Sub
ErrStore:
    If bTrans Then oConec.ConexionActiva.RollbackTrans
    Set oConec = Nothing
    MsgBox Err.Description, vbExclamation, "Error"
End Sub

Private Sub CargaDatos()
Dim sSql As String
Dim R As ADODB.Recordset
Dim oGen As DGeneral

    sSql = " select * from constsistema where nConsSisCod = 7"
    
    Set R = oConec.CargaRecordSet(sSql)
    If CDate(R!nConsSisValor) = gdFecSis Then
        ChkCierreAho.value = 1
        ChkCierreAho.Enabled = False
    End If
    R.Close
    
    sSql = " select * from constsistema where nConsSisCod = 11"
    Set R = oConec.CargaRecordSet(sSql)
    If CDate(R!nConsSisValor) = gdFecSis Then
        ChkCierreCred.value = 1
        ChkCierreCred.Enabled = False
    End If
    R.Close
    
    sSql = " select * from constsistema where nConsSisCod = 12"
    Set R = oConec.CargaRecordSet(sSql)
    If CDate(R!nConsSisValor) = gdFecSis Then
        ChkCierrePig.value = 1
        ChkCierrePig.Enabled = False
    End If
    R.Close
    bCierreMes = VerificaDiaHabil(gdFecSis, 3)
    If bCierreMes Then
        CheckConsol.Visible = True
    Else
        CheckConsol.Visible = False
    End If
    
    sSql = " select * from constsistema where nConsSisCod = 170"
    Set R = oConec.CargaRecordSet(sSql)
    If CDate(R!nConsSisValor) = gdFecSis Then
        CheckConsol.value = 1
        CheckConsol.Enabled = False
    End If
    R.Close
    
    
End Sub

Private Sub CierreAhorros()
Dim cmd As ADODB.Command
Dim prm As ADODB.Parameter
Dim sSql As String
    
    oConec.ConexionActiva.BeginTrans
    Set cmd = New ADODB.Command
    Set prm = New ADODB.Parameter
    cmd.CommandText = "CaptacCierreDiario"
    cmd.CommandType = adCmdStoredProc
    cmd.Name = "CaptacCierreDiario"
    Set prm = cmd.CreateParameter("FechaTran", adDate, adParamInput)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("Usuario", adChar, adParamInput, 4)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("Agencia", adChar, adParamInput, 2)
    cmd.Parameters.Append prm
    Set cmd.ActiveConnection = oConec.ConexionActiva
    cmd.CommandTimeout = 720
    cmd.Parameters.Refresh
    oConec.ConexionActiva.CaptacCierreDiario Format(gdFecSis & " " & GetHoraServer(), "yyyy-mm-dd hh:mm:ss"), gsCodUser, gsCodAge
    sSql = "UPDATE ConstSistema SET nConsSisValor = '" & Format(gdFecSis, "dd/mm/yyyy") & "' WHERE nConsSisCod = 7"
    oConec.ConexionActiva.Execute sSql
    oConec.ConexionActiva.CommitTrans
    oConec.ConexionActiva.BeginTrans
    If Not bCierreMes Then
        Call CapGeneraEstadistica
    End If

    
    oConec.ConexionActiva.CommitTrans
    Set cmd = Nothing
    Set prm = Nothing

End Sub

Private Sub SalvaColocacSaldo()
Dim cmd As New ADODB.Command
Dim prm As New ADODB.Parameter
Dim sSql As String
  
    
    'Salva Data A Consolidar
    If bCierreMes Then
            
        cmd.CommandText = "ColocCred_GeneraColocacSaldo"
        cmd.CommandType = adCmdStoredProc
        cmd.Name = "ColocCred_GeneraColocacSaldo"
        Set prm = cmd.CreateParameter("FechaTran", adDate, adParamInput)
        cmd.Parameters.Append prm
        Set cmd.ActiveConnection = oConec.ConexionActiva
        cmd.CommandTimeout = 72000
        cmd.Parameters.Refresh
        oConec.ConexionActiva.ColocCred_GeneraColocacSaldo Format(gdFecSis & " " & GetHoraServer(), "dd/mm/yyyy hh:mm:ss")
        Set cmd = Nothing
        Set prm = Nothing
                
    End If

End Sub

Private Sub GuardaConsolidada()
Dim cmd As New ADODB.Command
Dim prm As New ADODB.Parameter
Dim sSql As String

    
    
    'Salva Data A Consolidar
    If bCierreMes Then
        oConec.ConexionActiva.CommandTimeout = 72000
        cmd.CommandText = "GuardarDataAConsolidarCreditos"
        cmd.CommandType = adCmdStoredProc
        cmd.Name = "GuardarDataAConsolidarCreditos"
        Set prm = cmd.CreateParameter("FechaTran", adDate, adParamInput)
        cmd.Parameters.Append prm
        Set cmd.ActiveConnection = oConec.ConexionActiva
        cmd.CommandTimeout = 72000
        cmd.Parameters.Refresh
        oConec.ConexionActiva.GuardarDataAConsolidarCreditos Format(gdFecSis & " " & GetHoraServer(), "dd/mm/yyyy hh:mm:ss")
        Set cmd = Nothing
        Set prm = Nothing
    End If

End Sub

Private Sub GuardaConsolidadaDTS()
'CREADO POR CMACICA AUTOR:LMMD

    strNameDTS = "DTSCierreCreditosConsolidada"

Set objDTS = New DTS.Package

    objDTS.LoadFromSQLServer strServerName, strUsuarioSQL, strPasswordSQL, DTSSQLStgFlag_Default, _
                              , , , strNameDTS
                              
    objDTS.GlobalVariables("gFechaHora").value = Format(gdFecSis, "mm/dd/yyyy")
    objDTS.GlobalVariables("gCodUsers").value = gsCodUser
    objDTS.GlobalVariables("gCodAgen").value = Right(gsCodAge, 2)
    
                       
    objDTS.Execute
    bExito = True
    For Each objSteps In objDTS.Steps
        objSteps.ExecuteInMainThread = True
        If objSteps.ExecutionResult = DTSStepExecResult_Failure Then
            objSteps.GetExecutionErrorInfo nError, sSource, sDesc
            MsgBox "Error Cierre Créditos:" & objSteps.Description & " " & sDesc, vbExclamation, "Error"
            bExito = False
            Exit For
        End If
    Next
    objDTS.UnInitialize
    Set objDTS = Nothing
End Sub

Private Sub GuardaColocacSaldoDTS()

    strNameDTS = "DTSCierreCreditosColocacSaldo"

    Set objDTS = New DTS.Package

    objDTS.LoadFromSQLServer strServerName, strUsuarioSQL, strPasswordSQL, DTSSQLStgFlag_Default, _
                              , , , strNameDTS
                              
    objDTS.GlobalVariables("gFechaHora").value = Format(gdFecSis, "dd/mm/yyyy hh:mm:ss")
    objDTS.GlobalVariables("gCodUsers").value = gsCodUser
    objDTS.GlobalVariables("gCodAgen").value = Right(gsCodAge, 2)
    
                       
    objDTS.Execute
    bExito = True
    For Each objSteps In objDTS.Steps
        objSteps.ExecuteInMainThread = True
        If objSteps.ExecutionResult = DTSStepExecResult_Failure Then
            objSteps.GetExecutionErrorInfo nError, sSource, sDesc
            MsgBox "Error Cierre Créditos:" & objSteps.Description & " " & sDesc, vbExclamation, "Error"
            bExito = False
            Exit For
        End If
    Next
    objDTS.UnInitialize
    Set objDTS = Nothing
End Sub


Private Sub CierreCreditos()
Dim cmd As New ADODB.Command
Dim prm As New ADODB.Parameter
Dim sSql As String

    oConec.ConexionActiva.BeginTrans
    
    'Salva Data A Consolidar
'    If bCierreMes Then
'
'        cmd.CommandText = "ColocCred_GeneraColocacSaldo"
'        cmd.CommandType = adCmdStoredProc
'        cmd.Name = "ColocCred_GeneraColocacSaldo"
'        Set prm = cmd.CreateParameter("FechaTran", adDate, adParamInput)
'        cmd.Parameters.Append prm
'        Set cmd.ActiveConnection = oConec.ConexionActiva
'        cmd.CommandTimeout = 1420
'        cmd.Parameters.Refresh
'        oConec.ConexionActiva.ColocCred_GeneraColocacSaldo Format(gdFecSis & " " & GetHoraServer(), "dd/mm/yyyy hh:mm:ss")
'        Set cmd = Nothing
'        Set prm = Nothing
'
'        cmd.CommandText = "GuardarDataAConsolidarCreditos"
'        cmd.CommandType = adCmdStoredProc
'        cmd.Name = "GuardarDataAConsolidarCreditos"
'        Set prm = cmd.CreateParameter("FechaTran", adDate, adParamInput)
'        cmd.Parameters.Append prm
'        Set cmd.ActiveConnection = oConec.ConexionActiva
'        cmd.CommandTimeout = 1420
'        cmd.Parameters.Refresh
'        oConec.ConexionActiva.GuardarDataAConsolidarCreditos Format(gdFecSis & " " & GetHoraServer(), "dd/mm/yyyy hh:mm:ss")
'        Set cmd = Nothing
'        Set prm = Nothing
'
'    End If
       
    
    Set cmd = New ADODB.Command
    cmd.CommandText = "ColocCredCierreDiario"
    cmd.CommandType = adCmdStoredProc
    cmd.Name = "ColocCredCierreDiario"
    Set prm = cmd.CreateParameter("FechaTran", adDate, adParamInput)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("Usuario", adChar, adParamInput, 4)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("Agencia", adChar, adParamInput, 2)
    cmd.Parameters.Append prm
    Set cmd.ActiveConnection = oConec.ConexionActiva
    cmd.CommandTimeout = 720
    cmd.Parameters.Refresh
    oConec.ConexionActiva.ColocCredCierreDiario Format(gdFecSis & " " & GetHoraServer(), "dd/mm/yyyy hh:mm:ss"), gsCodUser, Right(gsCodAge, 2)
    
    'Estadistica diaria 12/03/2004
    
    Set cmd = New ADODB.Command
    cmd.CommandText = "ColocCredCierreEstadisticaDiaria"
    cmd.CommandType = adCmdStoredProc
    cmd.Name = "ColocCredCierreEstadisticaDiaria"
    Set prm = cmd.CreateParameter("FechaTran", adDate, adParamInput)
    cmd.Parameters.Append prm
    Set cmd.ActiveConnection = oConec.ConexionActiva
    cmd.CommandTimeout = 720
    cmd.Parameters.Refresh
    oConec.ConexionActiva.ColocCredCierreEstadisticaDiaria Format(gdFecSis & " " & GetHoraServer(), "dd/mm/yyyy hh:mm:ss")
    
    'Fin de estadistica diaria
    
    
    sSql = " UPDATE ConstSistema SET nConsSisValor = '" & Format(gdFecSis, "dd/mm/yyyy") & "' WHERE nConsSisCod = 11"
    oConec.ConexionActiva.Execute sSql
    
    oConec.ConexionActiva.CommitTrans
    
    Set cmd = Nothing
    Set prm = Nothing
 
        
    
End Sub

Private Sub CierrePignoraticio()
Dim cmd As New ADODB.Command
Dim prm As New ADODB.Parameter
Dim sSql As String
On Error GoTo mmm
    oConec.ConexionActiva.BeginTrans
    
    'If lnFiltro = 1 Then    'Trujillo
        cmd.CommandText = "ColocPigCierreDiario"
    'ElseIf lnFiltro = 2 Then    'Lima
     '   cmd.CommandText = "ColocPignoCierreDiario"
    'End If
    cmd.CommandType = adCmdStoredProc
    cmd.Name = "ColocPigCierreDiario"
    Set prm = cmd.CreateParameter("FechaTran", adDate, adParamInput)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("Usuario", adChar, adParamInput, 4)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("Agencia", adChar, adParamInput, 5)
    cmd.Parameters.Append prm
    Set cmd.ActiveConnection = oConec.ConexionActiva
    cmd.CommandTimeout = 720
    cmd.Parameters.Refresh
    oConec.ConexionActiva.ColocPigCierreDiario Format(gdFecSis & " " & GetHoraServer(), "dd/mm/yyyy hh:mm:ss"), gsCodUser, gsCodAge
    
    sSql = " UPDATE ConstSistema SET nConsSisValor = '" & Format(gdFecSis, "dd/mm/yyyy") & "' WHERE nConsSisCod = 12"
    oConec.ConexionActiva.Execute sSql
    
    oConec.ConexionActiva.CommitTrans
    
    Set cmd = Nothing
    Set prm = Nothing
Exit Sub
mmm:
MsgBox Err.Description
End Sub

Private Sub CmdAceptar_Click()
Dim sSql As String
Dim cmd As ADODB.Command
Dim prm As ADODB.Parameter
Dim nDias As Integer, i As Integer


If MsgBox("Esta Ud. seguro de efectuar el cierre?", vbQuestion + vbYesNo, "Advertencia") = vbNo Then
    Exit Sub
End If



strBaseSQL = ObtenerBaseDatos
sSql = " master..sp_dboption " & strBaseSQL & ",'trunc. log on chkpt.',true"
oConec.ConexionActiva.Execute sSql

nDias = DateDiff("d", CDate(txtFecIni), CDate(txtFecFin)) + 1
For i = 1 To nDias
    Screen.MousePointer = 11
    If i > 1 Then frmInicioDia.InicioDia
    bCierreMes = VerificaDiaHabil(gdFecSis, 3)
    GetDatosConexionDTS
    If bCierreMes Then
'        If CheckColocSaldo.value = 1 Then
'            Call SalvaColocacSaldo
'            MsgBox "Proceso Consolida Saldos Finalizado"
'        End If
        
        If CheckConsol.value = 1 And CheckConsol.Enabled = True Then
            'Call GuardaConsolidada
            Call GuardaConsolidadaDTS
            sSql = " UPDATE ConstSistema SET nConsSisValor = '" & Format(gdFecSis, "dd/mm/yyyy") & "' WHERE nConsSisCod = 170"
            oConec.ConexionActiva.Execute sSql
            MsgBox "Proceso Consolida Temporales Finalizado", vbInformation, "Aviso de Cierre"
        End If
        
    End If
    
    If ChkColoacSaldo.value = 1 And ChkColoacSaldo.Enabled = True Then
        Call GuardaColocacSaldoDTS
        sSql = " UPDATE ConstSistema SET nConsSisValor = '" & Format(gdFecSis, "dd/mm/yyyy") & "' WHERE nConsSisCod = 171"
        oConec.ConexionActiva.Execute sSql
        MsgBox "Proceso Consolida Saldos de Cartera", vbInformation, "Aviso de Cierre"
    End If
    
    If ChkCierreAho.value = 1 And ChkCierreAho.Enabled = True Then
        Call CierreAhorros
        If nDias = 1 Then MsgBox "Cierre de Ahorros Finalizado", vbInformation, "Aviso de Cierre"
    End If
    If ChkCierreCred.value = 1 And ChkCierreCred.Enabled = True Then
        Call CierreCreditos
        If nDias = 1 Then MsgBox "Cierre de Creditos Finalizado", vbInformation, "Aviso de Cierre"
    End If
    If ChkCierrePig.value = 1 And ChkCierrePig.Enabled = True Then
        Call CierrePignoraticio
        If nDias = 1 Then MsgBox "Cierre de Pignoraticio Finalizado", vbInformation, "Aviso de Cierre"
    End If
    
    If ChkCierrePig.value = 1 And ChkCierreAho.value = 1 And ChkCierreCred.value = 1 Then
        sSql = " UPDATE ConstSistema SET nConsSisValor = '" & Format(gdFecSis, "dd/mm/yyyy") & "' WHERE nConsSisCod = 13"
        oConec.ConexionActiva.Execute sSql
    End If
    MsgBox "Cierre del " & Format$(gdFecSis, "dd/mm/yyyy") & " Finalizado con éxito", vbInformation, "Aviso"
    
    If bCierreMes Then
        If MsgBox("El Sistema ha detectado que la Fecha Actual " & Format$(gdFecSis, "dd mmmm yyyy") & " es el último día" & Chr(13) + gPrnSaltoLinea _
            & "hábil del mes y por ello recomienda efectual el proceso de " & Chr(13) + gPrnSaltoLinea _
            & "CIERRE DE MES. Desea realizar el Cierre de Mes??", vbQuestion + vbYesNo, "Aviso") = vbYes Then
            
            oConec.ConexionActiva.BeginTrans
            
            Set cmd = New ADODB.Command
            Set prm = New ADODB.Parameter
            cmd.CommandText = "CaptacCierreMes"
            cmd.CommandType = adCmdStoredProc
            cmd.Name = "CaptacCierreMes"
            Set prm = cmd.CreateParameter("FechaTran", adDate, adParamInput)
            cmd.Parameters.Append prm
            Set prm = cmd.CreateParameter("Usuario", adChar, adParamInput, 4)
            cmd.Parameters.Append prm
            Set prm = cmd.CreateParameter("Agencia", adChar, adParamInput, 2)
            cmd.Parameters.Append prm
            Set cmd.ActiveConnection = oConec.ConexionActiva
            cmd.CommandTimeout = 720
            cmd.Parameters.Refresh
            oConec.ConexionActiva.CaptacCierreMes Format(gdFecSis & " " & GetHoraServer(), "dd/mm/yyyy hh:mm:ss"), gsCodUser, Right(gsCodAge, 2)
            
            sSql = " UPDATE ConstSistema SET nConsSisValor = '" & Format(gdFecSis, "dd/mm/yyyy") & "' WHERE nConsSisCod = 14"
            oConec.ConexionActiva.Execute sSql
            
            '''
            oConec.ConexionActiva.CommitTrans
            oConec.ConexionActiva.BeginTrans
            '''
            
            Call CapGeneraEstadistica
            oConec.ConexionActiva.CommitTrans
            'Cierre de Recuperaciones  ' LAYG
            Call CierreRecuperaciones
            
            'Salva Data A Consolidar
            If bCierreMes Then
                Set cmd = New ADODB.Command
                cmd.CommandText = "GuardaDataAConsolidarAhorros"
                cmd.CommandType = adCmdStoredProc
                cmd.Name = "GuardaDataAConsolidarAhorros"
                Set cmd.ActiveConnection = oConec.ConexionActiva
                cmd.CommandTimeout = 72000
                cmd.Parameters.Refresh
                oConec.ConexionActiva.GuardaDataAConsolidarAhorros
            End If
            Set cmd = Nothing
            Set prm = Nothing
        End If
    End If
    Screen.MousePointer = 0
Next i
sSql = " master..sp_dboption " & strBaseSQL & ",'trunc. log on chkpt.',false"
oConec.ConexionActiva.Execute sSql

MsgBox "Proceso de Cierre Finalizado", vbExclamation, "Aviso"
Unload Me
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdPreCuadre_Click()
Dim oPrevio As Previo.clsPrevio
Dim nBandera As Boolean
Dim sCadena As String

On Error GoTo err_
nBandera = False

sCadena = Devuelve_Errores_PreCuadre(txtFecIni.Text)
MsgBox "Precuadre ha culminado con exito." & Chr(13) & "Por favor revizar el reporte con detenimiento", vbInformation, "Aviso"
If Len(Trim(sCadena)) = 0 Then
    nBandera = True
Else
    nBandera = False
    Set oPrevio = New Previo.clsPrevio
    oPrevio.Show sCadena, "Errores de Precuadre"
    Set oPrevio = Nothing
End If
    
'If nBandera = True Then
    CmdAceptar.Enabled = True
'Else
'    If lnBitUsarValidacion = 1 Then
'        CmdAceptar.Enabled = False
'    Else
'        If MsgBox("El precuadre no pasó satisfactoriamente" & Chr(13) & Chr(13) & "Desea Ud. efectuar el cierre de todas maneras?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
'            CmdAceptar.Enabled = True
'        Else
'            CmdAceptar.Enabled = False
'        End If
'    End If
'End If
Exit Sub

err_:
    MsgBox "Error en precuadre", vbExclamation, "Aviso"
    Exit Sub
End Sub

Private Sub Form_Load()
Dim oGen As DGeneral
Me.Icon = LoadPicture(App.path & gsRutaIcono)
Set oGen = New DGeneral
lnFiltro = CInt(oGen.LeeConstSistema(104))
'lnBitUsarValidacion = CInt(oGen.LeeConstSistema(8005))
lnBitUsarValidacion = 1

Set oGen = Nothing
    
CentraForm Me
Set oConec = New DConecta
oConec.AbreConexion
Call CargaDatos
txtFecIni = gdFecSis
txtFecFin = gdFecSis
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    oConec.CierraConexion
    Set oConec = Nothing
End Sub

Private Sub TxtFecFin_GotFocus()
With txtFecFin
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub

Private Sub TxtFecFin_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    CmdAceptar.SetFocus
End If
End Sub

Private Sub txtFecIni_GotFocus()
With txtFecIni
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub

Private Sub TxtFecIni_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtFecFin.SetFocus
End If


End Sub


Private Sub CierreRecuperaciones()
Dim loRecup As NColRecCredito
Dim lsMovNro As String
Dim lsFechaHoraGrab As String
Dim loContFunct As NContFunciones
Dim loParam As NConstSistemas
Dim lnTipoCalcIntComp As Integer, lnTipoCalcIntMora   As Integer
Set loParam = New NConstSistemas
    lnTipoCalcIntComp = loParam.LeeConstSistema(151)
    lnTipoCalcIntMora = loParam.LeeConstSistema(152)
Set loParam = Nothing

    'Genera el Mov Nro
    Set loContFunct = New NContFunciones
        lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Set loContFunct = Nothing
    
    lsFechaHoraGrab = fgFechaHoraGrab(lsMovNro)
    
Set loRecup = New NColRecCredito

    Call loRecup.nCierreMesRecuperaciones(lsFechaHoraGrab, "131000", lsMovNro, lnTipoCalcIntComp, lnTipoCalcIntMora)
    
Set loRecup = Nothing
End Sub


Private Function Devuelve_Errores_PreCuadre(ByVal sFecha As String) As String

Dim sSql As String
Dim sCadena As String
Dim nBandera As Integer
Dim sCabecera As String
Dim rs As New ADODB.Recordset
Dim vLenNomb As Integer
Dim vespacio As Integer
Dim vPage As Integer
Dim lsTablaDiaria As String

'''''''''''''''''''''

If CDate(sFecha) = gdFecSis Then
    lsTablaDiaria = " MovDiario "
Else
    OpeDiaCreaTemporal Format(CDate(sFecha), "yyyymmdd"), gsCodUser
    lsTablaDiaria = " [dbo].[##Mov_" & gsCodUser & "] "
End If
''''''''''''''''''''

nBandera = 1
sCadena = ""
vLenNomb = 70
vespacio = vLenNomb + 54
vPage = 1

sCabecera = gsNomCmac & Space(vLenNomb + 6) & Format(gdFecSis & " " & Time, "dd/mm/yyyy hh:mm:ss") & Chr(10)
sCabecera = sCabecera & Space(vLenNomb + 16) & " Página :" & ImpreFormat(vPage, 5, 0) & Chr(10)
sCabecera = sCabecera & ImpreFormat(UCase(gsNomAge), 25) & Chr(10)
sCabecera = sCabecera & ImpreFormat(" P R E C U A D R E  DEL DIA " & Format(sFecha, "dd/mm/yyyy"), 44, 43) & Chr(10)
sCabecera = sCabecera & ImpreFormat(String(40, "="), 44, 42) & Chr(10) & Chr(10)
    
'Validaciones de Captaciones
'===========================

'Verificando que no existan cuentas de cmacs que no esten en CtaIfEspecial

sCadena = sCadena & sCabecera
sCadena = sCadena & ImpreFormat("  ( PRE CUADRE   DE   CAPTACIONES  )  ", 44, 43) & Chr(10) & Chr(10)

sSql = "Select a.cCtaCod from captaciones a "
sSql = sSql & " Inner join producto b on a.cctacod = b.cctacod "
sSql = sSql & " inner join productopersona c on a.cctacod = c.cctacod "
sSql = sSql & " inner join persona  d on c.cperscod = d.cperscod "
sSql = sSql & " where npersoneria = " & gPersonaJurCFLCMAC & " and left(nprdestado,2) not in (13,14) "
sSql = sSql & " and a.cctacod not in (select cCtaCod from cuentaifespecial) "

Set rs = oConec.CargaRecordSet(sSql)
If rs.BOF Then
Else
    nBandera = 2
    sCadena = sCadena & "** Existen Cuentas de CMACS no definidas en CtaIfEspecial" & Chr(10)
    sCadena = sCadena & "   ______________________________________________________" & Chr(10) & Chr(10)
    sCadena = sCadena & "   ** Cuentas            ** " & Chr(10)
    sCadena = sCadena & "   ** ================== ** " & Chr(10) & Chr(10)
                               
    Do While Not rs.EOF
        sCadena = sCadena & "   *  " & rs!cCtaCod & Chr(10)
        rs.MoveNext
    Loop
End If
rs.Close
Set rs = Nothing

'Valida Diferencias entre Cabecera y detalle Captaciones

sSql = " Select MC.nMovNro, MC.nMonto, SUM(MCD.nMonto) Detalle "
sSql = sSql & " From " & lsTablaDiaria & " M JOIN MovCap MC JOIN MovCapDet MCD ON MC.nMovNro = MCD.nMovNro "
sSql = sSql & " And MC.cCtaCod = MCD.cCtaCod And MC.cOpeCod = MCD.cOpeCod ON M.nMovNro = MC.nMovNro "
sSql = sSql & " Where M.cMovNro LIKE '" & Format(sFecha, "YYYYMMdd") & "%' And M.nMovFlag = 0 And MC.cCtaCod LIKE '111_____[12]%' "
sSql = sSql & " And (MCD.cOpeCod + Convert(Varchar(6),MCD.nConceptoCod)) "
sSql = sSql & " IN (Select cOpeCod + Convert(Varchar(6),nConcepto) From OpeCtaNeg Where cCtaContCod LIKE '11_102%') "
sSql = sSql & " Group by MC.nMovNro, MC.nMonto Having MC.nMonto <> SUM(MCD.nMonto) "

Set rs = oConec.CargaRecordSet(sSql)
If rs.BOF Then
Else
    nBandera = 2
    sCadena = sCadena & Chr(10) & Chr(10)
    sCadena = sCadena & "** Existen descuadres en el movCap y CovCapDet " & Chr(10)
    sCadena = sCadena & "   ___________________________________________ " & Chr(10) & Chr(10)
    sCadena = sCadena & "   ** nMovNro    ** " & Chr(10)
    sCadena = sCadena & "   ** ========== ** " & Chr(10) & Chr(10)
    
    Do While Not rs.EOF
        sCadena = sCadena & "   *  " & Left(rs!nMovNro & "          ", 10) & Chr(10)
        rs.MoveNext
    Loop
End If
rs.Close
Set rs = Nothing

' Valida la Caja, Asientos VS Planilla Consolidada Captaciones

sSql = " Select A.*, B.* From "
sSql = sSql & " (Select MC.nMovNro, nDebe = ISNULL(SUM(CASE WHEN O.cIngEgr = 'I' THEN Abs(MC.nMonto) END),0), "
sSql = sSql & " nHaber = ISNULL(SUM(CASE WHEN O.cIngEgr = 'E' THEN Abs(MC.nMonto) END),0) "
sSql = sSql & " From " & lsTablaDiaria & " M JOIN MovCap MC ON M.nMovNro = MC.nMovNro JOIN "
sSql = sSql & " (Select O.cOpeCod, G.cIngEgr From OpeTpo O JOIN GruposOpe GO JOIN GrupoOpe G ON "
sSql = sSql & " GO.cGrupoCod = G.cGrupoCod ON O.cOpeCod = GO.cOpeCod Where G.nEfectivo = 1 Group by O.cOpeCod, G.cIngEgr) O "
sSql = sSql & " ON MC.cOpeCod = O.cOpeCod "
sSql = sSql & " Where M.cMovNro LIKE '" & Format(sFecha, "YYYYMMdd") & "%' And M.nMovFlag = 0 And MC.cCtaCod LIKE '111_____[12]%' "
sSql = sSql & " Group by MC.nMovNro) A JOIN "
sSql = sSql & " (Select MC.nMovNro as nMovNro1, "
sSql = sSql & " nDebe1 = ISNULL(SUM(CASE WHEN C.cOpeCtaDH = 'D' THEN Abs(MCD.nMonto) END),0), "
sSql = sSql & " nHaber1 = ISNULL(SUM(CASE WHEN C.cOpeCtaDH = 'H' THEN Abs(MCD.nMonto) END),0) "
sSql = sSql & " From " & lsTablaDiaria & " M JOIN MovCap MC JOIN MovCapDet MCD "
sSql = sSql & " JOIN (Select cOpeCod, nConcepto, cOpeCtaDH From OpeCtaNeg Where cCtaContCod LIKE '11_102%' Group by "
sSql = sSql & " cOpeCod, nConcepto, cOpeCtaDH) C ON MCD.cOpeCod = C.cOpeCod And MCD.nConceptoCod = C.nConcepto "
sSql = sSql & " ON MC.nMovNro = MCD.nMovNro And MC.cCtaCod = MCD.cCtaCod And MC.cOpeCod = MCD.cOpeCod ON M.nMovNro = MC.nMovNro "
sSql = sSql & " Where M.cMovNro LIKE '" & Format(sFecha, "YYYYMMdd") & "%' And M.nMovFlag = 0 And MC.cCtaCod LIKE '1110____[12]%' "
sSql = sSql & " Group by MC.nMovNro) B ON A.nMovNro = B.nMovNro1 "
sSql = sSql & " Where (A.nDebe <> B.nDebe1 Or A.nHaber <> B.nHaber1)"

Set rs = oConec.CargaRecordSet(sSql)
If rs.BOF Then
Else
    nBandera = 2
    sCadena = sCadena & Chr(10) & Chr(10)
    sCadena = sCadena & "** Existen descuadres entre la Caja, Asientos y Planilla Consolidada Captaciones " & Chr(10)
    sCadena = sCadena & "   _____________________________________________________________________________ " & Chr(10) & Chr(10)
    sCadena = sCadena & "   ** nMovNro    Debe         nMovNro    Debe          ** " & Chr(10)
    sCadena = sCadena & "   ** ================================================ ** " & Chr(10) & Chr(10)
    
    Do While Not rs.EOF
        sCadena = sCadena & "   *  " & Left(rs!nMovNro & "          ", 10) & " " & ImpreFormat(rs!nDebe, 10, 2) & " " & ImpreFormat(rs!nhaber, 10, 2)
        sCadena = sCadena & " " & Left(rs!nMovNro1 & "          ", 10) & " " & ImpreFormat(rs!ndebe1, 10, 2) & " " & ImpreFormat(rs!nhaber1, 10, 2) & Chr(10)
        rs.MoveNext
    Loop
End If
rs.Close
Set rs = Nothing


'Validaciones de Colocaciones
'===========================
vPage = vPage + 1
sCadena = sCadena & Chr(12)
sCadena = sCadena & sCabecera
sCadena = sCadena & ImpreFormat("  ( PRE CUADRE   DE   COLOCACIONES )  ", 44, 43) & Chr(10) & Chr(10)

'Valida diferencias entre Cabecera detalle

sSql = " Select MC.nMovNro, MC.nMonto, SUM(MCD.nMonto) Detalle "
sSql = sSql & " From " & lsTablaDiaria & " M JOIN MovCol MC JOIN MovColDet MCD ON MC.nMovNro = MCD.nMovNro"
sSql = sSql & " And MC.cCtaCod = MCD.cCtaCod And MC.cOpeCod = MCD.cOpeCod ON M.nMovNro = MC.nMovNro"
sSql = sSql & " Where M.cMovNro LIKE '" & Format(sFecha, "YYYYMMdd") & "%' And M.nMovFlag = 0 And MC.cCtaCod LIKE '111_____[12]%' "
sSql = sSql & " And (MCD.cOpeCod + Convert(Varchar(6),MCD.nPrdConceptoCod)) "
sSql = sSql & " IN (Select cOpeCod + Convert(Varchar(6),nConcepto) From OpeCtaNeg Where cCtaContCod LIKE '11_102%') "
sSql = sSql & " Group by MC.nMovNro, MC.nMonto Having MC.nMonto <> SUM(MCD.nMonto)"

Set rs = oConec.CargaRecordSet(sSql)
If rs.BOF Then
Else
    nBandera = 2
    sCadena = sCadena & "** Existen descuadres en el movCol y CovColDet " & Chr(10)
    sCadena = sCadena & "   ___________________________________________ " & Chr(10) & Chr(10)
    sCadena = sCadena & "   ** nMovNro    ** " & Chr(10)
    sCadena = sCadena & "   ** ========== ** " & Chr(10) & Chr(10)
    
    Do While Not rs.EOF
        sCadena = sCadena & "   *  " & Left(rs!nMovNro & "          ", 10) & Chr(10)
        rs.MoveNext
    Loop
End If
rs.Close
Set rs = Nothing

'Valida que el saldo de calendario sea igual que el saldo de producto

sSql = "Select P.cCtaCod, P.nSaldo, "
sSql = sSql & " nSaldoCal = ( "
sSql = sSql & " Select SUM(nMonto) "
sSql = sSql & " From ( "
sSql = sSql & " Select SUM(nMonto-nMontoPagado) as nMonto "
sSql = sSql & " From ColocCalendDet CD "
sSql = sSql & " Inner Join ColocCalendario Cal ON Cal.cCtaCod = CD.cCtaCod AND Cal.nNroCalen = CD.nNroCalen AND Cal.nColocCalendEstado = 0"
sSql = sSql & " AND Cal.nColocCalendApl = CD.nColocCalendApl AND Cal.nCuota = CD.nCuota AND CD.nPrdConceptoCod = 1000 "
sSql = sSql & " Where CD.nNroCalen = CC.nNroCalen And CD.cCtaCod = CC.cCtaCod "
sSql = sSql & " AND CD.nColocCalendApl = 1 "
sSql = sSql & " Union All "
sSql = sSql & " Select ISNULL(SUM(nMonto-nMontoPagado),0) as nMonto "
sSql = sSql & " From ColocCalendDet CD "
sSql = sSql & " Inner Join ColocCalendario Cal ON Cal.cCtaCod = CD.cCtaCod AND Cal.nNroCalen = CD.nNroCalen AND Cal.nColocCalendEstado = 0"
sSql = sSql & " AND Cal.nColocCalendApl = CD.nColocCalendApl AND Cal.nCuota = CD.nCuota AND CD.nPrdConceptoCod = 1000 "
sSql = sSql & " Where CD.nNroCalen = CC.nNroCalPar And CD.cCtaCod = CC.cCtaCod "
sSql = sSql & " AND CD.nColocCalendApl = 1 "
sSql = sSql & " ) as T ), "
sSql = sSql & " TIPO=Case CC.CRFA"
sSql = sSql & " When 'RFA' Then 'RFA'"
sSql = sSql & " When 'RFC' Then 'RFC'"
sSql = sSql & " When 'DIF' Then 'DIF'"
sSql = sSql & " When 'NOR' Then 'NOR'"
sSql = sSql & " End"
sSql = sSql & " from Producto P "
sSql = sSql & " Inner Join ColocacCred CC ON CC.cCtaCod = P.cCtaCod  "
sSql = sSql & " Where P.nPrdEstado in (2020,2021,2022,2030,2031,2032) AND CC.CRFA not in ('RFA','RFC','DIF') "
sSql = sSql & " AND "
sSql = sSql & " P.nSaldo <> "
sSql = sSql & " ( "
sSql = sSql & " Select SUM(nMonto) "
sSql = sSql & " From ( "
sSql = sSql & " Select SUM(nMonto-nMontoPagado) as nMonto "
sSql = sSql & " From ColocCalendDet CD "
sSql = sSql & " Inner Join ColocCalendario Cal ON Cal.cCtaCod = CD.cCtaCod AND Cal.nNroCalen = CD.nNroCalen AND Cal.nColocCalendEstado = 0"
sSql = sSql & " AND Cal.nColocCalendApl = CD.nColocCalendApl AND Cal.nCuota = CD.nCuota AND CD.nPrdConceptoCod = 1000 "
sSql = sSql & " Where CD.nNroCalen = CC.nNroCalen And CD.cCtaCod = CC.cCtaCod "
sSql = sSql & " AND CD.nColocCalendApl = 1 "
sSql = sSql & " Union All "
sSql = sSql & " Select ISNULL(SUM(nMonto-nMontoPagado),0) as nMonto "
sSql = sSql & " From ColocCalendDet CD "
sSql = sSql & " Inner Join ColocCalendario Cal ON Cal.cCtaCod = CD.cCtaCod AND Cal.nNroCalen = CD.nNroCalen AND Cal.nColocCalendEstado = 0 "
sSql = sSql & " AND Cal.nColocCalendApl = CD.nColocCalendApl AND Cal.nCuota = CD.nCuota AND CD.nPrdConceptoCod = 1000 "
sSql = sSql & " Where CD.nNroCalen = CC.nNroCalPar And CD.cCtaCod = CC.cCtaCod "
sSql = sSql & " AND CD.nColocCalendApl = 1 "
sSql = sSql & " ) as T ) "


vPage = vPage + 1
sCadena = sCadena & Chr(12)

Set rs = oConec.CargaRecordSet(sSql)
If rs.BOF Then
Else
    nBandera = 2
    sCadena = sCadena & Chr(10) & Chr(10)
    sCadena = sCadena & "** Existen cuentas en las que no coincide el saldo de producto y el de calendario creditos NORMALES" & Chr(10)
    sCadena = sCadena & "   ______________________________________________________________________________ " & Chr(10) & Chr(10)
    sCadena = sCadena & "   ** Cuenta             S. Producto   S. Calendario ** " & Chr(10)
    sCadena = sCadena & "   ** ============================================== ** " & Chr(10) & Chr(10)
    
    Do While Not rs.EOF
        sCadena = sCadena & "   *  " & rs!cCtaCod & " " & ImpreFormat(rs!nSaldo, 10, 2) & " " & ImpreFormat(rs!nSaldoCal, 10, 2) & Chr(10)
        rs.MoveNext
    Loop
End If
rs.Close
Set rs = Nothing


'Valida que el saldo de calendario sea igual que el saldo de producto creditos RFA
sSql = "Select P.cCtaCod, P.nSaldo, "
sSql = sSql & " nSaldoCal = ( "
sSql = sSql & " Select SUM(nMonto) "
sSql = sSql & " From ( "
sSql = sSql & " Select SUM(nMonto-nMontoPagado) as nMonto "
sSql = sSql & " From ColocCalendDet CD "
sSql = sSql & " Inner Join ColocCalendario Cal ON Cal.cCtaCod = CD.cCtaCod AND Cal.nNroCalen = CD.nNroCalen AND Cal.nColocCalendEstado = 0"
sSql = sSql & " AND Cal.nColocCalendApl = CD.nColocCalendApl AND Cal.nCuota = CD.nCuota AND CD.nPrdConceptoCod = 1000 "
sSql = sSql & " Where CD.nNroCalen = CC.nNroCalen And CD.cCtaCod = CC.cCtaCod "
sSql = sSql & " AND CD.nColocCalendApl = 1 "
sSql = sSql & " Union All "
sSql = sSql & " Select ISNULL(SUM(nMonto-nMontoPagado),0) as nMonto "
sSql = sSql & " From ColocCalendDet CD "
sSql = sSql & " Inner Join ColocCalendario Cal ON Cal.cCtaCod = CD.cCtaCod AND Cal.nNroCalen = CD.nNroCalen AND Cal.nColocCalendEstado = 0"
sSql = sSql & " AND Cal.nColocCalendApl = CD.nColocCalendApl AND Cal.nCuota = CD.nCuota AND CD.nPrdConceptoCod = 1000 "
sSql = sSql & " Where CD.nNroCalen = CC.nNroCalPar And CD.cCtaCod = CC.cCtaCod "
sSql = sSql & " AND CD.nColocCalendApl = 1 "
sSql = sSql & " ) as T ), "
sSql = sSql & " TIPO=Case CC.CRFA"
sSql = sSql & " When 'RFA' Then 'RFA'"
sSql = sSql & " When 'RFC' Then 'RFC'"
sSql = sSql & " When 'DIF' Then 'DIF'"
sSql = sSql & " When 'NOR' Then 'NOR'"
sSql = sSql & " End"
sSql = sSql & " from Producto P "
sSql = sSql & " Inner Join ColocacCred CC ON CC.cCtaCod = P.cCtaCod  "
sSql = sSql & " Where P.nPrdEstado in (2020,2021,2022,2030,2031,2032) AND CC.CRFA in ('RFA','RFC','DIF') "
sSql = sSql & " AND "
sSql = sSql & " P.nSaldo <> "
sSql = sSql & " ( "
sSql = sSql & " Select SUM(nMonto) "
sSql = sSql & " From ( "
sSql = sSql & " Select SUM(nMonto-nMontoPagado) as nMonto "
sSql = sSql & " From ColocCalendDet CD "
sSql = sSql & " Inner Join ColocCalendario Cal ON Cal.cCtaCod = CD.cCtaCod AND Cal.nNroCalen = CD.nNroCalen AND Cal.nColocCalendEstado = 0"
sSql = sSql & " AND Cal.nColocCalendApl = CD.nColocCalendApl AND Cal.nCuota = CD.nCuota AND CD.nPrdConceptoCod = 1000 "
sSql = sSql & " Where CD.nNroCalen = CC.nNroCalen And CD.cCtaCod = CC.cCtaCod "
sSql = sSql & " AND CD.nColocCalendApl = 1 "
sSql = sSql & " Union All "
sSql = sSql & " Select ISNULL(SUM(nMonto-nMontoPagado),0) as nMonto "
sSql = sSql & " From ColocCalendDet CD "
sSql = sSql & " Inner Join ColocCalendario Cal ON Cal.cCtaCod = CD.cCtaCod AND Cal.nNroCalen = CD.nNroCalen AND Cal.nColocCalendEstado = 0 "
sSql = sSql & " AND Cal.nColocCalendApl = CD.nColocCalendApl AND Cal.nCuota = CD.nCuota AND CD.nPrdConceptoCod = 1000 "
sSql = sSql & " Where CD.nNroCalen = CC.nNroCalPar And CD.cCtaCod = CC.cCtaCod "
sSql = sSql & " AND CD.nColocCalendApl = 1 "
sSql = sSql & " ) as T ) "

vPage = vPage + 1
sCadena = sCadena & Chr(12)

Set rs = oConec.CargaRecordSet(sSql)
If rs.BOF Then
Else
    nBandera = 2
    sCadena = sCadena & Chr(10) & Chr(10)
    sCadena = sCadena & "** Existen cuentas en las que no coincide el saldo de producto y el de calendario creditos RFA" & Chr(10)
    sCadena = sCadena & "   ______________________________________________________________________________ " & Chr(10) & Chr(10)
    sCadena = sCadena & "   ** Cuenta             S. Producto   S. Calendario        TIPO ** " & Chr(10)
    sCadena = sCadena & "   ** ========================================================== ** " & Chr(10) & Chr(10)
    
    Do While Not rs.EOF
        sCadena = sCadena & "   *  " & rs!cCtaCod & " " & ImpreFormat(rs!nSaldo, 10, 2) & " " & ImpreFormat(rs!nSaldoCal, 10, 2) & ImpreFormat(rs!Tipo, 5) & Chr(10)
        rs.MoveNext
    Loop
End If
rs.Close
Set rs = Nothing



' Valida la Caja, Asientos VS Planilla Consolidada Colocaciones

sSql = "Select A.*, B.* From "
sSql = sSql & " (Select MC.nMovNro, nDebe = ISNULL(SUM(CASE WHEN O.cIngEgr = 'I' THEN Abs(MC.nMonto) END),0), "
sSql = sSql & " nHaber = ISNULL(SUM(CASE WHEN O.cIngEgr = 'E' THEN Abs(MC.nMonto) END),0) "
sSql = sSql & " From " & lsTablaDiaria & " M JOIN MovCol MC ON M.nMovNro = MC.nMovNro JOIN "
sSql = sSql & " (Select O.cOpeCod, G.cIngEgr From OpeTpo O JOIN GruposOpe GO JOIN GrupoOpe G ON "
sSql = sSql & " GO.cGrupoCod = G.cGrupoCod ON O.cOpeCod = GO.cOpeCod Where G.nEfectivo = 1 Group by O.cOpeCod, G.cIngEgr) O "
sSql = sSql & " ON MC.cOpeCod = O.cOpeCod "
sSql = sSql & " Where M.cMovNro LIKE '" & Format(sFecha, "YYYYMMdd") & "%' And M.nMovFlag = 0 And MC.cCtaCod LIKE '111_____[12]%' "
sSql = sSql & " Group by MC.nMovNro) A "
sSql = sSql & " Join "
sSql = sSql & " (Select MC.nMovNro as nMovNro1, "
sSql = sSql & " nDebe1 = ISNULL(SUM(CASE WHEN C.cOpeCtaDH = 'D' THEN ABS(MCD.nMonto) END),0), "
sSql = sSql & " nHaber1 = ISNULL(SUM(CASE WHEN C.cOpeCtaDH = 'H' THEN ABS(MCD.nMonto) END),0) "
sSql = sSql & " From " & lsTablaDiaria & " M JOIN MovCol MC JOIN MovColDet MCD "
sSql = sSql & " JOIN (Select cOpeCod, nConcepto, cOpeCtaDH From OpeCtaNeg Where cCtaContCod LIKE '11_102%' Group by "
sSql = sSql & " cOpeCod, nConcepto, cOpeCtaDH) C ON MCD.cOpeCod = C.cOpeCod And MCD.nPrdConceptoCod = C.nConcepto "
sSql = sSql & " ON MC.nMovNro = MCD.nMovNro And MC.cCtaCod = MCD.cCtaCod And MC.cOpeCod = MCD.cOpeCod ON M.nMovNro = MC.nMovNro "
sSql = sSql & " Where M.cMovNro LIKE '" & Format(sFecha, "YYYYMMdd") & "%' And M.nMovFlag = 0 And MC.cCtaCod LIKE '111_____[12]%' "
sSql = sSql & " Group by MC.nMovNro) B "
sSql = sSql & " ON A.nMovNro = B.nMovNro1 "
sSql = sSql & " Where (A.nDebe - A.nHaber <> B.nDebe1 - B.nHaber1) "

Set rs = oConec.CargaRecordSet(sSql)
If rs.BOF Then
Else
    nBandera = 2
    sCadena = sCadena & Chr(10) & Chr(10)
    sCadena = sCadena & "** Existen descuadres entre la Caja, Asientos y Planilla Consolidada Colocaciones " & Chr(10)
    sCadena = sCadena & "   ______________________________________________________________________________ " & Chr(10) & Chr(10)
    sCadena = sCadena & "   ** nMovNro    Debe         nMovNro    Debe          ** " & Chr(10)
    sCadena = sCadena & "   ** ================================================ ** " & Chr(10) & Chr(10)
    
    Do While Not rs.EOF
        sCadena = sCadena & "   *  " & Left(rs!nMovNro & "          ", 10) & " " & ImpreFormat(rs!nDebe, 10, 2) & " " & ImpreFormat(rs!nhaber, 10, 2)
        sCadena = sCadena & " " & Left(rs!nMovNro1 & "          ", 10) & " " & ImpreFormat(rs!ndebe1, 10, 2) & " " & ImpreFormat(rs!nhaber1, 10, 2) & Chr(10)
        rs.MoveNext
    Loop
End If
rs.Close
Set rs = Nothing

'Validaciones de CMAC - Llamada
'==============================
vPage = vPage + 1
sCadena = sCadena & Chr(12)
sCadena = sCadena & sCabecera
sCadena = sCadena & ImpreFormat("  ( PRE CUADRE   DE  CMAC/LLAMADAS )  ", 44, 43) & Chr(10) & Chr(10)

sSql = " Select substring(m1.cMovNro, 18,2) as cAgencia, right(m1.cMovNro, 4) as cUser, "
sSql = sSql & " m1.cOpeCod, m1.cOpeDesc,"
sSql = sSql & " m1.nMovNro as nMovCmac, m1.cMovNro as cMovCmac,  isnull(m2.nMovNroRef, 0) as nMovOpeVarias, "
sSql = sSql & " isnull(M2.cMovNro,'') as cMovOpeVarias,  isnull(m3.nMovNroRef, 0) as nMovRegu, isnull(M3.cMovNro,'') as cMovRegu "
sSql = sSql & " From "
sSql = sSql & " (  select m.cOpeCod, cOpeDesc=case  when m.copecod='260501' then 'LLAMADA DEPOSITO EFECTIVO CMAC' "
sSql = sSql & " when m.copecod='260503' then 'LLAMADA RETIRO EFECTIVO CMAC'  when m.copecod='260504' "
sSql = sSql & " then 'LLAMADA RETIRO ORDEN DE PAGO'  when m.copecod='107001' then 'LLAMADA PAGO DE CREDITO' "
sSql = sSql & " end, m.nMovNro , m.cMovNro from " & lsTablaDiaria & " m inner join movcmac mc on m.nmovnro=mc.nmovnro "
sSql = sSql & " inner join persona p on mc.cperscod=p.cperscod  where substring(m.cmovnro,1,8)='" & Format(sFecha, "YYYYMMdd") & "' and "
sSql = sSql & " m.copecod in(260501, 260503, 260504, 107001) and m.nmovflag=0 ) m1 "
sSql = sSql & " Left Join "
sSql = sSql & " (  Select MR.nMovNro, M1.cMovNro, MR.nMovNroRef "
sSql = sSql & " from MovRef MR Inner Join MovOpevarias MO  on MR.nMovNroRef=MO.nMovNro "
sSql = sSql & " Inner Join Mov M1 On M1.nMovNro=MR.nMovNroRef  ) m2 "
sSql = sSql & " on m1.nMovNro = m2.nMovNro "
sSql = sSql & " Left Join  (  Select MR.nMovNro, M1.cMovNro, MR.nMovNroRef, MC.nMonto as nImporteRegu, "
sSql = sSql & " substring(MC.cCtaCod, 9,1) as cMonedaRegu  from MovRef MR Inner Join MovCap MC "
sSql = sSql & " on MR.nMovNroRef=MC.nMovNro Inner Join Mov M1 On M1.nMovNro=MR.nMovNroRef  ) m3 "
sSql = sSql & " on m1.nMovNro = m3.nMovNro "
sSql = sSql & " Where m2.nMovNroRef Is Null Or m2.nMovNroRef = 0 Or m3.nMovNroRef Is Null Or m3.nMovNroRef = 0 "
sSql = sSql & " Order By substring(m1.cMovNro, 18,2),  right(m1.cMovNro, 4), "
sSql = sSql & " m1.cOpeCod , m1.cMovNro "

Set rs = oConec.CargaRecordSet(sSql)
If rs.BOF Then
Else
    nBandera = 2
    sCadena = sCadena & "** Existen llamadas a CMACS sin regularización y/o comision " & Chr(10)
    sCadena = sCadena & "   ________________________________________________________ " & Chr(10) & Chr(10)
    sCadena = sCadena & "   ** Ag User OpeCod Operacion                 nMovCmac  nMovOpeVar nMovRegu  **" & Chr(10)
    sCadena = sCadena & "   ** ======================================================================  **" & Chr(10) & Chr(10)
    Do While Not rs.EOF
        sCadena = sCadena & "   *  " & rs!cAgencia & " " & rs!Cuser & " " & rs!cOpeCod & " " & Left(rs!cOpeDesc & "                    ", 25) & " "
                  sCadena = sCadena & Left(rs!nMovCmac & "          ", 10) & Left(rs!nMovOpeVarias & "          ", 10) & " "
                  sCadena = sCadena & Left(rs!nMovRegu & "          ", 10) & Chr(10)
        rs.MoveNext
        
    Loop
End If
rs.Close
Set rs = Nothing

''''''''''''''''''''''''''''''

'Validaciones de Contabilidad
'=============================
vPage = vPage + 1
sCadena = sCadena & Chr(12)
sCadena = sCadena & sCabecera
sCadena = sCadena & ImpreFormat("  ( PRE CUADRE   DE  CONTABILIDAD  )  ", 44, 43) & Chr(10) & Chr(10)

sSql = "Select cctacontcod, cctacontdesc, cultimaactualizacion "
sSql = sSql & " From ctacont "
sSql = sSql & " Where Len(cctacontcod) > 3 "
sSql = sSql & " and substring(cctacontcod,1, len(cctacontcod)-2)  not in (select cctacontcod from ctacont) "

Set rs = oConec.CargaRecordSet(sSql)
If rs.BOF Then
Else
    nBandera = 2
    sCadena = sCadena & "** Existen Cuentas sin Cabecera " & Chr(10)
    sCadena = sCadena & "   ____________________________ " & Chr(10) & Chr(10)
    sCadena = sCadena & "   ** cCtaContCod         Descripcion                    **" & Chr(10)
    sCadena = sCadena & "   ** ================================================== **" & Chr(10) & Chr(10)
    Do While Not rs.EOF
        sCadena = sCadena & "   *  " & Left(rs!cCtaContCod & "                    ", 20) & " "
                  sCadena = sCadena & Left(rs!cCtaContDesc & "          ", 30) & Chr(10)
        rs.MoveNext
    Loop
End If
rs.Close
Set rs = Nothing

'Validaciones de Convenios
'=============================
vPage = vPage + 1
sCadena = sCadena & Chr(12)
sCadena = sCadena & sCabecera
sCadena = sCadena & ImpreFormat("  ( CREDITOS CON CONVENIO DSCTO. POR PLANILLA )  ", 44, 43) & Chr(10) & Chr(10)

'Busca Creditos de convenio - sin convenio
sSql = " Select cCtaCod From Producto where nPrdEstado in (2020,2021,2022,2030,2031,2032,2201) " _
     & " And Substring(cctacod,6,3) = '301' and cCtaCod not in (Select cCtaCod from ColocacConvenio) "
Set rs = oConec.CargaRecordSet(sSql)
If rs.BOF Then
Else
    nBandera = 2
    sCadena = sCadena & "** Busca Creditos de convenio - sin convenio  " & Chr(10)
    sCadena = sCadena & "   ____________________________ " & Chr(10) & Chr(10)
    sCadena = sCadena & "   ** cCtaCod             **" & Chr(10)
    sCadena = sCadena & "   ** ====================**" & Chr(10) & Chr(10)
    Do While Not rs.EOF
        sCadena = sCadena & "   *  " & rs!cCtaCod & Chr(10)
        rs.MoveNext
    Loop
End If
rs.Close
Set rs = Nothing

'Validaciones de Apoderado de Creditos
'=====================================
vPage = vPage + 1
sCadena = sCadena & Chr(12)
sCadena = sCadena & sCabecera
sCadena = sCadena & ImpreFormat("  ( DOBLE RELACION DE APODERADO )  ", 44, 43) & Chr(10) & Chr(10)

'Busca Doble Relacion de Apoderado
sSql = " Select cCtaCod, Count(*) From ProductoPersona where nPrdPersRelac = 29 " _
     & " Group By cCtaCod Having Count(*) > 1 "
Set rs = oConec.CargaRecordSet(sSql)
If rs.BOF Then
Else
    nBandera = 2
    sCadena = sCadena & "** Doble Relacion de Apoderado " & Chr(10)
    sCadena = sCadena & "   ____________________________ " & Chr(10) & Chr(10)
    sCadena = sCadena & "   ** cCtaCod             **" & Chr(10)
    sCadena = sCadena & "   ** ====================**" & Chr(10) & Chr(10)
    Do While Not rs.EOF
        sCadena = sCadena & "   *  " & rs!cCtaCod & Chr(10)
        rs.MoveNext
    Loop
End If
rs.Close
Set rs = Nothing

'Validaciones de Analista de Creditos
'====================================
vPage = vPage + 1
sCadena = sCadena & Chr(12)
sCadena = sCadena & sCabecera
sCadena = sCadena & ImpreFormat("  ( DOBLE RELACION DE ANALISTA )  ", 44, 43) & Chr(10) & Chr(10)

'Busca Doble Relacion de Analista
sSql = " Select cCtaCod, Count(*) From ProductoPersona where nPrdPersRelac = 28 " _
     & " Group By cCtaCod Having Count(*) > 1 "
Set rs = oConec.CargaRecordSet(sSql)
If rs.BOF Then
Else
    nBandera = 2
    sCadena = sCadena & "** Doble Relacion de Analista " & Chr(10)
    sCadena = sCadena & "   ____________________________ " & Chr(10) & Chr(10)
    sCadena = sCadena & "   ** cCtaCod             **" & Chr(10)
    sCadena = sCadena & "   ** ====================**" & Chr(10) & Chr(10)
    Do While Not rs.EOF
        sCadena = sCadena & "   *  " & rs!cCtaCod & Chr(10)
        rs.MoveNext
    Loop
End If
rs.Close
Set rs = Nothing


If nBandera = 1 Then
    sCadena = ""
End If

If CDate(sFecha) <> gdFecSis Then
    'OpeDiaEliminaTemporal gsCodUser
End If

Devuelve_Errores_PreCuadre = sCadena

End Function
