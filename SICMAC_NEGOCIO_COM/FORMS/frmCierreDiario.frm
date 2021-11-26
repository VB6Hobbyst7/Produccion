VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCierreDiario 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cierre de Dia"
   ClientHeight    =   5325
   ClientLeft      =   3495
   ClientTop       =   2505
   ClientWidth     =   4905
   Icon            =   "frmCierreDiario.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   4905
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
      Top             =   4800
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
      Height          =   2535
      Left            =   120
      ScaleHeight     =   2505
      ScaleWidth      =   4620
      TabIndex        =   8
      Top             =   1920
      Width           =   4650
      Begin VB.CheckBox chkCreditosDirferencias 
         Caption         =   "Chequear diferencias de  saldos"
         Enabled         =   0   'False
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
         Height          =   375
         Left            =   495
         TabIndex        =   16
         Top             =   320
         Value           =   1  'Checked
         Width           =   3840
      End
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
         Left            =   495
         TabIndex        =   15
         Top             =   690
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
         Left            =   495
         TabIndex        =   14
         Top             =   1020
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
         Left            =   495
         TabIndex        =   3
         Top             =   1695
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
         Left            =   495
         TabIndex        =   4
         Top             =   2040
         Width           =   2760
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
         Left            =   495
         TabIndex        =   2
         Top             =   1365
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
      Top             =   4800
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
      Left            =   1680
      TabIndex        =   6
      Top             =   4800
      Width           =   1425
   End
End
Attribute VB_Name = "frmCierreDiario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oConec As COMConecta.DCOMConecta
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
Dim ssql As String 'DAOR 20080606, Variable definido para superar errores del usuario ALPA
Private fMatCredPagoAuto() As CredPagoAuto 'WIOR 20140620
Dim N1 As Integer ' PEAC 20141017, variable para realizar el Log del proceso de cierre
Dim nTpoCierre As Integer ' PEAC 20150101


Function ObtenerPassword() As String
Dim oConec As COMConecta.DCOMConecta
Dim strCadenaConexion As String
Dim intPosI As Integer
Dim intPosF As Integer
Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    strCadenaConexion = oConec.CadenaConexion
    oConec.CierraConexion
    intPosI = InStr(1, strCadenaConexion, "Password")
    intPosF = InStr(intPosI, strCadenaConexion, ";")
    ObtenerPassword = Mid(strCadenaConexion, intPosI + Len("Password="), intPosF - (intPosI + Len("Password=")))
    Set oConec = Nothing
End Function

Function ObtenerUsuarioDatos() As String
Dim oConec As COMConecta.DCOMConecta
Dim strCadenaConexion As String
Dim intpos As Integer
Dim intPosF As Integer
Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    strCadenaConexion = oConec.CadenaConexion
    oConec.CierraConexion
    intpos = InStr(1, strCadenaConexion, "User ID=")
    intPosF = InStr(intpos, strCadenaConexion, ";")
    ObtenerUsuarioDatos = Mid(strCadenaConexion, intpos + Len("User ID="), intPosF - (intpos + Len("User ID=")))
Set oConec = Nothing
End Function

Function ObtenerServidorDatos() As String
Dim oConec As COMConecta.DCOMConecta
Dim strCadenaConexion As String
Dim intpos As Integer
Dim intPosF As Integer
Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    strCadenaConexion = oConec.CadenaConexion
    oConec.CierraConexion
    intpos = InStr(1, strCadenaConexion, "Data Source=")
    intPosF = InStr(intpos, strCadenaConexion, ";")
    ObtenerServidorDatos = Mid(strCadenaConexion, intpos + Len("Data Source="), intPosF - (intpos + Len("Data Source=")))
    
Set oConec = Nothing
End Function

Function ObtenerBaseDatos() As String
Dim oConec As COMConecta.DCOMConecta
Dim strCadenaConexion As String
Dim intpos As Integer
Dim intPosF As Integer
Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    strCadenaConexion = oConec.CadenaConexion
    oConec.CierraConexion
    intpos = InStr(1, strCadenaConexion, "Initial Catalog=")
    intPosF = InStr(intpos, strCadenaConexion, ";")
    ObtenerBaseDatos = Mid(strCadenaConexion, intpos + Len("Initial Catalog="), intPosF - (intpos + Len("Initial Catalog=")))
    
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
Dim nNumRet As Long, nNumDep As Long, nnumitf As Long, nITF As Double
Dim nMonDep As Double, nMonIntCap As Double, nSaldo As Double, nMonChq As Double
Dim nNumCom As Double, nMontoCom As Double
Dim sFecha As String, sCondicion As String
Dim oCap As COMDCaptaGenerales.DCOMCaptaGenerales
Dim rsEstad As ADODB.Recordset
Dim rsComi As ADODB.Recordset
Dim ssql As String

sFecha = Format$(dFecha, "yyyymmdd") & "%"
sCondicion = "___" & sAge & Trim(nProd) & Trim(nMon) & "%"

nNumAper = 0: nNumCancAct = 0: nNumCancInact = 0: nMonAper = 0
nMonCancAct = 0: nMonCancInact = 0: nMonRetInact = 0: nMonRetInt = 0
nMonRet = 0: nNumRet = 0: nNumDep = 0: nMonDep = 0: nMonIntCap = 0
nSaldo = 0: nMonChq = 0: nNumCom = 0: nMontoCom = 0

ssql = " Select ISNULL(Count(*),0)nNumCom , ISNULL(SUM(CD.nMonto),0)nMontoCom From " & _
       " CapMovTipo EO INNER JOIN Mov M INNER JOIN MovCap C INNER JOIN MovCapDet CD ON" & _
       " C.nMovNro = CD.nMovNro And C.cCtaCod = Cd.cCtaCod And CD.cOpeCod = '200358'  And C.cOpeCod = '200358'  ON M.nMovNro = C.nMovNro ON" & _
       " EO.cOpeCod = C.cOpeCod Where CD.nConceptoCod  = 11 And EO.nCapMovTpo = 6 And" & _
       " M.cMovNro LIKE '" & sFecha & "' And C.cCtaCod LIKE '" & sCondicion & "' And M.nMovFlag = 0  AND NOT C.COPECOD LIKE '99%'"
Set rsComi = New ADODB.Recordset
Set rsComi = oConec.CargaRecordSet(ssql)
If Not (rsComi.EOF And rsComi.BOF) Then
    nNumCom = rsComi!nNumCom
    nMontoCom = rsComi!nMontoCom
End If

Set oCap = New COMDCaptaGenerales.DCOMCaptaGenerales
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
    '----observacion
    nMonRet = rsEstad("nMonRet")
    nNumRet = rsEstad("nNumRet")
    '----
    nNumDep = rsEstad("nNumDep")
    nMonDep = rsEstad("nMonDep")
    nMonIntCap = rsEstad("nMonIntCap")
    nnumitf = rsEstad("nnumitf")
    nITF = rsEstad("nitf")

End If
' Sumatoria de comision al MontoRet / NumRet --- AVMM --- 10-11-2006
  If nMontoCom > 0 And nNumRet > 0 Then
    nMonRet = nMonRet + (nMontoCom * -1)
    nNumRet = nNumRet + nNumCom
  End If
' ------------------------------------------------------------------

'Calculamos El Saldo por Producto
Set rsEstad = oCap.GetCapEstadSaldoProd(sCondicion)
If Not (rsEstad.EOF And rsEstad.BOF) Then
    nSaldo = rsEstad("nSaldo")
End If
'Calculamos el monto en cheques movidos en la fecha indicada
Set rsEstad = oCap.GetCapEstadMovCheques(sFecha, sCondicion)
If Not (rsEstad.EOF And rsEstad.BOF) Then
    nMonChq = rsEstad("nMonChq")
End If
rsEstad.Close
Set rsEstad = Nothing
Set oCap = Nothing

'Ingresamos el Registro en la tabla de Estadistica
ssql = "Insert CapEstadMovimiento (cCodAge,dEstad,nProducto,nMoneda,nNumAper,nMonAper,nNumCanc,nMonCanc, " _
    & "nRetInt,nRetInac,nNumCancInac,nMonCancInac,nNumDep,nMonDep,nNumRet,nMonRet, nIntCap,nSaldo,nMonChq,nnumitf,nitf,cCodUsu) " _
    & "Values ('" & sAge & "','" & Format$(dFecha, "mm/dd/yyyy hh:mm:ss") & "'," & Trim(nProd) & "," & Trim(nMon) & ", " _
    & Trim(nNumAper) & ", " & Trim(nMonAper) & "," & Trim(nNumCancAct) & "," & Trim(nMonCancAct) & ", " _
    & Trim(nMonRetInt) & "," & Trim(nMonRetInact) & "," & Trim(nNumCancInact) & "," & Trim(nMonCancInact) & ", " _
    & Trim(nNumDep) & "," & Trim(nMonDep) & "," & Trim(nNumRet) & "," & Trim(nMonRet) & "," & Trim(nMonIntCap) & ", " _
    & Trim(nSaldo) & "," & Trim(nMonChq) & "," & Trim(nnumitf) & "," & Trim(nITF) & ",'" & sUsu & "')"
    oConec.ConexionActiva.Execute ssql

End Sub

Private Sub CapGeneraEstadistica()
Dim rsAgencia As Recordset
Dim sAgencia As String, ssql As String, sFecha As String
Dim clsGen As COMDConstSistema.DCOMGeneral
''''Dim cmd As adodb.Command
''''Dim prm As adodb.Parameter
Dim bTrans As Boolean
Dim dFecha As Date
On Error GoTo ErrStore
Dim SqlSent As String


bTrans = False
Set clsGen = New COMDConstSistema.DCOMGeneral
Set rsAgencia = clsGen.getAgencias()
Set clsGen = Nothing
sFecha = Format$(gdFecSis & " " & GetHoraServer(), "dd/mm/yyyy hh:mm:ss")
dFecha = CDate(sFecha)
bTrans = True
'ALPA 20080723****************************************************
oConec.CommadTimeOut = 0
'*****************************************************************
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

''''Set cmd = Nothing
''''Set cmd = New adodb.Command
''''cmd.CommandText = "CapGeneraEstadSaldo"
''''cmd.CommandType = adCmdStoredProc
''''cmd.Name = "CapGeneraEstadSaldo"
''''Set prm = cmd.CreateParameter("dEstad", adDate, adParamInput)
''''cmd.Parameters.Append prm
''''Set prm = cmd.CreateParameter("sCodCMACT", adChar, adParamInput, 4)
''''cmd.Parameters.Append prm
''''Set prm = cmd.CreateParameter("sUsuario", adChar, adParamInput, 4)
''''cmd.Parameters.Append prm
''''Set cmd.ActiveConnection = oConec.ConexionActiva
''''cmd.CommandTimeout = 720
''''cmd.Parameters.Refresh
ssql = ""
ssql = "exec CapGeneraEstadSaldo '" & Format(sFecha, "YYYY/MM/DD") & "' , '" & gsCodPersCMACT & "', '" & gsCodUser & "'"
oConec.ConexionActiva.Execute ssql
'''' oConec.ConexionActiva.CapGeneraEstadSaldo sFecha, gsCodPersCMACT, gsCodUser

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


''''Set cmd = Nothing
''''Set cmd = New adodb.Command
''''cmd.CommandText = "CapGeneraSaldosDiarios"
''''cmd.CommandType = adCmdStoredProc
''''cmd.Name = "CapGeneraSaldosDiarios"
''''Set prm = cmd.CreateParameter("dEstad", adDate, adParamInput)
''''cmd.Parameters.Append prm
''''Set prm = cmd.CreateParameter("sUsuario", adChar, adParamInput, 4)
''''cmd.Parameters.Append prm
''''Set cmd.ActiveConnection = oConec.ConexionActiva
''''cmd.CommandTimeout = 720
''''cmd.Parameters.Refresh
''''oConec.ConexionActiva.CapGeneraSaldosDiarios sFecha, gsCodUser
''''
''''Set prm = Nothing
''''Set cmd = Nothing

ssql = ""
ssql = "exec CapGeneraSaldosDiarios '" & Format(sFecha, "YYYY/MM/DD") & "' , '" & gsCodUser & "'"
oConec.ConexionActiva.Execute ssql
bTrans = False
Exit Sub
ErrStore:
    If bTrans Then oConec.ConexionActiva.RollbackTrans
    Set oConec = Nothing
    MsgBox err.Description, vbExclamation, "Error"
End Sub

Private Sub CargaDatos()
Dim ssql As String
Dim R As New ADODB.Recordset

    ' se incluyo Funcion Descuento de Inactivas
    bCierreMes = VerificaDiaHabil(gdFecSis, 3)
    If bCierreMes Then
        CheckConsol.Visible = True
    Else
        CheckConsol.Visible = False
    End If
    
    ssql = "Select nConsSisCod, nConsSisValor From ConstSistema Where nConsSisCod IN (7,11,12,170)"
    Set R = oConec.CargaRecordSet(ssql)
    
    Do While Not R.EOF
        If CDate(R("nConsSisValor")) = gdFecSis Then
            Select Case R("nConsSisCod")
                Case 7
                    ChkCierreAho.value = 1
                    ChkCierreAho.Enabled = False
                Case 11
                    ChkCierreCred.value = 1
                    ChkCierreCred.Enabled = False
                Case 12
                    ChkCierrePig.value = 1
                    ChkCierrePig.Enabled = False
                Case 170
                    CheckConsol.value = 1
                    CheckConsol.Enabled = False
            End Select
        End If
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
End Sub
'MIOL 20121009, SEGUN RQ12272 **************************************************
Private Sub ClientesBloqueadosOrdePago(ByVal dFecSist As String, ByVal cnromov As String)
    oConec.CommadTimeOut = 0
    ssql = ""
    ssql = "exec stp_upd_ActualizarClienteBloqueadoOrdenPAgo '" & dFecSist & "','" & cnromov & "'"
    oConec.ConexionActiva.Execute ssql
End Sub
'END MIOL **********************************************************************
Private Sub CierreAhorros()
''''Dim cmd As adodb.Command
''''Dim prm As adodb.Parameter
''''Dim sSQL As String
''''
''''    oConec.ConexionActiva.BeginTrans
''''    Set cmd = New adodb.Command
''''    Set prm = New adodb.Parameter
''''    cmd.CommandText = "CaptacCierreDiario"
''''    cmd.CommandType = adCmdStoredProc
''''    cmd.Name = "CaptacCierreDiario"
''''    Set prm = cmd.CreateParameter("FechaTran", adDate, adParamInput)
''''    cmd.Parameters.Append prm
''''    Set prm = cmd.CreateParameter("Usuario", adChar, adParamInput, 4)
''''    cmd.Parameters.Append prm
''''    Set prm = cmd.CreateParameter("Agencia", adChar, adParamInput, 2)
''''    cmd.Parameters.Append prm
''''    Set cmd.ActiveConnection = oConec.ConexionActiva
''''    cmd.CommandTimeout = 7200
''''    cmd.Parameters.Refresh
''''    oConec.ConexionActiva.CaptacCierreDiario Format(gdFecSis & " " & GetHoraServer(), "yyyy-mm-dd hh:mm:ss"), gsCodUser, gsCodAge
''''    sSQL = "UPDATE ConstSistema SET nConsSisValor = '" & Format(gdFecSis, "dd/mm/yyyy") & "' WHERE nConsSisCod = 7"
''''    oConec.ConexionActiva.Execute sSQL
''''    oConec.ConexionActiva.CommitTrans
''''    oConec.ConexionActiva.BeginTrans
''''    If Not bCierreMes Then
''''       Call CapGeneraEstadistica
''''    End If
''''    oConec.ConexionActiva.CommitTrans
''''    Set cmd = Nothing
''''    Set prm = Nothing
'**ALPA********************26/05/2008**********************************************************************************************************
    Dim ssql As String
    oConec.ConexionActiva.BeginTrans
    oConec.CommadTimeOut = 0
    ssql = ""
    ssql = "exec CaptacCierreDiario '" & Format(gdFecSis & " " & GetHoraServer(), "yyyy/mm/dd hh:mm:ss") & "', '" & gsCodUser & "', '" & gsCodAge & "'"
    oConec.ConexionActiva.Execute ssql
    ssql = "UPDATE ConstSistema SET nConsSisValor = '" & Format(gdFecSis, "dd/mm/yyyy") & "' WHERE nConsSisCod = 7"
    oConec.ConexionActiva.Execute ssql
    oConec.ConexionActiva.CommitTrans
        
    oConec.ConexionActiva.BeginTrans
    If Not bCierreMes Then
       Call CapGeneraEstadistica
    End If
    oConec.ConexionActiva.CommitTrans
'***********************************************************************************************************************************************
End Sub

Private Sub SalvaColocacSaldo()
'Dim cmd As New adodb.Command
'Dim prm As New adodb.Parameter
Dim ssql As String
  
    
    'Salva Data A Consolidar
    If bCierreMes Then
            
''        cmd.CommandText = "ColocCred_GeneraColocacSaldo"
''        cmd.CommandType = adCmdStoredProc
''        cmd.Name = "ColocCred_GeneraColocacSaldo"
''        Set prm = cmd.CreateParameter("FechaTran", adDate, adParamInput)
''        cmd.Parameters.Append prm
''        Set cmd.ActiveConnection = oConec.ConexionActiva
''        cmd.CommandTimeout = 72000
''        cmd.Parameters.Refresh
''        oConec.ConexionActiva.ColocCred_GeneraColocacSaldo Format(gdFecSis & " " & GetHoraServer(), "yyyy/mm/dd hh:mm:ss")
''        Set cmd = Nothing
''        Set prm = Nothing
        
        ssql = ""
        ssql = "exec ColocCred_GeneraColocacSaldo '" & Format(gdFecSis & " " & GetHoraServer(), "yyyy/mm/dd hh:mm:ss") & "'"
        oConec.ConexionActiva.Execute ssql
    End If

End Sub

Private Sub GuardaConsolidada()
'Dim cmd As New adodb.Command
'Dim prm As New adodb.Parameter
Dim ssql As String

    'Salva Data A Consolidar
    If bCierreMes Then
'''        oConec.ConexionActiva.CommandTimeout = 72000
'''        cmd.CommandText = "GuardarDataAConsolidarCreditos"
'''        cmd.CommandType = adCmdStoredProc
'''        cmd.Name = "GuardarDataAConsolidarCreditos"
'''        Set prm = cmd.CreateParameter("FechaTran", adDate, adParamInput)
'''        cmd.Parameters.Append prm
'''        Set cmd.ActiveConnection = oConec.ConexionActiva
'''        cmd.CommandTimeout = 72000
'''        cmd.Parameters.Refresh
'''        oConec.ConexionActiva.GuardarDataAConsolidarCreditos Format(gdFecSis & " " & GetHoraServer(), "yyyy/mm/dd hh:mm:ss")
'''        Set cmd = Nothing
'''        Set prm = Nothing
        
        ssql = ""
        ssql = "exec GuardarDataAConsolidarCreditos '" & Format(gdFecSis & " " & GetHoraServer(), "yyyy/mm/dd hh:mm:ss") & "'"
        oConec.ConexionActiva.Execute ssql
    End If

End Sub

Private Sub GuardaConsolidadaDTS()
'CREADO POR CMACICA AUTOR:LMMD
'ARCV 25-06-2007
'    strNameDTS = "DTSCierreCreditosConsolidada"
'
'Set objDTS = New DTS.Package
'
'    objDTS.LoadFromSQLServer strServerName, strUsuarioSQL, strPasswordSQL, DTSSQLStgFlag_Default, _
'                              , , , strNameDTS
'
'    objDTS.GlobalVariables("gFechaHora").value = Format(gdFecSis, "mm/dd/yyyy")
'    objDTS.GlobalVariables("gCodUsers").value = gsCodUser
'    objDTS.GlobalVariables("gCodAgen").value = Right(gsCodAge, 2)
'
'
'    objDTS.Execute
'    bExito = True
'    For Each objSteps In objDTS.Steps
'        objSteps.ExecuteInMainThread = True
'        If objSteps.ExecutionResult = DTSStepExecResult_Failure Then
'            objSteps.GetExecutionErrorInfo nError, sSource, sDesc
'            MsgBox "Error Cierre Créditos:" & objSteps.Description & " " & sDesc, vbExclamation, "Error"
'            bExito = False
'            Exit For
'        End If
'    Next
'    objDTS.UnInitialize
'    Set objDTS = Nothing
''''Dim cmd As adodb.Command
''''Dim prm As adodb.Parameter
''''    Set cmd = New adodb.Command
''''    cmd.CommandText = "GuardarDataAConsolidarCreditos"
''''    cmd.CommandType = adCmdStoredProc
''''    cmd.Name = "GuardarDataAConsolidarCreditos"
''''    Set prm = cmd.CreateParameter("FechaTran", adDate, adParamInput)
''''    cmd.Parameters.Append prm
''''    Set cmd.ActiveConnection = oConec.ConexionActiva
''''    '**DAOR 20080329 ***************************************
''''    'cmd.CommandTimeout = 3600
''''    cmd.CommandTimeout = 0
''''    '*******************************************************
''''    cmd.Parameters.Refresh
''''    oConec.ConexionActiva.GuardarDataAConsolidarCreditos Format(gdFecSis & " " & GetHoraServer(), "dd/mm/yyyy hh:mm:ss")
''''    Set cmd = Nothing
''''    Set prm = Nothing
    'ALPA*****25/05/2008**************************************************************************************************
    oConec.CommadTimeOut = 0
   
    ssql = ""
    ssql = "exec GuardarDataAConsolidarCreditos '" & Format(gdFecSis & " " & GetHoraServer(), "yyyy/mm/dd hh:mm:ss") & "'"
    oConec.ConexionActiva.Execute ssql
    
    
    '*********************************************************************************************************************
'---------
End Sub

Private Sub GuardaColocacSaldoDTS()
Dim ssql As String
''''Dim cmd As adodb.Command
''''Dim prm As adodb.Parameter
''''    Set cmd = New adodb.Command
''''    cmd.CommandText = "ColocCred_GeneraColocacSaldo"
''''    cmd.CommandType = adCmdStoredProc
''''    cmd.Name = "ColocCred_GeneraColocacSaldo"
''''    Set prm = cmd.CreateParameter("FechaTran", adDate, adParamInput)
''''    cmd.Parameters.Append prm
''''    Set cmd.ActiveConnection = oConec.ConexionActiva
''''    cmd.CommandTimeout = 1420
''''    cmd.Parameters.Refresh
''''    oConec.ConexionActiva.ColocCred_GeneraColocacSaldo Format(gdFecSis & " " & GetHoraServer(), "dd/mm/yyyy hh:mm:ss")
''''    Set cmd = Nothing
''''    Set prm = Nothing
    'ALPA***20080723
    
    oConec.CommadTimeOut = 0
    
    ''*** PEAC 20160130 - COMENTADO SEGUN INDICACIONES DE MARCOS PAREADES.
    
'    If bCierreMes Then
'        ssql = ""
'        ssql = "exec stp_pro_MigrarRefinanciadoANormal '" & Format(gdFecSis & " " & GetHoraServer(), "yyyy/mm/dd hh:mm:ss") & "','" & gsCodUser & "'"
'        oConec.ConexionActiva.Execute ssql
'    End If

    ssql = ""
    ssql = "exec ColocCred_GeneraColocacSaldo '" & Format(gdFecSis & " " & GetHoraServer(), "yyyy/mm/dd hh:mm:ss") & "'"
    oConec.ConexionActiva.Execute ssql
    
    'MAVM 20120927 Actualizar CapVencido***
    Dim R As ADODB.Recordset
    Set R = New ADODB.Recordset
    
    '*** PEAC 20141017 - proceso optimizado
    oConec.CommadTimeOut = 0
    ssql = ""
    ssql = "exec stp_sel_CreditosColocacSaldoNuevo '" & Format(gdFecSis & " " & GetHoraServer(), "yyyy/mm/dd hh:mm:ss") & "'"
    oConec.ConexionActiva.Execute ssql
    
'    oConec.CommadTimeOut = 0
'    ssql = "exec stp_sel_CreditosColocacSaldo '" & Format(gdFecSis & " " & GetHoraServer(), "yyyy/mm/dd hh:mm:ss") & "'"
'    Set R = oConec.CargaRecordSet(ssql)
'    If Not (R.BOF Or R.EOF) Then
'        Do While Not R.EOF
'            ssql = ""
'            ssql = "Update CS Set nCapVencido = ISNULL(dbo.CapitalVencido('" & R!cCtaCod & "','" & Format(gdFecSis, "yyyy/mm/dd") & "'),0) from ColocacSaldo CS Where cCtaCod = '" & R!cCtaCod & "' And DateDiff(d,dFecha,'" & Format(gdFecSis, "yyyy/mm/dd") & "')=0"
'            oConec.ConexionActiva.Execute ssql
'            R.MoveNext
'        Loop
'    End If
    '*** FIN PEAC
    
End Sub


Private Sub CierreCreditos()
'Dim cmd As adodb.Command
'Dim prm As adodb.Parameter
Dim ssql As String

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
       
    
'''''    Set cmd = New adodb.Command
'''''    cmd.CommandText = "ColocCredCierreDiario"
'''''    cmd.CommandType = adCmdStoredProc
'''''    cmd.Name = "ColocCredCierreDiario"
'''''    Set prm = cmd.CreateParameter("FechaTran", adDate, adParamInput)
'''''    cmd.Parameters.Append prm
'''''    Set prm = cmd.CreateParameter("Usuario", adChar, adParamInput, 4)
'''''    cmd.Parameters.Append prm
'''''    Set prm = cmd.CreateParameter("Agencia", adChar, adParamInput, 2)
'''''    cmd.Parameters.Append prm
'''''    Set cmd.ActiveConnection = oConec.ConexionActiva
'''''    cmd.CommandTimeout = 7200
'''''    cmd.Parameters.Refresh
'''''    oConec.ConexionActiva.ColocCredCierreDiario Format(gdFecSis & " " & GetHoraServer(), "dd/mm/yyyy hh:mm:ss"), gsCodUser, Right(gsCodAge, 2)
    '******ALPA*********************26/05/2008*************************************************************************************************************
    'oConec.ConexionActiva.BeginTrans
    'ALPA***20080723
    oConec.CommadTimeOut = 0
    ssql = ""
    ssql = "exec ColocCredCierreDiario '" & Format(gdFecSis & " " & GetHoraServer(), "yyyy/mm/dd hh:mm:ss") & "', '" & gsCodUser & "', '" & Right(gsCodAge, 2) & "'"
    oConec.ConexionActiva.Execute ssql
    
    'Actualiza Dias Atraso ColocacSaldo
    'ALPA 20110827
    ssql = ""
    ssql = "exec stp_upd_ActualizaDiasAtrasoCierreDiario '" & Format(gdFecSis & " " & GetHoraServer(), "yyyy/mm/dd hh:mm:ss") & "'"
    oConec.ConexionActiva.Execute ssql
    
    'oConec.ConexionActiva.BeginTrans
    ssql = ""
    ssql = "exec ColocCredCierreEstadisticaDiaria '" & Format(gdFecSis & " " & GetHoraServer(), "yyyy/mm/dd hh:mm:ss") & "'"
    oConec.ConexionActiva.Execute ssql
    
    'oConec.ConexionActiva.CommitTrans
    
    '******END**ALPA****************************************************************************************************************************************
    'EJVG20141211 *** Cartera Vencida de Recuperaciones
    ssql = "Exec stp_upd_ActualizaCapVencidoFecha '" & Format(gdFecSis, "yyyymmdd") & "'"
    oConec.ConexionActiva.Execute ssql
    
    ssql = "Exec GeneraColocRecupCartera '" & Format(DateAdd("D", 1, gdFecSis), "yyyymmdd") & "'"
    oConec.ConexionActiva.Execute ssql
    
    ssql = "Exec ActualizarCapRecuperado"
    oConec.ConexionActiva.Execute ssql
    'END EJVG *******
    
    'Estadistica diaria 12/03/2004
    
'''''    Set cmd = New adodb.Command
'''''    cmd.CommandText = "ColocCredCierreEstadisticaDiaria"
'''''    cmd.CommandType = adCmdStoredProc
'''''    cmd.Name = "ColocCredCierreEstadisticaDiaria"
'''''    Set prm = cmd.CreateParameter("FechaTran", adDate, adParamInput)
'''''    cmd.Parameters.Append prm
'''''    Set cmd.ActiveConnection = oConec.ConexionActiva
'''''    cmd.CommandTimeout = 7200
'''''    cmd.Parameters.Refresh
'''''    oConec.ConexionActiva.ColocCredCierreEstadisticaDiaria Format(gdFecSis & " " & GetHoraServer(), "dd/mm/yyyy hh:mm:ss")
    
    'Fin de estadistica diaria
    
    
    ssql = " UPDATE ConstSistema SET nConsSisValor = '" & Format(gdFecSis, "dd/mm/yyyy") & "' WHERE nConsSisCod = 11"
    oConec.ConexionActiva.Execute ssql
    
    oConec.ConexionActiva.CommitTrans
    
''    Set cmd = Nothing
''    Set prm = Nothing
 
        
    
End Sub

Private Sub CierrePignoraticio()
''''Dim cmd As New adodb.Command
''''Dim prm As New adodb.Parameter
Dim ssql As String
On Error GoTo MMM
    oConec.ConexionActiva.BeginTrans
    
    'If lnFiltro = 1 Then    'Trujillo
'''''        cmd.CommandText = "ColocPigCierreDiario"
    'ElseIf lnFiltro = 2 Then    'Lima
     '   cmd.CommandText = "ColocPignoCierreDiario"
    'End If
'''''    cmd.CommandType = adCmdStoredProc
'''''    cmd.Name = "ColocPigCierreDiario"
'''''    Set prm = cmd.CreateParameter("FechaTran", adDate, adParamInput)
'''''    cmd.Parameters.Append prm
'''''    Set prm = cmd.CreateParameter("Usuario", adChar, adParamInput, 4)
'''''    cmd.Parameters.Append prm
'''''    Set prm = cmd.CreateParameter("Agencia", adChar, adParamInput, 5)
'''''    cmd.Parameters.Append prm
'''''    Set cmd.ActiveConnection = oConec.ConexionActiva
'''''    cmd.CommandTimeout = 720
'''''    cmd.Parameters.Refresh
'''''    oConec.ConexionActiva.ColocPigCierreDiario Format(gdFecSis & " " & GetHoraServer(), "dd/mm/yyyy hh:mm:ss"), gsCodUser, gsCodAge
    'ALPA***20080723
    oConec.CommadTimeOut = 0
    ssql = ""
    ssql = "exec ColocPigCierreDiario '" & Format(gdFecSis & " " & GetHoraServer(), "yyyy/mm/dd hh:mm:ss") & "', '" & gsCodUser & "', '" & gsCodAge & "'"
    oConec.ConexionActiva.Execute ssql
    
    ssql = ""
    ssql = " UPDATE ConstSistema SET nConsSisValor = '" & Format(gdFecSis, "dd/mm/yyyy") & "' WHERE nConsSisCod = 12"
    oConec.ConexionActiva.Execute ssql
    
    oConec.ConexionActiva.CommitTrans
    
''''    Set cmd = Nothing
''''    Set prm = Nothing
Exit Sub
MMM:
MsgBox err.Description
End Sub
'PEAC 20141017, Proceso para generar el log de Cierre diario
Private Sub GeneraLogCierre(ByVal psTexto As String, Optional ByVal pnFecHor As Integer = 1)
    N1 = FreeFile()
    Open "C:\LOGCIERRE.TXT" For Append As #N1
    If pnFecHor <> 1 Then
        Print #N1, psTexto
    Else
        Print #N1, psTexto; Date; Time
    End If
    Close #1
End Sub


Private Sub cmdAceptar_Click()
Dim ssql As String
Dim Cmd As ADODB.Command
Dim Prm As ADODB.Parameter
Dim nDias As Integer, i As Integer
'ALPA 20120531****************************************
Dim loRsCredSal As ADODB.Recordset
Dim lsDiferenciaSaldos As String
Dim oImpre As COMFunciones.FCOMImpresion
Set oImpre = New COMFunciones.FCOMImpresion
Dim oPrevio As previo.clsprevio
'*****************************************************
Dim psMovAct As String
Dim loContFunct As COMNContabilidad.NCOMContFunciones
'ALPA 20100913***************************************************
Dim rsMov As ADODB.Recordset
Set loContFunct = New COMNContabilidad.NCOMContFunciones
psMovAct = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
Set loContFunct = Nothing
'****************************************************************
Set loRsCredSal = New ADODB.Recordset 'ALPA 20120531
If gsProyectoActual = "H" Then
    Dim oCred As COMDCredito.DCOMCredito
    Set oCred = New COMDCredito.DCOMCredito
    If oCred.VerificarIndiceVAC(gdFecSis) = False Then
        MsgBox "Debe ingresar el Factor VAC", vbInformation, "Mensaje"
        Exit Sub
    End If
    Set oCred = Nothing
End If

nTpoCierre = 1 ' 1= Cierre sin mensajes de aviso 0= cierre con avisos en cada proceso

If MsgBox("Esta Ud. seguro de efectuar el cierre?", vbQuestion + vbYesNo, "Advertencia") = vbNo Then
    Exit Sub
End If


If nTpoCierre = 1 Then
    '-----------------------------------------------------
    Call GeneraLogCierre("Comienza el cierre del " & Format(gdFecSis, "dd/mm/yyyy") & " - ")
    '-----------------------------------------------------
End If

'strBaseSQL = ObtenerBaseDatos
'sSql = " master..sp_dboption " & strBaseSQL & ",'trunc. log on chkpt.',true"
'oConec.ConexionActiva.Execute sSql
'2
nDias = DateDiff("d", CDate(txtFecIni), CDate(txtFecFin)) + 1
lsDiferenciaSaldos = oImpre.CabeceraPagina("REPORTE DE DIFERENCIA DE SALDOS", 0, 1, "Oficina Principal", "CMAC MAYNAS SA", gdFecSis, , False)
'ALPA 20120531**********
lsDiferenciaSaldos = BON & lsDiferenciaSaldos & oImpre.ImpreFormat("Credito", 18) & oImpre.ImpreFormat(" ", 12) & oImpre.ImpreFormat("Sald.Calend.", 12, 2) & oImpre.ImpreFormat("Sald.Product.", 12, 2) & oImpre.ImpreFormat(" ", 2) & oImpre.ImpreFormat("Diferencia", 15) & BOFF & Chr$(10)
If chkCreditosDirferencias.value = 1 Then
    oConec.CommadTimeOut = 0
    
    If nTpoCierre = 1 Then
        '----------------------------------------------------
        Call GeneraLogCierre("Comienza Validacion Dif saldos de producto y calendario ")
        '----------------------------------------------------
    End If
    
    ssql = "exec stp_sel_DiferenciasSaldosCreditos"
    Set loRsCredSal = oConec.CargaRecordSet(ssql)
    If Not (loRsCredSal.BOF Or loRsCredSal.EOF) Then
        Set oPrevio = New previo.clsprevio
        Do While Not loRsCredSal.EOF
        '*** PEAC 20120806 - se aumento el ancho de visualizacion de las cuentas encontradas de 15 a 18.
        lsDiferenciaSaldos = lsDiferenciaSaldos & oImpre.ImpreFormat(loRsCredSal!cCtaCod, 18) & oImpre.ImpreFormat(" ", 10) & oImpre.ImpreFormat(loRsCredSal!nSaldoCalendario, 12, 2) & oImpre.ImpreFormat(loRsCredSal!nSaldo, 12, 2) & oImpre.ImpreFormat(loRsCredSal!nSaldoCalendario - loRsCredSal!nSaldo, 12, 2) & Chr$(10)
        loRsCredSal.MoveNext
        
        Loop
        
        If nTpoCierre = 1 Then
            '----------------------------------------------------
            Call GeneraLogCierre("Encuentra Dif saldos de producto y calendario ")
            '----------------------------------------------------
        End If
            
        oPrevio.Show lsDiferenciaSaldos, "Direfencias de saldos de créditos"
        Set oPrevio = Nothing
        Exit Sub
    Else
    
        If nTpoCierre = 1 Then
            '----------------------------------------------------
            Call GeneraLogCierre("Termina Validacion Dif saldos de producto y calendario ")
            '----------------------------------------------------
        Else
            MsgBox "Validación de Saldos de Creditos no presenta diferencias"
        End If
    
    End If
    
End If

For i = 1 To nDias
    Screen.MousePointer = 11
    If i > 1 Then frmInicioDia.InicioDia
    bCierreMes = VerificaDiaHabil(gdFecSis, 3)
    'ALPA 20100913***************************************************
    If bCierreMes Then
        Set rsMov = New ADODB.Recordset
        ssql = "B2_stp_sel_MovVerificaCodOperacion '" & Mid(psMovAct, 1, 8) & "'"
        Set rsMov = oConec.ConexionActiva.Execute(ssql)
        If rsMov.BOF Or rsMov.EOF Then
            oConec.CommadTimeOut = 0
            ssql = "exec B2_stp_upd_Reclasificacion"
            oConec.ConexionActiva.Execute ssql

            ssql = "exec B2_stp_ins_llenarMovReclasificacion  '" & psMovAct & "'"
            oConec.ConexionActiva.Execute ssql
        End If
        rsMov.Close
        Set rsMov = Nothing
    End If
    '****************************************************************
    GetDatosConexionDTS
    
    ''Call PagoAutomaticoCredito(gdFecSis)   'WIOR 20140619
    
    'JUEZ 20150217 *************************************************************
    Call DebitoAutomaticoPagoCreditos(gdFecSis)
    If nTpoCierre = 1 Then
        Call GeneraLogCierre("Termina Proceso Pago de Credito con Debito Automatico ")
    Else
        MsgBox "Proceso Pago de Credito con Debito Automatico Finalizado", vbInformation, "Aviso de Cierre"
    End If
    Call DebitoAutomaticoPagoServicios(gdFecSis)
    If nTpoCierre = 1 Then
        Call GeneraLogCierre("Termina Proceso Pago de Servicios con Debito Automatico ")
    Else
        MsgBox "Proceso Pago de Servicios con Debito Automatico Finalizado", vbInformation, "Aviso de Cierre"
    End If
    'END JUEZ ******************************************************************
    
    If ChkColoacSaldo.value = 1 And ChkColoacSaldo.Enabled = True Then
        'ALPA 20080723*************************************************
        oConec.CommadTimeOut = 0
        '**************************************************************
        Call GuardaColocacSaldoDTS ' Ok ALPA
        ssql = "UPDATE ConstSistema SET nConsSisValor = '" & Format(gdFecSis, "dd/mm/yyyy") & "' WHERE nConsSisCod = 171"
        oConec.ConexionActiva.Execute ssql
        
            If nTpoCierre = 1 Then
                '----------------------------------------------------
                Call GeneraLogCierre("Termina Proceso Consolida Saldos de Cartera ")
                '----------------------------------------------------
            Else
                MsgBox "Proceso Consolida Saldos de Cartera", vbInformation, "Aviso de Cierre"
            End If
    End If
    
    'EJVG20151118 ***
    GuardarGarantiasDiario
    GeneraMovAsientoGarantiasDiario
    CierreGarantias
    'END EJVG *******
    
    If bCierreMes Then  'ARCV 10-07-2006
        
        
        If nTpoCierre = 1 Then
            '----------------------------------------------------
            Call GeneraLogCierre("Fecha Actual es ultimo dia, copia TEMPORALES antes de cierre de dia. ")
            '----------------------------------------------------
        Else
            MsgBox "El Sistema ha detectado que la Fecha Actual " & Format$(gdFecSis, "yyyy/mm/dd") & " es el último día" & Chr(13) + gPrnSaltoLinea _
            & "hábil del mes y por ello realizara la copia de TEMPORALES antes del proceso de CIERRE DE DIA", vbInformation, "Aviso"
        End If
        
        ' ARCV - SGA 06-07-2006 Se realiza este proceso para que los intereses pasen ya calculados a fin de mes
        'Cierre de Recuperaciones  ' LAYG
        'ALPA 20080723*************************************************
        oConec.CommadTimeOut = 0
        '**************************************************************
        Call CierreRecuperaciones 'Ok ALPA
        
        If CheckConsol.value = 1 Then
            'ALPA 20080723*************************************************
            oConec.CommadTimeOut = 0
            '**************************************************************
            Call GuardaConsolidadaDTS 'Ok ALPA
            ssql = " UPDATE ConstSistema SET nConsSisValor = '" & Format(gdFecSis, "dd/mm/yyyy") & "' WHERE nConsSisCod = 170"
            oConec.ConexionActiva.Execute ssql
            
            If nTpoCierre = 1 Then
                '----------------------------------------------------
                Call GeneraLogCierre("Termina Proceso Consolida Temporales ")
                '----------------------------------------------------
            Else
                MsgBox "Proceso Consolida Temporales Finalizado", vbInformation, "Aviso de Cierre"
            End If
        End If
    End If
    
    If ChkCierreAho.value = 1 And ChkCierreAho.Enabled = True Then
        'ALPA 20080723*************************************************
        oConec.CommadTimeOut = 0
        '**************************************************************
        'MIOL 20121009, SEGUN RQ12272 *********************************
        Call ClientesBloqueadosOrdePago(Format(gdFecSis, "yyyymmdd"), psMovAct)
        'END MIOL *****************************************************
        '***Agregado por ELRO el 20130715, según RFC1306270002****
        Call actualizarNroOperacionesDiariaServicioPago
        '***Fin Agregado por ELRO el 20130715, según RFC1306270002
        Call CancelarAutorizacionDepositos 'RIRO20140407 ERS011
        Call CierreAhorros ' Ok
        
        If nDias = 1 Then
            If nTpoCierre = 1 Then
                '----------------------------------------------------
                Call GeneraLogCierre("Termina Cierre de Ahorros ")
                '----------------------------------------------------
            Else
                MsgBox "Cierre de Ahorros Finalizado", vbInformation, "Aviso de Cierre"
            End If
        End If
        
        'If nDias = 1 Then MsgBox "Cierre de Ahorros Finalizado", vbInformation, "Aviso de Cierre"
        
    End If
    If ChkCierreCred.value = 1 And ChkCierreCred.Enabled = True Then
        'ALPA 20080723*************************************************
        oConec.CommadTimeOut = 0
        '**************************************************************
        Call CierreCreditos 'Ok
        'ALPA 20110827
        
'       ALPA 20120421
        If bCierreMes Then
            If CheckConsol.value = 1 Then
                oConec.CommadTimeOut = 0
                ssql = ""
                ssql = "exec stp_upd_ActualizaDiasAtrasoCierreMes  '" & Format(gdFecSis, "yyyy/mm/dd") & "' "
                oConec.ConexionActiva.Execute ssql
            End If
        End If
        
        'Nuevo Proceso
        'ALPA 20080723*************************************************
        oConec.CommadTimeOut = 0
        '**************************************************************
        Call CierreCartaFianza 'Ok
        
        If nDias = 1 Then
            If nTpoCierre = 1 Then
                '----------------------------------------------------
                Call GeneraLogCierre("Termina Cierre de Creditos ")
                '----------------------------------------------------
            Else
                MsgBox "Cierre de Creditos Finalizado", vbInformation, "Aviso de Cierre"
            End If
        End If
        
        'If nDias = 1 Then MsgBox "Cierre de Creditos Finalizado", vbInformation, "Aviso de Cierre"
        
    End If
    
    If ChkCierrePig.value = 1 And ChkCierrePig.Enabled = True Then
        'ALPA 20080723*************************************************
        oConec.CommadTimeOut = 0
        '**************************************************************
        Call CierrePignoraticio ' Ok
        
        If nDias = 1 Then
            If nTpoCierre = 1 Then
                '----------------------------------------------------
                Call GeneraLogCierre("Termina Cierre de Pignoraticio ")
                '----------------------------------------------------
            Else
                MsgBox "Cierre de Pignoraticio Finalizado", vbInformation, "Aviso de Cierre"
            End If
        End If
        
        'If nDias = 1 Then MsgBox "Cierre de Pignoraticio Finalizado", vbInformation, "Aviso de Cierre"
    End If
    
    If ChkCierrePig.value = 1 And ChkCierreAho.value = 1 And ChkCierreCred.value = 1 Then
        'ALPA 20080723*************************************************
        oConec.CommadTimeOut = 0
        '**************************************************************
        ssql = " UPDATE ConstSistema SET nConsSisValor = '" & Format(gdFecSis, "dd/mm/yyyy") & "' WHERE nConsSisCod = 13"
        oConec.ConexionActiva.Execute ssql
    End If
    
    If nTpoCierre = 1 Then
        '----------------------------------------------------
        Call GeneraLogCierre("Termina Cierre del ")
        '----------------------------------------------------
    Else
        MsgBox "Cierre del " & Format$(gdFecSis, "yyyy/mm/dd") & " Finalizado con éxito", vbInformation, "Aviso"
    End If
    
    If bCierreMes Then
        'ALPA 20080723*************************************************
        oConec.CommadTimeOut = 0
        '**************************************************************
        
        If nTpoCierre = 1 Then
            '----------------------------------------------------
            Call GeneraLogCierre("Fecha actual es ultimo dia. Realiza Cierre de Mes ")
            '----------------------------------------------------
            Call CierreDeMes
        Else
        
            If MsgBox("El Sistema ha detectado que la Fecha Actual " & Format$(gdFecSis, "dd mmmm yyyy") & " es el último día" & Chr(13) + gPrnSaltoLinea _
                & "hábil del mes y por ello recomienda efectual el proceso de " & Chr(13) + gPrnSaltoLinea _
             & "CIERRE DE MES. Desea realizar el Cierre de Mes??", vbQuestion + vbYesNo, "Aviso") = vbYes Then
            
             Call CierreDeMes
            
            End If
       End If
    End If
    '***Modificado por ELRO el 20130627, según TI-ERS019-2013****
    Call ActualizarDespuesGenerarSaldosDiariosCaptaciones '***comentado por Servicio de Pago
    '***Fin Modificado por ELRO el 20130627, según TI-ERS019-2013
    Screen.MousePointer = 0
Next i
'sSql = " master..sp_dboption " & strBaseSQL & ",'trunc. log on chkpt.',false"
'oConec.ConexionActiva.Execute sSql

    '----------------------------------------------------
    Call GeneraLogCierre("Proceso de Cierre Finalizado del " & Format(gdFecSis, "dd/mm/yyyy") & " - ")
    Call GeneraLogCierre("===============================================================================", 0)
    '----------------------------------------------------

    MsgBox "Proceso de Cierre Finalizado", vbExclamation, "Aviso"


Unload Me
End Sub

Private Sub CierreDeMes()
            oConec.ConexionActiva.BeginTrans
            
''''            Set cmd = New adodb.Command
''''            Set prm = New adodb.Parameter
''''            cmd.CommandText = "CaptacCierreMes"
''''            cmd.CommandType = adCmdStoredProc
''''            cmd.Name = "CaptacCierreMes"
''''            Set prm = cmd.CreateParameter("FechaTran", adDate, adParamInput)
''''            cmd.Parameters.Append prm
''''            Set prm = cmd.CreateParameter("Usuario", adChar, adParamInput, 4)
''''            cmd.Parameters.Append prm
''''            Set prm = cmd.CreateParameter("Agencia", adChar, adParamInput, 2)
''''            cmd.Parameters.Append prm
''''            Set cmd.ActiveConnection = oConec.ConexionActiva
''''            '**DAOR 20080329 ***************************************************
''''            'cmd.CommandTimeout = 1800
''''            cmd.CommandTimeout = 0
''''            '*******************************************************************
''''            cmd.Parameters.Refresh
''''            oConec.ConexionActiva.CaptacCierreMes Format(gdFecSis & " " & GetHoraServer(), "dd/mm/yyyy hh:mm:ss"), gsCodUser, Right(gsCodAge, 2)
            
            ssql = ""
            ssql = "exec CaptacCierreMes '" & Format(gdFecSis & " " & GetHoraServer(), "yyyy/mm/dd hh:mm:ss") & "', '" & gsCodUser & "', '" & Right(gsCodAge, 2) & "'"
            oConec.ConexionActiva.Execute ssql
             
            ssql = ""
            ssql = " UPDATE ConstSistema SET nConsSisValor = '" & Format(gdFecSis, "dd/mm/yyyy") & "' WHERE nConsSisCod = 14"
            oConec.ConexionActiva.Execute ssql
            
            '''
            oConec.ConexionActiva.CommitTrans
            oConec.ConexionActiva.BeginTrans
            '''
            
            Call CapGeneraEstadistica 'OK
            oConec.ConexionActiva.CommitTrans
            'Cierre de Recuperaciones  ' LAYG
            'Call CierreRecuperaciones
            
            'Salva Data A Consolidar
            If bCierreMes Then
'''''                Set cmd = New adodb.Command
'''''                cmd.CommandText = "GuardaDataAConsolidarAhorros"
'''''                cmd.CommandType = adCmdStoredProc
'''''                cmd.Name = "GuardaDataAConsolidarAhorros"
'''''                Set cmd.ActiveConnection = oConec.ConexionActiva
'''''                '**DAOR 20080329 **************************************************
'''''                'cmd.CommandTimeout = 72000
'''''                cmd.CommandTimeout = 0
'''''                '******************************************************************
'''''                cmd.Parameters.Refresh
'''''                oConec.ConexionActiva.GuardaDataAConsolidarAhorros
                
                ssql = ""
                ssql = "exec GuardaDataAConsolidarAhorros "
                oConec.ConexionActiva.Execute ssql
                'ALPA 20141231************************************
                ssql = ""
                'ssql = "exec stp_ins_ActualizarCaptacInactivasTotal '" & Format(gdFecSis, "dd/mm/yyyy") & "'" '' PEAC 20150101 - se cambio formato de fecha
                ssql = "exec stp_ins_ActualizarCaptacInactivasTotal '" & Format(gdFecSis, "yyyymmdd") & "'"
                oConec.ConexionActiva.Execute ssql
                '*************************************************
            End If
''''            Set cmd = Nothing
''''            Set prm = Nothing

End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdPreCuadre_Click()
Dim oPrevio As previo.clsprevio
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
    Set oPrevio = New previo.clsprevio
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
Dim oGen As COMDConstSistema.DCOMGeneral
Me.Icon = LoadPicture(App.path & gsRutaIcono)
Set oGen = New COMDConstSistema.DCOMGeneral
lnFiltro = CInt(oGen.LeeConstSistema(104))
'lnBitUsarValidacion = CInt(oGen.LeeConstSistema(8005))
lnBitUsarValidacion = 1

Set oGen = Nothing
    
Set oConec = New COMConecta.DCOMConecta
oConec.AbreConexion
'**DAOR 20080329 **********************************
oConec.CommadTimeOut = 0
'**************************************************
Call CargaDatos
txtFecIni = gdFecSis
txtFecFin = gdFecSis
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    oConec.CierraConexion
    Set oConec = Nothing
End Sub

Private Sub txtFecFin_GotFocus()
With txtFecFin
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub

Private Sub txtFecFin_KeyPress(KeyAscii As Integer)
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
Dim loRecup As COMNColocRec.NCOMColRecCredito
Dim lsMovNro As String
Dim lsFechaHoraGrab As String
Dim loContFunct As COMNContabilidad.NCOMContFunciones
Dim loParam As COMDConstSistema.NCOMConstSistema
Dim lnTipoCalcIntComp As Integer, lnTipoCalcIntMora   As Integer
Set loParam = New COMDConstSistema.NCOMConstSistema
    lnTipoCalcIntComp = loParam.LeeConstSistema(151)
    lnTipoCalcIntMora = loParam.LeeConstSistema(152)
Set loParam = Nothing

    'Genera el Mov Nro
    Set loContFunct = New COMNContabilidad.NCOMContFunciones
        lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Set loContFunct = Nothing
    
    lsFechaHoraGrab = fgFechaHoraGrab(lsMovNro)
    
Set loRecup = New COMNColocRec.NCOMColRecCredito

    Call loRecup.nCierreMesRecuperaciones(lsFechaHoraGrab, "131000", lsMovNro, lnTipoCalcIntComp, lnTipoCalcIntMora)
    
Set loRecup = Nothing
End Sub


Private Function Devuelve_Errores_PreCuadre(ByVal sFecha As String) As String

Dim ssql As String
Dim sCadena As String
Dim nBandera As Integer
Dim sCabecera As String
Dim rs As New ADODB.Recordset
Dim vLenNomb As Integer
Dim vespacio As Integer
Dim vPage As Integer
Dim lsTablaDiaria As String
Dim SFechaAnt As Date

SFechaAnt = DateAdd("d", -1, CDate(sFecha))

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

ssql = "Select a.cCtaCod from captaciones a "
ssql = ssql & " Inner join producto b on a.cctacod = b.cctacod "
ssql = ssql & " inner join productopersona c on a.cctacod = c.cctacod "
ssql = ssql & " inner join persona  d on c.cperscod = d.cperscod "
ssql = ssql & " where npersoneria = " & gPersonaJurCFLCMAC & " and left(nprdestado,2) not in (13,14) "
ssql = ssql & " and a.cctacod not in (select cCtaCod from cuentaifespecial) "

Set rs = oConec.CargaRecordSet(ssql)
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

ssql = " Select MC.nMovNro, MC.nMonto, SUM(MCD.nMonto) Detalle "
ssql = ssql & " From " & lsTablaDiaria & " M JOIN MovCap MC JOIN MovCapDet MCD ON MC.nMovNro = MCD.nMovNro "
ssql = ssql & " And MC.cCtaCod = MCD.cCtaCod And MC.cOpeCod = MCD.cOpeCod ON M.nMovNro = MC.nMovNro "
ssql = ssql & " Where M.cMovNro LIKE '" & Format(sFecha, "YYYYMMdd") & "%' And M.nMovFlag = 0 And MC.cCtaCod LIKE '111_____[12]%' "
ssql = ssql & " And (MCD.cOpeCod + Convert(Varchar(6),MCD.nConceptoCod)) "
ssql = ssql & " IN (Select cOpeCod + Convert(Varchar(6),nConcepto) From OpeCtaNeg Where cCtaContCod LIKE '11_102%') "
ssql = ssql & " Group by MC.nMovNro, MC.nMonto Having MC.nMonto <> SUM(MCD.nMonto) "

Set rs = oConec.CargaRecordSet(ssql)
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

ssql = " Select A.*, B.* From "
ssql = ssql & " (Select MC.nMovNro, nDebe = ISNULL(SUM(CASE WHEN O.cIngEgr = 'I' THEN Abs(MC.nMonto) END),0), "
ssql = ssql & " nHaber = ISNULL(SUM(CASE WHEN O.cIngEgr = 'E' THEN Abs(MC.nMonto) END),0) "
ssql = ssql & " From " & lsTablaDiaria & " M JOIN MovCap MC ON M.nMovNro = MC.nMovNro JOIN "
ssql = ssql & " (Select O.cOpeCod, G.cIngEgr From OpeTpo O JOIN GruposOpe GO JOIN GrupoOpe G ON "
ssql = ssql & " GO.cGrupoCod = G.cGrupoCod ON O.cOpeCod = GO.cOpeCod Where G.nEfectivo = 1 Group by O.cOpeCod, G.cIngEgr) O "
ssql = ssql & " ON MC.cOpeCod = O.cOpeCod "
ssql = ssql & " Where M.cMovNro LIKE '" & Format(sFecha, "YYYYMMdd") & "%' And M.nMovFlag = 0 And MC.cCtaCod LIKE '111_____[12]%' "
ssql = ssql & " Group by MC.nMovNro) A JOIN "
ssql = ssql & " (Select MC.nMovNro as nMovNro1, "
ssql = ssql & " nDebe1 = ISNULL(SUM(CASE WHEN C.cOpeCtaDH = 'D' THEN Abs(MCD.nMonto) END),0), "
ssql = ssql & " nHaber1 = ISNULL(SUM(CASE WHEN C.cOpeCtaDH = 'H' THEN Abs(MCD.nMonto) END),0) "
ssql = ssql & " From " & lsTablaDiaria & " M JOIN MovCap MC JOIN MovCapDet MCD "
ssql = ssql & " JOIN (Select cOpeCod, nConcepto, cOpeCtaDH From OpeCtaNeg Where cCtaContCod LIKE '11_102%' Group by "
ssql = ssql & " cOpeCod, nConcepto, cOpeCtaDH) C ON MCD.cOpeCod = C.cOpeCod And MCD.nConceptoCod = C.nConcepto "
ssql = ssql & " ON MC.nMovNro = MCD.nMovNro And MC.cCtaCod = MCD.cCtaCod And MC.cOpeCod = MCD.cOpeCod ON M.nMovNro = MC.nMovNro "
ssql = ssql & " Where M.cMovNro LIKE '" & Format(sFecha, "YYYYMMdd") & "%' And M.nMovFlag = 0 And MC.cCtaCod LIKE '1110____[12]%' "
ssql = ssql & " Group by MC.nMovNro) B ON A.nMovNro = B.nMovNro1 "
ssql = ssql & " Where (A.nDebe <> B.nDebe1 Or A.nHaber <> B.nHaber1)"

Set rs = oConec.CargaRecordSet(ssql)
If rs.BOF Then
Else
    nBandera = 2
    sCadena = sCadena & Chr(10) & Chr(10)
    sCadena = sCadena & "** Existen descuadres entre la Caja, Asientos y Planilla Consolidada Captaciones " & Chr(10)
    sCadena = sCadena & "   _____________________________________________________________________________ " & Chr(10) & Chr(10)
    sCadena = sCadena & "   ** nMovNro    Debe         nMovNro    Debe          ** " & Chr(10)
    sCadena = sCadena & "   ** ================================================ ** " & Chr(10) & Chr(10)
    
    Do While Not rs.EOF
        sCadena = sCadena & "   *  " & Left(rs!nMovNro & "          ", 10) & " " & ImpreFormat(rs!nDebe, 10, 2) & " " & ImpreFormat(rs!nHaber, 10, 2)
        sCadena = sCadena & " " & Left(rs!nmovnro1 & "          ", 10) & " " & ImpreFormat(rs!ndebe1, 10, 2) & " " & ImpreFormat(rs!nhaber1, 10, 2) & Chr(10)
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

ssql = " Select MC.nMovNro, MC.nMonto, SUM(MCD.nMonto) Detalle "
ssql = ssql & " From " & lsTablaDiaria & " M JOIN MovCol MC JOIN MovColDet MCD ON MC.nMovNro = MCD.nMovNro"
ssql = ssql & " And MC.cCtaCod = MCD.cCtaCod And MC.cOpeCod = MCD.cOpeCod ON M.nMovNro = MC.nMovNro"
ssql = ssql & " Where M.cMovNro LIKE '" & Format(sFecha, "YYYYMMdd") & "%' And M.nMovFlag = 0 And MC.cCtaCod LIKE '111_____[12]%' "
ssql = ssql & " And (MCD.cOpeCod + Convert(Varchar(6),MCD.nPrdConceptoCod)) "
ssql = ssql & " IN (Select cOpeCod + Convert(Varchar(6),nConcepto) From OpeCtaNeg Where cCtaContCod LIKE '11_102%') "
ssql = ssql & " Group by MC.nMovNro, MC.nMonto Having MC.nMonto <> SUM(MCD.nMonto)"

Set rs = oConec.CargaRecordSet(ssql)
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

ssql = "Select P.cCtaCod, P.nSaldo, "
ssql = ssql & " nSaldoCal = ( "
ssql = ssql & " Select SUM(nMonto) "
ssql = ssql & " From ( "
ssql = ssql & " Select SUM(nMonto-nMontoPagado) as nMonto "
ssql = ssql & " From ColocCalendDet CD "
ssql = ssql & " Inner Join ColocCalendario Cal ON Cal.cCtaCod = CD.cCtaCod AND Cal.nNroCalen = CD.nNroCalen AND Cal.nColocCalendEstado = 0"
ssql = ssql & " AND Cal.nColocCalendApl = CD.nColocCalendApl AND Cal.nCuota = CD.nCuota AND CD.nPrdConceptoCod = 1000 "
ssql = ssql & " Where CD.nNroCalen = CC.nNroCalen And CD.cCtaCod = CC.cCtaCod "
ssql = ssql & " AND CD.nColocCalendApl = 1 "
ssql = ssql & " Union All "
ssql = ssql & " Select ISNULL(SUM(nMonto-nMontoPagado),0) as nMonto "
ssql = ssql & " From ColocCalendDet CD "
ssql = ssql & " Inner Join ColocCalendario Cal ON Cal.cCtaCod = CD.cCtaCod AND Cal.nNroCalen = CD.nNroCalen AND Cal.nColocCalendEstado = 0"
ssql = ssql & " AND Cal.nColocCalendApl = CD.nColocCalendApl AND Cal.nCuota = CD.nCuota AND CD.nPrdConceptoCod = 1000 "
ssql = ssql & " Where CD.nNroCalen = CC.nNroCalPar And CD.cCtaCod = CC.cCtaCod "
ssql = ssql & " AND CD.nColocCalendApl = 1 "
ssql = ssql & " ) as T ), "
ssql = ssql & " TIPO=Case CC.CRFA "
ssql = ssql & " When 'RFA' Then 'RFA' "
ssql = ssql & " When 'RFC' Then 'RFC' "
ssql = ssql & " When 'DIF' Then 'DIF' "
ssql = ssql & " When 'NOR' Then 'NOR' "
ssql = ssql & " End "
ssql = ssql & " from Producto P "
ssql = ssql & " Inner Join ColocacCred CC ON CC.cCtaCod = P.cCtaCod  "
ssql = ssql & " Where P.nPrdEstado in (2020,2021,2022,2030,2031,2032) AND CC.CRFA not in ('RFA','RFC','DIF') "
ssql = ssql & " AND "
ssql = ssql & " P.nSaldo <> "
ssql = ssql & " ( "
ssql = ssql & " Select SUM(nMonto) "
ssql = ssql & " From ( "
ssql = ssql & " Select SUM(nMonto-nMontoPagado) as nMonto "
ssql = ssql & " From ColocCalendDet CD "
ssql = ssql & " Inner Join ColocCalendario Cal ON Cal.cCtaCod = CD.cCtaCod AND Cal.nNroCalen = CD.nNroCalen AND Cal.nColocCalendEstado = 0 "
ssql = ssql & " AND Cal.nColocCalendApl = CD.nColocCalendApl AND Cal.nCuota = CD.nCuota AND CD.nPrdConceptoCod = 1000  "
ssql = ssql & " Where CD.nNroCalen = CC.nNroCalen And CD.cCtaCod = CC.cCtaCod "
ssql = ssql & " AND CD.nColocCalendApl = 1  "
ssql = ssql & " Union All  "
ssql = ssql & " Select ISNULL(SUM(nMonto-nMontoPagado),0) as nMonto  "
ssql = ssql & " From ColocCalendDet CD "
ssql = ssql & " Inner Join ColocCalendario Cal ON Cal.cCtaCod = CD.cCtaCod AND Cal.nNroCalen = CD.nNroCalen AND Cal.nColocCalendEstado = 0 "
ssql = ssql & " AND Cal.nColocCalendApl = CD.nColocCalendApl AND Cal.nCuota = CD.nCuota AND CD.nPrdConceptoCod = 1000 "
ssql = ssql & " Where CD.nNroCalen = CC.nNroCalPar And CD.cCtaCod = CC.cCtaCod "
ssql = ssql & " AND CD.nColocCalendApl = 1 "
ssql = ssql & " ) as T ) "


vPage = vPage + 1
sCadena = sCadena & Chr(12)

Set rs = oConec.CargaRecordSet(ssql)
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
'--------*************PRENDARIO*************---------------
'Valida fecha renovacion y Fecha Vencimiento
'sFecha


ssql = " select p.cctacod,j.fechaven,j.fechapago,c.dvenc,p.nprdestado            "
ssql = ssql & " from  producto p"
ssql = ssql & " join ("
ssql = ssql & "  SELECT Fechaven=dateadd(d,30,cast(left(cmovnro,8) as datetime)),FechaPago=cast(left(cmovnro,8) as datetime),mc.cctacod,mc.copecod"
ssql = ssql & "  from  mov m"
ssql = ssql & "  join movcol mc on mc.nmovnro=m.nmovnro"
ssql = ssql & "  where mc.cctacod like '108__305%'"
ssql = ssql & "  and m.nmovflag=0 and (mc.copecod like '1211%' or mc.copecod like '1261%')"
ssql = ssql & "  and m.nmovnro in (select max(mg.nmovnro) from mov g join movcol mg on mg.nmovnro=g.nmovnro where mg.cctacod=mc.cctacod"
ssql = ssql & "  and (mg.copecod like '1211%' or mg.copecod like '1261%') )"
ssql = ssql & " ) J on j.cctacod =p.cctacod"
ssql = ssql & " join colocaciones c on c.cctacod=p.cctacod"
ssql = ssql & " where p.nprdestado not in (2102,2103,2108,2109,2110,2111,2112) and Fechaven<>c.dvenc   "


vPage = vPage + 1
sCadena = sCadena & Chr(12)

Set rs = oConec.CargaRecordSet(ssql)
If rs.BOF Then
Else
    nBandera = 2
    sCadena = sCadena & Chr(10) & Chr(10)
    sCadena = sCadena & "** Existen cuentas de Prendario en las que no coincide la FECHA DE PAGO + 30 dias con la FECHA ACTUAL DE VENCIMIENTO " & Chr(10)
    sCadena = sCadena & "   ______________________________________________________________________________ " & Chr(10) & Chr(10)
    sCadena = sCadena & "   ** Cuenta             Fecha Pago  Fecha Venc. Sugerida  Fecha Venc. Actual       ** " & Chr(10)
    sCadena = sCadena & "   ** ============================================================================= ** " & Chr(10) & Chr(10)
    
    Do While Not rs.EOF
        sCadena = sCadena & "   *  " & rs!cCtaCod & " " & Format(rs!Fechapago, "dd/MM/yyyy") & "  " & Format(rs!fechaven, "dd/MM/yyyy") & Space(12) & Format(rs!dVenc, "dd/MM/yyyy") & Chr(10)
        rs.MoveNext
    Loop
End If
rs.Close
Set rs = Nothing

'Valida Producto con Moviemiento Pagado del Día Prendario-AMORTIZACION

ssql = " Select c.nmovnro,c.cctacod,c.nmonto,c.nSaldoMov,c.nmonto,d.nsaldo,d.nsaldocap "
ssql = ssql & " From "
ssql = ssql & " ("
ssql = ssql & " select  m.nmovnro,m.cmovnro,mc.cctacod,mc.copecod,nSaldoMov=mc.nsaldocap,nmonto=round(md.nmonto,2)"
ssql = ssql & " from mov m"
ssql = ssql & " join movcol mc on mc.nmovnro=m.nmovnro"
ssql = ssql & " join movcoldet md on md.nmovnro=mc.nmovnro and md.copecod=mc.copecod and md.nprdconceptocod=2000"
ssql = ssql & " where  m.nmovflag=0 and (mc.copecod like '1210%' or mc.copecod like '1263%')"
ssql = ssql & " and m.cmovnro like '" & Format(sFecha, "YYYYMMdd") & "%'"
ssql = ssql & " ) C"
ssql = ssql & " Join"
ssql = ssql & " ("
ssql = ssql & " select s.dfecha,p.cctacod,p.nsaldo,p.nprdestado,s.nsaldocap"
ssql = ssql & " from producto p"
ssql = ssql & " join colocacsaldo s on s.cctacod=p.cctacod"
ssql = ssql & " where convert(char(8),s.dfecha,112) like '" & Format(SFechaAnt, "YYYYMMdd") & "%' and p.cctacod like '108__305%'"
ssql = ssql & " ) D on  d.cctacod=c.cctacod"
ssql = ssql & " Where d.nSaldo <> (d.nsaldocap - c.nMonto) Or d.nSaldo <> c.nSaldoMov"

vPage = vPage + 1
sCadena = sCadena & Chr(12)

Set rs = oConec.CargaRecordSet(ssql)
If rs.BOF Then
Else
    nBandera = 2
    sCadena = sCadena & Chr(10) & Chr(10)
    sCadena = sCadena & "** Existen cuentas de Prendario en las que no coincide el Saldo de Producto con Movimientos *AMORTIZACION* " & Chr(10)
    sCadena = sCadena & "   ________________________________________________________________________________ " & Chr(10) & Chr(10)
    sCadena = sCadena & "   ** Cuenta             NroMov  Monto       Saldo Ayer    Saldo Mov.  Saldo Actual       ** " & Chr(10)
    sCadena = sCadena & "   ** ============================================================================= ** " & Chr(10) & Chr(10)
    
    Do While Not rs.EOF
        sCadena = sCadena & "   *  " & rs!cCtaCod & " " & rs!nMovNro & "  " & ImpreFormat(rs!nMonto, 8, 2, False) & " " & ImpreFormat(rs!nSaldoCap, 8, 2) & " " & ImpreFormat(rs!nSaldoMov, 8, 2) & " " & ImpreFormat(rs!nSaldo, 8, 2) & Chr(10)
        rs.MoveNext
    Loop
End If
rs.Close
Set rs = Nothing

'Valida Producto con Moviemiento Pagado del Día Prendario-RENOVACION
ssql = " Select c.nmovnro,c.cctacod,c.nmonto,c.nsaldomov,c.nmonto,d.nsaldo,d.nsaldocap"
ssql = ssql & " From"
ssql = ssql & "( "
ssql = ssql & " select  m.nmovnro,m.cmovnro,mc.cctacod,mc.copecod,nsaldomov=mc.nsaldocap,nmonto=round(md.nmonto,2)"
ssql = ssql & " from mov m"
ssql = ssql & " join movcol mc on mc.nmovnro=m.nmovnro"
ssql = ssql & " join movcoldet md on md.nmovnro=mc.nmovnro and md.copecod=mc.copecod and md.nprdconceptocod=2000"
ssql = ssql & " where  m.nmovflag=0 and  mc.copecod like '12[16]1%'"
ssql = ssql & " and m.cmovnro like '" & Format(sFecha, "YYYYMMdd") & "%'"
ssql = ssql & " ) C"
ssql = ssql & " Join"
ssql = ssql & "("
ssql = ssql & " select s.dfecha,p.cctacod,p.nsaldo,p.nprdestado,s.nsaldocap"
ssql = ssql & " from producto p"
ssql = ssql & " join colocacsaldo s on s.cctacod=p.cctacod"
ssql = ssql & " where convert(char(8),s.dfecha,112) like '" & Format(SFechaAnt, "YYYYMMdd") & "%' and p.cctacod like '108__305%'"
ssql = ssql & " ) D on  d.cctacod=c.cctacod"
ssql = ssql & " Where d.nSaldo <> (d.nsaldocap - c.nMonto) Or d.nSaldo <> c.nSaldoMov "

vPage = vPage + 1
sCadena = sCadena & Chr(12)

Set rs = oConec.CargaRecordSet(ssql)
If rs.BOF Then
Else
    nBandera = 2
    sCadena = sCadena & Chr(10) & Chr(10)
    sCadena = sCadena & "** Existen cuentas de Prendario en las que no coincide el Saldo de Producto con Movimientos *RENOVACION*" & Chr(10)
    sCadena = sCadena & "   ________________________________________________________________________________ " & Chr(10) & Chr(10)
    sCadena = sCadena & "   ** Cuenta             NroMov  Monto       Saldo Ayer    Saldo Mov.  Saldo Actual       ** " & Chr(10)
    sCadena = sCadena & "   ** ============================================================================= ** " & Chr(10) & Chr(10)
    
    Do While Not rs.EOF
        sCadena = sCadena & "   *  " & rs!cCtaCod & " " & rs!nMovNro & "  " & ImpreFormat(rs!nMonto, 8, 2, False) & " " & ImpreFormat(rs!nSaldoCap, 8, 2) & " " & ImpreFormat(rs!nSaldoMov, 8, 2) & " " & ImpreFormat(rs!nSaldo, 8, 2) & Chr(10)
        rs.MoveNext
    Loop
End If
rs.Close
Set rs = Nothing




'-------************************************************
'Validacion de Producto Vs el MovimientoPagado del Dia
ssql = "Select CP.cCtaCod,CMOV.nMonto,CP.nDif,CP.nSaldo "
ssql = ssql & " From ("
ssql = ssql & " Select Sum(MD.nMonto)as nMonto,MD.cCtaCod"
ssql = ssql & "    From Mov M"
ssql = ssql & "         Inner Join MovColDet MD on M.nMovNro=MD.nMovNro"
ssql = ssql & "    Where Left(M.cMovNro,8)='" & Format(gdFecSis, "yyyymmdd") & "' and M.nMovFlag=0 and MD.nPrdConceptoCod in(1000,1109,1110,1010) and"
ssql = ssql & "              MD.cOpeCod not like '107[123456789]%' and MD.cOpeCod not in ('100101','100102','100103','100104','100105') and"
ssql = ssql & "              MD.cOpeCod <>'107002'  and MD.cOpeCod<>'107003' AND ISNULL(MD.cOpeCod,'')<>''"
ssql = ssql & "        Group By MD.cCtaCod)CMOV"
ssql = ssql & "     Inner Join ("
ssql = ssql & "            Select P.cCtaCod,P.nSaldo,CS.nSaldoCap-P.nSaldo as nDIf"
ssql = ssql & "            From Producto P"
ssql = ssql & "                 Inner Join ColocacSaldo CS on CS.cCtaCod=P.cCtaCod"
ssql = ssql & "             Where P.nPrdEstado in (2020,2021,2022,2030,2031,2032) and"
ssql = ssql & "                   Convert(Varchar(20),CS.dFecha,112)='" & Format(DateAdd("d", -1, gdFecSis), "YYYYMMDD") & "') CP on CP.cCtaCod=CMOV.cCtaCod"
ssql = ssql & " Where CMOV.nMonto<>CP.nDif "


Set rs = oConec.CargaRecordSet(ssql)
If rs.BOF Then
Else
    nBandera = 2
    sCadena = sCadena & Chr(10) & Chr(10)
    sCadena = sCadena & "** Existen cuentas en las que no coincide el saldo de producto y el de calendario creditos NORMALES" & Chr(10)
    sCadena = sCadena & "   ______________________________________________________________________________ " & Chr(10) & Chr(10)
    sCadena = sCadena & "   ** Cuenta             S. Movimiento   nDif    S.Producto** " & Chr(10)
    sCadena = sCadena & "   ** ============================================== ** " & Chr(10) & Chr(10)
    
    Do While Not rs.EOF
        sCadena = sCadena & "   *  " & rs!cCtaCod & " " & ImpreFormat(rs!nMonto, 10, 2) & " " & ImpreFormat(rs!nDif, 10, 2) & " " & ImpreFormat(rs!nSaldo, 10, 2) & Chr(10)
        rs.MoveNext
    Loop
End If
rs.Close
Set rs = Nothing


'******Validacion de la caja de creditos***** no cogiendo desaguio
ssql = "Select  m.nmovnro, m.copecod,mc.cCtaCod,round(mc.nMonto,2) as nMontoCaja, round(sum(mcd.nMonto),2) as nMontoDetalle,"
ssql = ssql & " round(mc.nMonto,2)- round(sum(mcd.nMonto),2) as ndif"
ssql = ssql & " from    mov m"
ssql = ssql & " join (select nmovnro, copecod, cCtaCod, sum(nMonto) nMonto"
ssql = ssql & " From movcol"
ssql = ssql & " group by nmovnro, copecod, cCtaCod) mc"
ssql = ssql & " on mc.nmovnro = m.nmovnro"
ssql = ssql & " join movcoldet mcd on mcd.nmovnro = mc.nmovnro and mcd.copecod = mc.copecod and mc.cctacod = mcd.cctacod"
ssql = ssql & " join opectaneg oc on oc.copecod = mcd.copecod and oc.nConcepto = mcd.nPrdConceptoCod and oc.cCtaContCod like '11_102%'"
ssql = ssql & " where   m.cmovnro like  '" & Format(gdFecSis, "yyyymmdd") & "%' and m.nmovflag = 0 and"
ssql = ssql & " not exists(Select * From MovColDet Where cCtacod=mc.cCtacod and nPrdConceptoCod=1106 and nMovNro=mc.nmovnro)"
ssql = ssql & " group by m.cmovnro, m.nmovnro, m.copecod, m.cmovdesc, mc.cCtaCod, mc.nMonto"
ssql = ssql & " having round(mc.nMonto,2) <> round(sum(mcd.nMonto),2)"


Set rs = oConec.CargaRecordSet(ssql)
If rs.BOF Then
Else
    nBandera = 2
    sCadena = sCadena & Chr(10) & Chr(10)
    sCadena = sCadena & "** La Caja de Creditos No Cuadra Porfavor revisar" & Chr(10)
    sCadena = sCadena & String(120, "_") & Chr(10) & Chr(10)
    sCadena = sCadena & "   **" & ImpreFormat("Nro Movimento", 20) & ImpreFormat(" Operacion", 10) & ImpreFormat("Credito", 20) & Space(12) & ImpreFormat("Monto Caja", 15) & ImpreFormat(" MontoDetalle", 15) & ImpreFormat(" Diferencia", 10) & "** " & Chr(10)
    sCadena = sCadena & "   ** " & String(120, "=") & " ** " & Chr(10) & Chr(10)
    
    Do While Not rs.EOF
        sCadena = sCadena & ImpreFormat(rs!nMovNro, 32, 0) & ImpreFormat(rs!cOpecod, 10) & ImpreFormat(rs!cCtaCod, 20) & ImpreFormat(rs!nMontoCaja, 15, 2) & ImpreFormat(rs!nMontoDetalle, 15, 2) & ImpreFormat(rs!nDif, 10, 2) & Chr(10)
        rs.MoveNext
    Loop
End If
rs.Close
Set rs = Nothing


'Valida que el saldo de calendario sea igual que el saldo de producto creditos RFA
ssql = "Select P.cCtaCod, P.nSaldo, "
ssql = ssql & " nSaldoCal = ( "
ssql = ssql & " Select SUM(nMonto) "
ssql = ssql & " From ( "
ssql = ssql & " Select SUM(nMonto-nMontoPagado) as nMonto "
ssql = ssql & " From ColocCalendDet CD "
ssql = ssql & " Inner Join ColocCalendario Cal ON Cal.cCtaCod = CD.cCtaCod AND Cal.nNroCalen = CD.nNroCalen AND Cal.nColocCalendEstado = 0"
ssql = ssql & " AND Cal.nColocCalendApl = CD.nColocCalendApl AND Cal.nCuota = CD.nCuota AND CD.nPrdConceptoCod = 1000 "
ssql = ssql & " Where CD.nNroCalen = CC.nNroCalen And CD.cCtaCod = CC.cCtaCod "
ssql = ssql & " AND CD.nColocCalendApl = 1 "
ssql = ssql & " Union All "
ssql = ssql & " Select ISNULL(SUM(nMonto-nMontoPagado),0) as nMonto "
ssql = ssql & " From ColocCalendDet CD "
ssql = ssql & " Inner Join ColocCalendario Cal ON Cal.cCtaCod = CD.cCtaCod AND Cal.nNroCalen = CD.nNroCalen AND Cal.nColocCalendEstado = 0"
ssql = ssql & " AND Cal.nColocCalendApl = CD.nColocCalendApl AND Cal.nCuota = CD.nCuota AND CD.nPrdConceptoCod = 1000 "
ssql = ssql & " Where CD.nNroCalen = CC.nNroCalPar And CD.cCtaCod = CC.cCtaCod "
ssql = ssql & " AND CD.nColocCalendApl = 1 "
ssql = ssql & " ) as T ), "
ssql = ssql & " TIPO=Case CC.CRFA"
ssql = ssql & " When 'RFA' Then 'RFA'"
ssql = ssql & " When 'RFC' Then 'RFC'"
ssql = ssql & " When 'DIF' Then 'DIF'"
ssql = ssql & " When 'NOR' Then 'NOR'"
ssql = ssql & " End"
ssql = ssql & " from Producto P "
ssql = ssql & " Inner Join ColocacCred CC ON CC.cCtaCod = P.cCtaCod  "
ssql = ssql & " Where P.nPrdEstado in (2020,2021,2022,2030,2031,2032) AND CC.CRFA in ('RFA','RFC','DIF') "
ssql = ssql & " AND "
ssql = ssql & " P.nSaldo <> "
ssql = ssql & " ( "
ssql = ssql & " Select SUM(nMonto) "
ssql = ssql & " From ( "
ssql = ssql & " Select SUM(nMonto-nMontoPagado) as nMonto "
ssql = ssql & " From ColocCalendDet CD "
ssql = ssql & " Inner Join ColocCalendario Cal ON Cal.cCtaCod = CD.cCtaCod AND Cal.nNroCalen = CD.nNroCalen AND Cal.nColocCalendEstado = 0"
ssql = ssql & " AND Cal.nColocCalendApl = CD.nColocCalendApl AND Cal.nCuota = CD.nCuota AND CD.nPrdConceptoCod = 1000 "
ssql = ssql & " Where CD.nNroCalen = CC.nNroCalen And CD.cCtaCod = CC.cCtaCod "
ssql = ssql & " AND CD.nColocCalendApl = 1 "
ssql = ssql & " Union All "
ssql = ssql & " Select ISNULL(SUM(nMonto-nMontoPagado),0) as nMonto "
ssql = ssql & " From ColocCalendDet CD "
ssql = ssql & " Inner Join ColocCalendario Cal ON Cal.cCtaCod = CD.cCtaCod AND Cal.nNroCalen = CD.nNroCalen AND Cal.nColocCalendEstado = 0 "
ssql = ssql & " AND Cal.nColocCalendApl = CD.nColocCalendApl AND Cal.nCuota = CD.nCuota AND CD.nPrdConceptoCod = 1000 "
ssql = ssql & " Where CD.nNroCalen = CC.nNroCalPar And CD.cCtaCod = CC.cCtaCod "
ssql = ssql & " AND CD.nColocCalendApl = 1 "
ssql = ssql & " ) as T ) "

vPage = vPage + 1
sCadena = sCadena & Chr(12)

Set rs = oConec.CargaRecordSet(ssql)
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

ssql = "Select A.*, B.* From "
ssql = ssql & " (Select MC.nMovNro, nDebe = ISNULL(SUM(CASE WHEN O.cIngEgr = 'I' THEN Abs(MC.nMonto) END),0), "
ssql = ssql & " nHaber = ISNULL(SUM(CASE WHEN O.cIngEgr = 'E' THEN Abs(MC.nMonto) END),0) "
ssql = ssql & " From " & lsTablaDiaria & " M JOIN MovCol MC ON M.nMovNro = MC.nMovNro JOIN "
ssql = ssql & " (Select O.cOpeCod, G.cIngEgr From OpeTpo O JOIN GruposOpe GO JOIN GrupoOpe G ON "
ssql = ssql & " GO.cGrupoCod = G.cGrupoCod ON O.cOpeCod = GO.cOpeCod Where G.nEfectivo = 1 Group by O.cOpeCod, G.cIngEgr) O "
ssql = ssql & " ON MC.cOpeCod = O.cOpeCod "
ssql = ssql & " Where M.cMovNro LIKE '" & Format(sFecha, "YYYYMMdd") & "%' And M.nMovFlag = 0 And MC.cCtaCod LIKE '111_____[12]%' "
ssql = ssql & " Group by MC.nMovNro) A "
ssql = ssql & " Join "
ssql = ssql & " (Select MC.nMovNro as nMovNro1, "
ssql = ssql & " nDebe1 = ISNULL(SUM(CASE WHEN C.cOpeCtaDH = 'D' THEN ABS(MCD.nMonto) END),0), "
ssql = ssql & " nHaber1 = ISNULL(SUM(CASE WHEN C.cOpeCtaDH = 'H' THEN ABS(MCD.nMonto) END),0) "
ssql = ssql & " From " & lsTablaDiaria & " M JOIN MovCol MC JOIN MovColDet MCD "
ssql = ssql & " JOIN (Select cOpeCod, nConcepto, cOpeCtaDH From OpeCtaNeg Where cCtaContCod LIKE '11_102%' Group by "
ssql = ssql & " cOpeCod, nConcepto, cOpeCtaDH) C ON MCD.cOpeCod = C.cOpeCod And MCD.nPrdConceptoCod = C.nConcepto "
ssql = ssql & " ON MC.nMovNro = MCD.nMovNro And MC.cCtaCod = MCD.cCtaCod And MC.cOpeCod = MCD.cOpeCod ON M.nMovNro = MC.nMovNro "
ssql = ssql & " Where M.cMovNro LIKE '" & Format(sFecha, "YYYYMMdd") & "%' And M.nMovFlag = 0 And MC.cCtaCod LIKE '111_____[12]%' "
ssql = ssql & " Group by MC.nMovNro) B "
ssql = ssql & " ON A.nMovNro = B.nMovNro1 "
ssql = ssql & " Where (A.nDebe - A.nHaber <> B.nDebe1 - B.nHaber1) "

Set rs = oConec.CargaRecordSet(ssql)
If rs.BOF Then
Else
    nBandera = 2
    sCadena = sCadena & Chr(10) & Chr(10)
    sCadena = sCadena & "** Existen descuadres entre la Caja, Asientos y Planilla Consolidada Colocaciones " & Chr(10)
    sCadena = sCadena & "   ______________________________________________________________________________ " & Chr(10) & Chr(10)
    sCadena = sCadena & "   ** nMovNro    Debe         nMovNro    Debe          ** " & Chr(10)
    sCadena = sCadena & "   ** ================================================ ** " & Chr(10) & Chr(10)
    
    Do While Not rs.EOF
        sCadena = sCadena & "   *  " & Left(rs!nMovNro & "          ", 10) & " " & ImpreFormat(rs!nDebe, 10, 2) & " " & ImpreFormat(rs!nHaber, 10, 2)
        sCadena = sCadena & " " & Left(rs!nmovnro1 & "          ", 10) & " " & ImpreFormat(rs!ndebe1, 10, 2) & " " & ImpreFormat(rs!nhaber1, 10, 2) & Chr(10)
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

ssql = " Select substring(m1.cMovNro, 18,2) as cAgencia, right(m1.cMovNro, 4) as cUser, "
ssql = ssql & " m1.cOpeCod, m1.cOpeDesc,"
ssql = ssql & " m1.nMovNro as nMovCmac, m1.cMovNro as cMovCmac,  isnull(m2.nMovNroRef, 0) as nMovOpeVarias, "
ssql = ssql & " isnull(M2.cMovNro,'') as cMovOpeVarias,  isnull(m3.nMovNroRef, 0) as nMovRegu, isnull(M3.cMovNro,'') as cMovRegu "
ssql = ssql & " From "
ssql = ssql & " (  select m.cOpeCod, cOpeDesc=case  when m.copecod='260501' then 'LLAMADA DEPOSITO EFECTIVO CMAC' "
ssql = ssql & " when m.copecod='260503' then 'LLAMADA RETIRO EFECTIVO CMAC'  when m.copecod='260504' "
ssql = ssql & " then 'LLAMADA RETIRO ORDEN DE PAGO'  when m.copecod='107001' then 'LLAMADA PAGO DE CREDITO' "
ssql = ssql & " end, m.nMovNro , m.cMovNro from " & lsTablaDiaria & " m inner join movcmac mc on m.nmovnro=mc.nmovnro "
ssql = ssql & " inner join persona p on mc.cperscod=p.cperscod  where substring(m.cmovnro,1,8)='" & Format(sFecha, "YYYYMMdd") & "' and "
ssql = ssql & " m.copecod in(260501, 260503, 260504, 107001) and m.nmovflag=0 ) m1 "
ssql = ssql & " Left Join "
ssql = ssql & " (  Select MR.nMovNro, M1.cMovNro, MR.nMovNroRef "
ssql = ssql & " from MovRef MR Inner Join MovOpevarias MO  on MR.nMovNroRef=MO.nMovNro "
ssql = ssql & " Inner Join Mov M1 On M1.nMovNro=MR.nMovNroRef  ) m2 "
ssql = ssql & " on m1.nMovNro = m2.nMovNro "
ssql = ssql & " Left Join  (  Select MR.nMovNro, M1.cMovNro, MR.nMovNroRef, MC.nMonto as nImporteRegu, "
ssql = ssql & " substring(MC.cCtaCod, 9,1) as cMonedaRegu  from MovRef MR Inner Join MovCap MC "
ssql = ssql & " on MR.nMovNroRef=MC.nMovNro Inner Join Mov M1 On M1.nMovNro=MR.nMovNroRef  ) m3 "
ssql = ssql & " on m1.nMovNro = m3.nMovNro "
ssql = ssql & " Where m2.nMovNroRef Is Null Or m2.nMovNroRef = 0 Or m3.nMovNroRef Is Null Or m3.nMovNroRef = 0 "
ssql = ssql & " Order By substring(m1.cMovNro, 18,2),  right(m1.cMovNro, 4), "
ssql = ssql & " m1.cOpeCod , m1.cMovNro "

Set rs = oConec.CargaRecordSet(ssql)
If rs.BOF Then
Else
    nBandera = 2
    sCadena = sCadena & "** Existen llamadas a CMACS sin regularización y/o comision " & Chr(10)
    sCadena = sCadena & "   ________________________________________________________ " & Chr(10) & Chr(10)
    sCadena = sCadena & "   ** Ag User OpeCod Operacion                 nMovCmac  nMovOpeVar nMovRegu  **" & Chr(10)
    sCadena = sCadena & "   ** ======================================================================  **" & Chr(10) & Chr(10)
    Do While Not rs.EOF
        sCadena = sCadena & "   *  " & rs!CAgencia & " " & rs!cUser & " " & rs!cOpecod & " " & Left(rs!cOpedesc & "                    ", 25) & " "
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

ssql = "Select cctacontcod, cctacontdesc, cultimaactualizacion "
ssql = ssql & " From ctacont "
ssql = ssql & " Where Len(cctacontcod) > 3 "
ssql = ssql & " and substring(cctacontcod,1, len(cctacontcod)-2)  not in (select cctacontcod from ctacont) "

Set rs = oConec.CargaRecordSet(ssql)
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
ssql = " Select cCtaCod From Producto where nPrdEstado in (2020,2021,2022,2030,2031,2032) " _
     & " And Substring(cctacod,6,3) = '301' and cCtaCod not in (Select cCtaCod from ColocacConvenio) "
Set rs = oConec.CargaRecordSet(ssql)
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
ssql = " Select cCtaCod, Count(*) From ProductoPersona where nPrdPersRelac = 29 " _
     & " Group By cCtaCod Having Count(*) > 1 "
Set rs = oConec.CargaRecordSet(ssql)
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
ssql = " Select cCtaCod, Count(*) From ProductoPersona where nPrdPersRelac = 28 " _
     & " Group By cCtaCod Having Count(*) > 1 "
Set rs = oConec.CargaRecordSet(ssql)
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

Public Sub CierreCartaFianza()
  Dim OCFN As COMNCartaFianza.NCOMCartaFianza
  Set OCFN = New COMNCartaFianza.NCOMCartaFianza
    OCFN.CierreCartaFianza gdFecSis
  Set OCFN = Nothing
End Sub
'***Agregado por ELRO el 20130627, según TI-ERS019-2012****
Private Sub ActualizarDespuesGenerarSaldosDiariosCaptaciones()
    oConec.ConexionActiva.BeginTrans
    oConec.CommadTimeOut = 0
    ssql = ""
    ssql = "exec stp_ins_ActualizarDespuesGenerarSaldosDiariosCaptaciones '" & Format(gdFecSis, "yyyy/mm/dd") & "', '" & gsCodUser & "'"
    oConec.ConexionActiva.Execute ssql
    oConec.ConexionActiva.CommitTrans
End Sub
'***Fin Agregado por ELRO el 20130627, según TI-ERS019-2012

'***Agregado por ELRO el 20130715, según RFC1306270002***
Private Sub actualizarNroOperacionesDiariaServicioPago()
    oConec.ConexionActiva.BeginTrans
    oConec.CommadTimeOut = 0
    ssql = ""
    ssql = "exec stp_upd_ActualizarNroOperacionesDiariaServicioPago"
    oConec.ConexionActiva.Execute ssql
    oConec.ConexionActiva.CommitTrans
End Sub
'***Fin Agregado por ELRO el 20130627, según RFC1306270002

'RIRO20140407 ERS011 ************************************************
Private Sub CancelarAutorizacionDepositos()
    oConec.ConexionActiva.BeginTrans
    oConec.CommadTimeOut = 0
    ssql = ""
    ssql = "exec stp_del_VisBueCapIndConCli"
    oConec.ConexionActiva.Execute ssql
    oConec.ConexionActiva.CommitTrans
End Sub
'END RIRO ***********************************************************

'WIOR 20140619 ******************************************************
'Comentado por JUEZ 20150217
'Private Sub PagoAutomaticoCredito(ByVal pdFecha As Date)
'Dim oCredito As COMDCredito.DCOMCredito
'Dim oCapta As COMNCaptaGenerales.NCOMCaptaGenerales
'Dim oNCredito As COMNCredito.NCOMCredito
'
'Dim rsCredito As ADODB.Recordset
'Dim rsAhorro As ADODB.Recordset
'Dim rsCapta As ADODB.Recordset
'
'Dim i As Integer
'Dim j As Integer
'Dim lsMsgError As String
'Dim lsOpeCod As String
'Dim lsImpreBoleta As String
'
'Dim MontoAPagar As Double
'Dim MontoPagado As Double
'
'Dim MatDebitos As Variant
'Dim CantDebitos As Integer
'Dim MontoDebito As Double
'Dim MontoDebitoTC As Double
'Dim nSaldoDisp As Double
'Dim nSaldoDispTC As Double
'Dim nTC As Double
'
'Dim oTipCambio As nTipoCambio
'Set oTipCambio = New nTipoCambio
'nTC = oTipCambio.EmiteTipoCambio(gdFecSis, TCFijoMes)
'Set oTipCambio = Nothing
'
'Set oCredito = New COMDCredito.DCOMCredito
'Set rsCredito = oCredito.RecuperaCredPagarAutomatico(pdFecha)
'
'If Not (rsCredito.EOF And rsCredito.BOF) Then
'    ReDim fMatCredPagoAuto(rsCredito.RecordCount - 1)
'    For i = 0 To rsCredito.RecordCount - 1
'        fMatCredPagoAuto(i).cctacod = Trim(rsCredito!cctacod)
'        fMatCredPagoAuto(i).cTpoCredCod = Trim(rsCredito!cTpoCredCod)
'        fMatCredPagoAuto(i).nDiasAtraso = CInt(rsCredito!nDiasAtraso)
'        fMatCredPagoAuto(i).nNroCalen = CInt(rsCredito!nNroCalen)
'        fMatCredPagoAuto(i).CuotaMin = CInt(rsCredito!CuotaMin)
'        fMatCredPagoAuto(i).CuotaMax = CInt(rsCredito!CuotaMax)
'        fMatCredPagoAuto(i).Pagar = CDbl(rsCredito!Pagar)
'        fMatCredPagoAuto(i).cMetLiquidacion = Trim(rsCredito!cMetLiquidacion)
'
'        Set rsAhorro = oCredito.RecuperaCuentasAhorroPagarAutomatico(fMatCredPagoAuto(i).cctacod)
'        If Not (rsAhorro.EOF And rsAhorro.BOF) Then
'            ReDim fMatCredPagoAuto(i).AhorroDebitar(rsAhorro.RecordCount - 1)
'            For j = 0 To rsAhorro.RecordCount - 1
'                fMatCredPagoAuto(i).AhorroDebitar(j) = Trim(rsAhorro!Ahorro)
'                rsAhorro.MoveNext
'            Next j
'        Else
'            ReDim fMatCredPagoAuto(i).AhorroDebitar(0)
'        End If
'        Set rsAhorro = Nothing
'
'        rsCredito.MoveNext
'    Next i
'Else
'    ReDim fMatCredPagoAuto(0)
'End If
'
'Set rsCredito = Nothing
'
'Set oCapta = New COMNCaptaGenerales.NCOMCaptaGenerales
'Set oNCredito = New COMNCredito.NCOMCredito
'lsOpeCod = gCredPagoCuotasAuto
'
'If Trim(fMatCredPagoAuto(0).cctacod) <> "" Then
'    For i = 0 To UBound(fMatCredPagoAuto)
'        MontoAPagar = fMatCredPagoAuto(i).Pagar
'        MontoPagado = 0
'        ReDim MatDebitos(2, 0)
'
'        If Trim(fMatCredPagoAuto(i).AhorroDebitar(0)) <> "" Then
'            CantDebitos = -1
'
'            For j = 0 To UBound(fMatCredPagoAuto(i).AhorroDebitar)
'                nSaldoDisp = 0
'                nSaldoDispTC = 0
'                MontoDebito = 0
'                MontoDebitoTC = 0
'                If (fMatCredPagoAuto(i).Pagar - MontoPagado) > 0 Then
'
'                    Set rsCapta = oCapta.GetDatosCuenta(fMatCredPagoAuto(i).AhorroDebitar(j))
'
'                    If Not (rsCapta.EOF And rsCapta.BOF) Then
'                        nSaldoDisp = CDbl(rsCapta("nSaldoDisp"))
'                        nSaldoDispTC = CDbl(rsCapta("nSaldoDisp"))
'                        If Mid(fMatCredPagoAuto(i).cctacod, 9, 1) <> Mid(fMatCredPagoAuto(i).AhorroDebitar(j), 9, 1) Then
'                            If Mid(fMatCredPagoAuto(i).cctacod, 9, 1) = "1" Then
'                                nSaldoDisp = nSaldoDisp * nTC
'                            Else
'                                nSaldoDisp = nSaldoDisp / nTC
'                            End If
'                        End If
'                    End If
'
'                    Set rsCapta = Nothing
'
'                    If nSaldoDisp > 0 Then
'                        MontoDebito = IIf(nSaldoDisp > MontoAPagar, nSaldoDisp - MontoAPagar, nSaldoDisp)
'
'                    End If
'
'                    If MontoDebito > 0 Then
'                        MontoPagado = MontoPagado + Round(MontoDebito, 2)
'                        CantDebitos = CantDebitos + 1
'
'                        ReDim Preserve MatDebitos(2, 0 To CantDebitos)
'                        MatDebitos(0, CantDebitos) = fMatCredPagoAuto(i).AhorroDebitar(j)
'                        MatDebitos(1, CantDebitos) = MontoDebito
'
'                        If Mid(fMatCredPagoAuto(i).cctacod, 9, 1) <> Mid(fMatCredPagoAuto(i).AhorroDebitar(j), 9, 1) Then
'                            If Mid(fMatCredPagoAuto(i).cctacod, 9, 1) = "1" Then
'                                MontoDebitoTC = MontoDebito / nTC
'                            Else
'                                MontoDebitoTC = MontoDebito * nTC
'                            End If
'                        End If
'
'                        MatDebitos(2, CantDebitos) = IIf(MontoDebitoTC > nSaldoDispTC, nSaldoDispTC, MontoDebitoTC)
'                        MontoAPagar = MontoAPagar - MontoPagado
'                    End If
'                Else
'                    Exit For
'                End If
'            Next j
'        End If
'
'        If MontoPagado > 0 Then
'            lsMsgError = oNCredito.GrabarPagosAutomaticos(fMatCredPagoAuto(i).cctacod, gsCodUser, Right(gsCodAge, 2), gdFecSis, lsImpreBoleta, _
'                            lsOpeCod, MontoPagado, MatDebitos, (fMatCredPagoAuto(i).Pagar - MontoPagado))
'        End If
'    Next i
'End If
'End Sub
'WIOR FIN ***********************************************************
'JUEZ 20150217 *****************************************************************
Private Sub DebitoAutomaticoPagoCreditos(ByVal pdFecSis As Date)
Dim oDCapGen As COMDCaptaGenerales.DCOMCaptaGenerales
Dim oNCapMov As COMNCaptaGenerales.NCOMCaptaMovimiento
Dim oNCred As COMNCredito.NCOMCredito
Dim oITF As COMDConstSistema.FCOMITF
Dim rs As ADODB.Recordset
Dim i As Integer
Dim nITF As Double
Dim cCtaCod As String
Dim cCtaCodAho As String
Dim nMontoPago As Double
Dim nComision As Double
Dim nMontoMax As Double
Dim bCObraITF As Boolean

Dim lsMsjError As String
Dim lsImpreBoleta As String

Set oDCapGen = New COMDCaptaGenerales.DCOMCaptaGenerales
Set rs = oDCapGen.ObtenerDebitoAutomaticoCredServ(gServCredito, pdFecSis)
Set oDCapGen = Nothing
If Not (rs.EOF And rs.BOF) Then
    Set oNCapMov = New COMNCaptaGenerales.NCOMCaptaMovimiento
    For i = 0 To rs.RecordCount - 1
        cCtaCod = rs!cCtaCod
        cCtaCodAho = rs!cCtaCodAho
        nMontoPago = rs!nMontoPago
        nComision = rs!nMontoCom
        nMontoMax = rs!nMontoMax
        bCObraITF = IIf(rs!nCobraITF = 1, True, False)
        
        If bCObraITF Then
            Set oNCred = New COMNCredito.NCOMCredito
            nITF = oNCred.DameMontoITF(nMontoPago)
            Set oNCred = Nothing
        Else
            nITF = 0
        End If

        If nMontoPago <= nMontoMax Then
            If oNCapMov.ValidaSaldoCuenta(cCtaCodAho, nMontoPago + nComision + nITF) Then
                Set oNCred = New COMNCredito.NCOMCredito
                lsMsjError = oNCred.GrabarPagoCreditoDebitoAutomatico(cCtaCod, gsCodUser, gsCodAge, gdFecSis, cCtaCodAho, nMontoPago, nITF, nComision, lsImpreBoleta)
                Set oNCred = Nothing
                oNCapMov.RegistrarCaptacServDebitoAutoLog cCtaCodAho, cCtaCod, nMontoPago + nITF, nMontoMax, nComision, IIf(lsMsjError = "", True, False), IIf(lsMsjError = "", "Pago Realizado", lsMsjError), gdFecSis
            Else
                oNCapMov.RegistrarCaptacServDebitoAutoLog cCtaCodAho, cCtaCod, nMontoPago + nITF, nMontoMax, nComision, False, "Cuenta a debitar NO posee saldo suficiente o NO está ACTIVA", gdFecSis
            End If
        Else
            oNCapMov.RegistrarCaptacServDebitoAutoLog cCtaCodAho, cCtaCod, nMontoPago + nITF, nMontoMax, nComision, False, "El Monto a Pagar excede del Monto Máximo ingresado en el registro", gdFecSis
        End If
        rs.MoveNext
    Next i
    Set oNCapMov = Nothing
End If
End Sub
Private Sub DebitoAutomaticoPagoServicios(ByVal pdFecSis As Date) 'Forma de pago según el módulo de Pago de Servicio de Recaudo implementado por RIRO

Dim oDCapGen As COMDCaptaGenerales.DCOMCaptaGenerales
Dim oDCapServ As COMDCaptaServicios.DCOMServicioRecaudo
Dim oNCapMov As COMNCaptaGenerales.NCOMCaptaMovimiento
Dim ClsMov As COMNContabilidad.NCOMContFunciones
Dim rsDeb As ADODB.Recordset
Dim rsRecaudo As ADODB.Recordset
Dim cCtaCodAho As String
Dim cCodConvenio As String
Dim cCodCliente As String
Dim nTipoDOI As Integer
Dim cDOI As String
Dim cCtaCodEmp As String
Dim nMontoMax As Double, nMontoPago As Double
Dim nComiDeb As Double, nComiEmp As Double, nComiCli As Double
Dim sNomCliente As String 'JUEZ 20160202
Dim bObtieneComision As Boolean
Dim sMovNroDeb As String, sMovNroPago As String
Dim nI As Integer, nJ As Integer
Dim sCad As String
Dim lsMsjError As String
Dim vMatCobro() As String
Dim vMatConceptos() As String

Dim sBoleta As String
Dim nITF As Double
Dim nRedondeoITF As Double

Set oDCapGen = New COMDCaptaGenerales.DCOMCaptaGenerales
Set rsDeb = oDCapGen.ObtenerDebitoAutomaticoCredServ(gServConvenio, pdFecSis)
Set oDCapGen = Nothing

If Not (rsDeb.EOF And rsDeb.BOF) Then
    For nI = 0 To rsDeb.RecordCount - 1
        cCtaCodAho = rsDeb!cCtaCodAho
        cCodCliente = rsDeb!cCodCliente
        cCodConvenio = rsDeb!cCodConvenio
        nMontoMax = rsDeb!nMontoMax
        nComiDeb = rsDeb!nMontoCom
        
        Set oDCapServ = New COMDCaptaServicios.DCOMServicioRecaudo
        Set rsRecaudo = oDCapServ.getBuscarUsuarioRecaudo(, , cCodCliente, cCodConvenio, 1)
        cCtaCodEmp = oDCapServ.getBuscaConvenioXCodigo(cCodConvenio)!cCtaCod
        nMontoPago = 0
        ReDim vMatCobro(0)
        ReDim vMatConceptos(3, 0)
        nJ = 0
        
        Do While Not rsRecaudo.EOF
            cDOI = rsRecaudo!cDOI
            nTipoDOI = rsRecaudo!nTipoDOI
            sNomCliente = rsRecaudo!cNomCliente
            nMontoPago = nMontoPago + CDbl(rsRecaudo!nImporte) + CDbl(rsRecaudo!nMora)
            
            nJ = nJ + 1
            ReDim Preserve vMatCobro(nJ)
            sCad = rsRecaudo!cId & "|" 'ID de la Trama
            sCad = sCad & Replace(IIf(Len(Trim(rsRecaudo!cServicio)) = 0, Space(200) & ".", Trim(rsRecaudo!cServicio)), ".", "") & "|" 'Servicio
            sCad = sCad & rsRecaudo!cConcepto & "|" 'Concepto
            sCad = sCad & CDbl(rsRecaudo!nImporte) & "|" 'Importe
            sCad = sCad & "0.00" & "|" 'Deuda Actual
            sCad = sCad & CDbl(rsRecaudo!nImporte) & "|" 'Monto Cobro
            sCad = sCad & Pagado & "|" 'Estado
            sCad = sCad & CDbl(rsRecaudo!nMora) & "|" 'Mora
            sCad = sCad & Format(CDate(rsRecaudo!dFechaVencimiento), "yyyyMMdd") & "|" 'Fecha Vencimiento

            vMatCobro(nJ - 1) = sCad
            
            ReDim Preserve vMatConceptos(3, nJ)
            vMatConceptos(1, nJ) = nJ
            vMatConceptos(2, nJ) = CDbl(rsRecaudo!nImporte) + CDbl(rsRecaudo!nMora)
            vMatConceptos(3, nJ) = 0
            rsRecaudo.MoveNext
        Loop
        Set rsRecaudo = Nothing
        
        If nMontoPago > 0 Then
            nITF = Format(fgITFCalculaImpuesto(nMontoPago), "#,##0.00")
            nRedondeoITF = fgDiferenciaRedondeoITF(CCur(nITF))
            If nRedondeoITF > 0 Then
               nITF = Format(CCur(nITF) - nRedondeoITF, "#,##0.00")
            End If
            
            Set oNCapMov = New COMNCaptaGenerales.NCOMCaptaMovimiento
            If nMontoPago <= nMontoMax Then
                bObtieneComision = oDCapServ.CalculaComisionRecaudo(cCodConvenio, cCodCliente, vMatConceptos, nComiCli, nComiEmp)
            
                If bObtieneComision Then
                    If oNCapMov.ValidaSaldoCuenta(cCtaCodAho, nMontoPago + nComiDeb + nITF + nComiEmp + nComiCli) Then
                        Set ClsMov = New COMNContabilidad.NCOMContFunciones
                        sMovNroDeb = ClsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
                        sMovNroPago = Left(sMovNroDeb, 19) + "01" + Right(sMovNroDeb, 4)
                        Set ClsMov = Nothing
                        Set oNCapMov = New COMNCaptaGenerales.NCOMCaptaMovimiento
                        'lsMsjError = oNCapMov.GrabarPagoServicioDebitoAutomatico(cCtaCodAho, nMontoPago, nComiDeb, vMatCobro, cCodConvenio, cCodCliente, nTipoDOI, cDOI, "", nComiEmp, nComiCli, cCtaCodEmp, nITF, gdFecSis, sMovNroDeb, sMovNroPago, gbImpTMU, sBoleta)
                        lsMsjError = oNCapMov.GrabarPagoServicioDebitoAutomatico(cCtaCodAho, nMontoPago, nComiDeb, vMatCobro, cCodConvenio, cCodCliente, nTipoDOI, cDOI, sNomCliente, nComiEmp, nComiCli, cCtaCodEmp, nITF, gdFecSis, sMovNroDeb, sMovNroPago, gbImpTMU, sBoleta) 'JUEZ 20160202
                        oNCapMov.RegistrarCaptacServDebitoAutoLog cCtaCodAho, cCodConvenio & "-" & cCodCliente, nMontoPago + nITF + nComiEmp + nComiCli, nMontoMax, nComiDeb, IIf(lsMsjError = "", True, False), IIf(lsMsjError = "", "Pago Realizado", lsMsjError), gdFecSis
                    Else
                        oNCapMov.RegistrarCaptacServDebitoAutoLog cCtaCodAho, cCodConvenio & "-" & cCodCliente, nMontoPago + nITF + nComiEmp + nComiCli, nMontoMax, nComiDeb, False, "Cuenta a debitar NO posee saldo suficiente o NO está ACTIVA", gdFecSis
                    End If
                Else
                    oNCapMov.RegistrarCaptacServDebitoAutoLog cCtaCodAho, cCodConvenio & "-" & cCodCliente, nMontoPago + nITF + nComiEmp + nComiCli, nMontoMax, nComiDeb, False, "No se pudieron obtener las comisiones del convenio y cliente", gdFecSis
                End If
            Else
                oNCapMov.RegistrarCaptacServDebitoAutoLog cCtaCodAho, cCodConvenio & "-" & cCodCliente, nMontoPago + nITF + nComiEmp + nComiCli, nMontoMax, nComiDeb, False, "El Monto a Pagar excede del Monto Máximo ingresado en el registro", gdFecSis
            End If
            
            Set oNCapMov = Nothing

        End If
        rsDeb.MoveNext
    Next nI
    Set oDCapServ = Nothing
End If
End Sub
'END JUEZ **********************************************************************
'EJVG20151118 ***
Private Sub GuardarGarantiasDiario()
    Dim ssql As String
    
    oConec.CommadTimeOut = 0
    ssql = "EXEC stp_upd_ERS0632014_GarantiaAVAL"
    oConec.ConexionActiva.Execute ssql
    
    oConec.CommadTimeOut = 0
    ssql = "EXEC sp_GeneraGarantiasDiario '" & Format(gdFecSis, "yyyymmdd") & "'"
    oConec.ConexionActiva.Execute ssql
End Sub
Private Sub GeneraMovAsientoGarantiasDiario()
    Dim ssql As String
    
    oConec.CommadTimeOut = 0
    ssql = "EXEC sp_GeneraMovAsientoGarantiasDiario '" & Format(gdFecSis & " " & GetHoraServer(), "yyyymmdd hh:mm:ss") & "', '" & gsCodUser & "', '" & Right(gsCodAge, 2) & "'"
    oConec.ConexionActiva.Execute ssql
End Sub
Private Sub CierreGarantias()
    Dim ssql As String
    
    oConec.CommadTimeOut = 0
    ssql = "EXEC stp_upd_ERS0632014_ConsolidaGarantia NULL,'" & Format(DateAdd("D", 1, gdFecSis), "yyyymmdd") & "'"
    oConec.ConexionActiva.Execute ssql
End Sub
'END EJVG *******
