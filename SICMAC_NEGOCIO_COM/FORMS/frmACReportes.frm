VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmACReportes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ahorros : Listados y Reportes"
   ClientHeight    =   6885
   ClientLeft      =   2700
   ClientTop       =   1590
   ClientWidth     =   6870
   Icon            =   "frmACReportes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraFecha 
      Caption         =   "Fechas"
      Height          =   630
      Left            =   60
      TabIndex        =   4
      Top             =   0
      Width           =   3765
      Begin MSMask.MaskEdBox txtFechaF 
         Height          =   300
         Left            =   2265
         TabIndex        =   5
         Top             =   225
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFecha 
         Height          =   300
         Left            =   480
         TabIndex        =   6
         Top             =   225
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Al:"
         Height          =   195
         Left            =   2010
         TabIndex        =   8
         Top             =   278
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Del:"
         Height          =   195
         Left            =   90
         TabIndex        =   7
         Top             =   278
         Width           =   285
      End
   End
   Begin MSComDlg.CommonDialog dlgGrabar 
      Left            =   960
      Top             =   6360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.TreeView treeRep 
      Height          =   5655
      Left            =   60
      TabIndex        =   0
      Top             =   750
      Width           =   6765
      _ExtentX        =   11933
      _ExtentY        =   9975
      _Version        =   393217
      HideSelection   =   0   'False
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   1
      Checkboxes      =   -1  'True
      HotTracking     =   -1  'True
      ImageList       =   "Img"
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CheckBox chkCondensado 
      Caption         =   "Condensado"
      Height          =   345
      Left            =   5520
      TabIndex        =   3
      Top             =   143
      Value           =   1  'Checked
      Width           =   1245
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5625
      TabIndex        =   2
      Top             =   6480
      Width           =   1185
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4350
      TabIndex        =   1
      Top             =   6480
      Width           =   1185
   End
   Begin MSComctlLib.ImageList Img 
      Left            =   60
      Top             =   6240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   17
      ImageHeight     =   17
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmACReportes.frx":030A
            Key             =   "Close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmACReportes.frx":06D0
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmACReportes.frx":0B62
            Key             =   "Hoja"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmACReportes.frx":10A4
            Key             =   "Hoja1"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmACReportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Centralizacion
'Usuario : NSSE
'fecha : 26/12/2000

Option Explicit
Option Base 1
Dim Flag As Boolean
Dim flag1 As Boolean
Dim Char12 As Boolean
Dim lbPrtCom As Boolean
Dim NumA As String
Dim SalA As String
Dim NumC As String
Dim NumP As String
Dim SalP As String
Dim SalC As String
Dim TB As Currency
Dim TBD As Currency
Dim ca As Currency
Dim CAE As Currency
Dim RegCmacS As Integer
Dim vBuffer As String
Dim lsCadena As String

Function CabCheques(pPag As Integer, psTitulo As String) As String
    Dim lsCad As String
    lsCad = lsCad & gsNomAge & Space(50) & ArmaFecha(gdFecSis) & "  -  " & Str(Time()) & Chr(10)
    lsCad = lsCad & Space(30) & psTitulo & Space(10) & "Pag :  " & Str(pPag) & Chr(10)
    lsCad = lsCad & String(130, "=") & Chr(10)
    lsCad = lsCad & "   CUENTA   " & Space(2) & " NUMCHQ " & Space(2) & "BANCO                         "
    lsCad = lsCad & Space(2) & "MONTO          " & Space(2) & "FECHA REG." & Space(2)
    lsCad = lsCad & "FECHA VAL." & Space(2) & "USUA" & Space(2) & "USUR" & Space(2)
    lsCad = lsCad & "AGENCIA" & Chr(10)
    lsCad = lsCad & String(130, "=") & Chr(10)
    CabCheques = lsCad
End Function

Private Function SaldosCmact(ByVal Moneda As String, ByVal dFec As Date) As String
Dim sql As String
Dim rs As New ADODB.Recordset
Dim bHoy As Boolean

bHoy = False
If DateDiff("d", dFec, gdFecSis) = 0 Then bHoy = True

If bHoy Then
    sql = "SELECT SUM(nNum) nNum, SUM(nSaldo) nSaldo From ( " _
        & "SELECT IsNull(Count(cCodCta),0) nNum, IsNull(SUM(nSaldCntAC),0) nSaldo FROM AhorroC WHERE " _
        & "cPersoneria = '6' And cEstCtaAC not in ('C','U') and Substring(cCodCta, 6,1) = '" & Moneda & "' " _
        & "UNION " _
        & "SELECT IsNull(Count(cCodCta),0) nNum, IsNull(SUM(nSaldCntPF),0) nSaldo FROM PlazoFijo WHERE " _
        & "cPersoneria = '6' And cEstCtaPF not in ('C','U') and SUBSTRING(cCodCta, 6,1) = '" & Moneda & "' " _
        & ") T"
Else
    sql = "SELECT SUM(nNum) nNum, SUM(nSaldo) nSaldo From ( " _
        & "SELECT nNumCMAC nNum, nSaldCMAC nSaldo FROM EstadDiaAC WHERE " _
        & "DateDiff(dd,dEstadAC,'" & Format$(dFec, gsFormatoFecha) & "') = 0 " _
        & "And cMoneda = '" & Moneda & "' " _
        & "UNION " _
        & "SELECT nNumCMAC nNum, nSaldCMAC nSaldo FROM EstadDiaPF WHERE " _
        & "DateDiff(dd,dEstadPF,'" & Format$(dFec, gsFormatoFecha) & "') = 0 " _
        & "And cMoneda = '" & Moneda & "' " _
        & ") T"
End If
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
'rs.Open sql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rs.ActiveConnection = Nothing
If rs.EOF And rs.BOF Then
    RegCmacS = 0
    SaldosCmact = "0"
Else
    RegCmacS = IIf(IsNull(rs!nNum), 0, rs!nNum)
    SaldosCmact = IIf(IsNull(rs!nSaldo), "0", Str(rs!nSaldo))
End If
rs.Close
Set rs = Nothing
End Function

Private Function BL(ByVal Moneda As String, ByVal ValMon As String, ByVal TipMon As String, ByVal dFec As Date) As String
Dim sql As String
Dim rs As ADODB.Recordset

sql = "SELECT IsNull(SUM(nCantidad),0) Monto FROM Billetaje " _
    & "WHERE cMoneda = '" & Moneda & "' and nMoneda = " & ValMon & " " _
    & "AND cTipMoneda = '" & TipMon & "' AND DateDiff(dd,dFecha,'" & Format(dFec, gsFormatoFecha) & "') = 0"

Set rs = New ADODB.Recordset
'rs.Open sql, dbCmact, adOpenForwardOnly, adLockReadOnly, adCmdText
If rs.EOF And rs.BOF Then
  BL = "0"
  Select Case Moneda
       Case "1"
         ca = 0
       Case "2"
         CAE = 0
   End Select
  
Else
  Select Case Moneda
     Case "1": ca = rs!Monto
        BL = CCur(rs!Monto * Val(ValMon))
        TB = TB + BL
     Case "2": CAE = rs!Monto
        BL = CCur(rs!Monto * Val(ValMon))
        TBD = TBD + BL
  End Select
End If
rs.Close
Set rs = Nothing
End Function

Private Sub DC(ByVal TipCre As String, ByVal Fondo As String, ByVal Moneda As String, ByVal dFec As Date)
Dim sql As String
Dim rs As ADODB.Recordset
Dim bHoy As Boolean

bHoy = False
If DateDiff("d", dFec, gdFecSis) = 0 Then bHoy = True

Select Case TipCre
    Case "2" ' CPE
        If bHoy Then
            sql = "SELECT IsNull(COUNT(cCodCta),0) nNum, IsNull(SUM(nSaldoCap),0) nSaldo FROM Credito " _
                & "WHERE substring(cCodLinCred,1,1) IN ('" & TipCre & "','1') AND " _
                & "substring(cCodLinCred,3,1) = '1' AND " _
                & "substring(cCodLinCred,6,1) IN ('" & Fondo & "') AND " _
                & "SUBSTRING(cCodLinCred,4,1) = '" & Moneda & "' " _
                & "AND cEstado = 'F'"
        Else
            sql = "SELECT IsNull(SUM(nNumSaldos),0) nNum, IsNull(SUM(nSaldoCap),0) nSaldo FROM EstadDiaCred " _
                & "WHERE substring(cCodLinCred,1,1) IN ('" & TipCre & "','1') AND " _
                & "substring(cCodLinCred,3,1) = '1' AND " _
                & "substring(cCodLinCred,6,1) IN ('" & Fondo & "') AND " _
                & "SUBSTRING(cCodLinCred,4,1) = '" & Moneda & "' " _
                & "AND DateDiff(dd,dFecha,'" & Format$(dFec, gsFormatoFecha) & "') = 0"
        End If
    Case "3" 'PP
        If bHoy Then
            sql = "SELECT IsNull(COUNT(cCodCta),0) nNum, IsNull(SUM(nSaldoCap),0) nSaldo FROM Credito " _
                & "WHERE substring(cCodLinCred,1,1) = '" & TipCre & "' AND " _
                & "substring(cCodLinCred,2,2) = '" & Fondo & "' AND " _
                & "SUBSTRING(cCodLinCred,4,1) = '" & Moneda & "' " _
                & "AND cEstado = 'F'"
        Else
            sql = "SELECT IsNull(SUM(nNumSaldos),0) nNum, IsNull(SUM(nSaldoCap),0) nSaldo FROM EstadDiaCred " _
                & "WHERE substring(cCodLinCred,1,1) = '" & TipCre & "' AND " _
                & "substring(cCodLinCred,2,2) = '" & Fondo & "' AND " _
                & "SUBSTRING(cCodLinCred,4,1) = '" & Moneda & "' " _
                & "AND DateDiff(dd,dFecha,'" & Format$(dFec, gsFormatoFecha) & "') = 0"
        End If
    Case "A" ' AGRIC
        If bHoy Then
            sql = "SELECT IsNull(COUNT(cCodCta),0) nNum, IsNull(SUM(nSaldoCap),0) nSaldo FROM Credito " _
                & "WHERE substring(cCodLinCred,1,1) IN ('1','2') AND " _
                & "substring(cCodLinCred,3,1) = '2' AND " _
                & "SUBSTRING(cCodLinCred,4,1) = '" & Moneda & "' " _
                & "AND cEstado = 'F'"
        Else
            sql = "SELECT IsNull(SUM(nNumSaldos),0) nNum, IsNull(SUM(nSaldoCap),0) nSaldo FROM EstadDiaCred " _
                & "WHERE substring(cCodLinCred,1,1) IN ('1','2') AND " _
                & "substring(cCodLinCred,3,1) = '2' AND " _
                & "SUBSTRING(cCodLinCred,4,1) = '" & Moneda & "' " _
                & "AND DateDiff(dd,dFecha,'" & Format$(dFec, gsFormatoFecha) & "') = 0"
        End If
End Select
Set rs = New ADODB.Recordset
'rs.Open sql, dbCmact, adOpenForwardOnly, adLockReadOnly, adCmdText
If rs.EOF And rs.BOF Then
    NumC = "0"
    SalC = "0"
Else
    NumC = Str(rs!nNum)
    SalC = Str(rs!nSaldo)
End If
rs.Close
Set rs = Nothing
End Sub

Private Sub DA(ByVal Producto As String, ByVal Moneda As String, ByVal dFec As Date, Optional chqProd As String = "1")
Dim sql As String
Dim bHoy As Boolean
Dim rs As ADODB.Recordset

bHoy = False
If DateDiff("d", dFec, gdFecSis) = 0 Then bHoy = True

Select Case Producto
    Case gsCodProAC
        If bHoy Then
            sql = "SELECT IsNull(COUNT(cCodCta),0) nNum, IsNull(SUM(nSaldCntAC),0) nSaldo " _
                & "FROM AHORROC WHERE cEstCtaAC NOT IN ('C','U') AND SUBSTRING(cCodCta,6,1)= '" & Moneda & "'"
        Else
            sql = "Select nCtaVigAC nNum, (nMonChqVal+nSaldoAC) nSaldo FROM " _
                & "EstadDiaAC WHERE DateDiff(dd,dEstadAC,'" & Format$(dFec, gsFormatoFecha) & "') = 0 " _
                & "And cMoneda = '" & Moneda & "'"
        End If
   Case gsCodProPF
        If bHoy Then
            sql = "SELECT IsNull(SUM(nNum),0) nNum, IsNull(SUM(nSaldo),0) nSaldo FROM " _
                & "(SELECT COUNT(cCodCta) nNum, SUM(nSaldCntPF) nSaldo " _
                & "FROM PLAZOFIJO WHERE cEstCtaPF NOT IN ('C','U')AND substring(cCodCta,6,1)= '" & Moneda & "' " _
                & "UNION " _
                & "SELECT COUNT(cCodCta) nNum, SUM(nSaldCntCTS) nSaldo " _
                & "FROM CTS WHERE cEstCtaCTS NOT IN ('C','U') AND substring(cCodCta,6,1)= '" & Moneda & "') T"
        Else
            sql = "SELECT IsNull(SUM(nNum),0) nNum, IsNull(SUM(nSaldo),0) nSaldo FROM " _
                & "(Select nNumVigPF nNum, (nSaldoPF+nMonChqVal) nSaldo FROM EstadDiaPF " _
                & "WHERE DateDiff(dd,dEstadPF,'" & Format$(dFec, gsFormatoFecha) & "') = 0 And cMoneda = '" & Moneda & "' " _
                & "UNION " _
                & "Select nNumVigCTS nNum, (nSaldoCTS+nMonChqVal) nSaldo FROM EstadDiaCTS " _
                & "WHERE DateDiff(dd,dEstadCTS,'" & Format$(dFec, gsFormatoFecha) & "') = 0 And cMoneda = '" & Moneda & "') T"
        End If
   Case "Cheques"
         Select Case chqProd
             Case gsCodProAC
                If bHoy Then
                    sql = "SELECT IsNull(COUNT(cCodCta),0) nNum, IsNull(sum(nMontoChq),0) nSaldo " _
                        & "FROM CHEQUE WHERE cEstChq = 'E' AND substring(cCodCta,6,1)= '" & Moneda & "' " _
                        & "and substring(cCodCta,3,3) = '" & gsCodProAC & "' AND substring(cCodCta,1,2) = '" & Mid(gsCodAge, 4, 2) & "'"
                Else
                    sql = "SELECT nNumChqVal nNum, nMonChqVal nSaldo FROM EstadDiaAC WHERE " _
                        & "DateDiff(dd,dEstadAC,'" & Format$(dFec, gsFormatoFecha) & "') = 0 " _
                        & "And cMoneda = '" & Moneda & "'"
                End If
             Case gsCodProPF
                If bHoy Then
                    sql = "SELECT IsNull(COUNT(cCodCta),0) nNum, IsNull(sum(nMontoChq),0) nSaldo " _
                        & "FROM CHEQUE WHERE cEstChq = 'E' AND substring(cCodCta,6,1)= '" & Moneda & "' " _
                        & "and substring(cCodCta,3,3) = '" & gsCodProPF & "' AND substring(cCodCta,1,2) = '" & Mid(gsCodAge, 4, 2) & "'"
                Else
                    sql = "SELECT nNumChqVal nNum, nMonChqVal nSaldo FROM EstadDiaPF WHERE " _
                        & "DateDiff(dd,dEstadPF,'" & Format$(dFec, gsFormatoFecha) & "') = 0 " _
                        & "And cMoneda = '" & Moneda & "'"
                End If
             Case gsCodProCTS
                If bHoy Then
                    sql = "SELECT IsNull(COUNT(cCodCta),0) nNum, IsNull(sum(nMontoChq),0) nSaldo " _
                        & "FROM CHEQUE WHERE cEstChq = 'E' AND substring(cCodCta,6,1)= '" & Moneda & "' " _
                        & "and substring(cCodCta,3,3) = '" & gsCodProCTS & "' AND substring(cCodCta,1,2) = '" & Mid(gsCodAge, 4, 2) & "'"
                Else
                    sql = "SELECT nNumChqVal nNum, nMonChqVal nSaldo FROM EstadDiaCTS WHERE " _
                        & "DateDiff(dd,dEstadCTS,'" & Format$(dFec, gsFormatoFecha) & "') = 0 " _
                        & "And cMoneda = '" & Moneda & "'"
                End If
          End Select
End Select
Set rs = New ADODB.Recordset
'rs.Open sql, dbCmact, adOpenForwardOnly, adLockReadOnly, adCmdText
If rs.EOF And rs.BOF Then
    NumA = "0"
    SalA = "0"
Else
    NumA = IIf(IsNull(rs!nNum), "0", Str(rs!nNum))
    SalA = IIf(IsNull(rs!nSaldo), "0", Str(rs!nSaldo))
End If
rs.Close
Set rs = Nothing
End Sub


Private Function DP(ByVal Estado As String, ByVal dFec As Date) As String
Dim sql As String
Dim rs As ADODB.Recordset
Dim bHoy As Boolean

bHoy = False
If DateDiff("d", dFec, gdFecSis) = 0 Then bHoy = True

Select Case Estado
    Case "V"
        If bHoy Then
            sql = "SELECT IsNull(count(cCodCta),0) nNum, IsNull(sum(nOroNeto),0) nOro, IsNull(sum(nSaldoCap),0) nSaldo " _
                & "FROM CredPrenda WHERE cEstado IN ('1','4','7','6')"
        Else
            sql = "Select nNumCredVig nNum, nOroVig nOro, nCapVig nSaldo FROM EstadDiaPrenda " _
                & "WHERE DateDiff(dd,dFecha,'" & Format$(dFec, gsFormatoFecha) & "') = 0"
        End If
    Case "A"
        If bHoy Then
            sql = " SELECT IsNull(count(cCodCta),0) nNum, IsNull(sum(nOroNeto),0) nOro, IsNull(sum(nSaldoCap),0) nSaldo " _
                & "FROM CREDPRENDA WHERE cEstado IN ('8')"
        Else
            sql = "Select nNumCredAdj nNum, nOroAdj nOro, nCapAdj nSaldo FROM EstadDiaPrenda " _
                & "WHERE DateDiff(dd,dFecha,'" & Format$(dFec, gsFormatoFecha) & "') = 0"
        End If
    Case "D"
        If bHoy Then
            sql = "SELECT IsNull(count(cCodCta),0) nNum, IsNull(sum(nOroNeto),0) nOro, IsNull(sum(nSaldoCap),0) nSaldo " _
                & "FROM CREDPRENDA WHERE cEstado IN ('2')"
        Else
            sql = "Select nNumCredDif nNum, nOroDif nOro, nSaldo = 0 FROM EstadDiaPrenda " _
                & "WHERE DateDiff(dd,dFecha,'" & Format$(dFec, gsFormatoFecha) & "') = 0"
        End If
End Select
Set rs = New ADODB.Recordset
'rs.Open sql, dbCmact, adOpenForwardOnly, adLockReadOnly, adCmdText
If rs.EOF And rs.BOF Then
    DP = "0"
    NumP = "0"
    SalP = "0"
Else
    DP = Str(rs!nNum)
    NumP = Str(rs!nOro)
    SalP = Str(rs!nSaldo)
End If
rs.Close
Set rs = Nothing
End Function

Private Function CuentasInactivas(Moneda As String) As String
Dim sql As String
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim SQL1 As String
Dim CuentaA As String
Dim vLineas As Integer
Dim vPag As Integer
Dim TDesc As Currency
Dim RCapital As Currency
Dim lsCad As String
     
sql = " SELECT A.cCodCta, A.nSaldDispAC,A.nSaldAntAC,T.nMonTran,A.dUltCntAC" _
     & " FROM AhorroC A,Trandiaria T WHERE A.cCodCta = T.cCodCta" _
     & " and T.cCodOpe in ('  gsACCanIna ',' gsACDesInaAct & ')" _
     & " and substring(A.cCodCta,6,1) = '" & Moneda & "' And (cFlag is null or cFlag <> 'X')"
'rs.Open sql, dbCmact, adOpenForwardOnly, adLockReadOnly, adCmdText
If RSVacio(rs) Then
Else
 vPag = 1
 
 lsCad = lsCad + Chr(10)
 lsCad = lsCad + gsNomAge + Space(5) + "Sección Ahorro" + Space(5) + "Trujillo " + ArmaFecha(gdFecSis) + Space(4) + "Pagina: " + Format(vPag, "####") + Chr(10) + Chr(10)
 lsCad = lsCad + Space(20) + "LISTADO DE CUENTAS INACTIVAS EN " + IIf(Moneda = "1", "SOLES", "DOLARES") + Chr(10) + Chr(10)
 lsCad = lsCad + "CUENTA " + Space(5) + "CUENTAANT" + Space(5) + "S.ANTERIOR" + Space(5) + "S.DISPONIBLE" + Space(3) + "R.CAPITAL" + Space(2) + "F.ULTCONTACTO" + Space(5) + Chr(10)
 lsCad = lsCad + String(90, "=") + Chr(10)
 Do While Not rs.EOF
        SQL1 = " SELECT cCodCtaAnt from Relcuentas where cCodCta = '" & rs!cCodCta & "'"
''        rs1.Open SQL1, dbCmact, adOpenForwardOnly, adLockReadOnly, adCmdText
        If RSVacio(rs1) Then
            CuentaA = "        "
        Else
            CuentaA = rs1!ccodctaant
        End If
        rs1.Close
        Set rs1 = Nothing
        RCapital = Abs(rs!nMonTran)
        
 'lsCad = lsCad + rs!cCodCta + Space(2) + CuentaA + JDNum(rs!nsaldantac, 13, True, 11, 2) + Space(2)
 'lsCad = lsCad + JDNum(rs!nsalddispac, 13, True, 11, 2) + Space(4) + JDNum(Str(RCapital), 6, True, 4, 2) + Space(8) + Format(rs!dUltCntAC, gsformatofechaview) + Chr(10)
  TDesc = TDesc + RCapital
  
  rs.MoveNext
  vLineas = vLineas + 1
  If vLineas > 52 Then
    vPag = vPag + 1
   'lsCad = lsCad + Chr(10) + "Resumen :" + Space(38) + JDNum(Str(TDesc), 13, True, 11, 2) + Chr(10)
   lsCad = lsCad + Chr(12)
   lsCad = lsCad + Chr(10)
   lsCad = lsCad + gsNomAge + Space(5) + "Sección Ahorro" + Space(5) + "Trujillo " + ArmaFecha(gdFecSis) + Space(4) + "Pagina: " + Format(vPag, "####") + Chr(10) + Chr(10)
   lsCad = lsCad + Space(20) + "LISTADO DE CUENTAS INACTIVAS EN " + IIf(Moneda = "1", "SOLES", "DOLARES") + Chr(10) + Chr(10)
   lsCad = lsCad + "CUENTA " + Space(5) + "CUENTAANT" + Space(5) + "S.ANTERIOR" + Space(5) + "S.DISPONIBLE" + Space(3) + "R.CAPITAL" + Space(2) + "F.ULTCONTACTO" + Space(5) + Chr(10)
   lsCad = lsCad + String(90, "=") + Chr(10)
    vLineas = 1
  End If
 Loop
   'lsCad = lsCad + Chr(10) + "Total Resumen :" + Space(32) + JDNum(Str(TDesc), 13, True, 11, 2) + Chr(10)
 
End If
rs.Close
Set rs = Nothing
CuentasInactivas = IIf(lsCad <> "", lsCad & Chr(12), "")
End Function

Private Sub ValorSolesDolares(Moneda As String, Estado As String, Mensaje As String, Titulo As String)
'    Dim sql As String
'    Dim fecha1 As String * 10
'    Dim fecha2 As String * 10
'    Dim vCuenta As String
'    Dim Monto As Currency
'    Dim Cuenta As String * 12
'    Dim nCheque As String * 8
'    Dim CtaBanco As String * 15
'    Dim CBanco As String * 25
'    Dim vLineas As Integer
'    Dim rs As New ADODB.Recordset
'    Dim Pag As String
'    sql = " select cCodCta,cNumChq,cCtaBco,cCodBco,nMontoChq,dRegChq,dValorChq" _
'        & " from cheque where cEstChq = '" & Estado & "' and substring(cCodCta,6,1) = '" & Moneda & "'"
'
'    If txtFechaF = "__/__/____" Then
'       MsgBox "Por favor Ingrese Fecha Para Busqueda de Cheques", vbInformation, "Aviso"
'       txtFechaF.SetFocus
'       Exit Sub
'    End If
'    If Estado <> "E" Then
'        sql = sql + " and  dRegChq between '" & Format(TXTFECHA, gsformatofecha) & "' and '" & Format(CDate(TXTFECHAF) + 1, gsformatofecha) & "'"
'    End If
''    rs.Open sql, dbCmact, adOpenStatic, adLockOptimistic, adCmdText
'    If RSVacio(rs) Then
'       MsgBox Mensaje
'       Char12 = False
'    Else
'       Char12 = True
'       vLineas = 7
'       Pag = "1"
'       With lscadena
'            lscadena = lscadena + Chr(10)
'            lscadena = lscadena + Chr$(10) + Chr$(10) + gsNomAge + Space(25) + "Sección AHORRO" + Space(10) + TXTFECHA + "  a  " + TXTFECHAF + Space(20) + "Pagina: " + Pag + Chr(10) + Chr(10)
'            lscadena = lscadena + Space(20) + Titulo + Chr$(10)
'            lscadena = lscadena + Space(35) + "C U E N T A S   D E   A H O R R O " + Chr$(10)
'            lscadena = lscadena + Chr$(10)
'            lscadena = lscadena + String(125, "=") + Chr(10)
'            lscadena = lscadena + "No DE" + Space(10) + "No DE CHEQUE" + Space(5) + "CUENTA DE CHEQUE" + Space(5) + "BANCO" + Space(30) + "MONTO" + Space(5) + "F.ENTREGA" + Space(5) + "F.VALORIZADOS" + Chr$(10)
'            lscadena = lscadena + "CUENTA" + Chr$(10)
'            lscadena = lscadena + String(125, "=") + Chr(10)
'
'       End With
'       Do While Not rs.EOF
'            nCheque = rs!cNumChq
'            CtaBanco = rs!cCtaBco
'            CBanco = DameNomBanco(rs!cCodBco)
'            fecha1 = Format(rs!dRegChq, gsformatofechaview)
'            fecha2 = Format(rs!dValorChq, gsformatofechaview)
'            With rs
'            lscadena = lscadena & !cCodCta & Space(7) & Space(8 - Len(Trim(nCheque)))
'            lscadena = lscadena & Trim(nCheque) & Space(6) & Space(15 - Len(Trim(CtaBanco)))
'            lscadena = lscadena & Trim(CtaBanco) & Space(5) & CBanco & Space(2)
'            lscadena = lscadena & Space(13 - Len(Format(!nMontoChq, "#,##0.00")))
'            lscadena = lscadena & Format(!nMontoChq, "#,##0.00")
'            lscadena = lscadena & Space(4) & fecha1 & Space(8)
'            lscadena = lscadena & fecha2 & Chr$(10)
'                vLineas = vLineas + 1
'                 If vLineas >= 55 Then
'                    lscadena = lscadena & Chr(12)
'                    Pag = Str(Val(Pag) + 1)
'                    vLineas = 7
'                    Call CABECERA1(gsNomAge, TXTFECHA + "  a  " + TXTFECHAF, Pag, Titulo)
'                End If
'                 Monto = Monto + !nMontoChq
'            End With
'            rs.MoveNext
'       Loop
'       lscadena = lscadena & Chr$(10) & Chr$(10)
'       lscadena = lscadena + String(125, "=") + Chr(10)
'       lscadena = lscadena & "RESUMEN TOTAL CHEQUES" & Space(59) & Space(13 - Len(Format(Monto, "#,##0.00"))) & Format(Monto, "#,##0.00") + Chr(10)
'    End If
End Sub
Private Sub CABECERA1(oficina As String, fFecha As String, Pag As String, Titulo As String)
    Dim i As Integer
    lsCadena = lsCadena + Chr(10)
    lsCadena = lsCadena + Chr$(10) + Chr$(10) + gsNomAge + Space(25) + "Sección AHORRO" + Space(10) + txtFecha + "  a  " + txtFechaF + Space(20) + "Pagina: " + Pag + Chr(10) + Chr(10)
    lsCadena = lsCadena + Space(20) + Titulo + Chr$(10)
    lsCadena = lsCadena + Space(35) + "C U E N T A S   D E   A H O R R O " + Chr$(10)
    lsCadena = lsCadena + Chr$(10)
    lsCadena = lsCadena + String(125, "=") + Chr(10)
    lsCadena = lsCadena + "No DE" + Space(10) + "No DE CHEQUE" + Space(5) + "CUENTA DE CHEQUE" + Space(5) + "BANCO" + Space(30) + "MONTO" + Space(5) + "F.ENTREGA" + Space(5) + "F.VALORIZADOS" + Chr$(10)
    lsCadena = lsCadena + "CUENTA" + Chr$(10)
    lsCadena = lsCadena + String(125, "=") + Chr(10)
End Sub
Private Sub TotalCheques(Titulo As String, Moneda As String, Mensaje As String)
Dim rs As New ADODB.Recordset
Dim sql As String
Dim vFecha As Date
Dim vCuenta As String
Dim vfecha1 As Date
Dim vLineas As Integer
Dim J As Integer
Dim Monto As Currency
Dim nBanco As String * 25
Dim Pag As String
sql = " select cCodCta,cNumChq,cCtaBco,cBcoDes,nMontoChq,dRegChq,dValorChq" _
    & " from cheque C," & gcCentralCom & "Bancos T" _
    & " where C.cCodBco = str(T.nBcoCod)" _
    & " and substring(cCodCta,6,1) = '" & Moneda & "'"
If Flag = True And flag1 = True Then
sql = sql + " and  datediff(day,dRegChq, '" & Format(gdFecSis, gsFormatoFecha) & "') = 0"
End If
If Flag = False And flag1 = False Then
If txtFechaF = "__/__/____" Then
   MsgBox "Por favor Ingrese una Fecha", vbInformation, "Aviso"
   txtFechaF.SetFocus
   Exit Sub
   End If
sql = sql + " and  dRegChq between '" & Format(txtFecha, gsFormatoFecha) & "' and '" & Format(CDate(txtFechaF.Text) + 1, gsFormatoFecha) & "'"
End If
If rs.State = 1 Then
   rs.Close
End If
'rs.Open sql, dbCmact, adOpenStatic, adLockOptimistic, adCmdText
If RSVacio(rs) Then
  MsgBox Mensaje, vbInformation, "Aviso"
  Char12 = False
Else
Pag = "01"
vLineas = 7
Char12 = True
Dim i As Integer
'*************************** CABECERA ******************************************************
            lsCadena = lsCadena + Chr(10)
            lsCadena = lsCadena & gsNomAge & Space(20) & "Sección AHORRO" & Space(10) & txtFecha + "  a " + txtFechaF & Space(15) & "Pagina: "
            lsCadena = lsCadena & Pag & Chr$(10) & Chr$(10)
            lsCadena = lsCadena & Space(30) & Titulo & Chr$(10)
            lsCadena = lsCadena & Space(45) & "C U E N T A S   D E   A H O R R O " & Chr$(10)
            lsCadena = lsCadena & Chr$(10)
            lsCadena = lsCadena + String(125, "=") + Chr(10)
            lsCadena = lsCadena + "No DE" + Space(10) + "No DE CHEQUE" + Space(5) + "CUENTA DE CHEQUE" + Space(5) + "BANCO" + Space(30) + "MONTO" + Space(5) + "F.ENTREGA" + Space(5) + "F.VALORIZADOS" + Chr$(10)
            lsCadena = lsCadena + "CUENTA" + Chr$(10)
            lsCadena = lsCadena + String(125, "=") + Chr(10)
'*****************************************************************************************

    Do While Not rs.EOF
    With rs
    nBanco = !cNomtab
    lsCadena = lsCadena & !cCodCta & Space(7) & Space(8 - Len(Trim(!cNumChq)))
    lsCadena = lsCadena & Trim(!cNumChq) & Space(6)
    lsCadena = lsCadena & Space(15 - Len(Trim(!cCtaBco)))
    lsCadena = lsCadena & Trim(!cCtaBco) & Space(5)
    lsCadena = lsCadena & nBanco
    lsCadena = lsCadena & Space(15 - Len(Format(!nMontoChq, "#,##0.00")))
    lsCadena = lsCadena & Format(!nMontoChq, "#,##0.00") & Space(4)
    lsCadena = lsCadena & Format(!dRegChq, gsFormatoFechaView) & Space(4) & Format(!dValorChq, gsFormatoFechaView) & Chr$(10)
        vLineas = vLineas + 1
        If vLineas >= 55 Then
           lsCadena = lsCadena + Chr(12)
           Pag = Str(Val(Pag) + 1)
           vLineas = 7
           Call Cabecera(gsNomAge, txtFecha + "  a  " + txtFechaF, Pag, Titulo)
        End If
        Monto = Monto + !nMontoChq
    End With
    rs.MoveNext
    Loop
    lsCadena = lsCadena & Chr$(10) & Chr$(10)
    lsCadena = lsCadena + String(125, "=") + Chr(10)
    lsCadena = lsCadena & "RESUMEN TOTAL CHEQUES" & Space(59) & Space(13 - Len(Format(Monto, "#,##0.00"))) & Format(Monto, "#,##0.00") + Chr(10)
    End If
rs.Close
Set rs = Nothing
End Sub
Private Sub ChequesTotal(Titulo As String, Mensaje As String)
 Call TotalCheques(Titulo + "  E N  S O L E S", "1", Mensaje + " en Soles")
 If Char12 = True Then
 lsCadena = lsCadena + Chr(12)
 End If
 Call TotalCheques(Titulo + "  E N  D O L A R E S ", "2", Mensaje + " en Dolares")
 If Char12 = True Then
 lsCadena = lsCadena + Chr(12)
 End If
 Flag = False
 flag1 = False
End Sub

Private Sub Cabecera(oficina As String, fFecha As String, Pag As String, Titulo As String)
Dim i As Integer
lsCadena = lsCadena + Chr(10)
lsCadena = lsCadena & oficina & Space(20) & "Sección AHORRO" & Space(10) & fFecha & Space(15) & "Pagina: "
lsCadena = lsCadena & Pag & Chr$(10) & Chr$(10)
lsCadena = lsCadena & Space(30) & Titulo & Chr$(10)
lsCadena = lsCadena & Space(45) & "C U E N T A S   D E   A H O R R O " & Chr$(10)
lsCadena = lsCadena & Chr$(10)
lsCadena = lsCadena + String(125, "=") + Chr(10)
lsCadena = lsCadena + "No DE" + Space(10) + "No DE CHEQUE" + Space(5) + "CUENTA DE CHEQUE" + Space(5) + "BANCO" + Space(30) + "MONTO" + Space(5) + "F.ENTREGA" + Space(5) + "F.VALORIZADOS" + Chr$(10)
lsCadena = lsCadena + "CUENTA" + Chr$(10)
lsCadena = lsCadena + String(125, "=") + Chr(10)
End Sub

Function Verifica() As Boolean
    Dim i As Integer
    For i = 1 To treeRep.Nodes.Count
        If treeRep.Nodes(i).Checked = True Then
           Verifica = True
           Exit Function
        End If
    Next
    Verifica = False
End Function

Function GetReportes() As String
    Dim lsCadena As String
    Dim i As Integer
    lsCadena = ""
    For i = 1 To treeRep.Nodes.Count
        If treeRep.Nodes(i).Checked = True And Len(treeRep.Nodes(i).Key) > 2 Then
           lsCadena = lsCadena & treeRep.Nodes(i).Key & ";"
        End If
    Next
    GetReportes = lsCadena
End Function

Private Sub cmdImprimir_Click()
    Dim oPrevio As Previo.clsPrevio
    Set oPrevio = New Previo.clsPrevio
    Dim lsCadena As String
    Dim lsRep As String
    Dim oRep As nCaptaReportes
    Set oRep = New nCaptaReportes
    
    If Verifica = False Then
         MsgBox " No Tiene Seleccionado ningun Reporte", vbInformation, "Aviso"
         treeRep.SetFocus
         Exit Sub
    ElseIf Not IsDate(txtFecha) Then
        MsgBox "Fecha no valida", vbInformation, "Aviso"
        Me.txtFecha.SetFocus
        Exit Sub
    ElseIf Not IsDate(txtFechaF) Then
        MsgBox "Fecha no valida", vbInformation, "Aviso"
        Me.txtFechaF.SetFocus
        Exit Sub
    End If
    
    gsCodAge = "11201"
    gsNomAge = "AG: PIZARRO"
    lsRep = GetReportes
    lsCadena = oRep.Reporte(lsRep, Me.txtFecha, Me.txtFechaF, gsNomAge, gsEmpresa, gdFecSis, gsCodAge)
    
    If chkCondensado.Value = 1 Then
        oPrevio.Show lsCadena, Caption, True, 66
    Else
        oPrevio.Show lsCadena, Caption, False, 66
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Function ConsolidadoInactivas() As String
Dim SQL1 As String
Dim SQL2 As String
Dim SQL3 As String
Dim rs1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim RS3 As New ADODB.Recordset
Dim lsCad As String
Dim lnNum1 As Long
Dim lnMon1 As Currency
Dim lnNum2 As Long
Dim lnMon2 As Currency
Dim lnNum3 As Long
Dim lnMon3 As Currency
Dim lnTotNumero As Long
Dim lnTotMonto As Currency
Dim i As Integer
 
lsCad = lsCad & Space(10) & "CAJA MUNICIPAL DE " & Space(50) & ArmaFecha(gdFecSis) & Chr(10)
lsCad = lsCad & Space(10) & "    TRUJILLO      " & Space(50) & Str(Time()) & Chr(10)
lsCad = lsCad & Space(10) & Space(30) & "RESUMEN DE DESCUENTO DE INACTIVAS" & Chr(10)
lsCad = lsCad & Space(10) & Space(30) & "---------------------------------" & Chr(10) & Chr(10)
lsCad = lsCad & Space(10) & "Agencia : " & gsNomAge & Chr(10) & Chr(10)
For i = 1 To 2
lnNum1 = 0: lnMon1 = 0: lnNum2 = 0: lnMon2 = 0: lnNum3 = 0
lnMon3 = 0: lnTotNumero = 0: lnTotMonto = 0
    'SQL1 = "SELECT COUNT(nNumTran) as Numero ,sum(nMonTran)as Monto FROM TRANDIARIA" _
         & " WHERE cCodOpe = '" & gsACActIna & "'" _
         & " and substring(cCodCta,6,1) = '" & i & "' And (cFlag is Null or cFlag <> 'X')"
    'SQL2 = "SELECT COUNT(nNumTran) as Numero,sum(nMonTran) as Monto FROM TRANDIARIA" _
         & " WHERE cCodOpe = '" & gsACDesInaAct & "'" _
         & " and substring(cCodCta,6,1) = '" & i & "' And (cFlag is Null or cFlag <> 'X')"
    'SQL3 = "SELECT COUNT(nNumTran) as Numero,sum(nMonTran) as Monto FROM TRANDIARIA" _
         & " WHERE cCodOpe = '" & gsACCanIna & "'" _
         & " and substring(cCodCta,6,1) = '" & i & "' And (cFlag is Null or cFlag <> 'X')"
    
    'rs1.Open SQL1, dbCmact, adOpenForwardOnly, adLockReadOnly, adCmdText
    'RS2.Open SQL2, dbCmact, adOpenForwardOnly, adLockReadOnly, adCmdText
    'RS3.Open SQL3, dbCmact, adOpenForwardOnly, adLockReadOnly, adCmdText
    lnNum1 = IIf(RSVacio(rs1), 0, rs1!Numero)
    lnMon1 = IIf(RSVacio(rs1) Or IsNull(rs1!Monto), 0, rs1!Monto)
    lnNum2 = IIf(RSVacio(RS2), 0, RS2!Numero)
    lnMon2 = IIf(RSVacio(RS2) Or IsNull(RS2!Monto), 0, RS2!Monto)
    lnNum3 = IIf(RSVacio(RS3), 0, RS3!Numero)
    lnMon3 = IIf(RSVacio(RS3) Or IsNull(RS3!Monto), 0, RS3!Monto)
   ' ' RClose rs1
   ' ' RClose RS2
   ' ' RClose RS3
    lnTotNumero = lnNum2 + lnNum3
    lnTotMonto = lnMon2 + lnMon3
    
lsCad = lsCad & Space(10) & "Moneda : " & IIf(i = 1, "SOLES", "DOLARES") & Chr(10)
lsCad = lsCad & Space(10) & "-------------------------------------------------------------------------------------" & Chr(10)
lsCad = lsCad & Space(10) & "Descripción                   " & Space(2) & "Nro     " & Space(2) & "     Monto     " & Chr(10)
lsCad = lsCad & Space(10) & "-------------------------------------------------------------------------------------" & Chr(10)
'lsCad = lsCad & Space(10) & "Paso de Activas a Inactivas : " & Space(2) & JDNum(Str(lnNum1), 6, False, 6, 0) & Space(2) & JDNum(Str(lnMon1), 13, True, 10, 2) & Chr(10)
'lsCad = lsCad & Space(10) & "Descuento de Inactivas      : " & Space(2) & JDNum(Str(lnNum2), 6, False, 6, 0) & Space(2) & JDNum(Str(lnMon2), 13, True, 10, 2) & Chr(10)
'lsCad = lsCad & Space(10) & "Cancelación de Inactivas    : " & Space(2) & JDNum(Str(lnNum3), 6, False, 6, 0) & Space(2) & JDNum(Str(lnMon3), 13, True, 10, 2) & Chr(10) & Chr(10)
lsCad = lsCad & Space(10) & "-------------------------------------------------------------------------------------" & Chr(10)
'lsCad = lsCad & Space(10) & "                      TOTAL : " & Space(2) & JDNum(Str(lnTotNumero), 6, False, 6, 0) & Space(2) & JDNum(Str(lnTotMonto), 13, True, 10, 2) & Chr(10)
lsCad = lsCad & Space(10) & "-------------------------------------------------------------------------------------" & Chr(10) & Chr(10)
Next
ConsolidadoInactivas = lsCad
End Function
Function CartasInactivas() As String
Dim sql As String
Dim rs As New ADODB.Recordset
Dim lsCad As String
Dim lmMes() As Variant
Dim lsBuff As String
Dim i As Long
Screen.MousePointer = vbHourglass

lmMes = Array("Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Setiembre", "Octubre", "Noviembre", "Diciembre")

sql = " SELECT AC.cCodCta,CONVERT(VARCHAR(50),P.cNomPers) as Nombre, convert(varchar(50),P.cDirPers) as Direccion ,P.cCodZon,AC.nSaldDispAC,AC.nSaldCntAC,AC.dUltCntAC,PC.cRelaCta" _
    & " FROM AHORROC AC INNER JOIN PERSCUENTA PC ON AC.cCodCta = PC.cCodCta" _
    & " INNER JOIN " & gcCentralPers & "PERSONA P ON PC.cCodPers = P.cCodPers" _
    & " WHERE bInactiva = 1 and cEstCtaAc NOT IN ('C','U') and cRelaCta = 'TI'"

'rs.Open sql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
If RSVacio(rs) Then
Else
 i = 0
 'lsCad = lsCad & PrnSeteo( "MI", 15)
 
  Do While Not rs.EOF
  lsCad = lsCad & Chr(10) & Chr(10) & Chr(10)
  lsCad = lsCad & Space(30) & "Trujillo " & Str(Day(gdFecSis)) & "  de  " & lmMes(Month(gdFecSis)) & " de " & Str(Year(gdFecSis)) & Chr(10) & Chr(10) & Chr(10) & Chr(10)
  'lsCad = lsCad & PrnSeteo("B+")
  lsCad = lsCad & "Sr(a)." & Chr(10)
  lsCad = lsCad & PstaNombre(rs!Nombre) & Chr(10)
  lsCad = lsCad & rs!Direccion & Chr(10)
  'lsCad = lsCad & Zona(rs!cCodZon) & Chr(10) & Chr(10) & Chr(10)
  'lsCad = lsCad & PrnSeteo("B-")
  lsCad = lsCad & "Estimado Cliente" & Chr(10)
  lsCad = lsCad & "Nos es grato dirigirnos a Usted para Saludarlo y a la vez informarle que hemos visto con preocupación " & Chr(10)
  lsCad = lsCad & "la falta de movimiento de las cuentas que mantiene con nosotros, la que las han convertido en cuentas " & Chr(10)
  lsCad = lsCad & "inactivas con un costo adicional por mantenimiento de las mismas." & Chr(10)
  lsCad = lsCad & "" & Chr(10)
  lsCad = lsCad & "Siendo el ahorro el unico modo de evitar este costo para usted, a continuación le detallamos las tasas" & Chr(10)
  lsCad = lsCad & "de interes que la Caja de Trujillo esta pagando por sus ahorros en moneda nacional y extranjera:" & Chr(10)
  lsCad = lsCad & "" & Chr(10) & Chr(10)
  'lsCad = lsCad & PrnSeteo("B+")
  lsCad = lsCad & Space(30) & "TASA DE INTERES EFECTIVA ANUAL" & Chr(10) & Chr(10)
  lsCad = lsCad & "PRODUCTO         " & Space(2) & "PLAZOS" & Space(15) & "INTERES EFECTIVO ANUAL " & Space(10) & "MONTO APERTURA" & Chr(10)
  lsCad = lsCad & "Ahorro Corriente " & Space(25) & "SOLES" & Space(10) & "DOLARES" & Space(5) & "SOLES" & Space(10) & "DOLARES" & Chr(10)
  'lsCad = lsCad & PrnSeteo("B-")
  lsCad = lsCad & Space(41) & "10.29%" & Space(10) & "5.00%" & Space(10) & " 50" & Space(10) & " 50" & Chr(10)
  lsCad = lsCad & "" & Chr(10)
  'lsCad = lsCad & PrnSeteo("B+")
  'lsCad = lsCad & "Plazo Fijo       " & PrnSeteo("B-") & " 31 a 89 dias " & Space(10) & "18.66%" & Space(10) & "6.00%" & Space(10) & "150" & Space(10) & "100" & Chr(10)
  lsCad = lsCad & Space(17) & " 90 a 179 dias" & Space(10) & "20.44%" & Space(10) & "6.00%" & Space(10) & "150" & Space(10) & "100" & Chr(10)
  lsCad = lsCad & Space(17) & "180 a 359 dias" & Space(10) & "23.11%" & Space(10) & "8.50%" & Space(10) & "150" & Space(10) & "100" & Chr(10) & Chr(10) & Chr(10)
  lsCad = lsCad & "Además le brindamos las siguientes ventajas:" & Chr(10)
  'lsCad = lsCad & PrnSeteo("B+")
  lsCad = lsCad & "AHORRO CORRIENTE " & Chr(10)
  'lsCad = lsCad & PrnSeteo("B-")
  lsCad = lsCad & "Usted puede depositar o retirar su dinero en cualquiera de nuestras agencias o " & Chr(10)
  lsCad = lsCad & "en las oficinas de las cajas Municipales a Nivel Nacional." & Chr(10)
  lsCad = lsCad & "Costo CERO de mantenimiento de Cuenta." & Chr(10) & Chr(10)
  'lsCad = lsCad & PrnSeteo("B+")
  lsCad = lsCad & "DEPOSITO A PLAZO FIJO" & Chr(10)
  'lsCad = lsCad & PrnSeteo("B-")
  lsCad = lsCad & "Retiro de sus interese mensualmente." & Chr(10)
  lsCad = lsCad & "Garantiza su Prestamo Personal de inmediato" & Chr(10)
  lsCad = lsCad & "Cancelación anticipada en casos de emergencia." & Chr(10)
  lsCad = lsCad & Chr(10) & Chr(10) & Chr(10)
  lsCad = lsCad & "Para mayor información estamos para atenderlos."
  lsCad = lsCad & "" & Chr(10) & Chr(10) & Chr(10) & Chr(10)
  lsCad = lsCad & "Atte." & Chr(10) & Chr(10) & Chr(10)
  lsCad = lsCad & "_____________________" & Chr(10)
  lsCad = lsCad & "   Administrador(a)    " & Chr(10) & Chr(12)
  
  i = i + 20
  If i Mod 40 = 0 Then
    lsBuff = lsBuff & lsCad
    lsCad = ""
  End If

  rs.MoveNext
  Loop
End If
' RClose rs
lsBuff = lsBuff & lsCad
CartasInactivas = lsBuff
Screen.MousePointer = vbNormal
End Function

Private Sub Form_Load()
    Dim oCon As DConecta
    Dim sqlV As String
    Dim rs As ADODB.Recordset
    Dim nodX As Node
    Dim lsCodCab As String
    Set rs = New ADODB.Recordset
    
    Set oCon = New DConecta
    
    sqlV = " Select  Case Right(cCapRepCod,2) When '00' Then Left(cCapRepCod,1) Else cCapRepCod End Codigo, cCapRepDescripcion Descripcion," _
         & " Case Right(cCapRepCod,2) When '00' Then 1 Else 2 End Nivel" _
         & " From CaptaReportes"
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sqlV)

    If Not (rs.EOF And rs.BOF) Then
        Set nodX = treeRep.Nodes.Add(, , "P", "TODO", "Close")
        nodX.Expanded = True
        While Not rs.EOF
            If rs!Nivel = 1 Then
                Set nodX = treeRep.Nodes.Add("P", tvwChild, rs!Codigo, rs!Descripcion, "Close")
                lsCodCab = rs!Codigo
            Else
                Set nodX = treeRep.Nodes.Add(lsCodCab, tvwChild, rs!Codigo, rs!Descripcion, "Hoja1")
            End If
            rs.MoveNext
        Wend
    End If
    
    txtFecha = Format(gdFecSis, gsFormatoFechaView)
    txtFechaF = Format(gdFecSis, gsFormatoFechaView)
End Sub

Private Sub treeRep_Click()
    Dim i As Integer
    Select Case treeRep.SelectedItem.Key
        Case "P"
            If treeRep.SelectedItem.Checked = True Then
              treeRep.SelectedItem.Image = "Open"
              For i = 1 To treeRep.Nodes.Count
                  treeRep.Nodes(i).Checked = True
                  Select Case treeRep.Nodes(i).Key
                      Case "P", "D", "M", "V"
                      Case Else
                          treeRep.Nodes(i).Image = "Hoja1"
                  End Select
              Next
            Else
            treeRep.SelectedItem.Image = "Close"
              For i = 1 To treeRep.Nodes.Count
                  treeRep.Nodes(i).Checked = False
                  Select Case treeRep.Nodes(i).Key
                      Case "P", "D", "M", "V"
                      Case Else
                          treeRep.Nodes(i).Image = "Hoja"
                  End Select
               Next
           End If
             
       Case "D", "M", "V"
            If treeRep.SelectedItem.Checked = True Then
            For i = 1 To treeRep.Nodes.Count
                If Mid(treeRep.Nodes(i).Key, 1, 1) = treeRep.SelectedItem.Key Then
                      treeRep.Nodes(i).Checked = True
                      treeRep.Nodes(i).Image = "Hoja"
                      treeRep.Nodes(i).ForeColor = vbBlue
                End If
            Next
              treeRep.SelectedItem.Image = "Open"
            Else
                For i = 1 To treeRep.Nodes.Count
                  If Mid(treeRep.Nodes(i).Key, 1, 1) = treeRep.SelectedItem.Key Then
                      treeRep.Nodes(i).Checked = False
                      treeRep.Nodes(i).Image = "Hoja1"
                     treeRep.Nodes(i).ForeColor = vbBlack
                 End If
                Next
                treeRep.SelectedItem.Image = "Close"
            End If
       Case Else
    '        treeRep.SelectedItem.Checked = Not treeRep.SelectedItem.Checked
            If treeRep.SelectedItem.Checked = True Then
            treeRep.SelectedItem.Image = "Hoja"
            treeRep.SelectedItem.ForeColor = vbBlue
            Else
            treeRep.SelectedItem.Image = "Hoja1"
            treeRep.SelectedItem.ForeColor = vbBlack
            End If
    End Select
End Sub
Private Sub treeRep_Collapse(ByVal Node As MSComctlLib.Node)
    Node.ExpandedImage = 1
End Sub
Private Sub treeRep_Expand(ByVal Node As MSComctlLib.Node)
 Node.ExpandedImage = 2
End Sub
Private Sub treeRep_NodeCheck(ByVal Node As MSComctlLib.Node)
    Dim i As Integer
    treeRep.SelectedItem = Node
    Select Case Node.Key
       Case "P"
         If Node.Checked = True Then
              For i = 1 To treeRep.Nodes.Count
                  treeRep.Nodes(i).Checked = True
              Next
         Else
              For i = 1 To treeRep.Nodes.Count
                  treeRep.Nodes(i).Checked = False
                 Next
         End If
       Case "D", "M", "V"
           If Node.Checked = True Then
              For i = 1 To treeRep.Nodes.Count
                If Mid(treeRep.Nodes(i).Key, 1, 1) = Node.Key Then
                      treeRep.Nodes(i).Checked = True
                      treeRep.Nodes(i).Image = "Hoja"
                      treeRep.Nodes(i).ForeColor = vbBlue
    
                End If
              Next
           Else
              For i = 1 To treeRep.Nodes.Count
                 If Mid(treeRep.Nodes(i).Key, 1, 1) = Node.Key Then
                      treeRep.Nodes(i).Checked = False
                      treeRep.Nodes(i).Image = "Hoja1"
                     treeRep.Nodes(i).ForeColor = vbBlack
                 End If
              Next
           End If
       Case Else
        
           If Node.Checked = True Then
               Node.Image = "Hoja"
               Node.ForeColor = vbBlue
           Else
               Node.Image = "Hoja1"
               Node.ForeColor = vbBlack
           End If
    End Select
End Sub

Private Sub txtFecha_GotFocus()
    fEnfoque txtFecha
End Sub

Private Sub txtFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtFechaF.SetFocus
    End If
End Sub

Private Sub txtFechaF_GotFocus()
fEnfoque txtFechaF
End Sub

Private Sub txtFechaF_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    treeRep.SetFocus
End If
End Sub

Public Sub PrintCartasSepelio()

'On Error GoTo ControlError
'    'Carga archivo rtfCartas
'    vBuffer = ""
'    'lsCadena.FileName = App.Path & cPlantillaAseguradoSepelio
'    'ImprimirCartasSepelio lsCadena
'    dlgGrabar.CancelError = True
'    dlgGrabar.InitDir = App.Path
'    dlgGrabar.Filter = "Archivos de Texto (*.TXT)|*.TXT"
'    dlgGrabar.ShowSave
'    If dlgGrabar.FileName <> "" Then
'        Open dlgGrabar.FileName For Output As #1
'        Print #1, vBuffer
'        Close #1
'    End If
'Exit Sub
'
'ControlError:   ' Rutina de control de errores.
'    If Err.Number = 32755 Then
'        MsgBox " Grabación Cancelada ", vbInformation, " Aviso "
'    Else
'        MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
'            " Avise al Area de Sistemas ", vbInformation, " Aviso "
'    End If
End Sub

Private Function GetCuentasSepelio(psCodPers As String) As String
'Dim rsCuentas As New ADODB.Recordset
'Dim lsCuentas As String
''VSQL= "Select cCodCta from PersCtaFonSep where cCodPers = '" & psCodPers & "' " _
'    & "ORDER BY cCodCta"
''rsCuentas.Open VSQL, dbCmact, adOpenForwardOnly, adLockReadOnly, adCmdText
'If RSVacio(rsCuentas) Then
'    MsgBox "Cliente no posee cuentas relacionadas.", vbExclamation, "Error"
'Else
'    lsCuentas = ""
'    Do While Not rsCuentas.EOF
'        lsCuentas = lsCuentas & String(16, " ") & Trim(rsCuentas!cCodCta) & vbCrLf
'        rsCuentas.MoveNext
'    Loop
'End If
'rsCuentas.Close
'Set rsCuentas = Nothing
'GetCuentasSepelio = lsCuentas
End Function

Public Sub ImprimirCartasSepelio(ByVal rtfCartas As RichTextBox)
'    Dim RegProcesar As New ADODB.Recordset
'    Dim sSQL As String
'    'Dim vFecAviso As Date
'    Screen.MousePointer = 11
'    gdFecSis = CDate("01/05/2000")
'    'sSQL = "SELECT P.cCodPers, P.cNomPers, P.cDirPers, Z.cDesZon FROM (" & gcCentralPers & "Persona P LEFT JOIN " & gcCentralCom & "Zonas Z ON P.cCodZon = Z.cCodZon) " _
'        & "INNER JOIN PerFonSep PS ON P.cCodPers = PS.cCodPers Order by P.cNomPers"
'    'RegProcesar.Open sSQL, dbCmact, adOpenForwardOnly, adLockReadOnly, adCmdText
'    If RSVacio(RegProcesar) Then
'        RegProcesar.Close
'        Set RegProcesar = Nothing
'        MsgBox "No existen clientes para el Fondo de Seguro de Sepelio.", vbInformation, " Aviso "
'        Screen.MousePointer = 0
'        Exit Sub
'    Else
'        'Llena cartas
'        Dim x As Integer
'        Dim vNom As String
'        Dim vDir As String * 75
'        Dim vFecha As String
'        Dim vMes As String
'        Dim vMesAnt As String
'        Dim vZona As String
'        'Dim vHorRem As String
'        Dim RTFTmp As String
'        Dim vPrevio As String
'        Dim lsCarta As String
'
'        lsCarta = ""
'        RTFTmp = ""
'        vBuffer = ""
'        x = 0
'        vFecha = Format(gdFecSis, "dddd, d mmmm yyyy")
'        vMes = UCase(Format(gdFecSis, "mmmm"))
'        vMesAnt = UCase(Format(DateAdd("m", -1, gdFecSis), "mmmm"))
'        lsCarta = rtfCartaslscadena
'        Do While Not RegProcesar.EOF
'            vNom = PstaNombre(RegProcesar!cNomPers, False)
'            vDir = RegProcesar!cDirPers
'            vZona = IIf(IsNull(RegProcesar!cDesZon), "", Trim(RegProcesar!cDesZon))
'            RTFTmp = lsCarta
'            RTFTmp = Replace(RTFTmp, "<<FECHA>>", vFecha, , 1, vbTextCompare)
'            RTFTmp = Replace(RTFTmp, "<<CLIENTE>>", vNom, , 1, vbTextCompare)
'            RTFTmp = Replace(RTFTmp, "<<DIRECCION>>", vDir, , 1, vbTextCompare)
'            RTFTmp = Replace(RTFTmp, "<<ZONA>>", vZona, , 1, vbTextCompare)
'            RTFTmp = Replace(RTFTmp, "<<MES>>", vMes, , 1, vbTextCompare)
'            'RTFTmp = Replace(RTFTmp, "<<MESANT>>", vMesAnt, , 1, vbTextCompare)
'            RTFTmp = Replace(RTFTmp, "<<CUENTAS>>", GetCuentasSepelio(RegProcesar!cCodPers), , 1, vbTextCompare)
'            RTFTmp = RTFTmp & Chr(12)
'            If x Mod 50 = 0 Then
'                vBuffer = vBuffer & vPrevio
'                vPrevio = ""
'            End If
'            vNom = "":   vDir = "": vZona = ""
'            vPrevio = vPrevio & RTFTmp
'            RTFTmp = ""
'            x = x + 1
'            RegProcesar.MoveNext
'        Loop
'    End If
'    vBuffer = vBuffer & vPrevio
'    Screen.MousePointer = 0
End Sub

Public Sub ImprimirCartasFirmas()
'    Dim dbAux As New ADODB.Connection
'    Dim RegProcesar As New ADODB.Recordset
'    Dim rsAux As New ADODB.Recordset
'    Dim lsConn As String
'    Screen.MousePointer = 11
'    gdFecSis = CDate("19/05/2000")
'    lsConn = "DSN=DSNRRHH;UID=;PWD"
'    dbAux.Open lsConn
'    'VSQL= "Select cCodCtaAnt from CartaFirma"
'    'RegProcesar.Open VSQL, dbAux, adOpenForwardOnly, adLockReadOnly, adCmdText
'    If RSVacio(RegProcesar) Then
'        RegProcesar.Close
'        Set RegProcesar = Nothing
'        MsgBox "No existen clientes para el Fondo de Seguro de Sepelio.", vbInformation, " Aviso "
'        Screen.MousePointer = 0
'        Exit Sub
'    Else
'        'On Error Resume Next
'        Do While Not RegProcesar.EOF
'            If Len(Trim(RegProcesar!ccodctaant)) = 8 Then
'                'VSQL= "SELECT P.cCodPers, P.cNomPers, P.cDirPers, Z.cDesZon FROM " _
'                    & "(" & gcCentralPers & "Persona P LEFT JOIN " & gcCentralCom & "Zonas Z ON P.cCodZon = Z.cCodZon) " _
'                    & "INNER JOIN PersCuenta PC INNER JOIN RelCuentas R ON " _
'                    & "PC.cCodCta = R.cCodCta ON P.cCodPers = PC.cCodPers AND " _
'                    & "R.cCodCtaAnt = '" & Trim(RegProcesar!ccodctaant) & "' AND PC.cRelaCta = 'TI'"
'
'            Else
'                'VSQL= "SELECT P.cCodPers, P.cNomPers, P.cDirPers, Z.cDesZon FROM " _
'                    & "(" & gcCentralPers & "Persona P LEFT JOIN " & gcCentralCom & "Zonas Z ON P.cCodZon = Z.cCodZon) " _
'                    & "INNER JOIN PersCuenta PC ON P.cCodPers = PC.cCodPers AND " _
'                    & "PC.cCodCta = '" & Trim(RegProcesar!ccodctaant) & "' AND PC.cRelaCta = 'TI'"
'            End If
'            If rsAux.State = adStateOpen Then rsAux.Close
'            'rsAux.Open VSQL, dbCmact, adOpenForwardOnly, adLockReadOnly, adCmdText
'            'If Not RSVacio(rsAux) And Not ExisteFirma(rsAux!cCodPers) Then
'                'VSQL= "INSERT INTO CartaFirma VALUES('" & rsAux!cCodPers & "', '" _
'                    & PstaNombre(rsAux!cNomPers, True) & "', '" & rsAux!cDirPers & "', '" _
'                    & rsAux!cDesZon & "')"
'                dbCmact.Execute VSQL
'            End If
'            RegProcesar.MoveNext
'        Loop
'
'    End If
'    RegProcesar.Close
'    Set RegProcesar = Nothing
'    dbAux.Close
'    Set dbAux = Nothing
'    MsgBox "Proceso finalizado con éxito", vbInformation, "Aviso"
'    Screen.MousePointer = 0
End Sub

Public Sub ImprimirCartasFirmasZofra()
'    Dim rsAux As New ADODB.Recordset
'    Dim lsCad As String
'    Screen.MousePointer = 11
'    'VSQL= "Select P.cCodPers, P.cNomPers, P.cDirPers, Z.cDesZon from (" & gcCentralPers & "Persona P LEFT JOIN " _
'        & gcCentralCom & "Zonas Z ON P.cCodZon = Z.cCodZon) INNER JOIN PersCuenta PC INNER JOIN AhorroC A " _
'        & "ON PC.cCodCta = A.cCodCta ON P.cCodPers = PC.cCodPers WHERE P.cNudoci = '00000000' " _
'        & "AND A.cEstCtaAC not in ('C','U')"
'    If rsAux.State = adStateOpen Then rsAux.Close
'    'rsAux.Open VSQL, dbCmact, adOpenForwardOnly, adLockReadOnly, adCmdText
'    If RSVacio(rsAux) Then
'        rsAux.Close
'        Set rsAux = Nothing
'        MsgBox "No existen clientes para el Fondo de Seguro de Sepelio.", vbInformation, " Aviso "
'        Screen.MousePointer = 0
'        Exit Sub
'    Else
'        'On Error Resume Next
'        lsCad = ""
'        Do While Not rsAux.EOF
'            lsCad = lsCad & rsAux!cCodPers & ";" & ImpreCarEsp(Trim(PstaNombre(rsAux!cNomPers))) & ";" & ImpreCarEsp(Trim(rsAux!cDirPers)) & ";" & ImpreCarEsp(Trim(rsAux!cDesZon)) & Chr$(10)
'            rsAux.MoveNext
'        Loop
'    End If
'    'VSQL= "Select P.cCodPers, P.cNomPers, P.cDirPers, Z.cDesZon from (" & gcCentralPers & "Persona P LEFT JOIN " _
'        & gcCentralCom & "Zonas Z ON P.cCodZon = Z.cCodZon) INNER JOIN PersCuenta PC INNER JOIN PlazoFijo A " _
'        & "ON PC.cCodCta = A.cCodCta ON P.cCodPers = PC.cCodPers WHERE P.cNudoci = '00000000' " _
'        & "AND A.cEstCtaPF not in ('C','U')"
'    If rsAux.State = adStateOpen Then rsAux.Close
'    'rsAux.Open VSQL, dbCmact, adOpenForwardOnly, adLockReadOnly, adCmdText
'    If RSVacio(rsAux) Then
'        rsAux.Close
'        Set rsAux = Nothing
'        MsgBox "No existen clientes para el Fondo de Seguro de Sepelio.", vbInformation, " Aviso "
'        Screen.MousePointer = 0
'        Exit Sub
'    Else
'        'On Error Resume Next
'        Do While Not rsAux.EOF
'            lsCad = lsCad & rsAux!cCodPers & ";" & ImpreCarEsp(Trim(PstaNombre(rsAux!cNomPers))) & ";" & ImpreCarEsp(Trim(rsAux!cDirPers)) & ";" & ImpreCarEsp(Trim(rsAux!cDesZon)) & Chr$(10)
'            rsAux.MoveNext
'        Loop
'    End If
'    rsAux.Close
'    Set rsAux = Nothing
'
'    lsCadena = lsCad
'    Screen.MousePointer = 0
'    frmPrevio.Previo lsCadena, "Firmas", False, 66
'
End Sub


