VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmGeneraDataProyFinan 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Genera Data para Proyecto Financiero"
   ClientHeight    =   1995
   ClientLeft      =   4380
   ClientTop       =   4365
   ClientWidth     =   5460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   5460
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ProgressBar prgBarra 
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Top             =   960
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComCtl2.DTPicker txtfecha 
      Height          =   330
      Left            =   2460
      TabIndex        =   5
      Top             =   60
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   582
      _Version        =   393216
      Format          =   63373313
      CurrentDate     =   38525
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   360
      Left            =   3165
      TabIndex        =   3
      Top             =   1380
      Width           =   1380
   End
   Begin VB.OptionButton optBal 
      Caption         =   "Estado de Perdida y Ganancias"
      Height          =   315
      Index           =   1
      Left            =   2085
      TabIndex        =   2
      Top             =   510
      Width           =   2670
   End
   Begin VB.CommandButton cmdProcesar 
      Caption         =   "Procesar"
      Height          =   360
      Left            =   1185
      TabIndex        =   1
      Top             =   1380
      Width           =   1380
   End
   Begin VB.OptionButton optBal 
      Caption         =   "Balance"
      Height          =   315
      Index           =   0
      Left            =   750
      TabIndex        =   0
      Top             =   495
      Value           =   -1  'True
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fecha de Información:"
      Height          =   195
      Left            =   660
      TabIndex        =   4
      Top             =   120
      Width           =   1590
   End
End
Attribute VB_Name = "frmGeneraDataProyFinan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdProcesar_Click()
Dim oCon As DConecta
Set oCon = New DConecta
Dim rsPlan As ADODB.Recordset
Dim rs As ADODB.Recordset

Dim sql As String
Dim lnTotal As Long

Dim pnItem As Long
Dim pcDescri As String
Dim pcCodRub As String
Dim pcTipCta As String
Dim pcgrupo1 As String
Dim pcSigno1 As String
Dim pcProcesa1 As String
Dim pcValor1 As Long
Dim pcNivGru As String
Dim pcColumn As String
Dim pnSalIni As Currency
Dim pnSalFin As Currency
Dim pcUnidad As String
Dim pnMonto As Currency
Dim pnMonto2 As Currency
Dim pnMonto3 As Currency
Dim lsTipoinfo As String
Dim lsAnio As String
Dim lsMes As String

lsAnio = Format(Me.txtfecha, "yyyy")
lsMes = Format(txtfecha, "mm")

If optBal(0).value Then
    lsTipoinfo = "20"
Else
    lsTipoinfo = "12"
End If

sql = "         SELECT  r.cCodInf as cCodInfPrinc, r.cDescri, r.cNivGru, R.cCodInf,"
sql = sql + "           R1.cGruCta1, R1.cSigno1, R1.nValor1, R1.cProcesa1, r1.cTipCta,"
sql = sql + "           r1.cColumn, r1.cCodRay "
sql = sql + "   FROM    RC6000 R"
sql = sql + "           join RC6001 R1 ON R.CCODINF = R1.CCODRAY"
sql = sql + "   WHERE   SUBSTRING(R.CCODINF,1,2)='" & lsTipoinfo & "'"
sql = sql + "   ORDER BY R1.CCODRAY, R1.cColumn "


oCon.AbreConexion
Set rsPlan = oCon.Ejecutar(sql)
lnTotal = rsPlan.RecordCount
Me.prgBarra.Max = lnTotal

sql = "DELETE FROM BALANCEPROYFINAN"
oCon.Ejecutar sql

Do While Not rsPlan.EOF
        pnItem = rsPlan!cCodInfPrinc
        pcDescri = Trim(rsPlan!CDESCRI)
        pcCodRub = Mid(rsPlan!cCodInfPrinc, 3, 4)
        pcTipCta = rsPlan!cTipCta
        pcgrupo1 = Trim(rsPlan!cGruCta1)
        pcSigno1 = rsPlan!cSigno1
        pcProcesa1 = rsPlan!cProcesa1
        pcValor1 = Val(rsPlan!nValor1)
        pcNivGru = rsPlan!CNIVGRU
        pcColumn = Trim(rsPlan!cColumn)
        
        'If pcCodRub = "1230" Then Stop
        
        sql = "SELECT * " _
            & " From    DBCMACICA.dbo.BalanceEstad " _
            & " Where   cBalanceCate = '1' And cBalanceTipo IN ('1','2') And cBalanceAnio = '" & lsAnio & "' And cBalanceMes = '" & lsMes & "' " _
            & "         and CCTACONTCOD ='" & pcgrupo1 & "' " _
            & " ORDER BY CCTACONTCOD "
                
        Set rs = oCon.CargaRecordSet(sql)
        Do While Not rs.EOF And Not rs.BOF
            pnSalIni = rs!nSaldoIniImporte
            pnSalFin = rs!nSaldoFinImporte
            pcUnidad = "01"
            If Len(Trim(pcgrupo1)) >= 10 Then ' verificando cuenta analitica
                Select Case pcTipCta
                    Case Is = "2"
                        'pnSalFin = IIf(pnSalFin < 0, pnSalFin * -1, 0)
                        pnSalFin = IIf(pnSalFin < 0, pnSalFin * -1, pnSalFin)
                    Case Else
                        pnSalFin = pnSalFin
                End Select
            End If
            
            If pcProcesa1 = "S" Then
                If Left(pcgrupo1, 8) = "28120101" Or Left(pcgrupo1, 8) = "28120102" Or Left(pcgrupo1, 8) = "28220101" Or Left(pcgrupo1, 8) = "28220102" Or Left(pcgrupo1, 8) = "28020101" Or Left(pcgrupo1, 8) = "28020102" Then
                    pnSalIni = -pnSalIni
                    pnSalFin = -pnSalFin
                End If
                If (Mid(pcgrupo1, 3, 1) = "1" Or Mid(pcgrupo1, 3, 1) = "2" Or Mid(pcgrupo1, 3, 1) = "0" Or Mid(pcgrupo1, 3, 1) = "6") And pcColumn = "1" Then ' soles
                    If pcValor1 = "2" Then
                        pnMonto = GetFormula(pcSigno1, pnMonto, pnSalFin)
                    Else
                        pnMonto = GetFormula(pcSigno1, pnMonto, pnSalIni)
                    End If
                End If
                If (Mid(pcgrupo1, 3, 1) = "1" Or Mid(pcgrupo1, 3, 1) = "2" Or Mid(pcgrupo1, 3, 1) = "0" Or Mid(pcgrupo1, 3, 1) = "6") _
                        And pcColumn = "2" Then  ' Dolares al tipo de cambio
                    If pcValor1 = "2" Then
                        pnMonto2 = GetFormula(pcSigno1, pnMonto2, pnSalFin)
                    Else
                        pnMonto2 = GetFormula(pcSigno1, pnMonto2, pnSalIni)
                    End If
                End If
                If (Mid(pcgrupo1, 3, 1) = "1" Or Mid(pcgrupo1, 3, 1) = "2" Or Mid(pcgrupo1, 3, 1) = "0" _
                    Or Mid(pcgrupo1, 3, 1) = "6") And pcColumn = "3" Then ' ajustado al tipo de cambio
                    If pcValor1 = "2" Then
                        pnMonto3 = GetFormula(pcSigno1, pnMonto3, pnSalFin)
                    Else
                        pnMonto3 = GetFormula(pcSigno1, pnMonto3, pnSalIni)
                    End If
                End If
            End If
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
    Me.prgBarra.value = rsPlan.Bookmark
    rsPlan.MoveNext
    If rsPlan.EOF Then
            sql = "INSERT INTO BALANCEPROYFINAN (CDESCRI,NMONMN,NMONME,NTOTAL,NMONAJU,DFECMES,CCODINS,CCODIGO,CUNIDAD,CNIVGRU ) " _
                & "VALUES('" & pcDescri & "'," & pnMonto & "," & pnMonto2 & "," & pnMonto + pnMonto2 & "," & pnMonto3 & "," _
                & "'04','108','" & pcCodRub & "','" & pcUnidad & "','" & pcNivGru & "')"
            
            oCon.Ejecutar sql
            
            pnMonto = 0
            pnMonto2 = 0
            pnMonto3 = 0
        Exit Do
    End If
    If pnItem <> rsPlan!cCodInfPrinc Then
        sql = "INSERT INTO BALANCEPROYFINAN (CDESCRI,NMONMN,NMONME,NTOTAL,NMONAJU,DFECMES,CCODINS,CCODIGO,CUNIDAD,CNIVGRU ) " _
            & "VALUES('" & pcDescri & "'," & pnMonto & "," & pnMonto2 & "," & pnMonto + pnMonto2 & "," & pnMonto3 & "," _
            & "'04','108','" & pcCodRub & "','" & pcUnidad & "','" & pcNivGru & "')"
        
        oCon.Ejecutar sql
        
        pnMonto = 0
        pnMonto2 = 0
        pnMonto3 = 0
    End If
    Me.Caption = "Procesando Plantilla " & rsPlan.Bookmark & " DE " & rsPlan.RecordCount
    DoEvents
Loop
rsPlan.Close
Set rsPlan = Nothing
MsgBox "Proceso Realizado con éxito", vbInformation, "Aviso"
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub
Function GetFormula(ByVal psSigno As String, ByVal pnMonto As Currency, ByVal pnDato As Currency)
Select Case psSigno
    Case Is = "+"
        GetFormula = pnMonto + pnDato
    Case Is = "-"
        GetFormula = pnMonto - pnDato
    Case Is = "*"
        GetFormula = Round(pnMonto * pnDato, 2)
    Case Is = "/"
        GetFormula = Round(pnMonto / pnDato, 2)
End Select

End Function

Private Sub Form_Load()
CentraForm Me
End Sub
