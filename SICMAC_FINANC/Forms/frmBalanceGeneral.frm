VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBalanceGeneral 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Balance General Forma A y B"
   ClientHeight    =   2865
   ClientLeft      =   1905
   ClientTop       =   3435
   ClientWidth     =   4890
   Icon            =   "frmBalanceGeneral.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   4890
   Begin VB.Frame fraAGE 
      Caption         =   "&Agencia"
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
      Height          =   735
      Left            =   90
      TabIndex        =   12
      Top             =   750
      Width           =   4770
      Begin VB.CheckBox chkAG 
         Appearance      =   0  'Flat
         Caption         =   "Todas"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3750
         TabIndex        =   15
         Top             =   0
         Width           =   825
      End
      Begin Sicmact.TxtBuscar txtAG 
         Height          =   285
         Left            =   120
         TabIndex        =   13
         Top             =   315
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   503
         Appearance      =   0
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   0
      End
      Begin VB.Label lblAG 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   855
         TabIndex        =   14
         Top             =   330
         Width           =   3810
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Expresado en ..."
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
      Height          =   615
      Left            =   90
      TabIndex        =   11
      Top             =   1485
      Width           =   4755
      Begin VB.OptionButton OptMoneda 
         Caption         =   "Miles de Soles"
         Height          =   225
         Index           =   1
         Left            =   2730
         TabIndex        =   3
         Top             =   270
         Value           =   -1  'True
         Width           =   1425
      End
      Begin VB.OptionButton OptMoneda 
         Caption         =   "Nuevos Soles"
         Height          =   225
         Index           =   0
         Left            =   510
         TabIndex        =   2
         Top             =   270
         Width           =   1425
      End
   End
   Begin MSComctlLib.ProgressBar PrgBarra 
      Height          =   270
      Left            =   1500
      TabIndex        =   6
      Top             =   2610
      Width           =   3405
      _ExtentX        =   6006
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   10
      Top             =   2580
      Width           =   4890
      _ExtentX        =   8625
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6068
            MinWidth        =   6068
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame3 
      Caption         =   "&Periodo"
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
      Height          =   735
      Left            =   75
      TabIndex        =   7
      Top             =   15
      Width           =   4785
      Begin VB.TextBox txtAnio 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   690
         MaxLength       =   4
         TabIndex        =   0
         Top             =   270
         Width           =   855
      End
      Begin VB.ComboBox cboMes 
         Height          =   315
         ItemData        =   "frmBalanceGeneral.frx":030A
         Left            =   2775
         List            =   "frmBalanceGeneral.frx":0332
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   270
         Width           =   1875
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Año :"
         Height          =   195
         Left            =   210
         TabIndex        =   9
         Top             =   330
         Width           =   375
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Mes :"
         Height          =   195
         Left            =   2265
         TabIndex        =   8
         Top             =   315
         Width           =   390
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   360
      Left            =   2670
      TabIndex        =   5
      Top             =   2175
      Width           =   1755
   End
   Begin VB.CommandButton cmdProcesar 
      Caption         =   "&Procesar"
      Height          =   360
      Left            =   480
      TabIndex        =   4
      Top             =   2160
      Width           =   1755
   End
End
Attribute VB_Name = "frmBalanceGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type TCuentas
    cCta As String
    nMN As Double
    nMNAj As Double
    nTotal As Double
    nTotAj As Double
    nMExt As Double
    bSaldoA As Boolean
    bSaldoD As Boolean
    TodoASoles As Boolean
End Type
Dim Cuentas() As TCuentas
Dim nCuentas As Integer
Dim CadImp As String
Dim dFecha As Date
Dim sTipoRepoFormula As String

Private Function DepuraMichi(sCodigo As String) As String
Dim R As New ADODB.Recordset
Dim sSql As String
Dim sCadFor As String
Dim sCadRes As String
Dim i As Integer
Dim bFinal As Boolean
Dim sCod As String
Dim TodoASoles As Boolean
Dim Aspersand As Boolean
Dim oRepFormula As New DRepFormula
    TodoASoles = False
    Aspersand = False
    If Mid(sCodigo, 1, 1) = "&" Then
        sCodigo = Mid(sCodigo, 2, Len(sCodigo) - 1)
        TodoASoles = True
        Aspersand = True
    End If
    Set R = oRepFormula.CargaRepFormula(sCodigo, gContBalanceFormaAB)
        sCadFor = Trim(R!cFormula)
    R.Close
    sCadRes = ""
    i = 1
    Do While i <= Len(sCadFor)
        If Mid(sCadFor, i, 1) <> "#" Then
            If TodoASoles And Mid(sCadFor, i, 1) >= "0" And Mid(sCadFor, i, 1) <= "9" And Aspersand Then
                sCadRes = sCadRes + "&"
                Aspersand = False
            Else
                If Not (Mid(sCadFor, i, 1) >= "0" And Mid(sCadFor, i, 1) <= "9") Then
                    Aspersand = True
                End If
            End If
            sCadRes = sCadRes + Mid(sCadFor, i, 1)
        Else
            i = i + 2
            bFinal = False
            sCod = ""
            Do While Not bFinal
                If Mid(sCadFor, i, 1) <> "]" Then
                    sCod = sCod + Mid(sCadFor, i, 1)
                Else
                    bFinal = True
                End If
                i = i + 1
            Loop
            sCadRes = sCadRes + DepuraMichi(sCod)
            i = i - 1
        End If
        i = i + 1
    Loop
    DepuraMichi = sCadRes
End Function
Private Function DepuraFormula(sFormula As String, Optional ByRef nConMichi As Integer = 0) As String
Dim sCad As String
Dim R As New ADODB.Recordset
Dim sSql As String
Dim i As Integer
Dim nContador As Integer
Dim sCadRes As String
Dim bFinal As Boolean
Dim sCod As String
Dim cSigno As String
Dim cSigno2 As String
    sCad = sFormula
    i = 1
    sCadRes = ""
    nContador = 0
    Do While i <= Len(sCad)
        If Mid(sCad, i, 1) <> "#" Then
        If nConMichi = 0 Or (nConMichi = 1 And (Mid(sCad, i, 1) > "0" And Mid(sCad, i, 1) <> "9")) Then
            sCadRes = sCadRes + Mid(sCad, i, 1)
        End If
        Else
            If nContador = 0 Then
                sCadRes = ""
            End If
            
            nContador = nContador + 1
            i = i + 2
            bFinal = False
            sCod = ""
            Do While Not bFinal
                'If Mid(sCad, I, 1) <> "]" Then
                If (Mid(sCad, i, 1) = "+" Or Mid(sCad, i, 1) = "-") Then
                    cSigno = Mid(sCad, i, 1)
                ElseIf (Mid(sCad, i + 1, 1) = "+" Or Mid(sCad, i + 1, 1) = "-") Then
                    cSigno = Mid(sCad, i + 1, 1)
                ElseIf (Mid(sCad, i + 2, 1) = "+" Or Mid(sCad, i + 2, 1) = "-") Then
                    cSigno = Mid(sCad, i + 2, 1)
                End If
                If (Mid(sCad, i, 1) >= "0" And Mid(sCad, i, 1) <= "9") Or (Mid(sCad, i, 1) = "+" Or Mid(sCad, i, 1) = "-") Then
                    sCod = sCod + Mid(sCad, i, 1)
                Else
                    bFinal = True
                End If
                i = i + 1
            Loop
            'ALPA 20090819********************************
            'sCadRes = sCadRes + DepuraMichi(sCod)
            
            If Trim(sCadRes) = "" Then
                sCadRes = sCod
            Else
                sCadRes = sCadRes & cSigno2 & sCod
            End If
            nConMichi = 1
            cSigno2 = cSigno
            '*********************************************
            i = i - 1
        End If
        i = i + 1
    Loop
    DepuraFormula = sCadRes
End Function
Private Function DameTitulo(ByVal sTit As String)
Dim i As Integer
Dim sCad As String
Dim bEnc As Boolean
    bEnc = False
    sCad = ""
    For i = 1 To Len(sTit)
        If Mid(sTit, i, 1) = "\" Then
            bEnc = True
        End If
        If bEnc Then
            sCad = sCad & Mid(sTit, i, 1)
        End If
    Next i
    DameTitulo = sCad
End Function
Private Function DepuraSaldoAD(ByVal sCta As String) As String
Dim i As Integer
Dim Cad As String
    Cad = ""
    For i = 1 To Len(sCta)
        If Mid(sCta, i, 1) >= "0" And Mid(sCta, i, 1) <= "9" Then
            Cad = Cad + Mid(sCta, i, 1)
        End If
    Next i
    DepuraSaldoAD = Cad
End Function

Private Sub procesar(psAgenciaCod As String)
Dim CTemp As String
Dim cTemp1 As String
Dim i As Integer
Dim sSql As String
Dim sSql2 As String
Dim R As New ADODB.Recordset
Dim RMon As New ADODB.Recordset
Dim MN As Currency
Dim MExt As Currency
Dim nTotal As Currency
Dim nTotAj As Currency
Dim MNAj As Double
Dim CadFormula1 As String
Dim CadFormula2 As String
Dim CadFormula3 As String
Dim Total As Integer
Dim Cont As Integer
Dim K As Integer
Dim j As Integer
Dim CadSql As String
Dim CadSql2 As String
Dim ContLin As Integer
Dim CadTit As String
Dim sFormula As String
Dim TodoASoles As Boolean
Dim Salto As Boolean
Dim dRep     As New DRepFormula
Dim dBalance As New DbalanceCont
Dim nBalance As New NBalanceCont
Dim nFormula As New NInterpreteFormula
'ALPA 20090815*****************************
Dim sMatriz() As String
ReDim Preserve sMatriz(1 To 5, 300)
Dim nMichi As Integer
Dim nCadFormula1 As Currency
Dim nCadFormula2 As Currency
Dim nCadFormula3 As Currency
Dim kI As Integer
Dim cSigno As String
Dim cSigno2 As String
Dim cCtaCtesAge As String
Dim nCtaCteAhorroS As Currency
Dim nCtaCteAhorroD As Currency
Dim oBalan As DbalanceCont
'******************************************
    Screen.MousePointer = 11
    prgBarra.value = 0
    dBalance.EliminaBalanceFormaAB CInt(sTipoRepoFormula), Format(dFecha, gsFormatoMovFecha)
    dBalance.EliminaBalanceTemp CInt(sTipoRepoFormula), "0"
    dBalance.InsertaBalanceTmpSaldos CInt(sTipoRepoFormula), "0", Format(dFecha, gsFormatoFecha)
    
    Set R = dRep.CargaRepFormula(, gContBalanceFormaAB)
      If Not R.BOF And Not R.EOF Then
          Total = R.RecordCount
      End If
      Cont = 0
      ContLin = 0
      Do While Not R.EOF
         Me.StatusBar1.Panels(1) = "Proceso: " & Format(prgBarra.value, "#0.00") & "%"
         TodoASoles = True = False
         If Trim(R!cCodigo) = "210" Then
             CTemp = ""
         End If
         If Trim(R!cDescrip) <> "-" And Trim(R!cFormula) <> "" Then
             'Obtener Cuentas
             CTemp = ""
             nCuentas = 0
             ReDim Cuentas(0)
             sFormula = Trim(R!cFormula)
             'ALPA 20090819******************************************************
             'sFormula = DepuraFormula(sFormula)
             nMichi = 0
             sFormula = DepuraFormula(sFormula, nMichi)
             '*******************************************************************
             'Carga las Cuentas de la formula en una Estructura de Datos
             'Inicio
             If nMichi = 0 Then
             For K = 1 To Len(Trim(sFormula))
                 If (Mid(Trim(sFormula), K, 1) >= "0" And Mid(Trim(sFormula), K, 1) <= "9") Or (Mid(Trim(sFormula), K, 1) = "S") Or (Mid(Trim(sFormula), K, 1) = "A") Or (Mid(Trim(sFormula), K, 1) = "D") Or (Mid(Trim(sFormula), K, 1) = "[") Or (Mid(Trim(sFormula), K, 1) = "]") Or (Mid(Trim(sFormula), K, 1) = "&") Then
                     CTemp = CTemp + Mid(Trim(sFormula), K, 1)
                 Else
                     If Len(CTemp) > 0 Then
                         nCuentas = nCuentas + 1
                         ReDim Preserve Cuentas(nCuentas)
                         Cuentas(nCuentas - 1).cCta = CTemp
                         If Mid(CTemp, 1, 2) = "SA" Or Mid(CTemp, 1, 2) = "SD" Then
                             Cuentas(nCuentas - 1).cCta = DepuraSaldoAD(CTemp)
                             If Mid(CTemp, 1, 2) = "SA" Then
                                 Cuentas(nCuentas - 1).bSaldoA = True
                                 Cuentas(nCuentas - 1).bSaldoD = False
                             Else
                                 Cuentas(nCuentas - 1).bSaldoA = False
                                 Cuentas(nCuentas - 1).bSaldoD = True
                             End If
                         Else
                             Cuentas(nCuentas - 1).bSaldoA = False
                             Cuentas(nCuentas - 1).bSaldoD = False
                         End If
                         If Mid(CTemp, 1, 1) = "&" Then
                             Cuentas(nCuentas - 1).TodoASoles = True
                             Cuentas(nCuentas - 1).cCta = Mid(Cuentas(nCuentas - 1).cCta, 2, Len(Cuentas(nCuentas - 1).cCta) - 1)
                         Else
                             Cuentas(nCuentas - 1).TodoASoles = False
                         End If
                     End If
                     CTemp = ""
                 End If
             Next K
             If Len(CTemp) > 0 Then
                 nCuentas = nCuentas + 1
                 ReDim Preserve Cuentas(nCuentas)
                 Cuentas(nCuentas - 1).cCta = CTemp
                 If Mid(CTemp, 1, 2) = "SA" Or Mid(CTemp, 1, 2) = "SD" Then
                     Cuentas(nCuentas - 1).cCta = DepuraSaldoAD(CTemp)
                     If Mid(CTemp, 1, 2) = "SA" Then
                         Cuentas(nCuentas - 1).bSaldoA = True
                         Cuentas(nCuentas - 1).bSaldoD = False
                     Else
                         Cuentas(nCuentas - 1).bSaldoA = False
                         Cuentas(nCuentas - 1).bSaldoD = True
                     End If
                 Else
                     Cuentas(nCuentas - 1).bSaldoA = False
                     Cuentas(nCuentas - 1).bSaldoD = False
                 End If
                 If Mid(CTemp, 1, 1) = "&" Then
                     Cuentas(nCuentas - 1).TodoASoles = True
                     Cuentas(nCuentas - 1).cCta = Mid(Cuentas(nCuentas - 1).cCta, 2, Len(Cuentas(nCuentas - 1).cCta) - 1)
                 Else
                     Cuentas(nCuentas - 1).TodoASoles = False
                 End If
             End If
             
             'Carga Valores de las Cuentas
             For K = 0 To nCuentas - 1
                If Cuentas(K).cCta = "410804" Then
'                nSaldo41_1 = nBalance.CalculaSaldoCuentaAD(Cuentas(k).cCta, "1", Cuentas(k).bSaldoA, sTipoRepoFormula, psAgenciaCod)
                End If
                 If Cuentas(K).bSaldoA Or Cuentas(K).bSaldoD Then
                     'Moneda Nacional Historico
                     MN = nBalance.CalculaSaldoCuentaAD(Cuentas(K).cCta, "1", Cuentas(K).bSaldoA, sTipoRepoFormula, psAgenciaCod)
                     
                     'Modificado por gitu con correo del area de contabilidad 13-01-2009
                     'el monto de la cuenta 2903 debe ser positivo solo para efectos de presentacion
                     'en el Balance forma A y B
                     If MN < 0 Then
                        MN = MN * -1
                     End If
                     'End Gitu
                     
                     'Moneda Extranjera
                     MExt = nBalance.CalculaSaldoCuentaAD(Cuentas(K).cCta, "2", Cuentas(K).bSaldoA, sTipoRepoFormula, psAgenciaCod)
                     'Moneda Ajustado
                     MNAj = nBalance.CalculaSaldoCuentaAD(Cuentas(K).cCta, "6", Cuentas(K).bSaldoA, sTipoRepoFormula, psAgenciaCod)
                 Else
                     'Moneda Nacional Historico
                     MN = nBalance.CalculaSaldoCuenta(Cuentas(K).cCta, "[13]", sTipoRepoFormula, , , psAgenciaCod)
                     'Moneda Extranjera
                     MExt = nBalance.CalculaSaldoCuenta(Cuentas(K).cCta, "2", sTipoRepoFormula, , , psAgenciaCod)
                     'Moneda Nacional Ajustado
                     MNAj = nBalance.CalculaSaldoCuenta(Cuentas(K).cCta, "6", sTipoRepoFormula, , , psAgenciaCod)
                 End If
                 
                 'Actualiza Montos
                 If Cuentas(K).TodoASoles Then
                     MN = MN + MExt
                     MExt = 0
                 End If
                 Cuentas(K).nMExt = MExt
                 Cuentas(K).nMN = MN
                 Cuentas(K).nMNAj = MNAj
                 Cuentas(K).nTotal = Cuentas(K).nMN + Cuentas(K).nMExt
                 Cuentas(K).nTotAj = Cuentas(K).nTotal + Cuentas(K).nMNAj
             Next K
             
             'Genero las 3 formulas para las 3 monedas
             CTemp = ""
             CadFormula1 = ""
             CadFormula2 = ""
             CadFormula3 = ""
             For K = 1 To Len(Trim(sFormula))
               If (Mid(Trim(sFormula), K, 1) >= "0" And Mid(Trim(sFormula), K, 1) <= "9") Or (Mid(Trim(sFormula), K, 1) = "S") Or (Mid(Trim(sFormula), K, 1) = "A") Or (Mid(Trim(sFormula), K, 1) = "D") Or (Mid(Trim(sFormula), K, 1) = "[") Or (Mid(Trim(sFormula), K, 1) = "]") Or (Mid(Trim(sFormula), K, 1) = "&") Then
                  CTemp = CTemp + Mid(Trim(sFormula), K, 1)
               Else
                  If Len(CTemp) > 0 Then
                      If Mid(CTemp, 1, 1) = "&" Then
                          CTemp = Mid(CTemp, 2, Len(CTemp) - 1)
                      End If
                      cTemp1 = CTemp
                      CTemp = DepuraSaldoAD(CTemp)
                      'busca su equivalente en monto
                      For j = 0 To nCuentas
                          If Cuentas(j).cCta = CTemp And (Left(cTemp1, 1) <> "S" Or (Left(cTemp1, 2) = "SA" And Cuentas(j).bSaldoA) Or (Left(cTemp1, 2) = "SD" And Cuentas(j).bSaldoD)) Then
                              CadFormula1 = CadFormula1 + Format(Cuentas(j).nMN, "#0.00")
                              CadFormula2 = CadFormula2 + Format(Cuentas(j).nMExt, "#0.00")
                              CadFormula3 = CadFormula3 + Format(Cuentas(j).nMNAj, "#0.00")
                              Exit For
                          End If
                      Next j
                  End If
                  CTemp = ""
                  CadFormula1 = CadFormula1 + Mid(Trim(sFormula), K, 1)
                  CadFormula2 = CadFormula2 + Mid(Trim(sFormula), K, 1)
                  CadFormula3 = CadFormula3 + Mid(Trim(sFormula), K, 1)
               End If
             Next K
             If Len(CTemp) > 0 Then
                 CTemp = DepuraSaldoAD(CTemp)
                 'busca su equivalente en monto
                 For j = 0 To nCuentas
                     If Cuentas(j).cCta = CTemp Then
                         CadFormula1 = CadFormula1 + Format(Cuentas(j).nMN, "#0.00")
                         CadFormula2 = CadFormula2 + Format(Cuentas(j).nMExt, "#0.00")
                         CadFormula3 = CadFormula3 + Format(Cuentas(j).nMNAj, "#0.00")
                         Exit For
                     End If
                 Next j
'            MN = nFormula.ExprANum(CadFormula1)
'            MExt = nFormula.ExprANum(CadFormula2)
'            MNAj = nFormula.ExprANum(CadFormula3)
             End If
            'EJVG20130218 NIIFs
            MN = nFormula.ExprANum(CadFormula1)
            MExt = nFormula.ExprANum(CadFormula2)
            MNAj = nFormula.ExprANum(CadFormula3)
            'END EJVG
             Else
'             CTemp = ""
'             nCadFormula1 = 0
'             nCadFormula2 = 0
'             nCadFormula3 = 0
'             sFormula = sFormula & "+"
'              For K = 1 To Len(Trim(sFormula))
'                 If (Mid(Trim(sFormula), K, 1) >= "0" And Mid(Trim(sFormula), K, 1) <= "9") Or (Mid(Trim(sFormula), K, 1) = "S") Or (Mid(Trim(sFormula), K, 1) = "A") Or (Mid(Trim(sFormula), K, 1) = "D") Or (Mid(Trim(sFormula), K, 1) = "[") Or (Mid(Trim(sFormula), K, 1) = "]") Or (Mid(Trim(sFormula), K, 1) = "&") Then
'                     CTemp = CTemp + Mid(Trim(sFormula), K, 1)
'                 Else
'                 If Trim(CTemp) <> "" Then
'                        nCadFormula1 = nCadFormula1 + IIf(sMatriz(1, CInt(CTemp)) = "", 0, Format(sMatriz(1, CInt(CTemp)), "#0.00"))
'                        nCadFormula2 = nCadFormula2 + IIf(sMatriz(2, CInt(CTemp)) = "", 0, Format(sMatriz(2, CInt(CTemp)), "#0.00"))
'                        nCadFormula3 = nCadFormula3 + IIf(sMatriz(3, CInt(CTemp)) = "", 0, Format(sMatriz(3, CInt(CTemp)), "#0.00"))
'                    End If
'                    CTemp = ""
'                 End If
'                Next K
'            MN = nCadFormula1
'            MExt = nCadFormula2
'            MNAj = nCadFormula3
             End If
            'Fin
            
            'If R!cCodigo = "176" Or R!cCodigo = "181" Or R!cCodigo = "191" Or R!cCodigo = "199" Then
'            If R!cCodigo = "176" Or R!cCodigo = "179" Or R!cCodigo = "181" Or R!cCodigo = "191" Or R!cCodigo = "199" Or R!cCodigo = "203" Then
'                If R!cCodigo = "176" Or R!cCodigo = "181" Or R!cCodigo = "179" Then
'                    If R!cCodigo = "176" Then
'                        If MN + MExt < 0 Then
'                            MN = 0
'                            MExt = 0
'                        Else
'                            MExt = MExt * -1
'                            MN = MN * -1
'                        End If
'                    Else
'                        If MN < 0 Then
'                           MN = 0
'                        End If
'                        If MExt < 0 Then
'                           MExt = 0
'                        End If
'                        If MNAj < 0 Then
'                           MNAj = 0
'                        End If
'                    End If
'                Else
'                    If R!cCodigo = "191" Then
'                    If MN + MExt > 0 Then
'                            MN = 0
'                            MExt = 0
'                    Else
'                            MExt = MExt * -1
'                            MN = MN * -1
'                    End If
'                    Else
'                        If MN < 0 Then
'                           MN = MN * -1
'                        Else
'                            MN = 0
'                        End If
'                        If MExt < 0 Then
'                           MExt = MExt * -1
'                        Else
'                            MExt = 0
'                        End If
'                        If MNAj < 0 Then
'                           MNAj = MNAj * -1
'                        Else
'                            MNAj = 0
'                        End If
'
'
'                    End If
'                End If
'            End If
'            'ALPA 20100322**************************************
'             If R!cCodigo = 7 Or R!cCodigo = 3 Then
'                nCtaCteAhorroS = 0
'                nCtaCteAhorroD = 0
'                Set oBalan = New DbalanceCont
'                nCtaCteAhorroS = oBalan.ObtenerDatosCtaCtesAhorroAyB(txtAnio.Text & Format(cboMes.ListIndex + 1, "00"), "461011", psAgenciaCod)
'                nCtaCteAhorroD = gnTipCambio * oBalan.ObtenerDatosCtaCtesAhorroAyB(txtAnio.Text & Format(cboMes.ListIndex + 1, "00"), "462011", psAgenciaCod)
'                MN = MN + nCtaCteAhorroS
'                MExt = MExt + nCtaCteAhorroD
'            End If
            '***************************************************
            'EJVG20130219 ***
            If R!cCodigo = 57 Then
                If MN + MExt < 0 Then
                    MN = 0
                    MExt = 0
                End If
            ElseIf R!cCodigo = 102 Then
                If MN + MExt >= 0 Then
                    MN = 0
                    MExt = 0
                End If
            End If
            'END EJVG *******
            sMatriz(1, CInt(R!cCodigo)) = MN
            sMatriz(2, CInt(R!cCodigo)) = MExt
            sMatriz(3, CInt(R!cCodigo)) = MNAj
            If nMichi = 0 Then
                sMatriz(4, CInt(R!cCodigo)) = 0
            Else
                sMatriz(4, CInt(R!cCodigo)) = 1
            End If
             sMatriz(5, CInt(R!cCodigo)) = sFormula
            If nMichi = 0 Then
                dBalance.InsertaBalanceGen R!cCodigo, MN, MExt, MN + MExt, MN + MExt + MNAj, Format(dFecha, gsFormatoMovFecha), sTipoRepoFormula
                Cont = Cont + 1
                prgBarra.value = (Cont / Total) * 100
            End If
         Else
             dBalance.InsertaBalanceGen R!cCodigo, 0, 0, 0, 0, Format(dFecha, gsFormatoMovFecha), sTipoRepoFormula
             Cont = Cont + 1
             prgBarra.value = (Cont / Total) * 100
         End If
         R.MoveNext
      Loop
      For K = 1 To 300
        If IIf(sMatriz(4, K) = "", "0", sMatriz(4, K)) = "1" Then
            CTemp = ""
             nCadFormula1 = 0
             nCadFormula2 = 0
             nCadFormula3 = 0
             sFormula = sMatriz(5, K) & "+"
              For kI = 1 To Len(Trim(sFormula))
                 If (Mid(Trim(sFormula), kI, 1) >= "0" And Mid(Trim(sFormula), kI, 1) <= "9") Or (Mid(Trim(sFormula), kI, 1) = "S") Or (Mid(Trim(sFormula), kI, 1) = "A") Or (Mid(Trim(sFormula), kI, 1) = "D") Or (Mid(Trim(sFormula), kI, 1) = "[") Or (Mid(Trim(sFormula), kI, 1) = "]") Or (Mid(Trim(sFormula), kI, 1) = "&") Then
                     CTemp = CTemp + Mid(Trim(sFormula), kI, 1)
'                 ElseIf Mid(Trim(sFormula), kI, 1) = "+" Or Mid(Trim(sFormula), kI, 1) = "-" Then
                    
                 Else
                 cSigno = Mid(Trim(sFormula), kI, 1)
                 If Trim(CTemp) <> "" Then
                    If (nCadFormula1 + nCadFormula2) = 0 Then
                       nCadFormula1 = IIf(sMatriz(1, CTemp) = "", 0, Format(sMatriz(1, CTemp), "#0.00"))
                       nCadFormula2 = IIf(sMatriz(2, CTemp) = "", 0, Format(sMatriz(2, CTemp), "#0.00"))
                       nCadFormula3 = IIf(sMatriz(3, CTemp) = "", 0, Format(sMatriz(3, CTemp), "#0.00"))
                    Else
                    If cSigno2 = "+" Then
                           nCadFormula1 = nCadFormula1 + IIf(sMatriz(1, CTemp) = "", 0, Format(sMatriz(1, CTemp), "#0.00"))
                           nCadFormula2 = nCadFormula2 + IIf(sMatriz(2, CTemp) = "", 0, Format(sMatriz(2, CTemp), "#0.00"))
                           nCadFormula3 = nCadFormula3 + IIf(sMatriz(3, CTemp) = "", 0, Format(sMatriz(3, CTemp), "#0.00"))
                       ElseIf cSigno2 = "-" Then
                           nCadFormula1 = nCadFormula1 - IIf(sMatriz(1, CTemp) = "", 0, Format(sMatriz(1, CTemp), "#0.00"))
                           nCadFormula2 = nCadFormula2 - IIf(sMatriz(2, CTemp) = "", 0, Format(sMatriz(2, CTemp), "#0.00"))
                           nCadFormula3 = nCadFormula3 - IIf(sMatriz(3, CTemp) = "", 0, Format(sMatriz(3, CTemp), "#0.00"))
                    End If
                    End If
                    End If
                    CTemp = ""
                    cSigno2 = cSigno
                    cSigno = ""
                 End If
                Next kI
            MN = nCadFormula1
            MExt = nCadFormula2
            MNAj = nCadFormula3
            sMatriz(1, K) = MN
            sMatriz(2, K) = MExt
            sMatriz(3, K) = MNAj
            sMatriz(4, K) = 1
            'sMatriz(5, CInt(R!cCodigo))
            'ALPA 20090930***********************************************
              dBalance.InsertaBalanceGen K, MN, MExt, MN + MExt, MN + MExt + MNAj, Format(dFecha, gsFormatoMovFecha), sTipoRepoFormula
             'dBalance.InsertaBalanceGen K, sMatriz(1, K), sMatriz(2, K), sMatriz(1, K) + sMatriz(2, K), sMatriz(1, K) + sMatriz(2, K) + sMatriz(3, K), Format(dFecha, gsFormatoMovFecha), sTipoRepoFormula
             '************************************************************
             Cont = Cont + 1
             prgBarra.value = (Cont / Total) * 100
        End If
      Next K
    Me.StatusBar1.Panels(1) = "Proceso: " & Format(prgBarra.value, "#0.00") & "%"
    Salto = False
    Screen.MousePointer = 0
R.Close
Set R = Nothing
Set dRep = Nothing
Set dBalance = Nothing
Set nBalance = Nothing
Set nFormula = Nothing
Me.StatusBar1.Panels(1) = "Proceso Terminado"
End Sub

Private Sub cboMes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If optMoneda(0).value Then
      optMoneda(0).SetFocus
   Else
      optMoneda(1).SetFocus
   End If
End If
End Sub

Private Sub chkAG_Click()
    Me.txtAG.Text = ""
    Me.lblAG.Caption = ""
End Sub

Private Sub cmdProcesar_Click()
Dim sImpre As String
Dim nResp  As VbMsgBoxResult
Dim oBalance As New NBalanceCont

If Me.chkAG.value = 0 And Me.txtAG.Text = "" Then
    MsgBox "Debe elegir una agencia.", vbInformation, "Aviso"
    Me.txtAG.SetFocus
    Exit Sub
End If

If Not ValidaDatos() Then
   Exit Sub
End If
dFecha = DateAdd("m", 1, "01/" & Format(cboMes.ListIndex + 1, "00") & "/" & txtAnio) - 1
If oBalance.BalanceGeneradoFormaAB(sTipoRepoFormula, Format(dFecha, gsFormatoMovFecha)) Then
   nResp = MsgBox("El Balance General para este Mes ya ha sido generado, ¿ Desea Volver ha procesarlo ? ", vbInformation + vbYesNoCancel, "Aviso")
   If nResp = vbCancel Then
      Exit Sub
   End If
   If nResp = vbYes Then
      Call procesar(IIf(Me.chkAG.value = 1, "", Me.txtAG.Text))
   End If
Else
   Call procesar(IIf(Me.chkAG.value = 1, "", Me.txtAG.Text))
End If
Screen.MousePointer = 11

Call BalanceGeneralAB(gContBalanceFormaAB, dFecha, sTipoRepoFormula, optMoneda(0).value, UCase(gsNomCmac))

'CadImp = oBalance.ImprimeBalanceFormaAB(gContBalanceFormaAB, dFecha, sTipoRepoFormula, optMoneda(0).value, UCase(gsNomCmac))
Set oBalance = Nothing
Screen.MousePointer = 0
'EnviaPrevio CadImp, "Balance General Forma A y B", gnLinPage, True
'CadImp = ""

End Sub

Public Function BalanceGeneralAB(ByVal psContBalanceFormaAB As String, ByVal dFecha As Date, ByVal psTipoRepoFormula As String, ByVal sMoneda As String, ByVal psNomCmac As String) As Boolean
    Dim liLineas As Integer
    Dim i As Integer
    
    Dim xlAplicacion As Excel.Application
    Dim xlLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet

    Dim lbConexion As Boolean
    Dim lbExisteHoja  As Boolean
    Dim lsNomHoja As String
    Dim glsarchivo As String
    
     
    glsarchivo = "ProcesoValidacion" & Format(Now, "yyyymmdd") & "_" & Format(Time(), "HHMMSS") & IIf(Me.chkAG.value = 1, "", "_AG" & Me.txtAG.Text) & ".XLS"

    Set xlAplicacion = New Excel.Application
    Set xlLibro = xlAplicacion.Workbooks.Add
    Set xlHoja1 = xlLibro.Worksheets.Add

    xlHoja1.PageSetup.CenterHorizontally = True
    xlHoja1.PageSetup.Zoom = 60
    xlHoja1.PageSetup.Orientation = xlLandscape
    
    Call ImprimeBalanceFormaABExcel(xlLibro, xlHoja1, xlAplicacion, psContBalanceFormaAB, dFecha, psTipoRepoFormula, sMoneda, psNomCmac, Me.txtAG.Text)
            
    xlHoja1.SaveAs App.path & "\SPOOLER\" & glsarchivo
    ExcelEnd App.path & "\Spooler\" & glsarchivo, xlAplicacion, xlLibro, xlHoja1
        'Cierra el libro de trabajo
    'xlLibro.Close
        ' Cierra Microsoft Excel con el método Quit.
    'xlAplicacion.Quit
        'Libera los objetos.
   'Set xlAplicacion = Nothing
    'Set xlLibro = Nothing
    MsgBox "Se ha generado el Archivo en " & App.path & "\SPOOLER\" & glsarchivo, vbInformation, "Aviso"
    'Call CargaArchivo(glsArchivo, App.Path & "\SPOOLER\")
    CargaArchivo App.path & "\SPOOLER\" & glsarchivo, App.path & "\SPOOLER\"
 
End Function




Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
CentraForm Me
txtAnio = Year(gdFecSis)
cboMes.ListIndex = Month(gdFecSis) - 1

sTipoRepoFormula = "3"

Dim oRHAreas As DActualizaDatosArea
Set oRHAreas = New DActualizaDatosArea

 txtAG.rs = oRHAreas.GetAgencias()

End Sub

Private Sub optMoneda_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
   cmdProcesar.SetFocus
End If
End Sub


Private Sub txtAG_EmiteDatos()
    Me.lblAG.Caption = Me.txtAG.psDescripcion
End Sub

Private Sub txtAnio_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
If KeyAscii = 13 Then
   If Not ValidaAnio(txtAnio) Then
      Exit Sub
   End If
   cboMes.SetFocus
End If
End Sub

Private Sub txtAnio_Validate(Cancel As Boolean)
   If Not ValidaAnio(txtAnio) Then
      Cancel = True
   End If
End Sub

Private Function ValidaDatos() As Boolean
ValidaDatos = False
   If Not ValidaAnio(txtAnio) Then
      Exit Function
   End If
ValidaDatos = True
End Function
'EJVG20130219 ***
Public Function ObtenerUtilidadEjercicio(pdFecha As Date, psMoneda As String, psAgenciaCod As String) As Currency
On Error GoTo ErrObtenerUtilidadEjercicio
    Dim oBal As New NBalanceCont
    Dim n5 As Currency
    Dim n62 As Currency
    Dim n63 As Currency
    Dim n64 As Currency
    Dim n65 As Currency
    Dim n4 As Currency
    Dim nUtilidad As Currency
    Dim Moneda As String
    
    'Moneda = pnMoneda
    'n5 = oBal.getImporteBalanceMes("5", 1, Moneda, Month(pdFecha), Year(pdFecha))
    n5 = oBal.CalculaSaldoCuenta("5", psMoneda, sTipoRepoFormula, , , psAgenciaCod)
    'n62 = oBal.getImporteBalanceMes("62", 1, Moneda, Month(pdFecha), Year(pdFecha))
    n62 = oBal.CalculaSaldoCuenta("62", psMoneda, sTipoRepoFormula, , , psAgenciaCod)
    'n63 = oBal.getImporteBalanceMes("63", 1, Moneda, Month(pdFecha), Year(pdFecha))
    n63 = oBal.CalculaSaldoCuenta("63", psMoneda, sTipoRepoFormula, , , psAgenciaCod)
    'n64 = oBal.getImporteBalanceMes("64", 1, Moneda, Month(pdFecha), Year(pdFecha))
    n64 = oBal.CalculaSaldoCuenta("64", psMoneda, sTipoRepoFormula, , , psAgenciaCod)
    'n65 = oBal.getImporteBalanceMes("65", 1, Moneda, Month(pdFecha), Year(pdFecha))
    n65 = oBal.CalculaSaldoCuenta("65", psMoneda, sTipoRepoFormula, , , psAgenciaCod)
    'n4 = oBal.getImporteBalanceMes("4", 1, Moneda, Month(pdFecha), Year(pdFecha))
    n4 = oBal.CalculaSaldoCuenta("4", psMoneda, sTipoRepoFormula, , , psAgenciaCod)
   
    nUtilidad = n5 + n62 + n64 - (n4 + n63 + n65)
    nUtilidad = nUtilidad - nUtilidad * 0.3
    ObtenerUtilidadEjercicio = nUtilidad
Exit Function
ErrObtenerUtilidadEjercicio:
   Call RaiseError(MyUnhandledError, "ObtenerUtilidadEjercicio Method")
End Function
'END EJVG *******
