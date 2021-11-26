VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRepBaseFormula 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reportes en Base a Fórmulas: "
   ClientHeight    =   6255
   ClientLeft      =   900
   ClientTop       =   2265
   ClientWidth     =   9750
   Icon            =   "frmRepBaseFormula.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   9750
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog cmdAbrir 
      Left            =   3975
      Top             =   5760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Busca Archivo ..."
   End
   Begin Sicmact.ProgressBarra prgBarra 
      Height          =   120
      Left            =   1395
      TabIndex        =   11
      Top             =   5760
      Width           =   2325
      _ExtentX        =   4101
      _ExtentY        =   212
   End
   Begin VB.Frame FraGenera 
      Height          =   855
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   9585
      Begin VB.CommandButton CmdProceso 
         Caption         =   "&Generar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8160
         TabIndex        =   0
         Top             =   240
         Width           =   1290
      End
      Begin MSMask.MaskEdBox txtFechaCorte 
         Height          =   345
         Left            =   6720
         TabIndex        =   15
         Top             =   240
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   609
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblMes 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   315
         Left            =   3360
         TabIndex        =   19
         Top             =   480
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label lblAnio 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   315
         Left            =   4560
         TabIndex        =   18
         Top             =   480
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label lblPatrimonio 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   315
         Left            =   2040
         TabIndex        =   17
         Top             =   480
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.Label lblTituloPatrimonio 
         Caption         =   "Patrimonio Efectivo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   480
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label lblFechaCorte 
         Caption         =   "Fecha corte"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5640
         TabIndex        =   14
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label txtMes 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   315
         Left            =   2010
         TabIndex        =   10
         Top             =   120
         Width           =   2325
      End
      Begin VB.Label txtAnio 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   315
         Left            =   990
         TabIndex        =   9
         Top             =   120
         Width           =   1005
      End
      Begin VB.Label lblReporte 
         Caption         =   "Periodo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   210
         TabIndex        =   8
         Top             =   180
         Width           =   765
      End
   End
   Begin VB.Frame fraLista 
      Height          =   4845
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   9585
      Begin VB.CommandButton cmdPatrimonio 
         Caption         =   "&Patrimonio"
         Height          =   375
         Left            =   2640
         TabIndex        =   20
         Top             =   4320
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdReporte 
         Caption         =   "&Reporte"
         Height          =   375
         Left            =   3960
         TabIndex        =   13
         Top             =   4320
         Width           =   1215
      End
      Begin VB.CommandButton cmdHoja 
         Caption         =   "&Hoja Trabajo"
         Height          =   375
         Left            =   6720
         TabIndex        =   12
         Top             =   4320
         Width           =   1290
      End
      Begin VB.CommandButton cmdOpciones 
         Caption         =   "&Opciones de Impresión..."
         Height          =   375
         Left            =   180
         TabIndex        =   2
         Top             =   4320
         Width           =   2130
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "&Imprimir"
         Height          =   375
         Left            =   5280
         TabIndex        =   3
         Top             =   4320
         Width           =   1290
      End
      Begin MSComctlLib.ListView lvRep 
         Height          =   4005
         Left            =   150
         TabIndex        =   1
         Top             =   240
         Width           =   9285
         _ExtentX        =   16378
         _ExtentY        =   7064
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripcion de Cuenta"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Consilidado H."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Consolidado A."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "M.Nacional H."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "M.Nacional Ajustado"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "M.Extranjera MN"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "M.Extranjera ME"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Ing Manual"
            Object.Width           =   176
         EndProperty
      End
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   375
         Left            =   8160
         TabIndex        =   4
         Top             =   4320
         Width           =   1320
      End
   End
   Begin VB.OLE OleExcel 
      AutoActivate    =   3  'Automatic
      Enabled         =   0   'False
      Height          =   405
      Left            =   2520
      SizeMode        =   1  'Stretch
      TabIndex        =   7
      Top             =   5280
      Visible         =   0   'False
      Width           =   825
   End
End
Attribute VB_Name = "frmRepBaseFormula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Option Base 0
Private Type TEstBal
    cCodCta  As String
    cDescrip As String
    cFormula As String
    bIngManual As Boolean
End Type
Private Type TCuentas
    cCta    As String
    nCOAj   As Double
    nCOHist As Double
    nMNAj   As Double
    nMNHist As Double
    nMES    As Double
    nMED    As Double
    cDescrip As String
    bIngManual As Boolean
    '***Agregado por ELRO el 20111224, según Acta N° 352-2011/TI-D
    nMEDAnterior As Double
    nDiferencia As Double
    '***Fin Agregado por ELRO*************************************
End Type
Dim EstBal() As TEstBal
Dim nContBal As Integer
'ALPA 20090512***************************
Dim EstBalReporte() As TEstBal
Dim nContBalReporte As Integer
Dim sCodOpeReporte As String
Dim CuentasReporte() As TCuentas
Dim MatrixReporte() As TCuentas
Dim nCuentasReporte As Integer
'****************************************
Dim Cuentas() As TCuentas
Dim nCuentas As Integer
Dim dFecha As Date
Dim sSql As String
Dim R As New ADODB.Recordset

Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet

Dim lCOAj As Integer
Dim lCOHi As Integer
Dim lMNAj As Integer
Dim lMNHi As Integer
Dim lMES  As Integer
Dim lMED  As Integer
Dim sTipoRepoFormula As String
Dim sTituRepoFormula As String
'***Agregado por ELRO el 20111223, según Acta N° 352-2011/TI-D
Private fnMEDAnterior As Integer
Private fnDiferencia As Integer
Private fnSaldoActivoMN As Currency
Private fnSaldoActivoME As Currency
Private fnSaldoPasivoMN As Currency
Private fnSaldoPasivoME As Currency
'***Fin Agregado por ELRO*************************************

Dim oNBal  As NBalanceCont
Dim oDBal  As DbalanceCont
Dim lnAnio As Integer
Dim lnMes  As Integer
Dim lsRepCod  As String
Dim lsRepDesc As String

Dim fsCodReport As String
Dim lsFormaOrdenRep As String
Dim lbLimpiaDescrip As String

Public Sub Inicio(pnAnio As Integer, pnMes As Integer, psRepCod As String, psRepDesc As String, Optional psFormaOrdenRep As String = "INT", Optional pbLimpiaDescrip As Boolean = True)
    lnAnio = pnAnio
    lnMes = pnMes
    lsRepCod = psRepCod
    fsCodReport = psRepCod
    lsRepDesc = psRepDesc
    lsFormaOrdenRep = psFormaOrdenRep
    lbLimpiaDescrip = pbLimpiaDescrip
    Me.Show 1
End Sub

Private Function DepuraEquivalentes(psEquival As String) As String
Dim j As Integer
Dim CadTemp As String
   CadTemp = ""
   For j = 1 To Len(psEquival)
       If Mid(psEquival, j, 1) <> "." Then
           CadTemp = CadTemp + Mid(psEquival, j, 1)
       End If
   Next j
   DepuraEquivalentes = CadTemp
End Function

Private Function DepuraFormula(sFormula As String) As String
Dim sCad As String
Dim R As New ADODB.Recordset
Dim sSql As String
Dim i As Integer
Dim sCadRes As String
Dim bFinal As Boolean
Dim sCod As String
    sCad = sFormula
    i = 1
    sCadRes = ""
    Do While i <= Len(sCad)
        If Mid(sCad, i, 1) <> "#" Then
            sCadRes = sCadRes + Mid(sCad, i, 1)
        Else
            i = i + 2
            bFinal = False
            sCod = ""
            Do While Not bFinal
                If Mid(sCad, i, 1) <> "]" Then
                    sCod = sCod + Mid(sCad, i, 1)
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
    DepuraFormula = sCadRes
End Function
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
    Set R = oRepFormula.CargaRepFormula(sCodigo, gsOpeCod)
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

Private Sub CargaDatos()
    Dim oRep As New DRepFormula
    Dim R    As New ADODB.Recordset
    Dim nReg As Integer
    nContBal = 0
    ReDim EstBal(0)
        Set R = oRep.CargaRepFormula(, gsOpeCod, , , , lsFormaOrdenRep)
          ProgressShow prgBarra, Me, eCap_CaptionPercent
          
          prgBarra.Titulo = sTituRepoFormula
          prgBarra.SubTitulo = "Cargando datos... "
          prgBarra.Max = R.RecordCount
          Do While Not R.EOF
              nContBal = nContBal + 1
              ReDim Preserve EstBal(nContBal)
              EstBal(nContBal - 1).cCodCta = Trim(R!cCodigo)
              If lbLimpiaDescrip Then
                EstBal(nContBal - 1).cDescrip = Trim(R!cDescrip)
              Else
                EstBal(nContBal - 1).cDescrip = R!cDescrip
              End If
              EstBal(nContBal - 1).cFormula = DepuraEquivalentes(Trim(R!cFormula))
              EstBal(nContBal - 1).bIngManual = IIf(R!bIngresoManual, True, False)
              R.MoveNext
              prgBarra.Progress nContBal
          Loop
          ProgressClose prgBarra, Me
        RSClose R
        Set oRep = Nothing
    Set R = Nothing
End Sub

Private Sub GeneraReporte()
Dim i As Integer
Dim K As Integer
Dim j As Integer
Dim CTemp As String
Dim sSql As String
Dim sSql2 As String
Dim R As New ADODB.Recordset
Dim CoHist As Double
Dim CoAj   As Double
Dim MNAj   As Double
Dim MNHist As Double
Dim MED    As Double
Dim MES    As Double
Dim CadSql As String
Dim CadFormula1 As String
Dim CadFormula2 As String
Dim CadFormula3 As String
Dim CadFormula4 As String
Dim CadFormula5 As String
Dim CadFormula6 As String
Dim L As ListItem
Dim nFormula As New NInterpreteFormula
Dim nImporte As Currency
Dim nTipC As Currency
'***Agregado por ELRo el 20111226, según Acta 352-2011/TI-D
Dim lsMovNro As String
Dim oNContFunciones As NContFunciones
Set oNContFunciones = New NContFunciones
'***Fin Agregado por ELRO**********************************

   ProgressShow prgBarra, Me, ePCap_CaptionPercent
   prgBarra.Titulo = sTituRepoFormula
   prgBarra.SubTitulo = "Cargando datos... "
   prgBarra.Max = nContBal - 1
   DoEvents
   'MIOL 20120705, SEGUN RQ12122 ***********************
   If sCodOpeReporte <> "770091" And sCodOpeReporte <> "763412" Then
'   If sCodOpeReporte <> "770091" Then
   'END MIOL *******************************************
        dFecha = DateAdd("m", 1, CDate("01/" & Format(lnMes, "00") & "/" & Format(lnAnio, "0000"))) - 1
   Else
        dFecha = txtFechaCorte.Text
   End If
   oDBal.EliminaBalanceTemp CInt(sTipoRepoFormula), "0" '** para el a es 1

   oDBal.InsertaBalanceTmpSaldos CInt(sTipoRepoFormula), "0", Format(dFecha, gsFormatoFecha)
    '***Agregado por ELRO el 20111222, según Acta N° 352-2011/TI-D
    'MIOL 20120705, SEGUN RQ 12122 *******************************
    If fsCodReport = "770090" Or fsCodReport = "763411" Then
'    If fsCodReport = "770090" Then
    'END MIOL ****************************************************
        lsMovNro = oNContFunciones.GeneraMovNro(CDate(txtFechaCorte), gsCodAge, gsCodUser)
        Call oDBal.eliminarBalanceTempHist(txtFechaCorte)
        Call oDBal.InsertaBalanceTempHistSaldos(lsMovNro)
    End If
    Set oNContFunciones = Nothing
    '***Fin Agregado por ELRO*************************************
   nTipC = oNBal.GetTipCambioBalance(Format(dFecha, gsFormatoMovFecha))
   lvRep.ListItems.Clear
   For i = 0 To nContBal - 1
       
      'Obtener Cuentas
      CTemp = ""
      nCuentas = 0
      EstBal(i).cFormula = DepuraFormula(EstBal(i).cFormula)
      
      ReDim Cuentas(0)
      For K = 1 To Len(EstBal(i).cFormula)
          If UCase(Mid(EstBal(i).cFormula, K, 3)) = "DBO" Then
               CTemp = CTemp + Left(Mid(EstBal(i).cFormula, K, InStr(1, Mid(EstBal(i).cFormula, K), ")")), 3) & "." & Mid(Mid(EstBal(i).cFormula, K, InStr(1, Mid(EstBal(i).cFormula, K), ")")), 4)
               K = K + InStr(1, Mid(EstBal(i).cFormula, K), ")") - 1
          ElseIf Mid(EstBal(i).cFormula, K, 1) >= "0" And Mid(EstBal(i).cFormula, K, 1) <= "9" Then
              CTemp = CTemp + Mid(EstBal(i).cFormula, K, 1)
          Else
              If Len(CTemp) > 0 Then
                  nCuentas = nCuentas + 1
                  ReDim Preserve Cuentas(nCuentas)
                  Cuentas(nCuentas - 1).cCta = CTemp
              End If
              CTemp = ""
          End If
      Next K
      If Len(CTemp) > 0 Then
          nCuentas = nCuentas + 1
          ReDim Preserve Cuentas(nCuentas)
          Cuentas(nCuentas - 1).cCta = CTemp
      End If
      'Carga Valres de las Cuentas
      For K = 0 To nCuentas - 1
        If UCase(Left(Cuentas(K).cCta, 4)) = "DBO." Then
           'Moneda Nacional Historico
           MNHist = nFormula.EjecutaFuncion(Cuentas(K).cCta)
           'Moneda Extranjera Soles
           MES = nFormula.EjecutaFuncion(Cuentas(K).cCta)
           'Moneda Nacional Ajustado
           MNAj = nFormula.EjecutaFuncion(Cuentas(K).cCta)
           'Consolidado Historico
           CoHist = nFormula.EjecutaFuncion(Cuentas(K).cCta)
           'Consolidado Ajustado
           CoAj = nFormula.EjecutaFuncion(Cuentas(K).cCta)
        Else
           'Moneda Nacional Historico
           MNHist = oNBal.CalculaSaldoCuenta(Cuentas(K).cCta, "[13]", sTipoRepoFormula, "0")
           'Moneda Extranjera Soles
           MES = oNBal.CalculaSaldoCuenta(Cuentas(K).cCta, "2", sTipoRepoFormula, "0")
           'Moneda Nacional Ajustado
           MNAj = oNBal.CalculaSaldoCuenta(Cuentas(K).cCta, "[136]", sTipoRepoFormula, "0")
           'Consolidado Historico
           CoHist = oNBal.CalculaSaldoCuenta(Cuentas(K).cCta, "[123]", sTipoRepoFormula, "0")
           'Consolidado Ajustado
           CoAj = oNBal.CalculaSaldoCuenta(Cuentas(K).cCta, "[1236]", sTipoRepoFormula, "0")
           
           If nTipC <> 0 Then
              MED = Round(MES / nTipC, 2)
           End If
        End If
        
        'Actualiza Montos
        Cuentas(K).nMES = MES
        Cuentas(K).nMED = MED
        Cuentas(K).nMNAj = MNAj
        Cuentas(K).nMNHist = MNHist
        Cuentas(K).nCOAj = CoAj
        Cuentas(K).nCOHist = CoHist
      Next K
      'Genero las 3 formulas para las 3 monedas
      CTemp = ""
      CadFormula1 = ""
      CadFormula2 = ""
      CadFormula3 = ""
      CadFormula4 = ""
      CadFormula5 = ""
      CadFormula6 = ""
                  
      'UCase(Left(EstBal(I).cFormula, 3)) = "DBO"
'            nImporte = nFormula.EjecutaFuncion(Left(EstBal(I).cFormula, 3) & "." & Mid(EstBal(I).cFormula, 4))
'            CadFormula1 = CadFormula1 + Format(nImporte, gsFormatoNumeroDato)
'            CadFormula2 = CadFormula2 + Format(nImporte, gsFormatoNumeroDato)
'            CadFormula3 = CadFormula3 + Format(nImporte, gsFormatoNumeroDato)
'            CadFormula4 = CadFormula4 + Format(nImporte, gsFormatoNumeroDato)
'            CadFormula5 = CadFormula5 + Format(nImporte, gsFormatoNumeroDato)
'            CadFormula6 = CadFormula6 + Format(nImporte, gsFormatoNumeroDato)

        For K = 1 To Len(EstBal(i).cFormula)
            If UCase(Mid(EstBal(i).cFormula, K, 3)) = "DBO" Then
                 CTemp = CTemp + Left(Mid(EstBal(i).cFormula, K, InStr(1, Mid(EstBal(i).cFormula, K), ")")), 3) & "." & Mid(Mid(EstBal(i).cFormula, K, InStr(1, Mid(EstBal(i).cFormula, K), ")")), 4)
                 K = K + InStr(1, Mid(EstBal(i).cFormula, K), ")") - 1
            ElseIf (Mid(EstBal(i).cFormula, K, 1) >= "0" And Mid(EstBal(i).cFormula, K, 1) <= "9") Or (Mid(EstBal(i).cFormula, K, 1) = ".") Then
                CTemp = CTemp + Mid(EstBal(i).cFormula, K, 1)
            ElseIf Mid(EstBal(i).cFormula, K, 1) = "$" Then
                 j = InStr(K + 1, EstBal(i).cFormula, "$") - 1
                 nImporte = Mid(EstBal(i).cFormula, K + 1, j - K)
                 CadFormula1 = CadFormula1 + Format(nImporte, gsFormatoNumeroDato)
                 CadFormula2 = CadFormula2 + Format(nImporte, gsFormatoNumeroDato)
                 CadFormula3 = CadFormula3 + Format(nImporte, gsFormatoNumeroDato)
                 CadFormula4 = CadFormula4 + Format(nImporte, gsFormatoNumeroDato)
                 CadFormula5 = CadFormula5 + Format(nImporte, gsFormatoNumeroDato)
                 CadFormula6 = CadFormula6 + Format(nImporte, gsFormatoNumeroDato)
                 K = j + 1
            Else
                 If Len(CTemp) > 0 Then
                     'busca su equivalente en monto
                     For j = 0 To nCuentas
                         If Cuentas(j).cCta = CTemp Then
                             CadFormula1 = CadFormula1 + Format(Cuentas(j).nMNAj, gsFormatoNumeroDato)
                             CadFormula2 = CadFormula2 + Format(Cuentas(j).nMNHist, gsFormatoNumeroDato)
                             CadFormula3 = CadFormula3 + Format(Cuentas(j).nMES, gsFormatoNumeroDato)
                             CadFormula4 = CadFormula4 + Format(Cuentas(j).nMED, gsFormatoNumeroDato)
                             CadFormula5 = CadFormula5 + Format(Cuentas(j).nCOAj, gsFormatoNumeroDato)
                             CadFormula6 = CadFormula6 + Format(Cuentas(j).nCOHist, gsFormatoNumeroDato)
                             Exit For
                         End If
                     Next j
                 End If
                 CTemp = ""
                 CadFormula1 = CadFormula1 + Mid(EstBal(i).cFormula, K, 1)
                 CadFormula2 = CadFormula2 + Mid(EstBal(i).cFormula, K, 1)
                 CadFormula3 = CadFormula3 + Mid(EstBal(i).cFormula, K, 1)
                 CadFormula4 = CadFormula4 + Mid(EstBal(i).cFormula, K, 1)
                 CadFormula5 = CadFormula5 + Mid(EstBal(i).cFormula, K, 1)
                 CadFormula6 = CadFormula6 + Mid(EstBal(i).cFormula, K, 1)
            End If
        Next K
         
         If Len(CTemp) > 0 Then
             'busca su equivalente en monto
             For j = 0 To nCuentas
                 If Cuentas(j).cCta = CTemp Then
                     CadFormula1 = CadFormula1 + Format(Cuentas(j).nMNAj, gsFormatoNumeroDato)
                     CadFormula2 = CadFormula2 + Format(Cuentas(j).nMNHist, gsFormatoNumeroDato)
                     CadFormula3 = CadFormula3 + Format(Cuentas(j).nMES, gsFormatoNumeroDato)
                     CadFormula4 = CadFormula4 + Format(Cuentas(j).nMED, gsFormatoNumeroDato)
                     CadFormula5 = CadFormula5 + Format(Cuentas(j).nCOAj, gsFormatoNumeroDato)
                     CadFormula6 = CadFormula6 + Format(Cuentas(j).nCOHist, gsFormatoNumeroDato)
                     Exit For
                 End If
             Next j
         End If
         MNAj = 0: MNHist = 0: MES = 0: MED = 0: CoAj = 0: CoHist = 0
         
         MNAj = nFormula.ExprANum(CadFormula1, EstBal(i).cCodCta)
         MNHist = nFormula.ExprANum(CadFormula2, EstBal(i).cCodCta)
         MES = nFormula.ExprANum(CadFormula3, EstBal(i).cCodCta)
         MED = nFormula.ExprANum(CadFormula4, EstBal(i).cCodCta)
         CoAj = nFormula.ExprANum(CadFormula5, EstBal(i).cCodCta)
         CoHist = nFormula.ExprANum(CadFormula6, EstBal(i).cCodCta)
    
        'Asignacion de Valores
         Set L = lvRep.ListItems.Add(, , EstBal(i).cCodCta)
         L.SubItems(1) = EstBal(i).cDescrip
         L.SubItems(2) = Format(CoHist, gsFormatoNumeroView)
         L.SubItems(3) = Format(CoAj, gsFormatoNumeroView)
         L.SubItems(4) = Format(MNHist, gsFormatoNumeroView)
         L.SubItems(5) = Format(MNAj, gsFormatoNumeroView)
         L.SubItems(6) = Format(MES, gsFormatoNumeroView)
         L.SubItems(7) = Format(MED, gsFormatoNumeroView)
         L.SubItems(8) = IIf(EstBal(i).bIngManual, "1", "0")
      
      prgBarra.SubTitulo = "Intepretando Fórmulas... "
      prgBarra.Progress i
   Next i
   ProgressClose prgBarra, Me
Set nFormula = Nothing
End Sub

Private Sub GenRepResumenContCartera()
Dim nCol  As Integer
Dim sCol  As String

Dim lsArchivo   As String
Dim lbLibroOpen As Boolean
Dim N           As Integer

lCOAj = frmRepBaseFormulaOpc.plCOAj
lCOHi = frmRepBaseFormulaOpc.plCOHist
lMNAj = frmRepBaseFormulaOpc.plMNAj
lMNHi = frmRepBaseFormulaOpc.plMNHist
lMES = frmRepBaseFormulaOpc.plMES
lMED = frmRepBaseFormulaOpc.plMED
 

On Error GoTo ErrImprime
sCol = Chr(67 + lCOAj + lCOHi + lMNAj + lMNHi + lMES + lMED)
   MousePointer = 11
   lsArchivo = App.path & "\Spooler\REP" & "_" & Format(lnMes, "00") & lnAnio & ".xls"
   OleExcel.Class = "ExcelWorkSheet"
   lbLibroOpen = ExcelBegin(lsArchivo, xlAplicacion, xlLibro)
   If lbLibroOpen Then
      Set xlHoja1 = xlLibro.Worksheets(1)
      ExcelAddHoja "R" & Right(gsOpeCod, 1), xlLibro, xlHoja1
      CabeceraExcel sCol
      For N = 1 To lvRep.ListItems.Count
      nCol = 3
         With lvRep.ListItems(N)
         xlHoja1.Cells(N + 7, 2) = .Text
         xlHoja1.Cells(N + 7, 3) = .SubItems(1)
         
         If lCOHi = 1 Then
            nCol = nCol + 1
            xlHoja1.Cells(N + 7, nCol) = .SubItems(2)
         End If
         If lCOAj = 1 Then
            nCol = nCol + 1
            xlHoja1.Cells(N + 7, nCol) = .SubItems(3)
         End If
         If lMNHi = 1 Then
            nCol = nCol + 1
            xlHoja1.Cells(N + 7, nCol) = .SubItems(4)
         End If
         If lMNAj = 1 Then
            nCol = nCol + 1
            xlHoja1.Cells(N + 7, nCol) = .SubItems(5)
         End If
         If lMES = 1 Then
            nCol = nCol + 1
            xlHoja1.Cells(N + 7, nCol) = .SubItems(6)
         End If
         If lMED = 1 Then
            nCol = nCol + 1
            xlHoja1.Cells(N + 7, nCol) = .SubItems(7)
         End If
         End With
      Next
      N = N + 7
      
      xlHoja1.Range("B" & 8 & ":" & sCol & N).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic
      xlHoja1.Range("B" & 8 & ":" & sCol & N).Borders(xlInsideVertical).LineStyle = xlContinuous
     
      OleExcel.Class = "ExcelWorkSheet"
      ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1
      OleExcel.SourceDoc = lsArchivo
      OleExcel.Verb = 1
      OleExcel.Action = 1
      OleExcel.DoVerb -1
   End If
   MousePointer = 0
Exit Sub
ErrImprime:
   MsgBox TextErr(Err.Description), vbInformation, "!Aviso!"
   MousePointer = 0
End Sub

Private Sub CabeceraExcel(sCol As String)
Dim nCol As Integer
xlHoja1.PageSetup.Orientation = xlLandscape

xlHoja1.Cells(1, 2) = gsNomCmac
xlHoja1.Cells(3, 2) = lsRepDesc
xlHoja1.Cells(4, 2) = "MES : " & txtMes & " " & txtAnio
xlHoja1.Range("B3:" & sCol & "3").Merge
xlHoja1.Range("B4:" & sCol & "4").Merge

xlHoja1.Cells(6, 2) = "CODIGO"
xlHoja1.Cells(6, 3) = "DESCRIPCION"

nCol = 3

If lCOHi Then
   nCol = nCol + 1
   xlHoja1.Cells(6, nCol) = "CONSOLIDADO"
   xlHoja1.Cells(7, nCol) = "HISTORICO"
End If
If lCOAj Then
   nCol = nCol + 1
   xlHoja1.Cells(6, nCol) = "CONSOLIDADO"
   xlHoja1.Cells(7, nCol) = "AJUSTADO"
End If
If lMNHi Then
   nCol = nCol + 1
   xlHoja1.Cells(6, nCol) = "M.NACIONAL"
   xlHoja1.Cells(7, nCol) = "HISTORICO"
End If
If lMNAj Then
   nCol = nCol + 1
   xlHoja1.Cells(6, nCol) = "M.NACIONAL"
   xlHoja1.Cells(7, nCol) = "AJUSTADO"
End If
If lMES Then
   nCol = nCol + 1
   xlHoja1.Cells(6, nCol) = "M.EXTRANJERA"
   xlHoja1.Cells(7, nCol) = "en MN"
End If
If lMED Then
   nCol = nCol + 1
   xlHoja1.Cells(6, nCol) = "M.EXTRANJERA"
   xlHoja1.Cells(7, nCol) = "en ME"
End If

xlHoja1.Range("B3:" & sCol & "3").Font.Size = 12
xlHoja1.Range("A1:" & sCol & "7").Font.Bold = True
xlHoja1.Range("A3:" & sCol & "7").HorizontalAlignment = xlHAlignCenter
xlHoja1.Range(xlHoja1.Cells(6, 2), xlHoja1.Cells(7, nCol)).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
xlHoja1.Range(xlHoja1.Cells(6, 2), xlHoja1.Cells(7, nCol)).Borders(xlInsideVertical).LineStyle = xlContinuous

xlHoja1.Range("A1:A1").ColumnWidth = 1
xlHoja1.Range("B1:B1").ColumnWidth = 8
xlHoja1.Range("C1:C1").ColumnWidth = 25
xlHoja1.Range("D1:D1").ColumnWidth = 14
xlHoja1.Range("E1:E1").ColumnWidth = 14
xlHoja1.Range("F1:F1").ColumnWidth = 14
xlHoja1.Range("G1:G1").ColumnWidth = 14
xlHoja1.Range("H1:H1").ColumnWidth = 14
xlHoja1.Range("I1:I1").ColumnWidth = 14
xlHoja1.Range("C1:" & sCol & "1").EntireColumn.NumberFormat = "#,##0.00;-#,##0.00"
xlHoja1.Range(xlHoja1.Cells(8, 2), xlHoja1.Cells(lvRep.ListItems.Count + 8, 2)).HorizontalAlignment = xlHAlignCenter

End Sub


Private Sub cmdImprimir_Click()
Select Case fsCodReport
    Case gRiesgoSBSA02A
        Call GenRep2AContingentesPondxRiesgo
    Case gRiesgoSBSA02B
        Call GenRep2BResumenContingentesxRiesgo
    Case gContAnx09
        Call GenAnex9RequerPatrimonialAfectasRiesgoCambiario
   
    Case Else
        Call GenRepResumenContCartera
End Select
End Sub

Private Sub cmdOpciones_Click()
'***Modificado por ELRO el 20120215, según Acta N° 352-2011/TI-D
'frmRepBaseFormulaOpc.Show 1, Me
frmRepBaseFormulaOpc.Inicio (fsCodReport)
End Sub

'***Agregado por ELRO el 20111222, según Acta N° 352-2011/TI-D
Private Sub cmdPatrimonio_Click()
Dim oForm As New frmRegPatrimonioEfectivo
oForm.Show 1, Me
lblPatrimonio = Format(CStr(oForm.PnPatrimonioEfectivo), gsFormatoNumeroDato)
lblMes = oForm.PsMes
lblAnio = oForm.PsAnio
End Sub
'***Fin Agregado por ELRO*************************************

Private Sub CmdProceso_Click()
Dim R As Double
Dim Cad As String
MousePointer = 11
gsOpeCod = lsRepCod '780021
'MIOL 20120705, SEGUN RQ12122 *****************************
If gsOpeCod = "763411" Then
    sTipoRepoFormula = "0"
Else
    sTipoRepoFormula = Right(gsOpeCod, 1) '780021
End If
'END MIOL *************************************************
sTituRepoFormula = lsRepDesc
If Not ValidaAnio(lnAnio) Then
   MousePointer = 0
   Exit Sub
End If
   Call CargaDatos
   Call GeneraReporte
   MousePointer = 0
End Sub

Private Sub cmdReporte_Click()
    Call CargaDatosReporte
    Call GeneraReporte770091
    Call GenRepResumenContCartera770091
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdHoja_Click()
    Call GenRepResumenContCartera
End Sub

Private Sub Form_Load()
   CentraForm Me
   Set oNBal = New NBalanceCont
   Set oDBal = New DbalanceCont
   Me.Caption = gsOpeDesc
   Dim oOpe As New DOperacion
   Set R = oOpe.CargaOpeCta(Mid(gContRepBaseFormula, 1, 4)) '**/** 760100
   Do While Not R.EOF
'     If fsCodReport = "760104" And R!cOpeCod = "760104" Then
'        cboTipo.AddItem R!cOpeDesc & Space(50) & Trim(R!cOpeCod)
'     ElseIf fsCodReport = "760110" And R!cOpeCod = "760110" Then
'        cboTipo.AddItem R!cOpeDesc & Space(50) & Trim(R!cOpeCod)
'     ElseIf fsCodReport = "760111" And R!cOpeCod = "760111" Then
'        cboTipo.AddItem R!cOpeDesc & Space(50) & Trim(R!cOpeCod)
'     ElseIf fsCodReport = "760112" And R!cOpeCod = "760112" Then
'        cboTipo.AddItem R!cOpeDesc & Space(50) & Trim(R!cOpeCod)
'     End If
     R.MoveNext
   Loop
   RSClose R
   frmRepBaseFormulaOpc.plCOAj = 1
   frmRepBaseFormulaOpc.plCOHist = 1
   frmRepBaseFormulaOpc.plMNAj = 1
   frmRepBaseFormulaOpc.plMNHist = 1
   frmRepBaseFormulaOpc.plMES = 1
   frmRepBaseFormulaOpc.plMED = 1
   '***Agregado por ELRO el 20111223, según Acta 352-2011/TI-D
   'MIOL 20120705, SEGUN RQ12122 *****************************
    If fsCodReport = "770090" Or fsCodReport = "763411" Then
'    If fsCodReport = "770090" Then
    'END MIOL ************************************************
        frmRepBaseFormulaOpc.pfnMEDAnterior = 1
        frmRepBaseFormulaOpc.pfnDiferencia = 1
    End If
   '***Fin Agregado por ELRO**********************************
   txtMes = frmReportes.cboMes
   'txtAnio = lnAnio
   'MIOL 20120705, SEGUN RQ12122 *****************************
   If fsCodReport = "770090" Then
        sCodOpeReporte = "770091"
   ElseIf fsCodReport = "763411" Then
        sCodOpeReporte = "763412"
   End If
   'END MIOL ************************************************
   
   
   txtFechaCorte.Text = gdFecSis
   txtAnio = Year(txtFechaCorte.Text)
   txtMes = UCase(IIf(Month(txtFechaCorte.Text) = "1", "Enero", IIf(Month(txtFechaCorte.Text) = "2", "Febrero", IIf(Month(txtFechaCorte.Text) = "3", "Marzo", IIf(Month(txtFechaCorte.Text) = "4", "Abril", IIf(Month(txtFechaCorte.Text) = "5", "Mayo", ""))))))
   If txtMes = "" Then
        txtMes = UCase(IIf(Month(txtFechaCorte.Text) = "6", "Junio", IIf(Month(txtFechaCorte.Text) = "7", "Julio", IIf(Month(txtFechaCorte.Text) = "8", "Agosto", IIf(Month(txtFechaCorte.Text) = "9", "Setiembre", IIf(Month(txtFechaCorte.Text) = "10", "Octubre", IIf(Month(txtFechaCorte.Text) = "11", "Noviembre", IIf(Month(txtFechaCorte.Text) = "12", "Diciembre", ""))))))))
   End If
   
   '***Agregado por ELRO el 20111226, según Acta N° 352-2011/TI-D
   'MIOL 20120705, SEGUN RQ12122 *****************************
    If fsCodReport = "770090" Or fsCodReport = "763411" Then
'   If fsCodReport = "770090" Then
   'END MIOL ************************************************
            
        Dim oDbalanceCont As DbalanceCont
        Set oDbalanceCont = New DbalanceCont
        Dim rsPatrimonio As ADODB.Recordset
        Set rsPatrimonio = New ADODB.Recordset
        
        lblTituloPatrimonio.Visible = True
        lblPatrimonio.Visible = True
        lblMes.Visible = True
        lblAnio.Visible = True
        cmdPatrimonio.Visible = True
        
        Set rsPatrimonio = oDbalanceCont.recuperarPatrimonioEfectivo()
      
        If Not rsPatrimonio.BOF And Not rsPatrimonio.EOF Then
            lblPatrimonio = rsPatrimonio.Fields(1)
            lblMes = UCase(rsPatrimonio.Fields(2))
            lblAnio = rsPatrimonio.Fields(3)
        Else
            lblPatrimonio = ""
            lblMes = ""
            lblAnio = ""
        End If
      
         Set rsPatrimonio = Nothing
         Set oDbalanceCont = Nothing
   End If
   

   '***Fin Agregado por ELRO*************************************
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set oNBal = Nothing
Set oDBal = Nothing
End Sub

' ** Reporte 2A - Activos y Creditos Contingentes Ponderados por Riesgo
Private Sub GenRep2AContingentesPondxRiesgo()
Dim lsSql As String
Dim lrReg As New ADODB.Recordset
Dim i As Integer, lnIIni As Integer
Dim lnRow As Integer
Dim lsArchivo   As String
Dim lbLibroOpen As Boolean
Dim lnFactor As Double

Dim lnValorAcumulado As Currency
Dim lnValorAcumuladoPond As Currency

lnValorAcumulado = 0
lnValorAcumuladoPond = 0

Dim lsCadenaTotal As String
Dim lsCadenaTotaPond As String


'On Error GoTo ErrImprime
   MousePointer = 11
   lsArchivo = App.path & "\Spooler\REP02" & "_" & txtMes.Caption & txtAnio & ".xls"
   lbLibroOpen = ExcelBegin(lsArchivo, xlAplicacion, xlLibro)
   If lbLibroOpen Then
      Set xlHoja1 = xlLibro.Worksheets(1)
      ExcelAddHoja "Rep2A_ContingxRiesgo", xlLibro, xlHoja1
      ' **** Cabecera **********
      xlHoja1.PageSetup.Orientation = xlPortrait
      xlHoja1.PageSetup.CenterHorizontally = True
      xlHoja1.PageSetup.Zoom = 55
      xlHoja1.Cells(1, 1) = gsNomCmac
      xlHoja1.Cells(2, 1) = "ACTIVOS Y CREDITOS CONTINGENTES PONDERADOS POR RIESGO"
      xlHoja1.Cells(3, 1) = "(Articulos 189 al 196 de la Ley Nro 26702) "
      xlHoja1.Cells(4, 1) = "AL " & Format(dFecha, "dd/mm/yyyy")
   
      xlHoja1.Range(xlHoja1.Cells(2, 1), xlHoja1.Cells(4, 4)).Font.Bold = True
      xlHoja1.Range(xlHoja1.Cells(2, 1), xlHoja1.Cells(4, 4)).Merge True
      xlHoja1.Range(xlHoja1.Cells(2, 1), xlHoja1.Cells(4, 4)).HorizontalAlignment = xlCenter

      xlHoja1.Range("A1:A150").ColumnWidth = 70
      xlHoja1.Range("B1:B150").ColumnWidth = 20
      xlHoja1.Range("C1:C150").ColumnWidth = 20
   
      xlHoja1.Range(xlHoja1.Cells(6, 3), xlHoja1.Cells(150, 3)).NumberFormat = "#,##0.00;#,##0.00"
    
      i = 5
      xlHoja1.Cells(i, 1) = "Ley General   -   Denominacion"
      xlHoja1.Cells(i, 2) = "Saldo "
      xlHoja1.Cells(i + 1, 2) = "Contable"
      xlHoja1.Cells(i, 3) = "Saldo"
      xlHoja1.Cells(i + 1, 3) = "Ponderado"
      xlHoja1.Range(xlHoja1.Cells(i, 1), xlHoja1.Cells(i + 1, 4)).HorizontalAlignment = xlCenter
      xlHoja1.Range(xlHoja1.Cells(i, 1), xlHoja1.Cells(i + 1, 4)).Cells.Borders.LineStyle = xlOutside
      xlHoja1.Range(xlHoja1.Cells(i, 1), xlHoja1.Cells(i + 1, 4)).Cells.Borders.LineStyle = xlInside
      i = i + 2
      lnIIni = i
      For lnRow = 1 To lvRep.ListItems.Count
         With lvRep.ListItems(lnRow)
            
            xlHoja1.Cells(i, 1) = .SubItems(1)
            xlHoja1.Cells(i, 2) = .SubItems(2) / 1000
            
            If lvRep.ListItems(lnRow).SubItems(8) = 1 Then
                xlHoja1.Range("B" & Trim(Str(i)) & ":B" & Trim(Str(i))).Interior.ColorIndex = 36
                xlHoja1.Range("B" & Trim(Str(i)) & ":B" & Trim(Str(i))).Interior.Pattern = xlSolid
            End If
            
            Select Case Mid(lvRep.ListItems(lnRow), 1, 1)
                Case "1":   lnFactor = 0
                Case "2":   lnFactor = 0.1
                Case "3":   lnFactor = 0.2
                Case "4":   lnFactor = 0.5
                Case "5":   lnFactor = 1
            End Select
            
            If .SubItems(2) <> "" Then
                xlHoja1.Cells(i, 3) = .SubItems(2) * lnFactor
                xlHoja1.Range(xlHoja1.Cells(i, 3), xlHoja1.Cells(i, 3)).Formula = "=" & ExcelColumnaString(2) & Trim(Str(i)) & "*" & Str(lnFactor)
            End If
         
            If Len(Trim(lvRep.ListItems(lnRow))) = 1 Then
                lnValorAcumulado = lnValorAcumulado + .SubItems(2)
                lnValorAcumuladoPond = lnValorAcumuladoPond + .SubItems(2) * lnFactor
            End If
            
            If .Text = "1" Or .Text = "2" Or .Text = "3" Or .Text = "4" Or .Text = "5" Then
                lsCadenaTotal = lsCadenaTotal & "+" & ExcelColumnaString(2) & Trim(Str(i))
                lsCadenaTotaPond = lsCadenaTotaPond & "+" & ExcelColumnaString(3) & Trim(Str(i))
            End If
            
            If lvRep.ListItems(lnRow) = "999999" Then
                xlHoja1.Range(xlHoja1.Cells(i, 2), xlHoja1.Cells(i, 2)).Formula = "=" & lsCadenaTotal
                xlHoja1.Range(xlHoja1.Cells(i, 3), xlHoja1.Cells(i, 3)).Formula = "=" & lsCadenaTotaPond
            End If
            
         End With
         i = i + 1
      Next
    
      xlHoja1.Range(xlHoja1.Cells(lnIIni, 1), xlHoja1.Cells(i - 1, 4)).Borders(xlInsideVertical).LineStyle = xlContinuous
      xlHoja1.Range(xlHoja1.Cells(lnIIni, 1), xlHoja1.Cells(i - 1, 4)).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic
   
     
      ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1
      CargaArchivo lsArchivo, App.path & "\Spooler\"
   End If
   MousePointer = 0
Exit Sub
ErrImprime:
   MsgBox TextErr(Err.Description), vbInformation, "!Aviso!"
   MousePointer = 0
End Sub


' ** Reporte 2B - Resumen Activos y Creditos Contingentes Ponderados por Riesgo
Private Sub GenRep2BResumenContingentesxRiesgo()
Dim lsSql As String
Dim lrReg As New ADODB.Recordset
Dim i As Integer, lnIIni As Integer
Dim lnRow As Integer
Dim lsArchivo   As String
Dim lbLibroOpen As Boolean
Dim lnFactor As Double
Dim lsCeldaI As String, lsCeldaII As String, lsCeldaIII As String, lsCeldaIV As String
Dim sRuta As String

Dim psArchivoA_Leer As String
Dim xlApp As Excel.Application
Dim xlLibroApp As Excel.Workbook
Dim xlHojaApp As Excel.Worksheet
Dim bExiste As Boolean
Dim bEncontrado As Boolean
Dim fs As New Scripting.FileSystemObject
Dim lnIApp As Integer

Dim lnDatos(5, 3) As Currency

Me.cmdAbrir.DialogTitle = "Seleccione la Ruta "
cmdAbrir.ShowOpen
sRuta = cmdAbrir.FileName

If sRuta = "" Then
    MsgBox "No se encontro Archivo de Reporte 2A.", vbInformation, "Aviso"
    Exit Sub
End If

psArchivoA_Leer = sRuta
bExiste = fs.FileExists(psArchivoA_Leer)

If bExiste = False Then
    MsgBox "Ud debe generar previamente el reporte " & Chr(13) & psArchivoA_Leer, vbExclamation, "Aviso!!!"
    Exit Sub
End If



Set xlApp = New Excel.Application
Set xlLibroApp = xlApp.Workbooks.Open(psArchivoA_Leer)
'''''''''''''''''''''''''''''
bEncontrado = False
For Each xlHojaApp In xlLibroApp.Worksheets
    If UCase(xlHojaApp.Name) = UCase("Rep2A_ContingxRiesgo") Then
        bEncontrado = True
        xlHojaApp.Activate
        Exit For
    End If
Next

If bEncontrado = False Then
    ExcelEnd psArchivoA_Leer, xlApp, xlLibroApp, xlHojaApp, True
    MsgBox "No existen datos con la fecha especificada", vbExclamation, "Aviso!!!"
    Exit Sub
End If

For lnIApp = 7 To 1000
    If UCase(xlHojaApp.Cells(lnIApp, 1)) = UCase("I ACTIVOS CON RIESGO 0% - Factor 0.0") Then
        lnDatos(1, 0) = 1
        lnDatos(1, 1) = xlHojaApp.Cells(lnIApp, 2)
        lnDatos(1, 2) = xlHojaApp.Cells(lnIApp, 3)
    End If
    If UCase(xlHojaApp.Cells(lnIApp, 1)) = UCase("II. ACTIVOS CREDITICIOS CON REISGO 10% - FACTOR 0,1") Then
        lnDatos(2, 0) = 1
        lnDatos(2, 1) = xlHojaApp.Cells(lnIApp, 2)
        lnDatos(2, 2) = xlHojaApp.Cells(lnIApp, 3)
    End If
    If UCase(xlHojaApp.Cells(lnIApp, 1)) = UCase("III.   ACTIVOS CREDITICIOS CON RIESGO 20% - FACTOR 0.2") Then
        lnDatos(3, 0) = 1
        lnDatos(3, 1) = xlHojaApp.Cells(lnIApp, 2)
        lnDatos(3, 2) = xlHojaApp.Cells(lnIApp, 3)
    End If
    If UCase(xlHojaApp.Cells(lnIApp, 1)) = UCase("IV. ACTIVOS CREDITICIOS CON RIESGO 50% - FACTOR 0.5") Then
        lnDatos(4, 0) = 1
        lnDatos(4, 1) = xlHojaApp.Cells(lnIApp, 2)
        lnDatos(4, 2) = xlHojaApp.Cells(lnIApp, 3)
    End If
    If UCase(xlHojaApp.Cells(lnIApp, 1)) = UCase("V. ACTIVOS CON RiESGO CIEN (100%)") Then
        lnDatos(5, 0) = 1
        lnDatos(5, 1) = xlHojaApp.Cells(lnIApp, 2)
        lnDatos(5, 2) = xlHojaApp.Cells(lnIApp, 3)
    End If
Next lnIApp

ExcelEnd psArchivoA_Leer, xlApp, xlLibroApp, xlHojaApp, False

'On Error GoTo ErrImprime
   MousePointer = 11
    lsArchivo = App.path & "\Spooler\REP" & "_" & Format(lnMes, "00") & lnAnio & ".xls"
   lbLibroOpen = ExcelBegin(lsArchivo, xlAplicacion, xlLibro)
   If lbLibroOpen Then
      Set xlHoja1 = xlLibro.Worksheets(1)
      ExcelAddHoja "Rep2B_ResumenxRiesgo", xlLibro, xlHoja1
      ' **** Cabecera **********
      xlHoja1.PageSetup.Orientation = xlPortrait
      xlHoja1.PageSetup.CenterHorizontally = True
      xlHoja1.PageSetup.Zoom = 70
      xlHoja1.Cells(1, 1) = gsNomCmac
      xlHoja1.Cells(2, 1) = "Reporte 2-B"
      xlHoja1.Cells(3, 1) = "RESUMEN DE ACTIVOS Y CREDITOS CONTINGENTES PONDERADOS POR RIESGO"
      xlHoja1.Cells(4, 1) = "CREDITICIO Y REQUERIEMIENTOS PATRIMONIALES POR RIESGO DE MERCADO"
      xlHoja1.Cells(5, 1) = "(Articulos 189, 196, 199 de la Ley Nro 26702, Modificada por la Ley Nro 27008 Y la Ley 27102) "
      xlHoja1.Cells(6, 1) = "AL " & Format(dFecha, "dd/mm/yyyy")
   
      xlHoja1.Range(xlHoja1.Cells(2, 1), xlHoja1.Cells(6, 4)).Font.Bold = True
      xlHoja1.Range(xlHoja1.Cells(2, 1), xlHoja1.Cells(6, 4)).Merge True
      xlHoja1.Range(xlHoja1.Cells(2, 1), xlHoja1.Cells(6, 4)).HorizontalAlignment = xlCenter

      xlHoja1.Range("A1:A150").ColumnWidth = 20
      xlHoja1.Range("B1:B150").ColumnWidth = 70
      xlHoja1.Range("C1:C150").ColumnWidth = 20
      xlHoja1.Range("D1:D150").ColumnWidth = 20
      
      xlHoja1.Range(xlHoja1.Cells(6, 3), xlHoja1.Cells(150, 4)).NumberFormat = "#,##0"
    
      i = 8
      xlHoja1.Cells(i, 1) = "A) Activos y Creditos contingentes Ponderados por Riesgos Crediticio: "
      i = i + 1
      xlHoja1.Cells(i, 2) = "Categorias"
      xlHoja1.Cells(i, 3) = "Saldos Contables"
      xlHoja1.Cells(i, 4) = "Saldos Ponderados"
      
      xlHoja1.Range(xlHoja1.Cells(i, 2), xlHoja1.Cells(i, 4)).HorizontalAlignment = xlCenter
      xlHoja1.Range(xlHoja1.Cells(i, 2), xlHoja1.Cells(i, 4)).Cells.Borders.LineStyle = xlOutside
      xlHoja1.Range(xlHoja1.Cells(i, 2), xlHoja1.Cells(i, 4)).Cells.Borders.LineStyle = xlInside
      
      i = i + 1
      lnIIni = i
      lnIApp = 1
      For lnRow = 1 To lvRep.ListItems.Count
         With lvRep.ListItems(lnRow)
            If Mid(lvRep.ListItems(lnRow), 1, 4) = "1010" Then
                xlHoja1.Cells(i, 2) = .SubItems(1)
                'xlHoja1.Cells(I, 3) = .SubItems(2) / 1000
                xlHoja1.Cells(i, 3) = lnDatos(lnIApp, 1)
                Select Case Mid(lvRep.ListItems(lnRow), 5, 1)
                    Case "1":   lnFactor = 0
                    Case "2":   lnFactor = 0.1
                    Case "3":   lnFactor = 0.2
                    Case "4":   lnFactor = 0.5
                    Case "5":   lnFactor = 1
                End Select
                'xlHoja1.Cells(I, 4) = .SubItems(2) * lnFactor / 1000
                xlHoja1.Cells(i, 4) = lnDatos(lnIApp, 2)
                lnIApp = lnIApp + 1
                i = i + 1
            End If
         End With
      Next
      
      xlHoja1.Cells(i, 3) = "TOTAL (I)"
      xlHoja1.Range("D" & i, "D" & i).Formula = "=SUM(D" & lnIIni & ":D" & i - 1 & ")"
      lsCeldaI = "D" & i
    
      xlHoja1.Range(xlHoja1.Cells(lnIIni, 2), xlHoja1.Cells(i - 1, 4)).Borders(xlInsideVertical).LineStyle = xlContinuous
      xlHoja1.Range(xlHoja1.Cells(lnIIni, 2), xlHoja1.Cells(i - 1, 4)).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic
   
      i = i + 3
      xlHoja1.Cells(i, 1) = "B) Requerimiento Patrimonial Minimo por riesgo de Mercado "
      i = i + 1
      xlHoja1.Cells(i, 2) = "Aplicable a :"
      xlHoja1.Cells(i, 3) = ""
      xlHoja1.Cells(i, 4) = "Total (II) "
      xlHoja1.Range(xlHoja1.Cells(i, 2), xlHoja1.Cells(i + 3, 3)).Merge True
      i = i + 1
      lnIIni = i
      For lnRow = 1 To lvRep.ListItems.Count
         With lvRep.ListItems(lnRow)
            If Mid(lvRep.ListItems(lnRow), 1, 4) = "2010" Then
                xlHoja1.Cells(i, 2) = .SubItems(1)
                xlHoja1.Cells(i, 4) = .SubItems(6) / 1000
                i = i + 1
            End If
         End With
      Next
      lsCeldaII = "D" & lnIIni
      
      xlHoja1.Range(xlHoja1.Cells(lnIIni, 2), xlHoja1.Cells(i - 1, 4)).Borders(xlInsideVertical).LineStyle = xlContinuous
      xlHoja1.Range(xlHoja1.Cells(lnIIni, 2), xlHoja1.Cells(i - 1, 4)).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic
     
      i = i + 2
      xlHoja1.Cells(i, 1) = "C) Patrimonio Efectivo Total"
      xlHoja1.Cells(i, 4) = "( III )"
      lsCeldaIII = "C" & i
      For lnRow = 1 To lvRep.ListItems.Count
         With lvRep.ListItems(lnRow)
            If Trim(lvRep.ListItems(lnRow)) = "3" Then
                xlHoja1.Cells(i, 3) = .SubItems(2) / 1000
                i = i + 1
            End If
         End With
      Next
      xlHoja1.Range(xlHoja1.Cells(i - 1, 3), xlHoja1.Cells(i - 1, 3)).BorderAround xlContinuous, xlThick, xlColorIndexAutomatic
      i = i + 1
      xlHoja1.Cells(i, 1) = " - Patrimonio asignado a Riesgo crediticio"
      xlHoja1.Cells(i, 4) = "( IV )"
      lsCeldaIV = "C" & i
      For lnRow = 1 To lvRep.ListItems.Count
         With lvRep.ListItems(lnRow)
            If Trim(lvRep.ListItems(lnRow)) = "301" Then
                'xlHoja1.Cells(I, 3) = .SubItems(2) / 1000
                xlHoja1.Range(xlHoja1.Cells(i, 3), xlHoja1.Cells(i, 3)).Formula = "=" & lsCeldaIII & "-" & lsCeldaII
                i = i + 1
            End If
         End With
      Next
      xlHoja1.Range(xlHoja1.Cells(i - 1, 3), xlHoja1.Cells(i - 1, 3)).BorderAround xlContinuous, xlThick, xlColorIndexAutomatic
      i = i + 1
      xlHoja1.Cells(i, 1) = " - Patrimonio asignado a Riesgo de mercado"
      xlHoja1.Cells(i, 4) = "( V )"
      For lnRow = 1 To lvRep.ListItems.Count
         With lvRep.ListItems(lnRow)
            If Trim(lvRep.ListItems(lnRow)) = "301" Then
                'xlHoja1.Cells(I, 3) = .SubItems(2) / 1000
                xlHoja1.Range(xlHoja1.Cells(i, 3), xlHoja1.Cells(i, 3)).Formula = "=" & lsCeldaII
                i = i + 1
            End If
         End With
      Next
      xlHoja1.Range(xlHoja1.Cells(i - 1, 3), xlHoja1.Cells(i - 1, 3)).BorderAround xlContinuous, xlThick, xlColorIndexAutomatic
     
      xlHoja1.Range(xlHoja1.Cells(i, 3), xlHoja1.Cells(150, 3)).NumberFormat = "#,##0.00;#,##0.00"
      i = i + 2
      xlHoja1.Cells(i, 1) = "D) Ratio de apalancamiento"
      i = i + 1
      xlHoja1.Cells(i, 1) = "    Relacion : (I) / (IV) "
      xlHoja1.Range("C" & i, "C" & i).Formula = "=" & lsCeldaI & "/" & lsCeldaIV & " "
      xlHoja1.Range(xlHoja1.Cells(i, 3), xlHoja1.Cells(i, 3)).BorderAround xlContinuous, xlThick, xlColorIndexAutomatic
      i = i + 2
      
      xlHoja1.Cells(i, 1) = "E) Ratio de apalancamiento global por riesgo crediticio"
      i = i + 1
      xlHoja1.Cells(i, 1) = "y de mercado Relacion [I + (11*II)] / (III)"
      xlHoja1.Range("C" & i, "C" & i).Formula = "= (" & lsCeldaI & " + (11 * " & lsCeldaII & ") ) / (" & lsCeldaIII & ") "
      xlHoja1.Range(xlHoja1.Cells(i, 3), xlHoja1.Cells(i, 3)).BorderAround xlContinuous, xlThick, xlColorIndexAutomatic
      
      i = i + 5
     xlHoja1.Cells(i, 1) = "1/ Segun lo establecido en el Anexo N 9 del Manual de Contabilidad, al que hace alusion en la Circular Nro B-2030-99,"
     xlHoja1.Cells(i + 1, 1) = "F-0371-99, CM-0220-99, CR-0090-99, EDPYME-0038-99 del 28.02.99"
     xlHoja1.Cells(i + 2, 1) = "2/ Segun lo establecido en el Anexo N 23 del Manual de contabilidad, al que hace alusion en la Circular Nro B-2063-99"
     xlHoja1.Cells(i + 3, 1) = "del 15.11.99"
    
    xlHoja1.Cells.Select
    xlHoja1.Cells.Font.Name = "Arial"
    xlHoja1.Cells.Font.Size = 8
    xlHoja1.Cells.EntireColumn.AutoFit
      
      
      ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1
      CargaArchivo lsArchivo, App.path & "\Spooler\"
   End If
   MousePointer = 0
Exit Sub
ErrImprime:
   MsgBox TextErr(Err.Description), vbInformation, "!Aviso!"
   MousePointer = 0
End Sub

' ** Anexo 9 - Requereimiento patrimonial por posiciones afectadas a riesgo
Private Sub GenAnex9RequerPatrimonialAfectasRiesgoCambiario()
Dim lsSql As String
Dim lrReg As New ADODB.Recordset
Dim i As Integer, lnIIni As Integer
Dim lnRow As Integer
Dim lsArchivo   As String
Dim lbLibroOpen As Boolean
Dim lnFactor As Double

Dim lnMontoPatrimonio As Currency

lnMontoPatrimonio = frmImput.Ini("Patrimonio Efectivo", "Ingrese el Patrimonio")
Set frmImput = Nothing

'On Error GoTo ErrImprime
   MousePointer = 11
   lsArchivo = App.path & "\Spooler\Anex9ReqPatrimxRiesgoCamb" & "_" & txtMes & txtAnio & ".xls"
   lbLibroOpen = ExcelBegin(lsArchivo, xlAplicacion, xlLibro)
   If lbLibroOpen Then
      
      Set xlHoja1 = xlLibro.Worksheets(1)
      ExcelAddHoja "HojaTrabajo", xlLibro, xlHoja1
      ' **** Cabecera **********
      xlHoja1.PageSetup.Orientation = xlPortrait
      xlHoja1.PageSetup.CenterHorizontally = True
      xlHoja1.PageSetup.Zoom = 60
      
      xlHoja1.Cells(1, 1) = "SUPERINTENDENCIA DE BANCA Y SEGUROS"
      xlHoja1.Cells(2, 1) = "EMPRESA: " & gsNomCmac
      xlHoja1.Cells(3, 1) = "CODIGO : " & gsCodCMAC
      
      xlHoja1.Cells(4, 1) = "POSICIONES AFECTAS A RIESGO CAMBIARIO"
      xlHoja1.Cells(5, 1) = "(En Nuevos Soles)"
      xlHoja1.Cells(6, 1) = "AL " & Format(dFecha, "dd/mm/yyyy")
   
      xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(6, 2)).Font.Bold = True
      xlHoja1.Range(xlHoja1.Cells(4, 3), xlHoja1.Cells(4, 10)).MergeCells = True
      xlHoja1.Range(xlHoja1.Cells(5, 3), xlHoja1.Cells(5, 10)).MergeCells = True
      xlHoja1.Range(xlHoja1.Cells(6, 3), xlHoja1.Cells(6, 10)).MergeCells = True
      xlHoja1.Range(xlHoja1.Cells(6, 1), xlHoja1.Cells(6, 2)).MergeCells = True
      xlHoja1.Range(xlHoja1.Cells(4, 3), xlHoja1.Cells(6, 10)).HorizontalAlignment = xlCenter
      xlHoja1.Range("A1:A150").ColumnWidth = 50
      xlHoja1.Range("B1:B150").ColumnWidth = 20

      xlHoja1.Range(xlHoja1.Cells(7, 2), xlHoja1.Cells(150, 2)).NumberFormat = "#,##0.00;#,##0.00"
      i = 8
      lnIIni = i
      For lnRow = 1 To lvRep.ListItems.Count
         With lvRep.ListItems(lnRow)
            If Mid(lvRep.ListItems(lnRow), 1, 1) = "1" Then
                xlHoja1.Cells(i, 1) = .SubItems(1)
                xlHoja1.Cells(i, 2) = .SubItems(6)
                i = i + 1
            End If
         End With
      Next
      xlHoja1.Range(xlHoja1.Cells(lnIIni, 1), xlHoja1.Cells(i - 1, 2)).Cells.Borders.LineStyle = xlOutside
      xlHoja1.Range(xlHoja1.Cells(lnIIni, 1), xlHoja1.Cells(i - 1, 2)).Cells.Borders.LineStyle = xlInside
      
      i = i + 2
      lnIIni = i
      For lnRow = 1 To lvRep.ListItems.Count
         With lvRep.ListItems(lnRow)
            If Mid(lvRep.ListItems(lnRow), 1, 1) = "2" Then
                xlHoja1.Cells(i, 1) = .SubItems(1)
                xlHoja1.Cells(i, 2) = .SubItems(6)
                i = i + 1
            End If
         End With
      Next
      xlHoja1.Range(xlHoja1.Cells(lnIIni, 1), xlHoja1.Cells(i - 1, 2)).Cells.Borders.LineStyle = xlOutside
      xlHoja1.Range(xlHoja1.Cells(lnIIni, 1), xlHoja1.Cells(i - 1, 2)).Cells.Borders.LineStyle = xlInside
      
    
    '*******************************************
    '**** ANEXO 9 ******************************

      Set xlHoja1 = xlLibro.Worksheets(1)
      ExcelAddHoja "Anexo09", xlLibro, xlHoja1
      ' **** Cabecera **********
      xlHoja1.PageSetup.Orientation = xlLandscape
      xlHoja1.PageSetup.CenterHorizontally = True
      xlHoja1.PageSetup.Zoom = 60
      xlHoja1.Range("A1:A150").ColumnWidth = 20
      xlHoja1.Range("B1:S150").ColumnWidth = 12
      
      xlHoja1.Cells(1, 1) = "SUPERINTENDENCIA DE BANCA Y SEGUROS"
      xlHoja1.Cells(2, 1) = "EMPRESA: " & gsNomCmac
      xlHoja1.Cells(3, 1) = "CODIGO : " & gsCodCMAC
      xlHoja1.Cells(1, 13) = "ANEXO Nº 9"
      xlHoja1.Cells(4, 5) = "POSICIONES AFECTAS A RIESGO CAMBIARIO"
      xlHoja1.Cells(5, 5) = "(En Nuevos Soles)"
      xlHoja1.Cells(6, 5) = "AL " & Format(dFecha, "dd/mm/yyyy")
      xlHoja1.Cells(6, 1) = "I.EXPOSICION EN MONEDA EXTRANJERA 1/"
   
      xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(6, 14)).Font.Bold = True
      xlHoja1.Range(xlHoja1.Cells(4, 4), xlHoja1.Cells(4, 14)).MergeCells = True
      xlHoja1.Range(xlHoja1.Cells(5, 4), xlHoja1.Cells(5, 14)).MergeCells = True
      xlHoja1.Range(xlHoja1.Cells(6, 4), xlHoja1.Cells(6, 14)).MergeCells = True
      xlHoja1.Range(xlHoja1.Cells(4, 4), xlHoja1.Cells(6, 14)).HorizontalAlignment = xlCenter

   
      xlHoja1.Range(xlHoja1.Cells(7, 2), xlHoja1.Cells(150, 19)).NumberFormat = "#,##0.00"
    
      i = 7
      xlHoja1.Cells(8, 1) = "DIVISAS"
      xlHoja1.Range(xlHoja1.Cells(8, 1), xlHoja1.Cells(12, 1)).MergeCells = True
      
      xlHoja1.Cells(8, 2) = "ACTIVOS"
      xlHoja1.Cells(9, 2) = "POR"
      xlHoja1.Cells(10, 2) = "DIVISAS"
      'xlHoja1.Range(xlHoja1.Cells(8, 2), xlHoja1.Cells(11, 2)).MergeCells = True
      xlHoja1.Cells(12, 2) = "(A)"
      
      'xlHoja1.Range(xlHoja1.Cells(9, 2), xlHoja1.Cells(9, 5)).MergeCells = True
      'xlHoja1.Cells(11, 2) = "Disponible"
      'xlHoja1.Cells(11, 3) = "Inversiones Financ. Temp."
      'xlHoja1.Range(xlHoja1.Cells(11, 3), xlHoja1.Cells(11, 4)).MergeCells = True
      'xlHoja1.Cells(12, 3) = "Renta Fija"
      'xlHoja1.Cells(12, 4) = "Renta Variab"
      'xlHoja1.Cells(11, 5) = "Otros"
      
      xlHoja1.Cells(8, 3) = "PASIVOS"
      xlHoja1.Cells(9, 3) = "POR"
      xlHoja1.Cells(10, 3) = "DIVISAS"
      'xlHoja1.Range(xlHoja1.Cells(8, 3), xlHoja1.Cells(11, 3)).MergeCells = True
      xlHoja1.Cells(12, 3) = "(B)"
'      xlHoja1.Range(xlHoja1.Cells(9, 6), xlHoja1.Cells(9, 9)).MergeCells = True
'      xlHoja1.Cells(11, 6) = "Depósitos"
'      xlHoja1.Cells(11, 7) = "Adeudados"
'      xlHoja1.Range(xlHoja1.Cells(11, 7), xlHoja1.Cells(11, 8)).MergeCells = True
'      xlHoja1.Cells(12, 7) = "Casa Matriz"
'      xlHoja1.Cells(12, 8) = "Otros"
'      xlHoja1.Cells(11, 9) = "Otros"
      
      xlHoja1.Cells(8, 4) = "POSICION"
      xlHoja1.Cells(9, 4) = "CONTABLE"
      xlHoja1.Cells(10, 4) = "EN M.E."
      xlHoja1.Cells(11, 4) = "(C)"
      xlHoja1.Cells(12, 4) = "(A-B)"
      
      xlHoja1.Cells(8, 5) = "COMPRAS"
      xlHoja1.Cells(9, 5) = "A FUTURO"
      xlHoja1.Cells(10, 5) = "FORWARD"
      xlHoja1.Cells(11, 5) = "DE M.E."
      xlHoja1.Cells(12, 5) = "(D)"
      
      xlHoja1.Cells(8, 6) = "VENTAS A"
      xlHoja1.Cells(9, 6) = "FUTURO"
      xlHoja1.Cells(10, 6) = "FORWARD"
      xlHoja1.Cells(11, 6) = "DE M.E."
      xlHoja1.Cells(12, 6) = "(E)"
      
      xlHoja1.Cells(8, 7) = "COMPRAS"
      xlHoja1.Cells(9, 7) = "EN OTROS "
      xlHoja1.Cells(10, 7) = "DERIVADOS"
      xlHoja1.Cells(11, 7) = "DE M.E."
      xlHoja1.Cells(12, 7) = "(F)"
      
      xlHoja1.Cells(8, 8) = "VENTAS"
      xlHoja1.Cells(9, 8) = "EN OTROS"
      xlHoja1.Cells(10, 8) = "DERIVADOS"
      xlHoja1.Cells(11, 8) = "DE M.E."
      xlHoja1.Cells(12, 8) = "(G)"
      
      xlHoja1.Cells(8, 9) = "POSICION"
      xlHoja1.Cells(9, 9) = "GLOBAL EN ME"
      xlHoja1.Cells(10, 9) = "(H)"
      xlHoja1.Cells(11, 9) = "(C+D+E+F+G)"
      
      xlHoja1.Cells(8, 10) = ""
      xlHoja1.Range(xlHoja1.Cells(8, 10), xlHoja1.Cells(8, 12)).Merge True
      xlHoja1.Cells(9, 10) = "SENSIBILIDAD DE OPCIONES"
      xlHoja1.Range(xlHoja1.Cells(9, 10), xlHoja1.Cells(9, 12)).Merge True
      
      xlHoja1.Cells(11, 10) = "DELTA"
      xlHoja1.Cells(12, 10) = "(J)"
      xlHoja1.Cells(11, 11) = "GAMMA"
      xlHoja1.Cells(12, 11) = "(K)"
      xlHoja1.Cells(11, 12) = "VEGA"
      xlHoja1.Cells(12, 12) = "(L)"
      
      xlHoja1.Cells(8, 13) = "VALOR DE"
      xlHoja1.Cells(9, 13) = "MERCADO DE"
      xlHoja1.Cells(10, 13) = "LAS OPCIONES"
      xlHoja1.Cells(11, 13) = "DE ME"
      xlHoja1.Cells(12, 13) = "(M)"
      
      xlHoja1.Range(xlHoja1.Cells(8, 1), xlHoja1.Cells(12, 14)).Font.Bold = True
      xlHoja1.Range(xlHoja1.Cells(8, 1), xlHoja1.Cells(12, 19)).HorizontalAlignment = xlCenter
      xlHoja1.Range(xlHoja1.Cells(8, 1), xlHoja1.Cells(12, 19)).VerticalAlignment = xlCenter
      xlHoja1.Range(xlHoja1.Cells(8, 1), xlHoja1.Cells(12, 13)).Cells.Borders.LineStyle = xlOutside
      xlHoja1.Range(xlHoja1.Cells(8, 1), xlHoja1.Cells(12, 13)).Cells.Borders.LineStyle = xlInside
      xlHoja1.Range(xlHoja1.Cells(8, 1), xlHoja1.Cells(12, 13)).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic
      
      xlHoja1.Cells(14, 1) = "DOLAR AMERICANO"
      xlHoja1.Cells(16, 1) = "LIBRA ESTERLINA"
      xlHoja1.Cells(18, 1) = "YEN JAPONES"
      xlHoja1.Cells(20, 1) = "DOLAR CANADIENSE"
      xlHoja1.Cells(22, 1) = "EURO"
      xlHoja1.Cells(24, 1) = "OTRAS DIVISAS"
      xlHoja1.Cells(26, 1) = "ORO"
      xlHoja1.Cells(28, 1) = "TOTAL ME."
      xlHoja1.Cells(30, 1) = "TOTAL M.E./P.E."
      
      i = i + 2

      For lnRow = 1 To lvRep.ListItems.Count
         With lvRep.ListItems(lnRow)
            If Mid(lvRep.ListItems(lnRow), 1, 3) = "101" Then
                xlHoja1.Cells(14, 2) = .SubItems(6) / 1
            End If
         End With
         i = i + 1
      Next
    
'      For lnRow = 1 To lvRep.ListItems.Count
'         With lvRep.ListItems(lnRow)
'            If Trim(lvRep.ListItems(lnRow)) = "104" Then
'                xlHoja1.Cells(14, 5) = .SubItems(6) / 1000
'            End If
'         End With
'         i = i + 1
'      Next

      For lnRow = 1 To lvRep.ListItems.Count
         With lvRep.ListItems(lnRow)
            If Trim(lvRep.ListItems(lnRow)) = "201" Then
                xlHoja1.Cells(14, 3) = .SubItems(6) / 1
            End If
         End With
         i = i + 1
      Next
      
'      For lnRow = 1 To lvRep.ListItems.Count
'         With lvRep.ListItems(lnRow)
'            If Trim(lvRep.ListItems(lnRow)) = "203" Then
'                xlHoja1.Cells(14, 8) = .SubItems(6) / 1000
'            End If
'         End With
'         i = i + 1
'      Next
'      For lnRow = 1 To lvRep.ListItems.Count
'         With lvRep.ListItems(lnRow)
'            If Trim(lvRep.ListItems(lnRow)) = "204" Then
'                xlHoja1.Cells(14, 9) = .SubItems(6) / 1000
'            End If
'         End With
'         i = i + 1
'      Next
    xlHoja1.Range("D14:D14").Formula = "=B14-C14"
    xlHoja1.Range("I14:I14").Formula = "=D14"
     
    xlHoja1.Range("B28:B28").Formula = "=SUM(B14:B27)"
    xlHoja1.Range("C28:C28").Formula = "=SUM(C14:C27)"
    xlHoja1.Range("D28:D28").Formula = "=SUM(D14:D27)"
    xlHoja1.Range("E28:E28").Formula = "=SUM(E14:E27)"
    xlHoja1.Range("F28:F28").Formula = "=SUM(F14:F27)"
    xlHoja1.Range("G28:G28").Formula = "=SUM(G14:G27)"
    xlHoja1.Range("H28:H28").Formula = "=SUM(H14:H27)"
    xlHoja1.Range("I28:I28").Formula = "=SUM(I14:I27)"
    xlHoja1.Range("J28:J28").Formula = "=SUM(J14:J27)"
    xlHoja1.Range("K28:K28").Formula = "=SUM(K14:K27)"
    xlHoja1.Range("L28:L28").Formula = "=SUM(L14:L27)"
    xlHoja1.Range("M28:M28").Formula = "=SUM(M14:M27)"
    
    If lnMontoPatrimonio <> 0 Then
        lnMontoPatrimonio = lnMontoPatrimonio / 1
        xlHoja1.Range("B30:B30").Formula = "=B28/" & lnMontoPatrimonio
        xlHoja1.Range("C30:C30").Formula = "=C28/" & lnMontoPatrimonio
        xlHoja1.Range("D30:D30").Formula = "=D28/" & lnMontoPatrimonio
        xlHoja1.Range("E30:E30").Formula = "=E28/" & lnMontoPatrimonio
        xlHoja1.Range("F30:F30").Formula = "=F28/" & lnMontoPatrimonio
        xlHoja1.Range("G30:G30").Formula = "=G28/" & lnMontoPatrimonio
        xlHoja1.Range("H30:H30").Formula = "=H28/" & lnMontoPatrimonio
        xlHoja1.Range("I30:I30").Formula = "=I28/" & lnMontoPatrimonio
        xlHoja1.Range("J30:J30").Formula = "=J28/" & lnMontoPatrimonio
        xlHoja1.Range("K30:K30").Formula = "=K28/" & lnMontoPatrimonio
        xlHoja1.Range("L30:L30").Formula = "=L28/" & lnMontoPatrimonio
        xlHoja1.Range("M30:M30").Formula = "=M28/" & lnMontoPatrimonio
    End If
    
    
    xlHoja1.Range(xlHoja1.Cells(13, 1), xlHoja1.Cells(30, 14)).Borders(xlInsideVertical).LineStyle = xlContinuous
    xlHoja1.Range("A13:M30").BorderAround xlContinuous, xlThin, xlColorIndexAutomatic
    xlHoja1.Range(xlHoja1.Cells(28, 1), xlHoja1.Cells(37, 14)).Font.Bold = True
    
    xlHoja1.Cells(35, 1) = "II.REQUERIMIENTO PARIMONIAL POR POSICIONES AFECTAS A RIESGO CAMBIARIO"
    xlHoja1.Cells(36, 1) = "Metodo"
    xlHoja1.Range(xlHoja1.Cells(36, 1), xlHoja1.Cells(37, 2)).Cells.Merge
    xlHoja1.Range(xlHoja1.Cells(36, 1), xlHoja1.Cells(37, 2)).HorizontalAlignment = xlCenter
    xlHoja1.Cells(36, 3) = "Posicion"
    xlHoja1.Range(xlHoja1.Cells(36, 3), xlHoja1.Cells(37, 5)).Cells.Merge
    xlHoja1.Range(xlHoja1.Cells(36, 3), xlHoja1.Cells(37, 5)).HorizontalAlignment = xlCenter
    xlHoja1.Cells(36, 6) = "Monto"
    xlHoja1.Range(xlHoja1.Cells(36, 6), xlHoja1.Cells(37, 6)).Cells.Merge
    xlHoja1.Range(xlHoja1.Cells(36, 6), xlHoja1.Cells(37, 6)).HorizontalAlignment = xlCenter
    xlHoja1.Cells(36, 7) = "Cargo"
    xlHoja1.Cells(37, 7) = "Capital"
    xlHoja1.Cells(36, 8) = "Requerimiento"
    xlHoja1.Cells(37, 8) = "Patrimonial"
    xlHoja1.Cells(36, 9) = "Total"
    xlHoja1.Range(xlHoja1.Cells(36, 9), xlHoja1.Cells(37, 9)).Cells.Merge
    
    xlHoja1.Range("A36:B37").BorderAround xlContinuous, xlThin, xlColorIndexAutomatic
    xlHoja1.Range("C36:E37").BorderAround xlContinuous, xlThin, xlColorIndexAutomatic
    xlHoja1.Range(xlHoja1.Cells(36, 5), xlHoja1.Cells(37, 8)).Borders(xlInsideVertical).LineStyle = xlContinuous
    xlHoja1.Range("F36:I37").BorderAround xlContinuous, xlThin, xlColorIndexAutomatic
    xlHoja1.Range(xlHoja1.Cells(36, 4), xlHoja1.Cells(37, 12)).HorizontalAlignment = xlCenter
    
    xlHoja1.Cells(38, 1) = "A. Estandarizado"
    xlHoja1.Cells(39, 1) = "   mas metodo"
    xlHoja1.Cells(40, 1) = "   simplificado"
    xlHoja1.Range(xlHoja1.Cells(38, 1), xlHoja1.Cells(38, 2)).Merge True
    xlHoja1.Range(xlHoja1.Cells(39, 1), xlHoja1.Cells(39, 2)).Merge True
    xlHoja1.Range(xlHoja1.Cells(40, 1), xlHoja1.Cells(40, 2)).Merge True
    xlHoja1.Cells(38, 3) = " Agregada Total Divisas"
    xlHoja1.Cells(39, 3) = " En Oro (Larga o Corta)"
    xlHoja1.Cells(40, 3) = " De opciones compradas"
    xlHoja1.Range(xlHoja1.Cells(38, 3), xlHoja1.Cells(38, 5)).Merge True
    xlHoja1.Range(xlHoja1.Cells(39, 3), xlHoja1.Cells(39, 5)).Merge True
    xlHoja1.Range(xlHoja1.Cells(40, 3), xlHoja1.Cells(40, 5)).Merge True
    xlHoja1.Range("F38:F38").Formula = "=I14"
    xlHoja1.Range(xlHoja1.Cells(38, 7), xlHoja1.Cells(43, 7)).NumberFormat = "0.00%"
    xlHoja1.Cells(38, 7) = 0.091
    xlHoja1.Cells(39, 7) = 0.091
    xlHoja1.Cells(40, 7) = 1
    'xlHoja1.Cells(38, 7) = "(I)"
    'xlHoja1.Cells(39, 7) = "(II)"
    'xlHoja1.Cells(40, 8) = "(III)"
    xlHoja1.Range("H38:H38").Formula = "=F38*G38"
    xlHoja1.Range("A38:B40").BorderAround xlContinuous, xlThin, xlColorIndexAutomatic
    xlHoja1.Range(xlHoja1.Cells(38, 3), xlHoja1.Cells(40, 5)).Borders(xlInsideVertical).LineStyle = xlContinuous
    xlHoja1.Range("F38:I40").BorderAround xlContinuous, xlThin, xlColorIndexAutomatic
    
    xlHoja1.Columns("M:M").ColumnWidth = 14.29
    
    xlHoja1.Range("I38:I39").Interior.ColorIndex = 48
    xlHoja1.Range("I38:I39").Interior.Pattern = xlSolid
    xlHoja1.Range("I40:I40").Formula = "=H38"
    
    xlHoja1.Cells(41, 1) = "B. Estandarizado mas"
    xlHoja1.Cells(42, 1) = "   metodologia"
    xlHoja1.Cells(43, 1) = "   Delta plus"
    xlHoja1.Range(xlHoja1.Cells(41, 1), xlHoja1.Cells(41, 2)).Merge True
    xlHoja1.Range(xlHoja1.Cells(42, 1), xlHoja1.Cells(42, 2)).Merge True
    xlHoja1.Range(xlHoja1.Cells(43, 1), xlHoja1.Cells(43, 2)).Merge True
    xlHoja1.Range("F41:F41").Formula = "=I14"
    xlHoja1.Cells(41, 3) = " Agregada Total Divisas"
    xlHoja1.Cells(42, 3) = " En Oro (Larga o Corta)"
    xlHoja1.Cells(43, 3) = " Monto Gamma y Vega de Opciones"
    xlHoja1.Range(xlHoja1.Cells(41, 3), xlHoja1.Cells(41, 5)).Merge True
    xlHoja1.Range(xlHoja1.Cells(42, 3), xlHoja1.Cells(42, 5)).Merge True
    xlHoja1.Range(xlHoja1.Cells(43, 3), xlHoja1.Cells(43, 5)).Merge True
    xlHoja1.Cells(41, 7) = 0.091
    xlHoja1.Cells(42, 7) = 0.091
    xlHoja1.Cells(43, 7) = 1
    xlHoja1.Range("H41:H41").Formula = "=F41*G41"
    'xlHoja1.Cells(41, 11) = "(I)"
    'xlHoja1.Cells(42, 11) = "(II)"
    'xlHoja1.Cells(43, 12) = "(III)"
    
    xlHoja1.Range("I41:I42").Interior.ColorIndex = 48
    xlHoja1.Range("I41:I42").Interior.Pattern = xlSolid
    xlHoja1.Range("I43:I43").Formula = "=H41"
    
    xlHoja1.Range("A41:B43").BorderAround xlContinuous, xlThin, xlColorIndexAutomatic
    xlHoja1.Range("C41:H43").BorderAround xlContinuous, xlThin, xlColorIndexAutomatic
    xlHoja1.Range("G38:G43").BorderAround xlContinuous, xlThin, xlColorIndexAutomatic
    xlHoja1.Range("H38:H43").BorderAround xlContinuous, xlThin, xlColorIndexAutomatic
    xlHoja1.Range(xlHoja1.Cells(41, 9), xlHoja1.Cells(43, 10)).Borders(xlInsideVertical).LineStyle = xlContinuous
    xlHoja1.Range("F41:I43").BorderAround xlContinuous, xlThin, xlColorIndexAutomatic
    
    xlHoja1.Cells(46, 1) = "III.MODELOS DE VALOR EN RIESGO"
    xlHoja1.Range(xlHoja1.Cells(46, 1), xlHoja1.Cells(52, 14)).Font.Bold = True
    
    xlHoja1.Cells(48, 1) = "DIVISAS"
    xlHoja1.Range(xlHoja1.Cells(48, 1), xlHoja1.Cells(52, 1)).Cells.Merge
    xlHoja1.Range(xlHoja1.Cells(48, 1), xlHoja1.Cells(52, 10)).HorizontalAlignment = xlCenter
    xlHoja1.Range(xlHoja1.Cells(48, 1), xlHoja1.Cells(52, 10)).VerticalAlignment = xlCenter
    xlHoja1.Cells(48, 2) = "METODO REGULATORIO"
    xlHoja1.Range(xlHoja1.Cells(48, 2), xlHoja1.Cells(48, 7)).Merge True
    xlHoja1.Cells(49, 2) = "POSICION"
    xlHoja1.Cells(50, 2) = "GLOBAL "
    xlHoja1.Cells(51, 2) = "EN ME"
    xlHoja1.Cells(52, 2) = "(H)"
    xlHoja1.Cells(49, 3) = ""
    xlHoja1.Cells(50, 3) = "VOLATILIDAD"
    xlHoja1.Cells(51, 3) = ""
    xlHoja1.Cells(52, 3) = ""
    
    xlHoja1.Cells(50, 4) = "VALOR EN RIESGO (VaR)"
    xlHoja1.Range(xlHoja1.Cells(49, 4), xlHoja1.Cells(51, 7)).Merge True
    
    xlHoja1.Cells(52, 4) = "POS. GLOBAL"
    xlHoja1.Cells(52, 5) = "GAMMA"
    xlHoja1.Cells(52, 6) = "VEGA"
    xlHoja1.Cells(52, 7) = "TOTAL (VAR)"
    
    xlHoja1.Cells(48, 8) = "MODELO INTERNO"
    xlHoja1.Range(xlHoja1.Cells(48, 8), xlHoja1.Cells(48, 10)).Merge True
    xlHoja1.Cells(49, 8) = "EXPOSICION"
    xlHoja1.Cells(50, 8) = "EN"
    xlHoja1.Cells(51, 8) = "ME"
    xlHoja1.Cells(49, 9) = ""
    xlHoja1.Cells(50, 9) = "VOLATILIDAD"
    xlHoja1.Cells(51, 9) = ""
    xlHoja1.Cells(49, 10) = ""
    xlHoja1.Cells(50, 10) = "VAR"
    xlHoja1.Cells(51, 10) = ""
    
    xlHoja1.Cells(54, 1) = "DOLAR AMERICANO"
    xlHoja1.Cells(56, 1) = "LIBRA ESTERLINA"
    xlHoja1.Cells(58, 1) = "YEN JAPONES"
    xlHoja1.Cells(60, 1) = "DOLAR CANADIENSE"
    xlHoja1.Cells(62, 1) = "EURO"
    xlHoja1.Cells(64, 1) = "OTRAS DIVISAS"
    xlHoja1.Cells(66, 1) = "ORO"
    xlHoja1.Cells(68, 1) = "TOTAL VaR"
    xlHoja1.Cells(69, 1) = "TOTAL 3*VaR"
    xlHoja1.Cells(70, 1) = "TOTAL 3*VaR / P.E."
    
    xlHoja1.Range("A48:A70").Borders(xlEdgeLeft).LineStyle = xlContinuous
    xlHoja1.Range("A48:A70").Borders(xlEdgeLeft).Weight = xlThin
    xlHoja1.Range("A48:A70").Borders(xlEdgeLeft).ColorIndex = xlAutomatic
    
    xlHoja1.Range("A48:A70").Borders(xlEdgeRight).LineStyle = xlContinuous
    xlHoja1.Range("A48:A70").Borders(xlEdgeRight).Weight = xlThin
    xlHoja1.Range("A48:A70").Borders(xlEdgeRight).ColorIndex = xlAutomatic
    
    xlHoja1.Range("B48:B67").Borders(xlEdgeRight).LineStyle = xlContinuous
    xlHoja1.Range("B48:B67").Borders(xlEdgeRight).Weight = xlThin
    xlHoja1.Range("B48:B67").Borders(xlEdgeRight).ColorIndex = xlAutomatic
    
    xlHoja1.Range("C48:C67").Borders(xlEdgeRight).LineStyle = xlContinuous
    xlHoja1.Range("C48:C67").Borders(xlEdgeRight).Weight = xlThin
    xlHoja1.Range("C48:C67").Borders(xlEdgeRight).ColorIndex = xlAutomatic
    
    xlHoja1.Range("C48:C67").Borders(xlEdgeRight).LineStyle = xlContinuous
    xlHoja1.Range("C48:C67").Borders(xlEdgeRight).Weight = xlThin
    xlHoja1.Range("C48:C67").Borders(xlEdgeRight).ColorIndex = xlAutomatic
    
    xlHoja1.Range("D48:D67").Borders(xlEdgeRight).LineStyle = xlContinuous
    xlHoja1.Range("D48:D67").Borders(xlEdgeRight).Weight = xlThin
    xlHoja1.Range("D48:D67").Borders(xlEdgeRight).ColorIndex = xlAutomatic
    
    xlHoja1.Range("E48:E67").Borders(xlEdgeRight).LineStyle = xlContinuous
    xlHoja1.Range("E48:E67").Borders(xlEdgeRight).Weight = xlThin
    xlHoja1.Range("E48:E67").Borders(xlEdgeRight).ColorIndex = xlAutomatic
    
    xlHoja1.Range("F48:F70").Borders(xlEdgeRight).LineStyle = xlContinuous
    xlHoja1.Range("F48:F70").Borders(xlEdgeRight).Weight = xlThin
    xlHoja1.Range("F48:F70").Borders(xlEdgeRight).ColorIndex = xlAutomatic
    
    xlHoja1.Range("G48:G70").Borders(xlEdgeRight).LineStyle = xlContinuous
    xlHoja1.Range("G48:G70").Borders(xlEdgeRight).Weight = xlThin
    xlHoja1.Range("G48:G70").Borders(xlEdgeRight).ColorIndex = xlAutomatic
    
    xlHoja1.Range("H48:H70").Borders(xlEdgeRight).LineStyle = xlContinuous
    xlHoja1.Range("H48:H70").Borders(xlEdgeRight).Weight = xlThin
    xlHoja1.Range("H48:H70").Borders(xlEdgeRight).ColorIndex = xlAutomatic
    
    xlHoja1.Range("I48:I70").Borders(xlEdgeRight).LineStyle = xlContinuous
    xlHoja1.Range("I48:I70").Borders(xlEdgeRight).Weight = xlThin
    xlHoja1.Range("I48:I70").Borders(xlEdgeRight).ColorIndex = xlAutomatic
    
    xlHoja1.Range("J48:J70").Borders(xlEdgeRight).LineStyle = xlContinuous
    xlHoja1.Range("J48:J70").Borders(xlEdgeRight).Weight = xlThin
    xlHoja1.Range("J48:J70").Borders(xlEdgeRight).ColorIndex = xlAutomatic
    
    xlHoja1.Range("D52:G52").Borders(xlEdgeTop).LineStyle = xlContinuous
    xlHoja1.Range("D52:G52").Borders(xlEdgeTop).Weight = xlThin
    xlHoja1.Range("D52:G52").Borders(xlEdgeTop).ColorIndex = xlAutomatic
    
    xlHoja1.Range("A48:J48").Borders(xlEdgeTop).LineStyle = xlContinuous
    xlHoja1.Range("A48:J48").Borders(xlEdgeTop).Weight = xlThin
    xlHoja1.Range("A48:J48").Borders(xlEdgeTop).ColorIndex = xlAutomatic
    
    xlHoja1.Range("A49:J49").Borders(xlEdgeTop).LineStyle = xlContinuous
    xlHoja1.Range("A49:J49").Borders(xlEdgeTop).Weight = xlThin
    xlHoja1.Range("A49:J49").Borders(xlEdgeTop).ColorIndex = xlAutomatic
    
    xlHoja1.Range("A53:J53").Borders(xlEdgeTop).LineStyle = xlContinuous
    xlHoja1.Range("A53:J53").Borders(xlEdgeTop).Weight = xlThin
    xlHoja1.Range("A53:J53").Borders(xlEdgeTop).ColorIndex = xlAutomatic
    
    xlHoja1.Range("A68:J68").Borders(xlEdgeTop).LineStyle = xlContinuous
    xlHoja1.Range("A68:J68").Borders(xlEdgeTop).Weight = xlThin
    xlHoja1.Range("A68:J68").Borders(xlEdgeTop).ColorIndex = xlAutomatic
    xlHoja1.Range("A69:J69").Borders(xlEdgeTop).LineStyle = xlContinuous
    xlHoja1.Range("A69:J69").Borders(xlEdgeTop).Weight = xlThin
    xlHoja1.Range("A69:J69").Borders(xlEdgeTop).ColorIndex = xlAutomatic
    xlHoja1.Range("A70:J70").Borders(xlEdgeTop).LineStyle = xlContinuous
    xlHoja1.Range("A70:J70").Borders(xlEdgeTop).Weight = xlThin
    xlHoja1.Range("A70:J70").Borders(xlEdgeTop).ColorIndex = xlAutomatic
    xlHoja1.Range("A71:J69").Borders(xlEdgeTop).LineStyle = xlContinuous
    xlHoja1.Range("A71:J71").Borders(xlEdgeTop).Weight = xlThin
    xlHoja1.Range("A71:J71").Borders(xlEdgeTop).ColorIndex = xlAutomatic
    
    xlHoja1.Range("B68:F70").Cells.Merge
    xlHoja1.Range("H68:I70").Cells.Merge
    
    xlHoja1.Range("B68:F70").Interior.ColorIndex = 15
    xlHoja1.Range("B68:F70").Interior.Pattern = xlSolid
    xlHoja1.Range("H68:I70").Interior.ColorIndex = 15
    xlHoja1.Range("H68:I70").Interior.Pattern = xlSolid
    xlHoja1.Range("G53:G67").Interior.ColorIndex = 15
    xlHoja1.Range("G53:G67").Interior.Pattern = xlSolid
    
    xlHoja1.Range("B76:D76").Cells.Merge
    xlHoja1.Range("G76:I76").Cells.Merge
    xlHoja1.Range("L76:N76").Cells.Merge
    
    xlHoja1.Range("B76:D76").Borders(xlEdgeTop).LineStyle = xlContinuous
    xlHoja1.Range("B76:D76").Borders(xlEdgeTop).Weight = xlThin
    xlHoja1.Range("B76:D76").Borders(xlEdgeTop).ColorIndex = xlAutomatic
    xlHoja1.Range("G76:I76").Borders(xlEdgeTop).LineStyle = xlContinuous
    xlHoja1.Range("G76:I76").Borders(xlEdgeTop).Weight = xlThin
    xlHoja1.Range("G76:I76").Borders(xlEdgeTop).ColorIndex = xlAutomatic
    xlHoja1.Range("L76:N76").Borders(xlEdgeTop).LineStyle = xlContinuous
    xlHoja1.Range("L76:N76").Borders(xlEdgeTop).Weight = xlThin
    xlHoja1.Range("L76:N76").Borders(xlEdgeTop).ColorIndex = xlAutomatic
    
    xlHoja1.Cells(76, 2) = "GERENTE GENERAL"
    xlHoja1.Cells(76, 7) = "CONTADOR GENERAL"
    xlHoja1.Cells(76, 12) = "ELABORADO POR"
    
    xlHoja1.Range("A76:L76").HorizontalAlignment = xlCenter
    xlHoja1.Range("A76:L76").Font.Bold = True
    
    
   ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1
   CargaArchivo lsArchivo, App.path & "\Spooler\"
   End If
   MousePointer = 0
Exit Sub
ErrImprime:
   MsgBox TextErr(Err.Description), vbInformation, "!Aviso!"
   MousePointer = 0
End Sub

Private Sub CargaDatosReporte()
    Dim oRep As New DRepFormula
    Dim R    As New ADODB.Recordset
    Dim nReg As Integer
    nContBalReporte = 0
    ReDim EstBalReporte(0)
        Set R = oRep.CargaRepFormula(, sCodOpeReporte, , , , "STRING")
        'Set R = oRep.CargaRepFormula(, sCodOpeReporte, , , , lsFormaOrdenRep)
          ProgressShow prgBarra, Me, eCap_CaptionPercent
          
          prgBarra.Titulo = sTituRepoFormula
          prgBarra.SubTitulo = "Cargando datos... "
          prgBarra.Max = R.RecordCount
          Do While Not R.EOF
              nContBalReporte = nContBalReporte + 1
              ReDim Preserve EstBalReporte(nContBalReporte)
              EstBalReporte(nContBalReporte - 1).cCodCta = Trim(R!cCodigo)
              If lbLimpiaDescrip Then
                EstBalReporte(nContBalReporte - 1).cDescrip = Trim(R!cDescrip)
              Else
                EstBalReporte(nContBalReporte - 1).cDescrip = R!cDescrip
              End If
              EstBalReporte(nContBalReporte - 1).cFormula = DepuraEquivalentes(Trim(R!cFormula))
              EstBalReporte(nContBalReporte - 1).bIngManual = IIf(R!bIngresoManual, True, False)
              R.MoveNext
              prgBarra.Progress nContBalReporte
          Loop
          ProgressClose prgBarra, Me
        RSClose R
        Set oRep = Nothing
    Set R = Nothing
End Sub

Private Sub GeneraReporte770091()
Dim i As Integer
Dim K As Integer
Dim j As Integer
Dim CTemp As String
Dim sSql As String
Dim sSql2 As String
Dim R As New ADODB.Recordset
Dim CoHist As Double
Dim CoAj   As Double
Dim MNAj   As Double
Dim MNHist As Double
Dim MED    As Double
Dim MES    As Double
Dim CadSql As String
Dim CadFormula1 As String
Dim CadFormula2 As String
Dim CadFormula3 As String
Dim CadFormula4 As String
Dim CadFormula5 As String
Dim CadFormula6 As String
Dim L As ListItem
Dim nFormula As New NInterpreteFormula
Dim nImporte As Currency
Dim nTipC As Currency
'***Agregado por ELRO el 20111226, según Acta N° 352-2011/TI-D
Dim MEDAnterior As Double
Dim MESAnterior As Double
Dim CadFormula7 As String
Dim nTipCambioAnterior As Currency
'***Fin Agregado por ELRO*************************************

   ProgressShow prgBarra, Me, ePCap_CaptionPercent
   prgBarra.Titulo = sTituRepoFormula
   prgBarra.SubTitulo = "Cargando datos... "
   prgBarra.Max = nContBalReporte - 1
   DoEvents
   
   'MIOL 20120705, SEGUN RQ12122 *****************************
   If sCodOpeReporte <> "770091" And sCodOpeReporte <> "763412" Then
'   If sCodOpeReporte <> "770091" Then
        dFecha = DateAdd("m", 1, CDate("01/" & Format(lnMes, "00") & "/" & Format(lnAnio, "0000"))) - 1
   'END MIOL *************************************************
   Else
        dFecha = txtFechaCorte.Text
   End If
   'oDBal.EliminaBalanceTemp CInt(sTipoRepoFormula), "0" '** para el a es 1
   'oDBal.InsertaBalanceTmpSaldos CInt(sTipoRepoFormula), "0", Format(dFecha, gsFormatoFecha)
       
   nTipC = oNBal.GetTipCambioBalance(Format(dFecha, gsFormatoMovFecha))
   '***Agregado por ELRO el 20111226, según Acta N° 352-2011/TI-D
    'MIOL 20120705, SEGUN RQ12122 *******************************
    If fsCodReport = "770090" Or fsCodReport = "763411" Then
'    If fsCodReport = "770090" Then
    'END MIOL ***************************************************
        nTipCambioAnterior = oNBal.GetTipCambioBalance(Format(DateAdd("d", -1, dFecha), gsFormatoMovFecha))
    End If
   '***Fin Agregado por ELRO*************************************
   'lvRep.ListItems.Clear
   ReDim MatrixReporte(0)
   ReDim Preserve MatrixReporte(0 To nContBalReporte)
   For i = 0 To nContBalReporte
      
      'Obtener Cuentas
      CTemp = ""
      nCuentasReporte = 0
      EstBalReporte(i).cFormula = DepuraFormula(EstBalReporte(i).cFormula)
      
      ReDim CuentasReporte(0)
      For K = 1 To Len(EstBalReporte(i).cFormula)
          If UCase(Mid(EstBalReporte(i).cFormula, K, 3)) = "DBO" Then
               CTemp = CTemp + Left(Mid(EstBalReporte(i).cFormula, K, InStr(1, Mid(EstBalReporte(i).cFormula, K), ")")), 3) & "." & Mid(Mid(EstBalReporte(i).cFormula, K, InStr(1, Mid(EstBalReporte(i).cFormula, K), ")")), 4)
               K = K + InStr(1, Mid(EstBalReporte(i).cFormula, K), ")") - 1
          ElseIf Mid(EstBalReporte(i).cFormula, K, 1) >= "0" And Mid(EstBalReporte(i).cFormula, K, 1) <= "9" Then
              CTemp = CTemp + Mid(EstBalReporte(i).cFormula, K, 1)
          Else
              If Len(CTemp) > 0 Then
                  nCuentasReporte = nCuentasReporte + 1
                  ReDim Preserve CuentasReporte(nCuentasReporte)
                  CuentasReporte(nCuentasReporte - 1).cCta = CTemp
              End If
              CTemp = ""
          End If
      Next K
      If Len(CTemp) > 0 Then
          nCuentasReporte = nCuentasReporte + 1
          ReDim Preserve CuentasReporte(nCuentasReporte)
          CuentasReporte(nCuentasReporte - 1).cCta = CTemp
      End If
      'Carga Valres de las Cuentas
      For K = 0 To nCuentasReporte - 1
        If UCase(Left(CuentasReporte(K).cCta, 4)) = "DBO." Then
           'Moneda Nacional Historico
           MNHist = nFormula.EjecutaFuncion(CuentasReporte(K).cCta)
           'Moneda Extranjera Soles
           MES = nFormula.EjecutaFuncion(CuentasReporte(K).cCta)
           'Moneda Nacional Ajustado
           MNAj = nFormula.EjecutaFuncion(CuentasReporte(K).cCta)
           'Consolidado Historico
           CoHist = nFormula.EjecutaFuncion(CuentasReporte(K).cCta)
           'Consolidado Ajustado
           CoAj = nFormula.EjecutaFuncion(CuentasReporte(K).cCta)
        Else
           'Moneda Nacional Historico
           MNHist = oNBal.CalculaSaldoCuenta(CuentasReporte(K).cCta, "[13]", sTipoRepoFormula, "0")
           'Moneda Extranjera Soles
           MES = oNBal.CalculaSaldoCuenta(CuentasReporte(K).cCta, "2", sTipoRepoFormula, "0")
           '***Agregado por ELRO el 20111226, según Acta N° 352-2011/TI-D
           'MIOL 20120705, SEGUN RQ12122 ********************************
            If fsCodReport = "770090" Or fsCodReport = "763411" Then
'           If fsCodReport = "770090" Then
           'END MIOL ****************************************************
                MESAnterior = oNBal.CalculaSaldoCuentaBalanceTempHist(CuentasReporte(K).cCta, "2", sTipoRepoFormula, "0", , , Format(DateAdd("d", -1, dFecha), gsFormatoMovFecha))
           End If
           '***Fin Agregado por ELRO*************************************
           'Moneda Nacional Ajustado
           MNAj = oNBal.CalculaSaldoCuenta(CuentasReporte(K).cCta, "[136]", sTipoRepoFormula, "0")
           'Consolidado Historico
           CoHist = oNBal.CalculaSaldoCuenta(CuentasReporte(K).cCta, "[123]", sTipoRepoFormula, "0")
           'Consolidado Ajustado
           CoAj = oNBal.CalculaSaldoCuenta(CuentasReporte(K).cCta, "[1236]", sTipoRepoFormula, "0")
           
           If nTipC <> 0 Then
              MED = Round(MES / nTipC, 2)
              '***Agregado por ELRO el 20111226, según Acta N° 352-2011/TI-D
              'MIOL 20120705, SEGUN RQ12122 ********************************
              If fsCodReport = "770090" Or fsCodReport = "763411" Then
'              If fsCodReport = "770090" Then
              'END MIOL ****************************************************
                MEDAnterior = Round(MESAnterior / nTipCambioAnterior, 2)
              End If
              '***Fin Agregado por ELRO*************************************
           End If
        End If
        
        'Actualiza Montos
        CuentasReporte(K).nMES = MES
        CuentasReporte(K).nMED = MED
        CuentasReporte(K).nMNAj = MNAj
        CuentasReporte(K).nMNHist = MNHist
        CuentasReporte(K).nCOAj = CoAj
        CuentasReporte(K).nCOHist = CoHist
        '***Agregado por ELRO el 20111226, según Acta N° 352-2011/TI-D
        'MIOL 20120705, SEGUN RQ12122 ********************************
        If fsCodReport = "770090" Or fsCodReport = "763411" Then
'        If fsCodReport = "770090" Then
        'END MIOL ****************************************************
            CuentasReporte(K).nMEDAnterior = MEDAnterior
        End If
        '***Fin Agregado por ELRO*************************************
      Next K
      'Genero las 3 formulas para las 3 monedas
      CTemp = ""
      CadFormula1 = ""
      CadFormula2 = ""
      CadFormula3 = ""
      CadFormula4 = ""
      CadFormula5 = ""
      CadFormula6 = ""
      '***Agregado por ELRO el 20111226, según Acta N° 352-2011/TI-D
      'MIOL 20120705, SEGUN RQ12122 ********************************
       If fsCodReport = "770090" Or fsCodReport = "763411" Then
'      If fsCodReport = "770090" Then
        CadFormula7 = ""
       End If
      '***Fin Agregado por ELRO*************************************
                  
        For K = 1 To Len(EstBalReporte(i).cFormula)
            If UCase(Mid(EstBalReporte(i).cFormula, K, 3)) = "DBO" Then
                 CTemp = CTemp + Left(Mid(EstBalReporte(i).cFormula, K, InStr(1, Mid(EstBalReporte(i).cFormula, K), ")")), 3) & "." & Mid(Mid(EstBalReporte(i).cFormula, K, InStr(1, Mid(EstBalReporte(i).cFormula, K), ")")), 4)
                 K = K + InStr(1, Mid(EstBalReporte(i).cFormula, K), ")") - 1
            ElseIf (Mid(EstBalReporte(i).cFormula, K, 1) >= "0" And Mid(EstBalReporte(i).cFormula, K, 1) <= "9") Or (Mid(EstBalReporte(i).cFormula, K, 1) = ".") Then
                CTemp = CTemp + Mid(EstBalReporte(i).cFormula, K, 1)
            ElseIf Mid(EstBalReporte(i).cFormula, K, 1) = "$" Then
                 j = InStr(K + 1, EstBalReporte(i).cFormula, "$") - 1
                 nImporte = Mid(EstBalReporte(i).cFormula, K + 1, j - K)
                 CadFormula1 = CadFormula1 + Format(nImporte, gsFormatoNumeroDato)
                 CadFormula2 = CadFormula2 + Format(nImporte, gsFormatoNumeroDato)
                 CadFormula3 = CadFormula3 + Format(nImporte, gsFormatoNumeroDato)
                 CadFormula4 = CadFormula4 + Format(nImporte, gsFormatoNumeroDato)
                 CadFormula5 = CadFormula5 + Format(nImporte, gsFormatoNumeroDato)
                 CadFormula6 = CadFormula6 + Format(nImporte, gsFormatoNumeroDato)
                 '***Agregado por ELRO el 20111226, según Acta N° 352-2011/TI-D
                 'MIOL 20120705, SEGUN RQ12122 ********************************
                 If fsCodReport = "770090" Or fsCodReport = "763411" Then
'                 If fsCodReport = "770090" Then
                 'END MIOL ****************************************************
                    CadFormula7 = CadFormula7 + Format(nImporte, gsFormatoNumeroDato)
                 End If
                 '***Fin Agregado por ELRO*************************************
                 K = j + 1
            Else
                 If Len(CTemp) > 0 Then
                     'busca su equivalente en monto
                     For j = 0 To nCuentasReporte
                         If CuentasReporte(j).cCta = CTemp Then
                             CadFormula1 = CadFormula1 + Format(CuentasReporte(j).nMNAj, gsFormatoNumeroDato)
                             CadFormula2 = CadFormula2 + Format(CuentasReporte(j).nMNHist, gsFormatoNumeroDato)
                             CadFormula3 = CadFormula3 + Format(CuentasReporte(j).nMES, gsFormatoNumeroDato)
                             CadFormula4 = CadFormula4 + Format(CuentasReporte(j).nMED, gsFormatoNumeroDato)
                             CadFormula5 = CadFormula5 + Format(CuentasReporte(j).nCOAj, gsFormatoNumeroDato)
                             CadFormula6 = CadFormula6 + Format(CuentasReporte(j).nCOHist, gsFormatoNumeroDato)
                             '***Agregado por ELRO el 20111226, según Acta N° 352-2011/TI-D
                             'MIOL 20120705, SEGUN RQ12122 ********************************
                             If fsCodReport = "770090" Or fsCodReport = "763411" Then
'                             If fsCodReport = "770090" Then
                             'END MIOL ****************************************************
                                CadFormula7 = CadFormula7 + Format(CuentasReporte(j).nMEDAnterior, gsFormatoNumeroDato)
                             End If
                             '***Fin Agregado por ELRO*************************************
                             Exit For
                         End If
                     Next j
                 End If
                 CTemp = ""
                 CadFormula1 = CadFormula1 + Mid(EstBalReporte(i).cFormula, K, 1)
                 CadFormula2 = CadFormula2 + Mid(EstBalReporte(i).cFormula, K, 1)
                 CadFormula3 = CadFormula3 + Mid(EstBalReporte(i).cFormula, K, 1)
                 CadFormula4 = CadFormula4 + Mid(EstBalReporte(i).cFormula, K, 1)
                 CadFormula5 = CadFormula5 + Mid(EstBalReporte(i).cFormula, K, 1)
                 CadFormula6 = CadFormula6 + Mid(EstBalReporte(i).cFormula, K, 1)
                 '***Agregado por ELRO el 20111226, según Acta N° 352-2011/TI-D
                 'MIOL 20120705, SEGUN RQ12122 ********************************
                 If fsCodReport = "770090" Or fsCodReport = "763411" Then
'                 If fsCodReport = "770090" Then
                 'END MIOL ****************************************************
                    CadFormula7 = CadFormula7 + Mid(EstBalReporte(i).cFormula, K, 1)
                 End If
                 '***Fin Agregado por ELRO*************************************
            End If
        Next K
         
         If Len(CTemp) > 0 Then
             'busca su equivalente en monto
             For j = 0 To nCuentasReporte
                 If CuentasReporte(j).cCta = CTemp Then
                     CadFormula1 = CadFormula1 + Format(CuentasReporte(j).nMNAj, gsFormatoNumeroDato)
                     CadFormula2 = CadFormula2 + Format(CuentasReporte(j).nMNHist, gsFormatoNumeroDato)
                     CadFormula3 = CadFormula3 + Format(CuentasReporte(j).nMES, gsFormatoNumeroDato)
                     CadFormula4 = CadFormula4 + Format(CuentasReporte(j).nMED, gsFormatoNumeroDato)
                     CadFormula5 = CadFormula5 + Format(CuentasReporte(j).nCOAj, gsFormatoNumeroDato)
                     CadFormula6 = CadFormula6 + Format(CuentasReporte(j).nCOHist, gsFormatoNumeroDato)
                     '***Agregado por ELRO el 20111226, según Acta N° 352-2011/TI-D
                     'MIOL 20120705, SEGUN RQ12122 ********************************
                     If fsCodReport = "770090" Or fsCodReport = "763411" Then
'                     If fsCodReport = "770090" Then
                     'END MIOL ****************************************************
                        CadFormula7 = CadFormula7 + Format(CuentasReporte(j).nMEDAnterior, gsFormatoNumeroDato)
                     End If
                     '***Fin Agregado por ELRO*************************************
                     Exit For
                 End If
             Next j
         End If
         MNAj = 0: MNHist = 0: MES = 0: MED = 0: CoAj = 0: CoHist = 0: MEDAnterior = 0
         
         MNAj = nFormula.ExprANum(CadFormula1, EstBalReporte(i).cCodCta)
         MNHist = nFormula.ExprANum(CadFormula2, EstBalReporte(i).cCodCta)
         MES = nFormula.ExprANum(CadFormula3, EstBalReporte(i).cCodCta)
         MED = nFormula.ExprANum(CadFormula4, EstBalReporte(i).cCodCta)
         CoAj = nFormula.ExprANum(CadFormula5, EstBalReporte(i).cCodCta)
         CoHist = nFormula.ExprANum(CadFormula6, EstBalReporte(i).cCodCta)
         '***Agregado por ELRO el 20111224, según Acta N° 352-2011/TI-D
          'MIOL 20120705, SEGUN RQ12122 ********************************
          If fsCodReport = "770090" Or fsCodReport = "763411" Then
'          If fsCodReport = "770090" Then
          'END MIOL ****************************************************
             MEDAnterior = nFormula.ExprANum(CadFormula7, EstBalReporte(i).cCodCta)
             If EstBalReporte(i).cCodCta = "1" Then
                fnSaldoActivoMN = MES
                fnSaldoActivoME = MED
             End If
            If EstBalReporte(i).cCodCta = "2" Then
                fnSaldoPasivoMN = MES
                fnSaldoPasivoME = MED
             End If
         End If
         '***Fin Agregado por ELRO*************************************
         'ReDim Preserve MatrixReporte(0 To I)
         
        'Asignacion de Valores
         'Set L = lvRep.ListItems.Add(, , EstBalReporte(I).cCodCta)
         MatrixReporte(i).cCta = EstBalReporte(i).cCodCta
         MatrixReporte(i).cDescrip = EstBalReporte(i).cDescrip
         MatrixReporte(i).nCOHist = Format(CoHist, gsFormatoNumeroView)
         MatrixReporte(i).nCOAj = Format(CoAj, gsFormatoNumeroView)
         MatrixReporte(i).nMNHist = Format(MNHist, gsFormatoNumeroView)
         MatrixReporte(i).nMNAj = Format(MNAj, gsFormatoNumeroView)
         MatrixReporte(i).nMES = Format(MES, gsFormatoNumeroView)
         MatrixReporte(i).nMED = Format(MED, gsFormatoNumeroView)
         MatrixReporte(i).bIngManual = IIf(EstBalReporte(i).bIngManual, "1", "0")
         '***Agregado por ELRO el 20111226, según Acta N° 352-2011/TI-D
         'MIOL 20120705, SEGUN RQ12122 ********************************
         If fsCodReport = "770090" Or fsCodReport = "763411" Then
'         If fsCodReport = "770090" Then
         'END MIOL ****************************************************
            MatrixReporte(i).nMEDAnterior = Format(MEDAnterior, gsFormatoNumeroView)
            MatrixReporte(i).nDiferencia = MatrixReporte(i).nMED - MatrixReporte(i).nMEDAnterior
         End If
         '***Fin Agregado por ELRO*************************************
'         L.SubItems(1) = EstBalReporte(I).cDescrip
'         L.SubItems(2) = Format(CoHist, gsFormatoNumeroView)
'         L.SubItems(3) = Format(CoAj, gsFormatoNumeroView)
'         L.SubItems(4) = Format(MNHist, gsFormatoNumeroView)
'         L.SubItems(5) = Format(MNAj, gsFormatoNumeroView)
'         L.SubItems(6) = Format(MES, gsFormatoNumeroView)
'         L.SubItems(7) = Format(MED, gsFormatoNumeroView)
'         L.SubItems(8) = IIf(EstBalReporte(I).bIngManual, "1", "0")
      nCuentasReporte = i
      prgBarra.SubTitulo = "Intepretando Fórmulas... "
      prgBarra.Progress i
   Next i
   'nCuentasReporte = I
   ProgressClose prgBarra, Me
Set nFormula = Nothing
End Sub

Private Sub GenRepResumenContCartera770091()
Dim nCol  As Integer
Dim sCol  As String

Dim lsArchivo   As String
Dim lbLibroOpen As Boolean
Dim N           As Integer
Dim nPosicion As Integer
Dim oDBalanceContA As DbalanceCont
Set oDBalanceContA = New DbalanceCont
'***Agregado por ELRO el 20111226, según Acta N° 352-2011/TI-D
Dim lnPosicion2 As Integer
Dim lsSobreVenta, lsSobreCompra  As String
'***Fin Agregado por ELRO*************************************
lCOAj = frmRepBaseFormulaOpc.plCOAj
lCOHi = frmRepBaseFormulaOpc.plCOHist
lMNAj = frmRepBaseFormulaOpc.plMNAj
lMNHi = frmRepBaseFormulaOpc.plMNHist
lMES = frmRepBaseFormulaOpc.plMES
lMED = frmRepBaseFormulaOpc.plMED
'***Agregado por ELRO el 20111226, según Acta N° 352-2011/TI-D
'MIOL 20120705, SEGUN RQ12122 ********************************
If fsCodReport = "770090" Or fsCodReport = "763411" Then
'If fsCodReport = "770090" Then
    fnMEDAnterior = frmRepBaseFormulaOpc.pfnMEDAnterior
    fnDiferencia = frmRepBaseFormulaOpc.pfnDiferencia
End If
'***Fin Agregado por ELRO*************************************

On Error GoTo ErrImprime
'***Modificado por ELRO el 20111226, según Acta N° 352-2011/TI-D
'nPosicion = 8
'sCol = Chr(67 + lCOAj + lCOHi + lMNAj + lMNHi + lMES + lMED )
 'MIOL 20120705, SEGUN RQ12122 ********************************
 If fsCodReport = "770090" Or fsCodReport = "763411" Then
' If fsCodReport = "770090" Then
 'END MIOL ****************************************************
    nPosicion = 16
    lnPosicion2 = 16
    sCol = Chr(67 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 2)
 Else
    nPosicion = 8
    sCol = Chr(67 + lCOAj + lCOHi + lMNAj + lMNHi + lMES + lMED)
 End If
'***Fin Modificado por ELRO*************************************

   MousePointer = 11
   '***Modificado por ELRO el 20111226, según Acta N° 352-2011/TI-D
   'lsArchivo = App.path & "\Spooler\REP770091" & "_" & Format(lnMes, "00") & lnAnio & ".xls"
   'OleExcel.Class = "ExcelWorkSheet"
   'lbLibroOpen = ExcelBegin(lsArchivo, xlAplicacion, xlLibro)
    If fsCodReport = "770090" Then
        lsArchivo = App.path & "\Spooler\Anexo9.xls"
        OleExcel.Class = "ExcelWorkSheet"
        lbLibroOpen = ExcelBegin(lsArchivo, xlAplicacion, xlLibro, False)
    'MIOL 20120705, SEGUN RQ12122 **********************************
    ElseIf fsCodReport = "763411" Then
        lsArchivo = App.path & "\Spooler\PosCambiariaDetalle.xls"
        OleExcel.Class = "ExcelWorkSheet"
        lbLibroOpen = ExcelBegin(lsArchivo, xlAplicacion, xlLibro, False)
    'END MIOL ******************************************************
    Else
        lsArchivo = App.path & "\Spooler\REP770091" & "_" & Format(lnMes, "00") & lnAnio & ".xls"
        OleExcel.Class = "ExcelWorkSheet"
        lbLibroOpen = ExcelBegin(lsArchivo, xlAplicacion, xlLibro)
    End If
   
   '***Fin Modificado por ELRO*************************************
   If lbLibroOpen Then
      Set xlHoja1 = xlLibro.Worksheets(1)
      ExcelAddHoja "R" & Right(gsOpeCod, 1), xlLibro, xlHoja1
      '***Modificado por ELRO el 20111226, según Acta N° 352-2011/TI-D
      'CabeceraExcel sCol
      'MIOL 20120705, SEGUN RQ12122 **********************************
      If fsCodReport = "770090" Or fsCodReport = "763411" Then
'      If fsCodReport = "770090" Then
      'END MIOL ******************************************************
        xlHoja1.Range("B4:B4") = "MES " & txtMes & " " & txtAnio
        xlHoja1.Range("C7:C7") = "PATRIMONIO EFECTIVO " & UCase(lblMes) & " " & lblAnio
        xlHoja1.Range("D6:D6").Formula = "=" & fnSaldoActivoMN & "-" & fnSaldoPasivoMN
        xlHoja1.Range("E6:E6").Formula = "=" & fnSaldoActivoME & "-" & fnSaldoPasivoME
        xlHoja1.Range("D7:D7") = Format(lblPatrimonio, gsFormatoNumeroDato)
        xlHoja1.Range("E7:E7").Formula = "=D6/E6"
        xlHoja1.Range("D8:D8").Formula = "=(D6/D7)*100"
        xlHoja1.Range("L14:L15") = "M.EXTRANJERAExp. en ME del " & CStr(DateAdd("d", -1, CDate(txtFechaCorte)))
      Else
        CabeceraExcel sCol
      End If
      '***Fin Modificado por ELRO*************************************
      
      For N = 0 To nCuentasReporte - 1
      '***Modificado por ELRO el 20111226, según Acta N° 352-2011/TI-D
      'nCol = 3
      'MIOL 20120705, SEGUN RQ12122 **********************************
      If fsCodReport = "770090" Or fsCodReport = "763411" Then
'      If fsCodReport = "770090" Then
      'END MIOL ******************************************************
        nCol = 5
      Else
        nCol = 3
      End If
      '***Fin Modificado por ELRO*************************************
         With MatrixReporte(N)
         xlHoja1.Cells(N + nPosicion, 2) = .cCta
         xlHoja1.Cells(N + nPosicion, 3) = .cDescrip
         
         If lCOHi = 1 Then
            nCol = nCol + 1
            xlHoja1.Cells(N + nPosicion, nCol) = .nCOHist
         '***Agregado por ELRO el 20111226, según Acta N° 352-2011/TI-D
            'MIOL 20120705, SEGUN RQ12122 ****************************
            If fsCodReport = "770090" Or fsCodReport = "763411" Then
'            If fsCodReport = "770090" Then
            'END MIOL ************************************************
                xlHoja1.Range("F" & N + nPosicion & ":F" & nCol).ColumnWidth = 15
            End If
         Else
            'MIOL 20120705, SEGUN RQ12122 ****************************
            If fsCodReport = "770090" Or fsCodReport = "763411" Then
'            If fsCodReport = "770090" Then
            'END MIOL ************************************************
                nCol = nCol + 1
                xlHoja1.Cells(N + nPosicion, nCol) = .nCOHist
                xlHoja1.Range("F" & N + nPosicion & ":F" & nCol).ColumnWidth = 0
            End If
         '***Fin Agregado por ELRO*************************************
         End If
         If lCOAj = 1 Then
            nCol = nCol + 1
            xlHoja1.Cells(N + nPosicion, nCol) = .nCOAj
         '***Agregado por ELRO el 20111226, según Acta N° 352-2011/TI-D
            'MIOL 20120705, SEGUN RQ12122 ****************************
            If fsCodReport = "770090" Or fsCodReport = "763411" Then
'            If fsCodReport = "770090" Then
            'END MIOL ************************************************
                xlHoja1.Range("G" & N + nPosicion & ":G" & nCol).ColumnWidth = 15
            End If
         Else
            'MIOL 20120705, SEGUN RQ12122 ****************************
            If fsCodReport = "770090" Or fsCodReport = "763411" Then
'            If fsCodReport = "770090" Then
            'END MIOL ************************************************
                nCol = nCol + 1
                xlHoja1.Cells(N + nPosicion, nCol) = .nCOAj
                xlHoja1.Range("G" & N + nPosicion & ":G" & nCol).ColumnWidth = 0
            End If
         '***Fin Agregado por ELRO*************************************
         End If
         If lMNHi = 1 Then
            nCol = nCol + 1
            xlHoja1.Cells(N + nPosicion, nCol) = .nMNHist
         '***Agregado por ELRO el 20111226, según Acta N° 352-2011/TI-D
            'MIOL 20120705, SEGUN RQ12122 ****************************
            If fsCodReport = "770090" Or fsCodReport = "763411" Then
'            If fsCodReport = "770090" Then
            'END MIOL ************************************************
                xlHoja1.Range("H" & N + nPosicion & ":H" & nCol).ColumnWidth = 15
            End If
         Else
            'MIOL 20120705, SEGUN RQ12122 ****************************
            If fsCodReport = "770090" Or fsCodReport = "763411" Then
'            If fsCodReport = "770090" Then
            'END MIOL ************************************************
                nCol = nCol + 1
                xlHoja1.Cells(N + nPosicion, nCol) = .nMNHist
                xlHoja1.Range("H" & N + nPosicion & ":H" & nCol).ColumnWidth = 0
            End If
         '***Fin Agregado por ELRO*************************************
         End If
         If lMNAj = 1 Then
            nCol = nCol + 1
            xlHoja1.Cells(N + nPosicion, nCol) = .nMNAj
         '***Agregado por ELRO el 20111226, según Acta N° 352-2011/TI-D
            'MIOL 20120705, SEGUN RQ12122 ****************************
            If fsCodReport = "770090" Or fsCodReport = "763411" Then
'            If fsCodReport = "770090" Then
            'END MIOL ************************************************
                xlHoja1.Range("I" & N + nPosicion & ":I" & nCol).ColumnWidth = 15
            End If
         Else
            'MIOL 20120705, SEGUN RQ12122 ****************************
            If fsCodReport = "770090" Or fsCodReport = "763411" Then
'            If fsCodReport = "770090" Then
            'END MIOL ************************************************
                nCol = nCol + 1
                xlHoja1.Cells(N + nPosicion, nCol) = .nMNAj
                xlHoja1.Range("I" & N + nPosicion & ":I" & nCol).ColumnWidth = 0
            End If
         '***Fin Agregado por ELRO*************************************
         End If
         If lMES = 1 Then
            nCol = nCol + 1
            xlHoja1.Cells(N + nPosicion, nCol) = .nMES
         '***Agregado por ELRO el 20111226, según Acta N° 352-2011/TI-D
            'MIOL 20120705, SEGUN RQ12122 ****************************
            If fsCodReport = "770090" Or fsCodReport = "763411" Then
'            If fsCodReport = "770090" Then
            'END MIOL ************************************************
            xlHoja1.Range("J" & N + nPosicion & ":J" & nCol).ColumnWidth = 15
            End If
         Else
            'MIOL 20120705, SEGUN RQ12122 ****************************
            If fsCodReport = "770090" Or fsCodReport = "763411" Then
'            If fsCodReport = "770090" Then
            'END MIOL ************************************************
                nCol = nCol + 1
                xlHoja1.Cells(N + nPosicion, nCol) = .nMES
                xlHoja1.Range("J" & N + nPosicion & ":J" & nCol).ColumnWidth = 0
            End If
         '***Fin Agregado por ELRO*************************************
         End If
         If lMED = 1 Then
            nCol = nCol + 1
            xlHoja1.Cells(N + nPosicion, nCol) = .nMED
         '***Agregado por ELRO el 20111226, según Acta N° 352-2011/TI-D
            'MIOL 20120705, SEGUN RQ12122 ****************************
            If fsCodReport = "770090" Or fsCodReport = "763411" Then
'            If fsCodReport = "770090" Then
            'END MIOL ************************************************
                xlHoja1.Range("K" & N + nPosicion & ":K" & nCol).ColumnWidth = 15
            End If
         Else
            'MIOL 20120705, SEGUN RQ12122 ****************************
            If fsCodReport = "770090" Or fsCodReport = "763411" Then
'            If fsCodReport = "770090" Then
            'END MIOL ************************************************
                nCol = nCol + 1
                xlHoja1.Cells(N + nPosicion, nCol) = .nMED
                xlHoja1.Range("K" & N + nPosicion & ":K" & nCol).ColumnWidth = 0
            End If
         '***Fin Agregado por ELRO*************************************
         End If
         '***Agregado por ELRO el 20111226, según Acta N° 352-2011/TI-D
            'MIOL 20120705, SEGUN RQ12122 ****************************
         If fsCodReport = "770090" Or fsCodReport = "763411" Then
'            If fsCodReport = "770090" Then
            'END MIOL ************************************************
            If fnMEDAnterior = 1 Then
               nCol = nCol + 1
               xlHoja1.Cells(N + nPosicion, nCol) = .nMEDAnterior
               xlHoja1.Range("L" & N + nPosicion & ":L" & nCol).ColumnWidth = 17
            Else
               nCol = nCol + 1
               xlHoja1.Cells(N + nPosicion, nCol) = .nMEDAnterior
               xlHoja1.Range("L" & N + nPosicion & ":L" & nCol).ColumnWidth = 0
            End If
            If fnDiferencia = 1 Then
               nCol = nCol + 1
               xlHoja1.Cells(N + nPosicion, nCol) = .nDiferencia
               xlHoja1.Range("M" & N + nPosicion & ":M" & nCol).ColumnWidth = 17
            Else
               nCol = nCol + 1
               xlHoja1.Cells(N + nPosicion, nCol) = .nDiferencia
               xlHoja1.Range("M" & N + nPosicion & ":M" & nCol).ColumnWidth = 0
            End If
         End If
         '***Fin Agregado por ELRO*************************************
         End With
      Next
      N = N + nPosicion
      '***Modificado por ELRO el 20111226, según Acta N° 352-2011/TI-D
      'xlHoja1.Range("B" & 8 & ":" & sCol & N).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic
      'xlHoja1.Range("B" & 8 & ":" & sCol & N).Borders(xlInsideVertical).LineStyle = xlContinuous
      'xlHoja1.Range(xlHoja1.Cells(8, 2), xlHoja1.Cells(nCuentasReporte + 8, 2)).HorizontalAlignment = xlHAlignLeft

      'Call oDBalanceContA.InsertaCtaContSaldoDiario("Rep9_1_1", txtFechaCorte.Text, "763406", xlHoja1.Cells(8, 8))
      'Call oDBalanceContA.InsertaCtaContSaldoDiario("Rep9_2_1", txtFechaCorte.Text, "763406", xlHoja1.Cells(38, 8))

      'Call oDBalanceContA.InsertaCtaContSaldoDiario("Rep9_1_2", txtFechaCorte.Text, "763406", xlHoja1.Cells(8, 9))
      'Call oDBalanceContA.InsertaCtaContSaldoDiario("Rep9_2_2", txtFechaCorte.Text, "763406", xlHoja1.Cells(38, 9))

       'MIOL 20120705, SEGUN RQ12122 ****************************
      If fsCodReport = "770090" Or fsCodReport = "763411" Then
'       If fsCodReport = "770090" Then
       'END MIOL ************************************************
        xlHoja1.Range("B" & lnPosicion2 & ":" & sCol & N).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic
        'xlHoja1.Range("B" & lnPosicion2 & ":" & sCol & N).Borders(xlInsideVertical).LineStyle =  xlContinuous
        xlHoja1.Range(xlHoja1.Cells(lnPosicion2, 2), xlHoja1.Cells(nCuentasReporte + lnPosicion2, 2)).HorizontalAlignment = xlHAlignLeft

        Call oDBalanceContA.InsertaCtaContSaldoDiario("Rep9_1_1", txtFechaCorte.Text, "763406", fnSaldoActivoMN)
        Call oDBalanceContA.InsertaCtaContSaldoDiario("Rep9_2_1", txtFechaCorte.Text, "763406", fnSaldoPasivoMN)
      
        Call oDBalanceContA.InsertaCtaContSaldoDiario("Rep9_1_2", txtFechaCorte.Text, "763406", fnSaldoActivoME)
        Call oDBalanceContA.InsertaCtaContSaldoDiario("Rep9_2_2", txtFechaCorte.Text, "763406", fnSaldoPasivoME)

      Else
        xlHoja1.Range("B" & 8 & ":" & sCol & N).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic
        xlHoja1.Range("B" & 8 & ":" & sCol & N).Borders(xlInsideVertical).LineStyle = xlContinuous
        xlHoja1.Range(xlHoja1.Cells(8, 2), xlHoja1.Cells(nCuentasReporte + 8, 2)).HorizontalAlignment = xlHAlignLeft
        
        Call oDBalanceContA.InsertaCtaContSaldoDiario("Rep9_1_1", txtFechaCorte.Text, "763406", xlHoja1.Cells(8, 8))
        Call oDBalanceContA.InsertaCtaContSaldoDiario("Rep9_2_1", txtFechaCorte.Text, "763406", xlHoja1.Cells(38, 8))
    
        Call oDBalanceContA.InsertaCtaContSaldoDiario("Rep9_1_2", txtFechaCorte.Text, "763406", xlHoja1.Cells(8, 9))
        Call oDBalanceContA.InsertaCtaContSaldoDiario("Rep9_2_2", txtFechaCorte.Text, "763406", xlHoja1.Cells(38, 9))
        
      End If
      
      
      '***Fin Modificado por ELRO*************************************
      
      OleExcel.Class = "ExcelWorkSheet"
      ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1
      
      OleExcel.SourceDoc = lsArchivo
      OleExcel.Verb = 1
      OleExcel.Action = 1
      OleExcel.DoVerb -1
   End If
   MousePointer = 0
Exit Sub
ErrImprime:
   MsgBox TextErr(Err.Description), vbInformation, "!Aviso!"
   MousePointer = 0
End Sub




'ALPA 20110525
'***Modificado por ELRO el 20111222, según Acta N° 352-2011/TI-D
'Private Sub txtFechaCorte_Change()
'    If Len(txtFechaCorte.Text) = "10" Then
'       If Mid(txtFechaCorte.Text, 4, 2) <= 12 Then
'            txtAnio = Mid(txtFechaCorte.Text, 7, 4)
'            txtMes = IIf(CInt(Mid(txtFechaCorte.Text, 4, 2)) = "1", "Enero", IIf(CInt(Mid(txtFechaCorte.Text, 4, 2)) = "2", "Febrero", IIf(CInt(Mid(txtFechaCorte.Text, 4, 2)) = "3", "Marzo", IIf(CInt(Mid(txtFechaCorte.Text, 4, 2)) = "4", "Abril", IIf(CInt(Mid(txtFechaCorte.Text, 4, 2)) = "5", "Mayo", "")))))
'            If txtMes = "" Then
'                 txtMes = IIf(CInt(Mid(txtFechaCorte.Text, 4, 2)) = "6", "Junio", IIf(CInt(Mid(txtFechaCorte.Text, 4, 2)) = "7", "Julio", IIf(CInt(Mid(txtFechaCorte.Text, 4, 2)) = "8", "Agosto", IIf(CInt(Mid(txtFechaCorte.Text, 4, 2)) = "9", "Setiembre", IIf(CInt(Mid(txtFechaCorte.Text, 4, 2)) = "10", "Octubre", IIf(CInt(Mid(txtFechaCorte.Text, 4, 2)) = "11", "Noviembre", IIf(CInt(Mid(txtFechaCorte.Text, 4, 2)) = "12", "Diciembre", "")))))))
'            End If
'       Else
'       MsgBox "Ingresar el mes correctamente"
'       End If
'    End If
'End Sub
Private Sub txtFechaCorte_LostFocus()
    If Len(txtFechaCorte.Text) = "10" Then
        If Mid(txtFechaCorte.Text, 4, 2) <= 12 Then
            txtAnio = Mid(txtFechaCorte.Text, 7, 4)
            txtMes = IIf(CInt(Mid(txtFechaCorte.Text, 4, 2)) = "1", "Enero", IIf(CInt(Mid(txtFechaCorte.Text, 4, 2)) = "2", "Febrero", IIf(CInt(Mid(txtFechaCorte.Text, 4, 2)) = "3", "Marzo", IIf(CInt(Mid(txtFechaCorte.Text, 4, 2)) = "4", "Abril", IIf(CInt(Mid(txtFechaCorte.Text, 4, 2)) = "5", "Mayo", "")))))
            If txtMes = "" Then
                 txtMes = IIf(CInt(Mid(txtFechaCorte.Text, 4, 2)) = "6", "Junio", IIf(CInt(Mid(txtFechaCorte.Text, 4, 2)) = "7", "Julio", IIf(CInt(Mid(txtFechaCorte.Text, 4, 2)) = "8", "Agosto", IIf(CInt(Mid(txtFechaCorte.Text, 4, 2)) = "9", "Setiembre", IIf(CInt(Mid(txtFechaCorte.Text, 4, 2)) = "10", "Octubre", IIf(CInt(Mid(txtFechaCorte.Text, 4, 2)) = "11", "Noviembre", IIf(CInt(Mid(txtFechaCorte.Text, 4, 2)) = "12", "Diciembre", "")))))))
            End If
            CmdProceso.SetFocus
        Else
            MsgBox "Ingresar el mes correctamente"
        End If
    End If
End Sub
'***Fin Modificado por ELRO*************************************

