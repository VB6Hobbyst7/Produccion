VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBalanceSec 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Balance Sectorial"
   ClientHeight    =   5880
   ClientLeft      =   1290
   ClientTop       =   3510
   ClientWidth     =   8685
   Icon            =   "frmBalanceSec.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   8685
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraGenera 
      Height          =   750
      Left            =   90
      TabIndex        =   13
      Top             =   30
      Width           =   8475
      Begin VB.TextBox txtAnio 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3450
         MaxLength       =   4
         TabIndex        =   1
         Top             =   270
         Width           =   855
      End
      Begin VB.ComboBox CboMes 
         Height          =   315
         ItemData        =   "frmBalanceSec.frx":030A
         Left            =   765
         List            =   "frmBalanceSec.frx":0332
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   255
         Width           =   1755
      End
      Begin VB.CommandButton cmdProcesar 
         Caption         =   "&Generar"
         Height          =   390
         Left            =   6360
         TabIndex        =   2
         Top             =   240
         Width           =   1800
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Año :"
         Height          =   195
         Left            =   2925
         TabIndex        =   16
         Top             =   330
         Width           =   375
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Mes :"
         Height          =   195
         Left            =   255
         TabIndex        =   14
         Top             =   300
         Width           =   390
      End
   End
   Begin VB.Frame fraLista 
      Height          =   5010
      Left            =   90
      TabIndex        =   12
      Top             =   780
      Width           =   8475
      Begin VB.CommandButton cmdArchivo 
         Caption         =   "&Generar Archivo"
         Height          =   330
         Left            =   4280
         TabIndex        =   20
         Top             =   4560
         Width           =   2000
      End
      Begin MSComctlLib.ListView LstBalance 
         Height          =   4215
         Left            =   120
         TabIndex        =   3
         Top             =   210
         Width           =   8265
         _ExtentX        =   14579
         _ExtentY        =   7435
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripcion de Cuenta"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "MN. Ajustado"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "MN. Historico"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "M.E."
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "&Editar"
         Height          =   330
         Left            =   120
         TabIndex        =   9
         Top             =   4560
         Width           =   2000
      End
      Begin VB.TextBox txtMNHist 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   5640
         TabIndex        =   7
         Top             =   4080
         Width           =   1230
      End
      Begin VB.TextBox TxtMExt 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   6870
         TabIndex        =   8
         Top             =   4080
         Width           =   1485
      End
      Begin VB.TextBox TxtMNaj 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   4410
         TabIndex        =   6
         Top             =   4080
         Width           =   1230
      End
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   330
         Left            =   6360
         TabIndex        =   11
         Top             =   4560
         Width           =   2000
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "&Imprimir"
         Height          =   330
         Left            =   2200
         TabIndex        =   10
         Top             =   4560
         Width           =   2000
      End
      Begin MSComctlLib.ProgressBar PrgBarra1 
         Height          =   225
         Left            =   120
         TabIndex        =   17
         Top             =   4500
         Visible         =   0   'False
         Width           =   8235
         _ExtentX        =   14526
         _ExtentY        =   397
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   330
         Left            =   120
         TabIndex        =   18
         Top             =   4560
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   330
         Left            =   1350
         TabIndex        =   19
         Top             =   4560
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Label LblCodigo 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   120
         TabIndex        =   4
         Top             =   4080
         Width           =   1500
      End
      Begin VB.Label Lbldescrip 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1620
         TabIndex        =   5
         Top             =   4080
         Width           =   2790
      End
   End
   Begin RichTextLib.RichTextBox rtf 
      Height          =   390
      Left            =   9015
      TabIndex        =   15
      Top             =   375
      Visible         =   0   'False
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   688
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmBalanceSec.frx":039A
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
End
Attribute VB_Name = "frmBalanceSec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type TEstBal
    cCtaCod As String
    cDescrip As String
    cEquival As String
End Type
Private Type TCuentas
    cCta As String
    nMNAj As Double
    nMNHist As Double
    nMExt As Double
End Type
Dim EstBal()         As TEstBal
Dim nContBal         As Integer
Dim Cuentas()        As TCuentas
Dim nCuentas         As Integer
Dim sFecha           As String
Dim sTipoRepoFormula As String

Dim oNBal  As NBalanceCont
Dim oDBal  As DbalanceCont
Dim prgBarra As New clsProgressBar
Attribute prgBarra.VB_VarHelpID = -1

Private Function DepuraEquivalentes(psEquival As String) As String
Dim J As Integer
Dim CadTemp As String
   CadTemp = ""
   For J = 1 To Len(psEquival)
       If Mid(psEquival, J, 1) <> "." Then
           CadTemp = CadTemp + Mid(psEquival, J, 1)
       End If
   Next J
   DepuraEquivalentes = CadTemp
End Function
Private Sub CargaDatos()
Dim oRep As New DRepFormula
Dim R    As New ADODB.Recordset
Dim nReg As Integer
nContBal = 0
ReDim EstBal(0)
    Set R = oRep.CargaRepFormula(, gContBalanceSectorial)
      prgBarra.ShowForm Me
      prgBarra.CaptionSyle = eCap_CaptionPercent
      prgBarra.Max = R.RecordCount
      Do While Not R.EOF
          nContBal = nContBal + 1
          ReDim Preserve EstBal(nContBal)
          EstBal(nContBal - 1).cCtaCod = Trim(R!cCodigo)
          EstBal(nContBal - 1).cDescrip = Trim(R!cDescrip)
          EstBal(nContBal - 1).cEquival = DepuraEquivalentes(Trim(R!cFormula))
          R.MoveNext
          prgBarra.Progress nContBal, "BALANCE SECTORIAL", "", "Cargando datos... "
      Loop
      prgBarra.CloseForm Me
    RSClose R
    Set oRep = Nothing
Set R = Nothing
End Sub

Private Sub GeneraBalance()
Dim I As Integer
Dim K As Integer
Dim J As Integer
Dim CTemp As String
Dim sSql As String
Dim sSql2 As String
Dim R As New ADODB.Recordset
Dim MNAj As Currency
Dim MNHist As Currency
Dim MExt As Currency
Dim CadSql As String
Dim CadFormula1 As String
Dim CadFormula2 As String
Dim CadFormula3 As String
Dim dFecha As Date
Dim L As ListItem
Dim nFormula As New NInterpreteFormula

   prgBarra.ShowForm Me
   prgBarra.CaptionSyle = eCap_CaptionPercent
   prgBarra.Max = nContBal - 1
   DoEvents
   prgBarra.Progress 0, "BALANCE SECTORIAL", "Cargando...", ""
   oDBal.EliminaBalanceTemp CInt(sTipoRepoFormula), "0"
   oDBal.InsertaBalanceTmpSaldos CInt(sTipoRepoFormula), "0", Format(CDate(sFecha), gsFormatoFecha)
   For I = 0 To nContBal - 1
      'Obtener Cuentas
      CTemp = ""
      nCuentas = 0
      ReDim Cuentas(0)
      For K = 1 To Len(EstBal(I).cEquival)
          If Mid(EstBal(I).cEquival, K, 1) >= "0" And Mid(EstBal(I).cEquival, K, 1) <= "9" Then
              CTemp = CTemp + Mid(EstBal(I).cEquival, K, 1)
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
          'Moneda Nacional Historico
          MNHist = oNBal.CalculaSaldoCuenta(Cuentas(K).cCta, "1", sTipoRepoFormula)
          'Moneda Extranjera
          MExt = oNBal.CalculaSaldoCuenta(Cuentas(K).cCta, "2", sTipoRepoFormula)
          'Moneda Nacional Ajustado
          MNAj = oNBal.CalculaSaldoCuenta(Cuentas(K).cCta, "[136]", sTipoRepoFormula)
      
          'Actualiza Montos
          Cuentas(K).nMExt = MExt
          Cuentas(K).nMNAj = MNAj
          Cuentas(K).nMNHist = MNHist
      Next K
      'Genero las 3 formulas para las 3 monedas
      CTemp = ""
      CadFormula1 = ""
      CadFormula2 = ""
      CadFormula3 = ""
      For K = 1 To Len(EstBal(I).cEquival)
          If (Mid(EstBal(I).cEquival, K, 1) >= "0" And Mid(EstBal(I).cEquival, K, 1) <= "9") Or (Mid(EstBal(I).cEquival, K, 1) = ".") Then
              CTemp = CTemp + Mid(EstBal(I).cEquival, K, 1)
          Else
              If Len(CTemp) > 0 Then
                  'busca su equivalente en monto
                  For J = 0 To nCuentas
                      If Cuentas(J).cCta = CTemp Then
                          CadFormula1 = CadFormula1 + Format(Cuentas(J).nMNAj, "#0")
                          CadFormula2 = CadFormula2 + Format(Cuentas(J).nMNHist, "#0")
                          CadFormula3 = CadFormula3 + Format(Cuentas(J).nMExt, "#0")
                          Exit For
                      End If
                  Next J
              End If
              CTemp = ""
              CadFormula1 = CadFormula1 + Mid(EstBal(I).cEquival, K, 1)
              CadFormula2 = CadFormula2 + Mid(EstBal(I).cEquival, K, 1)
              CadFormula3 = CadFormula3 + Mid(EstBal(I).cEquival, K, 1)
          End If
      Next K
      If Len(CTemp) > 0 Then
          'busca su equivalente en monto
          For J = 0 To nCuentas
              If Cuentas(J).cCta = CTemp Then
                  CadFormula1 = CadFormula1 + Format(Cuentas(J).nMNAj, "#0")
                  CadFormula2 = CadFormula2 + Format(Cuentas(J).nMNHist, "#0")
                  CadFormula3 = CadFormula3 + Format(Cuentas(J).nMExt, "#0")
                  Exit For
              End If
          Next J
      End If
        
      MNAj = Round(nFormula.ExprANum(CadFormula1) / 1000, 0)
      MNHist = Round(nFormula.ExprANum(CadFormula2) / 1000, 0)
      MExt = Round(nFormula.ExprANum(CadFormula3) / 1000, 0)
      
      oDBal.InsertaBalanceSectorial EstBal(I).cCtaCod, MNHist, MExt, MNAj, sFecha
      prgBarra.Progress I, "BALANCE SECTORIAL", "", "Intepretando Fórmulas... "
   Next I
   prgBarra.CloseForm Me
Set nFormula = Nothing
End Sub
Private Function Valida() As Boolean
    If Len(Trim(txtAnio.Text)) = 0 Then
        MsgBox "Ingrese Año del Balance", vbInformation, "Aviso"
        Valida = False
    End If
    If CInt(txtAnio.Text) <= 1950 Then
        MsgBox "Año no valido", vbInformation, "Aviso"
        Valida = False
    End If
    Valida = True
End Function
Private Sub LeeBalanceSec()
Dim sSql As String
Dim R As New ADODB.Recordset
Dim L As ListItem
   Set R = oNBal.CargaBalanceSectorial(sFecha)
   If Not R.EOF Then
      prgBarra.ShowForm Me
      prgBarra.CaptionSyle = eCap_CaptionPercent
      prgBarra.Max = R.RecordCount
   End If
   LstBalance.ListItems.Clear
   Do While Not R.EOF
       Set L = LstBalance.ListItems.Add(, , R!cCodigo)
       L.SubItems(1) = Trim(R!cDescrip)
       L.SubItems(2) = Format(R!nMNAj, gsFormatoNumeroView)
       L.SubItems(3) = Format(R!nMNHist, gsFormatoNumeroView)
       L.SubItems(4) = Format(R!nME, gsFormatoNumeroView)
       prgBarra.Progress LstBalance.ListItems.Count, "BALANCE SECTORIAL", "", "Mostrando Balance... "
       R.MoveNext
   Loop
   If LstBalance.ListItems.Count > 0 Then
      prgBarra.CloseForm Me
   End If
   RSClose R
End Sub

Private Sub cboMes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        fEnfoque txtAnio
        txtAnio.SetFocus
    End If
End Sub

Private Sub cmdAceptar_Click()
    oDBal.ActualizaBalanceSectorial LblCodigo.Caption, nVal(txtMNHist.Text), nVal(TxtMExt.Text), nVal(TxtMNaj), sFecha
    LstBalance.SelectedItem.SubItems(2) = Format(TxtMNaj.Text, gsFormatoNumeroView)
    LstBalance.SelectedItem.SubItems(3) = Format(txtMNHist.Text, gsFormatoNumeroView)
    LstBalance.SelectedItem.SubItems(4) = Format(TxtMExt.Text, gsFormatoNumeroView)
    ActivaBotonEditar False
    LstBalance.SetFocus
End Sub

Private Sub cmdArchivo_Click()
Dim sImpre As String
If LstBalance.ListItems.Count = 0 Then
    'MsgBox "No existen datos para imprimir", vbInformation, "Aviso"
    Exit Sub
End If

    Screen.MousePointer = 11
    sImpre = oNBal.ImprimeBalanceSectorial(sFecha, gsNomCmac, True)
    Screen.MousePointer = 0
EnviaPrevio sImpre, "BALANCE SECTORIAL", gnLinPage, False
End Sub

Private Sub cmdCancelar_Click()
    LblCodigo.Caption = ""
    Lbldescrip.Caption = ""
    TxtMNaj.Text = ""
    txtMNHist.Text = ""
    TxtMExt.Text = ""
    LstBalance.SetFocus
    ActivaBotonEditar False
    LstBalance.SetFocus
End Sub

Private Sub CmdEditar_Click()
If LstBalance.ListItems.Count > 0 Then
   ActivaBotonEditar True
   LblCodigo.Caption = LstBalance.SelectedItem.Text
   Lbldescrip.Caption = LstBalance.SelectedItem.SubItems(1)
   TxtMNaj = LstBalance.SelectedItem.SubItems(2)
   txtMNHist = LstBalance.SelectedItem.SubItems(3)
   TxtMExt = LstBalance.SelectedItem.SubItems(4)
   TxtMNaj.SetFocus
 End If
End Sub

Private Sub cmdImprimir_Click()
'Dim sImpre As String
'    Screen.MousePointer = 11
If LstBalance.ListItems.Count = 0 Then
    MsgBox "No existen datos para imprimir", vbInformation, "Aviso"
    Exit Sub
End If
    Call ImprimeBalanceSectorial(sFecha, gsNomCmac, False)
    
'    sImpre = oNBal.ImprimeBalanceSectorial(sFecha, gsNomCmac, False)
'    Screen.MousePointer = 0
'EnviaPrevio sImpre, "BALANCE SECTORIAL", gnLinPage, True
End Sub

Private Sub ActivaBotonEditar(lActiva As Boolean)
    cmdAceptar.Visible = lActiva
    CmdCancelar.Visible = lActiva
    cmdArchivo.Visible = Not lActiva
    cmdEditar.Visible = Not lActiva
    cmdImprimir.Visible = Not lActiva
    cmdSalir.Visible = Not lActiva
    If lActiva Then
      LstBalance.Height = 3800
    Else
      LstBalance.Height = 4215
    End If
End Sub

Private Sub ActivaBotonBarra(lActiva As Boolean)
    cmdEditar.Visible = Not lActiva
    cmdImprimir.Visible = Not lActiva
    cmdSalir.Visible = Not lActiva
End Sub

Private Sub cmdProcesar_Click()
Dim R As Double
Dim Cad As String
Dim sSql As String
    If Not Valida Then
        Exit Sub
    End If
    sFecha = DateAdd("m", 1, CDate("01/" & Format(CboMes.ListIndex + 1, "00") & "/" & Format(txtAnio.Text, "0000"))) - 1
    ActivaBotonBarra True
    If oNBal.BalanceGeneradoSectorial(CDate(sFecha)) Then
       If MsgBox("Balance Sectorial de esta Fecha ya fue Generado. ¿ Desea Generar nuevamente ?", vbQuestion + vbYesNo, "¡Aviso!") = vbYes Then
            oDBal.EliminaBalanceSectorial CDate(sFecha)
            Call CargaDatos
            Call GeneraBalance
        End If
    Else
        Call CargaDatos
        Call GeneraBalance
    End If
    Call LeeBalanceSec
    ActivaBotonBarra False
    LstBalance.SetFocus
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
frmMdiMain.Enabled = False

CentraForm Me
    CboMes.ListIndex = Month(gdFecSis) - 1
    Set oNBal = New NBalanceCont
    Set oDBal = New DbalanceCont
sTipoRepoFormula = "3"
txtAnio = Year(gdFecSis)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set oNBal = Nothing
Set oDBal = Nothing
frmMdiMain.Enabled = True

End Sub

Private Sub txtAnio_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        cmdProcesar.SetFocus
    End If
End Sub

Private Sub TxtMExt_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(TxtMExt, KeyAscii, 16, 2)
    If KeyAscii = 13 Then
        cmdAceptar.SetFocus
    End If
End Sub

Private Sub TxtMNaj_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(TxtMNaj, KeyAscii, 16, 2)
    If KeyAscii = 13 Then
        txtMNHist.SetFocus
    End If
End Sub

Private Sub txtMNHist_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtMNHist, KeyAscii, 16, 2)
    If KeyAscii = 13 Then
        TxtMExt.SetFocus
    End If
End Sub
